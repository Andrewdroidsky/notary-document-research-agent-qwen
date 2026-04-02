from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import textwrap
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
ET.register_namespace("w", W_NS)

DEFAULT_AGENTS_PATH = Path("AGENTS.md")
DEFAULT_STATE_PATH = Path("PROJECT_STATE.md")
DEFAULT_WORKFLOW_PATH = Path("input/workflow/manual-workflow.md")
DEFAULT_MASTER_PROMPT_PATH = Path("input/master prompt/Промпт по поиску документов 18.md")
DEFAULT_APPROVED_TOPICS_PATH = Path("input/workflow/Утверждаю.md")
DEFAULT_ORDER_TEMPLATE_PATH = Path("input/order/Текст приказа 18  15.11.10..md")
DEFAULT_INTERACTION_GUIDE_PATH = Path("input/workflow/User and LLM Interaction in LLM Contest Window.md")
DEFAULT_OUTLINE_OVERRIDES_ROOT = Path("input/workflow/outline overrides")
DEFAULT_OUTPUT_EXAMPLE_MD_PATH = Path("input/output examples/15.11.10. Наследование имущества.md")
DEFAULT_OUTPUT_EXAMPLE_DOCX_PATH = Path("input/output examples/15.11.10. Наследование имущества.docx")
DEFAULT_OUTPUT_READY_ROOT = Path("input/output ready")
ARTIFACT_PROFILE_ENV = "NOTARY_AGENT_ARTIFACT_PROFILE"
ARTIFACT_PROFILE_LEAN = "lean"
ARTIFACT_PROFILE_FULL = "full"
DEFAULT_ARTIFACT_PROFILE = os.environ.get(ARTIFACT_PROFILE_ENV, ARTIFACT_PROFILE_LEAN).strip().lower() or ARTIFACT_PROFILE_LEAN
PACKET_MODE_ENV = "NOTARY_AGENT_PACKET_MODE"
PACKET_MODE_LEAN = "lean-default"
PACKET_MODE_LITERAL = "literal-safe"
DEFAULT_PACKET_MODE = os.environ.get(PACKET_MODE_ENV, PACKET_MODE_LEAN).strip().lower() or PACKET_MODE_LEAN
RUN_MODE_FRESH = "fresh-only"
RUN_MODE_SURGICAL = "surgical-redo"
STRICT_ORCHESTRATION_TARGET_PERCENT = 3.0
PUBLISH_MIN_WORDS = 16000
PUBLISH_MIN_CHARS = 122000
PUBLISH_MIN_URL1 = 100
PUBLISH_MIN_URL2 = 100
SURGICAL_ALLOWED_COMMANDS = {
    "capture-part-output",
    "prepare-part-02-web",
    "prepare-part-03-plan",
    "capture-part-03-range",
    "metric-check",
}

PART_EXECUTION_MODES = {
    1: "rules_confirmation_wait_go",
    2: "core_answer_after_go",
    3: "targeted_follow_up",
    4: "coverage_blocks",
    5: "layer_blocks",
    6: "three_filters_gap_check",
    7: "new_documents_without_repeats",
    8: "federal_level_delta_check",
    9: "delta_audit",
    10: "mini_summary",
    11: "diary_tasks_only",
}

FOLLOWUP_LITERAL_PART_NUMBERS = (6, 7, 8, 9, 10, 11)
REASONING_LAYER_PART_NUMBERS = tuple(range(2, 12))
OMISSION_AUDIT_PART_NUMBERS = (9, 10, 11)
SEMANTIC_DEDUP_PART_NUMBERS = (5, 6, 7, 8)
PRODUCTION_ENABLE_SEMANTIC_DEDUP = False
PRODUCTION_ENABLE_OMISSION_AUDIT = False

REASONING_LAYER_GUIDANCE = {
    2: {
        "goal": "Разложить подтему на юридические узлы, проверить правовое ядро и собрать доказательное ядро ответа без преждевременного сужения поиска.",
        "moves": [
            "Сначала декомпозировать подтему на объект, действие нотариуса, участников, процедуру, деньги, ограничения, электронный и международный слой.",
            "Строить не только прямые, но и скрытые гипотезы поиска: кодексный, профильный, подзаконный, ведомственный, ФНП, судебный, международный.",
            "Обязательно искать редакционную цепочку и свежие акты, которые меняют базовые приказы или профильные документы по теме.",
            "По каждому документу различать роль: прямоприменимый, опорный, карантинный или fail-safe фактор.",
            "Не ограничиваться стартовыми query-подсказками, если смысл темы требует расширения поиска.",
            "Решение о включении документа принимать по существованию, применимости, роли в теме и подтвержденному URL1; URL2 усиливает карточку, но сам по себе не решает вопрос включения.",
        ],
        "risks": [
            "Потеря скрытого слоя темы из-за слишком раннего выбора 1-2 базовых актов.",
            "Потеря свежих приказов об изменениях и редакционных цепочек к уже найденным актам.",
            "Смешение substantive и tariff подтем в одном ответе.",
            "Формальный вывод документа без реальной роли в теме.",
        ],
    },
    3: {
        "goal": "Пройти канонические блоки I-XXXVII как карту сфер права и не закрывать блок формальной отпиской без содержательной проверки.",
        "moves": [
            "Для каждого блока сначала решать, применим ли он к теме по смыслу, а уже потом искать документы.",
            "Если блок применим, искать не только прямой акт, но и смежный документный слой внутри той же сферы.",
            "Использовать fail-safe по блоку только после реального прохода, а не по инерции.",
            "Отдельно контролировать обязательные блоки II, III, IV, X, XI, XII и XIV даже при слабой очевидности темы.",
        ],
        "risks": [
            "Формальное 'не выявлено' без реального прохода по блоку.",
            "Подмена канонической сетки I-XXXVII выдуманными заголовками.",
            "Повтор Part 2 вместо карты сфер права.",
        ],
    },
    4: {
        "goal": "Построить карту органов и источников власти по их реальной роли в теме, а не по случайной встречаемости в поиске.",
        "moves": [
            "Для каждого подпункта 3.1 определять роль органа: нормотворец, контролер, оператор системы, держатель реестра, методический центр или международный участник.",
            "Отделять акт органа от площадки публикации или вторичного пересказа.",
            "Проверять, есть ли у органа собственный обязательный слой по теме или только косвенное упоминание.",
            "Если по органу найден только вторичный материал, не выдавать это как полноценный нормативный результат.",
        ],
        "risks": [
            "Смешение органов власти и площадок публикации.",
            "Приписывание органу регулирующей роли, которой у него нет.",
            "Потеря смежных органов при международном, цифровом или контрольном слое темы.",
        ],
    },
    5: {
        "goal": "Разложить найденный материал по слоям и сохранить только смыслово новые документы в каждом слое.",
        "moves": [
            "Для каждого слоя определять, что именно он добавляет к теме по сравнению с уже найденными документами.",
            "Дедуплицировать по смысловой роли документа, а не только по названию.",
            "Если документ уже найден, но в новом слое играет иную роль, прямо фиксировать эту роль, а не молча повторять карточку.",
            "Следить, чтобы слойность не превращалась в повтор Part 2-4 под новым заголовком.",
        ],
        "risks": [
            "Повтор одного и того же акта в разных слоях без нового содержания.",
            "Потеря действительно нового слоя из-за формального дедупа.",
            "Смешение кодексного, подзаконного и методического слоя.",
        ],
    },
    6: {
        "goal": "Применить фильтры к уже собранной базе и сохранить только документы, которые реально усиливают покрытие темы.",
        "moves": [
            "Смотреть на документ как на носитель роли: ядро, смежный слой, контрольный слой, международный слой, электронный слой.",
            "Отбрасывать повторы и слабые кандидаты, но оставлять документы, закрывающие самостоятельный юридический риск.",
            "Сначала делать дополнительный внешний поиск и только потом фильтровать уже найденное; не превращать Часть 6 в короткую отписку без повторного прохода.",
            "После фильтрации перепроверять, не появился ли новый пробел по обязательному узлу темы.",
        ],
        "risks": [
            "Слишком агрессивная фильтрация с потерей полезных документов.",
            "Слабый документ оставлен только потому, что найден позже других.",
            "Ответ `новых документов не выявлено` без реального повторного web-search по узлам темы.",
            "Фильтр без повторной проверки на пробелы.",
        ],
    },
    7: {
        "goal": "Добавить только действительно новые документы без косметического размножения уже найденной базы.",
        "moves": [
            "Считать новым только документ, который дает новый нормативный слой, новый орган, новый режим или новый структурный элемент.",
            "Не принимать за 'новый' тот же акт в другой редакции или на другой площадке без новой смысловой функции.",
            "Перед выводом обязательно делать свежий web-search по редакционным цепочкам, изменениям и свежим публикациям, а не только пересматривать уже собранный массив.",
            "Проверять, закрывает ли документ ранее незакрытый риск или узел темы.",
        ],
        "risks": [
            "Добавление псевдо-новых документов ради количества.",
            "Подмена новых документов новыми ссылками на старые акты.",
            "Формальное `новых документов не выявлено` без повторного внешнего поиска.",
            "Потеря точечной, но ценной нормы из-за слишком грубого дедупа.",
        ],
    },
    8: {
        "goal": "Провести федеральный delta-check без повторения уже подтвержденных федеральных документов.",
        "moves": [
            "Сверять каждый федеральный кандидат с уже подтвержденной федеральной базой по роли, реквизитам и структурному элементу.",
            "Оставлять только те федеральные документы, которые реально расширяют покрытие темы.",
            "Обязательно искать новые федеральные акты и акты об изменениях к уже найденным документам, если они влияют на актуальную редакцию.",
            "Если найден новый структурный элемент в уже известном федеральном акте, фиксировать это как дельту, а не как новый акт.",
        ],
        "risks": [
            "Повтор федеральной базы под видом дельты.",
            "Потеря свежих федеральных приказов об изменениях к базовым актам.",
            "Потеря нового структурного элемента внутри уже известного акта.",
            "Смешение федеральной дельты с региональными или методическими источниками.",
        ],
    },
    9: {
        "goal": "Провести содержательный дельта-аудит: что еще упущено, что противоречит, что выглядит недожатым.",
        "moves": [
            "Искать подозрительно пустые узлы темы, особенно после фильтрации и дедупликации.",
            "Сверять обязательные блоки, органы и слои с уже собранным ядром.",
            "Перед выводом делать повторный внешний поиск по слепым зонам, а не только логический пересмотр уже записанного массива.",
            "Фиксировать не только новые документы, но и структурные дыры, слабые верификации и сомнительные карантинные зоны.",
            "Использовать аудит как проверку полноты, а не как формальное повторение списка документов.",
        ],
        "risks": [
            "Формальный аудит без новых выводов.",
            "Потеря противоречий между частями 2-8.",
            "Недостаточный акцент на реальных пропусках и слабых местах доказательной базы.",
        ],
    },
    10: {
        "goal": "Собрать прикладной мини-конспект из уже подтвержденной доказательной базы без фантазирования и без переписывания всего документа.",
        "moves": [
            "Выделять только ключевые нормы, режимы, риски и опорные документы, которые реально управляют темой.",
            "Показывать связи между ядром, смежными слоями, отказами, ограничениями и практическими последствиями.",
            "Сжимать материал до полезного прикладного конспекта, но не высушивать его: каждый тезис должен быть оперт на именованные документы с карточками и ссылками.",
        ],
        "risks": [
            "Повтор всего документа вместо конспекта.",
            "Добавление непроверенных выводов, которых не было в собранной базе.",
            "Сухой narrative без полноценных карточек документов-опор.",
            "Потеря практических рисков и ограничений.",
        ],
    },
    11: {
        "goal": "Построить перечень практических заданий только из реально найденного нормативного материала и рисков темы.",
        "moves": [
            "Выводить задания из подтвержденных документов, процедур, ограничений, отказов, проверок и международных режимов темы.",
            "Давать задания прикладные, проверяемые и привязанные к найденным актам и структурным элементам.",
            "Формулировать перечень как записи для 4-й колонки дневника о выполненных стажером действиях по характеру задания.",
            "Не генерировать учебный шум, если его нельзя опереть на уже собранную базу.",
        ],
        "risks": [
            "Абстрактные задания без опоры на документы темы.",
            "Повтор мини-конспекта вместо практических задач.",
            "Императивные формулировки и уход в другие колонки дневника вместо перечня выполненных стажером действий.",
            "Уход в смежные темы, не подтвержденные текущей базой.",
        ],
    },
}

PART_02_REQUIRED_MARKERS = [
    "АНАЛИЗ ОБЛАСТИ ПРАВА",
    "АНАЛИЗ-КОДЕКСЫ И БАЗОВЫЕ АКТЫ",
    "A. РЕГУЛЯТОРНОЕ ЯДРО",
    "B. ОПОРНЫЕ ДОКУМЕНТЫ",
    "КАРАНТИН",
    "FAIL-SAFE CHECK",
]

FINAL_PART_TITLES = {
    2: "АНАЛИЗ ОБЛАСТИ ПРАВА",
    3: "БЛОКИ ПОИСКА СФЕР ПРАВА",
    4: "ПОИСК ПО ОРГАНАМ ВЛАСТИ",
    5: "СЛОЙНОСТЬ",
    6: "ФИЛЬТР",
    7: "ФИЛЬТР БЕЗ ПОВТОРЕНИЙ",
    8: "ФИЛЬТР БЕЗ ПОВТОРЕНИЙ ФЕДЕРАЛЬНЫХ ДОКУМЕНТОВ",
    9: "ДЕЛЬТА-АУДИТ",
    10: "РАСШИРЕННЫЙ МИНИ-КОНСПЕКТ",
    11: "ПЕРЕЧЕНЬ ПРАКТИЧЕСКИХ ЗАДАНИЙ ПО ЗАДАННОЙ ТЕМЕ",
}

FRESH_RUN_TRUSTED_SOURCE_ORIGINS = {
    "execute_part_01",
    "capture_part_output",
    "capture_part_03_range",
    "capture_part_04_range",
    "capture_part_05_range",
}

PROMINENT_INNER_HEADINGS = {
    "ТЕМА:",
    "АНАЛИЗ ОБЛАСТИ ПРАВА",
    "АНАЛИЗ-КОДЕКСЫ И БАЗОВЫЕ АКТЫ",
    "A. РЕГУЛЯТОРНОЕ ЯДРО",
    "B. ОПОРНЫЕ ДОКУМЕНТЫ",
    "КАРАНТИН",
    "FAIL-SAFE CHECK",
}

PART_02_FORBIDDEN_PUBLIC_BLOCKS_RE = re.compile(
    r"(?m)^\s*\*{0,2}(?:I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI|XVII|XVIII|XIX|XX|XXI|XXII|XXIII|XXIV|XXV|XXVI|XXVII|XXVIII|XXIX|XXX|XXXI|XXXII|XXXIII|XXXIV|XXXV|XXXVI|XXXVII)(?:\.|\s*[—-])"
)

PART_02_SOURCE_CASCADE = [
    {
        "rank": 1,
        "role": "URL1 official anchor",
        "domains": ["publication.pravo.gov.ru", "pravo.gov.ru"],
        "usage": "Официальная публикация или официальный правовой портал как якорь официальности.",
    },
    {
        "rank": 2,
        "role": "Official topical layer",
        "domains": ["minjust.gov.ru", "notariat.ru", "government.ru", "kremlin.ru"],
        "usage": "Профильные официальные источники: Минюст, ФНП, Правительство, Президент.",
    },
    {
        "rank": 3,
        "role": "Readable URL2 candidates",
        "domains": ["consultant.ru", "garant.ru", "docs.cntd.ru"],
        "usage": "Читаемые источники для VERIFIED URL2 после сверки реквизитов и заголовка страницы.",
    },
]

PART_03_SEGMENTS = [
    {
        "segment_id": 1,
        "label": "I–V",
        "request_text": "Запрос 1: Обработать блоки I–V. Строго по Приказу. По завершении доложить, запросить обработку следующего блока.",
    },
    {
        "segment_id": 2,
        "label": "VI–X",
        "request_text": "Запрос 2: Продолжить. Блоки VI–X. Строго по Приказу. По завершении доложить, запросить обработку следующего блока.",
    },
    {
        "segment_id": 3,
        "label": "XI–XV",
        "request_text": "Запрос 3: Продолжить. Блоки XI–XV. Строго по Приказу. По завершении доложить, запросить обработку следующего блока.",
    },
    {
        "segment_id": 4,
        "label": "XVI–XX",
        "request_text": "Запрос 4: Продолжить. Блоки XVI–XX. Строго по Приказу. По завершении доложить, запросить обработку следующего блока.",
    },
    {
        "segment_id": 5,
        "label": "XXI–XXV",
        "request_text": "Запрос 5: Продолжить. Блоки XXI–XXV. Строго по Приказу. По завершении доложить, запросить обработку следующего блока.",
    },
    {
        "segment_id": 6,
        "label": "XXVI–XXX",
        "request_text": "Запрос 6: Продолжить. Блоки XXVI–XXX. Строго по Приказу. По завершении доложить, запросить обработку следующего блока.",
    },
    {
        "segment_id": 7,
        "label": "XXXI–XXXV",
        "request_text": "Запрос 7: Продолжить. Блоки XXXI–XXXV. Строго по Приказу. По завершении доложить, запросить обработку следующего блока.",
    },
    {
        "segment_id": 8,
        "label": "XXXVI–XXXVII",
        "request_text": "Запрос 8: Закончить: Блоки XXXVI–XXXVII. Строго по Приказу. Доложить о полном завершении обработки блоков.",
    },
]

PART_03_CANONICAL_BLOCKS = {
    "I": "Базовые отрасли материального права",
    "II": "Процессуальное право (ОБЯЗАТЕЛЬНО)",
    "III": "Налоговое право (ОБЯЗАТЕЛЬНО)",
    "IV": "Финансовое право (ОБЯЗАТЕЛЬНО)",
    "V": "Административное право и контроль",
    "VI": "Информационное право и цифровая среда",
    "VII": "Архивное право и делопроизводство",
    "VIII": "Противодействие легализации доходов (ПОД/ФТ)",
    "IX": "Международное частное право",
    "X": "Подзаконные акты (ОБЯЗАТЕЛЬНО УЧИТЫВАТЬ)",
    "XI": "Акты нотариального сообщества",
    "XII": "Судебные разъяснения",
    "XIII": "Комплексные отрасли",
    "XIV": "Нотариальное право и организация нотариата",
    "XV": "Регистрационное право и государственные реестры",
    "XVI": "Жилищное право",
    "XVII": "Наследственное право",
    "XVIII": "Вещное право и сделки с имуществом",
    "XIX": "Банкротство",
    "XX": "Валютное регулирование и комплаенс сделок",
    "XXI": "Доказательственное право и обеспечение доказательств",
    "XXII": "Интеллектуальные права",
    "XXIII": "Потребительские и “соц”-кейсы в нотариальной форме",
    "XXIV": "Таможенное право",
    "XXV": "Уголовное право",
    "XXVI": "Экологическое право",
    "XXVII": "Вещное право",
    "XXVIII": "Обязательственное право",
    "XXIX": "Рекламное право",
    "XXX": "Природоресурсное право",
    "XXXI": "Градостроительное право",
    "XXXII": "Воздушное право / гражданская авиация",
    "XXXIII": "Транспортное право: морское и внутренний водный транспорт",
    "XXXIV": "Бюджетное право",
    "XXXV": "Антимонопольное / конкурентное право",
    "XXXVI": "Миграционное и гражданство",
    "XXXVII": "Энергетическое право",
}

PART_03_FORBIDDEN_LEGACY_HEADINGS = (
    "БАЗОВЫЕ ПРОФИЛЬНЫЕ АКТЫ",
    "СПЕЦИАЛЬНЫЕ НОТАРИАЛЬНЫЕ ПОЛНОМОЧИЯ НОТАРИУСА",
    "ПОЛНОМОЧИЯ ДОЛЖНОСТНЫХ ЛИЦ МЕСТНОГО САМОУПРАВЛЕНИЯ",
    "ПОЛНОМОЧИЯ КОНСУЛЬСКИХ ДОЛЖНОСТНЫХ ЛИЦ",
    "ПОРЯДОК СОВЕРШЕНИЯ НОТАРИАЛЬНОГО ДЕЙСТВИЯ",
    "ТРЕБОВАНИЯ К ПРЕДСТАВЛЯЕМЫМ ДОКУМЕНТАМ",
    "ТРЕБОВАНИЯ К НОТАРИАЛЬНО ОФОРМЛЯЕМОМУ ДОКУМЕНТУ",
    "УДОСТОВЕРИТЕЛЬНЫЕ НАДПИСИ И ФОРМЫ",
    "РЕГИСТРАЦИЯ НОТАРИАЛЬНОГО ДЕЙСТВИЯ",
    "ДЕЛОПРОИЗВОДСТВО И ХРАНЕНИЕ",
    "УСТАНОВЛЕНИЕ ЛИЧНОСТИ ЗАЯВИТЕЛЯ",
    "ПРОВЕРКА ДЕЕСПОСОБНОСТИ/ПРАВОСПОСОБНОСТИ/ПОЛНОМОЧИЙ",
    "ОТКАЗ В СОВЕРШЕНИИ ДЕЙСТВИЯ",
    "СУДЕБНОЕ ОБЖАЛОВАНИЕ",
    "ФЕДЕРАЛЬНЫЙ ТАРИФ",
    "НАЛОГОВЫЕ ЛЬГОТЫ",
    "РЕГИОНАЛЬНЫЙ ТАРИФ",
    "ЭЛЕКТРОННАЯ ФОРМА И ЕИС",
    "УДАЛЕННОЕ СОВЕРШЕНИЕ",
    "МЕЖДУНАРОДНЫЙ ЭЛЕМЕНТ / ИНОСТРАННЫЕ ДОКУМЕНТЫ",
    "ЛЕГАЛИЗАЦИЯ / АПОСТИЛЬ",
    "НОТАРИАЛЬНЫЕ ДОКУМЕНТЫ И АРХИВ",
    "ПОДЗАКОННЫЕ АКТЫ МИНЮСТА",
    "АКТЫ ФНП",
    "РАЗЪЯСНЕНИЯ ВЫСШЕЙ СУДЕБНОЙ ИНСТАНЦИИ",
    "ВЕДОМСТВЕННЫЕ АКТЫ ИНЫХ ОРГАНОВ",
    "СПЕЦИАЛЬНЫЕ ЗАПРЕТЫ И ОГРАНИЧЕНИЯ НА ВИДЫ ДОКУМЕНТОВ",
    "ЯЗЫК ДОКУМЕНТА / ПЕРЕВОД",
    "КОПИЯ С КОПИИ",
    "ВЫПИСКА ИЗ ДОКУМЕНТА",
    "МНОГОСТРАНИЧНЫЕ ДОКУМЕНТЫ, СКРЕПЛЕНИЕ, ЦЕЛОСТНОСТЬ",
    "ОТМЕТКА О ТОМ, ЧТО ЧАСТЬ ОРИГИНАЛА СОДЕРЖИТ КОПИЮ ИНОГО ДОКУМЕНТА",
    "ПРОВЕРКА СОДЕРЖАНИЯ ДОКУМЕНТА НА ЗАКОННОСТЬ",
    "СПЕЦИАЛЬНЫЕ ПРАВИЛА ДЛЯ ОРГАНОВ МСУ",
    "СПЕЦИАЛЬНЫЕ ПРАВИЛА ДЛЯ КОНСУЛОВ",
    "ДИСЦИПЛИНАРНО-КОНТРОЛЬНЫЙ СЛОЙ",
    "РЕГИОНАЛЬНО-ЛОКАЛЬНЫЙ СЛОЙ НОТАРИАЛЬНЫХ ПАЛАТ",
)

PART_04_SEGMENTS = [
    {
        "segment_id": 1,
        "label": "1–5",
        "request_text": "Запрос 1: Обработать подпункты 1–5. Строго по Приказу. По завершении доложить, запросить обработку следующего блока подпунктов.",
    },
    {
        "segment_id": 2,
        "label": "6–10",
        "request_text": "Запрос 2: Продолжить. подпункты 6–10. Строго по Приказу. По завершении доложить, запросить обработку следующего блока подпунктов.",
    },
    {
        "segment_id": 3,
        "label": "11–15",
        "request_text": "Запрос 3: Продолжить. подпункты 11–15. Строго по Приказу. По завершении доложить, запросить обработку следующего блока подпунктов.",
    },
    {
        "segment_id": 4,
        "label": "16–18",
        "request_text": "Запрос 4: Продолжить. подпункты 16–18. Строго по Приказу. Доложить о полном завершении обработки блоков подпунктов.",
    },
]

PART_05_SEGMENTS = [
    {
        "segment_id": 1,
        "label": "1–2",
        "request_text": "Запрос 1: Обработать слои 1–2. Строго по Приказу. По завершении доложить, запросить обработку следующего блока слоев.",
    },
    {
        "segment_id": 2,
        "label": "3–4",
        "request_text": "Запрос 2: Продолжить. Слои 3–4. Строго по Приказу. По завершении доложить, запросить обработку следующего блока слоев.",
    },
    {
        "segment_id": 3,
        "label": "5–6",
        "request_text": "Запрос 3: Продолжить. Слои 5–6. Строго по Приказу. Доложить о полном завершении обработки слоев.",
    },
]

PART_04_CONTROL_POINTS = [
    "Росфинмониторинг",
    "Банк России",
    "Акты Федеральной нотариальной палаты",
    "Акты нотариальных палат субъектов",
    "Методические материалы/рекомендации (в составе акта ФНП)",
    "Росреестр",
    "Федеральная налоговая служба",
    "Минюст РФ/территориальные органы",
    "МВД",
    "Минфин",
    "ФОИВ",
    "Правительство РФ",
    "Президент РФ",
    "ГОСТ/Росстандарт",
    "Архивные правила",
    "Персональные данные/ИБ",
    "Электронная подпись",
    "Иное",
]

PART_05_LAYER_POINTS = [
    "слой 1: базовые кодексы/законы (при применимости)",
    "слой 2: специальное нотариальное регулирование",
    "слой 3: процессуальный/контрольный слой",
    "слой 4: подзаконные НПА уполномоченных органов",
    "слой 5: акты нотариального сообщества, обязательные для нотариусов",
    "слой 6: судебные разъяснения",
]

QUERY_STOPWORDS = {
    "и",
    "или",
    "в",
    "во",
    "на",
    "по",
    "при",
    "для",
    "из",
    "с",
    "со",
    "о",
    "об",
    "от",
    "до",
    "над",
    "под",
    "надо",
    "это",
    "как",
    "др",
    "иные",
    "них",
    "него",
    "нее",
    "отдельных",
    "видов",
    "размера",
    "федерального",
    "регионального",
    "тарифа",
}


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def safe_slug(value: str) -> str:
    value = value.strip().replace("\\", "-").replace("/", "-")
    value = re.sub(r"\s+", "-", value)
    value = re.sub(r"[^\w\-.А-Яа-яЁё]+", "-", value, flags=re.UNICODE)
    value = re.sub(r"-{2,}", "-", value)
    return value.strip("-.") or "topic"


def safe_filename_component(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "", value)
    cleaned = re.sub(r"\s+", " ", cleaned).strip().rstrip(".")
    return cleaned or "document"


def shorten_subtopic_title(title: str, max_length: int = 46) -> str:
    cleaned = safe_filename_component(clean_markdown_text(title))
    if len(cleaned) <= max_length:
        return cleaned
    truncated = cleaned[: max_length + 1]
    if " " in truncated:
        truncated = truncated[: truncated.rfind(" ")]
    truncated = truncated.rstrip(" .,:;!-")
    return truncated or cleaned[:max_length].rstrip(" .,:;!-")


def build_published_output_filename(entry: OutlineEntry, suffix: str) -> str:
    short_title = shorten_subtopic_title(entry.title)
    return f"{entry.item_id}. {short_title}{suffix}"


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def write_json(path: Path, payload: Any) -> None:
    write_text(path, json.dumps(payload, ensure_ascii=False, indent=2))


def read_text(path: Path) -> str:
    encodings = ["utf-8-sig", "utf-8", "cp1251", "utf-16", "latin-1"]
    for encoding in encodings:
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("text", b"", 0, 1, f"Unable to decode {path}")


def read_docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        xml_bytes = archive.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    paragraphs: list[str] = []
    for paragraph in root.findall(f".//{{{W_NS}}}p"):
        texts = []
        for text_node in paragraph.findall(f".//{{{W_NS}}}t"):
            texts.append(text_node.text or "")
        line = "".join(texts).strip()
        if line:
            paragraphs.append(line)
    return "\n".join(paragraphs)


def parse_text_source(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return read_docx_text(path)
    if suffix in {".md", ".txt", ".json"}:
        return read_text(path)
    raise ValueError(f"Unsupported text source: {path}")


def simple_frontmatter(text: str) -> tuple[dict[str, str], str]:
    if not text.startswith("---\n"):
        return {}, text
    parts = text.split("\n---\n", 1)
    if len(parts) != 2:
        return {}, text
    raw_meta, body = parts
    meta: dict[str, str] = {}
    for line in raw_meta.splitlines()[1:]:
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        meta[key.strip()] = value.strip()
    return meta, body.lstrip("\n")


def to_frontmatter(meta: dict[str, str], body: str) -> str:
    if not meta:
        return body
    lines = ["---"]
    for key, value in meta.items():
        lines.append(f"{key}: {value}")
    lines.extend(["---", "", body.rstrip(), ""])
    return "\n".join(lines)


def clean_markdown_text(value: str) -> str:
    value = value.replace("**", "")
    value = value.replace("\\.", ".")
    value = re.sub(r"<br\s*/?>", " ", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+", " ", value, flags=re.UNICODE)
    return value.strip()


def normalize_search_key(value: str) -> str:
    value = clean_markdown_text(value).lower().replace("ё", "е")
    value = re.sub(r"[^\wа-я]+", " ", value, flags=re.UNICODE)
    return re.sub(r"\s+", " ", value).strip()


def ensure_trailing_period(value: str) -> str:
    value = value.strip()
    if not value:
        return value
    return value if value.endswith(".") else value + "."


def trim_terminal_period(value: str) -> str:
    return value.strip().rstrip(".").strip()


def is_tariff_title(title: str) -> bool:
    lowered = trim_terminal_period(title).lower()
    return lowered.startswith("исчисление размера федерального и регионального тарифа")


def extract_tariff_action_phrase(title: str) -> str:
    cleaned = trim_terminal_period(title)
    match = re.search(r"\bпри\s+(.+)$", cleaned, flags=re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return cleaned


def infer_query_keywords(title: str) -> list[str]:
    raw_tokens = re.split(r"[\s,;:()«»\"-]+", trim_terminal_period(title))
    keywords: list[str] = []
    for token in raw_tokens:
        normalized = token.strip().lower()
        if not normalized or len(normalized) < 3:
            continue
        if normalized in QUERY_STOPWORDS:
            continue
        if normalized not in keywords:
            keywords.append(normalized)
    return keywords


def write_text_if_needed(path: Path, content: str, overwrite: bool) -> None:
    if path.exists() and not overwrite:
        return
    write_text(path, content)


def parse_markdown_table_row(line: str) -> list[str]:
    stripped = line.strip()
    if not stripped.startswith("|"):
        return []
    return [cell.strip() for cell in stripped.split("|")[1:-1]]


def copy_if_exists(source: Path, dest: Path) -> None:
    if not source.exists():
        return
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, dest)


def detect_topic_id_and_title(name: str) -> tuple[str | None, str]:
    match = re.match(r"^(?P<id>\d+(?:\.\d+)+)\.?\s*(?P<title>.*)$", name.strip())
    if match:
        topic_id = match.group("id").strip()
        title = match.group("title").strip(" .")
        return topic_id, title or name.strip()
    return None, name.strip()


def discover_files(source_dir: Path) -> dict[str, list[Path]]:
    prompt_files: list[Path] = []
    order_files: list[Path] = []
    note_files: list[Path] = []
    for item in sorted(source_dir.iterdir(), key=lambda p: p.name.lower()):
        if item.is_dir():
            continue
        name_lower = item.name.lower()
        if name_lower.startswith("промпт"):
            prompt_files.append(item)
            continue
        if name_lower.startswith("текст приказа"):
            order_files.append(item)
            continue
        if item.suffix.lower() in {".docx", ".md", ".txt"}:
            note_files.append(item)
    return {
        "prompt_files": prompt_files,
        "order_files": order_files,
        "note_files": note_files,
    }


def choose_primary_template(order_files: list[Path], note_files: list[Path]) -> Path | None:
    if order_files:
        return order_files[0]
    for path in note_files:
        if path.suffix.lower() == ".docx":
            return path
    return None


def load_json_config(path: Path | None) -> dict[str, Any]:
    defaults: dict[str, Any] = {
        "model": "gpt-5",
        "reasoning_effort": "low",
        "store": True,
        "include_sources": True,
        "request_timeout_seconds": 600,
        "output_root": "runs",
        "web_allowed_domains": [],
    }
    if path is None:
        return defaults
    loaded = json.loads(read_text(path))
    defaults.update(loaded)
    return defaults


def normalize_import_dest(dest_root: Path, topic_id: str | None, title: str) -> Path:
    folder_name = topic_id or safe_slug(title)
    return dest_root / folder_name


def markdown_from_source(path: Path) -> str:
    body = parse_text_source(path).strip()
    header = f"# {path.name}\n\n"
    return header + body + "\n"


def import_topic(source_dir: Path, dest_root: Path) -> Path:
    files = discover_files(source_dir)
    topic_id, title = detect_topic_id_and_title(source_dir.name)
    if not title:
        title = source_dir.name
    target_dir = normalize_import_dest(dest_root, topic_id, title)
    target_dir.mkdir(parents=True, exist_ok=True)

    prompt_file = files["prompt_files"][0] if files["prompt_files"] else None
    template_file = choose_primary_template(files["order_files"], files["note_files"])

    meta = {
        "id": topic_id or "",
        "title": title,
        "source_dir": str(source_dir),
    }
    if template_file is not None:
        meta["template_docx"] = "assets/template.docx"

    body_lines = [
        f"# {topic_id + '. ' if topic_id else ''}{title}".strip(),
        "",
        "Рабочее описание темы для агента.",
        "",
        "Здесь можно вручную дописать нюансы по теме, если они не попали в исходные документы.",
    ]
    write_text(target_dir / "topic.md", to_frontmatter(meta, "\n".join(body_lines)))

    if prompt_file is not None:
        write_text(target_dir / "master-prompt.md", markdown_from_source(prompt_file))

    orders_raw_dir = target_dir / "orders" / "raw"
    orders_clean_dir = target_dir / "orders" / "clean"
    for index, order_file in enumerate(files["order_files"], start=1):
        content = markdown_from_source(order_file)
        filename = f"part-{index:02d}.md"
        write_text(orders_raw_dir / filename, content)
        clean_stub = (
            "<!--\n"
            "Приведите приказ в порядок перед запуском агента.\n"
            "Удалите лишние абзацы, следы старых диалогов, промежуточные команды и дубли.\n"
            "Агент читает именно эту cleaned-версию.\n"
            "-->\n\n"
            + content
        )
        write_text(orders_clean_dir / filename, clean_stub)

    notes_dir = target_dir / "source-notes"
    for note_file in files["note_files"]:
        write_text(notes_dir / f"{safe_slug(note_file.stem)}.md", markdown_from_source(note_file))

    if template_file is not None and template_file.suffix.lower() == ".docx":
        assets_dir = target_dir / "assets"
        assets_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template_file, assets_dir / "template.docx")

    manifest = {
        "imported_at": utc_now_iso(),
        "source_dir": str(source_dir),
        "topic_id": topic_id,
        "title": title,
        "prompt_files": [str(p) for p in files["prompt_files"]],
        "order_files": [str(p) for p in files["order_files"]],
        "note_files": [str(p) for p in files["note_files"]],
        "template_source": str(template_file) if template_file else None,
    }
    write_text(target_dir / "manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
    return target_dir


def scan_source_tree(root: Path) -> list[Path]:
    candidates: list[Path] = []
    for current_root, dirs, files in os.walk(root):
        del dirs
        file_names = set(files)
        if not any(name.lower().startswith("промпт") or name.lower().startswith("текст приказа") for name in file_names):
            continue
        current = Path(current_root)
        topic_id, _ = detect_topic_id_and_title(current.name)
        if topic_id:
            candidates.append(current)
    return sorted(set(candidates))


@dataclass
class TopicBundle:
    topic_dir: Path
    metadata: dict[str, str]
    topic_body: str
    master_prompt: str
    order_parts: list[tuple[str, str]] = field(default_factory=list)
    source_notes: list[tuple[str, str]] = field(default_factory=list)

    @property
    def topic_id(self) -> str:
        return self.metadata.get("id", "").strip() or self.topic_dir.name

    @property
    def title(self) -> str:
        return self.metadata.get("title", "").strip() or self.topic_dir.name

    @property
    def template_docx(self) -> Path | None:
        rel = self.metadata.get("template_docx", "").strip()
        if not rel:
            return None
        path = self.topic_dir / rel
        return path if path.exists() else None


@dataclass
class ApprovedSubtopic:
    source_index: int
    source_text: str
    month: str = ""

    @property
    def title(self) -> str:
        return ensure_trailing_period(clean_markdown_text(self.source_text))


@dataclass
class ApprovedTheme:
    theme_id: str
    heading: str
    month: str = ""
    subtopics: list[ApprovedSubtopic] = field(default_factory=list)
    raw_lines: list[str] = field(default_factory=list)

    @property
    def full_title(self) -> str:
        return clean_markdown_text(self.heading)

    @property
    def short_title(self) -> str:
        match = re.match(r"^Тема\s+\d+\.\s*(.*)$", self.full_title)
        return match.group(1).strip() if match else self.full_title


@dataclass
class OrderTemplateInfo:
    main_theme_line: str
    toc_entries: list[str]
    focus_topic_line: str


@dataclass
class OutlineEntry:
    item_id: str
    title: str

    @property
    def line(self) -> str:
        return f"{self.item_id}. {ensure_trailing_period(self.title)}"

    @property
    def group_id(self) -> str:
        parts = self.item_id.split(".")
        return ".".join(parts[:2]) if len(parts) >= 2 else self.item_id


@dataclass
class MainThemeWorkspace:
    workspace_root: Path
    paths: dict[str, Path]
    theme: ApprovedTheme
    theme_folder: Path
    context_dir: Path
    outline_dir: Path
    orders_dir: Path
    packets_dir: Path
    final_md_dir: Path
    final_docx_dir: Path
    template_text: str
    template_info: OrderTemplateInfo
    master_prompt_text: str
    interaction_guide_text: str
    default_outline_entries: list[OutlineEntry]
    outline_entries: list[OutlineEntry]
    override_path: str | None
    packet_mode: str = DEFAULT_PACKET_MODE
    packet_mode_reason: str = ""


@dataclass
class OrderPart:
    number: int
    heading: str
    content: str

    @property
    def filename(self) -> str:
        return f"part-{self.number:02d}.md"


@dataclass
class SubtopicRunWorkspace:
    theme_workspace: MainThemeWorkspace
    subtopic_entry: OutlineEntry
    run_root: Path
    run_dir: Path
    context_dir: Path
    stage_inputs_dir: Path
    stage_outputs_dir: Path
    final_dir: Path
    web_plan_dir: Path
    order_path: Path
    packet_path: Path
    final_md_target: Path
    final_docx_target: Path
    order_text: str
    packet_text: str
    intro_block: str
    parts: list[OrderPart]
    artifact_profile: str = ARTIFACT_PROFILE_LEAN
    packet_mode: str = DEFAULT_PACKET_MODE
    packet_mode_reason: str = ""
    run_mode: str = RUN_MODE_FRESH
    allowed_parts: tuple[int, ...] = ()
    strict_lean_mode: bool = True
    orchestration_target_percent: float = STRICT_ORCHESTRATION_TARGET_PERCENT

    @property
    def lean_artifacts(self) -> bool:
        return self.artifact_profile != ARTIFACT_PROFILE_FULL

    @property
    def is_surgical_redo(self) -> bool:
        return self.run_mode == RUN_MODE_SURGICAL


def normalize_artifact_profile(value: str | None) -> str:
    candidate = (value or DEFAULT_ARTIFACT_PROFILE).strip().lower()
    if candidate not in {ARTIFACT_PROFILE_LEAN, ARTIFACT_PROFILE_FULL}:
        return ARTIFACT_PROFILE_LEAN
    return candidate


def normalize_packet_mode(value: str | None) -> str:
    candidate = (value or DEFAULT_PACKET_MODE).strip().lower()
    if candidate not in {PACKET_MODE_LEAN, PACKET_MODE_LITERAL}:
        return PACKET_MODE_LEAN
    return candidate


def normalize_packet_mode_reason(value: str | None) -> str:
    return (value or "").strip()


def normalize_run_mode(value: str | None) -> str:
    candidate = (value or RUN_MODE_FRESH).strip().lower()
    if candidate not in {RUN_MODE_FRESH, RUN_MODE_SURGICAL}:
        return RUN_MODE_FRESH
    return candidate


def normalize_allowed_parts(value: Any) -> tuple[int, ...]:
    if value in (None, "", []):
        return ()
    if isinstance(value, str):
        raw_items = [item.strip() for item in value.split(",")]
    else:
        raw_items = [str(item).strip() for item in value]
    normalized: list[int] = []
    for item in raw_items:
        if not item:
            continue
        number = int(item)
        if number < 1 or number > 11:
            raise RuntimeError(f"Allowed part is out of range: {number}")
        if number not in normalized:
            normalized.append(number)
    return tuple(normalized)


def load_existing_run_execution_contract(run_dir: Path) -> dict[str, Any]:
    manifest_path = run_dir / "manifest.json"
    manifest = load_json_if_exists(manifest_path) or {}
    target_percent = manifest.get("orchestration_target_percent", STRICT_ORCHESTRATION_TARGET_PERCENT)
    try:
        normalized_target = float(target_percent)
    except (TypeError, ValueError):
        normalized_target = STRICT_ORCHESTRATION_TARGET_PERCENT
    return {
        "run_mode": normalize_run_mode(manifest.get("run_mode")),
        "allowed_parts": normalize_allowed_parts(manifest.get("allowed_parts")),
        "strict_lean_mode": bool(manifest.get("strict_lean_mode", True)),
        "orchestration_target_percent": normalized_target,
    }


def assert_run_command_allowed(
    run_workspace: SubtopicRunWorkspace,
    command_name: str,
    *,
    part_number: int | None = None,
) -> None:
    if run_workspace.is_surgical_redo:
        if command_name not in SURGICAL_ALLOWED_COMMANDS:
            raise RuntimeError(
                f"Run {run_workspace.run_dir} is locked in surgical-redo mode. "
                f"Command `{command_name}` is forbidden."
            )
        if part_number is not None and run_workspace.allowed_parts and part_number not in run_workspace.allowed_parts:
            raise RuntimeError(
                f"Run {run_workspace.run_dir} is locked in surgical-redo mode for parts "
                f"{', '.join(str(item) for item in run_workspace.allowed_parts)}. "
                f"Part {part_number} is forbidden."
            )


def load_existing_run_artifact_profile(run_dir: Path) -> str:
    manifest_path = run_dir / "manifest.json"
    manifest = load_json_if_exists(manifest_path)
    if not manifest:
        return DEFAULT_ARTIFACT_PROFILE
    return normalize_artifact_profile(manifest.get("artifact_profile"))


def load_existing_run_packet_mode(run_dir: Path) -> str:
    manifest_path = run_dir / "manifest.json"
    manifest = load_json_if_exists(manifest_path)
    if not manifest:
        return DEFAULT_PACKET_MODE
    return normalize_packet_mode(manifest.get("packet_mode"))


def load_existing_run_packet_mode_reason(run_dir: Path) -> str:
    manifest_path = run_dir / "manifest.json"
    manifest = load_json_if_exists(manifest_path)
    if not manifest:
        return ""
    return normalize_packet_mode_reason(manifest.get("packet_mode_reason"))


def parse_main_theme_heading(value: str) -> tuple[str, str]:
    cleaned = clean_markdown_text(value)
    match = re.match(r"^(Тема\s+(?P<id>\d+)\.\s*(?P<title>.+))$", cleaned)
    if not match:
        raise ValueError(f"Unable to parse main theme heading: {value}")
    return match.group("id"), match.group(1)


def parse_subtopic_cell(value: str) -> tuple[int | None, str]:
    cleaned = clean_markdown_text(value)
    match = re.match(r"^(?P<index>\d+)\.\s*(?P<title>.+)$", cleaned)
    if not match:
        return None, cleaned
    return int(match.group("index")), match.group("title").strip()


def parse_approved_themes(text: str) -> list[ApprovedTheme]:
    themes: list[ApprovedTheme] = []
    current: ApprovedTheme | None = None

    for raw_line in text.splitlines():
        cells = parse_markdown_table_row(raw_line)
        if not cells:
            continue
        left = clean_markdown_text(cells[0])
        month = clean_markdown_text(cells[1]) if len(cells) > 1 else ""

        if left.startswith("Тема "):
            theme_id, heading = parse_main_theme_heading(left)
            if current is not None:
                themes.append(current)
            current = ApprovedTheme(
                theme_id=theme_id,
                heading=heading,
                month=month,
                raw_lines=[raw_line],
            )
            continue

        if current is None:
            continue

        source_index, source_text = parse_subtopic_cell(left)
        if source_index is None:
            continue
        current.subtopics.append(
            ApprovedSubtopic(
                source_index=source_index,
                source_text=source_text,
                month=month,
            )
        )
        current.raw_lines.append(raw_line)

    if current is not None:
        themes.append(current)

    return themes


def find_approved_theme(themes: list[ApprovedTheme], query: str) -> ApprovedTheme:
    normalized_query = normalize_search_key(query)
    id_match = re.search(r"(?:^|\s)(\d{1,3})(?:[\.\s]|$)", query)
    if id_match:
        target_id = id_match.group(1)
        exact = [theme for theme in themes if theme.theme_id == target_id]
        if len(exact) == 1:
            return exact[0]

    exact_heading = [theme for theme in themes if normalize_search_key(theme.full_title) == normalized_query]
    if len(exact_heading) == 1:
        return exact_heading[0]

    contains = [
        theme
        for theme in themes
        if normalized_query and normalized_query in normalize_search_key(theme.full_title)
    ]
    if len(contains) == 1:
        return contains[0]

    raise ValueError(f"Unable to resolve approved theme from query: {query}")


def extract_order_template_info(template_text: str) -> OrderTemplateInfo:
    lines = template_text.splitlines()
    start_idx = None
    end_idx = None
    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped == "**III. ОГЛАВЛЕНИЕ**":
            start_idx = index
        elif stripped == "Найди все документы по прилагаемому промпту." and start_idx is not None:
            end_idx = index
            break
    if start_idx is None or end_idx is None:
        raise ValueError("Unable to locate Section III in order template")

    block_lines = [line.strip() for line in lines[start_idx + 1 : end_idx] if line.strip()]
    if not block_lines:
        raise ValueError("Order template Section III is empty")

    main_theme_line = clean_markdown_text(block_lines[0])
    toc_entries = [clean_markdown_text(line) for line in block_lines[1:] if re.match(r"^\d+(?:\.\d+)+\.", clean_markdown_text(line))]
    if not toc_entries:
        raise ValueError("Order template has no TOC entries")

    focus_topic_line = max(toc_entries, key=lambda item: template_text.count(item))
    return OrderTemplateInfo(
        main_theme_line=main_theme_line,
        toc_entries=toc_entries,
        focus_topic_line=focus_topic_line,
    )


def build_generated_outline(theme: ApprovedTheme) -> list[str]:
    outline: list[str] = []
    for subtopic in theme.subtopics:
        outline.append(f"{theme.theme_id}.{subtopic.source_index}. {subtopic.title}")
    return outline


def build_default_outline_entries(theme: ApprovedTheme) -> list[OutlineEntry]:
    return [
        OutlineEntry(
            item_id=f"{theme.theme_id}.{subtopic.source_index}",
            title=subtopic.title.rstrip("."),
        )
        for subtopic in theme.subtopics
    ]


def parse_outline_override_text(theme: ApprovedTheme, text: str) -> list[OutlineEntry]:
    entries: list[OutlineEntry] = []
    seen: set[str] = set()
    for raw_line in text.splitlines():
        cleaned = clean_markdown_text(raw_line)
        match = re.match(rf"^({re.escape(theme.theme_id)}(?:\.\d+)+)\.\s*(.+)$", cleaned)
        if not match:
            continue
        item_id = match.group(1)
        title = match.group(2).strip().rstrip(".")
        if item_id in seen:
            continue
        seen.add(item_id)
        entries.append(OutlineEntry(item_id=item_id, title=title))
    if not entries:
        raise ValueError(f"Outline override for theme {theme.theme_id} has no parsable items")
    return entries


def get_next_subtopic(
    workspace_root: Path,
    theme_number: str,
    current_subtopic_id: str,
) -> OutlineEntry | None:
    override_path = workspace_root / DEFAULT_OUTLINE_OVERRIDES_ROOT / f"{theme_number}.md"
    if not override_path.exists():
        raise FileNotFoundError(
            f"Approved outline file is missing for theme {theme_number}: {override_path}"
        )

    text = read_text(override_path)
    entries: list[OutlineEntry] = []
    seen: set[str] = set()
    pattern = re.compile(rf"^({re.escape(theme_number)}(?:\.\d+)+)\.\s*(.+)$")
    for raw_line in text.splitlines():
        cleaned = clean_markdown_text(raw_line)
        match = pattern.match(cleaned)
        if not match:
            continue
        item_id = match.group(1)
        title = match.group(2).strip().rstrip(".")
        if item_id in seen:
            continue
        seen.add(item_id)
        entries.append(OutlineEntry(item_id=item_id, title=title))

    if not entries:
        raise ValueError(f"Approved outline for theme {theme_number} has no parsable subtopics")

    for index, entry in enumerate(entries):
        if entry.item_id != current_subtopic_id:
            continue
        next_index = index + 1
        if next_index >= len(entries):
            return None
        return entries[next_index]

    raise ValueError(
        f"Current subtopic {current_subtopic_id} is not present in the approved outline for theme {theme_number}"
    )


def render_project_state_next_subtopic_line(
    workspace_root: Path,
    theme_number: str,
    current_subtopic_id: str,
) -> str:
    next_entry = get_next_subtopic(workspace_root, theme_number, current_subtopic_id)
    if next_entry is None:
        return 'Следующая рабочая подтема по approved outline (для Kodeks-агента): "Тема завершена".'
    return (
        "Следующая рабочая подтема по approved outline (для Kodeks-агента): "
        f"`{next_entry.line}`"
    )


def update_project_state_next_subtopic(
    workspace_root: Path,
    theme_number: str,
    current_subtopic_id: str,
    state_path: Path | None = None,
) -> Path:
    resolved_state_path = state_path or (workspace_root / DEFAULT_STATE_PATH)
    state_text = read_text(resolved_state_path)
    next_line = render_project_state_next_subtopic_line(
        workspace_root=workspace_root,
        theme_number=theme_number,
        current_subtopic_id=current_subtopic_id,
    )
    pattern = re.compile(
        r"(?m)^- Следующая рабочая подтема по approved outline \(для Kodeks-агента\): .*$"
    )
    if pattern.search(state_text):
        updated = pattern.sub(f"- {next_line}", state_text, count=1)
    else:
        updated = state_text.rstrip() + f"\n- {next_line}\n"
    if updated != state_text:
        write_text(resolved_state_path, updated)
    return resolved_state_path


def resolve_outline_entries(
    theme: ApprovedTheme,
    theme_folder: Path,
    workspace_root: Path,
) -> tuple[list[OutlineEntry], str | None]:
    override_candidates = [
        workspace_root / DEFAULT_OUTLINE_OVERRIDES_ROOT / f"{theme.theme_id}.md",
        theme_folder / "01-outline" / "outline.approved.md",
    ]
    for candidate in override_candidates:
        if candidate.exists():
            return parse_outline_override_text(theme, read_text(candidate)), str(candidate)
    return build_default_outline_entries(theme), None


def render_outline_markdown(theme: ApprovedTheme, outline_entries: list[OutlineEntry]) -> str:
    lines = [
        f"# {theme.full_title}",
        "",
        "Автоматически сформированное Оглавление по текущей основной теме.",
        "",
        f"**{theme.full_title}**",
        "",
    ]
    for item in outline_entries:
        lines.append(item.line)
        lines.append("")

    lines.extend(
        [
            "## Исходные позиции из Утверждаю",
            "",
        ]
    )
    for subtopic in theme.subtopics:
        lines.append(f"{subtopic.source_index}. {subtopic.title}")
    lines.append("")
    return "\n".join(lines)


def render_theme_block_markdown(theme: ApprovedTheme) -> str:
    lines = ["# Блок темы из Утверждаю", ""]
    lines.extend(theme.raw_lines)
    lines.append("")
    return "\n".join(lines)


def render_order_copy(
    template_text: str,
    template_info: OrderTemplateInfo,
    theme: ApprovedTheme,
    outline_entries: list[OutlineEntry],
    focus_topic_line: str,
) -> str:
    outline_block_lines = [f"**{theme.full_title}**", ""]
    for item in outline_entries:
        outline_block_lines.append(item.line)
        outline_block_lines.append("")
    outline_block = "\n".join(outline_block_lines).rstrip()

    pattern = re.compile(
        r"(\*\*III\. ОГЛАВЛЕНИЕ\*\*\s*)(.*?)(\nНайди все документы по прилагаемому промпту\.)",
        re.DOTALL,
    )
    match = pattern.search(template_text)
    if not match:
        raise ValueError("Unable to rewrite Section III in order template")

    rewritten = pattern.sub(rf"\1{outline_block}\3", template_text, count=1)

    old_line = ensure_trailing_period(template_info.focus_topic_line)
    new_line = ensure_trailing_period(focus_topic_line)
    replacements = {
        old_line: new_line,
        old_line.rstrip("."): new_line.rstrip("."),
        template_info.main_theme_line: theme.full_title,
    }
    for old_value, new_value in replacements.items():
        rewritten = rewritten.replace(old_value, new_value)
    return rewritten


def build_execution_packet(
    theme: ApprovedTheme,
    subtopic_line: str,
    master_prompt_text: str,
    generated_order_text: str,
    interaction_guide_text: str,
    packet_mode: str = DEFAULT_PACKET_MODE,
    packet_mode_reason: str = "",
) -> str:
    def guard_no_canonical_inline(payload: str) -> None:
        canonical_map = {
            "00-context/03-master-prompt.md": master_prompt_text.strip(),
            "00-context/06-order.copy.md": generated_order_text.strip(),
            "00-context/08-interaction-guide.md": interaction_guide_text.strip(),
        }
        for path, text in canonical_map.items():
            if text and text in payload:
                raise ValueError(f"Canonical file must not be inlined: {path}")

    packet_mode = normalize_packet_mode(packet_mode)
    packet_mode_reason = normalize_packet_mode_reason(packet_mode_reason)
    if packet_mode == PACKET_MODE_LITERAL:
        if not packet_mode_reason:
            raise RuntimeError(
                "literal-safe packet mode requires an explicit reason because full canonical texts are already available as run files."
            )
        lines = [
            f"# Execution Packet: {subtopic_line}",
            "",
            f"> Packet mode: `{PACKET_MODE_LITERAL}`",
            f"> Reason: {packet_mode_reason}",
            "",
            "## Основная тема",
            "",
            theme.full_title,
            "",
            "## Обязательный порядок",
            "",
            "1. Прочитать AGENTS.md.",
            "2. Прочитать PROJECT_STATE.md.",
            "3. Прочитать manual-workflow.md.",
            "4. Прочитать мастер-промпт.",
            "5. Прочитать выбранную тему из Утверждаю и Оглавление.",
            "6. Прочитать сгенерированную копию приказа по текущей подтеме.",
            "7. Использовать guide по механике взаимодействия в окне LLM.",
            "",
            "## Ссылки на канон",
            "",
            "[см. 00-context/03-master-prompt.md]",
            "[см. 00-context/06-order.copy.md]",
            "[см. 00-context/08-interaction-guide.md]",
            "",
        ]
        return "\n".join(lines)
    lines = [
        f"# Execution Packet: {subtopic_line}",
        "",
        "## Основная тема",
        "",
        theme.full_title,
        "",
        "## Текущая подтема",
        "",
        subtopic_line,
        "",
        "## Обязательный порядок",
        "",
        "1. Прочитать AGENTS.md.",
        "2. Прочитать PROJECT_STATE.md.",
        "3. Прочитать manual-workflow.md.",
        "4. Прочитать мастер-промпт.",
        "5. Прочитать выбранную тему из Утверждаю и Оглавление.",
        "6. Прочитать сгенерированную копию приказа по текущей подтеме.",
        "7. Использовать guide по механике взаимодействия в окне LLM.",
        "",
        "## Ссылки на канон",
        "",
        "[см. 00-context/00-agents.md]",
        "[см. 00-context/01-project-state.md]",
        "[см. 00-context/02-manual-workflow.md]",
        "[см. 00-context/03-master-prompt.md]",
        "[см. 00-context/06-order.copy.md]",
        "[см. 00-context/08-interaction-guide.md]",
        "",
    ]
    packet_text = "\n".join(lines)
    guard_no_canonical_inline(packet_text)
    return packet_text


def extract_execution_packet_index(text: str, max_items: int = 20) -> list[str]:
    items: list[str] = []
    seen: set[str] = set()
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        normalized = line.strip("*").strip()
        if not normalized:
            continue
        is_heading = False
        if line.startswith("#"):
            is_heading = True
        elif line.startswith("**") and line.endswith("**"):
            is_heading = True
        elif re.match(r"^(ЧАСТЬ\s+\d+|[IVXLCDM]+\.)", normalized):
            is_heading = True
        if not is_heading:
            continue
        normalized = re.sub(r"\s+", " ", normalized)
        if len(normalized) > 220:
            normalized = normalized[:217].rstrip() + "..."
        if normalized in seen:
            continue
        seen.add(normalized)
        items.append(f"- {normalized}")
        if len(items) >= max_items:
            break
    return items


def extract_execution_packet_excerpt(
    text: str,
    max_nonempty_lines: int = 40,
    max_chars: int = 8000,
) -> str:
    selected: list[str] = []
    total_chars = 0
    for raw_line in text.splitlines():
        line = raw_line.rstrip()
        if not line.strip():
            if selected and selected[-1] != "":
                selected.append("")
            continue
        selected.append(line)
        total_chars += len(line) + 1
        nonempty_lines = sum(1 for item in selected if item.strip())
        if nonempty_lines >= max_nonempty_lines or total_chars >= max_chars:
            break
    while selected and not selected[-1].strip():
        selected.pop()
    excerpt = "\n".join(selected).strip()
    if excerpt and len(excerpt) < len(text.strip()):
        excerpt = excerpt.rstrip() + "\n\n[... сокращено в execution packet; полный guide не дублируется здесь ...]"
    return excerpt


def find_outline_entry(outline_entries: list[OutlineEntry], subtopic_id: str) -> OutlineEntry:
    for entry in outline_entries:
        if entry.item_id == subtopic_id:
            return entry
    raise ValueError(f"Subtopic {subtopic_id} is not present in the approved outline")


def render_generated_orders(workspace: MainThemeWorkspace) -> list[dict[str, str]]:
    generated_orders: list[dict[str, str]] = []
    for entry in workspace.outline_entries:
        subtopic_id = entry.item_id
        subtopic_line = entry.line
        order_text = render_order_copy(
            template_text=workspace.template_text,
            template_info=workspace.template_info,
            theme=workspace.theme,
            outline_entries=workspace.outline_entries,
            focus_topic_line=subtopic_line,
        )
        order_name = f"{subtopic_id}.md"
        order_path = workspace.orders_dir / order_name
        write_text(order_path, order_text)

        packet_text = build_execution_packet(
            theme=workspace.theme,
            subtopic_line=subtopic_line,
            master_prompt_text=workspace.master_prompt_text,
            generated_order_text=order_text,
            interaction_guide_text=workspace.interaction_guide_text,
            packet_mode=workspace.packet_mode,
            packet_mode_reason=workspace.packet_mode_reason,
        )
        packet_path = workspace.packets_dir / order_name
        write_text(packet_path, packet_text)

        generated_orders.append(
            {
                "subtopic_id": subtopic_id,
                "subtopic_line": subtopic_line,
                "order_file": str(order_path),
                "packet_file": str(packet_path),
                "final_md_target": str(workspace.final_md_dir / f"{subtopic_id}.md"),
                "final_docx_target": str(workspace.final_docx_dir / f"{subtopic_id}.docx"),
            }
        )
    return generated_orders


def split_rendered_order(order_text: str) -> tuple[str, list[OrderPart]]:
    matches = list(
        re.finditer(
            r"^\*\*ЧАСТЬ\s+(?P<number>\d+)\s+ПРИКАЗА НА ВЫПОЛНЕНИЕ МАСТЕР-ПРОМПТА\*\*$",
            order_text,
            flags=re.MULTILINE,
        )
    )
    if not matches:
        raise ValueError("Rendered order has no staged parts")

    intro_block = order_text[: matches[0].start()].strip()
    parts: list[OrderPart] = []
    for index, match in enumerate(matches):
        start = match.start()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(order_text)
        block = order_text[start:end].strip()
        part_number = int(match.group("number"))
        parts.append(
            OrderPart(
                number=part_number,
                heading=clean_markdown_text(match.group(0)),
                content=block,
            )
        )

    expected_numbers = list(range(1, 12))
    actual_numbers = [part.number for part in parts]
    if actual_numbers != expected_numbers:
        raise ValueError(
            f"Rendered order must contain Parts 1-11 in sequence, got: {actual_numbers}"
        )

    return intro_block, parts


def build_part_output_stub(subtopic_line: str, part: OrderPart, input_path: Path) -> str:
    mode = PART_EXECUTION_MODES.get(part.number, "custom")
    status = "ready" if part.number == 1 else "blocked_until_go"
    lines = [
        f"# Response Stub: Part {part.number:02d}",
        "",
        f"Подтема: {subtopic_line}",
        f"Режим: {mode}",
        f"Статус: {status}",
        f"Источник: {input_path}",
        "",
        "Сюда сохраняется ответ по соответствующей Части приказа.",
        "",
    ]
    return "\n".join(lines)


def get_related_outline_entries(run_workspace: SubtopicRunWorkspace) -> list[OutlineEntry]:
    current_id = run_workspace.subtopic_entry.item_id
    group_id = run_workspace.subtopic_entry.group_id
    return [
        entry
        for entry in run_workspace.theme_workspace.outline_entries
        if entry.group_id == group_id and entry.item_id != current_id
    ]


def classify_outline_entry(entry: OutlineEntry) -> str:
    return "tariff" if is_tariff_title(entry.title) else "substantive"


def build_part_02_focus(run_workspace: SubtopicRunWorkspace) -> dict[str, Any]:
    entry = run_workspace.subtopic_entry
    title = trim_terminal_period(entry.title)
    tariff = is_tariff_title(title)
    action_phrase = extract_tariff_action_phrase(title) if tariff else title
    related_entries = get_related_outline_entries(run_workspace)
    related_payload = [
        {
            "item_id": related.item_id,
            "title": trim_terminal_period(related.title),
            "type": classify_outline_entry(related),
        }
        for related in related_entries
    ]
    keyword_source = action_phrase if tariff else title
    keywords = infer_query_keywords(keyword_source)
    keyword_phrase = " ".join(keywords[:6]) if keywords else keyword_source
    substantive_sibling = next(
        (item for item in related_payload if item["type"] == "substantive"),
        None,
    )
    tariff_sibling = next(
        (item for item in related_payload if item["type"] == "tariff"),
        None,
    )
    execution_note = (
        "Тарифный подпункт должен анализироваться только в привязке к базовому нотариальному действию."
        if tariff
        else "Самостоятельный тарифный разбор не смешивается с базовой подтемой, если для группы есть отдельный тарифный подпункт."
    )
    return {
        "subtopic_id": entry.item_id,
        "subtopic_line": entry.line,
        "title": title,
        "subtopic_type": "tariff" if tariff else "substantive",
        "query_focus": action_phrase,
        "keyword_phrase": keyword_phrase,
        "keywords": keywords,
        "related_entries": related_payload,
        "substantive_sibling": substantive_sibling,
        "tariff_sibling": tariff_sibling,
        "execution_note": execution_note,
    }


def dedupe_queries(queries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: set[str] = set()
    unique: list[dict[str, Any]] = []
    for query in queries:
        key = query["query"].strip()
        if key in seen:
            continue
        seen.add(key)
        unique.append(query)
    return unique


def build_part_02_queries(run_workspace: SubtopicRunWorkspace) -> list[dict[str, Any]]:
    focus = build_part_02_focus(run_workspace)
    title = focus["title"]
    query_focus = focus["query_focus"]
    keyword_phrase = focus["keyword_phrase"]
    queries: list[dict[str, Any]] = []

    def add_query(
        query_id: str,
        purpose: str,
        query: str,
        preferred_domains: list[str],
        notes: str,
    ) -> None:
        queries.append(
            {
                "id": query_id,
                "purpose": purpose,
                "query": query,
                "preferred_domains": preferred_domains,
                "notes": notes,
            }
        )

    add_query(
        "q01",
        "exact_phrase",
        f"\"{title}\" нотариус",
        ["publication.pravo.gov.ru", "pravo.gov.ru", "consultant.ru", "garant.ru"],
        "Стартовый точный поиск по полной формулировке подтемы.",
    )
    add_query(
        "q02",
        "official_anchor",
        f"\"{title}\" site:publication.pravo.gov.ru OR site:pravo.gov.ru",
        ["publication.pravo.gov.ru", "pravo.gov.ru"],
        "Искать официальный якорь URL1.",
    )
    add_query(
        "q03",
        "readable_url2",
        f"\"{title}\" site:consultant.ru OR site:garant.ru OR site:docs.cntd.ru",
        ["consultant.ru", "garant.ru", "docs.cntd.ru"],
        "Подбирать читаемый VERIFIED URL2 для усиления карточки только после сверки реквизитов.",
    )
    add_query(
        "q04",
        "notariat_basics",
        f"\"{query_focus}\" \"Основы законодательства Российской Федерации о нотариате\"",
        ["pravo.gov.ru", "consultant.ru", "garant.ru"],
        "Проверить профильный базовый слой нотариального регулирования.",
    )
    add_query(
        "q05",
        "minjust_procedure",
        f"\"{query_focus}\" нотариат Минюст порядок",
        ["minjust.gov.ru", "publication.pravo.gov.ru", "pravo.gov.ru"],
        "Процедурный и подзаконный слой Минюста.",
    )
    add_query(
        "q06",
        "fnp_materials",
        f"\"{query_focus}\" site:notariat.ru",
        ["notariat.ru"],
        "Методический слой и документы нотариального сообщества.",
    )
    add_query(
        "q07",
        "court_clarifications",
        f"\"{query_focus}\" нотариус обзор практика пленум",
        ["vsrf.ru", "consultant.ru", "garant.ru"],
        "Проверить применимые разъяснения и обзоры судебной практики.",
    )

    if focus["subtopic_type"] == "tariff":
        add_query(
            "q08",
            "tariff_direct",
            f"\"{query_focus}\" \"федеральный тариф\" \"региональный тариф\" нотариус",
            ["pravo.gov.ru", "notariat.ru", "consultant.ru", "garant.ru"],
            "Основной поисковый блок по федеральному и региональному тарифу.",
        )
        add_query(
            "q09",
            "tariff_article_22_1",
            f"\"{query_focus}\" \"22.1\" нотариат",
            ["pravo.gov.ru", "consultant.ru", "garant.ru"],
            "Проверить профильные нормы Основ законодательства о нотариате и смежных актов.",
        )
        add_query(
            "q10",
            "tariff_state_duty",
            f"\"{query_focus}\" \"государственная пошлина\" нотариус",
            ["nalog.gov.ru", "pravo.gov.ru", "consultant.ru", "garant.ru"],
            "Налогово-финансовый слой, если он влияет на расчет.",
        )
        add_query(
            "q11",
            "tariff_regional_layer",
            f"\"{query_focus}\" \"предельные размеры регионального тарифа\" нотариус",
            ["notariat.ru", "consultant.ru", "garant.ru"],
            "Проверить региональный тариф и решения органов нотариального сообщества.",
        )
    else:
        add_query(
            "q08",
            "procedure_variants",
            f"\"{query_focus}\" порядок совершения нотариального действия",
            ["pravo.gov.ru", "minjust.gov.ru", "consultant.ru", "garant.ru"],
            "Уточнить процедуру, участников, сроки и ограничения по действию.",
        )
        add_query(
            "q09",
            "forms_registers",
            f"\"{query_focus}\" форма реестр нотариус",
            ["minjust.gov.ru", "notariat.ru", "consultant.ru", "garant.ru"],
            "Проверить формы, реестровый, делопроизводственный и учетный слой.",
        )
        add_query(
            "q10",
            "amendment_chain",
            f"\"{query_focus}\" приказ Минюста изменения 2024 2025 156 226 224 225",
            ["publication.pravo.gov.ru", "rg.ru", "consultant.ru", "garant.ru"],
            "Добрать редакционную цепочку и свежие приказы об изменениях к базовым приказам Минюста.",
        )
        add_query(
            "q11",
            "latest_official_updates",
            f"\"{query_focus}\" актуальная редакция приказ Минюста нотариат 2025",
            ["publication.pravo.gov.ru", "pravo.gov.ru", "rg.ru", "consultant.ru", "garant.ru"],
            "Проверить свежие официально опубликованные обновления и редакции по теме.",
        )

    return dedupe_queries(queries)


def render_part_02_queries_markdown(run_workspace: SubtopicRunWorkspace, queries: list[dict[str, Any]]) -> str:
    lines = [
        f"# Part 02 Query Plan: {run_workspace.subtopic_entry.line}",
        "",
        "Ниже — стартовый web-first набор запросов. Это не финальный перечень документов, а план поиска и верификации.",
        "",
    ]
    for query in queries:
        domains = ", ".join(query["preferred_domains"])
        lines.extend(
            [
                f"## {query['id']} — {query['purpose']}",
                "",
                "```text",
                query["query"],
                "```",
                "",
                f"Предпочтительные домены: {domains}",
                f"Зачем: {query['notes']}",
                "",
            ]
        )
    return "\n".join(lines)


def render_part_02_source_cascade_markdown(source_cascade: list[dict[str, Any]]) -> str:
    lines = [
        "# Part 02 Source Cascade",
        "",
        "Каскад фиксирует приоритет слоев источников для URL1/URL2 и верификации.",
        "",
    ]
    for item in source_cascade:
        lines.extend(
            [
                f"## {item['rank']}. {item['role']}",
                "",
                f"Домены: {', '.join(item['domains'])}",
                f"Применение: {item['usage']}",
                "",
            ]
        )
    return "\n".join(lines)


def build_part_02_research_pack(run_workspace: SubtopicRunWorkspace) -> str:
    focus = build_part_02_focus(run_workspace)
    queries = build_part_02_queries(run_workspace)
    lines = [
        f"# Part 02 Research Pack: {focus['subtopic_line']}",
        "",
        "## Фокус подтемы",
        "",
        f"- Тип подтемы: `{focus['subtopic_type']}`",
        f"- Базовая формулировка: `{focus['title']}`",
        f"- Фраза для правового ядра поиска: `{focus['query_focus']}`",
        f"- Ключевые слова: `{', '.join(focus['keywords'])}`" if focus["keywords"] else "- Ключевые слова: нет выделенных ключевых слов",
        f"- Правило исполнения: {focus['execution_note']}",
        "",
    ]
    if focus["substantive_sibling"]:
        sibling = focus["substantive_sibling"]
        lines.append(
            f"- Связанная базовая подтема: `{sibling['item_id']}. {ensure_trailing_period(sibling['title'])}`"
        )
    if focus["tariff_sibling"]:
        sibling = focus["tariff_sibling"]
        lines.append(
            f"- Связанная тарифная подтема: `{sibling['item_id']}. {ensure_trailing_period(sibling['title'])}`"
        )
    if focus["related_entries"]:
        lines.extend(["", "## Смежные подпункты группы", ""])
        for item in focus["related_entries"]:
            lines.append(
                f"- `{item['item_id']}. {ensure_trailing_period(item['title'])}` ({item['type']})"
            )
    lines.extend(
        [
            "",
            "## Жесткие ограничения Part 02",
            "",
            "- Part 02 должен сразу давать финально-годное ядро документа.",
            "- Если документ назван в `АНАЛИЗ-КОДЕКСЫ И БАЗОВЫЕ АКТЫ`, `A`, `B`, `КАРАНТИН` или `FAIL-SAFE CHECK`, у него должны быть полное официальное наименование, структурный элемент, URL1 и статус URL2.",
            "- Сначала анализ области права и кодексов, только потом A/B/КАРАНТИН/FAIL-SAFE CHECK.",
            "- Публичная обработка блоков I-XXXVII, статусы `НАЙДЕНО/НЕ ВЫЯВЛЕНО` и любые диапазоны блоков в Part 02 запрещены: это относится к Части 3.",
            "- Внутренняя маркировка применимости блоков допустима только как внутренняя работа исполнителя и не должна появляться в публичном тексте Part 02.",
            "- Для тарифных подтем нельзя терять привязку к конкретному нотариальному действию.",
            "",
            "## Стартовый порядок поиска",
            "",
            "1. Найти URL1-источник на официальном правовом портале или ином официальном ресурсе.",
            "2. Сверить точные реквизиты и заголовок страницы.",
            "3. Определить существование документа, его роль в теме и применимость к конкретной Части.",
            "4. Найти читаемый URL2-кандидат и подтвердить совпадение реквизитов, если это удается довести до конца.",
            "5. Если документ подтвержден по существованию, роли, применимости и URL1, включать его в ядро или опорный слой даже при слабом URL2; карантин использовать только при реальной неопределенности по самому документу, а не по одной читабельной ссылке.",
            "",
            "## Стартовые запросы",
            "",
        ]
    )
    for query in queries:
        lines.append(f"- `{query['id']}`: {query['query']}")
    lines.append("")
    return "\n".join(lines)


def build_part_02_core_template(run_workspace: SubtopicRunWorkspace) -> str:
    focus = build_part_02_focus(run_workspace)
    lines = [
        f"ТЕМА: {focus['subtopic_line']}",
        "",
        "АНАЛИЗ ОБЛАСТИ ПРАВА",
        "[Сначала квалифицировать юридическую природу нотариального действия, участников, объект удостоверения/проверки, процедурный, подзаконный и контрольный слои.]",
        "",
        "АНАЛИЗ-КОДЕКСЫ И БАЗОВЫЕ АКТЫ",
        "[Каждый названный здесь акт должен идти как отдельный нумерованный мини-блок: полное официальное наименование, структурный элемент, URL1, URL2, VERIFIED URL2, Заголовок страницы URL2, Сверка реквизитов.]",
        "",
        "A. РЕГУЛЯТОРНОЕ ЯДРО",
        "[Сюда включаются только подтвержденные прямо применимые документы с полноценными карточками: Вид документа, Полное наименование, Орган, Структурный элемент, Значение для темы, URL1, URL2, VERIFIED URL2, Заголовок страницы URL2, Сверка реквизитов, Каноническая строка поиска.]",
        "",
        "B. ОПОРНЫЕ ДОКУМЕНТЫ",
        "[Сюда включаются дополнительные профильные документы в том же карточном формате.]",
        "",
        "КАРАНТИН",
        "[Сюда попадают только документы с материальной неопределенностью по самому документу, его действию, применимости или структурному элементу; даже здесь обязательны полное наименование, структурный элемент, URL1, статус URL2 и причина карантина.]",
        "",
        "FAIL-SAFE CHECK",
        "[Отдельно перечислить слепые зоны; если называются конкретные акты, у них тоже должны быть наименование, структурный элемент и ссылки.]",
        "",
        "## Внутренние заметки для исполнителя",
        "",
        f"- Тип подтемы: `{focus['subtopic_type']}`",
        f"- Фокус: `{focus['query_focus']}`",
        f"- Правило: {focus['execution_note']}",
        "- Не публиковать в Part 02 блоки I-XXXVII, статусы `НАЙДЕНО/НЕ ВЫЯВЛЕНО` и блоковые диапазоны: это отдельная стадия Части 3.",
    ]
    if focus["tariff_sibling"]:
        lines.append(
            f"- Не смешивать самостоятельный тарифный разбор с базовой подтемой: тариф вынесен в `{focus['tariff_sibling']['item_id']}`."
        )
    if focus["substantive_sibling"]:
        lines.append(
            f"- Тарифный блок должен быть привязан к базовому действию из `{focus['substantive_sibling']['item_id']}`."
        )
    lines.append("")
    return "\n".join(lines)


def build_part_02_web_readme(run_workspace: SubtopicRunWorkspace) -> str:
    return "\n".join(
        [
            f"# Part 02 Web Plan: {run_workspace.subtopic_entry.line}",
            "",
            "Этот каталог готовит web-first исполнение Части 2.",
            "",
            f"- Research pack: `{run_workspace.web_plan_dir / 'part-02.research-pack.md'}`",
            f"- Query plan markdown: `{run_workspace.web_plan_dir / 'queries.md'}`",
            f"- Query plan json: `{run_workspace.web_plan_dir / 'queries.json'}`",
            f"- Source cascade markdown: `{run_workspace.web_plan_dir / 'source-cascade.md'}`",
            f"- Part 02 core template: `{run_workspace.web_plan_dir / 'part-02.core-template.md'}`",
            f"- Part 02 launch packet: `{run_workspace.web_plan_dir / 'part-02.launch-packet.md'}`",
            f"- Research log: `{run_workspace.web_plan_dir / 'research-log.jsonl'}`",
            f"- Evidence dir: `{run_workspace.web_plan_dir / 'evidence'}`",
            "",
            "Назначение: не дать исполнителю каждый раз заново придумывать поисковые строки, порядок доменов и каркас финального ядра ответа.",
            "",
        ]
    )


def build_part_02_web_plan_paths(run_workspace: SubtopicRunWorkspace) -> dict[str, Path]:
    return {
        "web_plan_dir": run_workspace.web_plan_dir,
        "readme": run_workspace.web_plan_dir / "README.web.md",
        "operator_sequence": run_workspace.web_plan_dir / "00-operator-sequence.md",
        "focus_json": run_workspace.web_plan_dir / "focus.json",
        "queries_md": run_workspace.web_plan_dir / "queries.md",
        "queries_json": run_workspace.web_plan_dir / "queries.json",
        "source_cascade_md": run_workspace.web_plan_dir / "source-cascade.md",
        "source_cascade_json": run_workspace.web_plan_dir / "source-cascade.json",
        "research_pack": run_workspace.web_plan_dir / "part-02.research-pack.md",
        "core_template": run_workspace.web_plan_dir / "part-02.core-template.md",
        "launch_packet": run_workspace.web_plan_dir / "part-02.launch-packet.md",
        "message_01": run_workspace.web_plan_dir / "message-01.part-01.md",
        "message_02": run_workspace.web_plan_dir / "message-02.go.txt",
        "message_03": run_workspace.web_plan_dir / "message-03.part-02-launch-packet.md",
        "research_log": run_workspace.web_plan_dir / "research-log.jsonl",
        "evidence_dir": run_workspace.web_plan_dir / "evidence",
    }


def build_llm_operator_sequence(run_workspace: SubtopicRunWorkspace) -> str:
    paths = build_part_02_web_plan_paths(run_workspace)
    part_01_output = run_workspace.stage_outputs_dir / "part-01.md"
    lines = [
        f"# Operator Sequence: {run_workspace.subtopic_entry.line}",
        "",
        "Это операторская последовательность для внешней LLM-сессии. Она нужна, пока агент еще не отправляет сообщения в LLM сам.",
        "",
        "## Что отправлять по порядку",
        "",
        f"1. Сообщение 1: содержимое файла `{paths['message_01']}`",
        f"2. Сообщение 2: содержимое файла `{paths['message_02']}`",
        f"3. Сообщение 3: содержимое файла `{paths['message_03']}` — только если после `GO/СТАРТ` LLM не начал сам выдавать Part 2.",
        "",
        "## Что сделать после ответа LLM по Части 2",
        "",
        "Скопировать ответ целиком в clipboard и выполнить команду:",
        "",
        "```powershell",
        f"python .\\notary_agent.py capture-part-output {run_workspace.subtopic_entry.item_id} 2 --clipboard",
        "```",
        "",
        "## Проверка перед стартом",
        "",
        f"- Файл Part 1 должен быть не stub: `{part_01_output}`",
        f"- Launch packet уже собран: `{paths['launch_packet']}`",
        f"- Актуальный latest run: `{run_workspace.run_dir}`",
        "",
        "## Что не делать",
        "",
        "- Не склеивать все три сообщения в одно.",
        "- Не пропускать сообщение `GO/СТАРТ` между Частью 1 и Частью 2.",
        "- Если ответ на `GO/СТАРТ` уже начинается с `ТЕМА:` и разворачивает Part 2, не отправлять fallback `message-03.part-02-launch-packet.md` в эту же сессию.",
        "- Не вставлять в LLM сырой `part-02.md` из stage-inputs вместо launch packet.",
        "",
    ]
    return "\n".join(lines)


def build_part_02_launch_packet(run_workspace: SubtopicRunWorkspace) -> str:
    focus = build_part_02_focus(run_workspace)
    part_02_input = next(part.content.strip() for part in run_workspace.parts if part.number == 2)
    literal_context_md = build_literal_context_bundle(run_workspace, 2)
    reasoning_brief_md = build_reasoning_part_brief(run_workspace, 2)
    lines = [
        f"# Part 02 Launch Packet: {focus['subtopic_line']}",
        "",
        "Этот файл должен максимально буквально продолжать ручной сильный сценарий по Части 2, а не подменять его пересказом.",
        "",
        "## Режим",
        "",
        "- Не пересказывать мастер-промпт и не делать мета-анализ вместо исполнения.",
        "- Не сужать охват только до заранее подготовленных поисковых строк: они могут помогать, но не ограничивают поиск.",
        "- Считать этот пакет прямым продолжением той же подтемы после Части 1 и сигнала `GO/СТАРТ`.",
        "",
        "## Что должен вернуть LLM",
        "",
        f"- Первая строка: `ТЕМА: {focus['subtopic_line']}`",
        "- Вернуть только финально-годный ответ по Части 2.",
        "- Сохранить жесткий порядок блоков и не добавлять лишние разделы раньше времени.",
        "- В пользовательском ответе оформлять ссылки и верификацию по эталонному режиму: `URL1`, `URL2`, `VERIFIED URL2`, `Заголовок страницы URL2 (как на странице)`, `Сверка реквизитов`, `Каноническая строка поиска`.",
        "- Каждый акт, названный в `АНАЛИЗ-КОДЕКСЫ И БАЗОВЫЕ АКТЫ`, должен сопровождаться отдельным мини-блоком с полным официальным наименованием, структурным элементом и ссылками.",
        "- Если документ подтвержден по существованию, применимости, роли и URL1, он не должен теряться из-за слабого URL2: оставлять его в `A` или `B` по роли и явно помечать статус `VERIFIED URL2` и сверку реквизитов.",
        "- В `КАРАНТИН` переносить документ только при реальном сомнении в существовании, действии, точной идентификации, применимости или структурном элементе.",
        "- Режим `URL2: отсутствует (КАРАНТИН)` относится к полю `URL2` внутри карточки и сам по себе не означает перенос документа в раздел `КАРАНТИН`.",
        "- В `Части 4` и других follow-up частях запрещен формат `Прямой документ(ы): ...` со сваленными ссылками: каждый найденный документ должен иметь собственную карточку.",
        "- Не публиковать в Part 2 блоки I-XXXVII, статусы `НАЙДЕНО/НЕ ВЫЯВЛЕНО`, диапазоны блоков и иные результаты обработки СФЕРЫ ПОИСКА: это отдельная Часть 3.",
        "- Внутренняя маркировка применимости блоков допустима только внутри размышления исполнителя и не должна попадать в финальный текст ответа по Части 2.",
        "",
        "Обязательные маркеры Part 2:",
        "",
    ]
    for marker in PART_02_REQUIRED_MARKERS:
        lines.append(f"- `{marker}`")
    lines.extend(["", reasoning_brief_md, ""])
    if literal_context_md:
        lines.extend(["", literal_context_md, ""])
    lines.extend(
        [
            "## Точная инструкция Части 2",
            "",
            part_02_input,
            "",
            "## Внутренние guardrails этой подтемы",
            "",
            f"- Тип подтемы: `{focus['subtopic_type']}`",
            f"- Фокус: `{focus['query_focus']}`",
            f"- Правило: {focus['execution_note']}",
        ]
    )
    if focus["tariff_sibling"]:
        lines.append(
            f"- Не смешивать самостоятельный тарифный разбор с базовой подтемой: тариф вынесен в `{focus['tariff_sibling']['item_id']}`."
        )
    if focus["substantive_sibling"]:
        lines.append(
            f"- Тарифный блок должен быть привязан к базовому действию из `{focus['substantive_sibling']['item_id']}`."
        )
    lines.extend(
        [
            "",
            "## Требование к ответу",
            "",
            "Верни только готовый ответ по Части 2 без пояснений вне структуры ответа.",
            "",
            "## После получения ответа",
            "",
            "Сохранить ответ можно командой:",
            "",
            "```powershell",
            f"python .\\notary_agent.py capture-part-output {run_workspace.subtopic_entry.item_id} 2 --clipboard",
            "```",
            "",
        ]
    )
    return "\n".join(lines)


def get_order_part(run_workspace: SubtopicRunWorkspace, part_number: int) -> OrderPart:
    for part in run_workspace.parts:
        if part.number == part_number:
            return part
    raise RuntimeError(f"Unknown part number: {part_number}")


def get_part_output_path(run_workspace: SubtopicRunWorkspace, part_number: int) -> Path:
    return run_workspace.stage_outputs_dir / f"part-{part_number:02d}.md"


def load_completed_part_output(run_workspace: SubtopicRunWorkspace, part_number: int) -> str | None:
    output_path = get_part_output_path(run_workspace, part_number)
    if not output_path.exists():
        return None
    content = read_text(output_path).strip()
    if not content or is_response_stub(content):
        return None
    return content


def collect_literal_context_parts(
    run_workspace: SubtopicRunWorkspace,
    target_part_number: int,
) -> list[tuple[int, str]]:
    collected: list[tuple[int, str]] = []
    for part_number in range(2, target_part_number):
        content = load_completed_part_output(run_workspace, part_number)
        if content:
            name_lines = [
                line.strip()
                for line in content.splitlines()
                if "Полное наименование:" in line
            ]
            if name_lines:
                collected.append((part_number, "\n".join(name_lines)))
    return collected


def build_literal_context_bundle(
    run_workspace: SubtopicRunWorkspace,
    target_part_number: int,
) -> str:
    context_parts = collect_literal_context_parts(run_workspace, target_part_number)
    if not context_parts:
        return ""

    lines = [
        "## Живой контекст этой же подтемы",
        "",
        "Продолжать как следующий шаг той же работы. Уже подтвержденные части ниже не перезапускать и не переписывать заново без прямой необходимости.",
        "",
    ]
    for part_number, content in context_parts:
        title = FINAL_PART_TITLES.get(part_number, f"ЧАСТЬ {part_number}")
        lines.extend(
            [
                f"### ЧАСТЬ {part_number}. {title}",
                "",
                content,
                "",
            ]
        )
    return "\n".join(lines).rstrip()


def count_non_overlapping(text: str, needle: str) -> int:
    if not needle:
        return 0
    return text.count(needle)


def extract_first_substantive_paragraphs(
    content: str,
    max_paragraphs: int = 2,
    max_chars: int = 1800,
) -> str:
    paragraphs: list[str] = []
    for raw_paragraph in re.split(r"\n\s*\n", content):
        paragraph = raw_paragraph.strip()
        if not paragraph:
            continue
        if paragraph.startswith("```"):
            continue
        if paragraph.startswith("#"):
            continue
        if paragraph.startswith("ТЕМА:"):
            continue
        if paragraph.startswith("## "):
            continue
        paragraphs.append(re.sub(r"\s+", " ", paragraph))
        if len(paragraphs) >= max_paragraphs:
            break
    summary = "\n\n".join(paragraphs).strip()
    if len(summary) > max_chars:
        summary = summary[: max_chars - 3].rstrip() + "..."
    return summary


def extract_anchor_lines_for_followup(
    content: str,
    max_items: int = 8,
) -> list[str]:
    anchors: list[str] = []
    seen: set[str] = set()
    lines = content.splitlines()
    for index, raw_line in enumerate(lines):
        line = raw_line.strip()
        if not line or not re.match(r"^\d+\.\s+", line):
            continue
        if len(line) > 320:
            continue
        nearby = "\n".join(lines[index : min(index + 8, len(lines))])
        if "URL1:" not in nearby and "Новый правовой узел:" not in nearby and "Структурный элемент:" not in nearby:
            continue
        normalized = re.sub(r"\s+", " ", line)
        if normalized in seen:
            continue
        seen.add(normalized)
        anchors.append(f"- {normalized}")
        if len(anchors) >= max_items:
            break
    return anchors


def extract_last_substantive_paragraph(
    content: str,
    max_chars: int = 900,
) -> str:
    paragraphs = []
    for raw_paragraph in re.split(r"\n\s*\n", content):
        paragraph = raw_paragraph.strip()
        if not paragraph:
            continue
        if paragraph.startswith("```"):
            continue
        if "URL1:" in paragraph or "URL2:" in paragraph:
            continue
        paragraphs.append(re.sub(r"\s+", " ", paragraph))
    if not paragraphs:
        return ""
    tail = paragraphs[-1]
    if len(tail) > max_chars:
        tail = tail[: max_chars - 3].rstrip() + "..."
    return tail


def infer_part_decision_label(part_number: int) -> str:
    if part_number == 2:
        return "Базовый правовой и документный каркас подтвержден"
    if part_number == 3:
        return "Слои поиска и блоки покрытия разложены"
    if part_number == 4:
        return "Диапазоны документов и узлы покрытия расширены"
    if part_number == 5:
        return "Смежные правовые слои и специальные режимы добраны"
    if part_number in {6, 7, 8, 9}:
        return "Новый документный материал и дельта-усиление зафиксированы"
    if part_number == 10:
        return "Мини-конспект и опорные подтверждения собраны"
    if part_number == 11:
        return "Практические действия для дневника сформированы"
    return "Промежуточное решение по части зафиксировано"


def build_decision_history_bundle(
    run_workspace: SubtopicRunWorkspace,
    target_part_number: int,
) -> str:
    context_parts = collect_literal_context_parts(run_workspace, target_part_number)
    if not context_parts:
        return ""

    lines = [
        "## История решений этой же подтемы",
        "",
        "Ниже дана не история текста, а история уже принятых решений по этой же подтеме.",
        "Использовать ее как рабочую память: что уже подтверждено, на какие документы уже опираемся и что нужно переносить дальше.",
        "",
    ]
    for part_number, content in context_parts:
        title = FINAL_PART_TITLES.get(part_number, f"ЧАСТЬ {part_number}")
        url1_count = count_non_overlapping(content, "URL1:")
        url2_count = count_non_overlapping(content, "URL2:")
        verified_count = count_non_overlapping(content, "VERIFIED URL2:")
        anchors = extract_anchor_lines_for_followup(content)
        opening = extract_first_substantive_paragraphs(content)
        tail = extract_last_substantive_paragraph(content)
        lines.extend(
            [
                f"### ЧАСТЬ {part_number}. {title}",
                "",
                f"- Принятое решение: {infer_part_decision_label(part_number)}.",
                f"- Зафиксировано в части: URL1={url1_count}, URL2={url2_count}, VERIFIED_URL2={verified_count}.",
            ]
        )
        if opening:
            lines.extend(["", "Что уже установлено:", "", opening])
        if anchors:
            lines.extend(["", "На что уже можно опираться дальше:", ""])
            lines.extend(anchors)
        if tail:
            lines.extend(["", "Что переносится в следующую часть:", "", tail])
        lines.append("")
    return "\n".join(lines).rstrip()


def build_reasoning_layer_paths(run_workspace: SubtopicRunWorkspace) -> dict[str, Any]:
    reasoning_dir = run_workspace.web_plan_dir / "reasoning-layer"
    brief_files = {
        part_number: reasoning_dir / f"part-{part_number:02d}.reasoning.md"
        for part_number in REASONING_LAYER_PART_NUMBERS
    }
    return {
        "reasoning_dir": reasoning_dir,
        "readme": reasoning_dir / "README.reasoning-layer.md",
        "brief_files": brief_files,
    }


def build_reasoning_layer_readme(run_workspace: SubtopicRunWorkspace) -> str:
    paths = build_reasoning_layer_paths(run_workspace)
    lines = [
        f"# Reasoning Layer: {run_workspace.subtopic_entry.line}",
        "",
        "Этот каталог хранит reasoning-briefs по Частям 2–11.",
        "Они усиливают интеллектуальный режим поиска, фильтрации, дедупликации, дельта-аудита и синтеза по всей подтеме.",
        "",
    ]
    for part_number in REASONING_LAYER_PART_NUMBERS:
        title = FINAL_PART_TITLES.get(part_number, f"ЧАСТЬ {part_number}")
        lines.append(f"- Часть {part_number}. {title}: `{paths['brief_files'][part_number]}`")
    lines.append("")
    return "\n".join(lines)


def build_reasoning_part_brief(
    run_workspace: SubtopicRunWorkspace,
    part_number: int,
    heading_level: str = "##",
) -> str:
    guidance = REASONING_LAYER_GUIDANCE[part_number]
    title = FINAL_PART_TITLES.get(part_number, f"ЧАСТЬ {part_number}")
    focus = build_part_02_focus(run_workspace)
    completed_parts = [
        str(number)
        for number in range(1, part_number)
        if load_completed_part_output(run_workspace, number)
    ]
    completed_label = ", ".join(completed_parts) if completed_parts else "нет подтвержденных предыдущих частей"
    lines = [
        f"{heading_level} Reasoning-layer: Часть {part_number}. {title}",
        "",
        f"Цель reasoning-слоя: {guidance['goal']}",
        "",
        "Контекст подтемы:",
        f"- Подтема: `{run_workspace.subtopic_entry.line}`",
        f"- Тип подтемы: `{focus['subtopic_type']}`",
        f"- Фокус: `{focus['query_focus']}`",
        f"- Подтвержденные предыдущие части: {completed_label}",
        f"- Базовое правило исполнения: {focus['execution_note']}",
    ]
    if focus["substantive_sibling"]:
        sibling = focus["substantive_sibling"]
        lines.append(
            f"- Связанная базовая подтема: `{sibling['item_id']}. {ensure_trailing_period(sibling['title'])}`"
        )
    if focus["tariff_sibling"]:
        sibling = focus["tariff_sibling"]
        lines.append(
            f"- Связанная тарифная подтема: `{sibling['item_id']}. {ensure_trailing_period(sibling['title'])}`"
        )
    lines.extend(["", "Как думать на этой стадии:"])
    for item in guidance["moves"]:
        lines.append(f"- {item}")
    lines.extend(
        [
            "",
            "Модель верификации документа:",
            "- Сначала решать вопрос существования, применимости и роли документа в теме.",
            "- `URL1` обязателен как якорь официальности, если он реально найден.",
            "- `URL2` нужен для усиления карточки, но его неполная верификация не должна сама по себе выбрасывать документ.",
            "- Если документ подтвержден по существованию, применимости, роли и URL1, он остается в основном массиве с пометкой статуса `VERIFIED URL2`/`Сверка реквизитов`.",
            "- В `КАРАНТИН` документ уходит только при материальной неопределенности по самому документу, а не из-за одной слабой читабельной ссылки.",
        ]
    )
    lines.extend(["", "Типичные пропуски и ошибки:"])
    for item in guidance["risks"]:
        lines.append(f"- {item}")
    lines.extend(
        [
            "",
            "Общие запреты reasoning-слоя:",
            "- Не придумывать документы, роли или связи без подтвержденной документной опоры.",
            "- Не ослаблять охват только потому, что предыдущие части уже содержат большой объем материала.",
            "- Не превращать reasoning-layer в мета-комментарий вместо исполнения текущей Части.",
            "",
        ]
    )
    return "\n".join(lines)


def write_reasoning_layer(run_workspace: SubtopicRunWorkspace, overwrite: bool) -> dict[str, Any]:
    paths = build_reasoning_layer_paths(run_workspace)
    paths["reasoning_dir"].mkdir(parents=True, exist_ok=True)
    write_text_if_needed(paths["readme"], build_reasoning_layer_readme(run_workspace), overwrite)
    for part_number in REASONING_LAYER_PART_NUMBERS:
        write_text_if_needed(
            paths["brief_files"][part_number],
            build_reasoning_part_brief(run_workspace, part_number, heading_level="#"),
            overwrite,
        )
    return paths


def load_json_if_exists(path: Path) -> Any | None:
    if not path.exists():
        return None
    return json.loads(read_text(path))


def count_document_cards(text: str | None) -> int:
    if not text:
        return 0
    return text.count("Вид документа:")


def collect_manifest_part_statuses(run_workspace: SubtopicRunWorkspace) -> dict[int, str]:
    manifest_path = run_workspace.run_dir / "manifest.json"
    manifest = load_json_if_exists(manifest_path)
    if not manifest:
        return {}
    statuses: dict[int, str] = {}
    for item in manifest.get("parts", []):
        statuses[int(item.get("part_number", 0))] = str(item.get("status", ""))
    return statuses


def collect_manifest_part_metadata(run_workspace: SubtopicRunWorkspace) -> dict[int, dict[str, Any]]:
    manifest_path = run_workspace.run_dir / "manifest.json"
    manifest = load_json_if_exists(manifest_path)
    if not manifest:
        return {}
    metadata: dict[int, dict[str, Any]] = {}
    for item in manifest.get("parts", []):
        part_number = int(item.get("part_number", 0))
        if part_number:
            metadata[part_number] = item
    return metadata


def collect_untrusted_output_parts(
    run_workspace: SubtopicRunWorkspace, included_parts: list[int]
) -> list[dict[str, Any]]:
    metadata = collect_manifest_part_metadata(run_workspace)
    issues: list[dict[str, Any]] = []
    for part_number in included_parts:
        part_meta = metadata.get(part_number, {})
        source_origin = str(part_meta.get("source_origin", "")).strip()
        if source_origin not in FRESH_RUN_TRUSTED_SOURCE_ORIGINS:
            issues.append(
                {
                    "part_number": part_number,
                    "source_origin": source_origin or "unknown",
                    "status": str(part_meta.get("status", "")),
                }
            )
    return issues


def collect_present_parts(run_workspace: SubtopicRunWorkspace, max_part_number: int) -> list[int]:
    present: list[int] = []
    for part_number in range(1, max_part_number + 1):
        if load_completed_part_output(run_workspace, part_number):
            present.append(part_number)
    return present


def find_missing_canonical_part_03_blocks(text: str | None) -> list[str]:
    if not text:
        return list(PART_03_CANONICAL_BLOCKS.keys())
    found: set[str] = set()
    for roman, title in PART_03_CANONICAL_BLOCKS.items():
        pattern = re.compile(rf"(?m)^\s*\*{{0,2}}{re.escape(roman)}\.\s+{re.escape(title)}(?:$|\s|\*)")
        if pattern.search(text):
            found.add(roman)
    return [roman for roman in PART_03_CANONICAL_BLOCKS if roman not in found]


def get_capture_progress(capture_status_path: Path, expected_total: int) -> dict[str, Any]:
    payload = load_json_if_exists(capture_status_path) or {}
    captured = [int(item) for item in payload.get("captured_segment_ids", [])]
    remaining = [int(item) for item in payload.get("remaining_segment_ids", [])]
    return {
        "captured_count": len(captured),
        "remaining_count": len(remaining) if remaining else max(expected_total - len(captured), 0),
        "expected_total": expected_total,
    }


def build_omission_audit_paths(run_workspace: SubtopicRunWorkspace) -> dict[str, Any]:
    audit_dir = run_workspace.web_plan_dir / "omission-audit"
    brief_files = {
        part_number: audit_dir / f"part-{part_number:02d}.omission-audit.md"
        for part_number in OMISSION_AUDIT_PART_NUMBERS
    }
    return {
        "audit_dir": audit_dir,
        "readme": audit_dir / "README.omission-audit.md",
        "snapshot_json": audit_dir / "omission-audit.snapshot.json",
        "brief_files": brief_files,
    }


def build_omission_audit_snapshot(run_workspace: SubtopicRunWorkspace) -> dict[str, Any]:
    focus = build_part_02_focus(run_workspace)
    part_statuses = collect_manifest_part_statuses(run_workspace)
    completed_parts = collect_present_parts(run_workspace, 11)

    part_02_output = load_completed_part_output(run_workspace, 2)
    part_03_output = load_completed_part_output(run_workspace, 3)
    part_04_output = load_completed_part_output(run_workspace, 4)
    part_05_output = load_completed_part_output(run_workspace, 5)
    part_06_output = load_completed_part_output(run_workspace, 6)
    part_07_output = load_completed_part_output(run_workspace, 7)
    part_08_output = load_completed_part_output(run_workspace, 8)

    part_03_missing_blocks = find_missing_canonical_part_03_blocks(part_03_output)
    part_04_progress = get_capture_progress(
        build_part_04_plan_paths(run_workspace)["capture_status"],
        len(PART_04_SEGMENTS),
    )
    part_05_progress = get_capture_progress(
        build_part_05_plan_paths(run_workspace)["capture_status"],
        len(PART_05_SEGMENTS),
    )

    part_02_marker_state = {
        marker: bool(part_02_output and marker in part_02_output)
        for marker in PART_02_REQUIRED_MARKERS
    }

    heuristic_flags: list[str] = []
    missing_part_02_markers = [marker for marker, present in part_02_marker_state.items() if not present]
    if missing_part_02_markers:
        heuristic_flags.append(
            "В Части 2 отсутствуют обязательные маркеры: " + ", ".join(missing_part_02_markers)
        )
    if part_03_output and part_03_missing_blocks:
        heuristic_flags.append(
            f"В Части 3 отсутствуют канонические блоки: {', '.join(part_03_missing_blocks[:8])}"
            + (" ..." if len(part_03_missing_blocks) > 8 else "")
        )
    if part_04_progress["captured_count"] < part_04_progress["expected_total"]:
        heuristic_flags.append(
            f"Часть 4 закрыта не полностью: {part_04_progress['captured_count']} из {part_04_progress['expected_total']} диапазонов."
        )
    if part_05_progress["captured_count"] < part_05_progress["expected_total"]:
        heuristic_flags.append(
            f"Часть 5 закрыта не полностью: {part_05_progress['captured_count']} из {part_05_progress['expected_total']} диапазонов."
        )
    for part_number, output in ((6, part_06_output), (7, part_07_output), (8, part_08_output)):
        if not output:
            heuristic_flags.append(f"Часть {part_number} пока не подтверждена; omission audit должен учитывать этот пробел.")
    if part_02_output and "КАРАНТИН" in part_02_output:
        heuristic_flags.append("В Части 2 есть карантинный слой: нужно отдельно проверить, что из карантина так и осталось недожатым.")
    if focus["subtopic_type"] == "substantive" and focus["tariff_sibling"]:
        heuristic_flags.append(
            f"У подтемы есть отдельный тарифный sibling `{focus['tariff_sibling']['item_id']}`: omission audit не должен смешивать тарифный добор с базовой подтемой."
        )
    if focus["subtopic_type"] == "tariff" and focus["substantive_sibling"]:
        heuristic_flags.append(
            f"Тарифная подтема должна сохранять привязку к базовому действию `{focus['substantive_sibling']['item_id']}`."
        )

    audit_questions = [
        "Какие обязательные смысловые узлы темы выглядят недожатыми даже после Частей 2–8?",
        "Какие обязательные блоки, органы или слои пройдены формально, но слабо доказаны документами?",
        "Где фильтры и дедуп могли убрать полезный документ слишком рано?",
        "Какие карантинные или слабо верифицированные позиции стоит перепроверить отдельным добором?",
        "Какие скрытые режимы темы могли остаться вне поля зрения: международный, цифровой, контрольный, процессуальный, финансовый?",
    ]

    return {
        "generated_at": utc_now_iso(),
        "subtopic_id": run_workspace.subtopic_entry.item_id,
        "subtopic_line": run_workspace.subtopic_entry.line,
        "subtopic_type": focus["subtopic_type"],
        "query_focus": focus["query_focus"],
        "completed_parts": completed_parts,
        "part_statuses": part_statuses,
        "part_02": {
            "markers_present": part_02_marker_state,
            "document_cards": count_document_cards(part_02_output),
            "has_quarantine": bool(part_02_output and "КАРАНТИН" in part_02_output),
        },
        "part_03": {
            "document_cards": count_document_cards(part_03_output),
            "missing_canonical_blocks": part_03_missing_blocks,
            "covered_blocks_count": len(PART_03_CANONICAL_BLOCKS) - len(part_03_missing_blocks),
            "expected_blocks_count": len(PART_03_CANONICAL_BLOCKS),
        },
        "part_04": {
            "document_cards": count_document_cards(part_04_output),
            **part_04_progress,
        },
        "part_05": {
            "document_cards": count_document_cards(part_05_output),
            **part_05_progress,
        },
        "part_06": {"present": bool(part_06_output), "document_cards": count_document_cards(part_06_output)},
        "part_07": {"present": bool(part_07_output), "document_cards": count_document_cards(part_07_output)},
        "part_08": {"present": bool(part_08_output), "document_cards": count_document_cards(part_08_output)},
        "heuristic_flags": heuristic_flags,
        "audit_questions": audit_questions,
    }


def build_omission_audit_readme(run_workspace: SubtopicRunWorkspace) -> str:
    paths = build_omission_audit_paths(run_workspace)
    lines = [
        f"# Omission Audit Layer: {run_workspace.subtopic_entry.line}",
        "",
        "Этот каталог хранит второй проход на пропуски после Частей 2–8.",
        "Его задача — не повторить список документов, а подсветить слабые места покрытия и подготовить сильный Part 9.",
        "",
        f"- Snapshot: `{paths['snapshot_json']}`",
    ]
    for part_number in OMISSION_AUDIT_PART_NUMBERS:
        lines.append(f"- Brief for Part {part_number}: `{paths['brief_files'][part_number]}`")
    lines.append("")
    return "\n".join(lines)


def build_omission_audit_brief(run_workspace: SubtopicRunWorkspace, part_number: int) -> str:
    snapshot = build_omission_audit_snapshot(run_workspace)
    title = FINAL_PART_TITLES.get(part_number, f"ЧАСТЬ {part_number}")
    lines = [
        f"## Omission Audit Layer: Часть {part_number}. {title}",
        "",
        "Это второй проход на пропуски. Его цель — не переписать предыдущие части, а найти то, что еще недожато или выглядит слабо подтвержденным.",
        "",
        "Снимок текущего покрытия:",
        f"- Подтема: `{snapshot['subtopic_line']}`",
        f"- Тип подтемы: `{snapshot['subtopic_type']}`",
        f"- Фокус: `{snapshot['query_focus']}`",
        f"- Подтвержденные части: {', '.join(map(str, snapshot['completed_parts'])) if snapshot['completed_parts'] else 'нет'}",
        f"- Часть 2: карточек документов = {snapshot['part_02']['document_cards']}, карантин = {snapshot['part_02']['has_quarantine']}",
        f"- Часть 3: блоков закрыто = {snapshot['part_03']['covered_blocks_count']} из {snapshot['part_03']['expected_blocks_count']}",
        f"- Часть 4: диапазонов закрыто = {snapshot['part_04']['captured_count']} из {snapshot['part_04']['expected_total']}",
        f"- Часть 5: диапазонов закрыто = {snapshot['part_05']['captured_count']} из {snapshot['part_05']['expected_total']}",
        f"- Часть 6 подтверждена = {snapshot['part_06']['present']}",
        f"- Часть 7 подтверждена = {snapshot['part_07']['present']}",
        f"- Часть 8 подтверждена = {snapshot['part_08']['present']}",
        "",
        "Подозрительные места и флаги:",
    ]
    if snapshot["heuristic_flags"]:
        for item in snapshot["heuristic_flags"]:
            lines.append(f"- {item}")
    else:
        lines.append("- Явных структурных флагов по состоянию run не обнаружено, но omission audit все равно обязателен.")
    lines.extend(["", "Контрольные вопросы второго прохода:"])
    for item in snapshot["audit_questions"]:
        lines.append(f"- {item}")
    lines.extend(
        [
            "",
            "Режим omission audit:",
            "- Искать не количество документов, а реальный недожатый смысловой слой.",
            "- Проверять не только отсутствие документа, но и слабость верификации уже найденного документа.",
            "- Не повторять уже сильные карточки без новой дельты.",
            "",
        ]
    )
    return "\n".join(lines)


def write_omission_audit_layer(run_workspace: SubtopicRunWorkspace, overwrite: bool) -> dict[str, Any]:
    paths = build_omission_audit_paths(run_workspace)
    paths["audit_dir"].mkdir(parents=True, exist_ok=True)
    write_text_if_needed(paths["readme"], build_omission_audit_readme(run_workspace), overwrite)
    write_json(paths["snapshot_json"], build_omission_audit_snapshot(run_workspace))
    for part_number in OMISSION_AUDIT_PART_NUMBERS:
        write_text_if_needed(
            paths["brief_files"][part_number],
            build_omission_audit_brief(run_workspace, part_number),
            overwrite,
        )
    return paths


def strip_document_card_prefix(value: str) -> str:
    return re.sub(r"^\d+\.\s+", "", value).strip()


def normalize_card_field_label(label: str) -> str:
    lowered = strip_document_card_prefix(label).strip().lower().replace("ё", "е")
    if lowered.startswith("вид документа"):
        return "doc_type"
    if lowered.startswith("полное официальное наименование документа"):
        return "full_name"
    if lowered.startswith("полное наименование"):
        return "full_name"
    if lowered in {"орган", "орган принятия (издания)", "орган принятия", "орган издания"}:
        return "organ"
    if lowered.startswith("дата и номер"):
        return "date_number"
    if lowered == "дата":
        return "date"
    if "структурный элемент" in lowered or "структурные элементы" in lowered:
        return "structure"
    if lowered.startswith("значение для темы"):
        return "significance"
    if lowered.startswith("url1"):
        return "url1"
    if lowered.startswith("url2"):
        return "url2"
    if lowered.startswith("каноническая строка поиска"):
        return "search_query"
    return normalize_search_key(lowered).replace(" ", "_")


def extract_card_fields(card_lines: list[str]) -> dict[str, str]:
    fields: dict[str, str] = {}
    current_label: str | None = None
    in_code_block = False
    for raw_line in card_lines:
        stripped = raw_line.strip()
        if stripped.startswith("```"):
            in_code_block = not in_code_block
            continue
        match = re.match(r"^(?:\d+\.\s+)?([^:]{1,120}):\s*(.*)$", stripped)
        if match:
            current_label = normalize_card_field_label(match.group(1))
            value = match.group(2).strip()
            if value:
                fields[current_label] = value
            else:
                fields.setdefault(current_label, "")
            continue
        if current_label and stripped:
            separator = "" if fields.get(current_label, "").endswith("/") else " "
            existing = fields.get(current_label, "")
            fields[current_label] = f"{existing}{separator}{stripped}".strip()
        elif in_code_block and stripped:
            current_label = None
    return fields


def infer_document_role(part_number: int, scope_heading: str, fields: dict[str, str]) -> tuple[str, str]:
    normalized_scope = normalize_search_key(scope_heading) or "__global__"
    display_scope = scope_heading if scope_heading and scope_heading != "__global__" else "глобальный контекст части"
    if part_number == 2:
        if scope_heading == "A. РЕГУЛЯТОРНОЕ ЯДРО":
            bucket = "regulatory_core"
            label = "регуляторное ядро"
        elif scope_heading == "B. ОПОРНЫЕ ДОКУМЕНТЫ":
            bucket = "supporting_documents"
            label = "опорные документы"
        elif scope_heading == "КАРАНТИН":
            bucket = "quarantine"
            label = "карантин"
        elif scope_heading == "FAIL-SAFE CHECK":
            bucket = "fail_safe"
            label = "fail-safe"
        else:
            bucket = "part_02"
            label = "Часть 2"
    elif part_number == 3:
        bucket = "search_sphere"
        label = f"сфера поиска: {display_scope}"
    elif part_number == 4:
        bucket = "authority_search"
        label = f"орган/подпункт: {display_scope}"
    elif part_number == 5:
        bucket = "layering"
        label = f"слой: {display_scope}"
    elif part_number == 6:
        bucket = "filter"
        label = "фильтр"
    elif part_number == 7:
        bucket = "new_without_repeats"
        label = "новые документы без повторов"
    elif part_number == 8:
        bucket = "federal_delta"
        label = "федеральная дельта"
    else:
        bucket = f"part_{part_number:02d}"
        label = f"Часть {part_number}"

    significance = normalize_search_key(fields.get("significance", ""))
    if bucket in {"part_02", "supporting_documents"} and "карантин" in significance:
        bucket = "quarantine"
        label = "карантин"
    return f"{bucket}::{normalized_scope}", label


def build_document_identity_key(fields: dict[str, str]) -> str:
    full_name = normalize_search_key(fields.get("full_name", ""))
    doc_type = normalize_search_key(fields.get("doc_type", ""))
    organ = normalize_search_key(fields.get("organ", ""))
    date_number = normalize_search_key(fields.get("date_number", "") or fields.get("date", ""))
    url1 = normalize_search_key(fields.get("url1", ""))
    if full_name:
        pieces = [full_name]
        if date_number:
            pieces.append(date_number)
        return " | ".join(pieces)
    fallback = " | ".join(item for item in (doc_type, organ, date_number, url1) if item)
    return fallback or "__unknown_document__"


def parse_document_cards(text: str | None, part_number: int) -> list[dict[str, Any]]:
    if not text:
        return []
    cards: list[dict[str, Any]] = []
    current_scope = "__global__"
    in_code_block = False
    card_lines: list[str] = []
    card_scope = current_scope

    def flush_card() -> None:
        nonlocal card_lines, card_scope
        if not card_lines:
            return
        fields = extract_card_fields(card_lines)
        if "doc_type" not in fields:
            card_lines = []
            return
        role_key, role_label = infer_document_role(part_number, card_scope, fields)
        display_name = (
            fields.get("full_name")
            or fields.get("doc_type")
            or "Неопознанный документ"
        )
        cards.append(
            {
                "part_number": part_number,
                "scope_heading": card_scope,
                "role_key": role_key,
                "role_label": role_label,
                "identity_key": build_document_identity_key(fields),
                "display_name": display_name,
                "fields": fields,
                "raw": "\n".join(card_lines).strip(),
            }
        )
        card_lines = []

    for raw_line in text.splitlines():
        stripped = raw_line.strip()
        if stripped.startswith("```"):
            in_code_block = not in_code_block
            if card_lines:
                card_lines.append(raw_line)
            continue
        if not in_code_block and is_structural_heading(stripped):
            flush_card()
            current_scope = stripped
            continue
        if re.match(r"^(?:\d+\.\s+)?Вид документа:", stripped):
            flush_card()
            card_scope = current_scope
            card_lines = [raw_line]
            continue
        if card_lines:
            card_lines.append(raw_line)
    flush_card()
    return cards


def collect_semantic_dedup_cards(run_workspace: SubtopicRunWorkspace) -> list[dict[str, Any]]:
    cards: list[dict[str, Any]] = []
    for part_number in range(2, 9):
        cards.extend(parse_document_cards(load_completed_part_output(run_workspace, part_number), part_number))
    return cards


def build_semantic_dedup_paths(run_workspace: SubtopicRunWorkspace) -> dict[str, Any]:
    dedup_dir = run_workspace.web_plan_dir / "semantic-dedup"
    brief_files = {
        part_number: dedup_dir / f"part-{part_number:02d}.semantic-dedup.md"
        for part_number in SEMANTIC_DEDUP_PART_NUMBERS
    }
    return {
        "dedup_dir": dedup_dir,
        "readme": dedup_dir / "README.semantic-dedup.md",
        "snapshot_json": dedup_dir / "semantic-dedup.snapshot.json",
        "brief_files": brief_files,
    }


def build_semantic_dedup_snapshot(run_workspace: SubtopicRunWorkspace) -> dict[str, Any]:
    cards = collect_semantic_dedup_cards(run_workspace)
    clusters: dict[str, list[dict[str, Any]]] = {}
    for card in cards:
        clusters.setdefault(card["identity_key"], []).append(card)

    multi_role_documents: list[dict[str, Any]] = []
    same_role_duplicates: list[dict[str, Any]] = []
    semantic_delta_candidates: list[dict[str, Any]] = []

    for identity_key, cluster_cards in clusters.items():
        if len(cluster_cards) < 2:
            continue
        role_map: dict[str, list[dict[str, Any]]] = {}
        for card in cluster_cards:
            role_map.setdefault(card["role_key"], []).append(card)
        display_name = cluster_cards[0]["display_name"]
        roles = sorted({card["role_label"] for card in cluster_cards})
        parts = sorted({card["part_number"] for card in cluster_cards})
        if len(role_map) > 1:
            multi_role_documents.append(
                {
                    "document": display_name,
                    "identity_key": identity_key,
                    "roles": roles,
                    "parts": parts,
                    "count": len(cluster_cards),
                }
            )
        repeated_roles = [
            {
                "role_label": items[0]["role_label"],
                "count": len(items),
                "parts": sorted({item["part_number"] for item in items}),
            }
            for items in role_map.values()
            if len(items) > 1
        ]
        if repeated_roles:
            same_role_duplicates.append(
                {
                    "document": display_name,
                    "identity_key": identity_key,
                    "repeated_roles": repeated_roles,
                    "parts": parts,
                    "count": len(cluster_cards),
                }
            )
        if len(role_map) > 1 and any(item["part_number"] in {7, 8} for item in cluster_cards):
            semantic_delta_candidates.append(
                {
                    "document": display_name,
                    "roles": roles,
                    "parts": parts,
                }
            )

    def trim(items: list[dict[str, Any]], limit: int = 12) -> list[dict[str, Any]]:
        return items[:limit]

    return {
        "generated_at": utc_now_iso(),
        "subtopic_id": run_workspace.subtopic_entry.item_id,
        "subtopic_line": run_workspace.subtopic_entry.line,
        "cards_total": len(cards),
        "unique_documents_total": len(clusters),
        "duplicate_documents_total": sum(1 for cluster in clusters.values() if len(cluster) > 1),
        "same_role_duplicates_total": len(same_role_duplicates),
        "multi_role_documents_total": len(multi_role_documents),
        "same_role_duplicates": trim(sorted(same_role_duplicates, key=lambda item: (-item["count"], item["document"]))),
        "multi_role_documents": trim(sorted(multi_role_documents, key=lambda item: (-item["count"], item["document"]))),
        "semantic_delta_candidates": trim(sorted(semantic_delta_candidates, key=lambda item: (item["document"]))),
    }


def build_semantic_dedup_readme(run_workspace: SubtopicRunWorkspace) -> str:
    paths = build_semantic_dedup_paths(run_workspace)
    lines = [
        f"# Semantic Dedup Layer: {run_workspace.subtopic_entry.line}",
        "",
        "Этот каталог хранит карту повторов документов по Частям 2–8.",
        "Его задача — различать повтор того же акта в той же роли и допустимое повторное появление того же акта в новой роли.",
        "",
        f"- Snapshot: `{paths['snapshot_json']}`",
    ]
    for part_number in SEMANTIC_DEDUP_PART_NUMBERS:
        lines.append(f"- Brief for Part {part_number}: `{paths['brief_files'][part_number]}`")
    lines.append("")
    return "\n".join(lines)


def build_semantic_dedup_brief(run_workspace: SubtopicRunWorkspace, part_number: int) -> str:
    snapshot = build_semantic_dedup_snapshot(run_workspace)
    title = FINAL_PART_TITLES.get(part_number, f"ЧАСТЬ {part_number}")
    lines = [
        f"## Semantic Dedup Layer: Часть {part_number}. {title}",
        "",
        "Смысловой дедуп работает так: один и тот же документ удаляется только если он повторяется в той же роли; если документ появляется в новой роли, его можно сохранять.",
        "",
        f"- Всего карточек документов в Частях 2–8: {snapshot['cards_total']}",
        f"- Уникальных документов: {snapshot['unique_documents_total']}",
        f"- Документов с повторами: {snapshot['duplicate_documents_total']}",
        f"- Повторов в той же роли: {snapshot['same_role_duplicates_total']}",
        f"- Документов, которые уже живут в нескольких ролях: {snapshot['multi_role_documents_total']}",
        "",
        "Как применять этот слой:",
        "- Если документ уже найден в той же роли и в том же смысловом слое, не дублировать его снова.",
        "- Если тот же документ закрывает новый слой, новый блок или новую дельту, сохранить его, но явно показать новую роль.",
        "- При сомнении смотреть не на название документа, а на его функцию в текущей Части.",
        "",
        "Кандидаты на удаление как повторы в той же роли:",
    ]
    if snapshot["same_role_duplicates"]:
        for item in snapshot["same_role_duplicates"][:6]:
            repeated_roles = "; ".join(
                f"{role['role_label']} ({role['count']})"
                for role in item["repeated_roles"]
            )
            lines.append(f"- `{item['document']}` -> {repeated_roles}")
    else:
        lines.append("- Явных повторов в той же роли по текущему snapshot не найдено.")

    lines.extend(["", "Документы, которые уже используются в нескольких ролях:"])
    if snapshot["multi_role_documents"]:
        for item in snapshot["multi_role_documents"][:6]:
            lines.append(f"- `{item['document']}` -> роли: {', '.join(item['roles'])}")
    else:
        lines.append("- Пока не видно документов с несколькими подтвержденными ролями.")

    lines.extend(
        [
            "",
            "Контрольные вопросы semantic dedup:",
            "- Это действительно новый документ, или просто другая ссылка на уже найденный акт?",
            "- Если документ повторяется, он играет новую роль или снова делает ту же работу?",
            "- Если документ оставляется, какую новую смысловую функцию он закрывает именно в этой Части?",
            "",
        ]
    )
    return "\n".join(lines)


def write_semantic_dedup_layer(run_workspace: SubtopicRunWorkspace, overwrite: bool) -> dict[str, Any]:
    paths = build_semantic_dedup_paths(run_workspace)
    paths["dedup_dir"].mkdir(parents=True, exist_ok=True)
    write_text_if_needed(paths["readme"], build_semantic_dedup_readme(run_workspace), overwrite)
    write_json(paths["snapshot_json"], build_semantic_dedup_snapshot(run_workspace))
    for part_number in SEMANTIC_DEDUP_PART_NUMBERS:
        write_text_if_needed(
            paths["brief_files"][part_number],
            build_semantic_dedup_brief(run_workspace, part_number),
            overwrite,
        )
    return paths


def build_followup_part_boosters(part_number: int) -> list[str]:
    if part_number in {6, 7, 8, 9}:
        return [
            "Эта Часть обязана делать новый внешний web-search по своим узлам, а не только пересказывать уже собранный массив.",
            "Ответ `новых документов не выявлено` допустим только после повторного внешнего прохода по официальным источникам, редакционным цепочкам и свежим изменениям.",
            "Если найден новый акт или новый релевантный акт об изменениях, он должен быть оформлен полной карточкой документа, а не коротким упоминанием.",
            "Не сушить выдачу: если документ реально усиливает тему, его нужно оставить даже при пересечении с уже найденным корпусом.",
        ]
    if part_number == 10:
        return [
            "Мини-конспект должен быть насыщенным и прикладным, а не сухим narrative-текстом.",
            "После каждого пункта мини-конспекта нужно давать копируемые ссылки в code-блоках с полным наименованием документа и структурным элементом для подтверждения содержания этого пункта.",
            "Нумеровать только пункты; строки текста внутри пункта не нумеровать.",
        ]
    if part_number == 11:
        return [
            "Документы перечислять не надо: они уже были даны в предыдущих Частях.",
            "Нужно писать только для 4-й колонки дневника как перечень выполненных стажером действий по характеру задания.",
        ]
    return []


def build_followup_part_packet(run_workspace: SubtopicRunWorkspace, part_number: int) -> str:
    def guard_no_canonical_inline(payload: str) -> None:
        canonical_paths = [
            run_workspace.context_dir / "00-agents.md",
            run_workspace.context_dir / "01-project-state.md",
            run_workspace.context_dir / "02-manual-workflow.md",
            run_workspace.context_dir / "03-master-prompt.md",
            run_workspace.context_dir / "06-order.copy.md",
            run_workspace.context_dir / "08-interaction-guide.md",
        ]
        for canonical_path in canonical_paths:
            if not canonical_path.exists():
                continue
            canonical_text = read_text(canonical_path).strip()
            if canonical_text and canonical_text in payload:
                relative_path = canonical_path.relative_to(run_workspace.run_dir).as_posix()
                raise ValueError(f"Canonical file must not be inlined: {relative_path}")

    part = get_order_part(run_workspace, part_number)
    context_md = (
        build_literal_context_bundle(run_workspace, part_number)
        if run_workspace.packet_mode == PACKET_MODE_LITERAL
        else build_decision_history_bundle(run_workspace, part_number)
    )
    reasoning_brief_md = build_reasoning_part_brief(run_workspace, part_number)
    title = FINAL_PART_TITLES.get(part_number, f"ЧАСТЬ {part_number}")
    lines = [
        f"# Part {part_number:02d} Continuation Packet: {run_workspace.subtopic_entry.line}",
        "",
        "Режим: максимально буквальное продолжение той же LLM-сессии по этой подтеме.",
        "- Не перезапускать Части 1–5 и не возвращаться к `ТЕМА/АНАЛИЗ`, если текущая Часть этого не требует.",
        "- Не ослаблять охват: использовать уже найденный материал как опору и дожимать именно текущую Часть.",
        "",
        "## Ссылки на канон",
        "",
        "[см. 00-context/00-agents.md]",
        "[см. 00-context/01-project-state.md]",
        "[см. 00-context/02-manual-workflow.md]",
        "[см. 00-context/03-master-prompt.md]",
        "[см. 00-context/06-order.copy.md]",
        "[см. 00-context/08-interaction-guide.md]",
        "",
    ]
    if part_number == 11:
        lines.insert(
            5,
            "- Для Части 11 документы не перечислять; нужен только перечень выполненных стажером действий для 4-й колонки дневника по характеру задания.",
        )
    elif part_number == 10:
        lines.insert(
            5,
            "- Для Части 10 после каждого пункта давать копируемые ссылочные code-блоки с полным наименованием документа и структурным элементом.",
        )
    else:
        lines.insert(
            5,
            "- Все найденные документы оформлять карточками с `Вид документа`, `Полное наименование`, `Структурный элемент`, `URL1`, `URL2`, `VERIFIED URL2`, `Заголовок страницы URL2`, `Сверка реквизитов` и `Каноническая строка поиска` по мере применимости.",
        )
    lines.extend([reasoning_brief_md, ""])
    boosters = build_followup_part_boosters(part_number)
    if boosters:
        lines.extend(["## Усилители текущей Части", ""])
        for item in boosters:
            lines.append(f"- {item}")
        lines.append("")
    if context_md:
        lines.extend([context_md, ""])
    lines.extend(
        [
            f"## Точная инструкция Части {part_number}",
            "",
            part.content.strip(),
            "",
            "## Что вернуть",
            "",
            f"Вернуть только готовый ответ по Части {part_number} без перезапуска предыдущих частей и без служебных пояснений вне ответа.",
            "",
        ]
    )
    packet_text = "\n".join(lines)
    guard_no_canonical_inline(packet_text)
    return packet_text


def build_followup_part_packet_paths(run_workspace: SubtopicRunWorkspace) -> dict[str, Any]:
    plan_dir = run_workspace.web_plan_dir / "follow-up-parts"
    message_files = {
        part_number: plan_dir / f"message.part-{part_number:02d}.md"
        for part_number in FOLLOWUP_LITERAL_PART_NUMBERS
    }
    return {
        "plan_dir": plan_dir,
        "readme": plan_dir / "README.follow-up-parts.md",
        "message_files": message_files,
    }


def build_followup_part_packets_readme(run_workspace: SubtopicRunWorkspace) -> str:
    paths = build_followup_part_packet_paths(run_workspace)
    lines = [
        f"# Follow-up Part Packets: {run_workspace.subtopic_entry.line}",
        "",
        "Эти файлы нужны для Частей 6–11. В отличие от сырых stage-inputs они включают живой контекст уже завершенных частей этой же подтемы и прямые усилители против пересушки выдачи.",
        "",
    ]
    for part_number in FOLLOWUP_LITERAL_PART_NUMBERS:
        lines.append(f"- Часть {part_number}: `{paths['message_files'][part_number]}`")
    lines.append("")
    return "\n".join(lines)


def write_followup_part_packets(run_workspace: SubtopicRunWorkspace, overwrite: bool) -> dict[str, Any]:
    paths = build_followup_part_packet_paths(run_workspace)
    paths["plan_dir"].mkdir(parents=True, exist_ok=True)
    write_text_if_needed(paths["readme"], build_followup_part_packets_readme(run_workspace), overwrite)
    for part_number in FOLLOWUP_LITERAL_PART_NUMBERS:
        write_text_if_needed(
            paths["message_files"][part_number],
            build_followup_part_packet(run_workspace, part_number),
            overwrite,
        )
    return paths


def refresh_dynamic_part_packets(run_workspace: SubtopicRunWorkspace) -> None:
    if not run_workspace.lean_artifacts:
        write_text(run_workspace.run_dir / "README.run.md", build_subtopic_run_readme(run_workspace))

        reasoning_paths = build_reasoning_layer_paths(run_workspace)
        reasoning_paths["reasoning_dir"].mkdir(parents=True, exist_ok=True)
        write_text(reasoning_paths["readme"], build_reasoning_layer_readme(run_workspace))
        for part_number in REASONING_LAYER_PART_NUMBERS:
            write_text(
                reasoning_paths["brief_files"][part_number],
                build_reasoning_part_brief(run_workspace, part_number, heading_level="#"),
            )

    if PRODUCTION_ENABLE_SEMANTIC_DEDUP:
        semantic_dedup_paths = build_semantic_dedup_paths(run_workspace)
        semantic_dedup_paths["dedup_dir"].mkdir(parents=True, exist_ok=True)
        write_text(semantic_dedup_paths["readme"], build_semantic_dedup_readme(run_workspace))
        write_json(semantic_dedup_paths["snapshot_json"], build_semantic_dedup_snapshot(run_workspace))
        for part_number in SEMANTIC_DEDUP_PART_NUMBERS:
            write_text(
                semantic_dedup_paths["brief_files"][part_number],
                build_semantic_dedup_brief(run_workspace, part_number),
            )

    if PRODUCTION_ENABLE_OMISSION_AUDIT:
        omission_paths = build_omission_audit_paths(run_workspace)
        omission_paths["audit_dir"].mkdir(parents=True, exist_ok=True)
        write_text(omission_paths["readme"], build_omission_audit_readme(run_workspace))
        write_json(omission_paths["snapshot_json"], build_omission_audit_snapshot(run_workspace))
        for part_number in OMISSION_AUDIT_PART_NUMBERS:
            write_text(
                omission_paths["brief_files"][part_number],
                build_omission_audit_brief(run_workspace, part_number),
            )

    part_02_paths = build_part_02_web_plan_paths(run_workspace)
    write_text(part_02_paths["launch_packet"], build_part_02_launch_packet(run_workspace))
    write_text(part_02_paths["message_03"], read_text(part_02_paths["launch_packet"]).rstrip() + "\n")

    if not run_workspace.lean_artifacts:
        part_03_paths = build_part_03_plan_paths(run_workspace)
        for segment in PART_03_SEGMENTS:
            write_text(
                part_03_paths["message_files"][segment["segment_id"]],
                build_part_03_message(run_workspace, segment),
            )

        part_04_paths = build_part_04_plan_paths(run_workspace)
        for segment in PART_04_SEGMENTS:
            write_text(
                part_04_paths["message_files"][segment["segment_id"]],
                build_part_04_message(run_workspace, segment),
            )

        part_05_paths = build_part_05_plan_paths(run_workspace)
        for segment in PART_05_SEGMENTS:
            write_text(
                part_05_paths["message_files"][segment["segment_id"]],
                build_part_05_message(run_workspace, segment),
            )

        followup_paths = build_followup_part_packet_paths(run_workspace)
        followup_paths["plan_dir"].mkdir(parents=True, exist_ok=True)
        write_text(followup_paths["readme"], build_followup_part_packets_readme(run_workspace))
        for part_number in FOLLOWUP_LITERAL_PART_NUMBERS:
            write_text(
                followup_paths["message_files"][part_number],
                build_followup_part_packet(run_workspace, part_number),
            )


def write_part_02_web_plan(run_workspace: SubtopicRunWorkspace, overwrite: bool) -> dict[str, Path]:
    paths = build_part_02_web_plan_paths(run_workspace)
    paths["web_plan_dir"].mkdir(parents=True, exist_ok=True)
    paths["evidence_dir"].mkdir(parents=True, exist_ok=True)

    focus = build_part_02_focus(run_workspace)
    queries = build_part_02_queries(run_workspace)

    write_text_if_needed(paths["readme"], build_part_02_web_readme(run_workspace), overwrite)
    write_text_if_needed(paths["operator_sequence"], build_llm_operator_sequence(run_workspace), overwrite)
    write_json(paths["focus_json"], focus) if overwrite or not paths["focus_json"].exists() else None
    write_text_if_needed(paths["queries_md"], render_part_02_queries_markdown(run_workspace, queries), overwrite)
    write_json(paths["queries_json"], queries) if overwrite or not paths["queries_json"].exists() else None
    write_text_if_needed(
        paths["source_cascade_md"],
        render_part_02_source_cascade_markdown(PART_02_SOURCE_CASCADE),
        overwrite,
    )
    write_json(paths["source_cascade_json"], PART_02_SOURCE_CASCADE) if overwrite or not paths["source_cascade_json"].exists() else None
    write_text_if_needed(paths["research_pack"], build_part_02_research_pack(run_workspace), overwrite)
    write_text_if_needed(paths["core_template"], build_part_02_core_template(run_workspace), overwrite)
    write_text_if_needed(paths["launch_packet"], build_part_02_launch_packet(run_workspace), overwrite)
    part_01_output_path = run_workspace.stage_outputs_dir / "part-01.md"
    part_01_payload = read_text(part_01_output_path) if part_01_output_path.exists() else build_part_01_response(run_workspace)
    write_text_if_needed(paths["message_01"], part_01_payload.rstrip() + "\n", overwrite)
    write_text_if_needed(paths["message_02"], "GO/СТАРТ\n", overwrite)
    write_text_if_needed(paths["message_03"], read_text(paths["launch_packet"]).rstrip() + "\n", overwrite)
    if overwrite or not paths["research_log"].exists():
        write_text(paths["research_log"], "")
    return paths


def update_run_manifest_web_plan(run_workspace: SubtopicRunWorkspace, plan_paths: dict[str, Path]) -> None:
    manifest_path = run_workspace.run_dir / "manifest.json"
    if not manifest_path.exists():
        return
    manifest = json.loads(read_text(manifest_path))
    manifest["part_02_web_plan"] = {
        "generated_at": utc_now_iso(),
        "web_plan_dir": str(plan_paths["web_plan_dir"]),
        "operator_sequence": str(plan_paths["operator_sequence"]),
        "focus_json": str(plan_paths["focus_json"]),
        "queries_md": str(plan_paths["queries_md"]),
        "queries_json": str(plan_paths["queries_json"]),
        "source_cascade_md": str(plan_paths["source_cascade_md"]),
        "source_cascade_json": str(plan_paths["source_cascade_json"]),
        "research_pack": str(plan_paths["research_pack"]),
        "core_template": str(plan_paths["core_template"]),
        "launch_packet": str(plan_paths["launch_packet"]),
        "message_01": str(plan_paths["message_01"]),
        "message_02": str(plan_paths["message_02"]),
        "message_03": str(plan_paths["message_03"]),
        "research_log": str(plan_paths["research_log"]),
        "evidence_dir": str(plan_paths["evidence_dir"]),
    }
    write_json(manifest_path, manifest)


def update_run_manifest_reasoning_layer(run_workspace: SubtopicRunWorkspace, reasoning_paths: dict[str, Any]) -> None:
    manifest_path = run_workspace.run_dir / "manifest.json"
    if not manifest_path.exists():
        return
    manifest = json.loads(read_text(manifest_path))
    manifest["reasoning_layer"] = {
        "generated_at": utc_now_iso(),
        "reasoning_dir": str(reasoning_paths["reasoning_dir"]),
        "readme": str(reasoning_paths["readme"]),
        "brief_files": {
            str(part_number): str(path)
            for part_number, path in reasoning_paths["brief_files"].items()
        },
    }
    write_json(manifest_path, manifest)


def update_run_manifest_omission_audit(run_workspace: SubtopicRunWorkspace, omission_paths: dict[str, Any]) -> None:
    manifest_path = run_workspace.run_dir / "manifest.json"
    if not manifest_path.exists():
        return
    manifest = json.loads(read_text(manifest_path))
    manifest["omission_audit"] = {
        "generated_at": utc_now_iso(),
        "audit_dir": str(omission_paths["audit_dir"]),
        "readme": str(omission_paths["readme"]),
        "snapshot_json": str(omission_paths["snapshot_json"]),
        "brief_files": {
            str(part_number): str(path)
            for part_number, path in omission_paths["brief_files"].items()
        },
    }
    write_json(manifest_path, manifest)


def update_run_manifest_semantic_dedup(run_workspace: SubtopicRunWorkspace, dedup_paths: dict[str, Any]) -> None:
    manifest_path = run_workspace.run_dir / "manifest.json"
    if not manifest_path.exists():
        return
    manifest = json.loads(read_text(manifest_path))
    manifest["semantic_dedup"] = {
        "generated_at": utc_now_iso(),
        "dedup_dir": str(dedup_paths["dedup_dir"]),
        "readme": str(dedup_paths["readme"]),
        "snapshot_json": str(dedup_paths["snapshot_json"]),
        "brief_files": {
            str(part_number): str(path)
            for part_number, path in dedup_paths["brief_files"].items()
        },
    }
    write_json(manifest_path, manifest)


def drop_disabled_manifest_layers(run_workspace: SubtopicRunWorkspace) -> None:
    manifest_path = run_workspace.run_dir / "manifest.json"
    if not manifest_path.exists():
        return
    manifest = json.loads(read_text(manifest_path))
    changed = False
    if not PRODUCTION_ENABLE_SEMANTIC_DEDUP and "semantic_dedup" in manifest:
        manifest.pop("semantic_dedup", None)
        changed = True
    if not PRODUCTION_ENABLE_OMISSION_AUDIT and "omission_audit" in manifest:
        manifest.pop("omission_audit", None)
        changed = True
    if changed:
        write_json(manifest_path, manifest)


def build_part_03_plan_paths(run_workspace: SubtopicRunWorkspace) -> dict[str, Any]:
    plan_dir = run_workspace.web_plan_dir / "part-03"
    message_files = {
        segment["segment_id"]: plan_dir / f"message-{segment['segment_id']:02d}.part-03.range-{segment['segment_id']:02d}.md"
        for segment in PART_03_SEGMENTS
    }
    return {
        "plan_dir": plan_dir,
        "readme": plan_dir / "README.part-03.md",
        "operator_sequence": plan_dir / "00-operator-sequence.md",
        "canonical_skeleton": plan_dir / "part-03.canonical-skeleton.md",
        "capture_status": plan_dir / "capture-status.json",
        "message_files": message_files,
    }


def build_part_03_canonical_skeleton() -> str:
    lines = [
        "# Part 03 Canonical Skeleton",
        "",
        "Использовать дословные заголовки VERSION 18. Статус блока и вывод `найдено/не выявлено` выносить в следующий абзац, а не в строку заголовка.",
        "",
    ]
    for roman, title in PART_03_CANONICAL_BLOCKS.items():
        lines.extend(
            [
                f"{roman}. {title}",
                "Статус: [заполнить]",
                "Документ(ы) / FAIL-SAFE: [заполнить]",
                "",
            ]
        )
    return "\n".join(lines)


def build_part_04_canonical_skeleton() -> str:
    lines = [
        "# Part 04 Canonical Skeleton",
        "",
        "Использовать исходную нумерацию подпунктов 1–18. Для каждого подпункта: короткий статус, затем документы или FAIL-SAFE.",
        "",
    ]
    for idx, label in enumerate(PART_04_CONTROL_POINTS, start=1):
        lines.extend(
            [
                f"{idx}. {label}.",
                "Статус: [НАЙДЕНО/НЕ ВЫЯВЛЕНО]",
                "Документ(ы) / FAIL-SAFE: [заполнить]",
                "",
            ]
        )
    return "\n".join(lines)


def build_part_05_canonical_skeleton() -> str:
    lines = [
        "# Part 05 Canonical Skeleton",
        "",
        "Использовать слойность 1–6. По каждому слою: статус, затем документы или FAIL-SAFE.",
        "",
    ]
    for idx, label in enumerate(PART_05_LAYER_POINTS, start=1):
        lines.extend(
            [
                f"{idx}. {label}.",
                "Статус: [НАЙДЕНО/НЕ ВЫЯВЛЕНО]",
                "Документ(ы) / FAIL-SAFE: [заполнить]",
                "",
            ]
        )
    return "\n".join(lines)


def build_part_10_canonical_skeleton() -> str:
    lines = [
        "# Part 10 Canonical Skeleton",
        "",
        "Мини-конспект: строго 10-15 нумерованных пунктов. После каждого пункта обязателен ссылочный блок.",
        "",
    ]
    for idx in range(1, 11):
        lines.extend(
            [
                f"{idx}. [Смысловой пункт мини-конспекта]",
                "",
                "URL1: [заполнить]",
                "URL2: [заполнить]",
                "",
            ]
        )
    return "\n".join(lines)


def build_part_11_canonical_skeleton() -> str:
    lines = [
        "# Part 11 Canonical Skeleton",
        "",
        "Только перечень практических действий стажера для 4-й колонки дневника. Без документов и ссылок.",
        "",
    ]
    for idx in range(1, 21):
        lines.append(f"{idx}. [Практическое действие по характеру задания]")
    lines.append("")
    return "\n".join(lines)


def build_part_03_readme(run_workspace: SubtopicRunWorkspace) -> str:
    paths = build_part_03_plan_paths(run_workspace)
    lines = [
        f"# Part 03 Range Plan: {run_workspace.subtopic_entry.line}",
        "",
        "Этот каталог готовит сегментированное исполнение Части 3 по диапазонам блоков I–XXXVII.",
        "",
        f"- Operator sequence: `{paths['operator_sequence']}`",
        f"- Canonical skeleton: `{paths['canonical_skeleton']}`",
        f"- Capture status: `{paths['capture_status']}`",
        f"- Aggregated output target: `{run_workspace.stage_outputs_dir / 'part-03.md'}`",
        "",
        "Часть 3 работает диапазонами: один ответ = один явно заданный диапазон блоков.",
        "После каждого диапазона ответ нужно сохранить отдельно, затем агент соберет их в единый `part-03.md`.",
        "",
    ]
    for segment in PART_03_SEGMENTS:
        message_path = paths["message_files"][segment["segment_id"]]
        lines.append(f"- Диапазон {segment['segment_id']}: `{segment['label']}` -> `{message_path}`")
    lines.append("")
    return "\n".join(lines)


def build_part_03_operator_sequence(run_workspace: SubtopicRunWorkspace) -> str:
    paths = build_part_03_plan_paths(run_workspace)
    capture_file = Path("C:/Users/koper/OneDrive/Documents/New project/part-03-capture.md")
    lines = [
        f"# Operator Sequence Part 03: {run_workspace.subtopic_entry.line}",
        "",
        "Использовать ту же LLM-сессию, где уже была выполнена Часть 2 по этой подтеме.",
        "",
        "## Что отправлять",
        "",
        f"1. Сообщение 1: `{paths['message_files'][1]}`",
        "2. После ответа LLM сохранить его как диапазон 1.",
        f"3. Сообщение 2: `{paths['message_files'][2]}`",
        "4. Повторять цикл до диапазона 8.",
        "",
        "## Как сохранять каждый диапазон",
        "",
        f"1. Вставить полный ответ LLM в `{capture_file}`.",
        "2. Выполнить команду для соответствующего диапазона:",
        "",
        "```powershell",
        f"python .\\notary_agent.py capture-part-03-range {run_workspace.subtopic_entry.item_id} <номер_диапазона> --source-file \"{capture_file}\"",
        "```",
        "",
        "## Ограничения",
        "",
        "- Не начинать заново с `ТЕМА:` или `АНАЛИЗ ОБЛАСТИ ПРАВА`.",
        "- Не добавлять `A. РЕГУЛЯТОРНОЕ ЯДРО` и `B. ОПОРНЫЕ ДОКУМЕНТЫ`.",
        "- Для каждого найденного документа давать собственную карточку с полным наименованием, структурным элементом и ссылками; не сваливать документы в одну строку под подпунктом.",
        "- Один ответ = один диапазон блоков. Самостоятельный переход к следующему диапазону запрещен.",
        "",
    ]
    return "\n".join(lines)


def build_part_03_message(run_workspace: SubtopicRunWorkspace, segment: dict[str, Any]) -> str:
    part_03_input = next(part.content.strip() for part in run_workspace.parts if part.number == 3)
    literal_context_md = build_literal_context_bundle(run_workspace, 3)
    reasoning_brief_md = build_reasoning_part_brief(run_workspace, 3)
    segment_start, segment_end = parse_roman_range_label(segment["label"])
    canonical_lines = []
    canonical_items = list(PART_03_CANONICAL_BLOCKS.items())[segment_start - 1 : segment_end]
    for roman, title in canonical_items:
        canonical_lines.append(f"- `{roman}. {title}`")
    lines = [
        f"# Part 03 Range {segment['segment_id']}: {segment['label']}",
        "",
        "Режим: ТОЛЬКО ИСПОЛНЕНИЕ. Не начинать с `ТЕМА/АНАЛИЗ`, не добавлять `A/B`, не перезапускать формат Части 2.",
        "",
        "Требования диапазона:",
        f"- Обработать только блоки `{segment['label']}`.",
        "- Использовать исходную нумерацию и формулировки блоков I–XXXVII.",
        "- После каждого блока дать найдено/не выявлено и документы строго после соответствующего блока.",
        "- Для каждого найденного документа указывать конкретные документы и ссылки рядом с ними; не сваливать ссылки общим блоком без привязки к названию акта.",
        "",
        "Канонические блоки этого диапазона:",
        "",
    ]
    lines.extend(canonical_lines)
    lines.extend(["", "BLOCKER: любое переименование или подмена этих заголовков = ошибка.", ""])
    lines.extend([reasoning_brief_md, ""])
    if literal_context_md:
        lines.extend([literal_context_md, ""])
    lines.extend(
        [
            "## Базовая инструкция Части 3",
            "",
            part_03_input,
            "",
        ]
    )
    lines.extend(
        [
            "## Активный запрос",
            "",
            "```text",
            segment["request_text"],
            "```",
            "",
            "## Что вернуть",
            "",
            "Вернуть только ответ по этому диапазону Части 3 без перезапуска Части 2 и без перехода к следующему диапазону.",
            "",
        ]
    )
    return "\n".join(lines)


def write_part_03_plan(run_workspace: SubtopicRunWorkspace, overwrite: bool) -> dict[str, Any]:
    paths = build_part_03_plan_paths(run_workspace)
    paths["plan_dir"].mkdir(parents=True, exist_ok=True)
    write_text_if_needed(paths["readme"], build_part_03_readme(run_workspace), overwrite)
    write_text_if_needed(paths["operator_sequence"], build_part_03_operator_sequence(run_workspace), overwrite)
    write_text_if_needed(paths["canonical_skeleton"], build_part_03_canonical_skeleton(), overwrite)
    for segment in PART_03_SEGMENTS:
        write_text_if_needed(
            paths["message_files"][segment["segment_id"]],
            build_part_03_message(run_workspace, segment),
            overwrite,
        )
    capture_status = {
        "generated_at": utc_now_iso(),
        "part_number": 3,
        "captured_segment_ids": [],
        "remaining_segment_ids": [segment["segment_id"] for segment in PART_03_SEGMENTS],
        "aggregated_output_file": str(run_workspace.stage_outputs_dir / "part-03.md"),
    }
    if overwrite or not paths["capture_status"].exists():
        write_json(paths["capture_status"], capture_status)
    return paths


def update_run_manifest_part_03_plan(run_workspace: SubtopicRunWorkspace, plan_paths: dict[str, Any]) -> None:
    manifest_path = run_workspace.run_dir / "manifest.json"
    if not manifest_path.exists():
        return
    manifest = json.loads(read_text(manifest_path))
    manifest["part_03_plan"] = {
        "generated_at": utc_now_iso(),
        "plan_dir": str(plan_paths["plan_dir"]),
        "operator_sequence": str(plan_paths["operator_sequence"]),
        "capture_status": str(plan_paths["capture_status"]),
        "message_files": {
            str(segment_id): str(path)
            for segment_id, path in plan_paths["message_files"].items()
        },
    }
    write_json(manifest_path, manifest)


def build_subtopic_run_readme(run_workspace: SubtopicRunWorkspace) -> str:
    subtopic_line = run_workspace.subtopic_entry.line
    lines = [
        f"# Run Workspace: {subtopic_line}",
        "",
        f"- Основная тема: `{run_workspace.theme_workspace.theme.full_title}`",
        f"- Подтема: `{subtopic_line}`",
        f"- Копия приказа: `{run_workspace.order_path}`",
        f"- Execution packet: `{run_workspace.packet_path}`",
        f"- Stage inputs: `{run_workspace.stage_inputs_dir}`",
        f"- Stage outputs: `{run_workspace.stage_outputs_dir}`",
        f"- Локальный final-dir: `{run_workspace.final_dir}`",
        f"- Web plan dir: `{run_workspace.web_plan_dir}`",
        f"- Канонический final md target: `{run_workspace.final_md_target}`",
        f"- Канонический final docx target: `{run_workspace.final_docx_target}`",
        "",
        "## Ритм исполнения",
        "",
        "- Часть 1: подтверждение правил и остановка на `GO/СТАРТ`.",
        "- Часть 2: основной ответ `ТЕМА -> АНАЛИЗ ОБЛАСТИ ПРАВА -> A -> B -> КАРАНТИН -> FAIL-SAFE CHECK`.",
        "- Части 3-9: доборы, закрытие пропусков, anti-repeat и дельта-аудит.",
        "- Часть 10: расширенный мини-конспект.",
        "- Часть 11: перечень практических действий только для 4-й колонки дневника.",
        "",
        "## Web-first слой Части 2",
        "",
        "- `04-web-plan` хранит research pack, поисковые запросы, source cascade, research log и evidence.",
        "- `04-web-plan/reasoning-layer` хранит reasoning-briefs по Частям 2–11 и усиливает интеллектуальный режим всего цикла.",
        "- Этот слой нужен, чтобы агент или LLM не придумывали поисковые строки заново перед каждым прогоном.",
        "- `04-web-plan/follow-up-parts` хранит буквальные continuation-пакеты для Частей 6–11 с уже накопленным контекстом предыдущих частей.",
        "- `04-web-plan/part-03` хранит сегментированный операторский план для Части 3 по диапазонам I–XXXVII.",
        "",
    ]
    return "\n".join(lines)


def build_final_output_contract(run_workspace: SubtopicRunWorkspace) -> dict[str, Any]:
    return {
        "theme_title": run_workspace.theme_workspace.theme.full_title,
        "subtopic_id": run_workspace.subtopic_entry.item_id,
        "subtopic_line": run_workspace.subtopic_entry.line,
        "assembly_policy": "etalon_publication_format_with_document_cards",
        "accepted_example_md": str(run_workspace.theme_workspace.paths["output_example_md"]),
        "accepted_example_docx": str(run_workspace.theme_workspace.paths["output_example_docx"]),
        "canonical_final_md_target": str(run_workspace.final_md_target),
        "canonical_final_docx_target": str(run_workspace.final_docx_target),
        "assembly_order": [
            {
                "slot": "part_02_core",
                "source_part": 2,
                "description": "ТЕМА, АНАЛИЗ ОБЛАСТИ ПРАВА, АНАЛИЗ-КОДЕКСЫ И БАЗОВЫЕ АКТЫ, A, B, КАРАНТИН, FAIL-SAFE CHECK",
            },
            {
                "slot": "part_03_append",
                "source_part": 3,
                "description": "дополнительный алгоритм охвата, первый блок",
            },
            {
                "slot": "part_04_append",
                "source_part": 4,
                "description": "дополнительный алгоритм охвата, продолжение",
            },
            {
                "slot": "part_05_append",
                "source_part": 5,
                "description": "слои 1-6",
            },
            {
                "slot": "part_06_append",
                "source_part": 6,
                "description": "три фильтра и проверка пропусков",
            },
            {
                "slot": "part_07_append",
                "source_part": 7,
                "description": "новые документы без повторов",
            },
            {
                "slot": "part_08_append",
                "source_part": 8,
                "description": "федеральный уровень, delta check",
            },
            {
                "slot": "part_09_append",
                "source_part": 9,
                "description": "дельта-аудит",
            },
            {
                "slot": "part_10_append",
                "source_part": 10,
                "description": "мини-конспект",
            },
            {
                "slot": "part_11_append",
                "source_part": 11,
                "description": "перечень практических заданий для 4-й колонки дневника",
            },
        ],
    }


def build_final_skeleton_markdown(run_workspace: SubtopicRunWorkspace) -> str:
    subtopic_line = run_workspace.subtopic_entry.line
    lines = [
        f"# {subtopic_line}",
        "",
        f"## ЧАСТЬ 2. {FINAL_PART_TITLES[2]}",
        "",
        "ТЕМА:",
        f"{subtopic_line}",
        "",
        "АНАЛИЗ ОБЛАСТИ ПРАВА",
        "[Заполняется из результата по Части 2.]",
        "",
        "АНАЛИЗ-КОДЕКСЫ И БАЗОВЫЕ АКТЫ",
        "[Заполняется из результата по Части 2.]",
        "",
        "A. РЕГУЛЯТОРНОЕ ЯДРО",
        "[Заполняется из результата по Части 2.]",
        "",
        "B. ОПОРНЫЕ ДОКУМЕНТЫ",
        "[Заполняется из результата по Части 2.]",
        "",
        "КАРАНТИН",
        "[Заполняется из результата по Части 2.]",
        "",
        "FAIL-SAFE CHECK",
        "[Заполняется из результата по Части 2.]",
        "",
        f"## ЧАСТЬ 3. {FINAL_PART_TITLES[3]}",
        "[Сюда встраивается результат по Части 3 без переформатирования.]",
        "",
        f"## ЧАСТЬ 4. {FINAL_PART_TITLES[4]}",
        "[Сюда встраивается результат по Части 4 без переформатирования.]",
        "",
        f"## ЧАСТЬ 5. {FINAL_PART_TITLES[5]}",
        "[Сюда встраивается результат по Части 5 без переформатирования.]",
        "",
        f"## ЧАСТЬ 6. {FINAL_PART_TITLES[6]}",
        "[Сюда встраивается результат по Части 6 без переформатирования.]",
        "",
        f"## ЧАСТЬ 7. {FINAL_PART_TITLES[7]}",
        "[Сюда встраивается результат по Части 7 без переформатирования.]",
        "",
        f"## ЧАСТЬ 8. {FINAL_PART_TITLES[8]}",
        "[Сюда встраивается результат по Части 8 без переформатирования.]",
        "",
        f"## ЧАСТЬ 9. {FINAL_PART_TITLES[9]}",
        "[Сюда встраивается результат по Части 9 без переформатирования.]",
        "",
        f"## ЧАСТЬ 10. {FINAL_PART_TITLES[10]}",
        "[Сюда встраивается результат по Части 10.]",
        "",
        f"## ЧАСТЬ 11. {FINAL_PART_TITLES[11]}",
        "[Сюда встраивается результат по Части 11.]",
        "",
    ]
    return "\n".join(lines)


def build_final_contract_readme(run_workspace: SubtopicRunWorkspace) -> str:
    lines = [
        f"# Final Output Contract: {run_workspace.subtopic_entry.line}",
        "",
        "Этот каталог фиксирует структуру будущего финального `.md/.docx` по эталонному примеру.",
        "",
        f"- Эталон `.md`: `{run_workspace.theme_workspace.paths['output_example_md']}`",
        f"- Эталон `.docx`: `{run_workspace.theme_workspace.paths['output_example_docx']}`",
        f"- Канонический final md target: `{run_workspace.final_md_target}`",
        f"- Канонический final docx target: `{run_workspace.final_docx_target}`",
        "",
        "Главное правило сборки: ядро берется из Части 2, затем результаты Частей 3-11 добавляются в хронологическом порядке без потери структуры и без самовольного рефакторинга формата.",
        "Итог должен быть максимально близок к прямому результату staged-взаимодействия с LLM; допустима только минимальная постобработка.",
        "",
    ]
    return "\n".join(lines)


def build_subtopic_run_workspace_for_dir(
    theme_workspace: MainThemeWorkspace,
    subtopic_id: str,
    run_dir: Path,
    render_orders: bool = True,
    artifact_profile: str = DEFAULT_ARTIFACT_PROFILE,
    packet_mode: str | None = None,
    packet_mode_reason: str | None = None,
) -> SubtopicRunWorkspace:
    artifact_profile = normalize_artifact_profile(artifact_profile)
    resolved_packet_mode = normalize_packet_mode(packet_mode or theme_workspace.packet_mode)
    resolved_packet_mode_reason = normalize_packet_mode_reason(
        packet_mode_reason if packet_mode_reason is not None else theme_workspace.packet_mode_reason
    )
    if render_orders:
        theme_workspace.packet_mode = resolved_packet_mode
        theme_workspace.packet_mode_reason = resolved_packet_mode_reason
        render_generated_orders(theme_workspace)
    subtopic_entry = find_outline_entry(theme_workspace.outline_entries, subtopic_id)

    order_path = theme_workspace.orders_dir / f"{subtopic_id}.md"
    packet_path = theme_workspace.packets_dir / f"{subtopic_id}.md"
    if not order_path.exists() or not packet_path.exists():
        raise FileNotFoundError(
            f"Rendered order or execution packet is missing for subtopic {subtopic_id}"
        )

    order_text = read_text(order_path)
    packet_text = read_text(packet_path)
    intro_block, parts = split_rendered_order(order_text)

    run_root = theme_workspace.theme_folder / "05-runs" / subtopic_id
    context_dir = run_dir / "00-context"
    stage_inputs_dir = run_dir / "01-stage-inputs"
    stage_outputs_dir = run_dir / "02-stage-outputs"
    final_dir = run_dir / "03-final"
    web_plan_dir = run_dir / "04-web-plan"

    for directory in [context_dir, stage_inputs_dir, stage_outputs_dir, final_dir, web_plan_dir]:
        directory.mkdir(parents=True, exist_ok=True)

    execution_contract = load_existing_run_execution_contract(run_dir) if (run_dir / "manifest.json").exists() else {
        "run_mode": RUN_MODE_FRESH,
        "allowed_parts": (),
        "strict_lean_mode": True,
        "orchestration_target_percent": STRICT_ORCHESTRATION_TARGET_PERCENT,
    }

    return SubtopicRunWorkspace(
        theme_workspace=theme_workspace,
        subtopic_entry=subtopic_entry,
        run_root=run_root,
        run_dir=run_dir,
        context_dir=context_dir,
        stage_inputs_dir=stage_inputs_dir,
        stage_outputs_dir=stage_outputs_dir,
        final_dir=final_dir,
        web_plan_dir=web_plan_dir,
        order_path=order_path,
        packet_path=packet_path,
        final_md_target=theme_workspace.final_md_dir / build_published_output_filename(subtopic_entry, ".md"),
        final_docx_target=theme_workspace.final_docx_dir / build_published_output_filename(subtopic_entry, ".docx"),
        order_text=order_text,
        packet_text=packet_text,
        intro_block=intro_block,
        parts=parts,
        artifact_profile=artifact_profile,
        packet_mode=resolved_packet_mode,
        packet_mode_reason=resolved_packet_mode_reason,
        run_mode=execution_contract["run_mode"],
        allowed_parts=execution_contract["allowed_parts"],
        strict_lean_mode=execution_contract["strict_lean_mode"],
        orchestration_target_percent=execution_contract["orchestration_target_percent"],
    )


def prepare_subtopic_run_workspace(
    workspace_root: Path,
    subtopic_id: str,
    theme_query: str | None = None,
    artifact_profile: str = DEFAULT_ARTIFACT_PROFILE,
    packet_mode: str = DEFAULT_PACKET_MODE,
    packet_mode_reason: str = "",
) -> SubtopicRunWorkspace:
    artifact_profile = normalize_artifact_profile(artifact_profile)
    packet_mode = normalize_packet_mode(packet_mode)
    packet_mode_reason = normalize_packet_mode_reason(packet_mode_reason)
    resolved_theme_query = theme_query or subtopic_id.split(".", 1)[0]
    theme_workspace = prepare_main_theme_workspace(
        workspace_root,
        resolved_theme_query,
        packet_mode=packet_mode,
        packet_mode_reason=packet_mode_reason,
    )
    write_main_theme_context_files(theme_workspace)
    write_outline_phase_files(theme_workspace)

    if not theme_workspace.override_path:
        raise RuntimeError(
            f"Theme {theme_workspace.theme.full_title} has no approved outline. "
            "Run draft-main-theme-outline and approve the outline first."
        )

    run_root = theme_workspace.theme_folder / "05-runs" / subtopic_id
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = run_root / timestamp
    return build_subtopic_run_workspace_for_dir(
        theme_workspace,
        subtopic_id,
        run_dir,
        render_orders=True,
        artifact_profile=artifact_profile,
    )


def ensure_subtopic_run_workspace(
    workspace_root: Path,
    subtopic_id: str,
    theme_query: str | None = None,
    packet_mode: str = DEFAULT_PACKET_MODE,
    create_if_missing: bool = True,
) -> SubtopicRunWorkspace:
    resolved_theme_query = theme_query or subtopic_id.split(".", 1)[0]
    packet_mode = normalize_packet_mode(packet_mode)
    theme_workspace = prepare_main_theme_workspace(workspace_root, resolved_theme_query, packet_mode=packet_mode)
    write_main_theme_context_files(theme_workspace)
    write_outline_phase_files(theme_workspace)

    if not theme_workspace.override_path:
        raise RuntimeError(
            f"Theme {theme_workspace.theme.full_title} has no approved outline. "
            "Run draft-main-theme-outline and approve the outline first."
        )

    run_root = theme_workspace.theme_folder / "05-runs" / subtopic_id
    existing_runs = [path for path in run_root.iterdir() if path.is_dir()] if run_root.exists() else []
    if not existing_runs:
        if not create_if_missing:
            raise RuntimeError(
                f"No existing run found for subtopic {subtopic_id}. Surgical/partial operations require a warm run."
            )
        run_workspace = prepare_subtopic_run_workspace(
            workspace_root=workspace_root,
            subtopic_id=subtopic_id,
            theme_query=resolved_theme_query,
        )
        write_subtopic_run_files(run_workspace)
        return run_workspace

    latest_run = max(existing_runs, key=lambda path: path.name)
    existing_packet_mode = load_existing_run_packet_mode(latest_run)
    existing_packet_mode_reason = load_existing_run_packet_mode_reason(latest_run)
    return build_subtopic_run_workspace_for_dir(
        theme_workspace,
        subtopic_id,
        latest_run,
        render_orders=False,
        artifact_profile=load_existing_run_artifact_profile(latest_run),
        packet_mode=existing_packet_mode,
        packet_mode_reason=existing_packet_mode_reason,
    )


def is_response_stub(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return True
    return stripped.startswith("# Response Stub:") or "Сюда сохраняется ответ по соответствующей Части приказа." in stripped


def strip_fenced_code_blocks(text: str) -> str:
    return re.sub(r"```.*?```", "", text, flags=re.DOTALL)


def find_link_like_tokens_outside_code_blocks(text: str) -> list[str]:
    outside = strip_fenced_code_blocks(text)
    matches = re.findall(r"(https?://\S+|www\.\S+)", outside, flags=re.IGNORECASE)
    normalized: list[str] = []
    for match in matches:
        cleaned = match.strip().rstrip(".,);]")
        if cleaned not in normalized:
            normalized.append(cleaned)
    return normalized


def normalize_part_output(part_number: int, text: str) -> str:
    normalized = text
    if part_number >= 2:
        normalized = normalize_label_only_code_blocks(normalized)
    return normalized


def normalize_part_02_url_blocks(text: str) -> str:
    lines = text.splitlines()
    result: list[str] = []
    i = 0
    in_code_block = False
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if stripped.startswith("```"):
            in_code_block = not in_code_block
            result.append(line)
            i += 1
            continue
        if not in_code_block and stripped == "URL1:":
            url1_line = lines[i + 1] if i + 1 < len(lines) else ""
            j = i + 2
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines) and lines[j].strip() == "URL2:":
                url2_line = lines[j + 1] if j + 1 < len(lines) else ""
                result.extend(
                    [
                        "```text",
                        "URL1:",
                        url1_line.strip(),
                        "",
                        "URL2:",
                        url2_line.strip(),
                        "```",
                    ]
                )
                i = j + 2
                continue
        result.append(line)
        i += 1
    normalized = "\n".join(result)
    if text.endswith("\n"):
        normalized += "\n"
    return normalized


def line_has_link_token(text: str) -> bool:
    return bool(re.search(r"(https?://|www\.)\S+", text))


def is_link_group_label(text: str) -> bool:
    stripped = text.strip()
    return any(
        stripped.startswith(prefix)
        for prefix in [
            "URL1:",
            "URL2:",
            "Каноническая строка поиска:",
            "ПОИСК URL2:",
            "Проверенные страницы:",
        ]
    )


def is_loose_link_line(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return False
    if line_has_link_token(stripped):
        return True
    if is_link_group_label(stripped):
        return True
    if re.match(r"^\d+\)\s*(https?://|www\.)\S+", stripped):
        return True
    return False


def normalize_loose_link_groups(text: str) -> str:
    lines = text.splitlines()
    result: list[str] = []
    i = 0
    in_code_block = False
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if stripped.startswith("```"):
            in_code_block = not in_code_block
            result.append(line)
            i += 1
            continue
        if not in_code_block and is_loose_link_line(line):
            block: list[str] = []
            while i < len(lines):
                current = lines[i]
                current_stripped = current.strip()
                if current_stripped.startswith("```"):
                    break
                if not current_stripped:
                    if block:
                        next_nonempty = ""
                        for j in range(i + 1, len(lines)):
                            if lines[j].strip():
                                next_nonempty = lines[j].strip()
                                break
                        if not next_nonempty or not is_loose_link_line(next_nonempty):
                            break
                    i += 1
                    continue
                if not is_loose_link_line(current):
                    break
                block.append(current_stripped)
                i += 1
            if block:
                result.extend(["```text", *block, "```"])
                continue
        result.append(line)
        i += 1
    normalized = "\n".join(result)
    if text.endswith("\n"):
        normalized += "\n"
    return normalized


def extract_heading_section(text: str, start_heading: str, end_headings: list[str]) -> str:
    lines = text.splitlines()
    collecting = False
    collected: list[str] = []
    for raw_line in lines:
        stripped = raw_line.strip()
        if not collecting:
            if stripped == start_heading:
                collecting = True
            continue
        if stripped in end_headings:
            break
        collected.append(raw_line)
    return "\n".join(collected).strip()


def has_structural_element_marker(text: str) -> bool:
    return bool(re.search(r"(?im)^.*Структурн(?:ый элемент|ые элементы):", text))


def has_document_card_markers(text: str) -> bool:
    return "Вид документа:" in text and "Полное наименование:" in text and has_structural_element_marker(text)


LIKELY_LEGAL_REFERENCE_RE = re.compile(
    r"(?i)(федеральн(?:ый|ого)\s+закон|приказ|кодекс|основы законодательства|конвенц|постановлен(?:ие|ия)\s+пленума|консульский устав|методическ(?:ие|их)\s+рекомендац)"
)


def extract_part_31_subpoints(master_prompt_text: str) -> dict[int, str]:
    subpoints: dict[int, str] = {}
    in_section = False
    for raw_line in master_prompt_text.splitlines():
        line = raw_line.strip()
        if not in_section:
            if "3.1. ДОПОЛНИТЕЛЬНЫЙ АЛГОРИТМ ОХВАТА" in line:
                in_section = True
            continue
        if not line:
            continue
        if line.startswith("По каждому контрольному подпункту"):
            break
        match = re.match(r"^(?P<number>\d+)\\?\.\s*(?P<title>.+?)\s*,?$", line)
        if match:
            subpoints[int(match.group("number"))] = match.group("title").strip().rstrip(",")
    return subpoints


def parse_numeric_range_label(label: str) -> tuple[int, int]:
    match = re.match(r"^\s*(\d+)\D+(\d+)\s*$", label)
    if not match:
        raise ValueError(f"Unable to parse numeric range label: {label}")
    return int(match.group(1)), int(match.group(2))


def parse_roman_range_label(label: str) -> tuple[int, int]:
    keys = list(PART_03_CANONICAL_BLOCKS.keys())
    match = re.match(r"^\s*([IVXLCDM]+)\D+([IVXLCDM]+)\s*$", label)
    if not match:
        raise ValueError(f"Unable to parse roman range label: {label}")
    start_roman = match.group(1)
    end_roman = match.group(2)
    if start_roman not in keys or end_roman not in keys:
        raise ValueError(f"Unknown roman range label: {label}")
    return keys.index(start_roman) + 1, keys.index(end_roman) + 1


def find_foreign_subtopic_ids(text: str, expected_subtopic_id: str) -> list[str]:
    patterns = [
        re.compile(r"(?im)^\s*ТЕМА:\s*(?P<id>\d{1,2}\.\d{1,2}\.\d{1,2})\b"),
        re.compile(r"(?im)^\s*(?P<id>\d{1,2}\.\d{1,2}\.\d{1,2})\.\s+"),
        re.compile(r"(?i)подтем[аы]\s*[«\"]?(?P<id>\d{1,2}\.\d{1,2}\.\d{1,2})\b"),
        re.compile(r"(?i)тема\s*[«\"]?(?P<id>\d{1,2}\.\d{1,2}\.\d{1,2})\b"),
    ]
    candidates: set[str] = set()
    for pattern in patterns:
        for match in pattern.finditer(text):
            candidates.add(match.group("id"))
    return sorted(candidate for candidate in candidates if candidate != expected_subtopic_id)


def validate_part_output(run_workspace: SubtopicRunWorkspace, part_number: int, text: str) -> list[str]:
    issues: list[str] = []
    stripped = text.strip()
    if not stripped:
        return ["output is empty"]

    if part_number == 1:
        if "ЖДУ СИГНАЛ GO" not in stripped:
            issues.append("Part 1 must end with `ЖДУ СИГНАЛ GO`")
        if "I. Подтверждаю понимание Пунктов 1–8" not in stripped:
            issues.append("Part 1 must contain the rules-confirmation opening block")
        return issues

    if part_number == 2:
        first_nonempty = next((line.strip() for line in stripped.splitlines() if line.strip()), "")
        normalized_first = first_nonempty.replace("**", "")
        if not normalized_first.startswith("ТЕМА:"):
            issues.append("Part 2 must start with `ТЕМА:`")
        for marker in PART_02_REQUIRED_MARKERS:
            if marker not in stripped:
                issues.append(f"Part 2 is missing required marker: {marker}")
        if "ЖДУ СИГНАЛ GO" in stripped:
            issues.append("Part 2 must not contain the Part 1 stop line")
        if PART_02_FORBIDDEN_PUBLIC_BLOCKS_RE.search(stripped):
            issues.append("Part 2 must not publish block-by-block I-XXXVII coverage statuses; that belongs to Part 3")
        analysis_codes_block = extract_heading_section(
            stripped,
            "АНАЛИЗ-КОДЕКСЫ И БАЗОВЫЕ АКТЫ",
            ["A. РЕГУЛЯТОРНОЕ ЯДРО"],
        )
        if analysis_codes_block:
            if "URL1:" not in analysis_codes_block:
                issues.append("Part 2 analysis/codes block must include URL1 lines for the named acts")
            if not has_structural_element_marker(analysis_codes_block):
                issues.append("Part 2 analysis/codes block must include structural elements for the named acts")
        quarantine_block = extract_heading_section(
            stripped,
            "КАРАНТИН",
            ["FAIL-SAFE CHECK"],
        )
        if quarantine_block and re.search(r"(?im)^\s*\d+[.)]?\s", quarantine_block):
            if "URL1:" not in quarantine_block:
                issues.append("Part 2 quarantine entries must include URL1")
            if not has_structural_element_marker(quarantine_block):
                issues.append("Part 2 quarantine entries must include a structural element marker")
        issues.extend(validate_url2_presence_per_document_block(stripped, part_number))

    if part_number == 3:
        issues.extend(validate_part_03_canonical_structure(stripped))
        missing_blocks = find_missing_canonical_part_03_blocks(stripped)
        if missing_blocks:
            issues.append(
                "Part 3 is missing canonical VERSION 18 blocks: " + ", ".join(missing_blocks)
            )
        issues.extend(validate_part_03_applicable_blocks_have_url2(stripped))
        issues.extend(validate_url2_presence_per_document_block(stripped, part_number))

    if part_number == 4:
        status_count = stripped.count("Статус:")
        if status_count < 4:
            issues.append("Part 4 should use explicit `Статус:` markers for the covered authority/control blocks")
        url2_count = count_url2_markers(stripped)
        if status_count >= 3 and url2_count < 3:
            issues.append(
                f"Part 4 with at least 3 `Статус:` markers must contain at least 3 `URL2:` lines; found only {url2_count}"
            )
        issues.extend(validate_url2_presence_per_document_block(stripped, part_number))

    if part_number == 5:
        status_count = stripped.count("Статус:")
        if status_count < 3:
            issues.append("Part 5 should use explicit `Статус:` markers for the covered layers")
        url2_count = count_url2_markers(stripped)
        if status_count > 0 and url2_count == 0:
            issues.append("Part 5 with `Статус:` markers must contain at least one `URL2:` line")
        issues.extend(validate_url2_presence_per_document_block(stripped, part_number))

    if part_number in {6, 7, 8, 9}:
        first_nonempty = next((line.strip() for line in stripped.splitlines() if line.strip()), "")
        normalized_first = first_nonempty.replace("**", "")
        if normalized_first.startswith("ТЕМА:"):
            issues.append(f"Part {part_number} output must not restart with `ТЕМА:`")
        if "АНАЛИЗ ОБЛАСТИ ПРАВА" in stripped:
            issues.append(f"Part {part_number} output must not restart the Part 2 analysis blocks")
        if "A. РЕГУЛЯТОРНОЕ ЯДРО" in stripped or "B. ОПОРНЫЕ ДОКУМЕНТЫ" in stripped:
            issues.append(f"Part {part_number} output must not reintroduce `A/B` sections from Part 2")
        if "Прямой документ:" in stripped or "Прямые документы:" in stripped:
            issues.append(
                f"Part {part_number} output must use full document cards instead of grouped `Прямой документ(ы)` lines"
            )

        required_labels = [
            "Вид документа:",
            "Полное наименование:",
            "Структурный элемент:",
            "URL1:",
            "URL2:",
            "Статус верификации:",
        ]
        card_blocks = [
            block.strip()
            for block in re.split(r"(?m)(?=^(?:\d+\.\s+)?Вид документа:)", stripped)
            if "Вид документа:" in block
        ]
        for index, block in enumerate(card_blocks, start=1):
            for label in required_labels:
                match = re.search(rf"(?m)^(?:\d+\.\s+)?{re.escape(label)}\s*(.*)$", block)
                if not match:
                    issues.append(f"Part {part_number} card #{index} is missing required field `{label}`")
                elif not match.group(1).strip():
                    issues.append(f"Part {part_number} card #{index} has empty required field `{label}`")
        issues.extend(validate_url2_presence_per_document_block(stripped, part_number))

    if part_number == 10:
        items = extract_top_level_arabic_items(stripped)
        if not items:
            issues.append("Part 10 must contain a numbered mini-summary")
        else:
            expected = list(range(1, len(items) + 1))
            if items != expected:
                issues.append(
                    f"Part 10 must use continuous numbering starting from 1 without restarts or gaps; found: {items}"
                )
            if not (10 <= len(items) <= 15):
                issues.append(f"Part 10 must contain 10-15 numbered points; found: {len(items)}")
        issues.extend(validate_part_10_item_level_url2(stripped))

    if part_number == 11:
        items = extract_top_level_arabic_items(stripped)
        if len(items) < 15:
            issues.append(f"Part 11 should contain an expanded practical-task list; found only {len(items)} numbered items")
        if "URL1:" in stripped or "URL2:" in stripped or "VERIFIED URL2:" in stripped:
            issues.append("Part 11 must not contain document links")

    return issues


def extract_top_level_arabic_items(text: str) -> list[int]:
    items: list[int] = []
    for raw_line in text.splitlines():
        stripped = raw_line.strip()
        if stripped.startswith(("URL1:", "URL2:", "VERIFIED URL2:")):
            continue
        match = re.match(r"^(?P<number>\d+)\.\s+", stripped)
        if match:
            items.append(int(match.group("number")))
    return items



def split_document_blocks_by_full_name(text: str) -> list[str]:
    full_name_label = "Полное наименование:"
    pattern = rf"(?m)(?=^(?:\d+\.\s+)?{re.escape(full_name_label)})"
    return [
        block.strip()
        for block in re.split(pattern, text)
        if full_name_label in block
    ]


def validate_url2_presence_per_document_block(text: str, part_number: int) -> list[str]:
    issues: list[str] = []
    full_name_label = "Полное наименование:"
    for index, block in enumerate(split_document_blocks_by_full_name(text), start=1):
        if "URL2:" not in block:
            issues.append(
                f"Part {part_number} document block #{index} with `{full_name_label}` must contain `URL2:` in the same block"
            )
    return issues


def count_url2_markers(text: str) -> int:
    return len(re.findall(r"(?m)^\s*URL2:\s*.+", text))


def split_numbered_part_10_items(text: str) -> list[str]:
    pattern = re.compile(r"(?ms)^\s*\d+\.\s+.*?(?=^\s*\d+\.\s+|\Z)")
    return [match.group(0).strip() for match in pattern.finditer(text)]


def validate_part_10_item_level_url2(text: str) -> list[str]:
    issues: list[str] = []
    for index, item in enumerate(split_numbered_part_10_items(text), start=1):
        if "URL2:" not in item:
            issues.append(
                f"Part 10 item #{index} must contain `URL2:` inside the same item; link blocks only at the end of the whole summary are not allowed"
            )
    return issues

def extract_part_03_headings(text: str) -> list[tuple[str, str]]:
    headings: list[tuple[str, str]] = []
    roman_pattern = "|".join(PART_03_CANONICAL_BLOCKS.keys())
    pattern = re.compile(rf"(?m)^\s*(?P<roman>{roman_pattern})\.\s+(?P<title>.+?)(?:\s+[—-]\s+|\s*$)")
    for match in pattern.finditer(text):
        roman = match.group("roman")
        title = match.group("title").strip()
        title = re.sub(r"^\*+|\*+$", "", title).strip()
        headings.append((roman, title))
    return headings


def validate_part_03_canonical_structure(text: str) -> list[str]:
    issues: list[str] = []
    stripped = text.strip()
    for legacy_heading in PART_03_FORBIDDEN_LEGACY_HEADINGS:
        if legacy_heading in stripped:
            issues.append(
                f"Part 3 must use canonical VERSION 18 search-sphere headings, found legacy heading: {legacy_heading}"
            )

    headings = extract_part_03_headings(stripped)
    if not headings:
        issues.append("Part 3 must contain canonical I–XXXVII headings from VERSION 18")
        return issues

    for roman, title in headings:
        expected = PART_03_CANONICAL_BLOCKS.get(roman)
        if expected and title != expected:
            issues.append(
                f"Part 3 heading `{roman}. {title}` does not match canonical VERSION 18 heading `{roman}. {expected}`"
            )
    return issues


def split_part_03_blocks(text: str) -> list[tuple[str, str]]:
    roman_pattern = "|".join(PART_03_CANONICAL_BLOCKS.keys())
    pattern = re.compile(
        rf"(?ms)^\s*(?P<roman>{roman_pattern})\.\s+.*?(?=^\s*(?:{roman_pattern})\.\s+|\Z)"
    )
    blocks: list[tuple[str, str]] = []
    for match in pattern.finditer(text):
        blocks.append((match.group("roman"), match.group(0).strip()))
    return blocks


def part_03_block_is_explicitly_not_applicable(block_text: str) -> bool:
    lowered = block_text.lower()
    negative_markers = [
        "статус: не выявлено",
        "не применяется",
        "не применим",
        "не применима",
        "не применимо",
        "не относится",
        "не релевант",
    ]
    return any(marker in lowered for marker in negative_markers)


def validate_part_03_applicable_blocks_have_url2(text: str) -> list[str]:
    issues: list[str] = []
    for roman, block in split_part_03_blocks(text):
        if part_03_block_is_explicitly_not_applicable(block):
            continue
        if "URL2:" not in block:
            issues.append(
                f"Part 3 block `{roman}` is applicable to the theme and must contain at least one `URL2:` inside the same block"
            )
    return issues


def validate_part_03_segment_output(text: str) -> list[str]:
    issues: list[str] = []
    stripped = text.strip()
    if not stripped:
        return ["output is empty"]
    first_nonempty = next((line.strip() for line in stripped.splitlines() if line.strip()), "")
    normalized_first = first_nonempty.replace("**", "")
    if normalized_first.startswith("ТЕМА:"):
        issues.append("Part 3 range output must not restart with `ТЕМА:`")
    if "АНАЛИЗ ОБЛАСТИ ПРАВА" in stripped:
        issues.append("Part 3 range output must not restart the Part 2 analysis blocks")
    if "A. РЕГУЛЯТОРНОЕ ЯДРО" in stripped or "B. ОПОРНЫЕ ДОКУМЕНТЫ" in stripped:
        issues.append("Part 3 range output must not reintroduce `A/B` sections from Part 2")
    issues.extend(validate_part_03_canonical_structure(stripped))
    issues.extend(validate_part_03_applicable_blocks_have_url2(stripped))
    return issues


def get_part_03_segment(segment_id: int) -> dict[str, Any]:
    for segment in PART_03_SEGMENTS:
        if int(segment["segment_id"]) == int(segment_id):
            return segment
    raise RuntimeError(f"Unknown Part 3 segment: {segment_id}")


def get_part_03_segment_output_path(run_workspace: SubtopicRunWorkspace, segment_id: int) -> Path:
    return run_workspace.stage_outputs_dir / f"part-03.range-{segment_id:02d}.md"


def rebuild_part_03_aggregated_output(run_workspace: SubtopicRunWorkspace) -> tuple[Path, list[int], list[int]]:
    aggregated_path = run_workspace.stage_outputs_dir / "part-03.md"
    blocks: list[str] = []
    captured: list[int] = []
    for segment in PART_03_SEGMENTS:
        path = get_part_03_segment_output_path(run_workspace, segment["segment_id"])
        if not path.exists():
            continue
        content = read_text(path).strip()
        if not content or is_response_stub(content):
            continue
        captured.append(segment["segment_id"])
        blocks.append(content)
    if blocks:
        write_text(aggregated_path, "\n\n".join(blocks).rstrip() + "\n")
    elif aggregated_path.exists():
        aggregated_path.unlink()
    remaining = [segment["segment_id"] for segment in PART_03_SEGMENTS if segment["segment_id"] not in captured]
    return aggregated_path, captured, remaining


def update_part_03_capture_status(run_workspace: SubtopicRunWorkspace, captured: list[int], remaining: list[int]) -> None:
    paths = build_part_03_plan_paths(run_workspace)
    status_payload = {
        "generated_at": utc_now_iso(),
        "part_number": 3,
        "captured_segment_ids": captured,
        "remaining_segment_ids": remaining,
        "aggregated_output_file": str(run_workspace.stage_outputs_dir / "part-03.md"),
    }
    write_json(paths["capture_status"], status_payload)


def validate_part_04_segment_output(text: str) -> list[str]:
    issues: list[str] = []
    stripped = text.strip()
    if not stripped:
        return ["output is empty"]
    first_nonempty = next((line.strip() for line in stripped.splitlines() if line.strip()), "")
    normalized_first = first_nonempty.replace("**", "")
    if normalized_first.startswith("ТЕМА:"):
        issues.append("Part 4 range output must not restart with `ТЕМА:`")
    if "АНАЛИЗ ОБЛАСТИ ПРАВА" in stripped:
        issues.append("Part 4 range output must not restart the Part 2 analysis blocks")
    if "A. РЕГУЛЯТОРНОЕ ЯДРО" in stripped or "B. ОПОРНЫЕ ДОКУМЕНТЫ" in stripped:
        issues.append("Part 4 range output must not reintroduce `A/B` sections from Part 2")
    if "Прямой документ:" in stripped or "Прямые документы:" in stripped:
        issues.append("Part 4 range output must use full document cards instead of grouped `Прямой документ(ы)` lines")
    if "Статус: НАЙДЕНО" in stripped and not has_document_card_markers(stripped):
        issues.append("Part 4 range output with found documents must include `Вид документа`, `Полное наименование` and `Структурный элемент`")
    if "Статус: НАЙДЕНО" in stripped and "URL1:" not in stripped:
        issues.append("Part 4 range output with found documents must include URL1")
    return issues


def get_part_04_segment(segment_id: int) -> dict[str, Any]:
    for segment in PART_04_SEGMENTS:
        if int(segment["segment_id"]) == int(segment_id):
            return segment
    raise RuntimeError(f"Unknown Part 4 segment: {segment_id}")


def get_part_04_segment_output_path(run_workspace: SubtopicRunWorkspace, segment_id: int) -> Path:
    return run_workspace.stage_outputs_dir / f"part-04.range-{segment_id:02d}.md"


def rebuild_part_04_aggregated_output(run_workspace: SubtopicRunWorkspace) -> tuple[Path, list[int], list[int]]:
    aggregated_path = run_workspace.stage_outputs_dir / "part-04.md"
    blocks: list[str] = []
    captured: list[int] = []
    for segment in PART_04_SEGMENTS:
        path = get_part_04_segment_output_path(run_workspace, segment["segment_id"])
        if not path.exists():
            continue
        content = read_text(path).strip()
        if not content or is_response_stub(content):
            continue
        captured.append(segment["segment_id"])
        blocks.append(content)
    if blocks:
        write_text(aggregated_path, "\n\n".join(blocks).rstrip() + "\n")
    elif aggregated_path.exists():
        aggregated_path.unlink()
    remaining = [segment["segment_id"] for segment in PART_04_SEGMENTS if segment["segment_id"] not in captured]
    return aggregated_path, captured, remaining


def update_part_04_capture_status(run_workspace: SubtopicRunWorkspace, captured: list[int], remaining: list[int]) -> None:
    paths = build_part_04_plan_paths(run_workspace)
    status_payload = {
        "generated_at": utc_now_iso(),
        "part_number": 4,
        "captured_segment_ids": captured,
        "remaining_segment_ids": remaining,
        "aggregated_output_file": str(run_workspace.stage_outputs_dir / "part-04.md"),
    }
    write_json(paths["capture_status"], status_payload)


def build_part_04_plan_paths(run_workspace: SubtopicRunWorkspace) -> dict[str, Any]:
    plan_dir = run_workspace.web_plan_dir / "part-04"
    message_files = {
        segment["segment_id"]: plan_dir / f"message-{segment['segment_id']:02d}.part-04.range-{segment['segment_id']:02d}.md"
        for segment in PART_04_SEGMENTS
    }
    return {
        "plan_dir": plan_dir,
        "readme": plan_dir / "README.part-04.md",
        "operator_sequence": plan_dir / "00-operator-sequence.md",
        "canonical_skeleton": plan_dir / "part-04.canonical-skeleton.md",
        "capture_status": plan_dir / "capture-status.json",
        "message_files": message_files,
    }


def build_part_04_readme(run_workspace: SubtopicRunWorkspace) -> str:
    paths = build_part_04_plan_paths(run_workspace)
    lines = [
        f"# Part 04 Range Plan: {run_workspace.subtopic_entry.line}",
        "",
        "Этот каталог готовит сегментированное исполнение Части 4 по подпунктам 3.1 Дополнительного алгоритма охвата.",
        "",
        f"- Operator sequence: `{paths['operator_sequence']}`",
        f"- Canonical skeleton: `{paths['canonical_skeleton']}`",
        f"- Capture status: `{paths['capture_status']}`",
        f"- Aggregated output target: `{run_workspace.stage_outputs_dir / 'part-04.md'}`",
        "",
    ]
    for segment in PART_04_SEGMENTS:
        message_path = paths["message_files"][segment["segment_id"]]
        lines.append(f"- Диапазон {segment['segment_id']}: `{segment['label']}` -> `{message_path}`")
    lines.append("")
    return "\n".join(lines)


def build_part_04_operator_sequence(run_workspace: SubtopicRunWorkspace) -> str:
    paths = build_part_04_plan_paths(run_workspace)
    capture_file = Path("C:/Users/koper/OneDrive/Documents/New project/part-04-capture.md")
    lines = [
        f"# Operator Sequence Part 04: {run_workspace.subtopic_entry.line}",
        "",
        "Использовать ту же LLM-сессию после завершения Части 3 по этой подтеме.",
        "",
        "## Что отправлять",
        "",
        f"1. Сообщение 1: `{paths['message_files'][1]}`",
        "2. После ответа LLM сохранить его как диапазон 1.",
        f"3. Сообщение 2: `{paths['message_files'][2]}`",
        "4. Повторять цикл до диапазона 4.",
        "",
        "## Как сохранять каждый диапазон",
        "",
        f"1. Вставить полный ответ LLM в `{capture_file}`.",
        "2. Выполнить команду для соответствующего диапазона:",
        "",
        "```powershell",
        f"python .\\notary_agent.py capture-part-04-range {run_workspace.subtopic_entry.item_id} <номер_диапазона> --source-file \"{capture_file}\"",
        "```",
        "",
        "## Ограничения",
        "",
        "- Не начинать заново с `ТЕМА:` или `АНАЛИЗ ОБЛАСТИ ПРАВА`.",
        "- Не добавлять `A. РЕГУЛЯТОРНОЕ ЯДРО` и `B. ОПОРНЫЕ ДОКУМЕНТЫ`.",
        "- Каждый найденный документ оформлять собственной карточкой или явно привязанным мини-блоком со ссылками и структурным элементом.",
        "- Один ответ = один диапазон подпунктов. Самостоятельный переход к следующему диапазону запрещен.",
        "",
    ]
    return "\n".join(lines)


def build_part_04_message(run_workspace: SubtopicRunWorkspace, segment: dict[str, Any]) -> str:
    part_04_input = next(part.content.strip() for part in run_workspace.parts if part.number == 4)
    literal_context_md = build_literal_context_bundle(run_workspace, 4)
    reasoning_brief_md = build_reasoning_part_brief(run_workspace, 4)
    subpoints = extract_part_31_subpoints(run_workspace.theme_workspace.master_prompt_text)
    range_start, range_end = parse_numeric_range_label(str(segment["label"]))
    range_lines: list[str] = []
    for number in range(range_start, range_end + 1):
        title = subpoints.get(number)
        if title:
            range_lines.append(f"{number}. {title}")
    lines = [
        f"# Part 04 Range {segment['segment_id']}: {segment['label']}",
        "",
        "Режим: ТОЛЬКО ИСПОЛНЕНИЕ. Не начинать с `ТЕМА/АНАЛИЗ`, не добавлять `A/B`, не перезапускать формат Частей 2–3.",
        "",
        "Требования диапазона:",
        f"- Обработать только подпункты `{segment['label']}`.",
        "- Использовать исходную нумерацию и формулировки подпунктов 3.1 Дополнительного алгоритма охвата.",
        "- После каждого подпункта дать найдено/не выявлено и документы строго после соответствующего подпункта.",
        "- Если подпункт дает документы, выводить их только карточками: `Вид документа`, `Полное наименование`, `Орган`, `Структурный элемент`, `Подтверждение применимости`, `URL1`, `URL2`, `VERIFIED URL2`, `Заголовок страницы URL2`, `Сверка реквизитов`, `Каноническая строка поиска`.",
        "- Запрещено писать `Прямой документ:` или `Прямые документы:` со списком названий и общей пачкой ссылок; ссылки должны быть привязаны к каждой карточке.",
        "",
    ]
    lines.extend([reasoning_brief_md, ""])
    if literal_context_md:
        lines.extend([literal_context_md, ""])
    lines.extend(
        [
            "## Базовая инструкция Части 4",
            "",
            part_04_input,
            "",
        ]
    )
    if range_lines:
        lines.extend(
            [
                "## Дословные подпункты 3.1 VERSION 18 для этого диапазона",
                "",
                "```text",
                *range_lines,
                "```",
                "",
            ]
        )
    lines.extend(
        [
            "## Активный запрос",
            "",
            "```text",
            segment["request_text"],
            "```",
            "",
            "## Что вернуть",
            "",
            "Вернуть только ответ по этому диапазону Части 4 без перезапуска предыдущих частей и без перехода к следующему диапазону.",
            "",
        ]
    )
    return "\n".join(lines)


def write_part_04_plan(run_workspace: SubtopicRunWorkspace, overwrite: bool) -> dict[str, Any]:
    paths = build_part_04_plan_paths(run_workspace)
    paths["plan_dir"].mkdir(parents=True, exist_ok=True)
    write_text_if_needed(paths["readme"], build_part_04_readme(run_workspace), overwrite)
    write_text_if_needed(paths["operator_sequence"], build_part_04_operator_sequence(run_workspace), overwrite)
    write_text_if_needed(paths["canonical_skeleton"], build_part_04_canonical_skeleton(), overwrite)
    for segment in PART_04_SEGMENTS:
        write_text_if_needed(
            paths["message_files"][segment["segment_id"]],
            build_part_04_message(run_workspace, segment),
            overwrite,
        )
    capture_status = {
        "generated_at": utc_now_iso(),
        "part_number": 4,
        "captured_segment_ids": [],
        "remaining_segment_ids": [segment["segment_id"] for segment in PART_04_SEGMENTS],
        "aggregated_output_file": str(run_workspace.stage_outputs_dir / "part-04.md"),
    }
    if overwrite or not paths["capture_status"].exists():
        write_json(paths["capture_status"], capture_status)
    return paths


def update_run_manifest_part_04_plan(run_workspace: SubtopicRunWorkspace, plan_paths: dict[str, Any]) -> None:
    manifest_path = run_workspace.run_dir / "manifest.json"
    if not manifest_path.exists():
        return
    manifest = json.loads(read_text(manifest_path))
    manifest["part_04_plan"] = {
        "generated_at": utc_now_iso(),
        "plan_dir": str(plan_paths["plan_dir"]),
        "operator_sequence": str(plan_paths["operator_sequence"]),
        "capture_status": str(plan_paths["capture_status"]),
        "message_files": {
            str(segment_id): str(path)
            for segment_id, path in plan_paths["message_files"].items()
        },
    }
    write_json(manifest_path, manifest)


def validate_part_05_segment_output(text: str) -> list[str]:
    issues: list[str] = []
    stripped = text.strip()
    if not stripped:
        return ["output is empty"]
    first_nonempty = next((line.strip() for line in stripped.splitlines() if line.strip()), "")
    normalized_first = first_nonempty.replace("**", "")
    if normalized_first.startswith("ТЕМА:"):
        issues.append("Part 5 range output must not restart with `ТЕМА:`")
    if "АНАЛИЗ ОБЛАСТИ ПРАВА" in stripped:
        issues.append("Part 5 range output must not restart the Part 2 analysis blocks")
    if "A. РЕГУЛЯТОРНОЕ ЯДРО" in stripped or "B. ОПОРНЫЕ ДОКУМЕНТЫ" in stripped:
        issues.append("Part 5 range output must not reintroduce `A/B` sections from Part 2")
    return issues


def get_part_05_segment(segment_id: int) -> dict[str, Any]:
    for segment in PART_05_SEGMENTS:
        if int(segment["segment_id"]) == int(segment_id):
            return segment
    raise RuntimeError(f"Unknown Part 5 segment: {segment_id}")


def get_part_05_segment_output_path(run_workspace: SubtopicRunWorkspace, segment_id: int) -> Path:
    return run_workspace.stage_outputs_dir / f"part-05.range-{segment_id:02d}.md"


def rebuild_part_05_aggregated_output(run_workspace: SubtopicRunWorkspace) -> tuple[Path, list[int], list[int]]:
    aggregated_path = run_workspace.stage_outputs_dir / "part-05.md"
    blocks: list[str] = []
    captured: list[int] = []
    for segment in PART_05_SEGMENTS:
        path = get_part_05_segment_output_path(run_workspace, segment["segment_id"])
        if not path.exists():
            continue
        content = read_text(path).strip()
        if not content or is_response_stub(content):
            continue
        captured.append(segment["segment_id"])
        blocks.append(content)
    if blocks:
        write_text(aggregated_path, "\n\n".join(blocks).rstrip() + "\n")
    elif aggregated_path.exists():
        aggregated_path.unlink()
    remaining = [segment["segment_id"] for segment in PART_05_SEGMENTS if segment["segment_id"] not in captured]
    return aggregated_path, captured, remaining


def update_part_05_capture_status(run_workspace: SubtopicRunWorkspace, captured: list[int], remaining: list[int]) -> None:
    paths = build_part_05_plan_paths(run_workspace)
    status_payload = {
        "generated_at": utc_now_iso(),
        "part_number": 5,
        "captured_segment_ids": captured,
        "remaining_segment_ids": remaining,
        "aggregated_output_file": str(run_workspace.stage_outputs_dir / "part-05.md"),
    }
    write_json(paths["capture_status"], status_payload)


def build_part_05_plan_paths(run_workspace: SubtopicRunWorkspace) -> dict[str, Any]:
    plan_dir = run_workspace.web_plan_dir / "part-05"
    message_files = {
        segment["segment_id"]: plan_dir / f"message-{segment['segment_id']:02d}.part-05.range-{segment['segment_id']:02d}.md"
        for segment in PART_05_SEGMENTS
    }
    return {
        "plan_dir": plan_dir,
        "readme": plan_dir / "README.part-05.md",
        "operator_sequence": plan_dir / "00-operator-sequence.md",
        "canonical_skeleton": plan_dir / "part-05.canonical-skeleton.md",
        "capture_status": plan_dir / "capture-status.json",
        "message_files": message_files,
    }


def build_part_05_readme(run_workspace: SubtopicRunWorkspace) -> str:
    paths = build_part_05_plan_paths(run_workspace)
    lines = [
        f"# Part 05 Range Plan: {run_workspace.subtopic_entry.line}",
        "",
        "Этот каталог готовит сегментированное исполнение Части 5 по слоям 1–6.",
        "",
        f"- Operator sequence: `{paths['operator_sequence']}`",
        f"- Canonical skeleton: `{paths['canonical_skeleton']}`",
        f"- Capture status: `{paths['capture_status']}`",
        f"- Aggregated output target: `{run_workspace.stage_outputs_dir / 'part-05.md'}`",
        "",
    ]
    for segment in PART_05_SEGMENTS:
        message_path = paths["message_files"][segment["segment_id"]]
        lines.append(f"- Диапазон {segment['segment_id']}: `{segment['label']}` -> `{message_path}`")
    lines.append("")
    return "\n".join(lines)


def build_part_05_operator_sequence(run_workspace: SubtopicRunWorkspace) -> str:
    paths = build_part_05_plan_paths(run_workspace)
    capture_file = Path("C:/Users/koper/OneDrive/Documents/New project/part-05-capture.md")
    lines = [
        f"# Operator Sequence Part 05: {run_workspace.subtopic_entry.line}",
        "",
        "Использовать ту же LLM-сессию после завершения Части 4 по этой подтеме.",
        "",
        "## Что отправлять",
        "",
        f"1. Сообщение 1: `{paths['message_files'][1]}`",
        "2. После ответа LLM сохранить его как диапазон 1.",
        f"3. Сообщение 2: `{paths['message_files'][2]}`",
        "4. Повторить цикл до диапазона 3.",
        "",
        "## Как сохранять каждый диапазон",
        "",
        f"1. Вставить полный ответ LLM в `{capture_file}`.",
        "2. Выполнить команду для соответствующего диапазона:",
        "",
        "```powershell",
        f"python .\\notary_agent.py capture-part-05-range {run_workspace.subtopic_entry.item_id} <номер_диапазона> --source-file \"{capture_file}\"",
        "```",
        "",
        "## Ограничения",
        "",
        "- Не начинать заново с `ТЕМА:` или `АНАЛИЗ ОБЛАСТИ ПРАВА`.",
        "- Не добавлять `A. РЕГУЛЯТОРНОЕ ЯДРО` и `B. ОПОРНЫЕ ДОКУМЕНТЫ`.",
        "- Каждый найденный документ оформлять собственной карточкой или явно привязанным мини-блоком со ссылками и структурным элементом.",
        "- Один ответ = один диапазон слоев. Самостоятельный переход к следующему диапазону запрещен.",
        "",
    ]
    return "\n".join(lines)


def build_part_05_message(run_workspace: SubtopicRunWorkspace, segment: dict[str, Any]) -> str:
    part_05_input = next(part.content.strip() for part in run_workspace.parts if part.number == 5)
    literal_context_md = build_literal_context_bundle(run_workspace, 5)
    reasoning_brief_md = build_reasoning_part_brief(run_workspace, 5)
    lines = [
        f"# Part 05 Range {segment['segment_id']}: {segment['label']}",
        "",
        "Режим: ТОЛЬКО ИСПОЛНЕНИЕ. Не начинать с `ТЕМА/АНАЛИЗ`, не добавлять `A/B`, не перезапускать формат Частей 2–4.",
        "",
        "Требования диапазона:",
        f"- Обработать только слои `{segment['label']}`.",
        "- Сначала проверить слой свежим внешним поиском, а уже потом решать вопрос о повторе или новой роли документа.",
        "- Для каждого применимого слоя дать найдено/не выявлено и документы строго после соответствующего слоя.",
        "- Каждый найденный документ оформлять собственной карточкой или явно привязанным мини-блоком со ссылками и структурным элементом.",
        "",
    ]
    lines.extend([reasoning_brief_md, ""])
    if literal_context_md:
        lines.extend([literal_context_md, ""])
    lines.extend(
        [
            "## Базовая инструкция Части 5",
            "",
            part_05_input,
            "",
        ]
    )
    lines.extend(
        [
            "## Активный запрос",
            "",
            "```text",
            segment["request_text"],
            "```",
            "",
            "## Что вернуть",
            "",
            "Вернуть только ответ по этому диапазону Части 5 без перезапуска предыдущих частей и без перехода к следующему диапазону.",
            "",
        ]
    )
    return "\n".join(lines)


def write_part_05_plan(run_workspace: SubtopicRunWorkspace, overwrite: bool) -> dict[str, Any]:
    paths = build_part_05_plan_paths(run_workspace)
    paths["plan_dir"].mkdir(parents=True, exist_ok=True)
    write_text_if_needed(paths["readme"], build_part_05_readme(run_workspace), overwrite)
    write_text_if_needed(paths["operator_sequence"], build_part_05_operator_sequence(run_workspace), overwrite)
    write_text_if_needed(paths["canonical_skeleton"], build_part_05_canonical_skeleton(), overwrite)
    for segment in PART_05_SEGMENTS:
        write_text_if_needed(
            paths["message_files"][segment["segment_id"]],
            build_part_05_message(run_workspace, segment),
            overwrite,
        )
    capture_status = {
        "generated_at": utc_now_iso(),
        "part_number": 5,
        "captured_segment_ids": [],
        "remaining_segment_ids": [segment["segment_id"] for segment in PART_05_SEGMENTS],
        "aggregated_output_file": str(run_workspace.stage_outputs_dir / "part-05.md"),
    }
    if overwrite or not paths["capture_status"].exists():
        write_json(paths["capture_status"], capture_status)
    return paths


def update_run_manifest_part_05_plan(run_workspace: SubtopicRunWorkspace, plan_paths: dict[str, Any]) -> None:
    manifest_path = run_workspace.run_dir / "manifest.json"
    if not manifest_path.exists():
        return
    manifest = json.loads(read_text(manifest_path))
    manifest["part_05_plan"] = {
        "generated_at": utc_now_iso(),
        "plan_dir": str(plan_paths["plan_dir"]),
        "operator_sequence": str(plan_paths["operator_sequence"]),
        "capture_status": str(plan_paths["capture_status"]),
        "message_files": {
            str(segment_id): str(path)
            for segment_id, path in plan_paths["message_files"].items()
        },
    }
    write_json(manifest_path, manifest)


def update_run_manifest_part_status(
    run_workspace: SubtopicRunWorkspace,
    part_number: int,
    status: str,
    source_origin: str | None = None,
    validation_issues: list[str] | None = None,
    validation_passed: bool | None = None,
) -> None:
    manifest_path = run_workspace.run_dir / "manifest.json"
    if not manifest_path.exists():
        return
    manifest = json.loads(read_text(manifest_path))
    for part in manifest.get("parts", []):
        if int(part.get("part_number", 0)) == part_number:
            part["status"] = status
            part["updated_at"] = utc_now_iso()
            if source_origin is not None:
                part["source_origin"] = source_origin
            if validation_issues is not None:
                part["validation_issues"] = list(validation_issues)
            if validation_passed is not None:
                part["validation_passed"] = bool(validation_passed)
            break
    write_json(manifest_path, manifest)


def build_part_01_response(run_workspace: SubtopicRunWorkspace) -> str:
    subtopic_line = run_workspace.subtopic_entry.line
    lines = [
        "I. Подтверждаю понимание Пунктов 1–8 и принимаю их к исполнению по каждому пункту отдельно.",
        "",
        "**Пункт 1. Жёсткие ограничения**  ",
        "Понимаю: запрещено домысливание, обобщение, подмена точного поиска \"похожими\" актами, упоминание документов без точного наименования и официального подтверждения, опора на устаревшие редакции, сужение охвата до 1–2 \"главных\" актов, а также вывод перечня до прохождения всех применимых слоёв и блоков.",
        "",
        "**Пункт 2. Обязательная проверка для каждого документа**  ",
        "Понимаю: для каждого документа обязательны полное официальное наименование, вид, орган, дата, номер, статус действия, актуальность редакции, официальный источник и структурный элемент; читабельный второй источник и его полная верификация желательны и усиливают карточку, но сами по себе не определяют право документа на включение, если существование, применимость и роль уже подтверждены.",
        "",
        "**Пункт 3. Сфера поиска**  ",
        "Понимаю: работа начинается не с перечня актов, а с анализа области права, юридических узлов темы и применимых кодексов/базовых актов; затем обязателен полный проход по блокам I–XXXVII с внутренней маркировкой \"применимо / не применимо\", с обязательной проверкой процессуального и налогово-финансового слоя, а также подзаконного, нотариального и судебно-разъяснительного слоя.",
        "",
        "**Пункт 4. Формат вывода**  ",
        "Понимаю: вывод строго фиксированный — сначала строка `ТЕМА: ...`, затем блок анализа сфер права и кодексов со структурными элементами, затем перечень нормативных актов, разделённый на `A. РЕГУЛЯТОРНОЕ ЯДРО`, `B. ОПОРНЫЕ ДОКУМЕНТЫ`, `КАРАНТИН` и `FAIL-SAFE CHECK`.",
        "",
        "**Пункт 5. Fail-safe**  ",
        "Понимаю: при отсутствии действующих прямо применимых актов это должно быть прямо зафиксировано; карточка документа обязательна; при сомнении в существовании, действии, актуальности или структурном элементе применяется карантин по конкретному документу; перед финалом обязателен блок проверки слепых зон и правило остановки при критически непроверяемом факторе.",
        "",
        "**Пункт 6. Контроль галлюцинаций**  ",
        "Понимаю: перед финалом обязателен anti-omit по слоям источников, включая кодифицированные акты, специальное нотариальное регулирование, процессуальный слой, подзаконные акты, акты нотариального сообщества и применимые разъяснения высшей судебной инстанции; если слой применим, но подтверждённый документ не найден, это отдельно фиксируется, а не замалчивается.",
        "",
        "**Пункт 7. Запуск задания и анти-отказ**  ",
        "Понимаю: поиск по подтеме запускается только после завершения этапов I–II и сигнала `GO/СТАРТ`; ранняя остановка из-за \"сбоя\" запрещена; обязательна дожимка проверки до результата по каждому документу; критерий исполнения — закрытый цикл проверки, в котором документ получает один из конечных статусов: подтвержден по существованию/применимости/роли и `URL1`, усилен `VERIFIED URL2`, либо помещен в `КАРАНТИН` при реальной неопределенности по самому документу.",
        "",
        "**Пункт 8. Правила ссылок**  ",
        "Понимаю: перед выдачей каждой ссылки требуется открытие страницы и сверка реквизитов; `URL1` даётся как якорь официальности даже если неудобочитаем, но в финальном ответе должен сопровождаться полным официальным наименованием документа и читабельным `URL2`, когда он подтвержден; при неподтверждении `URL2` документ не исключается автоматически, а сохраняется по своей роли с пометкой статуса верификации.",
        "",
        "II. Подтверждаю понимание логики и дословного режима исполнения.",
        "",
        "**Что запрещено**  ",
        "Запрещено начинать поиск и перечисление документов до завершения этапов I–II и сигнала `GO/СТАРТ`. Запрещено сокращать охват, заменять проверку предположением, исключать документ только из-за отсутствия второй подтверждённой ссылки, обходить слои, пропускать структурный элемент, подменять действие сообщением о сбое, а также называть акт без полного официального наименования, структурного элемента и ссылочного сопровождения.",
        "",
        "**Что разрешено**  ",
        "Разрешены только реально действующие и применимые к теме нормативные акты и официальные документы: кодексы, федеральные законы, подзаконные акты, акты нотариального сообщества, письма и методические документы при подтверждённой применимости, а также иные официально опубликованные специальные акты соответствующих органов.",
        "",
        "**Какие документы необходимо найти**  ",
        f"Нужно искать все актуальные и прямо применимые к подтеме `{subtopic_line}` нормативные акты и официальные документы по всей сфере поиска, а не только базовые кодексы и Основы законодательства о нотариате.",
        "",
        "**Что является обязательным**  ",
        "Обязательны: анализ области права; проход по всем блокам сферы поиска; проверка применимости каждого слоя; карточка по каждому документу; структурный элемент; контроль актуальности; `URL1` как якорь официальности; `URL2` после проверки, когда это удается довести; явная пометка статуса верификации; карантин при сомнении в самом документе; блок слепых зон перед финалом.",
        "",
        "**Как начинать**  ",
        "Начинать нужно строго с Пункта 3: анализ области права, юридических узлов темы и применимых кодексов/базовых актов, с указанием структурных элементов. Только после этого допускается формирование перечня документов.",
        "",
        "**Сфера поиска и алгоритм охвата**  ",
        "Алгоритм понимаю так: сначала выделяются объект регулирования, нотариальное действие, участники, порядок, сроки, платежи, ограничения, контроль, электронное взаимодействие и международный элемент; затем каждый из блоков I–XXXVII получает статус `применимо / не применимо`; по каждому применимому блоку либо находится минимум один подтверждённый документ, либо фиксируется fail-safe `не выявлено`.",
        "",
        "**Структура перечня, контроль, fail-safe, структурный элемент**  ",
        "Перечень должен быть двухслойным: `A. РЕГУЛЯТОРНОЕ ЯДРО` и `B. ОПОРНЫЕ ДОКУМЕНТЫ`. Для каждого документа обязателен структурный элемент. Если структурный элемент не установлен точно, документ не выбрасывается автоматически, а выводится в карантине с пометкой о необходимости уточнения. Если подтверждены существование, применимость, роль, реквизиты и `URL1`, но не подтверждена читабельная вторая ссылка, документ сохраняется в `A` или `B` по роли с пометкой статуса `VERIFIED URL2` и канонической строкой поиска.",
        "",
        "**Контроль галлюцинаций, карантин, слойность**  ",
        "Слойность понимаю как обязательную проверку: Слой 1 — базовые кодифицированные и профильные акты; Слой 2 — специальное нотариальное регулирование; Слой 3 — процессуальный и контрольный слой; Слой 4 — подзаконные акты федеральных органов; Слой 5 — акты нотариального сообщества; Слой 6 — применимые судебные разъяснения. Если документ существует и применим, но есть пробел только в читабельной второй ссылке, он не уходит автоматически в карантин; карантин применяется при пробеле в самом документе, его роли, актуальности или структурном элементе.",
        "",
        "**Внутренняя проверка перед финалом**  ",
        "Перед выдачей нужно проверить три вещи по каждому документу: существует ли он реально, действует ли сейчас или подлежит применению нотариусом, есть ли официальный источник. При отрицательном ответе документ не может проходить как подтверждённый найденный документ.",
        "",
        "**Правила первой и второй ссылки**  ",
        "Понимаю: `URL1` — якорь официальности, подлежит указанию даже если неудобочитаем; `URL2` — только после полного цикла проверки страницы и сверки реквизитов. При неподтверждении `URL2` документ не исключается автоматически, а сохраняется в выдаче с пометками `VERIFIED URL2`, `Заголовок страницы URL2`, `Сверка реквизитов` и канонической строкой поиска; в карантин он уходит только если остается сомнение в самом документе. Ссылочный формат в финальном ответе должен быть читабельным и привязанным к карточке каждого документа.",
        "Дополнительно понимаю: режим `URL2: отсутствует (КАРАНТИН)` описывает только статус поля `URL2` внутри карточки документа и не равен автоматическому переносу самого документа в раздел `КАРАНТИН`.",
        "",
        "**Подтверждаю применение анти-отказа и запрета на незавершённую проверку**  ",
        "Подтверждаю: правило АО3 принимаю полностью. До исчерпания доступных веб-действий не допускается остановка под предлогом технической недоступности. Проверка по каждому документу должна быть доведена до конечного статуса по результату цикла, включая прямое открытие страниц и сверку заголовков, когда это доступно.",
        "",
        "ЖДУ СИГНАЛ GO",
        "",
    ]
    return "\n".join(lines)


def sanitize_substantive_part_output(run_workspace: SubtopicRunWorkspace, part_number: int, content: str) -> str:
    focus = build_part_02_focus(run_workspace)
    if focus["subtopic_type"] != "substantive" or not focus["tariff_sibling"]:
        return content

    sanitized = content

    if part_number == 2:
        sanitized = sanitized.replace(
            "Финансовый базовый слой образуют Основы законодательства РФ о нотариате и Налоговый кодекс Российской Федерации. По статье 22.1 Основ определяется федеральный тариф по нотариальным действиям, а статья 333.38 НК РФ закрепляет специальные льготы, в том числе по свидетельствованию верности копий документов, необходимых для предоставления льгот.\n\n",
            "",
        )
        sanitized = sanitized.replace("статья 22.1 о федеральном тарифе; ", "")
        sanitized = sanitized.replace("статья 22.1; ", "")
        sanitized = sanitized.replace(" и отказ, обжалование и тарифный базис.", ", отказ и обжалование.")
        sanitized = sanitized.replace(" статья 22.1", "")
        sanitized = re.sub(
            r"\nВид документа: кодекс Российской Федерации\.\nПолное наименование: Налоговый кодекс Российской Федерации \(часть вторая\)\..*?\nНК РФ статья 333\.38 пункт 1 подпункт 11 свидетельствование верности копий документов\n",
            "\n",
            sanitized,
            flags=re.S,
        )
        sanitized = re.sub(
            r"\nРегиональный тариф по данному нотариальному действию\..*?\nсвидетельствование верности копий документов региональный тариф 2026 нотариальная палата \[субъект РФ\]\n",
            "\n",
            sanitized,
            flags=re.S,
        )

    if part_number == 3:
        sanitized = re.sub(
            r"\nXV\. ФЕДЕРАЛЬНЫЙ ТАРИФ — НАЙДЕНО.*?\nXVIII\. ",
            "\nXVIII. ",
            sanitized,
            flags=re.S,
        )
        sanitized = re.sub(
            r"\nXXXVII\. РЕГИОНАЛЬНО-ЛОКАЛЬНЫЙ СЛОЙ НОТАРИАЛЬНЫХ ПАЛАТ — НАЙДЕНО.*?(?=\nОбработка блоков XXXVI–XXXVII завершена\.|\Z)",
            "\n",
            sanitized,
            flags=re.S,
        )

    if part_number == 4:
        sanitized = sanitized.replace(
            "По подтеме 16.1.1 акты нотариальных палат субъектов выявлены, но они носят регионально-локальный характер. Подтверждены как минимум два типа материалов: публикации региональных тарифов, где отдельно указано действие «свидетельствование верности копий документов и выписок из них», и региональные публикации/размещения методических материалов по этой теме. Без указания конкретного субъекта РФ такой слой не может быть исчерпывающе индивидуализирован одним актом.",
            "По подтеме 16.1.1 на уровне нотариальных палат субъектов подтверждено размещение профильных методических материалов по рассматриваемому нотариальному действию. Самостоятельный тарифный региональный слой в этой подтеме не публикуется, поскольку вынесен в отдельный тарифный подпункт.",
        )
        sanitized = re.sub(
            r"\nВид документа: официальный информационный материал нотариальной палаты субъекта Российской Федерации\.\nПолное наименование: Размеры регионального тарифа на 2026 год\..*?\nнотариальная палата размеры регионального тарифа на 2026 год свидетельствование верности копий документов и выписок из них\n",
            "\n",
            sanitized,
            flags=re.S,
        )
        sanitized = re.sub(
            r"\nВид документа: официальный информационный материал нотариальной палаты субъекта Российской Федерации\.\nПолное наименование: Размер оплаты за услуги правового и технического характера\..*?\nнотариальная палата размер оплаты за услуги правового и технического характера свидетельствование верности копий документов и выписок из них\n",
            "\n",
            sanitized,
            flags=re.S,
        )
        sanitized = re.sub(
            r"\nВид документа: официальный материал территориального органа Минюста России\.\nПолное наименование: Льготы при обращении за совершением нотариальных действий\..*?\nМинюст льготы при обращении за совершением нотариальных действий свидетельствование верности копий документов\n",
            "\n",
            sanitized,
            flags=re.S,
        )
        sanitized = re.sub(
            r"\nМинфин — НЕ ВЫЯВЛЕНО.*?(?=\nФОИВ — )",
            "\n",
            sanitized,
            flags=re.S,
        )

    if part_number == 5:
        sanitized = sanitized.replace(
            "Основы законодательства Российской Федерации о нотариате, Гражданский процессуальный кодекс Российской Федерации, Налоговый кодекс Российской Федерации (часть вторая), Консульский устав Российской Федерации, а также международный договор о легализации иностранных официальных документов.",
            "Основы законодательства Российской Федерации о нотариате, Гражданский процессуальный кодекс Российской Федерации, Консульский устав Российской Федерации, а также международный договор о легализации иностранных официальных документов.",
        )

    if part_number == 8:
        sanitized = sanitized.replace(
            "Федеральный слой по текущей теме закрыт базовыми актами, специальными подзаконными актами Минюста России, специальными актами для органов местного самоуправления и консульских должностных лиц, международным слоем, а также ранее добавленными корректирующими и тарифно-льготными документами.",
            "Федеральный слой по текущей теме закрыт базовыми актами, специальными подзаконными актами Минюста России, специальными актами для органов местного самоуправления и консульских должностных лиц, международным слоем, а также ранее добавленными корректирующими документами.",
        )

    sanitized = re.sub(
        r"Значение для темы: это основной прямоприменимый акт, который одновременно закрепляет полномочие на совершение действия, пределы действия, специальное регулирование копии с копии, отказ, обжалование и тарифный базис\.",
        "Значение для темы: это основной прямоприменимый акт, который одновременно закрепляет полномочие на совершение действия, пределы действия, специальное регулирование копии с копии, а также отказ и обжалование.",
        sanitized,
    )
    sanitized = re.sub(
        r"Структурные элементы: статья 22\.1;\s*статья 35;\s*статья 45;\s*статья 46;\s*статья 48;\s*статья 49;\s*статья 50;\s*глава XIII;\s*статья 77;\s*статья 79\.",
        "Структурные элементы: статья 35; статья 45; статья 46; статья 48; статья 49; статья 50; глава XIII; статья 77; статья 79.",
        sanitized,
    )
    sanitized = re.sub(
        r"Подтверждение применимости: это базовый прямоприменимый акт по подтеме; статья 77 регулирует свидетельствование верности копий документов и выписок из них, статья 79 — свидетельствование верности копии с копии документа, статья 22\.1 — федеральный тариф, статьи 48–50 — отказ, обжалование и регистрация нотариального действия\.",
        "Подтверждение применимости: это базовый прямоприменимый акт по подтеме; статья 77 регулирует свидетельствование верности копий документов и выписок из них, статья 79 — свидетельствование верности копии с копии документа, а статьи 48–50 — отказ, обжалование и регистрацию нотариального действия.",
        sanitized,
    )
    sanitized = re.sub(
        r"Основы законодательства Российской Федерации о нотариате 11\.02\.1993 № 4462-1 статья 77 статья 79 статья 22\.1",
        "Основы законодательства Российской Федерации о нотариате 11.02.1993 № 4462-1 статья 77 статья 79",
        sanitized,
    )

    sanitized = re.sub(r"\n{3,}", "\n\n", sanitized).strip()
    return sanitized + "\n"


def assemble_subtopic_final(run_workspace: SubtopicRunWorkspace, publish: bool) -> Path:
    assembled_blocks: list[str] = []
    included_parts: list[int] = []
    skipped_stub_parts: list[int] = []
    missing_parts: list[int] = []

    for part in run_workspace.parts:
        if part.number == 1:
            continue
        output_path = run_workspace.stage_outputs_dir / part.filename
        if not output_path.exists():
            missing_parts.append(part.number)
            continue
        content = read_text(output_path).strip()
        if is_response_stub(content):
            skipped_stub_parts.append(part.number)
            continue
        if part.number == 3:
            issues = validate_part_03_canonical_structure(content)
            if issues:
                raise RuntimeError("Part 3 output validation failed:\n- " + "\n- ".join(issues))
        content = sanitize_substantive_part_output(run_workspace, part.number, content).strip()
        if not content:
            skipped_stub_parts.append(part.number)
            continue
        included_parts.append(part.number)
        assembled_blocks.append(render_final_part_block(part.number, content))

    if 2 not in included_parts:
        raise RuntimeError("Part 2 output is missing or still a stub; final assembly cannot start")

    untrusted_parts = collect_untrusted_output_parts(run_workspace, included_parts)

    final_markdown = assemble_final_markdown_document(run_workspace, assembled_blocks)
    assembled_md = run_workspace.final_dir / "final.assembled.md"
    assembled_docx = run_workspace.final_dir / "final.assembled.docx"
    refresh_master_working_file(run_workspace)
    write_text(assembled_md, final_markdown)
    replace_docx_body_with_text(
        run_workspace.theme_workspace.paths["output_example_docx"],
        assembled_docx,
        final_markdown,
    )
    write_local_metric_checkpoint(run_workspace, "assembled_final", target="assembled")

    report: dict[str, Any] = {
        "generated_at": utc_now_iso(),
        "run_dir": str(run_workspace.run_dir),
        "subtopic_id": run_workspace.subtopic_entry.item_id,
        "included_parts": included_parts,
        "skipped_stub_parts": skipped_stub_parts,
        "missing_parts": missing_parts,
        "assembled_md": str(assembled_md),
        "assembled_docx": str(assembled_docx),
        "run_mode": "fresh-only",
        "trusted_fresh_run": not bool(untrusted_parts),
        "untrusted_parts": untrusted_parts,
        "published": False,
    }

    if publish:
        unresolved = sorted(set(skipped_stub_parts + missing_parts))
        if unresolved:
            raise RuntimeError(
                f"Cannot publish final output while parts are still missing/stub: {unresolved}"
            )
        if untrusted_parts:
            issues = ", ".join(
                f"Part {item['part_number']} ({item['source_origin']})" for item in untrusted_parts
            )
            raise RuntimeError(
                "Cannot publish this run as fresh-only: untrusted stage outputs detected: " + issues
            )
        enforce_publish_metric_floor(final_markdown)
        write_text(run_workspace.final_md_target, final_markdown)
        replace_docx_body_with_text(
            run_workspace.theme_workspace.paths["output_example_docx"],
            run_workspace.final_docx_target,
            final_markdown,
        )
        write_local_metric_checkpoint(run_workspace, "published_final", target="published")
        report["published"] = True
        report["published_md"] = str(run_workspace.final_md_target)
        report["published_docx"] = str(run_workspace.final_docx_target)

    write_json(run_workspace.final_dir / "assembly.report.json", report)
    return assembled_md


def read_clipboard_text() -> str:
    try:
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                (
                    "$ErrorActionPreference = 'Stop'; "
                    "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
                    "$text = Get-Clipboard -Raw; "
                    "if ($null -eq $text) { exit 3 }; "
                    "[Console]::Out.Write($text)"
                ),
            ],
            capture_output=True,
            text=False,
            check=False,
        )
    except Exception as exc:
        raise RuntimeError(f"Unable to read clipboard: {exc}") from exc
    if result.returncode == 3:
        raise RuntimeError("Clipboard is empty.")
    if result.returncode != 0:
        stderr = decode_clipboard_bytes(result.stderr or b"").strip()
        if stderr:
            raise RuntimeError(f"Unable to read clipboard: {stderr}")
        raise RuntimeError("Unable to read clipboard: PowerShell returned a non-zero exit code.")
    text = decode_clipboard_bytes(result.stdout or b"")
    if not text.strip():
        raise RuntimeError("Clipboard is empty.")
    return text


def decode_clipboard_bytes(payload: bytes) -> str:
    for encoding in ("utf-8", "utf-8-sig", "cp1251", "cp866", "cp1252"):
        try:
            return payload.decode(encoding)
        except UnicodeDecodeError:
            continue
    return payload.decode("utf-8", errors="replace")


def load_part_output_source(source_file: str | None, use_clipboard: bool) -> str:
    if bool(source_file) == bool(use_clipboard):
        raise RuntimeError("Specify exactly one source: --source-file or --clipboard")
    if source_file:
        content = read_text(Path(source_file).resolve())
    else:
        content = read_clipboard_text()
    if not content or not content.strip():
        raise RuntimeError("Part output source is empty.")
    return content


def write_subtopic_run_files(run_workspace: SubtopicRunWorkspace) -> None:
    paths = run_workspace.theme_workspace.paths
    copy_if_exists(paths["agents"], run_workspace.context_dir / "00-agents.md")
    copy_if_exists(paths["state"], run_workspace.context_dir / "01-project-state.md")
    copy_if_exists(paths["workflow"], run_workspace.context_dir / "02-manual-workflow.md")
    copy_if_exists(paths["master_prompt"], run_workspace.context_dir / "03-master-prompt.md")
    write_text(
        run_workspace.context_dir / "04-approved-theme.md",
        render_theme_block_markdown(run_workspace.theme_workspace.theme),
    )
    copy_if_exists(
        run_workspace.theme_workspace.outline_dir / "outline.active.md",
        run_workspace.context_dir / "05-outline.active.md",
    )
    copy_if_exists(run_workspace.order_path, run_workspace.context_dir / "06-order.copy.md")
    copy_if_exists(run_workspace.packet_path, run_workspace.context_dir / "07-execution-packet.md")
    copy_if_exists(paths["interaction_guide"], run_workspace.context_dir / "08-interaction-guide.md")

    write_text(
        run_workspace.stage_inputs_dir / "00-assignment-and-outline.md",
        run_workspace.intro_block.rstrip() + "\n",
    )
    for part in run_workspace.parts:
        input_path = run_workspace.stage_inputs_dir / part.filename
        write_text(input_path, part.content.rstrip() + "\n")
        write_text(
            run_workspace.stage_outputs_dir / part.filename,
            build_part_output_stub(run_workspace.subtopic_entry.line, part, input_path),
        )

    if not run_workspace.lean_artifacts:
        write_text(run_workspace.run_dir / "README.run.md", build_subtopic_run_readme(run_workspace))
        write_text(run_workspace.final_dir / "README.final.md", build_final_contract_readme(run_workspace))
        write_json(
            run_workspace.final_dir / "final.contract.json",
            build_final_output_contract(run_workspace),
        )
        final_skeleton_md = build_final_skeleton_markdown(run_workspace)
        write_text(run_workspace.final_dir / "final.skeleton.md", final_skeleton_md)
        replace_docx_body_with_text(
            run_workspace.theme_workspace.paths["output_example_docx"],
            run_workspace.final_dir / "final.skeleton.docx",
            final_skeleton_md,
        )
    master_working_md = refresh_master_working_file(run_workspace)
    lightweight_skeleton_paths = write_lightweight_skeletons(run_workspace)
    web_plan_paths = write_part_02_web_plan(run_workspace, overwrite=False)
    if run_workspace.lean_artifacts:
        part_03_light_paths = build_part_03_plan_paths(run_workspace)
        part_03_light_paths["plan_dir"].mkdir(parents=True, exist_ok=True)
        write_text_if_needed(part_03_light_paths["canonical_skeleton"], build_part_03_canonical_skeleton(), overwrite=False)
        part_04_light_paths = build_part_04_plan_paths(run_workspace)
        part_04_light_paths["plan_dir"].mkdir(parents=True, exist_ok=True)
        write_text_if_needed(part_04_light_paths["canonical_skeleton"], build_part_04_canonical_skeleton(), overwrite=False)
        part_05_light_paths = build_part_05_plan_paths(run_workspace)
        part_05_light_paths["plan_dir"].mkdir(parents=True, exist_ok=True)
        write_text_if_needed(part_05_light_paths["canonical_skeleton"], build_part_05_canonical_skeleton(), overwrite=False)
    part_03_plan_paths = write_part_03_plan(run_workspace, overwrite=False) if not run_workspace.lean_artifacts else None
    part_04_plan_paths = write_part_04_plan(run_workspace, overwrite=False) if not run_workspace.lean_artifacts else None
    part_05_plan_paths = write_part_05_plan(run_workspace, overwrite=False) if not run_workspace.lean_artifacts else None
    followup_part_paths = write_followup_part_packets(run_workspace, overwrite=False) if not run_workspace.lean_artifacts else None
    reasoning_layer_paths = write_reasoning_layer(run_workspace, overwrite=False) if not run_workspace.lean_artifacts else None
    semantic_dedup_paths = (
        write_semantic_dedup_layer(run_workspace, overwrite=False)
        if PRODUCTION_ENABLE_SEMANTIC_DEDUP
        else None
    )
    omission_audit_paths = (
        write_omission_audit_layer(run_workspace, overwrite=False)
        if PRODUCTION_ENABLE_OMISSION_AUDIT
        else None
    )
    refresh_dynamic_part_packets(run_workspace)
    write_json(
        run_workspace.run_dir / "manifest.json",
        {
            "generated_at": utc_now_iso(),
            "status": "subtopic_run_prepared",
            "run_mode": run_workspace.run_mode,
            "allowed_parts": list(run_workspace.allowed_parts),
            "strict_lean_mode": run_workspace.strict_lean_mode,
            "orchestration_target_percent": run_workspace.orchestration_target_percent,
            "artifact_profile": run_workspace.artifact_profile,
            "packet_mode": run_workspace.packet_mode,
            "packet_mode_reason": run_workspace.packet_mode_reason,
            "allow_reuse_stage_outputs": False,
            "theme_id": run_workspace.theme_workspace.theme.theme_id,
            "theme_title": run_workspace.theme_workspace.theme.full_title,
            "subtopic_id": run_workspace.subtopic_entry.item_id,
            "subtopic_line": run_workspace.subtopic_entry.line,
            "run_dir": str(run_workspace.run_dir),
            "order_file": str(run_workspace.order_path),
            "packet_file": str(run_workspace.packet_path),
            "canonical_final_md_target": str(run_workspace.final_md_target),
            "canonical_final_docx_target": str(run_workspace.final_docx_target),
            "master_working_md": str(master_working_md),
            "lightweight_skeletons": {
                "part_10": str(lightweight_skeleton_paths["part_10"]),
                "part_11": str(lightweight_skeleton_paths["part_11"]),
            },
            "part_02_web_plan": {
                "web_plan_dir": str(web_plan_paths["web_plan_dir"]),
                "operator_sequence": str(web_plan_paths["operator_sequence"]),
                "focus_json": str(web_plan_paths["focus_json"]),
                "queries_md": str(web_plan_paths["queries_md"]),
                "queries_json": str(web_plan_paths["queries_json"]),
                "source_cascade_md": str(web_plan_paths["source_cascade_md"]),
                "source_cascade_json": str(web_plan_paths["source_cascade_json"]),
                "research_pack": str(web_plan_paths["research_pack"]),
                "core_template": str(web_plan_paths["core_template"]),
                "launch_packet": str(web_plan_paths["launch_packet"]),
                "message_01": str(web_plan_paths["message_01"]),
                "message_02": str(web_plan_paths["message_02"]),
                "message_03": str(web_plan_paths["message_03"]),
                "research_log": str(web_plan_paths["research_log"]),
                "evidence_dir": str(web_plan_paths["evidence_dir"]),
            },
            "parts": [
                {
                    "part_number": part.number,
                    "heading": part.heading,
                    "execution_mode": PART_EXECUTION_MODES.get(part.number, "custom"),
                    "input_file": str(run_workspace.stage_inputs_dir / part.filename),
                    "output_file": str(run_workspace.stage_outputs_dir / part.filename),
                    "status": "ready" if part.number == 1 else "blocked_until_go",
                    "source_origin": "stub_template",
                }
                for part in run_workspace.parts
            ],
        },
    )
    manifest_path = run_workspace.run_dir / "manifest.json"
    manifest = json.loads(read_text(manifest_path))
    if not run_workspace.lean_artifacts:
        manifest["final_contract_file"] = str(run_workspace.final_dir / "final.contract.json")
        manifest["final_skeleton_md"] = str(run_workspace.final_dir / "final.skeleton.md")
        manifest["final_skeleton_docx"] = str(run_workspace.final_dir / "final.skeleton.docx")
        if part_03_plan_paths is not None:
            manifest["part_03_plan"] = {
                "plan_dir": str(part_03_plan_paths["plan_dir"]),
                "operator_sequence": str(part_03_plan_paths["operator_sequence"]),
                "capture_status": str(part_03_plan_paths["capture_status"]),
                "message_files": {
                    str(segment_id): str(path)
                    for segment_id, path in part_03_plan_paths["message_files"].items()
                },
            }
        if part_04_plan_paths is not None:
            manifest["part_04_plan"] = {
                "plan_dir": str(part_04_plan_paths["plan_dir"]),
                "operator_sequence": str(part_04_plan_paths["operator_sequence"]),
                "capture_status": str(part_04_plan_paths["capture_status"]),
                "message_files": {
                    str(segment_id): str(path)
                    for segment_id, path in part_04_plan_paths["message_files"].items()
                },
            }
        if part_05_plan_paths is not None:
            manifest["part_05_plan"] = {
                "plan_dir": str(part_05_plan_paths["plan_dir"]),
                "operator_sequence": str(part_05_plan_paths["operator_sequence"]),
                "capture_status": str(part_05_plan_paths["capture_status"]),
                "message_files": {
                    str(segment_id): str(path)
                    for segment_id, path in part_05_plan_paths["message_files"].items()
                },
            }
        if followup_part_paths is not None:
            manifest["followup_part_packets"] = {
                "plan_dir": str(followup_part_paths["plan_dir"]),
                "readme": str(followup_part_paths["readme"]),
                "message_files": {
                    str(part_number): str(path)
                    for part_number, path in followup_part_paths["message_files"].items()
                },
            }
        if reasoning_layer_paths is not None:
            manifest["reasoning_layer"] = {
                "reasoning_dir": str(reasoning_layer_paths["reasoning_dir"]),
                "readme": str(reasoning_layer_paths["readme"]),
                "brief_files": {
                    str(part_number): str(path)
                    for part_number, path in reasoning_layer_paths["brief_files"].items()
                },
            }
        write_json(manifest_path, manifest)
    if semantic_dedup_paths is not None:
        manifest_path = run_workspace.run_dir / "manifest.json"
        manifest = json.loads(read_text(manifest_path))
        manifest["semantic_dedup"] = {
            "dedup_dir": str(semantic_dedup_paths["dedup_dir"]),
            "readme": str(semantic_dedup_paths["readme"]),
            "snapshot_json": str(semantic_dedup_paths["snapshot_json"]),
            "brief_files": {
                str(part_number): str(path)
                for part_number, path in semantic_dedup_paths["brief_files"].items()
            },
        }
        write_json(manifest_path, manifest)
    if omission_audit_paths is not None:
        manifest_path = run_workspace.run_dir / "manifest.json"
        manifest = json.loads(read_text(manifest_path))
        manifest["omission_audit"] = {
            "audit_dir": str(omission_audit_paths["audit_dir"]),
            "readme": str(omission_audit_paths["readme"]),
            "snapshot_json": str(omission_audit_paths["snapshot_json"]),
            "brief_files": {
                str(part_number): str(path)
                for part_number, path in omission_audit_paths["brief_files"].items()
            },
        }
        write_json(manifest_path, manifest)
    drop_disabled_manifest_layers(run_workspace)


def default_workflow_paths(workspace_root: Path) -> dict[str, Path]:
    return {
        "agents": workspace_root / DEFAULT_AGENTS_PATH,
        "state": workspace_root / DEFAULT_STATE_PATH,
        "workflow": workspace_root / DEFAULT_WORKFLOW_PATH,
        "master_prompt": workspace_root / DEFAULT_MASTER_PROMPT_PATH,
        "approved_topics": workspace_root / DEFAULT_APPROVED_TOPICS_PATH,
        "order_template": workspace_root / DEFAULT_ORDER_TEMPLATE_PATH,
        "interaction_guide": workspace_root / DEFAULT_INTERACTION_GUIDE_PATH,
        "outline_overrides_root": workspace_root / DEFAULT_OUTLINE_OVERRIDES_ROOT,
        "output_example_md": workspace_root / DEFAULT_OUTPUT_EXAMPLE_MD_PATH,
        "output_example_docx": workspace_root / DEFAULT_OUTPUT_EXAMPLE_DOCX_PATH,
        "output_root": workspace_root / DEFAULT_OUTPUT_READY_ROOT,
    }


def prepare_main_theme_workspace(
    workspace_root: Path,
    theme_query: str,
    packet_mode: str = DEFAULT_PACKET_MODE,
    packet_mode_reason: str = "",
) -> MainThemeWorkspace:
    packet_mode = normalize_packet_mode(packet_mode)
    packet_mode_reason = normalize_packet_mode_reason(packet_mode_reason)
    paths = default_workflow_paths(workspace_root)
    required = [
        "agents",
        "state",
        "workflow",
        "master_prompt",
        "approved_topics",
        "order_template",
        "interaction_guide",
        "output_example_md",
        "output_example_docx",
    ]
    missing = [str(paths[key]) for key in required if not paths[key].exists()]
    if missing:
        raise FileNotFoundError("Missing required workflow files:\n" + "\n".join(missing))

    approved_text = read_text(paths["approved_topics"])
    themes = parse_approved_themes(approved_text)
    theme = find_approved_theme(themes, theme_query)

    template_text = read_text(paths["order_template"])
    template_info = extract_order_template_info(template_text)
    master_prompt_text = read_text(paths["master_prompt"])
    interaction_guide_text = read_text(paths["interaction_guide"])

    theme_folder = paths["output_root"] / safe_slug(theme.full_title)
    context_dir = theme_folder / "00-context"
    outline_dir = theme_folder / "01-outline"
    orders_dir = theme_folder / "02-orders"
    packets_dir = theme_folder / "03-packets"
    final_md_dir = theme_folder / "04-output" / "md"
    final_docx_dir = theme_folder / "04-output" / "docx"

    for directory in [context_dir, outline_dir, orders_dir, packets_dir, final_md_dir, final_docx_dir]:
        directory.mkdir(parents=True, exist_ok=True)

    default_outline_entries = build_default_outline_entries(theme)
    outline_entries, override_path = resolve_outline_entries(
        theme=theme,
        theme_folder=theme_folder,
        workspace_root=workspace_root,
    )

    return MainThemeWorkspace(
        workspace_root=workspace_root,
        paths=paths,
        theme=theme,
        theme_folder=theme_folder,
        context_dir=context_dir,
        outline_dir=outline_dir,
        orders_dir=orders_dir,
        packets_dir=packets_dir,
        final_md_dir=final_md_dir,
        final_docx_dir=final_docx_dir,
        template_text=template_text,
        template_info=template_info,
        master_prompt_text=master_prompt_text,
        interaction_guide_text=interaction_guide_text,
        default_outline_entries=default_outline_entries,
        outline_entries=outline_entries,
        override_path=override_path,
        packet_mode=packet_mode,
        packet_mode_reason=packet_mode_reason,
    )


def write_main_theme_context_files(workspace: MainThemeWorkspace) -> None:
    copy_if_exists(workspace.paths["agents"], workspace.context_dir / "00-agents.md")
    copy_if_exists(workspace.paths["state"], workspace.context_dir / "01-project-state.md")
    copy_if_exists(workspace.paths["workflow"], workspace.context_dir / "02-manual-workflow.md")
    copy_if_exists(workspace.paths["master_prompt"], workspace.context_dir / "03-master-prompt.md")
    write_text(workspace.context_dir / "04-approved-theme.md", render_theme_block_markdown(workspace.theme))
    copy_if_exists(workspace.paths["approved_topics"], workspace.context_dir / "04a-approved-topics-full.md")
    copy_if_exists(workspace.paths["order_template"], workspace.context_dir / "05-order-template.md")
    copy_if_exists(workspace.paths["interaction_guide"], workspace.context_dir / "06-interaction-guide.md")
    copy_if_exists(workspace.paths["output_example_md"], workspace.context_dir / "07-output-example.md")
    copy_if_exists(workspace.paths["output_example_docx"], workspace.context_dir / "08-output-example.docx")


def write_outline_phase_files(workspace: MainThemeWorkspace) -> None:
    write_text(
        workspace.outline_dir / "outline.generated.md",
        render_outline_markdown(workspace.theme, workspace.default_outline_entries),
    )
    if workspace.override_path:
        copy_if_exists(Path(workspace.override_path), workspace.outline_dir / "outline.approved.md")
    write_text(
        workspace.outline_dir / "outline.active.md",
        render_outline_markdown(workspace.theme, workspace.outline_entries),
    )
    write_json(
        workspace.outline_dir / "source-subtopics.json",
        [
            {
                "source_index": item.source_index,
                "generated_id": f"{workspace.theme.theme_id}.{item.source_index}",
                "title": item.title,
                "month": item.month,
            }
            for item in workspace.theme.subtopics
        ],
    )


def write_outline_review_stub(workspace: MainThemeWorkspace) -> Path:
    review_path = workspace.outline_dir / "outline.review.md"
    override_target = workspace.paths["outline_overrides_root"] / f"{workspace.theme.theme_id}.md"
    lines = [
        f"# Review Required: {workspace.theme.full_title}",
        "",
        "Утвержденного outline override пока нет.",
        "",
        "Что сделать:",
        "",
        f"1. Проверить draft в `{workspace.outline_dir / 'outline.generated.md'}`.",
        f"2. Подготовить финальный approved outline в `{override_target}`.",
        "3. Перезапустить `prepare-main-theme` после утверждения Оглавления.",
        "",
        "Ниже помещен текущий draft как стартовая точка для правки.",
        "",
    ]
    for entry in workspace.default_outline_entries:
        lines.append(entry.line)
    lines.append("")
    write_text(review_path, "\n".join(lines))
    return review_path


def load_topic_bundle(topic_dir: Path, include_existing_notes: bool) -> TopicBundle:
    topic_md = topic_dir / "topic.md"
    if not topic_md.exists():
        raise FileNotFoundError(f"Missing topic.md in {topic_dir}")
    meta, body = simple_frontmatter(read_text(topic_md))

    prompt_path = topic_dir / "master-prompt.md"
    master_prompt = read_text(prompt_path) if prompt_path.exists() else ""

    order_parts: list[tuple[str, str]] = []
    preferred_orders_dir = topic_dir / "orders" / "clean"
    fallback_orders_dir = topic_dir / "orders"
    orders_dir = preferred_orders_dir if preferred_orders_dir.exists() else fallback_orders_dir
    if orders_dir.exists():
        for path in sorted(orders_dir.glob("*.md")):
            order_parts.append((path.name, read_text(path)))

    source_notes: list[tuple[str, str]] = []
    if include_existing_notes:
        notes_dir = topic_dir / "source-notes"
        if notes_dir.exists():
            for path in sorted(notes_dir.glob("*.md")):
                source_notes.append((path.name, read_text(path)))

    return TopicBundle(
        topic_dir=topic_dir,
        metadata=meta,
        topic_body=body.strip(),
        master_prompt=master_prompt.strip(),
        order_parts=order_parts,
        source_notes=source_notes,
    )


def build_agent_instructions() -> str:
    return textwrap.dedent(
        """
        Ты автономный исполнитель по юридическому поиску для нотариальной стажировки в Российской Федерации.
        Работай строго по предоставленному мастер-промпту и материалам темы.
        Используй веб-поиск для проверки действующих документов и ссылок.
        Если мастер-промпт требует строгого формата ссылок, соблюдай это в итоговом тексте.
        Возвращай только JSON-объект без markdown-оберток и без пояснений вне JSON.

        JSON-формат:
        {
          "topic_id": "строка",
          "topic_title": "строка",
          "table_of_contents": ["строка", "..."],
          "final_report_markdown": "строка",
          "notes": ["строка", "..."]
        }
        """
    ).strip()


def build_model_input(bundle: TopicBundle) -> str:
    sections = [
        f"Текущая дата UTC: {utc_now_iso()}",
        "",
        "=== ТЕМА ===",
        bundle.topic_body,
        "",
    ]
    if bundle.master_prompt:
        sections.extend(["=== МАСТЕР-ПРОМПТ ===", bundle.master_prompt, ""])
    if bundle.order_parts:
        for name, content in bundle.order_parts:
            sections.extend([f"=== ЧАСТЬ ПРИКАЗА: {name} ===", content, ""])
    if bundle.source_notes:
        for name, content in bundle.source_notes:
            sections.extend([f"=== ПРЕДЫДУЩИЙ МАТЕРИАЛ: {name} ===", content, ""])
    sections.extend(
        [
            "=== ЗАДАЧА ===",
            "1. Определи структуру содержания по теме.",
            "2. Выполни юридический поиск по правилам мастер-промпта.",
            "3. Подготовь итоговый текст так, чтобы его можно было сразу вставить в рабочий документ.",
        ]
    )
    return "\n".join(sections).strip()


def build_request_payload(bundle: TopicBundle, config: dict[str, Any]) -> dict[str, Any]:
    tool: dict[str, Any] = {"type": "web_search"}
    allowed_domains = config.get("web_allowed_domains") or []
    if allowed_domains:
        tool["filters"] = {"allowed_domains": allowed_domains}

    payload: dict[str, Any] = {
        "model": config["model"],
        "instructions": build_agent_instructions(),
        "input": build_model_input(bundle),
        "tools": [tool],
        "tool_choice": "auto",
        "store": bool(config.get("store", True)),
    }
    reasoning_effort = config.get("reasoning_effort")
    if reasoning_effort:
        payload["reasoning"] = {"effort": reasoning_effort}
    if config.get("include_sources", True):
        payload["include"] = ["web_search_call.action.sources"]
    return payload


def call_openai(payload: dict[str, Any], api_key: str, timeout_seconds: int) -> dict[str, Any]:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
            raw = response.read().decode("utf-8")
            return json.loads(raw)
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"OpenAI API error {exc.code}: {details}") from exc


def extract_output_text(response_json: dict[str, Any]) -> str:
    chunks: list[str] = []
    for item in response_json.get("output", []):
        if item.get("type") != "message":
            continue
        for content in item.get("content", []):
            if content.get("type") == "output_text":
                chunks.append(content.get("text", ""))
    return "\n".join(chunk for chunk in chunks if chunk).strip()


def maybe_parse_json(text: str) -> dict[str, Any] | None:
    text = text.strip()
    if not text:
        return None
    attempts = [text]
    if text.startswith("```"):
        stripped = re.sub(r"^```(?:json)?\s*", "", text)
        stripped = re.sub(r"\s*```$", "", stripped)
        attempts.append(stripped.strip())
    first = text.find("{")
    last = text.rfind("}")
    if first != -1 and last != -1 and last > first:
        attempts.append(text[first : last + 1])
    for candidate in attempts:
        try:
            parsed = json.loads(candidate)
        except json.JSONDecodeError:
            continue
        if isinstance(parsed, dict):
            return parsed
    return None


def markdown_lines_from_output(
    topic_title: str,
    toc: list[str],
    report_markdown: str,
    notes: list[str],
) -> list[str]:
    lines = [f"# {topic_title}", ""]
    if toc:
        lines.extend(["## Содержание", ""])
        for item in toc:
            lines.append(f"- {item}")
        lines.append("")
    lines.extend(["## Итог", "", report_markdown.strip(), ""])
    if notes:
        lines.extend(["## Примечания", ""])
        for note in notes:
            lines.append(f"- {note}")
        lines.append("")
    return lines


def normalize_label_only_code_blocks(content: str) -> str:
    labels = ("Каноническая строка поиска:", "URL1:", "URL2:")
    normalized = content
    for label in labels:
        normalized = re.sub(
            rf"```(?:text)?\s*\n{re.escape(label)}\s*\n```\s*\n?",
            f"{label}\n",
            normalized,
            flags=re.I,
        )
    return normalized


def is_link_metadata_line(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return True
    if is_link_group_label(stripped):
        return True
    if stripped.startswith("VERIFIED URL2:"):
        return True
    if stripped.startswith("Заголовок страницы URL2"):
        return True
    if stripped.startswith("Сверка реквизитов:"):
        return True
    if line_has_link_token(stripped):
        return True
    if stripped.startswith("отсутствует"):
        return True
    return False


def flatten_link_only_code_blocks(content: str) -> str:
    lines = content.splitlines()
    result: list[str] = []
    in_code_block = False
    code_lines: list[str] = []

    def flush_code_block() -> None:
        nonlocal code_lines
        if not code_lines:
            return
        if any(line.strip() for line in code_lines) and all(is_link_metadata_line(line) for line in code_lines):
            result.extend(code_lines)
        else:
            result.append("```text")
            result.extend(code_lines)
            result.append("```")
        code_lines = []

    for raw_line in lines:
        stripped = raw_line.strip()
        if stripped.startswith("```"):
            if in_code_block:
                flush_code_block()
                in_code_block = False
            else:
                in_code_block = True
                code_lines = []
            continue
        if in_code_block:
            code_lines.append(raw_line)
            continue
        result.append(raw_line)

    if in_code_block:
        flush_code_block()

    return "\n".join(result)


def is_structural_heading(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    if stripped.startswith("ЧАСТЬ "):
        return True
    if stripped.startswith("ТЕМА:"):
        return True
    if stripped in PROMINENT_INNER_HEADINGS:
        return True
    if re.match(r"^[IVXLCDM]+\.\s", stripped):
        return True
    return False


def renumber_document_cards(content: str) -> str:
    counters: dict[str, int] = {}
    current_scope = "__global__"
    result_lines: list[str] = []

    for raw_line in content.splitlines():
        stripped = raw_line.strip()
        if is_structural_heading(stripped):
            current_scope = stripped
            counters.setdefault(current_scope, 0)
            result_lines.append(raw_line)
            continue
        if re.match(r"^\d+\.\s+Вид документа:", stripped):
            result_lines.append(raw_line)
            continue
        if stripped.startswith("Вид документа:"):
            counters[current_scope] = counters.get(current_scope, 0) + 1
            indent = raw_line[: len(raw_line) - len(raw_line.lstrip())]
            result_lines.append(
                f"{indent}{counters[current_scope]}. {stripped}"
            )
            continue
        result_lines.append(raw_line)

    return "\n".join(result_lines)


def normalize_final_part_content(content: str) -> str:
    normalized = flatten_link_only_code_blocks(content.strip())
    normalized = normalize_label_only_code_blocks(normalized)
    normalized = renumber_document_cards(normalized)
    normalized = re.sub(r"\n{3,}", "\n\n", normalized).strip()
    return normalized


def strip_leading_part_heading(content: str, part_number: int) -> str:
    lines = content.splitlines()
    if not lines:
        return content

    heading_re = re.compile(
        rf"^\s*(?:#+\s*)?(?:\*+\s*)?ЧАСТЬ\s+{part_number}\.(?:\s+.+)?(?:\s*\*+)?\s*$",
        re.IGNORECASE,
    )
    first_nonempty_idx = next((idx for idx, line in enumerate(lines) if line.strip()), None)
    if first_nonempty_idx is None:
        return content
    if not heading_re.match(lines[first_nonempty_idx].strip()):
        return content

    remaining = lines[first_nonempty_idx + 1 :]
    while remaining and not remaining[0].strip():
        remaining = remaining[1:]
    return "\n".join(remaining)


def render_final_part_block(part_number: int, content: str) -> str:
    title = FINAL_PART_TITLES.get(part_number, f"ЧАСТЬ {part_number}")
    normalized = normalize_final_part_content(strip_leading_part_heading(content, part_number))
    return f"## ЧАСТЬ {part_number}. {title}\n\n{normalized}"


def assemble_final_markdown_document(
    run_workspace: SubtopicRunWorkspace,
    assembled_blocks: list[str],
) -> str:
    lines = [f"# {run_workspace.subtopic_entry.line}", ""]
    for block in assembled_blocks:
        block = block.strip()
        if not block:
            continue
        lines.append(block)
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def build_master_working_markdown(run_workspace: SubtopicRunWorkspace) -> str:
    lines = [f"# {run_workspace.subtopic_entry.line}", ""]
    for part in run_workspace.parts:
        if part.number == 1:
            continue
        output_path = run_workspace.stage_outputs_dir / part.filename
        title = FINAL_PART_TITLES.get(part.number, f"ЧАСТЬ {part.number}")
        lines.append(f"## ЧАСТЬ {part.number}. {title}")
        lines.append("")
        if output_path.exists():
            content = read_text(output_path).strip()
            if not is_response_stub(content):
                if part.number == 3:
                    content = sanitize_substantive_part_output(run_workspace, part.number, content).strip()
                content = strip_leading_part_heading(content, part.number)
                lines.append(normalize_final_part_content(content) if content else "[Пока не заполнено]")
            else:
                lines.append("[Пока не заполнено]")
        else:
            lines.append("[Пока не заполнено]")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def build_lightweight_skeleton_paths(run_workspace: SubtopicRunWorkspace) -> dict[str, Path]:
    skeleton_dir = run_workspace.final_dir / "skeletons"
    return {
        "dir": skeleton_dir,
        "part_10": skeleton_dir / "part-10.canonical-skeleton.md",
        "part_11": skeleton_dir / "part-11.canonical-skeleton.md",
    }


def write_lightweight_skeletons(run_workspace: SubtopicRunWorkspace) -> dict[str, Path]:
    paths = build_lightweight_skeleton_paths(run_workspace)
    paths["dir"].mkdir(parents=True, exist_ok=True)
    write_text(paths["part_10"], build_part_10_canonical_skeleton())
    write_text(paths["part_11"], build_part_11_canonical_skeleton())
    return paths


def refresh_master_working_file(run_workspace: SubtopicRunWorkspace) -> Path:
    master_path = run_workspace.final_dir / "working.master.md"
    write_text(master_path, build_master_working_markdown(run_workspace))
    return master_path


def compute_text_metrics(text: str) -> dict[str, int]:
    lines = text.splitlines()
    words = len(re.findall(r"\S+", text))
    chars = len(text)
    url1 = sum(1 for line in lines if line.strip().startswith("URL1:"))
    url2 = sum(1 for line in lines if line.strip().startswith("URL2:"))
    verified_url2 = sum(1 for line in lines if line.strip().startswith("VERIFIED URL2:"))
    document_cards = count_document_cards(text)
    return {
        "words": words,
        "chars": chars,
        "url1": url1,
        "url2": url2,
        "verified_url2": verified_url2,
        "document_cards": document_cards,
    }


def collect_publish_metric_shortfalls(metrics: dict[str, int]) -> list[str]:
    shortfalls: list[str] = []
    if metrics["words"] < PUBLISH_MIN_WORDS:
        shortfalls.append(f"words<{PUBLISH_MIN_WORDS}")
    if metrics["chars"] < PUBLISH_MIN_CHARS:
        shortfalls.append(f"chars<{PUBLISH_MIN_CHARS}")
    if metrics["url1"] < PUBLISH_MIN_URL1:
        shortfalls.append(f"url1<{PUBLISH_MIN_URL1}")
    if metrics["url2"] < PUBLISH_MIN_URL2:
        shortfalls.append(f"url2<{PUBLISH_MIN_URL2}")
    return shortfalls


def enforce_publish_metric_floor(text: str) -> dict[str, int]:
    metrics = compute_text_metrics(text)
    shortfalls = collect_publish_metric_shortfalls(metrics)
    if shortfalls:
        raise RuntimeError(
            "Метрики не достигнуты: "
            f"{metrics['words']} слов, "
            f"{metrics['url1']} URL1, "
            f"{metrics['url2']} URL2. "
            "Требуется добор документов."
        )
    return metrics


def read_docx_page_count(path: Path) -> int | None:
    if not path.exists():
        return None
    try:
        with zipfile.ZipFile(path) as archive:
            xml_bytes = archive.read("docProps/app.xml")
        root = ET.fromstring(xml_bytes)
        pages_node = root.find(".//Pages")
        if pages_node is None or not (pages_node.text or "").strip():
            return None
        return int((pages_node.text or "").strip())
    except Exception:
        return None


def write_local_metric_checkpoint(run_workspace: SubtopicRunWorkspace, checkpoint_name: str, target: str = "master") -> Path:
    checkpoints_dir = run_workspace.final_dir / "checkpoints"
    checkpoints_dir.mkdir(parents=True, exist_ok=True)
    if target == "master":
        source_path = refresh_master_working_file(run_workspace)
        text = read_text(source_path)
        page_count = None
    elif target == "assembled":
        source_path = run_workspace.final_dir / "final.assembled.md"
        if not source_path.exists():
            return checkpoints_dir / f"{checkpoint_name}.missing.json"
        text = read_text(source_path)
        page_count = read_docx_page_count(run_workspace.final_dir / "final.assembled.docx")
    elif target == "published":
        source_path = run_workspace.final_md_target
        if not source_path.exists():
            return checkpoints_dir / f"{checkpoint_name}.missing.json"
        text = read_text(source_path)
        page_count = read_docx_page_count(run_workspace.final_docx_target)
    else:
        raise RuntimeError("Unsupported checkpoint target")
    metrics = compute_text_metrics(text)
    payload = {
        "generated_at": utc_now_iso(),
        "checkpoint": checkpoint_name,
        "target": target,
        "source_path": str(source_path),
        "words": metrics["words"],
        "chars": metrics["chars"],
        "url1": metrics["url1"],
        "url2": metrics["url2"],
        "verified_url2": metrics["verified_url2"],
        "document_cards": metrics["document_cards"],
        "page_count": page_count,
    }
    checkpoint_path = checkpoints_dir / f"{checkpoint_name}.{target}.json"
    write_json(checkpoint_path, payload)
    return checkpoint_path


def paragraph_element(
    text: str,
    *,
    style: str = "body",
) -> ET.Element:
    paragraph = ET.Element(f"{{{W_NS}}}p")
    p_pr = ET.SubElement(paragraph, f"{{{W_NS}}}pPr")
    spacing = ET.SubElement(p_pr, f"{{{W_NS}}}spacing")

    if style == "title":
        spacing.set(f"{{{W_NS}}}before", "0")
        spacing.set(f"{{{W_NS}}}after", "80")
    elif style in {"section", "subsection"}:
        spacing.set(f"{{{W_NS}}}before", "80")
        spacing.set(f"{{{W_NS}}}after", "40")
    elif style in {"code", "code-spacer"}:
        spacing.set(f"{{{W_NS}}}before", "0")
        spacing.set(f"{{{W_NS}}}after", "0")
        ind = ET.SubElement(p_pr, f"{{{W_NS}}}ind")
        ind.set(f"{{{W_NS}}}left", "240")
        ind.set(f"{{{W_NS}}}right", "240")
        shd = ET.SubElement(p_pr, f"{{{W_NS}}}shd")
        shd.set(f"{{{W_NS}}}val", "clear")
        shd.set(f"{{{W_NS}}}color", "auto")
        shd.set(f"{{{W_NS}}}fill", "F3F3F3")
    elif style == "spacer":
        spacing.set(f"{{{W_NS}}}before", "0")
        spacing.set(f"{{{W_NS}}}after", "0")
    else:
        spacing.set(f"{{{W_NS}}}before", "0")
        spacing.set(f"{{{W_NS}}}after", "20")

    run = ET.SubElement(paragraph, f"{{{W_NS}}}r")
    r_pr = ET.SubElement(run, f"{{{W_NS}}}rPr")
    if style == "title":
        ET.SubElement(r_pr, f"{{{W_NS}}}b")
        size = ET.SubElement(r_pr, f"{{{W_NS}}}sz")
        size.set(f"{{{W_NS}}}val", "30")
    elif style in {"section", "subsection"}:
        ET.SubElement(r_pr, f"{{{W_NS}}}b")
        size = ET.SubElement(r_pr, f"{{{W_NS}}}sz")
        size.set(f"{{{W_NS}}}val", "24")
    elif style in {"code", "code-spacer"}:
        fonts = ET.SubElement(r_pr, f"{{{W_NS}}}rFonts")
        fonts.set(f"{{{W_NS}}}ascii", "Courier New")
        fonts.set(f"{{{W_NS}}}hAnsi", "Courier New")
        fonts.set(f"{{{W_NS}}}cs", "Courier New")
        size = ET.SubElement(r_pr, f"{{{W_NS}}}sz")
        size.set(f"{{{W_NS}}}val", "20")
    text_el = ET.SubElement(run, f"{{{W_NS}}}t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        text_el.set(f"{{{XML_NS}}}space", "preserve")
    text_el.text = text
    return paragraph


def build_docx_paragraph_specs(content: str) -> list[tuple[str, str]]:
    specs: list[tuple[str, str]] = []
    in_code_block = False
    previous_blank = False

    for raw_line in content.splitlines():
        line = raw_line.rstrip()
        stripped = line.strip()

        if stripped.startswith("```"):
            in_code_block = not in_code_block
            previous_blank = False
            continue

        if not in_code_block and stripped.startswith("#"):
            level = len(stripped) - len(stripped.lstrip("#"))
            text = stripped[level:].strip()
            specs.append((text, "title" if level == 1 else "section"))
            previous_blank = False
            continue

        if stripped == "":
            if previous_blank:
                continue
            specs.append(("", "code-spacer" if in_code_block else "spacer"))
            previous_blank = True
            continue

        text = line.replace("**", "")
        if in_code_block:
            specs.append((text, "code"))
        elif (
            stripped.startswith("URL1:")
            or stripped.startswith("URL2:")
            or stripped.startswith("VERIFIED URL2:")
            or stripped.startswith("Заголовок страницы URL2")
            or stripped.startswith("Сверка реквизитов:")
            or stripped.startswith("Каноническая строка поиска:")
            or line_has_link_token(stripped)
        ):
            specs.append((text, "code"))
        elif is_structural_heading(stripped):
            specs.append((text, "subsection"))
        else:
            specs.append((text, "body"))
        previous_blank = False

    return specs


def overwrite_docx_document_xml(docx_path: Path, new_xml: bytes) -> None:
    temp_path = docx_path.with_suffix(docx_path.suffix + ".tmp")
    with zipfile.ZipFile(docx_path, "r") as src, zipfile.ZipFile(temp_path, "w") as dst:
        for item in src.infolist():
            if item.filename == "word/document.xml":
                continue
            dst.writestr(item, src.read(item.filename))
        dst.writestr("word/document.xml", new_xml)
    temp_path.replace(docx_path)


def replace_docx_body_with_text(template_path: Path, output_path: Path, content: str) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)
    with zipfile.ZipFile(output_path, "r") as archive:
        xml_bytes = archive.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    body = root.find(f".//{{{W_NS}}}body")
    if body is None:
        raise RuntimeError("word/document.xml has no body")

    sect_pr = body.find(f"{{{W_NS}}}sectPr")
    children = list(body)
    for child in children:
        if sect_pr is not None and child is sect_pr:
            continue
        body.remove(child)

    insert_at = len(body)
    if sect_pr is not None:
        insert_at = list(body).index(sect_pr)

    new_paragraphs = [
        paragraph_element(text, style=style)
        for text, style in build_docx_paragraph_specs(content)
    ]
    for offset, paragraph in enumerate(new_paragraphs):
        body.insert(insert_at + offset, paragraph)

    new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    overwrite_docx_document_xml(output_path, new_xml)


def append_markdown_to_docx(template_path: Path, output_path: Path, content: str) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)
    with zipfile.ZipFile(output_path, "r") as archive:
        xml_bytes = archive.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    body = root.find(f".//{{{W_NS}}}body")
    if body is None:
        raise RuntimeError("word/document.xml has no body")

    sect_pr = body.find(f"{{{W_NS}}}sectPr")
    insert_at = len(body)
    if sect_pr is not None:
        insert_at = list(body).index(sect_pr)

    new_paragraphs = [
        paragraph_element("", style="spacer"),
        paragraph_element("Сгенерированный результат", style="section"),
        paragraph_element("", style="spacer"),
    ]
    for text, style in build_docx_paragraph_specs(content):
        new_paragraphs.append(paragraph_element(text, style=style))

    for offset, paragraph in enumerate(new_paragraphs):
        body.insert(insert_at + offset, paragraph)

    new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    overwrite_docx_document_xml(output_path, new_xml)


def save_run_artifacts(
    run_dir: Path,
    payload: dict[str, Any],
    response_json: dict[str, Any] | None,
    output_text: str,
    parsed: dict[str, Any] | None,
    bundle: TopicBundle,
) -> None:
    run_dir.mkdir(parents=True, exist_ok=True)
    write_text(run_dir / "request.json", json.dumps(payload, ensure_ascii=False, indent=2))

    if response_json is not None:
        write_text(run_dir / "response.json", json.dumps(response_json, ensure_ascii=False, indent=2))
    if output_text:
        write_text(run_dir / "response.txt", output_text)

    topic_title = bundle.title
    toc: list[str] = []
    report_markdown = output_text
    notes: list[str] = []

    if parsed is not None:
        write_text(run_dir / "result.json", json.dumps(parsed, ensure_ascii=False, indent=2))
        topic_title = str(parsed.get("topic_title") or topic_title)
        toc = [str(item) for item in parsed.get("table_of_contents", []) if str(item).strip()]
        report_markdown = str(parsed.get("final_report_markdown") or output_text)
        notes = [str(item) for item in parsed.get("notes", []) if str(item).strip()]

    final_markdown = "\n".join(markdown_lines_from_output(topic_title, toc, report_markdown, notes)).rstrip() + "\n"
    write_text(run_dir / "final-report.md", final_markdown)

    if bundle.template_docx is not None:
        append_markdown_to_docx(bundle.template_docx, run_dir / "final-report.docx", final_markdown)


def run_single_topic(
    topic_dir: Path,
    config: dict[str, Any],
    dry_run: bool,
    include_existing_notes: bool,
) -> Path:
    bundle = load_topic_bundle(topic_dir, include_existing_notes=include_existing_notes)
    payload = build_request_payload(bundle, config)
    output_root = Path(config.get("output_root", "runs"))
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = output_root / safe_slug(bundle.topic_id) / timestamp

    if dry_run:
        save_run_artifacts(run_dir, payload, None, "", None, bundle)
        write_text(run_dir / "assembled-input.md", payload["input"])
        return run_dir

    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not set")

    response_json = call_openai(
        payload=payload,
        api_key=api_key,
        timeout_seconds=int(config.get("request_timeout_seconds", 600)),
    )
    output_text = extract_output_text(response_json)
    parsed = maybe_parse_json(output_text)
    save_run_artifacts(run_dir, payload, response_json, output_text, parsed, bundle)
    return run_dir


def collect_imported_topics(root: Path) -> list[Path]:
    topics: list[Path] = []
    for path in root.iterdir():
        if path.is_dir() and (path / "topic.md").exists():
            topics.append(path)
    return sorted(topics)


def cmd_prepare_surgical_redo(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    allowed_parts = normalize_allowed_parts(args.parts)
    if not allowed_parts:
        raise RuntimeError("Surgical redo requires at least one allowed part.")
    run_workspace = ensure_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
        create_if_missing=False,
    )
    manifest_path = run_workspace.run_dir / "manifest.json"
    manifest = load_json_if_exists(manifest_path) or {}
    manifest["run_mode"] = RUN_MODE_SURGICAL
    manifest["allowed_parts"] = list(allowed_parts)
    manifest["strict_lean_mode"] = True
    manifest["orchestration_target_percent"] = STRICT_ORCHESTRATION_TARGET_PERCENT
    manifest["surgical_redo_reason"] = (args.reason or "").strip()
    manifest["forbidden_operations"] = [
        "run-subtopic",
        "execute-part-01",
        "assemble-subtopic-final",
        "prepare-part-04-plan",
        "prepare-part-05-plan",
        "capture-part-04-range",
        "capture-part-05-range",
        "archive/cleanup/housekeeping inside the same run",
        "full reread of unchanged canonical files",
    ]
    write_json(manifest_path, manifest)
    print(run_workspace.run_dir)
    return 0


def cmd_prepare_main_theme(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    workspace = prepare_main_theme_workspace(
        workspace_root,
        args.theme_query,
        packet_mode=args.packet_mode,
        packet_mode_reason=args.packet_mode_reason,
    )
    write_main_theme_context_files(workspace)
    write_outline_phase_files(workspace)

    if not workspace.override_path:
        review_path = write_outline_review_stub(workspace)
        write_text(
            workspace.theme_folder / "README.session.md",
            "\n".join(
                [
                    f"# Рабочий пакет: {workspace.theme.full_title}",
                    "",
                    "Утвержденного Оглавления пока нет, поэтому копии приказов не собраны.",
                    "",
                    f"- Draft outline: `{workspace.outline_dir / 'outline.generated.md'}`",
                    f"- Review file: `{review_path}`",
                    f"- Куда сохранить approved outline: `{workspace.paths['outline_overrides_root'] / f'{workspace.theme.theme_id}.md'}`",
                    "",
                    "После утверждения Оглавления перезапустите `prepare-main-theme`.",
                    "",
                ]
            ),
        )
        write_json(
            workspace.theme_folder / "manifest.json",
            {
                "generated_at": utc_now_iso(),
                "status": "outline_review_required",
                "theme_id": workspace.theme.theme_id,
                "theme_title": workspace.theme.full_title,
                "workspace_root": str(workspace.workspace_root),
                "outline_override_path": None,
                "paths": {key: str(value) for key, value in workspace.paths.items()},
                "generated_orders": [],
            },
        )
        print(workspace.theme_folder)
        return 0

    generated_orders: list[dict[str, str]] = []
    for entry in workspace.outline_entries:
        subtopic_id = entry.item_id
        subtopic_line = entry.line
        order_text = render_order_copy(
            template_text=workspace.template_text,
            template_info=workspace.template_info,
            theme=workspace.theme,
            outline_entries=workspace.outline_entries,
            focus_topic_line=subtopic_line,
        )
        order_name = f"{subtopic_id}.md"
        order_path = workspace.orders_dir / order_name
        write_text(order_path, order_text)

        packet_text = build_execution_packet(
            theme=workspace.theme,
            subtopic_line=subtopic_line,
            master_prompt_text=workspace.master_prompt_text,
            generated_order_text=order_text,
            interaction_guide_text=workspace.interaction_guide_text,
        )
        packet_path = workspace.packets_dir / order_name
        write_text(packet_path, packet_text)

        generated_orders.append(
            {
                "subtopic_id": subtopic_id,
                "subtopic_line": subtopic_line,
                "order_file": str(order_path),
                "packet_file": str(packet_path),
                "final_md_target": str(workspace.final_md_dir / f"{subtopic_id}.md"),
                "final_docx_target": str(workspace.final_docx_dir / f"{subtopic_id}.docx"),
            }
        )

    summary_lines = [
        f"# Рабочий пакет: {workspace.theme.full_title}",
        "",
        "## Что создано",
        "",
        f"- Пакет темы: `{workspace.theme_folder}`",
        f"- Draft outline: `{workspace.outline_dir / 'outline.generated.md'}`",
        f"- Active outline: `{workspace.outline_dir / 'outline.active.md'}`",
        f"- Копии приказов: `{workspace.orders_dir}`",
        f"- Execution packets: `{workspace.packets_dir}`",
        f"- Папка итоговых `.md`: `{workspace.final_md_dir}`",
        f"- Папка итоговых `.docx`: `{workspace.final_docx_dir}`",
        "",
        "## Подтемы",
        "",
    ]
    for item in workspace.outline_entries:
        summary_lines.append(f"- {item.line}")
    summary_lines.append("")
    write_text(workspace.theme_folder / "README.session.md", "\n".join(summary_lines))

    manifest = {
        "generated_at": utc_now_iso(),
        "status": "ready_for_execution",
        "theme_id": workspace.theme.theme_id,
        "theme_title": workspace.theme.full_title,
        "subtopic_count": len(workspace.outline_entries),
        "workspace_root": str(workspace.workspace_root),
        "outline_override_path": workspace.override_path,
        "paths": {key: str(value) for key, value in workspace.paths.items()},
        "generated_orders": generated_orders,
    }
    write_json(workspace.theme_folder / "manifest.json", manifest)

    print(workspace.theme_folder)
    return 0


def cmd_draft_main_theme_outline(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    workspace = prepare_main_theme_workspace(workspace_root, args.theme_query)
    write_main_theme_context_files(workspace)
    write_outline_phase_files(workspace)
    review_path = write_outline_review_stub(workspace)

    write_text(
        workspace.theme_folder / "README.session.md",
        "\n".join(
            [
                f"# Draft Outline Packet: {workspace.theme.full_title}",
                "",
                f"- Draft outline: `{workspace.outline_dir / 'outline.generated.md'}`",
                f"- Active outline: `{workspace.outline_dir / 'outline.active.md'}`",
                f"- Review file: `{review_path}`",
                f"- Approved outline target: `{workspace.paths['outline_overrides_root'] / f'{workspace.theme.theme_id}.md'}`",
                "",
                "Этот режим не создает копии приказов. Он предназначен только для подготовки и утверждения Оглавления.",
                "",
            ]
        ),
    )

    write_json(
        workspace.theme_folder / "manifest.json",
        {
            "generated_at": utc_now_iso(),
            "status": "draft_outline_prepared",
            "theme_id": workspace.theme.theme_id,
            "theme_title": workspace.theme.full_title,
            "workspace_root": str(workspace.workspace_root),
            "outline_override_path": workspace.override_path,
            "paths": {key: str(value) for key, value in workspace.paths.items()},
            "generated_orders": [],
        },
    )

    print(workspace.theme_folder)
    return 0


def cmd_run_subtopic(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    run_workspace = prepare_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
        artifact_profile=ARTIFACT_PROFILE_FULL if args.full_artifacts else ARTIFACT_PROFILE_LEAN,
        packet_mode=args.packet_mode,
        packet_mode_reason=args.packet_mode_reason,
    )
    write_subtopic_run_files(run_workspace)
    print(run_workspace.run_dir)
    return 0


def cmd_execute_part_01(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    run_workspace = ensure_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
    )
    assert_run_command_allowed(run_workspace, "execute-part-01", part_number=1)
    output_path = run_workspace.stage_outputs_dir / "part-01.md"
    if output_path.exists() and not args.force:
        current = read_text(output_path)
        if not is_response_stub(current):
            raise RuntimeError(
                f"Part 01 already has non-stub content: {output_path}. Use --force to overwrite."
            )
    write_text(output_path, build_part_01_response(run_workspace))
    update_run_manifest_part_status(
        run_workspace,
        1,
        "completed_waiting_for_go",
        source_origin="execute_part_01",
    )
    refresh_dynamic_part_packets(run_workspace)
    refresh_master_working_file(run_workspace)
    print(output_path)
    return 0


def cmd_assemble_subtopic_final(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    run_workspace = ensure_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
    )
    assert_run_command_allowed(run_workspace, "assemble-subtopic-final")
    if args.publish:
        assembled_md = assemble_subtopic_final(run_workspace, publish=False)
        assembled_text = read_text(assembled_md)
        metrics = compute_text_metrics(assembled_text)
        shortfalls = collect_publish_metric_shortfalls(metrics)
        pre_publish_payload = {
            "subtopic_id": run_workspace.subtopic_entry.item_id,
            "target": "pre-publish",
            "source_path": str(assembled_md),
            "words": metrics["words"],
            "chars": metrics["chars"],
            "url1": metrics["url1"],
            "url2": metrics["url2"],
            "verified_url2": metrics["verified_url2"],
            "document_cards": metrics["document_cards"],
            "publish_thresholds": {
                "words": PUBLISH_MIN_WORDS,
                "chars": PUBLISH_MIN_CHARS,
                "url1": PUBLISH_MIN_URL1,
                "url2": PUBLISH_MIN_URL2,
            },
            "publish_ready": not bool(shortfalls),
            "shortfalls": shortfalls,
        }
        print(json.dumps(pre_publish_payload, ensure_ascii=False, indent=2))
        if shortfalls:
            enforce_publish_metric_floor(assembled_text)
    assembled_md = assemble_subtopic_final(run_workspace, publish=args.publish)
    if args.publish:
        theme_number = run_workspace.theme_workspace.theme.theme_id
        current_subtopic_id = run_workspace.subtopic_entry.item_id
        update_project_state_next_subtopic(
            workspace_root=workspace_root,
            theme_number=theme_number,
            current_subtopic_id=current_subtopic_id,
        )
    print(assembled_md)
    return 0


def cmd_metric_check(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    run_workspace = ensure_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
    )
    assert_run_command_allowed(run_workspace, "metric-check")
    target = (args.target or "master").strip().lower()
    if target == "master":
        source_path = refresh_master_working_file(run_workspace)
        text = read_text(source_path)
        page_count = None
    elif target == "assembled":
        source_path = run_workspace.final_dir / "final.assembled.md"
        if not source_path.exists():
            raise RuntimeError("Assembled markdown is missing. Run assemble-subtopic-final first.")
        text = read_text(source_path)
        page_count = read_docx_page_count(run_workspace.final_dir / "final.assembled.docx")
    elif target == "published":
        source_path = run_workspace.final_md_target
        if not source_path.exists():
            raise RuntimeError("Published markdown is missing. Run assemble-subtopic-final --publish first.")
        text = read_text(source_path)
        page_count = read_docx_page_count(run_workspace.final_docx_target)
    else:
        raise RuntimeError("target must be one of: master, assembled, published")

    metrics = compute_text_metrics(text)
    payload = {
        "subtopic_id": run_workspace.subtopic_entry.item_id,
        "target": target,
        "source_path": str(source_path),
        "words": metrics["words"],
        "chars": metrics["chars"],
        "url1": metrics["url1"],
        "url2": metrics["url2"],
        "verified_url2": metrics["verified_url2"],
        "document_cards": metrics["document_cards"],
        "page_count": page_count,
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def cmd_capture_part_output(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    run_workspace = ensure_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
    )
    part_number = int(args.part_number)
    if part_number < 1 or part_number > 11:
        raise RuntimeError("part_number must be between 1 and 11")
    assert_run_command_allowed(run_workspace, "capture-part-output", part_number=part_number)

    content = load_part_output_source(args.source_file, args.clipboard)
    content = normalize_part_output(part_number, content)
    foreign_subtopic_ids = find_foreign_subtopic_ids(content, run_workspace.subtopic_entry.item_id)
    if foreign_subtopic_ids:
        raise RuntimeError(
            "Part output references a different subtopic id:\n- "
            + "\n- ".join(foreign_subtopic_ids)
        )
    issues = validate_part_output(run_workspace, part_number, content)
    if issues:
        update_run_manifest_part_status(
            run_workspace,
            part_number,
            "validation_failed",
            validation_issues=issues,
            validation_passed=False,
        )
        raise RuntimeError("Part output validation failed:\n- " + "\n- ".join(issues))

    output_path = run_workspace.stage_outputs_dir / f"part-{part_number:02d}.md"
    write_text(output_path, content.rstrip() + "\n")

    if part_number == 1:
        status = "completed_waiting_for_go"
    elif part_number == 2:
        status = "completed_core_answer"
    else:
        status = "completed"
    update_run_manifest_part_status(
        run_workspace,
        part_number,
        status,
        source_origin="capture_part_output",
        validation_issues=[],
        validation_passed=True,
    )
    if not run_workspace.is_surgical_redo:
        refresh_dynamic_part_packets(run_workspace)
    refresh_master_working_file(run_workspace)
    if part_number == 5:
        write_local_metric_checkpoint(run_workspace, "after_parts_02_05", target="master")
    elif part_number == 11:
        write_local_metric_checkpoint(run_workspace, "after_parts_02_11", target="master")

    if part_number == 2 and args.auto_assemble and not run_workspace.is_surgical_redo:
        assemble_subtopic_final(run_workspace, publish=False)

    print(output_path)
    return 0


def cmd_prepare_part_02_web(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    run_workspace = ensure_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
    )
    assert_run_command_allowed(run_workspace, "prepare-part-02-web", part_number=2)
    plan_paths = write_part_02_web_plan(run_workspace, overwrite=args.force)
    reasoning_paths = None if run_workspace.is_surgical_redo else write_reasoning_layer(run_workspace, overwrite=args.force)
    if not run_workspace.is_surgical_redo:
        refresh_dynamic_part_packets(run_workspace)
    update_run_manifest_web_plan(run_workspace, plan_paths)
    if reasoning_paths is not None:
        update_run_manifest_reasoning_layer(run_workspace, reasoning_paths)
    if PRODUCTION_ENABLE_SEMANTIC_DEDUP and not run_workspace.is_surgical_redo:
        semantic_dedup_paths = write_semantic_dedup_layer(run_workspace, overwrite=args.force)
        update_run_manifest_semantic_dedup(run_workspace, semantic_dedup_paths)
    if PRODUCTION_ENABLE_OMISSION_AUDIT and not run_workspace.is_surgical_redo:
        omission_paths = write_omission_audit_layer(run_workspace, overwrite=args.force)
        update_run_manifest_omission_audit(run_workspace, omission_paths)
    drop_disabled_manifest_layers(run_workspace)
    print(plan_paths["web_plan_dir"])
    return 0


def cmd_prepare_part_03_plan(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    run_workspace = ensure_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
    )
    assert_run_command_allowed(run_workspace, "prepare-part-03-plan", part_number=3)
    plan_paths = write_part_03_plan(run_workspace, overwrite=args.force)
    reasoning_paths = None if run_workspace.is_surgical_redo else write_reasoning_layer(run_workspace, overwrite=args.force)
    if not run_workspace.is_surgical_redo:
        refresh_dynamic_part_packets(run_workspace)
    update_run_manifest_part_03_plan(run_workspace, plan_paths)
    if reasoning_paths is not None:
        update_run_manifest_reasoning_layer(run_workspace, reasoning_paths)
    if PRODUCTION_ENABLE_SEMANTIC_DEDUP and not run_workspace.is_surgical_redo:
        semantic_dedup_paths = write_semantic_dedup_layer(run_workspace, overwrite=args.force)
        update_run_manifest_semantic_dedup(run_workspace, semantic_dedup_paths)
    if PRODUCTION_ENABLE_OMISSION_AUDIT and not run_workspace.is_surgical_redo:
        omission_paths = write_omission_audit_layer(run_workspace, overwrite=args.force)
        update_run_manifest_omission_audit(run_workspace, omission_paths)
    drop_disabled_manifest_layers(run_workspace)
    print(plan_paths["plan_dir"])
    return 0


def cmd_capture_part_03_range(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    run_workspace = ensure_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
    )
    assert_run_command_allowed(run_workspace, "capture-part-03-range", part_number=3)
    segment = get_part_03_segment(int(args.segment_id))
    content = load_part_output_source(args.source_file, args.clipboard)
    content = normalize_part_output(3, content)
    foreign_subtopic_ids = find_foreign_subtopic_ids(content, run_workspace.subtopic_entry.item_id)
    if foreign_subtopic_ids:
        raise RuntimeError(
            "Part 3 range references a different subtopic id:\n- "
            + "\n- ".join(foreign_subtopic_ids)
        )
    issues = validate_part_03_segment_output(content)
    if issues:
        raise RuntimeError("Part 3 range validation failed:\n- " + "\n- ".join(issues))

    segment_output_path = get_part_03_segment_output_path(run_workspace, segment["segment_id"])
    write_text(segment_output_path, content.rstrip() + "\n")

    _, captured, remaining = rebuild_part_03_aggregated_output(run_workspace)
    update_part_03_capture_status(run_workspace, captured, remaining)
    if remaining:
        update_run_manifest_part_status(
            run_workspace,
            3,
            f"in_progress_segments_{len(captured)}_of_{len(PART_03_SEGMENTS)}",
            source_origin="capture_part_03_range",
        )
    else:
        update_run_manifest_part_status(
            run_workspace,
            3,
            "completed",
            source_origin="capture_part_03_range",
        )
    if not run_workspace.is_surgical_redo:
        refresh_dynamic_part_packets(run_workspace)
    if args.auto_assemble and not run_workspace.is_surgical_redo:
        assemble_subtopic_final(run_workspace, publish=False)
    print(segment_output_path)
    return 0


def cmd_prepare_part_04_plan(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    run_workspace = ensure_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
    )
    assert_run_command_allowed(run_workspace, "prepare-part-04-plan", part_number=4)
    plan_paths = write_part_04_plan(run_workspace, overwrite=args.force)
    update_run_manifest_part_04_plan(run_workspace, plan_paths)
    print(plan_paths["plan_dir"])
    return 0


def cmd_capture_part_04_range(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    run_workspace = ensure_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
    )
    assert_run_command_allowed(run_workspace, "capture-part-04-range", part_number=4)
    segment = get_part_04_segment(int(args.segment_id))
    content = load_part_output_source(args.source_file, args.clipboard)
    content = normalize_part_output(4, content)
    foreign_subtopic_ids = find_foreign_subtopic_ids(content, run_workspace.subtopic_entry.item_id)
    if foreign_subtopic_ids:
        raise RuntimeError(
            "Part 4 range references a different subtopic id:\n- "
            + "\n- ".join(foreign_subtopic_ids)
        )
    issues = validate_part_04_segment_output(content)
    if issues:
        raise RuntimeError("Part 4 range validation failed:\n- " + "\n- ".join(issues))

    segment_output_path = get_part_04_segment_output_path(run_workspace, segment["segment_id"])
    write_text(segment_output_path, content.rstrip() + "\n")

    _, captured, remaining = rebuild_part_04_aggregated_output(run_workspace)
    update_part_04_capture_status(run_workspace, captured, remaining)
    if remaining:
        update_run_manifest_part_status(
            run_workspace,
            4,
            f"in_progress_segments_{len(captured)}_of_{len(PART_04_SEGMENTS)}",
            source_origin="capture_part_04_range",
        )
    else:
        update_run_manifest_part_status(
            run_workspace,
            4,
            "completed",
            source_origin="capture_part_04_range",
        )
    refresh_dynamic_part_packets(run_workspace)
    if args.auto_assemble:
        assemble_subtopic_final(run_workspace, publish=False)
    print(segment_output_path)
    return 0


def cmd_prepare_part_05_plan(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    run_workspace = ensure_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
    )
    assert_run_command_allowed(run_workspace, "prepare-part-05-plan", part_number=5)
    plan_paths = write_part_05_plan(run_workspace, overwrite=args.force)
    update_run_manifest_part_05_plan(run_workspace, plan_paths)
    print(plan_paths["plan_dir"])
    return 0


def cmd_capture_part_05_range(args: argparse.Namespace) -> int:
    workspace_root = Path(args.workspace_root).resolve()
    run_workspace = ensure_subtopic_run_workspace(
        workspace_root=workspace_root,
        subtopic_id=args.subtopic_id,
        theme_query=args.theme_query,
    )
    assert_run_command_allowed(run_workspace, "capture-part-05-range", part_number=5)
    segment = get_part_05_segment(int(args.segment_id))
    content = load_part_output_source(args.source_file, args.clipboard)
    content = normalize_part_output(5, content)
    foreign_subtopic_ids = find_foreign_subtopic_ids(content, run_workspace.subtopic_entry.item_id)
    if foreign_subtopic_ids:
        raise RuntimeError(
            "Part 5 range references a different subtopic id:\n- "
            + "\n- ".join(foreign_subtopic_ids)
        )
    issues = validate_part_05_segment_output(content)
    if issues:
        raise RuntimeError("Part 5 range validation failed:\n- " + "\n- ".join(issues))

    segment_output_path = get_part_05_segment_output_path(run_workspace, segment["segment_id"])
    write_text(segment_output_path, content.rstrip() + "\n")

    _, captured, remaining = rebuild_part_05_aggregated_output(run_workspace)
    update_part_05_capture_status(run_workspace, captured, remaining)
    if remaining:
        update_run_manifest_part_status(
            run_workspace,
            5,
            f"in_progress_segments_{len(captured)}_of_{len(PART_05_SEGMENTS)}",
            source_origin="capture_part_05_range",
        )
    else:
        update_run_manifest_part_status(
            run_workspace,
            5,
            "completed",
            source_origin="capture_part_05_range",
        )
    refresh_dynamic_part_packets(run_workspace)
    if args.auto_assemble:
        assemble_subtopic_final(run_workspace, publish=False)
    print(segment_output_path)
    return 0


def cmd_init_workspace(args: argparse.Namespace) -> int:
    root = Path(args.path).resolve()
    for relative in ["imports", "runs"]:
        (root / relative).mkdir(parents=True, exist_ok=True)
    return 0


def cmd_import_topic(args: argparse.Namespace) -> int:
    source_dir = Path(args.source).resolve()
    dest_root = Path(args.dest_root).resolve()
    imported = import_topic(source_dir, dest_root)
    print(imported)
    return 0


def cmd_import_tree(args: argparse.Namespace) -> int:
    source_root = Path(args.source_root).resolve()
    dest_root = Path(args.dest_root).resolve()
    found = scan_source_tree(source_root)
    for path in found:
        imported = import_topic(path, dest_root)
        print(imported)
    return 0


def cmd_run_topic(args: argparse.Namespace) -> int:
    topic_dir = Path(args.topic_dir).resolve()
    config = load_json_config(Path(args.config).resolve() if args.config else None)
    run_dir = run_single_topic(
        topic_dir=topic_dir,
        config=config,
        dry_run=args.dry_run,
        include_existing_notes=args.include_existing_notes,
    )
    print(run_dir)
    return 0


def cmd_batch_run(args: argparse.Namespace) -> int:
    imports_root = Path(args.imports_root).resolve()
    config = load_json_config(Path(args.config).resolve() if args.config else None)
    topics = collect_imported_topics(imports_root)
    for topic_dir in topics:
        run_dir = run_single_topic(
            topic_dir=topic_dir,
            config=config,
            dry_run=args.dry_run,
            include_existing_notes=args.include_existing_notes,
        )
        print(run_dir)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="md-first notary internship agent")
    subparsers = parser.add_subparsers(dest="command", required=True)

    draft_main_theme_outline = subparsers.add_parser(
        "draft-main-theme-outline",
        help="Prepare a draft/review packet for one main theme outline without rendering order copies",
    )
    draft_main_theme_outline.add_argument("theme_query")
    draft_main_theme_outline.add_argument("--workspace-root", default=".")
    draft_main_theme_outline.set_defaults(func=cmd_draft_main_theme_outline)

    prepare_main_theme = subparsers.add_parser(
        "prepare-main-theme",
        help="Build a workflow-first packet for one main theme from Утверждаю",
    )
    prepare_main_theme.add_argument("theme_query")
    prepare_main_theme.add_argument("--workspace-root", default=".")
    prepare_main_theme.add_argument("--packet-mode", choices=[PACKET_MODE_LEAN, PACKET_MODE_LITERAL], default=DEFAULT_PACKET_MODE)
    prepare_main_theme.add_argument("--packet-mode-reason", default="")
    prepare_main_theme.set_defaults(func=cmd_prepare_main_theme)

    run_subtopic = subparsers.add_parser(
        "run-subtopic",
        help="Prepare a staged run workspace for one final subtopic from a rendered order copy",
    )
    run_subtopic.add_argument("subtopic_id")
    run_subtopic.add_argument("--theme-query")
    run_subtopic.add_argument("--workspace-root", default=".")
    run_subtopic.add_argument("--full-artifacts", action="store_true")
    run_subtopic.add_argument("--packet-mode", choices=[PACKET_MODE_LEAN, PACKET_MODE_LITERAL], default=DEFAULT_PACKET_MODE)
    run_subtopic.add_argument("--packet-mode-reason", default="")
    run_subtopic.set_defaults(func=cmd_run_subtopic)

    prepare_surgical_redo = subparsers.add_parser(
        "prepare-surgical-redo",
        help="Lock the latest warm run into strict surgical-redo mode for explicitly allowed parts only",
    )
    prepare_surgical_redo.add_argument("subtopic_id")
    prepare_surgical_redo.add_argument("--parts", required=True, help="Comma-separated allowed parts, for example 2,3")
    prepare_surgical_redo.add_argument("--reason", default="")
    prepare_surgical_redo.add_argument("--theme-query")
    prepare_surgical_redo.add_argument("--workspace-root", default=".")
    prepare_surgical_redo.set_defaults(func=cmd_prepare_surgical_redo)

    execute_part_01 = subparsers.add_parser(
        "execute-part-01",
        help="Generate the standard Part 01 response and save it into the latest run workspace",
    )
    execute_part_01.add_argument("subtopic_id")
    execute_part_01.add_argument("--theme-query")
    execute_part_01.add_argument("--workspace-root", default=".")
    execute_part_01.add_argument("--force", action="store_true")
    execute_part_01.set_defaults(func=cmd_execute_part_01)

    assemble_subtopic_final_parser = subparsers.add_parser(
        "assemble-subtopic-final",
        help="Assemble the final md/docx from staged part outputs in the latest run workspace",
    )
    assemble_subtopic_final_parser.add_argument("subtopic_id")
    assemble_subtopic_final_parser.add_argument("--theme-query")
    assemble_subtopic_final_parser.add_argument("--workspace-root", default=".")
    assemble_subtopic_final_parser.add_argument("--publish", action="store_true")
    assemble_subtopic_final_parser.set_defaults(func=cmd_assemble_subtopic_final)

    metric_check = subparsers.add_parser(
        "metric-check",
        help="Compute local metrics for master, assembled or published markdown/docx without using model tokens",
    )
    metric_check.add_argument("subtopic_id")
    metric_check.add_argument("--theme-query")
    metric_check.add_argument("--workspace-root", default=".")
    metric_check.add_argument("--target", choices=["master", "assembled", "published"], default="master")
    metric_check.set_defaults(func=cmd_metric_check)

    capture_part_output = subparsers.add_parser(
        "capture-part-output",
        help="Capture one staged part output from a file or clipboard into the latest run workspace",
    )
    capture_part_output.add_argument("subtopic_id")
    capture_part_output.add_argument("part_number", type=int)
    capture_part_output.add_argument("--theme-query")
    capture_part_output.add_argument("--workspace-root", default=".")
    capture_part_output.add_argument("--source-file")
    capture_part_output.add_argument("--clipboard", action="store_true")
    capture_part_output.add_argument("--no-auto-assemble", dest="auto_assemble", action="store_false")
    capture_part_output.set_defaults(auto_assemble=True)
    capture_part_output.set_defaults(func=cmd_capture_part_output)

    prepare_part_02_web = subparsers.add_parser(
        "prepare-part-02-web",
        help="Prepare or refresh the web-first Part 02 research pack in the latest run workspace",
    )
    prepare_part_02_web.add_argument("subtopic_id")
    prepare_part_02_web.add_argument("--theme-query")
    prepare_part_02_web.add_argument("--workspace-root", default=".")
    prepare_part_02_web.add_argument("--force", action="store_true")
    prepare_part_02_web.set_defaults(func=cmd_prepare_part_02_web)

    prepare_part_03_plan = subparsers.add_parser(
        "prepare-part-03-plan",
        help="Prepare the segmented operator plan for Part 03 ranges in the latest run workspace",
    )
    prepare_part_03_plan.add_argument("subtopic_id")
    prepare_part_03_plan.add_argument("--theme-query")
    prepare_part_03_plan.add_argument("--workspace-root", default=".")
    prepare_part_03_plan.add_argument("--force", action="store_true")
    prepare_part_03_plan.set_defaults(func=cmd_prepare_part_03_plan)

    capture_part_03_range = subparsers.add_parser(
        "capture-part-03-range",
        help="Capture one Part 03 range response and append it into the aggregated part-03 output",
    )
    capture_part_03_range.add_argument("subtopic_id")
    capture_part_03_range.add_argument("segment_id", type=int)
    capture_part_03_range.add_argument("--theme-query")
    capture_part_03_range.add_argument("--workspace-root", default=".")
    capture_part_03_range.add_argument("--source-file")
    capture_part_03_range.add_argument("--clipboard", action="store_true")
    capture_part_03_range.add_argument("--no-auto-assemble", dest="auto_assemble", action="store_false")
    capture_part_03_range.set_defaults(auto_assemble=True)
    capture_part_03_range.set_defaults(func=cmd_capture_part_03_range)

    prepare_part_04_plan = subparsers.add_parser(
        "prepare-part-04-plan",
        help="Prepare the segmented operator plan for Part 04 ranges in the latest run workspace",
    )
    prepare_part_04_plan.add_argument("subtopic_id")
    prepare_part_04_plan.add_argument("--theme-query")
    prepare_part_04_plan.add_argument("--workspace-root", default=".")
    prepare_part_04_plan.add_argument("--force", action="store_true")
    prepare_part_04_plan.set_defaults(func=cmd_prepare_part_04_plan)

    capture_part_04_range = subparsers.add_parser(
        "capture-part-04-range",
        help="Capture one Part 04 range response and append it into the aggregated part-04 output",
    )
    capture_part_04_range.add_argument("subtopic_id")
    capture_part_04_range.add_argument("segment_id", type=int)
    capture_part_04_range.add_argument("--theme-query")
    capture_part_04_range.add_argument("--workspace-root", default=".")
    capture_part_04_range.add_argument("--source-file")
    capture_part_04_range.add_argument("--clipboard", action="store_true")
    capture_part_04_range.add_argument("--no-auto-assemble", dest="auto_assemble", action="store_false")
    capture_part_04_range.set_defaults(auto_assemble=True)
    capture_part_04_range.set_defaults(func=cmd_capture_part_04_range)

    prepare_part_05_plan = subparsers.add_parser(
        "prepare-part-05-plan",
        help="Prepare the segmented operator plan for Part 05 ranges in the latest run workspace",
    )
    prepare_part_05_plan.add_argument("subtopic_id")
    prepare_part_05_plan.add_argument("--theme-query")
    prepare_part_05_plan.add_argument("--workspace-root", default=".")
    prepare_part_05_plan.add_argument("--force", action="store_true")
    prepare_part_05_plan.set_defaults(func=cmd_prepare_part_05_plan)

    capture_part_05_range = subparsers.add_parser(
        "capture-part-05-range",
        help="Capture one Part 05 range response and append it into the aggregated part-05 output",
    )
    capture_part_05_range.add_argument("subtopic_id")
    capture_part_05_range.add_argument("segment_id", type=int)
    capture_part_05_range.add_argument("--theme-query")
    capture_part_05_range.add_argument("--workspace-root", default=".")
    capture_part_05_range.add_argument("--source-file")
    capture_part_05_range.add_argument("--clipboard", action="store_true")
    capture_part_05_range.add_argument("--no-auto-assemble", dest="auto_assemble", action="store_false")
    capture_part_05_range.set_defaults(auto_assemble=True)
    capture_part_05_range.set_defaults(func=cmd_capture_part_05_range)

    init_workspace = subparsers.add_parser("init-workspace", help="Create base workspace folders")
    init_workspace.add_argument("path", nargs="?", default=".")
    init_workspace.set_defaults(func=cmd_init_workspace)

    import_topic_parser = subparsers.add_parser("import-topic", help="Import one topic folder into md-first format")
    import_topic_parser.add_argument("source")
    import_topic_parser.add_argument("--dest-root", default="imports")
    import_topic_parser.set_defaults(func=cmd_import_topic)

    import_tree_parser = subparsers.add_parser("import-tree", help="Recursively import topic folders")
    import_tree_parser.add_argument("source_root")
    import_tree_parser.add_argument("--dest-root", default="imports")
    import_tree_parser.set_defaults(func=cmd_import_tree)

    run_topic_parser = subparsers.add_parser("run-topic", help="Run one imported topic")
    run_topic_parser.add_argument("topic_dir")
    run_topic_parser.add_argument("--config")
    run_topic_parser.add_argument("--dry-run", action="store_true")
    run_topic_parser.add_argument("--include-existing-notes", action="store_true")
    run_topic_parser.set_defaults(func=cmd_run_topic)

    batch_run_parser = subparsers.add_parser("batch-run", help="Run all imported topics sequentially")
    batch_run_parser.add_argument("imports_root", nargs="?", default="imports")
    batch_run_parser.add_argument("--config")
    batch_run_parser.add_argument("--dry-run", action="store_true")
    batch_run_parser.add_argument("--include-existing-notes", action="store_true")
    batch_run_parser.set_defaults(func=cmd_batch_run)

    return parser


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return int(args.func(args))
    except Exception as exc:  # pragma: no cover - top-level CLI guard
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
