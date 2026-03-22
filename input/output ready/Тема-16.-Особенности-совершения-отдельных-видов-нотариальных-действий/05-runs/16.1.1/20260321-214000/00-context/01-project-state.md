# Project State

Этот файл хранит текущее рабочее состояние проекта, чтобы длинная работа не зависела от истории окна чата и не ломалась после auto-compact.

## Назначение

- фиксировать постоянные правила проекта;
- фиксировать текущую тему и текущий этап работы;
- фиксировать согласованные пути к входным и выходным файлам;
- фиксировать, что уже сделано и что остается сделать;
- давать агенту одну точку восстановления контекста.

## Источники Истины

- Компас процесса: `C:\Users\koper\OneDrive\Documents\New project\input\workflow\manual-workflow.md`
- Мастер-промпт: `C:\Users\koper\OneDrive\Documents\New project\input\master prompt\Промпт по поиску документов 18.md`
- Стратегия взаимодействия в окне LLM: `C:\Users\koper\OneDrive\Documents\New project\input\workflow\User and LLM Interaction in LLM Contest Window.md`
- Workflow: `C:\Users\koper\OneDrive\Documents\New project\input\workflow\manual-workflow.md`
- Перечень тем: `C:\Users\koper\OneDrive\Documents\New project\input\workflow\Утверждаю.md`
- Шаблон приказа: `C:\Users\koper\OneDrive\Documents\New project\input\order\Текст приказа 18  15.11.10..md`
- Эталон результата `.md`: `C:\Users\koper\OneDrive\Documents\New project\input\output examples\15.11.10. Наследование имущества.md`
- Эталон результата `.docx`: `C:\Users\koper\OneDrive\Documents\New project\input\output examples\15.11.10. Наследование имущества.docx`

## Зафиксированные Правила

- `manual-workflow.md` = компас процесса.
- `Промпт по поиску документов 18.md` = стратегия поиска, верификации и формата.
- `Текст приказа 18  15.11.10..md` = тактика, операционный план и руководство исполнения.
- `User and LLM Interaction in LLM Contest Window.md` = эталон механики и стиля взаимодействия в окне LLM.
- `input\workflow\outline overrides\<номер темы>.md` = утвержденное Оглавление по теме, если оно уже подготовлено и должно иметь приоритет над черновой автогенерацией.
- Перед началом работы агент всегда читает мастер-промпт.
- Перед практическим исполнением агент также читает файл `User and LLM Interaction in LLM Contest Window.md`.
- Запуск идет от команды пользователя с номером и наименованием основной темы.
- Агент читает `Утверждаю.md` и работает по теме, указанной пользователем.
- Агент юридически осмысленно формирует Оглавление, а не дробит тему механически.
- Для служебных подпунктов формулировка должна достраиваться по смыслу конкретной подтемы.
- Исходный файл приказа не изменяется.
- Для каждой подтемы создается отдельная копия приказа.
- Если по теме есть approved outline override, агент обязан строить копии приказов по конечным подпунктам из него, а не по сырому перечню из `Утверждаю.md`.
- В проекте созданы локальные skills:
  - `skills\notary-outline-builder`
  - `skills\notary-order-renderer`
  - `skills\notary-execution-cycle`
- Эти skills пока существуют как project-local слой и описание процедур; они еще не установлены как глобальные Codex skills в `$CODEX_HOME\skills`.
- По Оглавлению действует двухэтапный режим:
  - `draft-main-theme-outline` = draft + review
  - `prepare-main-theme` = рендер приказов только при наличии approved outline
- По конечной подтеме теперь есть отдельный режим:
  - `run-subtopic <subtopic_id>` = staged run-workspace с разбиением копии приказа на Части 1-11
- В приказе меняются только Раздел III и заголовки подтем в 11 Частях.
- Готовые результаты сохраняются в `C:\Users\koper\OneDrive\Documents\New project\input\output ready`.
- Финальный результат должен формироваться в `.md` и `.docx`.
- Формат результата должен соответствовать утвержденным образцам.
- После завершения одной основной темы агент ждет подтверждения пользователя перед переходом к следующей теме.

## Порядок Чтения И Исполнения

Рабочий порядок для агента фиксируется так:

1. Прочитать `C:\Users\koper\OneDrive\Documents\New project\PROJECT_STATE.md`.
2. Прочитать `C:\Users\koper\OneDrive\Documents\New project\input\workflow\manual-workflow.md`.
3. Прочитать `C:\Users\koper\OneDrive\Documents\New project\input\master prompt\Промпт по поиску документов 18.md`.
4. Прочитать `C:\Users\koper\OneDrive\Documents\New project\input\workflow\Утверждаю.md`.
5. Найти основную тему, указанную пользователем.
6. Юридически осмысленно сформировать Оглавление по выбранной основной теме.
7. Прочитать шаблон `C:\Users\koper\OneDrive\Documents\New project\input\order\Текст приказа 18  15.11.10..md`.
8. Создать копии приказа под каждую подтему.
9. Вставить сформированное Оглавление в Раздел III и подставить названия подтем в 11 Частей.
10. Прочитать `C:\Users\koper\OneDrive\Documents\New project\input\workflow\User and LLM Interaction in LLM Contest Window.md` как образец механики исполнения.
11. Исполнить рабочий цикл по подтемам и сформировать итоговые `.md` и `.docx` в `C:\Users\koper\OneDrive\Documents\New project\input\output ready`.

## Текущий Статус

- Текущий режим: практическая сборка локального workflow-first агента.
- Текущая основная тема: `Тема 16. Особенности совершения отдельных видов нотариальных действий.`
- Текущий этап: собран контур подготовки темы, approved outline, копий приказов, execution packets и staged run-workspace по подтеме.
- Проверенный пилот: `python .\notary_agent.py run-subtopic 16.1.1`
- Следующее действие: перейти от подготовки run-workspace к автоматизированному исполнению цикла по Частям 1-11 и сборке финального `.md/.docx`.

## Как Использовать Этот Файл

Агент перед длительной работой может перечитывать этот файл, чтобы восстановить:

- что уже согласовано;
- какие файлы обязательны к чтению;
- куда сохранять результат;
- на каком этапе работа остановилась.

Пользователь может при необходимости просить обновить этот файл, если меняются правила, пути или формат результата.
