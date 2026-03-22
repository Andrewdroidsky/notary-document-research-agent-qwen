# Operator Sequence: 16.1.1. Свидетельствование верности копий документов, копий с копий документов и выписок из них.

Это операторская последовательность для внешней LLM-сессии. Она нужна, пока агент еще не отправляет сообщения в LLM сам.

## Что отправлять по порядку

1. Сообщение 1: содержимое файла `C:\Users\koper\OneDrive\Documents\New project\input\output ready\Тема-16.-Особенности-совершения-отдельных-видов-нотариальных-действий\05-runs\16.1.1\20260321-221253\04-web-plan\message-01.part-01.md`
2. Сообщение 2: содержимое файла `C:\Users\koper\OneDrive\Documents\New project\input\output ready\Тема-16.-Особенности-совершения-отдельных-видов-нотариальных-действий\05-runs\16.1.1\20260321-221253\04-web-plan\message-02.go.txt`
3. Сообщение 3: содержимое файла `C:\Users\koper\OneDrive\Documents\New project\input\output ready\Тема-16.-Особенности-совершения-отдельных-видов-нотариальных-действий\05-runs\16.1.1\20260321-221253\04-web-plan\message-03.part-02-launch-packet.md` — только если после `GO/СТАРТ` LLM не начал сам выдавать Part 2.

## Что сделать после ответа LLM по Части 2

Скопировать ответ целиком в clipboard и выполнить команду:

```powershell
python .\notary_agent.py capture-part-output 16.1.1 2 --clipboard
```

## Проверка перед стартом

- Файл Part 1 должен быть не stub: `C:\Users\koper\OneDrive\Documents\New project\input\output ready\Тема-16.-Особенности-совершения-отдельных-видов-нотариальных-действий\05-runs\16.1.1\20260321-221253\02-stage-outputs\part-01.md`
- Launch packet уже собран: `C:\Users\koper\OneDrive\Documents\New project\input\output ready\Тема-16.-Особенности-совершения-отдельных-видов-нотариальных-действий\05-runs\16.1.1\20260321-221253\04-web-plan\part-02.launch-packet.md`
- Актуальный latest run: `C:\Users\koper\OneDrive\Documents\New project\input\output ready\Тема-16.-Особенности-совершения-отдельных-видов-нотариальных-действий\05-runs\16.1.1\20260321-221253`

## Что не делать

- Не склеивать все три сообщения в одно.
- Не пропускать сообщение `GO/СТАРТ` между Частью 1 и Частью 2.
- Если ответ на `GO/СТАРТ` уже начинается с `ТЕМА:` и разворачивает Part 2, не отправлять fallback `message-03.part-02-launch-packet.md` в эту же сессию.
- Не вставлять в LLM сырой `part-02.md` из stage-inputs вместо launch packet.
