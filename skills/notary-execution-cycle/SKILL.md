---
name: notary-execution-cycle
description: Use when executing one final notarial subtopic end-to-end through the order parts, preserving the GO/START gate, strict link formatting, anti-repeat follow-ups, quarantine handling, delta-audit, mini-summary, and diary-task output.
---

# Notary Execution Cycle

Execute one final subtopic using the exact staged rhythm defined by the project.

Use this skill when:
- a rendered order copy already exists for one final subtopic;
- you need to follow the order parts instead of producing one monolithic answer;
- the result must preserve the style captured in `User and LLM Interaction in LLM Contest Window.md`.

## Inputs

Read in this order:

1. `C:\Users\koper\OneDrive\Documents\New project\AGENTS.md`
2. `C:\Users\koper\OneDrive\Documents\New project\PROJECT_STATE.md`
3. `C:\Users\koper\OneDrive\Documents\New project\input\master prompt\Промпт по поиску документов 18.md`
4. rendered order copy for the current final subtopic
5. `C:\Users\koper\OneDrive\Documents\New project\input\workflow\User and LLM Interaction in LLM Contest Window.md`
6. accepted output examples

## Execution Rhythm

Follow the order parts as staged modes:

1. Part 1:
   - confirm understanding of rules only;
   - stop at `ЖДУ СИГНАЛ GO`.
2. Part 2:
   - start with `ТЕМА: ...`;
   - produce `АНАЛИЗ ОБЛАСТИ ПРАВА`;
   - then `A. РЕГУЛЯТОРНОЕ ЯДРО` -> `B. ОПОРНЫЕ ДОКУМЕНТЫ` -> `КАРАНТИН` -> `FAIL-SAFE CHECK`.
3. Parts 3-9:
   - perform targeted expansion and gap closure;
   - keep anti-repeat discipline;
   - use execution-only mode when required;
   - do not restart from `ТЕМА` unless the part requires a fresh full answer.
4. Part 10:
   - create the expanded mini-summary with supporting links after each point.
5. Part 11:
   - create the practical-actions text for the 4th diary column only;
   - do not repeat the legal document list.

## Hard Rules

- Do not skip the `GO/СТАРТ` gate.
- Do not collapse the whole workflow into one answer if the staged logic matters.
- Do not output anchor links as final results.
- All URLs and link-like elements must be inside code blocks only.
- `URL1` is mandatory even when unreadable.
- `URL2` must be verified or replaced with `отсутствует (КАРАНТИН)` according to the strategy.
- A document may not be excluded solely because `URL2` is missing.
- Follow-up parts must not duplicate already-issued documents inside the same theme unless the instruction explicitly requires re-checking.

## Output Discipline

- preserve the accepted structure and tone from the interaction guide;
- preserve the strict formatting contract for links;
- keep quarantine and fail-safe explicit;
- separate legal-search output from mini-summary output and diary-task output.

## Stop Conditions

Stop and escalate if:
- link formatting would be invalid;
- a part instruction conflicts with the approved order copy;
- the current execution packet does not identify the active final subtopic clearly.
