---
name: notary-outline-builder
description: Use when working on a main notarial internship theme from Утверждаю.md and you need to build or validate the final Оглавление, especially when tariff subtopics must be split into separate semantic items like 16.1.1 and 16.1.2.
---

# Notary Outline Builder

Build the final Оглавление for one main theme before rendering any приказ copies.

Use this skill when:
- the user gives a main theme such as `Тема 16`;
- `Утверждаю.md` contains broad source rows that must be split into final subtopics;
- tariff language must be rewritten into a context-specific subtopic;
- an existing auto-generated outline must be checked against an approved one.

## Inputs

Read in this order:

1. `C:\Users\koper\Downloads\GitHub Projects\Notary Qwen\AGENTS.md`
2. `C:\Users\koper\Downloads\GitHub Projects\Notary Qwen\PROJECT_STATE.md`
3. `C:\Users\koper\Downloads\GitHub Projects\Notary Qwen\input\workflow\manual-workflow.md`
4. `C:\Users\koper\Downloads\GitHub Projects\Notary Qwen\input\workflow\Утверждаю.md`
5. `C:\Users\koper\Downloads\GitHub Projects\Notary Qwen\input\workflow\outline overrides\<номер темы>.md` if it exists

## Workflow

1. Identify the main theme block in `Утверждаю.md`.
2. Treat the rows in `Утверждаю.md` as source material, not automatically as final outline items.
3. If an approved outline override exists, use it as the final source of truth.
4. If no override exists, build a draft outline semantically:
   - preserve legal meaning;
   - split tariff language into a separate item when needed;
   - use numbering like `16.1.1`, `16.1.2`, `16.7.1` ... `16.7.4`;
   - do not leave a naked tariff line without context.
5. Save:
   - draft outline as `outline.generated.md`;
   - approved outline as `outline.active.md`;
   - override copy as `outline.approved.md` when provided.

## Hard Rules

- Do not keep flat numbering like `16.1`, `16.2` when the user-approved outline is hierarchical.
- `Исчисление размера федерального и регионального тарифа` must be rewritten as a context-specific item.
- The final active outline must be the one used by downstream order rendering and execution.
- If an approved outline exists, do not invent a competing structure.

## Outputs

- `outline.generated.md` = draft or fallback auto-outline
- `outline.active.md` = actual outline that downstream steps must use
- `outline.approved.md` = copied approved override when present

## Failure Conditions

Stop and ask for correction if:
- the main theme cannot be located in `Утверждаю.md`;
- numbering in the approved outline is inconsistent;
- source rows and approved outline clearly conflict and the conflict cannot be reconciled mechanically.
