---
name: notary-order-renderer
description: Use when you need to render final приказ copies for each approved notarial subtopic by filling Section III and replacing all current-topic references in the 11 parts without changing the canonical order template.
---

# Notary Order Renderer

Create one приказ copy per final subtopic from the approved outline.

Use this skill when:
- `outline.active.md` already exists;
- the canonical order template must remain untouched;
- each final subtopic needs its own rendered order file;
- all references to the current topic inside the 11 parts must be updated.

## Inputs

Read in this order:

1. `C:\Users\koper\Downloads\GitHub Projects\Notary Qwen\AGENTS.md`
2. `C:\Users\koper\Downloads\GitHub Projects\Notary Qwen\PROJECT_STATE.md`
3. `C:\Users\koper\Downloads\GitHub Projects\Notary Qwen\input\order\Текст приказа 18  15.11.10..md`
4. `outline.active.md`
5. current main theme and current final subtopic

## Workflow

1. Read the canonical order template as a read-only source.
2. Replace only the mutable layer:
   - Section III `ОГЛАВЛЕНИЕ`;
   - current focus subtopic references inside the 11 parts.
3. Keep the full approved outline in Section III.
4. Create one file per final subtopic using short filenames such as:
   - `16.1.1.md`
   - `16.1.2.md`
5. Build a parallel execution packet for the same subtopic.

## Hard Rules

- Never edit the source order template in place.
- Never use long descriptive filenames for final order copies.
- The active outline is global for the main theme; the focus subtopic is local for the current rendered file.
- Replace all occurrences of the old focus subtopic, not only the first visible heading.
- Keep the fixed text of the order intact.

## Outputs

- rendered order copies in `02-orders`
- execution packets in `03-packets`
- updated manifest entries pointing to the new files

## Validation

Check before finishing:

- Section III shows the approved outline, not the raw `Утверждаю` rows.
- the current order file focuses on the correct final subtopic;
- the old template topic `15.11.10` is no longer present in the rendered copy;
- filenames match the final subtopic ids exactly.
