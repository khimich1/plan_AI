# Implementation Plan: КП — правка и удаление строки иконками

**Спека**: [ai_docs/specs/kp-row-edit-delete-icons.md](../../specs/kp-row-edit-delete-icons.md)  
**Идея**: [ai_docs/ideas/kp-row-edit-delete-icons.md](../../ideas/kp-row-edit-delete-icons.md)  
**Related**: [kp-multi-nomenclature-append.md](../../specs/kp-multi-nomenclature-append.md) (`DELETE .../lines/{line_id}`) · [unparsed-line-live-highlight.md](../../specs/unparsed-line-live-highlight.md)  
**Дата**: 2026-08-31  
**Статус**: PLAN ✅ · IMPLEMENT ✅

## Overview

На предпросмотре шага 1 (все 6 `Kp*PreviewPanel`, **после** сверки) и на шаге 3 (`CalculationResultStep`) — иконки карандаш-в-квадрате и мусорка. Qty PATCH без полного ILP. Марка = `source_text` через `spec.generate_preview` **этого** `product_type`; 0/1/N = замена строки. Невалидный ввод = 400, состав не трогаем. Wide/unpriced карточки as-is; карандаш не disable. Undo — клиентский снимок последней операции строки + тост ~8 с, не undo-last-batch.

## Architecture Decisions

- **Restore API — dedicated POST, not inverse PATCH.**  
  `POST /api/v1/commercial/drafts/{draft_id}/lines/restore`  
  Body: `{ "index": <int>, "lines": [ <order-line snapshots> ], "replace_line_ids": [] }`.  
  **Why (meets S7, simpler than inverse PATCH):** after DELETE the line is gone — inverse PATCH 404s and cannot restore. After 1→N, inverse PATCH `source_text` would **re-parse** (new `line_id`, price drift), not restore the snapshot. Qty-only undo uses inverse **PATCH qty** (same endpoint as edit) — no restore needed there. `replace_line_ids` lets undo-of-replace remove the new lines and splice the snapshot in **one** call (avoids N deletes + insert races).
- **Qty path** — mutate existing line `qty` / `line_total` / `weight`; `compute_totals` via `get_draft_details`. No `generate_preview`, no full-order ILP.
- **Mark path** — fragment `generate_preview` of **this** `product_type` only; splice stamped lines at the old index; scrub old `line_id` from `append_batches[].line_ids`, insert new ids in the same batch; copy `append_batch_id` onto new lines. Do **not** rewrite draft-level wide/unpriced metadata (cards as-is). Unparsed fragment → 400, no mute-delete.
- **Icons** — shared `LineRowActions` (inline SVG, ghost, `aria-label`). No npm/pip deps. Step 3: drop visible «Удалить».
- **Undo slot** — one client snapshot; toast ~8 s («Строка удалена» / «Количество изменено» / «Строка изменена» + «Отменить»). Reset on timeout, next row op, or leaving the wizard step. Not undo-last-batch.
- **line_id** — preview rows pass `lineId`; icons hidden when missing (same as step 3 today).

## Task List

### Phase 1 — Backend line PATCH + restore

#### Task 1: PATCH qty + tests

**Description:** `PATCH /drafts/{id}/lines/{line_id}` with `{ "qty": N }` updates that line only; same `line_id`; unit price unchanged; totals recomputed. 404 if missing. No `generate_preview`.

**Acceptance:**
- [x] S3: PATCH qty → qty and line/totals updated; `line_id` unchanged
- [x] 404 «Строка не найдена.» if line missing
- [x] qty &lt; 1 or empty body → 400, draft unchanged

**Verification:** pytest in `tests/test_commercial_draft_line_edit.py` (reuse append helpers)  
**Dependencies:** None  
**Files (orientir):** `app/schemas/commercial.py`, `app/api/v1/endpoints/commercial.py`, `app/services/commercial_draft_lifecycle.py`, `app/services/commercial_workflow_service.py`, `tests/test_commercial_draft_line_edit.py`  
**Estimated scope:** M

#### Task 2: PATCH source_text replace 1→1 / 1→N + invalid 400

**Description:** `{ "source_text": "..." }` runs fragment preview for the line’s `product_type`, replaces that line with 0..N stamped lines. Invalid/unparsed input → 400, order unchanged. `append_batches.line_ids` scrub old id, add new ids to the same batch.

**Acceptance:**
- [x] S4: valid 1→1 mark/price updated (new or reused stamped id OK; old id gone if identity changed)
- [x] S5: 1→2+ : old line removed, new lines at its index, new `line_id`s
- [x] S6: invalid `source_text` → 400, composition unchanged
- [x] `append_batches.line_ids` consistent (old id out, new ids in same batch)

**Verification:** pytest (plates 1→1 qty-in-text; piles or two-line fragment 1→N; garbage 400)  
**Dependencies:** Task 1 (shared PATCH endpoint)  
**Files (orientir):** lifecycle `patch_order_line`, `product_draft_config.get_spec`, tests  
**Estimated scope:** M

#### Task 3: POST restore + tests

**Description:** `POST .../lines/restore` inserts snapshot `lines` at `index` after optionally removing `replace_line_ids`. Restores `append_batches` membership via snapshot `append_batch_id` (recreate batch if delete dropped it). 404 if draft missing.

**Acceptance:**
- [x] S7: after DELETE, restore returns the same `line_id` at the old index
- [x] restore after 1→N (`replace_line_ids` = new ids) brings old line back
- [x] S8: existing undo-last-batch tests still pass (no behavior change)

**Verification:** pytest restore-after-delete + restore-after-replace; regress `test_delete_line_*` / `test_append_undo_last_*`  
**Dependencies:** Task 2  
**Files (orientir):** schema `CommercialRestoreLinesRequest`, lifecycle `restore_order_lines`, endpoint, tests  
**Estimated scope:** S

### Checkpoint: Backend API

- [x] PATCH qty / mark / 400 / restore pytest green
- [x] DELETE + undo-last-batch unchanged

### Phase 2 — Frontend icons + mutations + undo toast

#### Task 4: line_id on preview rows + LineRowActions + wire panels / step 3

**Description:** `buildKpPreviewRows` and typed `build*PreviewRows` expose `lineId` (+ source-text prefill). Shared `LineRowActions`: pencil-in-square + trash ghost icon buttons; inline qty + «как в списке»; Esc cancels without undo slot. Wire 6 preview panels + `CalculationResultStep` (remove «Удалить» label). Icons only when `lineId` present. Pencil not disabled by wide/unpriced flags.

**Acceptance:**
- [x] S1: preview (plates at least) + step 3 show edit/delete icons for rows with `line_id`
- [x] S2: step 3 has no button with visible text «Удалить»
- [x] S9: wide/unpriced sections still in plate panel; pencil not `disabled` because of those flags
- [x] RTL: icons, edit fields, delete callback, no icons without `line_id`

**Verification:** vitest `buildKpPreviewRows` + `LineRowActions` + `KpPlatePreviewPanel` + `CalculationResultStep`  
**Dependencies:** Task 1–3 API contract (client methods can land with this slice)  
**Files (orientir):** `LineRowActions.tsx`, `buildKpPreviewRows.ts`, `build*PreviewRows.ts`, 6 `Kp*PreviewPanel.tsx`, 6 `*InputStep.tsx`, `CalculationResultStep.tsx`, `commercialOfferApi.ts`  
**Estimated scope:** L (split if needed: shared component + plates first, then copy to other 5)

#### Task 5: Wizard mutations + undo toast

**Description:** Wizard: PATCH/DELETE/restore mutations; one last-op snapshot; toast ~8 s with «Отменить»; copy does not say «добавление». Delete/edit on step 1 **does not** `set-step` to result. Leaving the step clears the slot. Invalid PATCH surfaces error **on the row**, not silent.

**Acceptance:**
- [x] S6 UI: invalid mark → error at row
- [x] S7: trash → existing delete; toast; «Отменить» restores
- [x] Qty undo via inverse PATCH qty; delete/replace undo via POST restore
- [x] Undo-last-batch CTA untouched

**Verification:** RTL toast + undo; wizard/handler tests if present  
**Dependencies:** Task 4  
**Files (orientir):** `CommercialOfferWizard.tsx`, `useCommercialOfferWizard.ts`, `LineUndoToast` (or Alert slot), tests  
**Estimated scope:** M

### Checkpoint: Feature complete

- [x] S1–S10
- [x] `cd frontend && npm run test -- src/features/commercial-offer && npm run typecheck`
- [x] `pytest tests/test_commercial_draft_line_edit.py tests/test_commercial_draft_append.py` green; `test_commercial_web_flow.py` 3 pre-existing fails (schema/identity), not this feature
- [x] Spec checklists updated
- [x] No commit; do not revert existing uncommitted OCR / A.4 work

## Out of this plan

- Manual unit_price / empty new row / confirm-dialog / Excel editor
- Icons on batch-review `PlateListEditor`
- Server-side undo journal / persist across sessions
- Redesign wide / unpriced cards
- Full-order ILP on qty
- Disable pencil because wide/unpriced
- New npm/pip deps
- Commit unless explicitly asked
- `bot_archived`

## Open for implementer

- Toast placement: above composition table (preview Card) and above «Позиции» on step 3 — not `stepError` (that’s for wizard-level failures).
- 1→1 may mint a new `line_id` if identity key changes; tests should not require id stability on mark change.
