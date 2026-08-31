# Implementation Plan: Apply AI → sync списка + red/yellow на сверке

**Спека**: [ai_docs/specs/kp-review-apply-sync-and-highlights.md](../../specs/kp-review-apply-sync-and-highlights.md)  
**Идея**: [ai_docs/ideas/kp-review-apply-sync-and-highlights.md](../../ideas/kp-review-apply-sync-and-highlights.md)  
**Parent (wait + AI)**: [2026-08-31-kp-ocr-wait-and-ai-on-review.md](./2026-08-31-kp-ocr-wait-and-ai-on-review.md) · [kp-ocr-wait-and-ai-on-review.md](../../specs/kp-ocr-wait-and-ai-on-review.md)  
**Related**: [unparsed-line-live-highlight.md](../../specs/unparsed-line-live-highlight.md) · multi-page [2026-08-31-kp-multi-page-screenshots.md](./2026-08-31-kp-multi-page-screenshots.md)  
**Дата**: 2026-08-31  
**Статус**: PLAN ✅ · IMPLEMENT ✅

## Overview

Follow-on to Phase A.3 / A.3.1: (1) after successful «Применить инструкцию», hydrate draft **and** write updated text into the review editor + store so the list beside the photo updates; visible error on failure; (2) batch-review list uses existing PlateListEditor / source lint — **red** = parser reject / unparsed, **yellow** = soft signals already wired — no new н→п heuristic, no highlight of parser-accepted lines (e.g. `8н`), no OCR pipeline auto suffix rewrite.

## Architecture Decisions

- **Apply sync** — post-success path must update the same surfaces the manager sees next to the image (review editor + per-page / multi-page store), not only draft hydrate. R10 already forbids reset; this plan adds **visible text sync**.
- **Apply error** — use existing product error surface (toast / inline); never silent failure.
- **Highlights** — wire / reuse PlateListEditor + source lint on batch-review; **do not** add load-suffix `н` yellow heuristic; **do not** change OCR pipeline.

## Task List

### Phase A.4 — Apply sync + review highlights

#### Task 1: Apply AI → editor/store sync + tests

**Description:** After successful Apply AI on batch-review: hydrate draft **and** write the updated batch / active-page text into the review editor and store so the list beside the photo changes immediately. On Apply failure: show a **visible** error (not silence). Preserve R10 (no `resetSource` / multi-session wipe when `hasStarted`).

**Acceptance:**
- [x] S1: Apply success → review editor + store text updated; list beside image shows new text
- [x] S2: Apply failure → visible error
- [x] R10 intact: multi-session pages / `hasStarted` unchanged on Apply
- [x] Tests cover success sync + failure visibility

**Verification:** unit / RTL on wizard / review chrome / apply handler  
**Dependencies:** parent A.3.1 (R10–R11) done  
**Files (orientir):** `CommercialOfferWizard.tsx`, apply-AI handlers, multi-page store / page text sync, InputStep / review editor wiring, tests  
**Estimated scope:** M

#### Task 2: Red/yellow highlights on review via existing lint + tests

**Description:** Ensure batch-review list uses existing PlateListEditor highlights / source lint: **red** = parser reject / unparsed; **yellow** = other soft signals already in product (e.g. OCR corrections) **if already wired**. Do **not** add a new load-suffix `н` heuristic. If the parser accepts a line (e.g. `8н`), do **not** highlight it. No OCR pipeline auto н→п.

**Acceptance:**
- [x] S3: unparsed / parser-reject lines → red on review
- [x] S4: yellow only for soft signals already wired; no new н-suffix heuristic
- [x] S5: parser-accepted `8н` (or equivalent) → not highlighted
- [x] S6: no OCR pipeline auto suffix rewrite
- [x] S7: `npm run test -- src/features/commercial-offer` + typecheck green

**Verification:** RTL on review list editor / lint wiring; code review for no pipeline suffix rule  
**Dependencies:** Task 1 preferred first (shared review surface), or parallel if lint path isolated  
**Files (orientir):** `PlateListEditor` (or review list chrome), source lint hooks, `*InputStep` / review props, tests; **no** OCR pipeline suffix rules  
**Estimated scope:** M

### Checkpoint: Phase A.4 done

- [x] S1–S7 green
- [x] Spec status → IMPLEMENT ✅
- [x] No auto suffix rewrite; no highlight-all; no new AI model; no layout redesign

## Out of this plan

- Auto suffix rewrite in OCR pipeline  
- Yellow heuristic for н→п on accepted lines  
- Highlight all lines  
- New AI model / endpoint  
- Layout redesign of review / lightbox / progressive  
- Commit unless explicitly asked

## Open for implementer

- Confirm which soft yellow signals are already wired on review — wire those only  
- Confirm whether draft-level Apply requires refreshing all ready page editors or active page only — match API result shape
