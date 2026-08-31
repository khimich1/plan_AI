# Implementation Plan: Wait-баннер OCR + AI на экране сверки

**Спека**: [ai_docs/specs/kp-ocr-wait-and-ai-on-review.md](../../specs/kp-ocr-wait-and-ai-on-review.md)  
**Идея**: [ai_docs/ideas/kp-ocr-wait-and-ai-on-review.md](../../ideas/kp-ocr-wait-and-ai-on-review.md)  
**Parent multi-page**: [2026-08-31-kp-multi-page-screenshots.md](./2026-08-31-kp-multi-page-screenshots.md) · [kp-multi-page-screenshots.md](../../specs/kp-multi-page-screenshots.md)  
**Follow-on**: [2026-08-31-kp-review-apply-sync-and-highlights.md](./2026-08-31-kp-review-apply-sync-and-highlights.md) · [kp-review-apply-sync-and-highlights.md](../../specs/kp-review-apply-sync-and-highlights.md)  
**Дата**: 2026-08-31  
**Статус**: PLAN ✅ · IMPLEMENT ✅ Phase A.3 · Phase A.3.1 ✅ (R10–R11 closed 2026-08-31)

## Overview

Поверх готового multi-page progressive review: (1) баннер ожидания, пока `hasStarted` и ещё нет ни одной `ready`; (2) перенос AI-инструкции на batch-review для всех типов изделий. Без auto н→п в pipeline, без Phase B, без redesign lightbox/progressive.

## Architecture Decisions

- **Wait chrome** — условие из уже существующих флагов хука/wizard (`hasStarted`, page statuses). Не новый job/progress API.
- **AI on review** — reuse `aiInstruction` / `onApplyAi` wiring; поверхность UI на экране сверки во всех `*InputStep` (или общий review chrome), а не только в `SourceInputCard` «Дополнительно».
- **D-no-auto-suffix** — нулевой diff OCR pipeline по суффиксам; fix только UI.

## Task List

### Phase A.3 — Wait + AI on review

#### Task 1: Wait banner + spinner + tests

**Description:** Пока `hasStarted &&` нет страницы со статусом `ready` (и аналог для single-file path, если тот же gap), показать баннер + простой spinner с текстом «Идёт распознавание, подождите 1–2 минуты». Скрыть при первой `ready`. До `hasStarted` не показывать.

**Acceptance:**
- [x] W1: banner+spinner visible when started and no ready yet
- [x] W2: hidden after first ready; progressive review unchanged
- [x] W3: not shown before start

**Verification:** RTL on wizard / SourceInputCard / review shell as appropriate  
**Dependencies:** multi-page Phase A.2 done  
**Files (orientir):** `CommercialOfferWizard.tsx` and/or review chrome / `PageReviewNav` / `SourceInputCard`; tests  
**Estimated scope:** S

#### Task 2: Move AI UI to batch-review on all `*InputStep` + tests

**Description:** Показать блок AI instruction (textarea + «Применить инструкцию») на **batch-review** для plates, piles, marches, steps, bridge, fbs. Текст списка остаётся редактируемым. Reuse existing apply-AI handlers. Решить при IMPLEMENT: убрать/оставить дубль в «Дополнительно» (ask if unclear — default: surface on review; don’t break append).

**Acceptance:**
- [x] A1: AI UI available on review for all six product types
- [x] A2: list text still editable on review
- [x] A3: no OCR pipeline auto suffix rewrite
- [x] A4: commercial-offer tests + typecheck green

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer && npm run typecheck`  
**Dependencies:** Task 1 (можно параллельно после wait green)  
**Files (orientir):** six `*InputStep.tsx`, possibly shared review chrome, `SourceInputCard.tsx` (если убираем/сужаем «Дополнительно»), wizard props, tests  
**Estimated scope:** M

### Checkpoint: Phase A.3 done

- [x] W1–W3 + A1–A4 green
- [x] Spec status → IMPLEMENT ✅
- [x] No pipeline н→п; no Phase B; no lightbox/progressive redesign

### Phase A.3.1 — Review remediation (блокирует merge)

Источник: Code review 2026-08-31 в спеке (R10–R11, S20–S21). Companion batch helpers: R12–R13 в multi-page plan.

#### Task 3: R10 — Apply AI must not reset multi-session

**Description:** В `handleApplyAi` не вызывать `resetSource()` / `multiPage.reset()`, если `multiPage.hasStarted`. После AI — hydrate draft + sync active page text; pages/statuses сохраняются.

**Acceptance:**
- [x] S20: 2+ pages → Apply AI → pages.length unchanged, hasStarted true

**Verification:** wizard/hook test or RTL  
**Files:** `CommercialOfferWizard.tsx`, tests  
**Estimated scope:** S

#### Task 4: R11 — AI enabled during progressive OCR tail

**Description:** На review `AiInstructionBlock` не `disabled={isRecognizing}` по полному queue-busy. Достаточно `isAiProcessing` (+ наличие draft). Page1 ready + page2 pending → Apply enabled.

**Acceptance:**
- [x] S21 green

**Verification:** RTL InputStep / review chrome  
**Files:** six `*InputStep.tsx` (or shared prop), tests  
**Estimated scope:** S

### Checkpoint: Merge-ready (A.3.1)

- [x] R10–R11 + S20–S21 green
- [x] R12–R13 closed in multi-page plan (or explicitly deferred with risk accept)
- [x] Spec verdict → Approve

## Out of this plan

- Auto suffix rewrite in OCR pipeline  
- Phase B server job  
- Mandatory AI on every OCR  
- Redesign lightbox / progressive review  
- Commit unless explicitly asked

## Open for implementer

- ~~Дубль AI в «Дополнительно» vs только review~~ → **kept** AI in «Дополнительно» (append path) + surfaced on batch-review
- ~~ApplyAi scope: active page vs full draft~~ → **reuse current** draft-level `applyAi` wiring; R10 only changes post-apply reset
