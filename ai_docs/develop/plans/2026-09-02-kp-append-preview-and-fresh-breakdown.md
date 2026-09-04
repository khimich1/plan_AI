# Implementation Plan: КП — видимый состав при дописи + свежий breakdown

**Спека**: [ai_docs/specs/kp-append-preview-and-fresh-breakdown.md](../../specs/kp-append-preview-and-fresh-breakdown.md)  
**Идея**: [ai_docs/ideas/kp-append-preview-and-fresh-breakdown.md](../../ideas/kp-append-preview-and-fresh-breakdown.md)  
**Дата**: 2026-09-02  
**Статус**: PLAN ✅ · IMPLEMENT ✅

## Overview

Три среза: **A** Drawer left + №; **B** preview = product-type sealed ∪ current; **C** invalidate + regenerate breakdown после изменений состава. TDD на каждый срез.

## Architecture Decisions

- **B:** добавить `getProductTypeOrderData(draft, type)` — фильтр по типу **без** отсечения sealed. Все `build*PreviewRows` перевести на него. `getCurrentCycleOrderData` сохранить для grade/batch text.
- **Grade safety:** `build*LinesFromOrderData` пропускает `sealed`; grade-select disabled на sealed; wizard early-return на sealed index.
- **C:** helper `_invalidate_breakdown(metadata)`; вызов из patch/delete/restore/undo. `refresh_breakdown_if_needed` на GET/export — пересборка через `CommercialService.generate_preview` из plate-text order_data.
- **A:** `side="left"` + колонка № в таблице Drawer.

## Task List

### Phase A — Drawer

#### Task A1: Drawer left + row numbers + tests

**Description:** В `ProductTypePicker` передать `side="left"`; добавить колонку «№» (1-based) перед Наименование / Кол-во. Обновить RTL.

**Acceptance:**
- [x] S4: drawer left class / side prop
- [x] Заголовки №, Наименование, Кол-во; номера строк

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/components/ProductTypePicker`  
**Dependencies:** None  
**Files:** `ProductTypePicker.tsx`, `ProductTypePicker.test.tsx`

### Phase B — Preview visibility

#### Task B1: `getProductTypeOrderData` + unit tests (RED→GREEN)

**Description:** Новая функция: строки данного product_type включая sealed. Тесты: same-type sealed видны; other-type скрыты; unsealed mono как сейчас.

**Acceptance:**
- [x] Unit: sealed plates видны при productType=plates
- [x] Unit: sealed plates скрыты при productType=piles
- [x] `getCurrentCycleOrderData` поведение не ломаем (regress)

**Verification:** `npm run test -- src/features/commercial-offer/lib/currentCycleOrderData`  
**Dependencies:** None  
**Files:** `currentCycleOrderData.ts`, `currentCycleOrderData.test.ts`

#### Task B2: Switch `build*PreviewRows` + update expectations

**Description:** Все builders используют `getProductTypeOrderData`. Обновить тесты, где «empty when only sealed» для **того же** типа — теперь non-empty. Cross-type empty остаётся.

**Acceptance:**
- [x] S1–S3 для plates/piles (и аналог для остальных builders)
- [x] Не «Список пуст» при sealed same-type

**Verification:** focused vitest на `buildKpPreviewRows`, `buildPilePreviewRows`, при необходимости остальные  
**Dependencies:** B1  
**Files:** `build*PreviewRows.ts`, соответствующие `*.test.ts`

#### Task B3: Grade change uses current-cycle only

**Description:** Grade text rebuild skips sealed; sealed rows show label only (no select).

**Acceptance:**
- [x] Grade mutate не дублирует sealed через append полного списка
- [x] Preview по-прежнему показывает sealed ∪ current

**Verification:** unit on `buildPileLinesFromOrderData` + sealed filter  
**Dependencies:** B2  
**Files:** `CommercialOfferWizard.tsx`, `Kp*PreviewPanel.tsx`, builders

### Phase C — Fresh breakdown

#### Task C1: Invalidate breakdown on composition mutate (pytest RED→GREEN)

**Description:** После `patch_order_line` / `delete_order_line` / `restore_order_lines` / undo — `breakdown_tables=[]`, count 0, strip breakdown from `generated_files`.

**Acceptance:**
- [x] S5: pytest доказывает clear после patch qty

**Verification:** `pytest tests/test_commercial_draft_line_edit.py -q`  
**Dependencies:** None  
**Files:** `commercial_draft_lifecycle.py`, `test_commercial_draft_line_edit.py`

#### Task C2: Regenerate on GET/export when dirty

**Description:** `get_draft_breakdown` и `generate_files`: если tables пусты — `refresh_breakdown_if_needed` через `generate_preview`.

**Acceptance:**
- [x] S6: после patch + GET — свежие tables
- [x] FE invalidateQueries на breakdown key (уже в `invalidateDraft`)

**Verification:** pytest regenerate + web_flow breakdown; hook vitest  
**Dependencies:** C1  
**Files:** `commercial_draft_lifecycle.py`, `commercial_workflow_service.py`, `commercial.py` endpoint, tests

### Phase Verify

#### Task V1: typecheck + focused suites green; docs status

**Description:** `npm run typecheck`; обновить статусы idea/spec/plan → IMPLEMENT ✅.

**Acceptance:**
- [x] S7
- [x] Docs статусы обновлены

**Verification:** commands above  
**Dependencies:** A–C  
**Files:** ai_docs idea/spec/plan

## Execution Order

1. A1  
2. B1 → B2 → B3  
3. C1 → C2  
4. V1  

## Notes

- Не коммитить без просьбы.
- D/E вынесены в [kp-archive-only-save](../../specs/kp-archive-only-save.md) + [kp-breakdown-xlsx-format](../../specs/kp-breakdown-xlsx-format.md) — **IMPLEMENT ✅**.
- Не убивать `./run+logs.sh`.
