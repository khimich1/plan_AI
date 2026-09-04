# Spec: КП — видимый состав при дописи + свежий breakdown

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-09-02  
**One-pager**: [ai_docs/ideas/kp-append-preview-and-fresh-breakdown.md](../ideas/kp-append-preview-and-fresh-breakdown.md)  
**Plan**: [ai_docs/develop/plans/2026-09-02-kp-append-preview-and-fresh-breakdown.md](../develop/plans/2026-09-02-kp-append-preview-and-fresh-breakdown.md)  
**Related**: [kp-multi-type-picker-transparency.md](./kp-multi-type-picker-transparency.md), [kp-row-edit-delete-icons.md](./kp-row-edit-delete-icons.md), [kp-multi-nomenclature-append.md](./kp-multi-nomenclature-append.md)

## Objective

**Проблема.** (B) При «Добавить к списку» / дописи того же типа предпросмотр шага 1 показывает «Список пуст», хотя sealed-строки в draft есть. (C) После правки строки / wide / recalculate скачивание `breakdown.xlsx` и GET breakdown отдают старую разбивку. (A) Drawer (i) открывается справа и без нумерации строк.

**Цель.** Предпросмотр = полный состав типа (sealed ∪ текущий заход); breakdown инвалидируется и пересобирается до отдачи; Drawer слева с колонкой «№».

**Пользователь:** менеджер в мастере КП (picker, шаг ввода, Result / скачивание).

**Успех:** при дописи видны уже забитые позиции типа; после правки состава breakdown не stale; (i) — слева с №.

---

## ASSUMPTIONS I'M MAKING

1. **B root cause.** `getCurrentCycleOrderData` отбрасывает все строки с `append_batch_id`; `build*PreviewRows` используют его → после seal / start-append предпросмотр пуст, пока нет новых unsealed строк.
2. **B fix.** Preview берёт **все строки данного `product_type`** (sealed ∪ unsealed). Чужие типы по-прежнему скрыты. Одна таблица; опциональный бейдж «новый» — не обязателен в MVP.
3. **Grade re-ingest.** Смена класса бетона по-прежнему шлёт только **текущий цикл** (`getCurrentCycleOrderData` + append mode). У sealed строк grade-select disabled или не вызывает append-rebuild по полному списку.
4. **C root cause.** `patch_order_line` / `delete_order_line` / restore не трогают `metadata.breakdown_tables`; `get_draft_breakdown` и export читают кэш as-is.
5. **C fix.** При изменении состава/конфига — clear breakdown (+ strip generated breakdown file). При GET breakdown / export breakdown — **пересборка**, если tables пусты/dirty и есть данные для пересчёта; иначе пустой ответ / skip файла (не stale).
6. **A.** `Drawer side="left"`; колонки `№ | Наименование | Кол-во` (1-based).
7. **Коммиты агент не делает**, пока явно не попросите.
8. **Не убивать** уже запущенный `./run+logs.sh`.

→ **Assumptions locked 2026-09-02** (defaults from ideation).

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| **D-preview** | Состав на шаге 1 | Одна таблица: sealed ∪ current того же типа |
| **D-cycle** | Сверка / grade text | `getCurrentCycleOrderData` остаётся для batch/grade-rebuild |
| **D-breakdown** | Свежесть | Invalidate on mutate; regenerate on get/export if dirty |
| **D-drawer** | (i) | `side="left"` + колонка № |
| **D-out** | Out (then) | Pencil popover; **D/E done**: [kp-archive-only-save](./kp-archive-only-save.md), [kp-breakdown-xlsx-format](./kp-breakdown-xlsx-format.md) |

---

## User Stories

- Как **менеджер**, нажимая «Добавить к списку» плит, я вижу уже добавленные плиты в «Состав КП», а не «Список пуст».
- Как **менеджер**, открыв (i) на picker, вижу список слева с нумерацией.
- Как **менеджер**, после правки qty/марки и скачивания разбивки получаю актуальную версию, не прошлую.

---

## Tech Stack

| Слой | Стек |
|------|------|
| Frontend | React 19, TS, Vite, Vitest + Testing Library, TanStack Query |
| Backend | FastAPI, `CommercialDraftLifecycle`, `commercial_export_service`, `CommercialService` |
| API | REST `/api/v1/commercial/drafts/...` (без новых endpoint’ов, поведение breakdown/export) |

Новых пакетов нет.

## Commands

```
# Frontend
cd frontend && npm run test -- src/features/commercial-offer/components/ProductTypePicker
cd frontend && npm run test -- src/features/commercial-offer/lib/currentCycleOrderData
cd frontend && npm run test -- src/features/commercial-offer/lib/buildKpPreviewRows
cd frontend && npm run test -- src/features/commercial-offer/lib/buildPilePreviewRows
cd frontend && npm run test -- src/features/commercial-offer/hooks/useCommercialOfferWizard
cd frontend && npm run typecheck

# Backend
pytest tests/test_commercial_draft_line_edit.py tests/test_commercial_web_flow.py -q -k "breakdown or patch_line or delete"
```

Dev: не убивать `./run+logs.sh`.

## Project Structure

```
frontend/src/features/commercial-offer/
  components/ProductTypePicker.tsx          → side=left, № column
  components/ProductTypePicker.test.tsx
  lib/currentCycleOrderData.ts              → + getProductTypeOrderData
  lib/build*PreviewRows.ts                  → product-type full list
  components/CommercialOfferWizard.tsx      → grade change via current-cycle only
  hooks/useCommercialOfferWizard.ts         → invalidate breakdown (уже частично)
app/services/commercial_draft_lifecycle.py  → clear + refresh breakdown
app/services/commercial_export_service.py   → regenerate before write
tests/test_commercial_draft_line_edit.py
tests/test_commercial_web_flow.py
ai_docs/ideas|specs|develop/plans/...
```

## Code Style

- Русские `detail` как у существующих draft API.
- Preview helpers — чистые функции; TDD на `getProductTypeOrderData` / builders.
- Не менять layout Excel разбивки (E out of scope).

## Testing Strategy

| Уровень | Что |
|---------|-----|
| RTL ProductTypePicker | Drawer left (`app-drawer--left`); заголовки № / Наименование / Кол-во; номера 1..N |
| Unit currentCycle / build*Preview | sealed same-type видны в preview; other-type скрыты; empty только если нет строк типа |
| Unit/RTL grade | sealed grade не уходит в append-rebuild полного списка |
| Pytest patch/delete | после mutate `breakdown_tables` cleared / count 0 |
| Pytest get/export | после dirty — regenerate или не отдать stale rows |
| FE hook | invalidateQueries на breakdown key после line mutate (regress) |
| typecheck | зелёный |

## Boundaries

- **Always:** не показывать чужой product_type в preview; не отдавать stale breakdown.xlsx; a11y Drawer.
- **Ask first:** redesign Excel (E); убрать production save (D); edit из Drawer.
- **Never:** два отдельных экрана sealed vs current; новые npm/pip; убивать `./run+logs.sh`.

## Success Criteria

| # | Критерий |
|---|----------|
| S1 | Append same-type: preview показывает sealed ∪ current; не «Список пуст» при наличии строк типа |
| S2 | Cross-type append: в preview только активный тип |
| S3 | «Добавить к списку» не очищает видимый sealed-список того же типа |
| S4 | Drawer (i): `side=left`, колонки № · Наименование · Кол-во |
| S5 | После patch/delete/restore breakdown invalidated (не stale cache) |
| S6 | Export/GET breakdown после правки — свежие данные или пусто до готовности, не старая таблица |
| S7 | Vitest + typecheck + focused pytest зелёные |

## Out of Scope

- D: вырезать production из create / archive-only save  
- E: эталонный layout breakdown.xlsx  
- Pencil popover side  
- Перенос (i) на Result  

## Open Questions

_Нет — locked defaults 2026-09-02._

---

**Next:** Done (IMPLEMENT ✅).
