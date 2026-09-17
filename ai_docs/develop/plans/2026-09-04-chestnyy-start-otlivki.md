# Implementation Plan: Честный старт отливки в «Ёмкости»

> **Спека:** [`ai_docs/specs/chestnyy-start-otlivki.md`](../../specs/chestnyy-start-otlivki.md)
> **Идея:** [`ai_docs/ideas/chestnyy-start-otlivki.md`](../../ideas/chestnyy-start-otlivki.md)
> **Дата:** 2026-09-04
> **Статус:** в работе — approve на выполнение дан.

## Overview

Поправка к календарю периодов в диалоге «В производство» / drawer «Ёмкость»:
клетка снова показывает занятость плана, но знаменатель — ручка архива
(`knob`), не `DayInfo.max` из календарного плана. «Начало» — первый день, куда
заказ ещё влезает (`first_pour_date` + остаток дня), не понедельник корзины.
Укладка `pack_pour` режет заказ по дням с лимитом `weeks[].free`; окно, холд и
гейт берут даты из `PourPlan`, не из `allocate()`. `allocate()` остаётся в
модуле со своими unit-тестами. `MonthCalendarGrid` / `FactoryMiniCalendar` не
подключаем.

## Architecture Decisions

- **Потолок дня = knob**, occupied = `days_info.occupied`. `day_free = max(0,
  knob − occupied)`. Перебор (`occupied > knob`) — клетка `4/3`, не старт.
- **Укладка `pack_pour`**: в день `take = min(remaining, day_free, week.free −
  уже взятое)`. Чужие обещания — только лимит `weeks[].free`, не цвет клеток.
  Холды не вычитаются из `free`.
- **`allocate()` не удалять.** `build_quote` / холд / гейт берут окно из
  `PourPlan` (`week_allocations` по monday недели). Дневной журнал холда нет.
- **Жёлтое по дням:** `first_pour_date` … воскресенье недели `solo_date`. Дни
  до старта в той же неделе не жёлтые.
- **Occupancy на GET promise-quote** (как holidays), клиент не ходит в
  `/production/calendar` ради сетки.
- **Клик клетки:** `onSelectWeek` + `onSelectDay`; поле срока не меняется.
- **Пустой occupancy** — регресс существующих `test_build_quote_*` сохранить
  (пустой завод совпадает с дневной укладкой по knob).

```
pack_pour / week_allocations
        │
        ▼
build_quote (solo_*, earliest=monday(first_pour), window из PourPlan)
        │
        ▼
_quote_to_response + occupancy span
        │
        ▼
PromiseQuoteBlock / PromiseWindowBand / PromisePeriodCalendar
        │
        └──────── MoveToProductionDialog + PromiseWeekOccupants ────────┘
```

## Task List

### Phase 1: Domain + API

- [x] **Task 1: `day_free`, `pack_pour`, `week_allocations`**
  - **Description:** Чистые функции в `core/production/promise_buckets.py`.
    `PourPlan`: first_pour_date/free, solo_date, solo_week_end_date, дневные
    allocations. `week_allocations` — сумма take по `iso_week_start`.
    `allocate()` и его тесты не ломать.
  - **Acceptance:** пример A: first 9.09, free 2, take 2+3, solo 10.09,
    week_end 11.09, холд `(7.09, 5)`; пример B: promised съел free → first не
    на 7–13; перебор `day_free(4,3)==0`; горизонт → None.
  - **Verify:** `pytest tests/test_promise_buckets.py -q`
  - **Dependencies:** None
  - **Files:** `core/production/promise_buckets.py`, `tests/test_promise_buckets.py`
  - **Scope:** M

- [x] **Task 2: `build_quote` из PourPlan**
  - **Description:** `build_quote` принимает occupancy; solo_*,
    earliest_start_week=monday(first_pour), window/allocations из PourPlan.
    Поля `first_pour_date` / `first_pour_free` на `PromiseQuote`. Нет плана →
    window и даты как сейчас при отсутствии окна. Пустой occupancy — регресс
    `test_build_quote_*`.
  - **Acceptance:** пустой завод: существующие assert solo/window зелёные;
    с occupancy примера A — first 9.09, promised 11.09.
  - **Verify:** `pytest tests/test_promise_buckets.py -q`
  - **Dependencies:** T1
  - **Files:** `core/production/promise_buckets.py`, `tests/test_promise_buckets.py`
  - **Scope:** S

- [x] **Task 3: API + гейт**
  - **Description:** `PromiseQuoteResponse`: `first_pour_date`,
    `first_pour_free`, `occupancy` (span как holidays, дни с occupied ≠ 0).
    `_quote_to_response` прокидывает поля. Гейт / `create_hold` без смены
    сигнатур — берут `window` с PourPlan. Тест гейта примера A: дата 10.09
    отказ, earliest 11.09.
  - **Acceptance:** response содержит новые поля; гейт A; регресс
    allocate/holds/occupants/atomicity.
  - **Verify:** `pytest tests/test_promise_service.py tests/test_archive_endpoints.py tests/test_move_to_production_atomicity.py -q`
  - **Dependencies:** T2
  - **Files:** `app/schemas/archive.py`, `app/services/promise_service.py`, тесты
  - **Scope:** M

### Phase 2: Frontend

- [x] **Task 4: типы + QuoteBlock + WindowBand**
  - **Description:** `promiseQuote.ts`: `first_pour_date`, `first_pour_free`,
    `occupancy`. `PromiseQuoteBlock`: «Начало: D.MM · остаток N дор.»;
    нет даты → «Начало: —». `PromiseWindowBand`: левая дата
    `first_pour_date`, правая — воскресенье жёлтого интервала; маркер
    клиенту на `promised_date`.
  - **Acceptance:** vitest «Начало: 9.09» + остаток 2; соло 10.09; не 7.09;
    полоса 9.09, не 7.09.
  - **Verify:** `cd frontend && npm run test -- --run src/features/factory-capacity/components/PromiseQuoteBlock.test.tsx src/features/factory-capacity/components/PromiseWindowBand.test.tsx`
  - **Dependencies:** None (∥ Phase 1; контракт из спеки)
  - **Files:** `promiseQuote.ts`, `PromiseQuoteBlock.tsx`, `PromiseWindowBand.tsx`, тесты
  - **Scope:** S

- [x] **Task 5: `PromisePeriodCalendar` — дроби и жёлтое по дням**
  - **Description:** occupancy/knob; дробь `{occupied}/{knob}`; перебор
    `4/3`; жёлтое `pourFrom…pourToSunday` по дням (пн–вт до старта не
    жёлтые); маркер старта ≠ точка клиента; `onSelectWeek` + `onSelectDay`.
  - **Acceptance:** 7 и 8 не жёлтые при pour_from=9; 9 жёлтый; 3/3 и 1/3;
    4/3 перебор; клик 9-го → week 7.09 и day 9.
  - **Verify:** `cd frontend && npm run test -- --run src/features/factory-capacity/components/PromisePeriodCalendar.test.tsx`
  - **Dependencies:** T4 типы
  - **Files:** `PromisePeriodCalendar.tsx`, `PromisePeriodCalendar.test.tsx`
  - **Scope:** M

- [x] **Task 6: Occupants + диалог**
  - **Description:** `PromiseWeekOccupants`: «Свободно: X из Y» из `weeks[]`
    выбранной недели + строка кликнутого дня. `MoveToProductionDialog`:
    state дня, прокинуть occupancy/knob/pour; клик не меняет
    `execution_terms`.
  - **Acceptance:** «Свободно: 8 из 15»; строка дня; клик не пишет срок;
    `npm run typecheck`.
  - **Verify:**
    ```
    cd frontend && npm run test -- --run src/features/factory-capacity src/features/commercial-archive/components/MoveToProductionDialog.test.tsx
    cd frontend && npm run typecheck
    ```
  - **Dependencies:** T4, T5
  - **Files:** `PromiseWeekOccupants.tsx`, `MoveToProductionDialog.tsx`, тесты
  - **Scope:** M

### Checkpoint: Complete

- [x] Команды из спеки зелёные
- [x] Пример A: Начало 9.09, соло 10.09, обещать к 11.09; 7–8 не жёлтые; 9–13 жёлтые
- [x] Пример B: first не на 7–13 при free=0
- [ ] Human review перед merge

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| `test_build_quote_*` ждут allocate-окно | Med | пустой occupancy даёт ту же нарезку по неделям, что greedy allocate |
| Гейт существующих тестов сдвинется | Med | пустой завод: solo/promised как nth_workday; новый тест только на примере A |
| Жёлтая целая неделя останется | Med | убрать `isInWindow(weekStart)`; красить клетку по дате |

## Open Questions

Нет блокирующих.

## Stop

План записан. Approve на выполнение уже дан — реализация по T1–T6.
