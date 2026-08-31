# Implementation Plan: GSM — календарь машины в раскрытой строке

Дата: 2026-08-24. Статус: approved, к реализации.
Спека: [`../../specs/gsm-vehicle-month-calendar.md`](../../specs/gsm-vehicle-month-calendar.md)
(2026-08-24, direction confirmed).
Идея: [`../../ideas/gsm-vehicle-month-calendar.md`](../../ideas/gsm-vehicle-month-calendar.md).

## Overview

Фронтенд-срез: в раскрытой строке обзора над лентой ПЛ — сетка дней выбранного
периода. В клетке маркеры транзакции (`fuel`/`wash`) и ПЛ; дыра (tx без ПЛ)
отдельно от красного бака. Backend / `core/gsm` не трогаем.

## Architecture Decisions

- **Модель дня — чистая функция** (`buildVehicleDayCells` + раскладка слотов
  недели). UI только рисует. Склейка `ts.slice(0, 10)` vs `waybill.date`.
- **Данные:** существующие `useGsmWaybillsQuery` + `useGsmTransactionsQuery`
  (`vehicleId`, тот же `from`/`to`). Новых GET нет.
- **Сетка 7 колонок Пн–Вс.** Даты только внутри `[from, to]`; до/после —
  пустые слоты. Диапазон через границу месяца — одна сетка, не два календаря.
- **Клик:** день с ПЛ → текущий drawer; дыра → `focus` + `scrollIntoView` на
  «Сгенерировать»; пустая клетка и слот — не кнопки. Generate не вызываем.
- **Пустой период:** все клетки нейтральные + текст «нет движений». Период
  шапки не меняем.

## Task List

### Phase 1: Модель (TDD)

- [x] **Task 1: `vehicleDayMap` — дни, дыры, красный бак, слоты недели**
  - **Description:** Чистые функции: список дат `[from, to]`; `VehicleDayCell`
    (`hasTx`/`hasPl`/`isGap`/`isRed`/`waybill`); раскладка недель со слотами
    с понедельника до воскресенья. `other` не якорь. Несколько tx в день —
    один маркер. Несколько ПЛ — первый по `id`.
  - **Acceptance criteria:**
    - [x] `fuel` без ПЛ → `isGap`; `wash` без ПЛ → `isGap`; `other` без ПЛ → не gap.
    - [x] ПЛ без tx → не gap; ПЛ + `manual_intervention` → `isRed`, не gap даже если есть tx.
    - [x] Август 2026 (Сб…Пн): слоты перед 01.08 и после 31.08; ни один слот не имеет `date` вне диапазона.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/lib/vehicleDayMap.test.ts`
  - **Dependencies:** None
  - **Files:** `frontend/src/features/gsm/lib/vehicleDayMap.ts`,
    `frontend/src/features/gsm/lib/vehicleDayMap.test.ts`
  - **Scope:** S (1-2 файла)

### Phase 2: UI

- [x] **Task 2: `VehicleMonthCalendar` — сетка и клики**
  - **Description:** Компонент 7 колонок (Пн–Вс). Клетка дня: номер,
    маркеры tx/ПЛ. Дыра — warning (не danger). `isRed` — danger-кольцо.
    Слоты пустые. Пустой период (все клетки без tx и ПЛ) — сетка + текст
    «нет движений». `data-testid={cal-day-${date}}`.
  - **Acceptance criteria:**
    - [x] Дыра доступным именем отличается от красного ПЛ.
    - [x] Клик по ПЛ → `onDayClick(waybill)`; клик по дыре → `onGapClick()`;
          слот и пустой день без обработчика-кнопки.
    - [x] Пустой период: текст «нет движений» и клетки диапазона на месте.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/components/VehicleMonthCalendar.test.tsx`
  - **Dependencies:** Task 1
  - **Files:** `frontend/src/features/gsm/components/VehicleMonthCalendar.tsx`,
    `frontend/src/features/gsm/components/VehicleMonthCalendar.test.tsx`
  - **Scope:** S (1-2 файла)

- [x] **Task 3: Вставить календарь в журнал ПЛ**
  - **Description:** `VehicleWaybillJournal` грузит транзакции той же машины
    и периода (`useGsmTransactionsQuery`). Календарь над `VehiclePeriodStrip`.
    Кнопка «Сгенерировать» с ref: дыра → focus + scrollIntoView, `onGenerate`
    не вызывать. Ошибка tx — Alert, сетка по ПЛ без маркеров tx. Лента и
    таблица без регрессии.
  - **Acceptance criteria:**
    - [x] В тесте журнала: tx 05.08 без ПЛ → клетка-дыра; клик не зовёт generate.
    - [x] Клик по 03.08 с ПЛ открывает путь к drawer (как сейчас).
    - [x] Июль без tx/ПЛ: «нет движений»; период пропсов не меняется.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/components/VehicleWaybillJournal.test.tsx`
  - **Dependencies:** Task 2
  - **Files:** `frontend/src/features/gsm/components/VehicleWaybillJournal.tsx`,
    `frontend/src/features/gsm/components/VehicleWaybillJournal.test.tsx`
  - **Scope:** S (1-2 файла)

### Checkpoint: Frontend готов

- [x] `cd frontend && npm test -- --run src/features/gsm/` — зелёный
- [x] `cd frontend && npm run build` — зелёный

### Phase 3: Приёмка в браузере (чтение)

- [x] **Task 4: Июль / август Palisade на live `/gsm`**
  - **Description:** Обзор, Palisade. Июль 2026 — сетка + «нет движений»,
    шапка остаётся июлем. Август 2026 — дни с ПЛ; дыра (если есть tx без ПЛ)
    ≠ красный бак. Generate с клика дыры не уходит (только фокус). Write
    (generate/export) не делать.
  - **Acceptance criteria:**
    - [x] Июль: пустые клетки + «нет движений», период 01.07–31.07.
    - [x] Август: сетка над лентой, клик дня с ПЛ открывает drawer.
  - **Verification:** live DB read-only (Palisade id=1: July 0 tx/0 ПЛ; Aug 10 ПЛ, anchors covered — дыр нет, как в Risks); Vite HMR отдаёт `VehicleMonthCalendar`/`vehicleDayMap`/`VehicleWaybillJournal`. Chrome DevTools MCP в сессии нет — клик drawer/дыры закрыт `VehicleWaybillJournal.test.tsx`. Без записи в `plita.db`.
  - **Dependencies:** Task 3, Checkpoint
  - **Files:** нет (приёмка)
  - **Scope:** XS

## Risks and Mitigations

| Риск | Вероятность | Митигация |
|:---|:---|:---|
| `Button` не пробрасывает ref | Средняя | T3: native `button` wrapper или `forwardRef` у shared Button — минимально |
| Журнал Spinner, пока не пришли tx | Средняя | Календарь после waybills; tx loading — скелет/спиннер только сетки |
| Ошибка GET transactions роняет журнал | Низкая | Alert + сетка по ПЛ без hasTx |
| Дыры нет на живом августе Palisade (все tx покрыты) | Высокая | Success: стили дыры закрыты тестом; live проверяет сетку+drawer |
| `from > to` | Низкая | Обзор уже не должен так слать; функция возвращает `[]` |

## Open Questions

- Нет. Три дефолта UX подтверждены 2026-08-24 (сетка на пустом месяце;
  слоты до Вс; focus+scrollIntoView).
