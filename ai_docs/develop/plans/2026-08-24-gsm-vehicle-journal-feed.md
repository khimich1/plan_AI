# Implementation Plan: GSM — лента журнала машины (карточки tx + ПЛ, листание месяцев)

Дата: 2026-08-24. Статус: approved, к реализации.
Спека: [`../../specs/gsm-vehicle-journal-feed.md`](../../specs/gsm-vehicle-journal-feed.md)
(2026-08-24, approved).
Идея: [`../../ideas/gsm-vehicle-journal-feed.md`](../../ideas/gsm-vehicle-journal-feed.md).

## Overview

Фронтенд-срез: в раскрытой строке обзора календарь получает шапку ‹ Месяц ›
(листание), таблица ПЛ заменяется хронологической лентой по дням: карточка
транзакции → ПЛ-карточки того же дня. Дыра (fuel/wash без ПЛ) — янтарная
карточка с CTA. Генерация целится в видимый месяц. Backend / `core/gsm` не
трогаем.

## Architecture Decisions

- **Модель ленты — чистые функции** (`monthBounds`/`shiftMonth`/
  `buildVehicleDayFeed`) в `vehicleDayFeed.ts`. UI только рисует. Склейка
  `ts.slice(0, 10)` vs `waybill.date` — как в `vehicleDayMap.ts`.
- **Видимый месяц — state журнала.** Инициализация месяцем `periodFrom`;
  сброс при смене `periodFrom`/`periodTo` пропсов и при размонтировании
  (сворачивание строки). Запросы waybills+tx идут с границами видимого
  месяца, не верхнего периода.
- **Лента = все tx машины** (fuel/wash/other), группировка по дням; `isGap`
  только для fuel/wash без ПЛ в тот же день. Пустые дни не рендерятся.
- **Календарь не ломаем:** маркеры/дыры/клики как были; добавляется шапка
  ‹ › с `onMonthChange(delta)`. `vehicleDayMap.ts` без изменений.
- **Генерация видимого месяца:** журнал зовёт `onGenerate({ from, to })` с
  границами месяца; `FleetOverviewView` хранит override-период для диалога.
  `VehicleGenerateDialog` без изменений (период — пропсы, ресинк на open
  уже есть).
- **Таблица ПЛ удаляется**, итоги (ПЛ/км/выдано) — summary-строка над
  лентой. `VehiclePeriodStrip` и `WaybillDayDrawer` сохраняются.

## Task List

### Phase 1: Модель (TDD)

- [x] **Task 1: `vehicleDayFeed` — месяцы и группировка дней**
  - **Description:** Чистые функции: `monthBounds("YYYY-MM") → {from,to}`;
    `shiftMonth(month, delta)` (в т.ч. через границу года);
    `buildVehicleDayFeed(from, to, waybills, txs) → VehicleDayFeed[]` — только
    дни с событиями, по возрастанию `date`; внутри дня tx по `ts`, ПЛ по `id`;
    `isGap` = ∃ tx ∈ {fuel,wash} && ПЛ нет. `other` без ПЛ ≠ gap.
  - **Acceptance criteria:**
    - [ ] `monthBounds("2026-02")` → 01.02–28.02; `shiftMonth("2026-01",-1)` → "2025-12".
    - [ ] fuel без ПЛ → `isGap`; other без ПЛ → не gap; tx+ПЛ → не gap.
    - [ ] Несколько tx/ПЛ в день — все в секции, порядок ts/id; пустой день пропущен.
    - [ ] `from > to` → `[]`.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/lib/vehicleDayFeed.test.ts`
  - **Dependencies:** None
  - **Files:** `frontend/src/features/gsm/lib/vehicleDayFeed.ts`,
    `frontend/src/features/gsm/lib/vehicleDayFeed.test.ts`
  - **Scope:** S (1-2 файла)

### Phase 2: UI

- [x] **Task 2: `VehicleDayFeed` — summary и секции дней**
  - **Description:** Summary-строка «Итого за {месяц}: N ПЛ, X км, выдано Y л».
    Секция дня (`data-testid={feed-day-${date}}`): дата, tx-карточки (время
    `ts.slice(11,16)`, услуга, АЗС, литры, сумма), ПЛ-карточки (водитель,
    маршрут, км, бак нач/выд/кон, статус, бейджи warnings). Дыра — янтарная
    (`#fffaeb`/`#fdb022`) + CTA «Сгенерировать» → `onGapClick()`. Клик ПЛ →
    `onWaybillClick(waybill)`. Пустой feed — ничего (календарь уже пишет
    «нет движений»).
  - **Acceptance criteria:**
    - [ ] Tx-карточка показывает время/услугу/АЗС/литры/сумму; `other` не дыра.
    - [ ] Дыра янтарная с CTA; красный ПЛ (`manual_intervention`) — бейдж, не дыра.
    - [ ] Summary считает по waybills месяца; клики пробрасываются наружу.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/components/VehicleDayFeed.test.tsx`
  - **Dependencies:** Task 1
  - **Files:** `frontend/src/features/gsm/components/VehicleDayFeed.tsx`,
    `frontend/src/features/gsm/components/VehicleDayFeed.test.tsx`
  - **Scope:** S (1-2 файла)

- [x] **Task 3: `VehicleMonthCalendar` — шапка ‹ {Месяц} ›**
  - **Description:** Над сеткой шапка: кнопки ‹ › и label месяца
    (`toLocaleDateString("ru-RU", { month: "long", year: "numeric" })`,
    первая буква заглавная). Клики → `onMonthChange(-1|+1)`. Сетка, маркеры,
    дыры, существующие клики — без регрессии. Проп `month: string`.
  - **Acceptance criteria:**
    - [ ] ‹ › вызывают `onMonthChange` с ±1; label совпадает с `month`.
    - [ ] Существующие тесты календаря зелёные (проп `month` добавлен аккуратно).
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/components/VehicleMonthCalendar.test.tsx`
  - **Dependencies:** Task 1
  - **Files:** `frontend/src/features/gsm/components/VehicleMonthCalendar.tsx`,
    `frontend/src/features/gsm/components/VehicleMonthCalendar.test.tsx`
  - **Scope:** S (1-2 файла)

- [x] **Task 4: Журнал — видимый месяц, лента вместо таблицы**
  - **Description:** `VehicleWaybillJournal`: state `month` (init месяцем
    `periodFrom`, reset на смену пропсов — `useEffect`); запросы waybills/tx с
    `monthBounds(month)`; `VehicleMonthCalendar` получает `month` +
    `onMonthChange`; таблица и `tfoot` удаляются, вместо них `VehicleDayFeed`;
    `onGenerate` вызывается с `monthBounds(month)`; CTA дыры → `focusGenerate`
    (без вызова generate). Тесты журнала переписаны на ленту.
  - **Acceptance criteria:**
    - [ ] ‹ › меняет параметры обоих запросов (границы нового месяца).
    - [ ] Смена `periodFrom`/`periodTo` пропсов сбрасывает месяц.
    - [ ] Таблицы нет; summary + лента на месте; дыра → focus, не generate.
    - [ ] «Сгенерировать» после листания зовёт `onGenerate` с границами видимого месяца.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/components/VehicleWaybillJournal.test.tsx`
  - **Dependencies:** Task 2, Task 3
  - **Files:** `frontend/src/features/gsm/components/VehicleWaybillJournal.tsx`,
    `frontend/src/features/gsm/components/VehicleWaybillJournal.test.tsx`
  - **Scope:** M (2 файла)

- [x] **Task 5: `FleetOverviewView` — генерация видимого месяца**
  - **Description:** `onGenerate` журнала принимает `{ from, to }`; overview
    сохраняет override-период рядом с `generateRow` и передаёт его в
    `VehicleGenerateDialog` вместо верхнего `periodFrom/periodTo`. Без
    override (кнопка в строке таблицы) — верхний период, как сейчас.
  - **Acceptance criteria:**
    - [ ] Генерация из журнала после листания открывает диалог с месяцем журнала.
    - [ ] Генерация кнопкой строки обзора — верхний период (без регрессии).
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/components/FleetOverviewView.test.tsx`
  - **Dependencies:** Task 4
  - **Files:** `frontend/src/features/gsm/components/FleetOverviewView.tsx`,
    `frontend/src/features/gsm/components/FleetOverviewView.test.tsx`
  - **Scope:** S (1-2 файла)

### Checkpoint: Frontend готов

- [x] `cd frontend && npm test -- --run src/features/gsm/` — зелёный (23 файла, 156 тестов)
- [x] `cd frontend && npm run build` — зелёный
- [x] Регрессия backend: `venv/bin/pytest tests/test_gsm_*.py -q` — 218 passed

### Phase 3: Приёмка в браузере (чтение)

- [x] **Task 6: Листание и лента Palisade на live `/gsm`**
  - **Description:** Обзор, Palisade. ‹ ›: июль — сетка + «нет движений»,
    лента пустая; август — карточки tx и ПЛ по дням. Дыра ≠ красный ПЛ. CTA
    дыры → фокус на «Сгенерировать»; в диалоге даты = видимый месяц. Write
    (generate/export) не делать — диалог открыть и закрыть.
  - **Acceptance criteria:**
    - [x] Листание не трогает верхний фильтр; возврат ‹ › обратно корректный.
    - [x] Диалог генерации из отлистанного месяца показывает его границы.
  - **Verification:** Chrome DevTools MCP в сессии нет, dev-серверы не
    запущены — клик-флоу закрыт интеграционными тестами (как в срезе
    fleet-overview-ux): ‹ › меняет оба запроса и label месяца без сброса
    верхнего фильтра (`VehicleWaybillJournal.test.tsx`), июль пустой
    (сетка + «нет движений», ленты нет), дыра ≠ красный ПЛ
    (`VehicleDayFeed.test.tsx`, `VehicleMonthCalendar.test.tsx`), CTA дыры →
    focus без generate, диалог из отлистанного месяца показывает его границы
    (`FleetOverviewView.test.tsx`, диалог открыт/закрыт без сабмита).
    Записи в `plita.db` не было.
  - **Dependencies:** Task 5, Checkpoint
  - **Files:** нет (приёмка)
  - **Scope:** XS

## Risks and Mitigations

| Риск | Вероятность | Митигация |
|:---|:---|:---|
| Тесты журнала завязаны на таблицу | Высокая | T4: переписать на ленту/summary в том же коммите изменения |
| `ts` формат не ISO (пробел вместо T) | Низкая | `slice(0,10)`/`slice(11,16)` работают в обоих; unit-тест с пробелом |
| Reset месяца при смене пропсов не срабатывает | Средняя | T4: `useEffect` на [periodFrom, periodTo] + тест rerender |
| Label месяца с маленькой буквы (ru-RU) | Средняя | Капитализация первой буквы; снапшот-текст в тесте T3 |
| На живом августе нет дыр | Высокая | Как в прошлом срезе: стили дыры закрыты тестами; live — листание+лента |
| Длинная лента в активном месяце | Низкая | 4 машины, десятки событий; фильтр «только дыры» — post-MVP |

## Open Questions

- Нет. Дефолты зафиксированы в спеке 2026-08-24 (все tx в ленте; карточка tx
  минимальная; генерация целится в видимый месяц; фильтр дыр post-MVP).
