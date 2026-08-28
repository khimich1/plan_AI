# Spec: ГСМ — календарь машины в раскрытой строке

Дата: 2026-08-24. Статус: draft, на ревью.
Идея: [`../ideas/gsm-vehicle-month-calendar.md`](../ideas/gsm-vehicle-month-calendar.md)
(2026-08-24, direction confirmed).
Базовый UX: [`gsm-fleet-overview-ux.md`](gsm-fleet-overview-ux.md) (реализован).

## ASSUMPTIONS I'M MAKING

1. **Backend и схема БД не меняются.** Сетка собирается на клиенте из уже
   существующих `GET /gsm/waybills?vehicle_id&from&to` и
   `GET /gsm/transactions?vehicle_id&from&to`. Новых эндпоинтов, полей
   `/overview`, миграций нет.
2. **`core/gsm/*` не трогаем.** Генерация по-прежнему периодом якорей, не днём.
3. **Клик по машине не меняется.** Раскрытие строки обзора как сейчас;
   календарь появляется внутри `VehicleWaybillJournal`, над лентой дней.
4. **Период сетки = «С / По» шапки обзора.** Молча в другой месяц не прыгаем.
   Пустой июль Palisade — пустая сетка + текст, не автопереход в август.
5. **Дыра ≠ красный день.** Дыра: в дате есть якорь-транзакция (`fuel`/`wash`)
   и нет ПЛ с `date` в тот же календарный день. Красный день: у ПЛ
   `warnings` содержит `manual_intervention`. Ходовой ПЛ без транзакции
   не подсвечиваем.
6. **`service_type=other` не якорь** (как в `core/gsm/generator.py`,
   `_ANCHOR_SERVICES`). Маркер tx и дыра только по `fuel`/`wash`.
7. **Несколько транзакций / мойка+топливо в один день** — один маркер tx.
   Несколько ПЛ в один день (не должно быть в нормальных данных) — один
   маркер ПЛ; клик открывает первый по `id`.
8. **Склейка дат:** `substr(ts, 1, 10)` транзакции vs `waybill.date`
   (`YYYY-MM-DD`). Полночь/23:50 на соседний день не сдвигаем.
9. **Диапазон не календарный месяц** (15.07–14.08) — одна сетка по всем дням
   `[from, to]`, без двух отдельных календарей. Неделя начинается с
   понедельника.
10. **Клик по дыре** не генерирует и не открывает диалог: фокус на кнопку
    «Сгенерировать» текущего журнала. Клик по дню с ПЛ — существующий
    `WaybillDayDrawer`.
11. **Умный баннер хвоста (`open_before_from`)** — вне скоупа. Контракт
    `/overview` не расширяем.
12. **Роли без изменений:** экран уже за `REQUIRE_ACCOUNTING`.
13. **Пагинации нет.** Масштаб: 4 машины, ≤31 день, десятки tx/ПЛ на месяц.
14. **Пустой период:** сетка всех дней диапазона (нейтральные клетки) **и**
    текст «нет движений» — не прячем сетку. Так видно, что период выбран.
15. **Выравнивание недели:** пустые слоты с понедельника до первого дня
    диапазона и после последнего дня до воскресенья. Слоты — не даты вне
    `[from, to]`, клика нет.
16. **Клик по дыре:** `focus` + `scrollIntoView` на «Сгенерировать». Без
    диалога, без вызова generate.

→ Поправьте сейчас, иначе иду с этим в план/задачи.

## Objective

Бухгалтер (`accountant`) в раскрытой строке машины видит **карту выбранного
периода**: в каждом дне — есть ли заправка/мойка и есть ли путевой. День, где
транзакция-якорь не покрыта ПЛ, подсвечен. Красный бак не путается с этой
дырой. Пустой период не маскируется прыжком на другой месяц.

**Пользователь:** бухгалтер (`accountant`), администратор (`admin`).

**Критерий успеха MVP:** раскрыть Palisade за август 2026 — сетка показывает
дни с ПЛ и дни с транзакциями; раскрыть Palisade за июль 2026 — сетка пустая
с текстом «нет движений», период в шапке остаётся июлем. День с `fuel`/`wash`
без ПЛ визуально отличается от дня с `manual_intervention`.

## Модель дня (клиент)

Для каждой даты `d` в `[periodFrom, periodTo]` (включительно):

```
has_tx  = есть транзакция vehicle с service_type ∈ {fuel, wash}
          и дата ts = d
has_pl  = есть ПЛ vehicle с date = d
is_gap  = has_tx && !has_pl
is_red  = has_pl && warnings содержат manual_intervention
```

| Состояние | Маркеры | Стиль клетки | Клик |
|:---|:---|:---|:---|
| пусто | — | нейтральный | нет действия (не кнопка) |
| только ПЛ | ПЛ | нейтральный (+ якорь ленты не дублируем) | drawer этого ПЛ |
| только tx | tx | **дыра** (warning, не danger) | фокус «Сгенерировать» |
| tx + ПЛ | оба | нейтральный; если `is_red` — **danger-кольцо** | drawer ПЛ |
| tx + ПЛ красный | оба | дыры нет, красный бак есть | drawer ПЛ |

Лента `VehiclePeriodStrip` сохраняется как была (только дни с ПЛ). Календарь
её не заменяет и не дублирует бак/км в клетке.

## Tech Stack

- Frontend: React 18, TypeScript, Vite, TanStack Query, `frontend/src/features/gsm/`.
- Новых npm/pip зависимостей нет. Календарную сетку не берём из библиотеки.
- Backend без изменений.

## Commands

```bash
# Frontend
cd frontend && npm test -- --run src/features/gsm/
cd frontend && npm run build

# Регрессия GSM (backend не меняем, но не ломаем)
venv/bin/pytest tests/test_gsm_*.py -q
```

Приёмка в UI: `/gsm` обзор, раскрыть машину, периоды июль и август 2026
(live или копия — только чтение; generate не обязателен).

## Project Structure

```
frontend/src/features/gsm/
  lib/vehicleDayMap.ts                 → NEW: даты диапазона, склейка tx+ПЛ,
                                          is_gap / is_red (чистая функция)
  lib/vehicleDayMap.test.ts            → NEW
  components/VehicleMonthCalendar.tsx  → NEW: сетка 7 колонок (Пн–Вс)
  components/VehicleMonthCalendar.test.tsx → NEW
  components/VehicleWaybillJournal.tsx → CHANGED: грузит tx той же машины/периода,
                                          рендерит календарь над лентой;
                                          фокус кнопки Generate по клику дыры
  components/VehicleWaybillJournal.test.tsx → CHANGED: сетка, дыра, пустой период
  hooks/useGsmQueries.ts               → без новых хуков, если
                                          useGsmTransactionsQuery уже подходит
```

`core/`, `app/` — не трогаем.

## Code Style

```typescript
export type VehicleDayCell = {
  date: string; // YYYY-MM-DD
  hasTx: boolean;
  hasPl: boolean;
  isGap: boolean;
  isRed: boolean;
  waybill: GsmWaybill | null;
};

export const buildVehicleDayCells = (
  periodFrom: string,
  periodTo: string,
  waybills: GsmWaybill[],
  transactions: { ts: string; service_type: string }[],
): VehicleDayCell[] => {
  // inclusive day walk; tx date = ts.slice(0, 10)
  // hasTx if some tx in {fuel, wash}; hasPl if some waybill.date === date
};
```

- Даты ISO `YYYY-MM-DD`.
- Неделя с понедельника; слоты до `from` и после `to` до конца недели —
  не даты вне периода.
- `data-testid={`cal-day-${date}`}`; дыра — доступное имя
  «нет путевого на заправку/мойку».
- Суммы/литры в клетке не показываем.

## Testing Strategy

| Уровень | Что | Команда |
|---|---|---|
| Unit | Диапазон дней; tx без ПЛ = gap; wash без ПЛ = gap; `other` не gap; красный только при ПЛ+`manual_intervention`; ПЛ без tx ≠ gap | `npm test -- --run src/features/gsm/lib/vehicleDayMap.test.ts` |
| Component | Сетка рендерит дни; дыра стилем/ролью отличается от красного ПЛ; клик ПЛ → callback drawer; клик дыры → callback generate-focus; пустой период — текст | `VehicleMonthCalendar.test.tsx` |
| Integration | Журнал с моками waybills+transactions показывает календарь над лентой | `VehicleWaybillJournal.test.tsx` |
| Regression | Существующий GSM vitest зелёный | `npm test -- --run src/features/gsm/` |
| Manual | Июль Palisade пустой; август — дни с ПЛ; не путать дыру и «Ручная доработка» | `/gsm` |

## Boundaries

- **Always:** TDD (сначала красный тест модели дня); фронтенд-only; регрессия
  `src/features/gsm/`; `npm run build`.
- **Ask first:** новые зависимости; поле в `/overview`; генерация одного дня;
  изменение контрактов GET; две сетки месяцев вместо одной ленты дней.
- **Never:** схема БД; `core/gsm/*`; умный баннер хвоста; подсветка «ПЛ без
  транзакции»; автопрыжок периода; замена журнала/ленты; коммиты без просьбы.

## Success Criteria

1. Раскрытие машины рисует сетку всех дней `[from, to]` над лентой ПЛ.
2. День с `fuel` или `wash` без ПЛ подсвечен как дыра (не как danger бака).
3. День с ПЛ и `manual_intervention` — красный бак; если при этом есть tx,
   это не дыра.
4. День только с ПЛ (без tx) не подсвечен.
5. Клик по дню с ПЛ открывает существующий drawer; клик по дыре ставит фокус
   на «Сгенерировать», generate не вызывается.
6. Июль без tx и ПЛ: сетка + «нет движений», даты шапки не меняются.
7. `npm test -- --run src/features/gsm/` зелёный; `npm run build` зелёный.

## Open Questions

- Нет. Дефолты 2026-08-24: сетка+текст на пустом месяце; слоты до Пн и до
  Вс; клик дыры = focus+scrollIntoView; баннер хвоста отдельно.
