# Spec: ГСМ — лента журнала машины (карточки tx + ПЛ, листание месяцев)

Дата: 2026-08-24. Статус: draft, на ревью.
Идея: [`../ideas/gsm-vehicle-journal-feed.md`](../ideas/gsm-vehicle-journal-feed.md)
(2026-08-24, direction A confirmed).
Базовый UX: [`gsm-vehicle-month-calendar.md`](gsm-vehicle-month-calendar.md) (реализован).

## ASSUMPTIONS I'M MAKING

1. **Frontend-only.** Оба запроса (`useGsmWaybillsQuery`,
   `useGsmTransactionsQuery`) уже живут в `VehicleWaybillJournal`. Новых
   эндпоинтов, полей, миграций нет. `core/gsm/*` не трогаем.
2. **Пейджинг месяцев — внутреннее состояние журнала.** ‹ › в шапке календаря
   сдвигают видимый месяц на ±1; оба запроса рефетчатся с границами
   календарного месяца (`monthBounds`). Верхний фильтр «С/По» — только
   начальное значение: видимый месяц инициализируется месяцем `periodFrom`;
   при смене верхнего периода — сброс на него. При сворачивании/раскрытии
   строки — тоже сброс (state живёт внутри компонента журнала).
2a. **Генерация целится в видимый месяц** (подтверждено 2026-08-24): кнопка
   «Сгенерировать» журнала и CTA карточки-дыры открывают
   `VehicleGenerateDialog` с границами видимого месяца, а не верхнего периода.
   Диалог без изменений — период приходит пропсами; меняется только
   проброс в `FleetOverviewView`.
3. **Пейджинг свободный** (любой месяц назад/вперёд), данные подгружаются по
   требованию. Масштаб прежний: 4 машины, десятки tx/ПЛ на месяц.
4. **Лента заменяет таблицу ПЛ** (подтверждено 2026-08-24). Итоги из `tfoot`
   (км, выдано литров) переезжают в summary-строку над лентой.
5. **Лента = все транзакции машины** (fuel/wash/other), сгруппированные по
   дням вместе с ПЛ. Дырой подсвечивается только `fuel`/`wash` без ПЛ в тот же
   день (как в `vehicleDayMap.ts`). `other` без ПЛ — обычная карточка.
6. **Секции только для дней с событиями** (tx или ПЛ); пустые дни не
   рендерятся. Порядок — по возрастанию дат (как календарь). Внутри дня:
   сначала карточки транзакций (по `ts`), затем карточки ПЛ (по `id`).
7. **Карточка транзакции минимальная:** время (`ts.slice(11, 16)`), услуга
   (Топливо/Мойка/`service_type`), АЗС (`address` или «—»), литры, сумма.
   Карта и вид топлива не показываем (есть вкладка «Транзакции»).
8. **Карточка-дыра** (fuel/wash без ПЛ в тот же день) — янтарная (та же
   палитра, что дыра в календаре: `#fffaeb`/`#fdb022`), с CTA «Сгенерировать»
   → существующий `focusGenerate` (focus + scrollIntoView, без вызова
   generate).
9. **Карточка ПЛ:** водитель, маршрут (`formatRouteSummary`), км, бак
   нач/выд/кон, статус, бейджи warnings (существующие `warningMeta`). Клик —
   существующий `WaybillDayDrawer`. Красный ПЛ (`manual_intervention`) ≠ дыра.
10. **`VehiclePeriodStrip` сохраняется** между календарём и лентой.
    Поведение календаря (маркеры, дыры, клики) не меняется — добавляется
    только шапка с ‹ › и названием месяца.
11. **Фильтр «только дыры» — не в MVP.** Дыры и так янтарные; добавим, если
    лента раздуется на живых данных.
12. **Пустой месяц:** сетка + «нет движений» (как сейчас), лента пустая.
    Автопрыжков на соседние месяцы нет.
13. **Роли без изменений:** экран уже за `REQUIRE_ACCOUNTING`.

→ Поправьте сейчас, иначе иду с этим в план/задачи.

## Objective

Бухгалтер (`accountant`) в раскрытой строке машины **листает месяцы стрелками**
в календаре и видит под ним **хронологическую ленту месяца**: каждая
транзакция — карточка, за ней — путевые листы того же дня. Незакрытая
заправка/мойка видна мгновенно (янтарная карточка-дыра с CTA), без переключения
на вкладку «Транзакции».

**Пользователь:** бухгалтер (`accountant`), администратор (`admin`).

**Критерий успеха MVP:** раскрыть Palisade — календарь с ‹ ›; листание месяца
меняет и сетку, и ленту; в ленте день с заправкой показывает карточку
транзакции и ПЛ того же дня под ней; таблицы ПЛ больше нет, итоги — в
summary-строке. Июль 2026: сетка + «нет движений», лента пустая.

## Модель ленты (клиент)

Для видимого месяца `m` (`monthFrom`/`monthTo` — границы календарного месяца):

```
DayFeed = {
  date: string;               // YYYY-MM-DD
  txs: GsmTransaction[];      // все service_type, по возрастанию ts
  waybills: GsmWaybill[];     // по возрастанию id
  isGap: boolean;             // ∃ tx ∈ {fuel, wash} && waybills.length === 0
}

feed = days в [monthFrom, monthTo], где txs.length + waybills.length > 0,
       по возрастанию date
```

| Ситуация в дне | Рендер |
|:---|:---|
| tx fuel/wash, ПЛ нет | янтарная tx-карточка-дыра + CTA «Сгенерировать» |
| tx fuel/wash + ПЛ | обычная tx-карточка → ПЛ-карточки |
| tx other, ПЛ нет | обычная tx-карточка (не дыра) |
| только ПЛ | ПЛ-карточки (рабочий день) |
| ПЛ с `manual_intervention` | ПЛ-карточка с красным бейджем (не дыра) |
| нет tx и ПЛ | секция не рендерится |

Summary-строка над лентой: «Итого за {месяц}: {n} ПЛ, {км} км, выдано {л} л»
(данные waybills видимого месяца; заменяет удалённый `tfoot`).

## Tech Stack

- Frontend: React 18, TypeScript, Vite, TanStack Query, `frontend/src/features/gsm/`.
- Новых npm/pip зависимостей нет.
- Backend без изменений.

## Commands

```bash
# Frontend
cd frontend && npm test -- --run src/features/gsm/
cd frontend && npm run build

# Регрессия GSM (backend не меняем, но не ломаем)
venv/bin/pytest tests/test_gsm_*.py -q
```

Приёмка в UI: `/gsm` обзор, раскрыть машину, листать июль ↔ август 2026
(live или копия — только чтение; generate не обязателен).

## Project Structure

```
frontend/src/features/gsm/
  lib/vehicleDayFeed.ts              → NEW: monthBounds/shiftMonth, группировка
                                        tx+ПЛ по дням, isGap (чистые функции)
  lib/vehicleDayFeed.test.ts         → NEW
  components/VehicleDayFeed.tsx      → NEW: summary-строка + секции дней
                                        (tx-карточки, ПЛ-карточки, дыра+CTA)
  components/VehicleDayFeed.test.tsx → NEW
  components/VehicleMonthCalendar.tsx   → CHANGED: шапка ‹ {Месяц YYYY} ›,
                                          onMonthChange
  components/VehicleMonthCalendar.test.tsx → CHANGED
  components/VehicleWaybillJournal.tsx  → CHANGED: state видимого месяца,
                                          запросы по границам месяца, таблица
                                          удалена, лента + summary,
                                          onGenerate(monthBounds)
  components/VehicleWaybillJournal.test.tsx → CHANGED
  components/FleetOverviewView.tsx      → CHANGED: onGenerate принимает границы
                                          видимого месяца и передаёт их в
                                          VehicleGenerateDialog
```

`VehicleGenerateDialog` — без изменений (период приходит пропсами).

`core/`, `app/` — не трогаем. `vehicleDayMap.ts` переиспользуется календарём
без изменений.

## Code Style

```typescript
export type VehicleDayFeed = {
  date: string; // YYYY-MM-DD
  txs: GsmTransaction[];
  waybills: GsmWaybill[];
  isGap: boolean;
};

export const monthBounds = (month: string): { from: string; to: string } => {
  // month = "YYYY-MM" → первый/последний день, ISO
};

export const shiftMonth = (month: string, delta: number): string => {
  // "2026-01" + 1 → "2026-02"; "2026-01" - 1 → "2025-12"
};

export const buildVehicleDayFeed = (
  periodFrom: string,
  periodTo: string,
  waybills: GsmWaybill[],
  transactions: GsmTransaction[],
): VehicleDayFeed[] => {
  // только дни с событиями; isGap только для fuel/wash без ПЛ
};
```

- Даты ISO `YYYY-MM-DD`, месяц `YYYY-MM`; время карточки `ts.slice(11, 16)`.
- Название месяца: `toLocaleDateString("ru-RU", { month: "long", year: "numeric" })`.
- `data-testid={`feed-day-${date}`}`; дыра — доступное имя
  «нет путевого на заправку/мойку» (как в календаре).
- Стили — существующая палитра (`#fffaeb`/`#fdb022` дыра, `#fef3f2`/`#f04438`
  красный ПЛ), инлайн-стили как в соседних компонентах.

## Testing Strategy

| Уровень | Что | Команда |
|---|---|---|
| Unit | `monthBounds`/`shiftMonth` (в т.ч. через границу года); группировка по дням; gap только fuel/wash без ПЛ; `other` ≠ gap; несколько tx/ПЛ в день; пустые дни пропущены; сортировка | `npm test -- --run src/features/gsm/lib/vehicleDayFeed.test.ts` |
| Component | Лента: tx-карточка (время/АЗС/литры/сумма); дыра янтарная + CTA; ПЛ-карточка → callback drawer; summary-строка; пустой месяц | `VehicleDayFeed.test.tsx` |
| Component | Календарь: шапка ‹ ›, label месяца, onMonthChange | `VehicleMonthCalendar.test.tsx` |
| Integration | Журнал: таблицы нет; лента под strip'ом; ‹ › меняет запросы (границы месяца); смена верхнего периода сбрасывает месяц; CTA дыры → focus «Сгенерировать»; «Сгенерировать» после листания передаёт границы видимого месяца в onGenerate | `VehicleWaybillJournal.test.tsx` |
| Regression | Существующий GSM vitest зелёный | `npm test -- --run src/features/gsm/` |
| Manual | Июль пустой; август — карточки; дыра ≠ красный ПЛ; листание не дёргает верхний фильтр | `/gsm` |

## Boundaries

- **Always:** TDD (сначала красный тест `vehicleDayFeed`); фронтенд-only;
  регрессия `src/features/gsm/`; `npm run build`.
- **Ask first:** новые зависимости; поля в GET-контрактах; явная привязка
  tx↔ПЛ; фильтр «только дыры» (post-MVP); горизонтальная карусель.
- **Never:** схема БД; `core/gsm/*`; вкладка «Транзакции»; удаление
  `WaybillDayDrawer`/`VehiclePeriodStrip`; автопрыжок месяца на пустых
  данных; коммиты без просьбы.

## Success Criteria

1. Календарь в раскрытой строке имеет шапку ‹ {Месяц} ›; листание меняет
   видимый месяц, сетку и ленту (запросы с границами месяца).
2. Лента по дням заменяет таблицу ПЛ: день с транзакцией — карточка (время,
   услуга, АЗС, литры, сумма), под ней ПЛ-карточки того же дня.
3. Tx `fuel`/`wash` без ПЛ в тот же день — янтарная карточка-дыра с CTA →
   focus+scrollIntoView на «Сгенерировать»; generate не вызывается.
3a. «Сгенерировать» (и CTA дыры) после листания открывает диалог генерации с
   границами видимого месяца, а не верхнего периода.
4. Tx `other` без ПЛ — обычная карточка, не дыра. ПЛ без tx — обычная
   карточка рабочего дня.
5. Summary-строка: итого ПЛ, км, выдано литров за видимый месяц.
6. Пустой месяц: сетка + «нет движений», лента пустая, автопрыжков нет.
7. Смена верхнего периода и сворачивание строки сбрасывают видимый месяц.
8. Клик по ПЛ-карточке открывает `WaybillDayDrawer`; красный ПЛ ≠ дыра.
9. `npm test -- --run src/features/gsm/` зелёный; `npm run build` зелёный.

## Open Questions

- Нет. Дефолты 2026-08-24: все tx в ленте; фильтра «только дыры» нет;
  карточка tx минимальная; месяц сбрасывается при смене периода/сворачивании;
  порядок ленты по возрастанию дат; генерация целится в видимый месяц.
