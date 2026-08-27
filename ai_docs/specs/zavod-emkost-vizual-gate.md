# Spec: Визуал ёмкости завода + жёсткий гейт

> **Источник:** [`ai_docs/ideas/zavod-emkost-vizual-gate.md`](../ideas/zavod-emkost-vizual-gate.md)  
> **UX follow-up:** [`ai_docs/ideas/zavod-emkost-left-drawer.md`](../ideas/zavod-emkost-left-drawer.md) ·  
> [`ai_docs/specs/zavod-emkost-left-drawer.md`](zavod-emkost-left-drawer.md) — панель в left drawer по кнопке «Ёмкость»  
> **Plan:** [`ai_docs/develop/plans/2026-08-25-zavod-emkost-vizual-gate.md`](../develop/plans/2026-08-25-zavod-emkost-vizual-gate.md)  
> **Handoff:** [`ai_docs/develop/handoffs/2026-08-25-zavod-emkost-vizual-gate.md`](../develop/handoffs/2026-08-25-zavod-emkost-vizual-gate.md)  
> **Orchestration:** `orch-2026-08-25-16-50-zavod-emkost-gate`  
> **Дата:** 2026-08-25  
> **Статус:** approved → IMPLEMENT + UX drawer  
> **Связанные:** `docs/specs/delivery-schedule.md`, `core/delivery_schedule_check.py`,  
> `MoveToProductionDialog`, `DeliveryScheduleEditor`, `MonthCalendarGrid`

---

## Assumptions (locked after review)

```
ASSUMPTIONS:
1. Занятость = только план: days_info (max − occupied). СГП — OUT of MVP.
2. Симуляция и окно «свободно» стартуют с ЗАВТРА. check_batches параметризуем
   start_date; график поставок в рамках фичи тоже со старта завтра.
3. Пороги светофора = как у партий (green / yellow / red, slack 5 раб.дней,
   buffer 1.15, 101 м, 5 дор/день).
4. Жёлтый = только предупреждение (сохранить можно). Красный = жёсткий гейт.
5. Несколько партий → блок по худшему (любой red). Симуляция последовательная.
6. НЕТ отдельного override «клиент согласовал». Поле срока как сейчас:
   в «В производство» — одна строка (дни / срок), без новых полей.
   Пока статус red — сохранить/отправить нельзя; видна подсказка (hint).
   Менеджер сам увеличивает N дней / сдвигает produce_by после разговора
   с клиентом, пока red не снимется.
7. Аудит override и таблица capacity_gate_override — OUT of MVP.
8. Гейт на backend обязателен (+ UI disable + текст подсказки).
9. «В производство»: КП = одна виртуальная партия; target =
   завтра + N рабочих дней из строки срока (как сейчас парсится execution_terms).
10. Виджет в ДВУХ местах: (а) архив → «В производство»;
    (б) этап согласования / редактор графика поставок.
11. Календарь UI: от текущего месяца до месяца конца заказа (target /
    max produce_by). Навигация между этими месяцами; не фиксированные «всегда 2».
12. Бронь дорожек НЕ делаем. Мобилка не проектируется.
```

---

## Objective

Менеджер на ПК на шагах **«В производство»** и **согласование графика поставок**
видит мини-календарь загрузки завода + сжатые цифры и **не может сохранить
обещание при red**, пока не сдвинет срок в существующем поле (после разговора
с клиентом). Отдельных галочек и дат согласования нет.

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|----------------|---------|--------|
| US-1 | отправляю КП в производство | видеть календарь и «нужно vs свободно» + подсказку | не обещать нереальное |
| US-2 | согласую график поставок | видеть тот же виджет рядом с партиями | не перегрузить завод |
| US-3 | срок красный | не иметь возможности сохранить; понять из подсказки, что менять | позвонить клиенту и увеличить срок в той же строке |
| US-4 | срок жёлтый | сохранить с предупреждением | не тормозить пограничные кейсы |

### Acceptance criteria (MVP)

- [ ] Виджет в `MoveToProductionDialog` **и** в UI графика поставок
      (`DeliveryScheduleDialog` / Editor): мини-календарь + `нужно · свободно · Δ`
- [ ] Диапазон месяцев календаря: **текущий месяц … месяц target**
      (конец заказа / max `produce_by`)
- [ ] Статус по правилам `delivery_schedule_check`, старт = завтра
- [ ] При red: кнопки «В производство» / «Сохранить график» disabled;
      видна понятная подсказка (hint). Новых полей в форме нет
- [ ] Backend отклоняет те же операции при red (4xx + сообщение)
- [ ] Единственный способ снять блок — изменить срок так, что статус ≠ red
- [ ] Занятость только из плана; нет резерва; нет override/аудита override

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | FastAPI, Pydantic v2, SQLite |
| Domain | `core/delivery_schedule_check.py`, `core/production_capacity.py`, `core/work_calendar.py` |
| Frontend | React 19, TS, Vite, TanStack Query |
| UI | compact read-only календарь (паттерн `MonthCalendarGrid`) |

Новых зависимостей нет. Новой таблицы БД в MVP нет.

---

## Commands

```bash
source venv/bin/activate
pytest tests/test_delivery_schedule_check.py tests/test_capacity_gate.py -q
pytest tests/test_archive_endpoints.py tests/test_delivery_schedule_endpoints.py -q
uvicorn app.main:app --reload

cd frontend && npm run test -- --run
cd frontend && npm run typecheck
cd frontend && npm run build
cd frontend && npm run dev

./run+logs.sh
```

---

## Project Structure

```
core/
  delivery_schedule_check.py   # start_date; без смены порогов
  production_capacity.py
  work_calendar.py

app/
  schemas/                     # CapacitySnapshot (без override)
  services/
    capacity_gate_service.py   # occupancy + check + enforce (block if red)
    archive_service.py         # move_to_production → gate
    delivery_schedule_service.py  # PUT → gate
  api/v1/endpoints/            # GET capacity-snapshot

frontend/src/features/
  factory-capacity/            # api/ components/ hooks/ types/
    FactoryCapacityPanel.tsx   # календарь + шапка + hint при red
    FactoryMiniCalendar.tsx
  commercial-archive/
    MoveToProductionDialog.tsx
  delivery-schedule/
    DeliveryScheduleDialog.tsx / Editor

tests/
  test_capacity_gate.py
  test_delivery_schedule_check.py  # start=tomorrow
```

---

## Поведение

### Виджет `FactoryCapacityPanel`

**Шапка:** нужно · свободно в `[завтра … target]` · Δ (цвет по статусу).

**Календарь:** месяцы от текущего до месяца target; read-only; цвета empty/partial/full.

**При red:** Alert/hint из check («нужно +N дорожек…» / «увеличьте срок»).
Без checkbox, без второго поля даты.

### Move → производство

1. Как сейчас: одна строка срока (`execution_terms`, дни).
2. Snapshot пересчитывается при изменении строки.
3. Submit при red → UI block + backend 4xx.
4. Менеджер увеличивает N дней → статус yellow/green → можно отправить.

### График поставок (этап согласования)

1. Тот же виджет рядом с редактором партий.
2. PUT при любом red → reject.
3. Менеджер двигает `produce_by` / состав → live-пересчёт → сохранить.

### Один алгоритм

`core/delivery_schedule_check.check_batches` — единая правда для чипов партий и панели.

---

## Data Model

Новых таблиц в MVP **нет**.

---

## API (эскиз)

| Метод | Путь | Назначение |
|-------|------|------------|
| GET | `/api/v1/commercial/archive/{kp_id}/capacity-snapshot?target=…` | снимок для виджета (move) |
| GET | delivery-schedule (существующий) | статусы партий + при необходимости summary для панели |
| POST | move-to-production | enforce: red → 4xx |
| PUT | delivery-schedule | enforce: any red → 4xx |

Тела запросов **без** `client_agreed_date`.

---

## Code Style

- Математика в `core/`; enforce в services.
- Feature `factory-capacity`; не плодить вторую сетку без нужды
  (Ask first: рефактор `MonthCalendarGrid` vs отдельный mini).
- Пример snapshot:

```ts
type CapacitySnapshot = {
  start_date: string;
  target_date: string;
  tracks_needed: number;
  tracks_free_in_window: number;
  delta: number;
  status: "green" | "yellow" | "red";
  hint: string | null;
  days_info: Record<string, { occupied: number; max: number }>;
  holidays: string[];
  extra_workdays: string[];
  calendar_from_month: string; // YYYY-MM текущего
  calendar_to_month: string;   // YYYY-MM target
};
```

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit check | start=tomorrow; пороги без регрессий |
| Unit gate | red→block; yellow/green→allow; multi-batch worst |
| API | move/PUT 4xx на red; 2xx когда не red |
| Frontend | панель; submit disabled на red; hint виден; без override-полей |

---

## Boundaries

**Always:** серверный гейт; один алгоритм статуса; старт с завтра.

**Ask first:** большой рефактор `MonthCalendarGrid`; смена порогов 1.15 / slack 5.

**Never:** override-галочка; резерв дорожек; СГП в «свободно»; gейт только на клиенте; mobile layout.

---

## Success Criteria

1. Red-заказ: сохранить нельзя; подсказка видна; увеличение дней в той же строке снимает блок.
2. Виджет в архиве (move) **и** на этапе графика поставок; цвета согласованы с чипами партий.
3. Календарь покрывает текущий месяц … месяц конца заказа.
4. Жёлтый сохраняется. Регрессии check/archive зелёные.

---

## Open Questions

_Нет блокирующих._ (при Plan уточнить точную формулу target из строки «N дней», если парсер уже отдаёт дату — переиспользовать.)

---

## Not Doing

- Checkbox / второе поле «клиент согласовал»
- Таблица и аудит override
- СГП в формуле
- Мягкий warning на red без блока
- Сайдбар на весь архив
- Авто-подбор оптимальной даты

---

## Next step

`/orchestrate execute orch-2026-08-25-16-50-zavod-emkost-gate`  
Handoff: `ai_docs/develop/handoffs/2026-08-25-zavod-emkost-vizual-gate.md`
