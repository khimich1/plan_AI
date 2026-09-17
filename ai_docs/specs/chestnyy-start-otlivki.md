# Spec: Честный старт отливки в «Ёмкости»

> **Источник:** [`ai_docs/ideas/chestnyy-start-otlivki.md`](../ideas/chestnyy-start-otlivki.md) — decisions locked 2026-09-04
> **Дата:** 2026-09-04
> **Статус:** в работе — approve на выполнение дан.
> **План:** [`ai_docs/develop/plans/2026-09-04-chestnyy-start-otlivki.md`](../develop/plans/2026-09-04-chestnyy-start-otlivki.md)
> **Связанные:** [`kalendar-periodov-v-proizvodstvo.md`](kalendar-periodov-v-proizvodstvo.md) (сетка, жильцы, полоса; эта спека **поправляет** цвет клеток и «Начало»),
> [`nedelnye-korziny-obeshchaniy.md`](nedelnye-korziny-obeshchaniy.md) (корзины, гейт, холд),
> `PromisePeriodCalendar.tsx`, `PromiseQuoteBlock.tsx`, `PromiseWindowBand.tsx`, `core/production/promise_buckets.py`

Поправка к календарю периодов: клетка снова показывает занятость **плана**, но знаменатель — **ручка архива**. Клик по-прежнему выбирает неделю. Холды по-прежнему не красятся в дни.

---

## Assumptions (locked)

```
ASSUMPTIONS:
1. Недельная формула корзин не меняется:
   free = capacity − planned − promised; холд не ест freely.
   weeks[] и журнал hold/promise остаются.
2. Потолок дня для клеток, старта и укладки = knob (tracks_per_day
   архива), не DayInfo.max из календарного плана.
   occupied = days_info.occupied (тот же источник, что у котировки).
3. Укладка дат (A5): в день берём
   min(day_free, остаток weeks[week].free на эту укладку, remaining).
   weeks[].free уже = capacity − planned − promised; холд не ест free.
   Чужие обещания не раскладывать по пустым дням. Если free недели 0 —
   дни этой недели не старт, даже при дневном остатке.
4. first_pour_date = первый рабочий день >= завтра, где
   day_free > 0 И у недели ещё есть free (с учётом уже взятого в этой
   укладке). occupied > knob → day_free = 0, клетка «4/3», не старт.
5. solo_date = день последней порции. solo_week_end_date =
   promised_date = последний рабочий день ISO-недели solo_date.
   Гейт сравнивает срок с этим promised_date.
   Аллокации холда = сумма дневных take по monday недели.
6. Жёлтое: все календарные дни от first_pour_date включительно
   до воскресенья недели solo_date включительно. Нерабочие —
   жёлтый фон, серый текст, без дроби. Дни до first_pour_date
   в той же неделе (полные пн–вт) НЕ жёлтые.
7. earliest_start_week остаётся в API (понедельник недели
   first_pour_date) для совместимости. UI «Начало» рисует
   first_pour_date, не понедельник.
8. Клик по клетке: выбирает ISO-неделю (жильцы) И запоминает
   кликнутый день для строки «D.MM: occupied/knob, остаток N».
   Поле срока не меняется. Состав плит не показываем.
9. Occupancy на GET promise-quote: карта дата → occupied на span
   недель котировки (как holidays). Клиент не ходит в
   /production/calendar ради этой сетки.
10. Нет миграции SQLite. Нет новых зависимостей.
    MonthCalendarGrid / FactoryMiniCalendar не drop-in.
→ Correct me now if wrong.
```

---

## Objective

Менеджер в «В производство» видит занятость плана в клетках «Ёмкости» в тех же дорожках, что ручка, и дату **начала отливки**, совпадающую с первым днём, куда заказ ещё влезает — не с понедельником корзины.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | открыл «Ёмкость» при плане 7–9.09 | видеть `3/3`, `3/3`, `1/3`, не пустые клетки | не сверять с «Календарным планом» вручную |
| US-2 | смотрю котировку | «Начало: 9.09 · остаток 2 дор.», не 7.09 | не решить, что льём с понедельника |
| US-3 | смотрю «Если только его» | дату от укладки с 9.09 (~5 дор. → 10.09), не 8.09 пустого завода | третья дата не была раньше начала |
| US-4 | заказ влезает в остаток недели | жёлтое с 9-го до воскресенья 13.09 | полные 7–8 не выглядели окном этого КП |
| US-5 | заказ не влезает в остаток недели | жёлтое сплошь до воскресенья недели, куда укладка ушла | видеть перенос, не «окно с понедельника» |
| US-6 | кликнул 9.09 | жильцы недели + строка «9.09: 1/3, остаток 2» | понять день без состава плит |
| US-8 | на неделе чужое обещание съело `free` | «Начало» не 9.09, если корзина недели пуста; 10-е в клетках пустое | не сесть поверх чужой брони |

### Acceptance criteria (MVP)

- [ ] Клетки «Ёмкости»: на рабочих днях с `occupied > 0` дробь `{occupied}/{knob}`; при `occupied > knob` — `4/3` и отличимый вид перебора; пустые рабочие дни без дроби (как пустые в календарном плане).
- [ ] `PromiseQuoteBlock`: «Начало» = `first_pour_date` + остаток этого дня; «Если только его» = `solo_date` от укладки; «Соло + до конца недели» = `solo_week_end_date`; «Обещать к» = тот же `promised_date`.
- [ ] Жёлтый фон клеток: `first_pour_date ≤ day ≤ sunday(week(solo_date))`. Полоса в диалоге подписывает этот же интервал (левая дата — день старта, не понедельник, если они разные).
- [ ] Маркер «дата клиенту» на `promised_date` остаётся. Отдельный маркер/акцент на `first_pour_date`.
- [ ] Клик: неделя + occupants; строка кликнутого дня; `execution_terms` не меняется.
- [ ] Под сеткой: «Уже в плане: N» как сейчас **и** «Свободно: X из Y» из `weeks[]` выбранной недели (X = `free`, Y = `capacity`).
- [ ] Гейт перевода использует новый `promised_date` (укладка + конец недели).
- [ ] Чужое `promised`, дающее `free = 0` на неделе плана 7–9: `first_pour_date` не в этой неделе; клетки 10–11 без дроби обещания.
- [ ] Регресс: холды не в `free`; occupants; week_count; атомарность перевода. `MonthCalendarGrid` не подключён.

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2, SQLite |
| Domain | `core/production/promise_buckets.py` — новые чистые функции укладки |
| Frontend | React 19, TS; `features/factory-capacity`, `MoveToProductionDialog` |
| Тесты | pytest `tests/`, vitest `frontend/src/` |

Новых внешних зависимостей нет.

---

## Commands

```bash
source venv/bin/activate

pytest tests/test_promise_buckets.py tests/test_promise_service.py \
  tests/test_archive_endpoints.py tests/test_move_to_production_atomicity.py -q

cd frontend && npm run test -- --run \
  src/features/factory-capacity \
  src/features/commercial-archive/components/MoveToProductionDialog.test.tsx
cd frontend && npm run typecheck
```

---

## Project Structure

```
core/production/promise_buckets.py
  + day_free(occupied, knob) -> int
  + pack_pour(...) -> PourPlan | None   # дни + лимит weeks[].free
  build_quote: solo_*, window, холд-аллокации из PourPlan (не из allocate);
               weeks[] как сейчас; earliest_start_week = monday(first_pour)

app/schemas/archive.py
  + PromiseQuoteResponse.first_pour_date, first_pour_free
  + PromiseQuoteResponse.occupancy: dict[str, int]   # ISO date → occupied
  PromiseQuoteWindow: from_week/to_week как ISO-понедельники недель,
    пересекающих жёлтый интервал; promised_date с PourPlan
  (опционально не ломать from_week как monday — см. API)

app/services/promise_service.py
  _quote_to_response: occupancy span как holidays; новые поля

frontend/.../api/promiseQuote.ts
  типы first_pour_date, first_pour_free, occupancy

frontend/.../components/PromiseQuoteBlock.tsx
  Начало + остаток; соло с укладки

frontend/.../components/PromiseWindowBand.tsx
  левая дата = first_pour_date (проп или расширенный window)

frontend/.../components/PromisePeriodCalendar.tsx
  occupancy, knob, pourFrom, pourToSunday, firstPourDate;
  дробь / перебор / жёлтый интервал по дням

frontend/.../components/PromiseWeekOccupants.tsx
  + свободно X из Y (пропсы capacity, free с выбранной weeks[])

frontend/.../commercial-archive/MoveToProductionDialog.tsx
  selectedDay + selectedWeek; прокинуть occupancy/knob/pour
```

---

## Domain: дневная укладка

Чистые функции в `promise_buckets.py` (без I/O). `knob = clamp_promise_knob`.

```python
def day_free(occupied: int, knob: int) -> int:
    return max(0, int(knob) - max(0, int(occupied)))


@dataclass(frozen=True, slots=True)
class PourPlan:
    first_pour_date: date
    first_pour_free: int
    solo_date: date
    solo_week_end_date: date
    allocations: tuple[tuple[date, int], ...]  # дни с take > 0
```

`week_allocations(pour)` — сумма `take` по `iso_week_start` для журнала холда/promise.

Алгоритм `pack_pour(tracks, occupancy, weeks, *, today, knob, is_workday, horizon_days=366*2)`:

1. Если `tracks <= 0` → `None`.
2. Словарь недели → `WeekBucket` (ключ `week_start`). `taken_in_week: dict[date, int] = {}`.
3. Идём по дням `d = today+1` … горизонт. Нерабочий — пропуск.
4. `day_left = day_free(occupied(d), knob)`. `week = iso_week_start(d)`.
   `week_left = bucket.free − taken_in_week[week]` (нет корзины → 0).
   `take = min(remaining, day_left, week_left)`. Если `take > 0` — append, обновить remaining и taken_in_week.
5. Если ни одного take → `None`. Если remaining > 0 в конце горизонта → `None`.
6. `first_pour_date = allocs[0][0]`; `first_pour_free = day_free(occupied(first), knob)` (остаток **дня** до нашей порции, для подписи).
7. `solo_date = allocs[-1][0]`; `solo_week_end_date = last_workday_of_week(monday(solo))`.

`build_quote`:

- `weeks` / `build_weeks` **как сейчас** (план+обещано в `free`).
- `allocate()` для дат человека и холда **не использовать**. Окно = `PourPlan`: `from_week` / `to_week` = понедельники первой и последней недели с take; `promised_date = solo_week_end_date`; `allocations = week_allocations(pour)`.
- UI «Начало» / соло / жёлтое — из `PourPlan`, не `nth_workday` с пустого завода.
- Нет `PourPlan` → `window` и даты как при отсутствии окна сейчас.

Жёлтый интервал UI:

```
pour_from = first_pour_date
pour_to_sunday = iso_week_start(solo_date) + 6 days
```

Полоса диалога: подпись `{pour_from} – {pour_to_sunday}`; маркер клиенту на `promised_date`.

Укладка **может резать** заказ по неделям, даже если `tracks ≤` ёмкости одной полной недели: взяли остаток `free` этой (среда–пятница), хвост — следующая. Это намеренно расходится с `allocate()` «целиком в первую неделю, где free ≥ tracks».

### Пример A (план, обещано 0)

`today = 2026-09-04`, ручка 3, `tracks = 5`, occupied: 7.09→3, 8.09→3, 9.09→1. `promised = 0` → `free` недели ≈ 8.

| Поле | Значение |
|------|----------|
| first_pour_date | 2026-09-09 |
| first_pour_free | 2 |
| укладка | 9.09: 2; 10.09: 3 |
| solo_date | 2026-09-10 |
| promised_date / solo_week_end_date | 2026-09-11 |
| жёлтое | 2026-09-09 … 2026-09-13 |
| холд allocations | `(2026-09-07, 5)` |
| earliest_start_week | 2026-09-07 (JSON; UI не как «Начало») |

### Пример B (чужое обещание съело неделю)

Тот же план, `promised = 8` → `free = 0`. Дни 10–11 в клетках пустые.

| Поле | Значение |
|------|----------|
| take на 7–13 | нет |
| first_pour_date | первый день **следующей** недели с day_free > 0 и free > 0 (типично 2026-09-14) |
| жёлтое | с того дня до воскресенья недели конца укладки |
| холд | ISO-недели, откуда взяли дорожки, не неделя 7.09 |

---

## API contract

### Расширение `GET …/promise-quote`

Новые поля (остальное совместимо):

```python
first_pour_date: date | None = None
first_pour_free: int = Field(ge=0, default=0)
occupancy: dict[str, int] = Field(default_factory=dict)
# ключ YYYY-MM-DD, значение occupied ≥ 0; дни без ключа = 0
```

`occupancy` на span `weeks[0].week_start` … последняя неделя+6, только дни с occupied ≠ 0 (клиент считает 0 по умолчанию). Можно отдать все дни span — не обязательно.

`window.promised_date` = `PourPlan.solo_week_end_date`.  
`solo_date` / `solo_week_end_date` с `PourPlan`.  
`earliest_start_week` = понедельник `first_pour_date` (не показывать как старт отливки).

Холд: `promised_date` и `allocations` с `PourPlan` / `week_allocations` (не из старого `allocate()`).

---

## Code Style

Укладка — в `core/`, сервис только кладёт поля в response. Фронт не считает `pack` сам: даты с API; дроби клеток — `occupancy[iso]` и `quote.knob`.

```text
Начало: {D.MM} · остаток {first_pour_free} дор.
Если только его: {D.MM}
Соло + до конца недели: {D.MM}
```

Нет `first_pour_date` → «Начало: —», остаток не пишем.

Occupants, копирайт плана без изменений, плюс:

```text
Свободно: {free} из {capacity}
```

Строка дня (drawer, под планом/свободно или над жильцами):

```text
{D.MM}: {occupied}/{knob}, остаток {free}
```

Перебор: `{occupied}/{knob} · перебор`. Клик по пустому дню: `{D.MM}: свободно {knob}`. Нерабочий: `{D.MM}: нерабочий`.

---

## UI mechanics

**Диалог**

- `PromiseQuoteBlock` — новые подписи дат и остаток.
- `PromiseWindowBand` — левая дата `first_pour_date`, правая воскресенье жёлтого интервала или `promised_date` как «дата клиенту» (как сейчас справа). Не подписывать понедельник, если старт в среде.

**Drawer «Ёмкость»**

- Календарь: дробь, перебор, жёлтый **по дням** (не целая строка пн–вс, если пн–вт вне интервала).
- Маркер клиенту на `promised_date`; маркер старта на `first_pour_date` (не та же точка).
- Клик: `onSelectWeek(monday)` + `onSelectDay(iso)`.
- Occupants: план, свободно X из Y, строка дня, список холд/promise, подпись про холды.

**Не делать:** кисть; клик пишет срок; дробь `occupied/5`; жёлтая вся неделя при старте в среде.

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| pytest buckets | `day_free`; пример A: first 9, solo 10, week_end 11, холд (7.09, 5); пример B: promised съел free → first не на 7–13; перебор 4/3 → day_free 0; горизонт → None |
| pytest quote/service | response содержит first_pour_*, occupancy; «Начало» в UI-тестах не 31.08 если occupied закрыл начало недели; гейт сравнивает с новым promised_date |
| pytest регресс | allocate/weeks формула; holds не в free; occupants; atomicity |
| vitest calendar | 7 и 8 не aria/жёлтые при pour_from=9; 9 жёлтый; дробь 3/3 и 1/3; 4/3 перебор; клик 9-го → week 7.09 и selected day 9 |
| vitest quote block | «Начало: 9.09» + остаток 2; соло 10.09; не 7.09 |
| vitest band | левая дата 9.09, не 7.09 |
| vitest occupants | «Свободно: 8 из 15»; строка дня |
| vitest dialog | поле срока не меняется по клику |

---

## Boundaries

**Always:**

- Знаменатель клеток = knob котировки.
- Холды не вычитать из `weeks[].free`.
- Клик не пишет `execution_terms`.
- pytest + vitest по командам выше.

**Ask first:**

- Смена формулы `weeks[].planned` при `occupied > knob` (сейчас сумма occupied как есть).
- Удаление поля `earliest_start_week`.
- Имена КП из файлов плана в клетке.
- Вернуть `allocate()` «целиком в первую неделю» вместо нарезки хвоста по `free`.

**Never:**

- Drop-in `MonthCalendarGrid` / `FactoryMiniCalendar`.
- Знаменатель `DayInfo.max` (5) в «Ёмкости».
- «Начало» = `window.from_week` в UI.
- Соло = `nth_workday` с пустого завода.
- Холды и чужие обещания цветом дня (раскладка брони по пустым клеткам).
- Дневной журнал холда («с среды»).
- Фантомы в `days_info`.
- Новые зависимости / миграции.

---

## Success Criteria

1. На данных примера (занятость 3,3,1 при ручке 3, ~5 дорожек) диалог показывает Начало 9.09, соло 10.09, обещать к 11.09; 7.09 в «Начало» нет.
2. В «Ёмкости» 7–8 с дробью полные и не жёлтые; 9–13 жёлтые; точка клиенту 11.09; точка/акцент старта 9.09.
3. Клик по 9.09 не меняет поле срока; видны жильцы недели 7.09 и строка остатка 2.
4. Перевод с датой раньше 11.09 в примере — отказ гейта с earliest = 11.09.
5. Регресс холдов/корзин/occupants зелёный.
6. Пример B: при `free = 0` на неделе плана «Начало» не 9.09; пустые 10–11 без чужой брони в клетке.

---

## Not Doing

- Календарь планирования / DayDrawer / кисть / состав плит.
- Унификация ручки 3 и max 5 в мастере плана.
- Размазывать обещания/холды по пустым дням.
- Дневной холд «с среды».
- Мобильная вёрстка.
- Доска загрузки на весь архив.

---

## Open Questions

Нет блокирующих.

Неблокирующие (в MVP зафиксировано «как ниже»):

- Выходные внутри жёлтого: фон жёлтый, текст серый.
- `weeks[].planned` при occupied > knob: без изменения (сумма occupied). Клетка всё равно `4/3`.

---

## Resolved при написании спеки

- Знаменатель 5 отвергнут (ложь на 7-е при ручке 3).
- Старт ≠ первая пустая клетка (10.09): остаток среды считается.
- Соло пересчитывается от честного старта, не «пустой завод».
- Жёлтое с дня отливки до воскресенья последней недели, не пн–вс целиком.
- Гейт следует новому `promised_date`.
- `earliest_start_week` оставляем в JSON, прячем из подписи «Начало».
- Occupancy на quote, не второй calendar endpoint.
- A5: чужие обещания = лимит `weeks[].free`, не клетки. Холд неделями из укладки, не «с среды». `allocate()` для дат/холда не используем.
