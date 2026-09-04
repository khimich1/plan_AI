# Spec: Календарь периодов в «В производство»

> **Источник:** [`ai_docs/ideas/kalendar-periodov-v-proizvodstvo.md`](../ideas/kalendar-periodov-v-proizvodstvo.md) — decisions locked 2026-09-04
> **Дата:** 2026-09-04
> **Статус:** draft — SPECIFY+PLAN записаны, ждут approve. IMPLEMENT не начинать.
> **План:** [`ai_docs/develop/plans/2026-09-04-kalendar-periodov-v-proizvodstvo.md`](../develop/plans/2026-09-04-kalendar-periodov-v-proizvodstvo.md)
> **Связанные:** [`nedelnye-korziny-obeshchaniy.md`](nedelnye-korziny-obeshchaniy.md) (модель корзин, гейт, холд),
> [`zavod-emkost-vizual-gate.md`](zavod-emkost-vizual-gate.md) (подневной гейт — не копировать клетку=день в этот диалог),
> [`chestnyy-start-otlivki.md`](chestnyy-start-otlivki.md) (**поправка 2026-09-04:** дробь `occupied/ручка` в клетках и «Начало» = первый день с остатком; клик по-прежнему неделя, холды не красить в дни),
> `MoveToProductionDialog.tsx`, `PromiseWeekStrip.tsx`, `app/services/promise_service.py`

---

## Assumptions (locked)

```
ASSUMPTIONS:
1. Это UI поверх уже существующей модели корзин. Семантика
   core/production/promise_buckets.py, гейт move-to-production, журнал
   холд/promise, TTL холда, free = capacity − planned − promised —
   НЕ меняются. Холды по-прежнему не вычитаются из freely.
2. Единица загрузки = ISO-неделя (week_start = понедельник). Клетка дня
   в сетке — линейка дат + праздник/выходной. Занятость дорожек в клетке
   и «Начало» ≠ понедельник — спека chestnyy-start-otlivki (поправка);
   клик всё равно выбирает неделю, холды в дни не красить.
3. Жильцы v1: активные hold + promise аллокации этой недели поимённо
   (kp_id, customer_name, tracks, promised_date, kind). План — число
   planned из котировки («уже в плане: N»), без имён КП из файлов плана.
   created_by в строке жильца нет.
4. Заказы бывают далеко вперёд. week_count котировки = 26 (~полгода),
   не 12: иначе «нет окна» обрежет далёкий слот. Формула корзин та же.
   Сетка: месяцы от первой недели котировки до max(последняя неделя котировки,
   месяц даты в поле срока). Неделя вне weeks[] кликабельна; GET occupants
   работает на любой понедельник. Жёлтое окно — только quote.window
   (куда система кладёт дорожки), не месяц даты клиенту, если он позже.
5. Праздники/extra_workdays: поля GET promise-quote на span сетки
   (недели котировки ∪ месяц срока). GET /production/work-calendar менеджеру
   недоступен. Подневной capacity-snapshot в диалоге НЕ используем ни для
   раскраски дней, ни как красный гейт: в «В производство» источник правды —
   корзины (спека nedelnye-korziny). Snapshot остаётся у графика поставок.
6. Клик по любой клетке недели выделяет неделю для списка жильцов.
   Поле «Срок выполнения» клик НЕ меняет. Срок — только текстовое поле.
7. Роли как у котировки: admin, manager. Общий архив: в списке жильцов
   видны чужие КП. Текущее КП тоже в списке (is_current): иначе свой холд
   на неделе не виден, хотя бейдж есть. Котировка по-прежнему exclude
   текущее КП из сумм held/promised, чтобы не съесть своё место.
8. Drawer «Ёмкость» остаётся слева; ширина ≥ 380px, можно чуть шире ради
   клеток; раскладка всегда колонка: сетка → жильцы → ручка. Не две колонки.
9. PromiseWeekStrip (карточки план/обещано/холды/свободно) убираем из
   диалога и из drawer. В диалоге — компактная жёлтая полоса окна.
10. Нет миграции схемы SQLite. Нет новых зависимостей.
→ Correct me now if wrong.
```

---

## Objective

Менеджер, открыв «В производство», видит **дату клиенту на календарной линейке** и в панели «Ёмкость» — **какие недели заняты какими КП** (холд / обещано), не путая TTL холда с сроком заказа и не принимая клетку дня за бронь отливки.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | менеджер открыл диалог | видеть жёлтое окно и точку «дата клиенту» (18.09) | не искать пятницу в тексте котировки |
| US-2 | менеджер нажал «Ёмкость» | видеть месяц, полосы недель, кликом — кто на неделе | понять загрузку без сложения четырёх цифр |
| US-3 | на неделе холд и свободно оба > 0 | увидеть КП-холд и подпись, что место не жёсткое | не решить, что 15 холдов = неделя закрыта |
| US-4 | в архиве бейдж холда | видеть «к 18.09 · до вечера» | не прочитать TTL как дату клиенту |
| US-5 | менеджер кликнул 10-е в жёлтой полосе | открылся список недели, срок в поле не сдвинулся | не обещать «на десятое» |
| US-6 | срок клиенту через несколько месяцев | листать месяц этой даты в Ёмкости | увидеть, пустые ли дальние недели |

### Acceptance criteria (MVP)

- [ ] Диалог: `PromiseQuoteBlock` как сейчас; вместо `PromiseWeekStrip` — полоса окна (`from_week`…`to_week`, маркер `promised_date`). Поле срока и кнопки без изменения контракта.
- [ ] Drawer «Ёмкость»: месяц пн–вс; нерабочие дни серые; полоса недели (строка) — единица клика; жёлтый фон строк, пересекающих `window`; маркер на `promised_date`.
- [ ] Клик по дню = `week_start` этой ISO-недели; поле `execution_terms` не меняется; под сеткой список жильцов.
- [ ] `GET …/promise-weeks/{week_start}/occupants`: холды+обещания поимённо, `planned`, без `created_by`; текущее КП с `is_current`.
- [ ] Строка плана: «Уже в плане: N дорожек». Пустой журнал: «На этой неделе нет холдов и обещаний». Подпись про холды (не едят свободно).
- [ ] Бейдж архива и шапки карточки: `к Д.ММ · до вечера` из `hold.promised_date`.
- [ ] `PromiseWeekStrip` не рендерится в диалоге и drawer.
- [ ] Не используются `MonthCalendarGrid` / `FactoryMiniCalendar` как drop-in.
- [ ] `MoveToProductionDialog` не блокирует submit по `isCapacityRed` / capacity-snapshot (гейт корзин + парсинг срока остаются).
- [ ] `PromiseService` default `week_count=26`; регресс allocate/quote зелёный.
- [ ] Модель корзин (формула, холды не в free) и гейт перевода без изменений семантики.

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`) |
| Domain | reuse `promise_buckets.py`, `work_calendar.is_working_day`, `PromiseRepository` |
| Frontend | React 19, TS, Vite, TanStack Query; `features/factory-capacity` + `commercial-archive` |
| Тесты | pytest `tests/`, vitest `frontend/src/` |

Новых внешних зависимостей нет.

---

## Commands

```bash
source venv/bin/activate

# Backend — новый контракт + регресс корзин
pytest tests/test_promise_service.py tests/test_archive_endpoints.py \
  tests/test_move_to_production_atomicity.py -q

# Frontend
cd frontend && npm run test -- --run \
  src/features/factory-capacity \
  src/features/commercial-archive/components/MoveToProductionDialog.test.tsx \
  src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx \
  src/features/commercial-archive/components/ArchiveOfferList.test.tsx
cd frontend && npm run typecheck

./run+logs.sh
```

---

## Project Structure

```
app/schemas/archive.py
  + PromiseQuoteResponse.holidays, extra_workdays  (ISO dates, span of weeks[])
  + PromiseWeekOccupant, PromiseWeekOccupantsResponse

app/api/v1/endpoints/archive.py
  + GET /{kp_id}/promise-weeks/{week_start}/occupants
  GET promise-quote — те же роли, расширенный response_model

app/services/promise_service.py
  + list_week_occupants(kp_id, week_start, user)
  + holidays/extra в _quote_to_response (span weeks)

app/repositories/promise_repository.py
  + list_week_allocs(week_start, kinds=("hold","promise"))  # active; expire holds on read

frontend/src/features/factory-capacity/
  api/promiseQuote.ts          → types + usePromiseWeekOccupantsQuery
  components/PromiseWindowBand.tsx      → NEW: полоса окна в диалоге
  components/PromisePeriodCalendar.tsx  → NEW: месяц + клик по неделе
  components/PromiseWeekOccupants.tsx   → NEW: список под сеткой
  components/PromiseWeekStrip.tsx       → не использовать в диалоге/drawer
                                          (можно оставить файл, если нужен тест/история)

frontend/src/features/commercial-archive/components/
  MoveToProductionDialog.tsx   → band вместо strip; drawer = calendar + occupants + knob;
                                 убрать useCapacitySnapshotQuery / isCapacityRed из этого диалога
  OfferDetailsDrawer.tsx       → текст бейджа
  ArchiveOfferList.tsx         → текст бейджа
```

---

## API contract

### Расширение `GET /api/v1/commercial/archive/{kp_id}/promise-quote`

Как сейчас, плюс два поля (пустые списки, если нет):

```python
holidays: list[date]         # нерабочие в span(weeks[0].week_start … last+6)
extra_workdays: list[date]   # рабочие выходные в том же span
```

Считаются через тот же `is_workday`, что и корзины. Клиент не ходит в `/production/work-calendar`. Месяц срока за пределами `weeks[]`: клетки вне span — серые только как сб/вс (без extra_workdays); клик по неделе и occupants всё равно работают.

### Новый `GET /api/v1/commercial/archive/{kp_id}/promise-weeks/{week_start}/occupants`

- Роли: `admin`, `manager`.
- `kp_id`: существует, `assert_offer_read_access` (как quote).
- `week_start`: ISO `YYYY-MM-DD`, **понедельник**. Иначе 422.
- Ленивый expire холдов до выборки (как `sum_held_by_week`).
- Аллокации `status=active`, `kind in (hold, promise)` на эту неделю. **Не** exclude текущего `kp_id`.
- `planned`: как в котировке для этой недели (сумма occupied рабочих дней); если недели нет в горизонте — всё равно посчитать по occupancy+календарю, не 404.
- `customer_name` с карточки КП; пусто → `""` (UI покажет «—»).
- Нет `created_by`.

```python
class PromiseWeekOccupant(BaseModel):
    kp_id: int
    customer_name: str
    kind: Literal["hold", "promise"]
    tracks: int = Field(ge=1)
    promised_date: date
    is_current: bool  # kp_id == path kp_id

class PromiseWeekOccupantsResponse(BaseModel):
    week_start: date
    planned: int = Field(ge=0)
    occupants: list[PromiseWeekOccupant]
```

Порядок `occupants`: `kind` promise, затем hold; внутри — `kp_id`.

Ошибки: 404 КП нет; 422 не понедельник / мусор в дате; 503 occupancy как у quote.

---

## Code Style

Чистый список жильцов — в репозитории; сервис клеит `customer_name` и `is_current`. Роутер тонкий.

```python
# app/services/promise_service.py
def list_week_occupants(
    self, kp_id: int, week_start: date, *, user: dict
) -> PromiseWeekOccupantsResponse:
    raw = self._load_kp(kp_id)
    assert_offer_read_access(user, raw)
    if week_start.weekday() != 0:
        raise PromiseWeekInvalidError("week_start должен быть понедельником.")
    self.repository.expire_stale_holds(now=self._moment())
    rows = self.repository.list_week_allocs(week_start, kinds=("promise", "hold"))
    planned = self._planned_for_week(week_start)
    occupants = tuple(
        PromiseWeekOccupant(
            kp_id=row["kp_id"],
            customer_name=self._customer_name(row["kp_id"]),
            kind=row["kind"],
            tracks=row["tracks"],
            promised_date=row["promised_date"],
            is_current=row["kp_id"] == kp_id,
        )
        for row in rows
    )
    return PromiseWeekOccupantsResponse(
        week_start=week_start, planned=planned, occupants=list(occupants)
    )
```

Фронт: клик считает понедельник ISO-недели даты клетки; `onSelectWeek(weekStart)` — единственный колбек сетки. Нет `onSelectDate`.

Бейдж (видимый текст):

```text
к {formatQuoteDayMonth(hold.promised_date)} · до вечера
```

Tooltip с `created_by` можно оставить, в бейдж не выносить.

Копирайт пустого списка: «На этой неделе нет холдов и обещаний».  
Подпись под списком: «Холды не занимают свободно — до перевода место могут взять другие.»  
План: «Уже в плане: {n} дорожек».

---

## UI mechanics

**Диалог**

1. Котировка (`PromiseQuoteBlock`).
2. `PromiseWindowBand`: жёлтый отрезок окна; подпись дат; точка на `promised_date`. Нет `window` — полосу не показывать.
3. Баннер холда, если есть.
4. Поле срока, ошибки парсинга, кнопки. Submit **не** зависит от capacity-snapshot / `isCapacityRed`. Красный блок перевода — только гейт корзин (дата раньше «обещать к») и разбор поля.
5. Нет `PromiseWeekStrip`.

**Drawer «Ёмкость» (открыт с диалогом)**

1. `PromisePeriodCalendar`: месяц по умолчанию = месяц `promised_date` (иначе текущий).
2. Навигация ‹ ›: от месяца `weeks[0]` до max(месяц последней недели котировки, месяц распознанной даты в поле срока). Дальше 26 недель котировки месяц срока всё ещё открывается (клик + occupants).
3. Выбранная неделя по умолчанию = ISO-неделя `promised_date` (иначе `weeks[0]`).
4. Под сеткой `PromiseWeekOccupants` выбранной недели + запрос occupants.
5. Внизу `PromiseKnobSettings` как сейчас.

Клик по серой нерабочей клетке внутри недели всё равно выбирает **неделю** (не день).

Жёлтые строки: `week_start` ∈ [`window.from_week`, `window.to_week`].

**Не делать:** кисть дорожек, клик задаёт срок, красная заливка `occupied/max` по дням, две колонки, карточки четырёх счётчиков.

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| pytest service | `list_week_occupants`: понедельник ok; не понедельник → ошибка; холд+promise; expire холда; текущее КП `is_current`; planned ≥ 0; нет created_by в payload |
| pytest api | 200 форма; 404; 422 week_start; роли manager ok; production 403 |
| vitest calendar | клик 10.09 и 18.09 одной недели → один `week_start`; поле срока в диалоге не меняется (мокированный input) |
| vitest band | маркер promised_date; нет window → пусто |
| vitest occupants | строки КП+клиент+kind; план; пустой текст; холд-подпись |
| vitest badge | «к 18.09 · до вечера» в списке и drawer |
| vitest dialog | нет `promise-week-strip`; есть band; drawer содержит period-calendar; submit не disabled из-за red snapshot |
| pytest quote | default week_count даёт ≥ 26 недель в `weeks[]` (или явный параметр 26) |
| регресс | `test_move_to_production_atomicity.py`, существующие quote/hold тесты |

---

## Boundaries

**Always:**

- Гейт перевода (корзины) и формула (м/101, холды не в free) без правок «заодно». Default `week_count=26` — да, в скоупе.
- Occupants expire холды перед чтением.
- `week_start` только понедельник.
- pytest + vitest зелёные по командам выше.

**Ask first:**

- Имена КП из производственных планов (провал A1).
- Ширина drawer > 480px или перенос сетки в модалку.
- `week_count` > 26.
- Удаление файла `PromiseWeekStrip.tsx`, если на него ещё ссылаются тесты/другие экраны.

**Never:**

- Drop-in `MonthCalendarGrid` / `FactoryMiniCalendar`.
- Фантомные дорожки в `days_info`.
- Клик по дню пишет срок.
- Новые статусы КП / секции архива.
- Расширение `GET /production/work-calendar` на manager ради этой фичи.
- Красный гейт `capacity-snapshot` в диалоге «В производство» (оставить в графике поставок).

---

## Success Criteria

1. На КП с холдом к 18.09 менеджер без открытия ёмкости видит жёлтое окно и точку 18.09 в диалоге; в архиве бейдж «к 18.09 · до вечера».
2. В «Ёмкость» клик по любому дню недели 7.09–13.09 показывает одних и тех же жильцов; значение поля срока не меняется.
3. Неделя с холдами 15 и свободно 15 показывает именной холд и подпись «не занимает свободно», а не красные пустые дни.
4. Регресс перевода в производство (атомарность + гейт корзин) зелёный.
5. Поле срока «март следующего года»: в Ёмкости открывается этот месяц; жёлтого окна там нет, если allocate положил КП раньше; список жильцов недели пустой или чужие дальние обещания.
6. Кнопка «Перевести» не серая только из-за старого дневного светофора, если корзины зелёные.

---

## Not Doing

- Календарь планирования / кисть / DayDrawer.
- Поимённый план из `days` планов.
- Доска загрузки на весь архив.
- Смена формулы корзин (м/101, холды ≠ free). `week_count` 12→26 — в скоупе.
- Мобильная вёрстка.

---

## Open Questions

Нет блокирующих. Ширина drawer — на вёрстке (≥ 380px). A1 (имена из плана) — отдельная идея после ручной проверки.

---

## Resolved при написании спеки

- Праздники: поля на `promise-quote`, не production work-calendar (роли).
- Occupants — отдельный GET по клику, не раздувать каждую котировку списком всех недель.
- Текущее КП в списке жильцов (`is_current`): иначе свой холд не виден; котировка по-прежнему exclude его из сумм.
- `week_count=26` и месяц поля срока в сетке — дальние заказы.
- Красный snapshot-гейт убираем из этого диалога (корзины заменили его здесь; график поставок не трогаем).
- `PromiseWeekStrip` не удаляем файл, пока grep не покажет ноль ссылок; из UX диалога/drawer убираем.
