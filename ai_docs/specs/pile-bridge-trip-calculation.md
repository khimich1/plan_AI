# Spec: Рейсы свай и мостовых свай в КП

> **Источник идеи:** [`ai_docs/ideas/pile-bridge-trip-calculation.md`](../ideas/pile-bridge-trip-calculation.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS ✅ → IMPLEMENT ✅  
> **Статус:** implemented  
> **Plan:** [`ai_docs/develop/plans/2026-09-01-pile-bridge-trip-calculation.md`](../develop/plans/2026-09-01-pile-bridge-trip-calculation.md)  
> **Дата:** 2026-09-01  
> **Связанные модули:** `core/cargo_delivery_pricing.py`, `core/pile_catalog.py`, `core/commercial_pricing.py`, `CalculationResultStep`, archive drawer, `KpPersistenceService`  
> **Суперседит (логистика non-plates):** assumption 8 / Q2 в [`kp-multi-nomenclature-append.md`](kp-multi-nomenclature-append.md) — «рейсы только по плитам» — **только для плит**; сваи/мостовые получают **отдельную** строку доставки по этому spec.

---

## Decisions locked (ideation 2026-09-01)

| # | Тема | Решение |
|---|------|---------|
| D1 | Где | КП (итоги/скидка) + архив |
| D2 | Отгрузки | **Код не меняем.** Импорт в `pile_catalog` даст вес канону `С60.30` как побочный эффект |
| D3 | Справочник | Excel → `pile_catalog`. Отдельного `bridge_pile_catalog` нет |
| D4 | Мостовые | Резолв по **длине × сечению**: `C14-40T4` ≡ `С140.40`. Тn и B25/B30 на массу не влияют |
| D5 | Адреса | Один пул рейсов на КП |
| D6 | Формула (есть норма) | `floor(qty / pcs_per_20t)` по марке + `ceil(Σ остаток_кг / 19800)` по остаткам известных марок |
| D7 | Геометрия остатков | Не в MVP |
| D8 | Нет нормы | Не в полные и не в кучу остатков. Вопрос: **сколько машин (рейсов)** на эти шт. N в этом КП (черновик и архив), **не** в `pile_catalog` |
| D9 | Пока нет ответа | Доставку **свай не считаем** (состав и цена изделий остаются) |
| D10 | Смешанное КП | **Две** строки доставки: плиты (18,6 т) и сваи (D6+D8). В одну фуру не мешаем |
| D11 | Тариф | Вводит менеджер; система считает **число** рейсов |
| D12 | Ж/д / контейнер | Не в MVP |
| D13 | C18 и «-» в авто | Одинаково: нет `pcs_per_20t` → вопрос про машины |
| D14 | PDF/XLSX | Две строки, если обе > 0: «Доставка плит» и «Доставка свай» |
| D15 | Архив | N машин и свайный тариф **можно менять** после save (как `logistics_cost` плит); totals и файлы пересчитать |
| D16 | UI расшифровка | Только **итог** рейсов свай + список марок без нормы. Не показывать «39 полных + 3 остатка» |

---

## Assumptions I'm Making

1. **Web-only** — React-мастер КП + архив; бот вне scope.
2. **Типы свай в этом релизе:** `piles` и `bridge_piles`. ФБС / ЛС / ЛМ **не** получают рейсы по этому spec.
3. **Два тарифа рейса:** `logistics_cost` (плиты, как сейчас) и **`pile_logistics_cost`** (сваи/мостовые). В UI смешанного КП — два поля. В pile-only / bridge-only КП существующее поле «Стоимость рейса» пишет в `pile_logistics_cost`.
4. **НДС** — как сейчас: НДС с изделий; обе доставки плюсуются к «стоимость с НДС» без отдельного НДС на доставку.
5. **`pcs_per_20t IS NULL` или 0** = нет нормы, даже если `weight_kg` есть (`С160.*`).
6. **Пустой ответ ≠ 0.** `N` не введено → доставка свай не считается. Явный **0** — ответ «машин не нужно», доставку свай можно считать (остальные марки + 0).
7. **Ключ override:** нормализованная марка строки (`C18-40T8` / `С18-40Т8` → один ключ). Смена qty не сбрасывает N.
8. **Импорт CLI**, не UI. Расширяем `scripts/import_pile_catalog.py`: лист `Вес и объем` **или** `Лист1`; колонка «автомобильный г/п 20тн»; «-» → NULL.
9. **Lookup:** сначала точный `mark` в `pile_catalog` (C↔С); иначе геометрия length_m + section_mm. Несколько строк с одной геометрией — берём любую с непустым `pcs_per_20t`, иначе любую с весом.
10. **Резолв не найден и нет override** = нет нормы → вопрос про машины (как C18). Вес в кучу остатков не идёт.
11. **PDF/XLSX/архив** показывают те же две (или одну) строки доставки, что итоги КП.
12. **Схема:** `ALTER KP_offers ADD pile_logistics_cost`; overrides в JSON в draft metadata и в `kp_meta` (новая колонка `pile_trip_overrides_json` TEXT).
13. **Константа остатков 19800 кг** живёт рядом с `CARGO_DELIVERY_TRUCK_CAPACITY_KG` (18600 для плит), не подменяем 18600.
14. **Без новых npm/pip зависимостей.**

→ Поправьте до PLAN, если что-то неверно.

---

## Objective

Менеджер КП на сваи и/или мостовые сваи видит **число авторейсов из заводской таблицы** (шт. в машине + остатки по 19,8 т) и сумму доставки = тариф × рейсы. Для марок без нормы система **спрашивает число машин**. Плиты по-прежнему считаются отдельно (18,6 т).

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | делаю КП на мостовые/сваи | чтобы «стоимость рейса» работала | не оставлять доставку нулём |
| US-2 | все марки есть в Excel | увидеть машины без ручного Excel | опереться на заводскую загрузку |
| US-3 | есть C18-40T8 без нормы | чтобы спросили, сколько машин | не подставить выдуманный вес |
| US-4 | не ответил на вопрос | чтобы доставку свай не посчитали | не отдать клиенту ложную логистику |
| US-5 | в КП плиты и сваи | две строки доставки | не мешать фуры ПБ и свай |
| US-6 | открываю архив | те же рейсы и суммы | не пересчитывать в голове |

### Reframed success criteria

| Требование | Критерий |
|------------|----------|
| Гибрид | На фикстуре тендера без C18: 39 полных + 3 остатка = **42** рейса свай |
| C18 | 49 шт. не входят в 42; UI требует N; без N `pile_delivery = 0` и доставка свай не в total |
| Явный 0 | N=0 по C18 → 42 рейса, доставку свай считать можно |
| Плиты | Mono/mixed plates: `ceil(кг / 18600)` без регрессии |
| Mixed | `total_delivery = plate_delivery + pile_delivery` (если сваи «готовы») |
| Справочник | После импорта `банк знаний/сваи вес и объем.xlsx` — 44 марки в `pile_catalog` |
| Отгрузки | Нет diff в `app/services/shipment_*.py` / logistics endpoints |

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`) |
| Domain | `core/pile_catalog.py`, **новый** `core/pile_trip_pricing.py`, `core/cargo_delivery_pricing.py` |
| API | `app/api/v1/endpoints/commercial.py`, archive logistics patch |
| Frontend | React 19, Vite, TypeScript, Vitest |
| Tests | pytest `tests/`, Vitest `frontend/src/features/commercial-offer/` |

---

## Commands

```bash
source .venv/bin/activate

# Импорт справочника
python scripts/import_pile_catalog.py --xlsx "банк знаний/сваи вес и объем.xlsx" --sheet Лист1

# Backend tests
pytest tests/test_pile_catalog_import.py tests/test_pile_trip_pricing.py \
  tests/test_commercial_logistics_cost.py tests/test_commercial_calculation_service.py -q

# Frontend
cd frontend && npm run typecheck && npm run test -- --run
```

---

## Project Structure

```
core/
  pile_catalog.py                 # EXTEND — header detect, Лист1, parse_bridge geometry
  pile_trip_pricing.py            # NEW — hybrid trips + overrides
  cargo_delivery_pricing.py       # EXTEND — константа 19800; не менять 18600
  commercial_pricing.py           # EXTEND — две доставки в totals
  commercial_offer.py             # EXTEND — строки доставки в PDF
  commercial_offer_xlsx.py        # EXTEND
  kp_db_schema.py                 # EXTEND — pile_logistics_cost, overrides json
  kp_persistence_service.py       # EXTEND

scripts/
  import_pile_catalog.py          # EXTEND — default sheet fallback

app/schemas/commercial.py         # EXTEND — pile_logistics_cost, pile_trip_overrides
app/schemas/archive.py            # EXTEND
app/services/commercial_calculation_service.py
app/services/archive_service.py

frontend/src/features/commercial-offer/
  components/steps/CalculationResultStep.tsx   # EXTEND — поле свай + вопросы N
  utils/cargoDeliveryPricing.ts                # EXTEND или pileTripPricing.ts
  types/commercialOffer.ts

frontend/src/features/commercial-archive/      # EXTEND — две доставки, те же вопросы read-only/edit

tests/
  test_pile_trip_pricing.py       # NEW
  test_pile_catalog_import.py     # EXTEND — Лист1 fixture
  test_commercial_logistics_cost.py  # EXTEND — mixed two lines

ai_docs/ideas/pile-bridge-trip-calculation.md
ai_docs/specs/pile-bridge-trip-calculation.md   # этот файл
```

---

## Code Style

Парсер геометрии мостовых — рядом с `parse_pile_mark`, без правок отгрузок:

```python
def parse_bridge_pile_geometry(mark: str) -> tuple[float | None, int | None]:
    """C14-40T4 / С14-40Т4 → (14.0 м, 400 мм)."""
    ...
```

Расчёт рейсов — чистая функция (легко тестировать без FastAPI):

```python
@dataclass(frozen=True)
class PileTripBreakdown:
    full_trips: int
    remainder_kg: float
    remainder_trips: int
    override_trips: int
    pending_marks: tuple[str, ...]  # нет нормы и нет N
    total_trips: int  # 0 если pending_marks не пуст (доставку не считаем)

    @property
    def ready(self) -> bool:
        return not self.pending_marks
```

`pcs_per_20t` из Excel — заводская норма на **20 т**; **19800** только для кучи остатков. Не переименовывать колонку БД.

---

## Architecture

### Flow (шаг результата КП)

```
order_data (piles | bridge_piles | mixed)
    → resolve catalog (mark | geometry)
    → marks with pcs: hybrid D6
    → marks without pcs: require pile_trip_overrides[key]
    → if pending: pile_delivery hidden/0, banner + inputs
    → else: pile_delivery = pile_logistics_cost × total_trips
plates (if any): plate_delivery = logistics_cost × ceil(kg / 18600)   # без изменений
grand delivery = plate_delivery + (pile_delivery if ready else 0)
```

### UI «Итоги и скидка»

| Состав КП | Поля |
|-----------|------|
| Только плиты | Как сейчас: стоимость рейса → `logistics_cost` |
| Только сваи и/или мостовые | Одно поле «Стоимость рейса» → `pile_logistics_cost`; **включено** (сейчас disabled) |
| Плиты + сваи | Два поля: «Рейс плит» и «Рейс свай» |
| Есть марки без нормы | Блок: марка, qty, input «Машин, шт.» (integer ≥ 0). Пока пусто — не применять доставку свай |

Текст вопроса (норматив):  
«Для {марка} ({qty} шт.) нет нормы загрузки в справочнике. Сколько машин нужно?»

### Формула (готовые марки)

```
полные_i     = floor(qty_i / pcs_i)           # pcs_i ≥ 1
остаток_кг   = Σ (qty_i % pcs_i) × weight_i   # только марки с pcs
рейсы_остат  = 0 если остаток_кг == 0 else ceil(остаток_кг / 19800)
рейсы_ручные = Σ N_j                          # только введённые
рейсы_свай   = Σ полные_i + рейсы_остат + рейсы_ручные
```

Фикстура тендера (без C18): полные 17+2+3+3+3+11 = 39; остаток 46950 кг → 3; **42**.

### Data model

**`pile_catalog`** — без новых колонок (уже есть `weight_kg`, `pcs_per_20t`).

```sql
ALTER TABLE KP_offers ADD COLUMN pile_logistics_cost REAL DEFAULT 0;
ALTER TABLE kp_meta ADD COLUMN pile_trip_overrides_json TEXT;  -- {"C18-40T8": 12}
```

**Draft metadata:**

```python
logistics_cost: float = 0.0          # плиты
pile_logistics_cost: float = 0.0     # сваи
pile_trip_overrides: dict[str, int]  # mark_key → N машин
```

### API

| Method | Path | Изменение |
|--------|------|-----------|
| calculate / meta PATCH | commercial drafts | принимать `pile_logistics_cost`, `pile_trip_overrides` |
| save | commercial | писать оба тарифа + json overrides |
| archive details | | отдавать breakdown рейсов свай + `pile_delivery_ready` |
| PATCH logistics | archive | расширить payload: `logistics_cost` как сейчас + `pile_logistics_cost` + `pile_trip_overrides`; без ломания старого поля |

Totals calculate: добавить `plate_delivery_total`, `pile_delivery_total`, `pile_trips`, `pile_trip_pending_marks` (имена уточнить в PLAN, не ломать существующих клиентов без полей — поля новые).

### Импорт

`parse_pile_catalog_from_xlsx`:

1. Предпочесть `--sheet`; иначе `Вес и объем`, иначе `Лист1`.
2. Найти строку заголовка по ячейке «марка» / «вес».
3. Колонка шт. = «автомобильный» / «20т» / 4-я колонка как сейчас.
4. `pcs` «-» / пусто → NULL (вопрос в КП).
5. Upsert по `mark` как сейчас.

---

## Testing Strategy

| Уровень | Где | Что |
|---------|-----|-----|
| Unit | `tests/test_pile_trip_pricing.py` | гибрид 42; C18 pending → not ready; override N; явный 0; `%` остаток; pcs NULL |
| Unit | `tests/test_pile_catalog_import.py` | реальный/фикстурный Лист1; C14 geometry match С140.40 |
| Unit | `tests/test_commercial_logistics_cost.py` | plates 18600 без регрессии; mixed сумма двух доставок; pending не добавляет сваи |
| API | commercial calculate/save | metadata round-trip overrides |
| Vitest | `CalculationResultStep` | поле включено без плит; вопросы C18; mixed два тарифа |
| Не трогать | shipment tests | нет изменений кода отгрузок |

Покрытие: все ветки `PileTripBreakdown.ready` + mixed vs mono.

---

## Boundaries

- **Always:** pytest релевантных наборов до коммита; константы 18600 и 19800 не путать; не писать N в `pile_catalog`.
- **Ask first:** пересборка PDF/XLSX сразу при PATCH N в архиве или только цифры, файлы по кнопке (default PLAN: как у текущего PATCH тарифа плит).
- **Never:** код `shipment_*` / logistics propose; формула веса C18; ж/д колонки Excel в расчёт КП; смешивать плиты и сваи в одном ceil.

---

## Success Criteria

- [x] Импорт `сваи вес и объем.xlsx` → 44 строки, у C9/C10/C11 есть B25-вес и pcs, у С160 pcs NULL
- [x] Тендерные марки кроме C18 резолвятся; расчёт 42
- [x] C18: вопрос + без N доставка свай = 0; с N=k итог 42+k
- [x] КП только мостовые: поле рейса активно
- [x] КП плиты+сваи: две суммы доставки в итоге и в файлах
- [x] Архив показывает то же; PATCH N / свайного тарифа пересчитывает доставку свай
- [x] `git diff` без `app/services/shipment*.py` и logistics propose
- [x] PDF/XLSX mixed: две строки доставки

---

## Open Questions

Нет. D14–D16 закрыты 2026-09-01.

---

## Not in this spec

Отгрузки, геометрия кузова, два адреса, ж/д, ФБС/ЛС/ЛМ, алиасы `C14-40T4` для shipment lookup, запись ручного N в каталог.

## Status

**IMPLEMENT** (2026-09-01): [`../develop/plans/2026-09-01-pile-bridge-trip-calculation.md`](../develop/plans/2026-09-01-pile-bridge-trip-calculation.md). PT-001…PT-902.
