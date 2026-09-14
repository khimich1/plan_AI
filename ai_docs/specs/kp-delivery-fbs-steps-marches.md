# Spec: КП — доставка для ФБС / ступеней / маршей по известной массе

> **Источник идеи:** [`ai_docs/ideas/kp-delivery-fbs-steps-marches.md`](../ideas/kp-delivery-fbs-steps-marches.md)  
> **План и задачи:** [`ai_docs/develop/plans/2026-09-10-kp-delivery-fbs-steps-marches.md`](../develop/plans/2026-09-10-kp-delivery-fbs-steps-marches.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS ✅ → IMPLEMENT  
> **Статус:** ready for planning (OQ-1 решён 2026-09-10: gate по флагу metadata)  
> **Дата:** 2026-09-10  
> **Связанные модули:** `core/commercial_pricing.calculate_total_cost`, `core/cargo_delivery_pricing`, `core/kp_plate_weight`, `core/kp_order_data`, новый `core/product_weight_catalog`, `app/services/commercial_calculation_service`, `app/services/archive_service`, `frontend/.../CalculationResultStep.tsx`

---

## Assumptions I'm Making

1. **Scope — только типы визарда:** `fbs`, `steps`, `marches`. Перемычки, прогоны, тротуарные плиты, ЛП, кросс-блоки **не** добавляем: в pb.db нет их прайсов, в визарде нет product types. Их веса из файлов не импортируем (строки ЛП/КБ при импорте пропускаем с отчётом).
2. **Формула подтверждена пользователем 2026-09-10:** рейсы = `ceil(Σкг / 20 000)`; один общий котёл для трёх типов; тариф — общее поле `logistics_cost` («Стоимость рейса»). Третья константа грузоподъёмности рядом с `CARGO_DELIVERY_TRUCK_CAPACITY_KG=18 600` (плиты) и `PILE_REMAINDER_TRUCK_CAPACITY_KG=19 800` (остаток свай).
3. **Плиты и сваи не трогаем.** Совмещение плит+ФБС в одну фуру не учитываем (редко; ручная правка менеджером — подтверждено 2026-09-10).
4. **Веса из 1С — нетто изделия; объём не ограничивает** (риск по габаритным ЛС — валидация на 1–2 реальных КП до релиза).
5. **Марка без веса → поведение как у свай:** котёл `ready=False`, его доставка 0, марки попадают в `pending`-список в totals и UI. Итог не блокируется.
6. **Конвенции НДС/скидки прежние:** `vat_amount = изделия × VAT_RATE`, скидка на доставку не распространяется. Этот spec не чинит НДС обычного КП (вопрос бухгалтерии №1).
7. **Строка в документах по конвенции свай** (`kp_delivery_export_lines`): если котёл один — «Услуга по доставке грузов»; если котлов несколько — «Доставка ФБС/ЛС/ЛМ».
8. **ЛС-резолв:** точное совпадение нормализованной марки → fallback на семейство `ЛС-N` (веса внутри семейства однородны: все ЛС-14* = 150 кг) → иначе pending. Fallback логируем.
9. **Справочник в plita.db** (как `pile_catalog`, не в pb.db — прайсы отдельно, каталоги отдельно). Таблица `product_weight_catalog`; импорт скриптом по паттерну `scripts/import_fbs_prices_from_xlsx.py`.
10. **«Вшитый» XLSX** (`kp-xlsx-delivery-in-unit-price`, в реализации): после мержа обеих фич пул расширяется третьим котлом (`fbs`+`steps`+`marches`); до этого вшивает только плиты/сваи, как написано в его spec. Порядок мержа определяет, кто правит `PLATE_POOL`/`PILE_POOL`.
11. **Ретро-эффект на архив — закрыт (D10):** котёл включается только при флаге `fbs_lm_delivery_enabled` в metadata КП. Флаг выставляется при создании драфта после релиза и сохраняется в kp_meta вместе с КП. `calculate_total_cost` получает новый kwarg с дефолтом «выкл» → архивные КП без флага побайтово не меняются, включая `update_logistics_cost` в архиве.

---

## Decisions locked (ideation 2026-09-10)

| # | Тема | Решение |
|---|------|---------|
| D1 | Scope | Только ФБС / ЛС / ЛМ (типы уже в визарде) |
| D2 | Схема | Весовая: `ceil(Σкг{fbs,steps,marches} / 20 000) × logistics_cost` |
| D3 | Тариф | Общий, одно поле «Стоимость рейса» |
| D4 | Сваи | Расчёт не трогаем; расхождение весов 11 марок — отдельная data-задача |
| D5 | Смешанный груз | Не учитываем; редкий случай — ручная правка |
| D6 | Успех | Менеджер: итог сразу; бухгалтерия: документы сходятся |
| D7 | Источник весов | Выгрузки 1С (`банк знаний/Новая папка`), импорт с валидацией |
| D8 | Марка без веса | Pending-конвенция свай (ready=False, котёл 0, список марок) |
| D9 | Котёл | Один общий на три типа (не три отдельных) |
| D10 | Архив | Котёл только для новых КП: флаг `fbs_lm_delivery_enabled` в metadata; архивные КП без флага не меняются (OQ-1, 2026-09-10) |

---

## Objective

Менеджер собирает КП с ФБС, ступенями или маршами (mono или mixed), вводит стоимость рейса один раз — и доставка этих изделий считается автоматически по массе из справочника. Итог КП не требует ручных досчётов; PDF/XLSX показывают доставку по существующей конвенции строк; карточка архива и «в производство» не ломаются.

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | собрал КП только из ФБС | ввести цену рейса и увидеть доставку в итоге | не считать фуры в уме |
| US-2 | собрал mixed плиты+ФБС | два независимых котла (18,6т и 20т) | сваи/плиты не оплачивали чужую доставку |
| US-3 | в КП марка без веса в справочнике | видеть список таких марок, доставка котла 0 | знать, что допросить данные, а не молча отдать заниженный итог |
| US-4 | открыл архивное КП | предсказуемый итог (см. OQ-1) | не объяснять заказчику «почему сумма изменилась» |
| US-5 | бухгалтерия сверяет PDF | строку доставки по знакомой конвенции | не переучивать заказчиков |

### Reframed success criteria

| Требование | Критерий |
|------------|----------|
| Формула | Mono-ФБС 30 т, рейс 50 000 → доставка = `ceil(30000/20000)×50000` = 100 000 |
| Независимость котлов | Mixed плиты+ФБС: плитный котёл (18,6т) и ФБС-котёл (20т) считаются раздельно, суммы складываются |
| Нулевой котёл | Нет строк fbs/steps/marches или их масса 0 → котёл 0, строки доставки нет |
| Pending | Марка без веса → `fbs_lm_delivery_ready=False`, котёл 0, марки в `fbs_lm_pending_marks` |
| UI | Поле «Стоимость рейса» видно для mono ФБС/ЛС/ЛМ драфтов; в итогах — рейсы котла |
| Документы | Строка по конвенции A7; Σ строк доставки = `plate + pile + fbs_lm` котлы |
| Регрессия | Существующие pytest по плитам/сваям зелёные без правок; архивные плиты/сваи КП не меняются |
| Импорт | Валидатор отклоняет: плотность вне 1500–3500 кг/м³, код 1С в колонке объёма, дубли марки с разным весом; строки ЛП/КБ пропускаются с отчётом |

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2 |
| Domain | `core/commercial_pricing.py`, новый `core/product_weight_catalog.py`, `core/cargo_delivery_pricing.py` |
| Data | SQLite: справочник в **plita.db** (`product_weight_catalog`); прайсы не трогаем |
| Frontend | React 19, TS; `CalculationResultStep.tsx` (видимость поля рейса, плашки итогов) |
| Tests | pytest (`tests/test_product_weight_catalog.py`, `tests/test_commercial_fbs_lm_delivery.py`), Vitest на UI |

---

## Commands

```bash
source venv/bin/activate
pytest tests/test_product_weight_catalog.py tests/test_commercial_fbs_lm_delivery.py -q
# регрессия обязательна:
pytest tests/test_commercial_logistics_cost.py tests/test_pile_trip_pricing.py tests/test_commercial_export_mixed.py tests/test_archive_endpoints.py tests/test_embed_delivery_in_unit_price.py -q

python scripts/import_product_weights_from_xlsx.py "банк знаний/Новая папка/Блоки.xlsx" --type fbs --dry-run

cd frontend && npm run typecheck && npm run test -- --run src/features/commercial-offer && npm run build
./run+logs.sh
```

---

## Project Structure

```
ai_docs/ideas/kp-delivery-fbs-steps-marches.md           # идея + data findings
ai_docs/specs/kp-delivery-fbs-steps-marches.md           # этот spec
core/product_weight_catalog.py                           # NEW: schema, parse xlsx, upsert, resolve (точный→семейство ЛС→None)
core/commercial_pricing.py                               # третий котёл в calculate_total_cost + ключи totals
core/cargo_delivery_pricing.py                           # константа 20 000 + weight-агрегация по типам fbs/steps/marches
scripts/import_product_weights_from_xlsx.py              # NEW: CLI-импорт с валидацией
app/services/commercial_calculation_service.py           # прокидка catalog_db_path + флага (как pile_catalog_db_path)
app/services/commercial_draft_service.py                 # выставление fbs_lm_delivery_enabled при создании драфта
app/repositories/kp_repository.py                        # сохранение/чтение флага в kp_meta
app/services/archive_service.py                          # totals для карточки + документов: флаг из raw metadata
frontend/src/features/commercial-offer/components/steps/CalculationResultStep.tsx  # видимость поля рейса, рейсы котла
frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx          # pending-марки котла (как у свай)
tests/test_product_weight_catalog.py                     # NEW
tests/test_commercial_fbs_lm_delivery.py                 # NEW
```

`core/embed_delivery_in_unit_price.py` (соседняя фича): после мержа обеих — третий пул `{"fbs","steps","marches"}` с надбавкой из `fbs_lm_delivery_total`. Отдельным follow-up, чтобы не блокировать друг друга.

---

## Code Style

Чистые функции, деньги в рублях с round(..., 2), марки нормализуем как в `pile_catalog`/`fbs_price_db`:

```python
FBS_LM_TRUCK_CAPACITY_KG: float = 20_000.0
FBS_LM_PRODUCT_TYPES = frozenset({"fbs", "steps", "marches"})
FBS_LM_PRODUCT_KINDS = frozenset({"fbs", "step", "march"})

@dataclass(frozen=True)
class WeightedDeliveryBreakdown:
    cargo_kg: float
    trips: int            # 0 если pending_marks не пуст
    pending_marks: tuple[str, ...]
    @property
    def ready(self) -> bool: ...

def compute_fbs_lm_delivery(
    order_data: list[dict[str, Any]],
    *,
    trip_cost: float,
    catalog_db_path: str,
) -> WeightedDeliveryBreakdown:
    """Σ масса строк fbs/steps/marches → ceil(кг/20 000) × рейс.

    Тип строки: product_type ∈ FBS_LM_PRODUCT_TYPES (mixed-builder)
    или product_kind ∈ FBS_LM_PRODUCT_KINDS (mono/legacy).
    Марка без веса → pending_marks, котёл не готов (как у свай).
    """
```

Контракт totals (новые ключи, старые не меняются):

```python
{
    # ... существующие ключи без изменений ...
    "fbs_lm_delivery_total": 100_000.0,
    "fbs_lm_cargo_kg": 30_000.0,
    "fbs_lm_trips": 2,
    "fbs_lm_delivery_ready": True,
    "fbs_lm_pending_marks": [],
}
```

`kp_delivery_export_lines` — третья строка по конвенции: котёл один в документе → «Услуга по доставке грузов»; несколько котлов → «Доставка ФБС/ЛС/ЛМ» (`trips`, `unit_price=logistics_cost`, `amount`).

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Pytest unit: catalog | parse xlsx (колонки, пропуск пустых/ЛП/КБ), валидация плотности 1500–3500, код-в-объёме, дубли с разным весом; resolve: точный, семейный ЛС (`ЛС14-Б` → 150), Т↔T/пробелы, miss → None |
| Pytest unit: boiler | ceil на границе 20 000; qty=0 строки не участвуют; pending → trips=0/ready=False; три типа в одном котле; смешение с product_kind |
| Pytest integration | `calculate_total_cost` mixed плиты+ФБС: два котла независимо; `delivery_total = plate + pile + fbs_lm`; старые ключи totals не изменились |
| Pytest export | `kp_delivery_export_lines`: третья строка, лейблы по конвенции; PDF/XLSX smoke без регрессии |
| Pytest archive | карточка mono-ФБС с рейсом и флагом: totals с котлом; архивное mixed-КП **без** флага: totals побайтово как раньше (D10) |
| Vitest | поле рейса видно для fbs/steps/marches драфтов; плашка «Рейсов ФБС/ЛС/ЛМ»; pending-марки отображаются |

Регрессия обязательна: `test_commercial_logistics_cost.py`, `test_pile_trip_pricing.py`, `test_commercial_export_mixed.py`, `test_archive_endpoints.py`, `test_embed_delivery_in_unit_price.py`.

---

## Boundaries

**Always:**

- Считать все котлы внутри `calculate_total_cost` — не плодить параллельные расчёты
- Справочник только читается из расчёта; запись — только импорт-скриптом
- Валидация при импорте, не в рантайме КП
- Тесты на математику котла до правки UI

**Ask first:**

- Новая таблица `product_weight_catalog` в plita.db (схема)
- Отдельное поле тарифа вместо общего (если логистика ответит «свой транспорт»)
- Изменение `kp_delivery_export_lines` лейблов существующих строк

**Never:**

- Трогать `pile_catalog`, `pile_trip_pricing`, плитный котёл 18 600
- UPDATE прайсов/строк КП из импортёра весов
- Веса из файлов без валидации плотности
- Commit живых БД / секретов
- Правка НДС обычного КП в этом spec

---

## Success Criteria

1. Mono-КП из ФБС/ЛС/ЛМ с ценой рейса: доставка в итоге = `ceil(Σкг/20 000) × рейс`, без ручного ввода рейсов.
2. Mixed плиты+ФБС: котлы независимы; существующие сценарии плит/свай побайтово не изменились.
3. Марка без веса → pending-список в карточке/визарде, котёл 0, итог честный.
4. Импорт трёх файлов проходит с отчётом; 6 известных артефактов (идея, таблица B) отклонены.
5. pytest/vitest/typecheck зелёные, включая регрессию из Commands.
6. D10 покрыт тестом: архивное КП без флага не меняется, новое с флагом — считает котёл.

---

## Out of scope (Not Doing)

- Сваи: расчёт, `pile_catalog`, сверка 11 расхождений весов (отдельная data-задача)
- Новые типы продукции в визард (перемычки, прогоны, тротуарные, ЛП, кросс-блоки)
- Совмещение разных типов в одну фуру
- Отдельные тарифы по типам
- Третий пул во «вшитом» XLSX — follow-up после мержа соседней фичи
- Правка НДС обычного КП; вопросы бухгалтерии №1–7 (см. идею)
- Telegram-бот

---

## Open Questions

Блокирующих нет. Остались внешние проверки перед релизом (код не блокируют):

1. Подтверждение 20 000 кг и «веса нетто без палет» — логистика/производство.
2. Объёмное ограничение фуры для габаритных ЛС — проверка на 1–2 реальных КП.
3. Вопросы бухгалтерии №1–7 (см. файл идеи).
