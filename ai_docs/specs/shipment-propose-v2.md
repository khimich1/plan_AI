# Spec: Propose v2 — автонабор рейса по правилам укладки ПБ

> **Источник идеи:** [`ai_docs/ideas/shipment-propose-v2.md`](../ideas/shipment-propose-v2.md) (ideation 2026-08-02)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS ⏳ → IMPLEMENT ⏳  
> **Связанные:** [`shipment-logistics.md`](./shipment-logistics.md) (MVP-1, SHIP-201), `app/services/shipment_service.py`, `core/kp_plate_weight.py`, `scripts/shipment_propose_hitrate.py`  
> **Дата:** 2026-08-02  
> **Статус:** ✅ SPECIFY + PLAN approved (2026-08-02)  
> **План:** [`ai_docs/develop/plans/2026-08-02-shipment-propose-v2.md`](../develop/plans/2026-08-02-shipment-propose-v2.md)  
> **Scope:** замена алгоритма `propose` (SHIP-201) на движок правил укладки; **без** раскладки серии рейсов и **без** live-валидации ручных правок

---

## ASSUMPTIONS I'M MAKING

1. **Кандидат-пул = плиты СГП только из КП, привязанных к рейсу** (как сейчас). Расширение на непривязанные КП того же заказчика — **Out of scope MVP**, см. Open Questions.
2. **Длина для правил «маркировка»** — из **`plate_name`** (парсер `core/plate_line_parser.py`, номинал ПБ 64 → 6,4 м). `plate_name` можно доверять. Fallback при неразборе: `round(length_m, 1)`.
3. **Кусок** — `width_m < 1.2` (строго меньше 1,2 м).
4. **Вес** — расчётный через `resolve_kp_line_weight_kg` (как сейчас), не фактическое взвешивание.
5. **Класс ТС:** `t20` и **отсутствие класса** → движок v2 с лимитами t20 (19 800 кг, 13,2 м, 4 яруса). **`t30plus`** → **legacy propose** (весовой FIFO, как SHIP-201) до уточнения правил укладки.
6. **UX** — та же кнопка «Предложить состав»; изменения UI минимальны: тексты Alert + показ `reason` у `not_fit`, блок предупреждений и «остаток по заказу».
7. **5-й ярус** — в MVP **не реализуем** ручной допуск с подтверждением; движок жёстко режет на 4 яруса. Подтверждение 5-го яруса — Open Question / post-MVP.
8. **Свободные строки (сваи)** — `propose` v2 работает **только с plate-строками СГП**; сваи логист добавляет вручную после propose (как сейчас).

→ **Correct me now or I'll proceed with these.**

---

## Objective

Заменить текущий «весовой FIFO» в `ShipmentService.propose` (SHIP-201) на **упаковщик по полному набору правил укладки ПБ в кузов 13,2 м**, чтобы логист получал рейс, который можно грузить без ручной перекладки, и видел **причину** для каждой невлезшей плиты.

**Пользователь:** логист (роль `logistics`).

**Проблема сейчас:** `propose` набирает плиты FIFO по `completed_date` и режет только по суммарному весу класса ТС. Не учитывает геометрию кузова, стопки, ярусы, совместимость длин, добор кусков. Логист правит ~X% рейсов вручную (метрика: `scripts/shipment_propose_hitrate.py`).

**Успех:** логист нажимает «Предложить состав» и получает:
- `items` — состав, проходящий **жёсткие** ограничения (вес, высота стопки);
- `not_fit[]` — каждая позиция с **кодом и текстом причины**;
- `order_remainder[]` — сколько штук каждого размера осталось на следующий рейс по заказу;
- `warnings[]` — мягкие ограничения (микс КП, неоптимальные куски); длина кузова — **hard** (D11), override вручную.

Движок — **отдельный pure-модуль** (`core/shipment_packing/`), чтобы следующим этапом та же математика работала в раскладке КП на серию рейсов.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | логист | нажать «Предложить состав» и получить рейс ≤19,8 т с укладкой в 13,2 м | не перекладывать плиты на площадке |
| US-2 | логист | видеть, почему конкретная плита не влезла | понимать, что оставить на следующий рейс |
| US-3 | логист | видеть остаток по заказу после propose | планировать следующий рейс |
| US-4 | логист | получить предупреждение при миксе нескольких КП в одном рейсе | проверить документооборот (УПД/счета) |
| US-5 | логист | видеть предупреждения о миксе КП и неоптимальных кусках | принять осознанно или поправить вручную |
| US-6 | разработчик | иметь golden-тесты из реестра отгрузок | не регрессировать при доработках движка |

### Acceptance criteria (MVP)

- [ ] Новый модуль `core/shipment_packing/` — **pure Python**, без SQLite/FastAPI; вход: список кандидатов + лимиты ТС; выход: items, not_fit, remainder, warnings, layout metadata (опционально для печати).
- [ ] `ShipmentService.propose` делегирует в движок; поведение без `vehicle_class` сохраняется (все доступные плиты, без геометрии — или явно документированное отличие, см. Open Questions).
- [ ] **Жёсткие ограничения (hard):**
  - [ ] суммарный вес ≤ лимита класса ТС (default `t20`: **19 800 кг**);
  - [ ] высота стопки ≤ **4 ярусов**;
  - [ ] правила совместимости длин в стопке: разница маркировок ≤ **1,0 м**;
  - [ ] ёмкость по длине (ГОСТ-таблица): ≤3,3 м → 4 штабеля; 3,3–4,4 м → 3; 4,4–6,5 м → 2; >6,5 м → 1;
  - [ ] ярус — до **2 значений ширины** (узкие между широкими);
  - [ ] штабель = 2 в ширину × 4 в высоту; узкие плиты — до 4 в ширину или в ярус рядом.
- [ ] **Длина кузова 13,2 м — hard в движке (D11):** не помещается → `not_fit` с `body_length`; логист может добрать через «Добавить всё равно» (confirm не блокируется).
- [ ] **Мягкие ограничения (soft → warnings, не блокируют items):**
  - [ ] неоптимальное распределение кусков;
  - [ ] микс КП одного заказчика в одном рейсе — предупреждение с номерами КП.
- [ ] **Добор кусков** (`width_m < 1.2`): приоритеты ①–④ из ideation; «кусковая машина» (только куски) — **запрещена**.
- [ ] **FIFO по дате изготовления** внутри одинакового размера; между размерами — решает геометрия.
- [ ] API-ответ расширен: `not_fit[].reason_code`, `not_fit[].reason_text`; `warnings[]`; `order_remainder[]`.
- [ ] Лимиты класса ТС в конфиге — **структура** `{max_weight_kg, body_length_m, max_tiers}`, не только вес.
- [ ] Golden-тесты: ПБ 58→10 шт, 63,5→9, 73→8, 90→6; кластер 64+74 в одной стопке; добор куском той же длины.
- [ ] Существующие тесты SHIP-201 адаптированы; `propose_snapshot` сохраняет новый формат ответа.
- [ ] Frontend: блок `not_fit` показывает причину; новый блок warnings и остаток по заказу (только тексты, без смены flow propose→confirm).

### Out of scope (Not Doing)

- Раскладка КП на **серию рейсов** — следующий этап из того же движка.
- Live-валидация **ручных правок** состава — остаётся текущий контроль перегруза по весу.
- Оптимизатор парка (выбор t20/t30plus по тарифу).
- ГОСТ-паспорт / схема загрузки для водителя.
- Панелевоз (вертикальная перевозка).
- Подтягивание доборов из **непривязанных** КП того же заказчика.
- Подтверждение 5-го яруса в UI.

---

## Tech Stack

| Слой | Стек |
|------|------|
| Packing engine | **NEW** `core/shipment_packing/` — pure stdlib + dataclasses; без ORM |
| Service | `app/services/shipment_service.py` — `propose()` вызывает движок |
| Config | `core/config/settings.py` — `vehicle_class_limits` (расширение JSON) |
| Weight | `core/kp_plate_weight.resolve_kp_line_weight_kg` |
| Availability | `core/kp_db_shipments.available_qty` |
| API schema | `app/schemas/logistics.py` — расширение `ShipmentProposeResponse` |
| API endpoint | `POST /api/v1/logistics/shipments/{id}/propose` — контракт backward-compatible (+ новые поля) |
| Frontend | `frontend/src/features/logistics/` — типы + тексты Alert в `ShipmentItemsSection` |
| Tests | pytest `tests/test_shipment_packing.py` (NEW), `tests/test_shipment_service.py` (адаптация) |
| Metrics | `scripts/shipment_propose_hitrate.py` — без изменений (читает snapshot JSON) |

---

## Commands

```bash
# Backend dev
source .venv/bin/activate
uvicorn app.main:app --reload

# Unit tests — движок (NEW)
pytest tests/test_shipment_packing.py -q

# Regression — propose + логистика
pytest tests/test_shipment_service.py -q -k propose
pytest tests/test_logistics_api.py -q -k propose

# Full backend gate
pytest tests/ -k "shipment or logistics" -q

# Hit-rate метрика (после пилота на prod-данных)
./.venv/bin/python scripts/shipment_propose_hitrate.py --verbose

# Frontend types/build (при изменении схемы)
cd frontend && npm run build
cd frontend && npm test -- --run src/features/logistics
```

---

## Project Structure

```
core/shipment_packing/
  __init__.py           → публичный API: pack_shipment(...)
  models.py             → PlateUnit, Stack, Tier, VehicleLimits, PackResult
  marking.py            → marking_length_m(plate_name, length_m) → parse plate_name
  rules.py              → GOST stack count, tier width, length mix, piece detection
  engine.py             → основной алгоритм набора + добор кусков
  reasons.py            → коды причин not_fit / warnings (enum + human text)

app/services/shipment_service.py
  propose()             → fetch candidates → pack_shipment → map to ShipmentProposeResponse

app/schemas/logistics.py
  ShipmentProposeItem   → + reason_code, reason_text (optional, only not_fit)
  ShipmentProposeResponse → + warnings[], order_remainder[]

core/config/settings.py
  vehicle_class_limits_raw → JSON {"t20": {"max_weight_kg": 19800, "body_length_m": 13.2, "max_tiers": 4}, ...}

tests/test_shipment_packing.py   → golden cases, rule unit tests
tests/test_shipment_service.py   → integration propose

frontend/src/features/logistics/
  types/logistics.ts             → ProposedItem.reason_*, ProposeResponse.warnings, order_remainder
  components/ShipmentItemsSection.tsx → отображение reason / warnings / remainder
```

---

## Code Style

Следуем существующим конвенциям проекта: dataclasses в pure-модулях, Pydantic v2 в API, `ShipmentError` с `.code` в сервисе.

**Пример: код причины и результат движка**

```python
# core/shipment_packing/reasons.py
class NotFitReason(str, Enum):
    WEIGHT_LIMIT = "weight_limit"
    TIER_LIMIT = "tier_limit"
    LENGTH_MIX = "length_mix_in_stack"
    BODY_LENGTH = "body_length"
    NO_STACK_SLOT = "no_stack_slot"
    PIECE_PRIORITY = "piece_priority"
    RESERVED = "reserved"  # qty=0 на СГП

REASON_TEXT = {
    NotFitReason.WEIGHT_LIMIT: "Превышен лимит веса класса ТС",
    NotFitReason.TIER_LIMIT: "Не помещается в 4 яруса",
    # ...
}

# core/shipment_packing/__init__.py
def pack_shipment(
    candidates: Sequence[PlateCandidate],
    *,
    limits: VehicleLimits | None,
) -> PackResult:
    """Pure: без I/O. Все qty разложены: items + not_fit + remainder."""
    ...
```

**Пример: расширение API-ответа**

```python
class ShipmentProposeWarning(BaseModel):
    code: str
    message: str
    kp_ids: list[int] = Field(default_factory=list)

class ShipmentOrderRemainderItem(BaseModel):
    completed_plate_id: int
    kp_id: int
    plate_name: str
    qty_remaining: int

class ShipmentProposeResponse(BaseModel):
    items: list[ShipmentProposeItem] = Field(default_factory=list)
    not_fit: list[ShipmentProposeItem] = Field(default_factory=list)
    order_remainder: list[ShipmentOrderRemainderItem] = Field(default_factory=list)
    warnings: list[ShipmentProposeWarning] = Field(default_factory=list)
    total_weight_kg: float = 0.0
    overload: bool = False
    vehicle_class: str | None = None
    vehicle_class_limits: dict[str, VehicleClassLimits] = Field(default_factory=dict)
```

Naming: `reason_code` / `reason_text` — как в `core/plate_validation.py` и OCR diagnostics.

---

## Domain Rules (спека движка)

### Входные данные кандидата

| Поле | Источник | Использование |
|------|----------|---------------|
| `completed_plate_id`, `kp_id` | `completed_plates` | идентификация |
| `plate_name`, `length_m`, `width_m`, `load_class` | `completed_plates` | геометрия, вес |
| `qty` | `available_qty()` | доступное кол-во |
| `completed_date` | `completed_plates` | FIFO внутри размера |
| `marking_length_m` | из `plate_name` (`plate_line_parser`); fallback `round(length_m, 1)` | правила стопок |

### Лимиты класса ТС (config)

| Класс | max_weight_kg | body_length_m | max_tiers |
|-------|---------------|---------------|-----------|
| `t20` | **19 800** | 13.2 | 4 |
| `t30plus` | 30 000 (или 19 800 — уточнить) | 13.2 | 4 |

Default JSON в settings заменяет текущий `vehicle_class_limits_kg` **с миграцией**: property `vehicle_class_limits_kg` остаётся для backward compat (только вес).

### Алгоритм (high-level)

```
1. Собрать плоский пул unit-ов из candidates (развернуть qty).
2. Отсортировать: marking_length DESC (крупные первыми), внутри — completed_date ASC (FIFO).
3. Жадно формировать стопки:
   a. Группировать по совместимости marking_length (Δ ≤ 1.0 м).
   b. Проверять ёмкость по ГОСТ-таблице (число штабелей вдоль кузова).
   c. Укладывать ярусы (≤2 ширины, ≤4 высоты, 2×4 ширина×высота штабеля).
4. Добор кусков (width < 1.2) по приоритетам ①–④.
5. Проверить hard: weight, tiers.
6. Сформировать soft warnings.
7. Агрегировать обратно в qty по completed_plate_id.
8. remainder = available - items (по каждому candidate).
```

### Коды причин `not_fit`

| code | Когда |
|------|-------|
| `weight_limit` | unit не помещается без превышения max_weight_kg |
| `tier_limit` | не помещается в max_tiers |
| `length_mix` | нет стопки с Δ marking ≤ 1.0 м |
| `body_length` | не помещается в `body_length_m` (13,2 м) — **hard**, в `not_fit` |
| `no_stack_slot` | нет места в ширине/ярусе |
| `piece_priority` | кусок не прошёл приоритеты добора |
| `next_trip` | сознательно оставлено на следующий рейс после полного набора |

### Коды `warnings`

| code | Когда |
|------|-------|
| `kp_mix` | items содержат плиты из ≥2 kp_id |
| `piece_suboptimal` | кусок положен не по приоритету ①, но влез |
| `marking_fallback` | `plate_name` не разобран — marking из `round(length_m, 1)` |

---

## Testing Strategy

| Уровень | Где | Что покрываем |
|---------|-----|---------------|
| Unit | `tests/test_shipment_packing.py` | marking, GOST table, stack compatibility, tier width, piece priority, golden layouts |
| Integration | `tests/test_shipment_service.py` | propose → DB snapshot, reserved qty, multi-KP |
| API | `tests/test_logistics_api.py` | HTTP contract, новые поля |
| Frontend | `ShipmentDrawer.test.tsx`, `logisticsApi.test.ts` | отображение reason/warnings |
| Manual | 2–3 живых рейса с логистом | assumptions validation |

**Golden cases (обязательные):**

1. ПБ 58 (5,8 м marking) → **10 шт** в t20.
2. ПБ 63,5 → **9 шт**.
3. ПБ 73 → **8 шт**.
4. ПБ 90 → **6 шт**.
5. Кластер **ПБ 64 + ПБ 74** (Δ=1,0 м) — в **одной стопке**.
6. Добор **куском той же длины** в ярус к целым.

Каждый golden — фикстура с явными `length_m`, `width_m`, qty, expected items count и not_fit reasons.

Coverage: движок ≥90% line coverage на `core/shipment_packing/` (новый код).

---

## Boundaries

### Always

- Движок **pure** — тестируется без БД.
- Все qty инвариант: `sum(items.qty) + sum(not_fit.qty) + sum(remainder.qty)` = available по каждому candidate (ничего не теряется).
- Жёсткие ограничения **никогда** не попадают в `items`.
- Сохранять `propose_snapshot` после каждого propose.
- Запускать `pytest tests/test_shipment_packing.py tests/test_shipment_service.py -k propose` перед merge.

### Ask first

- Изменение default лимита t20 с 20 000 → **19 800** в production `.env`.
- Изменение поведения propose **без** `vehicle_class` (сейчас — все плиты).
- Добавление зависимостей (OR-Tools и т.п.) — MVP только жадный алгоритм.
- Расширение кандидат-пула на непривязанные КП.

### Never

- Блокировать confirm из-за soft warnings.
- Коммитить секреты / `.env`.
- Ломать backward compat API без версионирования (новые поля — optional с defaults).
- Удалять failing golden tests без согласования с логистом.

---

## Success Criteria

1. Golden-тесты 1–6 **PASS** стабильно.
2. `propose` с `vehicle_class=t20` **не возвращает** items с весом > 19 800 кг и > 4 ярусов.
3. Каждый элемент `not_fit` имеет непустые `reason_code` и `reason_text`.
4. `order_remainder` совпадает с `available - proposed` по каждой plate-строке.
5. Frontend показывает причину в блоке «Не влезло» (не только «лимит класса ТС»).
6. `scripts/shipment_propose_hitrate.py` работает на новом snapshot JSON.
7. Логист подтверждает 2–3 кейса на пилоте (assumptions checklist).

---

## Open Questions

| # | Вопрос | Решение (2026-08-02) |
|---|--------|----------------------|
| Q1 | Непривязанные КП того же заказчика? | **MVP: нет** (только привязанные). Расширение на заказчика — **post-MVP**, отдельная задача |
| Q2 | Допуск 5-го яруса | **Out of scope MVP** — hard 4 яруса |
| Q3 | Складские куски вне заказа | **Нет** в MVP |
| Q4 | 19 800 vs 20 000 | **19 800 кг** default для t20 |
| Q5 | Propose без `vehicle_class` | **t20 по умолчанию** — движок v2 с лимитами t20 |
| Q6 | Правила `t30plus` | **Legacy propose** (весовой FIFO) до уточнения правил с логистом |
| Q7 | Маркировка для Δ≤1 м | **`plate_name`** (доверяем); fallback `round(length_m, 1)` |

---

## Decisions locked

| # | Тема | Решение | Статус |
|---|------|---------|--------|
| D1 | Модуль движка | `core/shipment_packing/` pure | ✅ |
| D2 | Алгоритм MVP | Жадный, без ILP | ✅ |
| D3 | Лимит t20 weight | 19 800 кг default | ✅ |
| D4 | marking length | из `plate_name`; fallback `round(length_m, 1)` | ✅ |
| D5 | Кусок | `width_m < 1.2` | ✅ |
| D6 | API расширение | +reason, +warnings, +order_remainder (optional fields) | ✅ |
| D7 | Без vehicle_class | **t20 по умолчанию** (v2 движок) | ✅ |
| D8 | 5-й ярус | Out of scope MVP | ✅ |
| D9 | t30plus | Legacy propose до уточнения правил | ✅ |
| D10 | Кандидат-пул | Только привязанные КП (post-MVP: + заказчик) | ✅ |
| D11 | Длина кузова 13,2 м | **Hard** в движке → `not_fit`; override — «Добавить всё равно», confirm не блокируется | ✅ |
| D12 | t30plus legacy | Лимит веса **30 000 кг** (без изменений) | ✅ |
| D13 | Выкатка v2 | Сразу для t20/null, без feature-flag | ✅ |
| D14 | `order_remainder` | По строкам `completed_plates` | ✅ |

---

## Key Assumptions to Validate (from ideation)

- [ ] Разница длин по маркировке (0,1 м) — подтвердить с логистом на 2–3 живых рейсах
- [ ] Две ширины в ярусе — подтвердить с крановщиком/мастером отгрузки
- [ ] Предупреждения о миксе КП достаточно — проверить документооборот
- [ ] Лимит 19 800 не ломает процессы — согласовать до деплоя

---

## Next step

→ **Phase 2: PLAN** — [`ai_docs/develop/plans/2026-08-02-shipment-propose-v2.md`](../develop/plans/2026-08-02-shipment-propose-v2.md)  
→ **Phase 3: TASKS** — после ревью плана
