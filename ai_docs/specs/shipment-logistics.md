# Spec: Логистика и отгрузка (MVP-1 — реестр рейсов и списание СГП)

> **Источник идеи:** [`ai_docs/ideas/shipment-logistics.md`](../ideas/shipment-logistics.md) (ideation 2026-07-31, 15 decisions locked)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS ✅ → IMPLEMENT ✅  
> **План:** [`ai_docs/develop/plans/2026-07-31-shipment-logistics.md`](../develop/plans/2026-07-31-shipment-logistics.md)  
> **Отчёт:** [`ai_docs/develop/reports/2026-07-31-shipment-logistics-implementation.md`](../develop/reports/2026-07-31-shipment-logistics-implementation.md)  
> **Дата:** 2026-07-31  
> **Статус:** ✅ implemented (MVP-1)  
> **Scope:** **только MVP-1** (утверждён на SPECIFY-уточнении 2026-07-31). MVP-2 (договоры-заявки, полная карточка перевозчика, ПДн) — отдельная спека сразу следом.  
> **Связанные:** [`sgp-warehouse.md`](./sgp-warehouse.md) (СГП MVP, implemented), `SgpService`, `completed_plates`, `kp_plates`, `plate_status_log`, `app/domain/enums.py`, [`docs/specs/1c-integration-tz.md`](../../docs/specs/1c-integration-tz.md), `core/kp_plate_weight.py`

---

## Decisions locked

Перенесены из ideation без изменений (источник — таблица «Decisions locked» идеи), плюс уточнения SPECIFY (S1–S4):

| # | Тема | Решение |
|---|------|---------|
| 1 | Типы выдачи | Доставка + самовывоз, ~50/50 |
| 2 | Частичная отгрузка | Норма; гранулярность = рейс |
| 3 | Мульти-заказ в одном ТС | Да (несколько КП/ЯР в одном рейсе) |
| 4 | Закрытие КП | `KpStatus.DONE` при отгрузке полного заказного qty |
| 5 | Смешанный заказ (плиты+сваи) | Да, гибридные позиции обязательны (plate-строки + free-строки) |
| 6 | Сбор рейса | propose→confirm из СГП (паттерн «Закрыть со склада») |
| 7 | Рейсы без заказа | Не в MVP (рейс требует ≥1 КП при создании) |
| 8 | № заказа | = номер счёта `ЯР-XXXXXXX`; обязателен к отгрузке, не при создании |
| 9 | Статусы рейса | 2 статуса `в работе → обработано` + флаг «внимание» с комментарием |
| 10 | КПП/пропуска | Не фиксируем (0% заполнения за 5 лет) |
| 11 | СГП достоверен | Пока нет; gate пилота = сверка «в системе = на складе» |
| 12 | УПД | Готовится заранее; поле в рейсе, заполняется до выезда |
| 13 | Договор-заявка | MVP-2: генерация, номер `N/MM/YY` месячный |
| 14 | Налоговый режим | MVP-2: из карточки перевозчика |
| 15 | Scope | **MVP-1 (реестр+СГП) → MVP-2 (договоры) отдельной спекой** |
| S1 | Scope этой спеки | **Только MVP-1** (уточнение 2026-07-31) |
| S2 | Вес сваи | **Мини-справочник `pile_catalog` в MVP-1**: импорт листа «Вес и объем» прайса (44 марки, вес кг + объём м³); свободная строка = марка (автокомплит) + qty → автовес, ручная правка допустима |
| S3 | Пилот | **Cut-over на сегменте**: 1 неделя, один тип выдачи (доставка ИЛИ самовывоз), Excel — read-only |
| S4 | Дедуп перевозчиков | **Нормализация имени + авто-дедуп при импорте + UI слияния**. ИНН в источнике нет (это поле карточки предприятия, MVP-2) — дедуп по ИНН на импорте невозможен; импорт «как есть» даст мусорный автокомплит в пилоте |

---

## Objective

Перенести управление отгрузкой из общего Excel-реестра (17 279 строк, 2021–2026) в раздел «Логистика» веб «Шишов», чтобы:

1. Рейс собирался из **реальных позиций СГП** (ссылки на `completed_plates`), а не вбивался текстом.
2. Статус был **явным** (`в работе → обработано` + «внимание») вместо цвета ячеек.
3. Факт выезда **атомарно списывал СГП** (audit `sgp_ship`), обновлял прогресс КП «отгружено X/M» и готовил событие для 1С.
4. Логист планировал рейс **до документов**: ЯР-номер, ТС, водитель дозаполняются позже; обязательность — на закрытии.

**Пользователь:** логист (новая роль), вторично — мастер СГП (видит актуальный склад), менеджер (видит «отгружено X/M» в архиве).

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | логист | создать рейс на дату и выбрать заказ(ы) КП | строка Excel заводилась в системе, а не в общем файле |
| US-2 | логист | получить предложение состава из плит этого КП на СГП (вес, лимит ТС) | не смотреть резерв внутри счёта в 1С вручную |
| US-3 | логист | скорректировать состав: убрать/добавить плиты, добавить свободную строку (сваи) | рейс соответствовал реальности (гибридные заказы) |
| US-4 | логист | дозаполнить ТС, водителя, УПД, № заявки позже | планировать рейс до приезда машины и до документов |
| US-5 | логист | отметить выезд одной кнопкой | склад списался, прогресс КП обновился, событие для 1С готово |
| US-6 | логист | поставить флаг «внимание» с комментарием | особые события («Работа крана!», «БЕЗ ЦЕН!») не терялись |
| US-7 | логист/мастер СГП | распечатать «лист отгрузки» (позиции, порядок, заметки укладки) | склад грузил по бумаге без доступа к системе |
| US-8 | логист | выбирать перевозчика из чистого справочника | не плодить дубли 2 397 строк Excel |
| US-9 | менеджер | видеть в архиве «отгружено X/M» рядом с «N/M на СГП» | понимать прогресс заказа без звонка логисту |
| US-10 | владелец | чтобы КП закрывался в «выполнено» только при полной отгрузке | «выполнено» = факт, а не обещание |

### Acceptance criteria (MVP-1)

- [ ] Раздел «Логистика» (отдельный route + роль `logistics`): реестр рейсов с фильтрами **дата (диапазон), заказ (ЯР/КП), перевозчик, тип Д/С, «без УПД», «внимание»**
- [ ] Карточка рейса: ядро-поля (дата, тип Д/С, заказы КП) + важные (перевозчик/доверенность №, водитель, ТС, УПД, № заявки на фрахт, стоимость план); статусы `в работе → обработано` + флаг «внимание»
- [ ] Создание рейса требует ≥1 КП; ЯР-номер(а) — поле связи рейс↔КП, обязательно на закрытии (422), не при создании
- [ ] **Propose→confirm:** система предлагает состав из свободных плит СГП выбранных КП (FIFO по `completed_date`), с автовесом и учётом лимита класса ТС; логист утверждает или правит (убрать/qty/добавить plate-строку/free-строку)
- [ ] Свободные строки: марка сваи — автокомплит из `pile_catalog` (44 марки), автовес = вес×qty, ручная правка веса и произвольная марка допустимы
- [ ] Мульти-заказный рейс: несколько КП, у каждого свой ЯР; прогресс пишется по каждому
- [ ] Автовес рейса (плиты — через `core/kp_plate_weight.py`, сваи — `pile_catalog`) + **предупреждение перегруза** класса ТС (20т/30т+, лимиты в конфиге); не блокирует
- [ ] **Выезд → «обработано»:** атомарное списание СГП (split при частичном qty), audit `sgp_ship` в `plate_status_log`, прогресс «отгружено X/M» в архиве, `KpStatus.DONE` при полном qty
- [ ] Отмена рейса возможна только из `в работе`; состав освобождается (availability восстанавливается)
- [ ] Справочник перевозчиков: импорт из листов «Перевозчики»/«Транспортные Компании» с авто-дедупом (нормализация) + UI слияния дублей
- [ ] Печатная форма «Лист отгрузки» (XLSX): позиции в порядке укладки, вес, заметки; шапка рейса
- [ ] Событие `shipment_completed` (JSON по контракту папки обмена) — заложено в модель и код, **за feature-flag, выключено** до интеграции G
- [ ] Qty-инвариант plate_loss расширен членом «отгружено» и остаётся PASS

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | FastAPI, SQLite (`plita.db`), Pydantic v2 — как в СГП MVP |
| Domain | `app/domain/enums.py` (+`ShipmentStatus`, `DeliveryType`, `ShipmentItemType`, `PlateStatus.SHIPPED`, `PlateTransitionReason.SGP_SHIP`) |
| Service | `app/services/shipment_service.py` (NEW, по образцу `SgpService`: `db_path`, `ShipmentError(ValueError).code` → 422) |
| Вес плит | `core/kp_plate_weight.resolve_kp_line_weight_kg` (режим `WEIGHT_SOURCE`, default formula) |
| Schema | `core/kp_db_schema.py` — `CREATE TABLE IF NOT EXISTS` + `ALTER` try/except, как принято |
| API | `app/api/v1/endpoints/logistics.py` (NEW) |
| Frontend | React, Vite, TypeScript, TanStack Query; `frontend/src/features/logistics/` (NEW) |
| Auth | существующая модель (`app/schemas/auth.py`, `app/dependencies/auth.py`) + роль `logistics` |
| Tests | pytest (`tests/`), Vitest (`frontend/`) |
| Regression | `scripts/run_plate_loss_regression.py` |
| Импорт | `scripts/import_carriers.py`, `scripts/import_pile_catalog.py` (openpyxl уже в venv) |

---

## Commands

```bash
# Backend
source .venv/bin/activate
uvicorn app.main:app --reload
pytest tests/test_shipment_service.py -q
pytest tests/test_shipment_qty_balance.py -q
pytest tests/ -k "sgp or shipment or completion" -q

# Qty gate (расширенный баланс: + отгружено)
./.venv/bin/python scripts/run_plate_loss_regression.py

# Импорт справочников (одноразовые, аргумент — путь к файлу)
./.venv/bin/python scripts/import_carriers.py --xlsx "Реестр отгрузок от 12082021.xlsx"
./.venv/bin/python scripts/import_pile_catalog.py --xlsx "Прайс на цельные сваи от 27.07.2026.xlsx"

# Frontend
cd frontend && npm run dev
cd frontend && npm test -- --run
cd frontend && npm run build
```

---

## Project Structure

```
app/domain/enums.py                    → += ShipmentStatus, DeliveryType, ShipmentItemType,
                                         PlateStatus.SHIPPED, PlateTransitionReason.SGP_SHIP
app/services/shipment_service.py       → NEW: create/propose/confirm/complete/cancel,
                                         availability, DONE-check, event payload
app/services/carrier_service.py        → NEW: list/merge/normalize (или секция в shipment_service)
app/repositories/...                   → при необходимости тонкий слой SQL
app/api/v1/endpoints/logistics.py      → NEW: /api/v1/logistics/* роутер
app/schemas/logistics.py               → NEW: request/response models
app/schemas/auth.py                    → role Literal += "logistics"
app/dependencies/auth.py               → guard для logistics endpoints (роль logistics|admin)
app/core/settings.py                   → VEHICLE_CLASS_LIMITS_KG, SHIPMENT_EVENTS_ENABLED,
                                         EXCHANGE_EXPORT_DIR
core/kp_db_schema.py                   → += shipments, shipment_orders, shipment_items,
                                         carriers, pile_catalog; plate_status_log += shipment_id
core/kp_db_audit.py                    → audit_append: прокинуть shipment_id (опц. параметр)

frontend/src/features/logistics/
  api/logisticsApi.ts                  → NEW
  types/logistics.ts                   → NEW
  components/LogisticsRegistryView.tsx → реестр + фильтры
  components/ShipmentDrawer.tsx        → карточка: поля, propose→confirm, редактор состава
  components/CarriersView.tsx          → справочник + слияние дублей
  router                               → /logistics, /logistics/carriers; nav по роли

frontend/src/features/commercial-archive/
  → бейдж «отгружено X/M» рядом с «N/M на СГП»

scripts/
  import_carriers.py                   → NEW: импорт + нормализация + отчёт о дублях
  import_pile_catalog.py               → NEW: лист «Вес и объем» → pile_catalog

tests/
  test_shipment_service.py             → NEW: propose/confirm/complete/cancel, 422-коды
  test_shipment_qty_balance.py         → NEW: баланс qty с членом «отгружено»
  test_carrier_import.py               → NEW: нормализация, дедуп, merge
  test_pile_catalog_import.py          → NEW: парсинг 44 марок, quirk С137,5.40

ai_docs/ideas/shipment-logistics.md    → идея (источник)
ai_docs/specs/shipment-logistics.md    → эта спека
ai_docs/develop/plans/…                → появится после approval (Phase 2)
```

---

## Code Style

Ориентир — паттерн `SgpService` / `_record_plate_completion`: одна транзакция, audit рядом с UPDATE/INSERT, split при частичном qty, стабильные коды ошибок, сообщения на русском.

```python
# Списание plate-строки при выезде — внутри транзакции complete():
def _ship_plate_item(cur, *, item, cp_row, shipment_id, actor):
    if item.qty > _available_qty(cur, cp_row["id"]):
        raise ShipmentError(
            f"Недостаточно «{cp_row['plate_name']}» на СГП: свободно {_available_qty(cur, cp_row['id'])}, требуется {item.qty}",
            code="shipment_no_availability",
        )
    _deduct_completed_plate_qty(cur, cp_row["id"], item.qty)  # split remainder, как в SGP
    audit_append(
        cur,
        plate_id=cp_row["id"],
        kp_id=item.kp_id,
        plate_name=cp_row["plate_name"],
        shipment_id=shipment_id,                 # новая колонка plate_status_log
        from_status=PlateStatus.ON_SGP.value,    # "on_sgp"
        to_status=PlateStatus.SHIPPED.value,     # "shipped"
        qty=item.qty,
        reason=PlateTransitionReason.SGP_SHIP.value,  # "sgp_ship"
        actor=actor,
    )
```

- Слои: router → service → repository/core SQL; не смешивать SQL в endpoint.
- Enum-значения — единая точка правды в `app/domain/enums.py`; новые коды — английские snake_case (как `on_sgp`), UI-лейблы — русские.
- Частичные qty → **split** (`insert_kp_plate_remainder_row`-стиль), не silent overwrite.
- Коды ошибок стабильные: `shipment_no_availability`, `shipment_missing_ya_order`, `shipment_already_done`, `shipment_not_in_work`, `sgp_row_allocated`, `carrier_merge_conflict`.

---

## Domain model

### Новые таблицы (`plita.db`)

```sql
CREATE TABLE IF NOT EXISTS shipments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    shipment_date TEXT NOT NULL,               -- дата рейса
    delivery_type TEXT NOT NULL,               -- 'delivery' | 'pickup'
    status TEXT NOT NULL DEFAULT 'in_work',    -- 'in_work' | 'done'
    attention INTEGER NOT NULL DEFAULT 0,      -- флаг «внимание»
    attention_comment TEXT,
    carrier_id INTEGER REFERENCES carriers(id),
    driver_name TEXT,                          -- MVP-1: свободный текст (справочник водителей — MVP-2)
    vehicle_text TEXT,                         -- а/м, назначается позже
    vehicle_class TEXT,                        -- 't20' | 't30plus' — для лимита веса
    proxy_no TEXT,                             -- № доверенности (самовывоз)
    upd_no TEXT,                               -- УПД, готовится заранее
    freight_request_no TEXT,                   -- «№ Заявки на фрахт ТС» (MVP-2 → ссылка на договор)
    planned_cost REAL,                         -- стоимость план
    time_slot TEXT,                            -- опционально («утро/вечер», умирает — 9%)
    propose_snapshot TEXT,                     -- JSON снимок предложенного состава (метрика пилота)
    completed_at TEXT,
    actor TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS shipment_orders (   -- рейс ↔ КП (мульти-заказ)
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    shipment_id INTEGER NOT NULL REFERENCES shipments(id) ON DELETE CASCADE,
    kp_id INTEGER NOT NULL REFERENCES KP_offers(kp_id),
    ya_order_no TEXT,                          -- «ЯР-0001467»; обязателен к закрытию
    UNIQUE (shipment_id, kp_id)
);

CREATE TABLE IF NOT EXISTS shipment_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    shipment_id INTEGER NOT NULL REFERENCES shipments(id) ON DELETE CASCADE,
    item_type TEXT NOT NULL,                   -- 'plate' | 'free'
    completed_plate_id INTEGER REFERENCES completed_plates(id),  -- plate-строка: ссылка, не текст
    kp_id INTEGER,                             -- plate: снимок из completed_plates; free: к заказу (опц.)
    mark TEXT,                                 -- free-строка: марка («С60.30») или произвольный текст
    qty INTEGER NOT NULL,
    unit_weight_kg REAL,
    weight_kg REAL,                            -- вес строки; ручная правка допустима
    sort_order INTEGER NOT NULL DEFAULT 0,     -- порядок укладки (печатная форма)
    note TEXT                                  -- заметки укладки
);

CREATE TABLE IF NOT EXISTS carriers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,                        -- как в источнике (для отображения)
    name_normalized TEXT NOT NULL,             -- ключ дедупликации
    source_sheet TEXT,                         -- «Перевозчики» | «Транспортные Компании»
    note TEXT,
    active INTEGER NOT NULL DEFAULT 1,
    merged_into_id INTEGER REFERENCES carriers(id),
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);
-- частичный уникальный индекс: name_normalized UNIQUE WHERE merged_into_id IS NULL

CREATE TABLE IF NOT EXISTS pile_catalog (      -- из прайса, лист «Вес и объем» (44 марки)
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    mark TEXT NOT NULL UNIQUE,                 -- «С30.30» … «С160.40»
    length_m REAL,                             -- из марки: С60.30 → 6.0
    section_mm INTEGER,                        -- 300 | 350 | 400
    volume_m3 REAL,
    weight_kg REAL NOT NULL,
    pcs_per_20t INTEGER                        -- справочно (из колонки «автомобильный г/п 20тн»)
);

ALTER TABLE plate_status_log ADD COLUMN shipment_id INTEGER;  -- миграция try/except
```

### Статусы и enum

| Enum | Значение | Где |
|------|----------|-----|
| `ShipmentStatus.IN_WORK` / `DONE` | `in_work` / `done` | `shipments.status` |
| `DeliveryType.DELIVERY` / `PICKUP` | `delivery` / `pickup` | UI «Доставка» / «Самовывоз» |
| `ShipmentItemType.PLATE` / `FREE` | `plate` / `free` | `shipment_items.item_type` |
| `PlateStatus.SHIPPED` | `shipped` | audit `to_status` при выезде |
| `PlateTransitionReason.SGP_SHIP` | `sgp_ship` | `plate_status_log.reason` |
| `KpStatus.DONE` | `выполнено` | `kp_meta.status` — впервые выставляется системой |

### Availability (вычислимый резерв)

Физический резерв не храним — считаем:

```
available(cp_row) = cp_row.qty − Σ shipment_items.qty
    WHERE completed_plate_id = cp_row.id
      AND shipment.status = 'in_work'
```

- propose видит только строки с `available > 0`.
- `unlink`/`relink` в СГП оперируют только свободной частью; попытка тронуть зарезервированное → 422 `sgp_row_allocated`.
- При `complete` availability перепроверяется (гонка маловероятна, но дешёвая страховка).
- Отмена рейса `in_work` → items удаляются, availability восстанавливается автоматически.

### Операции (инварианты)

| Операция | Физика (`completed_plates`) | Учёт |
|----------|------------------------------|------|
| Create рейса | — | ≥1 `shipment_orders`; статус `in_work` |
| Propose | — | читает СГП FIFO, пишет `propose_snapshot` |
| Confirm состава | — | `shipment_items` заменяются целиком (транзакция) |
| **Complete (выезд)** | −qty по каждой plate-строке (split remainder) | audit `sgp_ship`; статус `done`; прогресс КП; DONE-check; событие |
| Cancel (только `in_work`) | — | items удаляются; availability восстанавливается |
| Carrier merge | — | `shipments.carrier_id` переписываются на target; дубль → `merged_into_id` |

**Qty-инвариант (расширение plate_loss):**

```
Σ qty(kp_plates where status in план|производство)
+ Σ qty(completed_plates)                    -- включая зарезервированное open-рейсами
+ Σ qty(shipment_items plate-строк в done-рейсах)   -- «отгружено»
= исходный заказной qty  (± явный брак/списание, если задокументировано)
```

Сваи/free-строки в инвариант **не входят** (не производятся на дорожках, нет учётного qty).

### Критерий `KpStatus.DONE` (decision 4)

```
shipped_qty(kp) = Σ shipment_items.qty plate-строк done-рейсов по этому kp_id
IF shipped_qty >= M_ordered
   AND Σ kp_plates(kp) == 0
   AND Σ completed_plates WHERE kp_id = kp == 0:
    → kp_meta = «выполнено»
ELSE: без изменений («в работе» / «На СГП» + бейджи)
```

Проверка выполняется для каждого КП рейса в той же транзакции `complete`.

### Propose-алгоритм

Вход: рейс с ≥1 `shipment_orders`, опционально `vehicle_class`.

1. Для каждого КП (в порядке добавления к рейсу): строки `completed_plates WHERE kp_id = X AND available > 0`, `ORDER BY completed_date, id` (FIFO).
2. Вес строки — `resolve_kp_line_weight_kg` по `(length_m, width_m, qty)` (формула; учётный вес в 1С, наш — ориентир).
3. Если задан `vehicle_class`: накапливать до лимита `VEHICLE_CLASS_LIMITS_KG` (default `t20=20000`, `t30plus=30000`); не влезшее — секция «не влезло» в ответе (не молча отбрасывать).
4. Ответ: предлагаемые items (plate, qty, вес) + итоговый вес + флаг перегруза. Логист правит и подтверждает (`PUT items`).
5. `propose_snapshot` сохраняется для метрики пилота «доля рейсов без ручной правки» (сравнение snapshot vs финальный состав done-рейса — скриптом отчёта, не в UI).

### Нормализация перевозчиков (импорт + дедуп)

`name_normalized`: `lower()`, `ё→е`, трим и схлопывание пробелов, удаление кавычек `«»""''`, удаление ОПФ-токенов (`ооо`, `ип`, `ао`, `пао`, `зао`, `оао`), удаление пунктуации. Импорт вставляет по одному активному на `name_normalized`, дубликаты — в отчёт скрипта (не в БД). Near-дубли (опечатки, дефисы) — ручной merge в UI: `POST /carriers/{id}/merge { into_id }` переписывает `shipments.carrier_id`, дубль получает `merged_into_id`, `active=0`.

### Событие `shipment_completed` (контракт-черновик под папку обмена 1С)

По [`docs/specs/1c-integration-tz.md`](../../docs/specs/1c-integration-tz.md): JSON UTF-8, файл в папку обмена, ошибки — не «тихий» пропуск.

```json
{
  "event": "shipment_completed",
  "version": 1,
  "shipment_id": 123,
  "shipment_date": "2026-08-05",
  "completed_at": "2026-08-05T14:32:00",
  "delivery_type": "delivery",
  "orders": [{ "kp_id": 456, "ya_order_no": "ЯР-0001467", "uid_kp": null }],
  "items": [
    { "type": "plate", "plate_name": "ПБ 60-12-8", "length_m": 6.0, "width_m": 1.2,
      "nomenclature_id": null, "qty": 10, "weight_kg": 27000 },
    { "type": "free", "mark": "С60.30", "qty": 14, "weight_kg": 19320 }
  ],
  "carrier": { "name": "ООО ТрансЛогистик" },
  "driver_name": "Иванов И.И.", "vehicle_text": "Volvo FH / а123бв77",
  "upd_no": "1234", "total_weight_kg": 46320
}
```

- Файл `shipment_completed_{shipment_id}_{ts}.json` в `EXCHANGE_EXPORT_DIR`.
- Запись **после COMMIT** закрытия рейса; ошибка записи → лог, закрытие не откатывается (событие можно выгрузить повторно вручную).
- `SHIPMENT_EVENTS_ENABLED=false` по умолчанию до интеграции G. Схема — черновик, финал с интегратором 1С (Open Question P1).

---

## API (черновик контракта)

| Method | Path | Назначение |
|--------|------|------------|
| `GET` | `/api/v1/logistics/shipments?date_from&date_to&kp_id&carrier_id&delivery_type&status&no_upd=1&attention=1` | Реестр с фильтрами |
| `POST` | `/api/v1/logistics/shipments` | Создать рейс: `{ shipment_date, delivery_type, kp_ids: [≥1] }` |
| `GET` | `/api/v1/logistics/shipments/{id}` | Карточка (поля + orders + items + доступный СГП по КП) |
| `PATCH` | `/api/v1/logistics/shipments/{id}` | Поля: ТС, водитель, УПД, внимание, перевозчик, стоимость… |
| `POST` | `/api/v1/logistics/shipments/{id}/propose` | `vehicle_class` → предложение состава (не сохраняет). **Реализация:** query-параметр, не body (зафиксировано в Phase 6) |
| `PUT` | `/api/v1/logistics/shipments/{id}/items` | Confirm: полная замена состава (plate + free строки) |
| `POST` | `/api/v1/logistics/shipments/{id}/complete` | Выезд: списание СГП, audit, прогресс, DONE-check, событие |
| `POST` | `/api/v1/logistics/shipments/{id}/cancel` | Отмена (только `in_work`) |
| `GET` | `/api/v1/logistics/shipments/{id}/sheet.xlsx` | Печатная форма «Лист отгрузки» |
| `GET` | `/api/v1/logistics/carriers?active=1&q=` | Справочник (автокомплит) |
| `POST` | `/api/v1/logistics/carriers/{id}/merge` | `{ into_id }` — слияние дублей |
| `GET` | `/api/v1/logistics/pile-catalog?q=` | Автокомплит марок свай (марка, вес) |
| — | расширить archive DTO | `shipped_progress: { x, m }` рядом с `sgp_progress` |

Ошибки:

- `422` — доменные: `shipment_no_availability`, `shipment_missing_ya_order`, `shipment_not_in_work`, `sgp_row_allocated`, `carrier_merge_conflict`
- `409` — конфликт версий (если появится конкурентное редактирование карточки)
- Роль `logistics`/`admin` — `403` через существующий guard

---

## UI

### Раздел «Логистика» (route `/logistics`, роль `logistics` + `admin`)

Отдельный раздел, **не вкладка производства** (decision ideation: у логистики свой контур ролей). Nav-item виден только ролям с доступом.

### Реестр рейсов

Плоская таблица: дата, заказы ЯР (стек), заказчик(и), тип Д/С, перевозчик/доверенность, водитель, ТС, вес, УПД, стоимость план, статус, «внимание» (иконка + тултип комментария).  
Фильтры: дата (диапазон), заказ (поиск ЯР/заказчика), перевозчик, тип, «без УПД», «внимание».  
Строка → карточка рейса. Кнопка «Новый рейс».

### Карточка рейса (ShipmentDrawer)

1. Шапка: дата, тип Д/С, заказы КП (добавить/убрать, ЯР-номер у каждого), статус + «внимание».
2. **Состав:** кнопка «Предложить состав» (propose) → таблица items с весами → правка (qty, удалить, добавить plate из доступных, добавить free-строку с автокомплитом марок свай) → «Утвердить состав».
3. Панель веса: Σ кг, лимит класса ТС, красное предупреждение при перегрузе (не блокирует).
4. Поля дозаполнения: перевозчик (автокомплит) / доверенность №, водитель, ТС (текст + класс), УПД, № заявки на фрахт, стоимость план.
5. Действия: «Лист отгрузки (XLSX)», «Отменить рейс» (только в работе), **«Выезд»** — подтверждение → `complete`.

### Справочник перевозчиков (`/logistics/carriers`)

Таблица (имя, источник, рейсов, active) + поиск. Действие «Слить с…» (merge dialog: выбрать целевого, подтверждение с числом переносимых рейсов).

### Архив КП

Бейдж **«отгружено X/M»** рядом с существующим «N/M на СГП». КП `выполнено` уходит во «Выполненные» (существующее поведение архива для DONE).

---

## Testing Strategy

| Уровень | Что | Где |
|---------|-----|-----|
| Unit | availability: вычет open-рейсов, cancel восстанавливает | `tests/test_shipment_service.py` |
| Unit | propose: FIFO, лимит класса, «не влезло», мульти-КП | `tests/test_shipment_service.py` |
| Unit | DONE-критерий: полный qty → «выполнено»; частичный/брак → нет | `tests/test_shipment_service.py` |
| Unit | нормализация имён (ОПФ, кавычки, ё), дедуп, merge переносит рейсы | `tests/test_carrier_import.py` |
| Unit | парсинг pile_catalog: 44 марки, `С137,5.40` (запятая), «-» → NULL | `tests/test_pile_catalog_import.py` |
| Integration | create→propose→confirm→complete: атомарность, rollback на 422, БД без изменений при ошибке | `tests/test_shipment_service.py` |
| Integration | audit: `sgp_ship` строки с qty и `shipment_id`; from `on_sgp` → to `shipped` | `tests/test_shipment_service.py` |
| Integration | unlink/relink по зарезервированной строке → 422 `sgp_row_allocated` | `tests/test_shipment_service.py` + sgp tests |
| Qty balance | Σ(kp_plates)+Σ(completed)+Σ(shipped) = const до/после complete/cancel | `tests/test_shipment_qty_balance.py` |
| Regression | `run_plate_loss_regression.py` остаётся PASS (orphan Σ=0) | script |
| Frontend | фильтры реестра; propose→confirm; предупреждение перегруза | vitest / manual |
| Manual | печатная форма; справочник merge; права роли | browser |

**Пилот (S3, процедура, не код):**

1. **Шаг 0:** инвентаризация — сверка «в системе = на складе», расхождения устранены до старта.
2. Неделя cut-over: все рейсы одного типа выдачи — только в системе; Excel read-only.
3. Метрики: доля рейсов без ручной правки propose (по `propose_snapshot`); расхождения склада в конце недели.

---

## Boundaries

### Always

- Атомарные транзакции для `complete` / `cancel` / confirm состава / merge (как P0 `complete_day`)
- Писать `plate_status_log` в той же транзакции (`sgp_ship`, `qty`, `shipment_id`, `actor`)
- Split при частичном qty; availability перепроверять при выезде
- Вес — **ориентир** для лимита ТС, не учётный (учётный — в 1С для УПД)
- Прогонять qty-balance + `run_plate_loss_regression.py` перед merge
- Сообщения API на русском; коды ошибок стабильные

### Ask first

- Изменение критерия `KpStatus.DONE` (полный qty)
- Включение `SHIPMENT_EVENTS_ENABLED` в проде / смена папки обмена
- Добавление полей MVP-2 в схему (реквизиты, ПДн водителей, drivers/vehicles)
- Обязательность ЯР на создании рейса (сейчас — только на закрытии)
- Перенос `ya_order_no` на сторону КП (после 1С-интеграции — связка `УИДКП/НомерКП`)
- Повторная выгрузка/переотправка события `shipment_completed` из UI

### Never

- Списывать `completed_plates` без audit-строки в той же транзакции
- Отменять/редактировать `done`-рейсы (возврат/откат отгрузки — отдельная фаза, как «откат дня» в СГП)
- Выставлять `KpStatus.DONE` при неполном отгруженном qty
- Требовать ЯР-номер при создании рейса (планирование идёт до документов)
- Хранить ПДн водителей (паспорта) в MVP-1
- Мигрировать 17k строк истории Excel (остаётся read-only архивом снаружи)
- Заводить рейс без КП (decision 7)
- Коммитить секреты / `.env` / реальные Excel с ПДн

---

## Success Criteria

Конкретные, проверяемые (код):

1. **Реестр:** фильтры дата/заказ/перевозчик/тип/«без УПД»/«внимание» работают на пилотных данных; строка = рейс (ТС), не заказ.
2. **Propose:** по рейсу с КП-X предложены только свободные linked-плиты КП-X (FIFO), с весом; при `t20` и 25т на СГП — 5т в секции «не влезло».
3. **Confirm + выезд:** после `complete` каждая plate-строка уменьшила `completed_plates` ровно на qty (split при частичном); в `plate_status_log` — `sgp_ship` с `qty`, `shipment_id`, actor.
4. **Частичная отгрузка:** отгружено 8 из 10 → КП «в работе»/«На СГП» + бейдж «отгружено 8/M»; `DONE` нет.
5. **Полная отгрузка:** shipped = M, kp_plates пуст, СГП linked = 0 → `kp_meta = «выполнено»` в той же транзакции.
6. **Мульти-заказ:** рейс с КП-A и КП-B: прогресс и DONE-check по каждому независимо; ЯР обязателен у обоих (422 `shipment_missing_ya_order` при пустом).
7. **Сваи:** free-строка «С60.30» × 14 → автовес 19 320 кг из `pile_catalog`; ручная правка сохраняется.
8. **Перегруз:** Σ веса > лимита класса → видимое предупреждение, закрытие не блокируется.
9. **Cancel:** отмена `in_work` рейса → availability восстановилась, `completed_plates` не тронуты, audit нет.
10. **Защита резерва:** unlink/relink зарезервированного open-рейсом qty → 422 `sgp_row_allocated`.
11. **Qty gate:** `test_shipment_qty_balance.py` + `run_plate_loss_regression.py` — PASS.
12. **Справочник:** импорт 2 397 строк → отчёт о схлопнутых дублях; merge переносит рейсы и гасит дубль.
13. **Событие:** при `SHIPMENT_EVENTS_ENABLED=1` в `EXCHANGE_EXPORT_DIR` появляется валидный JSON по контракту; при `0` — ничего, закрытие работает.
14. **Роль:** пользователь `logistics` видит раздел; `manager`/`production` — нет (403).

Пилот (gate, зафиксировать в конце недели):

- Доля рейсов без ручной правки propose **≥ 50%** (иначе propose бесполезен — пересмотр алгоритма).
- Расхождения «в системе = на складе» по итогам недели — 0 или задокументированы.
- Логист подтверждает замену Excel (buy-in), критичные замечания заведены.

---

## Out of Scope

### MVP-2 (отдельная спека, сразу следом)

- Полная карточка перевозчика: реквизиты, банк, налоговый режим; под-сущности `Driver` (паспорт — ПДн, доступ только роли логистики), `Vehicle` (класс 20т/30т+)
- Генерация договора-заявки DOCX из шаблона (константы заказчика — конфиг), номер `N/MM/YY` с месячным счётчиком
- Поля доверенности при самовывозе (номер/дата/срок/на кого — ждём образец акта)

### Фаза 2 / позже

- Возврат/откат отгрузки (редактирование `done`-рейса)
- НДС 5%/7%, факт-стоимость, маржа, учёт оплат фрахта (деньги — в 1С)
- Склад свай/прочей номенклатуры как полноценный учёт (free-строки остаются)
- Умный подбор состава (оптимизатор укладки, довесок)
- Услуги техники и входящая логистика (жёлтые строки: кран, манипулятор, доски) — тип события зарезервировать, не реализовывать
- ЭДО/ЭТрН, ГИС ЭПД — через 1С (направление H роадмапа); GPS, приложение водителя — нет своего парка
- Реальная отправка события в 1С (приходит с интеграцией G)
- Миграция истории Excel; «Номер счёта» отдельным полем; КПП/пропуска; 4-статусный конвейер из мёртвой инструкции

---

## Open Questions (на Plan, не блокируют SPEC approval)

| # | Вопрос | Дефолт в спеке |
|---|--------|----------------|
| P1 | Финальный JSON-контракт `shipment_completed` — согласовать с интегратором 1С (встреча по ТЗ) | Черновик выше; flag выключен |
| P2 | Точные лимиты классов ТС (20т/30т+ в кг, gross vs net) — уточнить у логиста | `t20=20000`, `t30plus=30000` в конфиге |
| P3 | Автоподстановка ЯР из предыдущего рейса того же КП (UX-сахар) | Решить на Plan; хранение — `shipment_orders.ya_order_no` |
| P4 | Обновление `pile_catalog` при новом прайсе | Ручной re-import скриптом (upsert по mark) |
| P5 | Поля доверенности самовывоза (ждём образец акта от предприятия) | MVP-1: одно поле `proxy_no` |
| P6 | Фиксировать ли а/м клиента при самовывозе | `vehicle_text` свободный, допустимо |

---

## As-Is → To-Be

| As-Is (Excel-реестр) | To-Be (раздел «Логистика») |
|----------------------|----------------------------|
| Строка вбивается текстом, позиции «по памяти» | Рейс собирается propose→confirm из реального СГП |
| Резерв смотрят в счёте в 1С вручную | Availability вычисляется системой (СГП − open-рейсы) |
| Статус = цвет ячейки (красный/голубой/жёлтый) | `в работе → обработано` + флаг «внимание» с комментарием |
| Списание склада — устное/в 1С постфактум | Выезд = атомарное списание + audit `sgp_ship` |
| Прогресс заказа не виден менеджеру | Бейдж «отгружено X/M» в архиве; `DONE` при полном qty |
| Перевозчик — свободный текст, 2 397 строк с дублями | Справочник с дедупом при импорте + merge в UI |
| Вес «ориентировочный» из головы | Автовес: плиты — формула, сваи — `pile_catalog`; лимит ТС с предупреждением |
| Событие для 1С отсутствует | `shipment_completed` JSON по контракту папки обмена (flag off) |
| КПП/пропуска — пустые колонки | Не фиксируем вообще |

---

## Verification (конец Phase SPECIFY)

- [x] Spec covers Objective, Commands, Structure, Style, Testing, Boundaries
- [x] Decisions locked (1–15 из ideation + S1–S4 уточнений SPECIFY)
- [x] Success criteria — конкретные, тестируемые (код + пилот-gate)
- [x] Boundaries Always/Ask first/Never определены
- [x] Spec сохранена в репозиторий (`ai_docs/specs/shipment-logistics.md`)
- [x] **Human reviewed и approved** (2026-07-31)
- [x] **Stop:** не IMPLEMENT до Plan + Tasks (Phase 2–3)

---

## Next

1. **Review этой спеки** → правки/approval
2. Phase PLAN → `ai_docs/develop/plans/2026-07-31-shipment-logistics.md` (компоненты, порядок, риски, чекпоинты; закрыть P1–P6)
3. Phase TASKS → чеклист SHIP-000… по образцу SGP-чеклиста
4. Implement с TDD по checkpoint'ам (`incremental-implementation`, `test-driven-development`)
