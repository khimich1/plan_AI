# Spec: Переиспользование рейса — клон транспорта

> **Источник идеи:** [`ai_docs/ideas/shipment-reuse-transport.md`](../ideas/shipment-reuse-transport.md) (ideation 2026-08-02)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS ✅ → IMPLEMENT ✅  
> **План:** [`ai_docs/develop/plans/2026-08-02-shipment-reuse-transport.md`](../develop/plans/2026-08-02-shipment-reuse-transport.md)  
> **Связанные:** [`shipment-logistics.md`](./shipment-logistics.md), `app/services/shipment_service.py`, `app/api/v1/endpoints/logistics.py`, `frontend/src/features/logistics/`  
> **Дата:** 2026-08-02  
> **Scope:** кнопка «На основе» в реестре → новый рейс с копией только транспортного блока. Без полного дубля, без шаблонов перевозчика, без входа из карточки.

---

## Decisions locked (SPECIFY → PLAN, 2026-08-02)

| # | Тема | Решение |
|---|------|---------|
| R1 | API | Атомарный `POST .../shipments/{source_id}/reuse-transport` (не client create+patch) |
| R2 | КП | Обязателен; кнопка открывает диалог, не мгновенный create |
| R3 | Confirm | Диалог создания = confirm; отдельного «Вы уверены?» нет |
| R4 | delivery_type | Не в белом списке клона; в диалоге префилл из источника |
| R5 | Дата | По умолчанию сегодня, редактируема |
| R6 | Статус источника | `in_work` и `done` допустимы |
| R7 | Пустой транспорт | Копируем null; кнопку не блокируем |
| R8 | UX строки | `stopPropagation` на кнопке — не открывать drawer источника |
| R9 | Роль / audit | `REQUIRE_LOGISTICS`; actor как у `create`; отдельный audit-event не заводим |
| R10 | Текст кнопки | «На основе», `title="Создать на основе"` |
| R11 | Битый carrier_id | Копировать id как есть (merge сохраняет id) |

### Белый список клона

`carrier_id`, `driver_name`, `vehicle_text`, `vehicle_class`, `proxy_no`

### Чёрный список (никогда)

`upd_no`, `freight_request_no`, `planned_cost`, `attention*`, `items`, source `orders`, `status`, `completed_at`, snapshots

---

## Objective

Логист из строки реестра создаёт новый рейс, не перепечатывая перевозчика, водителя, ТС, доверенность и класс ТС.

**Пользователь:** логист (роль `logistics`).

**Проблема сейчас:** транспортный блок вводится вручную в карточке (`ShipmentFieldsSection`) на каждый рейс; дублирования в API/UI нет. Create требует только дату/тип/КП и не принимает транспорт.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | логист | нажать «Создать на основе» в строке реестра | не искать прошлый рейс глазами и не копировать поля вручную |
| US-2 | логист | в диалоге выбрать дату, тип Д/С и КП | новый рейс сразу был привязан к заказу (инвариант create) |
| US-3 | логист | получить карточку с уже заполненным транспортом | править только отличия и сразу собирать состав |
| US-4 | логист | чтобы УПД / № заявки / состав / заказы **никогда** не копировались | не перепутать документы и не утащить чужие плиты |

### Acceptance criteria (MVP)

- [x] В реестре у каждой строки есть действие «Создать на основе» (кнопка/ссылка в отдельной колонке или в конце строки)
- [x] Клик по действию открывает диалог создания с префиллом: дата=сегодня, `delivery_type`=из источника; поиск/выбор КП — как в `CreateShipmentDialog`
- [x] Submit вызывает `POST .../shipments/{source_id}/reuse-transport` с `{shipment_date, delivery_type, kp_ids}`
- [x] Ответ — карточка нового рейса (`ShipmentCard`); `status=in_work`; транспорт скопирован: `carrier_id`, `driver_name`, `vehicle_text`, `vehicle_class`, `proxy_no`
- [x] В новом рейсе **пусты/null:** `upd_no`, `freight_request_no`, `planned_cost`, `attention`, `items`; `orders` — только выбранные в диалоге КП (не из источника)
- [x] Источник `404` → структурированная ошибка; невалидные `kp_ids` — те же коды, что у `create`
- [x] После успеха диалог закрывается, открывается drawer нового рейса, список реестра инвалидируется
- [x] Тесты: service (копирует только транспорт), API (TestClient), UI (кнопка + вызов + stopPropagation)
- [x] `pytest -k "shipment or logistics"` и `npm run build` зелёные

### Out of scope (Not Doing)

- Полный дубль рейса / копирование `items` / `orders` источника
- Копирование `upd_no`, `freight_request_no`, `planned_cost`, `attention*`
- Шаблоны «перевозчик+водитель» в справочнике
- Автоподстановка last-by-carrier без явной кнопки
- Кнопка в `ShipmentDrawer` (тот же API — post-MVP UI)
- Изменение инварианта «create требует ≥1 КП»
- Серии рейсов / N машин на один план

---

## API

### `POST /api/v1/logistics/shipments/{source_id}/reuse-transport`

**Auth:** `REQUIRE_LOGISTICS`

**Request body** (как у create, без транспорта — транспорт с сервера):

```json
{
  "shipment_date": "2026-08-02",
  "delivery_type": "delivery",
  "kp_ids": [154]
}
```

**Response:** `200` → `ShipmentCard` нового рейса.

**Поведение сервиса (псевдо):**

1. Загрузить source; если нет → `ShipmentError(shipment_not_found)`.
2. Создать рейс как `create(shipment_date, delivery_type, kp_ids, actor)`.
3. В той же транзакции проставить из source: `carrier_id`, `driver_name`, `vehicle_text`, `vehicle_class`, `proxy_no`.
4. Вернуть `get(new_id)`.

**Никогда не копировать из source:** `upd_no`, `freight_request_no`, `planned_cost`, `attention`, `attention_comment`, `items`, `orders`, `status`, `completed_at`, `freight`/`propose` snapshots.

**Ошибки:** те же `_ERROR_4XX`, что у logistics endpoints (`404` / `422` через `_raise_domain_error`).

---

## Tech Stack

| Слой | Стек |
|------|------|
| Service | `app/services/shipment_service.py` — метод `reuse_transport(source_id, ...)` |
| API | `app/api/v1/endpoints/logistics.py` — новый route |
| Schema | `app/schemas/logistics.py` — request = тот же shape, что `ShipmentCreateRequest` (можно алиас/reuse класса) |
| Frontend | `features/logistics/` — кнопка в `LogisticsRegistryView`, расширение `CreateShipmentDialog` (`sourceShipmentId`), API + hook |
| Tests | `tests/test_shipment_service.py`, `tests/test_logistics_api.py`, `LogisticsRegistryView.test.tsx` / dialog test |

## Commands

```bash
# Backend
source .venv/bin/activate
pytest tests/test_shipment_service.py tests/test_logistics_api.py -q
pytest tests/ -k "shipment or logistics" -q

# Frontend
cd frontend && npm test -- --run src/features/logistics
cd frontend && npm run build
```

## Project Structure

```
app/schemas/logistics.py              → ShipmentReuseTransportRequest (= Create) или reuse Create
app/services/shipment_service.py      → reuse_transport(...)
app/api/v1/endpoints/logistics.py     → POST /shipments/{id}/reuse-transport

frontend/src/features/logistics/
  api/logisticsApi.ts                 → reuseTransport(sourceId, payload)
  hooks/useLogisticsQueries.ts        → useReuseTransportMutation
  types/logistics.ts                  → типы payload (можно = CreateShipmentPayload)
  components/CreateShipmentDialog.tsx → optional sourceShipmentId + другой submit path
  components/LogisticsRegistryView.tsx → колонка/кнопка «На основе»
  components/LogisticsRegistryView.test.tsx → клик не открывает drawer источника
```

## Code Style

Backend — как существующий `create`/`patch`: `ShipmentError(code=...)`, без ORM-утечек в endpoint.

```python
def reuse_transport(
    self,
    source_id: int,
    *,
    shipment_date: str,
    delivery_type: str,
    kp_ids: list[int],
    actor: str | None,
) -> ShipmentCard:
    ...
```

Frontend — named export, русские тексты, `stopPropagation` на кнопке строки; переиспользовать UI `CreateShipmentDialog` через prop `sourceShipmentId?: number` вместо второго модала.

## Testing Strategy

| Уровень | Что проверяем | Где |
|---------|----------------|-----|
| Service | копируются 5 полей; НЕ копируются upd/freight/items/orders/attention/cost; новый status=in_work; 404 источника | `tests/test_shipment_service.py` |
| API | 200 + тело; 404; 422 без kp_ids; роль logistics | `tests/test_logistics_api.py` |
| Frontend | кнопка вызывает reuse с source id; клик не селектит строку; после успеха открывается новый id | registry/dialog tests |

## Boundaries

- **Always:** pytest после backend; `npm run build` после frontend; минимальный diff; копировать только белый список полей
- **Ask first:** ослабление «create ≥1 КП»; кнопка в drawer; копирование `planned_cost` / `delivery_type` как жёсткий clone
- **Never:** копировать `upd_no` / `freight_request_no` / `items` / source `orders`; коммит без просьбы; трогать packing/propose

## Success Criteria

1. Из рейса с заполненным транспортом за ≤3 клика получается новый рейс с тем же перевозчиком/водителем/ТС/доверенностью/классом и другим набором КП.
2. В новом рейсе УПД и № заявки пустые; состав пустой; заказы только выбранные в диалоге.
3. Регрессия logistics/shipment зелёная.

## Open Questions

Нет блокирующих. Post-MVP: кнопка в `ShipmentDrawer`; шаблоны перевозчика.
