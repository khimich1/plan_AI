# План: Карточка заказа «На согласовании»

**Дата:** 2026-09-28  
**Статус:** не начат. Код не писать до приёмки плана  
**Идея:** `ai_docs/ideas/archive-on-approval.md`  
**Спека:** `ai_docs/specs/archive-on-approval.md`

## Overview

В архиве четыре вкладки. Кнопка «Отправить в 1С» пишет файл «Счёт на оплату» в папку обмена и переводит заказ в «на согласовании». Карточка делится на шапку с одним действием, итоги с составом и полку документов. Номер счёта карточка только читает из `order_number_1c`: разбор обратного JSON и поле ввода не входят. Правка идёт прежним конструктором. «Отправить исправление» пишет второй файл, когда номер уже лежит в колонке и состав изменился. «В производство» включается при отметке «Оплачен» и непустом номере, затем открывает сегодняшний диалог срока.

## Architecture Decisions

- **Та же запись КП.** Отдельной сущности «счёт» нет. Статус — строка `kp_meta.status` = `на согласовании`. Секция списка — `on_approval`.
- **Сборщик не пишет файл.** `core/invoice_export.py` возвращает словарь и хеш снимка. Файл в `exchange_export_dir` пишет сервис архива после ворот, и только потом меняет статус.
- **Имя документа — «Счёт на оплату».** `attach_contract_number` принимает и его, и прежнее «КоммерческоеПредложение», чтобы штамп договора не сломать.
- **Суммы не пересчитываются.** В JSON попадают цифры карточки этого `kp_id`.
- **Номер не вводится.** Маршрута записи `order_number_1c` нет. Тесты, которым нужен номер, кладут его в колонку напрямую. Пустая колонка гасит исправление и производство.
- **Оплата — `kp_meta.paid_at`.** Снять можно только пока статус «на согласовании».
- **Хеш снимка** в `KP_offers.invoice_snapshot_hash`. Кнопка исправления сравнивает текущий состав с ним. Клиент хеш не считает: карточка отдаёт `correction_pending`.
- **Конструктор не возвращает статус в архив.** Дополнение и сохранение принимают «на согласовании».
- **Договор на согласовании только читается.** Создание по-прежнему только из «в архиве».
- **Крах после записи файла и до коммита статуса** может оставить второй `create` при повторе. Дедуп — `НомерВПриложении` на стороне 1С. Отдельной таблицы исходящих нет.

## Dependency graph

```
сборщик JSON + хеш + штамп «Счёт на оплату»
    │
    ├── колонки paid_at, invoice_snapshot_hash
    │       │
    │       └── секция on_approval
    │               │
    │               ├── POST invoice-export
    │               │       ├── POST invoice-correction
    │               │       ├── POST payment
    │               │       └── move_to_production (статус + paid_at + номер)
    │               │
    │               └── конструктор сохраняет «на согласовании»
    │
    └── чтение договора при статусе «на согласовании»
            │
            └── вкладка и карточка
```

Разбор `invoice_ack` в граф не входит.

## Task List

### Phase 1 — Документ без HTTP

- [x] **OA-001: Сборщик счёта**
  - **Description:** `build_invoice_document` и хеш снимка. `create` без `НомерЗаказа`. `update` требует номер. `НомерДоговора` ставит существующий штамп. Штамп принимает документ «Счёт на оплату» и не ломает прежнее имя «КоммерческоеПредложение».
  - **Acceptance:**
    - [x] `Документ` = «Счёт на оплату», `event` = `invoice_export`, `НомерДоговора` сразу после `Контрагент`
    - [x] `create` не содержит `НомерЗаказа`; `update` без номера не собирается
    - [x] Суммы равны переданным цифрам карточки, сборщик их не пересчитывает
    - [x] Одинаковый состав даёт одинаковый хеш, изменённое количество — другой
  - **Verify:** `pytest tests/test_archive_invoice_export.py tests/test_supply_contract.py -q -k "invoice or contract_number or stamp"`
  - **Dependencies:** —
  - **Files:** `core/invoice_export.py`, `core/supply_contract.py`, `tests/test_archive_invoice_export.py`, `tests/test_supply_contract.py`
  - **Scope:** M

### Checkpoint: Phase 1

- [x] Сборщик зелёный без схемы, без файла на диске и без HTTP

### Phase 2 — Список и первая отправка

- [x] **OA-002: Колонки и секция**
  - **Description:** `paid_at` на `kp_meta`, `invoice_snapshot_hash` на `KP_offers`. Повтор схемы не падает. `get_all_kp_list` кладёт статус `на согласовании` только в `on_approval`. `ArchiveSection` на бэкенде и во фронтовых типах принимает `on_approval`.
  - **Acceptance:**
    - [x] Старые статусы остаются в своих секциях: «в архиве», «в работе», «На СГП», «выполнено»
    - [x] `section=on_approval` отдаёт только «на согласовании»
  - **Verify:** `pytest tests/test_archive_endpoints.py tests/test_archive_service.py -q -k "section or on_approval or schema"`
  - **Dependencies:** —
  - **Files:** `core/kp_db_schema.py`, `core/kp/offers_read.py`, `app/schemas/archive.py`, `frontend/src/features/commercial-archive/types/archive.ts`, `tests/test_archive_endpoints.py`
  - **Scope:** M

- [x] **OA-003: Отправить в 1С**
  - **Description:** `POST /commercial/archive/{kp_id}/invoice-export`. Ворота: статус «в архиве», контрагент, действующий договор, `assert_invoice_guids`. После записи `invoice_create_{kp_id}_{ts}.json` статус становится «на согласовании», пишется хеш. Ошибка ворот или записи файла статус не меняет. Ответ — карточка.
  - **Acceptance:**
    - [x] Нет контрагента, нет договора или нет GUID — файла нет, статус «в архиве», текст по-русски
    - [x] Успех создаёт файл и переводит статус
    - [x] Повтор из «на согласовании» отказывает и второй create не пишет
  - **Verify:** `pytest tests/test_archive_invoice_export.py tests/test_archive_service.py tests/test_archive_endpoints.py -q -k "invoice_export or move_to_production"`
  - **Dependencies:** OA-001, OA-002
  - **Files:** `app/services/archive_service.py`, `app/api/v1/endpoints/archive.py`, `app/schemas/archive.py`, `tests/test_archive_service.py`, `tests/test_archive_endpoints.py`
  - **Scope:** M

### Checkpoint: Phase 2

- [x] Из архива уходит файл и заказ виден только в `on_approval`
- [x] Каталог выгрузки в тестах временный

### Phase 3 — Номер, оплата, исправление, конструктор

- [x] **OA-004: Оплата и производство**
  - **Description:** `POST /commercial/archive/{kp_id}/payment` с `{ "paid": true|false }`. `move_to_production` принимает только «на согласовании» при непустых `paid_at` и `order_number_1c`. Из «в архиве» отказ. Снятие оплаты в «в работе» отказ. Диалог срока не меняется. Номер в этих тестах записывается в колонку напрямую, не через API.
  - **Acceptance:**
    - [x] Оплата без номера не переводит в производство
    - [x] Номер без оплаты не переводит
    - [x] Оба условия открывают прежний переход в «в работе»
    - [x] `paid: false` очищает `paid_at` только в «на согласовании»
  - **Verify:** `pytest tests/test_archive_service.py tests/test_archive_endpoints.py -q -k "payment or move_to_production"`
  - **Dependencies:** OA-002, OA-003
  - **Files:** `app/services/archive_service.py`, `app/api/v1/endpoints/archive.py`, `tests/test_archive_service.py`
  - **Scope:** M

- [x] **OA-005: Исправление**
  - **Description:** `POST /commercial/archive/{kp_id}/invoice-correction`. Пишет `invoice_update_{kp_id}_{ts}.json` с `НомерЗаказа` только при статусе «на согласовании», непустом `order_number_1c` и хеше, отличном от сохранённого. После успеха хеш обновляется, статус тот же. Карточка отдаёт `correction_pending`.
  - **Acceptance:**
    - [x] Без номера — отказ, файла нет
    - [x] С номером и тем же снимком — отказ
    - [x] Изменённый состав — файл с `НомерЗаказа`, статус «на согласовании»
  - **Verify:** `pytest tests/test_archive_invoice_export.py tests/test_archive_service.py -q -k "invoice_correction or snapshot"`
  - **Dependencies:** OA-001, OA-003
  - **Files:** `app/services/archive_service.py`, `app/api/v1/endpoints/archive.py`, `app/schemas/archive.py`, `tests/test_archive_service.py`
  - **Scope:** M

- [x] **OA-006: Конструктор не сбрасывает статус**
  - **Description:** Дополнение и сохранение черновика принимают «на согласовании». Успешное сохранение оставляет этот статус. Привязка контрагента по-прежнему только из «в архиве».
  - **Acceptance:**
    - [x] Resume и save из «на согласовании» не возвращают «в архиве»
    - [x] Из «в работе» дополнить по-прежнему нельзя
  - **Verify:** `pytest tests/test_archive_endpoints.py tests/test_commercial_multi_append_flow.py -q -k "resume or archive or approval"`
  - **Dependencies:** OA-002
  - **Files:** `app/services/commercial_draft_lifecycle.py`, `tests/test_archive_endpoints.py`, `tests/test_commercial_multi_append_flow.py`
  - **Scope:** S

- [x] **OA-007: Договор на согласовании только открывается**
  - **Description:** Чтение и скачивание действующего договора разрешены при статусе «на согласовании». Создание и замена по-прежнему требуют «в архиве».
  - **Acceptance:**
    - [x] GET договора для «на согласовании» отдаёт уже существующий номер
    - [x] POST создания из «на согласовании» отказывает
  - **Verify:** `pytest tests/test_supply_contract.py tests/test_archive_endpoints.py -q -k "supply_contract and status"`
  - **Dependencies:** OA-002
  - **Files:** `app/services/supply_contract_service.py`, `tests/test_supply_contract.py`
  - **Scope:** S

### Checkpoint: Phase 3

- [x] Исправление, оплата и производство закрыты тестами без UI
- [x] Маршрута записи номера нет

### Phase 4 — Карточка

- [x] **OA-008: Вкладка**
  - **Description:** Четвёртая вкладка «На согласовании» между «В архиве» и «В производстве». Список ходит в `section=on_approval`.
  - **Acceptance:**
    - [x] Вкладки четыре, в выбранной видны только заказы её секции
  - **Verify:** `cd frontend && npm run test -- src/features/commercial-archive/components/ArchiveSectionTabs.test.tsx src/pages/commercial-offer-archive/CommercialOfferArchivePage.test.tsx`
  - **Dependencies:** OA-002
  - **Files:** `frontend/src/features/commercial-archive/components/ArchiveSectionTabs.tsx`, `frontend/src/features/commercial-archive/api/archiveApi.ts`, `frontend/src/features/commercial-archive/hooks/useArchiveQueries.ts`, тесты вкладок и страницы
  - **Scope:** S

- [x] **OA-009: Три полки**
  - **Description:** Шапка показывает фазу и одну главную кнопку. «Отправить в 1С» жива только когда можно отправить; иначе серая с причиной (контрагент, договор, GUID). На согласовании без номера — серое «В производство» с подписью «нужен номер счёта» и текст «Исправление отправится, когда номер будет известен», если `correction_pending`. С номером и без оплаты — «после оплаты»; при `correction_pending` ещё «Отправить исправление». С номером и `paid_at` — фиолетовое «В производство», диалог срока прежний. «Оплачен» / снятие оплаты в шапке. Итоги на обоих статусах только для чтения, «(+ Добавить)» и «Редактировать» открывают конструктор. PDF, XLSX, схема и договор — полка «Документы». «Удалить КП» отдельно. График поставки и блок готовности на «на согласовании» скрыты. Для видов, где производство сегодня «скоро», кнопка производства остаётся недоступной и после оплаты.
  - **Acceptance:**
    - [x] В архиве «В производство» нет
    - [x] На согласовании нет поля ввода номера
    - [x] Документы не стоят в одном ряду с фиолетовой кнопкой
    - [x] Состояния кнопок совпадают с таблицей фаз в спеке
  - **Verify:** `cd frontend && npm run test -- src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx && npm run typecheck`
  - **Dependencies:** OA-003, OA-004, OA-005, OA-007, OA-008
  - **Files:** `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`, `OfferDetailsDrawer.test.tsx`, `frontend/src/features/commercial-archive/api/archiveApi.ts`, `frontend/src/features/commercial-archive/types/archive.ts`
  - **Scope:** M

### Checkpoint: Phase 4

- [x] `pytest tests/test_archive_invoice_export.py tests/test_archive_service.py tests/test_archive_endpoints.py -q`
- [x] `cd frontend && npm run test -- src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx src/features/commercial-archive/components/ArchiveSectionTabs.test.tsx`
- [x] `cd frontend && npm run typecheck`
- [ ] В карточке пройдены десять сценариев спеки, кроме живого ответа 1С: номер для сценариев 6–8 подставляется в колонку тестом или вручную в базе

## Risks and Mitigations

| Риск | Влияние | Что делаем |
|---|---|---|
| Пока нет разбора ответа, живой заказ не дойдёт до производства | Высокое для цеха, ожидаемое | Экран ждёт номер. Разбор — отдельная задача. В тестах номер пишется в колонку |
| Повторный `create`, если процесс умер после файла и до статуса | Среднее | Статус меняется только после записи. Дедуп по `НомерВПриложении` оставляем 1С |
| Штамп договора завязан на имя «КоммерческоеПредложение» | Среднее | OA-001 принимает оба имени. Старые тесты договора остаются зелёными |
| Старые тесты «В производство» из архива и конструктора | Среднее | OA-004 и OA-006 меняют ожидания вместе с правилом |
| НДС в примере JSON расходится с карточкой старых КП | Низкое для этого контура | Сборщик берёт цифры карточки и не вводит вторую формулу |

## Not Doing

- Разбор `invoice_ack` и sync-агент
- Поле ввода номера
- Спецификация к заказу
- HTTP в 1С
- График поставки и блок готовности на согласовании
- Смена контрагента после отправки

## Open Questions

Блокирующих нет. Имя «Счёт на оплату» и ключ `НомерЗаказа` можно переименовать при первой сверке с 1С, не меняя экран.
