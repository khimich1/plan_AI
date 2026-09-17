# План: Свободное имя клиента, контрагент 1С — на производстве

**Дата:** 2026-09-17  
**Статус:** реализовано  
**Идея:** `ai_docs/ideas/kp-counterparty-bind-on-production.md`  
**Спека:** `ai_docs/specs/kp-counterparty-bind-on-production.md`

## Overview

Ослабить обязательный `counterparty_id` на шаге «Клиент» и при первом `save_offer` в архив. Единственный жёсткий гейт — `move_to_production`. Привязка карточки 1С — PATCH в карточке архива. Схему БД не меняем.

## Architecture Decisions

- **Один идентификатор правды** — `KP_offers.counterparty_id`. Нет отдельной таблицы «черновых клиентов».
- **Снапшот** (`customer_name`, `customer_inn`, `customer_kpp`): без id — то, что ввёл менеджер; с id — только справочник (`require_active_client`).
- **Гейт на сервисе, не только в UI** — `ArchiveService.move_to_production` + тот же метод на `OffersService.create_offer(save_mode="work")`.
- **Bind — архивный PATCH**, по образцу `PATCH /discount`, не через черновик мастера.
- **Autocomplete без create:** вынуть `NewCounterpartyDialog` из визарда; тот же поиск переиспользовать в drawer архива.
- **Файлы на bind не трогаем** (D9).

## Dependency graph

```
save_offer / create_offer без id
    │
    ├── move_to_production требует id
    │
    └── PATCH bind (нужен, чтобы открыть производство)
            │
            ├── FE мастер: свободное имя, без «Добавить»
            └── FE архив: бейдж + bind + disabled производство
```

## Task List

### Phase 1 — Контракт сохранения и гейт

- [x] **CBP-001: `save_offer` без карточки 1С**
  - **Description:** В `commercial_draft_lifecycle.save_offer` при создании КП: если `metadata.counterparty_id` нет — писать `customer_name` из `client_name` (strip, непустой), `counterparty_id/inn/kpp = None`. Если id есть — поведение как сейчас (`require_active_client` + снапшот из карточки). Пустое имя → ошибка.
  - **Acceptance:**
    - [x] Черновик с `client_name` и без id сохраняется в архив
    - [x] С валидным id имя в КП = имя справочника, не формы
    - [x] Поставщик / inactive / несуществующий id по-прежнему 400
  - **Verify:** `pytest tests/test_offer_counterparty_validation.py tests/test_kp_counterparty_snapshot.py -q` (+ кейс без id)
  - **Dependencies:** —
  - **Files:** `app/services/commercial_draft_lifecycle.py`, `tests/test_offer_counterparty_validation.py`, `tests/test_kp_counterparty_snapshot.py`
  - **Scope:** M

- [x] **CBP-002: `CreateOfferRequest` / `OffersService.create_offer`**
  - **Description:** `counterparty_id: int | None = None`. Нет id → архив с `payload.customer_name`. `save_mode="work"` без активного клиента → `CounterpartyValidationError`. С id — снапшот из справочника как сейчас.
  - **Acceptance:**
    - [x] Schema принимает payload без id
    - [x] `work` без id падает
    - [x] `archive` без id пишет имя формы
  - **Verify:** `pytest tests/test_offer_counterparty_validation.py tests/test_offers_service.py -q`
  - **Dependencies:** —
  - **Files:** `app/schemas/offers.py`, `app/services/offers_service.py`, `tests/test_offer_counterparty_validation.py`, `tests/test_offers_service.py`
  - **Scope:** M

- [x] **CBP-003: Гейт `move_to_production`**
  - **Description:** Перед promise-gate вызвать `require_active_client(raw.counterparty_id)`. Сообщение по-русски. Статус не менять при ошибке.
  - **Acceptance:**
    - [x] КП в архиве без id → 400, остаётся «в архиве»
    - [x] С активным клиентом → путь как сейчас (promise / ёмкость не ломаем)
  - **Verify:** `pytest tests/test_archive_service.py tests/test_archive_endpoints.py -q`
  - **Dependencies:** CBP-001
  - **Files:** `app/services/archive_service.py`, `tests/test_archive_service.py`, `tests/test_archive_endpoints.py`
  - **Scope:** S

### Checkpoint: Phase 1

- [x] Можно сохранить КП без 1С и нельзя перевести его в производство через API
- [x] Путь «клиент уже в справочнике» зелёный по старым тестам

### Phase 2 — Привязка в архиве (API)

- [x] **CBP-004: PATCH counterparty + поле в выдаче**
  - **Description:** `PATCH /api/v1/commercial/archive/{kp_id}/counterparty` body `{ "counterparty_id": int }`. Write access, только статус «в архиве», `require_active_client`, UPDATE снапшота. В `ArchiveOfferDetails` и `ArchiveOfferListItem` добавить `counterparty_id: int | None`. Репозиторий: узкий `update_offer_counterparty(...)`.
  - **Acceptance:**
    - [x] Bind клиента → имя/ИНН/КПП из карточки, ответ details
    - [x] Supplier / inactive / чужой КП / не архив → ошибка
    - [x] Повторный bind на другую карточку в архиве — ок
    - [x] Список и details отдают id (null, если не привязан)
  - **Verify:** `pytest tests/test_archive_counterparty_bind.py tests/test_archive_endpoints.py -q`
  - **Dependencies:** CBP-001
  - **Files:** `app/schemas/archive.py`, `app/api/v1/endpoints/archive.py`, `app/services/archive_service.py`, `app/repositories/kp_repository.py`, `core/kp/offers_write.py` (узкий UPDATE), `tests/test_archive_counterparty_bind.py`
  - **Scope:** M (если >5 файлов — сначала UPDATE в `offers_write` + repo, затем endpoint)

### Checkpoint: Phase 2

- [x] curl/pytest: архив без id → PATCH → move-to-production 200

### Phase 3 — Мастер КП (UI)

- [x] **CBP-005: Шаг «Клиент» принимает свободное имя**
  - **Description:** `clientConditionsSchema`: `clientName` min 1, `counterpartyId` optional (null/0). Submit с текстом без id. Подсказка: без карточки 1С можно в архив, в производство — нет. Кнопка «Рассчитать» при непустом имени и менеджере.
  - **Acceptance:**
    - [x] Submit `{ clientName: "ИП Петров", counterpartyId: null }`
    - [x] Выбранная карточка по-прежнему шлёт id + реквизиты
    - [x] Пустое имя — нельзя
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-offer/components/steps/ClientConditionsStep.test.tsx src/features/commercial-offer/schemas/`
  - **Dependencies:** CBP-001 (контракт metadata)
  - **Files:** `frontend/src/features/commercial-offer/schemas/commercialOffer.ts`, `frontend/src/features/commercial-offer/components/steps/ClientConditionsStep.tsx`, `.../ClientConditionsStep.test.tsx`, при необходимости `CommercialOfferWizard.tsx` / `useCommercialOfferWizard.ts`
  - **Scope:** M

- [x] **CBP-006: Убрать создание контрагента из автокомплита**
  - **Description:** `CounterpartyAutocomplete` без `NewCounterpartyDialog` и без «Добавить контрагента». Пустой поиск — «Не найдено среди клиентов»; можно оставить имя. Компонент диалога и `createCounterparty` в UI мастера не использовать (файлы диалога можно оставить, если не orphan-тест — тогда удалить диалог + тесты + неэкспортируемый create, **если** больше нигде не зовётся).
  - **Acceptance:**
    - [x] В UI мастера нет «Добавить контрагента»
    - [x] `createCounterparty` не вызывается из autocomplete
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-offer/components/CounterpartyAutocomplete.test.tsx`
  - **Dependencies:** CBP-005
  - **Files:** `CounterpartyAutocomplete.tsx`, `CounterpartyAutocomplete.test.tsx`, опционально удаление `NewCounterpartyDialog.tsx` + `.test.tsx`
  - **Scope:** S

### Checkpoint: Phase 3

- [x] Ручной сценарий: нет в списке → имя → расчёт → «В архив»

### Phase 4 — Архив (UI)

- [x] **CBP-007: Бейдж и гейт кнопки**
  - **Description:** Список: если `counterparty_id == null` — бейдж «нет 1С». Drawer: тот же бейдж; «В производство» disabled + title «Сначала занесите контрагента из 1С» (простые типы по-прежнему «скоро», если так сейчас).
  - **Acceptance:**
    - [x] Бейдж только при отсутствии id
    - [x] С id кнопка как сейчас (с учётом simple-product)
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx src/features/commercial-archive/components/ArchiveOfferList.tsx`
  - **Dependencies:** CBP-004
  - **Files:** `frontend/src/features/commercial-archive/types/archive.ts`, `ArchiveOfferList.tsx`, `OfferDetailsDrawer.tsx`, тесты
  - **Scope:** M

- [x] **CBP-008: «Занести контрагента» в drawer**
  - **Description:** Если статус «в архиве» и (нет id **или** хотим сменить — достаточно поля, пока архив): `CounterpartyAutocomplete` + подтверждение. Вызов нового `archiveApi.bindCounterparty(kpId, id)`, invalidate details/list. После успеха — реквизиты как сейчас (ИНН/КПП).
  - **Acceptance:**
    - [x] Выбор карточки → PATCH → имя и ИНН в drawer
    - [x] Производство разблокируется без перезагрузки страницы
    - [x] Нет диалога создания контрагента
  - **Verify:** `OfferDetailsDrawer.test.tsx` (mock PATCH)
  - **Dependencies:** CBP-004, CBP-006 (autocomplete без create)
  - **Files:** `archiveApi.ts`, `useArchiveQueries` (если есть), `OfferDetailsDrawer.tsx`, тесты
  - **Scope:** M

### Checkpoint: Complete

- [x] Все CBP-001…008 закрыты
- [x] Команды из спеки зелёные
- [ ] Сценарий US-1…US-6 вручную: новое имя → архив → bind → производство (браузер не проверялся)
- [x] Регресс: клиент из справочника с шага 2; append; resume
- [x] Реализация выполнена по согласованному плану

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| PDF после bind со старым именем | Med | D9 явно; в drawer одна строка «файлы обновятся при следующем сохранении» |
| `POST /offers` + `work` без id | High | CBP-002: work всегда `require_active_client` |
| Грязный справочник через старый POST | Low | UI create убран; API не трогаем без отдельного решения |
| Список архива без `counterparty_id` в маппинге | Med | Поле уже в SQL list grouped — прокинуть в schema |
| Simple KP «В производство / скоро» | Low | Не включать кнопку, гейт на бэке всё равно |

## Parallelization

- CBP-001 ∥ CBP-002  
- CBP-005 можно параллельно с CBP-003 после контракта metadata  
- CBP-007 после CBP-004; CBP-008 после CBP-004 и CBP-006  

## Open Questions

Нет. Отклонения от D1–D11 — только правкой спеки.
