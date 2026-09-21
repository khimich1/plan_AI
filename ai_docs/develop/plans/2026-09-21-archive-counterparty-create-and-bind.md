# План: Архив — занести контрагента из 1С и привязать к КП

**Дата:** 2026-09-21  
**Статус:** IMPLEMENT  
**Идея:** `ai_docs/ideas/archive-counterparty-create-and-bind.md`  
**Спека:** `ai_docs/specs/archive-counterparty-create-and-bind.md`

## Overview

В карточке архива оставить поиск существующего клиента (PATCH как сейчас). Если карточки нет — имя из combobox + код 1С (+ опционально ИНН/КПП) одним `POST` создают строку в `counterparties` и привязывают КП. Схему БД не меняем. Мастер КП не трогаем. GUID и живой 1С — не в этом плане.

## Architecture Decisions

- **Два глагола, один путь URL:** `PATCH …/counterparty` — bind по id (без изменений контракта). `POST …/counterparty` — create-or-resolve по `code_1c` и bind. Фронт не вызывает `POST /api/v1/counterparties`.
- **Не `CounterpartiesService.create`:** он кидает `DuplicateCodeError` (409). Новый `resolve_active_client_for_bind`: нашёл код → `require_active_client`; нет → `insert(source="manual", is_client=True)`.
- **Порядок вместо shared-connection:** сначала те же гейты, что у bind (есть КП, write access, статус «в архиве»), потом resolve/insert, потом `update_counterparty` снапшотом с карточки. После прошедших гейтов insert без bind маловероятен; отдельный слой транзакции на два репозитория не заводим.
- **Снапшот только из справочника**, не из свободного имени КП (как PATCH).
- **UI:** combobox = имя (префилл `customer_name`, `onTextChange`). Поля код/ИНН/КПП видны, пока ничего не выбрано из списка. Одна кнопка: есть `bindSelected` → PATCH; иначе имя+код → POST.
- **ИНН необязателен.** Предупреждение о чужом ИНН в архиве не показываем.

## Dependency graph

```
resolve_active_client_for_bind
    │
    └── ArchiveService.create_and_bind_counterparty
            │
            └── POST /archive/{kp_id}/counterparty
                    │
                    ├── archiveApi + mutation
                    └── OfferDetailsDrawer (поля + ветка кнопки)
```

PATCH bind, бейдж «нет 1С», гейт производства — уже есть (CBP-004/007/008). Не переписывать.

## Task List

### Phase 1 — Справочник: resolve-or-create

- [x] **ACC-001: `resolve_active_client_for_bind`**
  - **Description:** В `CounterpartiesService` метод: trim имени и `code_1c`; пустые → `ValueError` (как create). Есть `code_1c` → `require_active_client` существующего id (поставщик/inactive — те же `CounterpartyValidationError`). Нет строки → `repo.insert` с `source="manual"`, `is_client=True`, без GUID. Не использовать `create()` (нет 409). Чужой ИНН не блокирует и warning наружу не отдаём.
  - **Acceptance:**
    - [x] Новый код → одна строка, searchable, `source=manual`
    - [x] Повтор кода активного клиента → тот же id, count не растёт, имя в справочнике **не** перезаписываем
    - [x] Повтор кода поставщика / inactive → `CounterpartyValidationError`
    - [x] Пустые имя или код → `ValueError`
  - **Verify:** `pytest tests/test_counterparties_service.py -q`
  - **Dependencies:** —
  - **Files:** `app/services/counterparties_service.py`, `tests/test_counterparties_service.py`
  - **Scope:** S

### Checkpoint: Phase 1

- [x] Resolve/create покрыт без архива и без HTTP
- [x] Старый `create` + 409 на `POST /counterparties` без регресса

### Phase 2 — Архив API

- [x] **ACC-002: `ArchiveService.create_and_bind_counterparty`**
  - **Description:** Гейты как у `bind_counterparty` (404 / write / «в архиве»). Затем `resolve_active_client_for_bind`, `update_counterparty` снапшотом из row, `get_details`. `ValueError` и `CounterpartyValidationError` → `ArchiveValidationError` (400 на HTTP). Перепривязка в архиве разрешена.
  - **Acceptance:**
    - [x] Нет карточки + имя/код → новая строка, КП с id и снапшотом, status «в архиве»
    - [x] Код существующего клиента → без insert, bind к нему (имя КП = справочник, не payload.name, если они разошлись)
    - [x] Код поставщика/inactive → ошибка, `counterparty_id` КП null
    - [x] Статус не «в архиве» → ошибка, справочник не менять (гейты до insert)
    - [x] После успешного create+bind `move_to_production` проходит `require_active_client`
  - **Verify:** `pytest tests/test_archive_counterparty_bind.py -q`
  - **Dependencies:** ACC-001
  - **Files:** `app/services/archive_service.py`, `tests/test_archive_counterparty_bind.py`
  - **Scope:** M

- [x] **ACC-003: POST endpoint + схема**
  - **Description:** `CreateAndBindCounterpartyRequest` в `app/schemas/archive.py` (лимиты как `CounterpartyCreateRequest`). `POST /{kp_id}/counterparty` рядом с существующим PATCH: admin/manager, те же 404/400 хелперы. Пустые поля — 422 Pydantic. PATCH не менять.
  - **Acceptance:**
    - [x] POST валидного тела зовёт `create_and_bind_counterparty` и отдаёт details
    - [x] PATCH `{counterparty_id}` как в `test_bind_counterparty_ok`
    - [x] Нет auth → 401
  - **Verify:** `pytest tests/test_archive_endpoints.py -q`
  - **Dependencies:** ACC-002
  - **Files:** `app/schemas/archive.py`, `app/api/v1/endpoints/archive.py`, `tests/test_archive_endpoints.py`
  - **Scope:** S

### Checkpoint: Phase 2

- [x] pytest: архив без id → POST create+bind → `move_to_production` 200 (сервисный тест)
- [x] Дубль `code_1c` не плодит строку
- [x] Мастер и `POST /counterparties` не затронуты

### Phase 3 — UI архива

- [x] **ACC-004: клиент API + мутация**
  - **Description:** `archiveApi.createAndBindCounterparty(kpId, { name, code_1c, inn, kpp })` → POST JSON. Хук по образцу `useBindCounterpartyMutation` (invalidate list/detail). Не звать `createCounterparty` из `counterpartiesApi.ts`.
  - **Acceptance:**
    - [x] Тест: POST `/api/v1/commercial/archive/42/counterparty` с name/code_1c
    - [x] PATCH-хелпер без изменений
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-archive/api/archiveApi.test.ts`
  - **Dependencies:** ACC-003 (контракт)
  - **Files:** `frontend/src/features/commercial-archive/api/archiveApi.ts`, `archiveApi.test.ts`, `hooks/useArchiveQueries.ts`
  - **Scope:** S

- [x] **ACC-005: форма в `OfferDetailsDrawer`**
  - **Description:** Пока статус «в архиве»: combobox + `onTextChange` (имя, префилл `customer_name`). Если нет `bindSelected` — поля код 1С / ИНН / КПП. Кнопка: выбран клиент → PATCH; иначе непустые имя и код → POST. Пока ни то ни другое — disabled. Ошибки обеих мутаций в Alert. После успеха — как сейчас (invalidate, бейдж, производство). Мок autocomplete в тесте: оставить pick; для create — `data-testid` на код и кнопку без pick.
  - **Acceptance:**
    - [x] Без выбора и без кода кнопка disabled
    - [x] Код без выбора → POST, не PATCH
    - [x] Pick из списка → PATCH как сейчас, POST не зовётся
    - [x] Нет диалога «Добавить контрагента» / вызова `createCounterparty`
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
  - **Dependencies:** ACC-004
  - **Files:** `OfferDetailsDrawer.tsx`, `OfferDetailsDrawer.test.tsx`
  - **Scope:** M

### Checkpoint: Complete

- [x] ACC-001…005 закрыты
- [x] Команды из спеки зелёные
- [x] Регресс: PATCH bind, мастер без create, `move_to_production` без id → 400
- [ ] Вручную (после кода, не в этом шаге): «ооо бармалей» + код 1С → бейдж снят → «В производство»
- [x] Выполнение начнётся только после явной команды

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Insert прошёл, UPDATE КП нет | Low | Гейты до insert; сиротская карточка всё равно валидный клиент для следующего bind |
| POST и PATCH на одном path путают клиент | Low | Разные методы; фронт ветвит явно; тесты обоих |
| Имя в combobox не уходит в POST | Med | Локальный `bindName` + `onTextChange`, префилл `customer_name` |
| Мок autocomplete ломает старый тест pick | Med | Поля кода — в drawer, не в моке; pick-тест не заполняет код |
| Опечатка `code_1c` | Med (продукт) | Спека: якорь для GUID; не чиним валидатором «похожести» |
| Реюз `create()` → 409 на дубле | High | Только `resolve_active_client_for_bind` |

## Parallelization

- После ACC-001: ACC-002 строго следом.
- ACC-004 можно писать по контракту спеки параллельно с ACC-002/003, но сливать после стабильного POST.
- ACC-005 после ACC-004.
- Мастер КП / импорт 1С — не параллелить сюда, не в скоупе.

## Open Questions

Нет. ИНН необязателен (ревью спеки). Отклонения от D1–D11 — только правкой спеки.
