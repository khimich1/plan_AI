# План: «Занести контрагента» по карточке договора

**Дата:** 2026-10-27  
**Статус:** PLAN  
**Спека:** `ai_docs/specs/archive-bind-contract-cards.md`  
**Не это:** мастер нового КП, переименование карточек, серверная база, каталог обмена с 1С.

## Overview

Окно архива берёт подсказки из карточек договоров и заносит КП по id выбранной карточки. Если у карточки ещё нет клиента, а ИНН один, сервис записывает этого клиента и на КП, и на карточку. Справочник целиком и ручной ввод кода 1С из окна уходят. Старые `PATCH` и `POST /counterparty` остаются, но это окно их не вызывает.

## Architecture Decisions

- Новый маршрут `PATCH /api/v1/commercial/archive/{kp_id}/contract-card` с телом `{ "card_id": <id карточки> }`. Существующий `PATCH /counterparty` по id справочника не перегружать: иначе старый вызов и новый легко перепутать.
- Поиск не копировать. Окно включает уже готовый `searchSource="contract-cards"`.
- Связь пишет `ArchiveService`, одной транзакцией на КП и карточку. Форма только передаёт id.
- Выбранная карточка важнее правила «самый поздний договор». Автовыбор позднего договора остаётся у кнопки «Отправить в 1С» и здесь не вызывается.
- Если у клиента уже есть другая действующая карточка, новую связь не переносить. КП получает уже связанного клиента.
- Имя на КП берётся из справочника, как в текущем `bind_counterparty`. `imported_name` не менять.
- Создание клиента из этого окна убрать только из интерфейса. `POST /counterparty` не удалять.

## Dependency graph

```
BIND-001 тест привязки по карточке
        │
        └── BIND-002 сервис
                    │
                    └── BIND-003 маршрут
                                │
                                └── BIND-004 окно архива
```

BIND-004 не начинается, пока маршрут отвечает на временной базе. Мастер КП в этой цепочке нет.

## Task List

- [ ] **BIND-001: Тест привязки КП к выбранной карточке**
  - **Description:** Временная база, не `plita.db`. Карточка с живым клиентом заносит на КП его id, имя, ИНН и КПП и не создаёт вторую строку справочника. Карточка без связи и с одним активным ИНН пишет того же клиента на КП и в `supply_contract.counterparty_id` этой карточки. Два активных клиента на ИНН и карточка без ИНН и без связи дают отказ, число строк не меняется. Если у клиента уже связана другая карточка, выбранная остаётся с пустым `counterparty_id`, а КП получает существующего клиента. Более поздняя карточка того же ИНН не перехватывает связь. Повторный вызов не создаёт вторую связь.
  - **Acceptance:** pytest красный, потому что сервиса ещё нет, а не из-за фикстуры. В тесте нет запроса к `GET /counterparties/search`.
  - **Verify:** `.venv/bin/python -m pytest tests/test_archive_bind_contract_card.py -q`
  - **Files:** `tests/test_archive_bind_contract_card.py`

- [ ] **BIND-002: Сервис занесения по id карточки**
  - **Description:** `ArchiveService.bind_contract_card` читает карточку, находит клиента по B5 и пишет КП по B7. Связь на выбранную карточку ставит только когда у клиента нет другой действующей. Чужую связь не снимает. Справочник не пополняет. КП не в архиве получает тот же отказ, что текущая привязка.
  - **Acceptance:** BIND-001 зелёный. `imported_name` и номер карточки в тесте не меняются.
  - **Verify:** `.venv/bin/python -m pytest tests/test_archive_bind_contract_card.py -q`
  - **Files:** `app/services/archive_service.py`

- [ ] **BIND-003: Маршрут карточки**
  - **Description:** `PATCH /api/v1/commercial/archive/{kp_id}/contract-card`, тело `card_id`. Роли те же, что у текущего `PATCH /counterparty`. Неизвестный id карточки — отказ, без записи. Старые PATCH и POST с кодом 1С отвечают как раньше.
  - **Acceptance:** вызов с id карточки доходит до сервиса. Вызов без `card_id` — 422. Старые тесты `test_bind_counterparty_ok` и `test_create_and_bind_counterparty_ok` зелёные.
  - **Verify:** `.venv/bin/python -m pytest tests/test_archive_bind_contract_card.py tests/test_archive_endpoints.py -q`
  - **Files:** `app/api/v1/endpoints/archive.py`, `app/schemas/archive.py`, `tests/test_archive_endpoints.py`

- [ ] **BIND-004: Окно архива**
  - **Description:** Блок «Занести контрагента» ставит `searchSource="contract-cards"` и placeholder «Найдите карточку договора». Кнопка активна только при выбранной карточке и шлёт `card_id`. Поля кода, ИНН и КПП не рисуются. `createAndBindCounterparty` из этого окна не вызывается. Подсказка строки — номер договора, её уже рисует режим карточек.
  - **Acceptance:** тест, который требовал не переключать окно на карточки, заменён обратным. Запрос «ОЛИМП» в тесте окна не вызывает `searchCounterparties`. Ошибка сервиса показывается в уже существующем `Alert`.
  - **Verify:** `cd frontend && npm run test -- src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx src/features/commercial-archive/api/archiveApi.test.ts`
  - **Files:** `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`, `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`, `frontend/src/features/commercial-archive/api/archiveApi.ts`, `frontend/src/features/commercial-archive/api/archiveApi.test.ts`

## Checkpoint

- [ ] Четыре задачи зелёные на временной базе
- [ ] Живая `plita.db` не открывалась
- [ ] Мастер КП и `GET /counterparties/search` не менялись
- [ ] В окне нельзя ввести код 1С без карточки

## Risks

- На сервере у большинства карточек `counterparty_id` пустой. Окно начнёт их показывать и связывать по ИНН только в момент занесения. Само по себе открытие списка базу не пишет.
- Две карточки одного клиента: номер в 1С будет у той, которую связали первой. Вторая кнопка клиента на КП ставит, но номер не переносит. Это сказано в B6, чтобы не выглядело как потерянный договор.
- Старый `POST /counterparty` останется в API. Если его оставить вызванным из окна, баг вернётся. BIND-004 это закрывает тестом.
