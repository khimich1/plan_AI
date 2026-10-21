# План: Предложение править найденную карточку договора

**Дата:** 2026-10-21  
**Статус:** IMPLEMENT. Три задачи сделаны. Код написан, тесты зелёные.  
**Спека:** `ai_docs/specs/contract-card-duplicate-offer.md`  
**Не это:** мастер КП, блок «Занести контрагента», импорт реестра, подстрочный поиск списка, разбор файла на уже открытой карточке.  
**Не править:** `ai_docs/specs/archive-counterparty-cards.md`. Живую `plita.db` не открывать и не менять. Приложение не запускать. Коммит только по отдельной просьбе.

## Overview

Форма «Новая карточка» перед `POST /contract-cards` спрашивает `GET /contract-cards/duplicates`. Совпал ИНН целиком по цифрам или отображаемое имя целиком — диалог «Такая карточка уже есть». «Редактировать» открывает найденную карточку как клик по списку, без подстановки бланка из файла. «Всё равно создать» отправляет то же тело, что форма или разбор отправили бы сегодня. Несколько совпадений — список по `id`. Файл сначала разбирается; при отказе разбора поиск не идёт; при совпадении строка не пишется, пока не нажата «Всё равно создать».

Прямой `POST /contract-cards` и `create_card` по-прежнему вставляют вторую карточку. Справочник `counterparties` поиск дубля не читает. Пустой ИНН и пустое имя никого не находят, в том числе карточку с пустым ИНН или пустым отображаемым именем.

## Architecture Decisions

- Принятый пункт 1: обрыв только на форме и на броске файла в зону этой формы. У `POST` нет флага подтверждения. Импорт реестра и прямой вызов не меняются.
- Принятый пункт 2: имя сравнивается с `_card_display_name` (`imported_name`, иначе `full_name`, иначе `short_name`). Не с остальными полями и не подстрокой. Ключ — свёртка пробелов и `casefold`. ИНН — только цифры, целиком, с ведущими нулями. Пустой ключ не равен другому пустому.
- Принятый пункт 3: «Редактировать» вызывает то же открытие, что клик по строке списка. `intake` разобранного файла туда не передаётся и в базу не пишется.
- Клиент дубль не считает. Диалог показывает `items` ответа GET. Порядок — порядок сервиса, тот же `list_for_search` по `id`.
- Маршрут `GET /contract-cards/duplicates` регистрируется раньше `/{card_id}`. Роли те же: `admin` и `manager`.
- «Всё равно создать» одна на диалог и шлёт отложенное тело в существующий `POST`. С файла номера в теле нет. Успех с файла открывает карточку с тем же `intake`, что сейчас после броска. Ошибка занятого номера остаётся на форме, карточка не открывается.
- Закрытие диалога без двух кнопок действия строку не создаёт. Введённые имя, ИНН и номер остаются. Разобранный файл после закрытия пропадает.

## Dependency graph

```
DUP-001 pytest поиска (красный)
        │
        └── DUP-002 find_duplicate_cards и GET
                    │
                    └── DUP-003 диалог, форма, vitest
```

DUP-001 не зелёный, пока нет метода. DUP-003 не переводит мастер КП и «Занести контрагента» на новый поиск.

## Task List

### Phase 1 — Поиск

- [x] **DUP-001: Падающий pytest совпадения** `(type: test, dependsOn: —)`
  - **Description:** Временная база, не `plita.db`. Поиск по ИНН `5024243403` находит карточку; `5024` при хранимом `5024243403` пуст. Хранимое `5024-243-403` и запрос `5024 243403` — одна карточка. `05024243403` и `5024243403` различны. Пустой ИНН, пробелы и «нет» не возвращают карточки без ИНН и не возвращают чужие. «ооо  ромашка» находит отображаемое «ООО Ромашка». «Строй» не находит «Строймонтаж» и «ООО Строй». «Ромашка» не находит карточку, у которой в колонке «Имя» «Завод», даже если `short_name` или `full_name` — «Ромашка». Пустое имя даёт пустой список и не цепляет карточку с пустым отображаемым именем. Карточка, совпавшая по имени и ИНН, одна. Две разные идут по возрастанию `id`. Строка только в `counterparties` не входит, число строк справочника не меняется. Статус «отмена не будем работать» входит. Вызов поиска не увеличивает `supply_contract`. GET делает то же и не пишет строку. Прямой POST второй карточки с тем же ИНН по-прежнему создаёт строку.
  - **Acceptance:**
    - [x] Сначала тесты падали из-за отсутствия поиска, теперь они зелёные
    - [x] Утверждения совпадают со спекой: ИНН целиком, имя только отображаемое, пустые не ищут, справочник не участвует
  - **Verify:** `/home/username/Code/plan_web/.venv/bin/python -m pytest tests/test_archive_contract_card_duplicates.py -q`
  - **Dependencies:** —
  - **Files:** `tests/test_archive_contract_card_duplicates.py`
  - **Scope:** M

- [x] **DUP-002: Поиск в сервисе и GET** `(type: feat-be, dependsOn: DUP-001)`
  - **Description:** `find_duplicate_cards` читает только `list_for_search`, строку не вставляет, `counterparties` не открывает. Ключи — `duplicate_name_key` и `duplicate_inn_key` из спеки. Пустые `name` и `inn` возвращают `{items: [], count: 0}` после проверки роли. `create_card` и `POST /contract-cards` не получают запрета дубля. `search_cards` не меняется. Маршрут стоит рядом с `parse` и `import`, раньше `/{card_id}`.
  - **Acceptance:**
    - [x] DUP-001 зелёный
    - [x] Подстрочный поиск списка по-прежнему зелёный
    - [x] Разбор файла карточки по-прежнему не пишет строку
  - **Verify:** `/home/username/Code/plan_web/.venv/bin/python -m pytest tests/test_archive_contract_card_duplicates.py tests/test_archive_contract_cards.py::test_search_finds_three_cards_by_name_inn_and_number_fragment tests/test_archive_contract_card_document_http.py::test_parse_contract_card_upload_returns_fields_and_writes_no_row tests/test_archive_contract_card_document_http.py::test_parse_contract_card_upload_rejects_a_non_card_and_writes_no_row -q`
  - **Dependencies:** DUP-001
  - **Files:** `app/services/supply_contract_service.py`, `app/api/v1/endpoints/archive.py`, `tests/test_archive_contract_card_duplicates.py`
  - **Scope:** M

### Checkpoint: Phase 1

- [x] Спека принята до старта DUP-001, три пункта Open Questions закрыты
- [x] `create_card` и прямой POST по-прежнему вставляют вторую карточку с тем же ИНН
- [x] Дальше не идти, если поиск читает `counterparties` или пишет `supply_contract`

### Phase 2 — Форма

- [x] **DUP-003: Диалог на форме и на броске файла** `(type: feat-fe, dependsOn: DUP-002)`
  - **Description:** `findDuplicateContractCards` зовёт GET. Форма и зона файла новой карточки сначала узнают совпадения. Пустой ответ — одно создание с тем же телом, что сегодня, диалога нет. Совпадения — `DuplicateContractCardDialog`: имя, ИНН (пустой как «—»), номер, «Редактировать» у строки, одна «Всё равно создать». «Редактировать» открывает блок «Карточка договора» этого номера и не вызывает `createContractCard`. С файла бланк при этом не заполняется. «Всё равно создать» вызывает создание один раз; с файла без номера и с бланком разбора, `saveContractRequisites` не вызывается. Закрытие диалога строку не создаёт, повторная отправка снова показывает диалог. Отказ «это не карточка контрагента» и пустое имя поиск не вызывают. Мастер КП и «Занести контрагента» этот диалог не получают.
  - **Acceptance:**
    - [x] vitest: диалог, «Редактировать» без create, «Всё равно создать» вызывает create, бросок файла при совпадении не создаёт строку
    - [x] Существующие тесты формы при пустом поиске по-прежнему ждут один `createContractCard` с тем же телом
    - [x] vitest мастера КП и «Занести контрагента» зелёные без новых сценариев дубля
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-archive/api/archiveApi.test.ts src/features/commercial-archive/components/ContractCardList.test.tsx src/features/commercial-offer/components/CounterpartyAutocomplete.test.tsx src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
  - **Dependencies:** DUP-002
  - **Files:** `frontend/src/features/commercial-archive/api/archiveApi.ts`, `frontend/src/features/commercial-archive/api/archiveApi.test.ts`, `frontend/src/features/commercial-archive/components/DuplicateContractCardDialog.tsx`, `frontend/src/features/commercial-archive/components/ContractCardList.tsx`, `frontend/src/features/commercial-archive/components/ContractCardList.test.tsx`
  - **Scope:** M

### Checkpoint: конец

- [x] Команды DUP-002 и DUP-003 зелёные
- [x] Живая `plita.db` не менялась, приложение не запускалось
- [x] В спеке фаза `SPECIFY ✅ → PLAN ✅ → IMPLEMENT`
- [x] Коммит только по отдельной просьбе

## Boundaries

- Always: сначала разбор файла, и только успешный разбор с именем идёт в поиск; потом создание. Пустой ключ не сравнивать. Тесты на временной базе.
- Ask first: уникальный индекс, отказ POST без кнопки, сравнение с `full_name` и `short_name` помимо отображаемого имени, подстановка бланка по «Редактировать».
- Never: править `archive-counterparty-cards.md`, открывать `plita.db`, менять подстрочный поиск списка, мастер КП и «Занести контрагента», откатывать имена 1С в `imported_name`, флаг «не полные», удаление, дроп файла и поиск мастера КП.
