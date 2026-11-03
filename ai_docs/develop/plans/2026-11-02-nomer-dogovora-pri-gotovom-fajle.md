# План: номер договора в момент готового файла

**Дата:** 2026-11-02
**Статус:** PLAN, на ревью. Кода ещё нет.
**Спека:** `ai_docs/specs/nomer-dogovora-pri-gotovom-fajle.md`
**Не это:** удаление `1032/11/26`, сверка Excel, пустой номер в схеме, черновик бланка в базе, отправка в ЭДО, запись договора в 1С.

## Overview

`POST /commercial/archive/directory/{id}/contract` без тела сейчас вставляет строку и следующий номер. Пустая карточка со статусом «нет» уже открывает счёт.

После этой работы пустой вызов ничего не пишет. Старый номер привязывается выбором карточки или текстом из Excel и 1С, seq не растёт. Следующий номер выдаётся один раз, в той же вставке, что и первый готовый бланк нового клиента. До этой вставки строки нет, и счёт остаётся закрыт. Повторное сохранение реквизитов номер не меняет.

Поиск клиента, «договор есть / нет» и «Привязать № …» для одной свободной карточки остаются. Живую `plita.db` тесты не открывают.

## Architecture Decisions

- Пустой `POST .../contract` отвечает отказом и не вызывает `_allocate_number`. Тело — `{ "number": "…" }` либо `{ "contract_id": N }`, не оба сразу.
- Известный номер ищется по тексту после `normalize_contract_number`. Свободная карточка с этим номером или переданный `contract_id` получает `counterparty_id`. Номер строки не переписывается. Чужой владелец — отказ без записи.
- Нового текста в базе ещё нет: вставка идёт отдельным методом, не через `insert_card` и не через `_allocate_number`. Дата — `contract_date_from_number`, иначе пустая. Имя и ИНН берутся из справочника. Статус «нет», реквизиты пустые.
- Карточка для выбора обязана быть свободной или уже принадлежать этому клиенту. Если у неё заполнен ИНН, он совпадает с клиентом. Чужой ИНН не привязывается.
- Список для экрана — свободные неотменённые карточки с ИНН клиента, даже если их несколько. В ответ поиска добавляется `candidates`. `linkable` по-прежнему заполнен только для одной однозначной пары.
- Новый договор: `POST .../contract-blank` с телом реквизитов. Пока `requisites_ready` ложно, вставки нет. Когда бланк готов и действующей строки нет, одна транзакция берёт `next_contract_seq` и пишет номер от календарной даты сервера. Действующая строка уже есть — реквизиты пишутся в неё, номер остаётся.
- Черновик до этого вызова живёт в состоянии формы. `POST /contract-cards/parse` строку не создаёт и остаётся разбором без id.
- Успех привязки старого номера снимает блок счёта. Открытие формы «Собрать договор» блок не снимает.

## Dependency graph

```
NUM-001 привязать вписанный номер, пустой POST ничего не пишет
        │
        ├── NUM-003 поле номера во вкладке
        │
        └── NUM-004 то же поле на КП

NUM-002 список свободных карточек ИНН
        │
        ├── NUM-003
        │
        └── NUM-004

NUM-005 готовый бланк выдаёт номер один раз
        │
        └── NUM-006 форма без строки вызывает этот бланк
```

NUM-001 и NUM-002 можно делать параллельно. NUM-005 можно делать параллельно с NUM-003 и NUM-004. NUM-006 ждёт NUM-005 и кнопки из NUM-003 и NUM-004.

## Task List

### Phase 1 — Старый номер без счётчика

- [ ] **NUM-001: Вписанный номер или отказ** `(type: feat-be)` (⏳ Pending)
  - **Description:** `POST /commercial/archive/directory/{counterparty_id}/contract` для `admin` и `manager`. Без `number` и без `contract_id` — отказ, строка не создаётся, seq не меняется. С номером или id: действующая строка этого клиента возвращается, если номер её; свободная подходящая карточка привязывается; неизвестный текст вставляется как есть; чужой владелец даёт отказ. Повтор того же номера возвращает ту же строку. `insert_card` и `_allocate_number` этот путь не вызывает. Тесты, которые ждут третью строку и `1032/11/26` от пустого POST, переписываются на отказ.
  - **Acceptance criteria:**
    - [ ] Пустое тело не создаёт строку и не увеличивает seq
    - [ ] Вписанный отсутствующий номер сохраняется буквально, дата берётся из номера, реквизиты пустые, seq не растёт
    - [ ] Свободный `contract_id` этого ИНН привязывается, номер карточки тот же
    - [ ] Номер другого клиента не меняет его строку
    - [ ] Повтор возвращает ту же строку
    - [ ] Две свободные карточки сами не выбираются и третья от пустого вызова не создаётся
    - [ ] Пустая привязанная карточка снимает «Нет действующего договора поставки»
  - **Verification:**
    - [ ] `.venv/bin/python -m pytest tests/test_archive_directory_contracts.py tests/test_supply_contract_number.py tests/test_kp_invoice_existing_contract.py -q`
  - **Dependencies:** нет
  - **Files likely touched:**
    - `app/api/v1/endpoints/archive.py`
    - `app/schemas/archive.py`
    - `app/services/supply_contract_service.py`
    - `app/repositories/supply_contract_repository.py`
    - `tests/test_archive_directory_contracts.py`
  - **Estimated scope:** Medium

- [ ] **NUM-002: Список карточек одного ИНН** `(type: feat-be)` (⏳ Pending)
  - **Description:** В строке `GET /commercial/archive/directory` поле `candidates`: свободные неотменённые договоры с ИНН клиента, `{id, number}`. Несколько штук остаются в списке. `linkable` не начинает выбирать одну из них. Карточки другого ИНН и уже занятые другим клиентом в список не входят.
  - **Acceptance criteria:**
    - [ ] Две свободные карточки одного ИНН обе видны в `candidates`, `linkable` пуст
    - [ ] Одна однозначная карточка по-прежнему в `linkable` и в `candidates`
    - [ ] Карточка другого клиента и другой ИНН в `candidates` не попадает
    - [ ] У клиента с действующим договором `candidates` пуст
  - **Verification:**
    - [ ] `.venv/bin/python -m pytest tests/test_archive_directory_contracts.py -q`
  - **Dependencies:** нет
  - **Files likely touched:**
    - `app/schemas/archive.py`
    - `app/services/supply_contract_service.py`
    - `app/repositories/supply_contract_repository.py`
    - `tests/test_archive_directory_contracts.py`
  - **Estimated scope:** Small

### Checkpoint: After NUM-001 and NUM-002

- [ ] Пустой POST на временной базе не создаёт договор
- [ ] Вписанный номер и выбранная карточка привязываются без нового seq
- [ ] Поиск отдаёт все свободные карточки ИНН

### Phase 2 — Экран старого номера

- [ ] **NUM-003: Номер во вкладке «Контрагенты»** `(type: feat-fe)` (⏳ Pending)
  - **Description:** Строка без договора больше не шлёт пустой POST по кнопке «Завести». Одна `linkable` по-прежнему «Привязать № …» и шлёт `contract_id`. Иначе показаны `candidates` и поле номера. Выбор и ввод вызывают NUM-001. Успех оставляет список открытым и показывает номер. Отказ остаётся текстом у строки.
  - **Acceptance criteria:**
    - [ ] «Завести» не вызывает POST без тела
    - [ ] «Привязать № …» шлёт `contract_id` и после успеха показывает этот номер
    - [ ] Выбор из двух карточек шлёт id выбранной
    - [ ] Ввод текста шлёт `number`
    - [ ] Пустой ввод запрос не отправляет
  - **Verification:**
    - [ ] `cd frontend && npx vitest run src/features/commercial-archive/components/ContractCardList.test.tsx`
  - **Dependencies:** NUM-001, NUM-002
  - **Files likely touched:**
    - `frontend/src/features/commercial-archive/components/ContractCardList.tsx`
    - `frontend/src/features/commercial-archive/components/ContractCardList.test.tsx`
    - `frontend/src/features/commercial-archive/api/archiveApi.ts`
    - `frontend/src/features/commercial-archive/hooks/useArchiveQueries.ts`
    - `frontend/src/features/commercial-archive/types/supplyContract.ts`
  - **Estimated scope:** Medium

- [ ] **NUM-004: То же на карточке КП** `(type: feat-fe)` (⏳ Pending)
  - **Description:** «в реестре договора нет» открывает тот же выбор, ящик КП не закрывается. Успешная привязка старого номера снимает блок счёта, как сейчас после записи. Открытие формы нового договора блок не снимает. Бейдж «нет 1С» и «Занести контрагента» не меняются.
  - **Acceptance criteria:**
    - [ ] Кнопка не шлёт пустой POST
    - [ ] Выбор или ввод номера не закрывает ящик КП
    - [ ] После привязки номера «Отправить в 1С» больше не держит «Нет действующего договора поставки»
    - [ ] Бейдж «нет 1С» остаётся у КП без `counterparty_id`
  - **Verification:**
    - [ ] `cd frontend && npx vitest run src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
  - **Dependencies:** NUM-001, NUM-002, NUM-003
  - **Files likely touched:**
    - `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`
    - `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
    - `frontend/src/features/commercial-archive/components/ContractCardList.tsx`
  - **Estimated scope:** Small

### Checkpoint: After NUM-003 and NUM-004

- [ ] Vitest вкладки и ящика КП зелёный
- [ ] Pytest NUM-001 по-прежнему зелёный
- [ ] Ни одна кнопка на этих экранах не создаёт следующий номер

### Phase 3 — Новый номер вместе с бланком

- [ ] **NUM-005: Первый готовый бланк** `(type: feat-be)` (⏳ Pending)
  - **Description:** `POST /commercial/archive/directory/{counterparty_id}/contract-blank`. Неготовые реквизиты — отказ, строка не создаётся. Готовые реквизиты и нет действующего договора — одна вставка: следующий seq, дата сервера, id клиента, поля бланка. Готовые реквизиты и договор уже есть — обновление той же строки, номер прежний. Отменённый номер остаётся занятым: следующий бланк получает seq больше него. Клиент без строки по-прежнему не проходит в счёт.
  - **Acceptance criteria:**
    - [ ] Пустое обязательное поле не создаёт строку и не меняет seq
    - [ ] Первый готовый бланк создаёт одну строку, номер — max+1 и месяц/год даты вставки
    - [ ] Второй готовый бланк сохраняет номер и id
    - [ ] После отмены следующий готовый бланк не занимает отменённый номер
    - [ ] До вставки счёт этого клиента отвечает «Нет действующего договора поставки»
    - [ ] После вставки счёт штампуется новым номером
  - **Verification:**
    - [ ] `.venv/bin/python -m pytest tests/test_archive_directory_contracts.py tests/test_supply_contract_number.py tests/test_kp_invoice_existing_contract.py tests/test_archive_contract_card_documents.py -q`
  - **Dependencies:** нет
  - **Files likely touched:**
    - `app/api/v1/endpoints/archive.py`
    - `app/services/supply_contract_service.py`
    - `app/repositories/supply_contract_repository.py`
    - `tests/test_archive_directory_contracts.py`
    - `tests/test_kp_invoice_existing_contract.py`
  - **Estimated scope:** Medium

- [ ] **NUM-006: Форма нового договора без строки** `(type: feat-fe)` (⏳ Pending)
  - **Description:** «Собрать договор» на вкладке и на КП открывает бланк без id строки. Разбор файла идёт в `POST /contract-cards/parse` и только заполняет поля. «Сохранить реквизиты» при неготовой форме запрос не шлёт, введённое остаётся. Готовая форма вызывает NUM-005 и показывает выданный номер. Повторное сохранение показывает тот же номер. Закрытие формы до сохранения ничего не восстанавливает.
  - **Acceptance criteria:**
    - [ ] До сохранения в базе нет вызова, который создаёт договор
    - [ ] Разбор файла не создаёт строку и подставляет поля в форму
    - [ ] Неготовая форма не вызывает `contract-blank`
    - [ ] Готовая форма один раз получает номер и показывает его
    - [ ] Повтор показывает тот же номер
  - **Verification:**
    - [ ] `cd frontend && npx vitest run src/features/commercial-archive/components/ContractCardDrawer.test.tsx src/features/commercial-archive/components/ContractCardList.test.tsx src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
  - **Dependencies:** NUM-003, NUM-004, NUM-005
  - **Files likely touched:**
    - `frontend/src/features/commercial-archive/components/ContractCardDrawer.tsx`
    - `frontend/src/features/commercial-archive/components/ContractCardDrawer.test.tsx`
    - `frontend/src/features/commercial-archive/components/ContractCardList.tsx`
    - `frontend/src/features/commercial-archive/api/archiveApi.ts`
    - `frontend/src/features/commercial-archive/hooks/useArchiveQueries.ts`
  - **Estimated scope:** Medium

### Checkpoint: Complete

- [ ] Pytest из NUM-005 зелёный
- [ ] Vitest из NUM-006 зелёный
- [ ] Пустой договор кнопкой не создаётся
- [ ] Старый номер привязывается без нового seq
- [ ] Новый номер появляется только с первым готовым бланком

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Старые тесты ждут `1032/11/26` от пустого POST и третью строку при двух карточках | Средний | NUM-001 переписывает эти проверки в том же изменении, что снимает выдачу |
| `insert_card` пишет пустой `counterparty_id` | Средний | Вписанный номер вставляется своим методом сразу с id клиента |
| Ящик КП сейчас сам гасит блок счёта после любого успеха | Средний | Гасить блок только после привязки номера. «Собрать договор» блок не трогает |
| Два готовых бланка одновременно | Низкий | Уникальный действующий договор на клиента. Второй вызов дописывает уже созданную строку и номер не берёт |

## Open Questions

Нет.
