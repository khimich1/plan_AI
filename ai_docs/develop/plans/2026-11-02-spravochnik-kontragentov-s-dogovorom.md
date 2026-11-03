# План: справочник контрагентов с договором

**Дата:** 2026-11-02  
**Статус:** PLAN, на ревью. Кода ещё нет.  
**Спека:** `ai_docs/specs/spravochnik-kontragentov-s-dogovorom.md`  
**Не это:** шаг «Клиент», «Занести контрагента», импорт xlsx, окно выбора из нескольких номеров, список карточек без клиента, поиск по коду и номеру, обязательный бланк до отправки.

## Overview

Вкладка «Контрагенты» сейчас ищет `supply_contract` и создаёт договор с пустым `counterparty_id`. КП при уже выбранном клиенте такой договор не видит и держит «Отправить в 1С» закрытой.

После этой работы вкладка ищет активных клиентов по имени и ИНН. Строка показывает действующий номер или его отсутствие. Одна кнопка записи применяет правило заново: действующий договор возвращается как есть, одна однозначная свободная карточка привязывается, иначе создаётся новый номер с id клиента. Ящик договора после записи не открывается. Та же кнопка вызывается со строки «в реестре договора нет» на КП, ящик КП остаётся.

`GET/POST /contract-cards` остаются для карточки и импорта. Живую `plita.db` тесты не открывают.

## Architecture Decisions

- Новый поиск не расширяет `GET /counterparties/search`. Ответ мастера остаётся коротким. Вкладка ходит в `GET /api/v1/commercial/archive/directory`.
- Запись одна: `POST /api/v1/commercial/archive/directory/{counterparty_id}/contract`. Подпись кнопки («Завести», «Привязать №…», «Открыть») берётся из поиска, а решение ещё раз считает сервис. Так повтор и гонка не создают второй договор.
- Свободная карточка привязывается только через `select_contract_client_links` для этого клиента. Несколько карточек и два кода дают пусто, затем «Завести» создаёт новый номер и старые не обновляет.
- Номер нового договора — `next_contract_seq` по всем занятым номерам и `format_contract_number` от даты вызова. Поля номера в запросе нет.
- Вставка с `counterparty_id` — отдельный метод репозитория. Текущий `insert_card` по-прежнему пишет пустой id для импорта.
- После записи вкладка обновляет строку поиска. `OfferDetailsDrawer` обновляет карточку КП и не переключает раздел архива.

## Dependency graph

```
DIR-001 поиск клиентов с пометкой договора
        │
        └── DIR-003 вкладка показывает строки и открывает готовый договор

DIR-002 запись: вернуть, привязать или завести
        │
        ├── DIR-003 кнопки «Завести» и «Привязать»
        │
        └── DIR-004 строка на КП вызывает ту же запись
```

DIR-001 и DIR-002 можно делать параллельно. DIR-003 ждёт оба. DIR-004 ждёт DIR-002 и не ждёт DIR-003.

## Task List

### Phase 1 — Чтение и запись

- [ ] **DIR-001: Поиск клиентов с пометкой** `(type: feat-be)` (⏳ Pending)
  - **Description:** `GET /commercial/archive/directory?q=` для `admin` и `manager`. Пустой запрос и один символ дают пустой список. Дальше имя и ИНН активных клиентов, менеджеру только `is_client = 1`, не больше 10. В строке id, имя, ИНН, КПП, код, действующий договор `{id, number}` либо `null`, и одна привязываемая карточка `{id, number}` либо `null`. Привязываемая карточка есть только когда `select_contract_client_links` даёт пару именно с этим клиентом. Код 1С в условие поиска не входит.
  - **Acceptance criteria:**
    - [ ] Пустой запрос и один символ возвращают 0 строк
    - [ ] Имя и ИНН находят клиента. Код 1С при том же тексте не находит, если имя и ИНН не совпали
    - [ ] У клиента с действующим договором в строке его номер, привязываемой карточки нет
    - [ ] Одна свободная карточка и один код дают эту карточку в `linkable`
    - [ ] Две свободные карточки и два кода дают `linkable = null`
    - [ ] Менеджер не видит строку с `is_client = 0`
  - **Verification:**
    - [ ] `.venv/bin/python -m pytest tests/test_archive_directory_contracts.py -q`
  - **Dependencies:** нет
  - **Files likely touched:**
    - `app/api/v1/endpoints/archive.py`
    - `app/schemas/archive.py`
    - `app/services/supply_contract_service.py`
    - `tests/test_archive_directory_contracts.py`
  - **Estimated scope:** Medium

- [ ] **DIR-002: Вернуть, привязать или завести** `(type: feat-be)` (⏳ Pending)
  - **Description:** `POST /commercial/archive/directory/{counterparty_id}/contract`. Клиент должен быть активным. Если действующий договор есть, вернуть его и ничего не вставлять. Если `select_contract_client_links` даёт одну карточку этого клиента, записать на неё `counterparty_id` и вернуть её. Иначе создать договор: номер от `next_contract_seq` и даты вызова, имя и ИНН из справочника, статус «нет». Повтор после создания возвращает тот же номер. Чужой договор и лишние свободные карточки не обновляются. Число строк `counterparties` не меняется.
  - **Acceptance criteria:**
    - [ ] Нет договора и нет однозначной карточки: одна новая строка с id клиента и следующим номером
    - [ ] Повтор не увеличивает число договоров и seq
    - [ ] Одна свободная карточка при одном коде привязывается, новой строки нет
    - [ ] Две свободные карточки остаются с пустым id, создаётся третья с id клиента
    - [ ] Два кода на ИНН не получают свободную карточку
    - [ ] Уже существующий действующий договор возвращается без второй строки
    - [ ] Неактивный id даёт отказ, строка договора не пишется
    - [ ] После создания пустой бланк больше не даёт «Нет действующего договора поставки» для КП этого клиента
  - **Verification:**
    - [ ] `.venv/bin/python -m pytest tests/test_archive_directory_contracts.py tests/test_supply_contract_number.py tests/test_kp_invoice_existing_contract.py -q`
  - **Dependencies:** нет
  - **Files likely touched:**
    - `app/api/v1/endpoints/archive.py`
    - `app/services/supply_contract_service.py`
    - `app/repositories/supply_contract_repository.py`
    - `tests/test_archive_directory_contracts.py`
  - **Estimated scope:** Medium

## Checkpoint: After DIR-001 and DIR-002

- [ ] Оба pytest зелёные на временной базе
- [ ] `GET/POST /contract-cards` и импорт карточки без id не сломаны
- [ ] Живая `plita.db` не открывалась

### Phase 2 — Экраны

- [ ] **DIR-003: Вкладка «Контрагенты»** `(type: feat-fe, dependsOn: DIR-001, DIR-002)` (⏳ Pending)
  - **Description:** `ContractCardList` ищет `GET .../directory`. Форма имени, ИНН, номера и зона файла с вкладки уходят. Пустой и однобуквенный запрос показывают прежнюю подсказку ввести имя или ИНН. Строка с договором открывает текущий ящик карточки. Строка без договора показывает «Завести» или «Привязать» с номером из `linkable`. После записи список обновляется, ящик карточки не открывается, номер виден в строке.
  - **Acceptance criteria:**
    - [ ] Запрос уходит на `/directory`, не на `/contract-cards`
    - [ ] Создание с вкладки не шлёт имя и номер отдельной формой
    - [ ] «Завести» и «Привязать» вызывают POST и оставляют список на экране
    - [ ] Строка с номером открывает карточку
    - [ ] Зона файла внутри уже открытой карточки остаётся
  - **Verification:**
    - [ ] `cd frontend && npx vitest run src/features/commercial-archive/components/ContractCardList.test.tsx`
  - **Dependencies:** DIR-001, DIR-002
  - **Files likely touched:**
    - `frontend/src/features/commercial-archive/components/ContractCardList.tsx`
    - `frontend/src/features/commercial-archive/api/archiveApi.ts`
    - `frontend/src/features/commercial-archive/hooks/useArchiveQueries.ts`
    - `frontend/src/features/commercial-archive/components/ContractCardList.test.tsx`
  - **Estimated scope:** Medium

- [ ] **DIR-004: Строка на КП** `(type: feat-fe, dependsOn: DIR-002)` (⏳ Pending)
  - **Description:** На архивном КП, где клиент выбран и блок отправки — «Нет действующего договора поставки», текст «в реестре договора нет» становится кнопкой того же POST. Ящик КП не закрывается и раздел архива не меняется. После успеха блок снимается. КП без `counterparty_id` по-прежнему показывает «нет 1С» и эту кнопку не получает.
  - **Acceptance criteria:**
    - [ ] Кнопка видна только при выбранном клиенте и блоке про договор
    - [ ] Нажатие вызывает POST с `counterparty_id` этого КП
    - [ ] Ящик остаётся открытым, «Отправить в 1С» становится доступной, когда блок снялся
    - [ ] Бейдж «нет 1С» и блок «Занести контрагента» не меняются
  - **Verification:**
    - [ ] `cd frontend && npx vitest run src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
  - **Dependencies:** DIR-002
  - **Files likely touched:**
    - `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`
    - `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
    - `frontend/src/features/commercial-archive/hooks/useArchiveQueries.ts`
  - **Estimated scope:** Small

## Checkpoint: After DIR-003 and DIR-004

- [ ] Vitest обоих компонентов зелёный
- [ ] Повторный pytest из DIR-002 зелёный
- [ ] На временных данных клиент без договора получает номер, клиент с договором второй номер не получает

## Risks

- `insert_card` сейчас всегда пишет пустой `counterparty_id`. Новый путь нельзя провести через него с последующим UPDATE в другом коммите: между ними уникальный индекс ещё не защищает намерение, а импорт должен остаться без id.
- Поиск на 10 строк может скрыть нужного тёзку. Это предел текущего справочника, в этой спеке он не расширяется.
- Пустой бланк снимет блок отправки. Это D5 и текущий `_invoice_export_block`. Менять проверку бланка здесь нельзя.
