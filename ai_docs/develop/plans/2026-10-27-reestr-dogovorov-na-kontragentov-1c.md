# План: реестр договоров на контрагентов 1С

**Дата:** 2026-10-27  
**Статус:** PLAN, на ревью. Кода ещё нет.  
**Спека:** `ai_docs/specs/reestr-dogovorov-na-kontragentov-1c.md`  
**Не это:** лист «Не вошли», экран выбора одного номера из нескольких, удаление карточек в продукте, GUID контрагента, правка бланка.

## Overview

Сейчас отправка сама пишет `counterparty_id` на свободный договор с тем же ИНН и при нескольких номерах берёт позднюю дату. Шаг «Клиент» ищет карточки договоров, поэтому в подсказке видны `0001/09/26` и `0006/09/26`.

После этой работы связь ставится только когда на ИНН один активный клиент и один неотменённый договор, и только отдельным вызовом после загрузки. Отправка и перевод в производство этот вызов не делают. Номер в JSON берётся из уже связанного действующего договора. На шаге «Клиент» снова ищется справочник 1С. Если клиент есть, а договора нет, на карточке архива виден текст «в реестре договора нет».

217 строк листа «ИНН без сомнений» загружаются в локальную `plita.db` последней задачей, после зелёных тестов. 44 строки листа «Не вошли» не грузятся. Живую базу тесты не открывают.

## Architecture Decisions

- Правило «один договор на ИНН» остаётся в `select_contract_client_links`. Несколько договоров одного ИНН дают пустой результат, поздняя дата больше не выбирает номер.
- `link_unique_inn_contracts` остаётся способом записать связь, но его зовут только после импорта. `_invoice_contract` и `move_to_production` его не зовут.
- Подсказка кнопки «Отправить в 1С» больше не показывает договор через `planned_contract_client_links`. Кнопка видит только `get_active_by_counterparty`.
- Если на КП пустой `counterparty_id`, а `customer_inn` уже указывает на клиента с действующим договором, отправка по-прежнему может запомнить этого клиента на КП. Договор она при этом не связывает.
- Текст «в реестре договора нет» — строка на карточке архива рядом с клиентом. Отказ кнопки остаётся «Нет действующего договора поставки».
- Шаг «Клиент» использует поиск справочника по умолчанию. Поиск карточек в архиве не удаляется.
- Очистка `supply_contract` и импорт xlsx не становятся эндпоинтом.

## Dependency graph

```
REG-001 один договор на ИНН, без поздней даты
        │
        └── REG-002 отправка и производство не связывают договор
                    │
                    ├── REG-003 текст «в реестре договора нет»
                    │
                    └── REG-005 очистка и загрузка 217 на этой plita.db

REG-004 шаг «Клиент» снова ищет справочник 1С
        │
        └── REG-005
```

REG-004 не ждёт REG-001. REG-005 только после зелёных REG-002, REG-003 и REG-004.

## Task List

### Phase 1 — Правило связи

- [ ] **REG-001: Один договор на ИНН** `(type: feat-be)` (⏳ Pending)
  - **Description:** `select_contract_client_links` возвращает пару, только если на ИНН ровно один неотменённый договор с пустым `counterparty_id` и ровно один активный клиент, и у этого клиента ещё нет действующего договора. Два и больше договоров дают пусто. Поздняя `contract_date` номер не выбирает.
  - **Acceptance criteria:**
    - [ ] Один договор и один клиент с тем же ИНН дают эту пару
    - [ ] Два договора одного ИНН дают пусто, ни поздний, ни ранний
    - [ ] Две карточки 1С на один ИНН и ИНН без карточки дают пусто
    - [ ] Отменённый, уже связанный и второй договор клиента не выбираются
  - **Verification:**
    - [ ] `.venv/bin/python -m pytest tests/test_kp_invoice_existing_contract.py -q`
  - **Dependencies:** нет
  - **Files likely touched:**
    - `app/repositories/contract_client_link.py`
    - `tests/test_kp_invoice_existing_contract.py`
  - **Estimated scope:** Small

### Phase 2 — Отправка только читает

- [ ] **REG-002: Отправка и производство не связывают договор** `(type: feat-be, dependsOn: REG-001)` (⏳ Pending)
  - **Description:** Запись счёта и перевод в производство не вызывают `link_unique_inn_contracts`. Подсказка кнопки не читает запланированную пару. Номер берётся из `get_active_by_counterparty`. Пустой или чужой `counterparty_id` договора номер не даёт.
  - **Acceptance criteria:**
    - [ ] Договор с пустым `counterparty_id` и совпавшим ИНН после отправки остаётся без связи, документ не собирается
    - [ ] Уже связанный договор этого клиента попадает в `НомерДоговора`
    - [ ] Договор другого клиента не становится номером этого КП
    - [ ] Два договора одного ИНН не получают `counterparty_id` от отправки
    - [ ] Перевод в производство не пишет связь
    - [ ] Число строк `counterparties` не меняется
  - **Verification:**
    - [ ] `.venv/bin/python -m pytest tests/test_kp_invoice_existing_contract.py tests/test_archive_invoice_export.py -q`
  - **Dependencies:** REG-001
  - **Files likely touched:**
    - `app/services/archive_service.py`
    - `tests/test_kp_invoice_existing_contract.py`
    - `tests/test_archive_invoice_export.py`
  - **Estimated scope:** Medium

## Checkpoint: After REG-002

- [ ] Выбор и отправка зелёные на временной базе
- [ ] Живая `plita.db` ещё не открывалась

### Phase 3 — Экран

- [ ] **REG-003: Текст «в реестре договора нет»** `(type: ui, dependsOn: REG-002)` (⏳ Pending)
  - **Description:** На карточке КП в архиве, когда `counterparty_id` заполнен, а действующего договора нет, рядом с клиентом виден текст «в реестре договора нет». Бейдж «нет 1С» при заполненном id не показывается. Кнопка отправки по-прежнему объясняет «Нет действующего договора поставки».
  - **Acceptance criteria:**
    - [ ] Клиент 1С без договора: виден текст «в реестре договора нет», бейджа «нет 1С» нет
    - [ ] Пустой `counterparty_id`: бейдж «нет 1С» есть, фразы про реестр нет
    - [ ] Связанный договор: ни бейджа, ни фразы про реестр
  - **Verification:**
    - [ ] `cd frontend && npx vitest run src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
  - **Dependencies:** REG-002
  - **Files likely touched:**
    - `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`
    - `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
  - **Estimated scope:** Small

- [ ] **REG-004: Шаг «Клиент» ищет справочник 1С** `(type: ui)` (⏳ Pending)
  - **Description:** Мастер нового КП больше не передаёт `searchSource="contract-cards"`. Подсказки идут из справочника: имя, ИНН, КПП, код. Номер договора в строке не пишется. Свободное имя по-прежнему сохраняется без `counterparty_id`. Поиск карточек в архиве не меняется.
  - **Acceptance criteria:**
    - [ ] Ввод имени вызывает поиск справочника и не вызывает поиск карточек
    - [ ] Выбор пишет id, имя, ИНН, КПП и код
    - [ ] Имя без выбора оставляет `counterparty_id` пустым
    - [ ] Блок «Занести контрагента» по-прежнему ищет справочник
  - **Verification:**
    - [ ] `cd frontend && npx vitest run src/features/commercial-offer/components/steps/ClientConditionsStep.test.tsx src/features/commercial-offer/components/CounterpartyAutocomplete.test.tsx src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
  - **Dependencies:** нет
  - **Files likely touched:**
    - `frontend/src/features/commercial-offer/components/steps/ClientConditionsStep.tsx`
    - `frontend/src/features/commercial-offer/components/steps/ClientConditionsStep.test.tsx`
    - `frontend/src/features/commercial-offer/components/CounterpartyAutocomplete.test.tsx`
  - **Estimated scope:** Small

## Checkpoint: After REG-003 and REG-004

- [ ] Тесты REG-002, REG-003 и REG-004 зелёные
- [ ] Живая `plita.db` ещё не открывалась

### Phase 4 — Локальная база

- [ ] **REG-005: Очистить договоры и загрузить 217** `(type: chore, dependsOn: REG-002, REG-003, REG-004)` (⏳ Pending)
  - **Description:** На этой машине очищаются `supply_contract` и `supply_contract_event`. Затем импортируется лист «ИНН без сомнений» из `/home/roman/Загрузки/Реестр договоров для загрузки.xlsx`. После импорта один раз вызывается `link_unique_inn_contracts`. Лист «Не вошли» не читается. Этот шаг не входит в pytest.
  - **Acceptance criteria:**
    - [ ] В `supply_contract` 217 строк
    - [ ] У каждой заполнен `counterparty_id` активного клиента с тем же ИНН
    - [ ] Номеров `0001/09/26` и `0006/09/26` нет
    - [ ] У ДСТрой 76 номер `0935/04/26`, у Хорса `0934/04/26`
    - [ ] Число строк `counterparties` не изменилось
  - **Verification:**
    - [ ] Ручной подсчёт строк и две проверки номеров на локальной `plita.db` после импорта
    - [ ] Повторный прогон тестов REG-002 не открывает эту базу
  - **Dependencies:** REG-002, REG-003, REG-004
  - **Files likely touched:** локальная `plita.db` только. Код импорта не меняется
  - **Estimated scope:** Small

## Risks

- Старые тесты ждут поздний номер в момент отправки. Их переписывают под пустой результат, а не удаляют отказ «Нет действующего договора поставки».
- Подсказка кнопки сейчас показывает ещё не связанный договор. Если оставить `planned_contract_client_links`, кнопка обещает номер, который отправка уже не возьмёт.
- REG-005 стирает семь программных карточек на этой базе, включая кривые бланки `0002/09/26` и `0003/09/26`. Это намеренно и только после зелёных тестов.
- Импорт сам `counterparty_id` не пишет. Без вызова `link_unique_inn_contracts` после него 217 строк останутся свободными и отправка их не подхватит.
