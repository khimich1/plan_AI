# План: Поиск клиента в мастере КП по карточкам договоров

**Дата:** 2026-10-21  
**Статус:** IMPLEMENT. Код написан. Имя в мастере и в сохранённом КП — имя карточки, не имя из counterparties.  
**Спека:** `ai_docs/specs/kp-client-search-contract-cards.md`  
**Не это:** карточки архива, импорт реестра, удаление карточки, флаг «не полные», приём файла. Их не откатывать и не помечать сделанными здесь.  
**Не править:** `ai_docs/specs/archive-counterparty-cards.md` — не записывать туда эту фичу как сделанную.

## Overview

На шаге «Клиент» нового КП подсказки идут только по строкам `supply_contract` (те же, что «Занесённые карточки»). Справочник 1С целиком в этот список не попадает. Архивный блок «Занести контрагента» остаётся на `GET /counterparties/search`. Карточка с активным клиентом подставляет его в КП. Карточка без связи пишет имя и ИНН, id оставляет пустым и строку справочника не создаёт. Гейт производства без 1С остаётся закрытым. Неполные карточки из выдачи не убираются.

Сейчас шаг «Клиент» и архивный bind делят `CounterpartyAutocomplete` и `GET /counterparties/search`. Поиск карточек уже есть: `SupplyContractService.search_cards`. При сохранении нового КП без id ИНН в строку КП не пишется — это закрывает отдельная задача плана, не текущий код.

## Architecture Decisions

- Отдельный маршрут мастера, не флаг на `GET /counterparties/search`. Иначе архив можно случайно сузить до карточек.
- `CounterpartyAutocomplete` получает режим. По умолчанию справочник. Режим карточек ставит только `ClientConditionsStep`. `OfferDetailsDrawer` режим не передаёт.
- Фильтр не копировать: сервис мастера вызывает `search_cards`. Короче двух символов — пусто, без выдачи всех карточек. `%` и `_` остаются буквальными, потому что так уже режет `search_cards`.
- Ответ мастера — карточка (id договора, имя, ИНН, номер, признак полноты бланка) и вложенный контрагент, если связь ведёт на активного клиента. Форма `CounterpartyShort` у `/search` не меняется.
- Сломанная связь (нет строки, не клиент, неактивен) не создаёт справочник и не кладёт id в КП. В КП идут имя и ИНН карточки.
- ИНН без id хранится в метаданных черновика и попадает в `customer_inn` только при создании нового КП в архиве. Ветка «в работе» без id по-прежнему зовёт отказ «контрагент не найден». `move_to_production` не ослаблять: пустой id по-прежнему не пускает.
- Свободное имя без выбора карточки не запрещать. Номер договора в КП не писать.
- Неполные карточки отдаёт уже `search_cards`. Подпись «не полные» — только текст строки подсказки.

## Dependency graph

```
KPC-001 тест поиска (красный)
        │
        └── KPC-002 маршрут и сервис
                    │
        KPC-003 тест ИНН и гейта (красный) ── KPC-004 запись ИНН без id
                                                    │
                                                    └── KPC-005 режим autocomplete и vitest
```

KPC-001 и KPC-003 можно написать до кода. KPC-002 не начинается, пока KPC-001 красный по причине «маршрута нет», а не из-за ошибки фикстуры. KPC-005 не переводит архив на новый поиск.

## Task List

### Phase 1 — Поиск

- [x] **KPC-001: Падающий pytest поиска мастера** `(type: test, dependsOn: —)`
  - **Description:** Тест на временной БД, не на живой `plita.db`. Есть карточка без `counterparty_id` и контрагент только в `counterparties`. Новый поиск по фрагменту имени карточки находит карточку (имя, ИНН, номер, контрагент пустой, неполный бланк не скрыт) и не находит контрагента только из справочника. Число строк `counterparties` до и после одинаковое. Старый `GET /counterparties/search` этого контрагента находит. Отдельный случай: `%` и `_` не открывают чужие карточки. Отдельный случай: карточка с активным клиентом отдаёт его id, код 1С, имя, ИНН и КПП и тоже не добавляет строку справочника.
  - **Acceptance:**
    - [x] Сначала тест падал из-за отсутствия маршрута, теперь маршрут есть и тест зелёный
    - [x] Утверждения совпадают со спекой: карточка есть, справочник-только нет, строка справочника не создана
  - **Verify:** `.venv/bin/python -m pytest tests/test_kp_counterparty_card_search.py -q`
  - **Dependencies:** —
  - **Files:** `tests/test_kp_counterparty_card_search.py`
  - **Scope:** S

- [x] **KPC-002: Маршрут мастера поверх search_cards** `(type: feat-be, dependsOn: KPC-001)`
  - **Description:** Роутер → сервис → репозиторий. Новый метод сервиса договоров вызывает `search_cards` и для заполненного `counterparty_id` читает уже существующую строку справочника. Insert в `counterparties` не вызывать. Неактивного и не-клиента во вложение не класть. `GET /counterparties/search` не менять. Роли `admin` и `manager`. `q` короче двух символов на границе — отказ или пустой список, но не все карточки.
  - **Acceptance:**
    - [x] KPC-001 зелёный
    - [x] `tests/test_counterparties_api.py` зелёный: старый поиск по-прежнему справочник
    - [x] Поиск карточек архива не сужен
  - **Verify:** `.venv/bin/python -m pytest tests/test_kp_counterparty_card_search.py tests/test_counterparties_api.py tests/test_archive_contract_cards.py -q`
  - **Dependencies:** KPC-001
  - **Files:** `app/services/supply_contract_service.py`, `app/api/v1/endpoints/counterparties.py`, `app/schemas/counterparties.py`, `tests/test_kp_counterparty_card_search.py`
  - **Scope:** M

### Checkpoint: Phase 1

- [x] Спека принята до старта KPC-001
- [x] Архивный `GET /counterparties/search` и `search_cards` ведут себя как до задачи
- [x] Дальше не идти, если выбор карточки без связи создаёт строку справочника

### Phase 2 — Запись в КП

- [x] **KPC-003: Падающий тест имени и ИНН без id** `(type: test, dependsOn: —)`
  - **Description:** Создание нового КП в архиве из черновика с именем и ИНН карточки и без `counterparty_id` должно записать имя и ИНН и оставить id пустым, не увеличив `counterparties`. Сохранение со статусом «в работе» без id по-прежнему отказ, даже если ИНН передан. Это ещё не перевод в производство: гейт `move_to_production` не ослаблять и в этом тесте не подменять.
  - **Acceptance:**
    - [x] Тест записи ИНН падает на текущем коде, где без id ИНН обнуляется
    - [x] Тест отказа «в работе» без id уже может быть зелёным, если гейт и так закрыт; тогда он остаётся регрессом и не считается причиной менять гейт
  - **Verify:** `.venv/bin/python -m pytest tests/test_offer_counterparty_validation.py -q`
  - **Dependencies:** —
  - **Files:** `tests/test_offer_counterparty_validation.py`
  - **Scope:** S

- [x] **KPC-004: Донести ИНН карточки до нового КП** `(type: feat-be, dependsOn: KPC-003)`
  - **Description:** Шаг «Клиент» умеет передать ИНН выбранной карточки в метаданные черновика. Создание нового КП в архиве без id пишет этот ИНН в `customer_inn` и имя в `customer_name`. Если передан `counterparty_id`, имя и ИНН по-прежнему берутся из справочника, а не из карточки. Обновление скидки и логистики ИНН не стирает. Свободное имя без ИНН пишет пустой ИНН. Ветка «в работе» без id не начинает сохранять КП.
  - **Acceptance:**
    - [x] KPC-003 зелёный
    - [x] Старые тесты снимка контрагента и отказа без id зелёные
    - [x] `move_to_production` по-прежнему требует активного клиента
  - **Verify:** `.venv/bin/python -m pytest tests/test_offer_counterparty_validation.py tests/test_kp_counterparty_snapshot.py tests/test_archive_service.py::test_move_to_production_requires_active_client -q`
  - **Dependencies:** KPC-003
  - **Files:** `app/schemas/commercial.py`, `app/api/v1/endpoints/commercial.py`, `app/services/commercial_draft_lifecycle.py`, `app/services/commercial_workflow_service.py`, фронт черновика только в части передачи ИНН
  - **Scope:** M

### Phase 3 — Экран

- [x] **KPC-005: Режим карточек только у мастера** `(type: feat-fe, dependsOn: KPC-002, KPC-004)`
  - **Description:** Общий autocomplete по умолчанию зовёт старый поиск. Шаг «Клиент» включает режим карточек: с двух символов новый маршрут, в строке имя, ИНН, номер и при неполном бланке «не полные». Выбор связанной карточки кладёт в форму контрагента 1С. Выбор карточки без связи кладёт имя и ИНН, id пустой. «Занести контрагента» режим не включает. Свободное имя без выбора по-прежнему уходит в расчёт без id.
  - **Acceptance:**
    - [x] vitest мастера: шаг «Клиент» вызывает новый поиск и не вызывает старый
    - [x] vitest autocomplete: без режима вызывается старый поиск и не вызывается новый
    - [x] vitest архива: блок «Занести контрагента» не получает режим карточек
    - [x] Неполная карточка есть в списке
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-offer/components/CounterpartyAutocomplete.test.tsx src/features/commercial-offer/components/steps/ClientConditionsStep.test.tsx src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
  - **Dependencies:** KPC-002, KPC-004
  - **Files:** `frontend/src/features/commercial-offer/api/counterpartiesApi.ts`, `frontend/src/features/commercial-offer/components/CounterpartyAutocomplete.tsx`, `frontend/src/features/commercial-offer/components/steps/ClientConditionsStep.tsx`, тесты рядом, `OfferDetailsDrawer.tsx` только если без правок режим и так не уходит в архив
  - **Scope:** M

### Checkpoint: конец

- [x] Команды KPC-002, KPC-004 и KPC-005 зелёные
- [x] Живая `plita.db` не менялась, приложение ради этой задачи не запускалось
- [x] `archive-counterparty-cards.md` не говорит, что поиск мастера уже сделан
- [x] Коммит только по отдельной просьбе

## Boundaries

- Always: тесты до кода на поиске и на записи ИНН. Точечные команды выше, не весь репозиторий.
- Ask first: убрать свободное имя, писать номер договора в КП, прятать неполные карточки, менять гейт производства.
- Never: начать имплементацию до приёмки спеки. Сузить архивный поиск. Откатить чужие незакоммиченные карточки, импорт, удаление, флаг «не полные» и приём файла.
