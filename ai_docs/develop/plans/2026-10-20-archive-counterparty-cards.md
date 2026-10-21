# План: Карточки договоров в архиве

**Дата:** 2026-10-20  
**Статус:** IMPLEMENT — 15 задач сделаны 2026-10-20. Бэкенд карточек 145 passed, фронтенд архива 157 passed. Коммита нет. Живой импорт xlsx в plita.db не запускался.  
**Оркестрация:** `orch-2026-10-20-18-00-contract-cards`  
**Идея:** `ai_docs/ideas/archive-counterparty-merged-card.md`  
**Спека:** `ai_docs/specs/archive-counterparty-cards.md`

Источник задач — этот файл. `tasks.json` в workspace хранит только статусы и DAG, без второго чеклиста.

## Overview

В архиве появляется раздел «Контрагенты»: поиск карточки договора по имени, ИНН или номеру, создание новой и правка всех полей, включая номер. Карточка — строка `supply_contract`. Разовая загрузка листа «ИНН без сомнений» (184 строки, колонки «номер договора», «имя контрагента», «ИНН») заводит неполные карточки: статус «нет», `counterparty_id` пустой, реквизиты бланка пустые. Уже лежащий `0001/10/26` виден в том же списке и импортом не переписывается. Блок договора (форма, статус, скан, Word, PDF, соглашение об ЭДО) открывается с карточки. Кнопка «Договор» на КП уходит. Справочник `counterparties`, привязка клиента к КП и гейт производства не меняются.

## Architecture Decisions

- **Слои.** Роутер только принимает HTTP. Номер, дата из номера и проверка «реквизиты заполнены» — `core/supply_contract.py`. Разбор xlsx — сервис, не репозиторий. SQL — `SupplyContractRepository`.
- **Не лист юриста.** `parse_lawyer_sheet` и `insert_imported` не используем. Они ждут шесть колонок, копируют имя в `full_name`/`short_name` и пишут пустые строки в NOT NULL-колонки. Карточка хранит имя в `imported_name`, реквизиты — NULL.
- **Пустые реквизиты — NULL, не `""`.** Новые карточки пишут NULL. Уже лежащие строки, где юрист записал `""`, не переписываем. «Пусто» для гейта Word — и NULL, и строка из пробелов.
- **Схема.** `CREATE TABLE IF NOT EXISTS` старую таблицу не ослабит. Снятие NOT NULL — пересоздание таблицы по образцу `_migrate_completed_plates_for_sgp`: `PRAGMA foreign_keys=OFF`, копия по именам колонок из `PRAGMA table_info`, drop, rename, заново частичный уникальный индекс. `number` остаётся `NOT NULL UNIQUE`. Индекс `idx_supply_contract_one_active` не трогаем по смыслу: он и так не включает строки с пустым `counterparty_id`, поэтому 184 карточки без связи со справочником совместимы.
- **Дата.** Из номера — день 01, только при загрузке и при создании, если номер передан в теле. Автономер берёт сегодняшнюю дату карточки целиком (`2026-10-20` → `1034/10/26`). Смена номера дату не трогает. Смена даты номер не трогает — это уже так в `apply_update`.
- **Счётчик.** `contract_seq` / `next_contract_seq` не меняем. В счётчик идёт только первый сегмент до `/`, если он целиком из цифр. `1033` и `966/06/26КП` занимают номер. `ТБ-1/10/25` и `0844-06-25` — нет: у дефисного номера нет `/`, весь текст не из цифр.
- **Поиск.** Не SQL `LIKE`. У SQLite без ICU `LIKE` не складывает регистр кириллицы, «секлюц» не найдёт «Секлюцкий». Строк договора сотни, не 4413 клиентов: репозиторий отдаёт строки `supply_contract`, сервис фильтрует `casefold()` по подстроке. `%` и `_` тогда буквальны. Короче двух символов — пустой список, справочник не читаем.
- **Имя в списке.** `imported_name`, если оно непустое, иначе `full_name`, иначе `short_name`. Так `0001/10/26` без `imported_name` всё равно ищется и показывается.
- **Чужой номер.** Текст только «Номер {number} уже есть». Не «уже есть в реестре» и не «У контрагента уже есть договор».
- **Блок на карточке — новые маршруты по `contract_id`.** См. расхождение R2. `replace_contract` и `create_for_kp` для этого не вызываем: первый создаёт новую строку и требует контрагента, второй завязан на КП.
- **Менеджер карточки.** В пользователе есть только `username`. Пустой `manager_name` заполняется логином того, кто сохранил блок реквизитов. Уже записанное имя (`0001/10/26`) не затираем. Отдельного поля в форме нет.
- **Раздел UI.** `counterparties` — значение только экрана архива. В `ArchiveSection` бэкенда его нет: список КП с таким `section` ответит 422. На этом разделе запрос списка КП не уходит.
- **Старый ящик договора.** `SupplyContractDrawer` с кнопки КП снимаем, файл и его тесты не удаляем. Реестр листа юриста остаётся без входа в интерфейсе: спека этот лист не переносит. Новый ящик — отдельный компонент, он переиспользует `SupplyContractFields`.

## Расхождения кода и спеки

Продуктовые решения спеки не меняются. Ниже — где код им противоречит и как это снять в реализации.

| # | Спека | Код сейчас | Как снять |
|---|--------|------------|-----------|
| R1 | Реквизиты и `contract_date` пустые, пока блок не сохранён. ИНН новой карточки можно не заполнять. Менеджер КП пуст до сохранения блока | `contract_date`, `manager_name`, `legal_form`, `full_name`, `short_name`, `signatory_name`, `signatory_verb`, `authority_basis`, `inn`, `ogrn`, `legal_address`, `email`, `bank_name`, `account`, `corr_account`, `bik` — `NOT NULL`. SQLite не снимает NOT NULL через `ALTER COLUMN` | Пересоздать таблицу. Эти колонки nullable. `number`, `status`, `created_at`, `created_by_user_id` остаются обязательными. Свежий `CREATE TABLE` в `_init_supply_contract_schema` сразу nullable, миграция для уже созданной `plita.db` |
| R2 | «Блок остаётся на нынешних маршрутах по `contract_id`»: статус, скан, дата, разбор, сохранение реквизитов, Word, PDF | По `contract_id` есть только PATCH статуса/скана/даты, link и replace. Word, PDF, ЭДО и разбор — `/{kp_id}/supply-contract/...`. Они ищут договор через `counterparty_id` КП. У загруженной карточки связи нет, эти маршруты её не видят. `replace` годится только для отменённого договора с контрагентом и вставляет новую строку | Оставить kp-маршруты (UI их больше не вызывает, тесты КП не сносим). Добавить соседей: `POST /supply-contracts/{id}/requisites`, `POST /supply-contracts/{id}/parse`, `GET /supply-contracts/{id}/document`, `GET /supply-contracts/{id}/edo-agreement`. Та же проверка полей `SupplyContractCreate` и те же сборщики docx. Гейт «Сначала заполните реквизиты договора» стоит и на новых маршрутах, и на старых kp-скачиваниях, чтобы пустая строка не попала в бланк |
| R3 | Ответ карточки с пустыми реквизитами | `SupplyContractOut` и `SupplyContractRegistryItem` требуют непустые строки (`legal_form`, `inn`, `contract_date`, …). PATCH статуса и `GET /supply-contracts` отдают все строки. После загрузки пустая дата даст 500 на валидации ответа | Ослабить эти модели: перечисленные поля `str \| None`. Карточные маршруты отвечают отдельной `ContractCardOut`, тоже с пустыми полями |
| R4 | Поиск без учёта регистра, «секлюц» → три карточки | SQL `LIKE` кириллицу не складывает | Фильтр в Python, `casefold`, подстрока. Не `LIKE` |
| R5 | Спека не описала тело ответа импорта и чтение одной карточки | Ящику нужен один договор после открытия и после сохранения | `GET /contract-cards/{id}`. Импорт отвечает `{created, skipped}`. Это форма ответа, не новое продуктовое правило |
| R6 | «Менеджер КП — пусто, пока не сохранят блок» | Колонка обязательна. Имя менеджера на КП берётся из карточки КП, у пользователя сессии есть только `username` | NULL до сохранения блока. При сохранении реквизитов, если поле пустое, пишем `user["username"]` |

Заголовки боевого файла проверены чтением, в git его не класть: лист «ИНН без сомнений», 184 уникальных номера, максимальный числовой префикс 1033, `0001/10/26` в файле нет, ИНН `440117229007` на трёх строках.

## Dependency graph

```
CDC-001 номер и дата          CDC-002 схема NULL          CDC-012 вкладка раздела
        │                            │                            │
        │                            └── CDC-003 вставка          │
        │                                        │                │
        │                                        └── CDC-004 SQL  │
        │                                                 │       │
        └──────────────────────────────────────────────── CDC-005 сервис
                                                          │
                                          ┌───────────────┴──────────────┐
                                          CDC-006 импорт                 CDC-008 HTTP карточек
                                          │                              │
                                          CDC-007 гейт Word              │
                                          │                              CDC-009 HTTP импорт
                                          └────────── CDC-010 HTTP блока ┘
                                                          │
                                                          CDC-011 клиент
                                                          │
                                          CDC-012 ──────── CDC-013 список
                                                          │
                                                          CDC-014 ящик
                                                          │
                                                          CDC-015 убрать кнопку «Договор»
```

## Task List

### Phase 1 — Номер и схема

- [ ] **CDC-001: Нормализация номера и дата из номера** `(type: feat-be, dependsOn: —)`
  - **Description:** В `core/supply_contract.py` добавить `normalize_contract_number` как в спеке и `contract_date_from_number`. `next_contract_seq` и `format_contract_number` не менять. Дата: режем по `/` и `-`. Год — последний токен, если он начинается ровно с двух цифр (хвост `КП` или `Б` допустим). Месяц — предыдущий токен, 1–2 цифры, значение 1–12. Год `00`–`69` → 2000-е, `70`–`99` → 1900-е, день 01. Иначе `None`. `0785/01/25`, `0786/1/25`, `0844-06-25`, `966/06/26КП`, `ТБ-1/10/25`, `ТБ-3/08/26Б` дают указанные в спеке даты. Номер без месяца и года — `None`.
  - **Acceptance:**
    - [ ] Пробелы схлопываются, пустая строка, строка без цифры и длиннее 32 символов — `ValueError` с текстами из спеки
    - [ ] Шесть примеров номеров дают даты из спеки, день 01
    - [ ] `next_contract_seq(["1033/09/26", "ТБ-1/10/25", "0844-06-25"])` остаётся 1034
  - **Verify:** `pytest tests/test_supply_contract_number.py -q`
  - **Dependencies:** —
  - **Files:** `core/supply_contract.py`, `tests/test_supply_contract_number.py`
  - **Scope:** S

- [ ] **CDC-002: Пустые реквизиты и дата в схеме** `(type: arch, dependsOn: —)`
  - **Description:** В `_init_supply_contract_schema` свежая таблица создаётся с nullable-колонками из R1. Если таблица уже есть и у `contract_date` или `legal_form` стоит `notnull=1`, пересоздать её: выключить FK, создать новую с тем же набором колонок, скопировать строки по именам, удалить старую, переименовать, заново создать `idx_supply_contract_one_active`. Повторный запуск миграции ничего не делает. `counterparties` не трогать. Существующие значения, включая `""` у листа юриста и заполненный `0001/10/26`, копируются как есть.
  - **Acceptance:**
    - [ ] Новая БД: перечисленные колонки nullable, `number` по-прежнему `NOT NULL UNIQUE`
    - [ ] Старая форма таблицы после `ensure_schema`: колонки nullable, строка на месте, частичный уникальный индекс есть
    - [ ] Две строки с `counterparty_id NULL` вставляются. Два одинаковых `number` — нет
  - **Verify:** `pytest tests/test_supply_contract_card_schema.py tests/test_supply_contract.py -q`
  - **Dependencies:** —
  - **Files:** `core/kp_db_schema.py`, `tests/test_supply_contract_card_schema.py`
  - **Scope:** M

### Checkpoint: Phase 1

- [ ] CDC-001 и CDC-002 зелёные, их можно делать параллельно
- [ ] `pytest tests/test_supply_contract.py tests/test_supply_contract_import.py -q` — создание полного договора и лист юриста не сломаны
- [ ] Дальше не идти, если миграция не идемпотентна

### Phase 2 — Репозиторий

- [ ] **CDC-003: Вставка неполной карточки и выборка** `(type: feat-be, dependsOn: CDC-002)`
  - **Description:** В репозитории метод вставки карточки: `counterparty_id` NULL, `imported_name`, `inn` или NULL, `number`, `contract_date` или NULL, `status` «нет», `manager_name` NULL, реквизиты NULL, `created_by_user_id` задан. Не вызывать `_insert_imported_row`. `get_by_number`. Выборка колонок для поиска: id, `imported_name`, `full_name`, `short_name`, `inn`, `number` и поля, которые отдаёт карточка. Занятый номер — ошибка уникальности наверх, без текста «в реестре».
  - **Acceptance:**
    - [ ] Вставка без реквизитов и без даты читается обратно с NULL и статусом «нет»
    - [ ] Повтор того же номера не создаёт вторую строку
    - [ ] `counterparties` этой вставкой не пополняется
  - **Verify:** `pytest tests/test_archive_contract_cards.py -q`
  - **Dependencies:** CDC-002
  - **Files:** `app/repositories/supply_contract_repository.py`, `tests/test_archive_contract_cards.py`
  - **Scope:** M

- [ ] **CDC-004: Смена номера, имя, ИНН и запись реквизитов** `(type: feat-be, dependsOn: CDC-003)`
  - **Description:** В одной транзакции: смена `number` и строка `supply_contract_event` (`field=number`, старое, новое, пользователь, время). Дата в этом методе не пишется. Тот же номер после нормализации — пустая операция, без события. Занятый номер — `IntegrityError` наверх. Отдельный метод меняет `imported_name` и `inn` без события. `save_requisites` обновляет поля бланка на той же строке, номер не меняет; дату пишет только если её передали. `manager_name` пишет только когда в строке он пустой.
  - **Acceptance:**
    - [ ] Смена номера оставляет `contract_date` и добавляет одно событие
    - [ ] Повтор чужого номера не коммитит ни номер, ни событие
    - [ ] Запись реквизитов не меняет номер и не затирает уже стоящего менеджера
  - **Verify:** `pytest tests/test_archive_contract_cards.py -q`
  - **Dependencies:** CDC-003
  - **Files:** `app/repositories/supply_contract_repository.py`, `tests/test_archive_contract_cards.py`
  - **Scope:** M

### Checkpoint: Phase 2

- [ ] Неполная строка, смена номера и уникальность закрыты без HTTP
- [ ] Лист юриста по-прежнему падает на дубле своим старым текстом, не текстом карточки

### Phase 3 — Сервис

- [ ] **CDC-005: Создание, поиск и правка карточки** `(type: feat-be, dependsOn: CDC-001, CDC-004)`
  - **Description:** Методы сервиса для admin/manager (`_assert_writer`). Создание: имя обязательно, ИНН необязателен, номер необязателен. Без номера дата карточки — `_today` или `date.today()`, номер — `format_contract_number(next_contract_seq(...), эта дата)`. С номером — `normalize_contract_number`, дата из номера или NULL, не «сегодня». Занятый номер — `SupplyContractValidationError` «Номер {number} уже есть». Правка имени, ИНН и номера: смена номера дату не пересчитывает. Поиск: короче 2 символов → пусто; иначе подстрока `casefold` по непустому `imported_name`, непустым `full_name` и `short_name`, ИНН и номеру. Имя в ответе — как в решении выше. Справочник не читается.
  - **Acceptance:**
    - [ ] При занятом префиксе `1033` и дате `2026-10-20` новая карточка без номера получает `1034/10/26` и эту дату
    - [ ] Ручной номер `0786/1/25` сохраняется буквально, дата `2025-01-01`. Повтор номера — ошибка с текстом спеки
    - [ ] Смена номера пишет событие, дата прежняя. Старый номер поиском не находится
    - [ ] «секлюц» и ИНН `440117229007` возвращают три карточки. `0793/02` — одну. Запрос из одного символа — пустой список. `%` не маска
  - **Verify:** `pytest tests/test_archive_contract_cards.py -q`
  - **Dependencies:** CDC-001, CDC-004
  - **Files:** `app/services/supply_contract_service.py`, `tests/test_archive_contract_cards.py`
  - **Scope:** M

- [ ] **CDC-006: Импорт трёх колонок** `(type: feat-be, dependsOn: CDC-005)`
  - **Description:** Новый разбор в `app/services/contract_card_import.py`, не в `supply_contract_import.py`. Лист «ИНН без сомнений», если он есть, иначе активный. Заголовки «номер договора», «имя контрагента», «ИНН» (без учёта регистра). Номер через `normalize_contract_number`, без `format_contract_number`. Дата из номера. Статус «нет», `counterparty_id` NULL, реквизиты NULL. Уже существующий номер пропускается, поля не обновляются. Пустое имя или номер, который не проходит нормализацию, — ошибка на весь файл, транзакция откатывается. Повтор номера внутри файла — пропуск, не откат. Ответ `{created, skipped}`. Только admin. Тест собирает xlsx в `tmp_path` на 184 строки: три номера одного ИНН `440117229007` (`0793/02/25`, `0794/02/25`, `0933/03/26`), `0786/1/25`, `0844-06-25`, `966/06/26КП`, `ТБ-3/08/26Б`, и номер с префиксом `1033`. Файл из «Загрузок» тест не открывает.
  - **Acceptance:**
    - [ ] 184 строки фикстуры создают 184 карточки. Число строк `counterparties` то же. `counterparty_id` у новых пустой
    - [ ] `0001/10/26`, вставленный до импорта и отсутствующий в файле, сохраняет имя, ИНН, реквизиты, дату и статус
    - [ ] Второй прогон не меняет имя, которое успели записать вручную
    - [ ] После импорта создание без номера при дате `2026-10-20` даёт `1034/10/26`
  - **Verify:** `pytest tests/test_archive_contract_cards.py tests/test_supply_contract_import.py -q`
  - **Dependencies:** CDC-005
  - **Files:** `app/services/contract_card_import.py`, `app/services/supply_contract_service.py`, `tests/test_archive_contract_cards.py`
  - **Scope:** M

- [ ] **CDC-007: Сохранение бланка и отказ Word до реквизитов** `(type: feat-be, dependsOn: CDC-006)`
  - **Description:** `requisites_ready`: все обязательные поля `SupplyContractCreate` непустые после `strip`, `legal_form` из `LEGAL_FORMS`. NULL и `""` — пусто. Пустая `contract_date` тоже «не готово»: сборщик docx вызывает `fromisoformat`. Сохранение бланка идёт в `save_requisites` той же строки, через уже существующий `_validated_fields`. Номер не меняется. Дата меняется только если она есть в теле. Пустой менеджер становится `username`. `document_for_contract` и `edo_for_contract` до сборки проверяют готовность и кидают `SupplyContractValidationError` «Сначала заполните реквизиты договора». Та же проверка в начале `document_for_kp` и `edo_agreement_for_kp`. Если бланк заполнен, а для ЭДО не хватает полей покупателя, остаётся прежний текст `MSG_EDO_REQUISITES`. `parse_for_contract` вызывает тот же разбор файла, без гейта КП и без требования `counterparty_id`. Не использовать `replace_contract` и `create_for_kp`.
  - **Acceptance:**
    - [ ] Word и PDF у карточки без реквизитов не собираются, текст ошибки «Сначала заполните реквизиты договора»
    - [ ] После сохранения бланка по правилам `SupplyContractCreate` тот же договор отдаёт байты docx. Номер и прежняя дата, если тело дату не прислало, на месте
    - [ ] Заполненный `0001/10/26` по-прежнему проходит гейт
  - **Verify:** `pytest tests/test_archive_contract_cards.py tests/test_supply_contract.py -q`
  - **Dependencies:** CDC-006
  - **Files:** `app/services/supply_contract_service.py`, `tests/test_archive_contract_cards.py`
  - **Scope:** M

### Checkpoint: Phase 3

- [ ] Обязательные случаи спеки закрыты на сервисе, без HTTP и без файла из «Загрузок»
- [ ] `pytest tests/test_supply_contract.py tests/test_supply_contract_import.py -q` зелёные
- [ ] CDC-007 идёт после CDC-006: оба правят `supply_contract_service.py` и один тестовый файл

### Phase 4 — HTTP

Маршруты `/contract-cards` и `/supply-contracts/...` объявить в `archive.py` выше `/{kp_id}`, рядом с уже существующим блоком `/supply-contracts`. Иначе сегмент может уйти в `kp_id: int`.

- [ ] **CDC-008: Схемы и маршруты карточки** `(type: api, dependsOn: CDC-005)`
  - **Description:** `ContractCardCreate` (`name`, необязательные `inn` и `number`), `ContractCardPatch`, `ContractCardOut`, список `{items, count}`. `GET /contract-cards?q=`, `GET /contract-cards/{id}`, `POST /contract-cards`, `PATCH /contract-cards/{id}`. Роли `admin` и `manager`, как у соседних маршрутов архива. Чужой номер — 400 с текстом сервиса. В этом же файле схем ослабить `SupplyContractOut` и `SupplyContractRegistryItem`: поля из R3 становятся `str | None`, чтобы PATCH статуса и список реестра не падали на пустой карточке. `SupplyContractCreate` не ослаблять.
  - **Acceptance:**
    - [ ] Поиск короче двух символов — `{items: [], count: 0}`
    - [ ] Создание без номера и смена на занятый номер отдают ответы сервиса, занятый — 400 «Номер {number} уже есть»
    - [ ] Нет auth — 401. PATCH статуса карточки с NULL-реквизитами — 200, не 500
  - **Verify:** `pytest tests/test_archive_contract_card_api.py -q`
  - **Dependencies:** CDC-005
  - **Files:** `app/schemas/supply_contract.py`, `app/api/v1/endpoints/archive.py`, `tests/test_archive_contract_card_api.py`
  - **Scope:** M

- [ ] **CDC-009: HTTP импорта** `(type: api, dependsOn: CDC-006, CDC-008)`
  - **Description:** `POST /commercial/archive/contract-cards/import`, файл xlsx, только `admin`. Ответ `{created, skipped}`. Менеджер — 403. Ошибка разбора — 400. Маршрут статический, не под `/{id}`.
  - **Acceptance:**
    - [ ] Admin загружает фикстуру и получает числа created/skipped
    - [ ] Manager получает 403, строк не прибавляется
    - [ ] `POST /supply-contracts/import` в этом файле не изменён
  - **Verify:** `pytest tests/test_archive_contract_card_api.py tests/test_supply_contract_import.py -q`
  - **Dependencies:** CDC-006, CDC-008
  - **Files:** `app/api/v1/endpoints/archive.py`, `app/schemas/supply_contract.py`, `tests/test_archive_contract_card_api.py`
  - **Scope:** S

- [ ] **CDC-010: HTTP блока по contract_id** `(type: api, dependsOn: CDC-007, CDC-009)`
  - **Description:** `POST /supply-contracts/{contract_id}/requisites` с телом `SupplyContractCreate`. `POST /supply-contracts/{contract_id}/parse`. `GET /supply-contracts/{contract_id}/document` и `/edo-agreement` с тем же `format`, что у kp-маршрутов. Роли `admin`, `manager`. Неготовые реквизиты — 400 «Сначала заполните реквизиты договора». kp-маршруты создания и скачивания не удалять.
  - **Acceptance:**
    - [ ] Сохранение реквизитов возвращает ту же карточку с заполненным бланком и прежним номером
    - [ ] Скачивание до сохранения — 400 с текстом спеки, без вызова LibreOffice
    - [ ] `POST /{kp_id}/supply-contract` по-прежнему отвечает в старых тестах
  - **Verify:** `pytest tests/test_archive_contract_card_api.py tests/test_supply_contract.py -q`
  - **Dependencies:** CDC-007, CDC-009
  - **Files:** `app/api/v1/endpoints/archive.py`, `tests/test_archive_contract_card_api.py`
  - **Scope:** M

### Checkpoint: Phase 4

- [ ] `pytest tests/test_supply_contract_number.py tests/test_supply_contract_card_schema.py tests/test_archive_contract_cards.py tests/test_archive_contract_card_api.py tests/test_supply_contract.py tests/test_supply_contract_import.py -q`
- [ ] kp-создание договора и лист юриста живы
- [ ] Фронт не начинать, пока тексты 400 не совпали со спекой

### Phase 5 — Интерфейс архива

- [ ] **CDC-011: Клиент и хуки карточки** `(type: feat-fe, dependsOn: CDC-010)`
  - **Description:** Методы в `archiveApi`: поиск, чтение, создание, правка номера/имени/ИНН, импорт не нужен UI в этом релизе. Плюс реквизиты, разбор, скачивание Word/PDF/ЭДО по `contractId`. Хуки по образцу существующих мутаций договора, ключи кэша отдельные от `supply-contract` КП. Тип карточки допускает пустые реквизиты и пустую дату. `SupplyContract.contract_date` и поля бланка в типе реестра — `string | null`, чтобы ответ ослабленной схемы не врал типу.
  - **Acceptance:**
    - [ ] Тест клиента: GET `/api/v1/commercial/archive/contract-cards?q=секлюц` и POST создания с одним именем
    - [ ] Скачивание документа зовёт `/supply-contracts/{id}/document`, не `/{kpId}/supply-contract/document`
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-archive/api/archiveApi.test.ts`
  - **Dependencies:** CDC-010
  - **Files:** `frontend/src/features/commercial-archive/api/archiveApi.ts`, `frontend/src/features/commercial-archive/api/archiveApi.test.ts`, `frontend/src/features/commercial-archive/hooks/useArchiveQueries.ts`, `frontend/src/features/commercial-archive/types/supplyContract.ts`
  - **Scope:** M

- [ ] **CDC-012: Вкладка «Контрагенты»** `(type: ui, dependsOn: —)`
  - **Description:** Тип экрана `ArchivePageSection = ArchiveSection | "counterparties"`. В union бэкенда значение не добавлять. Вкладка «Контрагенты» в `ArchiveSectionTabs`. `?section=counterparties` открывает раздел. На нём не вызывать `useArchiveListQuery` и не показывать фильтр типов продукции и поиск КП. Пока списка нет — короткая подсказка ввести имя, ИНН или номер. Остальные четыре раздела работают как раньше.
  - **Acceptance:**
    - [ ] Пять вкладок, «Контрагенты» выбирается и пишет `section=counterparties`
    - [ ] На этом разделе нет запроса списка КП
    - [ ] «В архиве» по-прежнему показывает список КП
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-archive/components/ArchiveSectionTabs.test.tsx src/pages/commercial-offer-archive/CommercialOfferArchivePage.test.tsx`
  - **Dependencies:** —
  - **Files:** `frontend/src/features/commercial-archive/types/archive.ts`, `frontend/src/features/commercial-archive/components/ArchiveSectionTabs.tsx`, `frontend/src/features/commercial-archive/components/ArchiveSectionTabs.test.tsx`, `frontend/src/pages/commercial-offer-archive/CommercialOfferArchivePage.tsx`, `frontend/src/pages/commercial-offer-archive/CommercialOfferArchivePage.test.tsx`
  - **Scope:** M

- [ ] **CDC-013: Список, поиск и новая карточка** `(type: ui, dependsOn: CDC-011, CDC-012)`
  - **Description:** `ContractCardList` на разделе «Контрагенты». Одна строка поиска. Короче двух символов запрос не уходит, видна подсказка ввести имя, ИНН или номер. Иначе список показывает имя, ИНН и номер. Форма новой карточки: имя обязательно, ИНН и номер необязательны. Успех открывает карточку. Ошибка «Номер … уже есть» видна в `Alert`.
  - **Acceptance:**
    - [ ] Ввод «секлюц» вызывает поиск с этим `q` и рисует три строки
    - [ ] Один символ не вызывает API
    - [ ] Создание с одним именем вызывает POST без номера
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-archive/components/ContractCardList.test.tsx`
  - **Dependencies:** CDC-011, CDC-012
  - **Files:** `frontend/src/features/commercial-archive/components/ContractCardList.tsx`, `frontend/src/features/commercial-archive/components/ContractCardList.test.tsx`, `frontend/src/pages/commercial-offer-archive/CommercialOfferArchivePage.tsx`
  - **Scope:** M

- [ ] **CDC-014: Блок договора на карточке** `(type: ui, dependsOn: CDC-013)`
  - **Description:** `ContractCardDrawer`. Не монтировать `SupplyContractDrawer` и `SupplyContractRegistry`. Форма — `SupplyContractFields`. Пустая карточка не подставляет `legal_form: "ooo"`: пока менеджер не выбрал форму, сохранение выключено. Статус, отметка скана и файл — тем же PATCH `/supply-contracts/{id}`, что реестр. Сохранение бланка — `POST .../requisites`. Word, PDF и соглашение об ЭДО — маршруты по `contract_id`. Пока реквизиты пустые, нажатие показывает ответ «Сначала заполните реквизиты договора». Разбор файла карточки — `POST .../parse` по id договора. Номер, имя и ИНН правятся здесь же через PATCH карточки; смена номера не перезаписывает дату на экране, пока ответ сервера её не изменил.
  - **Acceptance:**
    - [ ] Пустая карточка показывает имя, ИНН, номер и не шлёт создание договора от КП
    - [ ] Сохранение бланка вызывает requisites, не `createSupplyContract`
    - [ ] Word до заполнения показывает текст отказа сервера
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-archive/components/ContractCardDrawer.test.tsx`
  - **Dependencies:** CDC-013
  - **Files:** `frontend/src/features/commercial-archive/components/ContractCardDrawer.tsx`, `frontend/src/features/commercial-archive/components/ContractCardDrawer.test.tsx`, `frontend/src/features/commercial-archive/components/SupplyContractFields.tsx`, `frontend/src/features/commercial-archive/components/ContractCardList.tsx`
  - **Scope:** M

- [ ] **CDC-015: Убрать кнопку «Договор» с КП** `(type: ui, dependsOn: CDC-014)`
  - **Description:** В `OfferDetailsDrawer` убрать `SupplyContractArchiveButton`, состояние открытия и `<SupplyContractDrawer>`. Привязку контрагента, бейдж и «В производство» не трогать. Тесты, которые ищут кнопку «Договор», переписать: кнопки нет. `SupplyContractDrawer.tsx` и его тесты оставить.
  - **Acceptance:**
    - [ ] В карточке КП нет кнопки «Договор»
    - [ ] Блок привязки контрагента на месте
    - [ ] Тесты `SupplyContractDrawer.test.tsx` не переписываются под удаление компонента
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx src/features/commercial-archive/components/SupplyContractDrawer.test.tsx`
  - **Dependencies:** CDC-014
  - **Files:** `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`, `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
  - **Scope:** S

### Checkpoint: Complete

- [ ] CDC-001…015 закрыты
- [ ] Команды из секции Commands спеки зелёные
- [ ] Регресс: лист юриста, `POST /{kp_id}/supply-contract`, привязка контрагента к КП, `move_to_production` без id → 400
- [ ] Вручную (после кода, не в этом шаге): раздел «Контрагенты», поиск «секлюц», карточка без реквизитов не скачивает Word, кнопка «Договор» на КП отсутствует
- [ ] xlsx реестра в git не появляется
- [ ] Выполнение начнётся только после явной приёмки плана

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Пересоздание `supply_contract` на живой `plita.db` теряет колонку или рвёт FK событий | High | Копировать по `PRAGMA table_info`, FK выключить на время, индекс восстановить, тест на старой форме таблицы с живой строкой и событием |
| Пустая карточка в старом `SupplyContractOut` даёт 500 | High | R3 в CDC-008, до импорта и до UI |
| `LIKE` не находит «секлюц» | High | Поиск в Python, `casefold`, отдельный тест на регистр |
| `replace_contract` или `create_for_kp` использованы как «сохранить бланк» | High | Отдельный `save_requisites` на той же строке. В задаче прямо запрещены оба старых метода |
| Импорт перепишет `0001/10/26` или имя, которое менеджер уже правил | High | Совпавший номер только пропускается. Тест до/после |
| Кнопку «Договор» снимут раньше, чем заработает карточка | Med | CDC-015 зависит от CDC-014 |
| `section=counterparties` уйдёт в список КП | Med | Отдельный тип экрана, query списка КП выключен |
| Пустая форма по умолчанию «ООО» выглядит как заполненный бланк | Med | Не подставлять `legal_form`, сохранение выключено, пока форма не выбрана |
| Гейт Word сработает на заполненном договоре КП | Med | Гейт смотрит обязательные поля, не флаг. Тест на полный `0001/10/26` и старый `test_supply_contract.py` |
| Кириллица в номере `ТБ-…` ломает уникальность или счётчик | Low | Номер хранится как текст. Счётчик по-прежнему смотрит только цифровой префикс до `/` |
| Реестр листа юриста останется без кнопки входа | Low | Так и задумано этой спекой. Компонент не удаляем, endpoint импорта юриста не трогаем |

## Parallelization

- Сразу параллельно: CDC-001, CDC-002, CDC-012. Разные файлы.
- CDC-003 только после CDC-002. CDC-004 сразу за CDC-003. Оба правят репозиторий.
- После CDC-005: CDC-006 и CDC-008 параллельно. CDC-008 не ждёт импорт.
- CDC-007 после CDC-006: один сервис и один тестовый файл. Когда CDC-008 тоже готов, CDC-007 идёт параллельно с CDC-009: сервис против `archive.py`.
- CDC-009 после CDC-006 и CDC-008. CDC-010 после CDC-007 и CDC-009: один `archive.py`.
- CDC-011 после CDC-010. CDC-013 после CDC-011 и CDC-012. Дальше строго CDC-014 → CDC-015.
- `counterparties`, мастер КП, гейт производства, `POST /supply-contracts/import` — не параллелить сюда, не в скоупе.

## Open Questions

Нет. Отклонения от D1–D12 — только правкой спеки. R1–R6 закрыты в этом плане, спеку под них не переписываем.

## Приём файла карточки

**Дата:** 2026-10-21  
**Спека:** секция «Приём файла карточки» в `ai_docs/specs/archive-counterparty-cards.md`  
**Статус:** IMPLEMENT. `ContractCardDrawer.test.tsx` — 27 passed.

Файл принимается в блоке реквизитов уже открытой карточки. Жест и колонка источника — как у `SupplyContractDrawer`: выбор и бросок, из броска только первый файл, слева источник, справа бланк. JPEG, PNG и PDF без текста — картинка и масштаб. Word, Excel и PDF с текстом — имя файла и `source_text`, без масштаба. Разбор вызывает уже готовый `POST /supply-contracts/{id}/parse` и только заполняет бланк. Строку пишет кнопка «Сохранить реквизиты». Отказ «это не карточка контрагента» уже введённые поля не затирает. Импорт реестра и `OfferDetailsDrawer` не меняются. Форма «Новая карточка» тоже принимает файл — см. абзац ниже.

- [x] **CDC-016: Зона файла и превью в реквизитах** `(type: ui, dependsOn: CDC-014)`
  - **Description:** В `ContractCardDrawer` заменить голый input на зону выбора и перетаскивания. Бросок берёт `files[0]` и вызывает `parseContractCard`. Слева колонка источника: картинка и PDF без `source_text` — превью и те же кнопки масштаба, что у старого ящика; `.docx`, `.xlsx` и PDF с текстом — имя и `source_text`, без масштаба. Успешный разбор по-прежнему заполняет бланк, `doubtful`, счета и банки. Ошибка показывает текст и не меняет поля. С разбора `saveContractRequisites` не вызывается.
  - **Acceptance:**
    - [x] Бросок docx, xlsx или pdf вызывает `parseContractCard` с первым файлом, второй файл броска не уходит
    - [x] После текстового файла видны имя и `source_text`, кнопок «Увеличить» нет
    - [x] JPEG показывает картинку и масштаб; PDF без текста — превью и масштаб
    - [x] Отказ «это не карточка контрагента» оставляет ИНН, который уже был в бланке
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-archive/components/ContractCardDrawer.test.tsx`
  - **Dependencies:** CDC-014
  - **Files:** `frontend/src/features/commercial-archive/components/ContractCardDrawer.tsx`, `frontend/src/features/commercial-archive/components/ContractCardDrawer.test.tsx`
  - **Scope:** S

Форма «Новая карточка» принимает тот же файл: выбор и бросок, из броска первый файл, подпись «или перетащите файл сюда». Разбор — `POST /contract-cards/parse` (admin, manager): тот же ответ, что у parse, без записи строки. После успеха клиент создаёт карточку с именем (краткое, иначе полное, иначе имя файла без расширения) и ИНН, номер выдаёт программа, и открывает её с заполненным бланком. Реквизиты бланка по-прежнему пишет только «Сохранить реквизиты». Отказ разбора остаётся текстом на форме, строка в базе не появляется. Дроп внутри уже открытой карточки не меняется. Импорт реестра не используется.
