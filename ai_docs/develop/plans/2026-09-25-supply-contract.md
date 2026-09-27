# План: Договор поставки в архиве КП

**Дата:** 2026-09-25  
**Статус:** срезы SC-001…SC-013 в коде; выдача номеров на живой базе выключена, пока не импортирован реальный лист юриста  
**Идея:** `ai_docs/ideas/supply-contract.md`  
**Спека:** `ai_docs/specs/supply-contract.md`

## Overview

В карточке архива появляется «Договор». Менеджер подтверждает реквизиты покупателя и получает один номер на контрагента, docx и строку реестра. Повторное КП того же контрагента номер не плодит. Статус подписи меняется в реестре и в 1С не уходит. В шапку будущего JSON добавляется `НомерДоговора`. Живой обмен, спецификация и правка юридического текста бланка не входят.

## Architecture Decisions

- **Договор не колонка `counterparties`.** Реквизиты печати живут в `supply_contract`. Справочник по-прежнему затирается выгрузкой 1С.
- **КП номер не хранит.** Чтение действующего договора — по `KP_offers.counterparty_id`. На КП добавляются только `order_number_1c` и `order_status_1c` под будущий ответ, эта кнопка их не заполняет.
- **Один действующий.** Частичный уникальный индекс на `counterparty_id`, где статус не «отмена не будем работать». Повторный POST возвращает уже выданный номер и не увеличивает счётчик.
- **Номер в одной транзакции SQLite.** `BEGIN IMMEDIATE`, `seq = max + 1` по всем строкам, включая отменённые. Формат `f"{seq:04d}/{month:02d}/{year%100:02d}"`. Дата после выдачи меняется отдельным PATCH, номер не трогается.
- **Преамбула в `core/supply_contract.py`.** ООО и АО — «в лице … на основании …». ИП, КФХ и физлицо — без формулы ООО. Род глагола только из поля формы.
- **Печать — python-docx по файлу шаблона** `templates/supply_contract.docx` (копия бланка, без данных клиентов). Подстановка известных пустых мест шапки, преамбулы, п. 10.3(б) и блока «ПОКУПАТЕЛЬ». Приложение М-2 не заполняется.
- **Разбор карточки не выдаёт номер.** `parse` возвращает поля и список сомнительных. docx/pdf с текстом не вызывают зрение. Картинка идёт в существующий OCR-контур отдельным промптом полей, без парсера марок и без свободной инструкции модели.
- **Гейты архива те же:** статус «в архиве», `assert_offer_write_access`, роли `admin`/`manager`. Импорт листа — только `admin`.
- **Скан — файл на диске**, путь в строке договора. В git не попадает.

## Dependency graph

```
номер + преамбула + нормализация полей
    │
    ├── схема supply_contract
    │       │
    │       └── сервис: получить / создать / отменить / заменить
    │               │
    │               ├── HTTP архива
    │               │       └── drawer + реестр
    │               ├── docx
    │               └── parse (текст, затем картинка)
    │
    └── функция шапки JSON (НомерДоговора)
импорт листа — после схемы и сервиса, до включения выдачи номеров
```

## Task List

### Phase 1 — Правила без БД

- [x] **SC-001: Номер и преамбула**
  - **Description:** `format_contract_number`, выбор следующего seq из списка уже занятых номеров, текст преамбулы для пяти видов, нормализация ИНН/КПП/ОГРН/БИК/счетов и почты. Род глагола не выводится из ФИО.
  - **Acceptance:**
    - [x] `1027` и дата 25.09.2026 → `1028/09/26`
    - [x] Отменённый номер остаётся занятым
    - [x] Преамбула ИП и КФХ не содержит «Общество с ограниченной ответственностью»
    - [x] `760 4 01001` и `is-ag @mail.ru` склеиваются
  - **Verify:** `pytest tests/test_supply_contract_number.py -q`
  - **Dependencies:** —
  - **Files:** `core/supply_contract.py`, `tests/test_supply_contract_number.py`
  - **Scope:** S

### Checkpoint: Phase 1

- [x] Правила номера и преамбулы зелёные без схемы и без HTTP

### Phase 2 — Хранение и выдача

- [x] **SC-002: Таблицы**
  - **Description:** В `core/kp_db_schema.py` таблицы `supply_contract` и `supply_contract_event` как в спеке, частичный уникальный индекс одного действующего `counterparty_id`. На `KP_offers` колонки `order_number_1c`, `order_status_1c`.
  - **Acceptance:**
    - [x] Повторный запуск схемы не падает
    - [x] Второй действующий договор того же контрагента нарушает индекс
    - [x] Два отменённых на одного контрагента допустимы
  - **Verify:** `pytest tests/test_supply_contract.py -q -k schema`
  - **Dependencies:** SC-001
  - **Files:** `core/kp_db_schema.py`, `tests/test_supply_contract.py`
  - **Scope:** S

- [x] **SC-003: Создать и прочитать**
  - **Description:** Сервис по КП: нет контрагента или статус не «в архиве» — ошибка по-русски. Нет действующего договора — валидация полей (КПП обязателен только для ООО и АО), вставка, номер, событие создания. Договор уже есть — вернуть его, seq не растёт. Запись события смены статуса.
  - **Acceptance:**
    - [x] Первое подтверждение создаёт одну строку со статусом «нет»
    - [x] Второе подтверждение того же контрагента возвращает тот же номер
    - [x] ООО без КПП — ошибка, ИП без КПП — строка создана
    - [x] Смена даты не меняет номер
  - **Verify:** `pytest tests/test_supply_contract.py -q`
  - **Dependencies:** SC-002
  - **Files:** `app/repositories/supply_contract_repository.py`, `app/services/supply_contract_service.py`, `app/schemas/supply_contract.py`, `tests/test_supply_contract.py`
  - **Scope:** M

- [x] **SC-004: HTTP**
  - **Description:** `GET/POST /commercial/archive/{kp_id}/supply-contract` под ролями архива и теми же гейтами записи, что `bind_counterparty`. POST существующего договора — 200 и прежний номер, не 409.
  - **Acceptance:**
    - [x] Без контрагента — 400
    - [x] Не архив — 400
    - [x] Нет auth — 401
  - **Verify:** `pytest tests/test_archive_endpoints.py -q -k supply_contract`
  - **Dependencies:** SC-003
  - **Files:** `app/api/v1/endpoints/archive.py`, `tests/test_archive_endpoints.py`
  - **Scope:** S

### Checkpoint: Phase 2

- [x] Номер выдаётся через API и не дублируется на второго КП того же контрагента

### Phase 3 — Печать и разбор карточки

- [x] **SC-005: Docx**
  - **Description:** Положить бланк в `templates/supply_contract.docx`. Сборка docx подставляет номер, дату, преамбулу, e-mail п. 10.3(б) и блок покупателя. Блок поставщика и приложение М-2 остаются как в файле. `GET .../document` отдаёт файл.
  - **Acceptance:**
    - [x] В документе есть выданный номер и нет пустого «ДОГОВОР N /26»
    - [x] ИНН покупателя есть, реквизиты ЖБК СТАРТ не стёрты
    - [x] Повторное скачивание собирается из сохранённых полей
  - **Verify:** `pytest tests/test_supply_contract_docx.py -q`
  - **Dependencies:** SC-004
  - **Files:** `templates/supply_contract.docx`, `core/supply_contract_docx.py`, `app/api/v1/endpoints/archive.py`, `tests/test_supply_contract_docx.py`
  - **Scope:** M

- [x] **SC-006: Разбор текста карточки**
  - **Description:** `POST .../supply-contract/parse` для docx и pdf с текстовым слоем. Результат — поля и список сомнительных, без записи договора. Зрение не вызывается.
  - **Acceptance:**
    - [x] Карточка с ИНН, счётом и директором заполняет форму
    - [x] Разорванный пробелами ИНН приходит склеенным
    - [x] Пустой разбор не создаёт строку `supply_contract`
  - **Verify:** `pytest tests/test_supply_contract_card_parse.py -q`
  - **Dependencies:** SC-001
  - **Files:** `core/supply_contract_card_parse.py`, `app/services/supply_contract_service.py`, `tests/test_supply_contract_card_parse.py`
  - **Scope:** M

- [x] **SC-007: Картинка тем же OCR-контуром**
  - **Description:** Для изображения и скана без текста — препроцесс, Extract, Verify из `core/ocr/` с промптом полей договора. Второй проход подменяет поля только при непустых правках. Провайдер в тестах подменён. Свободной инструкции модели нет.
  - **Acceptance:**
    - [x] Ответ Extract попадает в ту же форму, что SC-006
    - [x] Пустые правки Verify не затирают Extract
    - [x] Пустой Verify помечает сверку неуспешной и договор не создаёт
  - **Verify:** `pytest tests/test_supply_contract_card_parse.py -q -k image`
  - **Dependencies:** SC-006
  - **Files:** `core/supply_contract_card_ocr.py`, `core/ocr/prompts.py`, `tests/test_supply_contract_card_parse.py`
  - **Scope:** M

### Checkpoint: Phase 3

- [x] Подтверждённые поля печатаются в docx
- [x] Разбор файла сам по себе номер не занимает

### Phase 4 — Реестр, отмена, экран

- [x] **SC-008: Статус, скан, реестр**
  - **Description:** `PATCH` статуса и отметки скана пишет `supply_contract_event`. Файл скана сохраняется на диск, в ответе только признак «файл есть». `GET /commercial/archive/supply-contracts` отдаёт колонки реестра.
  - **Acceptance:**
    - [x] Смена на «подписан по ЭДО» не удаляет номер и не создаёт второй договор
    - [x] В событии есть user_id и оба значения
    - [x] «Отмена» снимает договор с «действующих»
  - **Verify:** `pytest tests/test_supply_contract.py tests/test_archive_endpoints.py -q -k registry`
  - **Dependencies:** SC-004
  - **Files:** `app/services/supply_contract_service.py`, `app/api/v1/endpoints/archive.py`, `tests/test_supply_contract.py`
  - **Scope:** M

- [x] **SC-009: Новый договор после отмены**
  - **Description:** `POST .../supply-contracts/{id}/replace` разрешён только если текущий статус — отмена. Иначе ошибка «У контрагента уже есть договор …». Новый seq, старый номер на месте.
  - **Acceptance:**
    - [x] Replace без отмены — отказ, count договоров не растёт
    - [x] Replace после отмены — новый номер, старый занят
  - **Verify:** `pytest tests/test_supply_contract.py -q -k replace`
  - **Dependencies:** SC-008
  - **Files:** `app/services/supply_contract_service.py`, `app/api/v1/endpoints/archive.py`, `tests/test_supply_contract.py`
  - **Scope:** S

- [x] **SC-010: Кнопка и форма в архиве**
  - **Description:** В `OfferDetailsDrawer` кнопка «Договор». Нет `counterparty_id` — недоступна, подсказка как у «В производство». Есть договор — номер, дата, статус, скачивание. Нет — форма: файл или ручной ввод, поля, подсветка сомнительных, подтверждение. Картинка остаётся рядом с полями.
  - **Acceptance:**
    - [x] Без контрагента кнопка disabled
    - [x] Есть договор — формы создания нет
    - [x] Подтверждение вызывает POST и показывает номер
  - **Verify:** `cd frontend && npm run test -- src/features/commercial-archive/components/SupplyContractDrawer.test.tsx`
  - **Dependencies:** SC-004, SC-006
  - **Files:** `frontend/src/features/commercial-archive/components/SupplyContractDrawer.tsx`, `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`, `frontend/src/features/commercial-archive/api/archiveApi.ts`, `SupplyContractDrawer.test.tsx`
  - **Scope:** M

- [x] **SC-011: Реестр на экране**
  - **Description:** Из той же кнопки список колонок листа юриста. Смена статуса и отметки скана. «Новый договор» виден только у отменённой строки и не срабатывает сам при выборе «отмена».
  - **Acceptance:**
    - [x] Смена статуса вызывает PATCH
    - [x] Выбор «отмена» не вызывает replace
    - [x] Кнопка «Новый договор» есть только на отменённой строке
  - **Verify:** `cd frontend && npm run test -- src/features/commercial-archive/components/SupplyContractRegistry.test.tsx`
  - **Dependencies:** SC-008, SC-009, SC-010
  - **Files:** `frontend/src/features/commercial-archive/components/SupplyContractRegistry.tsx`, `SupplyContractRegistry.test.tsx`
  - **Scope:** M

### Checkpoint: Phase 4

- [ ] Менеджер проходит путь: пусто → форма → номер → реестр → отмена → явный новый договор
- [x] `cd frontend && npm run typecheck`

### Phase 5 — Шов JSON и лист юриста

- [x] **SC-012: Номер в шапке JSON**
  - **Description:** Функция, которая к уже собранному документу `КоммерческоеПредложение` добавляет `НомерДоговора` из действующего договора контрагента. Нет договора — ошибка, документ не возвращается. Живой отправки нет.
  - **Acceptance:**
    - [x] В шапке есть номер, в `Товары` его нет
    - [x] Без действующего договора функция падает предсказуемой ошибкой
  - **Verify:** `pytest tests/test_supply_contract.py -q -k json_header`
  - **Dependencies:** SC-003
  - **Files:** `core/supply_contract.py`, `tests/test_supply_contract.py`
  - **Scope:** S

- [x] **SC-013: Импорт листа**
  - **Description:** Admin-разбор xlsx реестра: номер, дата, имя, менеджер, статус, отметка скана. Строка создаётся без `counterparty_id`. Связка — отдельное действие менеджера, не по похожему имени. Файл листа в репозиторий не кладётся. Пока импорт не выполнен на живой базе, выдачу новых номеров в проде не включать.
  - **Acceptance:**
    - [x] `1027/09/26` из листа занимает seq
    - [x] Следующий созданный договор получает `1028/...`, а не `0001`
    - [x] Имя из листа само не проставляет `counterparty_id`
  - **Verify:** `pytest tests/test_supply_contract_import.py -q`
  - **Dependencies:** SC-003
  - **Files:** `app/services/supply_contract_import.py`, `app/api/v1/endpoints/archive.py`, `tests/test_supply_contract_import.py`
  - **Scope:** M

### Checkpoint: готово к реализации по срезам

- [ ] SC-001…SC-013 описаны, код не написан
- [ ] Следующий шаг после приёмки плана — SC-001, не весь контур сразу

## Risks and Mitigations

| Risk | Impact | Mitigation |
|---|---|---|
| В бланке нет именованных полей, пустые места — обычный текст | Подстановка промахнётся и оставит «N /26» | Тест на реальном файле шаблона: номер, покупатель, сохранённый блок поставщика |
| Лист юриста допишут после импорта | Два источника одного seq | Новые номера в программе включать только когда лист перестаёт быть местом выдачи; перед этим импорт повторить |
| Две живые строки одного клиента в листе | Нарушение «один договор» | Импорт не связывает сам; человек выбирает строку до связки |
| Картинки карточек разноформатные | Пустая форма при плохом чтении | Номер только после подтверждения; ручной ввод всегда рядом |

## Open Questions

Нет. Включение выдачи номеров на живой базе — после повторного импорта листа, это шаг эксплуатации, не развилка спеки.
