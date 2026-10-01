# План: Соглашение об ЭДО вместе с договором

**Дата:** 2026-09-30  
**Статус:** implemented (no commit)  
**Спека:** `ai_docs/specs/edo-agreement.md`  
**Идея:** `ai_docs/ideas/edo-agreement.md`

## Overview

Подтверждённый договор уже хранит покупателя. К нему добавляется необязательный основной ОКВЭД с карточки. Из тех же полей собирается соглашение об ЭДО. В карточке договора две строки, у каждой Word и PDF. PDF — конвертация только что собранного docx через LibreOffice. Образ бэкенда получает `libreoffice-writer-nogui`, иначе на сервере PDF нечем собрать.

## Architecture Decisions

- **ОКВЭД — колонка договора.** `supply_contract.okved TEXT NULL`. Существующая база не обновится одним `CREATE TABLE IF NOT EXISTS`: рядом `ALTER TABLE`, ошибка «колонка уже есть» глотается, как у других колонок `kp_db_schema`. Пустая строка с формы хранится как `NULL`.
- **Код, не описание.** Разбор берёт первый `\d{2}(?:\.\d{1,2}){1,3}` после подписи «ОКВЭД» и останавливается перед «Доп.». В `_HOLE_FIELDS` поле не входит. В промпт зрения добавляется ключ `okved`; повтор по сорванным цифрам его не трогает.
- **Соглашение — бланк с плейсхолдерами.** Текст пунктов лежит в `templates/edo_agreement.docx`. Код подставляет дату, преамбулу стороны-2 и реквизиты. Пункты в Python не собираются. Год берётся из `contract_date`.
- **Преамбула ИП отдельно.** ООО и АО идут через `build_preamble` с именем стороны «Сторона-2». ИП: «на основании ОГРНИП» и `ogrn`, даже если в договоре основание «устав».
- **PDF не второй макет.** Сервис пишет docx во временный файл и вызывает `soffice --convert-to pdf` с таймаутом. Это не `GsmExportService`. Юнит-тесты конвертер подменяют. Нет `soffice` или ненулевой код — 503, тело не docx.
- **Старый URL договора.** `GET .../supply-contract/document` без `format` по-прежнему docx. `format=pdf` конвертирует те же байты.
- **Импорт.** Соглашение требует непустые вид, полное имя, подписанта, ИНН, ОГРН, юридический адрес и почту. Иначе 400 с текстом из спеки. Скачивание договора для импорта не меняется.

## Dependency graph

```
разбор основного ОКВЭД
    │
    └── колонка и сохранение в договоре
            │
            ├── бланк соглашения и подстановка
            │       │
            │       └── PDF через soffice + оба GET
            │               │
            │               └── карточка: поле ОКВЭД и четыре файла
            │
            └── LibreOffice в образе бэкенда
```

Образ и карточка после GET не зависят друг от друга. Их можно делать в любом порядке после PDF-сервиса.

## Risks

- `soffice` в синхронном скачивании при `UVICORN_WORKERS=1` занимает процесс. Договор короткий. Таймаут 60 секунд, отдельный временный каталог на вызов, docx этот путь не ждёт.
- Образ раздуется пакетом Writer. Без него PDF на сервере — 503, Word жив. Это принятое поведение, не запасной docx.
- Длинное имя ООО сдвинет строки бланка. Это нормально: бланк переносится, а не штамп поверх PDF Бисерова.
- `CREATE TABLE` на уже существующей базе колонку не добавит. Без `ALTER` сохранение ОКВЭД упадёт только на старом файле SQLite.

## Task List

### Phase 1 — ОКВЭД в договоре

- [x] **ED-001: Основной код с карточки**
  - **Description:** Текстовый разбор кладёт в `okved` первый код после «ОКВЭД» и не берёт хвост «Доп.». Нет подписи — пусто и не `doubtful`. Подпись без кода — `doubtful`. Ключ добавлен в `FIELD_NAMES` и в JSON промпта зрения. Поле не в `_HOLE_FIELDS`.
  - **Acceptance:**
    - [x] Текст карточки ПСК даёт `23.61` и не содержит `52.29`
    - [x] Карточка без строки ОКВЭД не помечает поле сомнительным
    - [x] Подпись «ОКВЭД» без кода попадает в `doubtful`
  - **Verify:** `pytest tests/test_supply_contract_card_parse.py tests/test_supply_contract_card_ocr.py -q`
  - **Dependencies:** —
  - **Files:** `core/supply_contract_card_parse.py`, `core/ocr/prompts.py`, `tests/test_supply_contract_card_parse.py`
  - **Scope:** S

- [x] **ED-002: Колонка и ответ API**
  - **Description:** `okved` в `CREATE TABLE` и в `ALTER TABLE`. Репозиторий пишет и читает её, импорт оставляет `NULL`. Схемы создания, разбора и ответа принимают строку или пусто. Сервис сохраняет обрезку, пустое — `NULL`. Подтверждение без ОКВЭД проходит. `SupplyContractPatch` поле не принимает.
  - **Acceptance:**
    - [x] Повторный старт схемы на базе без колонки её добавляет
    - [x] Создание с `23.61` возвращает этот код; без поля — `null`
    - [x] Пустая строка сохраняется как `null`
  - **Verify:** `pytest tests/test_supply_contract.py tests/test_archive_endpoints.py -q`
  - **Dependencies:** ED-001
  - **Files:** `core/kp_db_schema.py`, `app/repositories/supply_contract_repository.py`, `app/schemas/supply_contract.py`, `app/services/supply_contract_service.py`, `frontend/src/features/commercial-archive/types/supplyContract.ts`
  - **Scope:** M

### Checkpoint: Phase 1

- [x] Договор подтверждается с ОКВЭД и без него. Карточка ПСК в тесте разбора даёт `23.61`

### Phase 2 — Бланк соглашения

- [x] **ED-003: Фразы стороны-2**
  - **Description:** Функции преамбулы и строк реквизитов. ООО и АО — `build_preamble` с «Сторона-2». ИП — ОГРНИП из `ogrn`. КФХ и физлицо — преамбула договора с тем же именем стороны. Реквизиты: ИНН, КПП только у ООО и АО, адрес, почта, почтовый адрес если отличается, строка ОКВЭД только при коде, оператор покупателя только если заполнен. Сторона-1 в этих функциях не собирается: она текст бланка.
  - **Acceptance:**
    - [x] ИП содержит «ОГРНИП» и номер, не «Устава»
    - [x] ООО содержит «в лице» и должность
    - [x] Пустой `okved` не даёт строку «ОКВЭД»
    - [x] Заполненные `edo_operator` или `edo_id` дают строку оператора, пустые — нет
  - **Verify:** `pytest tests/test_edo_agreement_docx.py -q`
  - **Dependencies:** ED-002
  - **Files:** `core/edo_agreement_docx.py`, `tests/test_edo_agreement_docx.py`
  - **Scope:** S

- [x] **ED-004: Бланк и подстановка**
  - **Description:** `templates/edo_agreement.docx` хранит пункты и приложение 1 образца от 22.09.2026, колонтитул «Сторона-1 / Сторона-2» и константы стороны-1 из спеки, включая почты и Тензор. Плейсхолдеры даты, преамбулы и реквизитов стороны-2 заполняет `render_edo_agreement_docx`. В файле бланка нет «Бисеров». Год не зашит отдельной веткой «2026».
  - **Acceptance:**
    - [x] Сборка на полях ПСК содержит ИНН `7814192061`, `info@gbi-psk.ru`, `ОКВЭД 23.61` и не содержит «Бисеров»
    - [x] Дата `2027-01-05` печатается с 2027 годом
    - [x] Пустой ОКВЭД не оставляет строку «ОКВЭД» у стороны-2
    - [x] В тексте есть «start@gbkstart.ru» и «Тензор»
  - **Verify:** `pytest tests/test_edo_agreement_docx.py -q`
  - **Dependencies:** ED-003
  - **Files:** `templates/edo_agreement.docx`, `core/edo_agreement_docx.py`, `tests/test_edo_agreement_docx.py`
  - **Scope:** M

### Checkpoint: Phase 2

- [x] Docx соглашения собирается без LibreOffice и без SQLite

### Phase 3 — Скачивание

- [x] **ED-005: PDF из docx**
  - **Description:** Функция конвертации во временном каталоге, таймаут 60 секунд. Нет бинарника или сбой — ошибка сервиса, которую endpoint отдаёт как 503 с русским текстом про LibreOffice. Успех — байты PDF.
  - **Acceptance:**
    - [x] Подменённый успех возвращает байты конвертера и `application/pdf`
    - [x] Отсутствие `soffice` даёт 503 и не отдаёт docx
  - **Verify:** `pytest tests/test_archive_endpoints.py -q -k "supply_contract or edo"`
  - **Dependencies:** ED-004
  - **Files:** `app/services/supply_contract_service.py`, `app/api/v1/endpoints/archive.py`, `tests/test_archive_endpoints.py`
  - **Scope:** M

- [x] **ED-006: Два GET**
  - **Description:** `document?format=pdf` конвертирует текущий docx договора. Без параметра — docx, как сейчас. `GET .../supply-contract/edo-agreement?format=docx|pdf` собирает соглашение. Пустые реквизиты импорта — 400, текст «В договоре не хватает реквизитов покупателя для соглашения об ЭДО». Имена файлов из спеки.
  - **Acceptance:**
    - [x] Старый URL без `format` остаётся docx
    - [x] Соглашение на пустом импорте — 400, файла нет
    - [x] Имя pdf соглашения `soglashenie-edo-1028-09-26.pdf`
  - **Verify:** `pytest tests/test_archive_endpoints.py tests/test_supply_contract_docx.py -q`
  - **Dependencies:** ED-005
  - **Files:** `app/services/supply_contract_service.py`, `app/api/v1/endpoints/archive.py`, `tests/test_archive_endpoints.py`
  - **Scope:** M

- [x] **ED-007: LibreOffice в образе**
  - **Description:** В `docker/backend/Dockerfile` пакет `libreoffice-writer-nogui`. Шрифт с кириллицей уже даёт `fonts-dejavu-core`. Образ в этом плане не пересобирается на машине разработчика, если сборка Docker не нужна для тестов.
  - **Acceptance:**
    - [x] Dockerfile содержит `libreoffice-writer-nogui` и по-прежнему чистит списки apt
  - **Verify:** просмотр `docker/backend/Dockerfile`. Живой `docker build` — только если образ всё равно собирают
  - **Dependencies:** ED-005
  - **Files:** `docker/backend/Dockerfile`
  - **Scope:** S

### Checkpoint: Phase 3

- [x] HTTP-тесты зелёные без установленного LibreOffice: конвертер подменён

### Phase 4 — Карточка

- [x] **ED-008: Поле и четыре файла**
  - **Description:** В форме договора поле «ОКВЭД», не обязательное. У действующего договора вместо «Скачать docx» две строки: договор Word и PDF, соглашение Word и PDF. Ошибка PDF показывается в карточке. Word идёт своим запросом.
  - **Acceptance:**
    - [x] Форма отправляется без ОКВЭД
    - [x] При действующем договоре видны четыре действия и нет кнопки «Скачать docx»
    - [x] Ошибка скачивания PDF видна текстом
  - **Verify:** `cd frontend && npx vitest run src/features/commercial-archive/components/SupplyContractDrawer.test.tsx`
  - **Dependencies:** ED-002, ED-006
  - **Files:** `frontend/src/features/commercial-archive/components/SupplyContractDrawer.tsx`, `frontend/src/features/commercial-archive/components/SupplyContractDrawer.test.tsx`, `frontend/src/features/commercial-archive/api/archiveApi.ts`, `frontend/src/features/commercial-archive/hooks/useArchiveQueries.ts`, `frontend/src/features/commercial-archive/types/supplyContract.ts`
  - **Scope:** M

### Checkpoint: готово к реализации по задачам

- [x] `pytest tests/test_supply_contract_card_parse.py tests/test_edo_agreement_docx.py tests/test_supply_contract.py tests/test_archive_endpoints.py -q`
- [x] vitest карточки договора зелёный
- [x] В соглашении на полях ПСК нет «Бисеров», есть `23.61` и Тензор

## Open Questions

Нет. LibreOffice в образе входит в ED-007.
