# План: PDF спецификации кнопкой рядом с Excel

**Дата:** 2026-10-01  
**Статус:** реализовано, коммита нет  
**Спека:** `ai_docs/specs/specification-pdf.md`  
**Продолжает:** `ai_docs/specs/specification-print-form.md`

## Overview

В панели спецификации после «Скачать Excel» появляется «Скачать PDF». Книга по-прежнему собирается `build_specification_xlsx`. Те же байты LibreOffice переводит в PDF тем же вызовом `soffice`, что уже собирает PDF договора. `GET .../specification/file` без `format` остаётся xlsx. Полка «Документы» и бланк № 650 не меняются.

## Architecture Decisions

- **Один бланк.** PDF — конвертация только что собранной книги. Reportlab и второй макет не используются. `core/specification_xlsx.py` и `templates/specification.xlsx` не трогаем.
- **Один вызов soffice.** Общий помощник рядом с `convert_docx_to_pdf` пишет исходник во временный каталог и читает `document.pdf`. Договорный docx вызывает тот же помощник с суффиксом `.docx`. Второй способ запускать `soffice` не появляется.
- **Формат на том же пути.** `download_specification(..., file_format="xlsx"|"pdf")`. По умолчанию книга. PDF меняет только байты и расширение имени: `Спецификация КП {id}.pdf` или `Спецификация по счету {номер}.pdf`.
- **503 без файла.** Нет `soffice`, ненулевой код или таймаут — `SupplyContractPdfError` и HTTP 503 с текстом «Не удалось собрать PDF. На сервере недоступен LibreOffice.» Тело — JSON, не xlsx.
- **Кнопка только в панели.** Тот же `submit`, что у Excel: проверка черновика, сохранение, затем скачивание. У PDF свой колбэк с `format=pdf`. «Скачать спецификацию» на полке по-прежнему качает xlsx без `format`.
- **Calc в образе, без пересборки.** В `docker/backend/Dockerfile` рядом с writer добавляется `libreoffice-calc-nogui`. Образ в этой задаче не собирается.
- **Тесты без живого soffice.** Сервис подменяет `convert_xlsx_to_pdf`. HTTP использует `_block_soffice` / `_stub_soffice`.

## Dependency graph

```
общий вызов soffice (docx не меняется)
    │
    └── download_specification(file_format)
            │
            ├── GET .../specification/file?format=
            │
            └── кнопка «Скачать PDF» → save → download format=pdf
```

Бланк, полка документов и папка обмена в граф не входят.

## Risks

- `soffice` для xlsx без Calc молча не делает PDF. В образе нужен `libreoffice-calc-nogui`. На машине разработчика образ не пересобирается, pytest конвертер не запускает.
- Имя PDF должно совпадать со стволом Excel. Расширение меняется в `_specification_filename`, а не второй строкой в эндпоинте.
- Подмена только `subprocess.run` без `document.pdf` уже принята в тестах договора. Новый конвертер читает тот же `document.pdf`.

## Task List

### Phase 1 — Конвертер и сервис

- [x] **SP-001: Общий soffice для xlsx**
  - **Description:** `convert_docx_to_pdf` и новый `convert_xlsx_to_pdf` зовут один помощник. Флаги, таймаут 60 с и текст ошибки те же. Docx по-прежнему пишет `document.docx`.
  - **Acceptance:**
    - [x] Нет `soffice` — `SupplyContractPdfError` с текстом про LibreOffice
    - [x] Успешный прогон читает `document.pdf`
    - [x] Договорный путь docx не меняет аргументы soffice
  - **Verify:** существующие тесты конвертера в `tests/test_archive_endpoints.py`
  - **Dependencies:** —
  - **Files:** `app/services/supply_contract_service.py`
  - **Scope:** S

- [x] **SP-002: PDF из сохранённой книги**
  - **Description:** `download_specification` принимает `file_format`. `pdf` отдаёт байты конвертера и имя с `.pdf`. Без сохранённой спецификации файл не создаётся и конвертер не вызывается.
  - **Acceptance:**
    - [x] Конвертер получает байты той же книги, что xlsx
    - [x] Имя заканчивается на `.pdf` и совпадает со стволом Excel
    - [x] Несохранённая спецификация не пишет файл
  - **Verify:** `pytest tests/test_specification_archive.py -q`
  - **Dependencies:** SP-001
  - **Files:** `app/services/archive_service.py`, `tests/test_specification_archive.py`
  - **Scope:** S

### Checkpoint: Phase 1

- [x] Сервисные тесты спецификации зелёные
- [x] Скачивание не пишет каталог обмена

### Phase 2 — HTTP и образ

- [x] **SP-003: query format**
  - **Description:** `GET /{kp_id}/specification/file` принимает `format=xlsx|pdf`, по умолчанию xlsx. PDF — `application/pdf`. Ошибка конвертации — 503, тело не xlsx. Прежний Content-Type книги сохраняется.
  - **Acceptance:**
    - [x] Без `format` — xlsx и прежний Content-Type
    - [x] `format=pdf` — `application/pdf`, тело от заглушки soffice
    - [x] Нет soffice — 503, тело не начинается с `PK`
  - **Verify:** `pytest tests/test_archive_endpoints.py -q`
  - **Dependencies:** SP-002
  - **Files:** `app/api/v1/endpoints/archive.py`, `tests/test_archive_endpoints.py`
  - **Scope:** S

- [x] **SP-004: Calc в Dockerfile**
  - **Description:** В список пакетов добавляется `libreoffice-calc-nogui`. Образ не пересобирается.
  - **Acceptance:**
    - [x] В `docker/backend/Dockerfile` есть `libreoffice-calc-nogui`
  - **Verify:** просмотр Dockerfile
  - **Dependencies:** —
  - **Files:** `docker/backend/Dockerfile`
  - **Scope:** XS

### Checkpoint: Phase 2

- [x] HTTP-тесты формата зелёные
- [x] Без `format` ответ остаётся книгой

### Phase 3 — Панель

- [x] **SP-005: Кнопка «Скачать PDF»**
  - **Description:** Кнопка сразу после «Скачать Excel». Тот же `submit`. Колбэк сохраняет черновик и качает `format=pdf`. Полка «Скачать спецификацию» по-прежнему xlsx.
  - **Acceptance:**
    - [x] В панели три кнопки: «Сохранить», «Скачать Excel», «Скачать PDF»
    - [x] PDF передаёт тот же черновик, что Excel
    - [x] «Скачать спецификацию» качает xlsx
  - **Verify:** `cd frontend && npm run test -- src/features/commercial-archive/components/SpecificationPanel.test.tsx src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx` и `npm run typecheck`
  - **Dependencies:** SP-003
  - **Files:** `SpecificationPanel.tsx`, `SpecificationPanel.test.tsx`, `OfferDetailsDrawer.tsx`, `OfferDetailsDrawer.test.tsx`, `archiveApi.ts`, `useArchiveQueries.ts`
  - **Scope:** S

### Checkpoint: Complete

- [x] Критерии спеки отмечены
- [x] Указанные pytest и vitest зелёные
- [x] Коммита нет

## Not in this plan

- PDF на полке «Документы»
- Отдельный макет PDF и Reportlab
- Хранение PDF в базе
- Правка бланка `templates/specification.xlsx`
- Пересборка Docker-образа
- Запись файла в папку обмена
