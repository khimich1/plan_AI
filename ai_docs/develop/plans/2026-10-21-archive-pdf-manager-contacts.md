# Implementation Plan: Контакты менеджера в PDF и XLSX из архива

**Дата:** 2026-10-21  
**Orchestration:** orch-2026-10-21-12-03-archive-mgr-contacts  
**Статус:** ✅ Реализовано — тесты из спеки зелёные, коммита нет  
**Спека:** `ai_docs/specs/archive-pdf-manager-contacts.md`  
**Идея:** `ai_docs/ideas/archive-pdf-manager-contacts.md`  
**Оценка:** около 2,5 часов

## Overview

Скачивание PDF и XLSX из архива сейчас передаёт в генераторы `manager_phone=None` и `manager_email=None`. Пустая пара печатает заводские `8 (4852) 595 000` и `info@zhbk.ru`. Нужно на момент скачивания найти менеджера по ФИО в справочнике и подставить его телефон и почту. Файл один и для клиента, и для менеджера.

Если имя в КП есть, а карточки нет или у карточки пустые оба контакта, строка шапки пустая. Завод не подставляется. Завод остаётся только когда имени менеджера нет вообще. Если заполнено только одно поле, печатается оно. В КП новые колонки не добавляются. Мастер и `offers_service` не меняются.

Контекст для воркера: скилл `plan-web-context` и эта спека. Не вызывать `resolve_manager_for_hydrate`.

## Architecture Decisions

- Подавление завода — явный `suppress_company_contacts: bool = False` в конце сигнатур `generate_commercial_offer_pdf` и `generate_commercial_offer_xlsx`. Пустая строка по-прежнему значит «контакта нет». Сентинель вроде `"\x00"` не использовать: сегодня оба пустых значения уже заняты смыслом «показать завод», и этот смысл остаётся у мастера и у `offers_service`.
- Архив ставит флаг `True` только когда `manager_name.strip()` непустой. Считает тройку `(phone, email, flag)` один раз до ветки `kind` и передаёт её в PDF и в XLSX. `xlsx_delivery_in_unit` сидит в той же ветке (`archive_service.py`, строка 1023) и получает те же значения. `embed_delivery_in_unit_price` не меняется.
- Поиск — чистые функции в новом `app/services/manager_contact_lookup.py`. Совпадение как у hydrate: `strip` + `casefold`, первая карточка в списке. Возврат — телефон и почта, не id. `resolve_manager_for_hydrate` не вызывать и не менять: при несовпадении он возвращает `(1, "", "")` (`commercial_draft_lifecycle.py`, строка 913).
- Справочник читается через `ManagerRepository.list_managers()` только если `db_path` — существующий файл (`str` или `Path`). Иначе пустой список и никакого `sqlite.connect`. Иначе `test_generate_document_pdf` создаст пустую базу или упадёт на `MagicMock`.
- Телефон в генератор идёт как `contact_number` из карточки. `format_phone` остаётся внутри генератора. Для `79621872265` шапка: `Тел.: +7 (962) 187-22-65, email: zubov@gbkstart.ru`.
- Подпись не трогать: PDF строка 811, XLSX строка 650.
- Ветки «только телефон» и «только почта» не менять по тексту. Перед ними оба значения `strip`, чтобы пробелы не считались контактом.
- Заводские контакты при отсутствии имени — нынешняя ветка `else` при флаге `False`. Архив в этом случае передаёт `None`, `None`, `False`.

## Dependency graph

```
AMC-001 флаг в PDF и XLSX
        \
         +--> AMC-003 подключить архив
        /
AMC-002 поиск контактов без id=1
```

AMC-001 и AMC-002 не делят файлы, их можно делать параллельно. AMC-003 ждёт оба. После AMC-003 система уже отвечает спеке. Отдельной фазы на фронт нет.

## Task List

### Phase 1: Foundation

- [x] **AMC-001: Флаг suppress_company_contacts в генераторах** `(type: feat-be, dependsOn: —)`
  - **Description:** В конец `generate_commercial_offer_pdf` (`core/commercial_offer.py`, сигнатура со строки 331) и `generate_commercial_offer_xlsx` (`core/commercial_offer_xlsx.py`, сигнатура со строки 183) добавить `suppress_company_contacts: bool = False`. В развилке шапки (PDF 513–521, XLSX 251–259) сначала `strip`. Непустые телефон и почта, только телефон и только почта печатаются как сейчас. Оба пустые и флаг `True` — пустая строка в том же слоте шапки. Оба пустые и флаг `False` — `Тел.: 8 (4852) 595 000, Email: info@zhbk.ru`. Подпись не менять. Порядок старых аргументов не менять.
  - **Acceptance criteria:**
    - [x] Совпавшая пара даёт `Тел.: +7 (962) 187-22-65, email: zubov@gbkstart.ru` в PDF и в XLSX; заводских строк нет; телефон в тексте один раз
    - [x] Только телефон и только почта печатают одно поле и не печатают завод
    - [x] Флаг `True` и пустая пара не печатают завод; без флага пустая пара печатает завод
    - [x] Ячейка XLSX под «С уважением,» равна ФИО и не содержит телефон или почту
  - **Verification:**
    - [x] Tests pass: `.venv/bin/python -m pytest tests/test_commercial_offer_header_contacts.py -q`
  - **Dependencies:** None
  - **Files likely touched:**
    - `core/commercial_offer.py`
    - `core/commercial_offer_xlsx.py`
    - `tests/test_commercial_offer_header_contacts.py`
  - **Estimated scope:** M (3 files)
  - **Tests to add:** `test_pdf_header_prints_formatted_phone_and_email`, `test_xlsx_header_prints_formatted_phone_and_email`, `test_pdf_header_phone_only_omits_email_and_company`, `test_xlsx_header_email_only_omits_phone_and_company`, `test_pdf_suppresses_company_when_both_contacts_empty`, `test_xlsx_suppresses_company_when_both_contacts_empty`, `test_pdf_and_xlsx_keep_company_fallback_without_flag`, `test_xlsx_signature_cell_is_manager_name_only`

- [x] **AMC-002: Поиск контактов по ФИО без заглушки id=1** `(type: feat-be, dependsOn: —)`
  - **Description:** Добавить `app/services/manager_contact_lookup.py` с `lookup_manager_contacts`, `header_contacts_for_archive` и `load_managers_if_db_exists` так, как в спеке. Несовпавшее или пустое имя не возвращает контакты другой карточки. `header_contacts_for_archive` возвращает `(None, None, False)` для пустого имени и `(phone or None, email or None, True)` для непустого. Загрузка справочника не создаёт файл базы. `commercial_draft_lifecycle.py` не менять.
  - **Acceptance criteria:**
    - [x] «зубов алексей юрьевич» находит `79621872265` и `zubov@gbkstart.ru` и не возвращает id
    - [x] Чужое ФИО при наличии карточки id=1 в списке возвращает `("", "")`
    - [x] Пустое имя даёт флаг `False`; непустое, даже без карточки, даёт флаг `True`
    - [x] Карточка только с телефоном или только с почтой отдаёт заполненное поле и пустое второе
    - [x] Несуществующий путь и не-путь возвращают `[]` и не создают файл
  - **Verification:**
    - [x] Tests pass: `.venv/bin/python -m pytest tests/test_archive_manager_contact_lookup.py -q`
  - **Dependencies:** None
  - **Files likely touched:**
    - `app/services/manager_contact_lookup.py`
    - `tests/test_archive_manager_contact_lookup.py`
  - **Estimated scope:** S (2 files)
  - **Tests to add:** `test_lookup_matches_fio_casefold_exact`, `test_lookup_unmatched_name_does_not_return_id1_contacts`, `test_lookup_blank_name_returns_empty_pair`, `test_lookup_partial_card_returns_only_filled_field`, `test_header_contacts_suppress_only_when_name_present`, `test_load_managers_missing_file_returns_empty_and_does_not_create_db`

### Checkpoint: Foundation

- [x] Спека принята до старта AMC-001
- [x] AMC-001 и AMC-002 зелёные по своим командам
- [x] `resolve_manager_for_hydrate` не импортирован новым модулем
- [x] Дальше не идти, если пустая пара без флага перестала печатать завод

### Phase 2: Core

- [x] **AMC-003: Подключить контакты в скачивании из архива** `(type: feat-be, dependsOn: AMC-001, AMC-002)`
  - **Description:** В `ArchiveService.generate_document` после чтения `manager_name` (строка 977) один раз вызвать `header_contacts_for_archive(manager_name, load_managers_if_db_exists(self.repository.db_path))`. Подставить результат вместо `manager_phone=None` / `manager_email=None` в вызов PDF (строки 1006–1007) и в вызов XLSX (строки 1031–1032) и передать `suppress_company_contacts`. Ветку `schema` не трогать. `commercial_export_service.py` и `offers_service.py` не редактировать. В тесте архива справочник подменять через monkeypatch `load_managers_if_db_exists`, генераторы можно оставить настоящими только если позиции считаются без живой базы; иначе мокать генератор и проверять kwargs, а текст шапки уже покрыт AMC-001. Фикстура менеджера: id=2, ФИО «Зубов Алексей Юрьевич», телефон `79621872265`, почта `zubov@gbkstart.ru`, плюс карточка id=1 с другими контактами, чтобы поймать заглушку id=1.
  - **Acceptance criteria:**
    - [x] PDF, XLSX и `xlsx_delivery_in_unit` получают телефон `79621872265`, почту `zubov@gbkstart.ru` и `suppress_company_contacts=True`, когда ФИО совпало
    - [x] Несовпавшее ФИО: оба контакта пустые и флаг `True`
    - [x] Пустое имя: `None`, `None`, флаг `False`
    - [x] Частичная карточка: в kwargs уходит только заполненное поле, флаг `True`
    - [x] `test_generate_document_pdf` зелёный и не создаёт sqlite-файл
  - **Verification:**
    - [x] Tests pass: `.venv/bin/python -m pytest tests/test_archive_document_manager_contacts.py tests/test_archive_service.py::test_generate_document_pdf tests/test_archive_pile.py tests/test_archive_march.py -q`
    - [x] Manual check: не открывать живую `plita.db` и не скачивать КП №11 из приложения в рамках задачи. Приёмка на фикстуре Зубова
  - **Dependencies:** AMC-001, AMC-002
  - **Files likely touched:**
    - `app/services/archive_service.py`
    - `tests/test_archive_document_manager_contacts.py`
  - **Estimated scope:** S (2 files)
  - **Tests to add:** `test_archive_pdf_passes_zubov_contacts_and_suppress_flag`, `test_archive_xlsx_and_delivery_in_unit_share_header_contacts`, `test_archive_unmatched_fio_sets_suppress_and_empty_contacts`, `test_archive_blank_manager_name_keeps_company_fallback`, `test_archive_partial_phone_passes_phone_and_suppress`

### Checkpoint: Complete

- [x] Команда из спеки зелёная
- [x] `commercial_export_service.py`, `offers_service.py`, `commercial_draft_lifecycle.py`, фронт и `archive.py` без диффа по этой задаче
- [x] В `KP_offers` не появилось колонок
- [x] Коммит только по отдельной просьбе

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Поиск возьмёт `resolve_manager_for_hydrate` и при чужом ФИО подставит пустые контакты карточки id=1 или её телефон | High | Отдельный модуль. Тест `test_lookup_unmatched_name_does_not_return_id1_contacts` с карточкой Александра в списке |
| Флаг забудут на одном из двух вызовов, XLSX с доставкой в цене останется заводским | High | Одна тройка до `if kind`. Тест сравнивает kwargs PDF, XLSX и `xlsx_delivery_in_unit` |
| Чтение справочника откроет sqlite на пути из теста и создаст файл или уронит `test_generate_document_pdf` | Med | `load_managers_if_db_exists` не коннектится, если пути-файла нет. Регресс-команда это проверяет |
| `strip` изменит шапку мастера, если в метаданных телефон из пробелов | Low | Для непустых цифр ветки те же. Пустая пара без флага по-прежнему завод. Это покрывает `test_pdf_and_xlsx_keep_company_fallback_without_flag` |
| Менеджера переименовали после сохранения КП | Low | Принято: шапка пустая, не заводская. В КП снимок не пишем |

## Open Questions

Нет. Решение пользователя: при несовпавшем ФИО строка контактов пустая; завод только если имени нет вообще.

Вне этой работы, не чинить сейчас: `offers_service.generate_pdf` и `generate_xlsx` по-прежнему передают `None` (`app/services/offers_service.py`, строки 195–196 и 219–220). Фронт архива этот API не вызывает.

## Boundaries

- Always: тесты из команд задач. Живую базу не открывать. Флаг передаёт только архив.
- Ask first: снимок контактов на КП, правка мастера, правка `offers_service`, нечёткий поиск ФИО.
- Never: начать код до приёмки спеки и этого плана. Коммитить без отдельной просьбы.
