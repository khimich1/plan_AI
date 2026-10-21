# Implementation Plan: ГСМ — табель Т-13 и допуск водителей в путевой

Дата: 2026-10-12. Статус: draft, к реализации.
Спека: [`../../specs/gsm-timesheet-driver-eligibility.md`](../../specs/gsm-timesheet-driver-eligibility.md).
Идея: [`../../ideas/gsm-timesheet-driver-eligibility.md`](../../ideas/gsm-timesheet-driver-eligibility.md).

## Overview

Бухгалтер загружает Т-13 за месяц; generate/bulk ставит в ПЛ только
допущенных (Я/В) и свободных на дату; запасные ротируются по счётчику
подстановок; три стопа (дубль человека в день, лист в отсутствии, день без
людей) на generate/create/patch/export. Без табеля generate и комплект
месяца закрыты.

## Architecture Decisions

- **Обратная совместимость `generate()`.** Сигнатуру не ломаем:
  `driver_id` остаётся (primary). Новые опциональные
  `drivers_by_date: Mapping[date, tuple[DriverCandidate, ...]] | None` и
  `substitute_counts: Mapping[int, int] | None`. `None` → легаси-режим
  (везде primary) — все существующие тесты генератора зелёные без правок.
  Снято с «Ask first» спеки: контракт расширяется аддитивно.
- **`DriverCandidate`** — frozen DTO в `core/gsm/driver_pool.py`:
  `(driver_id, code)`. Порядок в кортеже даты уже финальный (сервис
  отсортировал: primary → без-машинные → чужие; внутри группы по
  `(substitute_count, personnel_num)`); генератор берёт head, не сортирует.
- **Занятость считает сервис, не генератор.** Для каждой машины перед
  generate: `busy[date]` = все `gsm_waybill` на дату других машин (любой
  статус; свои дни внутри периода исключены — они перезапишутся). Bulk идёт
  последовательно в порядке запроса; после персиста машины следующая
  пересчитывает пул. Детерминизм = порядок входа.
- **Счётчик подстановок:** seed из БД (листы периода, где `driver_id` ≠
  primary той машины, по водителю) + локальный инкремент в генераторе по
  мере выбора. Окно = генерируемый период.
- **Гейт табеля — в generation service**, рядом с `evaluate_vehicle`:
  покрытие периода батчами, иначе 422 `gsm_timesheet_missing`.
- **Export-валидация — в export service до soffice**, на весь комплект:
  пары `(driver, date)` × машины, коды Я/В, покрытие табелем. Ловит и
  перезагрузку табеля после генерации.
- **Парсер Т-13 — pure в `core/gsm/timesheet.py`** (openpyxl workbook →
  DTO). Период из имени файла (`от DD.MM.YYYY`) или параметра `month`
  (UI fallback). Неизвестные коды сохраняются как есть и считаются запретом.
- **Фикстура Т-13 строится openpyxl в тесте** (helper), не бинарь в git.
  Реальные № 44/47 — только ручная приёмка на копии, не коммитим.
- **Шишов 241→438** — ручная правка справочника до приёмки (шаг в
  acceptance, не код).

## Task List

### Phase 1: Ядро (pure, TDD)

- [x] **Task 1: Парсер Т-13**
  - **Description:** `core/gsm/timesheet.py`: `parse_personnel_number`
    (`СЭ00-00470→470`, `СЭ0В-00001→1`, мусор→None), `parse_timesheet_xlsx`
    (строки: имя C, табельный E, сетка `I,K,M,N,P,R,T,V,X,Z,AB,AD,AF,AH,AK`
    по строкам r/r+2, день 31 из `AL` r+2; часы пропускаем),
    `period_from_filename`. DTO `TimesheetRow`, `ParsedTimesheet`
    (period, rows, unmatched). Коды сохраняем сырыми.
  - **Acceptance:**
    - [ ] номера: `СЭ00-00143→143`, `СЭ0В-00001→1`, `СЭ00-00438→438`, ``→None
    - [ ] июль-фикстура: 31 день у строки; `ОТ` Скрябина 1–9 на месте;
          часы не попадают в коды
    - [ ] неизвестный код (`Б`) сохраняется и не трактуется как Я/В
    - [ ] период из `... от 31.07.2026 (1).xlsx` → июль 2026; мусор → None
  - **Verification:** `.venv/bin/python -m pytest tests/test_gsm_timesheet.py -q`
  - **Dependencies:** нет
  - **Files:** `core/gsm/timesheet.py`, `tests/test_gsm_timesheet.py`
  - **Scope:** M

- [x] **Task 2: Пул водителей (выбор дня)**
  - **Description:** `core/gsm/driver_pool.py`: `DriverCandidate`,
    `ELIGIBLE_CODES = {"Я","В"}`, `order_candidates(primary_id, candidates,
    counts)` (группы + ранг `(count, personnel_num)`),
    `pick_driver(day_candidates) → (driver_id, warnings)` (`В` →
    `weekend_driver`; пусто → `no_available_driver`).
  - **Acceptance:**
    - [ ] primary свободен → он, без warning; primary занят/ОТ → head группы
    - [ ] ротация: счётчики 0/0 → меньший табельный; после инкремента —
          другой; серия при отпуске одного
    - [ ] `В` у выбранного → `weekend_driver`; `ОТ` не пропускается в кандидаты
    - [ ] пустой кортеж → `no_available_driver`
  - **Verification:** `.venv/bin/python -m pytest tests/test_gsm_driver_pool.py -q`
  - **Dependencies:** нет (параллельно T1)
  - **Files:** `core/gsm/driver_pool.py`, `tests/test_gsm_driver_pool.py`
  - **Scope:** S

- [x] **Task 3: Таблицы + репозиторий**
  - **Description:** `core/kp_db_schema.py`: `gsm_timesheet_batch`,
    `gsm_timesheet_day` + индекс `(driver_id, date)`. `GsmRepository`:
    insert_batch, replace_days(period, rows), list_days(period),
    list_coverage(), busy_drivers_by_date(period, exclude_vehicle_id),
    substitute_counts(period) (листы, где driver ≠ primary машины).
  - **Acceptance:**
    - [ ] replace_days: повторный импорт месяца затирает старые дни, батчи
          оба в истории
    - [ ] busy: лист Monjaro 12.07 занимает Лоншакову для других машин,
          свои дни Monjaro её не занимают
    - [ ] substitute_counts: день Поляковой на 952 считается, её же день на
          «своей» (нет) — нет
  - **Verification:** `.venv/bin/python -m pytest tests/test_gsm_repository.py -q`
  - **Dependencies:** нет (параллельно T1/T2)
  - **Files:** `core/kp_db_schema.py`, `app/repositories/gsm_repository.py`,
    `tests/test_gsm_repository.py`
  - **Scope:** M

### Phase 2: Импорт табеля (API)

- [x] **Task 4: Сервис импорта + endpoints**
  - **Description:** `GsmTimesheetService.import_xlsx(file, month=None)`:
    парсер → стыковка по `personnel_number` с `gsm_driver` (активные) →
    batch + replace_days → отчёт (period, rows_total, matched,
    unmatched[{full_name, personnel_number}], codes_summary). Период из
    имени файла, иначе `month` обязателен (400). `POST /timesheets/import`,
    `GET /timesheets?period=YYYY-MM` (сетка по сопоставленным). Схемы
    Pydantic, `extra="forbid"`.
  - **Acceptance:**
    - [ ] июль-файл: 5 сопоставленных, офисные — в unmatched, дни записаны
    - [ ] без имени-периода и без month → 400; чужая роль → 403
    - [ ] GET возвращает сетку (driver, date, code) за месяц
    - [ ] повторный импорт того же месяца — дни обновились
  - **Verification:** `.venv/bin/python -m pytest tests/test_gsm_timesheet_api.py -q`
  - **Dependencies:** T1, T3
  - **Files:** `app/services/gsm_timesheet_service.py`,
    `app/api/v1/endpoints/gsm.py`, `app/schemas/gsm.py`,
    `tests/test_gsm_timesheet_api.py`
  - **Scope:** M

### Phase 3: Солвер и сервис генерации

- [x] **Task 5: Генератор — водитель дня**
  - **Description:** `generate()` принимает `drivers_by_date` /
    `substitute_counts` (опционально). `_emit_day`/`_emit_maybe_manual`:
    водитель из пула дня (head), `В` → `weekend_driver`, пусто →
    `no_available_driver` + fallback primary. Выбор подстановки инкрементит
    локальный счётчик. Легаси-вызовы (без пула) — поведение прежнее.
  - **Acceptance:**
    - [ ] с пулом: каждый день — head кандидатов даты; `В` → жёлтый код
    - [ ] пустая дата → primary + `no_available_driver` в warnings дня
    - [ ] подстановка дня N сдвигает выбор дня N+k (чередование)
    - [ ] все существующие тесты генератора зелёные без правок
  - **Verification:** `.venv/bin/python -m pytest tests/test_gsm_generator.py -q`
  - **Dependencies:** T2
  - **Files:** `core/gsm/generator.py`, `core/gsm/__init__.py`,
    `tests/test_gsm_generator.py`
  - **Scope:** M

- [x] **Task 6: Сервис генерации — гейт, пул, bulk**
  - **Description:** `GsmGenerationService.generate`: гейт покрытия
    (422 `gsm_timesheet_missing`), сбор `drivers_by_date` (коды + busy,
    ранги), seed счётчиков. `generate_bulk`: последовательно по машинам;
    после персиста — пересчёт пула следующей. Водитель в `upsert` — из
    `WaybillDay` (уже выбран).
  - **Acceptance:**
    - [ ] без табеля за период → 422 `gsm_timesheet_missing`
    - [ ] bulk 848+952: Кулигин на каждую дату максимум в одном ПЛ;
          вторая Tugella — запасный по ротации
    - [ ] отпуск Лоншаковой 27.07–09.08: в её днях Monjaro — другой человек
    - [ ] одиночный generate после bulk: занятость видна из БД
  - **Verification:**
    `.venv/bin/python -m pytest tests/test_gsm_generation_api.py tests/test_gsm_generate_bulk_api.py -q`
  - **Dependencies:** T3, T5
  - **Files:** `app/services/gsm_generation_service.py`,
    `tests/test_gsm_generation_api.py`, `tests/test_gsm_generate_bulk_api.py`
  - **Scope:** L

### Phase 4: Стопы вне generate

- [x] **Task 7: create/patch/export-валидация**
  - **Description:** `create_waybill`/`patch_waybill`: занят другой
    машиной → 409 `gsm_driver_busy` (дата + госномер в details); табель за
    месяц есть и код не Я/В → 409 `gsm_driver_absent`; табеля нет — код не
    проверяем (generate и так закрыт). `GsmExportService.export_zip`: до
    soffice по комплекту — дубли `(driver,date)` × машины, коды не Я/В →
    409 `gsm_driver_rule_violation` со списком; нет покрытия →
    409 `gsm_timesheet_missing`.
  - **Acceptance:**
    - [ ] PATCH посадить Кулигина на дату, где он уже в 848 → 409 busy
    - [ ] PATCH посадить Лоншакову на её `ОТ` → 409 absent
    - [ ] комплект с парой (подложить в БД) → 409 до soffice, тексты с датами
    - [ ] комплект месяца без табеля → 409 `gsm_timesheet_missing`
  - **Verification:**
    `.venv/bin/python -m pytest tests/test_gsm_waybill_edit.py tests/test_gsm_export.py -q`
  - **Dependencies:** T3 (после T6 — единый стиль ошибок)
  - **Files:** `app/services/gsm_generation_service.py`,
    `app/services/gsm_export_service.py`, `tests/test_gsm_waybill_edit.py`,
    `tests/test_gsm_export.py`
  - **Scope:** M

### Checkpoint A (backend)

- [ ] `pytest tests/ -q` зелёный целиком

### Phase 5: Frontend

- [x] **Task 8: Загрузка и сетка табеля**
  - **Description:** `gsmApi.importTimesheet` (multipart, опц. `month`),
    `getTimesheetGrid`. `TimesheetImportDialog` (drop-зона как
    `TransactionsImportDialog`, fallback-выбор месяца, отчёт: сопоставленные,
    строки без водителя). `TimesheetGridView`: водители × дни, коды,
    подсветка `ОТ`. Карточка «Табель рабочего времени» в
    `GsmRegistriesView`: кнопка импорта, список батчей, сетка.
  - **Acceptance:**
    - [ ] загрузка файла → отчёт с 5 сопоставленными и офисными в unmatched
    - [ ] имя без даты → виден выбор месяца; без него — 400 с текстом
    - [ ] сетка июля: `ОТ` Скрябина 1–9 подсвечен
  - **Verification:**
    `cd frontend && npx vitest run src/features/gsm && npm run build`
  - **Dependencies:** T4
  - **Files:** `frontend/src/features/gsm/api/gsmApi.ts`,
    `components/TimesheetImportDialog.tsx`, `components/TimesheetGridView.tsx`,
    `components/GsmRegistriesView.tsx`, `hooks/useGsmQueries.ts`,
    `types/gsm.ts`, тесты рядом
  - **Scope:** M

- [x] **Task 9: Warnings, exportGate, тексты ошибок**
  - **Description:** `waybillWarnings.ts`: `weekend_driver` («Свой
    выходной», мягкий), `no_available_driver` («Нужен водитель», жёсткий).
    `exportGate.ts`: HARD += `no_available_driver`, SOFT +=
    `weekend_driver`. `gsmErrors.ts`: тексты `gsm_driver_busy`,
    `gsm_driver_absent`, `gsm_driver_rule_violation`,
    `gsm_timesheet_missing`. Drawer показывает 409-текст сервера.
  - **Acceptance:**
    - [ ] день только с `weekend_driver` не блокирует zip (confirm)
    - [ ] день с `no_available_driver` блокирует zip, причина читается
    - [ ] 409 busy в drawer: «Кулигин уже в О 848 ХР 44 на 12.07»
  - **Verification:**
    `cd frontend && npx vitest run src/features/gsm/lib src/features/gsm/components/WaybillDayDrawer.test.tsx`
  - **Dependencies:** нет (параллельно T8; контракт кодов — спека)
  - **Files:** `frontend/src/features/gsm/lib/waybillWarnings.ts`,
    `lib/exportGate.ts`, `lib/gsmErrors.ts`,
    `components/WaybillDayDrawer.tsx`, тесты рядом
  - **Scope:** S

### Checkpoint B (frontend)

- [ ] `cd frontend && npm test -- --run && npm run build` зелёные

### Phase 6: Приёмка на копии БД

- [ ] **Task 10: Июль + август на копии `plita.db`**
  - **Description:** Копия БД; ручная правка Шишов 241→438; импорт Т-13
    № 44 и № 47 через API; bulk generate июль и август по 4 машинам;
    сверка SC-1…SC-6; комплект июля и августа (zip) — закрывается,
    в комплекте ноль пар и ноль `ОТ`.
  - **Acceptance:**
    - [ ] SC-2/SC-3 из спеки: ротация на второй Tugella, отпуска пусты
    - [ ] `no_available_driver` = 0 на буднях; суб/вс-якоря — `weekend_driver`
    - [ ] экспорт комплектов без `gsm_driver_rule_violation`
    - [ ] живую `plita.db` не трогаем до зелёной копии
  - **Verification:** чек-лист в implementation report + скрин журнала
  - **Dependencies:** T1–T9
  - **Files:** report `ai_docs/develop/reports/2026-10-XX-gsm-timesheet-acceptance.md`
  - **Scope:** S

## Risks

- **Хрупкость формы Т-13** (другие юрлица, смещённые колонки): парсер
  валидирует шапку («ТАБЕЛЬ», колонки дней), несовпадение → 400 с текстом,
  не молчаливый сдвиг.
- **Перезагрузка табеля после generate:** дни уже стоят людьми, новый
  табель делает кого-то `ОТ` → export-валидация (T7) красным комплект;
  бухгалтер перегенерирует месяц. Осознанно, не баг.
- **Две загрузки одного месяца подряд в bulk:** replace_days
  идемпотентен по периоду; порядок батчей не влияет.
- **Производительность пула:** 4 машины × ~22 будня × 5 водителей —
  пересчёт на машину дешёвый, без оптимизаций.

## Out of Scope (этапом позже)

- UNIQUE `(driver_id, date)` в SQLite (после разведения якоря 30.06).
- Фильтрованный список допущенных водителей в drawer (сейчас 409 + текст).
- Скользящее окно счётчика ротации («по году»).
- Совет «отдайте пятничную заправку другой машине» (отдельная идея).
