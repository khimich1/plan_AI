# План: Контур GUID и цен — раунд 2 (ворота счёта, предупреждения, вкладка «Прайсы и 1С»)

**Дата:** 2026-09-17
**Статус:** план на согласовании (не приступать к реализации)
**Идея:** `ai_docs/ideas/guid-price-sync-loop.md`
**Спека:** `ai_docs/specs/guid-price-sync-round2.md`
**Предыдущий план:** `ai_docs/develop/plans/2026-09-15-guid-price-sync.md` (GPS-001…009 ✅, GPS-010 ⏳ → расширен в GPS-011…GPS-028)

## Контекст

Раунд 1 закрыт: `nomenclature_guid` + bootstrap 271 GUID, импорт .xls «Прайс-лист» через
диалог «Выгрузка 1С», Excel-очередь и импорт цен из неё, гигиена плит. Осталось:
GPS-010 (блок счёта без GUID) + решения раунда 2 (17.09): контур предупреждений
(шаг 3 → архив → вкладка → ворота), вкладка «Прайсы и 1С» как единое место работы
экономиста, парсер «Универсального отчёта» .xlsx с частичными выгрузками,
персистентные очереди 💰/⚠️ и единый резолв дублей (включая плиты через override).

## Architecture Decisions

- **Единый резолвер GUID** (`core/guid_gate.py`) для всех 6 видов продукции; используется
  превью-проверкой шага 3, уведомлением при архиве и будущими воротами счёта.
- **«у»-срез по суффиксу марки** — той же функцией, что линт-парити (D11): одно место правды.
- **Плиты не мигрируют** в `nomenclature_guid`; их выбор при дубле — override-таблица
  `plate_guid_choice`, `prays_plity` не правим.
- **Очереди живут в БД** (`price_queue`, `duplicate_candidates`), наполняются при каждом
  импорте 1С; Excel-скрипты остаются fallback и переиспользуют общий `core/guid_queue.py`.
- **Два формата 1С** (.xls полный / .xlsx с отборами): автодетект, у .xlsx partial-режим
  без `disappeared`.
- **Цена 💰 через ту же apply-функцию**, что Excel-импорт (одна семантика записи в
  прайс-таблицы); после записи GUID сразу в `nomenclature_guid` (auto).
- **Уведомления** — существующая система (`notifications` + колокольчик), тип
  `kp_guid_missing`, фан-аут economist.

## Task List

### Phase 1 — Хранилища и резолвер (backend core)

- [x] **GPS-011: Таблица `plate_guid_choice` + модуль**
  - **Description:** `core/plate_guid_choice.py`: `ensure_schema()` (идемпотентно),
    `set_choice / get_choice / clear_choice`, поля `(plate_name_norm PK, chosen_guid,
    chosen_name, decided_by, decided_at, note)`. Нормализация имени плиты — как в
    `lookup_nomenclature_by_plate_name`.
  - **Acceptance:** set→get возвращает выбор; повторный ensure_schema не падает;
    нормализация «ЛВ 60.12-4» ≡ «ЛВ60.12-4».
  - **Verify:** `pytest tests/test_plate_guid_choice.py -q`
  - **Dependencies:** —
  - **Files:** `core/plate_guid_choice.py`, `tests/test_plate_guid_choice.py`
  - **Scope:** S

- [x] **GPS-012: `duplicate_candidates` для не-плит**
  - **Description:** таблица `(kind, mark_norm, guid, name, price, first_seen, active,
    PK(kind, mark_norm, guid))`; `sync_pricelist` при `ambiguous` пишет кандидатов
    (с ценой/именем из файла 1С); helpers `list_candidates()`, `clear_candidates(kind, mark)`.
  - **Acceptance:** после импорта файла с дублем кандидаты видны; повторный импорт
    идемпотентен; `clear_candidates` вызывается резолвом (GPS-021).
  - **Verify:** `pytest tests/test_nomenclature_sync.py tests/test_duplicate_candidates.py -q`
  - **Dependencies:** —
  - **Files:** `core/duplicate_candidates.py`, `core/nomenclature_sync.py`,
    `tests/test_duplicate_candidates.py`, `tests/test_nomenclature_sync.py`
  - **Scope:** M

- [x] **GPS-013: Персистентная очередь `price_queue` (💰)**
  - **Description:** `core/price_queue_db.py`: `(guid PK, mark_norm, kind, name, state
    open|resolved, resolved_by, resolved_price, resolved_at, first_seen)`; наполнение из
    `unmatched_1c` в `nomenclature_import_service` (upsert, не сбрасывает resolved);
    `list_open()`.
  - **Acceptance:** импорт файла с изделием без нашей цены → open-задача; повторный
    импорт не дублирует; resolved не воскресает.
  - **Verify:** `pytest tests/test_price_queue.py -q`
  - **Dependencies:** —
  - **Files:** `core/price_queue_db.py`, `app/services/nomenclature_import_service.py`,
    `tests/test_price_queue.py`
  - **Scope:** M

- [x] **GPS-014: Резолвер `core/guid_gate.py`**
  - **Description:** `check_invoice_guids(order_lines, conn) -> GateReport(ready, missing)`;
    не-плиты — `get_guid_for_invoice` (статусы auto|manual), суффикс «у» → `guid_1c_u`
    (срез общей с линтом функцией); плиты — `plate_guid_choice` →
    `lookup_nomenclature_by_plate_name`; дубль без выбора/ambiguous/missing → `missing`
    с `reason` + `action_hint`; custom — без исключений. `assert_invoice_guids` →
    `InvoiceGuidBlockError(report)`.
  - **Acceptance:** все 6 видов; «у» без `guid_1c_u` → missing; плита-дубль без выбора →
    missing, с выбором → ready; custom без GUID → missing.
  - **Verify:** `pytest tests/test_guid_gate.py -q`
  - **Dependencies:** GPS-011
  - **Files:** `core/guid_gate.py`, `tests/test_guid_gate.py`
  - **Scope:** M

### Checkpoint: Phase 1
- [x] Все unit-тесты core зелёные; резолвер покрыт по видам и статусам.

### Phase 2 — Форматы 1С и API

- [x] **GPS-015: Парсер «Универсального отчёта» .xlsx**
  - **Description:** `core/universal_report_1c_parser.py`: плавающая шапка (поиск строки
    по «Номенклатура»/«Номенклатура.Ссылка»), колонки по именам («Ссылка», «Код»,
    «Номенклатура», «Количество», «Вес», «Объем»); детект блока отборов над шапкой →
    `partial=True`; строки → записи с guid/code/name/qty/weight/volume; невалидный файл →
    ValueError «Похоже, это не универсальный отчёт 1С». Тесты на синтетических fixture
    (полный, с отбором, без шапки).
  - **Acceptance:** парсит реальный образец (структуру — из образца 17.09); partial-детект
    работает; ошибка понятная.
  - **Verify:** `pytest tests/test_universal_report_1c_parser.py -q`
  - **Dependencies:** —
  - **Files:** `core/universal_report_1c_parser.py`, `tests/test_universal_report_1c_parser.py`
  - **Scope:** M

- [x] **GPS-016: Интеграция форматов в import-1c**
  - **Description:** `nomenclature_import_service` выбирает парсер по расширению/сигнатуре
    (.xls → старый, .xlsx → новый); детект вида продукции по маркам (regex семейств),
    неоднозначность → 400 с вопросом; `partial=True` → пропуск логики `disappeared`;
    `Import1cResponse` + поле `mode: full|partial` (+ `weights_updated` для GPS-017).
  - **Acceptance:** старый .xls работает как раньше (regress); .xlsx с отбором не пишет
    `disappeared`; ответ содержит `mode`.
  - **Verify:** `pytest tests/test_nomenclature_import.py -q` (+ существующие)
  - **Dependencies:** GPS-015
  - **Files:** `app/services/nomenclature_import_service.py`, `app/schemas/nomenclature.py`,
    `app/api/v1/endpoints/nomenclature.py`, `tests/test_nomenclature_import.py`
  - **Scope:** M

- [x] **GPS-017: Веса/объёмы из отчёта → справочник весов**
  - **Description:** при импорте .xlsx — upsert веса/объёма в справочник весов
    (существующий модуль weight-catalog): match по GUID, fallback по нормализованному
    имени; конфликт не перетирает вручную введённое (как в существующей логике каталога).
  - **Acceptance:** после импорта образца веса видны в справочнике; без мэтча — пропуск
    (счётчик в ответе).
  - **Verify:** `pytest tests/test_weight_catalog.py tests/test_nomenclature_import.py -q`
  - **Dependencies:** GPS-015 (после GPS-016)
  - **Files:** `app/services/nomenclature_import_service.py`, модуль weight-catalog, тесты
  - **Scope:** S

- [x] **GPS-018: `GET /nomenclature/tasks` — живые очереди**
  - **Description:** общий `core/guid_queue.py` (логика из `scripts/build_guid_queue.py`
    переносится и реюзается скриптом): 🏭 — наши марки без GUID (по всем 6 видам,
    «у»-срез для свай); 💰 — `price_queue` open; ⚠️ — не-плиты ambiguous +
    плиты-дубли (GROUP BY guid_1c HAVING count>1). Endpoint admin+economist.
  - **Acceptance:** ответ группирован `{to_create_1c: [...], to_price: [...],
    duplicates: [...]}`; скрипт Excel использует тот же core и даёт тот же состав.
  - **Verify:** `pytest tests/test_guid_queue.py -q`; прогон скрипта на копии pb.db
  - **Dependencies:** GPS-012, GPS-013
  - **Files:** `core/guid_queue.py`, `scripts/build_guid_queue.py` (тонкая обёртка),
    `app/api/v1/endpoints/nomenclature.py`, `app/schemas/nomenclature.py`, тесты
  - **Scope:** M

- [x] **GPS-019: `POST /nomenclature/price-queue/resolve` — ввод цены 💰**
  - **Description:** вынос apply-логики из `core/price_import_queue.py` в переиспользуемую
    функцию (та же семантика по видам/классам); endpoint `{guid, price}` → валидация
    (>0, конечная, потолок) → запись в прайс-таблицы → upsert `nomenclature_guid`
    (guid, status auto) → `price_queue.resolved*`. RBAC: admin+economist (REQUIRE_PRICES).
  - **Acceptance:** цена появляется в прайсе, GUID связан, задача closed; повтор → 409/ok
    идемпотентно; manager → 403; цена ≤ 0 → 400.
  - **Verify:** `pytest tests/test_price_queue.py tests/test_pile_price_import.py -q`
  - **Dependencies:** GPS-013
  - **Files:** `core/price_import_queue.py` (extract), `core/price_queue_db.py`,
    `app/api/v1/endpoints/nomenclature.py`, `app/schemas/nomenclature.py`, тесты
  - **Scope:** M

- [x] **GPS-020: Дубли: `GET /nomenclature/duplicates` + `POST .../resolve`**
  - **Description:** GET — не-плиты из `duplicate_candidates` + плиты-дубли с кандидатами
    (имя, цена 1С справочно). POST `{scope: plate|nonplate, key, chosen_guid, note}`:
    валидация кандидата (409 если исчез — A1); не-плиты → `nomenclature_guid` upsert
    `match_status='manual'` + аудит (chosen_by/at/note) + `clear_candidates`; плиты →
    `plate_guid_choice.set_choice`. RBAC: admin+economist.
  - **Acceptance:** резолв не-плиты виден в `get_guid_for_invoice`; резолв плиты виден в
    резолвере GPS-014; аудит заполнен; manager → 403.
  - **Verify:** `pytest tests/test_duplicate_resolve.py tests/test_guid_gate.py -q`
  - **Dependencies:** GPS-011, GPS-012 (резолвер-проверка — GPS-014)
  - **Files:** `core/nomenclature_guid.py` (set_manual_choice), `core/plate_guid_choice.py`,
    `core/duplicate_candidates.py`, `app/api/v1/endpoints/nomenclature.py`,
    `app/schemas/nomenclature.py`, `tests/test_duplicate_resolve.py`
  - **Scope:** M

- [x] **GPS-021: Проверка драфта `GET /commercial/drafts/{id}/guid-check`**
  - **Description:** собирает строки драфта (как calculate) → `check_invoice_guids` →
    `{missing: [{product_kind, mark, reason, action_hint}]}`. Роли — как у мастера
    (manager/admin/economist).
  - **Acceptance:** драфт с позицией без GUID → список; без дырок → пусто; чужой драфт
    менеджера — 403 по существующим правилам.
  - **Verify:** `pytest tests/test_commercial_draft_append.py tests/test_guid_gate_api.py -q`
  - **Dependencies:** GPS-014
  - **Files:** `app/api/v1/endpoints/commercial.py`, `app/schemas/commercial.py`, тесты
  - **Scope:** S

- [x] **GPS-022: Уведомление при архиве `kp_guid_missing`**
  - **Description:** в `offers_service.create_offer` после сохранения — резолвер по
    составу; при missing → `create_notification` активным economist (A4), тип
    `kp_guid_missing`, payload `{kp_id, seq, marks[]}`, link `/prices`; дедуп по
    (kp_id, mark) (A5).
  - **Acceptance:** архив с дыркой → 1 уведомление на марку у каждого economist; без
    дырок → тихо; повторный архив не дублирует; нет economist → warning в лог.
  - **Verify:** `pytest tests/test_notifications.py tests/test_offers_service.py -q`
  - **Dependencies:** GPS-014
  - **Files:** `app/services/offers_service.py`, `app/repositories/promise_repository.py`,
    `app/schemas/notifications.py`, тесты
  - **Scope:** M

### Checkpoint: Phase 2
- [x] Backend-тесты зелёные (unit). Ручной прогон импорта .xls/.xlsx на копии pb.db — не выполнялся (dev-сервер жив, прод-БД не трогаем).

### Phase 3 — Frontend

- [x] **GPS-023: Alert на шаге 3 мастера**
  - **Description:** `CalculationResultStep`: запрос `guid-check` по draft id; неблокирующий
    `Alert variant="warning"` со списком марок и текстом «…счёт в 1С не уйдёт; экономист
    будет уведомлён при сохранении в архив». Без дырок — ничего.
  - **Acceptance:** RTL: Alert есть/нет; переходы мастера не меняются.
  - **Verify:** `cd frontend && npm run test -- src/features/commercial-offer`
  - **Dependencies:** GPS-021
  - **Files:** `CalculationResultStep.tsx(+test)`, `api/commercialOfferApi.ts`, types
  - **Scope:** S

- [x] **GPS-024: Вкладка «Прайсы и 1С»: каркас + переезд импорта 1С**
  - **Description:** `PricesView` → секции: «Загрузка прайса завода» (как сейчас),
    «Выгрузка из 1С» (перенос `Import1cDialog` сюда; подсказка про частичные выгрузки
    и `mode` из ответа), плейсхолдеры «Задачи по изделиям»/«Дубли GUID». Кнопку
    «Выгрузка 1С» из `AppHeader` убрать. Роли без изменений (admin+economist).
  - **Acceptance:** RTL + ручное: диалог живёт на вкладке, шапка чистая, manager не
    видит вкладку.
  - **Verify:** `cd frontend && npm run test -- src/features/price-desk src/features/nomenclature src/app`
  - **Dependencies:** GPS-016
  - **Files:** `PricesView.tsx(+test)`, `Import1cDialog.tsx(+test)`, `AppHeader.tsx(+test)`
  - **Scope:** M

- [x] **GPS-025: Секция «Задачи по изделиям»**
  - **Description:** 🏭 — read-only список (вид, марка, подсказка «завести в 1С и
    загрузить отчёт»); 💰 — список с inline-полем цены → POST resolve → инвалидация
    queries. TanStack Query hooks.
  - **Acceptance:** RTL: рендер очередей, ввод цены закрывает 💰-задачу; пустые очереди —
    «задач нет».
  - **Verify:** `cd frontend && npm run test -- src/features/price-desk`
  - **Dependencies:** GPS-018, GPS-019, GPS-024
  - **Files:** `PricesView.tsx` (+ `TasksSection.tsx`), `api/nomenclatureApi.ts`,
    `hooks/useNomenclatureQueries.ts`, тесты
  - **Scope:** M

- [x] **GPS-026: Секция «Дубли GUID»**
  - **Description:** список дублей (не-плиты + плиты): кандидаты радиокнопками (имя,
    цена 1С справочно), поле заметки с подсказкой «по решению бухгалтерии», кнопка
    «Запомнить выбор» → POST resolve; 409 → показать «кандидат исчез, задача открыта».
  - **Acceptance:** RTL: выбор → задача исчезает из списка; ошибка 409 отображается.
  - **Verify:** `cd frontend && npm run test -- src/features/price-desk`
  - **Dependencies:** GPS-020, GPS-024
  - **Files:** `PricesView.tsx` (+ `DuplicatesSection.tsx`), `api/nomenclatureApi.ts`, тесты
  - **Scope:** M

- [x] **GPS-027: Уведомление `kp_guid_missing` во фронте**
  - **Description:** тип в `notifications.ts`, текст «КП №…: изделия без GUID — …»,
    клик → `/prices`.
  - **Acceptance:** RTL: рендер текста и переход по ссылке.
  - **Verify:** `cd frontend && npm run test -- src/features/notifications`
  - **Dependencies:** GPS-022
  - **Files:** `features/notifications/*`, `shared/ui` (при необходимости), тесты
  - **Scope:** S

### Checkpoint: Phase 3
- [x] `npm run typecheck` и targeted vitest зелёные. Ручной прогон вкладки и мастера в браузере не выполнялся (нет browser MCP в этой сессии).

### Phase 4 — Ворота

- [x] **GPS-028: Финальные ворота счёта**
  - **Description:** `assert_invoice_guids(order_lines, conn)` оформлен как публичная
    точка для будущей кнопки «В 1С»: docstring «куда вставить», интеграционный тест
    «счёт без GUID невозможен» на уровне сервиса (эмуляция вызова).
  - **Acceptance:** тест падает до вызова assert и зелёный после; отчёт содержит
    человекочитаемый список «что завести».
  - **Verify:** `pytest tests/test_guid_gate.py -q`
  - **Dependencies:** GPS-014 (логически после Phase 2)
  - **Files:** `core/guid_gate.py`, `tests/test_guid_gate.py`
  - **Scope:** S

### Checkpoint: Complete
- [ ] Все Success Criteria S1–S12 из спеки подтверждены.
- [ ] Коммиты только по явной просьбе пользователя.

## Risks and Mitigations

| Риск | Impact | Митигация |
|------|--------|-----------|
| Реальный .xlsx отчёта отличается от предположений | Med | GPS-015 TDD на синтетике + проверка образцом пользователя до GPS-016 |
| Частичная выгрузка без явного блока отборов | Med | Эвристика + при сомнении считать partial; `mode` в ответе |
| Расхождение «у»-среза резолвера и линта | Med | Общая функция среза (D11); кросс-тесты |
| Семантика цены 💰 по классам свай | Med | Резолв через ту же apply-функцию Excel-импорта (GPS-019) |
| Нет активных economist | Low | Лог-warning; задачи живут на вкладке (A4/A5) |
| Override дубля плиты устарел | Low | 409 при резолве; задача переоткрывается (A1) |

## Open Questions

- Внешние (не блокируют старт): полная выгрузка универсального отчёта без отборов;
  ритм reexport; нужен ли «Код» 1С. См. спеку, раздел Open Questions.
