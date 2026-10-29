# Реестр находок аудита

Стабильные ID для `/audit`. Один ID = одна проблема навсегда. Строки **не удалять**.
Новый ID — только если нет совпадения; следующий свободный номер в серии A / S / Q.

Статусы: `open` | `resolved` | `wontfix` | `by-design` | `unreproduced`

- `unreproduced` ≠ `resolved` (не нашли в этом прогоне ≠ закрыто)
- Не упомянули в отчёте ≠ закрыто
- P0 = open + Critical; P1 = open + High

Источник прогона 0: [2026-08-28-full-project-audit.md](./2026-08-28-full-project-audit.md).
Прогон 1 (сверка реестра): [2026-08-28-full-registry-sweep-audit.md](./2026-08-28-full-registry-sweep-audit.md).
KP (конструктор КП): [2026-09-04-kp-audit.md](./2026-09-04-kp-audit.md).
Архив + производство: [2026-09-04-archive-production-audit.md](./2026-09-04-archive-production-audit.md).
DS (график поставки): [2026-09-21-delivery-schedule-audit.md](./2026-09-21-delivery-schedule-audit.md).
Прогон 2 (--full): [2026-09-21-full-audit.md](./2026-09-21-full-audit.md).
Договор поставки (архив): [2026-09-27-supply-contract-audit.md](./2026-09-27-supply-contract-audit.md).
KP + архив (конструктор): [2026-10-01-kp-audit.md](./2026-10-01-kp-audit.md).
Производство: [2026-10-19-production-audit.md](./2026-10-19-production-audit.md).
Производство (план → СГП): [2026-10-26-production-audit.md](./2026-10-26-production-audit.md).
Прогон 3 (--full): [2026-10-26-full-audit.md](./2026-10-26-full-audit.md).
Форензик целостности плана (A24–A26): [идея](../../ideas/целостность-плана-производства.md) — логи + `plita.db`, 2026-09-21.
GSM High с [2026-08-26-gsm-audit.md](./2026-08-26-gsm-audit.md) перенумерованы (там были свои A1…, конфликт с этим реестром) — колонка Legacy.

Последнее обновление: 2026-10-26 (аудит --full / весь проект).

| ID | Scope | Sev | Status | P | Summary | Evidence / notes | Last seen | Legacy |
|----|-------|-----|--------|---|---------|------------------|-----------|--------|
| A1 | layout | Critical | open | P0 | Неявное мутабельное глобальное состояние заказа плит (`PlateMutableRuntime`, `config_and_data`, PEP 562) | **Partial 2026-10-26:** по умолчанию `PLATE_RUNTIME_STRICT=raise`. Пустой контекст — `RuntimeError`, TLS не создаётся (`core/plate_runtime_state.py:167-188`). Прокси `PLATES_*` и `apply_to_globals` удалены. `_build_sequence_from_plan` требует `layout_cfg` и getter не вызывает. `warn` по-прежнему отдаёт runtime потока. Полный `pytest tests/` не зелёный (3979 passed, 22 failed, 1 error); среди падений нет «заказ плит не привязан». Статус не `resolved` | 2026-10-26 | |
| A2 | platform | Critical | open | P0 | In-process state (rate limit, counters) блокирует несколько воркеров | `app/security/login_rate_limit.py:49-50`, `:84-85`, `:179-191`; `app/main.py:47-52`. Fail-fast только production+single_instance+явный workers>1. Тот же паттерн `commercial_upload_validation.py:24` | 2026-10-26 | |
| A3 | kp | High | resolved | — | God-module `CommercialWorkflowService` | 2026-10-26: фасад **873** строки; `calculate_draft` делегирует `commercial_workflow_service.py:691-692` | 2026-10-26 | |
| A4 | kp | High | open | P1 | Толстый API-слой: `commercial.py` разобран; remainder `production.py` (699 строк, 26 маршрутов). Журнал исключений и документы дня в роутере | `production.py` **705** строк, **26** маршрутов. `:18-24`, `:175-186` (журнал исключений глотает сбой), `:441-453`, `:474-485`, `:505-513` | 2026-10-26 | |
| A5 | kp | High | open | P1 | Сервисы обходят repository, raw SQL | `sgp_service.py:5`, `:717-729`; `day_view_service.py:231`, `:234-240` (**615** строк); `delivery_schedule_service.py:6`, `:91-93`; `shipment_service.py` **1001** строк, `import sqlite3`. delivery/readiness не сужать. ID не дробить | 2026-10-26 | |
| A6 | layout | High | open | P1 | Planning импортирует visualization / matplotlib at load | `core/production/planning.py:41-46`; `core/visualization/__init__.py:16-18` matplotlib.use; `core/visualization/layout.py:20-24`, `:253`. Порт `core/ports/visualization.py` есть, обхода пакета нет | 2026-10-26 | |
| A7 | platform | High | open | P1 | Неполный DI в FastAPI (фабрики без constructor injection) | `app/dependencies/services.py:121-122`; `supply_contract_service.py:122-136`. Рядом без Depends на репозиторий: `:102-103`, `:117-118` | 2026-10-26 | |
| A8 | layout | Medium | open | — | Параллельные подсистемы планирования | Два календаря: `app/planning/plan_calendar.py:136` и `app/services/plan_distribution_service.py:135`, `:197`. Живой HTTP: `production_service.py:511`. В тело P0/P1 не входит; действия на эту неделю нет | 2026-10-26 | |
| A9 | platform | Medium | open | — | Пустые app-сервисы-реэкспорты | `kp_persistence_service.py`, `rest_matching_service.py` | 2026-08-28 | |
| A10 | bot | Medium | open | — | Legacy-пути бота в persistence планов | `plan_storage.py`; не живой продукт | 2026-08-28 | |
| A11 | kp | Medium | open | — | ArchiveService god-orchestrator | `archive_service.py` **2038** строк, методов 65 (`    def `). Действия на эту неделю нет | 2026-10-26 | |
| A12 | kp | Medium | open | — | Frontend god-hook мастера КП | `useCommercialOfferWizard.ts` **379** строк, **19** useMutation (было 358/18). Mutations `:68–337`; split draft `:68–279` vs plate `:97–209` | 2026-10-01 | |
| A13 | — | — | — | — | **дырка** в нумерации 2026-08-28 — не занимать | | | |
| A14 | layout | Low | open | — | Монолит `core/visualization`, matplotlib `use('Agg')` at import | `core/visualization/__init__.py` | 2026-08-28 | |
| A15 | kp | Low | by-design | — | `owner_user_id` не в policy доступа КП | `offer_access.py:26–32`; [offer-access-policy.md](../architecture/offer-access-policy.md); связано с S1 | 2026-09-21 | |
| A16 | gsm | High | open | P1 | God-сервис жизненного цикла ПЛ (CRUD + confirm + rechain) | `gsm_generation_service.py` **1177** строк (было 1030), **24** метода. Класс `:56`; `:70` generate, `:328` list, `:352` patch, `:455` create, `:538` confirm, `:727` `_rechain_downstream` | 2026-10-26 | gsm-audit A1 |
| A17 | gsm | High | open | P1 | God-модуль `core/gsm/generator.py` | `core/gsm/generator.py` **1533** строки (было 1324), **59** функций, `generate` с `:134` | 2026-10-26 | gsm-audit A2 |
| A18 | gsm | High | open | P1 | Блокирующий LibreOffice/`soffice` в HTTP (и DoS) | `gsm_export_service.py:54-70`; `gsm.py:662-672`. Тот же дефект, новый ID не заводить: `supply_contract_service.py:778-806` timeout 60с, вызов `:840-841`, синхронные GET `archive.py:742-748` | 2026-10-26 | gsm-audit A3, S1 |
| A19 | gsm | High | open | P1 | Сезонная логика дублируется на фронте, расхождение контракта | `core/gsm/season.py:21-32`; генерация `gsm_generation_service.py:850-856`; API хвост журнала `gsm_registry_service.py:293-303`; схема `app/schemas/gsm.py:279-280`; фронт `downstreamPreview.ts:33-44` | 2026-10-26 | gsm-audit A4 |
| A20 | gsm | High | open | P1 | Импорт транзакций без unit-of-work | `gsm_transaction_service.py:65-114`; `gsm_repository.py:388` commit batch, `:404-424` commit строки | 2026-10-26 | gsm-audit A5 |
| A21 | gsm | Medium | open | — | `gsm_kit_gate` импортирует private `_chain_broken` из overview | `gsm_kit_gate.py:10` import, `:120` use | 2026-09-21 | |
| A22 | kp | High | open | P1 | God-сервис `ProductionCompletionService` + raw sqlite3 | `production_completion_service.py` **702** строки; sqlite3 `:4`, `:123-125`; `complete_day` `:281` | 2026-10-26 | |
| A23 | kp | High | open | P1 | God-module `DeliveryScheduleService` ~878 LOC, ~24 methods | `delivery_schedule_service.py` **878** строк, **24** метода, соединение `:75-93`, репозитория нет | 2026-10-26 | |
| A24 | layout | High | open | P1 | `concrete_grade` теряется в пайплайне планирования | Три сборки копируют `concrete_grade` с реза `viz_modules/layout_sequence/from_plan.py:693`, `:761`, `:965`. Тест `test_golden_root_items_carry_cut_grade_not_resolver_default`. Хеш марку выкидывает `tests/test_layout_hash.py:1-4`, `:40-55`. Skip/rescue `plate_attribution` и legacy `day_view` не перечитывались. **Не закрывать** | 2026-10-26 | |
| A25 | layout | Critical | open | P0 | Сироты/призраки: двойная атрибуция одной геометрии к разным КП (assignment-side vs track-side, нешаренные consumed) | Источник: `core/plate_attribution.py:111–114` (два взгляда, разные правила выбора); фикс-канал: unit_id (`finalize.py:224–226` ↔ `from_plan.py:693–694`). Кейс `plan_20260904_124248`: «ПБ 20,6-7,2-8п» заказано КП#4=3+КП#5=1, атрибуция 2/2 → сирота `kp_plates.id=197` + призрак КП#5 (3-й трек 2026-09-14). Симптом: `plan_commit.py:614–641` (pro-rated), `:650` (pool guard), `:703–708` (silent continue), `:711–743` (orphan-rollback мимо NULL-day) | 2026-09-21 | |
| A26 | layout | High | open | P1 | Слепота детекции surplus: coverage игнорирует пере-покрытие (ILP невиновен — доказано) | `coverage_verify.py:60` теперь `ok = not missing and not surplus` (сдвиг относительно улики «только missing»). ILP `ilp_model.py:261` по-прежнему `==`. `extract_cuts` и rescue в `plan_commit` не открывались. **Не закрывать** | 2026-10-26 | |
| A27 | kp | Medium | open | — | `update_contract_date` пишет в БД в обход `_assert_writer` (supply-contract) | `supply_contract_service.py:376–395` writes DB; `_assert_writer` `:397–398` not called. No HTTP route. Test happy-path ADMIN `tests/test_supply_contract.py:390–394` | 2026-10-01 | |
| A28 | kp | Medium | open | — | ArchiveService обходит SupplyContractService, зовёт SupplyContractRepository | `archive_service.py:527–529`, `:589–591`, `:746–748`. Action: read-port или `SupplyContractService.get_for_kp` | 2026-10-01 | |
| A29 | production | High | resolved | — | Сборка плана и резерв СГП коммитятся разными соединениями; ошибка компенсации глотается | 2026-10-26: `reserve_on_conn` на том же conn, один commit `core/production/planning.py:341-385`, вызов `production_service.py:458-476`. Тест `test_sgp_reserve_fail_after_build_compensates_no_plan_plates_not_in_plan`. `set_active` после commit `planning.py:392-398` — в A29/A30 не возвращать | 2026-10-26 | |
| A30 | production | High | resolved | — | Переход «плиты в производстве → в плане» не атомарен: `persist` коммитит пометку плит, затем отдельным соединением пишет документ плана; ошибка `return_plan_plates_to_production` глотается | 2026-10-26: тот же commit; удаление `commit_plan_deletion` `plan_storage.py:166-194`. Тест `test_plan_save_fail_after_plate_mark_rolls_back_no_plan_plates_stay_in_production` | 2026-10-26 | |
| S1 | kp | High | by-design | — | Менеджер видит чужие КП — общий архив (не IDOR для fix) | `offer_access.py:1-8`; ADR offer-access-policy.md; не IDOR. Связано с A15 | 2026-10-26 | |
| S2 | platform | High | resolved | — | CVE Starlette/FastAPI | 2026-10-26: `fastapi==0.141.1`, `starlette==1.6.0`; `pip-audit -r requirements.txt` — No known vulnerabilities found, exit 0 | 2026-10-26 | |
| S3 | platform | High | open | P1 | Rate limiting in-process (см. A2) | `commercial_upload_validation.py:23-24`, `:53-57`, `:141-142`; default `commercial_ocr_uploads_per_hour` = 0 → проверка сразу выходит `core/config/settings.py:150-154`; плюс строки A2 | 2026-10-26 | |
| S4 | platform | High | open | P1 | npm audit high | `npm audit --json` в `frontend/` exit 1; metadata high=5, critical=0, moderate=4, total=9. Прямые: `zod` 4.3.6 (`frontend/package.json:23`, диапазон `<4.6.6`), `vite` 8.0.16 (`package.json:34`, диапазон `8.0.0–8.1.5`). Транзитивные high: undici через jsdom, brace-expansion, source-map-js. До Critical не повышать | 2026-10-26 | |
| S5 | kp | High | open | P1 | OCR / vision во внешние LLM; supply-contract parse обходит `ensure_external_ocr_enabled` и rate limit | `archive.py:367-375`, `:710-718`, `:847-859`; `supply_contract_service.py:421-439`, `:1031-1074`; grep `ensure_external_ocr_enabled` в сервисе договора — 0. Картинка и текст | 2026-10-26 | |
| S6 | platform | Medium | open | — | SQLite без шифрования at rest | локальный завод | 2026-09-21 | |
| S7 | auth | Medium | open | — | CSP Report-Only + unsafe-inline | `security_headers.py:11–36` | 2026-09-21 | |
| S8 | platform | Medium | open | — | Утечка деталей ошибок в HTTP | `detail=str(exc)` `production.py:148`, `:181`, `:210`, `:216`, `:703` и много мест в `archive.py` (напр. `:206`, `:674`, `:805`). Часть маршрутов режется `raise_unexpected_server_error` `app/core/http_errors.py:67-71`. Severity не повышать | 2026-10-26 | |
| S9 | auth | Medium | open | — | CSRF парсит multipart до проверки токена | `app/middleware/csrf.py:36-47` по-прежнему читает multipart, если нет заголовка | 2026-10-26 | |
| S10 | auth | Medium | open | — | Сессия 12ч без refresh-ротации | TTL 12ч `core/config/settings.py:63-67`, без ротации | 2026-10-26 | |
| S11 | auth | Low | by-design | — | CSRF-cookie не HttpOnly (double-submit) | `app/security/csrf.py:21-28`. CSRF-cookie не HttpOnly — ожидаемо | 2026-10-26 | |
| S12 | auth | Low | open | — | Password policy messages на английском | auth schemas | 2026-09-21 | |
| S13 | auth | Low | open | — | `/health` метаданные вне production | `/health` в production без метаданных `health.py:16-19` | 2026-10-26 | |
| S14 | bot | Low | open | — | Legacy bot auth bypass при `BOT_AUTH_ENABLED=false` | `bot_archived/`; не живой контур | 2026-08-28 | |
| S15 | kp | Low | open | — | Draft в sessionStorage (XSS-вектор черновика) | `draftStorage.ts:8`, `:22` sessionStorage | 2026-10-26 | |
| S16 | gsm | High | open | P1 | Импорт без лимита файлов (DoS); per-file 50MB есть | `gsm.py:196-199`; на файл 50МБ `settings.py:145-148`; `core/gsm/transactions.py:199-202` без лимита строк; коммит строки `gsm_repository.py:404-424` | 2026-10-26 | gsm-audit S2 |
| S17 | kp | Medium | open | — | XLSX import: no magic bytes / no row cap | DS half not rechecked. **supply-contract**: `supply_contract_import.py:35–47` `load_workbook` + unbounded `iter_rows`; file size only `read_upload_file_capped` `archive.py:194` | 2026-10-01 | |
| S18 | kp | Medium | open | — | Formula injection в генерируемом XLSX | `core/delivery_schedule_xlsx.py` export — значения с `=`, `+`, `-`, `@` | 2026-09-21 | |
| S19 | kp | Medium | open | — | Unbounded PUT batches/items/name | `app/schemas/delivery_schedule.py` — нет `max_length` / `max_items` на PUT payload | 2026-09-21 | |
| S20 | production | High | resolved | — | Клиент сборки плана шлёт `expected_version`, схема `BuildPlanRequest` его отбрасывает; `persist` берёт version из БД | 2026-10-26: `BuildPlanRequest.expected_version` `app/schemas/production.py:230`, `production.py:140`. Тест `test_append_stale_expected_version_returns_409_plan_and_plates_unchanged`. `fill_targets` по-прежнему без проверки версии — это не переоткрывает S20 | 2026-10-26 | |
| S21 | production | High | resolved | — | `DELETE /plans/{plan_id}` не возвращает плиты в производство; `kp_plates` остаются «в плане». `plan_storage.delete_plan` вызывает `return_plan_plates_to_production`; HTTP-путь — нет. Два commit | 2026-10-26: DELETE → `commit_plan_deletion` `production.py:261-267`, `production_service.py:156-163`, `plan_storage.py:166-194` | 2026-10-26 | |
| Q1 | kp | High | resolved | — | Шесть копий product-type pipeline | 2026-10-26: `product_draft_config.py` — таблица типов | 2026-10-26 | |
| Q2 | kp | High | resolved | — | Copy-paste HTTP-обработчиков КП | 2026-10-26: тонкие `_run_product_*` `commercial.py:117-156` | 2026-10-26 | |
| Q3 | layout | High | resolved | — | ~720 строк мёртвого кода `build_layout_sequence` | 2026-10-26: builder ~265 строк, хеш `tests/test_layout_hash.py:30-37` | 2026-10-26 | |
| Q4 | kp | Medium | open | — | Пять копий `build_*_preview_metadata` | Копий `build_*_preview_metadata` уже **7** (было 5): `commercial_draft_service.py:197`, `:281`, `:354`, `:427`, `:500`, `:570`, `:643`. Свести в один сборщик до следующего типа изделия. Medium, не P1 | 2026-10-26 | |
| Q5 | kp | Medium | resolved | — | Дублирование resolve_wide/unpriced plates | 2026-10-26: `CommercialPlateResolve.resolve_wide_plates` / `resolve_unpriced_plates`, workflow только делегирует | 2026-10-26 | |
| Q6 | layout | Medium | open | — | Две реализации `get_global_calendar_info` | Оба `get_global_calendar_info` на месте: `plan_calendar.py:70`, `plan_distribution_service.py:135`. **Не закрывать** | 2026-10-26 | |
| Q7 | kp | Medium | open | — | Product-type duplication на фронте | partial `productTypeConfig.ts`; remainder `CalculationResultStep.tsx:52–74` | 2026-10-01 | |
| Q8 | kp | Medium | open | — | God-hook `useCreatePlanWizardState` | `useCreatePlanWizardState` **997** строк. В тело P1 не входит | 2026-10-26 | |
| Q9 | kp | Medium | open | — | God-component `OfferDetailsDrawer` | **1658** строк (было 1369) | 2026-10-01 | |
| Q10 | kp | Medium | open | — | Слабая типизация production API | JSON-маршруты без `response_model` на месте. Свежих номеров строк нет | 2026-10-26 | |
| Q11 | kp | Medium | open | — | preview: `Any` / `dict[str, Any]` | `commercial_draft_service.py:192,276,349,422,495`; `schemas/commercial.py:316–323` | 2026-10-01 | |
| Q12 | kp | Medium | open | — | ArchiveService скрывает частичные сбои | Список глотает `archive_service.py:1275-1288`; `move_to_production` рейзит `:993-997` | 2026-10-26 | |
| Q13 | layout | Medium | open | — | Нет прямых тестов `get_global_calendar_info` | mocks в `test_delivery_schedule_service.py`; после консолидации Q6 | 2026-09-21 | |
| Q14 | platform | Low | open | — | Однострочные delegate-обёртки | services | 2026-08-28 | |
| Q15 | layout | Low | open | — | Имя `_merge_plate_texts` вводит в заблуждение | layout utils | 2026-08-28 | |
| Q16 | kp | Low | open | — | `/parse` без `response_model` | `commercial.py:185–190` returns `dict`, no response_model | 2026-10-01 | |
| Q17 | gsm | Low | open | — | `GsmGenerationError` messages на английском | Тексты `GsmGenerationError` на английском `gsm_generation_service.py:122`, `:856` | 2026-10-26 | |
| Q18 | kp | Low | open | — | Подавление `react-hooks/exhaustive-deps` в 10 файлах | in-scope ранее: `MoveToProductionDialog.tsx:87`, `CounterpartyAutocomplete.tsx:78`. 2026-10-26: подавления в мастере на месте `useCreatePlanWizardState.ts:256`, `:368`. Новых Low нет. Глобально не закрыто | 2026-10-26 | |
| Q19 | gsm | High | open | P1 | Расхождение округления литров: `Math.round` vs banker's round | `downstreamPreview.ts:23-26` Math.round; `core/gsm/balance.py:15-17`; `tests/test_gsm_balance.py:73-77`; фронт-тест `downstreamPreview.test.ts:54-55` не на границе `*.xx5` | 2026-10-26 | gsm-audit Q1 |
| Q20 | gsm | High | open | P1 | Несогласованная обработка битого `season_switches` | `gsm_registry_service.py:343-361` глотает; генерация `gsm_generation_service.py:850-858`; экспорт `gsm_export_service.py:208-215`; отчёт `gsm_report_service.py:294-301`; `tests/test_gsm_season.py:358-371`. `switch_season` читает глотающий путь и перезаписывает | 2026-10-26 | gsm-audit Q2 |
| Q21 | gsm | Low | open | — | Две реализации `formatLiters` на фронте GSM | `importReport.ts:28–29` vs `waybillWarnings.ts:80–84` | 2026-08-28 | |
| Q22 | kp | Medium | open | — | Дублирование `is_*_draft` в трёх модулях | canon `commercial_calculation_service.py:42–58`; copy `commercial_export_service.py:212–234`; delegate `commercial_wizard_step_service.py:52–65` | 2026-10-01 | |
| Q23 | kp | Medium | open | — | Расхождение оценки дорожек: archive `round` vs production FE `ceil` | FE производства `Math.ceil(length/101)` `productionEstimate.ts:24`. Archive `round` не перепроверялся — ID не закрывать | 2026-10-26 | |
| Q24 | kp | Medium | open | — | Copy-paste day-document handlers: мёртвый raise после raise_not_found | Мёртвый raise после `raise_not_found` в разбивке и формовке `production.py:488-496`, `:515-523`. У схемы хвоста уже нет `:455-465`. Удалить два недостижимых блока. Medium, не P1 | 2026-10-26 | |
| Q25 | kp | Medium | open | — | `has_delivery_schedule` FE badge; list API не отдаёт поле | `ArchiveOfferList.tsx:173`; `types/archive.ts:26`; `schemas/archive.py:25–57` field absent; grep `app/` 0 hits `has_delivery_schedule` | 2026-10-01 | |
| Q27 | kp | Medium | open | — | `produce_by ≤ deliver_from` warning не реализован | schema: только `deliver_from ≤ deliver_to` (`schemas/delivery_schedule.py:55–64`); спека §215–216 | 2026-09-21 | |
| Q28 | kp | Medium | open | — | Три дублирующих SELECT из `kp_plates` | `delivery_schedule_service.py:496–500,682–686,703–707` | 2026-09-21 | |
| Q29 | kp | Medium | open | — | Нет service-тестов import_draft happy-path + GET red | gaps в `test_delivery_schedule_*` | 2026-09-21 | |
| Q30 | kp | Medium | open | — | Unmatched reasons на английском в UI | `delivery_schedule_xlsx.py:65–68`; `DeliveryScheduleEditor.tsx:160–161` | 2026-09-21 | |
| Q31 | kp | Medium | open | — | Форма договора: «доверенность» без `poa_number` / `poa_date` (supply-contract) | `SupplyContractDrawer.tsx:627–637` select only; BE `supply_contract_service.py:443–445` requires poa number/date | 2026-10-01 | |
| Q32 | kp | Medium | open | — | Мёртвый verify-путь OCR карточки договора | `apply_contract_verify` `core/supply_contract_card_ocr.py:65` 0 callers; `get_contract_card_verify_prompt` `core/ocr/prompts.py:153` 0 callers; `already_exists_message` `supply_contract_service.py:544` 0 callers | 2026-10-01 | |
| Q33 | kp | Medium | open | — | Дубли `extract_contract_fields` / `reread_contract_fields` | `supply_contract_service.py:625–651` both call `complete_vision_json` | 2026-10-01 | |
| Q34 | kp | Medium | open | — | God-component `SupplyContractDrawer` (~741 LOC) | **741** строк (было 601); modal + OCR preview + форма + реестр | 2026-10-01 | |
| Q35 | kp | Medium | open | — | FE `SupplyContract` слабее `SupplyContractOut` | `types/supplyContract.ts:15–29` `status: string` vs `schemas/supply_contract.py:99–131` | 2026-10-01 | |
| Q36 | kp | Medium | open | — | Формат даты договора: ISO в drawer vs DD.MM.YYYY в реестре | `SupplyContractDrawer.tsx:345` ISO vs `SupplyContractRegistry.tsx:21–27,88` DD.MM.YYYY | 2026-10-01 | |
| Q37 | kp | Medium | open | — | Нет теста отказа non-writer для `update_contract_date` | `tests/test_supply_contract.py:390–394` ADMIN only; pair A27 | 2026-10-01 | |
| Q38 | kp | Medium | open | — | Enum статусов/скана договора в трёх местах | `core/supply_contract.py:11–18`; `schemas/supply_contract.py:13–20`; `types/supplyContract.ts:1–11` | 2026-10-01 | |
| Q39 | kp | Medium | open | — | Enum kind спецификации (payment/term/delivery) triplicated | `core/specification_text.py:10–12`; `schemas/archive.py:286–288`; `types/archive.ts:211–220`; UI `SpecificationPanel.tsx:16–33` | 2026-10-01 | |
| Q40 | production | High | open | P1 | Актор аудита производства всегда пустой: код ищет email/login/user_id, в сессии только username и id | **Не закрывать.** Сборка плана уже пишет username иначе str(id): `production.py:111-113`, `production_service.py:469-476`, `sgp_service.py:717-726`. Пустые вызовы: `production.py:297`, `:399`, `:580`, `:600`; `production_capacity_service.py:25-35`. Сессия: `auth_repository.py:84-85`, `app/dependencies/auth.py:21` — id и username, не email/login/user_id | 2026-10-26 | |

## Счётчики (прогон 2 — --full, 2026-09-21)

| Scope | Open P0 | Open P1 |
|-------|---------|---------|
| layout | 1 (A1) | 1 (A6) |
| platform | 1 (A2) | 2 (A7, S3) |
| kp | 0 | 4 (A4, A5, A22, A23) |
| gsm | 0 | 8 (A16–A20, S16, Q19, Q20) |
| auth | 0 | 0 |
| **все open Critical/High** | **2 P0** | **15 P1** (A4, A5, A6, A7, S3, A16–A20, S16, Q19, Q20, A22, A23) |

*После [supply-contract 2026-09-27](./2026-09-27-supply-contract-audit.md): S5 Medium→High P1 → глобально **2 P0 / 16 P1** (+S5); scope **kp** **0 P0 / 5 P1** (A4, A5, A22, A23, S5).*

### Счётчики (supply-contract, 2026-09-27)

| Метрика | Значение |
|---------|----------|
| Закрыто в скоупе | 0 |
| Open P0 / P1 в скоупе | **0 / 2** (A7 platform, S5 kp) |
| Новые ID | A27, Q31–Q38 (Medium) |

### Счётчики (KP + архив, 2026-10-01)

| Метрика | Значение |
|---------|----------|
| Закрыто в скоупе | 0 |
| Подтверждено open | 28 |
| Новые ID | A28, Q39 (Medium) |
| Open P0 / P1 в скоупе | **0 / 3** (A7 platform, S3 platform, S5 kp) |

*Глобальный P0/P1 из прогона 2 и [supply-contract 2026-09-27](./2026-09-27-supply-contract-audit.md) **не пересчитывался**; новых Critical/High нет. A4, A5, A22, A23 остаются open P1 kp без изменения статуса (не проверялись). S5 остаётся High P1 (kp).*

### Счётчики (производство, 2026-10-19)

| Метрика | Значение |
|---------|----------|
| Закрыто в скоупе | 0 |
| Подтверждено open | 12 |
| Новые ID | A29, S20, S21, Q40 |
| Open P0 / P1 в скоупе | **0 / 8** (A4, A5, A6, A22, A29, S20, S21, Q40) |

*Глобальный P0/P1 из прогона 2 и последующих модульных аудитов не пересчитывался. Новые High этого прогона: A29, S20, S21, Q40.*

### Счётчики (производство, план → СГП, 2026-10-26)

| Метрика | Значение |
|---------|----------|
| Закрыто в скоупе | 0 |
| Подтверждено open | 16 |
| Новые ID | A30 (High, P1) |
| Open P0 / P1 в скоупе | **0 / 9** (A4, A5, A6, A22, A29, A30, S20, S21, Q40) |

*Глобальный P0/P1 не пересчитывался; в скоупе новый High A30.*

A18 = gsm LibreOffice+DoS (старые gsm A3+S1). Не плодить второй ID.
A21 (Medium), Q21 (Low) — не входят в P1.
Q26 не занят (пропущен).

Следующий свободный ID: **A31**, **S22**, **Q41**.

### Счётчики (--full, 2026-10-26)

Пересчёт по колонке Status таблицы после правки S4: open + Critical = P0, open + High = P1. Исторические блоки выше не переписывались. Блок «производство, план → СГП» всё ещё перечисляет A29, A30, S20, S21 как open P1; в таблице они resolved и в этот счётчик не входят.

| Scope | Open P0 | Open P1 |
|-------|---------|---------|
| layout | 2 (A1, A25) | 3 (A6, A24, A26) |
| platform | 1 (A2) | 3 (A7, S3, S4) |
| kp | 0 | 5 (A4, A5, A22, A23, S5) |
| gsm | 0 | 8 (A16, A17, A18, A19, A20, S16, Q19, Q20) |
| production | 0 | 1 (Q40) |
| auth | 0 | 0 |
| **все open Critical/High** | **3 P0** | **20 P1** |

| Метрика | Значение |
|---------|----------|
| Закрыто в скоупе | 0 |
| Переоткрыто | 1 (S4, High, P1) |
| Подтверждено open | 35 |
| Без изменения (last_seen не сдвигался) | 6 (A8, A9, A12, A21, A27, A28) |
| Не открывали | 34, включая A25 (open P0) |
| Новые ID | 0 |
| Open P0 / P1 в скоупе | **3 / 20** |

P0: A1, A2, A25. P1: A4, A5, A6, A7, A16, A17, A18, A19, A20, A22, A23, A24, A26, S3, S4, S5, S16, Q19, Q20, Q40.

Прогон 0 ([2026-09-21-full-audit.md](./2026-09-21-full-audit.md)) считал 2 P0 / 15 P1 и не включал A24–A26, S5 и Q40. Сейчас 3 / 20: в реестре учтены A25 (P0), A24 и A26 (P1), S5, Q40, и S4 снова open.

Подтверждено open (35): A1, A2, A4, A5, A6, A7, A11, A16, A17, A18, A19, A20, A22, A23, A24, A26, S3, S5, S8, S9, S10, S13, S15, S16, Q4, Q6, Q8, Q10, Q12, Q17, Q19, Q20, Q23, Q24, Q40.

Отчёт: [2026-10-26-full-audit.md](./2026-10-26-full-audit.md).

Следующий свободный ID: **A31**, **S22**, **Q41**. A13 и Q26 не занимать.

