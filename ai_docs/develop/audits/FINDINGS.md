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
Форензик целостности плана (A24–A26): [идея](../../ideas/целостность-плана-производства.md) — логи + `plita.db`, 2026-09-21.
GSM High с [2026-08-26-gsm-audit.md](./2026-08-26-gsm-audit.md) перенумерованы (там были свои A1…, конфликт с этим реестром) — колонка Legacy.

Последнее обновление: 2026-10-01 (аудит KP + архив — [2026-10-01-kp-audit.md](./2026-10-01-kp-audit.md); ранее 2026-09-27 supply-contract; 2026-09-21 прогон 2 — --full).

| ID | Scope | Sev | Status | P | Summary | Evidence / notes | Last seen | Legacy |
|----|-------|-----|--------|---|---------|------------------|-----------|--------|
| A1 | layout | Critical | open | P0 | Неявное мутабельное глобальное состояние заказа плит (`PlateMutableRuntime`, `config_and_data`, PEP 562) | `core/config_and_data.py:423–434` (`__getattr__` → `get_plate_mutable_runtime()`), `core/domain/plate_order.py:320–338` (`apply_to_globals`, deprecated), `viz_modules/layout_sequence/from_plan.py:194`; HTTP: `plate_runtime_isolation.py:12–17`; CPU: `cpu_bound.py:38–41`. PLATES_* в 6 live-файлах. **Partial**: middleware + CPU pool; globals живы | 2026-09-21 | |
| A2 | platform | Critical | open | P0 | In-process state (rate limit, counters) блокирует несколько воркеров | `login_rate_limit.py:84–85` (`_SlidingWindowRateLimiter`), `:179–191` `enforce_single_instance_workers` fail-fast только production+single_instance; `main.py:46–50`. **Partial NEW**: `draft_store.py:35–40` → FS + shared volume; rate limit in-memory | 2026-09-21 | |
| A3 | kp | High | resolved | — | God-module `CommercialWorkflowService` | Facade 802 строк (`commercial_workflow_service.py:43–66` delegates); identity / plate-resolve / lifecycle / ProductDraftHandler вынесены | 2026-09-21 | |
| A4 | kp | High | open | P1 | Толстый API-слой: `commercial.py` разобран; remainder `production.py` (698 строк, 26 routes) | `production.py` wc=698; `build_plan_from_filters` :104–180; day-documents :434–521 bypass ProductionService; SGP CRUD :535–608 | 2026-09-21 | |
| A5 | kp | High | open | P1 | Сервисы обходят repository, raw SQL | `sgp_service.py:78–89` (cur.execute + f-string WHERE), `delivery_schedule_service.py:494–500`, `kp_readiness_service.py:49–114`; нет *sgp* / *delivery* / *readiness* repos | 2026-09-21 | |
| A6 | layout | High | open | P1 | Planning импортирует visualization / matplotlib at load | `core/production/planning.py:39–44`, `core/visualization/__init__.py:16–18` (`matplotlib.use('Agg')` at import) | 2026-09-21 | |
| A7 | platform | High | open | P1 | Неполный DI в FastAPI (фабрики без constructor injection) | `app/dependencies/services.py:120–121` `return SupplyContractService(db_path=…)`; `SupplyContractService.__init__` `:94–108` builds repos inside. No Depends constructor injection | 2026-10-01 | |
| A8 | layout | Medium | open | — | Параллельные подсистемы планирования | `planning.py`, `plan_manager.py`, `plan_distribution*` | 2026-08-28 | |
| A9 | platform | Medium | open | — | Пустые app-сервисы-реэкспорты | `kp_persistence_service.py`, `rest_matching_service.py` | 2026-08-28 | |
| A10 | bot | Medium | open | — | Legacy-пути бота в persistence планов | `plan_storage.py`; не живой продукт | 2026-08-28 | |
| A11 | kp | Medium | open | — | ArchiveService god-orchestrator | `archive_service.py` **1892** строк, **87** methods (было 985/30). Lazy: CommercialWorkflowService `:184–187`, KpReadinessService `:200–204`, `:1282–1286`, SgpService `:1164–1167`. Блоки invoice `:363–621`, specification `:629–833`, promise `:903+`, `_to_details` `:1214–1349`. Split: ArchiveInvoiceService + ArchiveSpecificationService | 2026-10-01 | |
| A12 | kp | Medium | open | — | Frontend god-hook мастера КП | `useCommercialOfferWizard.ts` **379** строк, **19** useMutation (было 358/18). Mutations `:68–337`; split draft `:68–279` vs plate `:97–209` | 2026-10-01 | |
| A13 | — | — | — | — | **дырка** в нумерации 2026-08-28 — не занимать | | | |
| A14 | layout | Low | open | — | Монолит `core/visualization`, matplotlib `use('Agg')` at import | `core/visualization/__init__.py` | 2026-08-28 | |
| A15 | kp | Low | by-design | — | `owner_user_id` не в policy доступа КП | `offer_access.py:26–32`; [offer-access-policy.md](../architecture/offer-access-policy.md); связано с S1 | 2026-09-21 | |
| A16 | gsm | High | open | P1 | God-сервис жизненного цикла ПЛ (CRUD + confirm + rechain) | `gsm_generation_service.py` wc=**1030**. **Partial**: `gsm_kit_gate` извлечён | 2026-09-21 | gsm-audit A1 |
| A17 | gsm | High | open | P1 | God-модуль `core/gsm/generator.py` | `core/gsm/generator.py` wc=**1324** | 2026-09-21 | gsm-audit A2 |
| A18 | gsm | High | open | P1 | Блокирующий LibreOffice/`soffice` в HTTP (и DoS) | `gsm_export_service.py:64–70` sync `subprocess.run`; `gsm.py:644–660` `export_waybills` → sync export_zip | 2026-09-21 | gsm-audit A3, S1 |
| A19 | gsm | High | open | P1 | Сезонная логика дублируется на фронте, расхождение контракта | `core/gsm/season.py:35–43` vs `downstreamPreview.ts:35–44` | 2026-09-21 | gsm-audit A4 |
| A20 | gsm | High | open | P1 | Импорт транзакций без unit-of-work | `gsm_transaction_service.py:82–114`, `gsm_repository.py:391–424` per-row commit | 2026-09-21 | gsm-audit A5 |
| A21 | gsm | Medium | open | — | `gsm_kit_gate` импортирует private `_chain_broken` из overview | `gsm_kit_gate.py:10` import, `:120` use | 2026-09-21 | |
| A22 | kp | High | open | P1 | God-сервис `ProductionCompletionService` + raw sqlite3 | `production_completion_service.py` wc=702; sqlite3 :4, connect :125, ops :301+, :469+ | 2026-09-21 | |
| A23 | kp | High | open | P1 | God-module `DeliveryScheduleService` ~878 LOC, ~24 methods | `delivery_schedule_service.py` wc=878; replace/get/import_draft/generate_document/_enrich_with_traffic_light; PUT→get() re-runs I/O; нет repo | 2026-09-21 | |
| A24 | layout | High | open | P1 | `concrete_grade` теряется в пайплайне планирования | Точка потери: `viz_modules/layout_sequence/from_plan.py:676,747,944` (копируют `kp_id`/`plate_name`, не `concrete_grade`; cut несёт марку — `finalize.py:208`); skip-пути `core/plate_attribution.py:378–383` + rescue-ветка `:361–370`; legacy-фолбэк `day_view_service.py:415–421`; live: 0/558 root items с маркой vs 35/35 secondary в `plan_20260904_124248` | 2026-09-21 | |
| A25 | layout | Critical | open | P0 | Сироты/призраки: двойная атрибуция одной геометрии к разным КП (assignment-side vs track-side, нешаренные consumed) | Источник: `core/plate_attribution.py:111–114` (два взгляда, разные правила выбора); фикс-канал: unit_id (`finalize.py:224–226` ↔ `from_plan.py:693–694`). Кейс `plan_20260904_124248`: «ПБ 20,6-7,2-8п» заказано КП#4=3+КП#5=1, атрибуция 2/2 → сирота `kp_plates.id=197` + призрак КП#5 (3-й трек 2026-09-14). Симптом: `plan_commit.py:614–641` (pro-rated), `:650` (pool guard), `:703–708` (silent continue), `:711–743` (orphan-rollback мимо NULL-day) | 2026-09-21 | |
| A26 | layout | High | open | P1 | Слепота детекции surplus: coverage игнорирует пере-покрытие (ILP невиновен — доказано) | `core/optimization/coverage_verify.py:60` (`ok = not missing`, результат только в лог); итерация только по ключам спроса; рассинхрон нормализации ключей с post-correction. ILP не может перепроизводить: `ilp_model.py:261` (`==`), все vars `LpInteger`; живых surplus не найдено (4=4). `extract_cuts.py:61–62/:119/:140` ceil/round — латентный дефект. Surplus с identity уже роняет коммит (`plan_commit.py:494–502`); молчаливый канал — rescue-leftovers `:504–514` | 2026-09-21 | |
| A27 | kp | Medium | open | — | `update_contract_date` пишет в БД в обход `_assert_writer` (supply-contract) | `supply_contract_service.py:376–395` writes DB; `_assert_writer` `:397–398` not called. No HTTP route. Test happy-path ADMIN `tests/test_supply_contract.py:390–394` | 2026-10-01 | |
| A28 | kp | Medium | open | — | ArchiveService обходит SupplyContractService, зовёт SupplyContractRepository | `archive_service.py:527–529`, `:589–591`, `:746–748`. Action: read-port или `SupplyContractService.get_for_kp` | 2026-10-01 | |
| S1 | kp | High | by-design | — | Менеджер видит чужие КП — общий архив (не IDOR для fix) | `offer_access.py:26–32`; ADR offer-access-policy.md; тесты authorization | 2026-09-21 | |
| S2 | platform | High | resolved | — | CVE Starlette/FastAPI | fastapi 0.141.1, starlette 1.6.0; pip-audit clean | 2026-09-21 | |
| S3 | platform | High | open | P1 | Rate limiting in-process (см. A2) | `commercial_upload_validation.py:23–45`, `:53–57`; `prepare_commercial_ocr_upload` `:141–142` (`commercial.py`). In-process, not shared across workers | 2026-10-01 | |
| S4 | platform | High | resolved | — | npm audit high | high=0; 4 moderate (vitest/uuid/exceljs) отложены | 2026-09-21 | |
| S5 | kp | High | open | P1 | OCR / vision во внешние LLM; supply-contract parse обходит `ensure_external_ocr_enabled` и rate limit | Gate `commercial_upload_validation.py:60–66`, `:141–142`. Route `archive.py:562–576` → `parse_for_kp`. Service `supply_contract_service.py:307–319`, `:328–356`, `:358–369`, `:604–622`; grep: no ensure/check in service. При `OCR_EXTERNAL_ENABLED=false` коммерческий parse блокируется, supply-contract — нет | 2026-10-01 | |
| S6 | platform | Medium | open | — | SQLite без шифрования at rest | локальный завод | 2026-09-21 | |
| S7 | auth | Medium | open | — | CSP Report-Only + unsafe-inline | `security_headers.py:11–36` | 2026-09-21 | |
| S8 | platform | Medium | open | — | Утечка деталей ошибок в HTTP | `detail=str(exc)`: `archive.py` ~35 (напр. `:201`, `:263–275`, `:515–521`, `:587`); `commercial.py` ~7 (`:86`, `:95–97`, `:181`, `:528`, `:645–647`). Parse partly `raise_unexpected_server_error` `archive.py:589–592` | 2026-10-01 | |
| S9 | auth | Medium | open | — | CSRF парсит multipart до проверки токена | `app/middleware/csrf.py:41–47` `await request.form()` if header missing. Archive upload/parse `archive.py:166–179`, `:194`, `:569`. FE `httpClient.ts:130–132` | 2026-10-01 | |
| S10 | auth | Medium | open | — | Сессия 12ч без refresh-ротации | `settings.py:63–65` session_ttl_seconds=43200; `session.py:14–20`, `:29–41` | 2026-09-21 | |
| S11 | auth | Low | by-design | — | CSRF-cookie не HttpOnly (double-submit) | ожидаемо для паттерна | 2026-09-21 | |
| S12 | auth | Low | open | — | Password policy messages на английском | auth schemas | 2026-09-21 | |
| S13 | auth | Low | open | — | `/health` метаданные вне production | health endpoint | 2026-09-21 | |
| S14 | bot | Low | open | — | Legacy bot auth bypass при `BOT_AUTH_ENABLED=false` | `bot_archived/`; не живой контур | 2026-08-28 | |
| S15 | kp | Low | open | — | Draft в sessionStorage (XSS-вектор черновика) | `frontend/src/features/commercial-offer/store/draftStorage.ts:3`, `:18–22` sessionStorage | 2026-10-01 | |
| S16 | gsm | High | open | P1 | Импорт без лимита файлов (DoS); per-file 50MB есть | `gsm.py:171–181` — loop `for upload in files:` без `len(files)` cap; per-file 50MB: `settings.py:145–146` | 2026-09-21 | gsm-audit S2 |
| S17 | kp | Medium | open | — | XLSX import: no magic bytes / no row cap | DS half not rechecked. **supply-contract**: `supply_contract_import.py:35–47` `load_workbook` + unbounded `iter_rows`; file size only `read_upload_file_capped` `archive.py:194` | 2026-10-01 | |
| S18 | kp | Medium | open | — | Formula injection в генерируемом XLSX | `core/delivery_schedule_xlsx.py` export — значения с `=`, `+`, `-`, `@` | 2026-09-21 | |
| S19 | kp | Medium | open | — | Unbounded PUT batches/items/name | `app/schemas/delivery_schedule.py` — нет `max_length` / `max_items` на PUT payload | 2026-09-21 | |
| Q1 | kp | High | resolved | — | Шесть копий product-type pipeline | `product_draft_config.py` + `product_draft_handler.py` | 2026-09-21 | |
| Q2 | kp | High | resolved | — | Copy-paste HTTP-обработчиков КП | `commercial.py:98–319` runners; поглощён A4 | 2026-09-21 | |
| Q3 | layout | High | resolved | — | ~720 строк мёртвого кода `build_layout_sequence` | builder 991→~260; hash sequence `77c686bd…` | 2026-09-21 | |
| Q4 | kp | Medium | open | — | Пять копий `build_*_preview_metadata` | `commercial_draft_service.py:273,346,419,492,565` (файл 838 строк) | 2026-10-01 | |
| Q5 | kp | Medium | resolved | — | Дублирование resolve_wide/unpriced plates | `commercial_plate_resolve.py:66`, `:80` | 2026-09-21 | |
| Q6 | layout | Medium | open | — | Две реализации `get_global_calendar_info` | `plan_calendar.py:70` vs `plan_distribution_service.py:135`; callers `delivery_schedule_service.py:603`, `archive_service.py:903`; R2 fallback vs `plan_storage` MAX_TRACKS | 2026-09-21 | |
| Q7 | kp | Medium | open | — | Product-type duplication на фронте | partial `productTypeConfig.ts`; remainder `CalculationResultStep.tsx:52–74` | 2026-10-01 | |
| Q8 | kp | Medium | open | — | God-hook `useCreatePlanWizardState` | wc=997; `eslint-disable` :256, :368 | 2026-09-21 | |
| Q9 | kp | Medium | open | — | God-component `OfferDetailsDrawer` | **1658** строк (было 1369) | 2026-10-01 | |
| Q10 | kp | Medium | open | — | Слабая типизация production API | 9 routes без `response_model`: `production.py:87,95,220,266,321,385,525,611,619` | 2026-09-21 | |
| Q11 | kp | Medium | open | — | preview: `Any` / `dict[str, Any]` | `commercial_draft_service.py:192,276,349,422,495`; `schemas/commercial.py:316–323` | 2026-10-01 | |
| Q12 | kp | Medium | open | — | ArchiveService скрывает частичные сбои | `_to_list_item` `:1160–1173`, `_to_details` `:1287–1288` swallow; `move_to_production` re-raise `:887–891`; `_invoice_export_block` `:592–619` | 2026-10-01 | |
| Q13 | layout | Medium | open | — | Нет прямых тестов `get_global_calendar_info` | mocks в `test_delivery_schedule_service.py`; после консолидации Q6 | 2026-09-21 | |
| Q14 | platform | Low | open | — | Однострочные delegate-обёртки | services | 2026-08-28 | |
| Q15 | layout | Low | open | — | Имя `_merge_plate_texts` вводит в заблуждение | layout utils | 2026-08-28 | |
| Q16 | kp | Low | open | — | `/parse` без `response_model` | `commercial.py:185–190` returns `dict`, no response_model | 2026-10-01 | |
| Q17 | gsm | Low | open | — | `GsmGenerationError` messages на английском | `gsm_generation_service.py:78–119,659` | 2026-09-21 | |
| Q18 | kp | Low | open | — | Подавление `react-hooks/exhaustive-deps` в 10 файлах | in-scope: `MoveToProductionDialog.tsx:87`, `CounterpartyAutocomplete.tsx:78`. Production wizard half not rechecked | 2026-10-01 | |
| Q19 | gsm | High | open | P1 | Расхождение округления литров: `Math.round` vs banker's round | `downstreamPreview.ts:23–26` vs `balance.py:15–17`; burn 25km 10.1l → 2.53 vs 2.52; `tests/test_gsm_balance.py:73–77` | 2026-09-21 | gsm-audit Q1 |
| Q20 | gsm | High | open | P1 | Несогласованная обработка битого `season_switches` | `gsm_registry_service.py:356–361` swallows; `gsm_generation_service.py:703–711` throws; `tests/test_gsm_season.py:355–360`. **Partial** | 2026-09-21 | gsm-audit Q2 |
| Q21 | gsm | Low | open | — | Две реализации `formatLiters` на фронте GSM | `importReport.ts:28–29` vs `waybillWarnings.ts:80–84` | 2026-08-28 | |
| Q22 | kp | Medium | open | — | Дублирование `is_*_draft` в трёх модулях | canon `commercial_calculation_service.py:42–58`; copy `commercial_export_service.py:212–234`; delegate `commercial_wizard_step_service.py:52–65` | 2026-10-01 | |
| Q23 | kp | Medium | open | — | Расхождение оценки дорожек: archive `round` vs production FE `ceil` | archive `archive_service.py:927` `round(…)`. FE `productionEstimate.ts` NOT remeasured this run | 2026-10-01 | |
| Q24 | kp | Medium | open | — | Copy-paste day-document handlers: мёртвый raise после raise_not_found | `production.py:486–489` breakdown, `:513–516` formovka vs schema `:447–458` fixed (partial) | 2026-09-21 | |
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

A18 = gsm LibreOffice+DoS (старые gsm A3+S1). Не плодить второй ID.
A21 (Medium), Q21 (Low) — не входят в P1.
Q26 не занят (пропущен).

Следующий свободный ID: **A29**, **S20**, **Q40**.
