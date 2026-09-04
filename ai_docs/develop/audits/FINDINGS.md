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
GSM High с [2026-08-26-gsm-audit.md](./2026-08-26-gsm-audit.md) перенумерованы (там были свои A1…, конфликт с этим реестром) — колонка Legacy.

Последнее обновление: 2026-09-04 (прогон kp — конструктор КП).

| ID | Scope | Sev | Status | P | Summary | Evidence / notes | Last seen | Legacy |
|----|-------|-----|--------|---|---------|------------------|-----------|--------|
| A1 | layout | Critical | open | P0 | Неявное мутабельное глобальное состояние заказа плит (`PlateMutableRuntime`, `config_and_data`, PEP 562) | `core/config_and_data.py:383–411`, `core/domain/plate_order.py:331–383`, `viz_modules/layout_sequence/from_plan.py:194`; HTTP: `plate_runtime_isolation.py:12–17`; CPU: `cpu_bound.py:38–41`. **Partial**: middleware + CPU pool; globals живы | 2026-08-28 | |
| A2 | platform | Critical | open | P0 | In-process state (rate limit, drafts, counters) блокирует несколько воркеров | `login_rate_limit.py:84–85`, `draft_store.py:35–40`, `main.py:46–49`. **Partial**: fail-fast production+single_instance; shared store нет | 2026-08-28 | |
| A3 | kp | High | resolved | — | God-module `CommercialWorkflowService` | Facade 780 строк (`commercial_workflow_service.py:38–66` 47 delegates, `:629–636` hydrate); identity / plate-resolve / lifecycle / ProductDraftHandler вынесены (1001 строк) | 2026-09-04 | |
| A4 | kp | High | open | P1 | Толстый API-слой: `commercial.py` разобран; remainder `production.py` (642 строк, 26 routes) | `production.py` wc=642; `build_plan_from_filters` :104–167; day-documents :400–488 bypass ProductionService; SGP CRUD :501–574; `_date_range_inclusive` :594–600 | 2026-09-04 | |
| A5 | kp | High | open | P1 | Сервисы обходят repository, raw SQL | `sgp_service.py:78–89`, `delivery_schedule_service.py:494–500`, `kp_readiness_service.py:49–59`; нет *sgp* / *delivery* / *readiness* repos | 2026-09-04 | |
| A6 | layout | High | open | P1 | Planning импортирует visualization / matplotlib at load | `core/production/planning.py:39–44`, `core/visualization/__init__.py:16–18` | 2026-08-28 | |
| A7 | platform | High | open | P1 | Неполный DI в FastAPI (фабрики без constructor injection) | `app/dependencies/services.py:41–156` vs `get_auth_service` :159–162 | 2026-08-28 | |
| A8 | layout | Medium | open | — | Параллельные подсистемы планирования | `planning.py`, `plan_manager.py`, `plan_distribution*` | 2026-08-28 | |
| A9 | platform | Medium | open | — | Пустые app-сервисы-реэкспорты | `kp_persistence_service.py`, `rest_matching_service.py` | 2026-08-28 | |
| A10 | bot | Medium | open | — | Legacy-пути бота в persistence планов | `plan_storage.py`; не живой продукт | 2026-08-28 | |
| A11 | kp | Medium | open | — | ArchiveService god-orchestrator | `archive_service.py` 862 строк, 30 methods; lazy-coupling :124 CommercialWorkflowService, :140 KpReadinessService, :477 SgpService | 2026-09-04 | |
| A12 | kp | Medium | open | — | Frontend god-hook мастера КП | `useCommercialOfferWizard.ts` 517 строк, 28 useMutation :72–471 | 2026-09-04 | |
| A13 | — | — | — | — | **дырка** в нумерации 2026-08-28 — не занимать | | | |
| A14 | layout | Low | open | — | Монолит `core/visualization`, matplotlib `use('Agg')` at import | `core/visualization/__init__.py` | 2026-08-28 | |
| A15 | kp | Low | by-design | — | `owner_user_id` не в policy доступа КП | `offer_access.py:26–32`; [offer-access-policy.md](../architecture/offer-access-policy.md); связано с S1 | 2026-09-04 | |
| A16 | gsm | High | open | P1 | God-сервис жизненного цикла ПЛ (CRUD + confirm + rechain) | `gsm_generation_service.py` ~1030 строк. **Partial**: `gsm_kit_gate` извлечён | 2026-08-28 | gsm-audit A1 |
| A17 | gsm | High | open | P1 | God-модуль `core/gsm/generator.py` | `core/gsm/generator.py` ~1324 строк | 2026-08-28 | gsm-audit A2 |
| A18 | gsm | High | open | P1 | Блокирующий LibreOffice/`soffice` в HTTP (и DoS) | `gsm_export_service.py:64–70`, `gsm.py:620–662` | 2026-08-28 | gsm-audit A3, S1 |
| A19 | gsm | High | open | P1 | Сезонная логика дублируется на фронте, расхождение контракта | `core/gsm/season.py:35–43` vs `downstreamPreview.ts:28–45` | 2026-08-28 | gsm-audit A4 |
| A20 | gsm | High | open | P1 | Импорт транзакций без unit-of-work | `gsm_transaction_service.py:82–114`, `gsm_repository.py:391–424` per-row commit | 2026-08-28 | gsm-audit A5 |
| A21 | gsm | Medium | open | — | `gsm_kit_gate` импортирует private `_chain_broken` из overview | `gsm_kit_gate.py:10` import, `:120` use | 2026-08-28 | |
| S1 | kp | High | by-design | — | Менеджер видит чужие КП — общий архив (не IDOR для fix) | `offer_access.py:26–32`; ADR offer-access-policy.md; тесты authorization | 2026-09-04 | |
| S2 | platform | High | resolved | — | CVE Starlette/FastAPI | fastapi 0.141.1, starlette 1.6.0 | 2026-08-28 | |
| S3 | platform | High | open | P1 | Rate limiting in-process (см. A2) | `login_rate_limit.py:84–85`; kp-touchpoint: OCR rate limit `commercial_upload_validation.py:23–45`, `:141–142` | 2026-09-04 | |
| S4 | platform | High | resolved | — | npm audit high | high→0 2026-08-28; uuid/exceljs moderate отложен | 2026-08-28 | |
| S5 | kp | Medium | open | — | OCR шлёт изображения во внешние LLM | `OCR_EXTERNAL_ENABLED=true`: `openai.py:355–367`, `pipeline.py:50–53`, `commercial_upload_validation.py:60–66`, `commercial.py:102–104`, `:328–329`; `tests/test_commercial_ocr_policy.py` | 2026-09-04 | |
| S6 | platform | Medium | open | — | SQLite без шифрования at rest | локальный завод | 2026-08-28 | |
| S7 | auth | Medium | open | — | CSP Report-Only + unsafe-inline | `security_headers.py` | 2026-08-28 | |
| S8 | platform | Medium | open | — | Утечка деталей ошибок в HTTP | kp evidence: `production.py:140`, `:190`, `:196`, `:640`; `archive.py:296`, `:324` — `str(exc)` via `raise_unprocessable_client_error` | 2026-09-04 | |
| S9 | auth | Medium | open | — | CSRF парсит multipart до проверки токена | kp evidence: `csrf.py:42–47`; `commercial.py:194–202`, `:321–329`; `httpClient.ts:130–132` | 2026-09-04 | |
| S10 | auth | Medium | open | — | Сессия 12ч без refresh-ротации | `session.py` | 2026-08-28 | |
| S11 | auth | Low | by-design | — | CSRF-cookie не HttpOnly (double-submit) | ожидаемо для паттерна | 2026-08-28 | |
| S12 | auth | Low | open | — | Password policy messages на английском | auth schemas | 2026-08-28 | |
| S13 | auth | Low | open | — | `/health` метаданные вне production | health endpoint | 2026-08-28 | |
| S14 | bot | Low | open | — | Legacy bot auth bypass при `BOT_AUTH_ENABLED=false` | `bot_archived/`; не живой контур | 2026-08-28 | |
| S15 | kp | Low | open | — | Draft в sessionStorage (XSS-вектор черновика) | `draftStorage.ts:3`, `:18–19` | 2026-09-04 | |
| S16 | gsm | High | open | P1 | Импорт без лимита файлов (DoS); per-file 50MB есть | `gsm.py:147–166` — нет cap на число файлов | 2026-08-28 | gsm-audit S2 |
| Q1 | kp | High | resolved | — | Шесть копий product-type pipeline | `product_draft_config.py` + `product_draft_handler.py` | 2026-09-04 | |
| Q2 | kp | High | resolved | — | Copy-paste HTTP-обработчиков КП | `commercial.py:98–319` runners; поглощён A4 | 2026-09-04 | |
| Q3 | layout | High | resolved | — | ~720 строк мёртвого кода `build_layout_sequence` | builder 991→~260; hash sequence `77c686bd…` | 2026-08-28 | |
| Q4 | kp | Medium | open | — | Пять копий `build_*_preview_metadata` | `commercial_draft_service.py:270`, `:340`, `:410`, `:480`, `:550`; файл 820 строк | 2026-09-04 | |
| Q5 | kp | Medium | resolved | — | Дублирование resolve_wide/unpriced plates | `commercial_plate_resolve.py:66`, `:80` | 2026-09-04 | |
| Q6 | layout | Medium | open | — | Две реализации `get_global_calendar_info` | `plan_calendar.py` / `plan_distribution_service.py` | 2026-08-28 | |
| Q7 | kp | Medium | open | — | Product-type duplication на фронте | `commercialOfferApi.ts:145–243` + wizard `:81–275` | 2026-09-04 | |
| Q8 | kp | Medium | open | — | God-hook `useCreatePlanWizardState` | wc=724; `eslint-disable` :236 | 2026-09-04 | |
| Q9 | kp | Medium | open | — | God-component `OfferDetailsDrawer` | 1066 строк (было 901) | 2026-09-04 | |
| Q10 | kp | Medium | open | — | Слабая типизация production API | 9 routes без `response_model`: `production.py:86–90`, `:94–99`, `:207–212`, `:253–258`, `:308–312`, `:351–357`, `:491–496`, `:577–581`, `:585–590` | 2026-09-04 | |
| Q11 | kp | Medium | open | — | preview: `Any` / `dict[str, Any]` | `commercial_draft_service.py:192`, `:273`, `:343`, `:413`, `:483`, `:553`; `schemas/commercial.py:311–318` | 2026-09-04 | |
| Q12 | kp | Medium | open | — | ArchiveService скрывает частичные сбои | `:473–487`, `:541–542`, `:579–580`, `:783–785` swallow; `move_to_production` :243–247 re-raises | 2026-09-04 | |
| Q13 | layout | Medium | open | — | Нет прямых тестов `get_global_calendar_info` | после консолидации Q6 | 2026-08-28 | |
| Q14 | platform | Low | open | — | Однострочные delegate-обёртки | services | 2026-08-28 | |
| Q15 | layout | Low | open | — | Имя `_merge_plate_texts` вводит в заблуждение | layout utils | 2026-08-28 | |
| Q16 | kp | Low | open | — | `/parse` без `response_model` | `commercial.py:157` | 2026-09-04 | |
| Q17 | gsm | Low | open | — | `GsmGenerationError` messages на английском | generation service | 2026-08-28 | |
| Q18 | kp | Low | open | — | Подавление `react-hooks/exhaustive-deps` в 9 файлах | kp subset 3/9: `MoveToProductionDialog.tsx:66`, `DeliveryScheduleDialog.tsx:92`, `useCreatePlanWizardState.ts:236` | 2026-09-04 | |
| Q19 | gsm | High | open | P1 | Расхождение округления литров: `Math.round` vs banker's round | `downstreamPreview.ts:23–26` vs `balance.py:16–17`; burn 25km 10.1l → 2.53 vs 2.52 | 2026-08-28 | gsm-audit Q1 |
| Q20 | gsm | High | open | P1 | Несогласованная обработка битого `season_switches` | `gsm_registry_service.py:356–361` swallows; generation `703–711` throws. **Partial** | 2026-08-28 | gsm-audit Q2 |
| Q21 | gsm | Low | open | — | Две реализации `formatLiters` на фронте GSM | `importReport.ts:28–29` vs `waybillWarnings.ts:80–84` | 2026-08-28 | |
| Q22 | kp | Medium | open | — | Дублирование `is_*_draft` в трёх модулях | канон `commercial_calculation_service.py:28–45`; копия `commercial_export_service.py:178–191`; делегат `wizard_step_service.py:52–65` | 2026-09-04 | |

## Счётчики (прогон kp, 2026-09-04)

| Scope | Open P0 | Open P1 |
|-------|---------|---------|
| layout | 1 (A1) | 1 (A6) |
| platform | 1 (A2) | 2 (A7, S3) |
| kp | 0 | 2 (A4, A5) |
| gsm | 0 | 8 (A16–A20, S16, Q19, Q20) |
| auth | 0 | 0 |
| **все open Critical/High** | **2 P0** | **13 P1** (A4, A5, A6, A7, S3, A16–A20, S16, Q19, Q20) |

A18 = gsm LibreOffice+DoS (старые gsm A3+S1). Не плодить второй ID.
A21 (Medium), Q21 (Low) — не входят в P1.

Следующий свободный ID: **A22**, **S17**, **Q23**.
