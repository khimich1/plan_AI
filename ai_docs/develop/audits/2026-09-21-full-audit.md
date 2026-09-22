# Аудит: --full / весь проект

**Дата**: 2026-09-21  
**Скоуп**: полный проект (--full): критичные 20–30% live-кода — backend FastAPI (`app/`, `core/`), frontend React (`frontend/src/`), layout pipeline (`viz_modules/`), GSM, auth/platform, КП и производство; **исключено**: `bot/`, `bot_archived/`, `tests/archived/` как неживые контуры  
**Реестр**: [FINDINGS.md](./FINDINGS.md)  
**Прогон 0**: [2026-08-28-full-registry-sweep-audit.md](./2026-08-28-full-registry-sweep-audit.md)  
**Смежный прогон (реестр)**: [2026-09-21-delivery-schedule-audit.md](./2026-09-21-delivery-schedule-audit.md) — график поставки, same-day; ID A23, S17–S19, Q25, Q27–Q30 добавлены там, в этом --full не дублируются как NEW

## Дельта

| ID | Было | Стало | Суть |
|----|------|-------|------|
| A3 | resolved | resolved | Facade `commercial_workflow_service.py` 802 строк; delegates `:43–66`; регрессии нет |
| S2 | resolved | resolved | fastapi 0.141.1, starlette 1.6.0; pip-audit clean |
| S4 | resolved | resolved | npm audit high=0; 4 moderate (vitest/uuid/exceljs) отложены |
| Q1, Q2, Q3, Q5 | resolved | resolved | Без регрессии |
| S1, S11, A15 | by-design | by-design | Общий архив КП (`offer_access.py:26–32`); CSRF double-submit |
| A1 | open | open | **Partial**: middleware `:12–17`, CPU pool `:38–41`; PEP 562 globals живы `:423–434` |
| A2 | open | open | **Partial NEW**: `draft_store.py:35–40` → FS + shared volume; rate limit in-memory `:84–85` |
| A4 | open | open | `production.py` wc=698, 26 `@router`; day-documents bypass сервиса |
| A5 | open | open | Raw SQL без repo: sgp, delivery_schedule, kp_readiness |
| A6 | open | open | planning → visualization/matplotlib at import |
| A7 | open | open | Фабрики `return Service()` vs `get_auth_service` с Depends |
| A11 | open | open | ArchiveService **868→985** строк |
| A12 | open | open | **Partial**: wizard **517→358** строк, useMutation **28→18** |
| A16–A20, A22, A23 | open | open | God-modules GSM/KP подтверждены (wc snapshot ниже) |
| S3, S16 | open | open | In-process rate limit; GSM import без cap на число файлов |
| Q19, Q20 | open | open | Округление литров; season_switches policy |
| S17, S18, S19 | open | open | Не перепроверялись в --full; остаются open по DS-аудиту 2026-09-21 |
| A8, A9, A10, A14, S5–S10, S12–S15, Q4, Q6–Q18, Q21–Q30 | open | open | Medium/Low подтверждены или без изменения статуса |

- Закрыто: 0
- Открыто (подтверждено): 47 ID (включая resolved/by-design без регрессии)
- Новое: 0
- Не воспроизвелось: 3 (S17, S18, S19 — security --full не перепроверял DS Medium; в реестре статус **open** сохранён по уликам DS-аудита)

**Открытые P0 / P1 в скоупе**: 2 / 15

## Действия (максимум 8)

1. **[A1]** Завершить миграцию `PlateOrderContext` — убрать PEP 562 globals; покрыть BackgroundTasks/CLI/viz.
2. **[A2]/[S3]** Shared store для rate limit (и оставшихся in-process counters) **или** жёсткий single-worker contract во всех окружениях. Черновики уже на FS.
3. **[A18]** Вынести LibreOffice/`soffice` из HTTP-потока (очередь/worker); лимит параллельных конвертаций.
4. **[Q19]** Унифицировать округление литров: banker's round на фронте (`downstreamPreview.ts`) как в `balance.py`.
5. **[Q20]** Единая политика битого `season_switches`: throw в `get_settings` как в generation.
6. **[S16]** Лимит числа файлов импорта GSM (`gsm.py:171–181`); per-file 50MB уже есть.
7. **[A4]** Thin controller для `production.py` (`build_plan_from_filters`, day-documents, SGP CRUD).
8. **[A23]+[A5]** `DeliveryScheduleRepository` + разрезать `DeliveryScheduleService` (878 LOC).

## Открытые P0 / P1

### P0 — Critical

#### [A1] Неявное мутабельное глобальное состояние заказа плит

**Статус**: open (частичный прогресс)  
**Улика**: `core/config_and_data.py:423–434` (`__getattr__` → `get_plate_mutable_runtime()`), `core/domain/plate_order.py:320–338` (`apply_to_globals`, deprecated), `viz_modules/layout_sequence/from_plan.py:194`; partial HTTP: `app/middleware/plate_runtime_isolation.py:12–17`; CPU: `app/concurrency/cpu_bound.py:38–41`; PLATES_* в 6 live-файлах  
**Зачем**: Middleware и CPU-pool изолируют HTTP, но globals через PEP 562 живы — риск утечки состояния в BackgroundTasks, CLI и код без явного контекста

#### [A2] In-process state блокирует горизонтальное масштабирование

**Статус**: open (частичный прогресс)  
**Улика**: `app/security/login_rate_limit.py:84–85` (`_SlidingWindowRateLimiter`), `:179–191` `enforce_single_instance_workers` fail-fast только production+single_instance; `app/main.py:46–50`; **NEW vs 2026-08-28**: `draft_store.py:35–40` — FS + shared volume; rate limit по-прежнему in-memory  
**Зачем**: Rate limit и counters in-memory; fail-fast только в production — несколько workers/replicas ломают консистентность лимитов

---

### P1 — High

#### [A4] Толстый API-слой (remainder `production.py`)

**Статус**: open  
**Улика**: `app/api/v1/endpoints/production.py` wc=698, 26 `@router`; `build_plan_from_filters` `:104–180`; day-documents `:434–521` вызывают `generate_day_*` напрямую; SGP CRUD `:535–608`  
**Зачем**: Presentation дублирует orchestration; сложно тестировать и эволюционировать API производства

#### [A5] Сервисы обходят repository, raw SQL

**Статус**: open  
**Улика**: `app/services/sgp_service.py:78–89` (cur.execute + f-string WHERE), `delivery_schedule_service.py:494–500`, `kp_readiness_service.py:49–114`; grep `app/repositories` — нет *sgp* / *delivery* / *readiness* repos  
**Зачем**: SQL размазан по сервисам; нет единой границы persistence и изолированных unit-тестов

#### [A6] Planning импортирует visualization / matplotlib at load

**Статус**: open  
**Улика**: `core/production/planning.py:39–44`, `core/visualization/__init__.py:16–18` (`matplotlib.use('Agg')` at import)  
**Зачем**: Домен планирования тянет тяжёлый viz-стек при каждом импорте

#### [A7] Неполный DI в FastAPI

**Статус**: open  
**Улика**: `app/dependencies/services.py:45–185` — фабрики `return Service()` vs `get_auth_service` `:188–191` с `Depends`  
**Зачем**: Скрытая связность; граф зависимостей неединообразен и слабо тестируем

#### [A16] God-сервис жизненного цикла ПЛ GSM

**Статус**: open (частичный прогресс)  
**Улика**: `app/services/gsm_generation_service.py` wc=**1030**; `gsm_kit_gate` извлечён  
**Зачем**: CRUD + confirm + rechain в одном модуле — широкие регрессии при изменениях GSM

#### [A17] God-модуль `core/gsm/generator.py`

**Статус**: open  
**Улика**: `core/gsm/generator.py` wc=**1324**  
**Зачем**: Монолит генерации ПЛ; сложно тестировать изолированно

#### [A18] Блокирующий LibreOffice/`soffice` в HTTP

**Статус**: open  
**Улика**: `app/services/gsm_export_service.py:64–70` sync `subprocess.run`; `app/api/v1/endpoints/gsm.py:644–660` `export_waybills` → sync export_zip  
**Зачем**: Синхронная конвертация в request thread — DoS и таймауты при нагрузке

#### [A19] Сезонная логика дублируется на фронте

**Статус**: open  
**Улика**: `core/gsm/season.py:35–43` vs `frontend/src/features/gsm/lib/downstreamPreview.ts:35–44`  
**Зачем**: Расхождение контракта сезонных переключателей между API и UI

#### [A20] Импорт транзакций без unit-of-work

**Статус**: open  
**Улика**: `app/services/gsm_transaction_service.py:82–114` per-row insert; `app/repositories/gsm_repository.py:391–424` `conn.commit()` на каждый INSERT  
**Зачем**: Частичный импорт при сбое; нет атомарности batch

#### [A22] God-сервис `ProductionCompletionService` + raw sqlite3

**Статус**: open  
**Улика**: `app/services/production_completion_service.py` wc=**702**; sqlite3 `:4`, connect `:125`, ops `:301+`, `:469+`  
**Зачем**: God-service с прямым sqlite3 вместо repository-слоя

#### [A23] God-module `DeliveryScheduleService`

**Статус**: open  
**Улика**: `app/services/delivery_schedule_service.py` wc=**878**, ~24 метода; PUT→`get()` re-runs I/O; нет repo (подтверждено в --full и DS-аудите)  
**Зачем**: CRUD, светофор, импорт XLSX, документы и календарь в одном классе — высокая связность

#### [S3] Rate limiting in-process

**Статус**: open  
**Улика**: `app/security/login_rate_limit.py:84–85`; `commercial_upload_validation.py:23–45` OCR limiter, `:141–142`; partial: `main.py:46–50` single-instance  
**Зачем**: Brute-force и OCR лимиты обходятся при workers>1 (связано с A2)

#### [S16] Импорт GSM без лимита числа файлов

**Статус**: open  
**Улика**: `app/api/v1/endpoints/gsm.py:171–181` — loop `for upload in files:` без `len(files)` cap; per-file 50MB: `core/config/settings.py:145–146`  
**Зачем**: DoS через множество мелких upload при сохранении per-file лимита

#### [Q19] Расхождение округления литров

**Статус**: open  
**Улика**: `core/gsm/balance.py:15–17` vs `frontend/src/features/gsm/lib/downstreamPreview.ts:23–26`; burn 25km 10.1l → 2.52 vs 2.53; тест `tests/test_gsm_balance.py:73–77`  
**Зачем**: Preview и отчёт расходятся на граничных значениях

#### [Q20] Несогласованная обработка битого `season_switches`

**Статус**: open (частичный прогресс)  
**Улика**: `app/services/gsm_registry_service.py:356–361` swallows → `[]`; `gsm_generation_service.py:703–711` throws; `tests/test_gsm_season.py:355–360`  
**Зачем**: Разное поведение registry vs generation/export при повреждённых настройках

---

## Приложение

### God-module snapshot (wc)

| Модуль | Строк |
|--------|-------|
| `core/gsm/generator.py` | 1324 |
| `gsm_generation_service.py` | 1030 |
| `archive_service.py` | **985** (было 868) |
| `delivery_schedule_service.py` | 878 |
| `commercial_workflow_service.py` (facade) | 802 |
| `core/production/planning.py` | 790 |
| `production_completion_service.py` | 702 |
| `production.py` (endpoint) | 698 |

### Medium (action this week)

| ID | Scope | Summary | Улика |
|----|-------|---------|-------|
| A21 | gsm | kit_gate → private `_chain_broken` | `gsm_kit_gate.py:10`, `:120` |
| A11 | kp | ArchiveService god-orchestrator | `archive_service.py` **985** строк, 30 methods |
| A12 | kp | Frontend god-hook мастера КП | **358** строк (было 517), useMutation **18** (было 28) — partial |
| A8 | layout | Параллельные подсистемы планирования | `planning.py`, `plan_manager.py`, `plan_distribution*` |
| A9 | platform | Пустые app-сервисы-реэкспорты | `kp_persistence_service.py`, `rest_matching_service.py` |
| A10 | bot | Legacy bot paths | `plan_storage.py`; не живой контур |
| S5 | kp | OCR во внешние LLM | `openai.py:355–367`, `pipeline.py:50–53`, `commercial_upload_validation.py:60–66`, `commercial.py:113–115`, `:132–134`, `:229`; `tests/test_commercial_ocr_policy.py` |
| S6 | platform | SQLite без шифрования at rest | `plita.db`, `pb.db` |
| S7 | auth | CSP Report-Only + unsafe-inline | `security_headers.py:11–36` |
| S8 | platform | Утечка деталей ошибок в HTTP | `production.py` (~10), `archive.py` (~10), `delivery_schedule.py` (~8), `gsm.py` (~7), `commercial.py` (~6); пример `production.py:138–141`, `archive.py:276–279` |
| S9 | auth | CSRF парсит multipart до токена | `csrf.py:42–47`; `httpClient.ts:130–132` |
| S10 | auth | Сессия 12ч без refresh-ротации | `settings.py:63–65` session_ttl_seconds=43200; `session.py:14–20`, `:29–41` |
| S17 | kp | XLSX import без magic bytes / row cap | `core/delivery_schedule_xlsx.py`; не перепроверялось в --full security |
| S18 | kp | Formula injection в экспортируемом XLSX | `core/delivery_schedule_xlsx.py` export |
| S19 | kp | Unbounded PUT batches/items/name | `app/schemas/delivery_schedule.py` |
| Q4 | kp | Пять копий `build_*_preview_metadata` | `commercial_draft_service.py:271,342,413,484,555` (826 строк) |
| Q6 | layout | Две реализации `get_global_calendar_info` | `plan_calendar.py:70` vs `plan_distribution_service.py:135`; callers `delivery_schedule_service.py:603`, `archive_service.py:903` |
| Q7 | kp | Product-type duplication на фронте | partial: `productTypeConfig`; `commercialOfferApi.ts:156–173`, wizard `:76–130`, `CalculationResultStep.tsx:48–204` |
| Q8 | kp | God-hook `useCreatePlanWizardState` | wc=997; eslint-disable `:256`, `:368` |
| Q9 | kp | God-component `OfferDetailsDrawer` | **1091→1276** строк |
| Q10 | kp | Слабая типизация production API | `production.py:87,95,220,266,321,385,525,611,619` → dict, no response_model |
| Q11 | kp | preview: `Any` / `dict[str, Any]` | `commercial_draft_service.py:192,274,345,416,487,558`; `schemas/commercial.py:311–319` |
| Q12 | kp | ArchiveService скрывает частичные сбои | `archive_service.py:590–603,659,701` swallow; `move_to_production` re-raise `:326–332` |
| Q13 | layout | Нет прямых тестов `get_global_calendar_info` | mocks в `test_delivery_schedule_service.py` |
| Q22 | kp | Дублирование `is_*_draft` | канон `commercial_calculation_service.py:28–44`; копия `commercial_export_service.py:199–220`; делегат `commercial_wizard_step_service.py:52–65` |
| Q23 | kp | Расхождение оценки дорожек | `archive_service.py:368` round vs `productionEstimate.ts:24` ceil vs `delivery_schedule_check.py:199–200` ceil; 101m → 2 vs 1 |
| Q24 | kp | Copy-paste day-document handlers | `production.py:486–489` breakdown, `:513–516` formovka dead raise; schema `:447–458` fixed (partial) |
| Q25 | kp | `has_delivery_schedule` FE badge; list API без поля | `ArchiveOfferList.tsx:172`; `archive.ts:26`; `schemas/archive.py:25–57`; `archive_service.py:605–622` |
| Q27 | kp | `produce_by ≤ deliver_from` warning не реализован | `schemas/delivery_schedule.py:55–64` только deliver_from ≤ deliver_to |
| Q28 | kp | Три дублирующих SELECT из `kp_plates` | `delivery_schedule_service.py:496–500,682–686,703–707` |
| Q29 | kp | Нет service-тестов import_draft happy-path + GET red | gaps в `test_delivery_schedule_*` |
| Q30 | kp | Unmatched reasons на английском в UI | `delivery_schedule_xlsx.py:65–68`; `DeliveryScheduleEditor.tsx:160–161` |

### Low

| ID | Scope | Summary | Улика |
|----|-------|---------|-------|
| A14 | layout | Монолит `core/visualization`, matplotlib at import | `core/visualization/__init__.py` |
| S12 | auth | Password policy messages на английском | auth schemas |
| S13 | auth | `/health` метаданные вне production | health endpoint |
| S14 | bot | Legacy bot auth bypass | `bot_archived/`; не живой контур |
| S15 | kp | Draft в sessionStorage (XSS-вектор) | `draftStorage.ts:3`, `:18–22` |
| Q14 | platform | Однострочные delegate-обёртки | services |
| Q15 | layout | Имя `_merge_plate_texts` вводит в заблуждение | layout utils |
| Q16 | kp | `/parse` без `response_model` | `commercial.py:184–189` |
| Q17 | gsm | `GsmGenerationError` messages на английском | `gsm_generation_service.py:78–119,659` |
| Q18 | kp | Подавление `react-hooks/exhaustive-deps` | 10 файлов exhaustive-deps |
| Q21 | gsm | Две реализации `formatLiters` | `importReport.ts:28–29` vs `waybillWarnings.ts:80–84` |

### Подтверждённые resolved / by-design (без регрессии)

| ID | Статус | Суть |
|----|--------|------|
| A3 | resolved | Facade commercial workflow 802 строк; delegates `:43–66` |
| S2 | resolved | fastapi 0.141.1, starlette 1.6.0; pip-audit clean |
| S4 | resolved | npm audit high=0 |
| Q1, Q2, Q3, Q5 | resolved | product pipeline, runners, layout builder, plate resolve |
| S1, A15 | by-design | Общий архив КП — не IDOR для исправления |
| S11 | by-design | CSRF-cookie не HttpOnly (double-submit) |

### Частичный прогресс (still open)

- **A1**: HTTP middleware + CPU pool добавлены; PEP 562 globals и `apply_to_globals` живы.
- **A2**: Черновики мигрированы на FS + shared volume; rate limit и counters in-memory; fail-fast только production+single_instance.
- **A12**: Wizard hook уменьшен 517→358 строк, useMutation 28→18; god-hook остаётся.
- **A11**: ArchiveService вырос 868→985 строк — регресс по размеру.
- **A16**: `gsm_kit_gate` извлечён; generation ~1030 строк.
- **Q20**: generation/export бросают на corrupt JSON; registry по-прежнему глотает.

### Positive signals

- Draft ownership и path traversal guard на черновиках (FS store)
- `destructive_db_guard` на опасные операции БД
- ACL-тесты logistics/financial и GSM `REQUIRE_ACCOUNTING`
- HttpOnly session cookies; CSRF double-submit (S11 by-design)
- Ports/adapters для viz boundary (`core/ports/visualization.py`)
- GSM kit gate с unit-тестами (`test_gsm_kit_gate.py`)
- Delivery schedule: AuthZ GET/import, produced split, traffic_light_degraded (DS-аудит 2026-09-21)
