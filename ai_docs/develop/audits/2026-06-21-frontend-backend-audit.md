# Отчёт аудита: Frontend + Backend

**Дата:** 2026-06-21  
**Область:** `app/`, `core/`, `frontend/`, `viz_modules/` (релевантные `tests/` для контекста покрытия)  
**Исключено:** `bot/`, `bot_archived/`, `run_bot.py` — Telegram-бот постепенно отключается  
**Аудит провели:** senior-reviewer + security-auditor + reviewer

---

## Executive Summary

**Overall Health Score:** 0.0/10

| Severity | Architecture | Security | Code Quality | Total |
|----------|-------------|----------|--------------|-------|
| Critical | 3 | 0 | 0 | **3** |
| High | 9 | 4 | 4 | **17** |
| Medium | 7 | 8 | 9 | **24** |
| Low | 4 | 8 | 5 | **17** |

**Формула Health Score:** 10 − 6 (cap за critical) − 3 (cap за high) − 1 (cap за medium) = **0.0**

**Рекомендация:** Перед следующим релизом web-приложения необходимо устранить **3 критических архитектурных дефекта** — синхронная CPU-bound работа в HTTP workers, обход visualization ports и thread-local plate runtime. Эти проблемы создают риск деградации сервиса под нагрузкой, смешения данных между concurrent-запросами и нарушают границы модулей. Параллельно следует запланировать устранение **17 high-priority** находок по архитектуре, безопасности и качеству кода. Критических security-уязвимостей с прямым обходом аутентификации не обнаружено, но high-риски (rate limiting, OCR→OpenAI, frontend RBAC) требуют внимания в текущем спринте.

---

## Remediation status (P6 closure 2026-06-21)

> **Spec:** [`stabilizaciya-p6-architecture-2026-06-21.md`](../../specs/stabilizaciya-p6-architecture-2026-06-21.md) — **closed (WP9 stretch deferred)**  
> **Verify:** 953 pytest passed, 55 frontend tests, build OK

| Audit ID | Severity | Status | P6 WP / notes |
|----------|----------|--------|---------------|
| **A1** | Critical | **Resolved** | WP1 — `run_cpu_bound`, async hot endpoints |
| **A2** | Critical | **Resolved** | WP2 — `get_visualize_plan()` port on app hot paths |
| **A3** | Critical | **Resolved** | WP3 — explicit `PlateOrderContext`; no `get_plate_mutable_runtime` in `app/` |
| **S3** | High | **Resolved** | WP4 — fail-closed `destructive_db_guard` |
| **S4** | High | **Resolved** | WP4 — `RequireRole` route guards |
| **A10**, **S1** | High | **Mitigated** | WP7 — deploy contract [`deploy-contract.md`](../deploy-contract.md); startup warning on `workers > 1` |
| **S2** | High | **Mitigated** | WP7 — `OCR_EXTERNAL_ENABLED=false` default; opt-in for staging |
| **A4**, **A6**, **A11** | High | **Open → P7** | WP5 slices 4/6 deferred; god-modules partial |
| **A5** | High | **Partial** | WP5 slice 5 — `core/kp/offers_write.py` extracted |
| **A7**, **A8**, **A12** | High | **Open → P7** | WP6 planning/DI/legacy deferred |
| **A9** | High | **Partial** | WP6 — `AuthService`; endpoints use `Depends(get_auth_service)` |
| **Q1**–**Q4** | High | **Resolved** | WP8 — unified validation, `execution_terms`, hook tests |
| **A13**–**A19**, **S5**–**S12**, **Q5**–**Q13** | Medium | **Deferred** | WP9 stretch → P7 |

**Post-remediation Health Score (estimate):** ~6.5–7.0/10 — 0 critical, reduced high count; medium backlog remains for P7.

---

## Critical Issues (исправить немедленно)

### [A1] Sync CPU-bound работа блокирует HTTP workers

**Category:** Architecture  
**Location:** `app/api/v1/endpoints/commercial.py` (async handlers → sync `generate_preview`/`optimize`), `app/api/v1/endpoints/production.py` (`build_plan_from_filters` — sync), `app/services/optimization_service.py`, `app/services/commercial_workflow_service.py`  
**Impact:** ILP-оптимизация и сборка плана выполняются в worker-потоке FastAPI; под нагрузкой деградируют latency и throughput всех запросов на том же worker.  
**Fix:** Вынести тяжёлые вызовы в `asyncio.to_thread` / dedicated thread pool (как в `archive_service.py`) или job API с polling; для plan build — минимум `to_thread` + лимит concurrency.

---

### [A2] App обходит visualization ports для `visualize_plan`

**Category:** Architecture  
**Location:** `app/services/file_generation_service.py`, `app/services/archive_service.py`, `app/services/day_documents_service.py` → `from core.visualization import visualize_plan`; контракт ports: `core/ports/visualization.py`, wiring: `app/adapters/visualization.py`  
**Impact:** Граница core↔viz_modules нарушена на production hot paths; замена реализации и регрессионное тестирование усложняются (grep-gate покрывает только `viz_modules`, не `core.visualization`).  
**Fix:** Добавить port `visualize_plan` в `core/ports/visualization.py`, реализовать в `viz_modules/adapters/`, вызывать только через facade; app — только через port/adapter.

---

### [A3] Thread-local plate runtime остаётся SSOT для визуализации

**Category:** Architecture  
**Location:** `core/plate_runtime_state.py`, `core/visualization/__init__.py:96` (`get_plate_mutable_runtime()`), `core/optimization/context.py` (`OPT_CASCADING_PLAN` proxies), middleware: `app/middleware/plate_runtime_isolation.py`  
**Impact:** HTTP изолирован middleware, но `visualize_plan` и legacy core-пути жёстко завязаны на TLS/ContextVar; вызов вне `ctx.bound()` или без propagation в thread pool → риск смешения заказов (частично смягчено `run_in_order_context` в archive/day_documents).  
**Fix:** Рефактор `visualize_plan` на явный `PlateOrderContext`/snapshot parameter; deprecate `get_plate_mutable_runtime()` на web paths.

---

## High Priority Issues (исправить в ближайший спринт)

### Architecture

#### [A4] God module `CommercialWorkflowService`

**Category:** Architecture  
**Location:** `app/services/commercial_workflow_service.py` (~674 LOC)  
**Impact:** Orchestration OCR, draft, pricing, export, wizard steps в одном классе — нарушение SRP, высокий риск регрессий.  
**Fix:** Оставить thin orchestrator; вынести use-cases (`DraftLifecycleService`, `WizardStateService`, `OfferExportService`) с явными контрактами.

#### [A5] God module `core/kp_db_offers.py`

**Category:** Architecture  
**Location:** `core/kp_db_offers.py` (~581 LOC), потребители: `app/repositories/kp_repository.py`, `app/repositories/kp_offers_repository.py`  
**Impact:** Persistence, pricing side-effects и admin-операции в одном модуле; app repositories — thin pass-through без единой модели доступа.  
**Fix:** Разделить на `KpOffersRepository` (SQL/CRUD), `KpPricingService`, `KpAdminOperations`; app-слой зависит только от repository interfaces.

#### [A6] God module `core/visualization`

**Category:** Architecture  
**Location:** `core/visualization/__init__.py` (~593 LOC), `core/visualization/layout.py` (~667 LOC)  
**Impact:** Pricing, matplotlib, layout, runtime state в одном пакете; сложность тестов и рефакторинга.  
**Fix:** Split по доменам (`drawing/`, `layout/`, `export/`) с узким public API; `visualize_plan` — отдельный модуль с DI snapshot.

#### [A7] Дублирование planning между `app/planning/` и `core/production/`

**Category:** Architecture  
**Location:** `app/planning/plan_distribution.py`, `plan_storage.py`, `plan_manager.py`; `core/production/planning.py` (~709 LOC); оркестратор: `app/services/production_planning_service.py`, `app/services/plan_distribution_service.py`  
**Impact:** Две параллельные модели распределения/хранения планов; изменения требуют синхронизации, риск расхождения web-логики.  
**Fix:** SSOT в `core/production/planning.py` + ports; `app/planning/` — только thin adapters или deprecate.

#### [A8] DI без абстракций — сервисы self-construct зависимости

**Category:** Architecture  
**Location:** `app/dependencies/services.py`, `CommercialService.__init__`, `CommercialWorkflowService.__init__`, `ProductionPlanningService.__init__`  
**Impact:** Нарушение DIP; unit-tests требуют monkeypatch/module-level overrides; нельзя подменить repository без подмены всего сервиса.  
**Fix:** Protocol/ABC для repositories; constructor injection через `Depends`; единый composition root в `dependencies/services.py`.

#### [A9] Auth без application-слоя

**Category:** Architecture  
**Location:** `app/api/v1/endpoints/auth.py`, `app/dependencies/auth.py`, `app/repositories/auth_repository.py`, `app/web/legacy_routes.py`  
**Impact:** Login/register/change-password логика и rate-limit размазаны по endpoints; `AuthRepository()` создаётся inline в login/register (`auth.py:38,76`), минуя DI — дублирование и сложность тестирования.  
**Fix:** `AuthService` (authenticate, register, change_password, revoke_session); endpoints — только HTTP mapping + Depends.

#### [A10] In-process rate limiting не масштабируется

**Category:** Architecture  
**Location:** `app/security/login_rate_limit.py`, `app/services/commercial_upload_validation.py`  
**Impact:** При `workers > 1` лимиты умножаются на число процессов; архитектурный потолок горизонтального scaling.  
**Fix:** Redis/shared store или edge rate limiting; до миграции — enforce `workers=1` в deploy contract.

#### [A11] God modules production path

**Category:** Architecture  
**Location:** `app/services/production_completion_service.py` (~561 LOC), `app/services/day_view_service.py` (~522 LOC)  
**Impact:** Смешение aggregation, DB, file generation, business rules; сложно расширять без side effects.  
**Fix:** Декомпозиция по use-case (completion matching, day aggregation, document orchestration).

#### [A12] Legacy web routes параллельно SPA API

**Category:** Architecture  
**Location:** `app/web/legacy_routes.py`, `app/web/router.py`, `app/web/legacy_deprecation.py`  
**Impact:** Дублирование auth/commercial flows, двойная поверхность для поддержки и security review.  
**Fix:** Hard deprecate с redirect-only; удалить POST handlers после telemetry; единый API v1 path.

---

### Security

#### [S1] In-process rate limiting обходится при нескольких воркерах

**Category:** Security  
**Location:** `app/security/login_rate_limit.py`, `app/services/commercial_upload_validation.py`, `app/main.py`  
**Impact:** При `uvicorn --workers N>1` лимиты login/OCR/upload действуют per-process; brute-force паролей и злоупотребление OCR масштабируются пропорционально числу воркеров.  
**Fix:** Shared store (Redis) или rate limiting на reverse proxy/API gateway; в production — `--workers 1` до внедрения shared store.

#### [S2] Коммерческие документы и PII уходят во внешний OpenAI API

**Category:** Security  
**Location:** `core/ocr_gpt.py`, `app/api/v1/endpoints/commercial.py`, `app/services/commercial_upload_validation.py`  
**Impact:** JPEG/PNG/PDF с заказами, ценами и данными заказчиков передаются в GPT-4o Vision; риск утечки, нарушения 152-ФЗ/GDPR, отсутствие контроля над retention у провайдера.  
**Fix:** On-prem OCR, явный opt-in/consent, DPA с провайдером, redaction pipeline, feature-flag для отключения внешнего OCR в production.

#### [S3] Destructive DB reset разрешён при `APP_ENV=development`

**Category:** Security  
**Location:** `core/destructive_db_guard.py`, `app/api/v1/endpoints/admin.py`, `app/services/admin_service.py`  
**Impact:** При ошибочном деплое с `APP_ENV=development` админ может выполнить полное/частичное обнуление БД и планов без break-glass флагов.  
**Fix:** Fail-closed по умолчанию для всех non-local окружений; требовать оба флага (`ALLOW_DESTRUCTIVE_DB_RESET` + `DESTRUCTIVE_DB_RESET_BREAK_GLASS`) даже в staging; startup-check при `APP_ENV=production`.

#### [S4] Отсутствие route-level RBAC на фронтенде

**Category:** Security  
**Location:** `frontend/src/app/router/AppRouter.tsx`, `frontend/src/features/auth/components/ProtectedRoute.tsx`  
**Impact:** Любой аутентифицированный пользователь может открыть `/new`, `/archive`, `/production` напрямую; API отклоняет запросы (403), но UI и client-side state (sessionStorage) могут раскрывать структуру и кэшировать чувствительные данные wizard до ошибки API.  
**Fix:** Role-based route guards (`RequireRole`), редирект production → `/production`, manager/admin → commercial routes; не полагаться только на скрытие навигации в `AppHeader.tsx`.

---

### Code Quality

#### [Q1] Монолитная функция построения layout-sequence

**Category:** Code Quality  
**Location:** `viz_modules/layout_sequence/builder.py:27` (`build_layout_sequence`, ~965 строк, nesting до 10), связанные `viz_modules/layout_sequence/from_plan.py:183` (~837 строк)  
**Impact:** Любое изменение в визуализации/закупке требует правок в одном «комбайне»; высокий риск регрессий, сложный code review и невозможность точечного unit-тестирования веток.  
**Fix:** Разбить на фазы (загрузка плана → группировка плит → вторичные операции → сборка tracks) с отдельными функциями ≤30 строк и контрактными dataclass/TypedDict на границах.

#### [Q2] Дублирование правил валидации wizard с разной семантикой

**Category:** Code Quality  
**Location:** `app/services/commercial_calculation_service.py` (`wide_lines_blocking`, `meta_ready_for_calculate` vs `validate_calculate_prerequisites`), потребители: `app/services/commercial_wizard_step_service.py`, `app/services/commercial_workflow_service.py`  
**Impact:** Одни и те же бизнес-правила описаны дважды (bool vs raise); при изменении условий UI может показывать «можно продолжить», а `calculate_draft` упадёт с другой ошибкой. Тестов на `CommercialCalculationService` нет.  
**Fix:** Единый `validate_*` → список ошибок или Result; bool-методы — thin wrappers. Покрыть таблицей кейсов (пустые плиты, wide plates, manager, client terms).

#### [Q3] DRY: расходящаяся логика парсинга сроков исполнения

**Category:** Code Quality  
**Location:** `app/services/archive_service.py:437` (`_parse_execution_terms` — strict, `ArchiveValidationError`), `app/services/offers_service.py:180` (`_parse_execution_terms` — silent fallback +14 дней)  
**Impact:** Один и тот же доменный ввод обрабатывается по-разному: архив отклоняет невалидный срок, offers API молча подставляет дефолт. Риск тихих данных и расхождения UX между `/archive` и `/offers`.  
**Fix:** Общий helper в `core/execution_terms.py` или `ExecutionTermsService` с явной политикой (`strict` / `default_if_empty`); один набор тестов для обоих сервисов.

#### [Q4] Критичные frontend-хуки без тестов при высокой сложности

**Category:** Code Quality  
**Location:** `frontend/src/features/production/hooks/useCreatePlanWizardState.ts` (~442 строки), `frontend/src/features/commercial-offer/hooks/useCommercialOfferWizard.ts` (~201 строка); тестов: 7 файлов на ~105 исходников  
**Impact:** Логика выбора КП, conflict handling, step transitions и invalidation кэша не покрыты; регрессии wizard/plan-builder обнаруживаются только вручную.  
**Fix:** Unit-тесты с `@testing-library/react` + mocked React Query: переходы шагов, fillRequest, plan version conflict, мутации draft/plates/AI.

---

## Medium Priority Issues (запланировать на следующий спринт)

### Architecture

| ID | Проблема | Location | Impact | Fix |
|----|----------|----------|--------|-----|
| **A13** | Legacy PEP562 proxy `config_and_data` | `core/config_and_data.py`, потребители: `app/services/commercial_service.py`, `app/services/plate_parser_service.py`, множество core-модулей | Скрытые deprecated globals, неявные зависимости от runtime state; затрудняет статический анализ и decommission | Прямые импорты из `core/domain/`, `core/plate_order_context`; удалить `__getattr__` proxy по ADR |
| **A14** | `DraftStore` — file-based storage + ad-hoc instantiation | `app/services/draft_store.py`, `app/api/v1/endpoints/commercial.py:239,380` (`DraftStore()` inline), `app/dependencies/services.py:52` | Без shared volume multi-instance теряет консистентность черновиков; разные instances в workflow vs endpoint | Singleton через Depends `get_draft_store()`; для HA — shared storage + documented contract (`DRAFTS_DIR`) |
| **A15** | Frontend без route-level RBAC | `frontend/src/app/router/AppRouter.tsx`, `frontend/src/shared/lib/roleRoutes.ts`, `frontend/src/features/auth/components/ProtectedRoute.tsx` | Все authenticated roles видят все routes; защита только на API — UX leak и лишние 403 | `RequireRole` wrapper на `/new`, `/archive`, `/production`; redirect по `defaultRouteForRole` |
| **A16** | God hook `useCreatePlanWizardState` | `frontend/src/features/production/hooks/useCreatePlanWizardState.ts` (~443 LOC) | UI state, queries, mutations, fill-mode, calendar logic в одном hook — сложность тестирования и изменений | Split на `usePlanWizardSteps`, `useKpSelection`, `usePlanCalendar`, `useBuildPlanMutation` state |
| **A17** | Repositories смешивают raw SQL и core/kp_db | `app/repositories/kp_repository.py` (inline SQL + `kp_db_offers.*`), `app/repositories/plan_repository.py` (`core.kp_db_common._connect`) | Leaky abstraction; schema changes требуют правок в нескольких слоях; N+1 в `list_kps_in_production` | Весь SQL в repository methods; batch queries; запрет прямого `_connect` из app |
| **A18** | Неполные `response_model` на API | `app/api/v1/endpoints/commercial.py` (`/parse`, `/generate-preview` → `dict`), `app/api/v1/endpoints/production.py` (`get_plan`, `activate_plan` → `dict`) | OpenAPI contract drift, слабая типизация frontend integration | Pydantic schemas для всех public endpoints |
| **A19** | `OptimizationService.legacy_runtime` и module-level OPT proxies | `app/services/optimization_service.py:76-97`, `core/optimization/context.py` | Deprecated path мутирует module-level state; риск accidental use в новом коде | Удалить `legacy_runtime`; optimization только через `PlateOrderContext` + immutable snapshots |

### Security

| ID | Проблема | Location | Impact | Fix |
|----|----------|----------|--------|-----|
| **S5** | CSRF-cookie доступен JavaScript (не HttpOnly) | `app/security/csrf.py`, `app/middleware/csrf.py`, `frontend/src/shared/api/httpClient.ts` | При XSS злоумышленник читает `csrf_token` из `document.cookie` и выполняет state-changing запросы | Ужесточить CSP (enforce, без `unsafe-inline`); SameSite=Strict для CSRF-cookie |
| **S6** | Чувствительные данные КП в `sessionStorage` | `frontend/src/features/commercial-offer/store/draftStorage.ts` | Состояние wizard доступно на shared-станциях и при XSS | Хранить только `draft_id`; данные — на сервере (`DraftStore`); очищать при logout |
| **S7** | Черновики КП на FS без ACL per-user | `app/services/draft_store.py`, `.app_data/drafts/` | При компрометации сервера все JSON-черновики читаемы | Права каталога (700); шифрование at-rest или хранение в БД |
| **S8** | SQLite без шифрования at-rest | `core/config/settings.py` (`plita.db`, `pb.db`), `app/repositories/auth_repository.py` | Пароли, КП, производственные данные доступны при backup-доступе | SQLCipher/OS-level encryption; защита backup-файлов |
| **S9** | CSP только Report-Only с `unsafe-inline` | `app/middleware/security_headers.py` | XSS не блокируется браузером | Enforcing CSP; nonce/hash для Vite bundle |
| **S10** | Дублирующая legacy auth surface | `app/web/legacy_routes.py` (`POST /web/login`, `POST /web/offers/new`) | Параллельные endpoints увеличивают attack surface | Удалить legacy POST-handlers; redirect stubs |
| **S11** | Нет audit log для security-sensitive операций | `app/api/v1/endpoints/admin.py`, `app/api/v1/endpoints/auth.py`, `app/services/admin_service.py` | Login failures, destructive reset не пишутся в audit trail | Централизованный security logger (user_id, IP, action, outcome) |
| **S12** | Риск CSRF при первом login без prefetch cookie | `frontend/src/features/auth/model/AuthProvider.tsx`, `frontend/src/shared/api/httpClient.ts`, `app/middleware/csrf.py` | Race/prefetch failure → 403 на login или непредсказуемое поведение | Явный `GET /auth/csrf` при mount LoginPage |

### Code Quality

| ID | Проблема | Location | Impact | Fix |
|----|----------|----------|--------|-----|
| **Q5** | DRY: параллельные procurement breakdown | `viz_modules/procurement/breakdown.py:82`, `:393` | ~70% логики дублируется | Общий pipeline с strategy/flags (`mode: commercial \| production`) |
| **Q6** | Stub-метод с мёртвыми параметрами | `app/services/commercial_wizard_step_service.py:37` | Иллюзия динамики шага wizard | Реализовать логику или заменить константой |
| **Q7** | Длинные функции day-view без декомпозиции | `app/services/day_view_service.py:222,481,38` | Сложно локализовать баги агрегации | Extract: fetch → normalize → aggregate → build DTO |
| **Q8** | Silent failure и print вместо logging в KP DB | `core/kp_db_offers.py:53,481-484,557-558` | Ошибки без причины; print теряется в prod | `logger.exception` + typed errors |
| **Q9** | Тихий fallback при невалидной дате распределения | `app/planning/plan_distribution.py:56-59` | Некорректная дата silently → `datetime.now()` | Пробрасывать `ValueError`/`PlanBuildError` |
| **Q10** | Слабая типизация production API на frontend | `frontend/src/features/production/types/production.ts:167-181` | Поля плана не проверяются компилятором | Явные interfaces под OpenAPI; zod-parse на границе API |
| **Q11** | Строковые коды ошибок вместо typed exceptions | `app/services/offers_service.py:103-114` | Endpoint матчит строки; нет exhaustiveness | `OfferNotFoundError`, `OfferInvalidStatusError` |
| **Q12** | Boilerplate-делегирование в workflow facade | `app/services/commercial_workflow_service.py:46-65` | ~30% файла — шум passthrough | Consumers inject `CommercialWizardStepService` напрямую |
| **Q13** | Пробел unit-тестов OffersService | `app/services/offers_service.py` | create/update/generate без regression safety | `tests/test_offers_service.py` с бизнес-кейсами |

---

## Low Priority / Suggestions

### Architecture

| ID | Проблема | Location | Impact | Fix |
|----|----------|----------|--------|-----|
| **A20** | `plan_manager.py` sys.path hack | `app/planning/plan_manager.py:16-18` | Хрупкость при packaging/deploy | Package-relative imports; убрать `sys.path.insert` |
| **A21** | Repository зависит от app-layer planning utils | `app/repositories/plan_repository.py` → `app/planning/plan_storage.count_day_tracks` | Инверсия dependency direction | Перенести `count_day_tracks` в repository или `core/production` |
| **A22** | SQLite как единственный persistence backend | `core/kp_db*.py`, `app/repositories/*`, `app/core/settings.py` | Single-writer bottleneck; ограничение horizontal scaling | Defer до P6; PostgreSQL + migration layer |
| **A23** | Visualization ports wiring только at startup | `app/main.py:32-34`, `app/adapters/visualization.py` | Без `wire_visualization_ports()` → runtime failures | Lazy default registration или fail-fast import check |

### Security

| ID | Проблема | Location | Impact | Fix |
|----|----------|----------|--------|-----|
| **S13** | Различимые сообщения при смене пароля | `app/api/v1/endpoints/auth.py` | Minor information disclosure | Унифицировать ответ |
| **S14** | Password policy errors на английском | `app/security/password_policy.py`, `app/schemas/auth.py` | Утечка требований к паролю | Generic client message + детали в server log |
| **S15** | Длинная сессия без idle timeout | `app/security/session.py`, `core/config/settings.py` | Украденная cookie действует до `exp` | Sliding idle timeout; сократить TTL |
| **S16** | PBKDF2-SHA256 вместо Argon2id | `app/repositories/auth_repository.py` | Слабее против GPU offline attack | Миграция на Argon2id при login |
| **S17** | `GET /api/v1/managers` доступен роли `production` | `app/api/v1/endpoints/managers.py` | Business metadata disclosure | `require_roles("admin", "manager")` |
| **S18** | Небounded query param на production candidates | `app/api/v1/endpoints/production.py` (`limit: int = 500`) | Minor DoS через большой limit | `Query(default=500, ge=1, le=1000)` |
| **S19** | Traceback в stdout при OCR-ошибках | `core/ocr_gpt.py` | Information leakage в shared logging | `logger.exception()` без print |
| **S20** | Wizard state fallback для unknown role → `/new` | `frontend/src/shared/lib/roleRoutes.ts` | UX раскрывает intent при некорректной роли | Fail-safe redirect на `/login` |

### Code Quality

| ID | Проблема | Location | Impact | Fix |
|----|----------|----------|--------|-----|
| **Q14** | Dead dependency factory | `app/dependencies/services.py:49` (`get_commercial_wizard_step_service`) | Мёртвый код; риск отдельного DraftStore | Удалить или подключить к router |
| **Q15** | Неиспользуемый import | `core/kp_db_offers.py:8` (`import traceback`) | Шум static analysis | Удалить top-level import |
| **Q16** | Cross-stack DRY: дублирование формулы доставки | `frontend/.../cargoDeliveryPricing.ts`, `core/cargo_delivery_pricing.py` | Расхождение UI vs PDF/KP | Backend — source of truth |
| **Q17** | Массовый `user: dict` / `dict[str, Any]` в service layer | `offers_service.py`, `archive_service.py`, `commercial_workflow_service.py` | Опечатки ключей не ловятся mypy | TypedDict / Pydantic models |
| **Q18** | Неполное покрытие day_documents pipeline | `app/services/day_documents_service.py` | ZIP/PDF/XLSX без dedicated tests | Integration tests с temp dirs + mocks |

---

## Priority Matrix

| ID | Issue | Severity | Effort | Priority |
|----|-------|----------|--------|----------|
| A1 | Sync CPU-bound блокирует HTTP workers | Critical | Medium | **P0 — немедленно** |
| A2 | Обход visualization ports | Critical | Medium | **P0 — немедленно** |
| A3 | Thread-local plate runtime | Critical | High | **P0 — немедленно** |
| S3 | Destructive DB reset при misconfig | High | Low | **P0 — немедленно** |
| S4 | Frontend без route-level RBAC | High | Low | **P0 — немедленно** |
| A4 | God module CommercialWorkflowService | High | High | **P1 — этот спринт** |
| A5 | God module kp_db_offers | High | High | **P1 — этот спринт** |
| A6 | God module core/visualization | High | High | **P1 — этот спринт** |
| A7 | Дублирование planning app vs core | High | High | **P1 — этот спринт** |
| A8 | DI без абстракций | High | Medium | **P1 — этот спринт** |
| A9 | Auth без application-слоя | High | Medium | **P1 — этот спринт** |
| A10 | In-process rate limiting | High | Medium | **P1 — этот спринт** |
| A11 | God modules production path | High | High | **P1 — этот спринт** |
| A12 | Legacy web routes | High | Medium | **P1 — этот спринт** |
| S1 | Rate limiting per-worker bypass | High | Medium | **P1 — этот спринт** |
| S2 | OCR → OpenAI PII | High | Medium | **P1 — этот спринт** |
| Q1 | Layout builder monolith | High | High | **P1 — этот спринт** |
| Q2 | Wizard validation DRY | High | Medium | **P1 — этот спринт** |
| Q3 | Execution terms DRY | High | Low | **P1 — этот спринт** |
| Q4 | Frontend hooks untested | High | Medium | **P1 — этот спринт** |
| A13–A19 | Medium architecture findings | Medium | Mixed | **P2 — следующий спринт** |
| S5–S12 | Medium security findings | Medium | Mixed | **P2 — следующий спринт** |
| Q5–Q13 | Medium code quality findings | Medium | Mixed | **P2 — следующий спринт** |
| A20–A23 | Low architecture suggestions | Low | Low | **Backlog** |
| S13–S20 | Low security suggestions | Low | Low | **Backlog** |
| Q14–Q18 | Low code quality suggestions | Low | Low | **Backlog** |

---

## Next Steps

### 1. Immediate (до следующего релиза)

- **[A1]:** Вынести sync optimization/plan build из HTTP event loop — `asyncio.to_thread` или job API
- **[A2]:** Маршрутизировать `visualize_plan` через `core/ports/visualization` во всех app-сервисах
- **[A3]:** Рефактор plate runtime на явный `PlateOrderContext`; deprecate `get_plate_mutable_runtime()` на web paths
- **[S3]:** Fail-closed guard для destructive DB operations + startup-check production
- **[S4]:** Внедрить `RequireRole` route guards на frontend

### 2. This sprint (текущий спринт)

- **[A4]–[A6], [A11]:** Инкрементальная декомпозиция god-modules (commercial workflow, kp_db_offers, visualization, production)
- **[A7], [A8], [A9], [A12]:** Единое planning tree, DI abstractions, AuthService, legacy routes deprecation
- **[A10], [S1]:** Redis/shared rate limiter или deploy contract `workers=1`
- **[S2]:** Политика OCR / data residency (on-prem или opt-in)
- **[Q1]–[Q4]:** Layout builder split, wizard validation unify, execution terms helper, frontend hook tests

### 3. Next sprint (следующий спринт)

- **[A13]–[A19]:** PEP562 decommission, DraftStore singleton, response_model, legacy_runtime removal
- **[S5]–[S12]:** CSRF hardening, sessionStorage minimization, CSP enforce, audit log, legacy auth removal
- **[Q5]–[Q13]:** Procurement DRY, day-view refactor, OffersService tests, typed errors

### 4. Backlog

- **[A20]–[A23], [S13]–[S20], [Q14]–[Q18]:** sys.path cleanup, PostgreSQL migration, Argon2id/MFA, magic numbers, dead code

---

## Remediation Commands

| Тип проблемы | Команда | Примеры |
|--------------|---------|---------|
| Структурные / архитектурные / DRY / god-modules | `/refactor` | `/refactor app/services/file_generation_service.py`, `/refactor core/plate_runtime_state.py` |
| Security / behavioral / новая логика | `/implement` | `/implement asyncio.to_thread для commercial optimize`, `/implement RequireRole на frontend` |

**Рекомендуемый порядок remediation:**

1. `/implement asyncio.to_thread / job API для CPU-bound endpoints` — P0, блокирует production latency
2. `/refactor app/services/file_generation_service.py` — P0, visualization ports
3. `/implement PlateOrderContext на web paths` — P0, изоляция заказов
4. `/implement RequireRole route guards` — P0 security UX
5. `/refactor core/visualization/` — P1, инкрементальная декомпозиция

---

## Связанная документация

- Full-repo audit (включая bot): [`2026-06-21-full-project-audit.md`](./2026-06-21-full-project-audit.md)
- P5 stabilization spec (closed): [`../../specs/stabilizaciya-p5-architecture-2026-06-21.md`](../../specs/stabilizaciya-p5-architecture-2026-06-21.md)
- **P6 stabilization spec (closed):** [`../../specs/stabilizaciya-p6-architecture-2026-06-21.md`](../../specs/stabilizaciya-p6-architecture-2026-06-21.md) — WP1–WP4, WP7–WP8 closed; WP9 stretch deferred
- ADR core↔viz boundary: [`../architecture/core-viz-modules-boundary.md`](../architecture/core-viz-modules-boundary.md)

---

*Отчёт сгенерирован workflow `/audit` (frontend+backend lens). Health Score 0.0/10 отражает 3 critical + 17 high findings при формуле с caps. Telegram-бот исключён из scope — bot Critical (A1/A2 full-repo) сняты с active web path.*
