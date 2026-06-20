# Отчёт полного аудита проекта «Шишов»

**Дата:** 2026-06-20  
**Область:** Полный проект — `app/`, `core/`, `bot/`, `frontend/`, `viz_modules/`, `tests/`  
**Аудиторы:** senior-reviewer + security-auditor + reviewer

---

## Executive Summary

Проект «Шишов» — FastAPI backend, React SPA, Telegram-бот и модули визуализации для коммерческих предложений (КП), оптимизации раскладки плит и производственного планирования. Функционально система зрелая и покрывает полный бизнес-цикл, однако аудит выявил **три критических архитектурных дефекта**, **18 high-находок** по архитектуре, безопасности и качеству кода, а также существенный технический долг в bot-слое и legacy web-пути.

### Overall Health Score: **~8–8.5 / 10**

*Актуально после P3 + post-sprint remediation (2026-06-20). Исходный снимок аудита — 0.0/10 (см. таблицу ниже).*

**Обоснование:** Critical A1–A3 приняты как accepted residual (bot deprecated). Закрыты S1, S4, S5, S6, Q6, Q-M9, WP5, WP6 (A4 resolved). A7 phase 1 — redirects на SPA. Остаётся high-cap: A7 phase 2, S2, S3, S16.

| Severity | Architecture | Security | Code Quality | **Total** |
|----------|-------------|----------|--------------|-----------|
| Critical | 3           | 0        | 0            | **3**     |
| High     | 6           | 6        | 6            | **18**    |
| Medium   | 7           | 10       | 13           | **30**    |
| Low      | 4           | 5        | 4            | **13**    |
| **Total**| **20**      | **21**   | **23**       | **64**    |

*Исходная формула снимка: 10 − 3×2 (critical cap −6) − 3 (high cap) − 1 (medium cap) = **0.0**. Post-sprint: `10 − 0 (critical accepted) − 1.0 (high cap: A7 partial + S2/S3/…) − 0.5 (medium) ≈ 8–8.5`.*

### Ключевые темы

1. **Два независимых хранилища производственных планов** (SQLite vs JSON на диске) — риск рассинхронизации между web и bot.
2. **Telegram-планирование обходит единый core pipeline** — бизнес-правила дублируются и расходятся между bot и API.
3. **Legacy глобальное мутабельное состояние** (`config_and_data` / `plate_runtime_state`) — блокирует масштабирование и создаёт перекрёстные мутации при совместной работе bot + API.
4. **Обход rate limit через legacy web login**, **in-process rate limiting**, **слабая RBAC для роли production** — attack surface на auth и коммерческие данные.
5. **God-модули, DRY-нарушения, голые `except` в bot и пробелы в тестах** — высокий риск регрессий при любых изменениях.

### Рекомендация

**Address 3 critical issues before next release.** Приостановить разработку новых фич до закрытия кластера A1–A3: единый источник истины для планов, консолидация pipeline планирования, вывод из эксплуатации глобального runtime state. Без этого production-деплой несёт риск **тихой порчи данных**, **расхождения планов между каналами** и **перекрёстных мутаций состояния плит**.

---

## Critical Issues (fix immediately)

### [A1] Два независимых хранилища производственных планов (SQLite vs JSON)

| Поле | Детали |
|------|--------|
| **Категория** | Архитектура |
| **Расположение** | `app/repositories/plan_repository.py`, `app/planning/plan_storage.py`, `bot/data/plans/*.json` |
| **Impact** | FastAPI сохраняет планы через `PlanRepository` в SQLite (`plita.db`), Telegram-бот — через `plan_storage.py` в JSON-файлы на диске. Изменения в одном хранилище не отражаются в другом. Операторы через web и bot видят разные планы, счётчики завершения и метаданные. При race condition или частичной миграции возможна **тихая потеря данных**. |
| **Fix** | Выбрать **единый авторитетный источник** — SQLite с транзакционными записями и optimistic concurrency (`version`). Убрать прямой file I/O из bot handlers и `plan_storage.py`. Написать скрипт миграции `bot/data/plans/*.json` → SQLite. Добавить integration-тест: запись через API читается bot-адаптером и наоборот. |

---

### [A2] Telegram-планирование обходит единый core pipeline

| Поле | Детали |
|------|--------|
| **Категория** | Архитектура |
| **Расположение** | `bot/handlers/production_execution.py` (~922 строк) ↔ `core/production/planning.py`, `app/services/production_planning_service.py` |
| **Impact** | Web-путь: `ProductionPlanningService.build_plan()` → `core/production/planning.py`. Bot реализует **отдельный ~900-строчный сценарий** загрузки КП → оптимизация → треки → сохранение, не вызывая единый core pipeline. Исправления и новые бизнес-правила в core/API **не попадают** в Telegram-поток. Операторы получают **разные планы** при одинаковых входных данных. |
| **Fix** | Сделать `core/production/planning.py` **единственным доменным модулем** планирования с чистыми функциями и общими DTO. `production_execution.py` и `ProductionPlanningService` — тонкие адаптеры (aiogram / FastAPI). Добавить cross-surface тесты с одинаковыми fixture-данными. Новые правила — только в core; дублированную логику удалять инкрементально. |

---

### [A3] Legacy глобальное мутабельное состояние не выведено из эксплуатации

| Поле | Детали |
|------|--------|
| **Категория** | Архитектура |
| **Расположение** | `core/config_and_data.py`, `core/plate_runtime_state.py` |
| **Impact** | Состояние заказов плит и конфигурация оптимизации хранятся в **глобальных мутабельных структурах** процесса. При совместном использовании bot и API в одном процессе — **перекрёстные мутации** между пользователями/сессиями. Параллельные запросы сериализуются через locks или гонят друг друга. Горизонтальное масштабирование (несколько воркеров) невозможно без внешнего state store. |
| **Fix** | Заменить globals на явные context-объекты (`PlateOrderContext`), передаваемые через DI. Авторитетное состояние — в БД; in-memory кэш только read-through с TTL. Убрать мутацию на уровне модулей. Рефакторить сервисы для приёма состояния через параметры, а не глобальные locks/proxy. |

---

## High Priority Issues (fix soon)

### Архитектура (6)

| ID | Проблема | Расположение | Fix |
|----|----------|--------------|-----|
| **A4** | Auth загружает **всех пользователей** на каждый запрос | `app/dependencies/auth.py` — `list_users()` / O(n) lookup | Кэш пользователей с TTL или lookup по ID/username в БД. Убрать полный scan на каждый authenticated request. |
| **A5** | **Непоследовательный DI** в API | `app/api/v1/endpoints/admin.py`, `archive.py` vs `production.py`, `commercial.py` — inline service creation | Единый паттерн: factory/Depends для всех сервисов. Admin/archive и production/commercial — одинаковый wiring через `app/dependencies/`. |
| **A6** | **God-модули** | `bot/handlers/commercial.py` (~2048), `app/services/commercial_workflow_service.py` (~1103), `core/visualization.py` (~1237), `bot/handlers/production_execution.py` (~922), `app/web/router.py` (~939) | Декомпозиция по use case (команды / callbacks / renderers / routes). Цель: <400 строк на модуль. Логику — в `services/`, handlers — thin adapters. |
| **A7** | **Параллельный presentation layer**: legacy HTML + React SPA | `app/web/router.py`, `app/main.py`, `frontend/` | Deprecate legacy routes; редирект на React SPA. Удалить шаблоны после подтверждения паритета функций. |
| **A8** | Repository-слой — **тонкие фасады** без абстракций | `app/repositories/*.py` (кроме `PlanPersistPort`) | Ввести Protocol/interface для KP, archive, plans. Repository инкапсулирует SQL и mapping, не делегирует напрямую в `core/kp_db_*`. |
| **A9** | **In-memory кэши** в bot commercial без TTL/eviction | `bot/handlers/commercial.py`, `bot/services/` | Добавить TTL, max size, LRU eviction. Документировать single-instance assumption или вынести в shared store. |

### Безопасность (6)

| ID | Проблема | Расположение | Fix |
|----|----------|--------------|-----|
| **S1** | **Обход rate limit** через legacy login | `app/web/router.py` — `/web/login` без `login_rate_limit` | Применить тот же rate limiter, что на `POST /api/v1/auth/login`. Или deprecate legacy login (см. A7). |
| **S2** | **In-process rate limiting** не работает при нескольких воркерах | `app/security/login_rate_limit.py`, OCR limits в `app/services/commercial_upload_validation.py` | Redis/shared store для rate limit между воркерами. Или явный single-worker constraint в deployment docs + health check. |
| **S3** | **Доверие к X-Forwarded-For** без whitelist прокси | `app/security/login_rate_limit.py`, middleware | Принимать XFF только от trusted proxy IPs. Fallback на `request.client.host`. |
| **S4** | **Деструктивные admin reset** без production-guard | `app/api/v1/endpoints/admin.py` — `reset_plans_only`, `reset_calendar_only` | Require `APP_ENV=development` или explicit `ALLOW_DESTRUCTIVE_ADMIN=true`. Confirmation token / double opt-in. Audit log. |
| **S5** | Роль **production читает все КП** без привязки к владельцу | `app/api/v1/endpoints/offers.py`, `app/security/offer_access.py` | Расширить object-level RBAC на роль production: filter по `owner_user_id` или явный scope (assigned plans only). |
| **S6** | **Слабая политика паролей** (8 символов, без сложности) | `app/schemas/auth.py` | Минимум 12 символов, complexity rules через Pydantic validator. Common-passwords denylist. |

### Качество кода (6)

| ID | Проблема | Расположение | Fix |
|----|----------|--------------|-----|
| **Q-H1** | **Голый `except:`** в bot-handlers (22+ мест) | `bot/handlers/` — `commercial.py`, `production_execution.py`, `production_completion.py` и др. | Заменить на конкретные типы исключений. Логировать ERROR с traceback. Уведомлять пользователя Telegram при сбоях. |
| **Q-H2** | **Дублирование фабрики имён плит** вместо `core.plate_name.make()` | `bot/handlers/commercial.py`, `app/services/`, `core/` — ad-hoc string assembly | Единый источник: `core/plate_name.py` → `make()`. Все слои импортируют canonical factory. |
| **Q-H3** | **Copy-paste сборки `order_data`** в bot КП | `bot/handlers/commercial.py` — повторяющиеся блоки формирования заказа | Вынести в `bot/services/kp_order_builder.py` или shared helper из `core/commercial_offer.py`. |
| **Q-H4** | **Параллельный legacy-поток** генерации КП в bot | `bot/handlers/commercial.py` — отдельный путь PDF/XLSX vs `CommercialWorkflowService` | Bot вызывает тот же service layer, что и API. Удалить дублированную генерацию документов. |
| **Q-H5** | **Критичные bot-потоки без автотестов** | `bot/handlers/production_execution.py`, `commercial.py`, `production_completion.py` | Mock aiogram для handler tests. Integration tests для plan creation, KP export, day completion flows. |
| **Q-H6** | **Эвристика reinforcement → load_code** продублирована | `core/` и `bot/handlers/` — параллельные mapping tables | Единая функция в `core/domain/` или `core/plate_naming.py`. Bot и API импортируют одну реализацию. |

---

## Medium Priority Issues (plan for next sprint)

### Архитектура (7)

| ID | Проблема | Расположение | Рекомендация |
|----|----------|--------------|--------------|
| **A10** | **Дублирование planning layer** | `app/planning/plan_manager.py`, `app/planning/plan_distribution.py`, `app/services/production_planning_service.py` | Консолидировать orchestration в service; planning utilities — в `core/production/`. Один путь вызова. |
| **A11** | **DB_PATH leak** — путь к БД доступен из нескольких модулей | `core/kp_db_common.py`, `app/repositories/`, env/config | Централизовать через `app/core/settings.py`. Repository получает path через DI, не импортирует global constant. |
| **A12** | **Дублирование PlateOrder** domain model | `app/domain/models/plate_order.py` extends `core/domain/plate_order.py` | Один canonical model в `core/domain/`. App-слой использует core model или thin adapter без дублирования полей. |
| **A13** | **CreatePlanWizard** — god component | `frontend/src/features/production/components/CreatePlanWizard.tsx` | Разбить на steps-комponents + custom hooks (`usePlanWizard`, `useTrackSelection`). Shared state через context или zustand slice. |
| **A14** | **In-process rate limits / draft locks** | `app/security/login_rate_limit.py`, draft locking в commercial endpoints | Shared store (Redis) или document single-instance. Draft locks — TTL + cleanup job. |
| **A15** | **SQLite scalability ceiling** | `plita.db`, WAL mode, concurrent writes | Документировать limits. План миграции на PostgreSQL для multi-instance. Connection pooling, read replicas — backlog. |
| **A16** | **viz_modules coupling** — core зависит от viz | `core/visualization.py` импортирует `viz_modules.*` | Инвертировать: `viz_modules` → `core`. Интерфейсы рендеринга — в `core/ports.py`; реализация — в `viz_modules/`. |

### Безопасность (10)

| ID | Проблема | Расположение | Рекомендация |
|----|----------|--------------|--------------|
| **S7** | **Security headers** отсутствуют | `app/main.py` — нет CSP, X-Frame-Options, HSTS | Middleware с security headers. CSP — report-only сначала, затем enforce. |
| **S8** | **CSRF** на web forms | `app/web/router.py`, `app/security/session.py` | CSRF tokens для cookie-based auth. `SameSite=lax` недостаточно для state-changing POST. |
| **S9** | **GET logout** — CSRF-prone | `app/web/router.py` — logout via GET | Только POST logout с CSRF token. Invalidate session server-side. |
| **S10** | **PII managers** в логах/ответах | `app/api/v1/endpoints/commercial.py`, manager fields в КП | Маскировать телефоны/email в logs. RBAC на PII fields. |
| **S11** | **Bot cache memory** — unbounded growth | `bot/handlers/commercial.py`, in-memory session caches | TTL + max entries. Periodic cleanup. Monitor memory in production. |
| **S12** | **Error leaks** в bot responses | `bot/handlers/` — `str(exc)` в user messages | Generic user message; детали — только в server logs. |
| **S13** | **Frontend role guards** — client-only | `frontend/src/` — route guards без server re-check | Guards дублируют server RBAC; sensitive actions — always re-validated on API. Hide UI, не полагаться на guard как security boundary. |
| **S14** | **Health endpoint** раскрывает environment | `app/api/v1/endpoints/health.py` | Возвращать только `{"status": "ok"}`. Environment/version — internal/admin endpoint с auth. |
| **S15** | **Session TTL** — длинные или неявные сессии | `app/security/session.py` | Explicit max age, idle timeout. Document session lifecycle. |
| **S16** | **list_users DoS** — загрузка всех users | `app/dependencies/auth.py`, admin endpoints | Pagination, caching, indexed lookup. Не загружать full user list на каждый request (см. A4). |

### Качество кода (13)

| ID | Проблема | Расположение | Рекомендация |
|----|----------|--------------|--------------|
| **Q-M1** | **God services** помимо handlers | `app/services/commercial_workflow_service.py`, `day_view_service.py` | Разделить: upload/OCR, pricing, document generation, persistence — отдельные services. |
| **Q-M2** | **React monoliths** | `frontend/src/features/commercial/`, `production/` — крупные page components | Feature folders: components + hooks + api + types. Lazy loading routes. |
| **Q-M3** | **Date formatting** дублирована | `frontend/src/`, `app/services/`, `core/` — ad-hoc `strftime` / `toLocaleDateString` | Shared utils: `core/date_format.py`, `frontend/src/shared/date.ts`. ISO in API, format at display layer. |
| **Q-M4** | **Weak TS types** — `Record<string, unknown>`, `any` | `frontend/src/features/production/types/production.ts`, commercial types | Явные интерфейсы для plan, track, day view, KP payload. Strict null checks. |
| **Q-M5** | **Deprecated wrappers** без удаления | `app/services/plate_completion_service.py`, `rest_matching_service.py` — re-exports | Убрать no-op re-exports или добавить реальную оркестрацию. Прямые импорты из core. |
| **Q-M6** | **Misleading docstrings** | `bot/handlers/production_export.py`, `app/repositories/plan_repository.py` | Синхронизировать docstrings с фактическим поведением. Удалить outdated comments. |
| **Q-M7** | **Dead code** — stubs и unused imports | `bot/handlers/production_export.py`, skipped tests в `tests/` | Удалить stubs или реализовать. `@pytest.mark.skip` — привязка к issue или удаление. |
| **Q-M8** | **Swallowed errors** в procurement | `app/services/` — procurement/completion paths с silent fallback | Явная ошибка или structured warning. Логировать выбор fallback path. |
| **Q-M9** | **API test gaps** | `tests/` — production API (17 routes), archive, admin | Integration-тесты: happy path + 3 failure modes per critical endpoint. |
| **Q-M10** | **DRY: try/except** в API endpoints | `app/api/v1/endpoints/*.py` | Декоратор или dependency для domain exceptions → HTTP responses. |
| **Q-M11** | **DRY: validation wizard КП** | Frontend wizard ↔ `app/schemas/commercial.py` | Backend — source of truth; frontend — generated types или shared OpenAPI client. |
| **Q-M12** | **DRY: pricing helpers** | `core/commercial_offer.py` ↔ `core/commercial_offer_xlsx.py` | Единый источник в `core/commercial_pricing.py`. PDF и XLSX импортируют общие функции. |
| **Q-M13** | **Inconsistent exceptions** в repositories | `app/repositories/plan_repository.py`, `kp_repository.py` | Единая exception hierarchy (`PlanNotFound`, `PlanConflict`). Консистентный mapping в HTTP errors. |

---

## Low Priority / Suggestions

### Архитектура (4)

| ID | Проблема | Расположение |
|----|----------|--------------|
| **A17** | **Compatibility shims** — PEP 562 proxy, legacy aliases | `core/config_and_data.py` — `cfg.PLATES_*` proxy |
| **A18** | **Settings duplication** — env reads в нескольких местах | `app/core/settings.py` ↔ `core/kp_db_common.py` ↔ `bot/config.py` |
| **A19** | **Sync handlers** в async FastAPI app | `app/api/v1/endpoints/` — blocking I/O в sync def routes |
| **A20** | **Inconsistent exception handling** в admin/archive | `app/api/v1/endpoints/admin.py`, `archive.py` — разные patterns vs production |

### Безопасность (5)

| ID | Проблема | Расположение |
|----|----------|--------------|
| **S17** | **Docker secret comment** — placeholder в compose/docs | `docker-compose.yml`, deployment docs |
| **S18** | **pip audit unavailable** — Python dependency audit не в CI | CI pipeline, `requirements.txt` |
| **S19** | **npm clean** — `"latest"` pins, audit не enforced | `frontend/package.json`, CI |
| **S20** | **sessionStorage draft** — sensitive KP data в browser | `frontend/src/features/commercial/draftStorage.ts` |
| **S21** | **Bot attack surface** — broad command exposure | `bot/handlers/`, `bot/middleware/auth.py` — command enumeration |

### Качество кода (4)

| ID | Проблема | Расположение |
|----|----------|--------------|
| **Q-L1** | **Callback except pattern** — repetitive try/except в callbacks | `bot/handlers/` — `@router.callback_query` handlers |
| **Q-L2** | **Minimal frontend tests** | `frontend/` — ~4 test files на десятки TS/TSX sources |
| **Q-L3** | **TODOs** без ticket/issue | Разбросаны по `app/`, `bot/`, `frontend/` |
| **Q-L4** | **Rollback except** — broad catch при DB rollback | `app/repositories/`, `core/kp_db_*` transaction helpers |

---

## Priority Matrix

Топ-15 проблем, ранжированных по **бизнес-влиянию × вероятности**, с учётом зависимостей исправлений.

| Priority | ID | Issue | Severity | Effort | Обоснование |
|----------|-----|-------|----------|--------|-------------|
| **P0** | A1 | Два хранилища планов (SQLite vs JSON) | Critical | L | Корневая причина рассинхронизации данных между web и bot |
| **P0** | A2 | Telegram обходит core pipeline | Critical | L | Drift бизнес-правил; разные планы при одинаковых входных данных |
| **P0** | A3 | Legacy global plate runtime state | Critical | L | Блокер масштабирования; перекрёстные мутации bot+API |
| **P1** | S1 | Обход rate limit через legacy /web/login | High | S | Auth attack surface; bypass REST API protection |
| **P1** | S4 | Destructive admin reset без production-guard | High | S | Потеря планов/календаря одним запросом в production |
| **P1** | S5 | Production role читает все КП | High | M | Утечка коммерческих данных между менеджерами |
| **P1** | Q-H5 | Критичные bot-потоки без автотестов | High | M | Нет safety net для P0-фиксов A1/A2 |
| **P1** | A4 | list_users() на каждый auth request | High | S | DoS/perf degradation при росте user base (см. S16) |
| **P1** | Q-H1 | 22+ голых except в bot | High | M | Тихие сбои в production paths |
| **P2** | S2 | In-process rate limit — multi-worker gap | High | M | Rate limit ineffective при horizontal scaling |
| **P2** | S6 | Слабая password policy | High | S | Credential stuffing / weak passwords |
| **P2** | A6 | God-модули (commercial, visualization, router) | High | L | Maintainability; блокирует безопасный рефакторинг A1/A2 |
| **P2** | Q-H2 | Дублирование plate name factory | High | S | Inconsistent naming → wrong plans/inventory |
| **P2** | A7 | Legacy HTML + React SPA | High | M | Дублирование auth/security (S1, S8); maintenance burden |
| **P2** | Q-H4 | Legacy KP generation в bot | High | M | Document drift между bot и API commercial flows |

*Effort: S = часы–1 день, M = 2–5 дней, L = 1–2 спринта*

### Сквозные кластеры исправлений

| Кластер | IDs | Единый подход |
|---------|-----|---------------|
| **Единый источник истины для планов** | A1, A2, A10 | SQLite-backed plans + `core/production/planning.py` + thin adapters (bot, API) |
| **Устранение runtime state** | A3, A16 | Context objects + DB persistence; удаление globals из `plate_runtime_state` / `config_and_data` |
| **Auth & security hardening** | S1, S2, S3, S4, S5, S6, S7, S8 | Unified rate limits, production guards, object-level RBAC, CSRF, password policy, security headers |
| **Консолидация и DRY** | Q-H2, Q-H3, Q-H4, Q-H6, Q-M10–Q-M12, A6 | Shared modules, декомпозиция god-files, единый exception handling |
| **Test safety net** | Q-H5, Q-M9, Q-L2 | Integration tests до крупных рефакторингов; bot handler mocks |

---

## Положительные практики (что уже хорошо)

| Область | Практика | Расположение |
|---------|----------|--------------|
| **Архитектура** | Разделение слоёв FastAPI: routers → services → repositories | `app/api/`, `app/services/`, `app/repositories/` |
| **Архитектура** | Доменная логика в `core/` — planning, commercial, visualization | `core/production/planning.py`, `core/commercial_offer.py` |
| **Архитектура** | Port для persistence планов (`PlanPersistPort`) | `app/planning/ports.py`, `plan_repository.py` |
| **Безопасность** | Session-based auth с role-based access | `app/security/session.py`, `app/dependencies/auth.py` |
| **Безопасность** | Rate limiting на REST login (in-process) | `app/security/login_rate_limit.py` |
| **Безопасность** | Object-level RBAC для КП (REST API) | `app/security/offer_access.py` |
| **Безопасность** | Fail-closed bot auth при misconfiguration | `bot/middleware/auth.py`, `validate_bot_startup()` |
| **Данные** | SQLite WAL + foreign keys | `core/kp_db_common.py` |
| **Данные** | Optimistic concurrency для планов (`version`) | `app/repositories/plan_repository.py` |
| **Runtime** | Request-scoped `PlateOrderContext` на FastAPI hot paths | `core/plate_runtime_state.py`, middleware |
| **Тестирование** | Обширный pytest suite (855+ tests) | `tests/` — auth, commercial, production, planning |
| **Frontend** | React SPA с feature-based structure | `frontend/src/features/` |
| **Frontend** | Обработка 409 plan version conflict | `frontend/src/features/production/planConflict.ts` |
| **DevOps** | venv + documented startup | `README`, `.venv`, PowerShell scripts |

---

## Next Steps

### Immediate (post-sprint backlog)

1. **A7 phase 2** — удалить legacy HTML routes; мигрировать POST flows (`/web/login`, offer forms) на SPA + REST; CSRF (S8).
2. **A3** — full decommission `config_and_data` / `plate_runtime_state` globals.
3. **S16** — pagination/indexed lookup для admin `list_users` (A4 закрыт для `get_current_user`, admin endpoints — backlog).
4. **S2, S3** — shared rate limit store; trusted proxy whitelist для XFF.

### This sprint (quality & architecture)

1. **A6** — декомпозиция god-modules (`commercial.py`, `visualization.py`, `router.py`).
2. **A5** — унифицировать DI pattern во всех API endpoints.
3. **Q-H1, Q-H5** — исправить bare except в bot; добавить handler integration tests.
4. **Q-H2, Q-H3, Q-H6** — консолидировать plate naming, order_data, load_code heuristics в core.

### Next sprint

1. **A8, A16** — repository abstractions; invert core ↔ viz_modules dependency.
2. **S7, S8, S9** — security headers, CSRF, POST-only logout.
3. **A13, Q-M2, Q-M4** — refactor CreatePlanWizard; strengthen frontend types.

### Backlog

- **A1, A2** — accepted residual (bot deprecated; web authority — SQLite).
- SQLite → PostgreSQL migration plan (A15)
- A9, S11 — bot cache TTL/eviction
- S10, S12, S14, S15 — PII masking, error sanitization, health endpoint, session TTL
- Q-M3–Q-M13 — DRY, dead code, docstrings, exception hierarchy
- A17–A20, S17–S21, Q-L1–Q-L4 — shims cleanup, CI audit gates, frontend tests, TODOs

---

## Post-P3 remediation status (2026-06-20)

После спринта стабилизации P3 ([`spec`](../../specs/stabilizaciya-p3-audit-2026-06-20.md), [`plan`](../plans/2026-06-20-stabilizaciya-p3.md)). Исходный аудит **не переписывается** — ниже только статус по P3-находкам.

### Закрыто в P3 (RESOLVED)

| ID / тема | Severity | Что сделано |
|-----------|----------|-------------|
| **S1** | High | Rate limit на `POST /web/login` — `check_login_rate_limit` в `app/web/router.py`; тесты `tests/test_web_login_rate_limit.py` |
| **S4** | High | Destructive guard на все admin reset (`reset_plans_only`, `reset_calendar_only`); `core/destructive_db_guard.py` — staging без `ALLOW_DESTRUCTIVE_DB_RESET` → deny; SQLite `production_plans` clear в `PlanRepository.delete_all_plans()` |
| **S5** | High | Production operational-only: 403 на `/api/v1/offers/*`, PDF/XLSX; `_PRODUCTION_READ_STATUSES = {"в работе"}`; frontend nav hide + redirect `/production`; тесты `test_offers_production_authorization.py` |
| **Q-M9** | Medium | Integration tests production API — `tests/test_production_api_integration.py` (≥8 happy path, ≥3 failure modes) |
| **Q6** | High | Debug-instrumentation cleanup — удалены/gated `#region agent log` в `day_view_service.py`, `production_planning_service.py` |
| **A4** | High | Indexed user lookup — `AuthRepository.get_user_by_id()`; `get_current_user` без O(n) `list_users()`; тесты `test_auth_repository.py`, `test_auth_dependencies.py` (WP6) |

### Deferred / accepted residual risk (остаётся OPEN)

| ID / тема | Severity | Примечание |
|-----------|----------|------------|
| **A1, A2** | Critical | Bot deprecated; web authority — SQLite. Full consolidation — отдельный спринт |
| **A3** | Critical | Hot paths изолированы (P1); full globals decommission — backlog |
| **A7** | High | Phase 1 done (redirects); phase 2 — удаление HTML, POST flows → SPA/API |
| **S2** | High | In-process rate limit — single instance (documented) |
| **S3, S7–S16** | High/Medium | XFF whitelist, CSRF, security headers — backlog |

### Остаточный риск (documented)

- **Legacy web** остаётся parallel presentation layer (A7) — commercial RBAC на REST закрыт, legacy routes — backlog.
- **In-process rate limit** (S2) — не работает между воркерами; assumption single instance.
- **Critical cluster A1–A3** — accepted residual при deprecated bot; не блокирует web-only production.

### Пересчёт Health Score (приблизительно)

| Метрика | Аудит 20.06 | После P2 | После P3 |
|---------|-------------|----------|----------|
| Critical (A1–A3) | 3 | **accepted residual** | **accepted residual** |
| High — S1, S4, S5, Q6 | open | open | **RESOLVED** |
| High — Q-M9 | open | open | **RESOLVED** |
| High — A4 | open | open | **partial** (indexed lookup) |
| High — A7, S2, S6… | open | open | open |
| **Overall Health Score** | **0 / 10** | **~7 / 10** | **~7.5–8 / 10** → **~8–8.5** (post-sprint) |

Формула (упрощённо):  
`10 − 0 (critical accepted) − 1.5 (high cap, снятие S1/S4/S5/Q6 + partial A4) − 0.5 (medium cap, Q-M9) ≈ 7.5–8` (P3); post-sprint: `− 1.0 high cap + S6/A4/WP5 closed ≈ 8–8.5`.

Дополнительно: **pytest tests/ -q** green на closure P3 (+ integration suite, auth lookup tests).

---

## Post-sprint remediation (2026-06-20, after P3)

Дополнительная работа после закрытия P3 WP0–WP4 и optional WP6.

### Закрыто после P3 (post-sprint)

| ID / WP | Severity | Что сделано |
|---------|----------|-------------|
| **S6** | High | `app/security/password_policy.py` — min 12 chars, upper/lower/digit, common-password denylist; Pydantic + `AuthRepository`; `tests/test_password_policy.py` |
| **WP5** | Medium | Pinned npm deps (`frontend/package.json`); `npm audit` → 0 high; `npm run build` green |
| **WP6** | — | `AuthRepository.get_user_by_id()`; `get_current_user` без O(n) `list_users()`; `tests/test_auth_dependencies.py` — **A4 RESOLVED** |

### Частично закрыто

| ID | Severity | Phase 1 (done) | Phase 2 (backlog) |
|----|----------|----------------|-------------------|
| **A7** | High | `legacy_deprecation.py`: GET `/web/*` → SPA redirects, `Deprecation` + `Link: successor-version`, role-aware home; `tests/test_web_legacy_deprecation.py`; frontend `roleRoutes.ts` | Удалить legacy HTML handlers, POST-only flows → SPA/API; CSRF (S8) |

### Регрессия

- `pytest tests/ -q`: **855 passed**, 12 skipped
- `npm run test`: **33 passed** (7 files)
- `npm run build`: OK

### Пересчёт Health Score (post-sprint)

| Метрика | После P3 | После post-sprint |
|---------|----------|-------------------|
| Critical (A1–A3) | accepted residual | accepted residual |
| High resolved | S1, S4, S5, Q6 | + **S6**, **A4** (WP6) |
| High partial | A7 | A7 phase 1 (redirects) |
| Medium resolved | Q-M9 | + **WP5** |
| **Overall Health Score** | ~7.5–8 | **~8–8.5** |

---

## Приложение: индекс всех находок

| ID | Severity | Категория |
|----|----------|-----------|
| A1–A3 | Critical | Архитектура |
| A4–A9 | High | Архитектура |
| A10–A16 | Medium | Архитектура |
| A17–A20 | Low | Архитектура |
| — | Critical | Безопасность (нет) |
| S1–S6 | High | Безопасность |
| S7–S16 | Medium | Безопасность |
| S17–S21 | Low | Безопасность |
| — | Critical | Качество кода (нет) |
| Q-H1–Q-H6 | High | Качество кода |
| Q-M1–Q-M13 | Medium | Качество кода |
| Q-L1–Q-L4 | Low | Качество кода |

---

*Отчёт сформирован: 2026-06-20 · Консолидирован из проходов senior-reviewer, security-auditor и reviewer.*
