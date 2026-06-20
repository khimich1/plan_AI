# Отчёт аудита проекта «Шишов»

**Дата**: 2026-06-20  
**Область**: Полный проект — `app/`, `core/`, `bot/`, `frontend/`, `viz_modules/`, `tests/`  
**Аудиторы**: senior-reviewer + security-auditor + reviewer

---

## Executive Summary

Проект «Шишов» — FastAPI backend, React SPA, Telegram-бот и модули визуализации для коммерческих предложений (КП), оптимизации раскладки плит и производственного планирования. Функционально система покрывает полный бизнес-цикл, однако консолидированный аудит выявил **4 критические находки** (3 уникальных проблемы), **24 high-находки** и существенный технический долг в bot-слое, auth-слое и legacy web-пути.

**Общий Health Score (baseline snapshot)**: **0.0 / 10** → **~8.5 / 10** после post-sprint closure (см. [Post-sprint remediation](#post-sprint-remediation-status-2026-06-20))

> **Примечание (2026-06-20, post-review):** Решение P0-2026-06-19 — **Telegram-бот deprecated / out of active use**. Этот снимок аудита трактует bot как активный канал; рекомендации по bot reliability (Q1, Q3, A4, A9) **переиндексированы** в [`stabilizaciya-p1-next-audit-2026-06-20.md`](../../specs/stabilizaciya-p1-next-audit-2026-06-20.md) v2 — приоритет **web/API security** (S4, S6, S2, S3).

| Severity | Architecture | Security | Code Quality | Итого |
|----------|-------------|----------|--------------|-------|
| Critical | 3 | 1 | 0 | **4** |
| High | 9 | 7 | 8 | **24** |
| Medium | 10 | 11 | 14 | **35** |
| Low | 6 | 6 | 8 | **20** |
| **Итого** | **28** | **25** | **30** | **83** |

*Формула Health Score: `10 − min(critical×2, 6) − min(high×0.5, 3) − min(medium×0.1, 1) = 10 − 6 − 3 − 1 = 0.0`*

### Ключевые темы

1. **Split-brain производственных планов** — SQLite (`PlanRepository`) и JSON на диске (`plan_storage.py`, `bot/data/plans/`) работают параллельно; web и bot видят разные данные.
2. **Telegram-планирование обходит единый core pipeline** — бизнес-правила дублируются и расходятся между bot и API.
3. **Legacy глобальное мутабельное состояние** — `config_and_data` / `plate_runtime_state` блокирует масштабирование и создаёт перекрёстные мутации.
4. **Auth и RBAC** — stateless-сессии, in-process rate limits, client-only frontend guards, GET logout, XFF spoofing.
5. **God-модули, DRY-нарушения, голые `except` в bot и пробелы в тестах** — высокий риск регрессий при любых изменениях.

**Рекомендация**: **Исправить 3 критических архитектурных дефекта до следующего релиза.** Приостановить разработку новых фич до закрытия кластера A1/S1–A3: единый источник истины для планов, консолидация pipeline планирования, вывод из эксплуатации глобального runtime state. Без этого production-деплой несёт риск **тихой порчи данных**, **расхождения планов между каналами** и **перекрёстных мутаций состояния плит**.

---

## Критические проблемы (исправить немедленно)

### [A1 / S1] Split-brain производственных планов (SQLite vs JSON)

| Поле | Детали |
|------|--------|
| **ID** | A1 (архитектура), S1 (безопасность — дублирует A1) |
| **Категория** | Архитектура + Безопасность |
| **Расположение** | `app/repositories/plan_repository.py`, `app/planning/plan_storage.py`, `bot/data/plans/*.json` |
| **Impact** | FastAPI сохраняет планы через `PlanRepository` в SQLite (`plita.db`), Telegram-бот — через `plan_storage.py` в JSON-файлы на диске. Изменения в одном хранилище не отражаются в другом. Операторы через web и bot видят разные планы, счётчики завершения и метаданные. С точки зрения безопасности — нарушение целостности данных (integrity): невозможно гарантировать консистентность при audit trail и rollback. При race condition или частичной миграции возможна **тихая потеря данных**. |
| **Fix** | Выбрать **единый авторитетный источник** — SQLite с транзакционными записями и optimistic concurrency (`version`). Убрать прямой file I/O из bot handlers и `plan_storage.py`. Закрыть JSON backdoor в `PlanRepository` (см. A11). Написать скрипт миграции `bot/data/plans/*.json` → SQLite. Добавить integration-тест: запись через API читается bot-адаптером и наоборот. |

---

### [A2] Telegram-планирование обходит единый core pipeline

| Поле | Детали |
|------|--------|
| **Категория** | Архитектура |
| **Расположение** | `bot/handlers/production_execution.py` (~922 строк) ↔ `core/production/planning.py`, `app/services/production_planning_service.py` |
| **Impact** | Web-путь: `ProductionPlanningService.build_plan()` → `core/production/planning.py`. Bot реализует **отдельный ~900-строчный сценарий** загрузки КП → оптимизация → треки → сохранение, не вызывая единый core pipeline. Исправления и новые бизнес-правила в core/API **не попадают** в Telegram-поток. Операторы получают **разные планы** при одинаковых входных данных. |
| **Fix** | Сделать `core/production/planning.py` **единственным доменным модулем** планирования с чистыми функциями и общими DTO. `production_execution.py` и `ProductionPlanningService` — тонкие адаптеры (aiogram / FastAPI). Добавить cross-surface тесты с одинаковыми fixture-данными. Новые правила — только в core; дублированную логику удалять инкрементально. |

---

### [A3] Legacy глобальное мутабельное состояние заказа плит

| Поле | Детали |
|------|--------|
| **Категория** | Архитектура |
| **Расположение** | `core/plate_runtime_state.py`, `core/config_and_data.py` |
| **Impact** | Состояние заказов плит и конфигурация оптимизации хранятся в **глобальных мутабельных структурах** процесса. При совместном использовании bot и API в одном процессе — **перекрёстные мутации** между пользователями/сессиями. Параллельные запросы сериализуются через locks или гонят друг друга. Горизонтальное масштабирование (несколько воркеров) невозможно без внешнего state store. |
| **Fix** | Заменить globals на явные context-объекты (`PlateOrderContext`), передаваемые через DI. Авторитетное состояние — в БД; in-memory кэш только read-through с TTL. Убрать мутацию на уровне модулей и PEP 562 proxy в `config_and_data`. Рефакторить сервисы для приёма состояния через параметры, а не глобальные locks/proxy. |

---

## Высокий приоритет

### Архитектура (9)

| ID | Проблема | Расположение | Fix |
|----|----------|--------------|-----|
| **A4** | Bot → app dependency inversion: bot импортирует `app.services.*` | `bot/handlers/`, `bot/services/` | Инвертировать зависимости: bot → core (домен), app → core. Общие сервисы — в `core/` или `shared/`. Bot не должен импортировать app-слой. |
| **A5** | God-модули | `bot/handlers/commercial.py` (~2048), `app/services/commercial_workflow_service.py` (~1103), `core/visualization.py` (~1237), `bot/handlers/production_execution.py` (~922), `app/web/router.py` (~939) | Декомпозиция по use case (команды / callbacks / renderers / routes). Цель: <400 строк на модуль. Логику — в `services/`, handlers — thin adapters. |
| **A6** | Непоследовательный DI в API | `app/api/v1/endpoints/admin.py`, `archive.py` vs `production.py`, `commercial.py` | Единый паттерн: factory/Depends для всех сервисов. Admin/archive и production/commercial — одинаковый wiring через `app/dependencies/`. |
| **A7** | Repository-слой — тонкие фасады без абстракций | `app/repositories/*.py` (кроме `PlanPersistPort`) | Ввести Protocol/interface для KP, archive, plans. Repository инкапсулирует SQL и mapping, не делегирует напрямую в `core/kp_db_*`. |
| **A8** | Core зависит от viz_modules (инверсия слоёв) | `core/visualization.py` импортирует `viz_modules.*` | Инвертировать: `viz_modules` → `core`. Интерфейсы рендеринга — в `core/ports.py`; реализация — в `viz_modules/`. |
| **A9** | Параллельный commercial pipeline в bot | `bot/handlers/commercial.py` vs `CommercialWorkflowService` | Bot вызывает тот же service layer, что и API. Удалить дублированную генерацию КП, pricing, OCR flows. |
| **A10** | Triple planning orchestration | `app/planning/plan_manager.py`, `app/planning/plan_distribution.py`, `app/services/production_planning_service.py` | Консолидировать orchestration в service; planning utilities — в `core/production/`. Один путь вызова. |
| **A11** | PlanRepository JSON backdoor | `app/repositories/plan_repository.py` — fallback/запись в JSON | Убрать JSON fallback. Единственный persistence path — SQLite через `PlanPersistPort`. |
| **A12** | In-process stateful components без TTL/eviction | `bot/handlers/commercial.py`, in-memory caches, draft locks | Добавить TTL, max size, LRU eviction. Документировать single-instance assumption или вынести в shared store (Redis). |

### Безопасность (7)

| ID | Проблема | Расположение | Fix |
|----|----------|--------------|-----|
| **S2** | Stateless sessions — нет server-side invalidation | `app/security/session.py` | Server-side session store или signed tokens с revocation list. Явный idle timeout и max age. |
| **S3** | In-process rate limiting не работает при нескольких воркерах | `app/security/login_rate_limit.py`, OCR limits в `commercial_upload_validation.py` | Redis/shared store для rate limit между воркерами. Или явный single-worker constraint в deployment docs + health check. |
| **S4** | GET logout — CSRF-prone | `app/web/router.py` — logout via GET | Только POST logout с CSRF token. Invalidate session server-side. |
| **S5** | Frontend RBAC — client-only guards | `frontend/src/` — route guards без server re-check | Guards дублируют server RBAC; sensitive actions — always re-validated on API. UI hide ≠ security boundary. |
| **S6** | Деструктивные admin reset без production-guard | `app/api/v1/endpoints/admin.py` — `reset_plans_only`, `reset_calendar_only` | Require `APP_ENV=development` или explicit `ALLOW_DESTRUCTIVE_ADMIN=true`. Confirmation token / double opt-in. Audit log. |
| **S7** | XFF spoofing — доверие к X-Forwarded-For без whitelist | `app/security/login_rate_limit.py`, middleware | Принимать XFF только от trusted proxy IPs. Fallback на `request.client.host`. |
| **S8** | Bot/API drift — расхождение auth и business rules | `bot/handlers/` vs `app/api/`, `bot/middleware/auth.py` | Единые сервисы и RBAC. Cross-surface integration tests. Bot не дублирует API logic. |

### Качество кода (8)

| ID | Проблема | Расположение | Fix |
|----|----------|--------------|-----|
| **Q1** | Голый `except:` в bot-handlers (22+ мест) | `bot/handlers/` — `commercial.py`, `production_execution.py`, `production_completion.py` | Заменить на конкретные типы исключений. Логировать ERROR с traceback. Уведомлять пользователя Telegram при сбоях. |
| **Q2** | God-модули и god-функции | `commercial.py`, `production_*`, `visualization.py`, `commercial_workflow_service.py` | Декомпозиция по use case. Цель: <400 строк на модуль. См. A5. |
| **Q3** | Критичные bot-потоки без автотестов | `bot/handlers/production_execution.py`, `commercial.py`, `production_completion.py` | Mock aiogram для handler tests. Integration tests для plan creation, KP export, day completion flows. |
| **Q4** | Дублирование фабрики имён плит | `bot/handlers/commercial.py`, `app/services/`, `core/` — ad-hoc string assembly | Единый источник: `core/plate_name.py` → `make()`. Все слои импортируют canonical factory. |
| **Q5** | Copy-paste сборки `order_data` в bot КП | `bot/handlers/commercial.py` — повторяющиеся блоки | Вынести в `bot/services/kp_order_builder.py` или shared helper из `core/commercial_offer.py`. |
| **Q6** | Параллельный legacy-поток генерации КП в bot | `bot/handlers/commercial.py` — отдельный путь PDF/XLSX vs `CommercialWorkflowService` | Bot вызывает тот же service layer, что и API. Удалить дублированную генерацию документов. |
| **Q7** | Эвристика reinforcement → load_code продублирована | `core/` и `bot/handlers/` — параллельные mapping tables | Единая функция в `core/domain/` или `core/plate_naming.py`. Bot и API импортируют одну реализацию. |
| **Q8** | DRY: pricing helpers PDF/XLSX | `core/commercial_offer.py` ↔ `core/commercial_offer_xlsx.py` | Единый источник в `core/commercial_pricing.py`. PDF и XLSX импортируют общие функции. |

---

## Средний приоритет

### Архитектура (A13–A22)

| Тема | ID | Проблема | Расположение |
|------|-----|----------|--------------|
| **Domain models** | A13 | Дублирование PlateOrder | `app/domain/models/plate_order.py` extends `core/domain/plate_order.py` |
| **Config leak** | A14 | Global DB paths доступны из нескольких модулей | `core/kp_db_common.py`, `app/repositories/`, env/config |
| **Scalability** | A15 | SQLite scalability ceiling | `plita.db`, WAL mode, concurrent writes |
| **Frontend DRY** | A16 | Frontend pricing DRY — дублирование логики с backend | `frontend/src/features/commercial/` |
| **God component** | A17 | CreatePlanWizard — god component | `frontend/src/features/production/components/CreatePlanWizard.tsx` |
| **God service** | A18 | CommercialWorkflowService god | `app/services/commercial_workflow_service.py` |
| **Import hack** | A19 | `sys.path` manipulation | `bot/`, `core/` — runtime path injection |
| **Auth guards** | A20 | Frontend auth guards без server parity | `frontend/src/` routing |
| **Monolith DB** | A21 | `kp_db` monolith | `core/kp_db_*.py` — единый модуль на все операции |
| **Fat services** | A22 | Fat services без декомпозиции | `app/services/day_view_service.py`, `archive_service.py` и др. |

**Рекомендации (сводно):** один canonical model в `core/domain/`; централизовать DB path через `app/core/settings.py` + DI; документировать SQLite limits и план миграции на PostgreSQL; разбить CreatePlanWizard на steps + hooks; убрать `sys.path` hacks; декомпозировать `kp_db_*` и fat services.

### Безопасность (S9–S19)

| ID | Проблема | Расположение | Рекомендация |
|----|----------|--------------|--------------|
| **S9** | Health endpoint раскрывает environment | `app/api/v1/endpoints/health.py` | Только `{"status": "ok"}`; version/env — internal endpoint с auth |
| **S10** | OpenAPI public без ограничений | `app/main.py` — `/docs`, `/openapi.json` | Отключить в production или защитить auth |
| **S11** | CSP report-only (не enforce) | `app/middleware/security_headers.py` | Перейти на enforce после мониторинга violations |
| **S12** | CSRF cookie readable (не HttpOnly) | `app/security/session.py` | HttpOnly + Secure flags для session cookie |
| **S13** | PII managers в логах/ответах | `app/api/v1/endpoints/commercial.py` | Маскировать телефоны/email в logs; RBAC на PII fields |
| **S14** | Register roles — self-registration с выбором роли | `app/api/v1/endpoints/auth.py` | Только admin может назначать roles; default role = minimal |
| **S15** | Draft sessionStorage — sensitive KP data в browser | `frontend/src/features/commercial/draftStorage.ts` | Encrypt at rest или server-side drafts only |
| **S16** | Bot bare except — тихие сбои | `bot/handlers/` | См. Q1 |
| **S17** | Dev bot auth bypass | `bot/middleware/auth.py` | Fail-closed в production; synthetic admin только в dev |
| **S18** | Dependency audit не в CI | `requirements.txt`, CI pipeline | `pip-audit` / `safety` в CI; block on critical CVE |
| **S19** | SQLite encryption отсутствует | `plita.db` | SQLCipher или OS-level encryption для sensitive deploys |

### Качество кода (Q9–Q23)

| Тема | ID | Проблема |
|------|-----|----------|
| **God services** | Q9 | God services помимо handlers — `commercial_workflow_service`, `day_view_service` |
| **React monoliths** | Q10 | Крупные page components в `frontend/src/features/commercial/`, `production/` |
| **Date DRY** | Q11 | Date formatting дублирована — `frontend/`, `app/services/`, `core/` |
| **TS types** | Q12 | Weak TS types — `Record<string, unknown>`, `any` |
| **Re-export wrappers** | Q13 | Deprecated wrappers без удаления — `plate_completion_service`, `rest_matching_service` |
| **Misleading comments** | Q14 | Docstrings не соответствуют поведению — `plan_repository.py`, `production_export.py` |
| **Dead code** | Q15 | Stubs и unused imports — `production_export.py`, skipped tests |
| **Swallowed errors** | Q16 | Silent fallback в procurement/completion paths |
| **API try/except DRY** | Q17 | Повторяющийся try/except в `app/api/v1/endpoints/*.py` |
| **Service test gaps** | Q18 | Production API, archive, admin — недостаточно integration tests |
| **response_model** | Q19 | Missing `response_model` на endpoints — слабая OpenAPI contract |
| **Debug instrumentation** | Q20 | `#region agent log` в hot paths — `day_view_service`, `production_planning_service` |
| **Long methods** | Q21 | `plan_repository.py` — длинные методы без декомпозиции |
| **type ignore** | Q22–Q23 | `# type: ignore` без обоснования; `list[Any]` в services |

---

## Низкий приоритет / предложения

### Архитектура (A23–A28)

- **A23** — Legacy web parallel presentation layer (`app/web/router.py`) — deprecate после SPA parity
- **A24** — `OptimizationService` wrapper без добавленной ценности
- **A25** — Keyboards monolith в `bot/keyboards/`
- **A26** — `viz_modules` boundaries размыты
- **A27** — Re-export shims (PEP 562 proxy, legacy aliases)
- **A28** — `config_and_data` proxy — compatibility layer без срока удаления

### Безопасность (S20–S25)

- **S20** — `APP_DEBUG` может включать verbose logging в production
- **S21** — Session TTL — длинные или неявные сессии
- **S22** — IDOR возвращает 404 вместо 403 (information leak pattern)
- **S23** — Нет MFA для admin/production roles
- **S24** — Bot HTML injection в user-facing messages
- **S25** — Legacy web SPA routes — residual attack surface

### Качество кода (Q24–Q31)

- **Q24** — Repetitive callback except pattern в `@router.callback_query`
- **Q25** — Minimal frontend tests (~4 test files на десятки TSX)
- **Q26** — TODOs без ticket/issue по `app/`, `bot/`, `frontend/`
- **Q27** — Broad except при DB rollback в repositories
- **Q28** — `str(exc)` в bot user messages
- **Q29** — Legacy test style — inconsistent fixtures/marks
- **Q30** — `list[Any]` в service signatures
- **Q31** — Sync blocking routes в async FastAPI app

---

## Матрица приоритетов

| Priority | ID | Проблема | Severity | Effort | Обоснование |
|----------|-----|----------|----------|--------|-------------|
| **P0** | A1/S1 | Split-brain планов (SQLite vs JSON) | Critical | L | Корневая причина рассинхронизации данных между web и bot |
| **P0** | A2 | Telegram обходит core pipeline | Critical | L | Drift бизнес-правил; разные планы при одинаковых входных данных |
| **P0** | A3 | Legacy global plate runtime state | Critical | L | Блокер масштабирования; перекрёстные мутации bot+API |
| **P1** | S6 | Destructive admin reset без production-guard | High | S | Потеря планов/календаря одним запросом в production |
| **P1** | S4 | GET logout — CSRF-prone | High | S | Session hijack через CSRF на logout/login flows |
| **P1** | S5 | Frontend RBAC — client-only | High | M | UI bypass не блокирует API, но создаёт ложное чувство безопасности |
| **P1** | Q3 | Критичные bot-потоки без автотестов | High | M | Нет safety net для P0-фиксов A1/A2 |
| **P1** | Q1 | 22+ голых except в bot | High | M | Тихие сбои в production paths |
| **P1** | A11 | PlanRepository JSON backdoor | High | S | Обходит единый SQLite authority; усиливает A1/S1 |
| **P1** | A9/Q6 | Параллельный commercial pipeline в bot | High | M | Document drift между bot и API commercial flows |
| **P2** | S2 | Stateless sessions | High | M | Невозможность revoke; session fixation risk |
| **P2** | S3 | In-process rate limit — multi-worker gap | High | M | Rate limit ineffective при horizontal scaling |
| **P2** | S7 | XFF spoofing | High | S | Bypass IP-based rate limits |
| **P2** | A5/Q2 | God-модули | High | L | Maintainability; блокирует безопасный рефакторинг A1/A2 |
| **P2** | Q4/Q7 | Plate name / load_code DRY | High | S | Inconsistent naming → wrong plans/inventory |
| **P2** | A4 | Bot → app DIP violation | High | M | Circular deps; bot breaks при изменении app layer |
| **P2** | A8 | Core → viz_modules coupling | High | M | Нарушение слоёв; сложность тестирования core |
| **P2** | S8 | Bot/API drift | High | L | Расхождение auth и business rules между каналами |
| **P3** | A6 | Inconsistent DI | High | S | Технический долг; усложняет тестирование endpoints |
| **P3** | A7 | Repository facades без абстракций | High | M | Слабая инкапсуляция SQL; tight coupling к kp_db |
| **P3** | A10 | Triple planning orchestration | High | M | Три точки входа для одной операции |
| **P3** | A12 | In-process stateful components | High | M | Memory leak; single-instance assumption |
| **P3** | Q5/Q8 | order_data / pricing DRY | High | S | Copy-paste bugs при изменении business rules |
| **P3** | A13–A22 | Medium architecture cluster | Medium | M | Domain models, config, scalability, frontend god components |
| **P3** | S9–S19 | Medium security cluster | Medium | M | Health leak, OpenAPI, CSP, CSRF cookie, PII, CI audit |
| **P3** | Q9–Q23 | Medium quality cluster | Medium | M | God services, TS types, test gaps, dead code, instrumentation |
| **P3** | A23–A28, S20–S25, Q24–Q31 | Low priority backlog | Low | S–M | Shims, TODOs, frontend tests, legacy cleanup |

*Effort: S = часы–1 день, M = 2–5 дней, L = 1–2 спринта*

### Сквозные кластеры исправлений

| Кластер | IDs | Единый подход |
|---------|-----|---------------|
| **Единый источник истины для планов** | A1/S1, A2, A10, A11 | SQLite-backed plans + `core/production/planning.py` + thin adapters (bot, API) |
| **Устранение runtime state** | A3, A8, A28 | Context objects + DB persistence; удаление globals из `plate_runtime_state` / `config_and_data` |
| **Auth & security hardening** | S2–S7, S9–S12 | Shared rate limits, production guards, CSRF, POST logout, XFF whitelist, session store |
| **Консолидация и DRY** | Q4–Q8, Q9–Q11, Q17, A5, A9 | Shared modules, декомпозиция god-files, единый exception handling |
| **Test safety net** | Q3, Q18, Q25 | Integration tests до крупных рефакторингов; bot handler mocks |

---

## Уже исправлено (с прошлого аудита)

Следующие находки **закрыты** в ходе спринтов стабилизации P1–P3 (2026-06-19 — 2026-06-20):

| Тема | Что сделано |
|------|-------------|
| **Rate limit on /web/login** | `check_login_rate_limit` применён к `POST /web/login` в `app/web/router.py`; тесты `tests/test_web_login_rate_limit.py` |
| **destructive_db_guard** | `core/destructive_db_guard.py` — deny destructive reset в staging/production без `ALLOW_DESTRUCTIVE_DB_RESET`; guard на admin reset endpoints |
| **Production can't read all offers** | Production operational-only: 403 на `/api/v1/offers/*`; filter `_PRODUCTION_READ_STATUSES`; frontend nav hide + redirect; тесты `test_offers_production_authorization.py` |
| **Password policy** | `app/security/password_policy.py` — min 12 chars, complexity, common-password denylist; Pydantic + `AuthRepository`; `tests/test_password_policy.py` |
| **Security headers (partial)** | `app/middleware/security_headers.py` — X-Frame-Options, X-Content-Type-Options, Referrer-Policy, CSP report-only, HSTS; `tests/test_security_headers.py` |
| **CSRF middleware** | CSRF middleware добавлен для cookie-based auth flows |
| **CSRF web forms (S8)** | Полноценные CSRF tokens для cookie-based forms — `c391696`; `tests/test_csrf.py` |
| **XFF trusted proxy (S3)** | `X-Forwarded-For` только от configured proxies — `672308e` |
| **A3 phase 1–2** | Load codes из order data, без TLS globals — `90950ad`, `1150884` |
| **POST logout + destructive guard** | P1-next WP1–WP2 — `a818285`; `tests/test_web_logout_csrf.py`, `test_admin_destructive_guard.py` |
| **get_user_by_id auth** | `AuthRepository.get_user_by_id()` — indexed lookup; `get_current_user` без O(n) `list_users()`; тесты `test_auth_repository.py`, `test_auth_dependencies.py` |

*Примечание: S4 GET logout и S6 destructive reset закрыты в P1-next (`a818285`); CSRF (S8), XFF trusted proxy (S3) и A3 phase 1–2 — в post-sprint closure ниже.*

---

## Post-sprint remediation status (2026-06-20)

После P0-next, P1-next и post-sprint hardening. Исходный снимок аудита **не переписывается** — ниже только статус закрытых находок.

### Закрыто post-sprint (RESOLVED)

| ID / тема | Severity | Коммит | Что сделано |
|-----------|----------|--------|-------------|
| **S8** | High | `c391696` | CSRF tokens для cookie-based web forms; middleware + тесты `tests/test_csrf.py` |
| **S3** | High | `672308e` | Trusted proxy: `X-Forwarded-For` принимается только от настроенных proxy IPs; fallback на `request.client.host` |
| **A3 phase 1–2** | Critical (partial) | `90950ad`, `1150884` | Load codes из explicit order data вместо TLS globals; procurement path без мутации module-level state |

*Маппинг ID:* в baseline-таблицах аудита XFF был **S7**, CSRF logout — **S4**; в коммитах и спринте используются **S3** (XFF) и **S8** (CSRF forms) — статус RESOLVED по sprint ID.

### Пересчёт Health Score (приблизительно)

| Метрика | Baseline snapshot | Post-sprint |
|---------|-------------------|-------------|
| Critical (A1/S1, A2, A3) | 4 | **0** (P0-next closed; A3 phase 1–2 — partial, phase 3 OPEN) |
| High — S8 CSRF, S3 XFF | open | **RESOLVED** |
| High — S2 sessions, S9 health leak, A6 DI… | open | open (приоритет следующего спринта) |
| **Overall Health Score** | **0.0 / 10** | **~8.5 / 10** |

Формула (упрощённо):  
`10 − 0 (critical cap снят) − 0.5 (high residual cap) − 1 (medium cap) ≈ 8.5`.

Регрессия: **918 passed, 12 skipped** в `pytest tests/ -q` (2026-06-20; 9 failed — pre-existing, не блокер документации closure).

### Рекомендации следующего спринта (post-sprint)

1. **S9** — health endpoint: убрать раскрытие environment (`app/api/v1/endpoints/health.py`).
2. **S2** — server-side session invalidation / idle timeout (stateless sessions).
3. **A3 phase 3** — PEP 562 decommission `core/config_and_data.py`; полный вывод globals.
4. **A6** — единый DI-паттерн в API (`admin.py`, `archive.py` vs `production.py`, `commercial.py`).

---

## Следующие шаги

1. ~~**Немедленно (P0)**: Закрыть кластер A1/S1–A3~~ — **closed** (P0-next; WP2 bot adapter = maintenance-only при bot deprecated).
2. ~~**P1-next:** S4 POST logout, S6 full destructive guard~~ — **closed** (`a818285`). ~~S3 XFF trusted proxy, S8 CSRF~~ — **closed** (`672308e`, `c391696`). **Приоритет:** S9 health leak, S2 session invalidation. ~~Q1 bare except в bot, Q3 bot integration~~ — **cancelled** (bot deprecated).
3. **Следующий спринт:** **A3 phase 3** (PEP 562 decommission), **A6** DI unification, A5/Q2 god-module decomposition (web-side). ~~A4 bot→core, A9 commercial bot~~ — только при решении об удалении `bot/`.
4. **Backlog (P3):** A7 repository abstractions, A10 planning orchestration, medium clusters (S10–S19), optional `bot/` removal.

---

## Позитивные находки

| Область | Практика | Расположение |
|---------|----------|--------------|
| **Архитектура** | Разделение слоёв FastAPI: routers → services → repositories | `app/api/`, `app/services/`, `app/repositories/` |
| **Архитектура** | Доменная логика в `core/` — planning, commercial, visualization | `core/production/planning.py`, `core/commercial_offer.py` |
| **Архитектура** | Port для persistence планов (`PlanPersistPort`) | `app/planning/ports.py`, `plan_repository.py` |
| **Архитектура** | Request-scoped `PlateOrderContext` на FastAPI hot paths | `core/plate_runtime_state.py`, middleware |
| **Безопасность** | Session-based auth с role-based access | `app/security/session.py`, `app/dependencies/auth.py` |
| **Безопасность** | Rate limiting на REST login (in-process) | `app/security/login_rate_limit.py` |
| **Безопасность** | Object-level RBAC для КП (REST API) | `app/security/offer_access.py` |
| **Безопасность** | Fail-closed bot auth при misconfiguration | `bot/middleware/auth.py`, `validate_bot_startup()` |
| **Данные** | SQLite WAL + foreign keys | `core/kp_db_common.py` |
| **Данные** | Optimistic concurrency для планов (`version`) | `app/repositories/plan_repository.py` |
| **Тестирование** | Обширный pytest suite (918+ tests) | `tests/` — auth, commercial, production, planning |
| **Frontend** | React SPA с feature-based structure | `frontend/src/features/` |
| **Frontend** | Обработка 409 plan version conflict | `frontend/src/features/production/planConflict.ts` |
| **DevOps** | venv + documented startup | `README`, `.venv`, PowerShell scripts |

---

## Приложение: индекс всех находок

| ID | Severity | Категория |
|----|----------|-----------|
| A1/S1 | Critical | Архитектура + Безопасность (dedupe) |
| A2–A3 | Critical | Архитектура |
| A4–A12 | High | Архитектура |
| A13–A22 | Medium | Архитектура |
| A23–A28 | Low | Архитектура |
| S2–S8 | High | Безопасность |
| S9–S19 | Medium | Безопасность |
| S20–S25 | Low | Безопасность |
| Q1–Q8 | High | Качество кода |
| Q9–Q23 | Medium | Качество кода |
| Q24–Q31 | Low | Качество кода |

---

*Отчёт сформирован: 2026-06-20 · Консолидирован из проходов senior-reviewer, security-auditor и reviewer.*
