# Отчёт аудита проекта

**Дата:** 2026-06-21  
**Область:** Полный проект (`app/`, `core/`, `bot/`, `viz_modules/`, `tests/`)  
**Аудит провели:** senior-reviewer + security-auditor + reviewer

> **Remediation specs:** P3 closed · P4 closed (964 passed) · **[P5 draft](../../specs/stabilizaciya-p5-architecture-2026-06-21.md)** — bot soft decommission (D1), A7 god-modules ×2, Q6 CreatePlanWizard. Out of scope P5: S1 Redis, S3 OCR, A9 PostgreSQL → P6.

---

## Executive Summary

**Общий Health Score:** 0.0/10

| Severity | Architecture | Security | Code Quality | Total |
|----------|-------------|----------|--------------|-------|
| Critical | 3 | 0 | 0 | **3** |
| High | 7 | 5 | 11 | **23** |
| Medium | 8 | 9 | 11 | **28** |
| Low | 4 | 6 | 3 | **13** |

**Формула Health Score:** 10 − 6 (cap critical) − 3 (cap high) − 1 (cap medium) = **0.0**

**Рекомендация:** Перед следующим релизом необходимо устранить **3 критических архитектурных дефекта** — глобальное plate-runtime в bot, инверсию зависимостей bot → app и обход visualization ports в app. Эти проблемы создают риск утечки данных между concurrent-сессиями, нарушают границы слоёв и усложняют тестирование. Параллельно следует запланировать устранение 23 high-priority находок по безопасности, архитектуре и качеству кода.

---

## Critical Issues (исправить немедленно)

### [A1] Глобальное plate-runtime в bot

**Category:** Architecture  
**Location:** `core/config_and_data.py`, `core/plate_runtime_state.py`, `bot/handlers/commercial.py`, `production_*.py`, `optimize.py`, `bot/services/production_planning_adapter.py`  
**Impact:** Утечка данных между concurrent-сессиями Telegram-бота и HTTP-запросами; один пользователь может видеть или изменять состояние заказа другого.  
**Fix:** Ввести `PlateOrderContext` и явную передачу `PlateOrder` через все вызовы вместо глобального mutable runtime state.

---

### [A2] Инверсия зависимостей bot → app

**Category:** Architecture  
**Location:** `bot/services/production_planning_adapter.py`, `bot/handlers/commercial.py`, `production_execution.py`, `plan_manager.py`  
**Impact:** Bot-слой зависит от app-сервисов, нарушая направление зависимостей; расширяет attack surface, усложняет изоляцию и unit-тестирование bot без поднятия FastAPI.  
**Fix:** Перенести orchestration в `core/production/`; bot и app должны зависеть только от core-портов и адаптеров.

---

### [A3] App обходит visualization ports

**Category:** Architecture  
**Location:** `app/services/commercial_service.py:27-28`  
**Impact:** Прямой импорт из `viz_modules`/`core.visualization` минует `core.ports.visualization`; нарушает границу модулей, усложняет замену реализации и регрессионное тестирование.  
**Fix:** Маршрутизировать все вызовы визуализации через facades `core.ports.visualization` и app-адаптеры.

---

## High Priority Issues (исправить в ближайший спринт)

### Architecture

#### [A4] God module core/visualization.py

**Category:** Architecture  
**Location:** `core/visualization.py` (~1238 строк)  
**Impact:** Смешение ответственностей, высокий риск регрессий при изменениях, сложность code review и тестирования.  
**Fix:** Разбить на подмодули по доменам (drawing, layout, export) с явными публичными API.

#### [A5] God module core/kp_db_offers.py

**Category:** Architecture  
**Location:** `core/kp_db_offers.py` (~1043 строк)  
**Impact:** Смешение persistence, бизнес-логики и форматирования; затрудняет изоляцию и покрытие тестами.  
**Fix:** Выделить repository, service и mapper-слои; оставить в модуле только координацию.

#### [A6] God object CommercialWorkflowService

**Category:** Architecture  
**Location:** `app/services/commercial_service.py` (~775 строк)  
**Impact:** Единый сервис управляет загрузкой, OCR, pricing, persistence — нарушение SRP, сложность сопровождения.  
**Fix:** Декомпозировать на отдельные сервисы (upload, pricing, persistence) с orchestrator-facade.

#### [A7] Sync CPU-bound optimization в HTTP

**Category:** Architecture  
**Location:** HTTP endpoints, вызывающие sync optimization  
**Impact:** Блокировка worker-потоков FastAPI под нагрузкой; деградация latency и throughput.  
**Fix:** Вынести оптимизацию в background task / thread pool с явным job API или Celery/RQ.

#### [A8] Дублирование planning-деревьев

**Category:** Architecture  
**Location:** `core/production/`, `app/planning/`, `bot/`  
**Impact:** Расхождение логики планирования между web и bot; двойные баги и сложность синхронизации изменений.  
**Fix:** Единый источник истины в `core/production/planning.py`; app и bot — thin adapters.

#### [A9] Raw SQL в plan_distribution_service

**Category:** Architecture  
**Location:** `app/services/plan_distribution_service.py:319-330`  
**Impact:** Обход repository-слоя; сложнее тестировать, выше риск SQL-ошибок при рефакторинге схемы.  
**Fix:** Перенести запрос в `PlanRepository` с параметризованным ORM/SQLAlchemy.

#### [A10] Bot god-handlers

**Category:** Architecture  
**Location:** `bot/handlers/commercial.py` (~2048), `production_completion.py` (~1196), `production_day_view.py` (~1044)  
**Impact:** Handlers содержат бизнес-логику, DB и UI — невозможность unit-тестов без полного контекста aiogram.  
**Fix:** Вынести логику в core/app services; handlers — только routing и форматирование ответов.

---

### Security

#### [S1] In-process rate limiting не масштабируется

**Category:** Security  
**Location:** `app/security/login_rate_limit.py`, `commercial_upload_validation.py`  
**Impact:** При нескольких инстансах приложения лимиты не синхронизируются; brute-force и abuse возможны через round-robin.  
**Fix:** Redis-backed rate limiter или API gateway rate limiting.

#### [S2] Утечка plate-runtime в thread pool

**Category:** Security  
**Location:** `core/plate_runtime_state.py`, `bot/handlers/commercial.py`  
**Impact:** Thread-local/global state не изолирован между concurrent задачами — см. также [A1]; риск cross-session data leak.  
**Fix:** Явный context per request/task; запрет глобального mutable state в async/thread pool.

#### [S3] OCR документов в OpenAI API

**Category:** Security  
**Location:** `core/ocr_gpt.py`, commercial endpoints  
**Impact:** Коммерческие документы (PII, цены, контракты) передаются во внешний API; риск утечки и compliance.  
**Fix:** On-prem OCR, data processing agreement, redaction pipeline или opt-in с явным consent.

#### [S4] Dev-only обход Telegram-auth

**Category:** Security  
**Location:** `bot/middleware/auth.py`, `settings.py`  
**Impact:** При misconfiguration в production auth middleware может быть отключён.  
**Fix:** Fail-closed по умолчанию; dev bypass только при явном `ENV=development` + guard в startup.

#### [S5] Destructive DB reset в development

**Category:** Security  
**Location:** `destructive_db_guard.py`, admin endpoints  
**Impact:** При ошибке конфигурации destructive операции могут быть доступны вне dev.  
**Fix:** Двойной guard (env + feature flag); audit log всех destructive вызовов.

---

### Code Quality

#### [Q1] 283 функции ≥31 строки, mega-модули 500–965 строк

**Category:** Code Quality  
**Location:** По всему проекту (`core/`, `app/`, `bot/`, `viz_modules/`)  
**Impact:** Высокая cognitive complexity, сложность review и регрессий.  
**Fix:** Инкрементальный refactor: extract method, split modules по SRP.

#### [Q2] Drift fuzzy-tolerance web 0.005 vs bot 0.03

**Category:** Code Quality  
**Location:** Web services vs bot handlers  
**Impact:** Разные результаты lookup/plate matching между web и bot для одних данных.  
**Fix:** Единая константа в `core/`; shared helper + parity-тесты.

#### [Q3] Triplicated day-view/day-docs логика

**Category:** Code Quality  
**Location:** `app/services/day_view_service.py`, `day_documents_service.py`, bot handlers  
**Impact:** Тройное дублирование; баги исправляются только в одном месте.  
**Fix:** Consolidate в core service; app и bot — thin wrappers.

#### [Q4] Coupling через private API day_documents_service

**Category:** Code Quality  
**Location:** `app/services/day_documents_service.py` (private methods)  
**Impact:** Внешние модули зависят от `_`-prefixed API; refactor ломает callers silently.  
**Fix:** Оформить публичный интерфейс; private — только internal.

#### [Q5] Пробелы тестов: kp_db_offers, day_documents, bot handlers

**Category:** Code Quality  
**Location:** `core/kp_db_offers.py`, `day_documents_service.py`, `bot/handlers/`  
**Impact:** Критические пути без регрессионной защиты.  
**Fix:** Приоритетные integration/unit tests для top-risk модулей.

#### [Q6] Bare except: в bot handlers (17+ мест)

**Category:** Code Quality  
**Location:** `bot/handlers/` (multiple files)  
**Impact:** Глотаются все исключения включая KeyboardInterrupt/SystemExit; скрытые баги.  
**Fix:** Catch specific exceptions; log + re-raise или user-friendly error response.

#### [Q7] DRY error-handling в commercial API (13 блоков)

**Category:** Code Quality  
**Location:** Commercial API endpoints  
**Impact:** Copy-paste try/except; inconsistent error responses.  
**Fix:** Centralized exception handler или decorator для commercial routes.

#### [Q8] Слабая типизация HTTP (dict вместо response_model)

**Category:** Code Quality  
**Location:** Multiple API endpoints  
**Impact:** OpenAPI schema неполная; runtime validation слабее; IDE support хуже.  
**Fix:** Pydantic response_model для всех public endpoints.

#### [Q9] sys.path.insert в production_completion.py

**Category:** Code Quality  
**Location:** `bot/handlers/production_completion.py`  
**Impact:** Хрупкий import resolution; ломается при изменении структуры проекта.  
**Fix:** Proper package imports; editable install / PYTHONPATH в deployment.

#### [Q10] Дублирование procurement pipelines

**Category:** Code Quality  
**Location:** `viz_modules/procurement/`, app services  
**Impact:** Две реализации одного pipeline; drift и двойное сопровождение.  
**Fix:** Single pipeline в viz_modules; app — adapter only.

#### [Q11] layout_sequence/builder.py минимальное покрытие

**Category:** Code Quality  
**Location:** `viz_modules/` layout builder  
**Impact:** Layout bugs без автоматической детекции.  
**Fix:** Unit tests для edge cases (merge, orphan, secondary plates).

---

## Medium Priority Issues (запланировать на следующий спринт)

### Architecture

| ID | Проблема | Location |
|----|----------|----------|
| A11 | In-process rate limiting | `app/security/login_rate_limit.py` |
| A12 | Pass-through repositories | `app/repositories/` |
| A13 | Auth bypass DI | `app/dependencies/auth.py:25,63` |
| A14 | Service locator для viz ports | `app/adapters/visualization.py` |
| A15 | File-based DraftStore | `.app_data/drafts/` |
| A16 | Дублирование archive logic | `app/services/archive_service.py`, bot |
| A17 | Nested service construction per request | `app/dependencies/services.py` |
| A18 | Pass-through app services | Multiple app services |

### Security

| ID | Проблема | Location |
|----|----------|----------|
| S6 | CSRF-cookie без HttpOnly | Session/auth middleware |
| S7 | CSP только Report-Only | `app/main.py` security headers |
| S8 | Нет глобального API rate limiting | API layer |
| S9 | SQLite без шифрования at-rest | Database layer |
| S10 | Длинные stateless-сессии (12ч) | `app/security/session.py` |
| S11 | Auth endpoints обходят DI | Auth endpoints |
| S12 | Legacy `/web/login` без CSRF | Web routes |
| S13 | `/managers` для роли production | `app/api/v1/endpoints/managers.py` |
| S14 | Debug NDJSON с чувствительными данными | Debug logging |

### Code Quality

| ID | Проблема | Location |
|----|----------|----------|
| Q12 | Устаревший docstring tolerance | `day_view_service.py` |
| Q13 | Широкий `except Exception` | Multiple services |
| Q14 | Silent rollback failures | Transaction handlers |
| Q15 | OffersService без unit-тестов | `app/services/` |
| Q16 | gantt_excel.py без тестов | Export module |
| Q17 | Legacy List/Dict type hints | Multiple modules |
| Q18 | Неполные аннотации helpers | Helper functions |
| Q19 | sys.path.insert в тестах | `tests/conftest.py` |
| Q20 | Stub handlers в bot/export.py | `bot/handlers/export.py` |
| Q21 | Swallow-and-continue в archive_service | `archive_service.py` |
| Q22 | Нет тестов fuzzy lookup parity | Tests gap |

---

## Low Priority / Suggestions

### Architecture

| ID | Проблема | Location |
|----|----------|----------|
| A19 | Legacy path constants | `plan_storage.py` |
| A20 | sys.path hacks | Various modules |
| A21 | Import boundary только в unit-тестах | `tests/test_core_viz_import_boundary.py` |
| A22 | Inline DraftStore в dependencies | `app/dependencies/` |

### Security

| ID | Проблема | Location |
|----|----------|----------|
| S15 | PBKDF2 вместо argon2 | Password hashing |
| S16 | Нет MFA | Auth system |
| S17 | change-password раскрывает incorrect | Auth endpoint |
| S18 | pip audit недоступен | CI/tooling |
| S19 | Bot → app расширяет attack surface | Cross-layer deps |
| S20 | HSTS/CSP partial в development | Dev config |

### Code Quality

| ID | Проблема | Location |
|----|----------|----------|
| Q23 | Magic numbers load-code дублируются | Multiple modules |
| Q24 | Пустой subclass в schemas/commercial.py | `app/schemas/commercial.py` |
| Q25 | Временные артефакты | `_tmp_old.py`, `_spec_snip.txt` |

---

## Priority Matrix

| ID | Issue | Severity | Effort | Priority |
|----|-------|----------|--------|----------|
| A1 | Глобальное plate-runtime в bot | Critical | High | **P0 — немедленно** |
| A2 | Инверсия зависимостей bot → app | Critical | High | **P0 — немедленно** |
| A3 | App обходит visualization ports | Critical | Medium | **P0 — немедленно** |
| S2 | Утечка plate-runtime в thread pool | High | High | **P0 — немедленно** (связано с A1) |
| S4 | Dev-only обход Telegram-auth | High | Low | **P0 — немедленно** |
| S5 | Destructive DB reset guard | High | Low | **P0 — немедленно** |
| A7 | Sync CPU-bound optimization в HTTP | High | Medium | **P1 — этот спринт** |
| A8 | Дублирование planning-деревьев | High | High | **P1 — этот спринт** |
| A9 | Raw SQL в plan_distribution_service | High | Low | **P1 — этот спринт** |
| S1 | In-process rate limiting | High | Medium | **P1 — этот спринт** |
| S3 | OCR в OpenAI API | High | Medium | **P1 — этот спринт** |
| Q2 | Drift fuzzy-tolerance web vs bot | High | Low | **P1 — этот спринт** |
| Q3 | Triplicated day-view/day-docs | High | Medium | **P1 — этот спринт** |
| Q6 | Bare except в bot handlers | High | Medium | **P1 — этот спринт** |
| A4 | God module visualization.py | High | High | **P1 — этот спринт** |
| A5 | God module kp_db_offers.py | High | High | **P1 — этот спринт** |
| A6 | God object CommercialWorkflowService | High | High | **P1 — этот спринт** |
| A10 | Bot god-handlers | High | High | **P1 — этот спринт** |
| Q1 | Mega-functions и mega-modules | High | High | **P1 — этот спринт** (инкрементально) |
| Q5 | Пробелы тестов critical paths | High | Medium | **P1 — этот спринт** |
| Q7 | DRY error-handling commercial API | High | Low | **P1 — этот спринт** |
| Q8 | Слабая типизация HTTP | High | Medium | **P1 — этот спринт** |
| Q4 | Coupling private API day_documents | High | Low | **P1 — этот спринт** |
| Q9 | sys.path.insert production_completion | High | Low | **P1 — этот спринт** |
| Q10 | Дублирование procurement pipelines | High | Medium | **P1 — этот спринт** |
| Q11 | layout_sequence минимальное покрытие | High | Medium | **P1 — этот спринт** |
| A11–A18 | Medium architecture findings | Medium | Mixed | **P2 — следующий спринт** |
| S6–S14 | Medium security findings | Medium | Mixed | **P2 — следующий спринт** |
| Q12–Q22 | Medium code quality findings | Medium | Mixed | **P2 — следующий спринт** |
| A19–A22 | Low architecture suggestions | Low | Low | **Backlog** |
| S15–S20 | Low security suggestions | Low | Low | **Backlog** |
| Q23–Q25 | Low code quality suggestions | Low | Low | **Backlog** |

---

## Next Steps

### 1. Immediate (до следующего коммита / релиза)

- **[A1] + [S2]:** Устранить глобальное plate-runtime — `PlateOrderContext` + явная передача state
- **[A2]:** Начать декoupling bot → app; orchestration в `core/production/`
- **[A3]:** Подключить `app/services/commercial_service.py` через `core.ports.visualization`
- **[S4]:** Fail-closed guard для Telegram-auth bypass
- **[S5]:** Усилить double-guard для destructive DB operations

### 2. This sprint (текущий спринт)

- **[A7], [A8], [A9]:** Async optimization, единое planning tree, SQL → repository
- **[S1], [S3]:** Redis rate limiter; политика OCR / data residency
- **[A4]–[A6], [A10]:** Начать декомпозицию god-modules (инкрементально)
- **[Q1]–[Q11]:** Fuzzy parity, day-view consolidation, bare except cleanup, tests, typing

### 3. Next sprint (следующий спринт)

- **[A11]–[A18]:** DI cleanup, DraftStore abstraction, archive dedup, service construction
- **[S6]–[S14]:** HttpOnly CSRF, enforce CSP, global API rate limit, session TTL, RBAC review
- **[Q12]–[Q22]:** Exception handling, test coverage gaps, type hints migration

### 4. Backlog

- **[A19]–[A22], [S15]–[S20], [Q23]–[Q25]:** Legacy cleanup, argon2/MFA, magic numbers, temp artifacts

---

## Remediation Commands

Для устранения находок используйте workflow-команды Cursor:

| Тип проблемы | Команда | Примеры |
|--------------|---------|---------|
| Структурные / архитектурные / DRY / god-modules | `/refactor` | `/refactor core/plate_runtime_state.py`, `/refactor bot/handlers/commercial.py` |
| Security / behavioral / новая логика | `/implement` | `/implement Redis rate limiter для login`, `/implement PlateOrderContext` |

**Рекомендуемый порядок remediation:**

1. `/implement PlateOrderContext и изоляция plate-runtime` — P0, блокирует data leak
2. `/refactor bot/services/production_planning_adapter.py` — P0, bot → core
3. `/refactor app/services/commercial_service.py` — P0, visualization ports
4. `/implement fail-closed Telegram auth guard` — P0 security
5. `/refactor core/visualization.py` — P1, инкрементальная декомпозиция

---

*Отчёт сгенерирован workflow `/audit`. Health Score 0.0/10 отражает 3 critical + 23 high findings при формуле с caps. Critical security issues при production-конфиге не выявлены.*
