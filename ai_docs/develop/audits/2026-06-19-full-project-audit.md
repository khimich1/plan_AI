# Отчёт полного аудита проекта «Шишов»

**Дата:** 2026-06-19  
**Область:** Полный проект — `app/`, `core/`, `bot/`, `frontend/`, `viz_modules/`  
**Аудиторы:** senior-reviewer + security-auditor + reviewer

---

## Executive Summary

Проект «Шишов» — FastAPI backend, React frontend, Telegram-бот и модули визуализации для коммерческих предложений (КП), оптимизации раскладки плит и производственного планирования. Функционально система зрелая, но аудит выявил **критический архитектурный долг** (два хранилища планов, раздельные pipeline'ы, глобальное состояние), **уязвимость безопасности в Telegram-боте** и **существенные пробелы в качестве кода и тестовом покрытии**.

### Overall Health Score: **0 / 10**

| Severity | Architecture | Security | Code Quality | **Total** |
|----------|-------------|----------|--------------|-----------|
| Critical | 3           | 1        | 0            | **4**     |
| High     | 5           | 4        | 6            | **15**    |
| Medium   | 6           | 7        | 8            | **21**    |
| Low      | 4           | 6        | 5            | **15**    |
| **Total**| **18**      | **18**   | **19**       | **55**    |

*Формула: 10 − 6 (critical cap) − 3 (high cap) − 1 (medium cap) = **0***

### Ключевые темы

1. **Два независимых хранилища планов** (SQLite и JSON на диске) — риск рассинхронизации и потери данных.
2. **Telegram-планирование обходит core pipeline** — бизнес-правила дублируются и расходятся между bot и API.
3. **Глобальное мутабельное состояние плит** — блокирует горизонтальное масштабирование и создаёт перекрёстные утечки данных.
4. **Admin-доступ к боту при неверном `APP_ENV`** — критическая уязвимость авторизации.
5. **God-модули, DRY-нарушения и дыры в тестах** — высокий риск регрессий при любых изменениях.

### Рекомендация

**Приостановить разработку новых фич на один спринт** и сфокусироваться на стабилизации P0-кластера: единый источник истины для планов, закрытие S1, устранение глобального runtime state. Без этого production-деплой несёт риск **тихой порчи данных**, **полного admin-доступа к боту** и **регрессий, невидимых для CI**.

---

## Critical Issues (fix immediately)

### [A1] Два независимых хранилища планов (SQLite vs JSON)

| Поле | Детали |
|------|--------|
| **Категория** | Архитектура |
| **Расположение** | `app/repositories/plan_repository.py`, `app/planning/plan_storage.py`, `bot/data/plans/` |
| **Impact** | Производственные планы сохраняются параллельно в SQLite (`plita.db`) и как JSON-файлы на диске. FastAPI, Telegram-бот и фоновые процессы могут читать/писать в разные источники. Изменения в одном хранилище не отражаются в другом — операторы видят разные планы, счётчики завершения и метаданные. При сбое миграции или race condition возможна тихая потеря данных. |
| **Fix** | Выбрать **единый авторитетный источник** — SQLite с транзакционными записями. Реализовать `PlanRepository` с optimistic concurrency (версия строки). Убрать прямой file I/O из handlers и `plan_storage.py`. Написать скрипт миграции `bot/data/plans/*.json` → SQLite. Добавить integration-тест: запись через API читается ботом и наоборот. |

---

### [A2] Telegram-планирование не использует core pipeline

| Поле | Детали |
|------|--------|
| **Категория** | Архитектура |
| **Расположение** | `bot/handlers/production_execution.py` ↔ `core/production/planning.py` |
| **Impact** | Бот реализует собственный путь загрузки КП → оптимизация → треки → сохранение плана (~1500 строк), не вызывая `core/production/planning.py`. Исправления и новые бизнес-правила, внесённые в core или FastAPI-сервис, не попадают в Telegram-поток. Операторы через бота и через веб-API получают разные планы при одинаковых входных данных. |
| **Fix** | Сделать `core/production/planning.py` единственным доменным модулем планирования с чистыми функциями и общими DTO. `bot/handlers/production_execution.py` и `app/services/production_planning_service.py` — тонкие адаптеры (aiogram / FastAPI). Добавить cross-surface тесты с одинаковыми fixture-данными. Новые правила — только в core; дублированную логику удалять инкрементально. |

---

### [A3] Глобальное мутабельное состояние plate runtime

> **Статус (2026-06-19):** ✅ **RESOLVED** в P1 — hot paths на `PlateOrderContext`/DI; см. [Post-P1](#post-p1-remediation-status-2026-06-19). Cold legacy `cfg` proxy — backlog.

| Поле | Детали |
|------|--------|
| **Категория** | Архитектура |
| **Расположение** | `core/plate_runtime_state.py`, `core/config_and_data.py` |
| **Impact** | Состояние заказов плит и конфигурация оптимизации хранятся в глобальных мутабельных структурах процесса. Параллельные запросы FastAPI сериализуются через locks или гонят друг друга. При совместном использовании бота и API в одном процессе — перекрёстная утечка данных между пользователями/сессиями (связано с S5). Горизонтальное масштабирование (несколько воркеров) невозможно без внешнего state store. |
| **Fix** | Заменить globals на явные context-объекты (`PlateOrderContext`), передаваемые через DI. Авторитетное состояние — в БД; in-memory кэш только read-through с TTL. Убрать мутацию на уровне модулей. Рефакторить сервисы (`day_documents_service`, `archive_service`) для приёма состояния через параметры, а не глобальные locks. |

---

### [S1] Полный admin-доступ к Telegram-боту при неверном APP_ENV

> **Статус (2026-06-19):** ✅ **RESOLVED** в P1 — fail-closed auth, synthetic admin только в development; см. [Post-P1](#post-p1-remediation-status-2026-06-19).

| Поле | Детали |
|------|--------|
| **Категория** | Безопасность |
| **Расположение** | `bot/middleware/auth.py:69-78` |
| **Impact** | При `BOT_AUTH_ENABLED=false` или неверном `APP_ENV` middleware создаёт синтетического admin-пользователя для **всех** входящих сообщений. Любой, кто знает username бота, получает полный доступ к коммерческим данным, планам и админ-командам. В production-деплое с ошибочной конфигурацией — полная компрометация бота без аутентификации. |
| **Fix** | Убрать synthetic admin bypass из production paths. `BOT_AUTH_ENABLED=false` — только при `APP_ENV=development` с явным guard и warning в логах. Fail-closed: при отсутствии валидной auth-конфигурации бот отказывает в обработке. Добавить startup-check: если `APP_ENV=production` и auth отключён — `sys.exit(1)`. Покрыть тестом: production env + disabled auth → бот не стартует. |

---

## High Priority Issues (fix soon)

### Архитектура (5)

| ID | Проблема | Расположение | Рекомендация |
|----|----------|--------------|--------------|
| **A4** | bot → app dependency inversion | `bot/handlers/` импортируют `app.services.*`; данные в `bot/data/` | Инвертировать зависимости: bot → core (домен), app → core. Вынести общие сервисы в `core/` или `shared/`. Bot не должен импортировать app-слой. |
| **A5** | core зависит от viz_modules | `core/visualization.py` импортирует `viz_modules.*` | Направление зависимостей: только `viz_modules` → `core`. Интерфейсы рендеринга — в `core/ports.py`; реализация — в `viz_modules/`. |
| **A6** | God-модули в bot handlers | `commercial.py` (~2309), `production_completion.py` (~1648), `production_execution.py` (~1515) | Разделить по use case (команды / callbacks / state machines). Вынести логику в `bot/services/`. Цель: <300 строк на handler-модуль. |
| **A7** | Размытые границы repository vs domain для планов | `app/repositories/plan_repository.py` ↔ `app/planning/plan_manager.py` | Repository — только I/O и mapping. Domain/service — бизнес-правила (треки, календарь, агрегация). Убрать бизнес-логику из repository passthrough. |
| **A8** | Два runtime-пути для состояния плит (API vs bot) | FastAPI: `plate_runtime_isolation` middleware; bot: legacy globals через `core/config_and_data` | Унифицировать на DB-backed state (см. A3). Бот и API используют один `PlanRepository` и `PlateOrderContext`. |

### Безопасность (4)

| ID | Проблема | Расположение | Рекомендация |
|----|----------|--------------|--------------|
| **S2** | Brute-force на login без rate limiting | `app/api/v1/endpoints/auth.py` | ✅ **RESOLVED** в P2 — `app/security/login_rate_limit.py`, 5 req/min/IP → 429; см. [Post-P2](#post-p2-remediation-status-2026-06-20). |
| **S3** | RBAC без object-level authorization для КП | `app/api/v1/endpoints/offers.py`, `archive.py` | ✅ **RESOLVED** в P2 — `app/security/offer_access.py`, filter по `owner_user_id`; см. [Post-P2](#post-p2-remediation-status-2026-06-20). Gaps: legacy web, bot. |
| **S4** | npm CVE (react-router RCE, vite, undici) | `frontend/package.json` | `npm audit fix`; зафиксировать patched versions вместо `"latest"`. Добавить `npm audit` в CI pipeline. |
| **S5** | Перекрёстная утечка через global plate runtime | `core/plate_runtime_state.py`, `core/config_and_data.py` | Решается вместе с A3: request-scoped context, убрать shared mutable state. Добавить тест изоляции параллельных запросов. |

### Качество кода (6)

| ID | Проблема | Расположение | Рекомендация |
|----|----------|--------------|--------------|
| **Q1** | DRY: `calculate_total_cost` / `format_phone` дублированы | `core/commercial_offer.py` ↔ `core/commercial_offer_xlsx.py` | Единый источник в `core/commercial_pricing.py` или shared utils. XLSX и PDF импортируют общие функции. |
| **Q2** | God-файлы | `commercial.py`, `production_*`, `visualization.py` | Декомпозиция по use case / renderers. Цель: <400 строк на модуль. См. A5, A6. |
| **Q3** | Массовое подавление исключений в bot | `bot/handlers/` — голые `except:` и `pass` | Заменить на конкретные типы исключений. Логировать ERROR с traceback. Уведомлять пользователя Telegram при сбоях. |
| **Q4** | Тихий fallback на legacy в day_view_service | `app/services/day_view_service.py` | Убрать silent fallback. Явная ошибка или structured warning в API response. Логировать выбор legacy path. |
| **Q5** | Test coverage gaps | `day_documents_service`, production API, bot handlers | Integration-тесты для production API (17 routes). Mock aiogram для bot handlers. Golden-file tests для PDF/XLSX. |
| **Q6** | Debug-instrumentation (`#region agent log`) в hot path | `day_view_service.py`, `production_planning_service.py`, `production_execution.py` | Удалить или gate за `APP_DEBUG`. Не писать debug-файлы в production paths. |

---

## Medium Priority Issues (plan for next sprint)

### Архитектура (6)

| ID | Проблема | Расположение | Рекомендация |
|----|----------|--------------|--------------|
| **A9** | Repository passthrough к `core.kp_db_*` | `app/repositories/kp_repository.py`, `kp_archive_repository.py` | Ввести Protocol/interface. Repository инкапсулирует SQL и mapping, не делегирует напрямую в core DB-модули. |
| **A10** | Параллельный legacy web UI | `app/web/router.py`, подключён в `app/main.py` | Deprecate legacy routes; редирект на React SPA. Удалить шаблоны после подтверждения паритета. |
| **A11** | God-модуль `core/visualization.py` | `core/visualization.py` (~1300+ строк) | Декомпозировать на `core/visualization/` (layout, legend, export). Минимальный публичный API. |
| **A12** | Re-export facades в `app/services` | `plate_completion_service.py`, `rest_matching_service.py`, `kp_persistence_service.py` | Убрать no-op re-exports или добавить реальную оркестрацию. Прямые импорты из core где уместно. |
| **A13** | Дублирование PlateOrder domain model | `app/domain/models/plate_order.py` extends `core/domain/plate_order.py` | Один canonical model в `core/domain/`. App-слой использует core model или thin adapter. |
| **A14** | Дублирование production day flows | `day_view_service`, bot `production_day_view`, `production_completion` | Выделить общий use case в core/service. Bot и API — thin adapters. |

### Безопасность (7)

| ID | Проблема | Расположение | Рекомендация |
|----|----------|--------------|--------------|
| **S6** | `str(exc)` утечки в API | `production.py`, `archive.py`, другие endpoints | Централизованный exception handler. Клиенту — generic message; детали — только в server logs. |
| **S7** | OCR rate limit in-process only | `app/services/commercial_upload_validation.py` | Redis/shared store для rate limit между воркерами. Или single-worker constraint в deployment docs. |
| **S8** | Коммерческие изображения в OpenAI | `core/ocr_gpt.py`, `commercial_upload_validation.py` | Редактировать/токенизировать PII перед API. DPA review. On-prem OCR для sensitive deploys. |
| **S9** | Debug-логирование PII при `APP_DEBUG` | `core/debug_paths.py`, debug handlers | Gate PII logging за explicit flag. Никогда не логировать client data в production. |
| **S10** | Слабая password policy | `app/schemas/auth.py` | Минимум 12 символов, complexity rules через Pydantic validator. Common-passwords denylist. |
| **S11** | Нет CSRF на web forms | `app/web/router.py`, `app/security/session.py` | CSRF tokens для cookie-based auth. `SameSite=lax` недостаточно для state-changing POST. |
| **S12** | Health endpoint раскрывает environment | `app/api/v1/endpoints/health.py` | Возвращать только `{"status": "ok"}`. Environment/version — только для internal/admin endpoint с auth. |

### Качество кода (8)

| ID | Проблема | Расположение | Рекомендация |
|----|----------|--------------|--------------|
| **Q7** | DRY try/except в API endpoints | `app/api/v1/endpoints/*.py` | Декоратор или dependency для единообразной обработки domain exceptions → HTTP responses. |
| **Q8** | Дублирование валидации wizard КП | Frontend wizard ↔ backend commercial schemas | Single validation contract: backend — source of truth; frontend — generated types или shared schema. |
| **Q9** | TypeScript `Record<string, unknown>` в production types | `frontend/src/features/production/types/production.ts` | Явные интерфейсы для plan, track, day view. Убрать loose typing на критических моделях. |
| **Q10** | Dead code: export handlers-заглушки | `bot/handlers/production_export.py` | Удалить stubs или реализовать. Не оставлять misleading no-op handlers. |
| **Q11** | DRY monkeypatch в тестах | `tests/` — повторяющиеся patch patterns | Shared fixtures в `conftest.py`. Factory helpers для mock services. |
| **Q12** | Устаревшие skipped-тесты | `tests/` — `@pytest.mark.skip` без ticket | Удалить или восстановить с привязкой к issue. Не накапливать dead test debt. |
| **Q13** | `commercial_workflow_service` смешивает ответственности | `app/services/commercial_workflow_service.py` (~1000+ строк) | Разделить: upload/OCR, pricing, document generation, persistence — отдельные services. |
| **Q14** | Неоднородные ошибки в `plan_repository` | `app/repositories/plan_repository.py` | Единый exception hierarchy (`PlanNotFound`, `PlanConflict`). Консистентный mapping в HTTP errors. |

---

## Low Priority / Suggestions

### Архитектура (4)

| ID | Проблема | Расположение |
|----|----------|--------------|
| **A15** | Module-level singletons в bot | `bot/handlers/`, `bot/services/` |
| **A16** | Lazy import в `plan_distribution` | `app/planning/plan_distribution.py` |
| **A17** | Фрагментированный `kp_db_*` в core | `core/kp_db_offers.py`, `kp_db_schema.py`, `kp_db_common.py` |
| **A18** | Дублирование констант доставки frontend/backend | `core/cargo_delivery_pricing.py` ↔ `frontend/.../cargoDeliveryPricing.ts` |

### Безопасность (6)

| ID | Проблема | Расположение |
|----|----------|--------------|
| **S13** | Нет security headers | CSP, X-Frame-Options, HSTS отсутствуют в `app/` |
| **S14** | Stateless session без revocation | `app/security/session.py` — нет blacklist/revoke |
| **S15** | `get_current_user` загружает всех users | `app/dependencies/auth.py` — O(n) `list_users()` |
| **S16** | Draft в sessionStorage | `frontend/.../draftStorage.ts` |
| **S17** | `"latest"` pins в frontend | `frontend/package.json` |
| **S18** | pip audit недоступен | Python dependency audit не в CI |

### Качество кода (5)

| ID | Проблема | Расположение |
|----|----------|--------------|
| **Q15** | Misleading naming в export handler | `bot/handlers/production_export.py` |
| **Q16** | Emoji-комментарии в prod-коде | Разбросаны по `bot/`, `core/` |
| **Q17** | Frontend test coverage минимальна | ~4 test files на ~87 TS/TSX sources |
| **Q18** | viz_modules builder без unit-тестов | `viz_modules/layout_sequence/` |
| **Q19** | Inconsistent error-handling в admin/archive | `app/api/v1/endpoints/admin.py`, `archive.py` |

---

## Priority Matrix

Топ-проблемы, ранжированные по **бизнес-влиянию × вероятности**, с учётом зависимостей исправлений.

| Priority | ID | Issue | Severity | Effort | Обоснование |
|----------|-----|-------|----------|--------|-------------|
| **P0** | S1 | Полный admin-доступ к боту при неверном APP_ENV | Critical | S | Немедленная компрометация при misconfiguration |
| **P0** | A1 | Два хранилища планов (SQLite vs JSON) | Critical | L | Корневая причина рассинхронизации данных |
| **P0** | A2 | Telegram не использует core pipeline | Critical | L | Drift бизнес-правил между bot и API |
| **P0** | A3 | Глобальное мутабельное plate runtime | Critical | L | Блокер масштабирования; связано с S5 |
| **P1** | S2 | Brute-force на login без rate limiting | High | S | Прямая attack surface на auth |
| **P1** | S3 | RBAC без object-level auth для КП | High | M | Любой manager читает чужие КП |
| **P1** | S4 | npm CVE (react-router RCE, vite, undici) | High | S | Известные CVE в frontend supply chain |
| **P1** | S5 | Перекрёстная утечка через global runtime | High | M | Утечка данных между сессиями; решается с A3 |
| **P1** | Q5 | Test coverage gaps (production API, bot) | High | M | Нет safety net для P0-фиксов |
| **P1** | A4 | bot → app dependency inversion | High | M | Усложняет рефакторинг A1/A2 |
| **P2** | A6 | God-модули в bot handlers | High | L | Maintainability; после дизайна A2 |
| **P2** | Q3 | Массовое подавление исключений в bot | High | M | Тихие сбои в production paths |
| **P2** | Q6 | Debug-instrumentation в hot path | High | S | Шум, I/O overhead, риск утечки в логах |

*Effort: S = часы–1 день, M = 2–5 дней, L = 1–2 спринта*

### Сквозные кластеры исправлений

| Кластер | IDs | Единый подход |
|---------|-----|---------------|
| **Единый источник истины для планов** | A1, A2, A7, A8, A14 | SQLite-backed plans + `core/production/planning.py` + thin adapters (bot, API) |
| **Устранение runtime state** | A3, A8, S5 | Context objects + DB persistence; удаление globals из `plate_runtime_state` |
| **Auth & security hardening** | S1, S2, S3, S10, S11 | Fail-closed bot auth, rate limits, object-level RBAC, CSRF, password policy |
| **Консолидация и DRY** | Q1, Q2, Q7, Q8, A5, A11 | Shared modules, декомпозиция god-files, единый exception handling |
| **Test safety net** | Q5, Q11, Q12, Q17, Q18 | Integration tests до крупных рефакторингов |

---

## Next Steps

### Immediate (before next commit)

1. **Закрыть S1** — убрать synthetic admin bypass; fail-closed при `APP_ENV=production` + disabled auth.
2. **Начать spike A1** — зафиксировать SQLite как единственный store планов; draft migration script для `bot/data/plans/`.
3. **Добавить rate limiting на login** (S2) — 5 попыток/мин на IP.
4. **Удалить debug-instrumentation** из production paths (Q6).

### This sprint (стабилизация)

1. **Завершить кластер A1/A2** — единый `PlanRepository`, bot handler вызывает `core/production/planning.py`.
2. **Рефакторинг A3** — request-scoped `PlateOrderContext`; убрать globals из hot path.
3. **Object-level авторизация КП** (S3) на всех commercial/archive endpoints.
4. **Обновить npm dependencies** (S4); зафиксировать versions; `npm audit` в CI.
5. **Integration-тесты production API** (Q5) — happy path + 3 failure modes.
6. **Устранить silent exceptions в bot** (Q3) — конкретные типы, логирование, user notification.

### Next sprint

1. Инвертировать зависимости bot → core (A4); декомпозиция god-modules (A6, Q2).
2. Направление зависимостей core ↔ viz_modules (A5, A11).
3. Централизованный exception handler (S6, Q7, Q14).
4. Security headers (S13), sanitize health endpoint (S12), CSRF (S11).
5. Deprecate legacy web UI (A10).
6. Тесты bot handlers и `day_documents_service` (Q5).

### Backlog

- Repository abstractions (A9, A12, A13)
- Password policy (S10), session revocation (S14)
- Cached user lookup вместо `list_users()` (S15)
- Frontend test coverage (Q17), viz_modules unit tests (Q18)
- Консолидация `kp_db_*` (A17), константы доставки (A18)
- pip audit в CI (S18), `"latest"` pins (S17)
- Cleanup emoji-комментариев (Q16), dead export stubs (Q10)
- Lazy import cleanup (A16), module singletons (A15)

---

## Post-P0 remediation status (2026-06-19)

После спринта стабилизации P0 ([`spec`](../../specs/stabilizaciya-p0-audit-2026-06-19.md), [`plan`](../plans/2026-06-19-stabilizaciya-p0.md)). Исходный аудит **не переписывается** — ниже только статус по P0-находкам.

> **Важно:** ID в спеке P0 и в этом аудите **различаются** для Q1/Q3. В таблице — маппинг «что исправлено» ↔ «ID в аудите».

### Закрыто в P0 (RESOLVED)

| Спека P0 | Аудит (маппинг) | Что сделано |
|----------|-----------------|-------------|
| **Q1** — тихая потеря остатков | *(нет отдельного critical ID; data integrity)* | Транзакции, structured 422/500, тесты `test_production_completion_service` |
| **Q3** — fallback `area*4000` | *(нет отдельного critical ID; не путать с audit Q3 = bot except)* | 422 `unpriced_plates`, логирование, тесты `test_commercial_pricing_errors` |
| **A2** — планы в SQLite + `version` | **A1** — два хранилища (SQLite vs JSON) | `PlanRepository`, optimistic lock, миграция, web без file I/O |
| **A1** — pipeline в `core/` | **A2** — bot не использует core pipeline | Web → `core/production/planning.py`; бот **deprecated**, не консолидируется |

### Остаётся открытым (OPEN) — на момент Post-P0

> **Обновление:** A3 и S1 закрыты в P1 — см. [Post-P1 remediation status](#post-p1-remediation-status-2026-06-19).

| ID | Severity | Примечание |
|----|----------|------------|
| ~~**A3**~~ | ~~Critical~~ | **RESOLVED в P1** |
| ~~**S1**~~ | ~~Critical~~ | **RESOLVED в P1** |

### Пересчёт Health Score (приблизительно)

| Метрика | До P0 | После P0 |
|---------|-------|----------|
| Critical (audit) | 4 (A1, A2, A3, S1) | **2** (A3, S1) |
| Закрыто critical | — | A1 (планы), A2 (pipeline / bot frozen) |
| **Overall Health Score** | **0 / 10** | **~3 / 10** |

Формула (упрощённо, пропорционально оставшимся critical):  
`10 − 3 (2 из 4 critical ≈ половина штрафа −6) − 3 (high cap, без изменений) − 1 (medium cap) ≈ 3`.

Дополнительно закрыты **data-integrity** риски спеки (Q1/Q3), не отражённые как отдельные critical в исходном аудите.

### Рекомендации следующего спринта (после P0, устарело — см. Post-P1)

1. ~~**A3** + **S5**~~ — закрыто в P1.
2. ~~**S1**~~ — закрыто в P1.
3. **S2** — rate limiting на login.
4. **S3** — object-level RBAC для КП.
5. Frontend: reload плана при **409** `plan_version_conflict` (парсинг `code` уже в `apiError.ts`, UX reload — не сделан).

---

## Post-P1 remediation status (2026-06-19)

После спринта стабилизации P1 ([`spec`](../../specs/stabilizaciya-p1-runtime-security-2026-06-19.md), [`plan`](../plans/2026-06-19-stabilizaciya-p1.md)). Исходный аудит **не переписывается** — ниже только статус по P1-находкам.

### Закрыто в P1 (RESOLVED)

| ID | Severity | Что сделано |
|----|----------|-------------|
| **A3** | Critical | Request-scoped `PlateOrderContext` на FastAPI hot paths; `optimize()` под `ctx.bound()`; инвентаризация [`plate-runtime-globals-inventory.md`](../architecture/plate-runtime-globals-inventory.md); strangler для legacy `cfg` сохранён на cold paths |
| **S1** | Critical | Fail-closed bot auth: synthetic admin только `APP_ENV=development`; staging/unknown → deny; `validate_bot_startup()` → `sys.exit(1)`; тесты `test_bot_auth.py` |
| **S5** | High | Закрыт вместе с A3: параллельные HTTP isolation-тесты; `_visualize_lock` удалён с hot path |

### Deferred из P1 (OPEN → частично закрыто в P2)

| ID | Severity | Примечание |
|----|----------|------------|
| ~~**S2**~~ | ~~High~~ | **RESOLVED в P2** — login rate limit |
| ~~**S3**~~ | ~~High~~ | **RESOLVED в P2** — object-level RBAC REST API (gaps: legacy web, bot) |
| **S4** | High | npm CVE — deferred WP3 P2, не в scope P1 |
| *(backlog)* | — | Полное удаление PEP 562 proxy `cfg.PLATES_*`; bot god-modules (A6); ~~frontend 409 reload~~ **закрыто в P2** |

### Остаточный риск (documented, не critical)

- **Cold/warm paths** (bot handlers, scripts, часть `viz_modules/`) по-прежнему могут мутировать legacy `cfg` — задокументировано в inventory; миграция — backlog.
- **A8** (два runtime-пути API vs bot) — частично смягчено middleware бота; полная унификация не цель при deprecated bot.

### Пересчёт Health Score (приблизительно)

| Метрика | До P0 | После P0 | После P1 |
|---------|-------|----------|----------|
| Critical (audit) | 4 (A1, A2, A3, S1) | 2 (A3, S1) | **0** |
| Закрыто critical | — | A1, A2 | + A3, S1 |
| High (без изменений в формуле) | 15 | 15 | 15 (−S5 mitigated, остаётся в high backlog) |
| **Overall Health Score** | **0 / 10** | **~3 / 10** | **~6 / 10** |

Формула (упрощённо):  
`10 − 0 (critical cap снят) − 3 (high cap, без изменений) − 1 (medium cap) ≈ 6`.

Дополнительно: **744 passed, 12 skipped** в `pytest tests/ -q` на closure P1.

### Рекомендации следующего спринта (P2)

1. **S2** — rate limiting на `POST /api/v1/auth/login` (5 попыток/мин → 429).
2. **S3** — object-level authorization для КП/archive endpoints.
3. **S4** — `npm audit fix`, pinned versions, CI gate.
4. **Q5** — integration-тесты production API (happy path + failure modes).
5. Frontend: reload при **409** `plan_version_conflict`.
6. **Q6** — удалить debug-instrumentation из production paths.
7. Backlog: strangler `cfg.PLATES_*`; bot decomposition (A6) при необходимости поддержки.

---

## Post-P2 remediation status (2026-06-20)

После спринта безопасности P2 ([`spec`](../../specs/bezopasnost-p2-audit-2026-06-19.md), [`plan`](../plans/2026-06-19-bezopasnost-p2.md)). Исходный аудит **не переписывается** — ниже только статус по P2-находкам.

### Закрыто в P2 (RESOLVED)

| ID / тема | Severity | Что сделано |
|-----------|----------|-------------|
| **S2** | High | In-process login rate limit: `app/security/login_rate_limit.py`; 5 POST `/api/v1/auth/login` на IP / 60 с → **429** + `Retry-After`; тесты `tests/test_auth_login_rate_limit.py` |
| **S3** | High | Object-level RBAC: `app/security/offer_access.py`; колонка `owner_user_id` в `kp_meta`; filter в `kp_repository` / `kp_archive_repository`; endpoints `offers.py`, `archive.py`; тесты `test_offers_authorization.py`, `test_archive_authorization.py` |
| **FE-409** | UX (Post-P1 gap) | `planConflict.ts` — detect `plan_version_conflict`, toast + invalidate/refetch production queries; Vitest + `npm run build` green |

### Deferred / gaps (остаётся OPEN)

| ID / тема | Severity | Примечание |
|-----------|----------|------------|
| **S4** | High | npm CVE — WP3 optional **deferred** из P2 |
| **Legacy web** | Medium | `app/web/router.py` — без `offer_access`; role-level only |
| **Bot RBAC** | Low (deprecated bot) | Нет `owner_user_id` parity; не цель P2 |
| **Q5** | High | Integration tests production API |
| **Q6** | High | Debug-instrumentation cleanup |

### Остаточный риск (documented)

- **In-process rate limit** (S2, OCR S7) — не работает между воркерами; assumption single instance.
- **Legacy КП без `owner_user_id`** — admin-only; backfill — backlog.
- **High backlog** (A4–A8, Q1–Q6, S4) — без существенного сокращения счётчика находок, но auth attack surface REST API закрыт.

### Пересчёт Health Score (приблизительно)

| Метрика | После P0 | После P1 | После P2 |
|---------|----------|----------|----------|
| Critical (audit) | 2 | **0** | **0** |
| High — S2, S3 | open | open | **RESOLVED** |
| High — S4, Q5, Q6, A4–A8… | open | open | open (S4 deferred) |
| **Overall Health Score** | **~3 / 10** | **~6 / 10** | **~7 / 10** |

Формула (упрощённо):  
`10 − 0 (critical cap снят) − 2 (high cap, частичное снятие S2/S3/S5) − 1 (medium cap) ≈ 7`.

Дополнительно: **756 passed, 12 skipped** в `pytest tests/ -q` на closure P2 (+12 тестов vs P1 baseline 744).

### Рекомендации следующего спринта (после P2)

1. **Q5** — integration-тесты production API (happy path + failure modes).
2. **Q6** — удалить debug-instrumentation из production paths.
3. **S4** — `npm audit fix`, pinned versions, CI `npm audit --audit-level=high`.
4. **A10** — deprecate legacy web UI или портировать `offer_access` на `app/web/router.py`.
5. Backlog: S10 password policy, S6 exception sanitization, strangler `cfg.PLATES_*`, bot decomposition (A6).

---

## Приложение: индекс всех находок

| ID | Severity | Категория |
|----|----------|-----------|
| A1–A3 | Critical | Архитектура |
| A4–A8 | High | Архитектура |
| A9–A14 | Medium | Архитектура |
| A15–A18 | Low | Архитектура |
| S1 | Critical | Безопасность |
| S2–S5 | High | Безопасность |
| S6–S12 | Medium | Безопасность |
| S13–S18 | Low | Безопасность |
| — | Critical | Качество кода (нет) |
| Q1–Q6 | High | Качество кода |
| Q7–Q14 | Medium | Качество кода |
| Q15–Q19 | Low | Качество кода |

---

*Отчёт сформирован: 2026-06-19 · Консолидирован из проходов senior-reviewer, security-auditor и reviewer.*
