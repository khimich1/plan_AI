# Отчёт полного аудита проекта

**Дата:** 2026-06-03  
**Область:** Полный проект — критичные пути `app/`, `bot/`, `core/`, `tests/` (вспомогательно: `viz_modules/`)  
**Проверено:** senior-reviewer + security-auditor + reviewer  
**Статус:** Только отчёт — remediation **не** запускался в рамках данного документа  

---

## Краткое резюме (Executive Summary)

**Общий Health Score: 0.0 / 10**

Формула: `10 − min(3×2, 6) − min(17×0.5, 3) − min(26×0.1, 1) = 0` (округление до 1 знака, нижняя граница 0).

| Серьёзность | Архитектура | Безопасность | Качество кода | **Итого** |
|-------------|-------------|--------------|---------------|-----------|
| Critical    | 3           | 0            | 0             | **3**     |
| High        | 6           | 5            | 6             | **17**    |
| Medium      | 6           | 9            | 11            | **26**    |
| Low         | 3           | 5*           | 9             | **17**    |
| **Всего**   | **18**      | **19**       | **26**        | **63**    |

\* В подсчёте Low по безопасности **не** учитывается **[S-L6]** (положительные контроли); в матрице и тексте S-L6 вынесен в раздел «Положительные находки».

**Рекомендация:** Перед релизом закрыть **3 критические архитектурные** находки **[A1]–[A3]**. В первом спринте после критики — пакет **[S-H1]–[S-H3]** (auth, CSRF, IDOR) и пробелы тестов **[Q9]**. Для структурных изменений — `/refactor [путь]`; для точечных security/feature — `/implement [описание]`; для крупных эпиков — `/orchestrate`.

> **Контекст remediation (2026-06-03):** Часть рисков снижена в orchestration `orch-2026-06-03-arch-triage` и `orch-2026-06-03-bot-auth-s1` (см. [orchestration-arch-triage-2026-06-03.md](../reports/orchestration-arch-triage-2026-06-03.md)). Настоящий отчёт фиксирует **остаточное** состояние кодовой базы после этих работ; критические **[A1]–[A3]** остаются открытыми до завершения миграции.

---

## Критические проблемы (исправить немедленно)

### [A1] Остаточное thread-local / legacy `OPT_*` и глобальное runtime-состояние

**Категория:** Архитектура  
**Расположение:** `core/plate_order_context.py`, `core/plate_runtime_state.py`, `core/optimization/context.py`, `core/config_and_data.py`, `viz_modules/`, сервисы `app/services/optimization_service.py`, `day_documents_service.py`, `archive_service.py`  
**Статус remediation:** **PARTIAL** (middleware + `PlateOrderContext`, deprecate `apply_to_globals`)  

**Влияние (Impact):** Несмотря на введение request-scoped контекста, часть путей по-прежнему читает или мутирует thread-local, ContextVar и legacy-глобали (`OPT_*`, списки в `config_and_data`). При пропуске `ctx.bound()`, работе в пуле потоков без `run_in_order_context`, фоновых задачах или параллельной генерации документов заказы разных пользователей могут смешиваться; оптимизация и визуализация получают чужие данные. Горизонтальное масштабирование и надёжные интеграционные тесты остаются затруднёнными.  

**Исправление (Fix):** Полная инвентаризация обращений к `get_plate_mutable_runtime()`, `apply_to_globals`, `OPT_*` и глобальным спискам; запрет неявных чтений вне активного контекста (assert/lint); миграция `day_documents_service`, `archive_service`, `viz_modules` на явный `PlateOrderContext`; удаление дублирующих прокси после покрытия concurrency-тестами.  
**Команда:** `/refactor core/plate_order_context.py core/optimization viz_modules app/services`

---

### [A2] God-модуль `kp_db.py` (~3800+ строк)

**Категория:** Архитектура  
**Расположение:** `core/kp_db.py`, тонкие прокси `app/repositories/*`  
**Статус remediation:** **PARTIAL** (выделен `core/kp_db_nomenclature.py`, shim re-export)  

**Влияние (Impact):** В одном модуле смешаны CRUD коммерческих предложений, номенклатура, списание/возврат плит, логистика, статусы, отладочные записи и вспомогательная бизнес-логика. Любое изменение затрагивает несвязанные домены; высокий риск регрессий; невозможность изолированного тестирования слоёв. Facade-репозитории не создают реальной границы persistence.  

**Исправление (Fix):** Поэтапное разбиение на `repositories/kp.py`, `nomenclature.py`, `plates.py`, `logistics.py` с явными интерфейсами; вынесение SQL; сервисный слой только для оркестрации; `kp_db` как deprecated shim на переходный период.  
**Команда:** `/refactor core/kp_db.py`

---

### [A3] Бот обходит слой приложения (`app/services`)

**Категория:** Архитектура  
**Расположение:** `bot/handlers/*` — прямые вызовы `core.kp_db`, `config_and_data`, локальная логика планирования  
**Статус remediation:** **PARTIAL** (пилоты: `commercial.py` → `CommercialService`, `production_execution.py` → `ProductionPlanningService`)  

**Влияние (Impact):** Правила (даты исполнения, переход в производство, генерация документов) расходятся между Telegram и веб/API. Исправление бага в `app/` не гарантирует исправление в боте. Усиливается связность `bot → core` и обход единой политики ошибок и авторизации приложения.  

**Исправление (Fix):** Сценарии в `CommercialWorkflowService`, `ProductionPlanningService`, `DayDocumentsService`; handlers — только парсинг Telegram, DI контекста, вызов сервиса, форматирование ответа; общие DTO в `app/schemas` / `core/domain`.  
**Команда:** `/refactor bot/handlers`

---

## Высокий приоритет

### Архитектура

#### [A4] `PlateOrderContext` не универсален в API-слое

**Расположение:** `app/dependencies/plate_context.py`, `app/middleware/plate_runtime_isolation.py`, endpoints и сервисы без `Depends(get_plate_order_context)`  
**Влияние:** Middleware создаёт контекст, но многие endpoints и сервисы продолжают legacy-путь через глобали; рассинхрон `request.state` и фактического состояния оптимизатора.  
**Исправление:** Обязательный `PlateOrderContext` в критичных endpoints и сервисах; передача `ctx` по стеку; тесты изоляции двух параллельных HTTP-запросов.

#### [A5] `get_current_user` загружает всех пользователей (`list_users`)

**Расположение:** `app/dependencies/auth.py`, `app/repositories/auth_repository.py`  
**Влияние:** O(n) на каждый аутентифицированный запрос; деградация latency; избыточная PII в памяти (см. **[S-M3]**).  
**Исправление:** `get_user_by_id(session_user_id)`; кэш на уровне запроса через Depends.

#### [A6] Несогласованный dependency injection в API

**Расположение:** `app/api/v1/endpoints/*` — прямое `AdminService()`, `AuthRepository()` в handlers  
**Влияние:** Сложно мокать в тестах; дублирование инициализации; скрытые пути к БД.  
**Исправление:** Фабрики и `Depends` в `app/dependencies/`; запрет `Service()` в теле handler.

#### [A7] Синхронные CPU-heavy handlers блокируют asyncio event loop

**Расположение:** `app/api/v1/endpoints/production.py`, `commercial.py`, `archive.py` и связанные сервисы  
**Влияние:** Долгие sync-вычисления в `async def` блокируют весь loop — таймауты, «зависание» health-check.  
**Исправление:** `asyncio.to_thread` / executor для CPU-bound; очередь для PDF/XLSX.

#### [A8] God-модуль `plan_manager` + star import в боте

**Расположение:** `app/planning/plan_manager.py`, `bot/handlers/plan_manager.py` (или аналог)  
**Влияние:** Планирование, Gantt, файловое хранение и метаданные в одном модуле; домен «план» не отделён от persistence.  
**Исправление:** `PlanRepository`, `GanttService`, `CalendarService`; тонкий фасад; bot только через сервисы app.

#### [A9] Модульные in-memory кэши бота (`ORDER_CACHE`, `OPT_PLAN_CACHE`)

**Расположение:** `bot/handlers/*`  
**Влияние:** Состояние привязано к процессу; рост памяти без eviction; рассинхрон при нескольких инстансах (см. **[S-M2]**).  
**Исправление:** FSM aiogram + `plate_order_ctx`, Redis с TTL или явный state в callback data.

---

### Безопасность

#### [S-H1] Нет rate limiting на login

**Расположение:** `app/api/v1/endpoints/auth.py`, `app/web/router.py`  
**Влияние:** Brute-force по паролю; DoS на login endpoint.  
**Исправление:** Лимит по IP/username (slowapi, Redis, nginx); экспоненциальная задержка; единая политика API + web-form.

#### [S-H2] Cookie-authenticated web-формы без CSRF

**Расположение:** `app/web/router.py`, state-changing POST  
**Влияние:** Authenticated POST от имени жертвы с вредоносной страницы.  
**Исправление:** CSRF token в формах; проверка на сервере; `SameSite=strict` где возможно.

#### [S-H3] IDOR: offers / archive для роли manager

**Расположение:** `app/api/v1/endpoints` offers/archive, `offers_service.py`, `archive_service.py`  
**Влияние:** Manager может читать/менять чужие КП без фильтра по `manager_id`.  
**Исправление:** Row-level фильтрация в repository; policy object; тесты «manager A ≠ KP manager B».

#### [S-H4] Секреты бота на диске (`bot/bot.env`)

**Расположение:** `bot/bot.env`, `bot/bot_main.py`, загрузка токена  
**Влияние:** Утечка токена Telegram при компрометации диска/бэкапа/репозитория; риск попадания в VCS.  
**Исправление:** Только env vars / secret manager; `bot.env` в `.gitignore`; ротация токена; отдельные токены prod/stage.

#### [S-H5] Allowlist бота — MVP с ограниченной моделью угроз

**Расположение:** `bot/middleware/auth.py`, `core/config/settings.py` — `BOT_TELEGRAM_ALLOWLIST`  
**Влияние:** Статический allowlist без полноценного lifecycle пользователей, revoke, audit UI; при пустом allowlist в non-prod — открытый доступ; нет привязки к корпоративной IAM.  
**Исправление:** Хранение пользователей в БД; роли и отзыв доступа; обязательный непустой allowlist в production; аудит действий (расширить `bot/security/audit.py`).

---

### Качество кода

#### [Q1] Несогласованная типизация `load_code` в ключах dict заказа плит

**Расположение:** `core/domain/plate_order.py`  
**Влияние:** Ключи `(load_code, length, width)` то `int`, то `str` — промахи lookup, дубли позиций.  
**Исправление:** `LoadCode` как NewType/enum; единая `plate_key()`; тесты round-trip.

#### [Q2] Голый `except` в production bot handler

**Расположение:** `bot/handlers/production_execution.py`  
**Влияние:** КП исключаются из плана без логирования; неполный план в Telegram.  
**Исправление:** Конкретные исключения; лог `kp_id`; сводка пропущенных КП пользователю.

#### [Q3] God-функция `load_and_plan_production` (~1430 LOC)

**Расположение:** `bot/handlers/production_execution.py` (и вынесенные helpers)  
**Влияние:** Нет unit-тестов шагов; любой diff рискован.  
**Исправление:** Pipeline: load KPs → build order → optimize → persist → notify; шаги &lt; 80 строк.

#### [Q4] Мега-handlers commercial / completion (300–575+ строк)

**Расположение:** `bot/handlers/commercial.py`, `production_create.py`, `production_day_view.py`, связанные completion paths  
**Влияние:** Смешение UI Telegram и бизнес-логики; дубли внутри файла.  
**Исправление:** Подпакеты `bot/workflows/`, `bot/ui/`; тонкие handlers.

#### [Q9] Пробелы тестов: production API, offers, day documents, bot flows

**Расположение:** `tests/` — отсутствие или слабое покрытие `bot/handlers/*`, production/offers endpoints, day docs  
**Влияние:** Регрессии при рефакторинге **[A1]–[A3]** не ловятся CI.  
**Исправление:** pytest для сервисов; `TestClient` для API; 3–5 сценарных тестов главных команд бота.

#### [Q10] String-based dispatch ошибок в offers

**Расположение:** `app/api/v1/endpoints` (offers), `offers_service.py` — сравнение `str(exc)` с константами  
**Влияние:** Хрупкие ветвления; непредсказуемые HTTP-коды при смене текста ошибки.  
**Исправление:** Доменные исключения `OfferNotFound`, `OfferAccessDenied`; mapping в exception handler.

---

## Средний приоритет

### Архитектура (кратко)

| ID | Проблема | Расположение | Исправление (суть) |
|----|----------|--------------|-------------------|
| **[A10]** | Facade-репозитории без абстракции | `app/repositories/*` → `kp_db` | Protocol + Depends; замена kp_db [A2] |
| **[A11]** | Legacy `config_and_data` на hot path | `core/config_and_data.py`, bot, optimization | Deprecation; redirect на `PlateOrder` + ctx |
| **[A12]** | `viz_modules` и глобальный `OPT_PLAN` / lock | `viz_modules/*`, `file_generation_service.py` | Lock/snapshot на уровне `plate_order_ctx` |
| **[A13]** | Монолит `web/router.py` | `app/web/router.py` (~900+ строк) | Тонкие routes; логика в сервисах |
| **[A14]** | FS storage без multi-instance | `plan_manager`, settings `APP_STORAGE_LAYOUT` | Документировать single-instance или DB + locks |
| **[A15]** | Две модели `PlateOrder` (dual model) | `app/domain/models`, `core/domain`, adapters | SSOT в core; минимизировать app-only поля; ADR |

### Безопасность (кратко)

| ID | Проблема | Расположение | Исправление (суть) |
|----|----------|--------------|-------------------|
| **[S-M1]** | Debug NDJSON с бизнес-данными | `core/kp_db.py`, bot handlers, services | Удалить agent debug; structured logger |
| **[S-M2]** | Кэши бота без TTL/границ | `ORDER_CACHE`, `OPT_PLAN_CACHE` | См. [A9]; TTL, max size |
| **[S-M3]** | `list_users` на каждый auth-запрос | `app/dependencies/auth.py` | См. [A5] |
| **[S-M4]** | Нет server-side revoke сессий | `app/security/session.py` | Session store или rotation; runbook ротации `APP_SECRET_KEY` |
| **[S-M5]** | OCR rate limit только in-process | `commercial_upload_validation` / OCR pipeline | Redis/DB счётчик; лимит user_id + IP |
| **[S-M6]** | Утечка внутренних ошибок | `app/api/v1/endpoints/production.py`, `http_errors.py` | Стабильные коды клиенту; traceback только в log |
| **[S-M7]** | Telegram HTML без экранирования | bot handlers с `parse_mode=HTML` | `html.escape` для пользовательского ввода |
| **[S-M8]** | Dev bypass admin в боте | `bot/middleware/auth.py`, settings | Жёсткий запрет bypass при `APP_ENV=production` |
| **[S-M9]** | Остаточные пути `OPT_*` / globals | `core/optimization`, `config_and_data` | См. [A1]; lint/grep gate в CI |

### Качество кода (кратко)

| ID | Проблема | Расположение |
|----|----------|--------------|
| **[Q5]** | Дублированная фильтрация/загрузка КП | bot vs `production_planning_service`, `kp_db` |
| **[Q6]** | Трипликация day documents | `day_documents_service` vs `production_day_view` |
| **[Q7]** | Дубли «перевод в производство» / execution terms | bot vs `commercial_workflow_service` |
| **[Q8]** | Дубли mapping KP-plate → `order_data` | bot, `commercial_service`, `core/domain` |
| **[Q15]** | Дубли PDF/XLSX в `OffersService` | `app/services/offers_service.py` |
| **[Q16]** | Agent debug NDJSON в services | `app/services/*`, `core/kp_db.py` |
| **[Q17]** | Несогласованные HTTP-ошибки | endpoints vs `http_errors.py` |
| **[Q18]** | Бот обходит `CommercialWorkflowService` | `bot/handlers/commercial.py` |
| **[Q19]** | Смешанные типы plate (dict vs model) | handlers, services |
| **[Q20]** | Контракты `dict[str, Any]` | публичные API сервисов/repositories |
| **[Q21]** | Silent rest failure | production/bot paths без явного уведомления |

---

## Низкий приоритет / предложения

### Архитектура

- **[A16]** Устаревшие legacy shims (`config_and_data`, re-exports, `apply_to_globals`) — sunset и удаление по метрикам grep.
- **[A17]** Ad-hoc `sys.path` в `bot/bot_main.py` — `python -m bot.bot_main`, editable install в `pyproject.toml`.
- **[A18]** Debug-инструментация в domain-модулях — вынести в `core/debug_utils.py` с no-op по умолчанию.

### Безопасность (issues; S-L6 — в положительных находках)

- **[S-L1]** Отсутствуют HTTP security headers (HSTS, CSP, `X-Content-Type-Options`, `X-Frame-Options`) — middleware в `app/main.py`.
- **[S-L2]** `app_debug` / FastAPI `debug=settings.app_debug` — жёстко `False` при `APP_ENV=production`.
- **[S-L3]** Admin reset — большой blast radius (сброс паролей/данных без подтверждения) — двухэтапное подтверждение, audit log.
- **[S-L4]** Нет `pip-audit` / Dependabot в CI — автоматическое сканирование зависимостей.
- **[S-L5]** SQLite и draft/plan files без encryption at rest — ops: том FS, SQLCipher, права доступа.

### Качество кода

- **[Q11]** Неиспользуемый import в `production_execution.py`.
- **[Q12]** `user: dict` без типа — `UserPrincipal` / TypedDict.
- **[Q13]** Импорт приватных `_helper` через границы пакетов.
- **[Q14]** Нет ruff/mypy в CI — постепенный strict на `app/`, `core/domain`.
- **[Q22]** Дублированный logger в модуле.
- **[Q23]** TODO в export paths без трекинга.
- **[Q24]** Схемы «pass-through» без валидации полей.
- **[Q25]** Комментарии agent log / временная инструментация в production-коде.

---

## Положительные находки (S-L6 и недавний remediation)

### [S-L6] Положительные security-контроли (уже в кодовой базе)

| Контроль | Расположение | Описание |
|----------|--------------|----------|
| HMAC-сессии | `app/security/session.py` | Подписанные cookie; централизованная политика `secure`/`httponly`/`samesite` |
| PBKDF2 пароли | `app/repositories/auth_repository.py` (или эквивалент) | Хеширование паролей, не plaintext |
| Валидация `APP_SECRET_KEY` | `core/config/settings.py` | Min length, запрет placeholder, fail-fast при старте |
| Валидация draft/plate upload | commercial upload validation | Ограничения на загрузку черновиков (см. тесты commercial) |
| Единые HTTP-ошибки (commercial) | `app/core/http_errors.py` | Стабильные коды/сообщения для части commercial API |

### Недавний remediation (2026-06-03, не снимает Critical A1–A3)

| Область | Что сделано | Документация / тесты |
|---------|-------------|----------------------|
| **PlateOrderContext** | Request/update-scoped контекст; middleware FastAPI + aiogram; `run_in_order_context` | [plate-order-context-a1-001](../features/plate-order-context-a1-001-phase-1.md), [a1-002](../features/plate-order-context-a1-002-middleware-deprecations.md); `tests/test_plate_order_context.py` |
| **Session cookies** | `APP_SECRET_KEY` validator; `set_session_cookie` / policy по `APP_ENV` | [secure-session-cookies-a2-001](../features/secure-session-cookies-a2-001.md); `tests/test_app_session.py`, `test_settings_app_secret_key.py` |
| **Canonical PlateOrder** | SSOT `core/domain/plate_order.py`; adapters app↔core | [plate-order-canonical-a3-001](../features/plate-order-canonical-a3-001.md), [a3-002](../features/plate-order-migration-a3-002.md); `tests/test_plate_order_adapters.py` |
| **Bot auth MVP** | `bot/middleware/auth.py`, `role.py`; `BOT_TELEGRAM_ALLOWLIST`; audit | `orch-2026-06-03-bot-auth-s1`; `tests/test_bot_auth.py` |
| **kp_db срез** | `core/kp_db_nomenclature.py` вынесен из монолита | Первый шаг [A2] |
| **Пилот bot→app** | Часть handlers на `CommercialService` / `ProductionPlanningService` | Первый шаг [A3] |

**Ориентир тестов после remediation:** targeted/arch-triage — **136+ passed**; bot auth + session — **36/36** (по отчёту orchestration).

---

## Матрица приоритетов

| ID | Проблема | Серьёзность | Усилия | Приоритет |
|----|----------|-------------|--------|-----------|
| A1 | Thread-local / legacy OPT_* | Critical | High | **P0 — немедленно** |
| A2 | God-модуль kp_db ~3800 LOC | Critical | High | **P0 — немедленно** |
| A3 | Бот обходит app layer | Critical | High | **P0 — немедленно** |
| S-H1 | Нет rate limit на login | High | Medium | P1 — этот спринт |
| S-H2 | CSRF web forms | High | Medium | P1 |
| S-H3 | IDOR offers/archive | High | Medium | P1 |
| S-H4 | bot.env секреты на диске | High | Low | P1 |
| S-H5 | Allowlist MVP (ограничения) | High | Medium | P1 |
| A4 | PlateOrderContext не везде в API | High | Medium | P1 |
| A5 | get_current_user list_users | High | Low | P1 |
| A6 | Нет единого DI | High | Medium | P1 |
| A7 | Sync CPU блокирует loop | High | Medium | P1 |
| A8 | God plan_manager | High | High | P1 |
| A9 | ORDER_CACHE / OPT_PLAN_CACHE | High | Medium | P1 |
| Q1 | load_code typing | High | Medium | P1 |
| Q2 | Bare except в bot | High | Low | P1 |
| Q3 | load_and_plan_production ~1430 LOC | High | High | P1 |
| Q4 | Мега-handlers commercial/completion | High | High | P1 |
| Q9 | Пробелы тестов production/offers/day_docs | High | High | P1 |
| Q10 | String errors offers_service | High | Low | P1 |
| S-M1 | Debug NDJSON kp_db | Medium | Medium | P2 |
| S-M2 | Bot caches | Medium | Medium | P2 |
| S-M3 | list_users (дубль A5) | Medium | Low | P2 |
| S-M4 | Нет revoke сессий | Medium | Medium | P2 |
| S-M5 | OCR rate limit in-process | Medium | Medium | P2 |
| S-M6 | Error leakage production.py | Medium | Low | P2 |
| S-M7 | Telegram HTML escaping | Medium | Low | P2 |
| S-M8 | Bot dev admin bypass | Medium | Low | P2 |
| S-M9 | Residual OPT_* paths | Medium | Medium | P2 |
| A10 | Facade repositories | Medium | Medium | P2 |
| A11 | config_and_data legacy | Medium | Medium | P2 |
| A12 | viz_modules OPT_PLAN lock | Medium | Medium | P2 |
| A13 | web/router monolith | Medium | High | P2 |
| A14 | FS multi-instance | Medium | High | P2 |
| A15 | Dual PlateOrder models | Medium | Medium | P2 |
| Q5 | Дубли KP filter | Medium | Medium | P2 |
| Q6 | day_documents triplication | Medium | Medium | P2 |
| Q7 | execution terms dup | Medium | Medium | P2 |
| Q8 | order_data mapping dup | Medium | Medium | P2 |
| Q15 | offers PDF/XLSX dup | Medium | Medium | P2 |
| Q16 | agent debug NDJSON services | Medium | Medium | P2 |
| Q17 | inconsistent http errors | Medium | Low | P2 |
| Q18 | bot bypasses CommercialWorkflowService | Medium | Medium | P2 |
| Q19 | mixed plate types | Medium | Medium | P2 |
| Q20 | dict[str, Any] contracts | Medium | Medium | P2 |
| Q21 | silent rest failure | Medium | Low | P2 |
| A16 | Deprecated shims | Low | Low | P3 — бэклог |
| A17 | sys.path bot_main | Low | Low | P3 |
| A18 | debug in domain | Low | Medium | P3 |
| S-L1 | Security headers | Low | Low | P3 |
| S-L2 | app_debug production | Low | Low | P3 |
| S-L3 | admin reset blast radius | Low | Medium | P3 |
| S-L4 | no pip audit CI | Low | Low | P3 |
| S-L5 | SQLite FS encryption | Low | High | P3 |
| Q11 | unused import | Low | Low | P3 |
| Q12 | user dict untyped | Low | Medium | P3 |
| Q13 | private imports | Low | Medium | P3 |
| Q14 | no ruff/mypy | Low | High | P3 |
| Q22 | duplicate logger | Low | Low | P3 |
| Q23 | export TODO | Low | Low | P3 |
| Q24 | schema pass | Low | Medium | P3 |
| Q25 | agent log comments | Low | Low | P3 |
| S-L6 | Положительные контроли | — | — | Уже реализовано |

---

## Следующие шаги (Next Steps)

### 1. Немедленно (до production-релиза)

1. **`/refactor`** — завершить миграцию с thread-local/`OPT_*` на обязательный `PlateOrderContext` на всех entry points (**[A1]**, **[S-M9]**).
2. **`/refactor`** — продолжить декомпозицию `core/kp_db.py` (CRUD KP, plates, logistics) (**[A2]**).
3. **`/refactor`** — унифицировать bot handlers через `app/services/*` для всех commercial/production сценариев (**[A3]**, **[Q18]**).
4. **`/refactor`** — удалить debug NDJSON с бизнес-данными (**[S-M1]**, **[Q16]**).

### 2. Этот спринт

- **`/implement`** — rate limiting login (**[S-H1]**); CSRF для web (**[S-H2]**); row-level access manager (**[S-H3]**).
- **`/implement`** — `get_user_by_id` (**[A5]**, **[S-M3]**); безопасные ответы API (**[S-M6]**); вынести секреты из `bot.env` (**[S-H4]**).
- **`/refactor`** — DI endpoints (**[A6]**); `asyncio.to_thread` для тяжёлых API (**[A7]**); потребление `PlateOrderContext` (**[A4]**).
- **`/refactor`** — bot caches → FSM/ctx (**[A9]**, **[S-M2]**); bare `except` (**[Q2]**); `load_code` normalization (**[Q1]**).
- **`/implement`** — тесты production API, offers, day documents, ключевые bot commands (**[Q9]**).
- Усилить модель auth бота поверх allowlist (**[S-H5]**).

### 3. Следующий спринт

- **`/refactor`** — `plan_manager` (**[A8]**); `web/router.py` (**[A13]**); унификация KP/day docs/move-to-production (**[Q5]–[Q8]**).
- **`/implement`** — distributed OCR limits (**[S-M5]**); session hardening/revoke (**[S-M4]**); Telegram HTML escape (**[S-M7]**).
- Стратегия multi-instance storage (**[A14]**); разбиение `load_and_plan_production` (**[Q3]**).

### 4. Бэклог

- Legacy shims (**[A16]**); package entry бота (**[A17]**); ruff/mypy (**[Q14]**); `pip-audit` CI (**[S-L4]**).
- Security headers (**[S-L1]**); production guards debug (**[S-L2]**); encryption at rest ops (**[S-L5]**).

**Команды:** `/refactor [путь]` — структура; `/implement [описание]` — security/поведение; `/orchestrate` — эпики A1–A3 + kp_db.

---

## Связанная документация

- Workflow: `.cursor/skills/audit-workflow/SKILL.md`
- Конфиг путей: `.cursor/config.json` → `ai_docs/develop/audits`
- Remediation reports: [orchestration-arch-triage-2026-06-03.md](../reports/orchestration-arch-triage-2026-06-03.md)
- Features: `ai_docs/develop/features/plate-order-*.md`, `secure-session-cookies-a2-001.md`
- Plan arch-triage: [2026-06-03-architecture-triage-a1-a2-a3.md](../plans/2026-06-03-architecture-triage-a1-a2-a3.md)

---

*Консолидированный отчёт: Architecture (senior-reviewer), Security (security-auditor), Code Quality (reviewer). ID: A1–A18, S-H/M/L, Q1–Q25. Health Score по предрасчёту координатора аудита. Remediation в этом файле не выполнялся.*
