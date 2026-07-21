# Отчёт аудита проекта

**Дата**: 2026-06-10  
**Область**: Полный проект (Telegram-бот, FastAPI API, React frontend, core-модули, Docker-инфраструктура)  
**Аудиторы**: senior-reviewer + security-auditor + reviewer

---

## Краткое резюме

**Health Score: 2.0 / 10**

Формула: `10 − (2 critical × 2) − min(13 high × 0.5, 3) − min(24 medium × 0.1, 1) = 2.0`

| Категория | Critical | High | Medium | Low |
|-----------|----------|------|--------|-----|
| Архитектура (A*) | 1 | 4 | 4 | 4 |
| Безопасность (S*) | 1 | 4 | 9 | 5 |
| Качество кода (Q*) | 0 | 5 | 11 | 5 |
| **Итого** | **2** | **13** | **24** | **14** |

Проект функционально зрелый: есть разделение на слои, тесты на критичных модулях, защита от деструктивных операций с БД. Однако **две критические проблемы** — раздельное хранение планов без транзакционной границы (A1) и дефолтные учётные данные admin/admin123 в seed-БД (S1) — создают риск потери данных и немедленной компрометации при деплое.

**Рекомендация**: до любого production-релиза устранить A1 и S1, затем в течение одного спринта закрыть 13 findings высокого приоритета (архитектурная консолидация, auth/rate limiting/CSRF, удаление debug-инструментации, покрытие тестами ключевых API).

---

## Критические проблемы (исправить немедленно)

### [A1] Раздельная персистентность без транзакционной границы

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Расположение** | `bot/data/plans/` (JSON-планы), `plita.db` (SQLite), `app/planning/plan_manager.py` (~1488 строк), `core/plan_commit.py` |
| **Влияние** | План производства может оказаться в несогласованном состоянии: часть данных в JSON-файлах, часть в SQLite. При сбое между записью в файл и коммитом в БД — потеря или дублирование планов, «зависшие» плиты, невозможность восстановить единый источник истины. |
| **Исправление** | Ввести единый порт персистентности (repository interface) с атомарными транзакциями. Все операции commit/rollback плана должны проходить через один слой; JSON и SQLite — либо миграция в одно хранилище, либо двухфазный commit с компенсирующими действиями. |

---

### [S1] Дефолтные учётные данные admin/admin123 в seed-БД

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Расположение** | `docker/seed/plita.db` (файл в git), копируется при первом Docker-деплое |
| **Влияние** | Любой, кто знает стандартный seed, получает полный административный доступ к системе сразу после развёртывания. Критический риск компрометации в production. |
| **Исправление** | Удалить seed с предзаполненным admin из репозитория. Инициализацию администратора выполнять через `scripts/create_admin.py` при первом запуске. Ротировать все существующие credentials на уже развёрнутых инстансах. |

---

## Высокий приоритет

### Архитектура

#### [A2] Тройной слой представления

| Поле | Значение |
|------|----------|
| **Расположение** | `bot/handlers/commercial.py`, `bot/handlers/production_execution.py`, `app/services/*`, `app/web/router.py` |
| **Влияние** | Одна и та же бизнес-логика реализована в трёх местах. Изменение правила расчёта или планирования требует правок в нескольких файлах; высокий риск расхождения поведения бота, API и веб-интерфейса. |
| **Исправление** | Вынести use-cases в `core/` (или `app/domain/`). Handlers, services и router должны только оркестрировать вызовы общих use-case-функций. |

#### [A3] Auth загружает всех пользователей на каждый запрос

| Поле | Значение |
|------|----------|
| **Расположение** | `app/dependencies/auth.py`, `app/repositories/auth_repository.py` (вызов `init_schema` при чтении) |
| **Влияние** | O(n) по числу пользователей на каждый аутентифицированный запрос. При росте базы — деградация latency и потенциальный вектор DoS. Побочный init схемы при read-операциях нарушает принцип разделения ответственности. |
| **Исправление** | Добавить `get_user_by_id(session_user_id)` с точечным SELECT. Инициализацию схемы перенести в startup/lifespan приложения. |

#### [A4] God-модули

| Поле | Значение |
|------|----------|
| **Расположение** | `bot/handlers/commercial.py`, `app/planning/plan_manager.py`, `viz_modules/visualization.py`, `core/kp_db_offers.py`, `app/services/production_completion_service.py` |
| **Влияние** | Файлы на 1500–2100+ строк смешивают UI-оркестрацию, бизнес-правила, I/O и форматирование. Сложно тестировать, ревьюить и безопасно рефакторить. |
| **Исправление** | Разбить по ответственности: handlers → thin adapters, services → orchestration, core → pure business logic, repositories → data access. |

#### [A5] Bot импортирует app-слой — инвертированная зависимость

| Поле | Значение |
|------|----------|
| **Расположение** | Импорты из `app/` внутри `bot/` |
| **Влияние** | Нарушение направления зависимостей: инфраструктурный слой (бот) зависит от application layer. Циклические импорты, сложность изолированного тестирования бота. |
| **Исправление** | Перенести оркестрацию в `core/`. Bot и FastAPI должны зависеть только от core + своих адаптеров. |

---

### Безопасность

#### [S2] Нет rate limiting на auth-эндпоинтах

| Поле | Значение |
|------|----------|
| **Расположение** | `app/api/v1/endpoints/auth.py`, `app/web/router.py` (login) |
| **Влияние** | Неограниченные попытки входа — brute-force по паролям и username enumeration. |
| **Исправление** | Добавить rate limiting (slowapi, nginx limit_req, или Redis-based) на `/login` и связанные auth-маршруты. |

#### [S3] Cookie-сессия без CSRF на state-changing операциях

| Поле | Значение |
|------|----------|
| **Расположение** | Session-based auth по cookie, POST/PUT/DELETE без CSRF-токена |
| **Влияние** | Атакующий может выполнить действия от имени авторизованного пользователя через подставленную форму или cross-site запрос. |
| **Исправление** | Double-submit cookie, synchronizer token, или переход на SameSite=Strict + проверка Origin/Referer для всех мутирующих операций. |

#### [S4] npm audit: уязвимости во frontend-зависимостях

| Поле | Значение |
|------|----------|
| **Расположение** | `frontend/package.json` — react-router (RCE/DoS, high), postcss (moderate) |
| **Влияние** | Известные CVE в цепочке поставок frontend. RCE/DoS через react-router при определённых условиях эксплуатации. |
| **Исправление** | `npm audit fix`, обновить react-router и postcss до patched-версий, добавить `npm audit` в CI. |

#### [S5] Сессии валидны после смены пароля/роли

| Поле | Значение |
|------|----------|
| **Расположение** | Механизм сессий без revocation store |
| **Влияние** | После смены пароля или понижения роли старая сессия остаётся активной до истечения cookie. |
| **Исправление** | Версионирование сессий (session_version в user record), blacklist/invalidate при password change и role change. |

---

### Качество кода

#### [Q1] Остаточная agent-debug инструментация в production-путях

| Поле | Значение |
|------|----------|
| **Расположение** | Блоки `#region agent log` в production-коде |
| **Влияние** | Лишние I/O, утечка внутреннего контекста в логи, шум при отладке production-инцидентов. |
| **Исправление** | Удалить все `#region agent log` блоки. Использовать структурированное логирование через `core/logging_config.py`. |

#### [Q2] Bare except / except: pass глотает ошибки

| Поле | Значение |
|------|----------|
| **Расположение** | `bot/handlers/*`, `app/planning/plan_manager.py` |
| **Влияние** | Сбои остаются незамеченными; пользователь видит молчаливый fail или некорректное состояние без traceback в логах. |
| **Исправление** | Ловить конкретные исключения, логировать с `exc_info=True`, пробрасывать или возвращать понятную ошибку. |

#### [Q3] Ошибки валидации логируются, клиенту — generic message

| Поле | Значение |
|------|----------|
| **Расположение** | `app/core/http_errors.py` (или аналог) |
| **Влияние** | Frontend не может показать поле с ошибкой; разработчик видит детали только в логах. Ухудшенный UX и дольше отладка интеграции. |
| **Исправление** | Возвращать структурированный 422 с `detail` по полям (стандарт FastAPI/Pydantic), без утечки внутренних stack trace. |

#### [Q4] Экстремальная сложность модулей

| Поле | Значение |
|------|----------|
| **Расположение** | `bot/handlers/commercial.py` (~2124 строк), `app/services/production_completion_service.py` (~1608), `build_plan` (~275 строк в одной функции) |
| **Влияние** | Высокая когнитивная нагрузка, низкая тестируемость, риск регрессий при любом изменении. |
| **Исправление** | Декомпозиция на функции < 50 строк, выделение чистых helper'ов, покрытие unit-тестами после разбиения. |

#### [Q5] Отсутствуют тесты для ключевых API-модулей

| Поле | Значение |
|------|----------|
| **Расположение** | `app/api/v1/endpoints/offers.py`, `production.py`, `app/services/day_documents_service.py`, `app/services/offers_service.py` |
| **Влияние** | Регрессии в коммерческих предложениях и производстве не обнаруживаются до ручного тестирования. |
| **Исправление** | Добавить integration-тесты через `TestClient` и unit-тесты для service-слоя с mock repository. |

---

## Средний приоритет

### Архитектура

#### [A6] Непоследовательный repository pattern, raw sqlite3

| Поле | Значение |
|------|----------|
| **Расположение** | Разные модули в `app/repositories/`, `core/kp_db_*.py` |
| **Влияние** | Смешение стилей доступа к данным усложняет миграцию на async/pooling и единую обработку ошибок. |
| **Исправление** | Унифицировать интерфейсы repository, централизовать connection management. |

#### [A7] Legacy global mutable state

| Поле | Значение |
|------|----------|
| **Расположение** | `core/config_and_data.py` |
| **Влияние** | Глобальное состояние создаёт race conditions при concurrent-запросах (FastAPI workers) и усложняет тесты. |
| **Исправление** | Заменить на `PlateOrderContext` / request-scoped binding (паттерн уже частично внедрён). |

#### [A8] Непоследовательный FastAPI DI

| Поле | Значение |
|------|----------|
| **Расположение** | Разные endpoints — прямое создание сервисов vs `Depends` |
| **Влияние** | Сложнее мокать в тестах, дублирование wiring. |
| **Исправление** | Единые dependency-функции в `app/dependencies/`. |

#### [A9] File-based draft storage — проблемы масштабирования

| Поле | Значение |
|------|----------|
| **Расположение** | Черновики КП/планов на файловой системе |
| **Влияние** | При горизонтальном масштабировании (несколько API-инстансов) черновики на одном инстансе недоступны на другом. |
| **Исправление** | Shared storage (S3/MinIO) или хранение черновиков в БД с привязкой к user_id. |

---

### Безопасность

#### [S6] OpenAPI exposed в production

| Поле | Значение |
|------|----------|
| **Расположение** | `app/main.py` — `/docs`, `/openapi.json` |
| **Влияние** | Полная карта API доступна атакующему без аутентификации. |
| **Исправление** | Отключать docs в production через env-флаг (`docs_url=None`). |

#### [S7] Отсутствуют security headers

| Поле | Значение |
|------|----------|
| **Расположение** | FastAPI middleware, nginx config |
| **Влияние** | Нет защиты от clickjacking, MIME sniffing, XSS через заголовки. |
| **Исправление** | Добавить `X-Frame-Options`, `X-Content-Type-Options`, `Content-Security-Policy`, `Strict-Transport-Security`. |

#### [S8] Внутренние детали ошибок в ответах API

| Поле | Значение |
|------|----------|
| **Расположение** | `app/api/v1/endpoints/archive.py`, `production.py`, `admin.py` |
| **Влияние** | Stack trace или внутренние сообщения могут утечь клиенту. |
| **Исправление** | Централизованный exception handler: клиенту — generic, в лог — полный traceback. |

#### [S9] OCR отправляет загрузки в OpenAI

| Поле | Значение |
|------|----------|
| **Расположение** | `core/ocr_gpt.py` |
| **Влияние** | Пользовательские документы (возможно конфиденциальные) уходят на внешний API. Требуется явное согласие и DPA. |
| **Исправление** | Документировать политику данных; опционально local OCR; маскирование PII перед отправкой. |

#### [S10] OCR rate limit только in-process

| Поле | Значение |
|------|----------|
| **Расположение** | In-memory счётчик в OCR-модуле |
| **Влияние** | При нескольких workers лимит обходится (каждый процесс — свой счётчик). |
| **Исправление** | Redis или nginx rate limit на OCR endpoint. |

#### [S11] Auth загружает всех пользователей (perf/DoS)

| Поле | Значение |
|------|----------|
| **Расположение** | Дублирует A3 с точки зрения security |
| **Влияние** | Amplification: злоумышленник генерирует много запросов, каждый тянет полную таблицу users. |
| **Исправление** | См. A3 — `get_user_by_id`. |

#### [S12] Нет horizontal access control для managers

| Поле | Значение |
|------|----------|
| **Расположение** | Endpoints offers/archive/production |
| **Влияние** | Менеджер теоретически может получить доступ к чужим КП/планам по ID (IDOR). |
| **Исправление** | Проверка ownership: `resource.manager_id == current_user.id` или role-based scope. |

#### [S13] Split JSON + SQLite races

| Поле | Значение |
|------|----------|
| **Расположение** | Связано с A1 |
| **Влияние** | Concurrent-запись в JSON и SQLite без lock — corruption данных. |
| **Исправление** | File locking, transactional outbox, или единое хранилище. |

#### [S14] Health endpoint раскрывает environment

| Поле | Значение |
|------|----------|
| **Расположение** | Health/readiness probe |
| **Влияние** | Утечка версии, режима debug, путей — помощь в targeted attacks. |
| **Исправление** | Минимальный ответ `{"status": "ok"}` для публичного health; детали — только internal/admin. |

---

### Качество кода

#### [Q6] DRY — дублированный try/except в commercial API

| Поле | Значение |
|------|----------|
| **Расположение** | Commercial endpoints |
| **Исправление** | Общий decorator или dependency для обработки commercial-исключений. |

#### [Q7] Дублированный format_phone

| Поле | Значение |
|------|----------|
| **Расположение** | Несколько модулей bot и app |
| **Исправление** | Вынести в `core/utils/phone.py`. |

#### [Q8] Magic-string ошибки в offers_service

| Поле | Значение |
|------|----------|
| **Расположение** | `app/services/offers_service.py` |
| **Исправление** | Enum или константы ошибок в `core/exceptions.py`. |

#### [Q9] Неполная обработка ошибок на offers/production endpoints

| Поле | Значение |
|------|----------|
| **Расположение** | `offers.py`, `production.py` |
| **Исправление** | Явные HTTPException для domain errors, единый handler. |

#### [Q10] Импорт private helper из day_view_service

| Поле | Значение |
|------|----------|
| **Расположение** | Cross-module import `_private_*` |
| **Исправление** | Сделать функцию публичной в shared module или вынести в core. |

#### [Q11] Слабая типизация dict в response

| Поле | Значение |
|------|----------|
| **Расположение** | Несколько endpoints возвращают `dict` без schema |
| **Исправление** | Pydantic response models для всех публичных API. |

#### [Q12] Per-request construction сервисов

| Поле | Значение |
|------|----------|
| **Расположение** | Endpoints создают service внутри handler |
| **Исправление** | Singleton/factory через `Depends`. |

#### [Q13] print() в core-модулях

| Поле | Значение |
|------|----------|
| **Расположение** | `core/*` |
| **Исправление** | Заменить на `logger.debug/info/warning`. |

#### [Q14] Stub wizard helper

| Поле | Значение |
|------|----------|
| **Расположение** | Commercial offer wizard |
| **Исправление** | Реализовать или удалить мёртвый код. |

#### [Q15] Слабая типизация file_generation_service

| Поле | Значение |
|------|----------|
| **Расположение** | `app/services/file_generation_service.py` |
| **Исправление** | TypedDict или dataclass для параметров генерации. |

#### [Q16] Дублированные генераторы day documents

| Поле | Значение |
|------|----------|
| **Расположение** | Bot handlers и `day_documents_service` |
| **Исправление** | Единый генератор в service, bot — thin wrapper. |

---

## Низкий приоритет / рекомендации

### Архитектура

| ID | Проблема | Расположение | Рекомендация |
|----|----------|--------------|--------------|
| [A10] | Дублирование web UI стеков | Frontend + legacy web routes | Консолидировать на React SPA |
| [A11] | Re-export shims | Промежуточные `__init__.py` | Удалить после миграции импортов |
| [A12] | Дублированные бизнес-константы frontend/backend | `app/core/constants.py`, frontend config | Единый source of truth или codegen |
| [A13] | Нет connection pooling | Raw sqlite3 connections | SQLAlchemy pool или aiosqlite для async |

### Безопасность

| ID | Проблема | Рекомендация |
|----|----------|--------------|
| [S15] | Слабая политика паролей | Минимум 12 символов, проверка complexity |
| [S16] | Role values не в allowlist | Pydantic Literal/Enum для role |
| [S17] | Logout через unauthenticated GET | POST /logout с CSRF |
| [S18] | Нет Python dependency audit в CI | `pip-audit` или `safety` в pipeline |
| [S19] | PBKDF2 вместо argon2/bcrypt | Миграция при следующем login (upgrade hash) |

### Качество кода

| ID | Проблема | Рекомендация |
|----|----------|--------------|
| [Q17] | Deprecated legacy_runtime | Удалить после полной миграции на PlateOrderContext |
| [Q18] | Нереализованные export handlers | Реализовать или пометить 501 / скрыть UI |
| [Q19] | Непоследовательные HTTP error patterns | Единый `AppError` → HTTP mapping |
| [Q20] | Thin pass-through wrappers | Инлайн или явный facade с ценностью |
| [Q21] | CommercialCalculationService без тестов | Unit-тесты на граничные случаи расчёта |

---

## Матрица приоритетов

| ID | Проблема | Severity | Effort | Priority |
|----|----------|----------|--------|----------|
| A1 | Split persistence без транзакций | Critical | High | P0 |
| S1 | Default admin/admin123 в seed | Critical | Low | P0 |
| A2 | Тройной presentation layer | High | High | P1 |
| A3 | Auth loads all users | High | Medium | P1 |
| A4 | God modules | High | High | P1 |
| A5 | Bot imports app layer | High | Medium | P1 |
| S2 | No rate limiting on auth | High | Low | P1 |
| S3 | Cookie session без CSRF | High | Medium | P1 |
| S4 | npm audit vulnerabilities | High | Low | P1 |
| S5 | Sessions после password/role change | High | Medium | P1 |
| Q1 | Agent-debug instrumentation | High | Low | P1 |
| Q2 | Bare except / pass | High | Medium | P1 |
| Q3 | Generic validation errors to client | High | Low | P1 |
| Q4 | Extreme module complexity | High | High | P2 |
| Q5 | Missing tests offers/production | High | Medium | P1 |
| A6 | Inconsistent repository pattern | Medium | Medium | P2 |
| A7 | Legacy global mutable state | Medium | Medium | P2 |
| A8 | Inconsistent FastAPI DI | Medium | Low | P2 |
| A9 | File-based draft scaling | Medium | High | P3 |
| S6 | OpenAPI in prod | Medium | Low | P2 |
| S7 | Missing security headers | Medium | Low | P2 |
| S8 | Internal errors in API responses | Medium | Low | P2 |
| S9 | OCR → OpenAI | Medium | Low | P2 |
| S10 | OCR rate limit in-process | Medium | Medium | P2 |
| S11 | Auth perf/DoS (dup A3) | Medium | Medium | P1 |
| S12 | No horizontal access control | Medium | Medium | P2 |
| S13 | JSON+SQLite races (dup A1) | Medium | High | P0 |
| S14 | Health exposes environment | Medium | Low | P3 |
| Q6 | DRY commercial try/except | Medium | Low | P3 |
| Q7 | Duplicated format_phone | Medium | Low | P3 |
| Q8 | Magic-string errors | Medium | Low | P3 |
| Q9 | Missing error handling endpoints | Medium | Medium | P2 |
| Q10 | Private helper imports | Medium | Low | P3 |
| Q11 | Loose dict response typing | Medium | Medium | P2 |
| Q12 | Per-request service construction | Medium | Low | P3 |
| Q13 | print() in core | Medium | Low | P3 |
| Q14 | Stub wizard helper | Medium | Low | P3 |
| Q15 | Weak typing file_generation | Medium | Low | P3 |
| Q16 | Duplicate day doc generators | Medium | Medium | P2 |
| A10 | Duplicate web UI stacks | Low | High | P4 |
| A11 | Re-export shims | Low | Low | P4 |
| A12 | Duplicated constants | Low | Medium | P4 |
| A13 | No connection pooling | Low | Medium | P4 |
| S15 | Weak password policy | Low | Low | P3 |
| S16 | Role not allowlisted | Low | Low | P3 |
| S17 | Logout via GET | Low | Low | P3 |
| S18 | No Python dep audit CI | Low | Low | P3 |
| S19 | PBKDF2 vs argon2 | Low | Medium | P4 |
| Q17 | Deprecated legacy_runtime | Low | Medium | P4 |
| Q18 | Unimplemented export handlers | Low | Medium | P4 |
| Q19 | Inconsistent HTTP errors | Low | Medium | P4 |
| Q20 | Thin pass-through wrappers | Low | Low | P4 |
| Q21 | CommercialCalculationService untested | Low | Medium | P3 |

**Легенда Priority**: P0 — немедленно, P1 — этот спринт, P2 — следующий спринт, P3/P4 — бэклог.

---

## Следующие шаги

### 1. Немедленно (до следующего коммита)

- **[S1]** Удалить `docker/seed/plita.db` с дефолтным admin из git; добавить в `.gitignore` если нужен локальный seed.
- **[Q1]** Удалить все блоки `#region agent log` из production-кода.
- **[S4]** Выполнить `npm audit fix` в `frontend/`.
- Документировать процедуру первичного создания admin через `scripts/create_admin.py`.

### 2. Этот спринт

- **[A1] / [S13]** Спроектировать единый persistence port для планов; минимум — file lock + transactional commit в `core/plan_commit.py`.
- **[A3] / [S11]** Реализовать `get_user_by_id`, перенести `init_schema` в startup.
- **[S2]** Rate limiting на login (nginx или slowapi).
- **[S3]** CSRF-защита для cookie-based сессий.
- **[S5]** Session revocation при смене пароля/роли.
- **[Q2]**, **[Q3]**, **[Q5]** Исправить exception handling и добавить тесты для offers/production API.
- **[A5]** Начать вынос общей логики из bot handlers в `core/`.

### 3. Следующий спринт

- **[A2]**, **[A4]**, **[Q4]** Рефакторинг god-modules и консолидация use-cases.
- **[S6]–[S8]**, **[S12]** Hardening API: скрыть OpenAPI, security headers, IDOR checks.
- **[Q11]**, **[Q16]** Типизация ответов и унификация day documents.
- **[A6]–[A8]** Унификация repository pattern и FastAPI DI.

### 4. Бэклог

- **[A9]**, **[A10]**, **[A13]** Масштабирование: shared storage, connection pooling, UI consolidation.
- **[S9]**, **[S10]**, **[S14]**, **[S15]–[S19]** Улучшения security posture.
- **[Q6]–[Q21]**, **[A11]–[A12]** Tech debt и качество кода.

---

## Положительные моменты

Несмотря на низкий Health Score, в проекте уже есть зрелые решения, на которые стоит опереться при рефакторинге:

- **`PlateOrderContext`** (`core/plate_order_context.py`) — request-scoped изоляция runtime-состояния плит и оптимизации; middleware `plate_runtime_isolation` предотвращает утечку глобального state между запросами.
- **Parameterized SQL** — в repository-слое используются параметризованные запросы, снижающие риск SQL-инъекций.
- **`destructive_db_guard`** (`core/destructive_db_guard.py`) — явная блокировка опасных операций с БД без подтверждения; покрыто тестами (`tests/test_destructive_db_guard.py`).
- **Draft ownership** — черновики коммерческих предложений привязаны к пользователю, что закладывает основу для horizontal access control (нужно усилить проверки — S12).
- **Разделение core / app / bot** — задумка слоёвой архитектуры правильная; требуется довести до конца (устранить A5, A2).
- **Тестовое покрытие** — есть тесты на оптимизацию, persistence, auth repository, plate completion, archive service — хорошая база для расширения (Q5).
- **`scripts/create_admin.py`** — готовый механизм безопасного bootstrap администратора (нужно сделать единственным путём — S1).
- **Pydantic schemas** — коммерческие и auth схемы типизированы; направление для унификации остальных endpoints (Q11).

---

*Отчёт сформирован на основе аудита архитектуры, безопасности и качества кода. Следующий полный аудит рекомендуется после закрытия P0 и P1 findings.*
