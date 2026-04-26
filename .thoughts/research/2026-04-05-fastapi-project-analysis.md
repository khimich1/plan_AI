---
date: 2026-04-05
topic: fastapi-project-analysis
---

# Исследование: FastAPI-проект

## Резюме
Точка входа web/API части находится в `app/main.py`: приложение поднимает FastAPI, подключает API v1 под префиксом `/api/v1` и web-роутер без OpenAPI-схемы. Основной HTTP-поток построен по схеме endpoint -> service -> repository, при этом часть репозиториев делегирует операции в существующие модули `core/*` и `bot/handlers/plan_manager.py`. Авторизация реализована через cookie `app_session` с HMAC-подписью и проверкой роли через dependency.

## Подробные находки

**Расположение:** `app/main.py:15-41`, `main.py:1-3`  
**Что делает:** задает lifespan, на старте настраивает логирование, выполняет bootstrap admin-пользователя и сохраняет settings в `app.state`; создает `FastAPI(...)`, добавляет `/health`, подключает API и web роутеры; корневой `main.py` реэкспортирует `app` для ASGI-запуска.  
**Ключевые зависимости:** `app.api.v1.router`, `app.web.router`, `app.core.settings`, `app.repositories.auth_repository`, `core.logging_config`.  
**Паттерны:** единая фабрика `create_app()`, инициализация инфраструктуры в lifespan.

**Расположение:** `app/api/v1/router.py:5-12`  
**Что делает:** агрегирует endpoint-модули `health`, `auth`, `managers`, `commercial`, `production` в один `APIRouter`.  
**Ключевые зависимости:** `app.api.v1.endpoints.*`.  
**Паттерны:** централизованная регистрация роутеров уровня версии API.

**Расположение:** `app/api/v1/endpoints/auth.py:10-45`, `app/dependencies/auth.py:13-34`, `app/security/session.py:13-42`, `app/repositories/auth_repository.py:28-121`, `app/schemas/auth.py:6-8`  
**Что делает:**  
- `POST /api/v1/auth/login` принимает `LoginRequest`, аутентифицирует пользователя через `AuthRepository.authenticate`, создает signed session token, записывает cookie `app_session`, возвращает `{"user": ...}`.  
- `POST /api/v1/auth/logout` удаляет cookie.  
- `GET /api/v1/auth/me` возвращает текущего пользователя из dependency.  
- `get_current_user` читает cookie, декодирует токен, проверяет пользователя по `list_users()` и активность.  
- `require_roles` ограничивает доступ по `user.role`.  
**Ключевые зависимости:** `AuthRepository`, `create_session_token/decode_session_token`, `LoginRequest`.  
**Паттерны:** cookie-based session auth, role-based access через `Depends`.

**Расположение:** `app/api/v1/endpoints/managers.py:8-17`, `app/services/commercial_service.py:31-42`, `app/repositories/manager_repository.py:7-16`  
**Что делает:** `GET /api/v1/managers` после проверки ролей получает список менеджеров из `CommercialService.list_managers()`, который проксирует в `ManagerRepository.list_managers()` и далее в `core.kp_db.get_all_managers`.  
**Ключевые зависимости:** `require_roles`, `CommercialService`, `ManagerRepository`.  
**Паттерны:** тонкий endpoint + сервис-обертка + репозиторий-адаптер к legacy core.

**Расположение:** `app/api/v1/endpoints/commercial.py:10-78`, `app/services/commercial_service.py:31-150`, `app/services/plate_parser_service.py:21-214`, `app/services/optimization_service.py:14-45`, `app/services/draft_store.py:13-56`  
**Что делает:**  
- `POST /api/v1/commercial/parse` валидирует `CommercialParseRequest`, парсит текст заказа и возвращает структуру `order + diagnostics/warnings`.  
- `POST /api/v1/commercial/generate-preview` строит превью: parse -> optimize -> расчет прайса/разбивок -> сохранение черновика в JSON (`draft_id`) -> ответ с агрегацией данных.  
- `GET /api/v1/commercial/drafts/{draft_id}` поднимает черновик из диска, восстанавливает `PlateOrder` и `OptimizationContext`, возвращает `found` флаг и payload.  
- `PlateParserService.parse_plate_text()` нормализует текст, построчно парсит (`parse_line`), валидирует (`validate_plate_values`), заполняет `PlateOrder`, diagnostics и warnings, затем дополняет nomenclature cache из `pb.db`.  
- `OptimizationService.optimize()` преобразует `PlateOrder -> orders_2d`, вызывает `core.optimization.optimize_with_cascading_longitudinal_cuts`, собирает `OptimizationContext`; `legacy_runtime()` временно синхронизирует глобалы legacy-оптимизации.  
- `DraftStore` сохраняет/читает JSON-файлы черновиков в `settings.drafts_dir`.  
**Ключевые зависимости:** `CommercialParseRequest`, `CommercialPreviewRequest`, `PlateOrder`, `ParseResult`, `OptimizationContext`, `core.*`, `viz_modules.*`.  
**Паттерны:** orchestrator-сервис (`CommercialService`) + dataclass-модели домена + адаптация к legacy runtime через context managers.

**Расположение:** `app/api/v1/endpoints/production.py:9-97`, `app/services/production_service.py:13-86`, `app/repositories/plan_repository.py:7-61`, `app/repositories/work_calendar_repository.py:10-42`, `app/repositories/kp_repository.py:10-75`, `bot/handlers/plan_manager.py:90-1261`, `core/work_calendar.py:46-105`  
**Что делает:**  
- `GET/POST /api/v1/production/plans`, `GET /plans/{plan_id}`, `POST /plans/{plan_id}/activate`, `GET /calendar`, `GET /days/{target_date}`, `POST /days/{target_date}/complete`, `GET /candidates`, `GET/PUT /work-calendar`.  
- `ProductionService` маршрутизирует операции: планы/календарь/дни через `PlanRepository` и `plan_manager`, кандидатов через `KpRepository`, календарные настройки через `WorkCalendarRepository`.  
- `PlanRepository` является адаптером к функциям `bot.handlers.plan_manager` (metadata, load/save plan, add tracks, complete day, lookup day).  
- `WorkCalendarRepository` хранит JSON рабочего календаря и использует `core.work_calendar` для вычислений рабочих дней.  
- `KpRepository.list_production_candidates()` читает из `kp_plates` записи в статусах `'в производстве'` и `'в плане'`.  
**Ключевые зависимости:** `CreatePlanRequest`, `CompleteProductionDayRequest`, `SaveWorkCalendarRequest`, `plan_manager`, `core.work_calendar`.  
**Паттерны:** API-слой как фасад над файловыми планами в `bot/data/plans` и SQLite-таблицами.

**Расположение:** `app/web/router.py:14-125`  
**Что делает:** реализует HTML-интерфейс:  
- `GET/POST /web/login` — форма и логин (cookie `app_session`),  
- `GET /web` — dashboard (кол-во менеджеров/планов),  
- `GET /web/managers`, `/web/offers`, `/web/production` — табличные страницы.  
Маршруты защищены `get_current_user` или `require_roles(...)`.  
**Ключевые зависимости:** `AuthRepository`, `create_session_token`, `CommercialService`, `ProductionService`.  
**Паттерны:** server-rendered HTML (строки + escape), единая шапка `_nav(user)`.

**Расположение:** `app/schemas/auth.py:6-8`, `app/schemas/commercial.py:6-11`, `app/schemas/production.py:6-24`  
**Что делает:** Pydantic-модели входных payload API:  
- `LoginRequest`: `username`, `password` с `min_length=1`;  
- `CommercialParseRequest` / `CommercialPreviewRequest`: `text` с `min_length=1`;  
- `CreatePlanRequest`: поля плана (`name`, `start_date`, `tracks_per_day` с `ge=1, le=50`, lookup/orders/result по умолчанию пустые),  
- `CompleteProductionDayRequest`: `plan_id`,  
- `SaveWorkCalendarRequest`: списки ISO-дат `extra_holidays`, `extra_workdays`.  
**Ключевые зависимости:** используется в endpoint-параметрах FastAPI для body validation.  
**Паттерны:** request-only схемы, возвраты преимущественно как `dict`.

**Расположение:** `app/core/settings.py:18-73`, `app/core/constants.py:3-12`  
**Что делает:**  
- `Settings` через `pydantic-settings` читает `.env` и `bot/bot.env`, задает пути (`plita.db`, `pb.db`, `plans`, `work_calendar`, `drafts`, `logs`) и bootstrap-параметры админа.  
- `get_settings()` кэшируется через `lru_cache(1)` и гарантирует создание директорий.  
- `constants.py` определяет производственные/ценовые и role-константы.  
**Ключевые зависимости:** используется в `main`, репозиториях, security и сервисах.  
**Паттерны:** централизованная конфигурация + filesystem bootstrap.

## Ссылки на код

- `app/main.py:24-38` — создание FastAPI и подключение роутеров
- `app/main.py:15-21` — lifespan: логирование, bootstrap admin, app.state.settings
- `app/api/v1/router.py:7-12` — агрегирующий API v1 router
- `app/api/v1/endpoints/auth.py:13-45` — login/logout/me
- `app/dependencies/auth.py:13-34` — текущий пользователь и RBAC dependency
- `app/security/session.py:19-42` — создание/проверка signed session token
- `app/api/v1/endpoints/commercial.py:13-78` — parse/preview/draft endpoint’ы
- `app/services/plate_parser_service.py:26-193` — основной парсинг текста плит
- `app/services/commercial_service.py:44-70` — orchestration preview расчёта
- `app/services/draft_store.py:19-56` — сохранение и восстановление draft preview
- `app/api/v1/endpoints/production.py:12-97` — endpoint’ы планирования и календаря
- `app/services/production_service.py:20-78` — сервис production API
- `app/repositories/plan_repository.py:11-61` — адаптер к `plan_manager`
- `bot/handlers/plan_manager.py:990-1115` — сбор глобального календаря
- `bot/handlers/plan_manager.py:1118-1261` — сбор дорожек по дате из всех планов
- `app/web/router.py:14` — web-router скрыт из OpenAPI (`include_in_schema=False`)
- `app/schemas/production.py:6-24` — Pydantic-модели production запросов
- `app/core/settings.py:69-73` — кэшированный `get_settings()` и создание директорий

## Архитектурные наблюдения

- Маршруты API сгруппированы по доменам (`auth`, `managers`, `commercial`, `production`) и подключаются через единый v1-роутер (`app/api/v1/router.py:5-12`).
- В `app` используется разделение на endpoint/service/repository, при этом репозитории `PlanRepository`, `ManagerRepository`, `KpRepository` делегируют часть операций в legacy-модули `bot.handlers.plan_manager` и `core.kp_db` (`app/repositories/*.py`).
- Доменная модель коммерческого контура построена на dataclass-структурах (`PlateOrder`, `ParseResult`, `OptimizationContext`) и сериализуется в JSON для draft-хранилища (`app/domain/models/*.py`, `app/services/draft_store.py:27-56`).
- Web-интерфейс и API используют одну и ту же сессионную cookie и dependency-цепочку авторизации/ролей (`app/web/router.py:79-125`, `app/dependencies/auth.py:13-34`).
