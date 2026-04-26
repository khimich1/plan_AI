---
date: 2026-04-26
topic: после merge обзор кодовой базы
scope:
  - app
  - bot
  - frontend
  - core
  - viz_modules
  - factory_cost
  - tests
---

# Исследование: После merge обзор кодовой базы

## Резюме
Исследованы backend-слои FastAPI (`app/main.py`, `app/api/v1/endpoints/*`, `app/services/*`, `app/repositories/*`), web-роуты и SPA-shell (`app/web/router.py`), frontend SPA (`frontend/src/*`), Telegram-бот (`bot/*`), а также ключевые модули расчётов и хранения (`core/*`, `viz_modules/procurement.py`, `factory_cost/cost_engine.py`) и тесты (`tests/test_commercial_web_flow.py`). Поток данных в коммерческом сценарии проходит через `frontend`/`web`/`bot` к API и сервисам, затем в `core`/`viz_modules` и SQLite через репозитории. Производственный контур реализован через `production` endpoint-ы и `ProductionService`, который использует репозитории и модули планирования из `bot.handlers.plan_manager`. Авторизация в API и web-части основана на cookie `app_session` с HMAC-подписью и ролевой проверкой через зависимости FastAPI.

## Подробные находки

**Расположение:** `app/main.py:17-56`, `main.py:1-3`  
**Слой:** endpoint  
**Что делает:** создаёт приложение FastAPI через `create_app()`, подключает CORS, health route, API v1 и web router; на lifespan настраивает логирование и инициализирует схему пользователей.  
**Входы:** настройки из `get_settings()`.  
**Выходы:** экземпляр `FastAPI` (`app`).  
**Ключевые зависимости:** `app/api/v1/router.py`, `app/web/router.py`, `AuthRepository.init_schema()`, `core.logging_config.setup_logging()`.  
**Связи:** корневой `main.py` переэкспортирует `app` для ASGI.  
**Паттерны:** инициализация инфраструктуры в lifespan + явное подключение роутеров.

**Расположение:** `app/api/v1/router.py:5-14`  
**Слой:** endpoint  
**Что делает:** агрегирует endpoint-модули API v1 (`health`, `auth`, `managers`, `commercial`, `archive`, `production`, `admin`).  
**Входы:** импортированные роутеры.  
**Выходы:** единый `APIRouter`.  
**Ключевые зависимости:** `app/api/v1/endpoints/*`.  
**Связи:** монтируется в `app/main.py` под `/api/v1`.  
**Паттерны:** модульная сборка API через `include_router`.

**Расположение:** `app/dependencies/auth.py:13-34`, `app/security/session.py:19-42`  
**Слой:** dependency / security  
**Что делает:** `get_current_user()` читает cookie `app_session`, декодирует токен, проверяет активность пользователя; `require_roles()` реализует ролевую проверку; `create_session_token()/decode_session_token()` создают и валидируют HMAC-подписанный payload с `exp`.  
**Входы:** `Request.cookies`, `AuthRepository.list_users()`, секрет `APP_SECRET_KEY`.  
**Выходы:** словарь пользователя или `HTTPException` 401/403.  
**Ключевые зависимости:** `AuthRepository`, `hmac`, `hashlib`, `base64`, `json`.  
**Связи:** используются в API endpoint-ах и web-роутах через `Depends`.  
**Паттерны:** cookie-session + role guard через closure dependency.

**Расположение:** `app/web/router.py:19-997`  
**Слой:** web  
**Что делает:** обслуживает HTML-страницы `/web/*`, форму и preview КП, выдаёт SPA shell на `/commercial-offer/new` и фронтенд-ассеты `/commercial-offer/assets/{asset_path:path}`; реализует login/logout cookie flow, вызывает сервисы коммерции и производства.  
**Входы:** form-data (`/web/login`, `/web/offers/new`), cookies, файлы изображений, path params (`draft_id`, `asset_path`).  
**Выходы:** `HTMLResponse`, `RedirectResponse`, `FileResponse`, ошибки 404/400.  
**Ключевые зависимости:** `CommercialService`, `CommercialWorkflowService`, `ProductionService`, `AuthRepository`, `create_session_token`, `require_roles`.  
**Связи:** frontend dist читается из `settings.frontend_dist_dir`; web-страницы используют API `/api/v1/*` в embedded JS на странице менеджеров.  
**Паттерны:** server-rendered HTML + SPA shell в одном роутере, инстанцирование сервисов в обработчиках.

**Расположение:** `app/api/v1/endpoints/commercial.py:26-297`  
**Слой:** endpoint  
**Что делает:** API коммерческого потока: parse, создание/обновление draft, обработка wide plates, обновление meta, расчёт draft, генерация файлов, сохранение draft, скачивание файлов, получение draft.  
**Входы:** JSON (`Commercial*Request`) и multipart (`text`, `image`, `mode`), `draft_id`.  
**Выходы:** `CommercialDraftDetailsResponse`, `CommercialGenerateFilesResponse`, `CommercialSaveOfferResponse`, `FileResponse`, `HTTPException` 400/404.  
**Ключевые зависимости:** `CommercialWorkflowService`, `CommercialService`, `DraftStore`, `PlateParseError`, `require_roles("admin","manager")`.  
**Связи:** вызывается из SPA (`features/commercial-offer/api/commercialOfferApi.ts`) и web формы `/web/offers/new`.  
**Паттерны:** thin endpoint + бизнес-валидации/оркестрация в сервисе workflow.

**Расположение:** `app/api/v1/endpoints/offers.py:10-122`, `app/services/offers_service.py:14-237`  
**Слой:** endpoint / service  
**Что делает:** управление сохранёнными КП: список, карточка, создание, изменение скидки, перевод в производство, удаление, генерация PDF/XLSX.  
**Входы:** query (`status`, `limit`, `kp_id`), body (`CreateOfferRequest`, `UpdateOfferDiscountRequest`, `MoveOfferToProductionRequest`).  
**Выходы:** словари с агрегированными данными КП, бинарные документы через `Response`, HTTP 404/400.  
**Ключевые зависимости:** `KpRepository`, `core.kp_db`, `core.commercial_offer`, `core.commercial_offer_xlsx`.  
**Связи:** используется страницей менеджеров в `app/web/router.py` через fetch к `/api/v1/offers*`.  
**Паттерны:** endpoint маппит доменные ошибки сервиса в HTTP-ошибки; сервис возвращает summary/details с перерасчётом completion.

**Расположение:** `app/services/commercial_service.py:31-150`, `app/services/commercial_workflow_service.py:20-734`, `app/services/draft_store.py:17-105`  
**Слой:** service  
**Что делает:** `CommercialService` парсит заказ и формирует preview (оптимизация, price rows, breakdown, order_data); `CommercialWorkflowService` управляет draft lifecycle (text/image OCR input, merge/replace, wide plate decisions, metadata, calculate, files, save); `DraftStore` сохраняет/загружает draft JSON в `drafts_dir`.  
**Входы:** текст плит, OCR bytes, метаданные формы, решения по wide plates, режимы save/files.  
**Выходы:** draft-структуры, generated files metadata, данные сохранения в БД.  
**Ключевые зависимости:** `PlateParserService`, `OptimizationService`, `FileGenerationService`, `ManagerRepository`, `KpRepository`, `ExecutionTermsService`, `core.ocr_gpt`, `core.commercial_offer_xlsx.calculate_total_cost`.  
**Связи:** вызываются API `/commercial/*` и web `/web/offers/*`; используют repository/core/viz и файловое хранилище.  
**Паттерны:** отдельный workflow-оркестратор + serializable draft state через filesystem.

**Расположение:** `app/api/v1/endpoints/archive.py:27-168`, `app/services/archive_service.py:43-362`, `app/repositories/kp_archive_repository.py:12-40`  
**Слой:** endpoint / service / repository  
**Что делает:** архивный API по разделам (`archived/in_production/completed`), поиск КП, карточка, документы PDF/XLSX, изменение скидки, удаление, перевод в производство, оценка производства, выгрузка текущего плана Gantt.  
**Входы:** section/query/kp_id, payload для discount/move.  
**Выходы:** typed Pydantic модели (`ArchiveOffer*`), `FileResponse`, HTTP 404/400/500.  
**Ключевые зависимости:** `KpArchiveRepository` (обёртка над `core.kp_db`), `core.commercial_offer`, `core.commercial_offer_xlsx`, `core.gantt_excel`, отложенный импорт `bot.handlers.plan_manager`.  
**Связи:** frontend archive page использует эти endpoint-ы через hooks `useArchiveListQuery/useArchiveSearchQuery`.  
**Паттерны:** endpoint использует `Depends(get_archive_service)`; сервис маппит raw dict в схемы.

**Расположение:** `app/api/v1/endpoints/production.py:32-247`, `app/services/production_service.py:17-131`, `app/repositories/plan_repository.py:7-61`, `app/repositories/work_calendar_repository.py:10-42`  
**Слой:** endpoint / service / repository  
**Что делает:** API производства (планы, календарь, день, документы, кандидаты, рабочий календарь); сервис делегирует планирование в `ProductionPlanningService` и `bot.handlers.plan_manager`, хранение календаря — в JSON через `WorkCalendarRepository`.  
**Входы:** `BuildPlanRequest`, `CreatePlanRequest`, `CompleteProductionDayRequest`, `SaveWorkCalendarRequest`, path/query params.  
**Выходы:** typed ответы (`BuildPlanResponse`, `DayViewDetailResponse`, др.), документы (`schema/breakdown/formovka`), ошибки 404/422/500.  
**Ключевые зависимости:** `KpRepository`, `PlanRepository`, `WorkCalendarRepository`, `day_documents_service`, `day_view_service`, `plan_manager`.  
**Связи:** frontend `ProductionPage` и production feature вызывают `/api/v1/production/*`.  
**Паттерны:** service-фасад поверх репозиториев + reuse bot-логики через repository слой.

**Расположение:** `app/api/v1/endpoints/admin.py:13-100`, `app/services/admin_service.py:34-170`, `app/repositories/auth_repository.py:28-150`  
**Слой:** endpoint / service / repository  
**Что делает:** admin API для статистики БД, reset-операций (full/kp/plans/calendar), восстановления “застрявших” плит; `AuthRepository` хранит пользователей в `app_users`, выполняет hash/verify пароля и аутентификацию.  
**Входы:** admin role, без дополнительного payload для reset endpoints.  
**Выходы:** `DbStatsResponse`, `DbResetReport`, `RecoverPlatesResponse`; user payload из репозитория auth.  
**Ключевые зависимости:** `core.kp_db` операции reset/recover/stats, `PlanRepository`, `WorkCalendarRepository`, `sqlite3`, PBKDF2-hash.  
**Связи:** `AuthRepository.init_schema()` вызывается при старте app lifecycle.  
**Паттерны:** сервис централизует “опасные” операции и использует `Settings` пути.

**Расположение:** `app/core/settings.py:19-103`  
**Слой:** config  
**Что делает:** централизованные настройки через `BaseSettings`, загрузка `.env` + `bot/bot.env`, выдача путей к БД/директориям, CORS parsing, создание рабочих директорий.  
**Входы:** переменные окружения.  
**Выходы:** singleton `Settings` из `get_settings()`.  
**Ключевые зависимости:** `pydantic-settings`, `dotenv`.  
**Связи:** используется backend, bot config и сервисами/репозиториями по всему проекту.  
**Паттерны:** `@lru_cache` для конфигурации + computed field для CORS list.

**Расположение:** `run_bot.py:8-18`, `bot/bot_main.py:44-79`, `bot/handlers/__init__.py:11-31`, `bot/states.py:5-68`, `bot/keyboards.py:5-976`, `bot/handlers/commercial.py:45-2260`  
**Слой:** bot handler  
**Что делает:** отдельный запуск aiogram polling; регистрация роутеров; FSM-состояния для коммерции/производства/архива и др.; keyboard factory; крупный сценарий КП в боте (пошаговый FSM, OCR, preview xlsx, file generation callbacks, сохранение в БД/архив).  
**Входы:** Telegram updates (`message`, `callback_query`), FSM state, текст/фото пользователя.  
**Выходы:** Telegram сообщения/документы, сохранение КП через `core.kp_db.save_kp_to_db`, обновление state.  
**Ключевые зависимости:** `CommercialService`, `OptimizationService`, `core.*`, `viz_modules.procurement`, `aiogram`.  
**Связи:** bot использует те же базы и core-модули, что и backend.  
**Паттерны:** сценарии разделены на stateful handlers + keyboard-driven UX.

**Расположение:** `frontend/src/main.tsx:1-13`, `frontend/src/app/providers/AppProviders.tsx:1-13`, `frontend/src/app/router/AppRouter.tsx:1-23`, `frontend/src/shared/api/httpClient.ts:43-124`, `frontend/src/features/commercial-offer/api/commercialOfferApi.ts:56-116`, `frontend/src/pages/*`  
**Слой:** frontend app / shared / feature / page  
**Что делает:** SPA entrypoint, провайдеры (`react-query`, auth, wizard draft store), router (login/new/archive/production), HTTP client с `credentials: "include"` и unified error handling, feature API для коммерческого draft workflow, страницы архива/производства/логина.  
**Входы:** browser routes, формы (`react-hook-form + zod` в login), API path/payload.  
**Выходы:** UI-страницы и запросы в backend (`/api/v1/*`).  
**Ключевые зависимости:** `react-router-dom`, `@tanstack/react-query`, `react-hook-form`, `zod`.  
**Связи:** коммерческая фича ходит в `/api/v1/commercial/*` и `/api/v1/managers`; archive page использует archive hooks; production page переключает production tabs.  
**Паттерны:** FSD-подобное разделение `app/pages/features/shared`.

**Расположение:** `core/plate_line_parser.py:14-161`, `core/commercial_offer_xlsx.py:98-532`, `core/work_calendar.py:16-105`, `viz_modules/procurement.py:31-1326`, `factory_cost/cost_engine.py:20-253`  
**Слой:** core / viz / factory_cost  
**Что делает:** парсер строк плит (WxL и PB/ПК форматы), генерация XLSX КП и расчёт totals, рабочий календарь из JSON, построение закупки/price rows/breakdown с учётом планов и нагрузок, API чтения заводской себестоимости из SQLite.  
**Входы:** текст заказа, `order_data`, данные оптимизации (`OPT_*`), plate params/name, JSON calendar.  
**Выходы:** `LineParseResult`, xlsx buffer, calendar date utilities, procurement rows/tables, cost dicts.  
**Ключевые зависимости:** `core.config_and_data`, `core.optimization`, `core.price_db`, `sqlite3`, `openpyxl/pandas`.  
**Связи:** используются сервисами (`CommercialService`, `ArchiveService`, bot commercial handler).  
**Паттерны:** shared legacy-модули как backend/bot common runtime.

**Расположение:** `core/kp_db.py:60`, `core/kp_db.py:420`, `core/kp_db.py:509`, `core/kp_db.py:738`, `core/kp_db.py:3154`  
**Слой:** core  
**Что делает:** хранит и обслуживает ключевые DB-операции КП/менеджеров/статусов: инициализация схемы, сохранение КП, загрузка КП по id, обогащение номенклатурой, группировка списков КП по разделам.  
**Входы:** DB path, order data, параметры КП.  
**Выходы:** id/структуры КП и менеджеров, bool для мутаций.  
**Ключевые зависимости:** sqlite и внутренние SQL helpers модуля.  
**Связи:** используется через `KpRepository`, `KpArchiveRepository`, bot handlers, admin/archive/offers services.  
**Паттерны:** централизация SQLite-доступа в legacy-модуле и thin wrappers в `app/repositories`.

**Расположение:** `tests/test_commercial_web_flow.py:44-208`  
**Слой:** test  
**Что делает:** проверяет API/web коммерческого флоу через `TestClient`: создание draft (text/image), auth guard, скачивание generated file и path boundary check, web submit + redirect в draft page.  
**Входы:** monkeypatch сервисов (`CommercialWorkflowService`, `CommercialService`), auth cookie с `create_session_token`.  
**Выходы:** assertions по HTTP status/payload/location/content.  
**Ключевые зависимости:** `create_app`, `AuthRepository`, `CommercialWorkflowService`, `CommercialService`.  
**Связи:** покрывает `app/web/router.py` и `app/api/v1/endpoints/commercial.py`.  
**Паттерны:** интеграционный web/API flow test с изоляцией бизнес-слоя через monkeypatch.

## Поток данных
- `React SPA (frontend/src/pages/commercial-offer-create/CommercialOfferCreatePage.tsx)` → `features/commercial-offer/api/commercialOfferApi.ts` → `shared/api/httpClient.ts` → `FastAPI endpoint app/api/v1/endpoints/commercial.py` → `CommercialWorkflowService`/`CommercialService` → `viz_modules/procurement.py` + `core/*` → `DraftStore` JSON + `KpRepository`/`core.kp_db` + файлы в `outputs_dir`.
- `Web HTML (app/web/router.py)` → `Depends(get_current_user/require_roles)` → `CommercialWorkflowService`/`ProductionService`/`OffersService` → `KpRepository`/`PlanRepository`/`core.kp_db` → HTML-ответы/редиректы/скачивание файлов.
- `Telegram (bot/handlers/commercial.py)` → `FSM (bot/states.py)` → `CommercialService`/`OptimizationService` + `core.ocr_gpt`/`viz_modules.procurement` → `core.kp_db`/файлы в `OUTPUTS_DIR_STR` → сообщения и документы в Telegram.
- `Archive & Production API` (`app/api/v1/endpoints/archive.py`, `production.py`) → `ArchiveService`/`ProductionService` → `KpArchiveRepository`/`PlanRepository`/`WorkCalendarRepository` + `core.gantt_excel`/day docs services → SQLite/JSON/файлы документов.

## Ссылки на код
- `main.py:1` — переэкспорт ASGI `app`.
- `app/main.py:17` — lifespan с логированием и `AuthRepository.init_schema()`.
- `app/main.py:51` — монтирование API v1 под `/api/v1`.
- `app/main.py:52` — подключение web router.
- `app/api/v1/router.py:7-14` — сборка модулей API.
- `app/web/router.py:19` — web router без OpenAPI schema.
- `app/web/router.py:43-53` — отдача SPA shell из `frontend/dist/index.html`.
- `app/web/router.py:195-204` — login и установка cookie `app_session`.
- `app/dependencies/auth.py:13-25` — извлечение пользователя из cookie token.
- `app/dependencies/auth.py:28-34` — проверка ролей `require_roles`.
- `app/security/session.py:19-26` — создание session token.
- `app/security/session.py:29-42` — decode + проверка подписи/exp.
- `app/api/v1/endpoints/commercial.py:26` — префикс `/commercial`.
- `app/api/v1/endpoints/offers.py:10` — префикс `/offers`.
- `app/api/v1/endpoints/archive.py:27` — префикс `/commercial/archive`.
- `app/api/v1/endpoints/production.py:32` — префикс `/production`.
- `app/services/commercial_service.py:44-70` — генерация preview через optimization/procurement.
- `app/services/commercial_workflow_service.py:38-70` — create draft.
- `app/services/commercial_workflow_service.py:308-383` — генерация файлов draft.
- `app/services/commercial_workflow_service.py:443-481` — режимы save draft (`database/archive/skip`).
- `app/services/production_service.py:51-70` — делегирование build plan.
- `app/services/archive_service.py:149-199` — генерация документов архива.
- `app/repositories/kp_repository.py:17-43` — сохранение КП через `core.kp_db`.
- `app/repositories/auth_repository.py:34-57` — schema `app_users`.
- `app/core/settings.py:19-24` — конфигурация settings model.
- `app/core/settings.py:44-60` — пути БД/директорий (вкл. `frontend_dist_dir`).
- `bot/bot_main.py:44-76` — запуск polling aiogram.
- `bot/handlers/__init__.py:11-31` — регистрация роутеров бота.
- `bot/handlers/commercial.py:143-159` — старт сценария создания КП.
- `bot/handlers/commercial.py:952-1340` — генерация документов и подготовка save.
- `bot/states.py:5-68` — FSM state-группы.
- `frontend/src/main.tsx:7-13` — точка входа React.
- `frontend/src/app/providers/AppProviders.tsx:8-12` — `QueryClientProvider`, `AuthProvider`, draft store.
- `frontend/src/app/router/AppRouter.tsx:12-18` — маршруты login/new/archive/production.
- `frontend/src/shared/api/httpClient.ts:43-49` — fetch wrapper с `credentials: "include"`.
- `frontend/src/features/commercial-offer/api/commercialOfferApi.ts:59-115` — вызовы draft workflow API.
- `frontend/src/pages/login/LoginPage.tsx:26-29` — `useForm` + `zodResolver`.
- `core/plate_line_parser.py:89-161` — parse line to structured result.
- `core/commercial_offer_xlsx.py:98-145` — `calculate_total_cost`.
- `core/commercial_offer_xlsx.py:171-532` — генерация XLSX КП.
- `core/work_calendar.py:64-105` — рабочие дни и `nth_working_day`.
- `core/kp_db.py:60` — `init_schema`.
- `core/kp_db.py:509` — `save_kp_to_db`.
- `core/kp_db.py:738` — `get_kp_by_id`.
- `core/kp_db.py:3154` — `get_all_kp_list`.
- `viz_modules/procurement.py:277-340` — сборка procurement items.
- `viz_modules/procurement.py:541-690` — `build_price_rows`.
- `viz_modules/procurement.py:1006-1325` — `build_component_breakdown`.
- `factory_cost/cost_engine.py:20-87` — `get_cost_by_plate_name`.
- `factory_cost/cost_engine.py:89-179` — `get_cost_by_params`.
- `tests/test_commercial_web_flow.py:73-133` — тесты create draft from form.
- `tests/test_commercial_web_flow.py:173-208` — web flow `/web/offers/new`.

## Архитектурные наблюдения
- FastAPI-приложение создаётся в `app/main.py`, а корневой `main.py` переэкспортирует `app`.
- API v1 подключён под `/api/v1`, web-роуты подключены отдельно и исключены из OpenAPI (`include_in_schema=False`).
- Авторизация в API/web использует cookie `app_session` + декодирование токена и проверку ролей через dependency functions.
- Коммерческий сценарий разделён на endpoint слой (`app/api/v1/endpoints/commercial.py`) и workflow/service слой (`app/services/commercial_workflow_service.py`, `app/services/commercial_service.py`), с промежуточным хранением draft в файловой системе (`DraftStore`).
- Репозитории `app/repositories/*` в основном выступают thin-wrapper над legacy-модулями `core.kp_db` и `bot.handlers.plan_manager`.
- Telegram-бот запускается независимо через `run_bot.py`/`bot/bot_main.py`, но использует ту же предметную логику и источники данных, что backend.
- Frontend SPA построен модульно (`app/pages/features/shared`), использует `react-query` для data-fetching и `httpClient` с `credentials: "include"` для cookie-сессий.
