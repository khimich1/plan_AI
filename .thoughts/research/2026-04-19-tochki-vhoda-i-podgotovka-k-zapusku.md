---
date: 2026-04-19
topic: Точки входа приложения и подготовка к запуску (backend + Telegram-бот)
scope:
  - app
  - bot
  - core
---

# Исследование: точки входа и подготовка к запуску

## Резюме

Исследованы точки входа FastAPI (`app/main.py`) и Telegram-бота на aiogram 3 (`bot/bot_main.py`), загрузка настроек из переменных окружения и файлов `.env` / `bot/bot.env` (`app/core/settings.py`), сборка API v1 (`app/api/v1/router.py`) и HTML-маршрутов (`app/web/router.py`), регистрация хендлеров бота (`bot/handlers/__init__.py`), а также инициализация пользователей веб-кабинета через `AuthRepository.ensure_bootstrap_admin()` при старте приложения (`app/main.py:19`). В среде разработки проверены: создание venv, `pip install -r requirements.txt`, импорт `app.main:app` и краткий запуск `uvicorn`.

## Подробные находки

**Расположение:** `app/main.py:15-41`  
**Слой:** endpoint (корневое приложение), lifespan  
**Что делает:** Функция `lifespan` при старте вызывает `get_settings()`, `setup_logging(..., log_filename="backend.log")`, затем `AuthRepository(str(settings.plita_db_path)).ensure_bootstrap_admin()`, сохраняет `settings` в `app.state`. Функция `create_app()` создаёт `FastAPI` с `title` из настроек, `debug` из `settings.app_debug`, подключает `GET /health`, роутер API с префиксом `/api/v1` и web-router без префикса. Модуль экспортирует `app = create_app()`.  
**Входы:** Контекст ASGI при старте/остановке.  
**Выходы:** Экземпляр `FastAPI` в переменной `app`.  
**Ключевые зависимости:** `app.api.v1.router`, `app.web.router`, `app.core.settings.get_settings`, `app.repositories.auth_repository.AuthRepository`, `core.logging_config.setup_logging`.  
**Связи:** Запуск через ASGI-сервер (например uvicorn) с целевым объектом `app.main:app`.  
**Паттерны:** Использование `asynccontextmanager` для lifespan; bootstrap администратора при каждом старте, если заданы переменные окружения.

**Расположение:** `app/core/settings.py:11-73`  
**Слой:** dependency / configuration  
**Что делает:** `load_dotenv` вызывается для `PROJECT_ROOT / ".env"` и `BOT_DIR / "bot.env"` с `override=False`. Класс `Settings` наследует `BaseSettings`, читает `env_file` из тех же путей, поля включают `app_secret_key` (alias `APP_SECRET_KEY`), `bot_token` (`BOT_TOKEN`), пути к `pb.db`, `plita.db`, каталогам вывода, черновиков, логов, опционально `DATABASE_URL`, `REDIS_URL`, креды bootstrap-админа (`APP_ADMIN_USERNAME`, `APP_ADMIN_PASSWORD`, `APP_ADMIN_ROLE`). `get_settings()` кэшируется через `@lru_cache`, вызывает `ensure_directories()`.  
**Входы:** Переменные окружения и перечисленные env-файлы.  
**Выходы:** Экземпляр `Settings`.  
**Ключевые зависимости:** `pydantic_settings`, `python-dotenv`, `pathlib.Path`.  
**Связи:** Импортируется из `app.main`, репозиториев, сервисов, `bot/bot_config.py`.  
**Паттерны:** Единый корень проекта как `Path(__file__).resolve().parents[2]` от `app/core/settings.py`.

**Расположение:** `app/api/v1/router.py:7-12`  
**Слой:** endpoint (агрегация роутеров)  
**Что делает:** Создаётся `APIRouter()` без префикса на этом уровне; подключаются подроутеры `health`, `auth`, `managers`, `commercial`, `production`.  
**Входы:** HTTP-запросы к путям, определённым в подроутерах.  
**Выходы:** Ответы подроутеров.  
**Ключевые зависимости:** `app.api.v1.endpoints.*`.  
**Связи:** Включается в приложение в `app/main.py:36` с префиксом `/api/v1`.  
**Паттерны:** Модульная группировка по доменам.

**Расположение:** `app/web/router.py:16-339`  
**Слой:** endpoint (HTML), dependency  
**Что делает:** `APIRouter(include_in_schema=False)`; вспомогательные `_page`, `_nav`, формы КП; маршруты `/web/login` (GET/POST), `/web`, `/web/managers`, `/web/offers`, черновики КП, `/web/production`; используются `Depends(get_current_user)` и `Depends(require_roles(...))`, сервисы `CommercialService`, `CommercialWorkflowService`, `ProductionService`, `AuthRepository`, `create_session_token`; cookie `app_session` при успешном POST логина (`app/web/router.py:176-178`).  
**Входы:** HTTP, формы, файлы загрузки.  
**Выходы:** `HTMLResponse`, `RedirectResponse`.  
**Ключевые зависимости:** FastAPI, сервисы, `core.exceptions.PlateParseError`.  
**Связи:** Подключение в `app/main.py:37`.  
**Паттерны:** Инлайн-HTML в Python-строках.

**Расположение:** `app/dependencies/auth.py:8-34`  
**Слой:** dependency  
**Что делает:** `get_auth_repository()` возвращает `AuthRepository()`. `get_current_user` читает cookie `app_session`, декодирует через `decode_session_token`, сопоставляет с `repository.list_users()`. `require_roles` возвращает зависимость, проверяющую роль.  
**Входы:** `Request`, опционально репозиторий.  
**Выходы:** `dict` пользователя или `HTTPException` 401/403.  
**Ключевые зависимости:** `AuthRepository`, `decode_session_token`.  
**Связи:** Используется в `app/web/router.py` и `app/api/v1/endpoints/auth.py` (частично).  
**Паттерны:** Замыкание для `require_roles`.

**Расположение:** `app/repositories/auth_repository.py:28-111`  
**Слой:** repository  
**Что делает:** SQLite по пути `plita_db_path` из настроек (или переданный `db_path`); `init_schema` создаёт таблицу `app_users`; `ensure_bootstrap_admin` при наличии `bootstrap_admin_username` и `bootstrap_admin_password` вставляет или обновляет пользователя; `authenticate` проверяет пароль PBKDF2.  
**Входы:** Строки username/password, путь к БД.  
**Выходы:** Словари пользователей или `None`.  
**Ключевые зависимости:** `sqlite3`, `get_settings`.  
**Связи:** Вызывается из lifespan (`app/main.py:19`), web-login, API auth.  
**Паттерны:** Хранение пароля как salt+digest в одной строке.

**Расположение:** `app/security/session.py:13-41`  
**Слой:** security  
**Что делает:** HMAC-SHA256 подпись payload с секретом `app_secret_key`; токен сессии как `base64(json).signature`; проверка срока `exp` в `decode_session_token`.  
**Входы:** Словарь данных пользователя, строка токена.  
**Выходы:** Строка токена или `dict` / `None`.  
**Ключевые зависимости:** `get_settings`.  
**Связи:** Web и API login устанавливают cookie `app_session`.  
**Паттерны:** Самописный подписанный JWT-подобный токен в cookie.

**Расположение:** `bot/bot_config.py:1-19`  
**Слой:** configuration (бот)  
**Что делает:** Вызывает `get_settings()`, экспортирует `BOT_TOKEN`, пути `OUTPUTS_DIR`, `PRICES_DIR`, `DB_PATH` (pb), строковые алиасы; создаёт `OUTPUTS_DIR` на диске.  
**Входы:** Те же env/settings, что и backend.  
**Выходы:** Константы для импорта в хендлерах.  
**Ключевые зависимости:** `app.core.settings`.  
**Связи:** Импортируется из `bot/bot_main.py`.  
**Паттерны:** Общие настройки приложения и бота через один `Settings`.

**Расположение:** `bot/bot_main.py:7-87`  
**Слой:** bot entrypoint  
**Что делает:** Добавляет корень проекта в `sys.path`; настраивает логирование; `init_database()` логирует пути к `DB_PATH_STR` и `plita.db`, предупреждает если `pb.db` нет, не прерывает запуск; в `main()` проверяет `BOT_TOKEN` на пустоту/заглушку; создаёт `Bot` и `Dispatcher`, `register_all_handlers(dp)`, `dp.start_polling(bot)`; при `__main__` вызывает `asyncio.run(main())`.  
**Входы:** Переменная окружения/настройка `BOT_TOKEN`.  
**Выходы:** Долгоживущий polling-процесс.  
**Ключевые зависимости:** aiogram 3, `bot.handlers.register_all_handlers`, `core.kp_db` (импорт на уровне модуля).  
**Связи:** Запуск: `python bot/bot_main.py` из корня (или модуль с корнем в PYTHONPATH).  
**Паттерны:** Явное сообщение об отсутствии токена в `bot/bot.env` в логах (`bot/bot_main.py:47-48`).

**Расположение:** `bot/handlers/__init__.py:11-31`  
**Слой:** bot handler registration  
**Что делает:** Функция `register_all_handlers(dp)` последовательно подключает роутеры из подпакета `handlers` (main, instructions, kp, comparison, commercial, archive, production_*, work_calendar_manager, pb_info, export, admin).  
**Входы:** `Dispatcher`.  
**Выходы:** Побочный эффект — зарегистрированные хендлеры.  
**Ключевые зависимости:** Подмодули `bot.handlers.*`.  
**Связи:** Вызывается из `bot/bot_main.py:69`.  
**Паттерны:** Порядок `include_router` зафиксирован комментарием «в правильном порядке».

**Расположение:** `core/logging_config.py:9-26`  
**Слой:** core utility  
**Что делает:** Настраивает корневой логгер: консоль и ротационный файл в `log_dir` (по умолчанию `logs/` в корне проекта); если хендлеры уже есть — только уровень.  
**Входы:** Уровень логирования, опционально каталог и имя файла.  
**Выходы:** Настроенное логирование.  
**Ключевые зависимости:** `logging`, `RotatingFileHandler`.  
**Связи:** `app/main.py:18`, `bot/bot_main.py:23`.  
**Паттерны:** Идемпотентность при повторном вызове через проверку `root_logger.handlers`.

## Поток данных

- **HTTP API:** клиент → `FastAPI` (`app/main.py:26-37`) → `app/api/v1/router.py` → конкретный endpoint в `app/api/v1/endpoints/*.py` → сервисы/репозитории → SQLite/файлы/ответ JSON.

- **HTTP Web:** клиент → `app/web/router.py` → `get_current_user` / `require_roles` (`app/dependencies/auth.py`) → `AuthRepository` / сервисы → HTML или редирект; логин POST → `AuthRepository.authenticate` → `create_session_token` → cookie `app_session` (`app/web/router.py:170-178`).

- **Lifespan backend:** старт ASGI → `lifespan` (`app/main.py:16-21`) → `setup_logging` → `ensure_bootstrap_admin` → запись/обновление пользователя в SQLite `plita.db` при заданных `APP_ADMIN_*`.

- **Telegram:** Telegram API → aiogram `Dispatcher` (`bot/bot_main.py:66-75`) → зарегистрированные роутеры (`bot/handlers/__init__.py`) → хендлеры → при необходимости `core` / БД по путям из `bot_config` / `settings`.

## Ссылки на код

- `app/main.py:16-21` — lifespan: логи, bootstrap админ, `app.state.settings`
- `app/main.py:32-37` — `/health`, подключение API и web
- `app/main.py:41` — экспорт `app` для uvicorn
- `app/core/settings.py:14-16` — загрузка `.env` и `bot/bot.env`
- `app/core/settings.py:18-59` — поля настроек и alias переменных окружения
- `app/api/v1/router.py:7-12` — сборка API v1
- `app/web/router.py:16` — web-router без OpenAPI schema
- `app/dependencies/auth.py:17-24` — cookie `app_session` и проверка пользователя
- `app/repositories/auth_repository.py:34-56` — схема `app_users`
- `app/repositories/auth_repository.py:58-87` — `ensure_bootstrap_admin`
- `app/security/session.py:19-26` — создание токена сессии
- `bot/bot_config.py:7-12` — токен и пути из `get_settings()`
- `bot/bot_main.py:44-56` — проверка `BOT_TOKEN` перед polling
- `bot/bot_main.py:74-75` — `start_polling`
- `bot/handlers/__init__.py:11-31` — регистрация всех роутеров бота
- `requirements.txt:1-13` — ключевые зависимости (aiogram, fastapi, uvicorn и др.)

## Архитектурные наблюдения

- Две точки входа процессов: ASGI-приложение `app.main:app` и скрипт `bot/bot_main.py` с `asyncio.run(main())`.
- Конфигурация централизована в `app.core.settings.Settings`; бот читает токен и пути через `bot/bot_config.py`, опираясь на те же настройки.
- Пользователи веб-интерфейса и API-сессии завязаны на SQLite-файл `plita.db` (по умолчанию в корне проекта), таблица `app_users`.
- API v1 монтируется с префиксом `/api/v1`; HTML-маршруты используют префикс пути `/web`.
- В `.gitignore` игнорируются `bot.env`, `*.env`, файлы `*.db`, каталог `logs/`, что влияет на то, какие файлы должны существовать локально у разработчика, но не хранятся в репозитории.

## Подготовка к запуску (зафиксированные шаги)

Ниже перечислено то, что в репозитории и в коде предполагается для локального запуска; факт выполнения `pip install` и короткого старта uvicorn на машине агента: успешно (импорт `app.main:app`, `timeout 3 uvicorn app.main:app`).

1. **Виртуальное окружение и зависимости** (из корня репозитория `plan_wed_web`):

```bash
cd /home/username/Code/plan_wed_web
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. **Файлы конфигурации окружения:** в коде загружаются `PROJECT_ROOT/.env` и `bot/bot.env` (`app/core/settings.py:14-15`, `app/core/settings.py:19-20`). В `.gitignore` указаны `bot.env` и `*.env` (`.gitignore:41-43`), то есть при клонировании репозитория эти файлы отсутствуют в git и создаются вручную локально.

3. **Минимальные переменные для бота:** `BOT_TOKEN` должен быть непустым и не заглушкой (`bot/bot_main.py:46-56`); источник — настройки (`Settings.bot_token`, alias `BOT_TOKEN` в `app/core/settings.py:33`).

4. **Минимальные переменные для первого входа в web:** при старте backend вызывается `ensure_bootstrap_admin` (`app/main.py:19`); для создания/обновления администратора нужны `APP_ADMIN_USERNAME` и `APP_ADMIN_PASSWORD` (`app/repositories/auth_repository.py:58-63`, поля в `app/core/settings.py:57-58`). Секрет подписи cookie: `APP_SECRET_KEY` (`app/core/settings.py:28-31`, `app/security/session.py:14-16`).

5. **Базы данных:** по умолчанию `plita.db` и `pb.db` в корне проекта (`app/core/settings.py:38-39`). `AuthRepository` использует `plita.db` (`app/repositories/auth_repository.py:29-31`). Бот при отсутствии `pb.db` пишет предупреждение и продолжает работу (`bot/bot_main.py:34-38`).

6. **Запуск backend:**

```bash
cd /home/username/Code/plan_wed_web
source .venv/bin/activate
uvicorn app.main:app --host 127.0.0.1 --port 8000
```

Проверка готовности: `GET /health` (`app/main.py:32-34`) и `GET /api/v1/health` через подроутер (`app/api/v1/endpoints/health.py:10-17`).

7. **Запуск бота** (из корня, с активированным venv):

```bash
cd /home/username/Code/plan_wed_web
source .venv/bin/activate
python bot/bot_main.py
```

8. **Логи:** backend при lifespan пишет в файл с именем `backend.log` в `settings.logs_dir` (`app/main.py:18`); общая утилита логирования по умолчанию использует каталог `logs` под корнем (`core/logging_config.py:22-26`).
