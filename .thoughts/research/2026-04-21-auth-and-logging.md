---
date: 2026-04-21
topic: Логирование и авторизация (логин/пароль)
scope:
  - backend
  - frontend
  - core
---

# Исследование: Логирование и авторизация (логин/пароль)

## Резюме
В проекте логирование настраивается централизованно через `core/logging_config.py`: root logger пишет в консоль и в ротируемый файл. Backend включает логирование в `logs/backend.log`, а бот использует тот же механизм с файлом по умолчанию `logs/bot.log`. Логин реализован через `POST /api/v1/auth/login` или web-форму `POST /web/login`, после чего в cookie `app_session` кладется подписанный сессионный токен. Проверка логина/пароля выполняется по хэшу пароля в SQLite-таблице `app_users`.

## Подробные находки

**Расположение:** `core/logging_config.py:9-60`  
**Слой:** core  
**Что делает:** настраивает единый logging для процесса: formatter, console handler и `RotatingFileHandler`.  
**Входы:** `level`, `log_dir`, `log_filename`.  
**Выходы:** глобально настроенный root logger для всего процесса.  
**Ключевые зависимости:** `logging`, `RotatingFileHandler`, `sys`, `pathlib.Path`.  
**Связи:** вызывается из backend entrypoint и bot entrypoint.  
**Паттерны:** единая конфигурация логов в core + переиспользование в разных процессах.

**Расположение:** `app/main.py:25-27` и `app/core/settings.py:57,90-93`  
**Слой:** backend entrypoint  
**Что делает:** на старте backend вызывает `setup_logging(..., log_filename="backend.log")`, директория логов берется из settings (`logs_dir`).  
**Входы:** `settings.logs_dir`.  
**Выходы:** файл `logs/backend.log` и вывод в консоль backend-процесса.  
**Ключевые зависимости:** `core.logging_config.setup_logging`, `app.core.settings.get_settings`.  
**Связи:** lifespan FastAPI инициализирует логи до обслуживания запросов.  
**Паттерны:** инфраструктурная инициализация через lifespan.

**Расположение:** `bot/bot_main.py:22-24` и `core/logging_config.py:12`  
**Слой:** bot entrypoint  
**Что делает:** вызывает `setup_logging(level=logging.INFO)` без имени файла, поэтому используется дефолтный `bot.log`.  
**Входы:** уровень INFO.  
**Выходы:** файл `logs/bot.log` и консольные логи бота.  
**Ключевые зависимости:** `core.logging_config.setup_logging`.  
**Связи:** инициализация перед запуском polling aiogram.  
**Паттерны:** общий logging-конфиг для backend и bot.

**Расположение:** `start_bot.sh:12-16,109,118`  
**Слой:** script / runtime  
**Что делает:** дополнительно перенаправляет stdout/stderr backend и frontend в timestamp-файлы в `logs/`.  
**Входы:** текущая дата/время, команды запуска backend/frontend.  
**Выходы:** `logs/backend_*.log`, `logs/frontend_*.log`.  
**Ключевые зависимости:** shell redirection, `nohup`.  
**Связи:** параллельно с python logging даёт отдельные launch-логи.  
**Паттерны:** файловый аудит запуска через shell.

**Расположение:** `app/api/v1/endpoints/auth.py:10-45`, `app/schemas/auth.py:6-8`  
**Слой:** backend endpoint + backend schema  
**Что делает:** реализует login/logout/me API; login принимает `username/password` и при успехе ставит cookie `app_session`.  
**Входы:** JSON `LoginRequest(username, password)`.  
**Выходы:** `{"user": ...}` или 401 при невалидных учетных данных.  
**Ключевые зависимости:** `AuthRepository.authenticate`, `create_session_token`, `Depends(get_current_user)`.  
**Связи:** endpoint передает проверку в repository и использует security-модуль для токена.  
**Паттерны:** auth endpoint -> repository -> session cookie.

**Расположение:** `app/repositories/auth_repository.py:13-16,19-25,39-47,131-140`  
**Слой:** backend repository  
**Что делает:** создает/читает таблицу `app_users`, хранит `password_hash`, проверяет пароль через PBKDF2-HMAC-SHA256 и `hmac.compare_digest`.  
**Входы:** `username`, `password`, путь к SQLite БД (`settings.plita_db_path` по умолчанию).  
**Выходы:** payload пользователя без `password_hash` или `None`.  
**Ключевые зависимости:** `sqlite3`, `hashlib.pbkdf2_hmac`, `hmac`, `get_settings`.  
**Связи:** используется auth endpoints и скриптом создания администратора.  
**Паттерны:** repository управляет схемой БД и auth-операциями.

**Расположение:** `app/security/session.py:13-42`  
**Слой:** backend security  
**Что делает:** создает и валидирует signed session token формата `base64(payload).signature` с полем `exp`.  
**Входы:** payload пользователя, `APP_SECRET_KEY`, token из cookie.  
**Выходы:** строка токена или `None` при невалидной подписи/сроке.  
**Ключевые зависимости:** `hmac`, `hashlib.sha256`, `base64`, `json`, `time`.  
**Связи:** вызывается из login endpoint и dependency `get_current_user`.  
**Паттерны:** cookie-based signed session.

**Расположение:** `app/dependencies/auth.py:13-34`  
**Слой:** backend dependency  
**Что делает:** достает `app_session` из cookie, декодирует токен, подтягивает пользователя, проверяет `is_active`, ограничивает роли через `require_roles`.  
**Входы:** `Request.cookies`, repository, набор допустимых ролей.  
**Выходы:** user dict или HTTP 401/403.  
**Ключевые зависимости:** `decode_session_token`, `AuthRepository.list_users`, `fastapi.Depends`.  
**Связи:** используется защищенными API и web-роутами.  
**Паттерны:** dependency-based auth в FastAPI.

**Расположение:** `app/web/router.py:179-204`  
**Слой:** backend web (HTML)  
**Что делает:** отдает страницу логина и обрабатывает web-форму входа; при успехе ставит ту же cookie `app_session`.  
**Входы:** form fields `username/password`.  
**Выходы:** redirect на `/web/` при успехе или редирект обратно с `?error=...`.  
**Ключевые зависимости:** `AuthRepository.authenticate`, `create_session_token`.  
**Связи:** альтернативный вход в ту же систему auth, что и API.  
**Паттерны:** единый backend auth для API и server-rendered web.

**Расположение:** `frontend/src/shared/api/httpClient.ts:31-37`, `frontend/src/app/router/AppRouter.tsx:10-12`  
**Слой:** frontend API client + frontend routing  
**Что делает:** frontend запросы идут с `credentials: "include"` (браузер отправляет cookie), отдельной SPA-страницы логина в текущем роутере нет.  
**Входы:** `VITE_API_BASE_URL`.  
**Выходы:** HTTP-запросы к backend с cookie-сессией.  
**Ключевые зависимости:** Fetch API, env config.  
**Связи:** frontend опирается на backend cookie-auth.  
**Паттерны:** frontend без local token storage, с cookie include.

**Расположение:** `.env:1-4`, `app/core/settings.py:15-23,29-32`, `scripts/create_admin.py:25-27,40-47,63-67`  
**Слой:** config + script  
**Что делает:** settings загружает `.env` и `bot/bot.env`; скрипт `create_admin.py` вручную создает/обновляет пользователя в `app_users` по введенному паролю.  
**Входы:** env-переменные, аргументы CLI (`--username`, `--role`), пароль через prompt.  
**Выходы:** настройки runtime, запись/обновление пользователя в БД.  
**Ключевые зависимости:** `dotenv`, `pydantic-settings`, `AuthRepository`.  
**Связи:** script работает через тот же repository, что и login endpoint.  
**Паттерны:** bootstrap пользователя отдельным CLI-скриптом.

## Поток данных

- `HTTP /web/login` или `POST /api/v1/auth/login`
- `backend endpoint/web route` (`app/api/v1/endpoints/auth.py` или `app/web/router.py`)
- `backend repository` (`AuthRepository.authenticate`)
- `sqlite app_users` + проверка `password_hash`
- `backend security` (`create_session_token`)
- `Set-Cookie app_session`
- `следующие запросы` -> `get_current_user` -> `decode_session_token` -> доступ к защищенным endpoint

- `backend/bot startup`
- `setup_logging` из `core/logging_config.py`
- `console handler + rotating file handler`
- `logs/backend.log` или `logs/bot.log`
- `дополнительно при shell-запуске` -> `logs/backend_*.log`, `logs/frontend_*.log`

## Ссылки на код

- `core/logging_config.py:9-60` — единая конфигурация логирования  
- `app/main.py:25-27` — подключение backend logging  
- `bot/bot_main.py:22-24` — подключение bot logging  
- `start_bot.sh:12-16,109,118` — shell-логи запуска backend/frontend  
- `app/api/v1/endpoints/auth.py:10-45` — login/logout/me API  
- `app/web/router.py:179-204` — web-форма логина  
- `app/repositories/auth_repository.py:13-16,19-25,39-47,131-140` — хэширование/проверка пароля и таблица пользователей  
- `app/security/session.py:13-42` — создание/проверка session token  
- `app/dependencies/auth.py:13-34` — текущий пользователь и проверка ролей  
- `app/core/settings.py:15-23,29-32` — загрузка env и секрета подписи  
- `frontend/src/shared/api/httpClient.ts:31-37` — `credentials: "include"`  
- `.env:2-3` — текущие значения `APP_ADMIN_USERNAME` и `APP_ADMIN_PASSWORD`  
- `scripts/create_admin.py:25-27,40-47,63-67` — ручное создание/обновление admin

## Архитектурные наблюдения

- Backend использует cookie-based auth с FastAPI dependencies (`Depends`) для проверки пользователя и ролей.  
- Password verification делается на уровне repository через `password_hash` в SQLite, а не по plaintext-полю.  
- Секрет подписи сессии берется из `APP_SECRET_KEY` через централизованный settings.  
- Frontend работает с backend-сессией через cookie (`credentials: "include"`), без явного token storage в SPA-коде.  
- Логирование централизовано в core-модуле и переиспользуется в backend и bot.
