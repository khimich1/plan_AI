---
title: "Полное ревью кодовой базы"
date: 2026-05-07
scope: "Репозиторий plan_wed_web (app, bot, core, frontend, viz_modules, scripts, tests)"
model: "composer-2-fast"
source: "Команда /review — subagent reviewer (полный обзор)"
---

## Краткий обзор проекта

- **Назначение:** веб-бэкенд (FastAPI) для коммерческих предложений, производственных планов и архива; **Telegram-бот** (Aiogram 3) на общей бизнес-логике из `core/`; **React SPA** в `frontend/` с прокси на API; расчёты/визуализация — `viz_modules/`, `app/planning/`.
- **Стек:** Python 3, FastAPI, Pydantic v2 / pydantic-settings, SQLite (`plita.db`, `pb.db`), filelock, uvicorn; фронт — TypeScript/React (по структуре).
- **Точки входа:** `app/main.py` → `create_app()` (`/api/v1`, `/web`, `/health`); ASGI-обёртка `main.py` импортирует `app`; бот — `bot/bot_main.py` (`start_polling`).

---

## Выводы по категориям

### Critical

1. **Telegram-бот: нет явной авторизации на опасные действия**  
   Роутеры (в т.ч. `admin`) подключаются без общего фильтра по `allowed_user_ids` / ролям. Команда `/delete_kp` и прочие админ-сценарии доступны любому, кто может писать боту (если бот не изолирован иным способом снаружи репозитория).

```python
# bot/handlers/admin.py (фрагмент)
@router.message(Command("delete_kp"))
async def cmd_delete_kp(message: Message, state: FSMContext):
    """
    Удаляет КП из базы данных по номеру.
    Использование: /delete_kp 5
    """
    args = message.text.split()
    # ...
```

```python
# bot/handlers/__init__.py (фрагмент)
def register_all_handlers(dp: Dispatcher):
    """Регистрируем все роутеры в правильном порядке"""
    dp.include_router(main.router)
    # ...
    dp.include_router(admin.router)
```

2. **Cookie сессии без `Secure` в продакшене**  
   Утечка сессии при доступе по HTTP в смешанном окружении.

`app/api/v1/endpoints/auth.py` — `set_cookie(..., secure=False)`  
`app/web/router.py` — `login_submit` задаёт cookie без `secure`.

3. **Слабый дефолт секрета подписи сессии**  
   Если в проде не задать `APP_SECRET_KEY`, токены предсказуемы по умолчанию — `app/core/settings.py`: `app_secret_key` default `"change-this-secret-key-in-env"`.

*(Связка: HMAC-подпись в `app/security/session.py` корректна, но сила — только в секрете.)*

---

### Quality

1. **Аутентификация: полный список пользователей на каждый запрос** — `app/dependencies/auth.py` (`get_current_user` + `repository.list_users()`).
2. **Лимит OCR-загрузок только внутри процесса** — см. `app/services/commercial_upload_validation.py`.
3. **Несогласованность путей планов и настроек** — `Settings.plans_dir` vs жёстко `bot/data/plans` в `app/planning/plan_manager.py`.
4. **Health-эндпоинт раскрывает окружение** — `app/api/v1/endpoints/health.py`.
5. **Роль `production` и доступ к КП через API** — возможное избыточное раскрытие по бизнес-ролям — `app/api/v1/endpoints/offers.py`.
6. **Нет rate limiting на логин** — `/api/v1/auth/login` и `/web/login`.
7. **`ast.literal_eval` при разборе ключей плана** — `app/planning/plan_manager.py` (низкий риск для доверенных файлов).
8. **Публичный `/health` на корне** — `app/main.py`.

---

### Suggestion

- Привязать `Secure`, `domain` и срок cookie к `app_env` / настройкам; рассмотреть `SameSite=strict` для чисто same-site SPA.
- Для HTML-форм `/web/*` — CSRF-токены.
- Вынести `get_user_by_id` в `AuthRepository` вместо `list_users()` в `get_current_user`.
- Общий лимитер OCR (Redis или sticky + IP) при горизонтальном масштабе.
- Унифицировать `plans_dir` / метаданные планов с `Settings` для деплоя.
- Фронт: минимизировать чувствительные данные в `sessionStorage` / черновиках при строгих требованиях.

---

## Позитивные моменты

- **Черновики:** валидация `draft_id`, путь под `drafts_dir`, файловые блокировки, атомарная запись JSON (`app/services/draft_store.py`).
- **Скачивание сгенерированных файлов:** basename, whitelist метаданных, проверка родителя `outputs_dir` (`app/api/v1/endpoints/commercial.py`).
- **Загрузки OCR:** лимит размера, magic bytes (`app/services/commercial_upload_validation.py`).
- **Планы производства:** валидация `plan_id` и защита каталога (`app/planning/plan_manager.py`).
- **Пароли:** PBKDF2-SHA256, соль, `hmac.compare_digest` (`app/repositories/auth_repository.py`).
- **Сессия:** HMAC-подпись, `compare_digest` (`app/security/session.py`).
- **Ошибки коммерции:** обобщённые сообщения клиенту, детали в логах (`app/core/http_errors.py`).
- **HTML кабинет:** `html.escape` (`app/web/router.py`).
- **Настройки:** предупреждение про `APP_STORAGE_LAYOUT` в проде (`app/core/settings.py`).

---

## Тесты и CI

- **Структура:** в `tests/` — множество модулей (оптимизация, планы, парсер, КП, архив, интеграционные сценарии).
- **Пробелы:** не видно GitHub Actions / GitLab CI на корне; тест сессии — без негативных сценариев подделки; смешение утилит и тестов в `tests/`.

---

## Итоговая таблица

| Уровень     | Количество |
|------------|------------|
| Critical   | **3**      |
| Quality    | **8**      |
| Suggestion | **5**      |

---

## Резюме

Архитектура в целом взрослая (слои API → сервисы → репозитории, защита файлов и черновиков продумана). Главные риски для продакшена — **отсутствие авторизации опасных команд бота**, **жёстко выключенный `secure` у cookie** и **обязательность смены `APP_SECRET_KEY`**. Остальное — масштабирование аутентификации, лимиты, согласованность путей и политика ролей для API КП.

---

**Примечание:** отчёт подготовлен в режиме обзора без автоматических правок в коде.
