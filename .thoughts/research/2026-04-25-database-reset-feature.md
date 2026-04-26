---
date: 2026-04-25
topic: Перенос фичи «Очистка БД» из Telegram-бота в веб-приложение (полное и частичное обнуление)
scope:
  - app
  - bot
  - frontend
  - core
---

# Исследование: фича «Очистка базы данных» (полная и частичная)

## Резюме

В Telegram-боте уже реализован полный набор административных действий по очистке базы (раздел `⚙️ Управление БД` + slash-команды). Логика стоит на трёх уровнях: UI-кнопки в `bot/keyboards.py`, обработчики в `bot/handlers/admin.py` (FSM-состояния `bot/states.py:29-32`) и низкоуровневые SQL-функции в `core/kp_db.py`. На стороне FastAPI уже есть «частичные» удаления (одно КП через `DELETE /api/v1/commercial/archive/{kp_id}`, один план через `DELETE /api/v1/production/plans/{plan_id}`), но **массовой очистки нет ни в API, ни в SPA** — frontend (React/Vite) пока знает только три раздела навигации (`Создать КП`, `Архив`, `Производство`) и не имеет admin-страницы или admin-меню. В системе уже есть ролевая модель `admin / manager / production` (`app/dependencies/auth.py:28-34`), а данные хранятся одновременно в SQLite (`plita.db`) и в JSON-файлах (`bot/data/plans/`, `plans_metadata.json`, `current_plan.json`, `work_calendar.json`).

## Подробные находки

### 1. Бот: меню «Управление БД» и роутер

Расположение: `bot/handlers/admin.py:1-695`
Слой: bot handler (aiogram Router)
Что делает: единый `Router` с командами и callback'ами для администрирования.

Команды и кнопки:
- `/delete_kp <id>` — удаление одного КП с подтверждением через FSM (`bot/handlers/admin.py:25-128`).
- `/list_kp` — список всех КП по статусам (`bot/handlers/admin.py:131-183`).
- `/clear_all_kp` — старая команда «полная очистка КП» с подтверждением через FSM (`bot/handlers/admin.py:186-288`). Использует `kp_db.clear_all_kp` (только таблицы КП).
- Кнопка `⚙️ Управление БД` из `main_menu_kb` — открывает inline-меню (`bot/handlers/admin.py:293-301`).
- Callback `db_stats` → `kp_db.get_db_stats` → выводит сводку (`bot/handlers/admin.py:304-333`).
- Callback `db_clear_all` → подсчёт записей + предупреждение (`bot/handlers/admin.py:336-383`).
- Callback `db_clear_confirmed` → выполняет «суперполную» очистку через `kp_db.clear_all_plates_data` + локальную `clear_all_plans_data()` (`bot/handlers/admin.py:439-483`).
- Callback `db_clear_cancel` / `db_back_to_menu` — отмена/возврат (`bot/handlers/admin.py:486-503`).
- Callback `db_view_rests` — экспорт остатков в XLSX через openpyxl (`bot/handlers/admin.py:574-694`).
- `/recover_plates` — восстановление «застрявших» плит (статус `'в плане'` → `'в производстве'`) через `kp_db.recover_stuck_plates` (`bot/handlers/admin.py:506-571`).

Регистрация: `bot/handlers/__init__.py:11-32` — `dp.include_router(admin.router)` подключается последним.

Локальная функция `clear_all_plans_data()` (`bot/handlers/admin.py:386-436`) удаляет:
- `bot/data/current_plan.json`
- `bot/data/plans_metadata.json`
- всю папку `bot/data/plans/` через `shutil.rmtree`

Возвращает счётчики `{'current_plan', 'metadata', 'plan_files', 'total'}`.

### 2. Бот: клавиатуры

Расположение: `bot/keyboards.py:5-26` (`main_menu_kb`), `bot/keyboards.py:559-575` (`db_management_kb`), `bot/keyboards.py:578-590` (`db_clear_confirm_kb`).
Слой: bot keyboards
Что делает: формирует aiogram `ReplyKeyboardMarkup` / `InlineKeyboardMarkup`.

`main_menu_kb` содержит reply-кнопку `"⚙️ Управление БД"` в третьей строке (`bot/keyboards.py:21-23`).

`db_management_kb` (`bot/keyboards.py:568-575`):
- `🗑️ Очистить все данные` → `callback_data="db_clear_all"`
- `📋 Просмотр остатков` → `callback_data="db_view_rests"`
- `📊 Статистика БД` → `callback_data="db_stats"`
- `◀️ Назад` → `callback_data="db_back_to_menu"`

`db_clear_confirm_kb` (`bot/keyboards.py:585-590`):
- `⚠️ ДА, УДАЛИТЬ ВСЁ` → `callback_data="db_clear_confirmed"`
- `❌ Отмена` → `callback_data="db_clear_cancel"`

### 3. Бот: FSM-состояния

Расположение: `bot/states.py:29-32`
Слой: bot states
Что делает: класс `AdminStates(StatesGroup)` с двумя состояниями:
- `waiting_delete_confirmation` — после `/delete_kp <id>`, ждёт ввода `ДА`/`НЕТ` (`bot/handlers/admin.py:86-128`).
- `waiting_clear_all_confirmation` — после `/clear_all_kp`, ждёт `ДА`/`НЕТ` (`bot/handlers/admin.py:241-288`).

Кнопочное меню `db_management_kb` подтверждение через FSM не использует — оно реализовано через два callback-обработчика `db_clear_all` (показ предупреждения) → `db_clear_confirmed` (выполнение).

### 4. Core: низкоуровневые функции SQLite

Расположение: `core/kp_db.py`
Слой: core/legacy
Что делает: прямая работа с `plita.db` через `sqlite3`. Все функции принимают `db_path: str = DEFAULT_DB`.

Функции, относящиеся к очистке:

| Функция | Строки | Действие |
|---|---|---|
| `delete_kp_by_id(kp_id)` | `core/kp_db.py:953-994` | Удаляет одно КП. Включает `PRAGMA foreign_keys = ON`, `DELETE FROM KP_offers WHERE kp_id = ?`. CASCADE в схеме автоматически чистит связанные таблицы. Возвращает `bool`. |
| `clear_all_kp(db_path)` | `core/kp_db.py:1133-1203` | «Частичная» очистка: только таблицы `KP_offers`, `kp_plates`, `kp_files`, `kp_meta`. Сбрасывает `sqlite_sequence` для этих таблиц. Возвращает `dict[str,int]` со счётчиками удалённого. **НЕ трогает** `completed_plates` и `plate_rests`. |
| `clear_all_plates_data(db_path)` | `core/kp_db.py:997-1087` | «Полная» очистка: все таблицы — `KP_offers`, `kp_plates`, `completed_plates`, `plate_rests`, `kp_files`, `kp_meta` + сброс `sqlite_sequence` для всех. Возвращает `dict` с раздельным счётом и `total`. Делает `conn.rollback()` при ошибке и пробрасывает исключение. |
| `get_db_stats(db_path)` | `core/kp_db.py:1090-1130` | Возвращает `{kp_total, kp_in_work, kp_completed, plates_in_work, plates_completed, plate_rests}`. |
| `get_all_plate_rests(db_path)` | `core/kp_db.py:1885-...` | Список остатков для XLSX-экспорта. |
| `recover_stuck_plates(db_path)` | `core/kp_db.py:2461-2519` | UPDATE `kp_plates SET status='в производстве', plan_id=NULL WHERE status='в плане'`. Возвращает количество восстановленных записей. |

Различие двух «полных» очисток:
- `clear_all_kp` — старая команда `/clear_all_kp` (без `completed_plates` и `plate_rests`).
- `clear_all_plates_data` — новая кнопка `🗑️ Очистить все данные` (полная зачистка всех таблиц с плитами).

### 5. Backend: репозитории, через которые уже идёт удаление

Расположение: `app/repositories/kp_repository.py:76-77`
Слой: repository
Что делает: `KpRepository.delete_offer(kp_id)` → `kp_db.delete_kp_by_id(kp_id, self.db_path)`.

Расположение: `app/repositories/kp_archive_repository.py:33-34`
Слой: repository
Что делает: `KpArchiveRepository.delete(kp_id)` → `kp_db.delete_kp_by_id(kp_id, self.db_path)`.

Расположение: `app/repositories/plan_repository.py:26-27`
Слой: repository
Что делает: `PlanRepository.delete_plan(plan_id)` → `bot.handlers.plan_manager.delete_plan(plan_id)`. Бэкенд импортирует `bot.handlers.plan_manager` напрямую (см. `app/services/production_service.py:12`), поэтому JSON-планы тоже находятся под управлением backend.

Расположение: `app/repositories/auth_repository.py:28-150`
Слой: repository
Что делает: работает с таблицей `app_users` в `plita.db`. **Эта таблица НЕ затрагивается** ни одной из существующих очисток в `core/kp_db.py` — пользователи переживут любую «полную» зачистку.

### 6. Backend: где уже есть «удаление» в API v1

Расположение: `app/api/v1/router.py:7-13`
Слой: router
Что делает: подключает routers `health`, `auth`, `managers`, `commercial`, `archive`, `production`. **Отдельного `admin` router нет.**

Расположение: `app/api/v1/endpoints/archive.py:130-140`
Слой: endpoint
Что делает: `DELETE /api/v1/commercial/archive/{kp_id}` с `Depends(require_roles("admin", "manager"))`, возвращает `204 No Content`. Внутри `ArchiveService.delete_offer` (`app/services/archive_service.py:112-114`).

Расположение: `app/api/v1/endpoints/production.py:87-95`
Слой: endpoint
Что делает: `DELETE /api/v1/production/plans/{plan_id}` с `Depends(require_roles("admin", "production"))`, возвращает `DeletePlanResponse`. Внутри `ProductionService.delete_plan` (`app/services/production_service.py:35-37`) → `PlanRepository.delete_plan` → `plan_manager.delete_plan`.

Расположение: `app/services/archive_service.py:112-114`
Слой: service
Что делает: `delete_offer(kp_id)` → `repository.delete(kp_id)`; кидает `ArchiveNotFoundError`, если КП не существовал.

### 7. Backend: ролевая модель и сессии

Расположение: `app/dependencies/auth.py:13-34`
Слой: dependency
Что делает: `get_current_user` читает cookie `app_session` (`app/dependencies/auth.py:17`), декодирует через `decode_session_token` (`app/security/session.py`), сверяет с `AuthRepository.list_users()`. `require_roles(*allowed_roles)` — фабрика зависимости, возвращает 403, если роль не в списке.

Существующие роли в кодовой базе: `admin`, `manager`, `production` (используются во всех endpoint'ах и в шаблоне `app/web/router.py:223-840`).

Расположение: `app/repositories/auth_repository.py:34-56`
Слой: repository
Что делает: создаёт таблицу `app_users(id, username, password_hash, role, manager_id, is_active, created_at)` при `init_schema()`. Хеширование — PBKDF2-SHA256 c 200000 итераций (`app/repositories/auth_repository.py:13-25`).

### 8. Backend: настройки и пути к файлам

Расположение: `app/core/settings.py:19-95`
Слой: settings
Что делает: `Settings` (pydantic-settings) описывает пути:
- `plita_db_path = PROJECT_ROOT / "plita.db"` (строка 45) — все таблицы КП и `app_users`.
- `plans_dir = PROJECT_ROOT / "bot" / "data" / "plans"` (строка 54).
- `plans_metadata_path = PROJECT_ROOT / "bot" / "data" / "plans_metadata.json"` (строка 55).
- `work_calendar_path = PROJECT_ROOT / "bot" / "data" / "work_calendar.json"` (строка 56).

`current_plan.json` (используется в `bot/handlers/admin.py:409-412`) **не описан в `Settings`** — путь захардкожен в `BOT_DIR / 'data' / 'current_plan.json'`.

`ensure_directories()` (`app/core/settings.py:90-95`) создаёт `outputs_dir`, `logs_dir`, `plans_dir`, родителя `work_calendar_path`, `drafts_dir`.

### 9. Frontend: роутинг, навигация, существующие удаления

Расположение: `frontend/src/app/router/AppRouter.tsx:1-18`
Слой: frontend app
Что делает: `BrowserRouter` с тремя страницами под `AppLayout`:
- `/commercial-offer/new` → `CommercialOfferCreatePage`
- `/commercial-offer/archive` → `CommercialOfferArchivePage`
- `/production` → `ProductionPage`
- `*` → `Navigate` на `/commercial-offer/new`

**Admin-маршрута нет.**

Расположение: `frontend/src/app/layout/AppHeader.tsx:70-96`
Слой: frontend layout
Что делает: рендерит навигацию с тремя ссылками (`Создать КП`, `Архив`, `Производство`) через `NavLink`. Содержит модалку «У вас есть незавершённый черновик» через `frontend/src/shared/ui/Modal.tsx`.

Расположение: `frontend/src/features/commercial-archive/api/archiveApi.ts:29`
Слой: frontend feature/api
Что делает: `archiveApi.delete(kpId) → httpClient.delete<void>(\`${BASE}/${kpId}\`)`.

Расположение: `frontend/src/features/commercial-archive/hooks/useArchiveQueries.ts:60-69`
Слой: frontend feature/hook
Что делает: `useDeleteOfferMutation` через `useMutation` (TanStack Query); в `onSuccess` инвалидирует все ключи `archive`. Используется в `frontend/src/features/commercial-archive/components/DeleteConfirmDialog.tsx:1-52` (отрисовывает `Modal` + `Alert(tone="warning")` + `Button(variant="danger")`).

Расположение: `frontend/src/features/production/api/productionApi.ts:27-28`
Слой: frontend feature/api
Что делает: `productionApi.deletePlan(planId)`.

Расположение: `frontend/src/features/production/hooks/useProductionQueries.ts:87-95`
Слой: frontend feature/hook
Что делает: `useDeletePlanMutation` инвалидирует ключ `productionKeys.all`.

Расположение: `frontend/src/shared/api/httpClient.ts:31-48`, `93-106`
Слой: frontend shared
Что делает: `fetch`-обёртка с `credentials: "include"` (cookie `app_session` уходит автоматически), методы `get/post/put/patch/delete/download`. Ошибка → `ApiError` (`frontend/src/shared/lib/apiError.ts`). Базовый URL берётся из `frontend/src/shared/config/env.ts`.

Расположение: `frontend/src/shared/ui/Button.tsx:1-57`
Слой: frontend shared/ui
Что делает: компонент `Button` поддерживает `variant: "primary" | "secondary" | "ghost" | "danger"` — `danger`-стиль уже описан (`#fff1f2 / #b42318`), используется во всех модалках удаления.

Существующие переиспользуемые UI-блоки: `Modal`, `Alert`, `Card`, `Button`, `Spinner`, `Field`, `StepLayout`, `Drawer` (см. `frontend/src/shared/ui/`).

### 10. Web-роутер (HTML), интерфейс к удалению одного КП

Расположение: `app/web/router.py:763-774`
Слой: web HTML
Что делает: inline-JS в HTML-странице `/web/managers` (`app/web/router.py:223-840`) кнопкой «Удалить» дёргает `DELETE /api/v1/offers/{id}` через native `confirm()`. Доступно только при `canEdit = role in ("admin","manager")` (`app/web/router.py:305`).

`app/api/v1/endpoints/offers.py` — отдельный legacy-endpoint для одного КП (используется только из старого HTML-кабинета `/web/managers`); SPA пользуется `app/api/v1/endpoints/archive.py`.

## Поток данных

### Текущее состояние (как очистка работает в боте)

```
Telegram → "⚙️ Управление БД" (main_menu_kb)
        → bot/handlers/admin.py:293 (btn_db_management) → db_management_kb
        → callback "db_clear_all" (admin.py:336)
            → core.kp_db.get_db_stats → текст-предупреждение
            → db_clear_confirm_kb
        → callback "db_clear_confirmed" (admin.py:439)
            → core.kp_db.clear_all_plates_data(plita.db)
                  → DELETE FROM kp_plates / completed_plates / plate_rests / kp_files / kp_meta / KP_offers
                  → DELETE FROM sqlite_sequence WHERE name IN (...)
            → bot/handlers/admin.clear_all_plans_data()
                  → unlink current_plan.json, plans_metadata.json
                  → shutil.rmtree(bot/data/plans/)
            → отчёт пользователю
```

### Текущее состояние (одиночное удаление КП в SPA)

```
React (frontend/src/features/commercial-archive/components/DeleteConfirmDialog.tsx)
  → useDeleteOfferMutation (useArchiveQueries.ts:60)
  → archiveApi.delete (archiveApi.ts:29)
  → httpClient.delete (shared/api/httpClient.ts:101) с credentials:"include"
  → DELETE /api/v1/commercial/archive/{kp_id} (app/api/v1/endpoints/archive.py:130)
  → require_roles("admin","manager") (app/dependencies/auth.py:28)
  → ArchiveService.delete_offer (app/services/archive_service.py:112)
  → KpArchiveRepository.delete (app/repositories/kp_archive_repository.py:33)
  → core.kp_db.delete_kp_by_id (core/kp_db.py:953)
  → SQLite plita.db (CASCADE)
```

### Текущее состояние (одиночное удаление плана в SPA)

```
React (frontend/src/features/production/components/PlansList.tsx → useDeletePlanMutation)
  → productionApi.deletePlan (productionApi.ts:27)
  → DELETE /api/v1/production/plans/{plan_id} (app/api/v1/endpoints/production.py:87)
  → require_roles("admin","production")
  → ProductionService.delete_plan (app/services/production_service.py:35)
  → PlanRepository.delete_plan (app/repositories/plan_repository.py:26)
  → bot.handlers.plan_manager.delete_plan
  → удаление JSON-файла плана + правка plans_metadata.json
```

### Чего ещё нет в backend/SPA

- Нет endpoint'а массовой очистки (нет аналогов `clear_all_kp`, `clear_all_plates_data`, `clear_all_plans_data` в HTTP API).
- Нет endpoint'а статистики (`get_db_stats`).
- Нет endpoint'а восстановления застрявших плит (`recover_stuck_plates`).
- Нет admin-страницы и admin-навигации в SPA.
- Нет admin-router'а в `app/api/v1/router.py:7-13`.
- Нет сервиса `app/services/admin_service.py`.

## Ссылки на код

### Бот
- `bot/handlers/__init__.py:31` — регистрация `admin.router`.
- `bot/handlers/admin.py:25-128` — `/delete_kp` + FSM-подтверждение.
- `bot/handlers/admin.py:131-183` — `/list_kp`.
- `bot/handlers/admin.py:186-288` — `/clear_all_kp` + FSM-подтверждение.
- `bot/handlers/admin.py:293-301` — обработчик кнопки `⚙️ Управление БД`.
- `bot/handlers/admin.py:304-333` — callback `db_stats`.
- `bot/handlers/admin.py:336-383` — callback `db_clear_all` (предупреждение).
- `bot/handlers/admin.py:386-436` — `clear_all_plans_data()` (JSON-планы).
- `bot/handlers/admin.py:439-483` — callback `db_clear_confirmed` (полная очистка).
- `bot/handlers/admin.py:486-503` — `db_clear_cancel`, `db_back_to_menu`.
- `bot/handlers/admin.py:506-571` — `/recover_plates`.
- `bot/handlers/admin.py:574-694` — callback `db_view_rests` (XLSX-отчёт).
- `bot/keyboards.py:5-26` — `main_menu_kb` с кнопкой `⚙️ Управление БД`.
- `bot/keyboards.py:559-575` — `db_management_kb`.
- `bot/keyboards.py:578-590` — `db_clear_confirm_kb`.
- `bot/states.py:29-32` — `AdminStates`.

### Core
- `core/kp_db.py:953-994` — `delete_kp_by_id`.
- `core/kp_db.py:997-1087` — `clear_all_plates_data` (полная: все таблицы).
- `core/kp_db.py:1090-1130` — `get_db_stats`.
- `core/kp_db.py:1133-1203` — `clear_all_kp` (частичная: без completed_plates/plate_rests).
- `core/kp_db.py:1885-...` — `get_all_plate_rests`.
- `core/kp_db.py:2461-2519` — `recover_stuck_plates`.

### Backend
- `app/api/v1/router.py:7-13` — сборка API v1.
- `app/api/v1/endpoints/archive.py:130-140` — `DELETE /commercial/archive/{kp_id}`.
- `app/api/v1/endpoints/production.py:87-95` — `DELETE /production/plans/{plan_id}`.
- `app/dependencies/auth.py:13-34` — `get_current_user`, `require_roles`.
- `app/repositories/kp_repository.py:76-77` — `KpRepository.delete_offer`.
- `app/repositories/kp_archive_repository.py:33-34` — `KpArchiveRepository.delete`.
- `app/repositories/plan_repository.py:26-27` — `PlanRepository.delete_plan`.
- `app/repositories/auth_repository.py:28-150` — пользователи и роли.
- `app/services/archive_service.py:112-114` — `ArchiveService.delete_offer`.
- `app/services/production_service.py:12,35-37` — импорт `bot.handlers.plan_manager`, `delete_plan`.
- `app/core/settings.py:45,54-56` — пути `plita_db_path`, `plans_dir`, `plans_metadata_path`, `work_calendar_path`.
- `app/core/settings.py:90-95` — `ensure_directories`.
- `app/main.py:32-53` — `create_app`, монтирование `api_v1_router`/`web_router`.

### Frontend
- `frontend/src/app/router/AppRouter.tsx:7-17` — три маршрута (admin-маршрута нет).
- `frontend/src/app/layout/AppHeader.tsx:70-96` — навигация (admin-пункта нет).
- `frontend/src/features/commercial-archive/api/archiveApi.ts:29` — `archiveApi.delete`.
- `frontend/src/features/commercial-archive/hooks/useArchiveQueries.ts:60-69` — `useDeleteOfferMutation`.
- `frontend/src/features/commercial-archive/components/DeleteConfirmDialog.tsx:1-52` — паттерн модалки удаления.
- `frontend/src/features/production/api/productionApi.ts:27-28` — `productionApi.deletePlan`.
- `frontend/src/features/production/hooks/useProductionQueries.ts:87-95` — `useDeletePlanMutation`.
- `frontend/src/shared/api/httpClient.ts:31-48,93-106` — fetch-клиент c `credentials:"include"`.
- `frontend/src/shared/ui/Button.tsx:26-30` — `variant="danger"`.

## Архитектурные наблюдения

- **Бот хранит вся логика очистки централизованно в `bot/handlers/admin.py`** и переиспользует функции `core/kp_db.py` (single source of truth по SQLite).
- **Существуют ДВЕ отдельные функции «полной» очистки SQLite**: `clear_all_kp` (только КП-таблицы, без `completed_plates` и `plate_rests`) и `clear_all_plates_data` (все шесть таблиц + `sqlite_sequence`). Кнопка `⚙️ Управление БД → 🗑️ Очистить все данные` в боте использует именно `clear_all_plates_data`.
- **Очистка состояния «производство» состоит из двух хранилищ**: SQLite (`plita.db`) и JSON-файлы (`bot/data/plans/`, `plans_metadata.json`, `current_plan.json`). Полная зачистка в боте чистит оба — функция `clear_all_plans_data()` живёт прямо в `bot/handlers/admin.py:386-436`, а не в `core/`.
- **Таблица `app_users` (`plita.db`)** не задевается ни одной функцией очистки в `core/kp_db.py`, поэтому существующая модель полной очистки не выкидывает залогиненных пользователей.
- **Backend уже использует ролевое разграничение `admin / manager / production`** (`app/dependencies/auth.py:28-34`) и понимает cookie-сессии (`app/security/session.py`); `require_roles("admin")` — готовый замыкающий depend для admin-only endpoint'ов.
- **Backend импортирует `bot.handlers.plan_manager` напрямую** (`app/services/production_service.py:12`, `app/repositories/plan_repository.py:4`), поэтому функции работы с JSON-планами доступны из FastAPI без дублирования.
- **Frontend построен по FSD-подобной структуре** (`app/`, `pages/`, `features/`, `shared/`); каждая фича содержит подкаталоги `api/`, `hooks/`, `components/`, `types/`. Уже есть устоявшийся паттерн `useMutation` + модалка с `Alert(tone="warning")` + `Button(variant="danger")` (`DeleteConfirmDialog.tsx`).
- **HTTP-клиент `frontend/src/shared/api/httpClient.ts` единый для всего SPA**, использует `credentials: "include"` — cookie `app_session` уходит автоматически; ошибки выбрасываются как `ApiError` (статус, detail).
- **Endpoint'ы возвращают согласованные DTO**: `DeletePlanResponse` (`app/schemas/production.py`), `204 No Content` для удаления одиночного КП в архиве. Pydantic-схемы и сервисы лежат строго в `app/schemas/` и `app/services/`.
- **Старый HTML-кабинет `/web/managers`** держит дублирующий UI удаления одного КП на inline-JS (`app/web/router.py:763-774`), но не имеет UI массовой очистки. SPA-оболочка отдаётся через `/commercial-offer/new` (`app/web/router.py:982-997`).
- **`app/core/settings.py` НЕ описывает `current_plan.json`** — этот путь захардкожен в `bot/handlers/admin.py:409-412` и в коде `plan_manager`.
