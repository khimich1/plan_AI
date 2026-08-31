# Admin DB Reset Refresh

> Ideation session: 2026-07-23  
> Контекст: «Управление базой данных» → «Обнулить ВСЁ»; UI и тексты не совпадают с текущей архитектурой.

## Problem Statement

**Как сделать «Обнулить ВСЁ» честным и предсказуемым для локального dev-админа — чтобы UI описывал реальное поведение кода, без упоминания Telegram-бота и ложных обещаний про JSON-планы?**

## Recommended Direction

### Выбор: «Честный dev-reset» (направление A)

Три слоя изменений:

1. **UI truth layer** — убрать предупреждение про Telegram-бот; описать SQLite-планы как основной источник; переименовать статистику и все 4 кнопки reset; исправить тексты «Только планы (JSON)» (сейчас backend чистит и SQLite `production_plans`).

2. **Behavior layer** — `reset_full` дополнительно чистит legacy-артефакты:
   - SQLite: КП, плиты, остатки, журнал статусов, `production_plans` (как сейчас через `clear_all_plates_data` + `_clear_all_plans`);
   - `data/plans/`, `data/plans_metadata.json`, `data/current_plan.json`, `data/work_calendar.json` (best-effort);
   - **`bot_archived/data/plans/*.json`**, `bot_archived/data/plans_metadata.json`, `bot_archived/data/work_calendar.json` (21 legacy-файл в архиве — включить в полное обнуление);
   - **не трогать:** `app_users`, `managers` — админ остаётся залогиненным.

3. **DX layer** — при 403 от `DestructiveDbOperationBlocked` показывать конкретную инструкцию: «Добавьте `ALLOW_DESTRUCTIVE_DB_RESET=1` в `.env` и перезапустите backend»; задокументировать dev-preset в `.env.example`.

Кнопку «Только планы (JSON)» переименовать в «Удалить все планы» с честным описанием: SQLite `production_plans` + optional legacy cleanup в `data/`.

### Почему это направление

- **Для кого:** админ/разработчик на локальном стенде (сброс перед тестами/демо).
- **Успех:** полный refresh — тексты, статистика, кнопки, подсказка при 403.
- **Объём:** большой (frontend + backend + P6-aligned cleanup), но без смены security policy.
- Сейчас runtime хранит планы в SQLite (`PlanRepository`, `production_plans`); JSON в `data/plans/` не используется для read/write (`app/planning/plan_storage.py`). Telegram-бот soft-decommissioned с 2026-06-21 (`bot/README.md`).

### Env-guard (рекомендация)

**Не ослаблять guard автоматически** в development. Оставить `ALLOW_DESTRUCTIVE_DB_RESET=1`, но улучшить UX:

- явная ошибка в диалоге подтверждения (не общая «запрещено в текущем окружении»);
- блок в `.env.example` для локальной разработки;
- опционально: предупреждение в `run+logs.sh`, если флаг не выставлен.

Audit уже фиксировал риск misconfig при auto-allow в dev/staging.

## Key Assumptions to Validate

- [ ] **Legacy JSON в `bot_archived/data/plans/` — мусор**, не нужен для rollback → grep/спросить команду, нет ли скриптов, читающих оттуда в runtime
- [ ] **Единственный источник правды по планам — SQLite `production_plans`** → после wipe календарь и список планов в UI пустые
- [ ] **Локальный dev всегда с явным флагом** → один прогон full reset с `ALLOW_DESTRUCTIVE_DB_RESET=1`
- [ ] **Таблица `managers` должна сохраняться** → подтвердить: не входит в `clear_all_plates_data` (сейчас так)

## MVP Scope

**In:**

- Переписать `DbManagementModal`: warning, labels статистики, описания 4 кнопок
- Улучшить отображение 403 в `ResetConfirmDialog` / `getErrorMessage` (конкретная env-инструкция)
- `AdminService.reset_full`: cleanup `bot_archived/data/plans/` + metadata + calendar
- `get_stats`: честные счётчики — «Планов (SQLite): N», опционально «Legacy JSON: M»
- `.env.example` + короткая dev-инструкция

**Out (v1):**

- Auto-allow reset без флага в `APP_ENV=development`
- Удаление каталога `bot_archived/` из репозитория
- Изменение production break-glass policy (`DESTRUCTIVE_DB_RESET_BREAK_GLASS`)
- Wipe таблицы `managers` или всех `app_users`

## Not Doing (and Why)

- **Ослабление guard в dev** — audit ловил misconfig; лучше UX-подсказка, чем дыра в безопасности
- **Удаление `bot_archived/` из git** — отдельная задача P6-decommission, не смешивать с reset UX
- **Wipe `managers` / users** — полное обнуление бизнес-данных ≠ потеря доступа к админке
- **Отдельная кнопка «только legacy JSON»** — включено в full reset; лишняя сложность в UI

## Open Questions

1. Нужен ли **dry-run preview** перед подтверждением («будет удалено: 23 КП, 4 плана SQLite, 21 legacy JSON»)?
2. При full reset чистить ли **только** `bot_archived/data/plans/` или весь `bot_archived/data/`?
3. Достаточно ли **toast + refetch stats** после успеха, или показывать детальный отчёт `DbResetReport` в UI?

## Текущее расхождение UI ↔ код (baseline)

| UI / описание | Реальность (2026-07-23) |
|---------------|-------------------------|
| «Остановите Telegram-бот» | Бот deprecated, не используется |
| «JSON-планы производства» | Планы в SQLite `production_plans` (4 строки); `data/plans/` пуст |
| «Файлов планов (JSON)» в stats | Считает из SQLite `list_metadata()`, не файлы |
| «Только планы (JSON)» — SQLite не пострадает | `reset_plans_only` удаляет `production_plans` |
| Кнопка не «работает» | 403: нет `ALLOW_DESTRUCTIVE_DB_RESET=1` |

## Связанные файлы

- Frontend: `frontend/src/features/admin/components/DbManagementModal.tsx`, `ResetConfirmDialog.tsx`
- Backend: `app/services/admin_service.py`, `app/api/v1/endpoints/admin.py`
- Guard: `core/destructive_db_guard.py`
- SQLite wipe: `core/kp/offers_write.py` → `clear_all_plates_data`
- Plans: `app/repositories/plan_repository.py` → `delete_all_plans`
- P6 spec: `docs/specs/p6-legacy-decommission.md`
