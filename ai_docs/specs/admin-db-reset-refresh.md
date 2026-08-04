# Spec: Admin DB Reset Refresh

> **Источник:** [`ai_docs/ideas/admin-db-reset-refresh.md`](../ideas/admin-db-reset-refresh.md)  
> **Дата:** 2026-07-23  
> **Статус:** implemented — 2026-07-23  
> **Plan:** [`ai_docs/develop/plans/2026-07-23-admin-db-reset-refresh.md`](../develop/plans/2026-07-23-admin-db-reset-refresh.md)  
> **Связанные:** `DbManagementModal.tsx`, `AdminService`, `destructive_db_guard.py`, P6 legacy decommission

---

## Assumptions (validate before Plan)

```
ASSUMPTIONS I'M MAKING:
1. Основной пользователь — dev-админ на локальном стенде (сброс перед тестами/демо).
2. Telegram-бот не используется; предупреждение про бота удаляем из UI полностью.
3. Планы в runtime — только SQLite `production_plans`; JSON в `data/plans/` — legacy cleanup.
4. Full reset дополнительно чистит `bot_archived/data/plans/`, `bot_archived/data/plans_metadata.json`,
   `bot_archived/data/work_calendar.json` — НЕ весь каталог `bot_archived/`.
5. Security policy не меняем: reset по-прежнему требует ALLOW_DESTRUCTIVE_DB_RESET=1 в development.
6. HTTP 403 не раскрывает env-переменные (существующий тест test_raise_destructive_db_blocked_hides_env_details);
   dev-инструкция показывается только на frontend по известному 403-сообщению.
7. Dry-run preview перед подтверждением — OUT of MVP (follow-up).
8. После успешного reset достаточно refetch stats + success Alert в модалке (без toast-библиотеки).
9. Таблицы `app_users` и `managers` сохраняются при full reset (текущее поведение).
→ Correct me now or I'll proceed with these.
```

---

## Objective

Привести админское «Обнулить ВСЁ» в соответствие с текущей архитектурой: UI честно описывает SQLite-планы и реальный scope удаления; backend при full reset убирает legacy JSON в архиве; dev-админ понимает, почему reset заблокирован (403), и как это включить локально.

### User stories

| # | Как dev-админ… | Я хочу… | Чтобы… |
|---|----------------|---------|--------|
| US-1 | открываю «Управление базой данных» | видеть актуальную статистику (SQLite-планы, legacy JSON отдельно) | не путать JSON-файлы с реальным хранилищем |
| US-2 | нажимаю «Обнулить ВСЁ» | понимать, что именно удалится (КП, плиты, SQLite-планы, календарь, legacy) | не бояться скрытых сюрпризов |
| US-3 | reset заблокирован guard'ом | получить понятную инструкцию про `.env` | не гадать, почему кнопка «не работает» |
| US-4 | выполняю full reset на стенде | очистить и `bot_archived/data/plans/*.json` | не оставлять мусор после «полного» wipe |
| US-5 | после reset | остаться залогиненным | не терять доступ к админке |

### Reframed success criteria

| Требование | Конкретный критерий |
|------------|---------------------|
| «Актуальный UI» | Нет упоминаний Telegram-бота; нет формулировок «JSON-планы» как основного storage |
| «Честная статистика» | Отдельно: планы SQLite + legacy JSON files (если есть) |
| «Full reset полный» | После reset: 0 КП, 0 `production_plans`, пустой календарь, 0 legacy JSON в `data/plans/` и `bot_archived/data/plans/` |
| «403 понятен» | При blocked reset UI показывает шаги: `ALLOW_DESTRUCTIVE_DB_RESET=1` → restart backend |
| «Без регрессий security» | Production guard и тест `hides_env_details` остаются зелёными |

---

## Tech Stack

| Слой | Стек |
|------|------|
| Frontend | React 19, TypeScript, Vite, Vitest, TanStack Query |
| Backend | FastAPI, Pydantic v2, SQLite (`plita.db`) |
| Admin API | `/api/v1/admin/db/*` |
| Guard | `core/destructive_db_guard.py` |
| Tests | pytest (`tests/test_admin_service.py`, `tests/test_destructive_db_guard.py`, `tests/test_http_errors.py`) |

---

## Commands

```bash
# Dev
/home/roman/project/Шишов/run+logs.sh

# Backend tests (admin reset)
pytest tests/test_admin_service.py tests/test_destructive_db_guard.py tests/test_http_errors.py -q

# Frontend
cd frontend && npm test -- --run
cd frontend && npm run build
cd frontend && npm run typecheck

# Manual reset (local dev)
# 1. В .env: ALLOW_DESTRUCTIVE_DB_RESET=1
# 2. Перезапуск backend
# 3. UI → Управление БД → Обнулить ВСЁ → ввести ОБНУЛИТЬ
```

---

## Project Structure

```
app/
  services/admin_service.py       → reset_full, get_stats, _clear_all_plans, _clear_archived_legacy
  schemas/admin.py                → DbStatsResponse (новые поля / описания)
  api/v1/endpoints/admin.py       → без изменений контракта URL
  core/http_errors.py             → MSG_DESTRUCTIVE_DB_BLOCKED (без env в detail)

core/
  destructive_db_guard.py         → без изменений policy
  kp/offers_write.py              → clear_all_plates_data (без изменений scope)

frontend/src/features/admin/
  components/DbManagementModal.tsx    → тексты, stats labels, success report
  components/ResetConfirmDialog.tsx   → 403 hint (optional prop или helper)
  types/admin.ts                      → обновить DbStatsResponse
  lib/destructiveResetError.ts        → NEW: map 403 → dev instruction

tests/
  test_admin_service.py           → + archived legacy cleanup, + get_stats fields

.env.example                      → dev block для ALLOW_DESTRUCTIVE_DB_RESET
run+logs.sh                       → optional warning если флаг не выставлен
```

---

## Code Style

```python
def _count_legacy_json_files(self, *directories: Path) -> int:
    total = 0
    for directory in directories:
        if directory.is_dir():
            total += len(list(directory.glob("*.json")))
    return total

def _clear_archived_legacy(self) -> dict[str, int]:
    """Best-effort cleanup of bot_archived/data plan artifacts (full reset only)."""
    report = {"archived_plan_files": 0, "archived_metadata": 0, "archived_calendar": 0}
    # unlink / rmtree guarded paths under PROJECT_ROOT / "bot_archived" / "data"
    return report
```

```typescript
export const DESTRUCTIVE_DB_BLOCKED_HINT =
  "Операция запрещена в текущем окружении. Для локальной разработки добавьте " +
  "ALLOW_DESTRUCTIVE_DB_RESET=1 в .env и перезапустите backend.";

export function getDestructiveResetErrorMessage(error: unknown): string | null {
  if (!(error instanceof ApiError) || error.status !== 403) return null;
  if (error.detail !== MSG_DESTRUCTIVE_DB_BLOCKED_CLIENT) return null;
  return DESTRUCTIVE_DB_BLOCKED_HINT;
}
```

- Backend: пути только через `Settings` или явный `PROJECT_ROOT / "bot_archived"` — без hardcode `bot/data`
- Frontend: env-инструкция только в UI, не в API response
- Тесты: изолированный `tmp_path` для archived cleanup (не трогать реальный `bot_archived/` в CI)

---

## Testing Strategy

| Уровень | Что покрываем | Где |
|---------|---------------|-----|
| Unit (backend) | `reset_full` чистит archived JSON; `get_stats` возвращает legacy count | `tests/test_admin_service.py` |
| Unit (backend) | Guard policy без изменений | `tests/test_destructive_db_guard.py` |
| Unit (backend) | 403 detail не содержит env vars | `tests/test_http_errors.py` |
| Unit (frontend) | `getDestructiveResetErrorMessage` для 403 | `destructiveResetError.test.ts` (NEW) |
| Manual | Full reset с флагом → stats нули, production UI пуст | checklist ниже |

Coverage expectation: новые ветки в `AdminService` и frontend error helper — 100% happy path + blocked path.

---

## Boundaries

### Always

- Запускать `pytest tests/test_admin_service.py` и frontend tests перед merge
- Сохранять `app_users` / `managers` при full reset
- Fail-closed guard без auto-allow в development
- Не раскрывать env-переменные в HTTP `detail` (security test)

### Ask first

- Изменение `destructive_db_reset_allowed()` policy
- Удаление `bot_archived/` из репозитория
- Добавление dry-run endpoint
- Breaking rename полей admin API, если появятся внешние потребители

### Never

- Wipe `app_users` / `managers` в full reset
- Auto-enable `ALLOW_DESTRUCTIVE_DB_RESET` в production
- Читать/писать планы через JSON в runtime (откат P6)
- Удалять исходники `bot_archived/handlers/` — только data artifacts

---

## Functional Spec

### F1. UI texts (`DbManagementModal`)

| Элемент | As-Is | To-Be |
|---------|-------|-------|
| Top warning | «остановите Telegram-бот…» | «Действия необратимы. Убедитесь, что никто другой не работает с базой параллельно.» |
| Full reset description | «JSON-планы производства» | «все планы производства (SQLite)», «legacy JSON-файлы (если есть)», календарь |
| Stats: plans | «Файлов планов (JSON)» | «Планов (SQLite)» |
| Stats: legacy | — | «Legacy JSON-файлов» (новое поле) |
| Stats: current_plan | «current_plan.json» | убрать или «Legacy current_plan.json (не используется)» |
| Button: plans-only | «Только планы (JSON)» | «Удалить все планы» |
| Plans-only description | «SQLite не пострадает» | «Удаляются все записи production_plans и legacy JSON в data/plans/» |

### F2. Stats API (`DbStatsResponse`)

Расширить схему (additive, admin-only):

| Поле | Тип | Источник |
|------|-----|----------|
| `plans_count` | int | SQLite `list_metadata()` — **исправить description** в Pydantic |
| `legacy_json_files_count` | int | `glob("*.json")` в `settings.plans_dir` + `bot_archived/data/plans` |
| `current_plan_present` | bool | legacy file exists — оставить, но не акцентировать в UI |

`get_stats()` считает legacy files, не путая с SQLite.

### F3. Full reset behavior (`reset_full`)

Порядок (текущий + новое):

1. `require_destructive_db_reset()`
2. `clear_all_plates_data(db)` — КП/плиты/остатки/журнал
3. `_clear_all_plans()` — SQLite plans + `data/plans/*` + metadata + current_plan
4. **`_clear_archived_legacy()`** — NEW:
   - `bot_archived/data/plans/*.json` → rmtree/recreate or unlink each
   - `bot_archived/data/plans_metadata.json` → unlink if exists
   - `bot_archived/data/work_calendar.json` → unlink if exists (full reset; calendar-only still resets active `data/work_calendar.json`)
5. `_reset_calendar()` — пустой `data/work_calendar.json`

`DbResetReport.plans` расширить ключами: `archived_plan_files`, `archived_metadata`, `archived_calendar` (optional int fields).

### F4. 403 UX

- Backend: `MSG_DESTRUCTIVE_DB_BLOCKED` без изменений текста (security)
- Frontend: если `ApiError.status === 403` и `detail === MSG`, показать `DESTRUCTIVE_DB_BLOCKED_HINT` под error Alert в `ResetConfirmDialog`
- `.env.example`: раскомментированный пример для local dev
- `run+logs.sh` (optional): если `ALLOW_DESTRUCTIVE_DB_RESET` не truthy — жёлтое предупреждение при старте

### F5. Success feedback

После успешного reset:

- Закрыть confirm dialog (как сейчас)
- `statsQuery.refetch()` (как сейчас)
- **NEW:** краткий success Alert в главной модалке: «Удалено: N КП, M планов, K legacy JSON» из `DbResetReport`

---

## As-Is → To-Be (baseline 2026-07-23)

| As-Is | To-Be |
|-------|-------|
| UI про Telegram-бот | UI про параллельный доступ к БД |
| «JSON-планы» | «SQLite-планы + legacy cleanup» |
| Stats label «JSON» при SQLite count | «Планов (SQLite)» + «Legacy JSON» |
| Full reset не чистит `bot_archived/data/plans/` | Чистит archived legacy |
| 403: «запрещено в текущем окружении» | + frontend hint с `.env` шагами |
| `plans-only` UI врёт про SQLite | Честное описание |

---

## Success Criteria (Definition of Done)

- [ ] В `DbManagementModal` нет упоминаний Telegram-бота и «JSON-планы» как primary storage
- [ ] `GET /admin/db/stats` возвращает `legacy_json_files_count`; descriptions в schema актуальны
- [ ] `POST /admin/db/reset/full` удаляет файлы в `bot_archived/data/plans/` (тест на tmp_path mirror)
- [ ] После full reset: `KP_offers=0`, `production_plans=0`, `app_users>=1`
- [ ] 403 в UI показывает dev-инструкцию; HTTP detail по-прежнему без `ALLOW_DESTRUCTIVE`
- [ ] `pytest tests/test_admin_service.py tests/test_destructive_db_guard.py tests/test_http_errors.py` — green
- [ ] `cd frontend && npm test -- --run && npm run build` — green
- [ ] `.env.example` обновлён с dev preset

### Manual checklist

1. Запустить без `ALLOW_DESTRUCTIVE_DB_RESET` → full reset → 403 + hint в UI
2. Добавить флаг, restart → full reset с `ОБНУЛИТЬ` → stats нули, календарь пуст, планы исчезли из Production UI
3. Проверить, что login admin работает после reset

---

## Open Questions

1. **Убрать `current_plan_present` из UI полностью** или оставить как «legacy indicator»?
2. **`run+logs.sh` warning** — включать в MVP или отложить?
3. **Success Alert** — показывать всегда или только после full reset (не после partial)?

---

## Out of Scope (v1)

- Dry-run preview («будет удалено: …») перед confirm
- Auto-allow reset без флага в `APP_ENV=development`
- Удаление `bot_archived/` из git / hard P6 decommission
- Изменение production break-glass (`DESTRUCTIVE_DB_RESET_BREAK_GLASS`)
- Wipe `managers` table
- Отдельная кнопка «только legacy JSON»

---

## Next Step

Phase 4 **IMPLEMENT** → [`2026-07-23-admin-db-reset-refresh.md`](../develop/plans/2026-07-23-admin-db-reset-refresh.md) tasks `RESET-001`…`RESET-010`, one at a time.
