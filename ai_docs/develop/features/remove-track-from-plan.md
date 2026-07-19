# Удаление дорожки из плана

**Status:** ✅ Implemented  
**Date:** 2026-05-19  
**Orchestration:** `orch-remove-track-1b7316b8`  
**Plan:** [удаление_дорожки_из_плана_1b7316b8.plan.md](file:///c:/Users/Роман/.cursor/plans/удаление_дорожки_из_плана_1b7316b8.plan.md)

## Краткое описание

Пользователь удаляет **одну дорожку** конкретного производственного плана на выбранную дату через веб-интерфейс (DayDrawer). Плиты не исчезают: они возвращаются в `kp_plates` со статусом `в производстве` и снова доступны для `build_plan` / оптимизатора.

Операция атомарна по БД: сначала возврат строк в SQLite, затем правка JSON-плана. При ошибке БД JSON не меняется.

**Роли:** `admin`, `production` (как у остальных production routes).

## API

### `DELETE /api/v1/production/plans/{plan_id}/days/{date}/tracks/{track_index}`

| Параметр | Описание |
|----------|----------|
| `plan_id` | ID плана (stem JSON-файла) |
| `date` | Ключ дня в `plan['days']`, ISO `YYYY-MM-DD` |
| `track_index` | **0-based** индекс в `plan['days'][date]['tracks']` внутри **этого** плана |

**Request body:** нет.

**Response 200** (`RemoveTrackResponse`):

```json
{
  "plan_id": "plan_20260519_162417",
  "date": "2026-04-21",
  "track_index": 0,
  "plates_returned": 3,
  "saved_tracks_count": 1,
  "warnings": null
}
```

| Поле | Смысл |
|------|-------|
| `plates_returned` | Сколько физических единиц вернули в «в производстве» |
| `saved_tracks_count` | Оставшихся дорожек в дне после удаления; `0`, если день удалён из `plan['days']` |
| `warnings` | Зарезервировано; в v1 обычно `null` |

**Коды ошибок** (тело: `{ "detail": "<сообщение>" }`):

| HTTP | `code` (внутренний) | Когда |
|------|---------------------|-------|
| 404 | `plan_not_found` | План не найден |
| 404 | `day_not_found` | День отсутствует в плане |
| 409 | `day_already_completed` | `day.completed == true` |
| 409 | `incomplete_return` | В БД вернули меньше плит, чем ожидалось по дорожке |
| 400 | `invalid_track_index` | Индекс вне `[0, len(tracks))` |
| 400 | `no_plate_identity` | Нет `kp_plate_id` и legacy-идентичностей в items |
| 500 | `db_return_failed` | Ошибка SQLite при возврате |
| 500 | `plan_save_failed` | БД обновлена, но `save_plan` не удался (редкий кейс) |

**Код:** `app/api/v1/endpoints/production.py` → `ProductionService.remove_track` → `plan_manager.remove_track_from_plan`.

## Поток данных (DB → JSON)

Симметрично коммиту плана (`commit_plan_plates` → `save_plan`):

```
DayDrawer → productionApi.deleteTrack
         → DELETE endpoint
         → plan_manager.remove_track_from_plan
              1. load_plan, проверки (completed, index, identity)
              2. collect_plate_returns_from_track(track)
              3. kp_db.return_plate_rows_for_plan (транзакция)
                 — SELECT ... WHERE id=? AND plan_id=? AND status='в плане'
                 — audit reason='track_removed'
                 — при incomplete_return → rollback, JSON не трогаем
              4. tracks.pop(track_index); saved_tracks_count = len(tracks)
                 — если tracks пуст → del plan['days'][date]
              5. save_plan + update_plan_metadata
```

**Сбор плит:** `core/plan_track_removal.py` — `_iter_physical_items` (root + `secondary_cuts`), как при коммите (P9).

**Возврат в БД:** `core/kp_db.py` → `return_plate_rows_for_plan`:
- основной путь: `Counter[kp_plate_id → qty]`;
- legacy (без `kp_plate_id`): `Counter[(kp_id, plate_name) → qty]` с фильтром `AND plan_id=?`.

## Инварианты и защиты

| Инвариант | Поведение |
|-----------|-----------|
| Плиты не пропадают | `qty` сохраняется; статус `в плане` → `в производстве`, `plan_id = NULL` |
| Повторное планирование | Плиты снова в выборке `status='в производстве'` |
| Календарь занятости | −1 дорожка (`count_day_tracks` / `get_global_day_occupancy`) |
| Завершённый день | **409** — иначе рассинхрон с `completed_plates` |
| Чужой `plan_id` в строке | Строка не трогается (проверка в `return_plate_rows_for_plan`) |
| `incomplete_return` | Rollback транзакции; JSON без изменений; **409** |
| Пустой день | Ключ дня удаляется из `plan['days']` — календарь не показывает пустой день |

**Не трогаем:** `optimization_result`, `orders_2d`, `plate_lookup_*` — архив последней оптимизации; на доступность плит не влияет.

## Frontend

**Идентификация дорожки:** в day view `track_number` — сквозной номер по всем планам дня; для DELETE нужен **`plan_track_index`** — индекс внутри плана.

| Слой | Файл | Назначение |
|------|------|------------|
| Тип | `frontend/src/features/production/types/production.ts` | `DayTrackDetail.plan_track_index`, `RemoveTrackResponse` |
| API | `frontend/src/features/production/api/productionApi.ts` | `deleteTrack(planId, date, trackIndex)` |
| Hook | `frontend/src/features/production/hooks/useProductionQueries.ts` | `useDeleteTrackMutation` → invalidate `productionKeys.all` |
| UI | `frontend/src/features/production/components/DayDrawer.tsx` | Кнопка «Удалить дорожку», confirm, toast |

**DayDrawer:**
- кнопка только если `!plan.completed`;
- `window.confirm` с числом плит;
- вызов с `track.plan_track_index`, **не** `track.track_number`;
- после успеха — refetch day view и календаря;
- при **409** — «День уже завершён, удаление невозможно».

**Backend day view:** `get_tracks_for_date_from_all_plans` проставляет `source_track_index`; `day_view_service` отдаёт его как `plan_track_index` в `DayTrackDetail`.

## Тесты

Файл: [`tests/test_remove_track_from_plan.py`](../../../tests/test_remove_track_from_plan.py)

| Тест | Что проверяет |
|------|---------------|
| `test_remove_track_happy_path` | JSON без дорожки, БД `в производстве` |
| `test_replan_after_track_removal` | `build_plan` снова забирает те же плиты |
| `test_remove_track_completed_day_raises` | 409, без изменений |
| `test_wrong_plan_id_row_not_touched` | Изоляция по `plan_id` |
| `test_secondary_cuts_both_units_returned` | Root + secondary_cuts |
| `test_two_plans_same_date_isolation` | Удаление в плане A не трогает B |
| `test_saved_tracks_count_sync_after_removal` | `saved_tracks_count == len(tracks)` |
| `test_remove_track_no_plate_identity` | 400 при пустых identity |
| Unit: `collect_plate_returns_from_track_*` | Сбор счётчиков без БД |
| Unit: `return_plate_rows_for_plan_*` | Happy path, wrong plan_id, partial qty |

## Вне scope v1

- Telegram-бот
- Правка `optimization_result` / пересчёт ILP
- Удаление отдельной плиты внутри дорожки (только целая дорожка)
- Undo / корзина

## Связанные модули

| Модуль | Роль |
|--------|------|
| `core/plan_track_removal.py` | Сбор возвратов, `TrackRemovalError` |
| `core/kp_db.py` | `return_plate_rows_for_plan` |
| `app/planning/plan_manager.py` | `remove_track_from_plan`, `source_track_index` |
| `app/services/production_service.py` | HTTP-маппинг ошибок |
| `app/schemas/production.py` | `RemoveTrackResponse`, `DayTrackDetail.plan_track_index` |
| `app/services/day_view_service.py` | Прокидывание `plan_track_index` в API |
