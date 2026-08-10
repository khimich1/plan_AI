# Report: График поставки — Week 1 (ядро)

**Date:** 2026-08-07  
**Orchestration:** `orch-delivery-w1`  
**Status:** ✅ Week 1 complete, CP1 ready  
**Spec:** [`docs/specs/delivery-schedule.md`](../../../docs/specs/delivery-schedule.md)  
**Plan:** [`ai_docs/develop/plans/delivery-schedule-plan.md`](../plans/delivery-schedule-plan.md)

## Summary

Неделя 1 оркестрации завершена: схема БД (3 таблицы), общий модуль констант ёмкости и движок светофора. Все задачи T1–T3 **APPROVED**. Gate **CP1** готов к закрытию.

**Реализация недели 2 (T4–T7, API + XLSX import) не начата.**

## What Was Built

### T1 — Схема БД: 3 таблицы графика ✅

Таблицы `delivery_schedule` / `delivery_batch` / `delivery_batch_item` в `ensure_schema`: идемпотентность, `UNIQUE kp_id`, FK cascade от `KP_offers` / batch / plate, CHECK qty ≥ 1, индексы.

- **Files:** `core/kp_db_schema.py`, `tests/test_delivery_schedule_schema.py`
- **Tests:** 13 collected (`pytest tests/test_delivery_schedule_schema.py`)

### T2 — Константы ёмкости → `core/production_capacity.py` ✅

Вынесены `MAX_TRACK_LENGTH_M` (101 м) и `TRACKS_PER_DAY_DEFAULT` (5 дор./день). `archive_service` импортирует из нового модуля; поведение estimate не менялось.

- **Files:** `core/production_capacity.py` (новый), `app/services/archive_service.py`
- **Tests:** 23 зелёных в `tests/test_archive_service.py` (без правок тестов)

### T3 — Движок светофора `core/delivery_schedule_check.py` ✅

Чистая функция `check_batches(...) -> list[BatchCheck]`: зелёный/жёлтый/красный, вычет produced, пропуск выходных, конкуренция партий за ёмкость по `produce_by`, hint «+N дорожек» для красных, R2 (дата вне `days_info` = свободный день с дефолтным max).

- **Files:** `core/delivery_schedule_check.py`, `tests/test_delivery_schedule_check.py`
- **Tests:** 16 collected (`pytest tests/test_delivery_schedule_check.py`)

## CP1 Gate

| Критерий | Статус |
|----------|--------|
| `pytest tests/test_delivery_schedule_*.py` зелёный | ✅ (schema 13 + check 16) |
| Схема идемпотентна (двойной `ensure_schema`) | ✅ |
| Cascade от KP / batch / plate покрыт тестами | ✅ |
| Движок корректен на синтетике | ✅ |

## Tests

| Suite | Count |
|-------|------:|
| `test_delivery_schedule_schema.py` | 13 |
| `test_archive_service.py` (регрессия T2) | 23 |
| `test_delivery_schedule_check.py` | 16 |
| **CP1 delivery_schedule_*** | **29** |

## Review notes (minor)

- **Naming:** в планировании/occupancy исторически `MAX_TRACKS_PER_DAY`, в `production_capacity` — `TRACKS_PER_DAY_DEFAULT`. Семантика совпадает (5), имена разные — не блокер для CP1; унификация возможна позже.
- **Дублирование константы:** `MAX_TRACK_LENGTH_M = 101.0` по-прежнему продублирован в других модулях (`core/rescue_tracks.py`, `core/track_reconciliation.py`, …) помимо нового `production_capacity`. T2 сознательно трогал только archive-путь; полная консолидация — вне scope недели 1.
- **Edge-case тесты:** кейсы `produce_by < today` опциональны — можно добавить на неделе 2 при желании; на CP1 не влияют.

## Next Steps — Week 2

Задачи **T4–T7** (API + XLSX import), gates **CP2** / **CP3**:

| Task | Scope |
|------|--------|
| T4 | Pydantic-схемы + CRUD-сервис (PUT/GET без живого светофора) |
| T5 | GET с живым светофором (связка сервиса и T3) |
| T6 | Роутер + регистрация `/commercial/archive/{kp_id}/delivery-schedule` |
| T7 | ‖ XLSX-шаблон: build/parse + `/template`, `/import` |

- **CP2:** валидация светофора на 3–5 прошлых заказах (±20%), калибровка констант.
- **CP3:** ручной прогон PUT → GET → import через uvicorn; согласование макета документа.

Неделя 3 (T8–T10: документы XLSX/PDF + frontend + E2E / CP4) — после недели 2.

## Related Documentation

- Spec: [`docs/specs/delivery-schedule.md`](../../../docs/specs/delivery-schedule.md)
- Plan: [`ai_docs/develop/plans/delivery-schedule-plan.md`](../plans/delivery-schedule-plan.md)
- Workspace: `.cursor/workspace/active/orch-delivery-w1/`
