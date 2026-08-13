# Report: График поставки — Week 2 (API + import)

**Date:** 2026-08-07  
**Orchestration:** `orch-delivery-w2`  
**Status:** ✅ Week 2 complete  
**Spec:** [`docs/specs/delivery-schedule.md`](../../../docs/specs/delivery-schedule.md)  
**Plan:** [`ai_docs/develop/plans/delivery-schedule-plan.md`](../plans/delivery-schedule-plan.md)  
**Prev:** [`delivery-schedule-week1.md`](delivery-schedule-week1.md)

## Summary

Неделя 2 завершена: Pydantic-схемы + CRUD-сервис, GET с живым светофором, HTTP-роутер (GET/PUT), XLSX шаблон `/template` + `/import` (черновик без сохранения). Задачи **T4–T7 APPROVED/done**. Полный прогон week1+2: **90 passed**.

**Неделя 3 (T8–T12: документы XLSX/PDF + frontend + E2E / CP4) не начата.**

## What Was Built

### T4 — Pydantic-схемы + CRUD-сервис ✅ APPROVED

PUT принимает полный набор партий, валидирует инварианты (Σ ≤ qty, даты, qty ≥ 1), идемпотентен; GET отдаёт график или 404; статусы КП и `invoice_number`.

- **Files:** `app/schemas/delivery_schedule.py`, `app/services/delivery_schedule_service.py`, `tests/test_delivery_schedule_service.py`
- **Tests:** 16 (`test_delivery_schedule_service.py`)

### T5 — GET с живым светофором ✅ APPROVED

Связка сервиса с `check_batches`: остатки, occupancy, workdays; у партий `status` / `ready_date` / `hint`; R4 (`changed`) без падения.

- **Files:** `app/services/delivery_schedule_service.py`, `app/schemas/delivery_schedule.py`, `tests/test_delivery_schedule_service.py`
- **Tests:** покрыты тем же suite (16)

### T6 — Роутер + регистрация ✅ APPROVED

`/commercial/archive/{kp_id}/delivery-schedule` GET/PUT; роли admin/manager; 403 для customer; DI + OpenAPI.

- **Files:** `app/api/v1/endpoints/delivery_schedule.py`, `app/api/v1/router.py`, `app/dependencies/services.py`, `tests/test_delivery_schedule_endpoints.py`
- **Tests:** часть из 12 endpoint-тестов

### T7 — XLSX-шаблон build/parse + `/template`, `/import` ✅ APPROVED

`build_template` / `parse_template`; round-trip; `/import` возвращает черновик + unmatched, **не** пишет в БД.

- **Files:** `core/delivery_schedule_xlsx.py`, `app/api/v1/endpoints/delivery_schedule.py`, `app/services/delivery_schedule_service.py`, `tests/test_delivery_schedule_xlsx.py`, `tests/test_delivery_schedule_endpoints.py`
- **Tests:** 10 (xlsx) + endpoint-покрытие template/import

## Tests (week 1 + 2)

Команда:

```bash
.venv/bin/python -m pytest \
  tests/test_delivery_schedule_schema.py \
  tests/test_delivery_schedule_check.py \
  tests/test_delivery_schedule_service.py \
  tests/test_delivery_schedule_xlsx.py \
  tests/test_delivery_schedule_endpoints.py \
  tests/test_archive_service.py -q
```

| Suite | Count | Result |
|-------|------:|--------|
| `test_delivery_schedule_schema.py` | 13 | ✅ |
| `test_delivery_schedule_check.py` | 16 | ✅ |
| `test_delivery_schedule_service.py` | 16 | ✅ |
| `test_delivery_schedule_xlsx.py` | 10 | ✅ |
| `test_delivery_schedule_endpoints.py` | 12 | ✅ |
| `test_archive_service.py` (регрессия T2) | 23 | ✅ |
| **Total** | **90** | **passed** |

Warnings: `python_multipart` PendingDeprecation; Starlette cookies deprecation в TestClient (не блокеры).

## Gates

### CP2 — валидация светофора на прошлых заказах: **PARTIAL** (gate для человека)

Скрипт (вне репо): `/tmp/cp2_delivery_schedule_validate.py`  
БД: `plita.db` (`DEFAULT_DB`).

| kp_id | fact_days (creation→max completed) | len_m | est_days | est+buf | ±20% |
|------:|-----------------------------------:|------:|---------:|--------:|:----:|
| 1 | 4 | 1583.8 | 4 | 4 | YES |
| 2 | 8 | 338.4 | 1 | 1 | NO |

**Причина PARTIAL:** в локальной БД только **2** КП с `completed_plates` (нужно 3–5). Сэмпл недостаточен для калибровки констант; КП#2 сильно расходится (−88%). Неделю 2 **не блокирует** — нужен прогон на прод/стенд-данных человеком.

### CP3 — ручной uvicorn: **COVERED_BY_TESTS**

Ручной прогон через uvicorn не обязателен: TestClient покрыл:

| Операция | Покрытие |
|----------|----------|
| PUT create + GET return | `test_put_creates_and_get_returns` |
| GET 404 | `test_get_not_found` |
| PUT валидация 422 | `test_put_qty_exceeded_returns_422` |
| Auth / roles (GET/PUT/template/import) | `test_get_requires_auth`, `*_customer_forbidden` |
| GET `/template` | `test_get_template_as_admin` |
| POST `/import` (черновик, без записи schedule) | `test_post_import_valid_xlsx_returns_draft_without_db_schedule` |
| OpenAPI path | `test_openapi_contains_delivery_schedule_path` |
| XLSX round-trip | `tests/test_delivery_schedule_xlsx.py` |

Согласование макета **документа** (XLSX/PDF) — на неделе 3 (T8 / R5), не в scope CP3 API.

## Technical notes

- `/import` намеренно не сохраняет график — только draft + unmatched.
- Светофор на GET собирается через T3 (`check_batches`) + readiness/occupancy/work calendar.
- CP2 на 2 точках: крупный заказ совпал; мелкий — нет (возможны выходные, конкуренция планов, неполный объём в `completed_plates`).

## Next Steps — Week 3 (не начата)

| Task | Scope |
|------|--------|
| T8 | Документ XLSX + PDF + `/document` (макет с пользователем) |
| T9 | Frontend: api / types / hooks |
| T10 | Frontend: редактор партий + светофор + кнопка в drawer |
| T11 | Frontend: импорт + документы |
| T12 | E2E менеджера + полный зелёный CI / **CP4** |

Перед выкаткой: закрыть **CP2** на 3–5 реальных заказах (±20%) — калибровка `TRACKS_PER_DAY_DEFAULT` / buffer при необходимости.

## Related Documentation

- Spec: [`docs/specs/delivery-schedule.md`](../../../docs/specs/delivery-schedule.md)
- Plan: [`ai_docs/develop/plans/delivery-schedule-plan.md`](../plans/delivery-schedule-plan.md)
- Week 1: [`delivery-schedule-week1.md`](delivery-schedule-week1.md)
- Workspace: `.cursor/workspace/active/orch-delivery-w2/`
