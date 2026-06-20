# PLAN: Стабилизация P0-next (аудит 2026-06-20)

> **Дата:** 2026-06-20
> **Спека:** [`../../specs/stabilizaciya-p0-audit-2026-06-20.md`](../../specs/stabilizaciya-p0-audit-2026-06-20.md)
> **Источник:** [`../audits/2026-06-20-full-project-audit.md`](../audits/2026-06-20-full-project-audit.md)
> **Bot policy:** deprecated с P0-2026-06-19 — см. [`../../specs/stabilizaciya-p0-audit-2026-06-19.md`](../../specs/stabilizaciya-p0-audit-2026-06-19.md); WP2 = maintenance-only sunk cost

---

## Work Packages

| WP | Находка | Статус | Notes |
|----|---------|--------|-------|
| WP1 | A1/S1 — safety-net + SQLite authority | [x] done | `test_bot_plan_sqlite_authority.py`, `test_plan_storage_deprecation.py`; 23 passed (green, не RED — `plan_storage` уже на SQLite) |
| WP2 | A2 — bot thin adapter над core pipeline | [x] done | **Maintenance-only (bot deprecated):** adapter + parity test — sunk cost; не продлевать |
| WP3 | A3 — decommission globals | [~] partial | `PlateOrderContext` на planning hot paths; arch doc; PEP 562 proxy deferred |

---

## WP1 — A1/S1 (completed)

### Задачи

- [x] Рефактор `app/planning/plan_storage.py` — делегирование CRUD в `PlanRepository`
- [x] `plan_calendar` / `plan_aggregation` — через `plan_storage` (автоматически SQLite)
- [x] Bot handlers — `plan_manager` shim без прямого JSON I/O
- [x] Rollback в `production_export.py` — `repository.delete` вместо `os.remove`
- [x] Тесты: `test_bot_plan_sqlite_authority.py` (API↔bot cross-surface), `test_plan_storage_deprecation.py` (guard file I/O)
- [x] Тесты: `test_plan_sqlite_authority.py`, fixtures (`_repo_override`)
- [x] Migration: `scripts/migrate_plans_to_sqlite.py` (уже был из P0-2026-06-19)

### Verify

```powershell
pytest tests/test_bot_plan_sqlite_authority.py tests/test_plan_storage_deprecation.py tests/test_plan_sqlite_authority.py tests/test_plan_repository.py tests/test_migrate_plans_to_sqlite.py -q
```

**Safety-net (2026-06-20):** 23 passed, 0 failed. RED на текущем коде не воспроизведён — `plan_storage` уже делегирует в `PlanRepository`; guard-тесты фиксируют инвариант для WP2+.

---

## WP2 — A2 (completed)

- [x] Рефактор `bot/handlers/production_execution.py` → `ProductionPlanningService.run_planning_pipeline()` / `core/production/planning.py`
- [x] `bot/services/production_planning_adapter.py` — thin adapter (KP filters, rests, preview)
- [x] `ProductionPlanningService.run_planning_pipeline` + `build_plan_structure` (shared с `build_plan`)
- [x] Cross-surface test: `tests/test_bot_production_planning_parity.py`

### Verify

```powershell
pytest tests/test_bot_production_planning_parity.py tests/test_production_planning_service.py tests/test_core_production_planning.py -q
```

---

## WP3 — A3 (partial, 2026-06-20)

- [x] `PlateOrderContext` на hot paths: `core/production/planning.optimize`, `ProductionPlanningService.run_planning_pipeline`, API `POST /plans/build`, bot `production_execution` / `build_plan_preview`
- [x] Deployment constraint: [`architecture/plate-runtime-isolation.md`](./architecture/plate-runtime-isolation.md)
- [~] PEP 562 proxy в `config_and_data` — deferred (incremental strangler)

### Verify

```powershell
pytest tests/test_plate_runtime_isolation.py tests/test_plate_mutable_runtime_isolation.py tests/test_plate_runtime_request_isolation.py tests/test_bot_production_planning_parity.py -q
```

## Sprint closure (2026-06-20)

| WP | Итог |
|----|------|
| WP1 | [x] done — SQLite authority, cross-surface tests |
| WP2 | [x] done — bot adapter + parity |
| WP3 | [~] partial — context isolation on hot paths; PEP 562 deferred |

**Full suite:** `pytest tests/ -q` → **905 passed, 12 skipped, 0 failed**

**Spec status:** closed (partial A3 deferred)

*Обновлено: 2026-06-20 · Sprint closed (partial A3 deferred).*
