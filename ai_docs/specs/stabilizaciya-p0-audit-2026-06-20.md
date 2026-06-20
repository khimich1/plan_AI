# Spec: Стабилизация P0-next — аудит 2026-06-20

> **Тип:** remediation feature-spec
> **Дата:** 2026-06-20
> **Статус:** closed (partial A3 deferred)
> **Источник:** [`../develop/audits/2026-06-20-full-project-audit.md`](../develop/audits/2026-06-20-full-project-audit.md)
> **Предыдущий спринт:** [`stabilizaciya-p0-audit-2026-06-19.md`](./stabilizaciya-p0-audit-2026-06-19.md) (закрыт)

---

## Objective

Закрыть **3 critical** находки аудита 2026-06-20: split-brain планов (A1/S1), обход core pipeline ботом (A2), legacy globals (A3).

## Scope (P0-next)

| ID | Проблема | Fix |
|----|----------|-----|
| **A1/S1** | SQLite (`PlanRepository`) и JSON (`plan_storage` / `bot/data/plans/`) — параллельные источники | Единый SQLite authority; bot/web через repository |
| **A2** | Bot ~900 LOC дублирует planning pipeline | Thin adapter над `core/production/planning.py` |
| **A3** | `plate_runtime_state` / `config_and_data` globals | Context objects + DI (deferred partial) |

## WP1 — A1/S1: единый SQLite authority

**Acceptance:**
- [x] `plan_storage` CRUD делегирует в `PlanRepository` (нет file I/O для планов)
- [x] `plan_calendar` / occupancy читают через тот же path (via `plan_storage`)
- [x] Bot handlers используют `plan_manager` shim → SQLite
- [x] JSON backdoor в distribution/repository убран (A11 — plan_storage no-op)
- [x] Migration script существует: `scripts/migrate_plans_to_sqlite.py`
- [x] Integration test: API write → `plan_storage` read; bot `save_plan` → repository

**Verify:** `pytest tests/test_plan_sqlite_authority.py tests/test_plan_repository.py tests/test_migrate_plans_to_sqlite.py -q`

## WP2 — A2: bot → core pipeline

**Acceptance:**
- [x] `production_execution.py` вызывает `core/production/planning.py`
- [x] Cross-surface test: одинаковые fixtures → одинаковый план (web vs bot adapter)

## WP3 — A3: runtime globals (partial)

**Acceptance:**
- [x] `PlateOrderContext` на hot paths bot+API (planning, optimization entry points)
- [x] Документирован single-instance assumption или TTL cache — см. [`../develop/architecture/plate-runtime-isolation.md`](../develop/architecture/plate-runtime-isolation.md)

**Partial (deferred):** полный decommission PEP 562 proxy в `config_and_data` — вне scope WP3.

## Changelog / итоги спринта (2026-06-20)

| Дата | Изменение |
|------|-----------|
| 2026-06-20 | **WP1 [x]** — SQLite authority: `plan_storage` → `PlanRepository`, bot shim, guard-тесты без file I/O |
| 2026-06-20 | **WP2 [x]** — bot thin adapter: `production_execution` → `run_planning_pipeline`; parity test web vs bot |
| 2026-06-20 | **WP3 [~]** — `PlateOrderContext` на planning hot paths; [`plate-runtime-isolation.md`](../develop/architecture/plate-runtime-isolation.md); PEP 562 proxy deferred |
| 2026-06-20 | **Closure:** `pytest tests/ -q` → **905 passed, 12 skipped, 0 failed** |

### Deferred (post P0-next)

- **A3 phase 2** — полный decommission PEP 562 proxy в `core/config_and_data.py`
- Прямые `PlateOrderContext.fresh_empty()` fallback в core — incremental migration
- Межпроцессный TTL cache для plate state — не в scope WP3

### Следующий шаг

1. Backlog **A4** (bot → core, не app), **A5** (god-modules), **P1-next** (S4 logout, S6 admin guard, Q1 bare except)
2. Отдельный спринт **A3 full decommission** после стабилизации hot paths

*Создано: 2026-06-20 · Закрыто: 2026-06-20 (partial A3 deferred).*
