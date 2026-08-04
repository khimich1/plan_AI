# PEP 562 decommission — `core/config_and_data.py` (A3 phase 3+)

**Audit:** A3 (2026-06-20) · **Spec:** P2 WP1  
**Status:** in progress — web/core off proxy (P4 WP3 A4); **bot archived (P5 WP1)** — proxy backlog tests-only  
**Related:** [`plate-runtime-isolation.md`](./plate-runtime-isolation.md)

## Problem

`core/config_and_data.py` exposes mutable plate-order state via PEP 562 `__getattr__`:

```python
def __getattr__(name: str) -> Any:
    if name in MUTABLE_LEGACY_NAMES:
        return getattr(get_plate_mutable_runtime(), MUTABLE_ATTR_MAP[name])
    raise AttributeError(...)
```

Callers using `import core.config_and_data as cfg` and `cfg.PLATES_*` get **implicit global runtime** instead of explicit `PlateOrderContext` / `get_plate_mutable_runtime()`. This hides request boundaries and complicates testing.

**Out of scope for this ADR:** deleting the `config_and_data.py` module file entirely (only proxy + caller migration).

## Target APIs (replace proxy reads)

| Legacy (`cfg.X`) | Explicit replacement |
|------------------|-------------------|
| `PLATES_*`, cuts, strips, totals | `get_plate_mutable_runtime().<attr>` or `PlateOrderContext.plates` |
| `PLATE_LOAD_DETAILS`, `PLATE_EXACT_WIDTHS`, … | `PlateMutableRuntime` fields via `get_plate_mutable_runtime()` |
| `PlateOrder` / current order | `core.domain.plate_order.get_current_plate_order()` or `PlateOrderContext` |
| Constants (`TRACK_LENGTH_M`, prices) | `core.config.constants` |
| App settings | `core.config.app_settings.get_config()` |
| Parse / plate lists | `core.parsing.plate_lists.set_plate_lists_from_text`, `core.domain.plate_order` |
| Pure helpers (`make_plate_name`, `canonical_plate_key`, …) | Keep importing **named functions** from `config_and_data` (no proxy) |

**Rule:** no `cfg.PLATES_* = ...` assignment; only in-place mutation on objects returned from explicit runtime APIs inside `ctx.bound()` / middleware scope.

## Remaining proxy attributes (`MUTABLE_LEGACY_NAMES`)

**Inventory (2026-06-21, P4 WP3 A4):**

| Scope | `config_and_data as cfg` | `cfg.PLATES_*` / `cfg.PLATE_*` proxy |
|-------|--------------------------|--------------------------------------|
| `app/**` | **0** | **0** |
| `core/**` (excl. proxy module) | **0** | **0** |
| `viz_modules/**` | **0** | **0** (explicit `get_plate_mutable_runtime()`) |
| `bot_archived/**` | 6 files (named fn + paths; proxy: `optimize.py` only) | **archived** — off active inventory (P5 WP1) |
| `tests/**` (active) | ~15 files (isolation + parse/procurement fixtures) | proxy reads for semantics / legacy fixtures |
| `tests/archived/**` | bot tests (excluded via pytest.ini) | not collected |

Grep gate: `tests/test_config_and_data_proxy_boundary.py` — `app/` + `core/` must stay at **0**.

Narrowed proxy (**16** names, was **32** via full `MUTABLE_ATTR_MAP`):

| Group | Names still on PEP 562 proxy |
|-------|------------------------------|
| Plate width lists (bot — archived) | `PLATES_0_32`, `PLATES_0_46`, `PLATES_0_70`, `PLATES_0_72`, `PLATES_0_86`, `PLATES_0_88`, `PLATES_0_74`, `PLATES_0_48`, `PLATES_0_50`, `PLATES_0_34` |
| Plate width lists (tests) | `PLATES_1_2`, `PLATES_1_0`, `PLATES_1_08` |
| Metadata maps (tests / bot-adjacent) | `PLATE_LOAD_DETAILS`, `PLATE_EXACT_WIDTHS`, `PLATE_LENGTH_DM_RAW` |

**Removed from proxy (2026-06-21):** cuts/strips/waste totals (`LONGITUDINAL_CUTS`, `LENGTH_TRIMS`, `UNUSED_STRIPS_*`, `SCRAP_STRIPS_*`, `USABLE_STRIPS_*`, `WASTE_AREA_M2`), metadata caches (`PLATE_METADATA`, `PLATE_MAX_REINFORCEMENT_MAP`, `PLATE_NOMENCLATURE_CACHE`, `LAST_PARSE_DIAGNOSTICS`), unused list `PLATES_1_5_TO_1_2`. Access via `get_plate_mutable_runtime().<attr>` only.

`config_and_data.__dir__()` no longer advertises proxy names (web discovery off proxy).

Full map for migration reference remains in `MUTABLE_ATTR_MAP` (`core/plate_runtime_state.py`).

## Migration order (incremental PRs)

| Step | Area | Files (priority) | Notes |
|------|------|------------------|-------|
| 0 | **ADR** | this doc | Done in P2 WP1 prep |
| 1 | **app/** hot paths | `app/services/production_planning_service.py`, `app/services/commercial_service.py`, `app/services/plate_parser_service.py` | **Done (2026-06-20):** explicit imports + `PlateOrderContext`; zero `config_and_data as cfg` in `app/` |
| 2 | **core/production/** | `core/production/planning.py`, `core/plan_commit.py` | **Done (2026-06-21):** `planning.py` — explicit `normalize_load_code` from `core.domain.plate_order`; `plan_commit.py` — same; zero `config_and_data as cfg` in `core/production/` and `core/plan_commit.py` |
| 3 | **viz procurement** | `viz_modules/procurement/adapters_default.py`, `load_context.py`, `breakdown.py`, … | **Done (2026-06-21):** explicit `get_plate_mutable_runtime()`, `core.config.constants`, named imports; zero `config_and_data as cfg` in `viz_modules/procurement/` |
| 4 | **core/optimization** | `ilp_model.py`, `finalize.py`, `layout_runtime_snapshot.py`, … | **Done (2026-06-21):** `get_plate_mutable_runtime()`, `core.domain.plate_order.normalize_load_code`, `core.config.constants`, named imports; zero `config_and_data as cfg` in `core/optimization/` |
| 5 | **viz layout** | `viz_modules/layout_sequence/*` | **Done (2026-06-21):** `LayoutSequenceCfgSlice.from_plate_runtime()`, `build_layout_runtime_snapshot()`, `core.project_paths.PRICE_DB_PATH`, `core.domain.plate_order`; zero `config_and_data as cfg` in `viz_modules/layout_sequence/` |
| 6 | **Residual core** | `visualization.py`, `track_*`, `rescue_tracks.py`, … | **Done (2026-06-21):** `get_plate_mutable_runtime()`, `core.domain.plate_order`, `core.config.constants`, `core.project_paths`, named imports; zero `config_and_data as cfg` in `core/` |
| 7 | **Proxy removal** | `config_and_data.__getattr__` | **Partial (2026-06-21 P4 A4):** proxy narrowed 32→16 names; `__dir__` off proxy; grep gate `app/`+`core/`; DeprecationWarning retained |
| — | **bot_archived/** | archived P5 WP1 | Fix only if tests break; no handler refactors |

### Per-PR checklist

- [ ] Grep: `cfg\.(PLATES_|PLATE_LOAD|LONGITUDINAL)` in changed package → 0 new usages
- [ ] `pytest tests/test_plate_runtime_isolation.py tests/test_plate_runtime_request_isolation.py -q`
- [ ] `pytest tests/test_procurement_loads.py tests/test_config_and_data_plate_naming.py -q`
- [ ] If planning touched: `pytest tests/test_production_planning*.py tests/test_plan_*.py -q`
- [ ] No `bot/handlers` changes unless test-only minimal fix

## Full decommission criteria

`__getattr__` proxy may be **removed** when all are true:

1. `MUTABLE_LEGACY_NAMES` has **no** callers outside `config_and_data.py` and intentional compatibility tests
2. `app/**` and `core/production/**` use explicit runtime APIs only
3. Isolation tests green: `test_plate_runtime_*`, `test_procurement_*`
4. `__dir__` no longer advertises mutable legacy names (or module documents re-exports only)
5. Bot policy: bot archived (P5 WP1); **0 active bot consumers**; `__getattr__` removal — P6 after test-only backlog cleared

Named function imports (`make_plate_name`, `set_plate_lists_from_text`, …) may remain on `config_and_data` as a stable facade or move to submodules in a later cleanup.

## Rollback

| Risk | Mitigation |
|------|------------|
| Procurement/plan output drift | Small PRs; frozen fixtures in `tests/test_procurement_*`, `tests/test_plan_*` |
| Broken isolation | Revert single PR; middleware + `PlateOrderContext` unchanged |
| Missed caller | `rg "config_and_data as cfg"` before merge; deprecation `warnings.warn` on `__getattr__` optional gate |

**Rollback procedure:** revert the migration PR; proxy remains functional because `__getattr__` is unchanged until step 7.

## Verification commands

```bash
pytest tests/test_plate_runtime_isolation.py tests/test_plate_runtime_request_isolation.py tests/test_procurement_loads.py tests/test_config_and_data_plate_naming.py -q
pytest tests/test_production_planning*.py tests/test_plan_*.py -q
rg "config_and_data as cfg" app/ core/production/
```

## Changelog

| Date | Change |
|------|--------|
| 2026-06-20 | P2 WP1 — initial checklist / migration order |
| 2026-06-20 | Step 1 complete: `app/services/` migrated off module alias |
| 2026-06-21 | Step 6 complete: residual `core/` — visualization, track_top_up, track_reconciliation, rescue_tracks, reconciliation_xlsx, plates_preview_xlsx, kp_db_nomenclature, kp_db_schema |
| 2026-06-21 | P5 WP1: bot soft-decommission — `bot/` → `bot_archived/`; active bot PEP 562 consumers → **0**; proxy backlog tests-only |
