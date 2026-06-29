# Plan: core/optimization.py — God-module decomposition & security hardening

**Created:** 2026-05-05  
**Orchestration:** orch-2026-05-05-14-28-opt-refactor  
**Status:** ⏳ Planned  
**Priority:** Critical  
**Total Tasks:** 10  
**Estimated Total Time:** ~8–10 h

---

## Goal

Decompose the 2 700-line `core/optimization.py` god-module into a proper
sub-package (`core/optimization/`) with five focused modules, while:

1. Preserving **100 % backward compatibility** for all public import sites (no
   callers changed unless strictly necessary).
2. Eliminating **mutable global state** (`OPT_PLAN`, `OPT_CASCADING_PLAN`, …)
   that makes concurrent use unsafe.
3. Adding **input validation** at the public entry point and **redacting /
   gating sensitive data** in log output.

---

## Audit References

- [`ai_docs/develop/audits/2026-05-05-core-optimization-audit.md`](../audits/2026-05-05-core-optimization-audit.md) — full English audit (A1 god-module, security concerns)
- [`ai_docs/develop/audits/2026-05-05-core-optimization-audit.ru.md`](../audits/2026-05-05-core-optimization-audit.ru.md) — Russian summary
- [`ai_docs/develop/audits/2026-05-05-full-project-audit.md`](../audits/2026-05-05-full-project-audit.md) — project-wide context

---

## Target Package Layout

```
core/
└── optimization/               ← new sub-package (replaces optimization.py)
    ├── __init__.py             ← re-exports full public API (OPT-008)
    ├── geometry.py             ← slab sizes, cut options, narrowing (OPT-001)
    ├── ffd_packing.py          ← First-Fit-Decreasing track packing (OPT-002)
    ├── order_dispatch.py       ← order aggregation, canonical keys, slots (OPT-003)
    ├── ilp_model.py            ← PuLP variables, objective, constraints (OPT-004)
    └── orchestrator.py         ← top-level optimize_with_cascading_longitudinal_cuts (OPT-007)
```

---

## Public API that MUST stay importable from `core.optimization`

| Symbol | Type | Primary consumers |
|---|---|---|
| `optimize_with_cascading_longitudinal_cuts` | function | app, bot, tests, test_track_distribution, test_load_grouping |
| `verify_coverage` | function | tests/test_layout_secondary_*.py, tests/test_optimization_secondary_*.py |
| `_build_residual_balance_constraints` | function (semi-private) | tests/test_optimization_baseline.py |
| `OptimizationConfig` | dataclass/class | tests/experiments_compare.py |
| `OLD_CONFIG`, `DEFAULT_CONFIG` | module-level instances | tests/experiments_compare.py |
| `OPT_PLAN` | module global (→ thread-local accessor) | viz_modules/procurement.py, viz_modules/layout_sequence.py |
| `OPT_CASCADING_PLAN` | module global | viz_modules/procurement.py, viz_modules/layout_sequence.py |
| `OPT_CASCADING_PLAN_BY_LOAD` | module global | viz_modules/procurement.py |
| `OPT_WIDTH_PRIORITY` | module global | viz_modules/layout_sequence.py |
| `LOAD_TO_REINFORCEMENT_MAP` | module global | viz_modules/procurement.py |

---

## Dependencies Graph

```
OPT-001 ──────────────────────────────────────┐
OPT-002 (leaf) ────────────────────────────────┤
OPT-003 ←── OPT-001 ──────────────────────────┤
OPT-004 ←── OPT-001, OPT-003 ─────────────────┤
                                               ↓
OPT-005 ←── OPT-001..OPT-004 ─── OPT-006 ─→ OPT-007 ─→ OPT-008 ─→ OPT-009 ─→ OPT-010
```

---

## Progress

- ⏳ OPT-001: Extract geometry module (Pending)
- ⏳ OPT-002: Extract FFD packing module (Pending)
- ⏳ OPT-003: Extract order-dispatch module (Pending)
- ⏳ OPT-004: Extract ILP model builder (Pending)
- ⏳ OPT-005: Replace mutable globals with thread-safe context (Pending)
- ⏳ OPT-006: Add input validation and logging guards (Pending)
- ⏳ OPT-007: Create orchestrator module (Pending)
- ⏳ OPT-008: Create package `__init__.py` and retire old file (Pending)
- ⏳ OPT-009: Patch external call-sites and re-export shim (Pending)
- ⏳ OPT-010: Run full test suite and fix regressions (Pending)

---

## Tasks

---

### OPT-001 — Extract geometry module
**Priority:** Critical  
**Estimated time:** ~1 h  
**Dependencies:** None  
**Target file:** `core/optimization/geometry.py`

**Scope:** Move all code responsible for slab/plate geometry out of
`core/optimization.py`:

- Slab size/dimension lookup tables (constants and dicts).
- Functions that generate feasible cut-option lists given a plate dimension.
- Narrowing computation helpers (width tolerances, margin logic).
- `GeometryConfig` dataclass (if present) or equivalent named-tuple.

**Acceptance criteria:**

1. `core/optimization/geometry.py` exists and is importable standalone (`python -c "from core.optimization.geometry import ..."` succeeds).
2. All moved symbols have type-annotated signatures.
3. No circular imports: `geometry.py` must not import from other `core/optimization/` submodules.
4. Corresponding unit test (new or relocated) covers at least one representative cut-option generator.

---

### OPT-002 — Extract FFD packing module
**Priority:** Critical  
**Estimated time:** ~45 min  
**Dependencies:** None  
**Target file:** `core/optimization/ffd_packing.py`

**Scope:** Extract the First-Fit-Decreasing track-packing logic:

- The FFD algorithm function(s) (`pack_tracks` or equivalent).
- Any helper data-structures (track objects, bin containers) used only by FFD.

**Acceptance criteria:**

1. `core/optimization/ffd_packing.py` is importable standalone.
2. Public function signature: `pack_tracks(pieces: list[...]) -> list[Track]` (or equivalent typed signature).
3. No imports from `core/optimization/ilp_model.py` or `core/optimization/order_dispatch.py`  (pure leaf module).
4. Existing FFD-related test coverage still passes.

---

### OPT-003 — Extract order-dispatch module
**Priority:** High  
**Estimated time:** ~1.5 h  
**Dependencies:** OPT-001 (uses geometry config/constants)  
**Target file:** `core/optimization/order_dispatch.py`

**Scope:**

- Order aggregation from raw input into canonical representation.
- Canonical plate key construction (currently delegated to `core.config_and_data.canonical_plate_key`  — keep the delegation, just import it correctly).
- Slot attribution logic (`order_info_list` building, KP order dispatch).
- Any helpers that translate between user-facing order dicts and the internal slot/piece representation fed to ILP.

**Acceptance criteria:**

1. `core/optimization/order_dispatch.py` importable standalone.
2. All public functions have type-annotated parameters and return types.
3. `canonical_plate_key` is still sourced from `core.config_and_data`, not duplicated.
4. No direct PuLP imports (ILP concerns stay in `ilp_model.py`).

---

### OPT-004 — Extract ILP model builder
**Priority:** Critical  
**Estimated time:** ~1.5 h  
**Dependencies:** OPT-001, OPT-003  
**Target file:** `core/optimization/ilp_model.py`

**Scope:**

- PuLP `LpVariable` declarations.
- Objective function construction.
- All constraint builders, **including** `_build_residual_balance_constraints` (keep the name, keep it importable from `core.optimization` for test backward-compat).
- Return type: structured result (e.g. `ILPModel` dataclass: `problem`, `variables`, `metadata`).

**Acceptance criteria:**

1. `core/optimization/ilp_model.py` importable standalone.
2. `_build_residual_balance_constraints` importable from `core.optimization` without modification to existing test file.
3. `pulp` is imported only inside this module (not leaked to other submodules).
4. Existing `test_optimization_baseline.py` passes unchanged.

---

### OPT-005 — Replace mutable globals with thread-safe context
**Priority:** Critical  
**Estimated time:** ~1 h  
**Dependencies:** OPT-001, OPT-002, OPT-003, OPT-004  
**Target file:** `core/optimization/context.py` (new); updates to `orchestrator.py` and `__init__.py`

**Scope:**

Replace the five dangerous module-level mutable globals:

| Old global | Problem |
|---|---|
| `OPT_PLAN` | Overwritten on every call; data race under concurrency |
| `OPT_CASCADING_PLAN` | Same |
| `OPT_CASCADING_PLAN_BY_LOAD` | Same |
| `OPT_WIDTH_PRIORITY` | Same |
| `LOAD_TO_REINFORCEMENT_MAP` | Same |

**Implementation strategy:**

1. Create `core/optimization/context.py` with a `threading.local()` store and
   typed accessor functions: `get_opt_plan()`, `set_opt_plan(v)`, etc.
2. The orchestrator sets these after each solve; viz modules read them via the
   accessors.
3. In `core/optimization/__init__.py` expose the globals as **module-level
   properties** (using a module `__getattr__` shim) so that existing code like
   `from core.optimization import OPT_PLAN` still reads the thread-local
   value.

**Acceptance criteria:**

1. `OPT_PLAN` etc. still importable from `core.optimization` by existing viz modules without code change.
2. Two threads running `optimize_with_cascading_longitudinal_cuts` concurrently do not share state (verified by a simple threading test or doctest comment).
3. No bare module-level `OPT_PLAN = None` assignments remain in any submodule.

---

### OPT-006 — Add input validation and logging guards
**Priority:** High  
**Estimated time:** ~1 h  
**Dependencies:** OPT-005  
**Target file:** `core/optimization/orchestrator.py` (entry point); `core/optimization/context.py` or new `core/optimization/logging_utils.py`

**Scope:**

**Input validation** at `optimize_with_cascading_longitudinal_cuts`:

- Validate that `orders` is a non-empty list (raise `ValueError` with descriptive message).
- Validate that individual order dicts contain required keys (configurable required-key set).
- Validate numeric fields are within plausible bounds (e.g. slab length > 0, ≤ some `MAX_SLAB_MM` constant).
- Validate that total order count does not exceed `MAX_ORDER_COUNT` (e.g. 500 or from config).

**Logging guards:**

- Identify all `logger.debug` / `logger.info` calls that emit full order data or price-sensitive fields.
- Gate sensitive fields behind a `OPTIMIZATION_LOG_SENSITIVE=false` env-var / config flag.
- Add a redaction helper `redact_order(order: dict) -> dict` that strips price/client fields before logging.

**Acceptance criteria:**

1. `optimize_with_cascading_longitudinal_cuts(orders=[])` raises `ValueError`, not an obscure downstream exception.
2. With `OPTIMIZATION_LOG_SENSITIVE=false` (default), log output contains no price or client-name fields.
3. Existing tests pass (validation thresholds set conservatively so they never trigger on valid test fixtures).

---

### OPT-007 — Create orchestrator module
**Priority:** Critical  
**Estimated time:** ~1.5 h  
**Dependencies:** OPT-004, OPT-005, OPT-006  
**Target file:** `core/optimization/orchestrator.py`

**Scope:**

- Move `optimize_with_cascading_longitudinal_cuts` (and any cascade/retry logic) here.
- Compose calls to `geometry`, `order_dispatch`, `ilp_model`, `ffd_packing`, `context`.
- `verify_coverage` function lives here (or in a thin `coverage.py`; keep importable from `core.optimization`).
- Return a typed `PlanResult` (or existing named structure) instead of relying on global side-effects.

**Acceptance criteria:**

1. `orchestrator.py` contains a single top-level function `optimize_with_cascading_longitudinal_cuts` with the same signature as the original.
2. `verify_coverage` importable from `core.optimization` without caller changes.
3. Internal composition is readable: each major step (geometry setup, dispatch, ILP solve, FFD pack, coverage check) is a separate function call, not inlined monolith.
4. `OPT_PLAN` etc. are written via context accessors (OPT-005), not bare global assignments.

---

### OPT-008 — Create package `__init__.py` and retire old file
**Priority:** Critical  
**Estimated time:** ~45 min  
**Dependencies:** OPT-007  
**Target files:** `core/optimization/__init__.py` (create), `core/optimization.py` (delete)

**Scope:**

1. Create `core/optimization/__init__.py` that re-exports **all public symbols** listed in the "Public API" table above.
2. Include `module __getattr__` for the globals so `from core.optimization import OPT_PLAN` returns the thread-local value at access time.
3. Delete `core/optimization.py` (the old monolith file).
4. Verify `core/__init__.py` import (`from . import optimization`) still resolves correctly.

**Acceptance criteria:**

1. `python -c "from core.optimization import optimize_with_cascading_longitudinal_cuts, verify_coverage, OPT_PLAN, OptimizationConfig"` succeeds.
2. `python -c "from core.optimization import _build_residual_balance_constraints"` succeeds (test backward-compat).
3. No `ImportError` anywhere in the project after deletion of old file (verified by `python -m py_compile` on all affected files).

---

### OPT-009 — Patch external call-sites and re-export shim
**Priority:** High  
**Estimated time:** ~30 min  
**Dependencies:** OPT-008  
**Affected files:** `viz_modules/procurement.py`, `viz_modules/layout_sequence.py`, `bot/handlers/commercial.py`, `app/services/optimization_service.py`, `core/__init__.py`

**Scope:**

Review every external import of `core.optimization` globals. Where the existing
`module.__getattr__` shim in `__init__.py` handles them, no change is needed.
Where a caller stores a module-level reference at import time (e.g.
`from core.optimization import OPT_PLAN; ... use OPT_PLAN later`) and would
therefore get a stale `None`, update the call-site to use the accessor:

```python
# Before (stale reference anti-pattern)
from core.optimization import OPT_PLAN
# … later …
data = OPT_PLAN["some_key"]

# After
from core.optimization import get_opt_plan
data = get_opt_plan()["some_key"]
```

**Acceptance criteria:**

1. All files in `viz_modules/` import correctly with no `AttributeError` at runtime.
2. `bot/handlers/commercial.py` and `app/services/optimization_service.py` import successfully.
3. `core/__init__.py` `from . import optimization` resolves to the new package.

---

### OPT-010 — Run full test suite and fix regressions
**Priority:** Critical  
**Estimated time:** ~1 h  
**Dependencies:** OPT-009  
**Scope:** Verification, no new code.

Run the complete test suite and address any failures introduced by the refactor:

```powershell
python -m pytest tests/ -x -q 2>&1 | Tee-Object test_results.txt
```

Also run:
- `test_track_distribution.py` (root-level)
- `test_load_grouping.py` (root-level)
- `tests/test_optimization_baseline.py`
- `tests/test_optimization_secondary_parent_assignment.py`
- `tests/test_layout_secondary_null_parent_large_order.py`
- `tests/test_layout_secondary_unmatched_parent_user_list.py`
- `tests/test_visualization.py`

**Acceptance criteria:**

1. All previously-passing tests continue to pass.
2. No new `ImportError`, `AttributeError`, or `NameError` in any test.
3. Linter (`ruff` or `flake8`) reports zero new errors in changed files.
4. If any test was already broken before this cycle, document it in `ai_docs/develop/issues/` (not a regression of this plan).

---

## Architecture Decisions

1. **Package-over-file**: Converting `optimization.py` to a package is the
   standard Python approach for splitting a large module without breaking any
   existing `from core.optimization import X` statement.

2. **`threading.local()` for globals**: Chosen over passing context explicitly
   through every call frame because the public function signature cannot change
   without modifying all callers. `threading.local()` is a safe, standard
   library primitive; no extra dependency needed.

3. **`module.__getattr__`**: Allows `from core.optimization import OPT_PLAN`
   to work as a dynamic lookup at access time, consistent with the thread-local
   approach, without requiring callers to change to `get_opt_plan()` unless
   they hold a stale reference (see OPT-009).

4. **Leaf-modules-first ordering**: OPT-001 and OPT-002 are extracted first
   because they have no internal dependencies; this lets OPT-003 and OPT-004
   import cleanly and reduces the risk of circular imports.

5. **Old file stays intact until OPT-008**: Modules OPT-001 through OPT-007
   create *new* files without touching `core/optimization.py`. The switchover
   is atomic in OPT-008, making it easy to roll back if needed.
