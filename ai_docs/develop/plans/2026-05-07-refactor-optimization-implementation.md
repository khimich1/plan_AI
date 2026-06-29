# Plan: Decompose `core/optimization/_implementation.py`

**Created:** 2026-05-07  
**Orchestration:** `orch-2026-05-07-15-00-ref-impl-decompose`  
**Goal:** Split ~1900+ LOC `_implementation.py` into focused modules under `core/optimization/`, keep stable public surface (`core.optimization` / `from ._implementation import *`).

## Current `_implementation.py` snapshot (baseline)

**Imports:** `logging`, `pathlib.Path`, typing, `core.config_and_data`, `canonical_plate_key`, `debug_paths`, geometry (`GeometryConfig`, cut option generators), `ffd_packing`, `debug_log`, `order_dispatch`, `ilp_model`, `logging_utils`, `result_contract`; mid-file re-export of TLS globals from `context.py`.

**Top-level constructs (by section comments):**

| Section | Approx. LOC focus | Symbols |
|---------|-------------------|---------|
| PuLP helper | Early | `_opt_1d_pulp_nonneg_qty` |
| Helpers | `# ХЕЛПЕРЫ` | `verify_coverage` |
| Debug | Embedded | `_debug_runtime_write_648532`, several `_DEBUG_*` path constants |
| Config | `# КОНФИГУРАЦИЯ` | `OptimizationConfig`, `DEFAULT_CONFIG`, `OLD_CONFIG` |
| TLS | `# ГЛОБАЛЬНЫЕ...` | Re-import `OPT_*` from `context` (not duplicated) |
| Legacy | `# ЛЕГАСИ-АДАПТЕРЫ` | `_group_plate_lengths`, `_append_actions`, `apply_width_optimization`, `optimize_cuts_pulp` |
| Modern | `# СОВРЕМЕННЫЕ` | `_batch_sizes_for_secondary_z_sec`, `_optimize_2d_with_lengths` (~majority), `_optimize_1d_widths_only`, tail import `optimize_with_cascading_longitudinal_cuts` |

**Consumers:**  
- `core/optimization/__init__.py`: `from ._implementation import *` + `__all__` (same).  
- `orchestrator.py`: lazy `import core.optimization as pkg` → `pkg._optimize_2d_with_lengths`, `pkg._optimize_1d_widths_only` (avoid importing heavy graph at orchestrator load).  
- Direct: `tests/test_opt_1d_pulp_qty_extraction.py`, `tests/test_optimization_secondary_parent_assignment.py` import `_implementation`.

## Dependencies graph

```
OPT-REF-001 ──┐
              ├──► OPT-REF-007 ──► OPT-REF-008 ──► OPT-REF-009 ──┐
OPT-REF-005 ──┘                                                  │
              ┌──────────────────────────────────────────────────┤
OPT-REF-002 (parallel) ─────────────────────────────► OPT-REF-009
OPT-REF-003, OPT-REF-004 (parallel)
OPT-REF-006 depends on OPT-REF-002
OPT-REF-010 after: 009 + 006 + 004 + 003
```

## Proposed layout (incremental — names can be shortened)

Suggested new modules (all under `core/optimization/`):

- `optimization_config.py` — dataclass + `DEFAULT_CONFIG` / `OLD_CONFIG`
- `coverage_verify.py` — `verify_coverage`
- `pulp_qty.py` — `_opt_1d_pulp_nonneg_qty`
- Extend `debug_log.py` OR add `optimization_debug_impl.py` — session writers + `_DEBUG_*` paths currently local to `_implementation.py`
- `legacy_width_plan.py` — `_group_*`, `_append_actions`, `apply_width_optimization`, `optimize_cuts_pulp`
- `secondary_batches.py` — `_batch_sizes_for_secondary_z_sec` (exported name unchanged via package)
- `optimize_1d_widths.py` — `_optimize_1d_widths_only`
- **`optimize_2d/` package** OR three siblings (preferred if you avoid packages): introduce a small **`optimize_2d/state.py`** (dataclasses) shared by phases:

  - **`optimize_2d/prep_solve.py`**: normalize demand, slot ledger setup, geometry options, `build_two_d_cutting_ilp`, `prob.solve`, infeasible/undefined branch, slack/unmet diagnostics. **Output:** frozen/dataclass carrying `demand_2d`, `slot_lists`, `slot_cursors`, `ilp`, solver status strings, decrypted handles (`z_prim`, `x_sec`, dicts maps).
  - **`optimize_2d/extract_cuts.py`**: build `planned_primary_*`, sorting for factory rules, `z_sec` loop with **`_batch_sizes_for_secondary_z_sec`**.
  - **`optimize_2d/finalize.py`**: PlateAudit checkpoints, `_norm_key`/post-correction, `no_sources_keys` force-add, `verify_coverage`, primary/secondary `plate_assignments` via `_next_slot_info`, residual debug logs.

`_implementation.py` becomes a **thin facade**: imports submodules and defines `_optimize_2d_with_lengths` as orchestrating the three phases (or re-exports a single composed function from `optimize_2d/__init__.py`).

---

## Tasks (≤10)

### OPT-REF-001 — Extract `OptimizationConfig`  
**Priority:** High  
**Complexity:** Simple  
**Files:** ADD `optimization_config.py`; CHG `_implementation.py` (imports + removals)

**Acceptance criteria**

- `from core.optimization import OptimizationConfig, DEFAULT_CONFIG, OLD_CONFIG` unchanged.
- No behavior change; no new cyclic imports (module must not import `_implementation`).
- Lint clean on touched files.

---

### OPT-REF-002 — Extract `verify_coverage` + `_opt_1d_pulp_nonneg_qty`  
**Priority:** High  
**Complexity:** Simple  
**Files:** ADD `coverage_verify.py`, `pulp_qty.py` (or single `optimization_helpers_shared.py` if you prefer fewer files — **pick one**, avoid arbitrary splits); CHG `_implementation.py`, CHG tests that referenced `_implementation` for qty helper if imports move.

**Acceptance criteria**

- `pytest tests/test_opt_1d_pulp_qty_extraction.py` passes unchanged semantically (update import path if needed).
- `verify_coverage` still importable via `core.optimization` / same `__all__` entries.

---

### OPT-REF-003 — Consolidate debug instrumentation tied to `_implementation`  
**Priority:** Medium  
**Complexity:** Moderate  
**Files:** CHG `debug_log.py` and/or NEW `optimization_debug_impl.py`; CHG `_implementation.py` (+ any relocated `#region agent log`).

**Acceptance criteria**

- `OPT_DEBUG` / `_opt_debug_enabled()` behavior unchanged.
- Logging line `location": "core/optimization/_implementation.py:...`** — update paths** to reflect real call site file after moves (helps future debugging).

---

### OPT-REF-004 — Extract legacy width pipeline  
**Priority:** Medium  
**Complexity:** Simple  
**Files:** ADD `legacy_width_plan.py`; CHG `_implementation.py`

**Acceptance criteria**

- `apply_width_optimization`, `optimize_cuts_pulp`, `_group_plate_lengths`, `_append_actions` remain on `core.optimization` package API.
- `optimize_cuts_pulp` still calls cascading optimizer and mutates `OPT_PLAN` on success — **integration smoke:** run any test or script that exercised `optimize_cuts_pulp` if exists; otherwise manual `optimize_cuts_pulp({320: 1})`-style sanity with PuLP mock/disabled expectation.

---

### OPT-REF-005 — Extract `_batch_sizes_for_secondary_z_sec`  
**Priority:** Medium  
**Complexity:** Simple  
**Files:** ADD `secondary_batches.py`; CHG `_implementation.py`; CHG `tests/test_optimization_secondary_parent_assignment.py` import path if directed re-export stays on package only.

**Acceptance criteria**

- `pytest tests/test_optimization_secondary_parent_assignment.py` passes.
- `_batch_sizes_for_secondary_z_sec` still importable where tests/docs expect (`core.optimization` or `_implementation` re-export).

---

### OPT-REF-006 — Extract `_optimize_1d_widths_only`  
**Priority:** High  
**Complexity:** Moderate  
**Files:** ADD `optimize_1d_widths.py`; CHG `_implementation.py`

**Acceptance criteria**

- `orchestrator` unchanged (still `pkg._optimize_1d_widths_only`); re-export on package preserved.
- Run: `pytest tests/test_optimization_semantics_and_tracks.py` (or subset covering 1D path), `tests/test_optimization_baseline.py` if runtime tolerable.
- 1D PuLP error codes (`ERROR_PULP_MISSING`, `ERROR_SOLVER_*`) unchanged.

---

### OPT-REF-007 — Extract 2D phase A: prep + build ILP + solve  
**Priority:** Critical  
**Complexity:** Complex  
**Files:** ADD `optimize_2d/` (at least `state.py`, `prep_solve.py`); CHG `_implementation.py`

**Acceptance criteria**

- Introduce explicit **state object** (dataclass) returned from phase A; no reliance on Python closure for cross-phase data except via that object.
- Phase A ends at same logical point as today: after solver status handling for Infeasible/Undefined (early return) or continuation with partial solve.
- **Smoke:** one 2D integration test path still passes — e.g. `tests/test_optimization_baseline.py` (pick representative case) or `tests/test_order.py` if lightweight.

---

### OPT-REF-008 — Extract 2D phase B: extract solution → primary/secondary planned rows + factory ordering + parent assignment  
**Priority:** Critical  
**Complexity:** Complex  
**Dependencies:** OPT-REF-007, OPT-REF-005  
**Files:** ADD `optimize_2d/extract_cuts.py`; CHG `_implementation.py` / wire

**Acceptance criteria**

- Order of `primary_cuts` (solids first, sorted) preserved.
- Secondary `parent_instance_id` behavior unchanged; **regression:** `tests/test_optimization_secondary_parent_assignment.py`, `tests/test_layout_secondary_null_parent_large_order.py` if applicable.

---

### OPT-REF-009 — Extract 2D phase C: post-correction, audit, coverage, slot attribution  
**Priority:** Critical  
**Complexity:** Complex  
**Dependencies:** OPT-REF-008, OPT-REF-002  
**Files:** ADD `optimize_2d/finalize.py`; CHG `_implementation.py`

**Acceptance criteria**

- `verify_coverage` attached to result; PlateAudit loss detection unchanged.
- `_coverage_summary['ok']` semantics preserved for same inputs.
- **Regression focus:** layout / null-parent tests, any test asserting `identity_match_type` or slot exhaustion logs.

---

### OPT-REF-010 — Slim `_implementation.py` + `__all__` integrity  
**Priority:** High  
**Complexity:** Simple (mechanical)  
**Dependencies:** OPT-REF-009, OPT-REF-006, OPT-REF-004, OPT-REF-003  
**Files:** CHG `_implementation.py`, possibly `__init__.py` if `__all__` moves; CHG `orchestrator.py` **only if** public attribute path for `_optimize_*` must change (should **not** change).

**Acceptance criteria**

- `_implementation.py` essentially: imports, re-exports, `from orchestrator import optimize_with_cascading_longitudinal_cuts`, `__all__` matches pre-refactor surface.
- `python -c "import core.optimization as o; assert hasattr(o, '_optimize_2d_with_lengths')"` succeeds.
- Full suite (or agreed subset): `pytest tests/test_optimization_*.py` + `tests/test_order.py` green.

---

## Risk notes

| Risk | Mitigation |
|------|------------|
| **Circular imports** | `orchestrator` must keep **lazy** `import core.optimization as pkg`. New submodules must **not** import `_implementation` or `__init__` package. Prefer `from core.optimization.xxx import y` with tree leaves depending only upward toward `geometry`, `ilp_model`, `result_contract`, etc. |
| **`from ._implementation import *`** | After split, either keep `__all__` in `_implementation.py` re-exporting submodules, or move `__all__` to `__init__.py` explicitly — **one source of truth**; update once. |
| **Grep / hardcoded paths** | Search `optimization.py:` and `_implementation.py` in log `location` strings after moves. |
| **Phase boundary mistakes** | Wrong split between REF-007/008/009 can drop variables (`_norm_key`, counters). Use a single `TwoDPhaseState` dataclass; type-check optional. |
| **Behavior drift in partial solver** | Preserve comments about Not Optimal / post-correction; do not change order of post-correction vs `no_sources_keys` fix. |

## Progress (orchestrator updates)

- ⏳ OPT-REF-001 … OPT-REF-010: Pending
