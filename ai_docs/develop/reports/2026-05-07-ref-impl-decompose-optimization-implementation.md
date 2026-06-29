# Report: `_implementation.py` decomposition (ref‑impl facade)

**Date:** 2026-05-07  
**Orchestration:** `orch-2026-05-07-15-00-ref-impl-decompose`  
**Status:** Completed  

## Summary — goal

Replace the **`core/optimization/_implementation.py` monolith** (~1900+ LOC) with a **thin import/re‑export façade** (~100 LOC) plus **focused modules** under `core/optimization/`. Preserve the **existing public surface** for `core.optimization` (`from ._implementation import *` / stable `__all__`) and the **orchestrator contract** documented in `core/optimization/orchestrator.py`.

**Canonical plan:** `ai_docs/develop/plans/2026-05-07-refactor-optimization-implementation.md`

## Completed tasks (OPT‑REF‑001 … OPT‑REF‑010)

| ID | Name | Dependencies | Outcome |
|----|------|--------------|---------|
| **OPT‑REF‑001** | Extract `OptimizationConfig` | — | `optimization_config.py`: `OptimizationConfig`, `DEFAULT_CONFIG`, `OLD_CONFIG`. |
| **OPT‑REF‑002** | Extract coverage + PuLP qty helpers | — | `coverage_verify.py` (`verify_coverage`), `pulp_qty.py` (`_opt_1d_pulp_nonneg_qty`). |
| **OPT‑REF‑003** | Consolidate implementation debug helpers | — | `optimization_debug_impl.py` (alongside extended `debug_log.py` usage per plan); log `location` strings updated to match real call sites. |
| **OPT‑REF‑004** | Extract legacy width plan adapters | — | `legacy_width_plan.py`: `_group_plate_lengths`, `_append_actions`, `apply_width_optimization`, `optimize_cuts_pulp`. |
| **OPT‑REF‑005** | Extract secondary `z_sec` batch sizing | — | `secondary_batches.py`: `_batch_sizes_for_secondary_z_sec`. |
| **OPT‑REF‑006** | Extract 1D widths optimizer | OPT‑REF‑002 | `optimize_1d_widths.py`: `_optimize_1d_widths_only`. |
| **OPT‑REF‑007** | Extract 2D phase: prep + ILP solve | OPT‑REF‑001 | `optimize_2d/prep_solve.py`: `run_two_d_phase_a`. |
| **OPT‑REF‑008** | Extract 2D phase: cuts + ordering + parents | OPT‑REF‑007, OPT‑REF‑005 | `optimize_2d/extract_cuts.py`: `extract_two_d_phase_b` (ties in secondary batches). |
| **OPT‑REF‑009** | Extract 2D phase: post‑correction + audit + attribution | OPT‑REF‑008, OPT‑REF‑002 | `optimize_2d/finalize.py`: `run_two_d_phase_finalize` (coverage verification, assignments, residuals). |
| **OPT‑REF‑010** | Slim façade + stable `__all__` | OPT‑REF‑009, 006, 004, 003 | `_implementation.py` aggregates imports/`__all__`; orchestrates only via delegated symbols (e.g. `_optimize_2d_with_lengths` from `optimize_2d/with_lengths.py`). |

## New / primary module map (`core/optimization/`)

| Path | Responsibility |
|------|----------------|
| `_implementation.py` | Thin façade: re‑exports TLS/context, geometry/ILP/FFD symbols, delegates 2D entry via `optimize_2d/with_lengths`; **last-line** pull-in of `optimize_with_cascading_longitudinal_cuts` from orchestrator (see cyclic‑import notes). |
| `optimization_config.py` | Configuration dataclass and defaults. |
| `coverage_verify.py` | Demand/coverage verification (`verify_coverage`). |
| `pulp_qty.py` | PuLP nonnegative quantity helper for 1D path. |
| `optimization_debug_impl.py` | Debug instrumentation consolidated from former monolith sections. |
| `legacy_width_plan.py` | Legacy width pipeline / PuLP width helpers. |
| `secondary_batches.py` | Secondary `z_sec` batch sizing (`_batch_sizes_for_secondary_z_sec`). |
| `optimize_1d_widths.py` | Modern 1D width-only optimizer (`_optimize_1d_widths_only`). |
| `optimize_2d/__init__.py` | Subpackage exports: phase A/B entry points + `TwoDPhaseAState`, `norm_demand_key`. |
| `optimize_2d/state.py` | Shared state / normalization for 2D phases. |
| `optimize_2d/prep_solve.py` | Phase A: prep + build/solve ILP. |
| `optimize_2d/extract_cuts.py` | Phase B: extract primary/secondary cuts, parents/order rules. |
| `optimize_2d/finalize.py` | Phase C: audit checkpoints, normalization/post‑correction, `verify_coverage`, slot assignment. |
| `optimize_2d/with_lengths.py` | Composes phases A→B→C: `_optimize_2d_with_lengths`. |

Existing neighbors unchanged in role: **`orchestrator.py`** (public entry + lazy package import), **`context.py`** (TLS + `OPT_*`), **`ilp_model.py`**, **`ffd_packing.py`**, **`order_dispatch.py`**, **`geometry.py`**, **`validation.py`**, **`result_contract.py`**, **`debug_log.py`**.

## Regression tests (recommended suite)

Run from the repo root **with the project virtualenv / full `requirements.txt` installed**:

```bash
pytest tests/test_optimization*.py tests/test_opt_1d_pulp_qty_extraction.py -v
```

**Files explicitly in scope:**

- `tests/test_optimization_validation.py`
- `tests/test_optimization_baseline.py`
- `tests/test_optimization_config.py`
- `tests/test_optimization_result_contract.py`
- `tests/test_optimization_secondary_parent_assignment.py`
- `tests/test_optimization_semantics_and_tracks.py`
- `tests/test_optimization_thread_local_globals.py`
- `tests/test_optimization_verify_pulp_submodules_ref002.py`
- `tests/test_opt_1d_pulp_qty_extraction.py`

**Documentation-time verification:** Automated execution in this environment was **not** fully reproducible (`venv` absent here; bare Python lacked full dependency tree — e.g. `matplotlib` required via `core/__init__.py` → `visualization`). Re‑run in your Windows venv before release.

## Risks and cyclic import posture

1. **`core.optimization.__init__.py`** eagerly does `from ._implementation import *`, so `_implementation` loads early for any `import core.optimization`.
2. **`_implementation.py`** imports **`optimize_with_cascading_longitudinal_cuts`** from **`orchestrator.py`** at **module tail** (`# Последним: оркестратор подтягивает пакет лениво внутри API`).
3. **`orchestrator.py`** avoids importing the package at module import time inside the heavy path; it performs **`import core.optimization as pkg` inside** `optimize_with_cascading_longitudinal_cuts` when delegating to `pkg._optimize_2d_with_lengths` / `pkg._optimize_1d_widths_only`.

**Operational rule:** Avoid adding **top-level** `import core.optimization` (or imports that chain back to `_implementation`) inside **`orchestrator.py`** or other modules imported by `_implementation` before the façade finishes initializing — that would recreate an import cycle. Prefer **lazy local imports** in call paths (as orchestrator already does).

**Submodules:** `optimize_2d/*` imports should stay **one-way** toward `ilp_model`, `geometry`, `secondary_batches`, `optimization_config`, etc., and avoid pulling `_implementation` back in.

## Related documentation

- Plan: [`ai_docs/develop/plans/2026-05-07-refactor-optimization-implementation.md`](../plans/2026-05-07-refactor-optimization-implementation.md)
- Workspace: `.cursor/workspace/active/orch-2026-05-07-15-00-ref-impl-decompose/` (`tasks.json`, `progress.json`, `links.json`)

## Next steps (optional)

- Run the full **`tests/test_optimization*.py`** + **`tests/test_opt_1d_pulp_qty_extraction.py`** suite in CI/project venv after merges.
- If `links.json` automation is extended, wire **`report`** → this file path for discoverability.
