# Project Audit Report

**Date**: 2026-05-06  
**Scope**: `core/optimization/` + `viz_modules/layout_sequence.py`  
**Audited by**: senior-reviewer + security-auditor + reviewer (Composer 2)

---

## Executive Summary

**Overall Health Score**: 4.0/10

| Severity | Architecture | Security | Code Quality | Total |
|----------|-------------|----------|--------------|-------|
| Critical | 0 | 0 | 1 | **1** |
| High | 5 | 0 | 2 | **7** |
| Medium | 5 | 3 | 3 | **11** |
| Low | 3 | 2 | 2 | **7** |

**Recommendation**: Address the single Critical code-quality finding (silent solver/value extraction) before relying on cut quantities from `_implementation.py`; then prioritize architecture highs (monolith/coupling/unreachable branch) and medium security (debug logs).

---

## Critical Issues (fix immediately)

**[Q1]** — In `core/optimization/_implementation.py`, bare `except:` around extraction of PuLP solver values for 1D primary/secondary variables swallows every failure (including bad solver state or real bugs). Cuts can be **dropped without any signal**, so returned cut lists may under-represent reality while the rest of the pipeline continues — **silent corruption** of optimization output.

**Impact**: Wrong cut counts and downstream planning/visualization built on incomplete data.

**Fix direction**: Replace bare `except` with narrow exception handling (or explicit checks after solve), log or raise on unexpected failures, and add tests that assert behavior when variables are missing or infeasible.

---

## Architecture Findings (full text)

## Architecture Findings

### Critical
- None

### High
- **[A1]** `core/optimization/_implementation.py` — ~1.8k+ lines acting as a single "optimization monolith": PuLP wiring, option generation consumption, legacy width optimizer, coverage helpers, and cost heuristics (e.g. `plate_price = 12000` around lines 1797–1805) live together, which breaks SRP and makes changes high‑risk.
- **[A2]** `viz_modules/layout_sequence.py` (~1.78k lines) — one module owns reinforcement DB access (`Path(__file__).parent.parent / "pb.db"`, `get_reinforcement`), global/thread‑local plan consumption (`from core.optimization import OPT_PLAN, OPT_CASCADING_PLAN_BY_LOAD`, …), and full sequence construction; presentation/viz is tightly coupled to persistence and optimizer globals instead of a narrow, injected "plan DTO → sequence" service.
- **[A3]** `viz_modules/layout_sequence.py` — after `if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('primary_cuts'):` returns via `_build_sequence_from_plan` (lines 444–448), the next `if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('primary_cuts'):` (from ~450) is **unreachable** with the same condition; ~400+ lines of legacy branching (through ~1015) are dead, which is a major maintainability and audit hazard.
- **[A4]** `core/optimization/ilp_model.py` — imports `get_price` from `core.price_db` (line 14); the cutting ILP layer depends on pricing infrastructure, weakening the boundary between "geometry/assignment model" and "commercial DB".
- **[A5]** Circular package coupling: `core/optimization/orchestrator.py` lazily imports `core.optimization` as `pkg` (lines 39–48), while `core/optimization/_implementation.py` re-imports `optimize_with_cascading_longitudinal_cuts` from `orchestrator` at the bottom (line 1811). It works only because of import order/laziness; any refactor can reintroduce import cycles.

### Medium
- **[A6]** `core/optimization/__init__.py` — replaces the module's class via `sys.modules[__name__].__class__ = _OptimizationModule` (lines 30–30) to proxy `OPT_*` writes to TLS; powerful but non‑idiomatic and harder for tooling/static analysis than explicit context objects or `contextvars`.
- **[A7]** `core/optimization/context.py` — thread‑local proxies for `OPT_PLAN`, `OPT_CASCADING_PLAN`, etc. (lines 181–186) improve multi‑thread isolation vs process‑globals, but state is still **implicit** (callers mutate module‑level names); concurrent workloads on one thread (e.g. interleaved async tasks without a dedicated worker thread or reset) remain easy to get wrong compared to passing the plan dict through the call stack.
- **[A8]** `viz_modules/layout_sequence.py` — large duplicated narrative: `_build_sequence_from_plan` (from ~1089) repeats grouping/separator/secondary‑attach logic that overlaps the older `build_layout_sequence` body; divergence risk and review cost are high.
- **[A9]** `viz_modules/layout_sequence.py` — extensive ad‑hoc debug I/O (`_agent_seq_debug`, multiple `debug-*.log` paths, `print("[VISUAL] …")`) mixed with domain logic; no single logging abstraction or feature flag boundary.
- **[A10]** Scalability: the stack centers on synchronous PuLP/solver work inside `_implementation.py`/`ilp_model.py` with no obvious queue/worker boundary in these modules — vertical scaling only unless callers always offload to thread/process pools.

### Low
- **[A11]** `core/optimization/orchestrator.py` — uses `print` for mode selection (lines 43–50) instead of structured logging, unlike downstream modules that use `logging` in places.
- **[A12]** `core/optimization/geometry.py` vs `core/optimization/ffd_packing.py` — `ffd_packing.py` stays stdlib‑only and documented as isolated (lines 1–5); `geometry.py` still depends on `core.config_and_data` globally (line 7), a small layering leak compared to injecting config into generators.
- **[A13]** `viz_modules/layout_sequence.py` — sorts and mutates plan rows in place (e.g. adding `reinforcement` on `cut` dicts in solid/split loops ~558–573 / ~1280–1296), which can surprise callers if the same plan dict is reused elsewhere.

---

## Security Findings (full text)

## Security Findings

### Critical
None

### High
None

### Medium
- [S1] **Sensitive data written to disk without a debug gate** — `core/optimization/_implementation.py` and `viz_modules/layout_sequence.py` append NDJSON-style traces to fixed repo paths (e.g. `debug-7e420e.log`, agent logs, `debug-ef42ae.log`) **without** the `OPT_DEBUG_LOG` guard used in `core/optimization/debug_log.py` / `_dbg_open_append`. Payloads include truncated `plate_name`, assignment/instance identifiers, demand/shape keys, and secondary-cut structure → **commercial / identifying business data at rest** and uneven controls versus the rest of the optimizer package.
- [S2] **`ValueError` text may leak implementation detail** — `core/optimization/validation.py` raises `ValueError` with types and `repr`-style fragments (`got {type...}`, `{L!r}`, index paths). If an HTTP layer forwards these strings verbatim to clients, that supports **fingerprinting and clearer attack/debug mapping** (OWASP information disclosure / misconfiguration), even though this is not classic injection.
- [S3] **Thread-local `OPT_*` mirrors risky global semantics** — `core/optimization/__init__.py` + `context.py` keep optimization outputs in TLS-backed module proxies. Any pipeline that **reuses a worker thread** across unrelated jobs without resetting or copying state can surface **stale or mismatched plans** between logical sessions (integrity / wrong-artifact risk), especially where visualization reads globals (`layout_sequence` ← `core.optimization`) after optimization.

### Low
- [S4] **Inconsistent minimization of commercial fields in traces** — `viz_modules/layout_sequence.py` uses heavy `[VISUAL]` `print`/`logger` tracing (e.g. `[TRACE]` with `kp_id`) while `core/optimization/logging_utils.py` documents redaction for console order lines; posture is **mixed**, increasing odds of **PII/commercial fields in logs or operator-visible stdout**.
- [S5] **Accidental disclosure via `repr`** — `ThreadLocalDictProxy.__repr__` / `ThreadLocalListProxy.__repr__` in `context.py` embed full backing structures (`self._d()!r`); a stray `repr(OPT_CASCADING_PLAN)` in logs could dump **entire plan blobs**.

---

## Code Quality Findings (full text)

### Critical

- **[Q1]** Bare `except:` blocks around extraction of solver values for 1D primary/secondary quantities (`value(x_prim[i])`, `value(x_sec[i])`) in `core/optimization/_implementation.py` (~1732–1761). This swallows all failures (including legitimate bugs or bad solver state), **drops cuts silently**, and can return a cut list that understates reality while the rest of the pipeline keeps going — silent logical corruption of the optimization result.

---

### High

- **[Q2]** **API contract vs documentation drift:** `validate_optimize_entrypoint` docstring (`core/optimization/validation.py`, ~25–26) says empty `orders` and `orders_2d` together are not an error and references legacy `{}` behavior, while `orchestrator.optimize_with_cascading_longitudinal_cuts` returns `opt_error(ERROR_NO_INPUT, …)` when both are empty. Callers and future refactors can easily assume the wrong outcome class (new angle vs. raw `ValueError` leakage).

- **[Q3]** **Inconsistent observability in the visualization stack:** `viz_modules/layout_sequence.py` uses many `print("[VISUAL]…")` calls (e.g. `_choose_best_separator`, `_split_group_into_subgroups`) while `_build_sequence_from_plan` uses `logging`. Same feature area mixes styles, which hurts log filtering, production hygiene, and tests that need to assert on log output.

---

### Medium

- **[Q4]** **Test gap for newly split modules:** There are no focused unit tests for `core/optimization/geometry.py`, `ilp_model.py`, or `order_dispatch.py` (unlike `ffd_packing`, which has `tests/test_ffd_packing.py`). Regressions in option generation, ILP construction, or slot dispatch will mostly surface only through broad baselines/integration tests, increasing debug cost.

- **[Q5]** **Weak typing at important boundaries:** Examples include `CutOption = dict[str, Any]` and large `Any` surfaces in `geometry.py` / `ilp_model.py`, `optimize_tracks(items: list)` in `ffd_packing.py`, untyped public helpers in `layout_sequence.py` (`build_layout_sequence`, `_choose_best_separator`, …), and unstructured `dict` returns from `verify_coverage` in `_implementation.py`. The domain is combinatorial; the lack of `TypedDict`/protocols makes incorrect key access and silent shape drift more likely.

- **[Q6]** **Error handling as "black holes" for maintenance:** Very frequent `except Exception: pass` around debug and auxiliary blocks in `_implementation.py` and `order_dispatch.py` (and similar patterns in `layout_sequence.py`) means **instrumentation or side logic can fail permanently** with no trace — distinct from the security critique of NDJSON paths; this is about **lost signals when debugging optimization mismatches**.

---

### Low

- **[Q7]** **Magic sentinels** such as `999.0` / `999` for "missing reinforcement" in `layout_sequence.py` (`_choose_best_separator`, `_get_reinforcement_from_map` fallbacks, sorting keys) aren't centralized or named constants, inviting subtle inconsistencies if one path uses `999` and another uses `999.0` or different thresholds later.

- **[Q8]** **Nested helpers inside `build_layout_sequence`** (`_supplement_reinforcement_map_from_plan`, `plate_label`) increase cyclomatic complexity and block reuse in tests without importing the enclosing function or duplicating logic — a maintainability/testing smell even aside from overall file size.

---

## Priority Matrix

| ID | Issue | Severity | Effort guess | Priority |
|----|-------|----------|--------------|----------|
| Q1 | Bare `except` around PuLP `value()` extraction; silent cut drops | Critical | Medium | P0 — immediate |
| A1 | `_implementation.py` monolith / SRP violation | High (Arch) | High | P1 — this sprint |
| A2 | `layout_sequence.py` DB + globals + viz coupling | High (Arch) | High | P1 — this sprint |
| A3 | Unreachable duplicate branch; ~400+ lines dead | High (Arch) | Medium | P1 — this sprint |
| A4 | `ilp_model` → `price_db` boundary leak | High (Arch) | Medium | P2 — this sprint |
| A5 | Circular `orchestrator` ↔ package import risk | High (Arch) | Medium | P2 — this sprint |
| Q2 | `validation` doc vs `orchestrator` empty-input behavior | High (Quality) | Low | P2 — this sprint |
| Q3 | Mixed `print` vs `logging` in viz stack | High (Quality) | Medium | P2 — this sprint |
| S1 | Un-gated NDJSON / debug files; commercial data at rest | Medium (Sec) | Medium | P2 — this sprint |
| S2 | `ValueError` messages may leak via HTTP | Medium (Sec) | Low | P3 — next sprint |
| S3 | TLS `OPT_*` stale state across reused worker threads | Medium (Sec) | Medium | P3 — next sprint |
| A6 | `sys.modules` class replacement for TLS proxy | Medium (Arch) | High | Backlog |
| A7 | Implicit plan state vs explicit passing | Medium (Arch) | High | Backlog |
| A8 | Duplicated plan vs legacy sequence logic | Medium (Arch) | High | Backlog |
| Q4 | Missing unit tests for geometry / ILP / dispatch | Medium (Quality) | High | Backlog |

---

## Next Steps

1. **Immediate** — Fix **[Q1]**: remove bare `except` around solver value extraction; fail loud or log with context; add regression tests for extraction edge cases.
2. **This sprint** — Resolve **[A3]** (remove or wire dead branch; document intent), start **[A1]/[A2]** decomposition (extract services, inject plan DTOs), gate or relocate **[S1]** traces behind `OPT_DEBUG_LOG` (or equivalent), and align **[Q2]** (doc + `orchestrator` contract).
3. **Next sprint** — Address **[A4]/[A5]** (dependency boundaries, import graph), **[S2]** (sanitize map validation errors at API boundary), **[S3]** (explicit reset/copy of optimization context between jobs), and **[Q3]** (single logging strategy for `layout_sequence`).
4. **Backlog** — **[A6]–[A10]**, **[Q4]–[Q8]**, **[A11]–[A13]**, **[S4]–[S5]**, typing (`TypedDict`/protocols), magic-sentinel constants, and async/worker boundary for heavy solves per architecture notes.

---

## SOURCE FINDINGS TO INCLUDE VERBATIM

### Architecture (paste entire "## Architecture Findings" block from senior-reviewer):

## Architecture Findings

### Critical
- None

### High
- **[A1]** `core/optimization/_implementation.py` — ~1.8k+ lines acting as a single "optimization monolith": PuLP wiring, option generation consumption, legacy width optimizer, coverage helpers, and cost heuristics (e.g. `plate_price = 12000` around lines 1797–1805) live together, which breaks SRP and makes changes high‑risk.
- **[A2]** `viz_modules/layout_sequence.py` (~1.78k lines) — one module owns reinforcement DB access (`Path(__file__).parent.parent / "pb.db"`, `get_reinforcement`), global/thread‑local plan consumption (`from core.optimization import OPT_PLAN, OPT_CASCADING_PLAN_BY_LOAD`, …), and full sequence construction; presentation/viz is tightly coupled to persistence and optimizer globals instead of a narrow, injected "plan DTO → sequence" service.
- **[A3]** `viz_modules/layout_sequence.py` — after `if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('primary_cuts'):` returns via `_build_sequence_from_plan` (lines 444–448), the next `if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('primary_cuts'):` (from ~450) is **unreachable** with the same condition; ~400+ lines of legacy branching (through ~1015) are dead, which is a major maintainability and audit hazard.
- **[A4]** `core/optimization/ilp_model.py` — imports `get_price` from `core.price_db` (line 14); the cutting ILP layer depends on pricing infrastructure, weakening the boundary between "geometry/assignment model" and "commercial DB".
- **[A5]** Circular package coupling: `core/optimization/orchestrator.py` lazily imports `core.optimization` as `pkg` (lines 39–48), while `core/optimization/_implementation.py` re-imports `optimize_with_cascading_longitudinal_cuts` from `orchestrator` at the bottom (line 1811). It works only because of import order/laziness; any refactor can reintroduce import cycles.

### Medium
- **[A6]** `core/optimization/__init__.py` — replaces the module's class via `sys.modules[__name__].__class__ = _OptimizationModule` (lines 30–30) to proxy `OPT_*` writes to TLS; powerful but non‑idiomatic and harder for tooling/static analysis than explicit context objects or `contextvars`.
- **[A7]** `core/optimization/context.py` — thread‑local proxies for `OPT_PLAN`, `OPT_CASCADING_PLAN`, etc. (lines 181–186) improve multi‑thread isolation vs process‑globals, but state is still **implicit** (callers mutate module‑level names); concurrent workloads on one thread (e.g. interleaved async tasks without a dedicated worker thread or reset) remain easy to get wrong compared to passing the plan dict through the call stack.
- **[A8]** `viz_modules/layout_sequence.py` — large duplicated narrative: `_build_sequence_from_plan` (from ~1089) repeats grouping/separator/secondary‑attach logic that overlaps the older `build_layout_sequence` body; divergence risk and review cost are high.
- **[A9]** `viz_modules/layout_sequence.py` — extensive ad‑hoc debug I/O (`_agent_seq_debug`, multiple `debug-*.log` paths, `print("[VISUAL] …")`) mixed with domain logic; no single logging abstraction or feature flag boundary.
- **[A10]** Scalability: the stack centers on synchronous PuLP/solver work inside `_implementation.py`/`ilp_model.py` with no obvious queue/worker boundary in these modules — vertical scaling only unless callers always offload to thread/process pools.

### Low
- **[A11]** `core/optimization/orchestrator.py` — uses `print` for mode selection (lines 43–50) instead of structured logging, unlike downstream modules that use `logging` in places.
- **[A12]** `core/optimization/geometry.py` vs `core/optimization/ffd_packing.py` — `ffd_packing.py` stays stdlib‑only and documented as isolated (lines 1–5); `geometry.py` still depends on `core.config_and_data` globally (line 7), a small layering leak compared to injecting config into generators.
- **[A13]** `viz_modules/layout_sequence.py` — sorts and mutates plan rows in place (e.g. adding `reinforcement` on `cut` dicts in solid/split loops ~558–573 / ~1280–1296), which can surprise callers if the same plan dict is reused elsewhere.

### Security (paste entire "## Security Findings" block):

## Security Findings

### Critical
None

### High
None

### Medium
- [S1] **Sensitive data written to disk without a debug gate** — `core/optimization/_implementation.py` and `viz_modules/layout_sequence.py` append NDJSON-style traces to fixed repo paths (e.g. `debug-7e420e.log`, agent logs, `debug-ef42ae.log`) **without** the `OPT_DEBUG_LOG` guard used in `core/optimization/debug_log.py` / `_dbg_open_append`. Payloads include truncated `plate_name`, assignment/instance identifiers, demand/shape keys, and secondary-cut structure → **commercial / identifying business data at rest** and uneven controls versus the rest of the optimizer package.
- [S2] **`ValueError` text may leak implementation detail** — `core/optimization/validation.py` raises `ValueError` with types and `repr`-style fragments (`got {type...}`, `{L!r}`, index paths). If an HTTP layer forwards these strings verbatim to clients, that supports **fingerprinting and clearer attack/debug mapping** (OWASP information disclosure / misconfiguration), even though this is not classic injection.
- [S3] **Thread-local `OPT_*` mirrors risky global semantics** — `core/optimization/__init__.py` + `context.py` keep optimization outputs in TLS-backed module proxies. Any pipeline that **reuses a worker thread** across unrelated jobs without resetting or copying state can surface **stale or mismatched plans** between logical sessions (integrity / wrong-artifact risk), especially where visualization reads globals (`layout_sequence` ← `core.optimization`) after optimization.

### Low
- [S4] **Inconsistent minimization of commercial fields in traces** — `viz_modules/layout_sequence.py` uses heavy `[VISUAL]` `print`/`logger` tracing (e.g. `[TRACE]` with `kp_id`) while `core/optimization/logging_utils.py` documents redaction for console order lines; posture is **mixed**, increasing odds of **PII/commercial fields in logs or operator-visible stdout**.
- [S5] **Accidental disclosure via `repr`** — `ThreadLocalDictProxy.__repr__` / `ThreadLocalListProxy.__repr__` in `context.py` embed full backing structures (`self._d()!r`); a stray `repr(OPT_CASCADING_PLAN)` in logs could dump **entire plan blobs**.

### Code Quality (paste entire "## Code Quality Findings" block from reviewer):

## Code Quality Findings

### Critical

- **[Q1]** Bare `except:` blocks around extraction of solver values for 1D primary/secondary quantities (`value(x_prim[i])`, `value(x_sec[i])`) in `core/optimization/_implementation.py` (~1732–1761). This swallows all failures (including legitimate bugs or bad solver state), **drops cuts silently**, and can return a cut list that understates reality while the rest of the pipeline keeps going — silent logical corruption of the optimization result.

---

### High

- **[Q2]** **API contract vs documentation drift:** `validate_optimize_entrypoint` docstring (`core/optimization/validation.py`, ~25–26) says empty `orders` and `orders_2d` together are not an error and references legacy `{}` behavior, while `orchestrator.optimize_with_cascading_longitudinal_cuts` returns `opt_error(ERROR_NO_INPUT, …)` when both are empty. Callers and future refactors can easily assume the wrong outcome class (new angle vs. raw `ValueError` leakage).

- **[Q3]** **Inconsistent observability in the visualization stack:** `viz_modules/layout_sequence.py` uses many `print("[VISUAL]…")` calls (e.g. `_choose_best_separator`, `_split_group_into_subgroups`) while `_build_sequence_from_plan` uses `logging`. Same feature area mixes styles, which hurts log filtering, production hygiene, and tests that need to assert on log output.

---

### Medium

- **[Q4]** **Test gap for newly split modules:** There are no focused unit tests for `core/optimization/geometry.py`, `ilp_model.py`, or `order_dispatch.py` (unlike `ffd_packing`, which has `tests/test_ffd_packing.py`). Regressions in option generation, ILP construction, or slot dispatch will mostly surface only through broad baselines/integration tests, increasing debug cost.

- **[Q5]** **Weak typing at important boundaries:** Examples include `CutOption = dict[str, Any]` and large `Any` surfaces in `geometry.py` / `ilp_model.py`, `optimize_tracks(items: list)` in `ffd_packing.py`, untyped public helpers in `layout_sequence.py` (`build_layout_sequence`, `_choose_best_separator`, …), and unstructured `dict` returns from `verify_coverage` in `_implementation.py`. The domain is combinatorial; the lack of `TypedDict`/protocols makes incorrect key access and silent shape drift more likely.

- **[Q6]** **Error handling as "black holes" for maintenance:** Very frequent `except Exception: pass` around debug and auxiliary blocks in `_implementation.py` and `order_dispatch.py` (and similar patterns in `layout_sequence.py`) means **instrumentation or side logic can fail permanently** with no trace — distinct from the security critique of NDJSON paths; this is about **lost signals when debugging optimization mismatches**.

---

### Low

- **[Q7]** **Magic sentinels** such as `999.0` / `999` for "missing reinforcement" in `layout_sequence.py` (`_choose_best_separator`, `_get_reinforcement_from_map` fallbacks, sorting keys) aren't centralized or named constants, inviting subtle inconsistencies if one path uses `999` and another uses `999.0` or different thresholds later.

- **[Q8]** **Nested helpers inside `build_layout_sequence`** (`_supplement_reinforcement_map_from_plan`, `plate_label`) increase cyclomatic complexity and block reuse in tests without importing the enclosing function or duplicating logic — a maintainability/testing smell even aside from overall file size.
