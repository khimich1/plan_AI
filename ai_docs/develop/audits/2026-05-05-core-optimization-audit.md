# Consolidated Audit Report: `core/optimization.py`

**Date:** 2026-05-05  
**Audited by:** Senior Reviewer, Security Auditor, Code Quality Reviewer  
**Orchestration:** Full project audit (orchestrated review)

---

## Executive Summary

`core/optimization.py` is a critical 2.7k-line module serving the core business logic of order handling and ILP-based cutting optimization. The audit identified **1 Critical**, **7 High**, **11 Medium**, and **11 Low** severity findings across three domains: architecture, security, and code quality.

### Overall Health Score: **4.0 / 10.0** 🔴

**Calculation:** 10 − (1 × 2) − (7 × 0.5, capped −3) − (11 × 0.1, capped −1) = **4.0**

### Severity Distribution

| Domain | Critical | High | Medium | Low | Total |
|--------|----------|------|--------|-----|-------|
| **Architecture** | 1 | 3 | 3 | 3 | 10 |
| **Security** | 0 | 2 | 3 | 3 | 8 |
| **Code Quality** | 0 | 2 | 5 | 5 | 12 |
| **TOTAL** | **1** | **7** | **11** | **11** | **30** |

### Key Findings

- **Critical architectural flaw:** "God module" combining ILP solver, geometry, order accounting, legacy 1D optimization, and FFD track packing into a single unit—violation of SRP; any change to either ILP or order handling affects both.
- **Security risks:** Shared global mutable state under concurrency, unbounded input processing, and business-sensitive logging without controls.
- **Code quality degradation:** Inconsistent typing, poor error handling (swallowed exceptions), and fragmented test coverage.

**Recommended Next Steps:** Execute `/refactor` workflow to decompose the module; implement `/implement` for security hardening and observability.

---

## Critical Issues

### [A1] God Module: ILP Solver, Geometry, Order Accounting, and Packing Combined

**Severity:** CRITICAL  
**Category:** Architecture  
**Scope:** `core/optimization.py` (entire file)

#### Description

`core/optimization.py` combines multiple autonomous problem domains into a single monolithic module:

1. **ILP model assembly and solving** – Variables `x_prim`, `x_sec`, `z_prim`, `z_sec`, `unmet`; objective function; constraints; PuLP solver invocation (`PULP_CBC_CMD`).
2. **Cutting geometry and options** – `NARROWING_TABLE`, generation of `primary_options` and `secondary_options` based on plate dimensions.
3. **Residual balance accounting** – Logic in `_build_residual_balance_constraints` to track physical inventory state.
4. **Order/quote allocation** – Demand aggregation via `order_info_list`, order attribution logic in `_get_next_order_info`, `_peek_order_info`, slot lists, slot distribution.
5. **Legacy 1D optimization** – Functions `apply_width_optimization`, `_optimize_1d_widths_only`, `optimize_cuts_pulp` for backward compatibility.
6. **Track packing via FFD** – Classes `Piece`, `Track`; algorithm `first_fit_decreasing`; integration with 2D result.
7. **Coverage verification** – Helper `verify_coverage`.

#### Root Cause

These domains share a single logical flow but have no clean abstraction boundaries. A change to ILP parameters ripples into order attribution logic; modifications to quote handling affect solver setup.

#### Impact

- **Maintainability:** High cognitive load; navigation and understanding of even simple fixes requires reading cross-domain context.
- **Testing:** Difficult to unit-test individual subsystems; integration tests are heavyweight and fragile.
- **Extensibility:** Adding new demand types, load classes, or solver constraints requires touching the core module; risk of regressions.
- **Concurrency:** Shared globals (`OPT_PLAN`, `OPT_CASCADE_PLAN`, etc.) make parallel optimization unsafe without external synchronization.

#### Recommendation

Decompose into separate modules:

1. **`core/optimization/ilp_model.py`** – Objective, variables, constraints; export `ILPModel(demand_2d, options, cfg) → LpProblem`.
2. **`core/optimization/order_dispatch.py`** – Order aggregation, canonical keys, slot attribution; export `OrderDispatcher(orders_2d) → slots`.
3. **`core/optimization/geometry.py`** – Plate dimensions, cut options, narrowing; export `GeometryConfig`.
4. **`core/optimization/ffd_packing.py`** – FFD track packing; export `pack_tracks(pieces) → tracks`.
5. **`core/optimization/orchestrator.py`** – High-level `optimize_cuts_pulp(orders, cfg) → PlanResult`; compose submodules.

**Effort:** High (2–3 days refactoring + test rewrites).  
**Urgency:** High (foundational; unblocks all other improvements).

---

## High Severity Issues

### [A2] Misalignment Between ILP Model and Order Identity

**Severity:** HIGH  
**Category:** Architecture  
**Scope:** `_optimize_2d_with_lengths`, `_get_next_order_info`

#### Description

Demand is aggregated into ILP model via key `canonical_plate_key(length, width, ...)`, which lumps all orders with the same dimensions together. The solver minimizes cost across volumes but treats them as interchangeable. Order identity (`kp_id`, client, material grade) is attached separately and post-hoc through mutable `qty_remaining` counters and fallback logic in `_get_next_order_info`.

Attribution matching is lossy: only by `(length, width)` and adjacent length tolerance `LEN_TOL`, without load class or grade constraints.

#### Impact

- Optimally solved model may produce results that violate commercial constraints (e.g., wrong client assigned to plate, mismatched material grade).
- Silent semantic errors: ILP sees uniform demand, human sees distinct orders.
- Audit trail unclear: which demand unit maps to which quote?

#### Recommendation

Embed order identity into the ILP model:

- Extend `canonical_plate_key` to include load class and material grade.
- Introduce binary variable `y_order[i]` for each order `i`; link demand aggregation to order-level constraints.
- Refactor `_get_next_order_info` to track orders as first-class entities in the model, not post-hoc labels.

**Effort:** Medium (1–2 days).  
**Urgency:** High (correctness; impacts commercial reporting).

---

### [A3] Dual Data Sources for Order Input

**Severity:** HIGH  
**Category:** Architecture  
**Scope:** `optimize_with_cascading_longitudinal_cuts`, `optimize_cuts_pulp`

#### Description

Two parallel entry points feed orders:

1. Public API `optimize_with_cascading_longitudinal_cuts(orders=None)` – Branches on `orders` (per-width aggregation from legacy `cfg`) vs `orders_2d` (explicit 2D input).
2. Fallback `optimize_cuts_pulp(orders=None)` – If `orders` is `None`, auto-populates from `cfg.PLATES_*` counters in `core/config_and_data.py`.

No single source of truth. Global config state and explicit function arguments are both treated as valid. This creates confusion and silent fallback behavior.

#### Impact

- Calling code unsure which input is authoritative.
- Difficult to reason about state mutations: did the optimizer consume my input or the global config?
- Unit tests and mocks become complex; hard to isolate state.

#### Recommendation

1. **Deprecate fallback:** Remove auto-population from `cfg.PLATES_*`; require explicit `orders` or `orders_2d` argument.
2. **Unify input:** Single canonical `orders_2d` structure; normalize legacy `orders` at the API boundary, not deep in the module.
3. **Document contract:** Clear docstring specifying that global config is read-only reference; all mutable state is function-local.

**Effort:** Medium (1–2 days).  
**Urgency:** High (clarity; unblocks safe refactoring).

---

### [A4] Module-Level Global Caches with Undocumented Side Effects

**Severity:** HIGH  
**Category:** Architecture  
**Scope:** `OPT_PLAN`, `OPT_CASCADING_PLAN`, `OPT_CASCADING_PLAN_BY_LOAD`, `LOAD_TO_REINFORCEMENT_MAP`, `OPT_WIDTH_PRIORITY`

#### Description

Results and intermediate state are cached in module-level globals, mutated by `optimize_cuts_pulp`, `apply_width_optimization`. Downstream code in `core/visualization.py` imports and reads `OPT_CASCADING_PLAN`. Contract is implicit: "Call the optimizer; side effects land in these globals."

#### Impact

- **Testing:** Globals leak between test cases; requires manual teardown or pytest fixtures for isolation.
- **Concurrency:** Process-wide mutations; no safety under async, threading, or parallel requests.
- **Transparency:** Caller cannot predict API contract from signature alone.
- **Debugging:** State changes are invisible; hard to trace where a value came from.

#### Recommendation

1. Return all state explicitly from functions; do not mutate module globals.
2. Encapsulate result in `OptimizationResult` dataclass:
   ```python
   @dataclass
   class OptimizationResult:
       plan: dict
       cascading_plan: dict
       by_load: dict
       reinforcement_map: dict
       width_priority: dict
   ```
3. Inject result into visualization layer via dependency (context object or explicit argument), not global import.
4. Use `@functools.cache` or explicit session objects for legitimate caching (thread-safe, scoped).

**Effort:** High (1–2 days refactoring + test updates).  
**Urgency:** High (unblocks concurrency, testing, and API clarity).

---

### [S1] Shared Global Optimization State Under Concurrency

**Severity:** HIGH  
**Category:** Security  
**Scope:** Module-level globals; concurrent request handling

#### Description

The module maintains process-wide mutable state (`OPT_*` globals). Under concurrent requests (e.g., FastAPI with `workers > 1`, async tasks, threaded background jobs), results from one session/tenant may be visible to another if the globals are not properly isolated.

#### Impact

- **Data leak:** Session A's optimized plan visible in Session B's result.
- **Correctness:** Cross-session contamination of order/quote data.
- **Compliance:** Audit trail corrupted; cannot prove isolation between tenants or requests.

#### Recommendation

1. Eliminate module-level mutable globals (as per A4).
2. Use thread-safe session objects scoped to request/task lifecycle.
3. In FastAPI, use `Depends` to inject session context per request.
4. Document concurrency model explicitly.

**Effort:** Medium (aligned with A4 refactoring).  
**Urgency:** CRITICAL (if concurrency is enabled in production).

---

### [S2] Unbounded Input Processing and Resource Exhaustion

**Severity:** HIGH  
**Category:** Security  
**Scope:** `optimize_cuts_pulp`, `_optimize_2d_with_lengths`

#### Description

No validation on the size or cardinality of input `orders_2d` before ILP model construction. Large or malformed input can:

- Exhaust memory during variable/constraint generation.
- Cause solver to hang or timeout.
- TimeLimit is enforced only after the model is fully built (potentially minutes of CPU on failure case).

#### Impact

- **Denial of Service:** Attacker sends 1M+ orders; optimization process hangs indefinitely or crashes.
- **Resource leak:** No early validation gates.

#### Recommendation

1. Add input validation before model construction:
   ```python
   def _validate_orders_2d(orders_2d: list, max_orders: int = 1000):
       if len(orders_2d) > max_orders:
           raise ValueError(f"Too many orders: {len(orders_2d)} > {max_orders}")
   ```
2. Add cumulative demand threshold (total units).
3. Move timeLimit check to model setup, not solver invocation.

**Effort:** Low (validation layer; 1 day).  
**Urgency:** High (operational safety).

---

### [Q1] Attribution Mismatch: Slot Grouping vs. Canonical Keys

**Severity:** HIGH  
**Category:** Code Quality  
**Scope:** `_build_proportional_slot_lists`, `_get_next_order_info`

#### Description

`_build_proportional_slot_lists` groups slots by raw key (e.g., plate geometry tuple) but `_get_next_order_info` matches demand via `canonical_plate_key(...)`. If canonicalization logic changes or raw key collides, attribution fails silently.

#### Impact

- Orders misaligned to plates; wrong client/quote ID in output.
- Silent error; hard to detect in tests without explicit assertion on order IDs.

#### Recommendation

1. Use `canonical_plate_key` consistently in both functions.
2. Add unit test: `test_slot_grouping_matches_canonical_key()` that verifies all keys are identical.
3. Refactor to single KeyBuilder class/function.

**Effort:** Low (1 day).  
**Urgency:** High (correctness; audit trail).

---

### [Q2] Monolithic Procedure: `_optimize_2d_with_lengths` Complexity and Testability

**Severity:** HIGH  
**Category:** Code Quality  
**Scope:** `_optimize_2d_with_lengths` (entire function)

#### Description

Single 100+ line procedure handling:
- Variable declaration and naming
- Constraint assembly (demand, residual, load class, etc.)
- Objective function
- Solver invocation
- Result parsing and post-processing

No internal helper functions; complex control flow hard to trace; testing requires mocking PuLP entirely.

#### Impact

- Cognitive burden: changes risk unintended side effects in other constraints.
- Testing: must either mock PuLP (fragile) or run full solver (slow).
- Debugging: stack trace doesn't pinpoint which constraint or variable caused infeasibility.

#### Recommendation

Extract into testable subfunctions:

```python
def _build_variables(demand_2d, primary_options, secondary_options):
    """Return dict of PuLP variables."""
    x_prim, x_sec, z_prim, z_sec, unmet = ...
    return {x_prim, x_sec, z_prim, z_sec, unmet}

def _build_constraints(variables, demand_2d, options, cfg):
    """Return list of constraint expressions."""
    constraints = []
    constraints.extend(_demand_constraints(...))
    constraints.extend(_residual_constraints(...))
    ...
    return constraints

def _optimize_2d_with_lengths(...):
    variables = _build_variables(...)
    constraints = _build_constraints(...)
    problem = LpProblem(...)
    problem += variables, constraints  # simplified pseudo-code
    ...
```

**Effort:** Medium (1–2 days refactoring + unit tests).  
**Urgency:** High (maintainability).

---

## Medium Severity Issues

### [A5] Hard Dependency on Infrastructure: Pricing and Config Layer

**Severity:** MEDIUM  
**Category:** Architecture  
**Scope:** `_optimize_2d_with_lengths` (PuLP objective function), `core/price_db.py`

#### Description

Cut cost is sourced directly from `core/price_db.get_price(..., cfg.PRICE_DB_PATH)` inside the ILP objective function. No abstraction or port; the domain logic is tightly coupled to a specific file-based price store and config path.

#### Impact

- Testing requires mocking both `price_db` and config.
- Changing pricing source (e.g., switching to API-based pricing) requires modifying the solver core.
- Hard to inject test fixtures for unit tests.

#### Recommendation

1. Introduce `PricingProvider` interface:
   ```python
   class PricingProvider(ABC):
       def get_cut_cost(self, cut_type, dims) -> float: ...
   ```
2. Inject into `_optimize_2d_with_lengths` via `OptimizationConfig`.
3. Implement concrete `FileBasedPricingProvider(path)` and test `MockPricingProvider`.

**Effort:** Low (1 day).  
**Urgency:** Medium (testability; future extensibility).

---

### [A6] Underspecified Optimization Model Parameters

**Severity:** MEDIUM  
**Category:** Architecture  
**Scope:** `_optimize_2d_with_lengths`, `OptimizationConfig`

#### Description

`OptimizationConfig` only covers `unused_rest_penalty_coeff` and `secondary_reuse_bonus`. Many critical ILP tuning knobs are hardcoded inline:

- Penalty weights: `M_UNMET`, `M_SOLID`
- Objective bonus: `lpSum(x_prim.values()) * 5000` (minimize count)
- Solver timeout: `PULP_CBC_CMD(..., timeLimit=60, gapRel=0.005)`
- Narrowing geometry: `NARROWING_TABLE` (local dict)

Result: Optimization behavior is not reproducible or tunable without code edits. Experiments require recompiles.

#### Impact

- No A/B testing or model tuning without code changes.
- Parameter history lost; hard to trace why results changed across time.
- Config object is a lie; real config is scattered.

#### Recommendation

Extend `OptimizationConfig`:

```python
@dataclass
class OptimizationConfig:
    unused_rest_penalty_coeff: float
    secondary_reuse_bonus: float
    unmet_penalty: float = 10000  # M_UNMET
    solid_penalty: float = 5000   # M_SOLID
    piece_count_bonus: float = 5000  # minimize x_prim count
    solver_timeout_sec: int = 60
    solver_gap_rel: float = 0.005
    narrowing_table: dict = field(default_factory=lambda: {...})
```

Validate and expose in API; persist to config store.

**Effort:** Low (1 day).  
**Urgency:** Medium (operability; tuning).

---

### [A7] In-Place Mutation of Input Structures

**Severity:** MEDIUM  
**Category:** Architecture  
**Scope:** `_build_residual_balance_constraints` (line ~1500)

#### Description

Function normalizes and directly assigns `opt['load_code'] = ...` on elements of `primary_options` list in-place. Caller's data structure is modified as a side effect, not documented.

#### Impact

- **Surprise behavior:** Caller passes options; they are mutated mid-function without warning.
- **Debugging:** State changes unexpectedly; hard to trace who modified the dict.
- **Reusability:** Cannot call function twice on same input.

#### Recommendation

1. Accept `primary_options`, return normalized copy (or separate results).
2. Document any mutation in docstring with `# MUTATES` marker.
3. Prefer functional style: immutable transformations.

**Effort:** Low (1 day).  
**Urgency:** Medium (correctness; reusability).

---

### [S3] Durable Business-Sensitive Logging: `core/price_db.py`

**Severity:** MEDIUM  
**Category:** Security  
**Scope:** `core/price_db.get_price` (logging to `debug_logs/debug-db7a51.log`)

#### Description

Every pricing lookup is appended to an unbounded debug log file, including demand volumes and client context. File grows indefinitely; sensitive commercial data persists on disk.

#### Impact

- **Data leakage:** Historical demand/pricing data accessible to any process reading the log.
- **Compliance:** GDPR, HIPAA, etc. may require data purge; no retention policy.
- **Disk space:** Log file grows unbounded; can fill disk.

#### Recommendation

1. Implement structured logging with rotation and TTL:
   ```python
   import logging
   from logging.handlers import RotatingFileHandler
   
   handler = RotatingFileHandler(
       "debug_logs/debug-pricing.log",
       maxBytes=10_000_000,  # 10 MB
       backupCount=3
   )
   ```
2. Conditionally enable debug logging via `log_level` config, not unconditional append.
3. Redact client ID, demand volumes in logs (or use hashed/anonymized tokens).

**Effort:** Low (1 day).  
**Urgency:** Medium (compliance; operational safety).

---

### [S4] Unconditional Debug Log File Writes in Optimizer Core

**Severity:** MEDIUM  
**Category:** Security  
**Scope:** `core/optimization.py` (debug output in `_optimize_2d_with_lengths`, etc.)

#### Description

Debug logs write order demand, solver state, and solution metadata directly to filesystem without conditional logging. Business-critical data (demand aggregates, solution diagnostics) become durable artifacts.

#### Impact

- **Audit trail contamination:** Sensitive demand info in plaintext logs.
- **Performance:** Unbounded file I/O on production calls.
- **Operational visibility:** No control over what is logged; cannot disable for performance.

#### Recommendation

1. Implement conditional debug mode via `OptimizationConfig`:
   ```python
   @dataclass
   class OptimizationConfig:
       debug_enabled: bool = False
       debug_log_path: str | None = None
   ```
2. Only write logs if `debug_enabled and debug_log_path` is set.
3. Use structured logging (JSON) for future parsing/redaction.

**Effort:** Low (1 day).  
**Urgency:** Medium (compliance; observability).

---

### [S5] Uncontrolled stdout Prints in ILP Solver

**Severity:** MEDIUM  
**Category:** Security  
**Scope:** `_optimize_2d_with_lengths` (prints from PuLP/solver)

#### Description

stdout prints (solver progress, debug emoji/status messages) may be captured by production logging systems and become persistent operational logs, mixing business context with debug chatter.

#### Impact

- **Log pollution:** Difficulty filtering signal from noise in operational logs.
- **Data exposure:** Solver state visible in production logs.

#### Recommendation

1. Suppress PuLP solver output (pass `msg=0` to `PULP_CBC_CMD`).
2. Capture solver diagnostics into structured logger if debug enabled.
3. Sanitize all prints from core module; use logging only.

**Effort:** Low (1 day).  
**Urgency:** Medium (operability).

---

### [Q3] Code Duplication: `_get_next_order_info` Slot Handling

**Severity:** MEDIUM  
**Category:** Code Quality  
**Scope:** `_get_next_order_info` (repeated dict structure, try/log blocks)

#### Description

Multiple try/except blocks with identical logging pattern, repeated dict access patterns (`order['qty_remaining']`, `order['plate_id']`, etc.), and fallback logic duplicated across branches.

#### Impact

- **Maintenance burden:** Bug fixes must be applied to multiple locations.
- **Cognitive load:** Logic harder to follow.

#### Recommendation

Extract common patterns into helpers:

```python
def _get_order_remaining_qty(order: dict) -> int:
    return order.get('qty_remaining', 0)

def _log_order_match(order_id, plate_key, qty):
    logger.debug(f"Order {order_id} matched to {plate_key}: {qty} units")
```

**Effort:** Low (1 day).  
**Urgency:** Low (refactoring; maintainability).

---

### [Q4] Exception Swallowing: Widespread `except Exception: pass` Blocks

**Severity:** MEDIUM  
**Category:** Code Quality  
**Scope:** Multiple functions (`_get_next_order_info`, `_peek_order_info`, etc.)

#### Description

Multiple bare `except Exception: pass` statements silently ignore all errors, including programming bugs, resource exhaustion, and unexpected state.

#### Impact

- **Debugging:** Errors disappear; hard to diagnose failures.
- **Reliability:** Silent failures lead to incorrect results without warning.

#### Recommendation

1. Replace with specific exception handling:
   ```python
   try:
       ...
   except KeyError:
       logger.warning(f"Order key not found: {order_id}")
   except ValueError as e:
       logger.error(f"Invalid order data: {e}")
   ```
2. Use `logger.exception()` to capture stack traces.
3. Consider not catching if error should propagate.

**Effort:** Low (1 day).  
**Urgency:** Medium (reliability; debugging).

---

### [Q5] Inconsistent Type Annotations

**Severity:** MEDIUM  
**Category:** Code Quality  
**Scope:** Function signatures (`orders_2d: list`, raw dicts throughout)

#### Description

Type hints are sparse or overly generic. `orders_2d` typed as `list` instead of `list[dict]` or structured dataclass; `order` parameters are raw dicts, not typed.

#### Impact

- **IDE support:** Autocomplete fails; no type checking.
- **Clarity:** Contract unclear; what fields are required in `order`?
- **Refactoring:** Risky; no type checker to catch mismatches.

#### Recommendation

1. Define TypedDict or dataclass for Order, Plate, etc.:
   ```python
   class OrderInfo(TypedDict):
       kp_id: str
       qty_remaining: int
       plate_id: str
       client: str
       ...
   ```
2. Use throughout module.
3. Enable `mypy --strict` in CI.

**Effort:** Medium (1–2 days).  
**Urgency:** Medium (maintainability; IDE support).

---

### [Q6] Domain Thresholds Scattered vs. Config

**Severity:** MEDIUM  
**Category:** Code Quality  
**Scope:** `LEN_TOL`, `KERF_WIDTH_MM`, heuristic constants

#### Description

Tolerance and threshold constants (`LEN_TOL`, `KERF_WIDTH_MM`, etc.) are hardcoded as module-level constants rather than part of config. Tuning requires code edit.

#### Impact

- **Operability:** No way to adjust tolerances for different plants or materials without redeployment.
- **Testing:** Hard to test behavior at tolerance boundaries without parameterization.

#### Recommendation

1. Add to `OptimizationConfig`:
   ```python
   @dataclass
   class OptimizationConfig:
       length_tolerance_mm: float = 10  # LEN_TOL
       kerf_width_mm: float = 2.0
       ...
   ```
2. Pass config to all functions that use constants.
3. Load from config file or environment.

**Effort:** Low (1 day).  
**Urgency:** Low (operability; tuning).

---

### [Q7] Inconsistent Debug Logging Discipline

**Severity:** MEDIUM  
**Category:** Code Quality  
**Scope:** Scattered `_opt_debug_enabled` checks, hardcoded paths, mixed print/log patterns

#### Description

Some functions check `_opt_debug_enabled` before logging; others unconditionally write files; still others print to stdout. No consistent pattern.

#### Impact

- **Maintenance:** Inconsistent behavior; hard to disable debug globally.
- **Performance:** Debug code interspersed with production paths.

#### Recommendation

1. Centralize debug configuration:
   ```python
   @dataclass
   class OptimizationConfig:
       debug_mode: bool = False
       debug_log_path: str | None = None
   ```
2. Inject into all functions; use consistently.
3. Use `logging` module, not prints or custom flags.

**Effort:** Low (1 day).  
**Urgency:** Low (consistency; operability).

---

## Low Severity Issues

### [A8] FFD Track Packing Logic in Same Module as ILP

**Severity:** LOW  
**Category:** Architecture  
**Scope:** `Piece`, `Track`, `first_fit_decreasing` classes/functions

#### Description

First-fit-decreasing bin packing (track assembly) is implemented in the same file as ILP optimization, despite being a distinct algorithmic subdomain. No shared abstraction with cascading ILP; FFD is an alternative/parallel post-processing step.

#### Impact

- **Organization:** Cognitive load; two different algorithms in one file.
- **Reusability:** Hard to use FFD independently for other packing problems.
- **Testing:** Tests for FFD and ILP mixed together.

#### Recommendation

Extract to `core/optimization/ffd_packing.py`; expose as `pack_tracks(pieces: list[Piece]) -> list[Track]`. Import in orchestrator only when needed.

**Effort:** Low (1 day).  
**Urgency:** Low (organization; long-term maintainability).

---

### [A9] Module-Level Cleanup: Late Imports and Dead Constants

**Severity:** LOW  
**Category:** Architecture  
**Scope:** Late `from dataclasses import dataclass` near bottom; `KERF_WIDTH_MM`, unused constants

#### Description

Imports and constants scattered throughout file, some near bottom after large code blocks. Cognitive friction; reader must jump to understand dependencies.

#### Impact

- **Onboarding:** New contributors confused about what's available.
- **Cleanliness:** Code smell; suggests disorganized module growth.

#### Recommendation

1. Move all imports to top.
2. Group constants into `class Constants:` or config.
3. Remove unused constants (`KERF_WIDTH_MM` if not used, etc.).

**Effort:** Low (0.5 day).  
**Urgency:** Low (hygiene).

---

### [A10] Data Flow Implicit Through Global Config

**Severity:** LOW  
**Category:** Architecture  
**Scope:** Implicit dependency on global `cfg` object

#### Description

Module depends on global `cfg` (`core/config_and_data.py`) for plate counts, price DB path, and other state. Dependency is not explicit in function signatures; buried in function bodies.

#### Impact

- **Unclear contract:** Caller cannot see that function depends on global config without reading source.
- **Testing:** Global state makes unit tests fragile.
- **Concurrency:** Global config is shared state; unsafe under parallelism.

#### Recommendation

1. Make `cfg` an explicit parameter to top-level API functions.
2. Inject into sub-functions via dependency container or dataclass.
3. Use `@pytest.fixture` to mock `cfg` in tests.

**Effort:** Medium (1 day).  
**Urgency:** Low (testing; clarity; long-term maintainability).

---

### [S6] Untrusted Input: Order Structures Not Explicitly Validated

**Severity:** LOW  
**Category:** Security  
**Scope:** `orders_2d` input validation

#### Description

Input `orders_2d` structures are assumed to have required fields and valid types. No schema validation at entry point.

#### Impact

- **Robustness:** Malformed input causes cryptic AttributeError deep in solver.
- **Debugging:** Hard to trace back to caller.

#### Recommendation

1. Add input validation function:
   ```python
   def _validate_orders_2d(orders: list) -> None:
       for i, order in enumerate(orders):
           if not isinstance(order, dict):
               raise TypeError(f"Order {i} is not a dict: {type(order)}")
           if 'length' not in order or 'width' not in order:
               raise ValueError(f"Order {i} missing required fields: {order}")
           if not isinstance(order['length'], (int, float)):
               raise TypeError(f"Order {i} length not numeric: {order['length']}")
   ```
2. Call at entry point before processing.

**Effort:** Low (1 day).  
**Urgency:** Low (robustness).

---

### [S7] Access Control on Shared Price Database

**Severity:** LOW  
**Category:** Security  
**Scope:** `pb.db` ACLs, `PRICE_DB_PATH` assumption

#### Description

Price DB file path (`pb.db`) is assumed to exist and be readable. On shared hosts, file-level ACLs may not be enforced; any process can read/modify pricing data.

#### Impact

- **Confidentiality:** Pricing visible to other services/users on host.
- **Integrity:** Pricing can be corrupted by parallel processes.

#### Recommendation

1. Check file permissions at startup:
   ```python
   import os
   db_path = cfg.PRICE_DB_PATH
   if not os.path.exists(db_path):
       raise FileNotFoundError(f"Price DB not found: {db_path}")
   # On Unix: check mode; on Windows, use ACLs
   if os.stat(db_path).st_mode & 0o077:  # World-readable/writable
       logger.warning(f"Price DB has insecure permissions: {db_path}")
   ```
2. Use file locking for concurrent access (if applicable).
3. Document assumed ACLs in deployment guide.

**Effort:** Low (1 day).  
**Urgency:** Low (deployment; host-specific).

---

### [S8] Resource Limits: Unbounded Model Size Under Large Demand

**Severity:** LOW  
**Category:** Security  
**Scope:** `_optimize_2d_with_lengths` (model construction)

#### Description

Very large demands can cause PuLP model to consume excessive memory or solver to struggle (as per S2, but lower priority after input validation is added).

#### Impact

- **Availability:** Solver may timeout or OOM on edge cases.
- **Graceful degradation:** No fallback for "model too complex."

#### Recommendation

1. After input validation (S2), add model complexity estimation:
   ```python
   def _estimate_model_complexity(orders_2d, options):
       n_vars = len(orders_2d) * len(options)  # Rough estimate
       if n_vars > MAX_VARS:
           raise ValueError(f"Model too complex: {n_vars} variables")
   ```
2. Implement graceful fallback (greedy heuristic) if model exceeds limit.

**Effort:** Low (1 day).  
**Urgency:** Low (edge case; operational resilience).

---

### [Q8] Test Coverage Gaps: Width Optimization and FFD Tracks

**Severity:** LOW  
**Category:** Code Quality  
**Scope:** `tests/` directory (missing test files)

#### Description

Test suite covers baseline cascading cuts but lacks dedicated coverage for:
- `apply_width_optimization` / `_optimize_1d_widths_only` (legacy width logic)
- `optimize_cuts_pulp` and `OPT_PLAN` mutation side effects
- `first_fit_decreasing` and track assembly (FFD)

#### Impact

- **Regression risk:** Changes to width or FFD logic have no safety net.
- **Confidence:** Refactoring these functions is risky.

#### Recommendation

Add test files:
1. `tests/test_width_optimization.py` – Test `apply_width_optimization` end-to-end.
2. `tests/test_ffd_packing.py` – Unit tests for `first_fit_decreasing`, `Piece`, `Track`.
3. `tests/test_integration_optimize_cuts_pulp.py` – End-to-end for legacy endpoint.

**Effort:** Medium (1–2 days writing tests).  
**Urgency:** Low (coverage; risk mitigation).

---

### [Q9] Memory and Time Complexity: FFD Expansion of Large Orders

**Severity:** LOW  
**Category:** Code Quality  
**Scope:** `first_fit_decreasing` (loop over quantities)

#### Description

In FFD, each unit of quantity is expanded into a separate `Piece` object. For orders with qty=1000, this creates 1000 Piece instances; memory and time complexity is O(qty).

#### Impact

- **Performance:** Large orders slow down FFD.
- **Memory:** Linear memory usage in order size.
- **Scalability:** Does not scale to high-volume scenarios.

#### Recommendation

1. Optimize FFD to treat quantities as weights, not expand:
   ```python
   class Piece:
       def __init__(self, length, width, qty=1, ...):
           self.length = length
           self.width = width
           self.qty = qty  # Don't expand; keep aggregated
   ```
2. Adjust packing algorithm to handle quantities.
3. Benchmark impact.

**Effort:** Medium (1–2 days algorithm tuning).  
**Urgency:** Low (scalability; performance under high load).

---

### [Q10] Module Organization: Late Imports and Scattered Utilities

**Severity:** LOW  
**Category:** Code Quality  
**Scope:** Same as A9 (late imports, scattered constants)

**Already addressed in A9.** Reference A9 recommendation.

---

### [Q11] Dead Code: Unused Constants

**Severity:** LOW  
**Category:** Code Quality  
**Scope:** Constants like `KERF_WIDTH_MM` (if truly unused)

#### Description

Module may define constants that are never referenced in the current version of code, indicating legacy cleanup needed.

#### Impact

- **Cognitive load:** Reader wonders if constant is used.
- **Technical debt:** Dead code accumulates.

#### Recommendation

1. Run static analysis to identify unused constants:
   ```bash
   python -m vulture core/optimization.py
   ```
2. Remove unused definitions.
3. Document retained constants with comments if non-obvious.

**Effort:** Low (0.5 day).  
**Urgency:** Low (hygiene).

---

### [Q12] Mixed Output UX: Emoji, Russian/English Mix in Library Core

**Severity:** LOW  
**Category:** Code Quality  
**Scope:** stdout prints, log messages in core module

#### Description

Debug output contains mix of emoji (🎯, ✅), Russian text, and English; inconsistent tone for library code. Suitable for CLI, not for library core.

#### Impact

- **Professionalism:** Library core should be locale-neutral and structured.
- **Parsing:** Emoji/RU text hard to parse for monitoring systems.
- **Maintenance:** Future contributors unsure of output style.

#### Recommendation

1. Remove emoji and RU text from core module output.
2. Use structured logging (JSON) for machine parsing.
3. Keep UX flavor for CLI wrappers only (if any).

**Effort:** Low (0.5 day).  
**Urgency:** Low (professionalism; operability).

---

## Priority Matrix

| Priority | ID | Title | Category | Effort | Impact | Recommended Action |
|----------|----|----|----------|--------|--------|-------------------|
| 🔴 CRITICAL | A1 | God Module Decomposition | Architecture | High | Very High | `/refactor` – Immediate multi-day sprint |
| 🔴 CRITICAL | S1 | Concurrency: Shared Mutable State | Security | High | Very High | Fix before enabling multiprocessing |
| 🔴 CRITICAL | S2 | Input Validation & Resource Exhaustion | Security | Low | High | `/implement` – Priority security hardening |
| 🟠 HIGH | A2 | Order Identity in ILP Model | Architecture | Medium | High | `/refactor` – Design phase 1 |
| 🟠 HIGH | A3 | Dual Data Sources | Architecture | Medium | Medium | `/refactor` – Design phase 1 |
| 🟠 HIGH | A4 | Global Caches / Side Effects | Architecture | High | High | `/refactor` – Parallel with A1 |
| 🟠 HIGH | S3 | Durable Business Logging | Security | Low | Medium | `/implement` – Logging hardening |
| 🟠 HIGH | S4 | Unconditional Debug Writes | Security | Low | Medium | `/implement` – Logging hardening |
| 🟠 HIGH | Q1 | Slot Attribution Mismatch | Code Quality | Low | High | `/refactor` – Quick fix |
| 🟠 HIGH | Q2 | Monolithic Solver Procedure | Code Quality | Medium | High | `/refactor` – Design phase 2 |
| 🟡 MEDIUM | A5 | Infrastructure Dependency | Architecture | Low | Medium | `/implement` – Abstraction layer |
| 🟡 MEDIUM | A6 | Underspecified Config | Architecture | Low | Medium | `/implement` – Config expansion |
| 🟡 MEDIUM | A7 | Input Mutation Side Effects | Architecture | Low | Medium | `/refactor` – Cleanup |
| 🟡 MEDIUM | S5 | Uncontrolled stdout | Security | Low | Low | `/implement` – Output sanitization |
| 🟡 MEDIUM | Q3 | Code Duplication | Code Quality | Low | Low | `/refactor` – Cleanup |
| 🟡 MEDIUM | Q4 | Exception Swallowing | Code Quality | Low | Medium | `/refactor` – Error handling |
| 🟡 MEDIUM | Q5 | Type Annotations | Code Quality | Medium | Medium | `/implement` – Type safety |
| 🟡 MEDIUM | Q6 | Scattered Constants | Code Quality | Low | Low | `/refactor` – Config migration |
| 🟡 MEDIUM | Q7 | Logging Discipline | Code Quality | Low | Low | `/implement` – Logging strategy |
| 🟢 LOW | A8 | FFD in Same Module | Architecture | Low | Low | `/refactor` – Extract module |
| 🟢 LOW | A9 | Module Cleanup | Architecture | Low | Low | `/refactor` – Code hygiene |
| 🟢 LOW | A10 | Implicit Config Dependency | Architecture | Medium | Low | Future refactor |
| 🟢 LOW | S6 | Input Validation | Security | Low | Low | `/implement` – Validation layer |
| 🟢 LOW | S7 | DB Access Control | Security | Low | Low | Deployment checklist |
| 🟢 LOW | S8 | Resource Limits | Security | Low | Low | `/implement` – Edge case handling |
| 🟢 LOW | Q8 | Test Coverage Gaps | Code Quality | Medium | Medium | `/implement` – Test writing |
| 🟢 LOW | Q9 | FFD Memory Complexity | Code Quality | Medium | Low | Future optimization |
| 🟢 LOW | Q10 | Late Imports | Code Quality | Low | Low | `/refactor` – Cleanup |
| 🟢 LOW | Q11 | Dead Code | Code Quality | Low | Low | `/refactor` – Cleanup |
| 🟢 LOW | Q12 | Output UX | Code Quality | Low | Low | `/implement` – Output sanitization |

---

## Next Steps

### Immediate Actions (This Sprint)

1. **Run `/refactor` for Critical Architecture (A1)**
   - Decompose `core/optimization.py` into 5 modules: `ilp_model.py`, `order_dispatch.py`, `geometry.py`, `ffd_packing.py`, `orchestrator.py`.
   - Expected duration: 2–3 days.
   - Blocks: All other refactorings depend on this.

2. **Run `/implement` for Critical Security (S1, S2)**
   - Add input validation (S2, S6): Bounds checking, cardinality limits.
   - Eliminate global mutable state (S1, S4): Return results explicitly; inject session context.
   - Expected duration: 1–2 days.
   - Blocks: Before enabling production concurrency.

3. **Run `/implement` for High-Severity Code Quality (Q1, Q2)**
   - Fix slot attribution logic (Q1): Unit test for canonical key consistency.
   - Extract helper functions from `_optimize_2d_with_lengths` (Q2): Break into `_build_variables`, `_build_constraints`, `_build_objective`.
   - Expected duration: 1 day.

### Secondary Actions (Following Sprint)

4. **Run `/implement` for Security Hardening (S3, S4, S5)**
   - Implement structured logging with rotation (S3, S4).
   - Suppress solver stdout (S5).
   - Expected duration: 1 day.

5. **Run `/implement` for Type Safety (Q5)**
   - Define `OrderInfo`, `PlateOption`, `CutConfig` TypedDicts.
   - Enable `mypy --strict` in CI.
   - Expected duration: 1 day.

6. **Run `/implement` for Test Coverage (Q8)**
   - Write `tests/test_width_optimization.py`, `tests/test_ffd_packing.py`, `tests/test_integrate_cuts_pulp.py`.
   - Aim for >80% coverage of width and FFD logic.
   - Expected duration: 1–2 days.

### Operational Improvements

7. **Update Configuration (A5, A6)**
   - Introduce `PricingProvider` interface for testability.
   - Expand `OptimizationConfig` to include all tuning parameters: `unmet_penalty`, `solver_timeout_sec`, etc.
   - Expected duration: 1 day.

8. **Documentation and Cleanup (A8, A9, A10)**
   - Move FFD to separate module.
   - Group imports and constants at top.
   - Document global config assumptions in module docstring.
   - Expected duration: 1 day (after A1 refactoring).

### Long-Term Roadmap

- **Performance optimization (Q9):** Optimize FFD to avoid expanding large quantities.
- **Comprehensive test suite (Q8):** Ensure >85% coverage across all modules post-refactoring.
- **Monitoring and observability:** Implement structured logging, metrics export (latency, model complexity, solver gap).
- **Design review:** Once A1 is complete, conduct senior review of new modular architecture to lock in patterns.

---

## Audit Metadata

- **Report Generated:** 2026-05-05 11:45 UTC+3
- **Reviewed Files:** `core/optimization.py` (~2.7k LOC), `core/config_and_data.py` (supporting context), `core/price_db.py` (referenced in S3)
- **Audit Phases:** Senior Reviewer (architecture), Security Auditor (security), Code Quality Reviewer (quality)
- **Total Findings:** 30 (1 Critical, 7 High, 11 Medium, 11 Low)
- **Overall Assessment:** **Requires significant refactoring and security hardening before production use at scale.** Critical architecture flaw (A1) blocks safe concurrency and maintainability. Immediate focus on decomposition and state isolation.

---

**Report Status:** ✅ Complete  
**Distribution:** Development team, Architecture review, Security review  
**Follow-up:** Schedule refactoring sprint planning within 1 week.
