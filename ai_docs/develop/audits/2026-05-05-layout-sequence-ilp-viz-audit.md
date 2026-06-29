# Project Audit Report
## Layout Sequence + ILP Visualization Integration

**Date:** 2026-05-05  
**Scope:** `viz_modules/layout_sequence.py`, `core/config_and_data.py`, `core/optimization.py`, `core/visualization.py`  
**Health Score:** 2.0/10  
**Status:** Critical — Production readiness blocked by critical issues

---

## Executive Summary

### Audit Findings by Severity

| Severity | Architecture | Security | Code Quality | Total |
|----------|-------------|----------|--------------|-------|
| **Critical** | 1 | 1 | 0 | **2** |
| **High** | 3 | 2 | 2 | **7** |
| **Medium** | 3 | 1 | 6 | **10** |
| **Low** | 2 | 2 | 5 | **9** |
| **TOTAL** | **9** | **6** | **13** | **28** |

### Key Findings Summary

- **Architecture**: Globals-based pipeline (`build_layout_sequence` → `split_sequence_into_tracks`) lacks dependency injection; fragile test isolation and worker concurrency issues.
- **Security**: Multi-tenant isolation broken via mutable module globals; path traversal risks on `output_dir`; debug file disclosure.
- **Code Quality**: Duplicated layout pipelines, cyclomatic complexity debt, swallowed exceptions, type safety gaps.

### Recommendation

**Do NOT deploy to concurrent production without addressing the 2 critical issues** (A1, S1). These break isolation and enable cross-request data corruption.

---

## Critical Issues (Action Required)

### [A1] Request-Scoped Data as Process Globals

**Severity:** 🔴 Critical  
**Category:** Architecture  
**Location:** `core/config_and_data.py`, `viz_modules/layout_sequence.py`  

**Impact:**  
Global variables (`OPT_CASCADING_PLAN`, `OPT_PLAN`, `OPT_ORDERS`, etc.) store per-request state without isolation. In multi-worker or concurrent test scenarios, one request's data corrupts another's. Thread-unsafe; no greenlet-safe semantics. Violates request scoping contract.

**Current Code Pattern:**
```python
# Module globals (dangerous in concurrent context)
OPT_CASCADING_PLAN = None
OPT_PLAN = None
OPT_ORDERS = None

def build_layout_sequence():
    global OPT_PLAN
    # implicit: caller must pre-set OPT_PLAN, OPT_CASCADING_PLAN, OPT_ORDERS
```

**Fix:**
1. Inject pipeline state via function parameters or context manager
2. Use thread-local or request-scoped dependency injection (FastAPI `Depends`, context vars, or domain context objects)
3. Eliminate module-level mutations
4. Add test coverage for concurrent access

**Effort:** High  
**Priority:** P0

---

### [S1] Multi-Tenant / Worker Isolation Broken via Mutable Module Globals

**Severity:** 🔴 Critical  
**Category:** Security  
**Location:** `core/config_and_data.py`, `viz_modules/layout_sequence.py`  

**Impact:**  
If the system serves multiple tenants, customers, or parallel requests in the same process, mutable globals enable direct data leakage: one tenant's layout plan, orders, or optimization results become visible to another. Exfiltration vector in shared infrastructure.

**Scenario:**
```python
# Worker A sets global state
OPT_PLAN = {"customer_id": 123, "sensitive_layout": ...}

# Context switch: Worker B starts execution
# Worker B sees Worker A's data via globals
```

**Fix:**  
Same as A1 — implement request-scoped dependency injection. Audit existing request handlers for global state mutations.

**Effort:** High  
**Priority:** P0 (Security)

---

## High-Severity Issues

### [A2] Unreachable Duplicate Block After Early Return

**Severity:** 🟠 High  
**Category:** Architecture  
**Location:** `viz_modules/layout_sequence.py`, `build_layout_sequence()`  

**Impact:**  
Dead code path after early return when plan is present. Increases cognitive load, suggests incomplete refactoring, and masks intent.

**Current Pattern:**
```python
def build_layout_sequence():
    if OPT_PLAN:  # plan already set
        return OPT_PLAN  # early return
    
    # ... dead code below — never reached if plan was pre-populated
```

**Fix:**  
Remove unreachable block or clarify the conditional flow.

**Effort:** Low  
**Priority:** P1

---

### [A3] Dummy Optimization Corrupts OPT_PLAN State

**Severity:** 🟠 High  
**Category:** Architecture  
**Location:** `core/visualization.py`, `visualize_plan()`  

**Impact:**  
`visualize_plan()` calls `optimize_cuts_pulp()` with dummy order (`order_key="dummy"`), which clears and repopulates `OPT_PLAN` as side effect. Subsequent code sees corrupted state. Violates referential transparency; hides implicit state mutation.

**Current Pattern:**
```python
def visualize_plan():
    optimize_cuts_pulp("dummy")  # side effect: mutates OPT_PLAN globally
    # OPT_PLAN now contains toy optimization, not real data
```

**Fix:**  
1. Make optimization explicit via return value, not side effect
2. Add validation to detect state corruption
3. Consider separate `VisualizationContext` that is not mutated by optimization

**Effort:** Medium  
**Priority:** P1

---

### [A4] God Module: Mixed Concerns in layout_sequence.py

**Severity:** 🟠 High  
**Category:** Architecture  
**Location:** `viz_modules/layout_sequence.py`  

**Impact:**  
Single module handles: DB paths, domain rules, graph logic, debug I/O. Difficult to test, maintain, and extend. Violates single responsibility principle.

**Modules/Concerns:**
- Database initialization (DB path, connection)
- Domain business rules (track grouping, load balancing)
- Graph algorithms (sequence layout)
- Debug file I/O (log export)

**Fix:**  
1. Extract DB logic → `core/db_init.py`
2. Extract domain rules → `domain/layout_rules.py`
3. Extract graph algorithms → `algorithms/sequence_layout.py`
4. Keep `layout_sequence.py` as orchestrator

**Effort:** High  
**Priority:** P1

---

### [A5] Implicit Contract: Nullary build_layout_sequence()

**Severity:** 🟠 High  
**Category:** Architecture  
**Location:** `viz_modules/layout_sequence.py`, function signature  

**Impact:**  
`build_layout_sequence()` takes no parameters but depends on globals. Callers must understand and pre-set `OPT_CASCADING_PLAN`, `OPT_PLAN`, `OPT_ORDERS`. Implicit contract is error-prone and undocumented.

**Fix:**  
1. Add explicit parameters: `def build_layout_sequence(plan, orders, cascading_plan=None)`
2. Or use context object: `def build_layout_sequence(context: LayoutContext)`
3. Document preconditions if globals remain

**Effort:** Medium  
**Priority:** P1

---

### [S2] visualize_plan() Corrupts OPT_PLAN with Toy Optimization

**Severity:** 🟠 High  
**Category:** Security  
**Location:** `core/visualization.py`  

**Impact:**  
By replacing real optimization data with toy results, visualization may produce misleading output. If layout decisions are based on corrupted plan, yields incorrect customer deliverables or cost calculations.

**Fix:**  
Separate visualization context from optimization state. Add unit test that asserts `OPT_PLAN` unchanged after visualization.

**Effort:** Medium  
**Priority:** P1

---

### [S3] Path Traversal Risk on output_dir

**Severity:** 🟠 High  
**Category:** Security  
**Location:** `core/visualization.py`, `core/config_and_data.py`  

**Impact:**  
If `output_dir` is user-controlled, attacker can inject `../../../etc/passwd` or other paths. No validation on path components.

**Current Pattern:**
```python
output_dir = user_input  # e.g., "../../sensitive/"
filepath = os.path.join(output_dir, filename)  # no validation
```

**Fix:**  
1. Validate `output_dir` is within a base directory
2. Use `os.path.normpath()` and check it starts with base
3. Add path traversal test case
4. Consider allowlist of permitted paths

**Effort:** Low  
**Priority:** P1

---

### [Q1] Duplicated Layout Pipeline Implementations

**Severity:** 🟠 High  
**Category:** Code Quality  
**Location:** `viz_modules/layout_sequence.py`  

**Impact:**  
Parallel layout pipeline logic (grouped vs. flat layouts) duplicated across functions. Changes to one path may not propagate to the other. Maintenance burden.

**Affected Functions:**  
- `build_layout_sequence()` — main path
- `split_sequence_into_tracks()` — variants for grouped/flat

**Fix:**  
1. Extract common pipeline steps → shared function
2. Parameterize branching logic (e.g., `strategy` parameter)
3. Add integration tests covering both paths

**Effort:** Medium  
**Priority:** P2

---

### [Q2] Monolithic Functions Exceed Complexity Thresholds

**Severity:** 🟠 High  
**Category:** Code Quality  
**Location:** `viz_modules/layout_sequence.py`, `core/visualization.py`  

**Impact:**  
Long functions with nested conditionals, loops, and side effects. Difficult to test, reason about, and refactor. High bug surface.

**Examples:**  
- `build_layout_sequence()` — 200+ lines
- `split_sequence_into_tracks()` — cyclomatic complexity > 8

**Fix:**  
1. Break into smaller functions with single responsibility
2. Extract conditional branches into separate handlers
3. Add unit tests for each extracted function
4. Target max complexity ≤ 5

**Effort:** High  
**Priority:** P2

---

## Medium-Severity Issues

### [A6] Inconsistent strict_layout_integrity Default

**Severity:** 🟡 Medium  
**Category:** Architecture  
**Location:** `viz_modules/layout_sequence.py`, `split_sequence_into_tracks()`  

**Impact:**  
`strict_layout_integrity=False` default; but bot workflows may rely on strict semantics. Inconsistency between viz flows (non-strict) and bot flows (strict expected).

**Fix:**  
1. Document expected strictness per flow
2. Align defaults or make explicit in callers
3. Add test asserting strictness applies where expected

**Effort:** Low  
**Priority:** P2

---

### [A7] Duplicated Domain Rules and Fallback Tying Globals to Viz

**Severity:** 🟡 Medium  
**Category:** Architecture  
**Location:** `core/config_and_data.py`, `viz_modules/layout_sequence.py`  

**Impact:**  
Load fallback logic, config defaults, and domain rules duplicated across modules. Tightly couples config/data setup to visualization module. Changes to fallbacks must be synced manually.

**Fix:**  
1. Centralize domain rules → `domain/layout_rules.py`
2. Move fallback logic → `core/config_defaults.py`
3. Import cleanly into viz and other consumers

**Effort:** Medium  
**Priority:** P2

---

### [A8] print vs. logging Inconsistency

**Severity:** 🟡 Medium  
**Category:** Architecture  
**Location:** `viz_modules/layout_sequence.py`, `core/optimization.py`  

**Impact:**  
Mix of `print()` statements and proper logging. Difficult to control verbosity, capture logs in tests, or integrate with centralized logging.

**Fix:**  
1. Replace all `print()` with `logging.info()`, `logging.debug()`, etc.
2. Configure logging centrally in `core/logging.py`
3. Add log level to config
4. Add test mocking logging to verify messages

**Effort:** Low  
**Priority:** P2

---

### [S4] Debug Logs May Leak Commercial/Layout Data

**Severity:** 🟡 Medium  
**Category:** Security  
**Location:** `core/visualization.py`, `get_debug_log_path()`  

**Impact:**  
Debug logs may contain layout plans, pricing, or customer data. If logs are shipped to centralized storage or exposed via API, data leakage risk.

**Fix:**  
1. Audit what is logged at each level (debug/info/warn/error)
2. Mask sensitive fields (prices, customer IDs) in logs
3. Add environment control: `LOG_LEVEL=WARNING` in production
4. Document what is safe to log

**Effort:** Medium  
**Priority:** P2

---

### [S5] Hardcoded pb.db Path vs. Centralized Settings

**Severity:** 🟡 Medium  
**Category:** Security  
**Location:** `core/config_and_data.py`, `core/optimization.py`  

**Impact:**  
Hardcoded `pb.db` path scattered across code. Not configurable per environment (dev, staging, prod). No audit trail of DB access.

**Fix:**  
1. Move DB path to centralized config
2. Use environment variables: `DB_PATH=./db/pb.db`
3. Load via `BaseSettings` in FastAPI
4. Add migration: `Alembic` with versioning

**Effort:** Low  
**Priority:** P2

---

### [Q3] Duplicated split_sequence_into_tracks Branches

**Severity:** 🟡 Medium  
**Category:** Code Quality  
**Location:** `viz_modules/layout_sequence.py`, `split_sequence_into_tracks()`  

**Impact:**  
Separate branches for grouped vs. flat layout handling. Logic paths nearly identical except for grouping predicates. Divergence creates maintenance debt.

**Fix:**  
1. Extract grouping strategy pattern
2. Use strategy objects: `class GroupingStrategy`
3. Parameterize branching
4. Add test for each strategy

**Effort:** Medium  
**Priority:** P2

---

### [Q4] Heuristic-Based Grouped Detection Brittleness

**Severity:** 🟡 Medium  
**Category:** Code Quality  
**Location:** `viz_modules/layout_sequence.py`  

**Impact:**  
Grouped layout detected by heuristic on first element type. If first element anomalous, detection fails silently. No explicit schema or validation.

**Current Pattern:**
```python
if isinstance(sequence[0], GroupType):  # brittle heuristic
    # assume all are grouped
```

**Fix:**  
1. Add explicit schema validation at entry
2. Add test with mixed/anomalous input
3. Document heuristic or replace with explicit marker

**Effort:** Medium  
**Priority:** P2

---

### [Q5] Broad except/pass on Agent Log Regions

**Severity:** 🟡 Medium  
**Category:** Code Quality  
**Location:** `core/optimization.py`, agent logging sections  

**Impact:**  
Bare `except: pass` swallows exceptions silently. Hides bugs, makes debugging difficult. Silent failures accumulate into data corruption.

**Current Pattern:**
```python
try:
    log_agent_action(...)
except:
    pass  # silent — bug hidden
```

**Fix:**  
1. Replace bare except with specific exception types
2. Add conditional logging or re-raise
3. Add test ensuring exception paths are covered
4. Use `logging.exception()` for visibility

**Effort:** Low  
**Priority:** P2

---

### [Q6] Optional Types Misuse (list = None)

**Severity:** 🟡 Medium  
**Category:** Code Quality  
**Location:** `viz_modules/layout_sequence.py`  

**Impact:**  
Variables intended as lists assigned `None`. Missing type annotations or None checks. Runtime errors when code assumes list methods exist.

**Current Pattern:**
```python
tracks_per_file = None  # intended to be list, not None
# Later: tracks_per_file.append(...)  # TypeError if None
```

**Fix:**  
1. Add explicit type hints: `tracks_per_file: list[Track] = []`
2. Initialize to empty list instead of None
3. Run mypy to catch type errors
4. Add test with edge cases

**Effort:** Low  
**Priority:** P2

---

### [Q7] Variable Name Reuse Confusion (key)

**Severity:** 🟡 Medium  
**Category:** Code Quality  
**Location:** `viz_modules/layout_sequence.py`  

**Impact:**  
Variable `key` reused across loop scopes with different meanings (dictionary key, loop index, object ID). Reduces readability, risk of accidental misuse.

**Fix:**  
1. Rename for clarity: `order_key`, `track_index`, `item_id`
2. Add type hints to distinguish semantics
3. Add linter rule forbidding single-letter variable names

**Effort:** Low  
**Priority:** P2

---

### [Q8] Default load_code 800 vs 8 Inconsistency

**Severity:** 🟡 Medium  
**Category:** Code Quality  
**Location:** `core/config_and_data.py`  

**Impact:**  
Hardcoded default `load_code=800` vs. system-wide `default=8`. Inconsistency suggests incomplete refactoring. Users confused which default applies.

**Fix:**  
1. Consolidate defaults in single config file
2. Document override rules (code > env > config > hardcoded)
3. Add test asserting consistency

**Effort:** Low  
**Priority:** P2

---

## Low-Severity Issues

### [A9] visualization Module Re-exports layout_sequence, Blurs Boundaries

**Severity:** 🔵 Low  
**Category:** Architecture  
**Location:** `core/visualization.py`  

**Impact:**  
Re-exporting layout_sequence in visualization module masks dependency graph. Consumers unclear whether they import from viz or layout.

**Fix:**  
1. Remove re-export; let consumers import directly from layout_sequence
2. Document module responsibilities clearly
3. Add lint rule forbidding barrel exports

**Effort:** Low  
**Priority:** P3

---

### [S6] get_debug_log_path() Filename Not Validated

**Severity:** 🔵 Low  
**Category:** Security  
**Location:** `core/visualization.py`, `get_debug_log_path()`  

**Impact:**  
Future risk: if filename is user-controlled without validation, path traversal possible. Currently hardcoded, but defensive coding advised.

**Fix:**  
1. Add filename validation (no `../`, etc.)
2. Whitelist allowed characters
3. Test with malicious input

**Effort:** Low  
**Priority:** P3

---

### [Q9] print + dynamic __import__ in layout_sequence

**Severity:** 🔵 Low  
**Category:** Code Quality  
**Location:** `viz_modules/layout_sequence.py`  

**Impact:**  
Use of `__import__()` for dynamic imports is anti-pattern. Should use importlib. Also mixes print statements (already flagged in Q8).

**Fix:**  
1. Replace `__import__()` with `importlib.import_module()`
2. Add static type hints to make imports explicit

**Effort:** Low  
**Priority:** P3

---

### [Q10] Stale Tuple Comment in PLATE_LOAD_DETAILS Config

**Severity:** 🔵 Low  
**Category:** Code Quality  
**Location:** `core/config_and_data.py`  

**Impact:**  
Comment describing tuple schema doesn't match current code. Misleads maintainers.

**Fix:**  
1. Update comment to match current structure
2. Or convert to dataclass with explicit fields

**Effort:** Low  
**Priority:** P3

---

### [Q11] Comment Drift: MAX_TRACK_LENGTH in visualization

**Severity:** 🔵 Low  
**Category:** Code Quality  
**Location:** `core/visualization.py`  

**Impact:**  
Comment describing `MAX_TRACK_LENGTH` no longer accurate. Creates confusion during maintenance.

**Fix:**  
1. Update comment to reflect current behavior
2. Add test asserting constraint

**Effort:** Low  
**Priority:** P3

---

### [Q12] Unused next_group in Separator API Noise

**Severity:** 🔵 Low  
**Category:** Code Quality  
**Location:** `viz_modules/layout_sequence.py`  

**Impact:**  
Parameter `next_group` unused in separator function. Adds API noise, unclear intent.

**Fix:**  
1. Remove unused parameter
2. Update all call sites
3. Add linter rule to catch unused parameters

**Effort:** Low  
**Priority:** P3

---

### [Q13] Test Coverage Gap vs. Branching Surface

**Severity:** 🔵 Low  
**Category:** Code Quality  
**Location:** `tests/` (layout_sequence, visualization)  

**Impact:**  
High branching complexity (Q2) with low test coverage. Edge cases untested; regressions likely.

**Fix:**  
1. Add unit tests for each branch of `split_sequence_into_tracks()`
2. Add edge case tests (empty, single, grouped, flat, mixed)
3. Aim for 85%+ line coverage
4. Add parameterized tests for variant combinations

**Effort:** Medium  
**Priority:** P2

---

## Priority Matrix

| Priority | Critical | High | Medium | Low | Total |
|----------|----------|------|--------|-----|-------|
| **P0** (Blocker) | 2 | — | — | — | **2** |
| **P1** (Major) | — | 5 | — | — | **5** |
| **P2** (Important) | — | — | 7 | — | **7** |
| **P3** (Nice-to-have) | — | — | — | 5 | **5** |
| **TOTAL** | **2** | **5** | **7** | **5** | **19** |

---

## Next Steps

### Immediate Actions (This Sprint)

1. **A1 + S1: Refactor globals to request-scoped DI**
   - Extract `LayoutContext` dataclass with plan, orders, cascading_plan
   - Update all callers to pass context explicitly
   - Add test for concurrent access (simulated with threads or asyncio)
   - Estimate: 2–3 days

2. **S3: Add path validation on output_dir**
   - Implement allowlist check
   - Add test with path traversal payloads
   - Estimate: 2 hours

3. **Q8 & A8: Consolidate print → logging, defaults**
   - Replace all print with logging
   - Centralize config defaults
   - Estimate: 1 day

### Follow-Up (Next Sprint)

4. **A4: Split god module layout_sequence.py**
   - Extract DB, domain rules, algorithms
   - Refactor orchestrator
   - Estimate: 3–4 days

5. **Q1–Q3, Q13: Dedup layout pipelines + test coverage**
   - Parameterize branching logic
   - Add 15–20 new unit tests
   - Estimate: 2–3 days

6. **A3–A5: Explicit function signatures and optional cleanups**
   - Add parameters to `build_layout_sequence()`, etc.
   - Clean up dead code (A2)
   - Estimate: 1 day

### Long-Term Improvements

- Migrate to async/await architecture to handle concurrent requests safely
- Implement rate limiting and request queue to prevent resource exhaustion
- Add comprehensive observability (structured logging, tracing, metrics)
- Expand test suite to 90%+ coverage with mutation testing

---

## Conclusion

**Current Status:** ⚠️ Not production-ready for concurrent workloads.

**Health Score:** 2.0/10 — Critical architectural and security issues must be resolved before deploying to multi-tenant or high-concurrency environments.

**Recommendation:** Address P0 issues (global state refactor) immediately. Allocate 2–3 week sprint for stability and security hardening.

---

**Audit Completed:** 2026-05-05  
**Next Audit Recommended:** After P0/P1 fixes; target: 2026-05-26
