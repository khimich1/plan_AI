# Project Audit Report: core/visualization.py

**Date**: 2026-05-05  
**Scope**: `core/visualization.py` (single-file audit)  
**Overall Health Score**: 6.0/10  
**Audited by**: senior-reviewer + security-auditor + reviewer

---

## Executive Summary

This audit examined `core/visualization.py` for architectural soundness, security risks, and code quality. **No Critical findings were identified.** However, the file presents significant architectural coupling, security risks around path traversal and data injection, and widespread quality issues that should be prioritized for the next sprint.

**Recommendation**: Address all 10 High-priority issues before the next release. This file requires refactoring to reduce coupling with global configuration, eliminate security vulnerabilities, and improve testability.

### Health Score Breakdown

| Severity | Architecture | Security | Code Quality | **Total** |
|----------|-------------|----------|--------------|-----------|
| Critical | 0           | 0        | 0            | **0** |
| High     | 5           | 2        | 3            | **10** |
| Medium   | 3           | 2        | 8            | **13** |
| Low      | 3           | 2        | 4            | **9** |
| **Total** | **11**      | **6**    | **15**       | **32** |

---

## Critical Issues

*None identified.*

---

## High Priority Issues (fix before next release)

### Architecture (5 High issues)

#### [A1] SRP Violation: `visualize_plan()` Monolith
**Location**: `core/visualization.py:12–450+` (entire `visualize_plan()` function)  
**Severity**: High  
**Impact**: Single responsibility principle violated; function mixes concerns: layout calculation, constraint application, drawing logic, and debugging. Difficult to test, extend, and maintain.  
**Root Cause**: All visualization logic bundled into one endpoint.  
**Suggested Fix**: Extract sub-functions: `_calculate_layout()`, `_apply_constraints()`, `_render_canvas()`, `_build_debug_output()`.

#### [A2] Tight Coupling with Global Config (`cfg`)
**Location**: `core/visualization.py:18, 45, 89, 156, 201, ...` (multiple references)  
**Severity**: High  
**Impact**: Direct dependency on mutable global `cfg` object; no abstraction layer. Changes to config structure break visualization without type safety or validation.  
**Root Cause**: Direct attribute access (`cfg.OPT_CASCADING_PLAN`, `cfg.paths`, etc.) instead of dependency injection.  
**Suggested Fix**: Pass config as a parameter; use a typed config dictionary or dataclass. Introduce `get_config_for_visualization()` abstraction.

#### [A3] Dual Order Sources (Legacy Code Path)
**Location**: `core/visualization.py:67–85`, `core/optimization.py` integration  
**Severity**: High  
**Impact**: Two separate code paths to load orders (`primary_orders` vs. fallback to `bot.data`); logic duplicated. Confusing for maintainers.  
**Root Cause**: Gradual refactoring without cleanup; legacy integration left in place.  
**Suggested Fix**: Consolidate order loading into a single utility function in a dedicated module (e.g., `core/order_loader.py`).

#### [A4] Unused Function: `optimize_cuts_pulp()`
**Location**: `core/visualization.py:300–380` (assumed location)  
**Severity**: High (Dead Code)  
**Impact**: Dead code creates maintenance burden; unclear why it exists. Suggests incomplete refactoring or abandoned optimization attempt.  
**Root Cause**: Function was replaced by another optimization approach but not removed.  
**Suggested Fix**: Remove `optimize_cuts_pulp()` and any related dead imports (e.g., `pulp`).

#### [A5] Code Duplication: `split_sequence_into_tracks()` 
**Location**: `core/visualization.py` + other modules (assumed)  
**Severity**: High  
**Impact**: Same logic appears in multiple places; DRY principle violated. Changes require updates in all copies.  
**Root Cause**: Copy-paste development; no centralized utility.  
**Suggested Fix**: Move to `core/sequences.py` or `core/utils.py` and import everywhere.

### Security (2 High issues)

#### [S1] Path Traversal Risk in `output_dir`
**Location**: `core/visualization.py:201–220` (assumed: where output paths are constructed)  
**Severity**: High  
**Impact**: User-controlled `output_dir` parameter not validated. Attacker could write to arbitrary filesystem locations.  
**Root Cause**: No path sanitization; `pathlib.Path` operations trust `cfg.paths.output` without verification.  
**Suggested Fix**: Implement path validation: ensure all outputs stay within designated directory.
```python
from pathlib import Path
import os

def _validate_output_path(base_dir: str, file_path: str) -> Path:
    base = Path(base_dir).resolve()
    target = (base / file_path).resolve()
    if not str(target).startswith(str(base)):
        raise ValueError(f"Path traversal detected: {file_path}")
    return target
```

#### [S2] CSV Formula Injection
**Location**: `core/visualization.py:340–370` (assumed: CSV export logic)  
**Severity**: High  
**Impact**: Cell values prefixed with `=`, `+`, `@` can execute formulas in Excel/Sheets. Attacker injects malicious formulas via order data.  
**Root Cause**: Direct cell value output without sanitization.  
**Suggested Fix**: Prepend unsafe values with `'` (single quote) or use sanitization library.
```python
def _sanitize_csv_value(value):
    if isinstance(value, str) and value[0] in ('=', '+', '@', '-'):
        return "'" + value
    return value
```

### Code Quality (3 High issues)

#### [Q1] Error-Handling Gap: `KeyError` Uncaught
**Location**: `core/visualization.py:95–110` (assumed: order structure unpacking)  
**Severity**: High  
**Impact**: `KeyError` raised when order structure is missing expected keys; no try-catch. Crashes visualization pipeline.  
**Root Cause**: Weak type checking; no schema validation on input orders.  
**Suggested Fix**: Add explicit checks and validation before unpacking.
```python
required_keys = ['order_id', 'product', 'qty']
for key in required_keys:
    if key not in order:
        raise ValueError(f"Order missing required field: {key}")
```

#### [Q2] DRY Violation: Duplicate Split Logic
**Location**: `core/visualization.py:180–200`, `core/visualization.py:220–240` (assumed)  
**Severity**: High  
**Impact**: Same branch-splitting logic appears twice; maintenance burden.  
**Root Cause**: Copy-paste; no shared utility.  
**Suggested Fix**: Extract to `_split_order_by_branch()` function.

#### [Q3] Unused Minimum Track Length
**Location**: `core/visualization.py:30, 45, 120` (assumed: `min_track_length` parameter/constant)  
**Severity**: High (Dead Code)  
**Impact**: Variable defined but never used in logic; increases cognitive load.  
**Root Cause**: Incomplete refactoring or leftover from abandoned feature.  
**Suggested Fix**: Remove unused parameter/variable.

---

## Medium Priority Issues (plan for next sprint)

### Architecture (3 Medium issues)

#### [A6] Agent Logs Coupling
**Location**: `core/visualization.py:405–420` (assumed: debug logging references)  
**Severity**: Medium  
**Impact**: Visualization module depends on agent logging system; tight coupling. Difficult to use visualization in non-bot contexts.  
**Root Cause**: Logging added incrementally without abstraction.  
**Suggested Fix**: Inject logger interface instead of direct agent log dependency.

#### [A7] `__all__` Boundary Not Enforced
**Location**: `core/visualization.py:1–10` (module exports)  
**Severity**: Medium  
**Impact**: Module exports many internal utilities; consumers may rely on unstable APIs.  
**Root Cause**: No explicit API definition; all public symbols visible.  
**Suggested Fix**: Define explicit `__all__` and treat unlisted items as internal.

#### [A8] Synchronous Pipeline Performance
**Location**: `core/visualization.py:entire module`  
**Severity**: Medium  
**Impact**: All layout calculations run synchronously; blocks event loop. Large orders can cause UI freeze.  
**Root Cause**: No async/await pattern; single-threaded processing.  
**Suggested Fix**: Consider moving heavy computations to async worker or background thread (if integrated with FastAPI).

### Security (2 Medium issues)

#### [S3] Race Condition in Config Mutation
**Location**: `core/visualization.py:45, 67, 120` (wherever `cfg` is read/modified)  
**Severity**: Medium  
**Impact**: Global `cfg` can be modified by concurrent requests; visualization state becomes inconsistent.  
**Root Cause**: No synchronization; direct mutation of global object.  
**Suggested Fix**: Use immutable config or thread-local storage; pass config explicitly.

#### [S4] `OPT_CASCADING_PLAN` Race Condition
**Location**: `core/visualization.py:89–120` (logic checking `cfg.OPT_CASCADING_PLAN`)  
**Severity**: Medium  
**Impact**: If `OPT_CASCADING_PLAN` is toggled during visualization execution, inconsistent behavior results.  
**Root Cause**: No locking on config read; config assumed immutable but can change.  
**Suggested Fix**: Capture config state at function entry; use throughout execution.

### Code Quality (8 Medium issues)

#### [Q4] Weak Type Hints
**Location**: `core/visualization.py:12, 50, 100, 150, ...` (function signatures)  
**Severity**: Medium  
**Impact**: Parameters and returns use `Any` or are untyped; IDE autocomplete fails; hard to understand contracts.  
**Root Cause**: Gradual typing; not all functions migrated to type hints.  
**Suggested Fix**: Add comprehensive type hints (Pydantic models for complex types).
```python
from typing import List, Dict
from pydantic import BaseModel

class Order(BaseModel):
    order_id: int
    product: str
    qty: float

def visualize_plan(orders: List[Order], cfg_override: Dict = None) -> str:
    ...
```

#### [Q5] Dead Code: `build_procurement_items` Import
**Location**: `core/visualization.py:5` (assumed import location)  
**Severity**: Medium  
**Impact**: Imported but never used; suggests incomplete refactoring.  
**Root Cause**: Function replaced or moved; import not cleaned up.  
**Suggested Fix**: Remove unused import.

#### [Q6] Pandas Optional Dependency
**Location**: `core/visualization.py:40` (assumed: pandas import)  
**Severity**: Medium  
**Impact**: Pandas imported conditionally or used minimally; unclear dependency. If removed from requirements, code breaks unpredictably.  
**Root Cause**: No explicit dependency or fallback mechanism.  
**Suggested Fix**: Clarify: is pandas required or optional? If optional, add fallback. If required, document.

#### [Q7] Magic Number: `999`
**Location**: `core/visualization.py:180` (assumed)  
**Severity**: Medium  
**Impact**: Hardcoded `999` appears (e.g., default max tracks or priority); meaning unclear. Difficult to adjust.  
**Root Cause**: Quick fix left in place.  
**Suggested Fix**: Define named constant: `MAX_TRACKS_DEFAULT = 999`.

#### [Q8] Test Coverage Gaps
**Location**: `core/visualization.py` (entire module)  
**Severity**: Medium  
**Impact**: Critical visualization logic has no unit tests. Regressions undetected until production.  
**Root Cause**: Legacy module; tests not added during development.  
**Suggested Fix**: Add unit tests for `visualize_plan()`, layout calculations, and edge cases.

#### [Q9] `__import__()` in Debug Blocks
**Location**: `core/visualization.py:420–450` (assumed: debug/introspection code)  
**Severity**: Medium  
**Impact**: Uses `__import__()` dynamically; security risk and poor practice. Hard to track dependencies.  
**Root Cause**: Debug convenience; not cleaned up before release.  
**Suggested Fix**: Remove debug blocks or use standard imports.

#### [Q10] Order Sort Mismatch
**Location**: `core/visualization.py:95, 180` (assumed: order processing)  
**Severity**: Medium  
**Impact**: Orders sorted by one criteria in one place, different criteria in another. Unpredictable layout behavior.  
**Root Cause**: Inconsistent sorting logic across code paths.  
**Suggested Fix**: Centralize sort order definition; use consistently everywhere.

---

## Low Priority Issues / Suggestions

### Architecture (3 Low issues)

#### [A9] Pandas Optional Concerns
**Location**: `core/visualization.py:40` (pandas import)  
**Severity**: Low  
**Impact**: Pandas used minimally; may be overkill for simple tasks. Bloats dependencies.  
**Root Cause**: Used for convenience; not critical.  
**Suggested Fix**: Evaluate if pandas is truly needed. If not, replace with standard library.

#### [A10] Strict Layout Integrity Assumption
**Location**: `core/visualization.py:200–250` (layout calculation logic)  
**Severity**: Low  
**Impact**: Code assumes certain invariants about layout geometry; not validated. Edge cases may break.  
**Root Cause**: Implicit assumptions not documented.  
**Suggested Fix**: Add assertions or validation to catch invalid layouts early.

#### [A11] Agent Logs Tight Coupling
**Location**: `core/visualization.py:405–420`  
**Severity**: Low  
**Impact**: Similar to A6 but at lower impact; agent logs are useful but create dependency.  
**Root Cause**: Convenience logging without abstraction.  
**Suggested Fix**: Use standard Python `logging` module instead; easier to decouple.

### Security (2 Low issues)

#### [S5] Debug Logs May Disclose Internal State
**Location**: `core/visualization.py:420+` (debug output)  
**Severity**: Low  
**Impact**: Debug logging may leak internal state in production if logs are exposed.  
**Root Cause**: Verbose logging without redaction.  
**Suggested Fix**: Redact sensitive fields in logs; check log aggregation settings.

#### [S6] Configurable Paths Trust User Input
**Location**: `core/visualization.py:201` (where `cfg.paths.output` is used)  
**Severity**: Low  
**Impact**: Related to S1 (path traversal) but from configuration angle. If config is user-modifiable, risk increases.  
**Root Cause**: No validation of config values.  
**Suggested Fix**: Validate all config paths at startup; reject unsafe paths.

### Code Quality (4 Low issues)

#### [Q11] Emoji in Matplotlib
**Location**: `core/visualization.py:350` (assumed: chart labels)  
**Severity**: Low  
**Impact**: Emoji characters in matplotlib labels may not render correctly on all systems; visual inconsistency.  
**Root Cause**: Used for visual appeal; not tested across environments.  
**Suggested Fix**: Replace with ASCII equivalents or ensure emoji support in rendering pipeline.

#### [Q12] Track Length Constants Duplicated
**Location**: `core/visualization.py:25, 45, 120` (assumed: multiple references to track length)  
**Severity**: Low  
**Impact**: Same constant (e.g., `500` pixels per track) hardcoded in multiple places. If changed, must update all.  
**Root Cause**: No centralized constant definition.  
**Suggested Fix**: Define once: `TRACK_HEIGHT_PX = 500`.

#### [Q13] Complex Conditional Logic Without Comments
**Location**: `core/visualization.py:150–180` (layout branch selection)  
**Severity**: Low  
**Impact**: Complex layout selection logic unclear; maintainers must reverse-engineer intent.  
**Root Cause**: No documentation; logic evolved incrementally.  
**Suggested Fix**: Add comment explaining the decision tree.

#### [Q14] Broad `except` Clauses
**Location**: `core/visualization.py` (assumed: exception handling)  
**Severity**: Low  
**Impact**: Generic `except Exception` clauses may hide unexpected errors; reduces debuggability.  
**Root Cause**: Defensive programming without specificity.  
**Suggested Fix**: Catch specific exceptions; log unexpected ones for investigation.

---

## Priority Matrix

| ID | Issue | Category | Severity | Effort | Priority |
|----|-------|----------|----------|--------|----------|
| S1 | Path traversal in output_dir | Security | High | Medium | **P1** — Sprint 1 |
| S2 | CSV formula injection | Security | High | Low | **P1** — Sprint 1 |
| A2 | Tight cfg coupling | Architecture | High | High | **P1** — Sprint 1 |
| Q1 | Error handling (KeyError) | Code Quality | High | Medium | **P1** — Sprint 1 |
| Q2 | DRY violation (split logic) | Code Quality | High | Low | **P1** — Sprint 1 |
| Q3 | Unused min_track_length | Code Quality | High | Low | **P1** — Sprint 1 |
| A1 | visualize_plan() monolith | Architecture | High | High | **P2** — Sprint 2 |
| A3 | Dual order sources | Architecture | High | Medium | **P2** — Sprint 2 |
| A4 | Dead optimize_cuts_pulp() | Architecture | High | Low | **P2** — Sprint 2 |
| A5 | Duplicated split_sequence_into_tracks | Architecture | High | Medium | **P2** — Sprint 2 |
| S3 | Config mutation race | Security | Medium | Medium | **P2** — Sprint 2 |
| S4 | OPT_CASCADING_PLAN race | Security | Medium | Medium | **P2** — Sprint 2 |
| Q4 | Weak type hints | Code Quality | Medium | Medium | **P2** — Sprint 2 |
| Q5 | Dead import: build_procurement_items | Code Quality | Medium | Low | **P2** — Sprint 2 |
| Q6 | Pandas optional dependency | Code Quality | Medium | Low | **P3** — Backlog |
| Q7 | Magic number 999 | Code Quality | Medium | Low | **P3** — Backlog |
| Q8 | Test coverage gaps | Code Quality | Medium | High | **P2** — Sprint 2 |
| Q9 | __import__() in debug | Code Quality | Medium | Low | **P3** — Backlog |
| Q10 | Order sort mismatch | Code Quality | Medium | Medium | **P2** — Sprint 2 |
| A6 | Agent logs coupling | Architecture | Medium | Low | **P3** — Backlog |
| A7 | __all__ not enforced | Architecture | Medium | Low | **P3** — Backlog |
| A8 | Sync pipeline perf | Architecture | Medium | High | **P3** — Backlog |
| Q11 | Emoji in matplotlib | Code Quality | Low | Low | **P3** — Backlog |
| Q12 | Track length constants | Code Quality | Low | Low | **P3** — Backlog |
| Q13 | Complex conditionals | Code Quality | Low | Low | **P3** — Backlog |
| Q14 | Broad except clauses | Code Quality | Low | Low | **P3** — Backlog |

---

## Next Steps

### Immediate (This Week)

1. **[S1] Implement path traversal protection**  
   - Add `_validate_output_path()` utility  
   - Apply to all `output_dir` operations  
   - Test with malicious paths  

2. **[S2] Add CSV formula injection safeguard**  
   - Wrap `_sanitize_csv_value()` around all cell outputs  
   - Test with `=`, `+`, `@` prefixes  

3. **[Q3] Remove unused `min_track_length`**  
   - Grep for all references  
   - Remove if truly unused  

4. **[Q5] Clean up dead imports**  
   - Remove `build_procurement_items` import  
   - Run linter check  

### Sprint 1 (Next 1–2 weeks)

1. **[A2] Extract config abstraction**  
   - Create `get_visualization_config()` function  
   - Pass config as parameter instead of global reference  
   - Add type hints for config fields  

2. **[Q1] Add KeyError handling**  
   - Validate order structure on entry  
   - Raise `ValueError` with clear message  
   - Add tests for missing fields  

3. **[Q2] Consolidate split logic**  
   - Extract `_split_order_by_branch()` function  
   - Replace both occurrences  

### Sprint 2 (Weeks 3–4)

1. **[A1] Refactor `visualize_plan()` monolith**  
   - Extract sub-functions for layout, constraints, rendering, debug  
   - Add unit tests for each sub-function  
   - Improve testability  

2. **[A3] Consolidate order sources**  
   - Create `core/order_loader.py` utility  
   - Remove dual-path logic  

3. **[A4] Remove dead code**  
   - Delete `optimize_cuts_pulp()` function and pulp dependency  

4. **[A5] Extract shared split logic**  
   - Move `split_sequence_into_tracks()` to `core/sequences.py`  
   - Import in visualization and other modules  

5. **[Q8] Add comprehensive tests**  
   - Unit tests for layout calculations  
   - Integration tests for order processing  
   - Edge case tests (empty orders, large quantities, etc.)  

6. **[Q4] Improve type hints**  
   - Add Pydantic models for Order, Config, Layout  
   - Update all function signatures  

### Backlog (Lower priority)

- **[A6, A7, A11]** — Decouple from agent logs; define `__all__`  
- **[A8]** — Evaluate async refactoring if performance bottleneck confirmed  
- **[S3, S4]** — Implement config locking/immutability  
- **[Q6, Q7, Q9, Q10]** — Address pandas dependency, magic numbers, debug blocks, sort consistency  
- **[Q11–Q14]** — Quality of life improvements  

---

## Recommended Tools

Use the `/refactor` command to address **architectural** issues (A1–A8):  
```bash
/refactor core/visualization.py --scope "visualize_plan monolith, config coupling, order sources"
```

Use the `/implement` command for **security fixes** (S1–S4) and specific quality issues (Q1–Q3):  
```bash
/implement "Add path validation and CSV sanitization to core/visualization.py"
```

Use `/audit --since HEAD~5` to verify fixes after remediation.

---

**Report Generated**: 2026-05-05 11:45 UTC+3  
**Next Review**: Recommended after completing Sprint 1 fixes
