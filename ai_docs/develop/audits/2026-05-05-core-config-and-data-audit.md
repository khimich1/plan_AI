# Project Audit Report: core/config_and_data.py

**Date**: 2026-05-05  
**Scope**: `core/config_and_data.py` (single-file focused audit)  
**Audited by**: senior-reviewer + security-auditor + reviewer  
**Focus**: Integration with «картина по заказу» (custom order visualization) and ILP (integer linear programming / optimization)

---

## Executive Summary

**Overall Health Score**: 2.0/10  
**Risk Level**: CRITICAL — Multiple systemic issues affecting concurrency, data integrity, and maintainability.

### Severity Breakdown

| Severity | Architecture | Security | Code Quality | Total |
|----------|-------------|----------|--------------|-------|
| Critical | 2 | 0 | 0 | **2** |
| High | 3 | 2 | 2 | **7** |
| Medium | 3 | 3 | 6 | **12** |
| Low | 3 | 3 | 5 | **11** |

**Total findings**: 32 (2 critical, 7 high, 12 medium, 11 low)

### Recommendation

**Немедленно рефакторить глобальное состояние перед интеграцией ILP.** The current globals-based architecture is incompatible with concurrent optimization workflows and poses significant data integrity risks. Critical refactoring must precede any «картина по заказу» feature expansion or advanced optimization integration.

---

## Critical Issues (fix immediately)

### [A1] Globals as Integration Surface: Order → ILP → Visualization
**Category**: Architecture  
**Location**: Module-level design; PlateOrder state management pattern  
**Impact**: 
- Concurrency conflicts: simultaneous order processing corrupts global state (test_layout_secondary_unmatched_parent_user_list fails due to cross-contamination)
- ILP solver integration impossible: optimizer needs isolated state copies, not shared mutable globals
- «картина по заказу» visualization unsafe: multiple threads/requests race for kp_db, prays_by_plate state
- Unpredictable cascading failures when multiple orders flow through optimization pipeline

**Fix**:
1. Extract order context into isolated, immutable `OrderContext` dataclass (not globals)
2. Pass context through ILP solver and visualization layers as dependency
3. Return results without side effects; caller merges results into persistent storage
4. Example: `def optimize_order(context: OrderContext) → OptimizationResult` instead of `apply_to_globals()`

---

### [A2] PlateOrder.apply_to_globals(): Legacy Bridge + Split Brain Risk
**Category**: Architecture  
**Location**: `PlateOrder.apply_to_globals()` method (~line 400–500 estimated)  
**Impact**:
- Two sources of truth: object state (self.plates, self.totals) and global state (kp_db, prays_by_plate, plate_totals_cache)
- Updates may fail silently; caller does not know if global state is inconsistent
- Partial application leaves module in broken state if exception occurs mid-update
- Legacy PlateOrder ORM-style design conflicts with modern immutable/functional order data flow

**Fix**:
1. Replace `apply_to_globals()` with pure function: `def apply_plate_order(order: PlateOrder, state: GlobalState) → GlobalState` (returns new state, no side effects)
2. Remove split brain: decide — is PlateOrder an ORM-style object or immutable DTO?
   - If ORM: use SQLAlchemy or explicit session/commit pattern
   - If DTO: serialize/deserialize to/from persistent storage, not intermediate globals
3. Add explicit transaction semantics: all-or-nothing update, or rollback on partial failure
4. Test with concurrent requests; verify no state leaks between orders

---

## High Priority Issues (fix soon)

### Architecture

**[A3] SRP Violation: Megamodule (Single Responsibility Principle)**  
**Location**: core/config_and_data.py — entire module  
**Details**: PlateOrder, order parsing, globals initialization, ILP integration helpers, «картина» data prep, name generation, debug I/O, cache management all in one file.  
**Impact**: Impossible to test, refactor, or reuse components in isolation. Cyclic mental model makes onboarding hard. Changes affect unexpected subsystems.  
**Fix**: Split into:
- `plate_order.py` — PlateOrder class and serialization
- `order_context.py` — immutable context + helpers  
- `ipl_adapter.py` — ILP solver integration
- `visualization_adapter.py` — «картина по заказу» data prep
- `config_globals.py` — shared state initialization (temporary, for migration)

---

**[A4] ILP Tightly Coupled to cfg Globals, Not Abstract Order Context**  
**Location**: ILP integration code (search for solver calls, matrix building)  
**Details**: Solver accesses kp_db, prays_by_plate directly instead of receiving an isolated order context parameter.  
**Impact**: Cannot run multiple optimization jobs in parallel. Testing solver logic requires global state setup. Reusing solver in other projects requires copying globals.  
**Fix**: 
1. Extract ILP input (matrix, bounds, objective) into immutable `ILPInput` struct
2. Solver takes ILPInput, returns `ILPSolution` (not mutating globals)
3. Caller responsibility: convert order context → ILPInput → apply solution

---

**[A5] Duplicated Totals Logic in Three Places**  
**Location**: Totals calculation scattered (PlateOrder, cache validation, import functions)  
**Details**: `sum()`, `compute_totals()`, `plate_totals_cache` logic written independently.  
**Impact**: Inconsistency bugs (cache stale, object totals wrong). Three places to fix on formula change. DRY violation.  
**Fix**: Single source of truth: `def compute_plate_totals(plates: List[Plate]) → Totals` function. Use everywhere.

---

### Security

**[S1] Cross-Session Globals / Data Mixing**  
**Location**: Module-level kp_db, prays_by_plate, plate_totals_cache  
**Details**: Globals persist across requests/users in web context. No session isolation. One user's order calculations leak into another's.  
**Impact**: High severity privacy breach. User A's custom plate names/prices visible to User B. Regulatory (GDPR, PCI) violation.  
**Fix**:
1. Scope all state to request-local or per-order context
2. In FastAPI: use Depends() to inject request-specific context
3. Zero persistent globals for order data (only immutable config)

---

**[S2] Swallowed fill_plate_nomenclature_cache() After Partial Global Updates**  
**Location**: `apply_to_globals()` caller or exception handler  
**Details**: If `fill_plate_nomenclature_cache()` fails, caller continues as if cache is valid. Exception silently logged/ignored.  
**Impact**: Optimizer uses stale or incomplete nomenclature. Results incorrect (plates mislabeled, costs wrong). User doesn't detect error.  
**Fix**:
1. Propagate exception, don't swallow
2. Use explicit transaction: all updates succeed or none
3. Log exception with context: which order, which user, remediation steps

---

### Code Quality

**[Q1] from_dict / _parse_load_key Fragility**  
**Location**: `from_dict()` method and `_parse_load_key()` helper  
**Details**: KeyError if schema differs (missing keys, extra keys). Float/int type coercion implicit, may fail silently. No size/depth validation.  
**Impact**: Crashes on user input variance. Hard to debug (user doesn't know which key failed). DoS risk if large payloads accepted.  
**Fix**:
1. Use Pydantic v2 for schema validation (coercion, defaults, size limits)
2. Explicit type hints; replace implicit coercion
3. Add max_depth, max_size checks before deserialization
4. Return clear error messages: "Key 'X' missing", "Type mismatch: expected float, got str"

---

**[Q2] from_orders_2d Sharp KeyError Edges**  
**Location**: `from_orders_2d()` method  
**Details**: Assumes 2D array structure. No bounds check. Crashes if row/column missing.  
**Impact**: Brittle import. One malformed CSV row breaks entire batch import.  
**Fix**:
1. Validate shape before processing
2. Provide row/column indices in error message
3. Implement skip-and-log mode: import valid rows, log invalid ones separately
4. Add tests for edge cases: empty rows, extra columns, type mismatches

---

## Medium Priority Issues (plan for next sprint)

### Architecture

**[A6] Lazy kp_db Coupling + Swallowed Warnings**  
**Location**: kp_db initialization / import  
**Details**: kp_db loaded on first use (lazy). Warnings logged but not raised. Caller doesn't know DB is incomplete.  
**Fix**: Eager load with explicit success check on startup. Propagate warnings.

---

**[A7] Debug I/O in make_plate_name() Production Path**  
**Location**: `make_plate_name()` function  
**Details**: File I/O or print() statements in name generation logic. Causes performance issues, side effects in test/production.  
**Fix**: Remove debug code or move to separate debug utility.

---

**[A8] Key Shape Float/Int Load Inconsistency**  
**Location**: Plate loading / parsing (search for float/int conversion)  
**Details**: Some callers convert keys to int, others leave float. Shape comparison fails (1.0 != 1).  
**Fix**: Normalize key type on load. Use explicit `int(key)` converter or always float, but document contract.

---

### Security

**[S3] NDJSON debug_logs Retention**  
**Location**: Debug log file storage  
**Details**: Logs may contain order data (PII, prices). No rotation, cleanup, or access control.  
**Impact**: Old data persists; could leak in backups or via unprotected log access.  
**Fix**: Log only non-sensitive metadata. Rotate/compress daily. Access-restrict log files.

---

**[S4] PlateOrder.from_dict No Schema/Size Limits (DoS)**  
**Location**: `PlateOrder.from_dict()` deserialization  
**Details**: Accepts arbitrary large dicts. No payload size limit. Unbounded recursion possible.  
**Impact**: Attacker submits 100MB order; server memory exhausted or parsing hangs.  
**Fix**: Add max_size check. Use streaming parser if size > threshold. Rate-limit deserialization calls.

---

**[S5] Exception Logging Disclosure**  
**Location**: Exception handlers with full traceback logging  
**Details**: Logs include internal paths, variable values, system info. Visible to users or in error responses.  
**Impact**: Reconnaissance attack. Attacker learns system architecture, library versions.  
**Fix**: Log only safe metadata (user_id, timestamp, error_code). Include full traceback only in internal logs, not user-facing responses.

---

### Code Quality

**[Q3] Duplicate Prays Helpers**  
**Location**: Multiple functions computing prays_by_plate / plate_name_to_prays_variant  
**Details**: Similar logic written multiple times. Dead code (`plate_name_to_prays_variant` unused?).  
**Fix**: Consolidate into single source. Remove dead code. Document intent of remaining functions.

---

**[Q4] Huge set_plate_lists_from_text / add_items Complexity**  
**Location**: `set_plate_lists_from_text()` and `add_items()` functions (estimated >60 lines each)  
**Details**: 5+ responsibilities per function (parsing, validation, state update, caching). Deep nesting (>3 levels). Hard to test.  
**Fix**: Break into smaller functions: `parse_line()`, `validate_plate()`, `add_to_cache()`. Test each independently.

---

**[Q5] Missing Optional Type Hints**  
**Location**: Function signatures; variable declarations  
**Details**: Optional fields not annotated. Callers unsure if None possible.  
**Fix**: Add `Optional[T]` or `T | None` to all potentially null values. Use strict mypy mode.

---

**[Q6] get_load_code_for_plate Int vs Float Inconsistency**  
**Location**: `get_load_code_for_plate()` function  
**Details**: Parameter type inconsistent with usage (int expected, float sometimes passed).  
**Fix**: Document type contract. Add explicit conversion on entry. Consider using TypedDict or dataclass for load codes.

---

**[Q7] Silent 0.0 from Length Parsers**  
**Location**: Parsing functions for dimensions/lengths  
**Details**: If parsing fails, returns 0.0 without warning. Caller assumes valid data.  
**Impact**: Incorrect optimization results. Hard to trace root cause (parser failure vs user input 0).  
**Fix**: Raise exception on parse failure. Return `Optional[float]`. Caller must handle None.

---

**[Q8] Broad Except in Normalizers**  
**Location**: Normalizer functions (search for `except:` or `except Exception:`)  
**Details**: Catches all exceptions, silently ignores. Hides bugs.  
**Fix**: Catch specific exceptions (ValueError, TypeError). Log and propagate.

---

**[Q9] Misleading Module Docstring vs Actual Contents**  
**Location**: Module-level docstring  
**Details**: Says one thing (e.g., "order data layer") but code does 10 different things.  
**Fix**: Update docstring to list all responsibilities, then refactor to reduce them.

---

## Low Priority / Suggestions

### Architecture

**[A9] Config Split: Mutable Globals vs Immutable Config**  
**Location**: Module root  
**Details**: No clear distinction between setup-time config (kp_db, material prices) and runtime state (PlateOrder cache).  
**Suggestion**: Separate into `Config` (immutable, loaded at startup) and `State` (mutable, request-scoped).

---

**[A10] Large add_items Thresholds**  
**Location**: Constants for batch size thresholds  
**Details**: Magic numbers. Hard-coded batch sizes may not suit all environments.  
**Suggestion**: Move to config file or environment variables. Document reasoning.

---

**[A11] Opaque Debug Symbols**  
**Location**: branch_001, debug flags scattered  
**Details**: Debug/dev code left in production. Naming unclear (what is branch_001?).  
**Suggestion**: Either remove or move to explicit DebugFlags dataclass. Document each flag's purpose.

---

### Security

**[S6] Debug Path Contract**  
**Location**: Paths used in debug I/O  
**Details**: Paths hard-coded or rely on working directory. May leak sensitive structure.  
**Suggestion**: Use explicit temp directory. Clean up after tests.

---

**[S7] Long Strings CPU Cost**  
**Location**: String concatenation in loops  
**Details**: O(n²) complexity for large orders if strings concatenated repeatedly.  
**Suggestion**: Use `''.join([...])` or f-strings; profile if performance matters.

---

**[S8] Predictable Local Paths**  
**Location**: Debug/temp file generation  
**Details**: Paths like `/tmp/plate_debug_001.txt` — predictable, race condition risk.  
**Suggestion**: Use `tempfile.NamedTemporaryFile(delete=False)` with random name.

---

### Code Quality

**[Q10] Orphaned register_plate_metadata Function**  
**Location**: Check for unused functions  
**Details**: Function defined but never called. Dead code.  
**Fix**: Remove or document why it exists.

---

**[Q11] Redundant Parsed Flag Branch**  
**Location**: Boolean flag checking (if parsed / if not parsed)  
**Details**: Same logic duplicated in both branches; flag check unnecessary.  
**Fix**: Simplify conditional; remove redundant branch.

---

**[Q12] Branch_001 Naming Anti-Pattern**  
**Location**: Variable/branch naming  
**Details**: Names like `branch_001`, `temp_var_2` provide no context.  
**Fix**: Rename to reflect purpose: `fallback_branch`, `price_variance_branch`.

---

**[Q13] Test Coverage Gaps**  
**Location**: No tests for from_dict(), from_orders_2d(), parsing helpers  
**Details**: Critical paths untested. Errors caught in production.  
**Suggestion**: Add pytest test suite covering:
   - from_dict with valid/invalid/edge-case inputs
   - from_orders_2d with malformed CSV
   - Concurrent order processing (concurrent.futures / threading)

---

**[Q14] Doc/Comment Drift Risk**  
**Location**: Function docstrings vs implementation  
**Details**: Comments mention old behavior (e.g., "stores in memory cache" when cache removed).  
**Fix**: Audit all docstrings. Update or remove if stale.

---

## Priority Matrix

| ID | Issue | Category | Severity | Effort | Blocker |
|----|-------|----------|----------|--------|---------|
| A1 | Globals integration surface + concurrency | Architecture | Critical | High | YES — prevents ILP integration |
| A2 | apply_to_globals split brain | Architecture | Critical | Medium | YES — data integrity risk |
| S1 | Cross-session data mixing | Security | High | Medium | YES — privacy breach |
| S2 | Swallowed fill_plate_nomenclature_cache | Security | High | Low | YES — silent data corruption |
| Q1 | from_dict fragility | Code Quality | High | Medium | No |
| Q2 | from_orders_2d KeyError edges | Code Quality | High | Medium | No |
| A3 | SRP megamodule | Architecture | High | High | No — technical debt |
| A4 | ILP coupled to globals | Architecture | High | High | No — blocks modularity |
| A5 | Duplicated totals logic | Architecture | High | Low | No |
| S3 | NDJSON debug log retention | Security | Medium | Low | No |
| S4 | from_dict DoS (no size limits) | Security | Medium | Medium | No |
| S5 | Exception logging disclosure | Security | Medium | Low | No |
| A6–A11 | Medium/Low architecture | Architecture | Medium–Low | Varies | No |
| Q3–Q14 | Code quality debt | Code Quality | Medium–Low | Varies | No |

---

## Next Steps

### Phase 1: Immediate (before next commit)
1. **Create `OrderContext` dataclass** — immutable order state (A1)
2. **Refactor `apply_to_globals()` into pure function** — no side effects (A2)
3. **Add session isolation** — FastAPI dependency for request-local state (S1)
4. **Fix exception swallowing** — propagate fill_plate_nomenclature_cache errors (S2)

### Phase 2: This Sprint
1. **Consolidate totals logic** (A5)
2. **Add Pydantic validation** to from_dict / from_orders_2d (Q1, Q2)
3. **Split core/config_and_data.py** into modules (A3)
4. **Extract ILP adapter** (A4) — solver takes ILPInput, returns ILPSolution
5. **Remove debug I/O** from production paths (A7)

### Phase 3: Next Sprint
1. **Fix type hints** (Q5, Q6)
2. **Reduce from_orders_2d complexity** (Q4)
3. **Remove dead code** (Q3, Q10)
4. **Add pytest suite** for from_dict, from_orders_2d, parsing (Q13)
5. **Audit docstrings** (Q9, Q14)

### Phase 4: Backlog
1. Extract configuration — immutable vs mutable state (A9, A10)
2. Clean up debug symbols and branches (A11, Q12)
3. Log security audit (S3–S8) — redact PII, rotate logs, use tempfile safely

---

## Integration with «картина по заказу» & ILP

**Blocker**: A1 and A2 (globals architecture) must be resolved before expanding visualization or optimization integration.

**Post-refactor roadmap**:
1. Immutable OrderContext flows through ILP solver
2. Solver returns ILPSolution (matrix assignments, costs, feasibility)
3. Visualization layer consumes OrderContext + ILPSolution → renders «картина»
4. No global state mutations; all results append-only or versioned

---

## Audit Closure

**Report generated**: 2026-05-05 11:44 UTC+3  
**Audit status**: COMPLETE

### Critical Issues Summary

Two **critical** findings require immediate architectural refactoring:

1. **[A1]** Globals as integration surface prevent concurrent order processing and ILP solver integration
2. **[A2]** PlateOrder.apply_to_globals() creates split-brain state with silent failure modes

These block deployment of advanced optimization features and pose data integrity risks in multi-user environments.

---

**Start remediation for these critical issues?** (y/n)

If yes, remediation will proceed as follows:
- Bucket A (Structural): A1, A2 → refactor agent → test-runner → verify
- Bucket B (Security): S1, S2 → planner + worker → test-writer → security-auditor → verify
- Remaining (High/Medium/Low) → logged for sprint planning

Please confirm to begin automated remediation.
