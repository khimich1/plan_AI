# Consolidated Audit Report: Core Optimization, Config, Layout Sequence

**Date:** 2026-05-07  
**Scope:** `core/optimization/`, `core/config_and_data.py`, `viz_modules/layout_sequence.py`  
**Audited by:** senior-reviewer + security-auditor + reviewer (Composer 2)

---

## Executive Summary

This audit covers the optimization pipeline, central configuration/data loading, and the layout-sequence visualization module. The area mixes solver configuration, ILP modeling, 2D orchestration, and heavy UI-adjacent geometry—creating cohesion and isolation risks around global proxies, module size, and request-scoped state.

**Overall health score:** **4.0 / 10**

Scoring follows the project formula: start at 10; −2 per critical (cap −6); −0.5 per high (cap −3); −0.1 per medium (cap −1); floor at 0, rounded to one decimal. With **1 critical**, **8 high**, and **13 medium** findings, deductions are −2, −3 (capped), and −1 (capped), yielding **4.0**.

**Recommendation:** Treat **A1** (TLS/global `OPT_*` proxy isolation) as a release blocker for any multi-tenant or parallel-use scenario. Schedule **high** items (SRP split of `config_and_data`, `layout_sequence` decomposition, DIP for `ilp_model`, gated debug I/O, and **S1** cross-request state) in the current sprint before expanding features in this layer.

---

## Severity Summary

| Severity  | Architecture | Security | Code Quality | **Total** |
|-----------|-------------:|---------:|-------------:|----------:|
| Critical  | 1            | 0        | 0            | **1**     |
| High      | 4            | 1        | 3            | **8**     |
| Medium    | 5            | 2        | 6            | **13**    |
| Low       | 2            | 3        | 4            | **9**     |

---

## Critical Issues (fix immediately)

### [A1] TLS/global module proxy `OPT_*` isolation risk

**Category:** Architecture  
**Location:** `core/optimization/__init__.py`, `core/optimization/context.py` (`optimization_context_scope` and related)

**Impact:** Module-level or context-scoped exposure of optimization toggles can leak configuration between callers, tests, or concurrent flows. Wrong `OPT_*` values can change solver behavior silently (incorrect plans, non-deterministic runs, hard-to-reproduce prod bugs).

**Summary:** TLS/ContextVar-based proxies for `OPT_*` without strict boundaries risk cross-talk where the same process serves multiple logical “requests” or imports order changes defaults.

**Fix hint:** Narrow the public surface; document lifetime rules; ensure every entry point uses `optimization_context_scope` (or equivalent) consistently; add tests that prove isolation under nested/concurrent contexts.

---

## High Priority Issues (fix soon)

### [A2] `config_and_data` monolith (SRP violation)

**Category:** Architecture  
**Location:** `core/config_and_data.py`

**Impact:** Changes to one concern (paths, DB, caching, parsing) risk regressions elsewhere; harder onboarding and unsafe edits.

**Summary:** Single module aggregates loading, paths, and data access concerns.

**Fix hint:** Split into focused modules (e.g. paths, SQLite access, DTO/loaders) behind small facades; keep backward-compatible imports if needed.

### [A3] `layout_sequence.py` god module

**Category:** Architecture  
**Location:** `viz_modules/layout_sequence.py`

**Impact:** High change cost, review fatigue, and hidden coupling between build, I/O, and geometry.

**Summary:** Very large module mixing build pipeline, formatting, and side effects.

**Fix hint:** Extract builders, render/format helpers, and file/debug paths into submodules with explicit dependencies.

### [A4] `ilp_model` DIP violation (price DB / config)

**Category:** Architecture  
**Location:** `core/optimization` (ILP model construction vs. price DB and config sources)

**Impact:** Hard to test, swap pricing sources, or run in isolation; tight coupling to global/config state.

**Summary:** Model layer reaches concrete config/DB instead of injected ports.

**Fix hint:** Introduce interfaces or callables for price lookups and settings; wire in orchestrator/factory only.

### [A5] Ungated debug I/O in `layout_sequence`

**Category:** Architecture / operability  
**Location:** `viz_modules/layout_sequence.py`

**Impact:** Unexpected disk writes, performance hits, and noisy production logs when flags are mis-set.

**Summary:** Debug output not consistently behind a single debug gate or level.

**Fix hint:** Route all debug writes through `debug_log` / `optimization_config` (or one shared gate) and default off in production paths.

### [S1] Cross-request state mixing (thread-local / ContextVar modules)

**Category:** Security / reliability  
**Location:** `core/optimization/context.py`, consumers of context/thread-local state across `core/optimization/`

**Impact:** In multi-worker or async-heavy use, state bleed can mix plans, user data, or toggles—integrity and confidentiality risk at the application boundary.

**Summary:** Modules that rely on thread-local or `ContextVar` without invariant enforcement can confuse logical request boundaries.

**Fix hint:** Audit all readers/writers; document “one context per run”; add integration tests for async and threaded entry points; prefer explicit request-scoped bags where feasible.

### [Q1] Megafunction: `layout_sequence` build path

**Category:** Code quality  
**Location:** `viz_modules/layout_sequence.py`

**Impact:** Hard to test, reason about branches, and refactor safely.

**Summary:** Primary “build” flow is excessively long and multi-responsibility.

**Fix hint:** Extract phases with clear inputs/outputs; add unit tests per phase.

### [Q2] Megafunction: `finalize.run_two_d_phase_finalize`

**Category:** Code quality  
**Location:** `core/optimization/optimize_2d/finalize.py`

**Impact:** Regression risk when touching 2D completion logic.

**Summary:** `run_two_d_phase_finalize` concentrates too many steps.

**Fix hint:** Split validation, mutation, and persistence steps; align with `prep_solve` structure.

### [Q3] Megafunction: `prep_solve.run_two_d_phase_a`

**Category:** Code quality  
**Location:** `core/optimization/optimize_2d/prep_solve.py`

**Impact:** Same as Q2 for the “phase A” preparation path.

**Summary:** `run_two_d_phase_a` is a large orchestration blob.

**Fix hint:** Same as Q2—named substeps and tests per boundary.

---

## Medium Priority Issues (plan for next sprint)

### [A6] Hardcoded `pb.db` (or equivalent product DB name)

**Category:** Architecture / configuration  
**Location:** References across `core/config_and_data.py` / call sites

**Summary:** Database filename embedded instead of centralized config.

**Fix hint:** Single source of truth in settings or `db_config`; env override for deployments.

### [A7] `geometry.py` wide config import

**Category:** Architecture  
**Location:** Geometry helpers (under `core/` / optimization-related geometry)

**Summary:** Pulls a large config surface for small geometric needs—increases import graph churn.

**Fix hint:** Pass required scalars or a minimal protocol into geometry functions.

### [A8] Duplicate `load_code` logic

**Category:** Architecture / DRY  
**Location:** `core/optimization/layout_runtime_snapshot.py` vs. `core/config_and_data.py`

**Summary:** Two implementations risk drift in encoding and edge cases.

**Fix hint:** One canonical `load_code` in a shared util; thin wrappers elsewhere.

### [A9] N+1 `get_reinforcement` loop

**Category:** Architecture / performance  
**Location:** Hot paths calling reinforcement lookup per item (orchestration / planning)

**Summary:** Repeated DB or service calls in a loop.

**Fix hint:** Batch fetch, cache per plan slice, or pre-index by key.

### [A10] `print` in `orchestrator`

**Category:** Architecture / observability  
**Location:** `core/optimization/orchestrator.py`

**Summary:** Unstructured stdout instead of logging.

**Fix hint:** Replace with module logger; respect log level and debug gates.

### [S2] Sensitive disk logs without unified debug gate

**Category:** Security  
**Location:** Various `core/optimization/` and `viz_modules/layout_sequence.py` debug paths

**Summary:** Log files may contain layout or order details without uniform redaction or enablement policy.

**Fix hint:** Central debug policy (path, retention, env flag), avoid logging secrets/PII; document what may appear in logs.

### [S3] Hardcoded SQLite path vs. `db_config`

**Category:** Security / portability  
**Location:** SQLite path construction (near `core/config_and_data.py` or callers)

**Summary:** Divergence between hardcoded paths and `db_config` risks wrong DB or accidental overwrite.

**Fix hint:** Always resolve DB via one module; tests for relative/absolute path behavior on Windows.

### [Q4] Validation documentation drift

**Category:** Code quality  
**Location:** `core/optimization/validation.py` and related docs/comments

**Summary:** Documented rules no longer match implementation.

**Fix hint:** Sync docstrings or `ai_docs` with code; add tests that encode the contract.

### [Q5] Split `load_code` rules unclear

**Category:** Code quality  
**Location:** Load/encode helpers split across modules

**Summary:** Maintainers unsure which entry point to extend.

**Fix hint:** Document single owner (see A8); add docstring matrix “when to use X.”

### [Q6] `order_dispatch` complexity

**Category:** Code quality  
**Location:** `core/optimization/order_dispatch.py`

**Summary:** Branchy dispatch hard to follow and test.

**Fix hint:** Table-driven dispatch or smaller functions with explicit strategy objects.

### [Q7] Weak typing in `orchestrator` / `layout_sequence`

**Category:** Code quality  
**Location:** `core/optimization/orchestrator.py`, `viz_modules/layout_sequence.py`

**Summary:** Ambiguous `Any`-like or missing annotations on public flows.

**Fix hint:** Typed dicts/dataclasses for phase results; mypy/ruff tightening on touched files.

### [Q8] Bare `except` / pass patterns

**Category:** Code quality  
**Location:** Scattered try/except in optimization and layout paths

**Summary:** Swallowed errors hide production failures.

**Fix hint:** Catch specific exceptions; log and re-raise or return structured errors.

### [Q9] Parse fallbacks to `0.0`

**Category:** Code quality  
**Location:** Numeric parsing in config/layout inputs

**Summary:** Invalid input silently becomes zero—bad layouts or silent data loss.

**Fix hint:** Validate and fail fast or surface warnings; distinguish “missing” vs. “invalid.”

---

## Low Priority / Suggestions

### [A11] Wide `__all__` in `_implementation`

**Category:** Architecture  
**Location:** `core/optimization/_implementation.py`

**Summary:** Exposes a large surface from an internal module.

**Fix hint:** Trim `__all__` to supported API; re-export from `__init__.py` only intentionally.

### [A12] Agent/log noise in `layout_sequence`

**Category:** Architecture / DX  
**Location:** `viz_modules/layout_sequence.py`

**Summary:** Verbose logging hampers local debugging and CI output.

**Fix hint:** Gate noisy blocks behind DEBUG; reduce default verbosity.

### [S4] Print disclosure

**Category:** Security (informational)  
**Location:** Various `print` calls in optimization/viz

**Summary:** Stdout may leak internal structure in shared consoles.

**Fix hint:** Prefer logging with levels; no prints in library code paths.

### [S5] SQLite churn loop

**Category:** Security / reliability (DoS-ish locally)  
**Location:** Tight loops opening/closing SQLite

**Summary:** Excessive connect/disconnect or small transactions risk locking and slow runs.

**Fix hint:** Connection context manager per batch; WAL pragmas if appropriate.

### [S6] No auth in library layer

**Category:** Security (acceptance)  
**Location:** `core/optimization/` as a library

**Summary:** Library correctly has no HTTP auth—but callers must enforce boundaries.

**Fix hint:** Document trust model; never expose raw optimization entrypoints without service-layer auth.

### [Q10] Dynamic `__import__` for `json`

**Category:** Code quality  
**Location:** Minor lazy-import pattern in optimization or viz

**Summary:** Unusual pattern hurts static analysis readability.

**Fix hint:** Standard top-level `import json` unless cycle forces otherwise.

### [Q11] Doc mismatch (separator / format)

**Category:** Code quality  
**Location:** Comments or docs describing separators in layout/export code

**Summary:** Docs say one delimiter; code uses another.

**Fix hint:** One source of truth constant; update docs.

### [Q12] Test gaps for 2D modules

**Category:** Code quality  
**Location:** `core/optimization/optimize_2d/*`

**Summary:** Critical finalize/prep paths under-tested.

**Fix hint:** Golden-file or property tests on small instances.

### [Q13] Noisy comments

**Category:** Code quality  
**Location:** Various touched files

**Summary:** Outdated or shouty comments distract from real invariants.

**Fix hint:** Delete stale comments; keep “why” only where non-obvious.

---

## Priority Matrix (top issues)

| ID  | Issue | Severity | Effort (indicative) | Priority |
|-----|--------|----------|---------------------|----------|
| A1  | TLS/global `OPT_*` isolation | Critical | Medium | **P0 — now** |
| S1  | Cross-request state (ContextVar / thread-local) | High | Medium | **P0 — now** |
| A2  | `config_and_data` monolith | High | High | **P1 — this sprint** |
| A3  | `layout_sequence` god module | High | High | **P1 — this sprint** |
| A4  | `ilp_model` DIP / price DB | High | Medium | **P1 — this sprint** |
| A5  | Ungated debug I/O (`layout_sequence`) | High | Low–Med | **P1 — this sprint** |
| Q1–Q3 | Megafunctions (layout build, finalize, prep_solve) | High | Medium | **P1 — this sprint** |

---

## Next Steps

1. **Immediate (P0):** Resolve **A1** and harden **S1**—prove isolation with concurrent/async tests; review every `OPT_*` read/write path and context manager usage.
2. **This sprint (P1):** Split **`config_and_data`** (**A2**) and start **`layout_sequence`** modularization (**A3**); inject pricing/config into **ILP** (**A4**); unify debug gates (**A5**, **S2**).
3. **Next sprint:** Address **medium** architecture items (**A6–A10**), security hygiene (**S3**), and quality items **Q4–Q9** (validation docs, dispatch refactor, typing, error handling, parse validation).
4. **Backlog:** **Low** items **A11–A12**, **S4–S6**, **Q10–Q13** as ongoing cleanup during feature touch.

**Remediation routing:** Use `/refactor` for structural/DRY work (A2, A3, A8, megafunctions); use `/implement` or planner/worker for behavioral security and validation (S1, S2, S3, Q8, Q9).

---

## Remediation Applied

**Date:** 2026-05-07

### Fixed (Critical A1 — partial / core API)

- **[A1]** Публичный вход `optimize_with_cascading_longitudinal_cuts` обёрнут в `optimization_context_scope()`, чтобы каждый вызов получал новое состояние `OPT_*` (ContextVar / согласованная изоляция с TLS). Файл: `core/optimization/orchestrator.py`.
- Добавлен regression-тест: `tests/test_optimization_context_and_snapshot.py` — `test_optimize_entrypoint_opens_fresh_context_each_call` (два вызова подряд создают два разных state).

### Fixed (S1 — plate mutable runtime)

- **TLS по умолчанию:** ленивая инициализация `get_plate_mutable_runtime()` использует `new_plate_mutable_runtime_empty()` вместо демо-заказа (`factory_demo_order`), чтобы «чужой» демо-набор не выглядел как данные пользователя на новом потоке.
- **`fresh_plate_mutable_request_scope()`** в `core/plate_runtime_state.py` — обёртка для HTTP/бота.
- **FastAPI:** `PlateMutableRuntimeIsolationMiddleware` в `app/middleware/plate_runtime_isolation.py`, подключение в `app/main.py` (после CORS, внешний слой на запрос).
- **Aiogram:** `PlateMutableRuntimeIsolationMiddleware` в `bot/middleware/plate_runtime_isolation.py`, регистрация `dp.update.middleware(...)` в `bot/handlers/__init__.py`.
- **Тесты:** `tests/test_plate_mutable_runtime_isolation.py` (пустой TLS в новом потоке, asyncio-изоляция, парсинг внутри `fresh_plate_mutable_request_scope`). Обновлены ожидания в `tests/test_plates_preview_xlsx.py` (типы ячеек openpyxl, марка ширины `-3-`, пустая колонка A на второй строке блока).
- Демо-данные по-прежнему: `new_plate_mutable_runtime_from_demo()`.

### Остаётся (вне этого прохода)

- Скрипты/CLI без middleware: при работе с `cfg.PLATES_*` вручную оборачивать в `fresh_plate_mutable_request_scope()` или привязывать свой `PlateMutableRuntime`.
- **S2–S6**, **A2–A5**, прочие пункты отчёта — по бэклогу.

---

*End of consolidated audit report.*
