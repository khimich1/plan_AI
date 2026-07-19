# Report: Architecture Triage Implementation

**Date:** 2026-06-03  
**Orchestration:** `orch-2026-06-03-arch-triage`  
**Status:** ✅ Completed (5/5 tasks)  
**Plan:** [2026-06-03-architecture-triage-a1-a2-a3.md](../plans/2026-06-03-architecture-triage-a1-a2-a3.md)  
**Audit source:** [2026-06-03-full-project-audit.md](../audits/2026-06-03-full-project-audit.md)

---

## Summary

The architecture triage orchestration closed **five critical audit findings (A1, A2, A3, S1)** through a sequenced refactor: session hardening first, then explicit request-scoped plate order context, mandatory middleware, canonical `PlateOrder` type with boundary adapters, and full call-site migration off `apply_to_globals()`.

Cross-request plate/optimizer state leaks are mitigated via mandatory FastAPI and bot middleware plus `PlateOrderContext` hydration. Session cookies use a centralized policy with fail-fast `APP_SECRET_KEY` validation. The duplicate `PlateOrder` model problem is resolved with `core/domain/plate_order.py` as SSOT and a thin app subclass for commercial-only `nomenclature_cache`.

**Final test slice:** **136 passed**, 0 failed (full `tests/` suite at orchestration close).

---

## Task Outcomes

### A2-001 — Secure session cookies and APP_SECRET_KEY validation

**Status:** ✅ Completed  
**Review:** APPROVED  
**Audit findings addressed:** A2, S1, S3 (cookie `secure` policy)

**Outcome:**

- Mandatory `APP_SECRET_KEY` from environment with min length (32), forbidden placeholders, fail-fast at settings load.
- Centralized cookie helpers in `app/security/session.py`: `session_cookie_policy()`, `set_session_cookie()`, `clear_session_cookie()` with matching attributes on delete.
- API (`POST /api/v1/auth/login|logout`) and web (`POST /web/login`, `GET /web/logout`) use the same policy; no hardcoded `secure=False`.
- `cookie_secure_enabled` inferred from `APP_ENV=production` or overridden via `COOKIE_SECURE`.
- `.env.example` documents session-related variables; `tests/conftest.py` sets a valid test key before collection.

**Doc:** [secure-session-cookies-a2-001.md](../features/secure-session-cookies-a2-001.md)

---

### A1-001 — Explicit request-scoped order context (DI)

**Status:** ✅ Completed (Phase 1 infrastructure)  
**Audit finding addressed:** A1 (foundation)

**Outcome:**

- Introduced `PlateOrderContext` in `core/plate_order_context.py` as SSOT for mutable plates + optimization dict per request/update.
- `PlateMutableRuntimeIsolationMiddleware` on FastAPI (`request.state.plate_order_ctx`) and aiogram (`data["plate_order_ctx"]`).
- FastAPI DI via `Depends(get_plate_order_context)` in `app/dependencies/plate_context.py`.
- `run_in_order_context()` for sync legacy code in `asyncio.to_thread` workers with correct binding.
- Strangler pattern: legacy `get_plate_mutable_runtime()` and `OPT_*` proxies continue working inside `ctx.bound()`.

**Doc:** [plate-order-context-a1-001-phase-1.md](../features/plate-order-context-a1-001-phase-1.md)

---

### A1-002 — Mandatory middleware + deprecate apply_to_globals

**Status:** ✅ Completed  
**Audit finding addressed:** A1 (enforcement + migration)

**Outcome:**

- Mandatory isolation middleware on all HTTP requests and Telegram updates (enforcement model documented).
- `PlateOrder.apply_to_globals()` and `get_current_plate_order()` emit `DeprecationWarning` (stacklevel=2).
- `PlateOrderContext.hydrate_from_order()`, `load_optimization_snapshot()`, `load_production_snapshot()` replace global hydration.
- Migrated app services and bot handlers to explicit context hydration + `run_in_order_context` where needed.
- Bot DI helper in `bot/dependencies/plate_context.py`.

**Doc:** [plate-order-context-a1-002-middleware-deprecations.md](../features/plate-order-context-a1-002-middleware-deprecations.md)

---

### A3-001 — Canonical PlateOrder type and boundary adapters

**Status:** ✅ Completed  
**Audit finding addressed:** A3 (type unification)

**Outcome:**

- **Canonical SSOT:** `core/domain/plate_order.py` — single field definition, serialization, `from_orders_2d`, totals.
- **App subclass:** `app/domain/models/plate_order.py` extends core; adds `nomenclature_cache` only (no duplicated core fields).
- **Boundary adapters:** `app/domain/adapters/plate_order.py` — `to_core_order()` / `from_core_order()`.
- **Core coercion:** `coerce_core_plate_order()` strips app-only fields for hydration/optimization paths.
- Legacy re-export in `core/config_and_data.py` preserved for compatibility.

**Doc:** [plate-order-canonical-a3-001.md](../features/plate-order-canonical-a3-001.md)

---

### A3-002 — Migrate call sites and remove duplicate model

**Status:** ✅ Completed  
**Audit finding addressed:** A3 (call-site migration)

**Outcome:**

- Production and commercial flows no longer use `AppPlateOrder → core → apply_to_globals()` chains.
- All migrated paths use `to_core_order()` + `ctx.hydrate_from_order()` / `load_production_snapshot()` + `ctx.bound()` or `run_in_order_context()`.
- **`apply_to_globals()` grep:** only definition in `core/domain/plate_order.py` + intentional test in `tests/test_plate_order_context.py`.
- App `PlateOrder` retained as thin commercial/API domain type (not removed — by design).
- Deprecated shims (`apply_to_globals`, `get_current_plate_order`) kept for backward compatibility; safe to delete in a follow-up cleanup.

**Doc:** [plate-order-migration-a3-002.md](../features/plate-order-migration-a3-002.md)

---

## Files Changed (High Level)

| Area | Key files |
|------|-----------|
| **Settings / security** | `core/config/settings.py`, `app/security/session.py`, `app/api/v1/endpoints/auth.py`, `app/web/router.py`, `.env.example` |
| **Plate order context** | `core/plate_order_context.py`, `core/domain/plate_order.py` |
| **Middleware / DI** | `app/middleware/plate_runtime_isolation.py`, `bot/middleware/plate_runtime_isolation.py`, `app/dependencies/plate_context.py`, `bot/dependencies/plate_context.py` |
| **Domain / adapters** | `app/domain/models/plate_order.py`, `app/domain/adapters/plate_order.py`, `app/domain/adapters/__init__.py` |
| **App services** | `commercial_service.py`, `commercial_workflow_service.py`, `file_generation_service.py`, `day_documents_service.py`, `archive_service.py`, `production_planning_service.py` |
| **Bot handlers** | `commercial.py`, `production_day_view.py`, `production_execution.py`, `production_create.py`, `kp.py`, `optimize.py` |
| **Tests** | `tests/conftest.py`, `test_settings_app_secret_key.py`, `test_app_session.py`, `test_plate_order_context.py`, `test_plate_order_adapters.py`, `test_commercial_web_flow.py`, `test_archive_endpoints.py`, `test_admin_service.py` |

**Approximate scope:** ~35 Python modules touched across `app/`, `core/`, `bot/`, and `tests/`.

---

## Test Results

| Milestone | Result |
|-----------|--------|
| A2-001 slice | 51 passed (session + secret validation) |
| **Final orchestration slice** | **136 passed**, 0 failed |

**Key test modules added or extended:**

- `tests/test_settings_app_secret_key.py` — secret validation, `cookie_secure_enabled` matrix
- `tests/test_app_session.py` — cookie policy, API/web login/logout parity
- `tests/test_plate_order_context.py` — context, middleware, DI, hydration, deprecation warnings, snapshots
- `tests/test_plate_order_adapters.py` — `to_core_order` / `from_core_order` roundtrips

**Isolation suites kept green:** `test_plate_mutable_runtime_isolation.py`, `test_optimization_context_and_snapshot.py`, `test_optimization_thread_local_globals.py`.

---

## Documentation Created

| Document | Task |
|----------|------|
| [secure-session-cookies-a2-001.md](../features/secure-session-cookies-a2-001.md) | A2-001 |
| [plate-order-context-a1-001-phase-1.md](../features/plate-order-context-a1-001-phase-1.md) | A1-001 |
| [plate-order-context-a1-002-middleware-deprecations.md](../features/plate-order-context-a1-002-middleware-deprecations.md) | A1-002 |
| [plate-order-canonical-a3-001.md](../features/plate-order-canonical-a3-001.md) | A3-001 |
| [plate-order-migration-a3-002.md](../features/plate-order-migration-a3-002.md) | A3-002 |

**Also referenced:** orchestration plan, full-project audit report.

---

## Technical Decisions

| Decision | Choice | Rationale |
|----------|--------|-----------|
| Canonical `PlateOrder` | `core/domain/plate_order.py` | Tied to optimization/runtime; largest field set |
| App layer type | Subclass + adapters | Commercial `nomenclature_cache` without field duplication |
| Session store | Stateless HMAC cookie (hardened) | Minimize scope; server-side sessions / JWT rotation deferred |
| Context delivery | Middleware + `Depends` / bot `data` injection | Matches existing `plate_mutable_runtime_scope` API |
| Migration strategy | Strangler + deprecate, then migrate call sites | Avoid big-bang; keep isolation tests green |
| `apply_to_globals` removal | Deferred shim delete | Zero production callers; tests still assert warning behavior |

---

## Remaining Follow-ups

### P0 (still open — not in this orchestration)

| ID | Item | Notes |
|----|------|-------|
| **S2** | Telegram bot without user authorization | Whitelist `user_id`, roles, audit — **highest remaining security risk** |

### P1 — recommended next sprint

| ID | Item | Notes |
|----|------|-------|
| **S4** | Rate limiting on login | IP/username limits; Redis or slowapi |
| **S5** | CSRF protection | Cookie auth + `allow_credentials=True` in CORS |
| **S6** | Admin destructive ops without step-up | `/db/reset/*` endpoints |
| **S7** | Permissive CORS | Narrow origins, headers, methods |
| **A4** | God module `core/kp_db.py` (~4200+ lines) | Split into repositories |
| **A5** | God module `app/planning/plan_manager.py` | Extract PlanRepository, GanttService |
| **A6** | `list_users()` on every auth request | `get_user_by_id` with request-scoped cache |
| **A7** | No DI for services in endpoints | Factories in `app/dependencies/` |
| **A8** | Sync CPU-heavy work in async API | `run_in_executor` / worker queue |
| **A9** | Core → viz_modules dependency inversion | Ports in `core/ports/` |
| **Q1** | Agent debug NDJSON in production code | `core/kp_db.py`, bot handlers |
| **Q4** | No bot handler tests | `tests/test_bot_*` |
| **Q5** | Frontend test gap | Vitest + RTL for auth, production, archive |
| **Q6** | `day_documents_service` untested | Unit tests for schema/statement generation |

### Orchestration cleanup (low effort)

- Remove deprecated `apply_to_globals()` / `get_current_plate_order()` shims from `core/domain/plate_order.py` once remaining tests/docs updated.
- Migrate `tests/test_procurement_loads.py` off `get_current_plate_order()`.
- Optional: replace `AppPlateOrder` imports with core type in parser/draft-only modules (no global side effect today).
- Production deploy checklist: set strong `APP_SECRET_KEY`, `APP_ENV=production`, TLS at reverse proxy; plan secret rotation (invalidates all sessions).

---

## Metrics

| Metric | Value |
|--------|-------|
| Tasks completed | 5 / 5 |
| Critical audit items addressed | A1, A2, A3, S1 (+ S3 cookie secure) |
| Critical audit items remaining | S2 (bot auth) |
| Feature docs created | 5 |
| Final test slice | 136 passed |
| Production `apply_to_globals()` callers | 0 |

---

## Related Documentation

- **Plan:** [2026-06-03-architecture-triage-a1-a2-a3.md](../plans/2026-06-03-architecture-triage-a1-a2-a3.md)
- **Audit:** [2026-06-03-full-project-audit.md](../audits/2026-06-03-full-project-audit.md)
- **Workspace:** `.cursor/workspace/active/orch-2026-06-03-arch-triage/` (plan, tasks, progress — ready for archival per workspace cleanup policy)

---

## Next Steps

1. **Before production deploy:** implement bot user authorization (S2); verify `.env` has production-grade `APP_SECRET_KEY`.
2. **This sprint (P1):** login rate limits, CSRF, admin step-up, CORS tightening, `get_user_by_id` auth lookup.
3. **Structural (P1):** begin `kp_db.py` decomposition; add bot handler and day-documents test coverage.
4. **Cleanup:** delete deprecated plate-order global shims after final test migration.
