# Explicit request-scoped plate order context — Phase 1 (A1-001)

**Status:** Phase 1 implemented (strangler / infrastructure)  
**Date:** 2026-06-03  
**Orchestration:** `orch-2026-06-03-arch-triage`  
**Plan:** [2026-06-03-architecture-triage-a1-a2-a3.md](../plans/2026-06-03-architecture-triage-a1-a2-a3.md) (task A1-001, steps 1–2 of scope)  
**Depends on:** [secure-session-cookies-a2-001.md](secure-session-cookies-a2-001.md) (recommended, not blocking)

---

## Summary

Phase 1 introduces a single **request- or update-scoped** container for mutable plate order state and optimizer (`OPT_*`) state. Each HTTP request and each Telegram update gets a fresh `PlateOrderContext`; middleware binds it to existing thread-local / ContextVar shims so legacy `get_plate_mutable_runtime()` and `core.optimization` proxies keep working without migrating every call site yet.

FastAPI routes can inject context via `Depends(get_plate_order_context)`. Background sync work in worker threads uses `run_in_order_context` so the same binding applies inside `asyncio.to_thread`.

**Out of scope for Phase 1 (later A1-001 / A1-002):** migrating business code off bare `get_plate_mutable_runtime()`, deprecating `apply_to_globals`, and enforcing context-only access in services.

---

## Problem

Plate lists and optimizer state historically lived in thread-local storage, ContextVars, and module-level aliases (`cfg.PLATES_1_2`, `OPT_PLAN`, etc.). Without a consistent bind at request/update boundaries—and especially inside thread pools—orders could leak across concurrent users or requests.

Phase 1 does not remove globals; it **standardizes where binding happens** and gives new code an explicit object to pass or inject.

---

## Architecture (Phase 1)

```mermaid
sequenceDiagram
    participant Client
    participant MW as PlateMutableRuntimeIsolationMiddleware
    participant Ctx as PlateOrderContext
    participant App as Handler / legacy code
    participant TL as plate_mutable_runtime_scope + optimization_context_scope

    Client->>MW: HTTP or Telegram update
    MW->>Ctx: fresh_empty()
    MW->>MW: store on request.state / data["plate_order_ctx"]
    MW->>TL: ctx.bound()
    TL->>App: get_plate_mutable_runtime() / OPT_* proxies
    App-->>Client: response
    Note over TL: scopes exit; no cross-request state
```

| Layer | Role |
|-------|------|
| `PlateOrderContext` | SSOT object: `plates` (`PlateMutableRuntime`) + `optimization` (`dict`) |
| `bound()` | Nests `plate_mutable_runtime_scope` + `optimization_context_scope` for current asyncio task / thread |
| HTTP middleware | One fresh context per request; `request.state.plate_order_ctx` |
| Bot middleware | One fresh context per update; `data["plate_order_ctx"]` |
| `get_plate_order_context` | FastAPI `Depends` → reads `request.state`; 500 if missing/invalid |
| `run_in_order_context` | Runs sync callable in `asyncio.to_thread` with `ctx.bound()` |

---

## `PlateOrderContext`

**Location:** `core/plate_order_context.py`

| Member | Description |
|--------|-------------|
| `plates: PlateMutableRuntime` | Mutable order lists and load details (same type as legacy runtime) |
| `optimization: dict[str, Any]` | Fresh OPT state from `new_optimization_context_state()` |
| `fresh_empty()` | Empty plates + new optimization dict (same baseline as middleware S1) |
| `bound()` | Context manager; yields `self` while legacy accessors point at this instance |

`fresh_empty()` intentionally does **not** load demo/sample orders—each request/update starts empty unless handler logic fills `ctx.plates`.

---

## Middleware

### FastAPI

**Location:** `app/middleware/plate_runtime_isolation.py`  
**Registered in:** `app/main.py` (after `CORSMiddleware`, outermost on incoming requests)

```python
ctx = PlateOrderContext.fresh_empty()
request.state.plate_order_ctx = ctx
with ctx.bound():
    return await call_next(request)
```

### Telegram bot

**Location:** `bot/middleware/plate_runtime_isolation.py`  
**Registered in:** `bot/handlers/__init__.py` → `dp.update.middleware(PlateMutableRuntimeIsolationMiddleware())`

```python
ctx = PlateOrderContext.fresh_empty()
data["plate_order_ctx"] = ctx
with ctx.bound():
    return await handler(event, data)
```

Bot handlers that need explicit access should read `data["plate_order_ctx"]` (typed as `PlateOrderContext`). FastAPI `Depends` is not available in aiogram handlers.

---

## Dependency injection (FastAPI)

**Location:** `app/dependencies/plate_context.py`

```python
from fastapi import Depends
from app.dependencies.plate_context import get_plate_order_context
from core.plate_order_context import PlateOrderContext

@router.get("/example")
def example(ctx: PlateOrderContext = Depends(get_plate_order_context)):
    ...
```

| Behavior | Detail |
|----------|--------|
| Success | Returns `request.state.plate_order_ctx` when it is a `PlateOrderContext` |
| Failure | `HTTP 500` with detail `"Plate order context not initialized"` if attribute missing or wrong type |

Middleware must be registered **before** routes that depend on this; otherwise `get_plate_order_context` raises.

---

## `run_in_order_context`

Use when async code calls **sync** functions via `asyncio.to_thread` that still read `get_plate_mutable_runtime()` or `OPT_*` proxies. The worker thread gets the same `ctx.bound()` as the request handler.

```python
from core.plate_order_context import PlateOrderContext, run_in_order_context

async def handler(ctx: PlateOrderContext = Depends(get_plate_order_context)):
    result = await run_in_order_context(ctx, heavy_sync_fn, arg1, kw=2)
    return result
```

Implementation: wraps `fn` in a closure that enters `with ctx.bound():` inside the thread pool worker.

---

## Strangler pattern (legacy compatibility)

Phase 1 keeps existing call sites working:

- `get_plate_mutable_runtime()` → resolves to `ctx.plates` while `bound()` is active
- `cfg.PLATES_1_2` and similar module aliases → still mutate the bound runtime
- `core.optimization.OPT_*` → backed by `ctx.optimization` while `optimization_context_scope` is active

New code should prefer **`PlateOrderContext` passed or injected** rather than new uses of `get_plate_mutable_runtime()` in services (enforced in A1-002+).

---

## Files (Phase 1)

| File | Change |
|------|--------|
| `core/plate_order_context.py` | `PlateOrderContext`, `bound()`, `run_in_order_context` |
| `app/middleware/plate_runtime_isolation.py` | HTTP isolation middleware |
| `app/dependencies/plate_context.py` | `get_plate_order_context` |
| `app/main.py` | Registers `PlateMutableRuntimeIsolationMiddleware` |
| `bot/middleware/plate_runtime_isolation.py` | Bot isolation middleware |
| `bot/handlers/__init__.py` | Registers bot middleware on `dp.update` |
| `tests/test_plate_order_context.py` | Unit + middleware + DI + `run_in_order_context` |

**Unchanged in Phase 1 (still used via `bound()`):** `core/plate_runtime_state.py`, `core/optimization/context.py`, `core/config_and_data.py`.

---

## Tests

```bash
pytest tests/test_plate_order_context.py -q
```

Coverage includes:

- `fresh_empty()` — empty plates, fresh optimization dict
- `bound()` — legacy runtime and `OPT_*` see context state; nested scopes restore outer
- `run_in_order_context` — thread worker sees correct plates and `opt_plan`
- `get_plate_order_context` — reads `request.state`; 500 on missing/invalid
- Middleware — consecutive requests get isolated fresh context (not shared demo data)

Existing isolation suites should remain green:

```bash
pytest tests/test_plate_mutable_runtime_isolation.py tests/test_optimization_context_and_snapshot.py tests/test_optimization_thread_local_globals.py -q
```

---

## Phase 2+ (A1-002 and beyond)

| Task | Focus | Doc |
|------|--------|-----|
| A1-001 (remainder) | Migrate services/handlers to explicit `ctx`; restrict `get_plate_mutable_runtime()` to shim layer | — |
| A1-002 | Mandatory middleware everywhere; deprecate `apply_to_globals`; `hydrate_from_order` + `run_in_order_context` migration | [plate-order-context-a1-002-middleware-deprecations.md](plate-order-context-a1-002-middleware-deprecations.md) |
| A3-* | Canonical `PlateOrder` type and duplicate model removal | plan A3-001 / A3-002 |

---

## Related documentation

- Plan task A1-001 / A1-002: [architecture triage plan](../plans/2026-06-03-architecture-triage-a1-a2-a3.md)
- Prior security work: [secure-session-cookies-a2-001.md](secure-session-cookies-a2-001.md)
