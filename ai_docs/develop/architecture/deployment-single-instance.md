# Deployment: single-instance / single-worker

**Status:** accepted  
**Date:** 2026-08-02  
**Updated:** 2026-08-03 (A2 hard-fail)  
**Related:** [`rate-limiting.md`](./rate-limiting.md), [`../deploy-contract.md`](../deploy-contract.md)  
**Audit:** A2 / S2 (2026-08-02), A2 enforcement (2026-08-03)

## Context

The web app uses:

- **SQLite** (`plita.db` / `pb.db`) for KP, production, logistics
- **File drafts** under `DRAFTS_DIR` for commercial wizard state
- **In-process** login/OCR rate-limit counters (`threading.Lock` + memory)

None of these are safe to share across multiple OS processes without a shared store (Redis / PostgreSQL) that is **out of scope** until P7.

## Decision

Run production as **one application process** with **uvicorn `--workers 1`** until a shared rate-limit store and multi-writer storage land.

| Constraint | Value |
|------------|-------|
| Uvicorn workers | `1` |
| Preferred env | `UVICORN_WORKERS=1` |
| Horizontal replicas | Not supported (sticky single instance) |

Compose / Docker CMD must pass `--workers 1` explicitly and set `UVICORN_WORKERS=1` so startup warning and `/health` metadata stay honest.

### Startup enforcement (A2)

When **all** of the following hold, lifespan **refuses to start** (`RuntimeError`):

1. `APP_ENV=production`
2. `APP_STORAGE_LAYOUT=single_instance`
3. `configured_workers > 1` (from `UVICORN_WORKERS` / `WEB_CONCURRENCY` / `APP_WORKERS`)

If `configured_workers` is **undeclared** (`None`), startup does **not** hard-fail; a WARNING is logged and `/health` `rate_limiting` metadata includes `workers_undeclared: true`.

Non-production environments with `workers > 1` still only warn (no hard-fail).

## Consequences

- Rate limits match configured ceilings (not `N × limit`).
- SQLite writers stay in one process; drafts are not split across hosts.
- Misconfigured production multi-worker deploys fail fast instead of silently splitting limits/data.
- Scaling out requires Redis (rate limits) + shared DB/volume strategy — tracked for P7, not this ADR.

## Non-goals

- Hard-fail when worker env vars are unset (undeclared remains warning + health metadata only).
- Hard-fail in non-production when `workers > 1`.
- Implementing Redis rate limiting in this change set.
