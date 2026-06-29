# Rate limiting — deployment constraints

**Audit:** S3 (2026-06-20) · **Spec:** P1-next WP4  
**Status:** in-process limiters; single-worker deployment constraint documented

## Scope

| Limiter | Module | Key |
|---------|--------|-----|
| Login (API + web) | `app/security/login_rate_limit.py` | Client IP |
| Commercial OCR uploads | `app/services/commercial_upload_validation.py` | User ID |

Both use an **in-process sliding window** (`threading.Lock` + in-memory dict). Each OS process maintains its own counters.

## Multi-instance behaviour

With `N` uvicorn/gunicorn workers:

- Limits are **not** shared — effective allowance scales roughly with `N`.
- Worker restarts clear counters for that process.
- Horizontal scaling (multiple hosts) multiplies the same issue.

This is acceptable when the app runs as a **single worker** (current production recommendation).

## Production deployment

### Option A — single worker (recommended)

```bash
uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 1
```

Local dev scripts (`run_local.sh`, `run+logs.sh`) already use a single process with `--reload`.

Set for observability (startup warning + health metadata):

```bash
export UVICORN_WORKERS=1
# or
export WEB_CONCURRENCY=1
```

### Option B — sticky sessions (weaker)

If you must run `--workers > 1` without Redis:

- Configure the load balancer for **session affinity** (typically by client IP).
- Accept that limits are still per-worker and reset on deploy/restart.
- Set `UVICORN_WORKERS` / `WEB_CONCURRENCY` so startup logs the misconfiguration warning.

### Option C — shared store (not implemented)

`REDIS_URL` exists in settings but is **not** wired to rate limiters. A future shared store would replace `_SlidingWindowRateLimiter` and `_CommercialOcrUploadLimiter` backends.

## Health check

`GET /health` and `GET /api/v1/health` include a `rate_limiting` object:

```json
{
  "store": "in-process",
  "shared_across_workers": false,
  "configured_workers": 1,
  "single_worker_required": true,
  "deployment_note": "Rate limits are in-process only. Use uvicorn --workers 1 or sticky sessions."
}
```

If `configured_workers > 1`, a `warning` field is added. Ops can alert on that field.

## Startup warning

When `UVICORN_WORKERS` or `WEB_CONCURRENCY` is set to a value `> 1`, the app logs a WARNING during lifespan startup (after logging is configured).

## Related settings

| Variable | Default | Purpose |
|----------|---------|---------|
| `AUTH_LOGIN_ATTEMPTS_PER_MINUTE` | (see settings) | Login attempts per IP per window |
| `AUTH_LOGIN_RATE_LIMIT_WINDOW_SECONDS` | 60 | Login sliding window |
| `COMMERCIAL_OCR_UPLOADS_PER_HOUR` | (see settings) | OCR uploads per user per hour |
