# Deploy contract (P6 WP7)

Operational constraints until shared infrastructure lands in P7 (Redis rate limiting, PostgreSQL).

## Uvicorn workers

Login and OCR upload rate limits use **in-process** counters. They are **not** shared across uvicorn/gunicorn workers.

**Requirement:** run with a single worker until `RATE_LIMIT_SHARED_STORE` is implemented:

```bash
uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 1
```

### Worker count env vars

Set one of these so startup can log a warning when misconfigured:

| Variable | Purpose |
|----------|---------|
| `UVICORN_WORKERS` | Preferred when using uvicorn CLI |
| `WEB_CONCURRENCY` | Common on PaaS (Heroku, Railway, etc.) |
| `APP_WORKERS` | Documented alias for custom process managers |

Example:

```env
UVICORN_WORKERS=1
```

If any of these is `> 1` without a shared rate-limit store, the app logs a **warning** at startup. `GET /health` (non-production) includes `rate_limiting` metadata for ops checks.

### Shared store (P7)

`RATE_LIMIT_SHARED_STORE=redis` is reserved for P7 and currently raises `NotImplementedError` at startup. Do not set until Redis support is released.

## External OCR

External image recognition (OpenAI GPT-4o) is **disabled by default**:

```env
OCR_EXTERNAL_ENABLED=false
```

When disabled, commercial endpoints that accept image uploads or AI plate instructions return **503** with a clear message. Text-only parsing (`POST /api/v1/commercial/parse`) remains available.

Enable explicitly for staging or admin-only environments:

```env
OCR_EXTERNAL_ENABLED=true
OPENAI_API_KEY=sk-...
```

## Related

- `app/security/login_rate_limit.py` — login/password-change limits
- `app/services/commercial_upload_validation.py` — OCR upload limits
- Spec: `ai_docs/specs/stabilizaciya-p6-architecture-2026-06-21.md` (WP7)
