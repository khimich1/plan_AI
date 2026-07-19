# Secure session cookies and APP_SECRET_KEY validation (A2-001)

**Status:** Implemented  
**Date:** 2026-06-03  
**Orchestration:** `orch-2026-06-03-arch-triage`  
**Plan:** [2026-06-03-architecture-triage-a1-a2-a3.md](../plans/2026-06-03-architecture-triage-a1-a2-a3.md) (task A2-001)  
**Audit finding:** [A2] in [2026-06-03-full-project-audit.md](../audits/2026-06-03-full-project-audit.md)

---

## Summary

Hardened stateless HMAC session cookies: mandatory strong `APP_SECRET_KEY` at startup, centralized cookie policy (`httponly`, `samesite`, `secure`, `max_age`), and the same helpers for API (`/api/v1/auth/*`) and web (`/web/login`, `/web/logout`) login/logout. The app no longer hardcodes `secure=False` on API cookies.

---

## What was implemented

### APP_SECRET_KEY validation (fail-fast)

Settings load in `core/config/settings.py` rejects weak or missing secrets before the app serves traffic:

| Rule | Detail |
|------|--------|
| Required | Must come from environment (no safe default in code) |
| Minimum length | 32 characters (`APP_SECRET_KEY_MIN_LENGTH`) |
| Forbidden values | `""`, `changeme`, `secret`, `change-this-secret-key-in-env`, and similar placeholders |
| Normalization | Leading/trailing whitespace is stripped |

On failure, Pydantic raises `ValidationError` with a message that includes a generator hint:

```bash
python -c "import secrets; print(secrets.token_urlsafe(48))"
```

`get_settings()` is cached; tests and local reloads should call `get_settings.cache_clear()` after env changes.

### Centralized cookie policy

`app/security/session.py` exposes:

| Function | Role |
|----------|------|
| `session_cookie_policy()` | Returns `httponly`, `samesite`, `secure`, `max_age` from settings |
| `set_session_cookie(response, token)` | `Set-Cookie` for `app_session` using the policy |
| `clear_session_cookie(response)` | `delete_cookie` with **matching** `secure` / `samesite` / `httponly` (required for browsers to clear the cookie) |

Cookie name: `app_session` (`SESSION_COOKIE_NAME`).

Session tokens remain **stateless HMAC-signed JSON** (`create_session_token` / `decode_session_token`): payload + HMAC-SHA256 signature, TTL from `SESSION_COOKIE_MAX_AGE` (default 12 hours).

### Unified API and web cookies

Both entry points use the same helpers (no divergent cookie attributes):

| Surface | Login | Logout |
|---------|-------|--------|
| REST API | `POST /api/v1/auth/login` → `set_session_cookie` | `POST /api/v1/auth/logout` → `clear_session_cookie` |
| Web UI | `POST /web/login` → `set_session_cookie` | `GET /web/logout` → `clear_session_cookie` |

Auth dependency `get_current_user` still reads `request.cookies["app_session"]` and validates via `decode_session_token`.

### `cookie_secure_enabled` logic

Computed property on `Settings`:

1. If `COOKIE_SECURE` is set in env → use that boolean explicitly.
2. Otherwise → `secure=True` when `APP_ENV` is `production` (case-insensitive), else `False`.

This removes hardcoded `secure=False` on API login and aligns web/API behavior.

---

## Environment variables

Documented in [.env.example](../../../.env.example) at project root.

| Variable | Required | Default | Purpose |
|----------|----------|---------|---------|
| `APP_SECRET_KEY` | **Yes** | — | HMAC signing key for session cookies; min 32 chars, no known placeholders |
| `APP_ENV` | No | `development` | When `production` and `COOKIE_SECURE` unset → cookies get `Secure` |
| `COOKIE_SECURE` | No | inferred | Override `Secure` flag (`true` / `false`) |
| `COOKIE_SAMESITE` | No | `lax` | `lax`, `strict`, or `none` (use `none` only with HTTPS + `Secure`) |
| `SESSION_COOKIE_MAX_AGE` | No | `43200` (12h) | Cookie `Max-Age` in seconds; bounds 60 … 30 days |

Example local `.env`:

```env
APP_SECRET_KEY=<paste-output-of-secrets.token_urlsafe(48)>
APP_ENV=development
# COOKIE_SECURE=false   # optional override for local HTTP
# COOKIE_SAMESITE=lax
# SESSION_COOKIE_MAX_AGE=43200
```

`tests/conftest.py` sets a valid test key before collection so pytest does not depend on a developer `.env`.

---

## Production configuration

### Minimum checklist

1. **Generate a strong secret** (once per environment; store in secrets manager, not in git):

   ```bash
   python -c "import secrets; print(secrets.token_urlsafe(48))"
   ```

2. Set **`APP_ENV=production`** so `Secure` cookies are enabled automatically behind HTTPS.

3. Terminate TLS at the reverse proxy (nginx, Caddy, cloud load balancer). Browsers require `Secure` cookies only over HTTPS.

4. Align **CORS** with your frontend origin (`BACKEND_CORS_ALLOWED_ORIGINS`).

5. Optional explicit override if you need `Secure` in non-production (e.g. staging behind HTTPS):

   ```env
   APP_ENV=staging
   COOKIE_SECURE=true
   ```

6. For cross-site SPA on another origin (uncommon for this app): `COOKIE_SAMESITE=none` and **`COOKIE_SECURE=true`** (mandatory with `SameSite=None`).

### `APP_ENV` vs `APP_DEBUG`

| Setting | Effect on cookies |
|---------|-------------------|
| `APP_ENV=production` | `secure=True` if `COOKIE_SECURE` not set |
| `APP_ENV=development` | `secure=False` if `COOKIE_SECURE` not set |
| `COOKIE_SECURE=true/false` | Always wins over `APP_ENV` inference |

`app_debug` does not control cookie security; use `APP_ENV` and `COOKIE_SECURE`.

### Secret rotation (current behavior)

Rotating `APP_SECRET_KEY` **invalidates all existing sessions immediately** (stateless HMAC). Plan maintenance windows or accept forced re-login. See [Future work](#future-work).

---

## Future work

Not implemented in A2-001; documented for later security iterations:

| Direction | Why |
|-----------|-----|
| **Server-side sessions** (Redis/DB) | Revoke sessions per user; rotate signing material without logging everyone out at once |
| **JWT with key ids (`kid`)** | Support zero-downtime key rotation with multiple active verification keys |
| **Sticky sessions / shared storage** | Already noted in settings for multi-replica drafts; session store would be a separate concern |

In-code reference: `app/security/session.py` (comment above `create_session_token`).

---

## Files changed

| File | Change |
|------|--------|
| `core/config/settings.py` | `APP_SECRET_KEY` validation; `COOKIE_SECURE`, `COOKIE_SAMESITE`, `SESSION_COOKIE_MAX_AGE`; `cookie_secure_enabled` computed field |
| `app/core/settings.py` | Re-export only (no logic) |
| `app/security/session.py` | `session_cookie_policy`, `set_session_cookie`, `clear_session_cookie`; rotation note in comments |
| `app/api/v1/endpoints/auth.py` | Login/logout use shared cookie helpers |
| `app/web/router.py` | Web login/logout use shared cookie helpers |
| `.env.example` | Document session-related env vars |
| `tests/conftest.py` | Valid `APP_SECRET_KEY` before imports at collection |
| `tests/test_settings_app_secret_key.py` | Secret validation and `cookie_secure_enabled` matrix |
| `tests/test_app_session.py` | Cookie policy, set/clear helpers, API/web login logout `Set-Cookie` attributes |

Related (unchanged behavior, uses cookies): `app/dependencies/auth.py` — reads `app_session` cookie.

---

## Tests

```bash
pytest tests/test_settings_app_secret_key.py tests/test_app_session.py -q
```

Coverage includes: weak/missing secret rejection, policy from settings, `APP_ENV` → `secure` default, API and web login/logout `Set-Cookie` / clear cookie parity.

---

## Related documentation

- Plan task A2-001: [architecture triage plan](../plans/2026-06-03-architecture-triage-a1-a2-a3.md)
- Web guide (legacy `APP_SECRET_KEY` mention): [docs/web-interface-guide.md](../../../docs/web-interface-guide.md)
- Next orchestration tasks: A1 (request-scoped plate context), A3 (canonical `PlateOrder`)
