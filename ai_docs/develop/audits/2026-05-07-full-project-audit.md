# Full project audit — Shishov monorepo

**Date:** 2026-05-07  
**Scope:** full project (`app/`, `bot/`, `core/`, critical paths)  
**Audited by:** senior-reviewer + security-auditor + reviewer (Composer 2)

---

## 1. Executive Summary

This report consolidates architecture, security, and code-quality findings. Counts aggregate **unique** findings by primary category and severity.

**Overall Health Score:** **0.0 / 10**

**Scoring formula (capped deductions):** start at **10**; subtract **min(critical × 2, 6)**; subtract **min(high × 0.5, 3)**; subtract **min(medium × 0.1, 1)**; round to one decimal.  
With **5 critical**, **15 high**, **19 medium**:  
`10 − min(10, 6) − min(7.5, 3) − min(1.9, 1) = 10 − 6 − 3 − 1 = 0.0`.

**Aggregated counts**

| Severity | Architecture | Security | Code Quality | Total |
|----------|---------------|----------|--------------|-------|
| Critical | 4 | 1 | 0 | **5** |
| High | 7 | 4 | 4 | **15** |
| Medium | 5 | 8 | 6 | **19** |
| Low | 4 | 4 | 4 | **12** |

**Recommendation:** Address **5 critical** issues before the next release; the health score reflects **capped** deductions—remaining high and medium items still materially affect scalability, availability, and abuse resistance.

---

## 2. Critical Issues

### Architecture

- **[A1] Inverted dependency (shared → app):** `core/db_config.py` imports `app.core.settings.get_settings`, so the supposedly shared `core/` layer depends on the FastAPI app package. That breaks inward-only dependency flow, makes `core` non-reusable, and risks import/bootstrap cycles if `app` ever pulls modules that load `core.db_config` early.

- **[A2] Global optimization state + concurrency risk:** `core.optimization` keeps mutable module-level plans/maps; `app/services/optimization_service.py` `legacy_runtime` swaps those globals around `yield`. Any overlapping requests (multiple workers / threads) can clobber each other's optimization results—scales poorly and violates safe service design.

- **[A3] Session security defaults:** `app/api/v1/endpoints/auth.py` sets the session cookie with `secure=False` unconditionally; `app/core/settings.py` ships a default `APP_SECRET_KEY` string. On HTTPS production, cookies are exposed to downgrade/MITM-adjacent risks, and a forgotten default secret invalidates the HMAC model in `app/security/session.py`.

- **[A4] Auth hot path loads entire user table:** `app/dependencies/auth.py` `get_current_user` calls `repository.list_users()` and builds a full `{id: user}` map on **every** authenticated request. This is O(n) per request, painful at scale, and loads more identity data than needed (should be `get_user_by_id` only).

### Security

- **[S1] Hardcoded default signing secret for sessions** — `APP_SECRET_KEY` falls back to the fixed string `change-this-secret-key-in-env` in `app/core/settings.py` (lines 32–35). Anyone who can reach the app with that default (or who knows a leaked/stale deploy) can forge valid `app_session` HMAC tokens (`app/security/session.py` uses this key at lines 14–16). **Must** require a strong random secret in production and fail fast if unset or equal to the default.

### Code Quality

- *(No critical code-quality findings.)*

---

## 3. High Priority (consolidated)

Summarized by theme; each item is **High** in its primary track. **Location → impact → fix hint.**

### Architecture (A5–A11)

| ID | Location / pattern | Impact | Fix hint |
|----|--------------------|--------|----------|
| **A5** | `core/kp_db.py` — very large single module | Hard to test, isolate regressions, and evolve schema/CRUD vs workflows | Split by bounded context (schema, CRUD, workflows, diagnostics); narrow public API |
| **A6** | `app/planning/plan_manager.py` | JSON paths, calendar rules, occupancy, tracks, IDs tightly coupled; blocks swapping storage | Introduce a port (interface) + adapter; isolate file layout behind one module |
| **A7** | `core/config_and_data.PlateOrder` vs `app/domain/models/plate_order.py`; bot imports both | Field drift, subtle mapping bugs | Single canonical model + explicit adapters for legacy |
| **A8** | `app/web/router.py` (~1000+ lines) + parallel REST | Double surface for flows; larger security/review perimeter | Push behavior to REST/services; slim HTML shells |
| **A9** | Sync `sqlite3` + sync `def` routes in `app/repositories/*.py`, e.g. `auth_repository`, production endpoints | Event-loop blocking under concurrency; scalability/DoS-ish tail latency | `async`/aiosqlite or dedicated threadpool; timeouts |
| **A10** | `bot/handlers/commercial.py`, `bot/handlers/production_completion.py` — 1000+ lines | Presentation/business mixed; brittle tests | Extract services, slim handlers to orchestration |
| **A11** | `app/repositories/plan_repository.py` delegates only to `plan_manager` | “Repository” does not abstract storage | Real repository boundary or rename to facade |

### Security (S2–S5)

| ID | Risk | Impact | Fix hint |
|----|-----|--------|----------|
| **S2** | Session cookies `secure=False` (`app/api/v1/endpoints/auth.py` ~26–32); HTML login in `app/web/router.py` ~203–205 | Cookie exposure on insecure/mixed transport | `Secure=True` in prod + `SameSite`; HTTPS-only assumption |
| **S3** | No rate limiting on `POST /api/v1/auth/login` and `/web/login` | Online guessing, abusive load vs PBKDF2 | Middleware / Redis-backed limiter, lockout backoff |
| **S4** | Sync SQLite on async stack (`AuthRepository`, `list_users`), called from async paths | Availability / blocked workers | Same as architectural async story + avoid full-table reads per request |
| **S5** | `_CommercialOcrUploadLimiter` in `app/services/commercial_upload_validation.py` (18–37) — in-memory | Per-process limit multiplies across workers | Shared store (Redis) or centralized gateway limits |

### Code Quality — High (Q1–Q4)

| ID | Issue | Impact | Fix hint |
|----|-------|--------|----------|
| **Q1** | Bare `except:` in multiple bot handlers swallows all exceptions | Failures hidden, debugging harder | Narrow exceptions, log and re-raise; see `bot/handlers/archive.py`, `production_execution.py`, `production_completion.py`, etc. |
| **Q2** | Oversized modules: `core/kp_db.py`, `app/web/router.py`, `app/planning/plan_manager.py` | Hard reviews/refactors | Split by responsibility; align with **A5**, **A6**, **A8** |
| **Q3** | Two parallel `PlateOrder` implementations (`core/config_and_data` vs `app/domain/models`) | Domain drift, mapping bugs | Single canonical model + explicit adapters (**A7**) |
| **Q4** | `plan_manager.py`: bare `except:` on parse/coerce; duplicated merge logic for `plate_lookup_by_length` | Bad input masked; duplicate bug surface | Replace bare except; extract shared merge helper |

---

## 4. Medium / Low (by pillar)

Per audit-workflow template, this section lists **only** Medium and Lower for each pillar. Critical and High items appear in **§2** and **§3**. The full architecture register (all severities **A1–A20**) is preserved verbatim in **Appendix A**; security Critical/High (**S1–S5**) in **Appendix B**; complete code-quality ladder (**Q1–Q14**) in **Appendix C**.

### 4.1 Architecture (medium / low only)

#### Medium

- **[A12]** **Non-structured diagnostics in optimizer entrypoint:** `core/optimization/orchestrator.py` uses `print()` for mode logging instead of the project logger (`core/logging_config.py` / `logging`), which complicates production observability and log aggregation.

- **[A13]** **Multiple sources of truth for DB paths:** e.g. `core/kp_db.py` defines `DEFAULT_DB` from `__file__` layout while `app/core/settings.py` and `core/db_config.py` use settings; `bot/bot_main.py` logs a hard-coded `plita.db` path separately from `bot/bot_config.DB_PATH_STR` (`pb.db`). Operational confusion and misconfiguration risk.

- **[A14]** **Unused configuration surface:** `app/core/settings.py` defines `database_url` and `redis_url` with no other references in the codebase—noise for operators and a hint of abandoned or premature infrastructure planning.

- **[A15]** **Heavy orchestration services:** `app/services/production_planning_service.py` and `app/services/day_documents_service.py` pull in many `core.*` concerns (optimization, visualization, serialization, commit, debug paths). They behave as application monoliths; further decomposition would clarify use-cases.

- **[A16]** **Legacy coupling via global `core.optimization`:** Aside from concurrency, the pattern in `app/services/optimization_service.py` (mutating `legacy_optimization` module attributes) leaks implementation detail and complicates parallel tests and future non-global APIs.

#### Low

- **[A17]** **Ad-hoc service construction in routes:** Endpoints often do `ProductionService()` / `CommercialService()` / `CommercialWorkflowService()` inline instead of FastAPI `Depends` providers (`app/api/v1/endpoints/production.py`, `app/api/v1/endpoints/commercial.py`)—works, but weakens a single composition root and consistent overrides in tests.

- **[A18]** **Star-import shim:** `bot/handlers/plan_manager.py` re-exports `app.planning.plan_manager` with `from app.planning.plan_manager import *`, obscuring which symbols handlers rely on and hindering refactors.

- **[A19]** **Import-time path mutation:** `app/planning/plan_manager.py` mutates `sys.path` when imported—side effect that can surprise tooling, duplicate path entries, and order-dependent behavior.

- **[A20]** **Embedded debug log constants in `kp_db`:** `core/kp_db.py` wires multiple `_DEBUG_*` paths at module level—ties business data access to developer debug artifacts and clutters production modules.

---

### 4.2 Security (medium / low only)

#### Medium

- **[S6]** **`get_current_user` loads the full user roster every request** — `repository.list_users()` builds a dict of all users on each authenticated call (`app/dependencies/auth.py`, lines 21–24). Besides cost, it pulls every account’s metadata into memory (usernames, roles, `manager_id`, `is_active`) for **every** API call using this dependency—unnecessary exposure surface and timing/footprint vs. a single-row lookup by `id`.

- **[S7]** **State-changing HTML forms without CSRF tokens** — `POST /web/login`, `POST /web/offers/new`, `POST /web/offers/drafts/...` (`app/web/router.py`, e.g. 197–205, 867–927, 947–978) rely on cookie auth with `samesite="lax"` but no synchronizer token. Defense is mostly browser SameSite behavior; **defense-in-depth** (CSRF token or strict `SameSite=Strict` where compatible) is missing for legacy form flows.

- **[S8]** **Permissive CORS headers** — `app/main.py` (lines 41–46) sets `allow_headers=["*"]` with `allow_credentials=True`. Combined with a mis-set `BACKEND_CORS_ALLOWED_ORIGINS`, this increases impact of origin misconfiguration (preflight allows any header from allowed origins).

- **[S9]** **`/api/v1/health` leaks deployment context** — `app/api/v1/endpoints/health.py` (lines 10–17) returns `environment` from `settings.app_env`, aiding attackers in choosing exploits or fingerprinting deployments.

- **[S10]** **Debug flag on the ASGI app** — `FastAPI(debug=settings.app_debug)` in `app/main.py` (line 38). If `APP_DEBUG` / `app_debug` is true in production, expect verbose errors and Swagger UI exposure beyond what hardened APIs should show.

- **[S11]** **Duplicated surface: REST + large `app/web/router.py`** — Same business actions exist over `/api/v1/...` and `/web/...` (e.g. commercial flows). More routes and mixed paradigms increase the chance one path gets weaker checks, headers, or future drift (OWASP expanded attack surface / configuration errors).

- **[S12]** **Plan and draft data on local paths** — Defaults under `bot/data/` (`plans`, `current_plan.json`, `plans_metadata.json` in `app/core/settings.py`, lines 57–59) and `.app_data/drafts/` (line 62) hold operational JSON. **Integrity and confidentiality** depend on OS permissions and backups; any local write access (or confused-deputy in another component) can tamper with production state—amplified by dual order/plan representations called out in architecture review.

- **[S13]** **`core/db_config.py` imports `app.core.settings`** (`core/db_config.py`, lines 7–12) — Low-level DB path resolution is tied to the full app settings stack. A packaging/misconfig bug or unexpected import order can widen blast radius or make “core-only” tooling accidentally load web config/secrets.

#### Low

- **[S14]** **Logout cookie clearing may be incomplete** — `logout` in `app/api/v1/endpoints/auth.py` (lines 37–40) calls `delete_cookie("app_session")` without matching `path` / `secure` / `samesite` used on `set_cookie`; some clients can retain cookies until expiry depending on how the cookie was set (esp. if `path` diverges later).

- **[S15]** **No standard security headers** — `app/main.py` does not add CSP, `X-Frame-Options`, `Referrer-Policy`, etc., increasing baseline XSS/clickjacking risk for HTML-served pages (`app/web/router.py` server-rendered HTML and `/commercial-offer` shell).

- **[S16]** **Bot token handling** — `BOT_TOKEN` comes from settings (`bot/bot_config.py`, lines 1–8; `app/core/settings.py`, line 42). No hardcoded production token in code paths reviewed; residual risk is **operational** (token in env files, logs, or screen output from `run_bot.py` / `bot/bot_main.py` messages)—ensure logging never prints the token.

- **[S17]** **Verbose exception logging in some `core` modules** — e.g. `traceback.print_exc()` patterns in `core/ocr_gpt.py`, `core/commercial_offer.py`, `core/kp_db.py` (among others) can write paths and internals to stdout/logs; keep log destinations restricted and avoid shipping debug logs to untrusted parties.

---

### 4.3 Code Quality (medium / low only)

#### Medium (Q5–Q10)

- **[Q5]** **DI / testability** — Routes construct **new service instances per call** (`ProductionService()` throughout `app/api/v1/endpoints/production.py`, `OffersService()` in every handler in `app/api/v1/endpoints/offers.py`) instead of FastAPI `Depends()` providers — harder to swap fakes at the HTTP boundary without patching classes.

- **[Q6]** **`_user: dict` on protected endpoints** (`app/api/v1/endpoints/offers.py`, `production.py`, `admin.py`, etc.) drops structure for roles/identity; typos and shape drift are only caught at runtime.

- **[Q7]** **Cross-layer duplicate audit write path** — `core/kp_db.py` documents that `_audit_append` mirrors `app/repositories/plate_audit_repository.py` — maintainers must update two implementations for one business rule change.

- **[Q8]** **Test gap** — No dedicated coverage for JSON API `app/api/v1/endpoints/offers.py` (unlike extensive `plan_manager` / production tests under `tests/`), so offer CRUD/discount/move endpoints are brittle to regressions.

- **[Q9]** **`app/planning/plan_manager.py` mutates `sys.path` on import** (~20–23) — environment-dependent import behavior and “works only from certain cwd/launchers” risk.

- **[Q10]** **`app/services/optimization_service.py` `legacy_runtime`** temporarily **reassigns globals** in `core.optimization` — easy to get wrong in async/concurrent or partial test runs; high mental overhead compared to a state object passed explicitly.

#### Low (Q11–Q14)

- **[Q11]** **Inconsistent user-facing error strings** in the same API surface — e.g. English `"Plan not found"` in `app/api/v1/endpoints/production.py` vs Russian messages in other responses.

- **[Q12]** **`bot/handlers/plan_manager.py`** uses `from app.planning.plan_manager import *` — hides the real public API and weakens tooling (go-to-definition, unused-import detection).

- **[Q13]** **`core/kp_db.py`** keeps many **session-scoped debug log constants** and `_debug_session_write` with broad `try` / **silent `pass`** — dead-weight paths and accidental silent failures in shared infrastructure code.

- **[Q14]** **`bot/handlers/export.py`** contains **TODO** comments (~28, ~65) for unimplemented `domain.export` — explicit unfinished layer, easy to forget.

---

## 5. Priority Matrix (P0–P2)

Top items derived from severity, blast radius, and fix cost. (Lower granularities can be scheduled as P3+.)

| Priority | IDs (representative) | Rationale |
|----------|---------------------|-----------|
| **P0** | **S1**, **A1**, **A2**, **A3**, **A4** | Forged sessions / unsafe defaults; dependency inversion breaking `core` reuse; global optimizer races; cookie + secret misconfiguration; O(n) auth on every request |
| **P1** | **S2**, **S3**, **S4**, **S5**, **A5**, **A6**, **A7**, **A8**, **A9**, **A10**, **A11**, **Q1–Q4** | Abuse resistance (cookie, rate limits, SQLite blocking, OCR limits); structural debt (god modules, dual domain, dual web/API surface, sync I/O, fat handlers, repository honesty) |
| **P2** | **A12–A16**, **S6–S13**, **Q5–Q10** | Observability, config sprawl, global optimization coupling, CSRF/CORS/health/debug metadata, filesystem trust; CQ: `Depends`, typed user DTO, dual audit path, offers test gap, `sys.path`, optimizer globals |
| **P3** (backlog) | **A17–A20**, **S14–S17**, **Q11–Q14** | Inline services, star-import shim, `kp_db` silent paths, i18n errors, export TODOs; cookie delete parity, security headers, bot token ops hygiene, verbose `traceback` in logs |

---

## 6. Next Steps

1. **Immediate (before next release):** Enforce non-default **APP_SECRET_KEY** and startup failure on weak values **[S1]**; remove **`core → app`** import by injecting settings/URI into `core` from composition root **[A1]**; isolate **optimization** state per request/task (no shared mutable module globals) **[A2]**; set **Secure** / **HttpOnly** / **SameSite** appropriately for prod (**[S2]**, aligns with **[A3]**); replace **`list_users`** in `get_current_user` with **`get_user_by_id`** + caching as needed **[A4]**.
2. **Current sprint:** Add **authentication rate limiting** and lockout policy **[S3]**; route blocking SQLite behind **threads or async drivers** (**[S4]**, aligns with **[A9]**); move **OCR limits** to a shared store for multi-worker deployments **[S5]**; begin **bounded-context splits** for `kp_db` and `plan_manager` **[A5]/[A6]/[Q2]/[Q3]**.
3. **Next sprint:** Converge **PlateOrder** representations **[A7]/[Q3]/[Q4]**; reduce **web/router** monolith and duplicate flows **[A8]**; tighten **bare `except:`** in bot handlers **[Q1]**; CSRF tokens or SameSite policy for web forms **[S7]**; tighten **CORS** headers to required set **[S8]**.
4. **Backlog:** Structured logging in optimizer **[A12]**; unify **DB path** resolution **[A13]**; remove dead **settings** keys or implement them **[A14]**; security headers **[S15]**; logging hygiene **[S17]**; `Depends()` for services and tests **[Q5]**; API tests for **offers** **[Q8]**; replace star-import shim **[Q12]**; drop `sys.path` side effects **[Q9]/[A19]**.

---

## Appendix A — Architecture findings (verbatim register)

### Critical

- [A1] **Inverted dependency (shared → app):** `core/db_config.py` imports `app.core.settings.get_settings`, so the supposedly shared `core/` layer depends on the FastAPI app package. That breaks inward-only dependency flow, makes `core` non-reusable, and risks import/bootstrap cycles if `app` ever pulls modules that load `core.db_config` early.

- [A2] **Global optimization state + concurrency risk:** `core.optimization` keeps mutable module-level plans/maps; `app/services/optimization_service.py` `legacy_runtime` swaps those globals around `yield`. Any overlapping requests (multiple workers / threads) can clobber each other's optimization results—scales poorly and violates safe service design.

- [A3] **Session security defaults:** `app/api/v1/endpoints/auth.py` sets the session cookie with `secure=False` unconditionally; `app/core/settings.py` ships a default `APP_SECRET_KEY` string. On HTTPS production, cookies are exposed to downgrade/MITM-adjacent risks, and a forgotten default secret invalidates the HMAC model in `app/security/session.py`.

- [A4] **Auth hot path loads entire user table:** `app/dependencies/auth.py` `get_current_user` calls `repository.list_users()` and builds a full `{id: user}` map on **every** authenticated request. This is O(n) per request, painful at scale, and loads more identity data than needed (should be `get_user_by_id` only).

### High

- [A5] **God module in `core`:** `core/kp_db.py` is a very large single module (thousands of lines) mixing schema, CRUD, production workflows, managers, rests, debug logging, and KPI-style helpers. Single Responsibility and testability are severely strained; regressions in one area are hard to isolate.

- [A6] **God module for planning I/O:** `app/planning/plan_manager.py` centralizes JSON/metadata, calendar rules, plan IDs, occupancy, track distribution, and file paths under `bot/data/plans`. High coupling and many reasons to change; hard to evolve storage (e.g. DB) behind a narrow port.

- [A7] **Dual domain representations:** Legacy `core/config_and_data.PlateOrder` coexists with `app/domain/models/plate_order.py` (and bot code imports both paths, e.g. `bot/handlers/production_execution.py`, `bot/handlers/production_day_view.py`). Mapping drift and subtle field mismatches are likely; boundaries are unclear.

- [A8] **Monolithic web layer:** `app/web/router.py` is a large HTML/SSR-style surface (~1000 lines) that parallels REST in `app/api/v1/endpoints/`, multiplying maintenance and security review surface (two ways to do the same flows).

- [A9] **Blocking I/O on async stack:** Most API handlers and repositories use synchronous `sqlite3` and sync `def` routes (`app/repositories/*.py`, e.g. `app/repositories/auth_repository.py`; endpoints like `app/api/v1/endpoints/production.py`). Under load, worker threads block the event loop pattern typical of FastAPI + async server deployments.

- [A10] **Fat Telegram handlers:** `bot/handlers/commercial.py` and `bot/handlers/production_completion.py` are very large modules (1000+ lines) mixing FSM/UI, file handling, and service calls—classic presentation-layer god objects, brittle to test and change.

- [A11] **"Repository" as pass-through:** `app/repositories/plan_repository.py` only delegates to `app.planning.plan_manager`; the repository pattern does not hide storage details or enable swapping implementations—callers remain tightly bound to JSON plan files.

### Medium

- [A12] **Non-structured diagnostics in optimizer entrypoint:** `core/optimization/orchestrator.py` uses `print()` for mode logging instead of the project logger (`core/logging_config.py` / `logging`), which complicates production observability and log aggregation.

- [A13] **Multiple sources of truth for DB paths:** e.g. `core/kp_db.py` defines `DEFAULT_DB` from `__file__` layout while `app/core/settings.py` and `core/db_config.py` use settings; `bot/bot_main.py` logs a hard-coded `plita.db` path separately from `bot/bot_config.DB_PATH_STR` (`pb.db`). Operational confusion and misconfiguration risk.

- [A14] **Unused configuration surface:** `app/core/settings.py` defines `database_url` and `redis_url` with no other references in the codebase—noise for operators and a hint of abandoned or premature infrastructure planning.

- [A15] **Heavy orchestration services:** `app/services/production_planning_service.py` and `app/services/day_documents_service.py` pull in many `core.*` concerns (optimization, visualization, serialization, commit, debug paths). They behave as application monoliths; further decomposition would clarify use-cases.

- [A16] **Legacy coupling via global `core.optimization`:** Aside from concurrency, the pattern in `app/services/optimization_service.py` (mutating `legacy_optimization` module attributes) leaks implementation detail and complicates parallel tests and future non-global APIs.

### Low

- [A17] **Ad-hoc service construction in routes:** Endpoints often do `ProductionService()` / `CommercialService()` / `CommercialWorkflowService()` inline instead of FastAPI `Depends` providers (`app/api/v1/endpoints/production.py`, `app/api/v1/endpoints/commercial.py`)—works, but weakens a single composition root and consistent overrides in tests.

- [A18] **Star-import shim:** `bot/handlers/plan_manager.py` re-exports `app.planning.plan_manager` with `from app.planning.plan_manager import *`, obscuring which symbols handlers rely on and hindering refactors.

- [A19] **Import-time path mutation:** `app/planning/plan_manager.py` mutates `sys.path` when imported—side effect that can surprise tooling, duplicate path entries, and order-dependent behavior.

- [A20] **Embedded debug log constants in `kp_db`:** `core/kp_db.py` wires multiple `_DEBUG_*` paths at module level—ties business data access to developer debug artifacts and clutters production modules.

---

## Appendix B — Security findings (verbatim Critical & High)

### Critical

- [S1] **Hardcoded default signing secret for sessions** — `APP_SECRET_KEY` falls back to the fixed string `change-this-secret-key-in-env` in `app/core/settings.py` (lines 32–35). Anyone who can reach the app with that default (or who knows a leaked/stale deploy) can forge valid `app_session` HMAC tokens (`app/security/session.py` uses this key at lines 14–16). **Must** require a strong random secret in production and fail fast if unset or equal to the default.

### High

- [S2] **Session cookies not marked `Secure`** — `response.set_cookie(..., secure=False)` in `app/api/v1/endpoints/auth.py` (lines 26–32). The HTML flow sets the same cookie without `secure` in `app/web/router.py` (lines 203–205). On any HTTP path or mixed content, the session cookie is easier to steal or replay (MITM / insecure transport).

- [S3] **No rate limiting on authentication** — `POST /api/v1/auth/login` (`app/api/v1/endpoints/auth.py`, lines 13–34) and `POST /web/login` (`app/web/router.py`, lines 197–205) accept unlimited login attempts (no lockout, no IP/user throttling). Enables online password guessing and load-based DoS against `AuthRepository.authenticate` (PBKDF2 at 200k iterations in `app/repositories/auth_repository.py`, lines 13–16, 131–139).

- [S4] **Blocking SQLite on the async/concurrent stack** — `AuthRepository` and related code use synchronous `sqlite3.connect` everywhere (e.g. `list_users` at `app/repositories/auth_repository.py`, lines 142–149; `get_current_user` calls this per request in `app/dependencies/auth.py`, lines 21–24). Under `async` routes and multiple concurrent clients, workers block on the event loop → **availability / slowloris-style DoS** and harder-to-reason timing behavior; also increases contention on `plita.db` if other code holds locks.

- [S5] **OCR upload rate limit is per-process only** — `_CommercialOcrUploadLimiter` in `app/services/commercial_upload_validation.py` (lines 18–37) is in-memory. With multiple Uvicorn workers or processes, each has its own counter; effective limits multiply and abuse capacity increases.

---

## Appendix C — Code quality findings (Q1–Q14)

### High

- **[Q1]** **Bare `except:`** in multiple bot handlers swallows all exceptions (including unexpected ones), hides failures, and complicates debugging — e.g. `bot/handlers/archive.py`, `bot/handlers/production_execution.py`, `bot/handlers/production_completion.py`, `bot/handlers/production_export.py`, `bot/handlers/instructions.py`, `bot/handlers/pb_info.py`, `bot/handlers/production_plans_list.py`, `bot/handlers/admin.py` (non-exhaustive; grep shows many occurrences).

- **[Q2]** **Oversized modules** hurt readability, review, and safe refactors: `core/kp_db.py` (~3410 lines, dozens of public functions), `app/web/router.py` (~3700 lines), `app/planning/plan_manager.py` (~3008 lines).

- **[Q3]** **DRY / domain drift**: two parallel `PlateOrder` implementations — `app/domain/models/plate_order.py` (dataclass + `to_orders_2d`) vs `core/config_and_data.py` (legacy class + `from_dict` / `from_orders_2d`) — every schema change must be reconciled manually.

- **[Q4]** **`app/planning/plan_manager.py`**: bare `except:` at date parsing / float coercion (~424–427, ~989–992, ~1279–1282) masks bad input; **duplicated merge logic** for `plate_lookup_by_length` (same float-coercion + merge pattern appears twice in the Gantt/combine flow around ~984–999 and ~1274–1297).

### Medium

- **[Q5]** **DI / testability**: routes construct **new service instances per call** (`ProductionService()` throughout `app/api/v1/endpoints/production.py`, `OffersService()` in every handler in `app/api/v1/endpoints/offers.py`) instead of FastAPI `Depends()` providers — harder to swap fakes at the HTTP boundary without patching classes.

- **[Q6]** **`_user: dict`** on protected endpoints (`app/api/v1/endpoints/offers.py`, `production.py`, `admin.py`, etc.) drops structure for roles/identity; typos and shape drift are only caught at runtime.

- **[Q7]** **Cross-layer duplicate audit write path**: `core/kp_db.py` documents that `_audit_append` mirrors `app/repositories/plate_audit_repository.py` — maintainers must update two implementations for one business rule change.

- **[Q8]** **Test gap**: no dedicated coverage for JSON API `app/api/v1/endpoints/offers.py` (unlike extensive `plan_manager` / production tests under `tests/`), so offer CRUD/discount/move endpoints are brittle to regressions.

- **[Q9]** **`app/planning/plan_manager.py`** mutates **`sys.path` on import** (~20–23) — environment-dependent import behavior and “works only from certain cwd/launchers” risk.

- **[Q10]** **`app/services/optimization_service.py`** `legacy_runtime` temporarily **reassigns globals** in `core.optimization` — easy to get wrong in async/concurrent or partial test runs; high mental overhead compared to a state object passed explicitly.

### Low

- **[Q11]** **Inconsistent user-facing error strings** in the same API surface — e.g. English `"Plan not found"` in `app/api/v1/endpoints/production.py` vs Russian messages in other responses.

- **[Q12]** **`bot/handlers/plan_manager.py`** uses `from app.planning.plan_manager import *` — hides the real public API and weakens tooling (go-to-definition, unused-import detection).

- **[Q13]** **`core/kp_db.py`** keeps many **session-scoped debug log constants** and `_debug_session_write` with broad `try` / **silent `pass`** — dead-weight paths and accidental silent failures in shared infrastructure code.

- **[Q14]** **`bot/handlers/export.py`** contains **TODO** comments (~28, ~65) for unimplemented `domain.export` — explicit unfinished layer, easy to forget.

---

*Consolidated audit report (Composer 2 subagents only). **Appendix A** — architecture register; **Appendix B** — security Critical/High; **Appendix C** — code-quality ladder from reviewer. **§4.2** medium/low security matches security-auditor wording for **S6–S17**.*

