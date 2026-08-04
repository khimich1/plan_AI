# Security sprint (WP1–WP5)

Спринт hardening веб-приложения. Канонический путь: SPA + `/api/v1/auth/*`.
Legacy `/web/*` — только redirect/stub до P6 (см. [p6-legacy-decommission.md](./p6-legacy-decommission.md)).

**Статус на 2026-07-06:** WP1–WP2 закрыты; WP3 закрыт в коде (спека обновлена); WP4 — variant B; WP5 — stretch.

---

## WP1 — S1: logout POST + CSRF

**Acceptance:**

- [x] Logout доступен только через **POST** (GET `/web/logout` → 405 без invalidate session)
- [x] Cookie-based logout требует валидный CSRF token (как login/forms)
- [x] Frontend: форма или fetch POST с token
- [x] Тесты: `tests/test_web_logout_csrf.py`, `tests/test_csrf.py`

**Verify:**

```bash
pytest tests/test_csrf.py tests/test_web_logout_csrf.py -q
```

---

## WP2 — S6: full destructive guard

**Acceptance:**

- [x] Инвентаризация всех destructive endpoints (admin reset, archive wipe, bulk delete)
- [x] Каждый path вызывает единый guard (`ALLOW_DESTRUCTIVE_DB_RESET` / env fail-closed)
- [x] Production/staging: 403 без явного override; audit log при deny
- [x] Тесты покрывают ранее неохваченные routes (`test_admin_destructive_guard.py`)

**Verify:**

```bash
pytest tests/test_admin_destructive_guard.py -q
```

---

## WP3 — S2: session invalidation on logout

**Acceptance:**

- [x] Logout **инвалидирует** server-side session (`app_users.session_version` bump + clear cookie)
- [x] Модель задокументирована: stateless HMAC cookie + `sv` claim; max age = `SESSION_COOKIE_MAX_AGE` при выдаче (отдельный idle sliding window **не** реализован — осознанный variant B)
- [x] Тест: после logout старый session cookie не даёт доступ (`test_stale_session_cookie_rejected_after_logout`)

**Реализация:** `app/security/session.py`, `app/api/v1/endpoints/auth.py`, `app/repositories/auth_repository.py`

**Verify:**

```bash
pytest tests/test_app_session.py tests/test_web_logout_csrf.py -q
```

**Out of scope (future):** per-device revocation без global `session_version` bump; sliding idle timeout.

---

## WP4 — S3: rate limit multi-worker

**Acceptance:**

- [ ] **Вариант A (preferred при multi-worker):** shared store для login/OCR rate limits — **отложено (P7 / Redis)**
- [x] **Вариант B (single instance):** deployment constraint зафиксирован; rate limit in-process
- [x] Тесты не регрессируют существующий login rate limit

**Реализация:** `app/security/login_rate_limit.py` — `warn_if_multi_worker_without_shared_store()`

**Verify:**

```bash
pytest tests/test_auth_login_rate_limit.py tests/test_web_login_rate_limit.py tests/test_rate_limit_deployment.py -q
```

**Deployment constraint:** один uvicorn worker **или** `APP_STORAGE_LAYOUT=shared_volume` + будущий Redis (P7).

---

## WP5 (stretch) — A3 phase 2 backlog prep

**Acceptance (optional):**

- [x] Checklist: [pep562-config-and-data-decommission.md](../pep562-config-and-data-decommission.md)
- [x] Нет обязательного code removal в этом спринте — удаление proxy в P6-B

---

## Definition of Done (спринт)

- [x] WP1–WP2 acceptance выполнены
- [x] WP3 закрыт (код + тесты)
- [x] WP4 — variant B задокументирован
- [x] **Нет** новых bot-specific WP или расширения parity tests
- [x] `pytest tests/ -q` зелёный
