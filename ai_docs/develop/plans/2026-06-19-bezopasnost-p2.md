# PLAN: Безопасность P2 — rate limit, object RBAC, frontend 409 (аудит 2026-06-19)

> **Фаза SDD:** PLAN → IMPLEMENT → **CLOSED**
> **Дата:** 2026-06-19
> **Дата закрытия:** 2026-06-20
> **Orchestration:** `orch-2026-06-20-bezopasnost-p2` — **закрыт**
> **Спека:** [`../../specs/bezopasnost-p2-audit-2026-06-19.md`](../../specs/bezopasnost-p2-audit-2026-06-19.md)
> **Baseline:** [`../../specs/project-baseline.md`](../../specs/project-baseline.md)
> **Предшественник:** [`2026-06-19-stabilizaciya-p1.md`](./2026-06-19-stabilizaciya-p1.md) — **закрыт**
> **Источник:** [`../audits/2026-06-19-full-project-audit.md`](../audits/2026-06-19-full-project-audit.md) → Post-P1

---

## 0. Резюме плана

Три обязательных work package + один optional:

- **WP0 (S2):** rate limiting на `POST /api/v1/auth/login` — 5 req/min/IP → 429.
- **WP1 (S3):** object-level RBAC для КП и archive — filter by `manager_id`, admin full access.
- **WP2 (FE-409):** frontend reload UX при `plan_version_conflict`.
- **WP3 (S4, optional):** npm audit fix, pinned versions, CI gate.

**Health Score цель:** ~6/10 → **~7/10** (S2 + S3 сняты из high backlog; UX conflict resolved).

### Граф зависимостей

```
WP0 (S2: login rate limit)              [старт сразу, независим]
        │
        │  (параллельно)
        ▼
WP1 (S3: object-level RBAC КП/archive)  [старт сразу, независим от WP0]
        │
        │  (параллельно — API 409 уже стабилен с P0/P1)
        ▼
WP2 (FE-409: reload plan on conflict)   [может стартовать параллельно с WP1]

WP3 (S4: npm CVE) — OPTIONAL, после WP0–WP2 или параллельно, не блокирует P2 closure
```

**Параллельно:** `WP0`, `WP1`, `WP2` (разные агенты/разработчики).
**Строго последовательно:** нет hard deps между WP0–WP2; WP3 optional в конце.

---

## Прогресс (closure 2026-06-20)

| WP | Статус | Находка |
|----|--------|---------|
| WP0 | ✅ done | S2 — login rate limit (`app/security/login_rate_limit.py`) |
| WP1 | ✅ done | S3 — object-level RBAC (`app/security/offer_access.py`, `owner_user_id`) |
| WP2 | ✅ done | FE-409 — reload on `plan_version_conflict` |
| WP3 | ⏸ deferred | S4 — npm CVE (optional, не блокирует P2) |

### Gates

| Gate | Статус |
|------|--------|
| G0 (WP0) | ✅ closed |
| G1 (WP1) | ✅ closed |
| G2 (WP2) | ✅ closed |
| **G3 — P2 closure** | ✅ **closed** |
| G4 (WP3 optional) | ⏸ deferred |

---

## WP0 — S2: rate limiting на login

**Зачем:** закрыть high-priority `S2` из audit Post-P1. Brute-force на единственной точке аутентификации web.

**Текущее состояние:**
- `app/api/v1/endpoints/auth.py` — `login()` без throttling.
- Паттерн in-process limiter уже есть: `app/services/commercial_upload_validation.py` (`check_commercial_ocr_rate_limit`, `reset_commercial_ocr_rate_limiter_for_tests`).

**Работы:**
1. Создать `app/security/login_rate_limit.py` (или `app/services/auth_rate_limit.py`):
   - Sliding/fixed window counter keyed by client IP.
   - `check_login_rate_limit(request: Request) -> None` raises HTTP 429.
   - `reset_login_rate_limiter_for_tests()`.
2. Вызвать check в начале `login()` **до** `authenticate` (fail fast).
3. Response 429: `Retry-After` header; body — structured или string `detail` (consistent с проектом).
4. **Тесты** `tests/test_auth_rate_limit.py`:
   - `test_login_rate_limit_allows_five_attempts`
   - `test_login_rate_limit_blocks_sixth_with_429_and_retry_after`
   - `test_login_rate_limit_does_not_apply_to_me`
5. Убедиться `tests/test_app_session.py` зелёный.

**Files (~3–4):**
- `app/security/login_rate_limit.py` (new)
- `app/api/v1/endpoints/auth.py`
- `tests/test_auth_rate_limit.py` (new)

**Verify:**
```powershell
pytest tests/test_auth_rate_limit.py -q
pytest tests/test_app_session.py -q
```

**Зависимости:** нет.

**Gate G0:** rate limit tests green; 6-й login → 429.

**Complexity:** Simple · **~0.5–1 день**

---

## WP1 — S3: object-level RBAC для КП и archive

**Зачем:** закрыть high-priority `S3`. Manager не должен читать/менять чужие КП.

**Текущее состояние:**
- `require_roles("admin", "manager")` на `offers.py`, `archive.py`.
- `OffersService.list_offers` / `ArchiveService.list_offers` — без filter by user.
- `manager_id` хранится в metadata КП (`commercial_workflow_service`).

**Работы:**
1. **Authorization helper** — `app/security/offer_access.py` (new) или расширить `app/dependencies/auth.py`:
   - `def can_access_offer(user: dict, offer: dict) -> bool`
   - `def assert_offer_access(user: dict, offer: dict) -> None` → HTTP 403
   - `def offer_filter_for_user(user: dict) -> dict | None` — `{"manager_id": user["id"]}` для manager; `None` для admin.
2. **Repository layer** (preferred):
   - `kp_repository.list_offers_grouped(manager_id: int | None = None)`
   - `kp_archive_repository` — аналогичные filter params для list/search.
3. **Service layer:** прокинуть `user` из endpoints в `list_offers`, `get_offer`, mutate methods.
4. **Endpoints** — минимальный scope:
   - `app/api/v1/endpoints/offers.py` — all routes
   - `app/api/v1/endpoints/archive.py` — list, search, get, patch, download
   - (если нужно) `commercial.py` — get/update flows с kp_id
5. **Create offer:** auto-set `manager_id = user["id"]` для role manager; admin может указать явно.
6. **Legacy КП без manager_id:** policy — visible only to admin (document + test).
7. **Тесты:**
   - `tests/test_offers_authorization.py` (new)
   - `tests/test_archive_authorization.py` (new)
   - Fixtures: `manager_a`, `manager_b`, `admin_user`, offers owned by each.

**Files (~6–10):**
- `app/security/offer_access.py` (new)
- `app/dependencies/auth.py` (optional re-export)
- `app/api/v1/endpoints/offers.py`
- `app/api/v1/endpoints/archive.py`
- `app/services/offers_service.py`
- `app/services/archive_service.py`
- `app/repositories/kp_repository.py`
- `app/repositories/kp_archive_repository.py`
- `tests/test_offers_authorization.py` (new)
- `tests/test_archive_authorization.py` (new)

**Verify:**
```powershell
pytest tests/test_offers_authorization.py tests/test_archive_authorization.py -q
pytest tests/test_archive_endpoints.py tests/test_commercial_web_flow.py -q
```

**Зависимости:** нет (параллельно с WP0).

**Gate G1:** cross-manager access denied; admin bypass works; list filtered.

**Complexity:** Moderate · **~2–3 дня**

---

## WP2 — FE-409: reload UX при plan_version_conflict

**Зачем:** закрыть Post-P1 frontend gap. Пользователь не застревает на stale plan после concurrent edit.

**Текущее состояние:**
- Backend: 409 + `code: plan_version_conflict` + `expected_version` в details (`production.py`, `errors.py`).
- Frontend: `parseApiErrorPayload` парсит `code`; mutations не обрабатывают reload.
- Types: `PlanMetaSummary.version` уже в `production.ts`.

**Работы:**
1. **Helper** `frontend/src/shared/lib/planConflict.ts` (new):
   - `isPlanVersionConflict(error: unknown): boolean`
   - `handlePlanVersionConflict(queryClient, options?)` — toast message + invalidate queries.
2. **Query keys** — централизовать production keys (если ещё нет): `productionKeys.plans`, `productionKeys.plan(id)`, `productionKeys.calendar`, `productionKeys.dayView(date)`.
3. **Integrate** в mutation `onError`:
   - `PlansList.tsx` — delete, activate
   - `DayDrawer.tsx` — complete day, delete track (не путать с другими 409)
   - `WorkCalendarEditor.tsx` — save calendar
   - `CreatePlanWizard.tsx` — build plan (if returns 409)
4. **UX copy:** «План был изменён. Данные обновлены — повторите действие.»
5. **Tests:** Vitest для `isPlanVersionConflict`; optional component test.

**Files (~5–7):**
- `frontend/src/shared/lib/planConflict.ts` (new)
- `frontend/src/shared/lib/planConflict.test.ts` (new)
- `frontend/src/features/production/components/PlansList.tsx`
- `frontend/src/features/production/components/DayDrawer.tsx`
- `frontend/src/features/production/components/WorkCalendarEditor.tsx`
- `frontend/src/features/production/components/CreatePlanWizard.tsx` (if applicable)
- (optional) `frontend/src/features/production/hooks/useProductionMutations.ts`

**Verify:**
```powershell
cd frontend
npm run test
npm run typecheck
npm run build
```

**Зависимости:** нет hard dep; параллельно с WP1.

**Gate G2:** vitest green; build green; manual two-tab conflict scenario.

**Complexity:** Moderate · **~1–2 дня**

---

## WP3 — OPTIONAL: S4 npm CVE

**Зачем:** high-priority supply chain из audit; **не блокирует P2 closure**.

**Работы:**
1. `cd frontend && npm audit` — зафиксировать baseline.
2. `npm audit fix` (review lockfile diff).
3. Replace `"latest"` in `package.json` with pinned semver (react, vite, react-router-dom, undici transitive via vite).
4. Document CI step `npm audit --audit-level=high` (implement CI only Ask first).
5. Re-run `npm run test && npm run build`.

**Files (~2):**
- `frontend/package.json`
- `frontend/package-lock.json`
- (optional) `.github/workflows/*.yml`

**Verify:**
```powershell
cd frontend
npm audit --audit-level=high
npm run test
npm run build
```

**Зависимости:** none; optional after G3.

**Gate G4 (optional):** npm audit clean or documented exceptions.

**Complexity:** Simple · **~0.5–1 день**

---

## Риски и митигации (на уровне плана)

| Риск | Где | Митигация |
|------|-----|-----------|
| Rate limit flaky on shared IP (NAT) | WP0 | Document; single-tenant deploy assumption |
| S3 breaks commercial flow tests | WP1 | Update conftest fixtures; admin tests separate |
| Null manager_id legacy rows | WP1 | Admin-only policy + migration note |
| FE double-toast on conflict | WP2 | Guard flag in handler; debounce invalidate |
| npm audit fix breaks build | WP3 | Pin incrementally; test after each bump |
| Scope creep legacy web UI | WP1 | Explicit out of scope in spec |

---

## Контрольные точки верификации (gates)

| Gate | Условие перехода |
|------|------------------|
| **G0** (после WP0) | `test_auth_rate_limit` green; login 429 behavior confirmed |
| **G1** (после WP1) | Authorization tests green; manager isolation proven |
| **G2** (после WP2) | `npm run test && npm run build` green |
| **G3 — P2 closure** | G0 + G1 + G2; `pytest tests/ -q` green; audit Post-P2 draft |
| **G4 — optional** (после WP3) | `npm audit --audit-level=high` acceptable |

### Full regression (G3) — ✅ пройден 2026-06-20

```powershell
Set-Location "c:\Users\Роман\Desktop\Шишов"
.\.venv\Scripts\Activate.ps1

pytest tests/test_auth_login_rate_limit.py -q
pytest tests/test_offers_authorization.py tests/test_archive_authorization.py -q
pytest tests/test_http_errors.py tests/test_app_session.py -q
pytest tests/ -q   # 756 passed, 12 skipped

cd frontend
npm run test
npm run build
```

**Результат:** backend 756 passed / 12 skipped; frontend test + build green.

---

## Оценка трудозатрат (ориентир)

| WP | Complexity | Оценка |
|----|------------|--------|
| WP0 | Simple | 0.5–1 день |
| WP1 | Moderate | 2–3 дня |
| WP2 | Moderate | 1–2 дня |
| WP3 (opt) | Simple | 0.5–1 день |

**Итого P2 (WP0–WP2):** ~4–6 рабочих дней.

---

## Task IDs (orchestration)

| Task ID | WP | Name |
|---------|-----|------|
| P2-S2-001 | WP0 | Login rate limiter module |
| P2-S2-002 | WP0 | Wire login endpoint + tests |
| P2-S3-001 | WP1 | Offer access authorization helper |
| P2-S3-002 | WP1 | Repository filters by manager_id |
| P2-S3-003 | WP1 | Endpoint/service integration |
| P2-S3-004 | WP1 | Authorization test suite |
| P2-FE-001 | WP2 | planConflict helper + tests |
| P2-FE-002 | WP2 | Mutation onError integration |
| P2-S4-001 | WP3 | npm audit + pin versions (optional) |

---

## Deferred (явно вне обязательного closure P2)

| ID | Тема | Следующий спринт |
|----|------|------------------|
| **S4** | npm CVE | WP3 optional или отдельный chore |
| **Q5** | Integration tests production API | Quality sprint |
| **Q6** | Debug-instrumentation cleanup | Chore sprint |
| **S10** | Password policy | Security backlog |
| **S6** | Exception sanitization | Security backlog |

---

## Следующий шаг (после closure P2)

1. ~~Execute orchestration~~ — **завершено**.
2. **Quality sprint:** Q5 (integration tests production API), Q6 (debug cleanup).
3. **WP3 / S4:** `npm audit fix`, pinned versions, CI gate — отдельный chore.
4. **Gaps P2:** legacy web RBAC parity (`app/web/router.py`); bot без `owner_user_id` (deprecated).
5. Audit Post-P2: [`../audits/2026-06-19-full-project-audit.md#post-p2-remediation-status-2026-06-20`](../audits/2026-06-19-full-project-audit.md#post-p2-remediation-status-2026-06-20).

**Spec:** [`../../specs/bezopasnost-p2-audit-2026-06-19.md`](../../specs/bezopasnost-p2-audit-2026-06-19.md) — **closed (implemented)**
