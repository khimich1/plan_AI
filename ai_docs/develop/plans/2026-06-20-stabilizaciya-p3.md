# PLAN: Стабилизация P3 — legacy login, production RBAC, destructive guard (аудит 2026-06-20)

> **Фаза SDD:** PLAN → TASKS → IMPLEMENT
> **Дата:** 2026-06-20
> **Спека:** [`../../specs/stabilizaciya-p3-audit-2026-06-20.md`](../../specs/stabilizaciya-p3-audit-2026-06-20.md)
> **Baseline:** [`../../specs/project-baseline.md`](../../specs/project-baseline.md)
> **Предшественник:** [`2026-06-19-bezopasnost-p2.md`](./2026-06-19-bezopasnost-p2.md) — **закрыт**
> **Источник:** [`../audits/2026-06-20-full-project-audit.md`](../audits/2026-06-20-full-project-audit.md)

---

## 0. Резюме плана

Пять обязательных work package + два optional:

- **WP0 (S1):** rate limit на `POST /web/login` — переиспользовать `check_login_rate_limit`.
- **WP1 (S5):** production operational-only — 403 на commercial/offers; frontend nav + redirect.
- **WP2 (S4):** guard на все admin reset + `reset_plans_only` чистит SQLite `production_plans`.
- **WP3 (Q-M9):** integration tests production API.
- **WP4 (Q6):** debug-instrumentation cleanup.
- **WP5/WP6 (opt):** npm CVE, user lookup cache.

**Health Score цель:** ~7/10 → **~7.5–8/10** (S1, S4, S5 сняты).

### Граф зависимостей

```
WP0 (S1: /web/login rate limit)     [старт сразу]
        │
WP1 (S5: production RBAC + FE)      [параллельно с WP0, WP2]
        │
WP2 (S4: destructive guard + SQLite plans reset)  [параллельно]
        │
        ▼
WP3 (Q-M9: production API integration tests)
        │
        ▼
WP4 (Q6: debug cleanup)

WP5 (npm) / WP6 (A4 cache) — OPTIONAL, не блокируют closure
```

**Параллельно:** WP0, WP1, WP2 (разные файлы/слои).
**Последовательно:** WP3 после WP1 (стабильный RBAC); WP4 в конце.

---

## WP0 — S1: Rate limit на legacy `POST /web/login`

**Зачем:** закрыть bypass brute-force через legacy path (аудит 20.06 S1).

**Текущее состояние:**
- `app/api/v1/endpoints/auth.py` — `check_login_rate_limit(request)` ✅
- `app/web/router.py:211` — `login_submit` без throttling ❌

**Работы:**
1. Импорт `check_login_rate_limit` из `app/security/login_rate_limit.py`.
2. В `login_submit` — inject `Request`, вызвать check **до** `authenticate`.
3. При 429 — вернуть HTML error page или redirect с query (consistent с legacy UX); предпочтение: `HTTPException(429)` с `Retry-After` если FastAPI route поддерживает — иначе redirect `/web/login?error=...`.
4. **Тесты** `tests/test_web_login_rate_limit.py`:
   - `test_web_login_rate_limit_blocks_sixth_attempt`
   - reuse `reset_login_rate_limiter_for_tests` из OCR/login pattern
5. Регрессия: `tests/test_auth_login_rate_limit.py`

**Files (~2–3):**
- `app/web/router.py`
- `tests/test_web_login_rate_limit.py` (new)
- (optional) `tests/test_app_session.py`

**Verify:**
```powershell
pytest tests/test_web_login_rate_limit.py tests/test_auth_login_rate_limit.py -q
```

**Gate G0:** 6-й POST `/web/login` → 429 (или эквивалент block).

**Complexity:** Simple · **~2–4 часа**

**Deferred:** полный deprecate `/web/login` → backlog **A7**.

---

## WP1 — S5: Production RBAC + frontend UX

**Зачем:** закрыть утечку коммерческих данных (суммы, менеджеры) для role `production`.

**Текущее состояние:**
- `offer_access.py`: `_PRODUCTION_READ_STATUSES = {"в работе", "выполнено"}`; production может read offers
- `offers.py`: `require_roles(..., "production")` на list/get/pdf/xlsx
- `archive.py`: только `admin`, `manager` — production уже blocked ✅
- Frontend: все роли видят «Создать КП» / «Архив»; login redirect → `/new`

**Работы — Backend:**
1. **`app/api/v1/endpoints/offers.py`:**
   - Убрать `"production"` из `require_roles` на всех routes **или** добавить dependency `forbid_production_commercial` → 403.
   - Минимальный diff: убрать production из allowed roles (list, get, pdf, xlsx).
2. **`app/security/offer_access.py`:**
   - `_PRODUCTION_READ_STATUSES = frozenset({"в работе"})` — для остаточных code paths / kp_repository filters.
   - Упростить `can_read_offer` для production → `False` (commercial API closed) **или** удалить production branch если API закрыт.
3. **`app/repositories/kp_repository.py`** — `list_kps_in_production()`: filter только `в работе` (если сейчас включает выполненные).
4. **Тесты** `tests/test_offers_production_authorization.py`:
   - fixture `production_user`
   - GET `/api/v1/offers` → 403
   - GET `/api/v1/offers/{id}/pdf` → 403
   - GET `/api/v1/production/kp-candidates` → 200
   - manager/admin regression spot-check

**Работы — Frontend:**
1. **`AppHeader.tsx`:** если `user.role === "production"` — не рендерить ссылки «Создать КП», «Архив».
2. **`LoginPage.tsx`:** `DEFAULT_REDIRECT` по роли — production → `/production`, иначе `/new`; то же в `Navigate` when already logged in.
3. **`AppRouter.tsx`:**
   - `index` + `*` fallback: role-aware redirect (через small helper `defaultRouteForRole(role)`).
   - Optional: `RoleHomeRedirect` component внутри `ProtectedRoute`.
4. **Vitest:** unit test `defaultRouteForRole` (optional).

**Files (~6–8):**
- `app/api/v1/endpoints/offers.py`
- `app/security/offer_access.py`
- `app/repositories/kp_repository.py` (if needed)
- `tests/test_offers_production_authorization.py` (new)
- `frontend/src/pages/login/LoginPage.tsx`
- `frontend/src/app/layout/AppHeader.tsx`
- `frontend/src/app/router/AppRouter.tsx`
- `frontend/src/shared/lib/roleRoutes.ts` (new, optional)

**Verify:**
```powershell
pytest tests/test_offers_production_authorization.py tests/test_offers_authorization.py -q
cd frontend && npm run test && npm run build
```

**Gate G1:** production → 403 offers; 200 production/kp-candidates; FE nav + redirect.

**Complexity:** Moderate · **~1–2 дня**

---

## WP2 — S4: Destructive admin guard + SQLite plans reset

**Зачем:** закрыть S4 — reset без guard; `reset_plans_only` не чистит authoritative SQLite store.

**Текущее состояние:**
- `reset_full`, `reset_kp_only` — `DestructiveDbOperationBlocked` ✅
- `reset_plans_only`, `reset_calendar_only` — **без guard** ❌
- `destructive_db_reset_allowed()`: любой `APP_ENV != production` → allow (staging без flag) ⚠️
- `admin_service.reset_plans_only` → только `_clear_all_plans()` (JSON files) ❌

**Работы — Guard:**
1. **`core/destructive_db_guard.py`** — уточнить policy:
   ```python
   # Allowed: APP_ENV == "development" OR ALLOW_DESTRUCTIVE_DB_RESET=1
   # Denied: production, staging, unknown — без flag
   ```
2. **`app/api/v1/endpoints/admin.py`:**
   - Обернуть `reset_plans_only`, `reset_calendar_only` в try/except `DestructiveDbOperationBlocked` (как full/kp-only).
3. Обновить **`tests/test_destructive_db_guard.py`** — staging без flag → deny.
4. **`tests/test_admin_destructive_guard.py`** (new):
   - monkeypatch `APP_ENV=production` → POST plans-only → 403
   - `APP_ENV=development` → 200

**Работы — SQLite plans reset:**
1. **`app/repositories/plan_repository.py`:**
   - `delete_all_plans() -> int` — `DELETE FROM production_plans`; return count.
2. **`app/services/admin_service.py`:**
   - `_clear_all_plans()` → rename/split: `_clear_sqlite_plans()` + `_clear_legacy_json_plans()` (best-effort).
   - `reset_plans_only()` вызывает оба; report включает `sqlite_plans_deleted`.
   - `reset_full()` — убедиться, что SQLite plans тоже очищаются.
3. **Обновить тесты** `tests/test_admin_service.py`:
   - Переименовать/переписать `test_reset_plans_only_does_not_touch_db` → assert KP tables intact, `production_plans` empty, JSON cleaned.
   - Seed plan row in `production_plans` fixture.

**Files (~5–6):**
- `core/destructive_db_guard.py`
- `app/api/v1/endpoints/admin.py`
- `app/services/admin_service.py`
- `app/repositories/plan_repository.py`
- `tests/test_admin_destructive_guard.py` (new)
- `tests/test_admin_service.py`, `tests/test_destructive_db_guard.py`

**Verify:**
```powershell
pytest tests/test_admin_destructive_guard.py tests/test_admin_service.py tests/test_destructive_db_guard.py -q
```

**Gate G2:** guarded resets; SQLite `production_plans` cleared on plans-only.

**Complexity:** Moderate · **~1 день**

**Ask first:** ужесточение staging guard — может затронуть deploy scripts; документировать `ALLOW_DESTRUCTIVE_DB_RESET=1`.

---

## WP3 — Q-M9: Integration tests production API

**Зачем:** safety net перед дальнейшими рефакторингами (deferred из P2).

**Работы:**
1. **`tests/test_production_api_integration.py`** (new) с fixtures: admin/production users, sample plan in SQLite.
2. **Happy path (минимум 8):**
   - `GET /production/plans`
   - `GET /production/plans/{id}`
   - `POST /production/plans/build` (или create flow)
   - `POST /production/plans/{id}/activate`
   - `GET /production/kp-candidates`
   - `GET /production/day/{date}`
   - `GET /production/work-calendar`
   - `PATCH /production/work-calendar`
3. **Failure modes (≥3):**
   - 401 без cookie
   - 403 manager на admin-only route (if applicable)
   - 409 plan version conflict (mock/stale version)
4. Shared fixtures в `conftest.py` — не дублировать monkeypatch из других tests.

**Files (~2–3):**
- `tests/test_production_api_integration.py` (new)
- `tests/conftest.py` (extend if needed)

**Verify:**
```powershell
pytest tests/test_production_api_integration.py -q
pytest tests/ -q
```

**Gate G3:** integration file green; full suite green.

**Complexity:** Moderate · **~2–3 дня**

**Зависимости:** желательно после WP1 (production RBAC stable).

---

## WP4 — Q6: Debug-instrumentation cleanup

**Зачем:** убрать agent log / debug file I/O из hot paths.

**Работы:**
1. ripgrep `#region agent log`, `debug-db`, `agent log` в `app/services/`.
2. Удалить или обернуть в `if get_settings().app_debug` (или `APP_DEBUG`).
3. Заменить оставшиеся `print()` → `logging`.
4. Регрессия: `test_production_planning_service`, day view tests.

**Files (~2–4):**
- `app/services/day_view_service.py`
- `app/services/production_planning_service.py`
- (others if found)

**Verify:**
```powershell
pytest tests/test_production_planning_service.py -q
pytest tests/ -q
```

**Gate G4 — P3 closure:** G0 + G1 + G2 + G3 + G4; full suite + frontend build.

**Complexity:** Simple · **~0.5–1 день**

---

## WP5 — OPTIONAL: npm CVE

См. spec P3 WP5 / P2 deferred. `npm audit fix`, pin versions.

**Gate G5 (optional):** `npm audit --audit-level=high`

---

## WP6 — OPTIONAL: A4 user lookup cache

`get_current_user` без O(n) `list_users()` — TTL cache или indexed lookup.

**Не блокирует P3 closure.**

---

## Риски и митигации

| Риск | Где | Митигация |
|------|-----|-----------|
| Web login 429 — HTML vs JSON | WP0 | Match legacy error UX; test both paths |
| Production заблокирован от offers, но legacy web открыт | WP1 | Out of scope A7; document |
| Staging reset сломан | WP2 | Docs + `ALLOW_DESTRUCTIVE_DB_RESET=1` |
| delete_all_plans FK/orphans | WP2 | Transaction; test with seeded plan |
| Flaky integration tests | WP3 | Isolated DB tmp_path fixtures |

---

## Контрольные точки (gates)

| Gate | Условие |
|------|---------|
| **G0** | `test_web_login_rate_limit` green |
| **G1** | production 403 offers; FE redirect/nav |
| **G2** | admin destructive guarded; SQLite plans reset |
| **G3** | `test_production_api_integration` green |
| **G4 — P3 closure** | G0–G3 + WP4; `pytest tests/ -q`; `npm run test && npm run build` |
| **G5 (opt)** | npm audit |

### Full regression (G4)

```powershell
Set-Location "c:\Users\Роман\Desktop\Шишов"
.\.venv\Scripts\Activate.ps1

pytest tests/test_web_login_rate_limit.py -q
pytest tests/test_offers_production_authorization.py -q
pytest tests/test_admin_destructive_guard.py -q
pytest tests/test_production_api_integration.py -q
pytest tests/ -q

cd frontend
npm run test
npm run build
```

---

## Оценка трудозатрат

| WP | Оценка |
|----|--------|
| WP0 | 0.5 дня |
| WP1 | 1–2 дня |
| WP2 | 1 день |
| WP3 | 2–3 дня |
| WP4 | 0.5–1 день |
| **Итого WP0–WP4** | **~5–7 дней** |

---

## Task IDs (для orchestration / TASKS)

| Task ID | WP | Name |
|---------|-----|------|
| P3-S1-001 | WP0 | Wire rate limit on `/web/login` |
| P3-S1-002 | WP0 | `test_web_login_rate_limit.py` |
| P3-S5-001 | WP1 | Remove production from offers routes |
| P3-S5-002 | WP1 | `в работе` only in kp candidates filter |
| P3-S5-003 | WP1 | Frontend nav hide + login redirect |
| P3-S5-004 | WP1 | `test_offers_production_authorization.py` |
| P3-S4-001 | WP2 | Guard plans-only + calendar-only endpoints |
| P3-S4-002 | WP2 | Tighten `destructive_db_reset_allowed` policy |
| P3-S4-003 | WP2 | `PlanRepository.delete_all_plans` + admin service |
| P3-S4-004 | WP2 | Admin destructive tests |
| P3-Q9-001 | WP3 | Production API integration test suite |
| P3-Q6-001 | WP4 | Remove/gate agent log blocks |

---

## Следующий шаг

1. Ревью PLAN (этот документ).
2. Фаза **TASKS** — задачи ≤5 файлов с acceptance + verify.
3. **IMPLEMENT** — порядок: WP0 ∥ WP1 ∥ WP2 → WP3 → WP4.

**Deferred после P3:**
- **A7** — deprecate legacy web UI целиком
- **A3** phase 2 — full globals decommission
- **WP5/WP6** — npm CVE, user cache
