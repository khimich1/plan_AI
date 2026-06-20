# Spec: Безопасность P2 — rate limit, object RBAC, frontend 409 (аудит 2026-06-19)

> **Тип:** remediation feature-spec (стабилизационный спринт P2)
> **Фаза SDD:** SPEC → PLAN → IMPLEMENT
> **Дата:** 2026-06-19
> **Ревизия:** v2 (closure)
> **Статус:** **closed (implemented)**
> **Дата закрытия:** 2026-06-20
> **Baseline:** [`project-baseline.md`](./project-baseline.md)
> **Предшественник:** [`stabilizaciya-p1-runtime-security-2026-06-19.md`](./stabilizaciya-p1-runtime-security-2026-06-19.md) — **закрыт**
> **Источник находок:** [`../develop/audits/2026-06-19-full-project-audit.md`](../develop/audits/2026-06-19-full-project-audit.md) → Post-P1

---

## Стратегия (одной фразой)

> Закрыть high-priority security backlog Post-P1: brute-force на login (S2), object-level RBAC для КП/archive (S3) и UX перезагрузки плана при optimistic-lock conflict (409).

**Контекст после P1:** Health Score **~6/10**. Все **4 critical** закрыты (A1/A2 в P0, A3/S1 в P1). High-priority кластер auth & RBAC + UX для `plan_version_conflict` — **закрыт в P2** (см. [Post-P2 в audit](../develop/audits/2026-06-19-full-project-audit.md#post-p2-remediation-status-2026-06-20)).

**Deferred из этого спринта:** S4 (npm CVE), Q5 (integration tests), Q6 (debug cleanup) — отдельные спринты/backlog.

---

## ASSUMPTIONS I'M MAKING

1. **P0 и P1 закрыты** — `pytest tests/ -q` зелёный (~744 passed, 12 skipped на closure P1); critical = 0; hot paths на `PlateOrderContext`/DI.
2. **Деплой по умолчанию — single instance** (`APP_STORAGE_LAYOUT=single_instance`); in-process rate limiter для S2 **достаточен** для текущего production; Redis — не цель P2.
3. **Auth-модель:** cookie-based session (`app_session`), роли `admin` | `manager` | `production`; `require_roles()` проверяет только role-level, **не** object-level (S3 gap подтверждён аудитом).
4. **Ownership КП:** для object-level RBAC введено поле **`owner_user_id`** в `kp_meta` (колонка + filter в repository); при create auto-set из session. Legacy `manager_id` в metadata сохранён для совместимости, но авторизация идёт по `owner_user_id`.
5. **Admin bypass:** роль `admin` видит все КП/archive; `manager` — только свои (`manager_id == user["id"]`); `production` — read-only scope по существующим endpoint rules (не расширять права в P2).
6. **409 conflict:** backend уже возвращает structured error `code: plan_version_conflict` + `expected_version` (`app/schemas/errors.py`, `app/api/v1/endpoints/production.py`); frontend парсит `code` в `apiError.ts`, но **не делает auto-reload** плана.
7. **OCR rate limit** (`commercial_upload_validation.py`) — отдельный механизм; S2 касается **только** `POST /api/v1/auth/login`.
8. **S4 optional WP3:** выполняется только при наличии времени после WP0–WP2; не блокирует closure P2.
9. PLAN — [`../develop/plans/2026-06-19-bezopasnost-p2.md`](../develop/plans/2026-06-19-bezopasnost-p2.md).

→ Подтвердите или поправьте допущения перед IMPLEMENT (особенно: mapping `manager_id` ↔ `user.id`, scope `production` role на archive).

---

## 1. Objective & Problem Statement

### Objective

Закрыть **3 приоритета Post-P1** из аудита и поднять устойчивость auth/data-access слоя без новых продуктовых фич. Целевой Health Score после P2: **~7/10** (снятие S2/S3 из high backlog + UX для version conflicts).

### Problem Statement (после P1)

1. **Brute-force на login (S2):** `POST /api/v1/auth/login` (`app/api/v1/endpoints/auth.py`) не ограничивает частоту попыток по IP. Атакующий может перебирать пароли без throttling.
2. **RBAC без object-level auth (S3):** endpoints `offers.py`, `archive.py`, часть `commercial.py` проверяют только роль (`require_roles("admin", "manager")`), но **не фильтруют** списки и **не запрещают** доступ к чужому `kp_id`. Любой manager видит все КП.
3. **Frontend 409 UX:** при concurrent edit плана API возвращает 409 `plan_version_conflict`; пользователь видит generic error без перезагрузки актуальной версии — риск повторных конфликтов и потери данных.

### Reframe: что значит «стало лучше»

| Требование | Reframed success criteria |
|------------|---------------------------|
| «Защита login» | 6-й запрос с одного IP за минуту → **429** + `Retry-After`; легитимный login не ломается |
| «Manager видит только свои КП» | List/search/get/mutate archive & offers для manager фильтруются или отклоняются (**403/404**) для чужих `kp_id` |
| «409 не dead-end» | UI предлагает перезагрузить план, подтягивает свежий `version`, пользователь может повторить действие |

---

## 2. Scope

### In scope

| ID | Категория | Название | Механизм |
|----|-----------|----------|----------|
| **S2** | Безопасность | Rate limiting на login | In-process limiter по IP: 5 req/min → 429 |
| **S3** | Безопасность | Object-level RBAC КП/archive | Filter by `manager_id`; admin — full access; authorization tests |
| **FE-409** | UX / Frontend | Reload при `plan_version_conflict` | Detect `code`, invalidate query, toast/dialog, refetch plan |

### Deferred (явно вне обязательного closure P2)

| ID / тема | Причина переноса |
|-----------|------------------|
| **S4** | npm CVE — optional WP3; supply-chain, не блокирует S2/S3 |
| **Q5** | Integration tests production API — отдельный quality-спринт |
| **Q6** | Debug-instrumentation cleanup — отдельный chore |
| **S10** | Password policy | Backlog |
| **S6** | `str(exc)` sanitization | Backlog |
| Bot/archive parity | Bot deprecated; не расширять bot RBAC в P2 |

### Out of scope

- Новые продуктовые фичи (wizard КП, новые роли, SSO).
- Redis/distributed rate limiting, account lockout после N failures (можно backlog).
- PostgreSQL migration, горизонтальное масштабирование.
- Полный audit closure всех 55 findings.
- Legacy web UI (`app/web/router.py`) object-level RBAC — только если явно используется; приоритет — REST API v1 + React SPA.

---

## 3. Tech Stack

Без изменений относительно baseline.

Ключевое для P2:
- **S2:** FastAPI middleware или dependency на login route; паттерн как `commercial_upload_validation.check_commercial_ocr_rate_limit` (in-process, test reset hook).
- **S3:** `app/dependencies/auth.py` — helper `assert_offer_access(user, offer)` / `filter_offers_for_user`; сервисный слой `offers_service`, `archive_service`, `kp_repository`.
- **FE-409:** React Query (`queryClient.invalidateQueries`), `ApiError.code`, production feature components (`PlansList`, `DayDrawer`, `CreatePlanWizard` — по месту мутаций).

**Зависимости (Ask first):** `slowapi` **или** lightweight custom limiter без новых pip-пакетов (предпочтение — custom, по аналогии с OCR limiter).

---

## 4. Commands

```powershell
Set-Location "c:\Users\Роман\Desktop\Шишов"
.\.venv\Scripts\Activate.ps1

# WP0 — S2 rate limit login
pytest tests/test_auth_login_rate_limit.py -q

# WP1 — S3 object-level RBAC
pytest tests/test_offers_authorization.py -q
pytest tests/test_archive_authorization.py -q
pytest tests/test_commercial_authorization.py -q   # если затронут commercial endpoints

# WP2 — frontend 409
cd frontend
npm run test
npm run typecheck
npm run build
cd ..

# Регрессия перед закрытием спринта
pytest tests/ -q
pytest tests/test_http_errors.py -q
pytest tests/test_app_session.py -q
```

---

## 5. Acceptance Criteria по находкам

### WP0 — S2: Rate limiting на login

#### S2 — Brute-force на login

- **Где:** `app/api/v1/endpoints/auth.py` (`POST /login`), опционально `app/main.py` (middleware registration)
- **Текущее состояние:** нет rate limit; каждый запрос вызывает `AuthRepository.authenticate`.
- **Acceptance:**
  - [x] Лимит: **5 попыток POST /api/v1/auth/login на IP за скользящее окно 60 секунд**.
  - [x] 6-й и последующие запросы с того же IP в окне → **HTTP 429** с телом structured error (или `detail` string) и заголовком **`Retry-After`** (секунды до reset).
  - [x] Успешный login **не сбрасывает** счётчик failed attempts (простой IP throttle; account lockout — out of scope).
  - [x] Неудачный login (401) **учитывается** в лимите так же, как успешный (защита от enumeration через timing — out of scope).
  - [x] Тестовый hook `reset_*_for_tests()` по аналогии с OCR limiter для детерминированных pytest (`reset_login_rate_limiter_for_tests`).
  - [x] Лимит **не применяется** к `/auth/logout`, `/auth/me` (только login).
  - [x] За reverse proxy: использовать `request.client.host`; документировать limitation для multi-worker (single instance assumption).
- **Verify:**
  - `tests/test_auth_login_rate_limit.py` — 5×401/200 OK, 6-й → 429 + `Retry-After`.
  - `tests/test_app_session.py` — существующие login tests зелёные.

---

### WP1 — S3: Object-level RBAC для КП и archive

#### S3 — RBAC без object-level authorization

- **Где (минимальный scope):**
  - `app/api/v1/endpoints/offers.py` — list, get, patch, delete, download
  - `app/api/v1/endpoints/archive.py` — list, search, get details, mutate (discount, logistics, move-to-production, file download)
  - `app/services/offers_service.py`, `app/services/archive_service.py`
  - `app/repositories/kp_repository.py`, `app/repositories/kp_archive_repository.py` (filter at query level preferred)
  - `app/dependencies/auth.py` — shared authorization helpers
- **Текущее состояние:** `require_roles("admin", "manager")` без проверки `manager_id` на object access.
- **Acceptance — правила доступа:**
  - [x] **`admin`:** полный доступ ко всем КП (без изменения текущего поведения).
  - [x] **`manager`:** list/search возвращают **только** КП где `kp_meta.owner_user_id == user["id"]`.
  - [x] **`manager`:** get/update/delete/download по чужому `kp_id` → **403 Forbidden** (единый стиль 403).
  - [x] **`production`:** существующие read endpoints в `offers.py` — только КП, доступные production workflow (фильтр по status сохранён).
  - [x] Create offer: manager auto-set `owner_user_id` из session.
  - [x] Admin — полный доступ; reassign через существующие admin flows (без расширения scope P2).
- **Acceptance — тесты:**
  - [x] Минимум **2 manager fixtures** (manager A, manager B) + admin.
  - [x] Test: manager A не видит kp_id manager B в list.
  - [x] Test: manager A → GET/PATCH/DELETE kp_id B → 403.
  - [x] Test: admin видит оба.
  - [x] Coverage: offers list/get + archive list/get + mutate (discount).
- **Verify:**
  - Новые файлы `tests/test_offers_authorization.py`, `tests/test_archive_authorization.py`.
  - Регрессия `tests/test_commercial_web_flow.py` зелёная (adjust fixtures if needed).

---

### WP2 — Frontend: reload UX при 409 `plan_version_conflict`

#### FE-409 — Optimistic lock UX

- **Где:**
  - `frontend/src/shared/lib/apiError.ts` — уже парсит `code`
  - `frontend/src/shared/api/httpClient.ts` — throws `ApiError` with `code`
  - Production mutations: `PlansList.tsx`, `DayDrawer.tsx`, `WorkCalendarEditor.tsx`, `CreatePlanWizard.tsx` (и shared hook если выделят)
  - `frontend/src/features/production/api/productionApi.ts`
- **Текущее состояние:** `apiError.test.ts` покрывает парсинг `plan_version_conflict`; UI не обрабатывает reload; `DayDrawer` проверяет `status === 409` только для delete track (generic).
- **Acceptance:**
  - [x] Shared helper `isPlanVersionConflict(error): boolean` (`frontend/src/shared/lib/planConflict.ts`).
  - [x] При 409 `plan_version_conflict` на мутациях плана (delete plan, delete track, complete day, save work calendar, build plan если applicable):
    - Показать понятное сообщение («План был изменён. Данные обновлены — повторите действие.»).
    - **Invalidate + refetch** production queries: `["production", "plans"]`, активный plan detail, calendar/day view по контексту.
    - Auto-refetch после сообщения.
  - [x] После refetch UI показывает актуальный `version` из API (`PlanMetaSummary.version`).
  - [x] Не ломать другие 409 (`day_already_completed`, `incomplete_return`) — обрабатывать **только** `code === "plan_version_conflict"`.
  - [x] Vitest: unit test helper + mock mutation handler.
- **Verify:**
  - `npm run test` — green
  - `npm run build` — green
  - Manual: два таба, concurrent edit → conflict → reload flow

---

### Optional WP3 — S4: npm CVE (deferred from mandatory closure)

#### S4 — npm supply chain

- **Acceptance (если WP3 выполнен):**
  - [ ] `npm audit` без critical/high для react-router, vite, undici (или documented exceptions).
  - [ ] Заменить `"latest"` pins в `frontend/package.json` на semver ranges или exact versions.
  - [ ] CI step `npm audit --audit-level=high` (Ask first if no CI yet — document in plan only).
- **Не блокирует** closure P2 WP0–WP2.

---

## 6. Testing Strategy

| Находка | Тип теста | Файл |
|---------|-----------|------|
| S2 | Unit/integration (HTTP) | `tests/test_auth_rate_limit.py` (новый) |
| S2 | Regression | `tests/test_app_session.py` |
| S3 | Authorization (HTTP + fixtures) | `tests/test_offers_authorization.py`, `tests/test_archive_authorization.py` (новые) |
| S3 | Regression | `tests/test_commercial_web_flow.py`, `tests/test_archive_endpoints.py` |
| FE-409 | Unit (Vitest) | `frontend/src/shared/lib/planConflict.ts` или расширение `apiError.test.ts` |
| FE-409 | Component (optional) | production feature tests |
| Регрессия | Full suite | `pytest tests/ -q` |

Safety-net: authorization tests **до** массового рефакторинга repository queries.

---

## 7. Boundaries

### Always
- Сначала failing test (S2 throttle, S3 cross-manager access, FE conflict helper), затем реализация.
- Object-level checks в **одном месте** (`app/dependencies/auth.py` или `app/security/offer_access.py`) — не копировать if-ы в каждый endpoint.
- Prefer filter-at-query (repository) over filter-in-Python для list endpoints (performance + correctness).
- Сохранить structured error format для 409/403 (`app/schemas/errors.py`).
- `pytest tests/ -q` зелёный перед closure.
- Документировать in-process rate limit limitation в spec/plan (single worker).

### Ask first
- Новая pip-зависимость (`slowapi`).
- 403 vs 404 для чужих КП (security through obscurity).
- Расширение S3 на legacy `app/web/router.py`.
- CI changes для `npm audit`.
- Изменение прав role `production`.

### Never
- Ослаблять S1 fail-closed bot auth (P1).
- Отключать auth на login rate limit в production via flag без документации.
- Коммитить без явной просьбы пользователя.
- Тащить Q5/Q6 в обязательный closure P2.
- Добавлять Redis dependency только для S2 в P2.

---

## 8. Success Criteria (спринт «готово»)

- [x] ASSUMPTIONS прочитаны и подтверждены (ownership → `owner_user_id`, 403 везде).
- [x] **WP0 закрыт:** S2 acceptance; `tests/test_auth_login_rate_limit.py` зелёный.
- [x] **WP1 закрыт:** S3 acceptance; authorization tests зелёные; manager isolation доказан тестами.
- [x] **WP2 закрыт:** FE-409 acceptance; `npm run test && npm run build` зелёные.
- [x] `pytest tests/ -q` зелёный — **756 passed, 12 skipped** (2026-06-20).
- [x] Health Score пересчитан в audit Post-P2 → S2/S3 RESOLVED; **~7/10**.
- [x] WP3 (S4) — optional; статус **deferred** (не блокирует closure P2).

### Deferred / out of closure scope

- **S4** npm CVE (optional WP3).
- **Q5** integration tests production API.
- **Q6** debug-instrumentation cleanup.

---

## 9. Risks & Mitigations

| Риск | Митигация |
|------|-----------|
| In-process rate limit обходится при multi-worker | Document single-instance; S4 backlog for Redis |
| S3 ломает существующие E2E tests с одним manager | Fixtures с двумя managers; admin bypass tests |
| `manager_id` null у legacy КП | Policy: admin-only access или orphan bucket; document in WP1 |
| FE reload race (double refetch) | Debounce invalidate; React Query `refetchType: 'active'` |
| 403 vs 404 inconsistency | Единый helper + table в WP1 |
| Scope creep в commercial.py | Strict endpoint list from audit S3 |

---

## 10. Decisions (предварительные — уточнить на ревью)

1. **Порядок спринта:** WP0 (S2) и WP1 (S3) можно параллелить; WP2 зависит от стабильного API 409 (уже есть) — параллельно с WP1 после контракта errors.
2. **Rate limiter:** custom in-process (как OCR), без `slowapi`, если не нужны global middleware features.
3. **Ownership key:** `manager_id` в metadata КП == `user["id"]` для role `manager`.
4. **403 для unauthorized object access** (единый стиль).
5. **S4 optional:** WP3 только после G2 (frontend) или в конце спринта.
6. **Legacy web UI:** out of scope unless actively used in production.

---

## Closure notes (2026-06-20)

### Реализовано

| WP | Модуль / артефакт | Тесты |
|----|-------------------|-------|
| WP0 S2 | `app/security/login_rate_limit.py`, wire в `auth.py` | `tests/test_auth_login_rate_limit.py` (+ регрессия session) |
| WP1 S3 | `app/security/offer_access.py`, filters в repositories, offers + archive endpoints | `tests/test_offers_authorization.py`, `tests/test_archive_authorization.py` |
| WP2 FE-409 | `frontend/src/shared/lib/planConflict.ts`, mutation `onError` в production UI | Vitest + `npm run build` |

### Известные gaps (вне обязательного closure P2)

| Gap | Риск | Рекомендация |
|-----|------|--------------|
| **Legacy web UI** (`app/web/router.py`) | Object-level RBAC не через `offer_access`; только role-level checks на части routes | Deprecate (A10) или портировать helper при активном использовании |
| **Telegram bot** | Нет `owner_user_id` / object-level RBAC parity; bot deprecated | Не расширять; при revival — thin adapter к `offer_access` |
| **Legacy КП без `owner_user_id`** | Admin-only access (policy в `offer_access`) | Backfill migration при необходимости |
| **S4 npm CVE** | Supply-chain high backlog | Отдельный chore / WP3 |

---

## Следующий шаг (после P2)

1. ~~Спринт quality~~ → **[`stabilizaciya-p3-audit-2026-06-20.md`](./stabilizaciya-p3-audit-2026-06-20.md)** (S1 legacy login, S4 admin guard, S5 production RBAC, Q-M9, Q6).
2. **S4** npm + CI audit — optional WP5 в P3.
3. Backlog: S10 password policy, S6 exception sanitization, strangler `cfg.PLATES_*`, A7 legacy web deprecate.
4. ~~Post-P2 секция в audit~~ — добавлена.

**PLAN:** [`../develop/plans/2026-06-19-bezopasnost-p2.md`](../develop/plans/2026-06-19-bezopasnost-p2.md) — **закрыт**
