# Spec: Стабилизация P3 — security gaps + quality (аудит 2026-06-20)

> **Тип:** remediation feature-spec (стабилизационный спринт P3)
> **Фаза SDD:** SPECIFY → **PLAN** (решения §10 закрыты 2026-06-20)
> **Дата:** 2026-06-20
> **Ревизия:** v2 (decisions closed)
> **Статус:** **closed (implemented)** — closure **2026-06-20**
> **Baseline:** [`project-baseline.md`](./project-baseline.md)
> **Предшественник:** [`bezopasnost-p2-audit-2026-06-19.md`](./bezopasnost-p2-audit-2026-06-19.md) — **закрыт**
> **Источник находок:** [`../develop/audits/2026-06-20-full-project-audit.md`](../develop/audits/2026-06-20-full-project-audit.md)
> **PLAN:** [`../develop/plans/2026-06-20-stabilizaciya-p3.md`](../develop/plans/2026-06-20-stabilizaciya-p3.md)

---

## Стратегия (одной фразой)

> Закрыть **web-facing security gaps** после P2 (legacy login bypass, production RBAC, destructive admin + SQLite plans reset) и заложить **safety net** integration-тестами — без bot/critical-кластеров A1/A2/A3.

**Контекст после P2:** Health Score **~7/10** (аудит 19.06, Post-P2). REST login rate limit, object-level RBAC для `manager`, FE-409 — в проде.

---

## Remediation context (маппинг ID)

| Аудит 20.06 | Scope P3 (WP) |
|-------------|---------------|
| **S1** — обход rate limit `/web/login` | **WP0** |
| **S5** — production читает все КП / коммерческие данные | **WP1** |
| **S4** — destructive admin без guard; reset планов не чистит SQLite | **WP2** |
| **Q-M9** — integration tests production API | **WP3** |
| **Q6** — debug-instrumentation | **WP4** |
| **A7** — deprecate legacy web | **Deferred** (после WP0) |

### Accepted residual risk (вне P3)

- **A1, A2** — bot deprecated; web authority — SQLite.
- **A3** — hot paths изолированы (P1); full decommission — отдельный спринт.
- **S2** multi-worker rate limit — single instance (documented).

---

## ASSUMPTIONS (подтверждены 2026-06-20)

1. **P0, P1, P2 закрыты** — `pytest tests/ -q` ~756 passed; React SPA — основной UI.
2. **Telegram-бот deprecated** — не в проде.
3. **Деплой — single instance**; in-process rate limit достаточен.
4. **Legacy `/web/login`** остаётся; полный deprecate — backlog **A7** (effort M).
5. **Destructive guard** использует существующий флаг **`ALLOW_DESTRUCTIVE_DB_RESET`** (новый `ALLOW_DESTRUCTIVE_ADMIN` не вводим).
6. **Production role** — только операционный контур `/production/*`; commercial API — 403.

---

## 1. Objective & Problem Statement

### Objective

Закрыть **S1, S4, S5** аудита 20.06 для web-пути, исправить **reset_plans_only** (SQLite), добавить integration safety net.

### Reframe: success criteria

| Требование | Критерий |
|------------|----------|
| Нет bypass login | `POST /web/login` — тот же rate limit, что REST login |
| Production изолирован | `/offers`, `/commercial/archive`, PDF/XLSX — **403**; работа только через `/production/*` |
| Admin reset безопасен | Reset только при `APP_ENV=development` **или** `ALLOW_DESTRUCTIVE_DB_RESET=1` |
| Планы сбрасываются полностью | `reset_plans_only` чистит **`production_plans`** в SQLite, не только legacy JSON |
| Safety net | Integration tests на критичные production routes |

---

## 2. Scope

### In scope — обязательный closure P3

| WP | ID | Название |
|----|-----|----------|
| **WP0** | S1 | Rate limit на `POST /web/login` |
| **WP1** | S5 | Production RBAC + frontend UX |
| **WP2** | S4 | Destructive admin guard + SQLite plans reset |
| **WP3** | Q-M9 | Production API integration tests |
| **WP4** | Q6 | Debug-instrumentation cleanup |

### Optional

| WP | Название |
|----|----------|
| **WP5** | npm CVE (`npm audit fix`) |
| **WP6** | `list_users()` cache (A4) |

### Out of scope

- Deprecate legacy web UI целиком (**A7** — отдельный спринт).
- Bot consolidation, A3 full decommission, Redis/PostgreSQL.

---

## 3. Tech Stack

Без изменений. Ключевое:

- **WP0:** `app/security/login_rate_limit.py` → `app/web/router.py`
- **WP1:** `offer_access.py`, `offers.py`, `archive.py` (verify), `LoginPage`, `AppHeader`, `AppRouter`
- **WP2:** `core/destructive_db_guard.py`, `admin.py`, `admin_service.py`, `PlanRepository`
- **WP3–WP4:** pytest, Vitest (регрессия)

---

## 4. Commands

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

## 5. Acceptance Criteria

### WP0 — S1: Rate limit на legacy `POST /web/login`

- **Где:** `app/web/router.py` — `login_submit`
- **Acceptance:**
  - [x] В начале handler — `check_login_rate_limit(request)` (5 req/min/IP → 429 + `Retry-After`).
  - [x] Тест: 6-й POST `/web/login` → 429.
  - [x] `tests/test_auth_login_rate_limit.py` зелёный.
- **Не в scope:** deprecate `/web/login` — backlog **A7**.
- **Verify:** `tests/test_web_login_rate_limit.py`

---

### WP1 — S5: Production operational-only + frontend

#### Backend

| Область | Решение |
|---------|---------|
| Commercial API | Роль `production` → **403** на `GET/POST/PATCH/DELETE /api/v1/offers/*` и `/api/v1/commercial/archive/*` |
| PDF/XLSX | **403** для `production` на `/{kp_id}/pdf`, `/{kp_id}/xlsx` |
| Operational contour | Только `/api/v1/production/*` (`kp-candidates`, планы, дни, календарь) |
| Kp candidates | Только КП со статусом **`в работе`** (убрать `выполнено` из `_PRODUCTION_READ_STATUSES` в `offer_access.py` и filter в `list_kps_in_production` при необходимости) |

- **Acceptance:**
  - [x] `require_roles` на offers убирает `production` **или** единый guard → 403.
  - [x] Archive endpoints — `production` не в allowed roles (уже `admin`, `manager` — проверить регрессию).
  - [x] `_PRODUCTION_READ_STATUSES = frozenset({"в работе"})` (без `выполнено`).
  - [x] Тесты: production user → 403 на offers list/get/pdf; 200 на `/production/kp-candidates`.
  - [x] Manager/admin — без регрессии P2.

#### Frontend

- [x] Скрыть «Создать КП» и «Архив» в `AppHeader` для `user.role === "production"`.
- [x] После login и при уже залогиненном user — redirect **`/production`** (не `/new`).
- [x] `AppRouter`: index route и `*` fallback для production → `/production`; manager/admin — как сейчас (`/new`).
- [x] Vitest smoke для redirect helper (optional).

- **Verify:** `tests/test_offers_production_authorization.py`; `npm run test && npm run build`

---

### WP2 — S4: Destructive admin guard + SQLite plans reset

#### Guard (все reset-операции)

- **Где:** `app/api/v1/endpoints/admin.py`, `core/destructive_db_guard.py`
- **Правило:** разрешено только если `APP_ENV == "development"` **ИЛИ** `ALLOW_DESTRUCTIVE_DB_RESET=1`.
- **Затронутые endpoints:** `reset_full`, `reset_kp_only`, **`reset_plans_only`**, **`reset_calendar_only`**.
- **Acceptance:**
  - [x] `reset_plans_only` / `reset_calendar_only` оборачивают `require_destructive_db_reset()` (как full/kp-only).
  - [x] При блокировке → structured error через `raise_destructive_db_blocked_error` (без утечки env secret в client body).
  - [x] Уточнить `destructive_db_reset_allowed()`: **staging без флага — deny** (не только `production`).
  - [x] Тест: `APP_ENV=production`, без flag → 403 на plans-only и calendar-only.
  - [x] Тест: `APP_ENV=development` → reset работает.

#### SQLite plans reset

- **Где:** `app/services/admin_service.py`, `app/repositories/plan_repository.py`
- **Текущий баг:** `reset_plans_only` чистит только legacy JSON (`_clear_all_plans`), не `production_plans`.
- **Acceptance:**
  - [x] `PlanRepository.delete_all_plans()` — удаляет все строки из `production_plans` (транзакционно).
  - [x] `reset_plans_only` вызывает SQLite clear **и** legacy JSON cleanup (best-effort).
  - [x] Обновить `test_reset_plans_only_does_not_touch_db` → assert `production_plans` count = 0; KP tables не тронуты.
  - [x] `reset_full` также очищает SQLite plans (если ещё не делает).

- **Verify:** `tests/test_admin_destructive_guard.py`, `tests/test_admin_service.py`

---

### WP3 — Q-M9: Integration tests production API

- [x] `tests/test_production_api_integration.py` — ≥8 routes happy path; ≥3 failure modes (401, 403, 409).
- [x] Shared fixtures из `conftest.py`.

---

### WP4 — Q6: Debug-instrumentation cleanup

- [x] Удалить или gate `#region agent log` в `day_view_service.py`, `production_planning_service.py`.
- [x] ripgrep в `app/services/` → 0 ungated agent log blocks.

---

## 6. Testing Strategy

| WP | Файл |
|----|------|
| WP0 | `tests/test_web_login_rate_limit.py` |
| WP1 | `tests/test_offers_production_authorization.py` |
| WP2 | `tests/test_admin_destructive_guard.py`, `tests/test_admin_service.py`, `tests/test_destructive_db_guard.py` |
| WP3 | `tests/test_production_api_integration.py` |
| WP4 | regression production/day tests |

---

## 7. Boundaries

### Always
- Переиспользовать `login_rate_limit.py`, `destructive_db_guard.py`, `offer_access.py`.
- Fail-closed destructive ops по умолчанию.
- Не ослаблять P1/P2 closure.

### Ask first
- Deprecate `/web/login` (A7).
- Ужесточение guard для staging (если ломает существующие deploy scripts).
- CI `npm audit` (WP5).

### Never
- Новый env `ALLOW_DESTRUCTIVE_ADMIN`.
- Bot consolidation в P3.
- Коммит без явной просьбы.

---

## 8. Success Criteria (спринт «готово»)

- [x] **WP0–WP4** по acceptance; тесты зелёные.
- [x] `pytest tests/ -q` + frontend test/build green.
- [x] Post-P3 секция в audit (remediation delta).
- [x] WP5 — optional, **deferred** (npm CVE backlog).
- [x] WP6 — optional, **implemented** (indexed `get_user_by_id` lookup, closure 2026-06-20).

**Target Health Score:** ~7.5–8/10 (S1/S4/S5 сняты; A1/A2/A3 — accepted residual).

**Closure date:** 2026-06-20

---

## 9. Risks & Mitigations

| Риск | Митигация |
|------|-----------|
| Production не может сверить выполненные заказы | Operational-only by design; redacted summary — backlog при запросе |
| Staging deploy ломается без flag | Документировать `ALLOW_DESTRUCTIVE_DB_RESET=1` для миграций |
| reset_plans_only ломает активный план | Только admin; guard + confirm в UI (уже есть) |
| Legacy web login 429 ломает скрипты | Тот же лимит, что REST; документировать |

---

## 10. Decisions (закрыты 2026-06-20)

| Тема | Решение |
|------|---------|
| **WP0** | Rate limit на `POST /web/login` **сейчас**; полный deprecate `/web/login` — backlog **A7** (effort M) |
| **Production scope** | **Operational-only:** статус только **`в работе`**; `/production/*` — да; `GET /offers`, archive, PDF/XLSX — **403**; фронт скрывает «Создать КП»/«Архив», redirect → `/production` |
| **Destructive admin** | `APP_ENV=development` **ИЛИ** `ALLOW_DESTRUCTIVE_DB_RESET=1` на **все** reset; без нового флага; `reset_plans_only` чистит **SQLite `production_plans`**, не только JSON |
| **Порядок WP** | WP0 ∥ WP1 ∥ WP2 параллельно → WP3 → WP4 |
| **Health Score** | Post-P3 delta в audit, не переписывать audit 20.06 целиком |

---

## Следующий шаг

1. ~~Ревью §10~~ — **закрыто**.
2. ~~**PLAN**~~ — **закрыто**.
3. ~~**IMPLEMENT** WP0–WP4~~ — **закрыто 2026-06-20**.
4. **Backlog:** A7 (legacy web deprecate), A3 phase 2, WP5 (npm CVE).
