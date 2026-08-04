# Spec: стабилизация P6 — web production readiness

> **Тип:** remediation feature-spec (архитектурный спринт)
> **Фаза SDD:** closed
> **Дата:** 2026-06-21
> **Ревизия:** v1 closed
> **Статус:** closed (WP9 stretch deferred)
> **Источник:** [`../develop/audits/2026-06-21-frontend-backend-audit.md`](../develop/audits/2026-06-21-frontend-backend-audit.md)
> **Predecessor (закрыт):** [`stabilizaciya-p5-architecture-2026-06-21.md`](./stabilizaciya-p5-architecture-2026-06-21.md) — WP1–WP5 closed; bot soft-decommission; god-module slices 1–2; CreatePlanWizard shell
> **Successor:** [`stabilizaciya-p7-architecture-2026-06-22.md`](./stabilizaciya-p7-architecture-2026-06-22.md) — god-modules closure, arch hygiene, medium backlog, infra stretch
> **Bot policy (наследуется):** [`stabilizaciya-p0-audit-2026-06-19.md`](./stabilizaciya-p0-audit-2026-06-19.md) § «Решение по Telegram-боту»
> **ADR:** [`../develop/architecture/core-viz-modules-boundary.md`](../develop/architecture/core-viz-modules-boundary.md), [`../develop/architecture/pep562-config-and-data-decommission.md`](../develop/architecture/pep562-config-and-data-decommission.md)

---

## Стратегия (одной фразой)

> Закрыть **3 Critical** и **quick-win High** из frontend-backend аудита на web path, довести инкрементальную декомпозицию god-modules (начатую в P5) до production-safe состояния, и зафиксировать **deploy contract** до миграции на Redis/PostgreSQL в P7.

---

## Маппинг ID (важно)

Нумерация **P6 backlog** использует **frontend-backend audit 2026-06-21** как primary source. P5 deferred items перенесены сюда с префиксом наследования.

| P6 WP / ID | Аудит frontend-backend | P5 (deferred / partial) | Статус до P6 |
|------------|------------------------|-------------------------|--------------|
| **WP1** HTTP CPU offload | **A1** sync CPU в workers | Out of scope P5 («audit A7 HTTP») | **Closed** — `run_cpu_bound`, async hot endpoints |
| **WP2** `visualize_plan` ports | **A2** app bypass ports | Partial P4/P5 — `CommercialService` на ports | **Closed** — `get_visualize_plan()` на app hot paths |
| **WP3** plate runtime web | **A3** TLS/ContextVar SSOT | Partial P3/P4 — middleware + `run_in_order_context` | **Closed** — explicit `plate_order_ctx`; `rg get_plate_mutable_runtime app/` → 0 |
| **WP4** guards + RBAC | **S3**, **S4** | S4 out of scope P5 (S8) | **Closed** — fail-closed guard; `RequireRole` routes |
| **WP5** god-modules slice 4–6 | **A4**, **A5**, **A6**, **A11** | P5 A7-1/2/3 partial | **Partial** — slice 5 (`core/kp/offers_write.py`); slices 4/6 deferred P7 |
| **WP6** arch hygiene | **A7**, **A8**, **A9**, **A12** | — | **Partial** — `AuthService` (A9); planning/DI/legacy deferred P7 |
| **WP7** rate limit + OCR policy | **A10**, **S1**, **S2** | P5 S1/S3 deferred | **Closed** — [`deploy-contract.md`](../develop/deploy-contract.md); `OCR_EXTERNAL_ENABLED` default false |
| **WP8** quality P1 | **Q1**–**Q4** | — | **Closed** — Q2/Q3 unit tests; Q4 hook tests (2 files) |
| **WP9** medium backlog | **A13**–**A19**, **S5**–**S12**, **Q5**–**Q13** | P5 Q8 partial | **Deferred** — stretch; не блокировал closure |
| **P7 preview** | **A20**–**A23**, **S13**–**S20**, **Q14**–**Q18** | P5 bot hard delete, full PEP562 | Backlog |

*Примечание: в full-repo аудите **A1/A2** = bot issues — **сняты P5 D1**. В frontend-backend lens **A1** = HTTP CPU — это **WP1 P6**.*

---

## Objective

После closure P5 (**914 passed**, bot archived, wizard shell **98 LOC**, layout/kp read slices):

1. Снять **3 Critical** web (CPU blocking, visualization ports, plate runtime coupling).
2. Закрыть **P0 security UX** (destructive DB guard, frontend route RBAC).
3. Продолжить god-module декомпозицию без big-bang.
4. Зафиксировать **deploy contract** (`workers=1` или shared rate limiter) до P7 infra.

**Целевой Health Score (frontend-backend lens):** с **0.0/10** к **~6.0–7.0/10** после WP1–WP4 (0 critical, ≤8 high open); **~7.5–8.0/10** после WP1–WP8.

**Verify baseline (P5 closure):** `pytest tests/ -q` → **≥914 passed**, 9 skipped, 0 failed; `cd frontend && npm run build` green.

---

## Контекст после P5

| P5 WP | Закрыто | Остаётся для P6 |
|-------|---------|-----------------|
| WP1 D1 | `bot_archived/`, stub `run_bot.py` | Hard delete `bot_archived/` → **P7** |
| WP2 A7-1 | `core/visualization/layout.py` | `visualize_plan`, drawing, export domains |
| WP3 A7-2 | `core/kp/offers_read.py` | Write/admin paths в `kp_db_offers.py` |
| WP4 Q6 | Shell + step components | `useCreatePlanWizardState` **442 LOC**; manual fill-mode AC |
| WP5 stretch | `CommercialWizardStepService`, Q8 partial | Workflow **724 LOC**; commercial/production `dict` responses |
| Deferred P5 | S1 Redis, S3 OCR, PostgreSQL, PEP562 full | **WP7, WP9, P7** |

**LOC snapshot (post-P5, 2026-06-21):**

| Модуль | LOC | Роль |
|--------|-----|------|
| `core/visualization/__init__.py` | ~628 | `visualize_plan`, orchestration, runtime hooks |
| `core/visualization/layout.py` | ~667 | Track/layout pure functions (P5 slice) |
| `core/kp_db_offers.py` | ~581 | Write/admin + re-export reads |
| `app/services/commercial_workflow_service.py` | ~724 | OCR, draft, export orchestration |
| `app/services/production_completion_service.py` | ~561 | Completion matching |
| `app/services/day_view_service.py` | ~522 | Day aggregation |
| `frontend/.../useCreatePlanWizardState.ts` | ~442 | Plan wizard state (новый god-hook) |
| `viz_modules/layout_sequence/builder.py` | ~965 | Layout sequence monolith |

**Уже есть (не дублировать в P6):**

- `archive_service.py`, `day_documents_service.py` — частично `asyncio.to_thread` / `run_in_order_context`
- `app/middleware/plate_runtime_isolation.py` — HTTP request isolation
- `core/plate_order_context.py` — `run_in_order_context` helper
- `CommercialService` — visualization ports (P4)
- `destructive_db_guard.py` — базовый guard (нужно ужесточить для non-local)

---

## Product decisions (accepted defaults — closure 2026-06-21)

| # | Вопрос | Принятое решение | Альтернатива (deferred) |
|---|--------|-------------------|-------------------------|
| **D1** | CPU-bound offload strategy | **`async def` + `run_cpu_bound`** (`asyncio.to_thread` + concurrency cap) на hot endpoints | Job API + polling → P7 при необходимости |
| **D2** | Rate limiting до Redis | **Deploy contract `workers=1`** — [`deploy-contract.md`](../develop/deploy-contract.md); Redis — **P7** | Redis в P6 WP7 |
| **D3** | OCR / OpenAI policy | **`OCR_EXTERNAL_ENABLED=false` default prod** + admin-only opt-in; on-prem OCR — P7 | Полный отказ от внешнего OCR в P6 |
| **D4** | Frontend RBAC model | **`RequireRole` wrapper** на `/new`, `/archive`, `/production`; redirect по `defaultRouteForRole` | Server-driven route manifest API |
| **D5** | Plate runtime refactor depth | **Slice:** explicit `PlateOrderContext` param в `visualize_plan`; web paths без raw TLS | Full removal TLS (P7) |
| **D6** | Bot hard delete | **Defer P7** — сохранить `bot_archived/` для reference | Delete в P6 |

**Обоснование D1:** минимальный diff, паттерн уже в `archive_service` / `day_documents_service`; job API — отдельная спека при необходимости UX polling.

---

## Выбор scope

| Трек | IDs | Effort | Решение |
|------|-----|--------|---------|
| **In scope P6 P0** | WP1 A1 HTTP CPU | M | Блокирует latency под нагрузкой |
| **In scope P6 P0** | WP2 A2 visualize_plan ports | M | Закрывает Critical + ADR boundary |
| **In scope P6 P0** | WP3 A3 plate runtime slice | L | Critical; инкрементальный slice |
| **In scope P6 P0** | WP4 S3 + S4 guards/RBAC | S–M | Quick security wins |
| **In scope P6 P1** | WP5 god-modules continuation | L | Продолжение P5 A7 |
| **In scope P6 P1** | WP6 planning/DI/Auth/legacy | L | Архитектурная гигиена |
| **In scope P6 P1** | WP7 rate limit contract + OCR flag | M | Deploy safety без Redis |
| **In scope P6 P1** | WP8 Q1–Q4 quality | M–L | DRY + hook tests |
| **Stretch P6 P2** | WP9 medium backlog | L | Не блокирует closure |
| **Deferred P7** | PostgreSQL, Redis shared store, hard delete bot, full PEP562, Argon2id, MFA, CSP enforce | L | Infra / compliance sprint |
| **Out of scope P6** | Новые product features, onboarding UX, job API (unless D1 rejected) | — | Отдельные спеки |

---

## Scope

### In scope (summary)

| Приоритет | WP | Проблема | Fix (кратко) |
|-----------|-----|----------|--------------|
| P0 | **WP1** | Sync CPU в HTTP workers | `async def` endpoints + `to_thread` / `run_in_order_context` |
| P0 | **WP2** | `visualize_plan` bypass ports | Port + adapter; app только через port |
| P0 | **WP3** | TLS plate runtime на web | Explicit context в `visualize_plan`; web paths без raw TLS |
| P0 | **WP4** | DB reset + frontend RBAC | Fail-closed guard; `RequireRole` routes |
| P1 | **WP5** | God modules (workflow, kp_db, viz, production) | Инкрементальные срезы 4–6 |
| P1 | **WP6** | Planning dup, DI, AuthService, legacy routes | SSOT + thin adapters |
| P1 | **WP7** | Rate limit + OCR | `workers=1` contract; `OCR_EXTERNAL_ENABLED` |
| P1 | **WP8** | Layout monolith, wizard DRY, hook tests | Фазы + unified validation + vitest |
| P2 | **WP9** | Medium audit items | PEP562 prep, DraftStore DI, CSP prep, audit log |

### Out of scope (explicit)

- **PostgreSQL** migration (audit A22)
- **Redis** shared rate limiter (полная реализация — P7; в P6 только contract)
- **Hard delete** `bot_archived/`
- **Full** PEP 562 `__getattr__` removal
- **CSP enforce** (S9) — подготовка в WP9, enforce P7
- **SQLCipher / encryption at rest** (S8)
- **Argon2id / MFA** (S16, backlog)
- **Новые product features**

---

## WP1 — HTTP CPU offload (audit A1)

**Цель:** ILP-оптимизация, plan build и тяжёлые preview не блокируют event loop / worker thread pool FastAPI.

**Effort:** M

**Проблема:** endpoints объявлены как sync `def` — весь запрос занимает worker целиком:

- `app/api/v1/endpoints/commercial.py` — `generate_preview`, `calculate_draft`, OCR paths
- `app/api/v1/endpoints/production.py` — `build_plan_from_filters`, `create_plan`
- `app/services/commercial_service.py` — `generate_preview` → `optimization_service.optimize`
- `app/services/production_service.py` — `build_plan_from_filters`

**Файлы:**

- `app/api/v1/endpoints/commercial.py` — `async def` + offload
- `app/api/v1/endpoints/production.py` — `async def` + offload
- `app/services/commercial_service.py` — sync core; вызывается из thread
- `app/services/production_service.py` — sync core; вызывается из thread
- `app/services/optimization_service.py` — ensure thread-safe / context propagation
- NEW (optional): `app/concurrency/cpu_bound.py` — shared executor + semaphore (max concurrent optimizations)
- `tests/test_commercial_web_flow.py`, `tests/test_production_api_integration.py`

**Паттерн (default D1):**

```python
# endpoint
@router.post("/generate-preview")
async def generate_preview(...) -> CommercialPreviewResponse:
    return await run_cpu_bound(
        lambda: service.generate_preview(...),
        plate_order_ctx=plate_order_ctx,
    )
```

Использовать `run_in_order_context` где нужен `PlateOrderContext.bound()` (как в `archive_service`).

**Acceptance:**

- [x] Hot endpoints (`generate_preview`, `build_plan_from_filters`, `calculate_draft`) — `async def`
- [x] CPU work выполняется вне event loop (`asyncio.to_thread` или shared executor)
- [x] `PlateOrderContext` propagates в worker thread на optimize/visualize paths
- [x] Concurrency cap (env `CPU_BOUND_MAX_CONCURRENT`, default 2) — optional but recommended
- [x] Integration tests green; нет регрессии response shape
- [x] Manual: два параллельных preview не деградируют health endpoint latency заметно

**Verify:**

```bash
pytest tests/test_commercial_web_flow.py tests/test_production_api_integration.py -q
pytest tests/test_plate_runtime_request_isolation.py -q
```

**Риски:**

| Риск | Митигация |
|------|-----------|
| ContextVar не propagates в thread | `run_in_order_context` / `ctx.bound()` wrapper |
| Thread pool exhaustion | Semaphore + documented cap |
| Starlette sync `def` в других routers | Grep audit; fix only listed hot paths in WP1 |

---

## WP2 — `visualize_plan` через ports (audit A2)

**Цель:** все app hot paths к визуализации идут через `core/ports/visualization`, не через `from core.visualization import visualize_plan`.

**Effort:** M

**Файлы:**

- `core/ports/visualization.py` — NEW protocol `VisualizePlanFn` + register/get
- `viz_modules/adapters/visualization_ports.py` — register implementation
- `app/adapters/visualization.py` — optional thin wrapper for app layer
- `app/services/file_generation_service.py`
- `app/services/archive_service.py`
- `app/services/day_documents_service.py`
- `tests/test_core_viz_import_boundary.py` — extend grep gate
- `ai_docs/develop/architecture/core-viz-modules-boundary.md` — update port map

**Acceptance:**

- [x] `VisualizePlanFn` в ports; wired at startup (`wire_visualization_ports`)
- [x] `rg "from core.visualization import visualize_plan" app/` → **0**
- [x] App services вызывают `get_visualize_plan()` (или injected port)
- [x] Import boundary tests green
- [x] PDF/XLSX generation behaviour unchanged (archive + day documents integration tests)

**Verify:**

```bash
rg "from core.visualization import visualize_plan" app/ --glob "*.py"
pytest tests/test_archive*.py tests/test_day_documents*.py tests/test_core_viz_import_boundary.py -q
```

**Риски:**

| Риск | Митигация |
|------|-----------|
| Signature drift port vs implementation | Typed protocol; single adapter |
| Missed call site | CI grep gate on `app/` |

---

## WP3 — Plate runtime web slice (audit A3)

**Цель:** `visualize_plan` и optimize paths на web принимают явный snapshot/context; снизить зависимость от `get_plate_mutable_runtime()`.

**Effort:** L (инкрементальный slice — не full TLS removal)

**Файлы:**

- `core/visualization/__init__.py` — `visualize_plan(..., plate_order_ctx: PlateOrderContext | None = None)`
- `core/plate_order_context.py` — snapshot builder for visualization
- `app/services/archive_service.py`, `day_documents_service.py`, `file_generation_service.py` — pass context
- `core/optimization/context.py` — document deprecated OPT_* proxies
- `app/services/optimization_service.py` — remove `legacy_runtime` path (или WP9)
- `tests/test_plate_order_context.py`, `tests/test_plate_runtime_request_isolation.py`

**Acceptance (slice):**

- [x] `visualize_plan` accepts optional `PlateOrderContext`; uses it when provided
- [x] Web services pass context explicitly (не полагаются только на middleware TLS)
- [x] `rg "get_plate_mutable_runtime" app/` → **0** (web layer)
- [x] Isolation tests green under concurrent requests
- [x] Deprecation comment on `get_plate_mutable_runtime()` for web paths in ADR

**Verify:**

```bash
pytest tests/test_plate_order_context.py tests/test_plate_runtime_request_isolation.py tests/test_plate_mutable_runtime_isolation.py -q
rg "get_plate_mutable_runtime" app/ --glob "*.py"
```

**Риски:**

| Риск | Митигация |
|------|-----------|
| Legacy core callers still on TLS | Incremental; core internal may use ctx.bound() |
| Bot archived code unchanged | Out of scope — web path only |

---

## WP4 — Security quick wins (audit S3, S4)

**Цель:** fail-closed destructive DB; role-based frontend routes.

**Effort:** S–M

### S3 — Destructive DB guard

**Файлы:**

- `core/destructive_db_guard.py`
- `app/main.py` — startup check when `APP_ENV=production`
- `app/api/v1/endpoints/admin.py`, `app/services/admin_service.py`
- `tests/test_destructive_db_guard.py` (new or extend)

**Acceptance:**

- [x] `APP_ENV=development` **не** разрешает reset без явного `ALLOW_DESTRUCTIVE_DB_RESET=1`
- [x] Staging/production требуют **оба** флага (уже частично есть — ужесточить dev)
- [x] Startup warning/error if production + destructive flags set
- [x] Admin endpoint tests cover blocked/allowed matrix

### S4 — Frontend `RequireRole`

**Файлы:**

- NEW: `frontend/src/features/auth/components/RequireRole.tsx`
- `frontend/src/app/router/AppRouter.tsx`
- `frontend/src/shared/lib/roleRoutes.ts` — `canAccessRoute(role, path)`
- NEW: `frontend/src/features/auth/components/RequireRole.test.tsx`
- `frontend/src/shared/lib/roleRoutes.test.ts` — extend

**Матрица (default D4):**

| Route | Roles |
|-------|-------|
| `/new` | admin, manager |
| `/archive` | admin, manager |
| `/production` | admin, production |

**Acceptance:**

- [x] `production` user → redirect с `/new` и `/archive` на `/production`
- [x] `manager` user → redirect с `/production` на `/new`
- [x] Unauthenticated → `/login` (existing `ProtectedRoute`)
- [x] Vitest tests для matrix
- [x] `npm run build` green

**Verify:**

```bash
pytest tests/test_admin*.py -q
cd frontend && npm run test -- --run RequireRole roleRoutes
cd frontend && npm run build
```

---

## WP5 — God-modules continuation (audit A4, A5, A6, A11)

**Цель:** продолжить P5 A7 slices без big-bang; каждый slice — отдельный mergeable PR.

**Effort:** L

### Slice 4 — `core/visualization` drawing/export

**Файлы:**

- NEW: `core/visualization/drawing.py` or `core/visualization/export.py`
- `core/visualization/__init__.py` — shrink to facade (< 400 LOC target)
- Tests: existing visualization/layout suite

**Acceptance:**

- [ ] Drawing/export helpers extracted; `__init__.py` LOC −200 minimum *(deferred — slice 4 → P7)*
- [ ] No new `viz_modules` imports in `core/` *(deferred — slice 4 → P7)*

### Slice 5 — `kp_db_offers` write/admin

**Файлы:**

- NEW: `core/kp/offers_write.py` (save, update, delete, clear)
- `core/kp_db_offers.py` — thin re-export facade
- `app/repositories/kp_repository.py`, `kp_offers_repository.py` — delegate

**Acceptance:**

- [x] Write functions single definition; `kp_db_offers.py` < 350 LOC
- [x] `tests/test_kp_db_*.py` green

### Slice 6 — `CommercialWorkflowService` OCR/draft/export

**Файлы:**

- NEW: `app/services/commercial_draft_lifecycle_service.py` (draft save/load/calculate orchestration)
- NEW: `app/services/commercial_offer_export_service.py` (xlsx/pdf export)
- `app/services/commercial_workflow_service.py` — target **< 500 LOC**

**Acceptance:**

- [ ] Draft lifecycle extracted; workflow = thin orchestrator *(deferred — slice 6 → P7)*
- [x] `tests/test_commercial*.py` green

### Slice 7 (stretch) — production god services

**Файлы:**

- `app/services/production_completion_service.py` — extract matching module
- `app/services/day_view_service.py` — fetch → normalize → aggregate phases

**Acceptance (stretch):**

- [ ] One production service −150 LOC via extraction
- [ ] Day view tests green

**Verify:**

```bash
pytest tests/test_commercial*.py tests/test_kp_db_*.py tests/test_layout_*.py tests/test_production*.py -q
```

---

## WP6 — Architecture hygiene (audit A7, A8, A9, A12)

**Цель:** SSOT planning, DI abstractions для ключевых сервисов, AuthService, legacy route deprecation.

**Effort:** L

### A7 — Planning SSOT

- `core/production/planning.py` — canonical
- `app/planning/` — thin adapters or deprecate modules with re-exports
- `app/services/plan_distribution_service.py` — delegate to core

### A8 — DI protocols

- NEW: `app/repositories/protocols.py` — `KpRepositoryProtocol`, `PlanRepositoryProtocol`, etc.
- `app/dependencies/services.py` — constructor injection; remove inline `AuthRepository()` from endpoints

### A9 — AuthService

- NEW: `app/services/auth_service.py` — authenticate, register, change_password
- `app/api/v1/endpoints/auth.py` — HTTP only

### A12 — Legacy web routes

- `app/web/legacy_routes.py` — POST handlers → 410/redirect; keep GET redirects only
- `app/web/legacy_deprecation.py` — telemetry log

**Acceptance:**

- [x] `AuthRepository()` not constructed in `auth.py` endpoints — `AuthService` via `Depends(get_auth_service)`
- [ ] `rg "from app.planning" app/services` — only adapters (or 0 after deprecate) *(deferred — A7 → P7)*
- [ ] Legacy POST `/web/login`, `/web/offers/new` removed or 410 *(deferred — A12 → P7)*
- [x] Auth + production integration tests green

**Verify:**

```bash
pytest tests/test_auth*.py tests/test_production_api_integration.py -q
rg "AuthRepository\(\)" app/api/ --glob "*.py"
```

---

## WP7 — Deploy contract + OCR policy (audit A10, S1, S2)

**Цель:** безопасный deploy при `workers > 1` до Redis; контролируемый внешний OCR.

**Effort:** M

### Rate limiting (D2 default)

**Файлы:**

- `app/main.py` — startup warning if `workers > 1` without `RATE_LIMIT_SHARED_STORE`
- NEW: `docs/deploy.md` or section in `ai_docs/` — **contract: `uvicorn --workers 1`** until Redis
- `app/security/login_rate_limit.py` — document per-process semantics

**Acceptance:**

- [x] Production deploy doc states `workers=1` requirement
- [x] Startup logs warning when violated
- [x] Optional: `RATE_LIMIT_SHARED_STORE=redis` stub raises NotImplemented → P7

### OCR policy (D3 default)

**Файлы:**

- `core/config/settings.py` — `ocr_external_enabled: bool = False`
- `app/api/v1/endpoints/commercial.py` — block parse if disabled
- `app/services/commercial_upload_validation.py` — respect flag
- `.env.example` — document

**Acceptance:**

- [x] Default prod: external OCR off → 503/403 with clear message
- [x] Admin can enable via env for staging only
- [x] Tests for disabled/enabled paths

**Verify:**

```bash
pytest tests/test_commercial*.py tests/test_login_rate_limit.py -q
```

---

## WP8 — Code quality P1 (audit Q1–Q4)

**Effort:** M–L

### Q1 — Layout builder phases

- `viz_modules/layout_sequence/builder.py` — extract phases (load → group → secondary → tracks)
- Target: no function > 80 LOC in new modules

### Q2 — Wizard validation unify

- `app/services/commercial_calculation_service.py` — single `validate_calculate_prerequisites` → errors list
- `commercial_wizard_step_service.py`, `commercial_workflow_service.py` — use unified validator
- NEW: `tests/test_commercial_calculation_service.py`

### Q3 — Execution terms DRY

- NEW: `core/execution_terms.py` — `parse_execution_terms(strict=...)`
- `archive_service.py`, `offers_service.py` — delegate

### Q4 — Frontend hook tests

- NEW: `frontend/src/features/production/hooks/useCreatePlanWizardState.test.ts`
- NEW: `frontend/src/features/commercial-offer/hooks/useCommercialOfferWizard.test.ts`
- Mock React Query; cover step transitions, conflict handling

**Acceptance:**

- [x] Q2/Q3 unit tests added
- [x] ≥2 hook test files; critical paths covered
- [x] `pytest` + `npm run test` green

---

## WP9 (stretch) — Medium backlog (audit A13–A19, S5–S12, Q5–Q13)

**Цель:** закрыть medium findings; не блокирует P6 closure.

**Effort:** L (optional)

| Cluster | IDs | Краткий fix |
|---------|-----|-------------|
| PEP562 prep | A13 | Direct imports; shrink `config_and_data` consumers |
| DraftStore DI | A14 | `get_draft_store()` Depends; no inline `DraftStore()` |
| response_model | A18 | Remaining `dict` on commercial/production |
| legacy_runtime | A19 | Remove `OptimizationService.legacy_runtime` |
| CSRF/CSP prep | S5, S9, S12 | CSRF prefetch; CSP roadmap |
| sessionStorage | S6 | draft_id only in client |
| audit log | S11 | Security logger for login/admin |
| Procurement DRY | Q5 | Shared breakdown pipeline |
| OffersService tests | Q13 | `tests/test_offers_service.py` |

**Acceptance (stretch):**

- [ ] ≥3 medium clusters closed with tests
- [ ] Changelog documents deferrals to P7

---

## Приоритеты и порядок работ

| Порядок | WP | Обоснование |
|---------|-----|-------------|
| 1 | **WP4** | S–M effort; security wins; независим от рефакторинга |
| 2 | **WP1** | Critical A1; разблокирует production load |
| 3 | **WP2** | Critical A2; ports — prerequisite для чистого WP3 |
| 4 | **WP3** | Critical A3; зависит от WP2 signature |
| 5 | **WP7** | Deploy contract параллельно с WP1 |
| 6 | **WP8 Q3, Q2** | Low-risk DRY |
| 7 | **WP5** | God-modules incremental |
| 8 | **WP6** | После стабилизации hot paths |
| 9 | **WP8 Q1, Q4** | Layout + frontend tests |
| 10 | **WP9** | Stretch |

**Параллелизация:** WP4 frontend RBAC ∥ WP4 backend guard ∥ WP7 deploy doc. WP5 slices — отдельные PR после WP1–WP3.

---

## Definition of Done (спринт P6)

- [x] WP1–WP4 acceptance выполнены (**обязательно для closure**)
- [x] WP5 slice 4 **или** slice 5 выполнен (минимум один god-module срез) — **slice 5** (`core/kp/offers_write.py`)
- [x] WP7 deploy contract задокументирован
- [x] Product decisions **D1–D5** зафиксированы в changelog
- [x] `pytest tests/ -q` — **953 passed**, 0 failed
- [x] `cd frontend && npm run build && npm run test` — green (**55** frontend tests)
- [x] Health Score frontend-backend lens — **≥6.0/10** (0 critical)
- [x] Manual: wizard fill mode; parallel preview smoke
- [x] Spec status → `closed` or `closed (WP9 stretch deferred)`
- [x] Cross-ref в [`2026-06-21-frontend-backend-audit.md`](../develop/audits/2026-06-21-frontend-backend-audit.md)

---

## Риски (спринт)

| Риск | Вероятность | Митигация |
|------|-------------|-----------|
| WP3 scope creep → full TLS removal | Medium | Explicit slice acceptance; P7 for full removal |
| Async migration breaks sync tests | Medium | Incremental endpoints; keep service sync |
| God-module refactor regressions | Medium | One slice per PR; full test suite |
| OCR disable breaks staging workflow | Low | Feature-flag; document admin override |
| Workers>1 in prod despite contract | Medium | Startup warning + deploy doc review |

---

## P7 preview (следующий спринт)

| Тема | Audit IDs | Примечание |
|------|-----------|------------|
| Hard delete `bot_archived/` | P5 D6 defer | После 30d без rollback need |
| Full PEP 562 removal | A13 | `core/config_and_data.py` delete `__getattr__` |
| Redis rate limiting | S1, A10 | Shared store |
| PostgreSQL | A22 | Migration layer |
| CSP enforce | S9 | Nonce/hash for Vite |
| SQLCipher / at-rest encryption | S8 | Compliance |
| Argon2id password hash | S16 | Migration on login |
| Job API for long optimize | A1 alt | If to_thread insufficient |

---

## Changelog

| Дата | Изменение |
|------|-----------|
| 2026-06-21 | v1 draft — P6 после closure P5; scope: WP1–WP4 P0, WP5–WP8 P1, WP9 stretch; baseline 914 passed |
| 2026-06-21 | **P6 closed (WP9 stretch deferred)** — verify: **953 pytest passed**, **55 frontend tests**, build OK |
| 2026-06-21 | **Closed WPs:** WP1 (A1 CPU offload), WP2 (A2 visualization ports), WP3 (A3 plate runtime slice), WP4 (S3 destructive guard + S4 `RequireRole`), WP7 (A10/S1 deploy contract + S2 OCR flag), WP8 (Q2–Q4 quality) |
| 2026-06-21 | **Partial:** WP5 slice 5 (`core/kp/offers_write.py`); WP6 `AuthService` (A9) — endpoints via `Depends(get_auth_service)` |
| 2026-06-21 | **Deferred → P7:** WP9 medium backlog (A13–A19, S5–S12, Q5–Q13); WP5 slices 4/6 (visualization drawing/export, workflow OCR/draft/export); WP6 planning SSOT (A7), DI protocols (A8), legacy routes (A12) |
| 2026-06-21 | **Product decisions accepted (defaults):** D1 `run_cpu_bound`/`asyncio.to_thread`; D2 `workers=1` deploy contract; D3 `OCR_EXTERNAL_ENABLED=false`; D4 `RequireRole` route guards; D5 explicit `PlateOrderContext` slice |
| 2026-06-22 | **P7 spec created** — successor [`stabilizaciya-p7-architecture-2026-06-22.md`](./stabilizaciya-p7-architecture-2026-06-22.md); deferred items formalized as WP1–WP5; plan: [`../develop/plans/p7-stabilization-plan-2026-06-22.md`](../develop/plans/p7-stabilization-plan-2026-06-22.md) |

---

*Создано: 2026-06-21 · v1 closed: web production readiness, 0 critical on web path; Health Score frontend-backend lens ~6.5–7.0/10; successor P7: god-modules, arch hygiene, Redis/ADR, medium backlog.*
