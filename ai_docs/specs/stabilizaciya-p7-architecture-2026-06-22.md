# Spec: стабилизация P7 — god-modules closure, arch hygiene, medium backlog

> **Тип:** remediation feature-spec (архитектурный спринт)
> **Фаза SDD:** draft
> **Дата:** 2026-06-22
> **Ревизия:** v1 draft
> **Статус:** draft
> **Источник:** [`../develop/audits/2026-06-21-frontend-backend-audit.md`](../develop/audits/2026-06-21-frontend-backend-audit.md) (Remediation status + open items)
> **Predecessor (закрыт):** [`stabilizaciya-p6-architecture-2026-06-21.md`](./stabilizaciya-p6-architecture-2026-06-21.md) — WP1–WP4, WP7–WP8 closed; WP5/6 partial; WP9 stretch deferred
> **Successor:** — (draft)
> **Bot policy (наследуется):** [`stabilizaciya-p0-audit-2026-06-19.md`](./stabilizaciya-p0-audit-2026-06-19.md) § «Решение по Telegram-боту»
> **ADR:** [`../develop/architecture/core-viz-modules-boundary.md`](../develop/architecture/core-viz-modules-boundary.md), [`../develop/architecture/pep562-config-and-data-decommission.md`](../develop/architecture/pep562-config-and-data-decommission.md), [`../develop/architecture/rate-limiting.md`](../develop/architecture/rate-limiting.md)
> **Deploy contract (наследуется):** [`../develop/deploy-contract.md`](../develop/deploy-contract.md)

---

## Стратегия (одной фразой)

> Закрыть **оставшиеся High** из frontend-backend аудита (god-modules, planning/DI/legacy), инкрементально снять **medium backlog** (WP9 P6), и подготовить **infra-ready** контракты для Redis/PostgreSQL/CSP — без big-bang и без регрессии **953 pytest / 55 frontend tests**.

---

## Маппинг ID (важно)

Нумерация **P7 backlog** наследует deferred items из P6 changelog + P7 preview + open rows в audit Remediation status.

| P7 WP / ID | Аудит frontend-backend | P6 (deferred / partial) | Статус до P7 |
|------------|------------------------|-------------------------|--------------|
| **WP1** | **S1**, **A10** | WP7 deploy contract only (`workers=1`) | **Open** — mitigated, not resolved |
| **WP2** | **A4**, **A6**, **A11** (partial) | WP5 slices 4/6 deferred | **Open** — god-modules continuation |
| **WP3** | **A7**, **A8**, **A12** | WP6 partial (A9 `AuthService` closed) | **Open** — planning/DI/legacy |
| **WP4** | **A13**–**A19**, **S5**–**S12**, **Q5**–**Q13** | WP9 stretch deferred | **Open** — medium clusters |
| **WP5** | **A22**, **S8**, **S9**, **S16**, P5 **D6** | P6/P7 preview infra stretch | **Backlog** — infra / compliance |
| **P8 preview** | **A20**–**A23**, **S13**–**S20**, **Q14**–**Q18** | Low backlog | Future |

*Примечание: **A9** AuthService закрыт в P6 WP6. **A15** (frontend RBAC) закрыт в P6 WP4 (`RequireRole`).*

---

## Objective

После closure P6 (**953 passed**, **55** frontend tests, **0 critical** на web path):

1. Завершить инкрементальную декомпозицию god-modules (P6 WP5 slices 4/6 + production path).
2. Закрыть архитектурную гигиену: planning SSOT, DI protocols, legacy route removal.
3. Закрыть ≥3 medium clusters из WP9 (PEP562 prep, DraftStore DI, response_model, …).
4. Либо реализовать Redis shared rate limiting, либо зафиксировать **infra prerequisite** и ADR (если Redis недоступен в sprint).
5. Опционально (stretch): PostgreSQL preview, bot hard delete, CSP enforce, Argon2id migration path.

**Целевой Health Score (frontend-backend lens):** с **~6.5–7.0/10** к **~7.5–8.5/10** после WP2–WP4; **~8.5+/10** при WP1 Redis + WP5 stretch.

**Verify baseline (P6 closure):** `pytest tests/ -q` → **953 passed**, 9 skipped, 0 failed; `cd frontend && npm run build && npm run test` → **55** tests green.

---

## Контекст после P6

| P6 WP | Закрыто | Остаётся для P7 |
|-------|---------|-----------------|
| WP1 A1 | `run_cpu_bound`, async hot endpoints | Job API (alt) → P8 preview |
| WP2 A2 | `get_visualize_plan()` port | — |
| WP3 A3 | explicit `PlateOrderContext` slice | Full TLS removal → P8 |
| WP4 S3/S4 | destructive guard + `RequireRole` | — |
| WP5 | slice 5 (`offers_write.py`) | slices 4/6 (viz drawing/export, workflow OCR/draft/export); slice 7 production |
| WP6 | `AuthService` (A9) | A7 planning SSOT, A8 DI protocols, A12 legacy routes |
| WP7 | deploy contract + OCR flag | Redis implementation (WP1) |
| WP8 | Q2–Q4 | Q1 layout builder phases (partial) |
| WP9 | — | A13–A19, S5–S12, Q5–Q13 → **WP4** |

**LOC snapshot (post-P6, 2026-06-22):**

| Модуль | LOC (approx) | Роль |
|--------|--------------|------|
| `core/visualization/__init__.py` | ~593 | `visualize_plan`, orchestration |
| `core/visualization/layout.py` | ~667 | layout pure functions (P5 slice) |
| `core/kp_db_offers.py` | ~350 | thin facade + re-exports (P6 slice 5) |
| `app/services/commercial_workflow_service.py` | ~674 | OCR, draft, export orchestration |
| `app/services/production_completion_service.py` | ~561 | Completion matching |
| `app/services/day_view_service.py` | ~522 | Day aggregation |
| `frontend/.../useCreatePlanWizardState.ts` | ~442 | Plan wizard state hook |
| `viz_modules/layout_sequence/builder.py` | ~965 | Layout sequence monolith (Q1) |

**Уже есть (не дублировать в P7):**

- `app/concurrency/cpu_bound.py`, `run_cpu_bound`
- `core/ports/visualization.py`, `app/adapters/visualization.py`
- `app/services/auth_service.py`, `Depends(get_auth_service)`
- `core/kp/offers_read.py`, `core/kp/offers_write.py`
- `frontend/.../RequireRole.tsx`, `roleRoutes.ts`
- [`deploy-contract.md`](../develop/deploy-contract.md) — `workers=1` until Redis

---

## Product decisions (требуют подтверждения на ревью)

| # | Вопрос | Default для draft | Альтернатива |
|---|--------|-------------------|--------------|
| **D1** | Redis в P7 | **Implement WP1** если Redis URL доступен в staging; иначе **ADR-only** + оставить `workers=1` | Edge rate limiting (nginx) без app changes |
| **D2** | God-module slice order | **WP2 first:** viz drawing/export (slice 4) → workflow OCR/draft/export (slice 6) | Workflow first (higher business churn) |
| **D3** | Planning SSOT migration | **Deprecate `app/planning/`** с re-exports → `core/production/planning.py`; adapters в `app/services/` | Big-bang delete `app/planning/` |
| **D4** | Legacy web routes | **410 POST** + redirect GET only; telemetry 30d → delete handlers | Keep POST with deprecation headers |
| **D5** | Bot hard delete | **Defer to WP5 stretch** — 30d после P5 archive без rollback need | Delete in WP5 mandatory |
| **D6** | PostgreSQL | **Preview only** — connection settings + migration spike doc; no production cutover in P7 | Full migration in P7 |
| **D7** | PEP 562 | **Prep phase:** direct imports + shrink consumers; physical `__getattr__` delete — end of WP4 or P8 | Full delete in WP4 mandatory |

**Обоснование D2:** slice 4 (viz) имеет сильный test net (`test_layout_*`, viz boundary); slice 6 (workflow) зависит от стабильных draft/export contracts — меньше coupling если viz сначала.

---

## Выбор scope

| Трек | IDs | Effort | Решение |
|------|-----|--------|---------|
| **In scope P7 P1** | WP2 god-modules 4/6 | L | Продолжение P5/P6 A7 |
| **In scope P7 P1** | WP3 A7, A8, A12 | L | Архитектурная гигиена |
| **In scope P7 P1** | WP4 medium clusters (≥3) | M–L | Закрыть WP9 P6 deferred |
| **In scope P7 P0/P1** | WP1 S1, A10 | M | Redis **или** infra prerequisite ADR |
| **Stretch P7 P2** | WP5 PostgreSQL preview, bot delete, CSP, Argon2id | L | Infra / compliance |
| **Out of scope P7** | Job API + polling, MFA, SQLCipher production, full PostgreSQL cutover | — | P8 / отдельные спеки |

---

## Scope

### In scope (summary)

| Приоритет | WP | Проблема | Fix (кратко) |
|-----------|-----|----------|--------------|
| P0/P1 | **WP1** | Per-process rate limits при `workers > 1` | Redis store **или** ADR + enforce deploy contract |
| P1 | **WP2** | God modules viz + workflow (+ optional production) | Slices 4/6 (+7 stretch) |
| P1 | **WP3** | Planning dup, DI, legacy routes | SSOT + protocols + 410 legacy POST |
| P1/P2 | **WP4** | Medium backlog WP9 | PEP562, DraftStore DI, response_model, audit log, … |
| P2 | **WP5** | Infra stretch | PostgreSQL preview, bot hard delete, CSP enforce, Argon2id path |

### Out of scope (explicit)

- **Full** PostgreSQL production migration (A22 complete)
- **Job API** для long optimize (A1 alt) — unless `to_thread` proven insufficient
- **Full** TLS/plate runtime removal (A3 remainder)
- **MFA**, **SQLCipher** production rollout
- **Новые product features**

---

## WP1 — Redis / shared rate limiting (audit S1, A10)

**Цель:** снять deploy constraint `workers=1` **или** формально зафиксировать infra prerequisite до появления Redis.

**Effort:** M (implementation) / S (ADR-only path)

**Проблема:** `login_rate_limit.py` и OCR upload limits — in-process; при `workers > 1` лимиты умножаются.

**Файлы:**

- `app/security/login_rate_limit.py` — `RedisRateLimitStore` or protocol + in-memory fallback
- `app/services/commercial_upload_validation.py` — shared store injection
- `app/main.py` — wire store from `RATE_LIMIT_REDIS_URL`; remove `NotImplementedError` stub for `RATE_LIMIT_SHARED_STORE=redis`
- `core/config/settings.py` — `rate_limit_redis_url`, `rate_limit_shared_store`
- NEW or update: `ai_docs/develop/architecture/rate-limiting.md`
- `tests/test_login_rate_limit.py`, `tests/test_rate_limit_deployment.py`
- `.env.example`

**Паттерн (implementation path):**

```python
# settings
rate_limit_shared_store: Literal["memory", "redis"] = "memory"
rate_limit_redis_url: str | None = None

# startup: if workers > 1 and store == memory → fail-fast or warning (existing)
# if store == redis → connect pool, health check
```

**Acceptance (implementation):**

- [ ] `RATE_LIMIT_SHARED_STORE=redis` + valid URL → limits shared across processes (integration test with 2 fake workers or redis fixture)
- [ ] `workers=1` + memory store — unchanged behaviour
- [ ] `deploy-contract.md` updated: Redis path documented
- [ ] Startup fail-fast or clear error when `workers > 1` without redis

**Acceptance (ADR-only fallback — if D1 rejected):**

- [ ] ADR states Redis as **required** before `workers > 1`
- [ ] CI/deploy checklist references contract
- [ ] No silent regression — existing warning preserved

**Verify:**

```bash
pytest tests/test_login_rate_limit.py tests/test_rate_limit_deployment.py tests/test_password_change_rate_limit.py -q
```

**Риски:**

| Риск | Митигация |
|------|-----------|
| Redis unavailable in dev | Docker compose optional; memory fallback for `workers=1` |
| Race in redis INCR | Use atomic operations; document TTL semantics |

---

## WP2 — God-modules continuation (audit A4, A6, A11)

**Цель:** закрыть P6 WP5 deferred slices; снизить LOC и SRP violations на hot paths.

**Effort:** L

### Slice 4 — `core/visualization` drawing/export

**Файлы:**

- NEW: `core/visualization/drawing.py` and/or `core/visualization/export.py`
- `core/visualization/__init__.py` — shrink to facade (**< 400 LOC** target)
- `tests/test_layout_*.py`, `tests/test_core_viz_import_boundary.py`

**Acceptance:**

- [ ] Drawing/export helpers extracted; `__init__.py` LOC −200 minimum
- [ ] No new `viz_modules` imports in `core/`
- [ ] Visualization tests green

### Slice 6 — `CommercialWorkflowService` OCR/draft/export

**Файлы:**

- NEW: `app/services/commercial_draft_lifecycle_service.py`
- NEW: `app/services/commercial_offer_export_service.py`
- `app/services/commercial_workflow_service.py` — target **< 500 LOC**
- `app/dependencies/services.py` — wire new services
- `tests/test_commercial*.py`

**Acceptance:**

- [ ] Draft lifecycle extracted; workflow = thin orchestrator
- [ ] Export paths isolated; OCR policy flag respected
- [ ] `tests/test_commercial_web_flow.py` green

### Slice 7 (stretch) — production god services

**Файлы:**

- `app/services/production_completion_service.py` — extract matching module
- `app/services/day_view_service.py` — fetch → normalize → aggregate phases

**Acceptance (stretch):**

- [ ] One production service −150 LOC via extraction
- [ ] Day view / completion tests green

**Verify:**

```bash
pytest tests/test_commercial*.py tests/test_layout_*.py tests/test_core_viz_import_boundary.py tests/test_production*.py tests/test_day_view_service.py -q
```

---

## WP3 — Architecture hygiene (audit A7, A8, A12)

**Цель:** SSOT planning, DI protocols, legacy route deprecation.

**Effort:** L

### A7 — Planning SSOT

- `core/production/planning.py` — canonical
- `app/planning/` — deprecate with re-exports; `app/services/plan_distribution_service.py` delegates to core
- `app/services/production_planning_service.py` — thin adapter

### A8 — DI protocols

- NEW: `app/repositories/protocols.py` — `KpRepositoryProtocol`, `PlanRepositoryProtocol`, …
- `app/dependencies/services.py` — constructor injection for `CommercialService`, `CommercialWorkflowService`, `ProductionPlanningService`
- Remove inline `DraftStore()` from endpoints (overlap WP4 A14)

### A12 — Legacy web routes

- `app/web/legacy_routes.py` — POST handlers → **410 Gone** or redirect; GET redirects only
- `app/web/legacy_deprecation.py` — structured telemetry log
- `tests/test_web_*.py` or extend auth tests

**Acceptance:**

- [ ] `rg "from app.planning" app/services` — only adapters (or **0** after deprecate)
- [ ] Repository protocols used in ≥2 service constructors via `Depends`
- [ ] Legacy POST `/web/login`, `/web/offers/new` → 410 or removed; GET redirects preserved
- [ ] Auth + production integration tests green

**Verify:**

```bash
pytest tests/test_auth*.py tests/test_production_api_integration.py tests/test_plan_consistency.py -q
rg "from app.planning" app/services --glob "*.py"
rg "AuthRepository\(\)" app/api/ --glob "*.py"
```

---

## WP4 — Medium backlog clusters (audit A13–A19, S5–S12, Q5–Q13)

**Цель:** закрыть ≥3 medium clusters из P6 WP9; остальное — documented deferrals to P8.

**Effort:** M–L

| Cluster | IDs | Краткий fix | Priority |
|---------|-----|-------------|----------|
| PEP562 prep | **A13** | Direct imports; shrink `config_and_data` consumers; optional `__getattr__` delete (D7) | P1 |
| DraftStore DI | **A14** | `get_draft_store()` Depends; no inline `DraftStore()` | P1 |
| response_model | **A18** | Pydantic schemas for `/parse`, `/generate-preview`, `get_plan`, `activate_plan` | P1 |
| legacy_runtime | **A19** | Remove `OptimizationService.legacy_runtime` | P2 |
| CSRF/CSP prep | **S5**, **S9**, **S12** | CSRF prefetch on LoginPage; CSP roadmap (enforce → WP5) | P2 |
| sessionStorage | **S6** | `draft_id` only in client | P2 |
| audit log | **S11** | Security logger for login/admin/destructive | P2 |
| Procurement DRY | **Q5** | Shared breakdown pipeline | P2 |
| OffersService tests | **Q13** | `tests/test_offers_service.py` | P2 |
| Wizard hook split | **A16** | Split `useCreatePlanWizardState` (optional) | P2 |
| Repository SQL | **A17** | Batch queries; no `_connect` from app | P2 |

**Acceptance (minimum for P7 closure):**

- [ ] ≥3 clusters closed with tests (recommended: **A13 prep**, **A14**, **A18**)
- [ ] Changelog documents remaining deferrals → P8
- [ ] No regression: full pytest + frontend test suite green

**Verify:**

```bash
pytest tests/test_offers_service.py tests/test_commercial*.py -q  # after Q13/A18
rg "DraftStore\(\)" app/api/ --glob "*.py"  # target → 0
```

---

## WP5 (stretch) — Infra & compliance preview

**Цель:** подготовить следующий infra sprint без production cutover.

**Effort:** L (optional — не блокирует P7 closure)

| Item | IDs | Краткий fix |
|------|-----|-------------|
| PostgreSQL preview | **A22** | Settings + Alembic spike; dual-read doc only |
| Bot hard delete | P5 **D6** | Remove `bot_archived/`; grep gate; update ADR |
| CSP enforce | **S9** | Nonce/hash for Vite; move Report-Only → enforce in staging |
| Argon2id path | **S16** | Hash on login migration; settings flag |
| SQLCipher note | **S8** | ADR for at-rest encryption — no implementation |

**Acceptance (stretch):**

- [ ] ≥1 infra item delivered (recommended: **bot hard delete** or **CSP staging enforce**)
- [ ] PostgreSQL spike doc in `ai_docs/develop/architecture/` if attempted

---

## Приоритеты и порядок работ

| Priority | WP | Обоснование |
|----------|-----|-------------|
| **P1** | **WP2** | Продолжение P6; закрывает High A4/A6; без infra deps |
| **P1** | **WP3** | Foundation для DI и planning; разблокирует WP4 A17 |
| **P1** | **WP4** (A13, A14, A18) | Quick-medium wins; улучшает OpenAPI и draft consistency |
| **P0/P1** | **WP1** | Параллельно с WP2 **если** Redis доступен; иначе ADR в начале спринта |
| **P2** | **WP4** (остальные clusters) | S11, S6, Q5, Q13, A19 |
| **P2** | **WP5** | Stretch после closure criteria |

**Рекомендуемый порядок спринта:**

| Порядок | WP | Checkpoint |
|---------|-----|------------|
| 0 | **Git hygiene** | Commit uncommitted P3–P6 work; tag baseline `p6-closure` |
| 1 | **WP2 slice 4** | Viz drawing/export extracted; layout tests green |
| 2 | **WP2 slice 6** | Workflow < 500 LOC; commercial tests green |
| 3 | **WP3** | Planning SSOT + legacy 410; integration tests green |
| 4 | **WP4** (A14, A18, A13) | ≥3 medium clusters |
| 5 | **WP1** | Redis **or** ADR signed off |
| 6 | **WP5** | Stretch if time |

**Параллелизация:** WP2 slice 4 ∥ WP4 A18 (different layers). WP1 blocked on infra decision (D1).

---

## Definition of Done (спринт P7)

- [ ] WP2 slice 4 **и** slice 6 acceptance выполнены (**обязательно для closure**)
- [ ] WP3 acceptance выполнен (**обязательно**)
- [ ] WP4 — ≥3 medium clusters closed
- [ ] WP1 — Redis implemented **или** ADR infra prerequisite accepted (D1)
- [ ] Product decisions **D1–D7** зафиксированы в changelog
- [ ] `pytest tests/ -q` — **≥953 passed**, 0 failed
- [ ] `cd frontend && npm run build && npm run test` — green (**≥55** tests)
- [ ] Health Score frontend-backend lens — **≥7.5/10** (0 critical, reduced high)
- [ ] Cross-ref в [`2026-06-21-frontend-backend-audit.md`](../develop/audits/2026-06-21-frontend-backend-audit.md) Remediation status
- [ ] Spec status → `closed` or `closed (WP5 stretch deferred)`

---

## Риски (спринт)

| Риск | Вероятность | Митигация |
|------|-------------|-----------|
| WP2+WP3 scope creep | Medium | One slice per PR; strict acceptance per slice |
| Redis infra not ready | High | D1 ADR-only path; keep `workers=1` |
| Legacy route removal breaks bookmarked URLs | Low | GET redirects preserved; telemetry before delete |
| PEP562 delete breaks imports | Medium | Prep phase first; grep gate on consumers |
| Uncommitted P3–P6 confuses baseline | High | Git hygiene checkpoint before WP2 |

---

## P8 preview (следующий спринт)

| Тема | Audit IDs | Примечание |
|------|-----------|------------|
| Full PostgreSQL migration | A22 | After P7 preview |
| Job API for long optimize | A1 alt | If CPU offload insufficient |
| Full plate runtime / TLS removal | A3 remainder | Explicit context everywhere |
| MFA, idle session timeout | S15 | Security hardening |
| Low backlog | A20–A23, S13–S20, Q14–Q18 | Incremental |

---

## Changelog

| Дата | Изменение |
|------|-----------|
| 2026-06-22 | v1 draft — P7 после closure P6; scope: WP1 Redis/ADR, WP2 god-modules 4/6, WP3 arch hygiene, WP4 medium backlog, WP5 infra stretch; baseline 953 pytest / 55 frontend |

---

*Создано: 2026-06-22 · v1 draft: god-modules closure, arch hygiene, medium backlog; наследование P6 deferred items + audit open High/Medium.*
