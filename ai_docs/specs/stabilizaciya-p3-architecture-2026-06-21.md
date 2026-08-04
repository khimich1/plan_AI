# Spec: стабилизация P3 — web/core после аудита 2026-06-21

> **Тип:** remediation feature-spec (архитектурный спринт)
> **Фаза SDD:** CLOSED
> **Дата:** 2026-06-21
> **Ревизия:** v1 closed (WP5 deferred)
> **Статус:** closed (WP5 stretch deferred)
> **Источник:** [`../develop/audits/2026-06-21-full-project-audit.md`](../develop/audits/2026-06-21-full-project-audit.md)
> **Predecessor (закрыт):** [`stabilizaciya-p2-architecture-2026-06-20.md`](./stabilizaciya-p2-architecture-2026-06-20.md) — WP1–WP6
> **Bot policy (наследуется):** [`stabilizaciya-p0-audit-2026-06-19.md`](./stabilizaciya-p0-audit-2026-06-19.md) § «Решение по Telegram-боту»
> **Idea one-pager:** [`../ideas/post-p2-p3-direction.md`](../ideas/post-p2-p3-direction.md)

---

## Стратегия (одной фразой)

> Закрыть **P0 web/core gaps** после P2: thin repository (A2), DI для offers/managers (A6 gap), первый срез инверсии `core`↔`viz_modules` (A1), security quick wins (S2, S7, S9-residual); bot **не трогаем**.

---

## Маппинг ID аудитов (важно)

Нумерация в аудите **2026-06-21** не совпадает с аудитом **2026-06-20** / P2.

| Аудит 2026-06-21 | Аудит 2026-06-20 / P2 | Статус до P3 |
|------------------|------------------------|--------------|
| **A1** core↔viz_modules | **A8** (deferred P2) | Open |
| **A2** PlanRepository orchestration | **A10** partial (entry point only) | Open — другой DoD |
| **A3** DI offers/managers | **A6** partial | Open — gap P2 scope |
| **A4** global plate runtime | **A3** phase 3+ partial | Open — proxy active |
| A1 split-brain SQLite/JSON | **A1** | ✅ Closed P0-next |
| A2 bot planning bypass | **A2** | ✅ Closed (bot deprecated) |

---

## Bot deprecation strategy

| Тема | Решение |
|------|---------|
| **Статус** | Bot **deprecated**. Активный канал — **web (React SPA + FastAPI API)**. |
| **Вне scope P3** | Bot god-modules, `bare except`, bot tests (Q4–Q10), bot→app DIP (A8 audit-21) |
| **PEP 562** | Web/core: довести proxy removal где безопасно; bot `config_and_data` — только minimal fix при поломке тестов |
| **`run_bot.py`** | Не использовать в production; удаление `bot/` — отдельное согласование |

---

## Objective

Закрыть критический и high-priority **web/core** backlog из аудита 2026-06-21 после closure P2. Цель — устранить архитектурные блокеры релиза (A1, A2), закрыть gap DI (A3), применить low-effort security fixes (S2, S7, S9-residual). **≥947 passed, 12 skipped** — baseline не должен регрессировать.

## Контекст после P2

| P2 WP | Закрыто | Остаётся для P3 |
|-------|---------|-----------------|
| WP1 A3 | Hot paths off PEP 562; DeprecationWarning на proxy | Полное удаление proxy; bot residual |
| WP2 A6 | DI: production, commercial, admin, archive | **offers.py, managers.py** inline `Service()` |
| WP3 A10 | `ProductionPlanningService` — canonical entry | **Thin PlanRepository** — оркестрация в service |
| WP4 A5 | Router split; draft/export services | Дальнейшая декомпозиция — stretch |
| WP5 S5 | Server-side RBAC + tests | Frontend role guards — UX only (не security WP) |
| WP6 | Health redacts `environment`/`app` | **`rate_limiting` metadata** в public health |

**Health Score:** P2 closure ~8.5–9/10 (web/core scope). Аудит 2026-06-21 full-repo = **2.0/10** — не откат P2, другой lens.

---

## Product decisions (требуют подтверждения на ревью)

| # | Вопрос | Default для draft | Альтернатива |
|---|--------|-------------------|--------------|
| D1 | `GET /managers` для role `production` | **Accepted** (как в P2 WP5 inventory) | Restrict PII (audit S10) |
| D2 | Multi-worker deployment в ближайшем релизе | **No** → S1 Redis defer | Yes → S1 в P0 |
| D3 | Thin repository (A2) | **In scope P3** | Отложить — только документировать границы |

---

## Выбор scope

| Трек | IDs (audit 2026-06-21) | Effort | Решение |
|------|------------------------|--------|---------|
| **In scope (P3 P0)** | **A1** (slice), **A2**, **A3**, **S2**, **S7** | M–L | Блокеры / quick wins web path |
| **In scope (P3 P1)** | **A5** (sqlite3→repo), **S9-residual**, **A4** (proxy removal slice) | M | Maintainability + hardening |
| **Stretch** | **S1** Redis rate limit (если D2=Yes), **Q1–Q3** DRY pricing | M | По решению D2 |
| **Deferred** | Bot Q4–Q10, **A8** bot↔app, **A9** scaling/PostgreSQL, **Q6** CreatePlanWizard, **S3** OCR consent, **S6** CSP enforce | L | P4+ или отдельные спеки |
| **Out of scope** | Полное удаление `bot/`, MFA, encryption at rest, frontend role guards как security | L | Отдельное согласование |

---

## Scope

### In scope

| Приоритет | ID | Проблема | Fix (кратко) |
|-----------|-----|----------|--------------|
| P0 | **A3** | `offers.py`, `managers.py` — inline `OffersService()` / `CommercialService()` | `get_offers_service`, `get_managers_service` в `app/dependencies/services.py`; все handlers через `Depends` |
| P0 | **A2** | `PlanRepository` (~645 LOC) — business orchestration | Вынести track distribution / aggregation orchestration в service; repository = CRUD + SQL |
| P0 | **A1** (slice 1) | `core/` импортирует `viz_modules/` | Protocol/ports в `core/`; первый adapter в `app/` или `viz_modules/`; CI grep/lint gate `core/` ↛ `viz_modules/` |
| P0 | **S2** | Swagger/OpenAPI публичен в production | `docs_url=None`, `redoc_url=None`, `openapi_url=None` при `app_env=production` |
| P0 | **S7** | Регистрация принимает произвольные role strings | Pydantic validator / enum whitelist (`admin`, `manager`, `production`) |
| P1 | **A5** | Raw `sqlite3.connect` в `core/production/planning.py` | SQL через `PlanRepository` / port |
| P1 | **S9-residual** | `/health` раскрывает `rate_limiting` deployment metadata | Минимизировать public payload; детали — internal endpoint или auth |
| P1 | **A4** (slice) | PEP 562 proxy active | Удалить/sузить `__getattr__` для web-only callers; ADR update |

### Deferred

- **S1** in-process rate limiting → Redis (unless D2=Yes)
- **S3** OCR third-party consent / data minimization
- **S6** CSP enforce mode
- **S8** frontend role-based route guards (UX; server RBAC уже есть)
- **S10** managers PII для production (unless D1=revisit)
- **A7, Q4–Q6** god-module декомпозиция (bot, frontend wizard)
- **A9, S14** PostgreSQL, shared DraftStore

### Out of scope (explicit)

- Новые product features
- Bot reliability и commercial pipeline consolidation
- Полное удаление `config_and_data.py` module file
- Decommission `legacy_routes` / `bot/`

---

## WP1 — A3 gap: DI для offers и managers

**Цель:** `rg "Service\(\)" app/api/v1/endpoints/` → **0**.

**Файлы:**
- `app/dependencies/services.py` — `get_offers_service()`, при необходимости thin wrapper для managers
- `app/api/v1/endpoints/offers.py` — заменить 8× `OffersService()` на `Depends(get_offers_service)`
- `app/api/v1/endpoints/managers.py` — `Depends(get_commercial_service)` или dedicated factory
- Тесты: расширить endpoint integration / `dependency_overrides` pattern

**Acceptance:**
- [x] Factory functions в `app/dependencies/services.py`
- [x] Ни один handler в `offers.py` / `managers.py` не вызывает `Service()` в body
- [x] `rg "Service\(\)" app/api/v1/endpoints/ --glob "*.py"` → 0
- [x] Override pattern документирован (как WP2 A6)

**Verify:**
```bash
pytest tests/test_offers*.py tests/test_*managers* tests/test_commercial*.py -q
rg "Service\(\)" app/api/v1/endpoints/ --glob "*.py"
```

---

## WP2 — A2: thin PlanRepository

**Цель:** repository содержит только persistence; orchestration — в service layer.

**Контекст P2:** WP3 закрыл **единый entry** (`ProductionPlanningService`). P3 закрывает **разделение ответственности** repository vs service.

**Файлы:**
- `app/repositories/plan_repository.py` — CRUD + SQL only (~280 LOC)
- `app/services/plan_distribution_service.py` — track add/remove orchestration, multi-plan aggregation
- `app/services/production_planning_service.py` — canonical entry; persist via `PlanPersistAdapter`
- `app/planning/plan_distribution.py`, `plan_aggregation.py` — helpers (pure или service-called)
- Тесты: `tests/test_plan_repository.py`, `tests/test_production_planning*.py`, lifecycle integration

**Acceptance:**
- [x] `PlanRepository` не импортирует `add_tracks_to_plan` для business rules inline в persist methods
- [x] Build/add-tracks flow: API → ProductionPlanningService → helpers → PlanRepository CRUD
- [x] `PlanRepository` < ~300 LOC **или** documented split (repository + persistence helpers module)
- [x] Integration test create → build → activate без регрессии track count / version
- [x] ADR-абзац или update P2 orchestration diagram

**Verify:**
```bash
pytest tests/test_production_planning*.py tests/test_plan_*.py tests/test_plan_lifecycle_create_build_activate_orchestration.py -q
```

---

## WP3 — A1 slice 1: core ↛ viz_modules

**Цель:** первый инкрементальный срез инверсии зависимости; gate на новые нарушения.

**Файлы (приоритет):**
- `core/visualization.py` — убрать direct imports `viz_modules/*` (или вынести facade)
- `core/production/planning.py` — `build_layout_sequence` через port
- Новый: `core/ports/visualization.py` (Protocol) + adapter в `app/adapters/` или `viz_modules/adapters/`
- CI/docs: grep check или import-linter contract

**Acceptance:**
- [x] ADR: граница core/viz_modules, список портов
- [x] Минимум **2** бывших `core→viz_modules` import path переведены на port+adapter
- [x] `rg "from viz_modules|import viz_modules" core/ --glob "*.py"` — сокращение зафиксировано в spec changelog (target: 0 для hot paths)
- [x] Существующие layout/procurement tests зелёные

**Verify:**
```bash
pytest tests/test_layout_*.py tests/test_procurement_*.py tests/test_production_planning*.py -q
rg "viz_modules" core/ --glob "*.py"
```

---

## WP4 — Security quick wins (S2, S7, S9-residual)

**Цель:** закрыть high/medium security без Redis/CSP full rollout.

**Файлы:**
- `app/main.py` — conditional OpenAPI docs off in production
- `app/schemas/auth.py` — role whitelist on register
- `app/api/v1/endpoints/health.py` — trim `rate_limiting` details in production public response
- Тесты: `tests/test_rate_limit_deployment.py`, новый/расширенный test для docs disabled + role validation

**Acceptance:**
- [x] Production: `/docs`, `/redoc`, `/openapi.json` недоступны (404 или disabled)
- [x] `RegisterUserRequest.role` — только допустимые значения; invalid → 422
- [x] Production `GET /health` и `GET /api/v1/health` — без deployment-sensitive `rate_limiting` internals (или redacted subset)
- [x] Non-production dev/staging behaviour сохранён или документирован

**Verify:**
```bash
pytest tests/test_rate_limit_deployment.py tests/test_auth*.py -q
```

---

## WP5 (P1) — A5 + A4 slice

**Цель:** SQL только через repository; следующий шаг PEP 562 decommission.

**Файлы:**
- `core/production/planning.py` — `_load_kp_list` и аналоги через repository/port
- `core/config_and_data.py` — сузить proxy (web callers only)
- `ai_docs/develop/architecture/pep562-config-and-data-decommission.md` — update status

**Acceptance (optional P1 — не блокирует closure P3 если WP1–WP4 green):**
- [ ] Нет новых raw `sqlite3.connect` в `core/production/planning.py` для plan flows
- [ ] ADR checklist updated with remaining proxy consumers count

---

## Приоритеты и порядок работ

| Порядок | WP | ID | Обоснование |
|---------|-----|-----|-------------|
| 1 | WP1 | A3 | Быстрый win; закрывает gap P2 A6 |
| 2 | WP4 | S2, S7, S9 | Low effort security |
| 3 | WP2 | A2 | Зависит от стабильного DI; высокий architecture impact |
| 4 | WP3 | A1 | После A2 — меньше coupling при рефакторинге viz |
| 5 | WP5 | A5, A4 | Stretch P1 |

**Параллелизация:** WP1 и WP4 независимы — можно параллельно.

---

## Definition of Done (спринт)

- [x] WP1–WP4 acceptance выполнены (**обязательно**)
- [x] WP5 — optional; не блокирует closure (deferred)
- [x] **Нет** bot-specific WP
- [x] Product decisions D1–D3 зафиксированы в changelog
- [x] `pytest tests/ -q` — **955 passed, 12 skipped** (target ≥947)
- [x] Spec status → `closed (WP5 deferred)`
- [x] Audit 2026-06-21 обновлён: cross-ref P3 closure items

---

## Риски

| Риск | Вероятность | Митигация |
|------|-------------|-----------|
| A2 refactor меняет plan output | Medium | Lifecycle integration test; compare track count / version |
| A1 port extraction ломает visualization | Medium | Incremental; run layout + procurement suites each step |
| S2 docs off ломает staging workflows | Low | Env-specific: disable only `production` |
| Scope creep в bot | Low | Explicit out-of-scope; grep gate on `bot/handlers` |

---

## Следующий шаг

1. **Ревью v1** этой спеки — подтвердить D1–D3 и порядок WP.
2. **IMPLEMENT P3:** WP1 (A3) ∥ WP4 (S2/S7/S9) → WP2 (A2) → WP3 (A1) → WP5 stretch.
3. **P4 preview:** S1 Redis, S3 OCR consent, Q6 CreatePlanWizard, A9 PostgreSQL, bot removal decision.

---

## Changelog

| Дата | Изменение |
|------|-----------|
| 2026-06-21 | v1 draft — P3 после closure P2 + audit 2026-06-21 |
| 2026-06-21 | WP2 closed — `PlanDistributionService` + thin `PlanRepository` (~280 LOC) |
| 2026-06-21 | WP3 A1 slice 1 — `core/ports/visualization.py` + adapters; `planning.py` и `visualization.py` hot paths off direct `viz_modules` (`build_layout_sequence`, `load_price_table_from_xlsx`); grep gate `tests/test_core_viz_import_boundary.py` |
| 2026-06-21 | **P3 closure** — WP1–WP4 complete; WP5 (A5, A4 slice) deferred; product decisions D1 (managers for production — accepted), D2 (multi-worker — no, S1 defer), D3 (thin repo — in scope, done); `pytest tests/ -q` → **955 passed, 12 skipped, 0 failed** |

---

*Создано: 2026-06-21 · v1: web/core P0 gaps, bot frozen, ID mapping audit 2026-06-20 ↔ 2026-06-21. Closed 2026-06-21 (WP1–WP4).*
