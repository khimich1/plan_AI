# Spec: стабилизация P2 — архитектура web/core (аудит 2026-06-20)

> **Тип:** remediation feature-spec (архитектурный спринт)
> **Фаза SDD:** SPECIFY (черновик на ревью)
> **Дата:** 2026-06-20
> **Ревизия:** v1
> **Статус:** closed
> **Источник:** [`../develop/audits/2026-06-20-full-project-audit.md`](../develop/audits/2026-06-20-full-project-audit.md)
> **Predecessor (закрыт):** [`stabilizaciya-p1-next-audit-2026-06-20.md`](./stabilizaciya-p1-next-audit-2026-06-20.md) — WP1–WP4 (web/API security)
> **Successor (draft):** [`stabilizaciya-p3-architecture-2026-06-21.md`](./stabilizaciya-p3-architecture-2026-06-21.md) — gaps после full audit 2026-06-21
> **Bot policy (наследуется):** [`stabilizaciya-p0-audit-2026-06-19.md`](./stabilizaciya-p0-audit-2026-06-19.md) § «Решение по Telegram-боту»

---

## Стратегия (одной фразой)

> Снизить архитектурный долг **web/core** после закрытия P0/P1-next: decommission PEP 562 proxy (A3 phase 3+), единый DI (A6), консолидация planning orchestration (A10), первый срез декомпозиции god-модулей web-side (A5) и server-side RBAC re-check (S5); bot **не трогаем**.

---

## Bot deprecation strategy

| Тема | Решение |
|------|---------|
| **Статус** | Bot **deprecated** с 2026-06-19. Единственный активный канал — **web (React SPA + FastAPI API)**. |
| **Что оставить** | `bot/` заморожен (read-only archive). SQLite authority через `plan_storage` shim — без изменений. |
| **Что не трогать** | A4 bot→app DIP, bot god-modules (`commercial.py`, `production_execution.py`), Q1/Q3 bot tests, A9 bot commercial pipeline — **вне scope** P2. |
| **`run_bot.py`** | WARNING при старте; не использовать в production. Полное удаление `bot/` — **отдельное согласование**, не блокер P2. |
| **PEP 562 / config_and_data** | Миграция **web/core hot paths** в приоритете; bot-импорты `config_and_data` — только если ломаются тесты (минимальный fix, без рефакторинга bot handlers). |

---

## Objective

Закрыть **архитектурный и security-backlog кластер P2** из аудита 2026-06-20 после стабилизации data plane (P0-next) и web/API security (P1-next). Цель — убрать legacy globals proxy, выровнять DI, упростить planning orchestration, начать декомпозицию web god-модулей и усилить server-side RBAC. **927 passed, 12 skipped** — baseline не должен регрессировать.

## Контекст закрытых спринтов

| Спринт | Закрыто | Остаётся для P2 |
|--------|---------|-----------------|
| P0-next | A1/S1, A2, A3 phase 1–2 (`PlateOrderContext`, load codes) | A3 phase 3+ (PEP 562) |
| P1-next | S4 POST logout, S6 destructive guard, S2 `session_version`, S3 variant B | S5 RBAC re-check; S7 XFF — review; S9-audit health leak |
| Post-sprint | Health Score **8.5–9/10** | A5, A6, A10, medium clusters |

**S2 (session invalidation):** в P1-next закрыт через `session_version` + bump on logout (`fb122d4`). В P2 **не** открываем заново, кроме регрессий.

---

## Выбор scope: architecture vs security vs bot

| Трек | IDs | Effort | Решение |
|------|-----|--------|---------|
| **In scope (P2)** | **A3 phase 3+**, **A6**, **A10**, **A5** (web-side) | M–L | Прямой maintainability/scalability risk на активном web/core path |
| **In scope (P2 security)** | **S5** | M | Client-only frontend guards ≠ security boundary |
| **Stretch / optional** | **S9-audit**, **S7** (review), **ADR** (WP1 prep) | S | Low effort, закрывает audit residual |
| **Deferred** | **A4, A5 bot**, **A8** core→viz_modules, **A9**, **Q1–Q8 bot** | M–L | Bot frozen; A8 — отдельный спринт после A3 |
| **Out of scope** | Полное удаление `bot/`, PostgreSQL migration (A15), frontend god components (A17) | L | Отдельное согласование / P3+ |

---

## Scope

### In scope

| Приоритет | ID | Проблема | Fix (кратко) |
|-----------|-----|----------|--------------|
| P0 | **A3 phase 3+** | PEP 562 `__getattr__` proxy в `core/config_and_data.py` — скрытый global state | Инкрементальный decommission: явные импорты из `core.plate_runtime_state`, `core.domain`, `core.config`; удалить/сузить proxy |
| P0 | **A6** | Непоследовательный DI: `production.py`/`commercial.py` — `Service()` inline; `admin.py`/`archive.py` — `Depends(get_*_service)` | Единые factory в `app/dependencies/`; все endpoints через `Depends` |
| P1 | **A10** | Triple planning orchestration: `plan_manager`, `plan_distribution`, `production_planning_service` | Один orchestration entry в service; utilities в `core/production/` |
| P1 | **A5** (web) | God-модули web-side: `app/web/router.py` (~939), `commercial_workflow_service.py` (~1103) | Первый срез: routes/forms vs business logic; целевой модуль <600 LOC за итерацию |
| P1 | **S5** | Frontend RBAC — client-only; API должен быть авторитетом | Inventory sensitive actions; server re-check + тесты 403 |
| P2 (stretch) | **S9-audit** | `/health` раскрывает `environment`, `app` | Public: `{"status":"ok"}` + безопасные metadata; env — internal/auth |
| P2 (stretch) | **S7** | XFF hardening review | Подтвердить `trusted_proxy_ips` в deployment; расширить coverage при необходимости |
| P2 (stretch) | **ADR** | Checklist decommission PEP 562 (наследие P1-next WP5) | `ai_docs/develop/architecture/` — шаги миграции и rollback |

### Deferred

- **A8** — инверсия `core` ↔ `viz_modules` (после A3 phase 3+)
- **A5 bot** — `bot/handlers/commercial.py`, `production_execution.py`
- **A11** — JSON backdoor (закрыт в P0-next для plan_storage; полный audit repository — backlog)
- **A17** — `CreatePlanWizard.tsx` god component
- **S10–S19** medium security cluster (кроме S9-audit stretch)

### Out of scope (explicit)

- Новые product features
- Bot reliability (Q1, Q3), bot commercial consolidation (A9, Q6)
- Полный decommission `config_and_data` module file (только proxy + migration callers)
- CSP enforce (S11), OpenAPI lockdown (S10)

---

## WP1 — A3 phase 3+: PEP 562 decommission (incremental)

**Цель:** убрать скрытый доступ к mutable runtime через `from core import config_and_data as cfg` + PEP 562 `__getattr__` на **web/core hot paths**.

**Файлы (приоритет миграции):**
- `core/config_and_data.py` — сузить/удалить `__getattr__` proxy
- `app/services/production_planning_service.py`, `core/production/planning.py`
- `viz_modules/procurement/adapters_default.py`, `core/plan_commit.py`
- Тесты: `tests/test_config_and_data_plate_naming.py`, `tests/test_plate_runtime_*.py`, `tests/test_procurement_*.py`

**Acceptance:**
- [x] ADR или checklist в `ai_docs/develop/architecture/` — порядок миграции, список оставшихся proxy-атрибутов, критерий полного удаления
- [x] Hot paths **app/** и **core/production/** не используют PEP 562 mutable proxy (`PLATES_*`, `PLATE_*` через `cfg.X =`)
  - [x] **app/services/** (step 1): `commercial_service`, `plate_parser_service` — явные импорты; `production_planning_service`, `optimization_service` — уже на `PlateOrderContext`; `rg "config_and_data as cfg" app/` → 0
  - [x] **core/production/planning.py** (step 2): explicit `normalize_load_code` from `core.domain.plate_order`; `rg "config_and_data as cfg" core/production/planning.py` → 0
  - [x] **core/plan_commit.py** (step 2b): explicit `normalize_load_code` from `core.domain.plate_order`; `rg "config_and_data as cfg" core/plan_commit.py` → 0
  - [x] **viz_modules/procurement/** (step 3): `get_plate_mutable_runtime()`, `core.config.constants`, named imports; `rg "config_and_data as cfg" viz_modules/procurement/` → 0
  - [x] **core/optimization/** (step 4): `get_plate_mutable_runtime()`, `PlateOrderContext` snapshot ports, `core.domain.plate_order`, `core.config.constants`, named imports; `rg "config_and_data as cfg" core/optimization/` → 0
  - [x] **viz_modules/layout_sequence/** (step 5): `LayoutSequenceCfgSlice.from_plate_runtime()`, `build_layout_runtime_snapshot()`, `core.project_paths`, named imports; `rg "config_and_data as cfg" viz_modules/layout_sequence/` → 0
  - [x] **Residual core/** (step 6): `get_plate_mutable_runtime()`, `core.domain.plate_order`, `core.config.constants`, `core.project_paths`, named imports; `rg "config_and_data as cfg" core/` → 0 (excl. `config_and_data.py`)
  - [x] **core/ residual** (step 6): `visualization.py`, `track_*`, `rescue_tracks.py`, `plates_preview_xlsx.py`, `reconciliation_xlsx.py`, `kp_db_*`; explicit runtime/constants/named imports; `rg "config_and_data as cfg" core/` → 0
  - [x] **viz_modules/** (step 6b): `visualization_drawing.py`, `price_utils.py` off module alias
- [x] Callers переведены на явные API: `PlateOrderContext`, `get_plate_mutable_runtime()`, `core.domain.plate_order`, `core.config.constants` (web/core hot paths)
- [x] `__getattr__` в `config_and_data` — `DeprecationWarning` на чтение `MUTABLE_LEGACY_NAMES` (step 7 partial; полное удаление proxy — после bot agreement)
- [x] Существующие isolation-тесты зелёные; нет регрессии procurement/planning flows (927 passed, 12 skipped)

**Verify:**
```bash
pytest tests/test_plate_runtime_isolation.py tests/test_plate_runtime_request_isolation.py tests/test_procurement_loads.py tests/test_config_and_data_plate_naming.py -q
pytest tests/test_production_planning*.py tests/test_plan_*.py -q
```

**Риск:** большой blast radius — **инкрементальные PR внутри WP1**, не big-bang.

---

## WP2 — A6: единый DI в API endpoints

**Цель:** все service dependencies через `app/dependencies/`, без inline `Service()` в handlers.

**Файлы:**
- `app/dependencies/` — новый модуль `services.py` (или split по доменам)
- `app/api/v1/endpoints/production.py` — заменить `ProductionService()` на `Depends(get_production_service)`
- `app/api/v1/endpoints/commercial.py` — `CommercialService`, `CommercialWorkflowService`
- `app/api/v1/endpoints/admin.py`, `archive.py` — выровнять с общим паттерном (перенос factories из endpoints)
- Тесты: `tests/test_auth_dependencies.py`, endpoint integration suites

**Acceptance:**
- [x] Factory functions для `ProductionService`, `CommercialService`, `CommercialWorkflowService`, `AdminService`, `ArchiveService` в `app/dependencies/`
- [x] Ни один endpoint в `production.py` / `commercial.py` не вызывает `SomeService()` напрямую в handler body
- [x] Паттерн совпадает с `admin.py`/`archive.py`: `service: X = Depends(get_x_service)`
- [x] Override в тестах через `app.dependency_overrides` документирован (fixture helper при необходимости)
- [x] Нет циклических импортов (dependencies → services → repositories only)

**Verify:**
```bash
pytest tests/test_production_api*.py tests/test_commercial*.py tests/test_archive_endpoints.py tests/test_admin*.py -q
rg "Service\(\)" app/api/v1/endpoints/ --glob "*.py"
```

---

## WP3 — A10: консолидация planning orchestration

**Цель:** один явный orchestration path для build/save/activate плана вместо трёх параллельных точек входа.

**Архитектура (A10):**

```
API (production.py)
  └─ ProductionService          — CRUD, activate, calendar, day view
       └─ ProductionPlanningService   — canonical build orchestrator
            └─ core.production.planning (load → optimize → persist)
                 └─ PlanRepository.build_plan_from_tracks
                      └─ plan_distribution.add_tracks_to_plan (helpers)
```

`app/planning/plan_manager` — legacy-фасад для frozen bot; web/core hot path его не вызывает.
Распределение по дням и merge lookup — `plan_distribution`; календарь/агрегация — `plan_calendar` / `plan_aggregation`; SQLite I/O — `PlanRepository`.

**Файлы:**
- `app/services/production_planning_service.py` — canonical orchestrator
- `app/planning/plan_manager.py`, `app/planning/plan_distribution.py` — utilities или thin delegates
- `app/api/v1/endpoints/production.py` — вызывает только service layer
- `core/production/planning.py` — domain functions без app imports

**Acceptance:**
- [x] Диаграмма или ADR-абзац: единственный entry `ProductionPlanningService` (или переименованный orchestrator) для build plan flow
- [x] `plan_manager` / `plan_distribution` не содержат дублирующей business logic build/save (только helpers: calendar, aggregation, I/O)
- [x] API endpoint `build_plan` / `create_plan` не импортирует `app.planning.plan_manager` напрямую
- [x] Integration test: create → build → activate через API использует один pipeline
- [x] Нет новых imports `app` → `bot` или обратно

**Verify:**
```bash
pytest tests/test_production_planning*.py tests/test_plan_sqlite_authority.py tests/test_plan_repository.py -q
rg "from app.planning.plan_manager|from app.planning.plan_distribution" app/api/
```

---

## WP4 — A5 (web-side): первый срез god-модулей

**Цель:** уменьшить связность и размер двух крупнейших web-модулей без full rewrite.

**Приоритетные модули:**
1. `app/web/router.py` (~939 LOC) — вынести auth/forms/logout в `app/web/routes/` или `app/web/handlers/`
2. `app/services/commercial_workflow_service.py` (~1103 LOC) — вынести 1–2 use cases (например draft lifecycle, export) в отдельные service modules

**Acceptance:**
- [x] `app/web/router.py` < 600 LOC **или** split на ≥2 модуля с `APIRouter` include (документировать структуру)
- [x] `commercial_workflow_service.py` — минимум один extracted module (`commercial_draft_service.py` / `commercial_export_service.py`) с перенесёнными методами
- [x] Публичные API endpoints **без** изменения контрактов (paths, schemas)
- [x] Handlers остаются thin: валидация → service → response
- [x] **Не** декомпозировать `bot/handlers/*` — только web

**Verify:**
```bash
pytest tests/test_web_*.py tests/test_commercial*.py tests/test_csrf.py -q
wc -l app/web/router.py app/services/commercial_workflow_service.py
```

---

## WP5 — S5: server-side RBAC re-check

**Цель:** sensitive actions всегда проверяются на API; frontend hide — UX only.

**Файлы:**
- `app/dependencies/auth.py`, `app/security/offer_access.py`
- `app/api/v1/endpoints/` — inventory admin, archive destructive, commercial PII, production operational
- `frontend/src/app/layout/AppHeader.tsx`, `frontend/src/shared/lib/roleRoutes.ts` — документировать как non-security
- Новые/расширенные тесты: `tests/test_*_authorization.py`

### Sensitive endpoints inventory (server guards)

| Endpoint group | Path pattern | Required roles | Object-level check |
|----------------|--------------|----------------|-------------------|
| Admin users/stats | `GET /admin/users`, `GET /admin/db/stats` | `admin` | — |
| Admin destructive reset | `POST /admin/db/reset/*` | `admin` | `require_destructive_db_reset` (env guard) |
| Admin recover | `POST /admin/db/recover-plates` | `admin` | — |
| Offers (PII) | `/offers/*` | `admin`, `manager` | `assert_offer_read/write_access` |
| Archive | `/commercial/archive/*` | `admin`, `manager` | `assert_offer_read/write_access` in service |
| Commercial drafts | `/commercial/drafts/*`, `/commercial/parse`, … | `admin`, `manager` | `verify_draft_ownership` (+ admin bypass) |
| Production ops | `/production/*` | `admin`, `production` | — |
| Managers list | `GET /managers` | `admin`, `manager`, `production` | — (non-PII reference data) |

**Acceptance:**
- [x] Таблица sensitive endpoints (admin reset, archive delete, offer read/write, production-only routes) с required roles
- [x] Каждый sensitive endpoint использует `require_roles` / `offer_access` — **нет** endpoints с только client-side guard
- [x] Тесты: production user → 403 на admin/commercial restricted; manager → 403 на admin destructive
- [x] Frontend: комментарий «UI hide ≠ authorization» в `roleRoutes.ts`
- [x] Регрессия P3 production read filter (`test_offers_production_authorization.py`) — зелёная

**Verify:**
```bash
pytest tests/test_offers_production_authorization.py tests/test_admin*.py tests/test_archive_endpoints.py tests/test_auth_dependencies.py tests/test_rbac_server_side.py -q
```

---

## WP6 (stretch) — S9-audit health + S7 XFF review + ADR closure

**Цель:** закрыть optional audit residuals низким effort.

**Файлы:**
- `app/api/v1/endpoints/health.py` — убрать `environment` / `app` из public response **или** split `/health` vs `/internal/health`
- `app/security/login_rate_limit.py`, `core/config/settings.py` — `trusted_proxy_ips`
- `tests/test_rate_limit_deployment.py`, `tests/test_client_ip_resolution.py`

**Acceptance (optional):**
- [x] Public `GET /api/v1/health` (и `/health` если unified) не возвращает `environment` / secrets
- [x] Deployment doc: `trusted_proxy_ips` обязателен за reverse proxy; пример nginx (docstring в `resolve_client_ip` + Field comment в settings)
- [x] S7 review checklist: XFF используется только от trusted proxy; negative tests актуальны
- [x] WP1 ADR (если не сделан в WP1) — финализирован

**Verify:**
```bash
pytest tests/test_rate_limit_deployment.py tests/test_client_ip_resolution.py tests/test_auth_login_rate_limit.py -q
```

---

## Приоритеты и порядок работ

| Порядок | WP | ID | Обоснование |
|---------|-----|-----|-------------|
| 1 | WP1 (ADR item) | ADR | Checklist до code removal — снижает риск A3 |
| 2 | WP2 | A6 | DI — фундамент для рефакторинга services/orchestration |
| 3 | WP1 | A3 | PEP 562 — критический architecture debt |
| 4 | WP3 | A10 | Orchestration — проще после DI и cleaner imports |
| 5 | WP4 | A5 | God-modules — после стабилизации service boundaries |
| 6 | WP5 | S5 | Security — можно параллельно с WP4 |
| 7 | WP6 | S9/S7 | Stretch — когда WP1–WP5 зелёные |

**Параллелизация:** WP5 независим от WP1–WP4 после inventory. WP4 и WP3 — последовательно предпочтительнее.

---

## Definition of Done (спринт)

- [x] WP1–WP3 acceptance выполнены (**обязательно**)
- [x] WP4 **или** WP5 — минимум один закрыт полностью
- [x] WP6 — optional; не блокирует closure
- [x] **Нет** новых bot-specific WP или расширения parity tests
- [x] `pytest tests/ -q` — **≥927 passed**, skipped без роста failed
- [x] Spec status → `closed` или `closed (stretch deferred)`
- [x] Audit doc обновлён: ссылка на closure + Health Score note

---

## Риски

| Риск | Вероятность | Митигация |
|------|-------------|-----------|
| A3 migration ломает optimization/procurement | High | Инкрементальные commits; run procurement + planning test suites каждый шаг |
| A10 refactor меняет plan output | Medium | Parity test с frozen fixtures; сравнение `version` / track count |
| A5 split ломает CSRF/web forms | Medium | `test_csrf.py`, `test_web_logout_csrf.py` в CI каждого PR |
| Scope creep в bot/ | Low | Explicit out-of-scope; grep gate на `bot/handlers` changes |
| S5 false sense of security | Medium | Таблица endpoints + negative tests обязательны |

---

## Следующий шаг

1. **Ревью v1** этой спеки — подтвердить scope и порядок WP.
2. **IMPLEMENT P2:** WP1 ADR → WP2 (A6) → WP1 code (A3) → WP3 (A10) → WP4 (A5) → WP5 (S5) → WP6 stretch.
3. **Не планировать:** bot god-modules, A8 viz inversion, PostgreSQL — до закрытия P2 или отдельного решения.

## Следующий спринт после P2 (preview)

→ Реализовано как draft: [`stabilizaciya-p3-architecture-2026-06-21.md`](./stabilizaciya-p3-architecture-2026-06-21.md)

1. **P3 architecture:** A8 core↔viz_modules inversion, A17 frontend `CreatePlanWizard`, fat services (A22)
2. **P3 security:** S10 OpenAPI prod lockdown, S11 CSP enforce, S12 session cookie flags
3. **Optional:** полное удаление `bot/` + `run_bot.py` — **отдельное согласование**

---

## Post-closure delta (2026-06-21)

Сверка с [`../develop/audits/2026-06-21-full-project-audit.md`](../develop/audits/2026-06-21-full-project-audit.md). **P2 остаётся `closed`** — пункты ниже это gaps / другой DoD, не reopen.

| Тема | P2 DoD (closed) | Остаётся (→ P3) |
|------|-----------------|-----------------|
| **A6 DI** | `production`, `commercial`, `admin`, `archive` через `Depends` | `offers.py` (8×), `managers.py` — inline `Service()` |
| **A10 orchestration** | Единый entry `ProductionPlanningService`; API без `plan_manager` | Thin `PlanRepository` — оркестрация всё ещё в repository (~645 LOC) |
| **A3 PEP 562** | Web/core hot paths off proxy; `DeprecationWarning` | Proxy active; bot legacy path; полное удаление — P3/P4 |
| **A8 viz inversion** | Explicitly deferred | Audit-21 **[A1] Critical** — P3 WP3 |
| **S5 RBAC** | Server authoritative; `test_rbac_server_side.py` | Frontend guards (S8) — UX; `/managers`+production — **by design** (D1 в P3) |
| **S9 health** | `environment`/`app` redacted in production | `rate_limiting` metadata still public — P3 WP4 |
| **Bot / full-repo score** | Out of scope P2 | Audit Health 2.0/10 — full repo lens, не откат P2 |

**ID mapping:** audit 2026-06-21 использует новую нумерацию (A1 = viz inversion, не split-brain). Таблица — в P3 spec § «Маппинг ID аудитов».

---

## Changelog

| Дата | Изменение |
|------|-----------|
| 2026-06-20 | v1 draft — спринт P2 architecture после closure P1-next |
| 2026-06-21 | WP1 step 6–7: residual `core/` + viz drawing/price_utils off alias; DeprecationWarning on PEP 562 proxy |
| 2026-06-21 | WP3 A10: ProductionPlanningService canonical orchestrator; plan_repository → plan_distribution; API integration test |
| 2026-06-21 | WP4 A5 slice 1: web router split (legacy_routes/spa_routes/shell); commercial_draft_service + commercial_export_service |
| 2026-06-21 | WP5 S5: server-side RBAC inventory + test_rbac_server_side.py; admin draft bypass; roleRoutes UI≠auth comment |
| 2026-06-21 | WP6 S9-audit: production `/health` redacts `environment`/`app`; S7 XFF docstring; tests in test_rate_limit_deployment.py |
| 2026-06-21 | **P2 closed** — WP1–WP6 complete; pytest ≥947 passed |
| 2026-06-21 | Post-closure delta + link to P3 spec draft (audit 2026-06-21 reconciliation) |

---

*Создано: 2026-06-20 · v1: architecture + security backlog (web/core only, bot frozen).*
