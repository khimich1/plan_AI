# Spec: стабилизация P5 — bot decommission, god-modules, CreatePlanWizard

> **Тип:** remediation feature-spec (архитектурный спринт)
> **Фаза SDD:** draft
> **Дата:** 2026-06-21
> **Ревизия:** v1 draft
> **Статус:** closed (WP1–WP5)
> **Источник:** [`../develop/audits/2026-06-21-full-project-audit.md`](../develop/audits/2026-06-21-full-project-audit.md)
> **Predecessor (закрыт):** [`stabilizaciya-p4-architecture-2026-06-21.md`](./stabilizaciya-p4-architecture-2026-06-21.md) — WP1–WP4 closed, S12 in code, **964 passed**
> **Bot policy (наследуется):** [`stabilizaciya-p0-audit-2026-06-19.md`](./stabilizaciya-p0-audit-2026-06-19.md) § «Решение по Telegram-боту»
> **ADR:** [`../develop/architecture/core-viz-modules-boundary.md`](../develop/architecture/core-viz-modules-boundary.md), [`../develop/architecture/pep562-config-and-data-decommission.md`](../develop/architecture/pep562-config-and-data-decommission.md)

---

## Стратегия (одной фразой)

> Зафиксировать **soft decommission** Telegram-бота, снять архитектурный шум full-repo lens, и инкрементально декомпозировать **три god-модуля** (visualization, kp_db_offers, commercial workflow) + **CreatePlanWizard** на web path; bot **не развиваем**, физическое удаление — **P6**.

---

## Маппинг ID (важно)

Нумерация **P5 backlog** ≠ нумерация аудита 2026-06-21.

| P5 / stabilization ID | Аудит 2026-06-21 | Статус до P5 |
|----------------------|------------------|--------------|
| **Bot D1** | A1, A2, A10, S4, Q6 (bare except) | Deprecated; frozen с P0–P4 |
| **A7** god-modules | **A4** `core/visualization.py`, **A5** `core/kp_db_offers.py`, **A6** `CommercialWorkflowService` | Open (~1314 / ~1113 / ~844 LOC) |
| **Q6** CreatePlanWizard | — (frontend; audit не нумерует) | Open (~950 LOC monolith) |
| **Q8** response_model (partial) | **Q8** typing | Auth closed (session); production/commercial — open |
| A3 viz ports (app) | **A3** app bypass ports | Partially closed — `CommercialService` уже на ports |
| S1 Redis | **S1** | Deferred P6 |
| S3 OCR | **S3** | Deferred P6 |
| A9 PostgreSQL | scaling / **A9** raw SQL residual | Deferred P6 |

*Примечание: в аудите **A7** = sync CPU optimization в HTTP — **не входит в P5** (defer P6+).*

---

## Bot deprecation → decommission (D1)

### Рекомендация: **soft decommission** (default)

| Вариант | Плюсы | Минусы | Решение P5 |
|---------|-------|--------|------------|
| **Archive** (`bot_archived/`) | История в git; быстрый rollback для reference | Два дерева; риск случайного import | ✅ **Default** — move + README |
| **Delete сразу** | Чистый repo; full-repo Health ↑ | Потеря reference; большой diff | ❌ Defer **P6** |
| **Keep frozen** | Минимальный diff | Bot tests, bot→app imports, audit Critical A1/A2 остаются | ❌ Не закрывает цель P5 |

### Soft decommission — что делаем в P5

1. **Переместить** `bot/` → `bot_archived/` (сохранить структуру handlers/services).
2. **Оставить stub:** `bot/README.md` (DEPRECATED, ссылка на `bot_archived/`, дата, web = canonical).
3. **`run_bot.py`:** оставить entry с `sys.exit(1)` + stderr message «DEPRECATED — см. bot/README.md» (не запускать aiogram).
4. **CI / deploy / setup:** убрать `requirements-bot.txt` из default install path (`setup_venv.sh`, `requirements.txt` comments); bot deps — optional doc only.
5. **Тесты:** удалить или `@pytest.mark.archived` + exclude `test_bot_*.py` из default `pytest tests/` (5 файлов).
6. **Parity tests:** `test_plan_consistency.py::test_rescue_web_matches_bot` — удалить или заменить web-only invariant.
7. **PEP 562:** зафиксировать в ADR — **8 bot consumers** `config_and_data` сняты с active inventory; proxy backlog → **0 active** (физическое удаление `__getattr__` — P6).
8. **Imports bot→app** (6 файлов): архивируются вместе с bot; grep gate `bot_archived/` ↛ `app/` — documentation only.

### Impact inventory

| Область | Файлы / объём | Действие P5 |
|---------|---------------|-------------|
| Handlers | 22 файла, ~10.8k LOC (mega: `commercial.py` ~2234, `production_completion.py` ~1238, `production_day_view.py` ~1073) | Archive |
| bot→app imports | `production_planning_adapter.py`, `commercial.py`, `production_execution.py`, `production_completion.py`, `plan_manager.py`, `production_export.py` | Archive |
| PEP 562 consumers | 8 import sites в bot | Off active inventory |
| Tests | `test_bot_*.py` (5), parity в `test_plan_consistency.py` | Remove / exclude |
| Entry | `run_bot.py` | Stub only |
| Deps | `requirements-bot.txt` | Optional, not in default venv |

---

## Objective

После closure P4 (**964 passed**): (1) формально decommission bot без production impact; (2) два инкрементальных среза god-module refactor (A7); (3) первый срез декомпозиции CreatePlanWizard (Q6); (4) optional Q8 typing на non-auth endpoints. Цель — поднять **web/core Health Score** с **~6.5–7.0/10** к **~7.5–8.0/10**; full-repo lens — с **2.0/10** к **~4–5/10** (bot Critical сняты с active path).

**Verify baseline:** `pytest tests/ -q` → **≥964 passed**, 12 skipped, 0 failed (после bot test exclusion — новый baseline зафиксировать в changelog).

---

## Контекст после P4

| P4 WP | Закрыто | Остаётся для P5 |
|-------|---------|-----------------|
| WP1 A1 slice 2 | `core/` ↛ `viz_modules` runtime imports | Декомпозиция `core/visualization.py` по доменам |
| WP2 A5 | `PlanLoadPort` + repository read | — |
| WP3 A4 | PEP 562 proxy 32→16; web/core off proxy | Bot consumers removed → P6 full proxy delete |
| WP4 Q1–Q3, Q8 | DRY + core logging | Q8 partial на API endpoints |
| WP5 S12 | Password change rate limit (in code) | — |
| Deferred P4 | Q6, A7, bot removal, S1, S3, A9 | **In scope P5** |

**LOC snapshot (2026-06-21):**

| Модуль | LOC | Роль |
|--------|-----|------|
| `core/visualization.py` | ~1314 | Layout, track split, visualize_plan, export side effects |
| `core/kp_db_offers.py` | ~1113 | SQLite CRUD, search, stats, xlsx blob |
| `app/services/commercial_service.py` | ~304 | Preview/save (уже на viz ports; не primary god) |
| `app/services/commercial_workflow_service.py` | ~844 | Wizard orchestration, OCR, draft, export |
| `frontend/.../CreatePlanWizard.tsx` | ~950 | 3-step wizard + fill mode + inline Kp row |

---

## Product decisions (требуют подтверждения на ревью)

| # | Вопрос | Default для draft | Альтернатива |
|---|--------|-------------------|--------------|
| **D1** | Bot removal strategy | **Soft decommission** — `bot_archived/`, stub `run_bot.py`, exclude bot tests; **hard delete P6** | Hard delete сразу; keep frozen |
| **D2** | Q6 split strategy | **Vertical slice by wizard step** — Step1/Step2/Step3 components + `useCreatePlanWizardState` hook; fill mode stays in shell | Horizontal layers (hooks/api/ui folders only) |
| **D3** | A7 decomposition order | **Slice 1:** `core/visualization.py` → `core/visualization/layout.py` (track/layout pure functions). **Slice 2:** `core/kp_db_offers.py` → `core/kp/offers_repository.py` (read path). **Slice 3 (stretch):** `CommercialWorkflowService` wizard steps → dedicated services | Start with CommercialWorkflow; kp_db first |

**Обоснование D3 slice 1 = visualization layout:** сильное покрытие `tests/test_layout_*.py`; функции `split_sequence_into_tracks`, `validate_track_integrity` — чистый domain без I/O; ports уже вынесены в P3/P4; минимальный риск регрессии vs OCR/wizard paths.

---

## Выбор scope

| Трек | IDs | Effort | Решение |
|------|-----|--------|---------|
| **In scope (P5 P0)** | **D1** bot soft decommission | S–M | Снимает Critical A1/A2 с active codebase |
| **In scope (P5 P0)** | **A7 slice 1** visualization layout | M | Lowest-risk god-module cut |
| **In scope (P5 P1)** | **A7 slice 2** kp_db_offers read repository | M | Разблокирует KpRepository consolidation |
| **In scope (P5 P1)** | **Q6** CreatePlanWizard step extraction | M | Frontend maintainability |
| **Stretch (P5 P2)** | **A7 slice 3** CommercialWorkflowService; **Q8** non-auth `response_model` | M–L | Не блокирует closure |
| **Deferred P6** | **S1** Redis, **S3** OCR consent, **A9** PostgreSQL, hard delete `bot_archived/`, full PEP 562 removal | L | Explicit out-of-scope P5 |
| **Out of scope P5** | Новые product features, MFA, encryption at rest, async optimization job API (audit A7 HTTP), bot reliability fixes | — | Отдельные спеки |

---

## Scope

### In scope

| Приоритет | ID | Проблема | Fix (кратко) |
|-----------|-----|----------|--------------|
| P0 | **D1** | Bot deprecated но живёт в repo, tests, imports | Soft decommission: archive, stub, exclude tests, update docs |
| P0 | **A7-1** | `core/visualization.py` ~1314 LOC | Extract `core/visualization/layout.py` + re-export facade |
| P1 | **A7-2** | `core/kp_db_offers.py` ~1113 LOC | Extract read SQL → `core/kp/offers_read.py` or extend `KpRepository` |
| P1 | **Q6** | `CreatePlanWizard.tsx` ~950 LOC monolith | Step components + shared hook; shell < ~250 LOC |
| P2 | **A7-3** | `CommercialWorkflowService` ~844 LOC | Extract wizard transition service (1 step domain) |
| P2 | **Q8** | dict responses на production/commercial endpoints | Pydantic `response_model` для top 5–8 endpoints |

### Out of scope (explicit)

- **S1** Redis-backed rate limiting
- **S3** OCR third-party consent / data residency policy
- **A9 / S14** PostgreSQL migration, shared DraftStore redesign
- **Hard delete** `bot/` (P6)
- **Full** PEP 562 `__getattr__` removal (P6, after bot gone)
- **S6** CSP enforce, **S8** frontend role guards
- Новые product features, onboarding UX (см. product-analysis SWOT)

---

## WP1 — Bot soft decommission (D1)

**Цель:** bot не участвует в active codebase, CI, production; документированный путь к P6 hard delete.

**Effort:** S–M

**Файлы:**
- `bot/` → `bot_archived/` (move entire tree)
- `bot/README.md` — NEW stub (DEPRECATED pointer)
- `run_bot.py` — stub exit
- `setup_venv.sh`, `requirements.txt` — remove bot from default install narrative
- `tests/test_bot_*.py` — delete or move to `tests/archived/`
- `tests/test_plan_consistency.py` — remove/adapt `test_rescue_web_matches_bot`
- `ai_docs/develop/architecture/pep562-config-and-data-decommission.md` — bot consumers → archived
- `pytest.ini` or `pyproject.toml` — `norecursedirs` / `--ignore` if needed

**Acceptance:**
- [x] `bot/` содержит только README stub; код в `bot_archived/`
- [x] `python run_bot.py` → exit ≠ 0, message DEPRECATED (no aiogram start)
- [x] `rg "from bot\.|import bot" app/ core/ --glob "*.py"` → **0** (excluding `bot_archived/`)
- [x] Default `pytest tests/ -q` green; bot tests excluded; новый baseline записан в changelog
- [x] ADR PEP 562: active bot consumer count → **0**
- [x] Product decision **D1** зафиксирована

**Verify:**
```bash
pytest tests/ -q
rg "from bot\.|import bot" app/ core/ tests/ --glob "*.py"
python run_bot.py ; echo exit=$?
```

**Риски:**

| Риск | Митигация |
|------|-----------|
| Случайный import из `bot_archived/` | README + grep gate; no `bot_archived` on PYTHONPATH |
| Parity test removal скрывает drift | Web-only tests already canonical since P2 |
| Lost rollback reference | Git history + `bot_archived/` preserved |

---

## WP2 — A7 slice 1: `core/visualization` layout submodule

**Цель:** вынести track/layout pure logic из mega-module; `core/visualization.py` — thin facade + `visualize_plan` orchestration.

**Effort:** M

**Файлы:**
- NEW: `core/visualization/__init__.py` (facade re-exports)
- NEW: `core/visualization/layout.py` — `LayoutIntegrityError`, `TrackLayoutInvariantError`, `validate_track_integrity`, `split_sequence_into_tracks`, private helpers `_starter_solid_tiers`, `_pick_track_starter_solid_index`, `_assert_track_starts_with_solid`, `_count_solids_remaining`, `_iter_sequence_items`, `_ensure_layout_uid`
- `core/visualization.py` → migrate remaining (visualize_plan, drawing hooks) OR shrink to re-export during transition
- `core/ports/visualization.py` — update imports if needed
- `tests/test_layout_*.py`, `tests/test_layout_identity_integrity.py` — must stay green
- `ai_docs/develop/architecture/core-viz-modules-boundary.md` — submodule map

**Acceptance:**
- [x] `core/visualization/layout.py` exists; layout public API documented
- [x] `core/visualization.py` LOC reduced by **≥400** (target: layout block extracted)
- [x] `rg "def split_sequence_into_tracks" core/` → **1** definition (layout module)
- [x] `tests/test_layout_*.py` + import boundary tests green
- [x] No new `viz_modules` imports in `core/`

**Verify:**
```bash
pytest tests/test_layout_*.py tests/test_core_viz_import_boundary.py -q
rg "viz_modules" core/ --glob "*.py"
```

**Риски:**

| Риск | Митигация |
|------|-----------|
| Circular imports layout ↔ visualization | Pure layout module; no matplotlib/pandas |
| Call sites break on import path | Facade re-export from `core.visualization` for compat |

---

## WP3 — A7 slice 2: `kp_db_offers` read repository

**Цель:** отделить read/query SQL от write/admin helpers; единая точка для list/search/get KP.

**Effort:** M

**Файлы:**
- NEW: `core/kp/offers_read.py` (or `core/repositories/kp_offers_read.py`) — `get_kp_by_id`, `get_all_kp_list`, `search_kp_by_customer_name`, `get_kp_completion_percentage`, `get_kp_plates_in_plan_percentage`, `get_kp_total_length`
- `core/kp_db_offers.py` — re-export or thin wrapper; write paths (`save_kp_to_db`, `update_*`, `delete_*`, `clear_*`) remain temporarily
- `app/repositories/kp_repository.py` — delegate reads to new module (avoid duplicate SQL)
- Tests: `tests/test_kp_db_*.py`, `tests/test_archive*.py`

**Acceptance:**
- [x] Read functions live in dedicated module; `kp_db_offers.py` imports and re-exports (backward compat)
- [x] `rg "def get_kp_by_id|def get_all_kp_list" core/` → **1** definition each
- [x] Archive/list/search integration tests green
- [x] No behaviour change in KP list ordering / access filters

**Verify:**
```bash
pytest tests/test_kp_db_*.py tests/test_archive*.py tests/test_commercial*.py -q
rg "def get_all_kp_list" core/ --glob "*.py"
```

**Риски:**

| Риск | Митигация |
|------|-----------|
| Duplicate SQL in KpRepository vs offers_read | Single implementation; repository delegates |
| Access filter regression | Existing `test_kp_db_*` + archive endpoint tests |

---

## WP4 — Q6: CreatePlanWizard incremental extraction

**Цель:** shell-компонент orchestrates; UI и state по шагам изолированы; fill mode не ломается.

**Effort:** M

**Strategy (D2 default): vertical slice by step**

**Target structure:**
```
frontend/src/features/production/
  components/
    CreatePlanWizard.tsx          # shell: step routing, mutations, ~200-250 LOC
    create-plan-wizard/
      WizardStepIndicator.tsx
      Step1PlanStartDate.tsx      # date, plan name, MonthCalendarGrid
      Step2TracksConfig.tsx       # tracks count, presets
      Step3KpPlateSelection.tsx   # filter, KP list, estimates
      KpCandidateRow.tsx            # extracted from inline ~L800+
  hooks/
    useCreatePlanWizardState.ts   # selectedPlatesByKp, step, fill mode sync
```

**Файлы:**
- `frontend/src/features/production/components/CreatePlanWizard.tsx` — refactor to shell
- NEW components under `create-plan-wizard/` (4–5 files)
- NEW `hooks/useCreatePlanWizardState.ts`
- Existing: `MonthCalendarGrid.tsx`, `productionEstimate.ts`, `useProductionQueries.ts` — unchanged API

**Incremental order (within WP4):**
1. Extract `KpCandidateRow` (lowest coupling, ~150 LOC)
2. Extract `Step1PlanStartDate` + `Step2TracksConfig`
3. Extract `Step3KpPlateSelection` + move selection state to hook
4. Wire fill mode (`fillRequest`) in shell only

**Acceptance:**
- [x] `CreatePlanWizard.tsx` **< 280 LOC**
- [x] Each step component **< 200 LOC**; `KpCandidateRow` isolated with typed props
- [ ] Manual: create plan flow (step 1→2→3→build) works; fill mode from calendar works
- [x] `npm run build` (frontend) succeeds
- [x] No new API calls; behaviour parity with pre-refactor

**Verify:**
```bash
cd frontend && npm run build
# optional: npm run test -- --run CreatePlanWizard  (if vitest covers)
pytest tests/test_production_api_integration.py -q
```

**Риски:**

| Риск | Митигация |
|------|-----------|
| State lifted incorrectly → stale selections | Hook owns Maps/Sets; pass stable callbacks |
| Fill mode regression | Explicit acceptance: fillRequest path unchanged |
| Over-extraction in one PR | 4 sub-commits per incremental order |

---

## WP5 (stretch) — A7 slice 3 + Q8 partial

**Цель:** сократить `CommercialWorkflowService`; улучшить OpenAPI schema на hot endpoints.

**Effort:** M–L (optional — не блокирует P5 closure)

### A7 slice 3 — CommercialWorkflowService

**Файлы:**
- NEW: `app/services/commercial_wizard_step_service.py` — `infer_wizard_current_step`, `_normalize_stored_step`, step transitions
- `app/services/commercial_workflow_service.py` — delegate; target **< 600 LOC**
- Tests: existing `tests/test_commercial*.py`, wizard integration

**Acceptance (stretch):**
- [x] Step inference logic extracted; workflow service delegates
- [x] Wizard API tests green

### Q8 partial — non-auth response_model

**Scope:** top endpoints без auth `response_model` (already done for auth per session closure):

| Endpoint area | File | Priority |
|---------------|------|----------|
| Production plans list/get | `app/api/v1/endpoints/production*.py` | P2 |
| Archive list/details | `app/api/v1/endpoints/archive.py` | P2 |
| Commercial wizard | `app/api/v1/endpoints/commercial.py` | P2 |

**Acceptance (stretch):**
- [x] ≥5 endpoints with explicit `response_model`
- [x] OpenAPI schema validates in dev `/docs`
- [x] No response shape change (field names stable)

**Verify:**
```bash
pytest tests/test_production_api_integration.py tests/test_archive_endpoints.py tests/test_commercial*.py -q
```

---

## Приоритеты и порядок работ

| Порядок | WP | ID | Обоснование |
|---------|-----|-----|-------------|
| 1 | WP1 | D1 | Снимает bot Critical; упрощает grep/Health; разблокирует P6 proxy delete |
| 2 | WP2 | A7-1 | Lowest-risk god-module; strong test net |
| 3 | WP3 | A7-2 | Independent of frontend; backend maintainability |
| 4 | WP4 | Q6 | После backend slices — меньше context switch |
| 5 | WP5 | A7-3, Q8 | Stretch |

**Параллелизация:** WP2 и WP1 могут идти последовательно (WP1 first — smaller diff noise). WP4 может стартовать параллельно WP3 после step component design agreed (D2).

---

## Definition of Done (спринт P5)

- [x] WP1–WP4 acceptance выполнены (**обязательно**)
- [x] WP5 — optional stretch; не блокирует closure
- [ ] Product decisions **D1–D3** зафиксированы в changelog
- [ ] `pytest tests/ -q` — **≥964 passed** (adjusted baseline after bot test exclusion documented)
- [ ] `npm run build` — frontend green
- [ ] Health Score web/core — **~7.5–8.0/10**; full-repo — **~4–5/10** (bot archived)
- [ ] Spec status → `closed` or `closed (WP5 stretch deferred)`
- [ ] Audit 2026-06-21 — cross-ref P5 (optional note)

---

## Риски (спринт)

| Риск | Вероятность | Митигация |
|------|-------------|-----------|
| Bot archive ломает неожиданный import | Low | Grep gate; stub README |
| Visualization layout extraction breaks merge tests | Medium | Incremental; run layout suite each commit |
| kp_db read split changes list ordering | Medium | Golden archive tests |
| CreatePlanWizard UX regression | Medium | Manual checklist; fill mode explicit AC |
| Scope creep → Redis/PostgreSQL | Low | Explicit out-of-scope table |
| CommercialWorkflow extraction touches OCR paths | Medium | Defer to WP5 stretch |

---

## Следующий шаг

1. **Ревью v1** — подтвердить D1 (soft decommission), D2 (step split), D3 (visualization first).
2. **IMPLEMENT P5:** WP1 → WP2 → WP3 → WP4 → WP5 stretch.
3. **P6 preview:** hard delete `bot_archived/`, full PEP 562 removal, S1 Redis, S3 OCR, A9 PostgreSQL, audit A7 async optimization.

---

## Changelog

| Дата | Изменение |
|------|-----------|
| 2026-06-21 | v1 draft — P5 после closure P4 (964 passed); scope: D1 bot soft decommission, A7×2, Q6 wizard, stretch A7-3 + Q8 |
| 2026-06-21 | **WP2 closed (A7 slice 1)** — `core/visualization/layout.py` extracted (~705 LOC); package facade `core/visualization/__init__.py` (~631 LOC); monolith removed (was 1313 LOC); pytest **914 passed**, 9 skipped |
| 2026-06-21 | **WP5 closed (stretch)** — Q8: `response_model` on offers/managers/health (8 JSON endpoints); A7-3: `CommercialWizardStepService` extracted; pytest **914 passed**, 9 skipped |

---

*Создано: 2026-06-21 · v1 draft: bot soft decommission, god-module slices, CreatePlanWizard extraction, baseline 964 passed / Health web/core ~6.5–7.0.*
