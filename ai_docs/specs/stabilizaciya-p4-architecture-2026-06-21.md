# Spec: стабилизация P4 — web/core после closure P3

> **Тип:** remediation feature-spec (архитектурный спринт)
> **Фаза SDD:** CLOSED
> **Дата:** 2026-06-21
> **Ревизия:** v1 closed
> **Статус:** closed (WP5 stretch S12 deferred)
> **Источник:** [`../develop/audits/2026-06-21-full-project-audit.md`](../develop/audits/2026-06-21-full-project-audit.md)
> **Predecessor (закрыт):** [`stabilizaciya-p3-architecture-2026-06-21.md`](./stabilizaciya-p3-architecture-2026-06-21.md) — WP1–WP4 closed, WP5 deferred
> **Bot policy (наследуется):** [`stabilizaciya-p0-audit-2026-06-19.md`](./stabilizaciya-p0-audit-2026-06-19.md) § «Решение по Telegram-боту»
> **ADR (slice 1):** [`../develop/architecture/core-viz-modules-boundary.md`](../develop/architecture/core-viz-modules-boundary.md)

---

## Стратегия (одной фразой)

> Завершить **deferred P3 backlog** на web/core path: второй срез инверсии `core`↔`viz_modules` (A1 slice 2), SQL через repository (A5), сужение PEP 562 proxy (A4), DRY + logging hygiene (Q1–Q3, Q8); bot **не трогаем**.

---

## Маппинг ID аудитов (важно)

Нумерация в аудите **2026-06-21** не совпадает с аудитом **2026-06-20** / P2.

| Аудит 2026-06-21 | Аудит 2026-06-20 / P2 | Статус до P4 |
|------------------|------------------------|--------------|
| **A1** core↔viz_modules | **A8** (deferred P2) | ✅ Closed P4 WP1 — slice 2 (procurement/drawing ports); `core/` ↛ `viz_modules` |
| **A2** PlanRepository orchestration | **A10** partial | ✅ Closed P3 WP2 |
| **A3** DI offers/managers | **A6** partial | ✅ Closed P3 WP1 |
| **A4** global plate runtime / PEP 562 | **A3** phase 3+ | ✅ Closed P4 WP3 — web/core off proxy; bot-only backlog |
| **A5** raw sqlite3 в domain | — | ✅ Closed P4 WP2 — `PlanLoadPort` + repository read methods |
| **S2, S7, S9-residual** | — | ✅ Closed P3 WP4 |
| **Q1–Q3, Q8** | — | ✅ Closed P4 WP4 — DRY pricing/phone + core hot-path logging |
| **S12** password change rate limit | — | ⏸ Deferred (WP5 stretch) |
| A1 split-brain SQLite/JSON | **A1** | ✅ Closed P0-next |
| A2 bot planning bypass | **A2** | ✅ Closed (bot deprecated) |

---

## Bot deprecation strategy

| Тема | Решение |
|------|---------|
| **Статус** | Bot **deprecated**. Активный канал — **web (React SPA + FastAPI API)**. |
| **Вне scope P4** | Bot god-modules, `bare except`, bot tests (Q4–Q7, Q9), bot→app DIP (A8 audit-21), bot `print()` (Q8 bot paths) |
| **PEP 562** | P4 WP3: сузить proxy для **web callers**; bot `config_and_data` — только minimal fix при поломке тестов |
| **`run_bot.py`** | Не использовать в production; удаление `bot/` — отдельное согласование |

---

## Objective

Закрыть оставшийся **web/core** backlog после closure P3 (WP1–WP4): полный A1 hot-path gate, A5 repository boundary, A4 proxy decommission slice, quality quick wins Q1–Q3/Q8. Цель — поднять Health Score web/core с **~6.0/10** к **~7.5–8/10**. **≥955 passed, 12 skipped** — baseline не должен регрессировать.

## Контекст после P3

| P3 WP | Закрыто | Остаётся для P4 |
|-------|---------|-----------------|
| WP1 A3 | DI offers/managers через `Depends` | — |
| WP2 A2 | Thin `PlanRepository` + `PlanDistributionService` | — |
| WP3 A1 slice 1 | Ports для `build_layout_sequence`, `load_price_table_from_xlsx`; `planning.py` ↛ `viz_modules` | — |
| P4 WP1 A1 slice 2 | Procurement + drawing ports; `core/visualization.py` ↛ `viz_modules`; grep gate 0 | — |
| WP4 S2/S7/S9 | OpenAPI off in prod; role whitelist; health redaction | — |
| WP5 (deferred) | — | **A5** sqlite3 → repo; **A4** PEP 562 proxy slice |

**Health Score:** P3 closure ~**6.0/10** (web/core scope). Full-repo lens **2.0/10** — без изменений (bot, god-modules).

**Verify baseline:** `pytest tests/ -q` → **955 passed, 12 skipped, 0 failed**.

---

## Product decisions (требуют подтверждения на ревью)

| # | Вопрос | Default для draft | Альтернатива |
|---|--------|-------------------|--------------|
| D1 | S12 rate limit backend | **In-process** (как login S1 pattern) — без Redis | Defer S12 до Redis (S1) |
| D2 | Q8 `print()` scope | **Core hot paths only** (`production/planning`, `visualization`, `optimization/*`) | Full `core/` sweep |
| D3 | A4 PEP 562 | **Narrow proxy** — web path off proxy; inventory remaining consumers | Full removal (blocked by bot legacy) |
| D4 | Q2 archive mapper | **Web-only** — единый mapper в `archive_service`; bot duplicate untouched | Bot calls service (bot scope) |

---

## Выбор scope

| Трек | IDs (audit 2026-06-21) | Effort | Решение |
|------|------------------------|--------|---------|
| **In scope (P4 P0)** | **A1** (slice 2), **A5**, **A4** (proxy slice) | M–L | Архитектурные блокеры после P3 |
| **In scope (P4 P1)** | **Q1**, **Q2** (web), **Q3**, **Q8** (core hot paths) | M | Maintainability + observability |
| **Stretch** | **S12** password change rate limit | S | По решению D1 |
| **Deferred** | **S1** Redis rate limit, **S3** OCR consent, **A9** PostgreSQL, **Q6** CreatePlanWizard, bot Q4–Q10, **A8** bot↔app | L | P5+ или отдельные спеки |
| **Out of scope** | Bot reliability, frontend wizard split, MFA, encryption at rest, decommission `bot/` | L | Отдельное согласование |

---

## Scope

### In scope

| Приоритет | ID | Проблема | Fix (кратко) |
|-----------|-----|----------|--------------|
| P0 | **A1** (slice 2) | `core/visualization.py` — direct `viz_modules.procurement`, `viz_modules.visualization_drawing` (+ lazy imports) | Новые Protocol в `core/ports/visualization.py`; adapters в `viz_modules/adapters/`; grep gate → 0 `viz_modules` в `core/` |
| P0 | **A5** | Raw `sqlite3.connect` в `core/production/planning.py` (`_load_kp_list`, `_load_plates_for_kps`) | SQL через `PlanRepository` или dedicated read port; planning — domain orchestration only |
| P0 | **A4** (slice) | PEP 562 proxy в `core/config_and_data.py` — web callers | Inventory web imports; migrate to explicit `DraftStore` / session state; сузить `__getattr__`; ADR update |
| P1 | **Q1** | Дублированный расчёт итогов заказа | Единая `calculate_order_totals()` в shared module; `commercial_offer.py` + `commercial_offer_xlsx.py` → import |
| P1 | **Q2** (web) | Дублированный KP→order_data mapping | Единый mapper в `app/services/archive_service.py` (web path) |
| P1 | **Q3** | Дублированный `format_phone()` | `core/utils/phone.py` (или `app/utils/formatters.py`); удалить копии в offer modules |
| P1 | **Q8** | `print()` в core production hot paths | `logging.getLogger(__name__)`; уровни info/warning/error; без изменения bot paths |

### Stretch

| ID | Проблема | Fix (кратко) |
|----|----------|--------------|
| **S12** | Нет rate limiting на `POST /change-password` | Reuse in-process sliding window pattern (`login_rate_limit.py`); 3 attempts / 15 min per IP+user |

### Deferred

- **S1** in-process rate limiting → Redis (unless multi-worker policy changes)
- **S3** OCR third-party consent / data minimization
- **S6** CSP enforce mode
- **S8** frontend role-based route guards (UX; server RBAC уже есть)
- **S10** managers PII для production (D1 P3 — accepted)
- **A7, Q4–Q6** god-module декомпозиция (bot, frontend wizard, mega `core/visualization.py` split)
- **A9, S14** PostgreSQL, shared DraftStore

### Out of scope (explicit)

- Новые product features
- Bot reliability и commercial pipeline consolidation
- Полное удаление `config_and_data.py` module file
- Decommission `legacy_routes` / `bot/`
- **Redis S1**, **OCR S3**, **PostgreSQL A9**, **CreatePlanWizard Q6**

---

## WP1 — A1 slice 2: procurement + drawing off direct viz_modules

**Цель:** `rg "viz_modules" core/ --glob "*.py"` → **0**; расширить ports/adapters pattern из slice 1.

**Контекст P3:** slice 1 закрыл `build_layout_sequence`, `load_price_table_from_xlsx`; deferred — procurement + drawing в `core/visualization.py` (см. ADR).

**Файлы:**
- `core/visualization.py` — убрать top-level и lazy `from viz_modules.procurement import ...`, `from viz_modules.visualization_drawing import ...`
- `core/ports/visualization.py` — новые Protocol + facades (procurement breakdown, drawing helpers — по фактическим call sites)
- `viz_modules/adapters/visualization_ports.py` — register implementations
- `app/adapters/visualization.py` — wiring без изменений или расширение
- `ai_docs/develop/architecture/core-viz-modules-boundary.md` — update slice 2
- `tests/test_core_viz_import_boundary.py` — расширить gate на весь `core/`

**Acceptance:**
- [x] ADR updated: slice 2 ports listed; `core/` ↛ `viz_modules/` contract complete
- [x] `rg "viz_modules" core/ --glob "*.py"` → **0** (runtime imports; docstring/comment refs only)
- [x] `tests/test_core_viz_import_boundary.py` — gate на `core/visualization.py` и full `core/` scan
- [x] Layout, procurement, production planning tests зелёные

**Verify:**
```bash
pytest tests/test_core_viz_import_boundary.py tests/test_layout_*.py tests/test_procurement_*.py tests/test_production_planning*.py -q
rg "viz_modules" core/ --glob "*.py"
```

---

## WP2 — A5: sqlite3 в planning.py → PlanRepository/port

**Цель:** domain module не открывает SQLite напрямую; единая точка доступа к KP/plate data.

**Контекст P3:** deferred WP5; функции `_load_kp_list`, `_load_plates_for_kps` (~L485, L560) используют `sqlite3.connect`.

**Файлы:**
- `core/production/planning.py` — удалить `import sqlite3`; load helpers через injected port или `PlanRepository` methods
- `app/repositories/plan_repository.py` — read methods: `load_kp_list(...)`, `load_plates_for_kps(...)` (или `KpReadRepository` если split cleaner)
- `app/services/production_planning_service.py` — wire repository into planning entry
- Тесты: `tests/test_production_planning*.py`, lifecycle integration

**Acceptance:**
- [x] `rg "sqlite3" core/production/planning.py` → **0**
- [x] SQL для KP/plate load живёт только в repository layer
- [x] Create → build → activate flow без регрессии track count / version
- [x] No new raw `sqlite3.connect` in `core/` for plan flows

**Verify:**
```bash
pytest tests/test_production_planning*.py tests/test_plan_lifecycle_create_build_activate_orchestration.py -q
rg "sqlite3" core/production/planning.py
rg "sqlite3.connect" core/ --glob "*.py"
```

---

## WP3 — A4: narrow/remove PEP 562 proxy for web callers

**Цель:** web path не зависит от legacy `__getattr__` proxy; зафиксировать inventory оставшихся consumers.

**Файлы:**
- `core/config_and_data.py` — сузить `__getattr__`; DeprecationWarning для web-deprecated names
- `core/plate_runtime_state.py` — explicit state where web callers migrated
- Web callers inventory: `rg "config_and_data\.(PLATE|PLATES)" app/ frontend/` → migrate to session/DraftStore
- `ai_docs/develop/architecture/pep562-config-and-data-decommission.md` — create or update status checklist
- Тесты: regression on production/commercial flows using plate state

**Acceptance:**
- [x] Inventory document: web vs bot proxy consumers (count before/after)
- [x] **0** web/API handlers используют PEP 562 proxy attributes (grep gate or documented exceptions)
- [x] `__getattr__` proxy list сужен; bot paths не ломаются (minimal fix only)
- [x] ADR checklist: remaining proxy consumers → bot-only backlog

**Verify:**
```bash
pytest tests/test_production_api_integration.py tests/test_commercial*.py -q
rg "from core import config_and_data|config_and_data\." app/ --glob "*.py"
```

---

## WP4 — Q1–Q3, Q8: DRY pricing/phone + logging in core hot paths

**Цель:** убрать дублирование и `print()` в production-critical core modules.

**Файлы:**
- **Q1:** `core/commercial_offer.py`, `core/commercial_offer_xlsx.py` → shared `calculate_order_totals()` (e.g. `core/domain/pricing.py`)
- **Q2 (web):** `app/services/archive_service.py` — canonical KP→order_data mapper; dedupe web-only copies
- **Q3:** `core/utils/phone.py` — `format_phone()`; remove duplicates in offer modules
- **Q8:** replace `print()` in:
  - `core/production/planning.py`
  - `core/visualization.py`
  - `core/optimization/` hot paths (as touched by WP1–WP2)
- Тесты: existing commercial/archive suites; optional unit tests for shared helpers

**Acceptance:**
- [x] Single `format_phone()` implementation; grep → 1 definition
- [x] Order totals calculation — single source; both offer modules import it
- [x] Web archive mapping — no duplicate mapper outside `archive_service` (bot excluded per D4)
- [x] `rg "print\(" core/production/planning.py core/visualization.py core/optimization/ --glob "*.py"` → **0**
- [x] Logging uses module-level `logger`; no behaviour change in happy path

**Verify:**
```bash
pytest tests/test_commercial*.py tests/test_archive*.py -q
rg "def format_phone" core/ app/ --glob "*.py"
rg "print\(" core/production/planning.py core/visualization.py core/optimization/ --glob "*.py"
```

---

## WP5 (stretch) — S12: password change rate limit

**Цель:** brute-force mitigation на `POST /api/v1/auth/change-password` без Redis (D1 default).

**Файлы:**
- `app/security/login_rate_limit.py` — extract shared limiter or add `password_change_rate_limit` helper
- `app/api/v1/endpoints/auth.py` — apply limit on `change_password`
- `tests/test_password_policy.py` or new `tests/test_password_change_rate_limit.py`

**Acceptance (optional — не блокирует closure P4 если WP1–WP4 green):**
- [ ] 4th attempt within window → 429 with stable error message
- [ ] Limit keyed by IP (+ user id if authenticated)
- [ ] Test covers allow → block → window reset (or documented TTL)

**Verify:**
```bash
pytest tests/test_password_policy.py tests/test_password_change_rate_limit.py -q
```

---

## Приоритеты и порядок работ

| Порядок | WP | ID | Обоснование |
|---------|-----|-----|-------------|
| 1 | WP1 | A1 slice 2 | Закрывает critical A1 residual; разблокирует viz refactor |
| 2 | WP2 | A5 | Зависит от stable planning; repository boundary |
| 3 | WP3 | A4 | После A5 — меньше legacy state в planning path |
| 4 | WP4 | Q1–Q3, Q8 | Независимый quality pass; частично параллелен WP1–WP3 |
| 5 | WP5 | S12 | Stretch security |

**Параллелизация:** WP4 (Q3, Q1) может идти параллельно WP1 после port design frozen.

---

## Definition of Done (спринт)

- [x] WP1–WP4 acceptance выполнены (**обязательно**)
- [x] WP5 (S12) — optional stretch; не блокирует closure (deferred)
- [x] **Нет** bot-specific WP (кроме minimal PEP 562 fix if tests break)
- [x] Product decisions D1–D4 зафиксированы в changelog
- [x] `pytest tests/ -q` — **959 passed, 12 skipped**, 0 failed
- [x] Health Score web/core — **~6.5–7.0/10** (A1 full, A5 closed, A4 partial; bot backlog unchanged)
- [x] Spec status → `closed (WP5 stretch S12 deferred)`
- [x] Audit 2026-06-21 обновлён: cross-ref P4 closure items (v1.3)

---

## Риски

| Риск | Вероятность | Митигация |
|------|-------------|-----------|
| A1 slice 2 ломает procurement breakdown | Medium | Incremental port extraction; run procurement + layout suites each step |
| A5 repo methods меняют plan load semantics | Medium | Lifecycle integration test; compare KP/plate counts |
| A4 proxy removal ломает bot tests | Medium | Narrow web only; bot minimal fix; grep inventory |
| Q1 refactor меняет KP totals | Low | Golden-file or existing commercial tests |
| Scope creep в bot / Redis / PostgreSQL | Low | Explicit out-of-scope; defer to P5 |

---

## Следующий шаг

1. **Ревью v1** этой спеки — подтвердить D1–D4 и порядок WP.
2. **IMPLEMENT P4:** WP1 (A1 slice 2) → WP2 (A5) → WP3 (A4) ∥ WP4 (Q1–Q3, Q8) → WP5 stretch (S12).
3. **P5 preview:** S1 Redis, S3 OCR consent, A9 PostgreSQL, Q6 CreatePlanWizard, A7 god-module splits, bot removal decision.

---

## Changelog

| Дата | Изменение |
|------|-----------|
| 2026-06-21 | v1 draft — P4 после closure P3 (955 passed); scope: A1 slice 2, A5, A4, Q1–Q3/Q8, stretch S12 |
| 2026-06-21 | **WP1 complete** — A1 slice 2: procurement + drawing ports; `core/` ↛ `viz_modules` runtime imports; ADR + import-boundary gate updated |
| 2026-06-21 | **WP2 complete** — A5: `PlanLoadPort` + `PlanRepository.fetch_*`; `core/production/planning.load` без raw `sqlite3.connect` |
| 2026-06-21 | **WP3 complete** — A4: PEP 562 proxy narrowed 32→16; app/core grep gate 0; ADR inventory updated |
| 2026-06-21 | **WP4 complete** — Q1–Q3, Q8: shared `calculate_order_totals`, `format_phone`, web archive mapper; `print()` → logging in core hot paths |
| 2026-06-21 | **P4 CLOSED** — WP1–WP4 done; WP5 (S12) deferred; `pytest tests/ -q` → **959 passed, 12 skipped, 0 failed**; Health web/core ~6.5–7.0/10 |

---

*Создано: 2026-06-21 · v1 draft: web/core deferred backlog, bot frozen, baseline 955 passed / Health ~6.0/10 web/core. · **Closed 2026-06-21:** 959 passed / Health ~6.5–7.0 web/core.*
