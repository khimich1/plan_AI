# PLAN: Стабилизация P7 — god-modules, arch hygiene, medium backlog

> **Фаза SDD:** PLAN → TASKS → IMPLEMENT
> **Дата:** 2026-06-22
> **Спека:** [`../../specs/stabilizaciya-p7-architecture-2026-06-22.md`](../../specs/stabilizaciya-p7-architecture-2026-06-22.md)
> **Предшественник (закрыт):** [`../../specs/stabilizaciya-p6-architecture-2026-06-21.md`](../../specs/stabilizaciya-p6-architecture-2026-06-21.md)
> **Источник:** [`../audits/2026-06-21-frontend-backend-audit.md`](../audits/2026-06-21-frontend-backend-audit.md) — Remediation status (open → P7)

---

## 0. Текущее состояние (post-P6)

| Метрика | Значение |
|---------|----------|
| Pytest | **953 passed**, 9 skipped, 0 failed |
| Frontend tests | **55** passed |
| Frontend build | green |
| Critical (web) | **0** (A1–A3 resolved) |
| Health Score (estimate) | ~6.5–7.0/10 |
| Deploy contract | [`deploy-contract.md`](../deploy-contract.md) — `workers=1` until Redis |
| Git | P2 closure committed; **P3–P6 largely uncommitted** in working tree |

**Закрыто в P6:** CPU offload, visualization ports, plate runtime slice, destructive guard, `RequireRole`, deploy contract, OCR flag, `AuthService`, `offers_write` slice, Q2–Q4.

**Открыто → P7:** god-modules slices 4/6, planning SSOT, DI protocols, legacy routes, WP9 medium backlog, Redis, infra stretch.

---

## 1. Цель спринта

Закрыть оставшиеся **High** (god-modules, planning/DI/legacy) и ≥3 **Medium** clusters; подготовить Redis path или ADR prerequisite; поднять Health Score к **~7.5–8.5/10** без регрессии тестов.

---

## 2. Git hygiene (checkpoint 0 — до кода P7)

> **Важно:** рабочее дерево содержит незакоммиченные изменения P3–P6. Перед WP2 зафиксировать baseline.

**Рекомендуемые шаги:**

1. Review `git status` / `git diff` — убедиться, что scope = P3–P6 stabilization only.
2. Создать один или несколько логических коммитов (например: P3, P4, P5, P6) **по запросу владельца репо**.
3. Опционально: tag `p6-closure` на коммите с verify 953 passed.
4. Запустить verify commands (§6) на чистом HEAD.

**Без коммита P3–P6:** риск смешения baseline P7 с незавершённым diff; новые PR P7 сложнее review.

---

## 3. WP execution order

### Граф зависимостей

```
Checkpoint 0 (git hygiene)
        │
        ▼
WP2 slice 4 (viz drawing/export)     [старт — recommended first]
        │
        ▼
WP2 slice 6 (workflow draft/export)
        │
        ├──────────────────┐
        ▼                  ▼
WP3 (planning/DI/legacy)   WP4 A18 (response_model)  [параллельно после slice 6]
        │                  │
        └────────┬─────────┘
                 ▼
WP4 (A14 DraftStore, A13 PEP562 prep, …)
                 │
                 ▼
WP1 (Redis implementation OR ADR-only)  [параллельно с WP4 если infra ready]
                 │
                 ▼
WP5 stretch (bot delete, CSP, PostgreSQL preview)  [optional]
```

### Checkpoints

| # | Checkpoint | Gate | Verify |
|---|------------|------|--------|
| **C0** | Git baseline | P3–P6 committed or explicitly scoped branch | `git status` clean for P7 branch |
| **C1** | WP2 slice 4 | `core/visualization/__init__.py` −200 LOC; tests green | `pytest tests/test_layout_*.py tests/test_core_viz_import_boundary.py -q` |
| **C2** | WP2 slice 6 | `commercial_workflow_service.py` < 500 LOC | `pytest tests/test_commercial*.py -q` |
| **C3** | WP3 | Planning SSOT + legacy 410; DI protocols ≥2 services | `pytest tests/test_production_api_integration.py tests/test_auth*.py -q` |
| **C4** | WP4 | ≥3 medium clusters (A13/A14/A18) | `pytest tests/ -q` full suite |
| **C5** | WP1 | Redis shared limits **or** ADR signed | `pytest tests/test_rate_limit_deployment.py -q` |
| **C6** | Closure | DoD P7 spec | §6 full verify |

**Рекомендуемый первый WP:** **WP2 slice 4** (viz drawing/export) — продолжение P6 без infra deps, сильный test net.

---

## 4. Work packages (кратко)

| WP | Приоритет | Effort | Ключевой результат |
|----|-----------|--------|-------------------|
| **WP2** | P1 | L | God-modules slices 4/6 (+7 stretch) |
| **WP3** | P1 | L | `core/production/planning` SSOT; protocols; legacy 410 |
| **WP4** | P1/P2 | M–L | ≥3 medium clusters (PEP562, DraftStore, response_model) |
| **WP1** | P0/P1 | M | Redis rate limit **or** infra ADR |
| **WP5** | P2 | L | PostgreSQL preview, bot hard delete, CSP, Argon2id (stretch) |

Детальные acceptance criteria — в [спеке P7](../../specs/stabilizaciya-p7-architecture-2026-06-22.md).

---

## 5. Product decisions (нужно подтвердить)

| ID | Вопрос | Default |
|----|--------|---------|
| **D1** | Redis в этом спринте? | Implement if URL available; else ADR-only |
| **D2** | Порядок god-modules | Viz slice 4 → workflow slice 6 |
| **D5** | Hard delete `bot_archived/` | WP5 stretch only |
| **D6** | PostgreSQL | Preview doc only, no cutover |
| **D7** | PEP 562 `__getattr__` delete | End of WP4 or defer P8 |

---

## 6. Verify commands

**Backend (baseline):**

```powershell
cd c:\Users\Роман\Desktop\Шишов
.\venv\Scripts\activate
pytest tests/ -q
# Expect: >= 953 passed, 9 skipped, 0 failed
```

**Frontend:**

```powershell
cd frontend
npm run build
npm run test
# Expect: 55 tests passed, build OK
```

**Per-WP (examples):**

```powershell
# WP2
pytest tests/test_layout_*.py tests/test_core_viz_import_boundary.py tests/test_commercial*.py -q

# WP3
pytest tests/test_production_api_integration.py tests/test_plan_consistency.py tests/test_auth*.py -q
rg "from app.planning" app/services --glob "*.py"

# WP1
pytest tests/test_login_rate_limit.py tests/test_rate_limit_deployment.py -q

# WP4
rg "DraftStore\(\)" app/api/ --glob "*.py"
```

**Grep gates (post-WP2/WP3):**

```powershell
rg "from core.visualization import visualize_plan" app/ --glob "*.py"
rg "get_plate_mutable_runtime" app/ --glob "*.py"
```

---

## 7. Риски и митигации

| Риск | Митигация |
|------|-----------|
| Uncommitted P3–P6 | Checkpoint C0 before WP2 |
| Redis unavailable | D1 ADR path; keep deploy contract |
| Scope creep WP5 | Stretch only; does not block closure |
| God-module regressions | One slice per PR; full pytest each checkpoint |

---

## 8. Definition of Done (plan level)

- [ ] Checkpoints C1–C4 complete
- [ ] WP1 resolved (implement or ADR)
- [ ] P7 spec DoD satisfied
- [ ] Audit Remediation status updated
- [ ] Plan status → completed (link to implementation report when available)

---

*Создано: 2026-06-22 · Plan for P7 stabilization sprint; baseline 953 pytest / 55 frontend tests.*
