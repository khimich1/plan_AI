# Spec: стабилизация P1-next — аудит 2026-06-20

> **Тип:** remediation feature-spec (стабилизационный спринт)
> **Фаза SDD:** SPECIFY (черновик на ревью)
> **Дата:** 2026-06-20
> **Статус:** draft
> **Источник:** [`../develop/audits/2026-06-20-full-project-audit.md`](../develop/audits/2026-06-20-full-project-audit.md)
> **Predecessor (закрыт):** [`stabilizaciya-p0-audit-2026-06-20.md`](./stabilizaciya-p0-audit-2026-06-20.md) — WP1–WP3 (partial A3)

---

## Objective

Закрыть **высокоприоритетный security/quality кластер P1** из аудита 2026-06-20, пока архитектурный фундамент планов (A1/S1, A2) стабилизирован в P0-next.

## Выбор scope: P1 security (S4/S6/Q1) vs A4/A5

| Трек | IDs | Effort | Почему сейчас / позже |
|------|-----|--------|------------------------|
| **Рекомендуется (этот спринт)** | **S4, S6, Q1** (+ точечно **Q3**) | S–M | Малый объём, прямой production-risk; логичное продолжение после унификации bot/web planning в P0-next |
| Отложить | **A4, A5** | M–L | Крупные рефакторинги (DIP bot→core, god-modules); выигрывают от «замороженного» data plane и меньше регрессий после security hardening |

**Решение:** стартовать **P1-next security + bot reliability**, не A4/A5. A4/A5 — отдельный **P2 architecture** спринт после POST-logout, полного destructive-guard и чистки bare `except` на bot hot paths.

## Scope (P1-next)

| ID | Проблема | Fix |
|----|----------|-----|
| **S4** | GET `/web/logout` — CSRF-prone session teardown | POST-only logout + CSRF token; убрать side-effect GET |
| **S6** | Destructive admin reset без единого production-guard | Расширить `destructive_db_guard` на все destructive admin/archive paths; fail-closed в prod/staging |
| **Q1** | 22+ bare `except` в bot | Заменить на typed exceptions + logging; не глотать ошибки на production/commercial/planning paths |
| **Q3** | Критичные bot-пути без integration safety net | Минимальный набор integration tests поверх P0-next parity (planning + plan persist) |

**Out of scope (explicit):** A3 phase 2 (PEP 562 decommission), A4 DIP, A5 god-module split, S2/S3 shared rate-limit store, S5 frontend RBAC server enforcement.

---

## WP1 — S4: POST logout + CSRF

**Acceptance:**
- [ ] Logout доступен только через **POST** (GET возвращает 405 или redirect без invalidate session)
- [ ] Cookie-based logout требует валидный CSRF token (как login/forms)
- [ ] Шаблоны/frontend: форма или fetch POST с token
- [ ] Тесты: `tests/test_web_logout_csrf.py` (или расширение csrf suite) — positive/negative

**Verify:** `pytest tests/test_csrf.py tests/test_web_logout_csrf.py -q`

---

## WP2 — S6: full destructive guard

**Acceptance:**
- [ ] Инвентаризация всех destructive endpoints (admin reset, archive wipe, bulk delete)
- [ ] Каждый path вызывает единый guard (`ALLOW_DESTRUCTIVE_DB_RESET` / env fail-closed)
- [ ] Production/staging: 403 без явного override; audit log при deny
- [ ] Тесты покрывают ранее неохваченные routes (дополнить `test_admin_destructive_guard.py`)

**Verify:** `pytest tests/test_admin_destructive_guard.py -q`

---

## WP3 — Q1: bare except в bot (hot paths)

**Acceptance:**
- [ ] `production_execution`, `production_export`, `commercial` (top handlers): нет голых `except:` / `except Exception: pass`
- [ ] Ошибки логируются с context (user_id, callback data); пользователю — безопасное сообщение
- [ ] Регрессии: существующие bot unit tests зелёные

**Verify:** `pytest tests/test_bot_production_planning_parity.py tests/test_bot_plan_sqlite_authority.py -q` + grep/linters по bot/handlers

---

## WP4 — Q3: bot integration safety net

**Acceptance:**
- [ ] Integration test: bot adapter → SQLite plan persist → read back через repository (end-to-end без real Telegram)
- [ ] Integration test: planning failure surfaces controlled error (не silent success)
- [ ] Документировано в spec changelog, какие paths покрыты

**Verify:** `pytest tests/test_bot_plan_sqlite_authority.py tests/test_bot_production_planning_parity.py -q`

---

## WP5 (stretch) — A3 phase 2 backlog prep

**Acceptance (optional в этом спринте):**
- [ ] ADR или checklist: шаги decommission PEP 562 proxy в `core/config_and_data.py`
- [ ] Нет обязательного code removal в P1-next; только plan + 1 strangler PR если остаётся capacity

---

## Definition of Done (спринт)

- [ ] WP1–WP3 acceptance выполнены; WP4 — минимум 1 новый integration scenario
- [ ] `pytest tests/ -q` зелёный
- [ ] Spec status → `closed` или `closed (stretch deferred)`

## Следующий спринт после P1-next (preview)

1. **P2 architecture:** A4 (bot→core, не app), A5 (decompose `commercial.py` / `production_execution` остатки)
2. **A3 full:** PEP 562 proxy removal по [`plate-runtime-isolation.md`](../develop/architecture/plate-runtime-isolation.md)
3. **P2 security:** S2/S3 shared rate limit, S7 hardening review

---

*Создано: 2026-06-20 (draft после closure P0-next).*