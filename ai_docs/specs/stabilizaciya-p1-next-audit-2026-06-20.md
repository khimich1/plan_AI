# Spec: стабилизация P1-next — аудит 2026-06-20

> **Тип:** remediation feature-spec (стабилизационный спринт)
> **Фаза SDD:** SPECIFY (черновик на ревью)
> **Дата:** 2026-06-20
> **Ревизия:** v2 — учёт bot deprecated (2026-06-20)
> **Статус:** draft
> **Источник:** [`../develop/audits/2026-06-20-full-project-audit.md`](../develop/audits/2026-06-20-full-project-audit.md)
> **Predecessor (закрыт):** [`stabilizaciya-p0-audit-2026-06-20.md`](./stabilizaciya-p0-audit-2026-06-20.md) — WP1–WP3 (partial A3)
> **Bot policy (наследуется):** [`stabilizaciya-p0-audit-2026-06-19.md`](./stabilizaciya-p0-audit-2026-06-19.md) § «Решение по Telegram-боту»

---

## Стратегия (одной фразой)

> Закрыть **web/API security** high-priority (S4, S6, S2, S3 по ID аудита 20.06). Telegram-бот **deprecated / out of active use** — дальнейшие инвестиции в bot **минимальны** (maintenance-only, без новых WP).

---

## Bot deprecation strategy

| Тема | Решение |
|------|---------|
| **Статус** | Bot **deprecated** с 2026-06-19. Активная поддержка и новые фичи **не** планируются. Единственный канал — **web (React SPA + FastAPI API)**. |
| **Что оставить** | Код `bot/` в репо **заморожен** (read-only archive). SQLite authority через `plan_storage` shim (P0-next WP1) — чтобы legacy paths не писали в JSON. Существующие parity/guard-тесты — **freeze** (не расширять). |
| **Что не трогать** | God-modules bot (`commercial.py`, `production_execution.py`), bare `except` (Q1 20.06), bot integration tests (Q3 20.06), A4/A5 bot→core — **вне scope**, пока нет решения об удалении `bot/`. |
| **`run_bot.py`** | Оставить с **WARNING** при старте (`run_bot.py`, `bot/bot_main.py`). **Не** использовать в production. Полное отключение entrypoint — опциональный follow-up (Ask first). |
| **`BOT_AUTH_ENABLED`** | В production — **true** (fail-closed). `BOT_AUTH_ENABLED=false` — только `APP_ENV=development` (см. `core/config/settings.py`). При полном отказе от bot: можно зафиксировать `BOT_AUTH_ENABLED=false` + не деплоить `run_bot.py` — отдельное согласование. |
| **P0-next WP2 (sunk cost)** | Bot adapter + parity tests **уже сделаны** — maintenance-only; **не** продлевать cross-surface parity как product goal. |

---

## Objective

Закрыть **высокоприоритетный web/API security кластер** из аудита 2026-06-20, пока data plane планов стабилизирован в P0-next (SQLite authority). Bot reliability **не** цель спринта.

## Выбор scope: security vs bot vs architecture

| Трек | IDs | Effort | Решение |
|------|-----|--------|---------|
| **In scope (этот спринт)** | **S4**, **S6**, **S2**, **S3** | S–M | Прямой production-risk на **web/API**; логичное продолжение после P0-next |
| **Cancelled (bot deprecated)** | **Q1**, **Q3** (bot, аудит 20.06) | — | Закрыты **deprecation**, как Q2 в P0-2026-06-19; фикс bare `except` / bot integration — не делаем |
| **Deferred** | **A4, A5, A9, Q6** (bot architecture/commercial) | M–L | Bot frozen; только при решении об удалении `bot/` |
| **Deferred** | **A3 phase 2** (PEP 562 decommission) | M | После security hardening |
| **Out of scope** | **S5** frontend RBAC server re-check (частично), **S7** XFF (частично mitigated P2/P3) | — | Backlog P2 |

**Решение (v2):** **P1-next = web/API security only**. Bot-specific WP (бывшие WP3–WP4) **удалены** из scope.

## Scope (P1-next)

| Приоритет | ID | Проблема | Fix |
|-----------|-----|----------|-----|
| P0 | **S4** | GET `/web/logout` — CSRF-prone session teardown | POST-only logout + CSRF token; убрать side-effect GET |
| P0 | **S6** | Destructive admin reset без единого production-guard | Расширить `destructive_db_guard` на все destructive admin/archive paths; fail-closed в prod/staging |
| P1 | **S2** | Stateless sessions — нет server-side invalidation | Явный idle/max age; invalidate on logout; документировать ограничения |
| P1 | **S3** | In-process rate limit не работает при нескольких воркерах | Shared store **или** явный single-worker constraint в deployment docs + health check |

**Resolved by deprecation (не в scope):**

| ID | Почему |
|----|--------|
| **Q1** (bare `except` в bot) | Bot deprecated; риск ≈ 0 при отсутствии production-запуска |
| **Q3** (bot integration tests) | Parity tests из P0-next WP2 — **freeze**; новые bot integration scenarios не пишем |

**Out of scope (explicit):** A3 phase 2, A4/A5/A9, bot commercial consolidation, S5 full server RBAC re-check.

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

## WP3 — S2: session invalidation on logout

**Acceptance:**
- [ ] Logout (WP1) **инвалидирует** server-side session / cookie (не только client-side clear)
- [ ] Документированы idle timeout и max session age (или явное ограничение stateless model)
- [ ] Тест: после logout старый session cookie не даёт доступ

**Verify:** `pytest tests/test_app_session.py tests/test_web_logout_csrf.py -q`

---

## WP4 — S3: rate limit multi-worker

**Acceptance:**
- [ ] **Вариант A (preferred при multi-worker):** shared store для login/OCR rate limits **или**
- [ ] **Вариант B (single instance):** deployment docs + health check фиксируют single-worker; rate limit остаётся in-process
- [ ] Тесты не регрессируют существующий login rate limit

**Verify:** `pytest tests/test_auth_login_rate_limit.py tests/test_web_login_rate_limit.py -q`

---

## WP5 (stretch) — A3 phase 2 backlog prep

**Acceptance (optional):**
- [ ] ADR или checklist: шаги decommission PEP 562 proxy в `core/config_and_data.py`
- [ ] Нет обязательного code removal в P1-next

---

## Definition of Done (спринт)

- [ ] WP1–WP2 acceptance выполнены (обязательно)
- [ ] WP3–WP4 — минимум один из двух закрыт или задокументирован deployment constraint (S3 variant B)
- [ ] **Нет** новых bot-specific WP или расширения parity tests
- [ ] `pytest tests/ -q` зелёный
- [ ] Spec status → `closed` или `closed (stretch deferred)`

## Следующий шаг (актуально на 2026-06-20)

1. **Ревью v2** этой спеки — подтвердить отказ от bot WP (Q1/Q3).
2. **IMPLEMENT P1-next:** WP1 (S4) → WP2 (S6) → WP3 (S2) → WP4 (S3).
3. **Не планировать:** bot bare except, bot integration, A4 bot→core — до решения об удалении `bot/`.

## Следующий спринт после P1-next (preview)

1. **P2 architecture (web/core):** A5 god-modules **web-side**, A10 planning orchestration, A3 full PEP 562 decommission
2. **P2 security:** S7 XFF hardening review, S5 server-side RBAC re-check для sensitive actions
3. **Optional cleanup:** полное удаление `bot/` + `run_bot.py` — **отдельное согласование**, не блокер

---

*Создано: 2026-06-20 · v2: 2026-06-20 — bot deprecated, scope пересмотрен (security-first).*
