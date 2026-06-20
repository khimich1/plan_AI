# Spec: Стабилизация P1 — runtime isolation + bot auth (аудит 2026-06-19)

> **Тип:** remediation feature-spec (стабилизационный спринт)
> **Фаза SDD:** IMPLEMENT — **закрыт**
> **Дата:** 2026-06-19
> **Ревизия:** v2 (closure)
> **Статус:** **closed (implemented) 2026-06-19**
> **Baseline:** [`project-baseline.md`](./project-baseline.md)
> **Предшественник:** [`stabilizaciya-p0-audit-2026-06-19.md`](./stabilizaciya-p0-audit-2026-06-19.md) — **закрыт**
> **Источник находок:** [`../develop/audits/2026-06-19-full-project-audit.md`](../develop/audits/2026-06-19-full-project-audit.md) → Post-P0

---

## Стратегия (одной фразой)

> Устранить оставшийся critical architecture (runtime globals → `PlateOrderContext`/DI) и закрыть дыру bot auth (fail-closed, без synthetic admin в production).

**Контекст после P0:** Health Score **~3/10** (было 0/10). Закрыты data-integrity (`Q1`/`Q3` спеки P0), единое хранилище планов (`A1` аудита), pipeline в `core/` для web (`A2` аудита).

**Итог P1 (2026-06-19):** закрыты **2 remaining critical** — `A3`, `S1`; связанный **S5** (утечки через global runtime на hot paths). Health Score **~6/10** (см. Post-P1 в audit). **WP5 (S2 rate limit)** — deferred.

**Решение по Telegram-боту:** бот **deprecated** (P0), не цель развития. **`S1` всё равно в scope P1** — при ошибочном `APP_ENV` или запуске бота в prod misconfiguration даёт полный admin-доступ. Код бота не удаляем.

---

## ASSUMPTIONS I'M MAKING

1. **P0 закрыт** — `pytest tests/ -q` зелёный (~726 passed на момент закрытия P0); планы в SQLite через `PlanRepository`; web вызывает `core/production/planning.py`.
2. **Деплой по умолчанию — single instance** (`APP_STORAGE_LAYOUT=single_instance`); горизонтальное масштабирование без shared state **не цель P1**, но globals не должны ломать конкурентные запросы в одном процессе.
3. **Частичная инфраструктура A3 уже есть:** `core/plate_order_context.py`, `core/plate_runtime_state.py` (thread-local + ContextVar), `app/middleware/plate_runtime_isolation.py`, `app/dependencies/plate_context.py`, тест `tests/test_plate_mutable_runtime_isolation.py`. P1 — **довести до production-grade**, а не писать с нуля.
4. **Частичный S1 уже есть:** `Settings.validate_bot_telegram_auth` запрещает `BOT_AUTH_ENABLED=false` при `APP_ENV=production`; middleware в prod при disabled auth не вызывает handler. **Остаётся:** synthetic admin в non-production paths без явного dev-guard, `bot_main` не делает `sys.exit(1)`, нет полного набора middleware/startup тестов по аудиту.
5. **Strangler для A3:** legacy `import core.config_and_data as cfg` + PEP 562 `__getattr__` сохраняются на переходный период; новый код и hot paths получают явный `PlateOrderContext` через DI/`bound()`. Полное удаление `cfg.PLATES_*` — **после P1** (backlog).
6. **`_visualize_lock`** в `day_documents_service` / `archive_service` — симптом globals; цель P1 — убрать после гарантированной изоляции контекста на hot paths.
7. **S2/S3, frontend 409 reload, декомпозиция bot god-modules** — **deferred** (P1.1 или optional WP). Не блокируют закрытие P1.
8. PLAN — [`../develop/plans/2026-06-19-stabilizaciya-p1.md`](../develop/plans/2026-06-19-stabilizaciya-p1.md). **WP0–WP4 реализованы**; closure подтверждён `pytest tests/ -q` → **744 passed, 12 skipped**.

→ Допущения подтверждены в ходе IMPLEMENT; остаточный риск — cold paths (`cfg.PLATES_*` proxy), optional **S2** (WP5).

---

## 1. Objective & Problem Statement

### Objective

Закрыть **оставшиеся 2 critical-находки** аудита (`A3`, `S1`) и связанный **S5** (перекрёстная утечка через global plate runtime). Поднять Health Score с **~3/10** к **~5–6/10** за счёт устранения critical architecture + security hole.

### Problem Statement (после P0)

1. **Глобальное мутабельное состояние плит (`A3`, `S5`):** `core/config_and_data.py` проксирует `PLATES_*` / `PLATE_*` в process-wide runtime (`core/plate_runtime_state.py`). Часть web-путей обёрнута middleware, но optimization/viz/commercial hot paths всё ещё мутируют legacy globals; `_visualize_lock` сериализует параллельные запросы. Риск утечки данных между сессиями и блокер масштабирования.
2. **Admin bypass в боте (`S1`):** при `BOT_AUTH_ENABLED=false` в non-production middleware выдаёт synthetic `BotUser(role="admin")` всем. При misconfiguration (`APP_ENV` не `production`, но бот доступен извне) — полная компрометация. Startup не гарантирует fail-closed (`sys.exit`).

### Reframe: что значит «стало лучше»

| Требование | Reframed success criteria |
|------------|---------------------------|
| «Убрать globals» | Hot paths FastAPI не полагаются на неявный module-level state; `PlateOrderContext` передаётся явно или через `request.state` + `bound()` |
| «Нет утечки между запросами» | Параллельные asyncio-задачи / HTTP-запросы с разными заказами не видят чужие `PLATES_*` (тест изоляции на production endpoints) |
| «Закрыть S1» | В production невозможен старт бота без auth; synthetic admin только при явном `APP_ENV=development` + warning; иначе fail-closed |
| «Убрать locks как костыль» | `_visualize_lock` удалён или заменён на no-op там, где контекст изолирован |

---

## 2. Scope

### In scope — Этап A: Security (можно стартовать сразу)

| ID | Категория | Название | Механизм |
|----|-----------|----------|----------|
| **S1** | Безопасность | Synthetic admin bypass в Telegram-боте | Fail-closed auth; startup guard; убрать admin bypass вне явного development |
| *(часть S1)* | Безопасность | Startup guard бота | `validate_bot_startup()` → `sys.exit(1)` при fatal misconfiguration |

### In scope — Этап B: Runtime isolation (`A3` + `S5`)

| ID | Категория | Название | Механизм |
|----|-----------|----------|----------|
| **A3** | Архитектура | Глобальное мутабельное plate runtime | Инвентаризация → request-scoped `PlateOrderContext`/DI на FastAPI hot paths → миграция optimization/config hot path → тесты изоляции |
| **S5** | Безопасность | Перекрёстная утечка через global runtime | Решается вместе с A3; integration-тесты параллельных запросов |

### Deferred (P1.1 / optional WP — вне обязательного closure P1)

| ID / тема | Причина переноса |
|-----------|------------------|
| **S2** | Rate limiting на login — optional WP5 при наличии времени |
| **S3** | Object-level RBAC для КП — отдельный security-спринт |
| Frontend 409 reload | UX для `plan_version_conflict`; `apiError.ts` уже парсит `code` |
| Bot god-module decomposition (`A6`) | Бот deprecated; не блокер P1 |
| Полное удаление `cfg.PLATES_*` / PEP 562 | Strangler: после hot paths; отдельный backlog |
| **A8** | Унификация bot vs API runtime — частично покрыто middleware бота; полная унификация не цель при deprecated bot |

### Out of scope (в этой спеке)

- Новые продуктовые фичи.
- PostgreSQL, Redis session store, горизонтальное масштабирование.
- **S4** (npm CVE), **S6–S18**, **A4–A7**, **A9–A26**, **Q*** (кроме тестов изоляции как часть A3).
- Полное удаление кода бота.
- Нормализация треков планов (`A7`).

---

## 3. Tech Stack

Без изменений относительно baseline (см. [`project-baseline.md`](./project-baseline.md)).

Ключевое для P1:
- `PlateOrderContext` + `plate_mutable_runtime_scope` / `optimization_context_scope` (`core/plate_order_context.py`).
- FastAPI middleware `PlateMutableRuntimeIsolationMiddleware` + `Depends(get_plate_order_context)`.
- Bot: `bot/middleware/auth.py`, `bot/middleware/plate_runtime_isolation.py`, `core/config/settings.py` (`BOT_AUTH_*`).
- pytest + FastAPI `TestClient` + asyncio isolation tests.

---

## 4. Commands

```powershell
Set-Location "c:\Users\Роман\Desktop\Шишов"
.\.venv\Scripts\Activate.ps1

# Этап A — S1 (bot auth)
pytest tests/test_bot_auth.py -q

# Этап B — A3 / S5 (runtime isolation)
pytest tests/test_plate_mutable_runtime_isolation.py -q
pytest tests/test_plate_runtime_request_isolation.py -q          # новый (WP4)
pytest tests/test_production_day_documents_isolation.py -q       # новый (WP4), если создан

# Границы слоёв
pytest tests/test_core_no_app_import.py -q

# Регрессия перед закрытием спринта
pytest tests/ -q

# Frontend (не блокер P1, unless touched)
cd frontend; npm run test; npm run build
```

---

## 5. Acceptance Criteria по находкам

### ЭТАП A — S1: fail-closed bot auth

#### S1 — Synthetic admin bypass

- **Где:** `bot/middleware/auth.py` (~69–78), `bot/bot_main.py` (`validate_bot_startup`), `core/config/settings.py` (`validate_bot_telegram_auth`)
- **Текущее состояние (post-P0 partial):** Pydantic запрещает `BOT_AUTH_ENABLED=false` + `APP_ENV=production`; middleware в production при disabled auth не вызывает handler. **Synthetic admin остаётся** при `BOT_AUTH_ENABLED=false` и `APP_ENV != production`.
- **Acceptance:**
  - [x] Synthetic `BotUser(role="admin")` выдаётся **только** при одновременном: `APP_ENV=development` **и** `BOT_AUTH_ENABLED=false` **и** явный dev-mode (документировано в `bot/README.md`).
  - [x] Любой другой `APP_ENV` (включая `staging`, `test`, пустой, опечатки) при `BOT_AUTH_ENABLED=false` → **fail-closed** (handler не вызывается, security event в лог).
  - [x] `APP_ENV=production` + disabled auth → бот **не стартует**: `validate_bot_startup()` → `sys.exit(1)` (не только `return` из `main`).
  - [x] `APP_ENV=production` + `BOT_AUTH_ENABLED=true` + пустой allowlist (при `bot_auth_fail_closed_enabled`) → старт запрещён (`sys.exit(1)` или `ValidationError` до polling).
  - [x] Warning в логах при dev open-access (сохранить/усилить существующий warning в `bot_main`).
  - [x] Бот помечен deprecated — warning при старте сохраняется.
- **Verify:**
  - `tests/test_bot_auth.py` — расширить: middleware production+disabled auth; staging+disabled auth → denied.
  - Новый тест: `validate_bot_startup` / entrypoint при production misconfig → exit code ≠ 0 (monkeypatch `sys.exit`).

### ЭТАП B — A3 + S5: runtime globals → context/DI

#### A3 / S5 — Глобальное мутабельное состояние и утечки

- **Где (ключевые):**
  - `core/plate_runtime_state.py`, `core/config_and_data.py` (PEP 562 legacy proxy)
  - `core/plate_order_context.py`, `app/middleware/plate_runtime_isolation.py`, `app/dependencies/plate_context.py`
  - Hot paths: `app/services/day_documents_service.py`, `app/services/archive_service.py`, `app/services/commercial_workflow_service.py`, `core/optimization/*`, `viz_modules/*` (через `cfg`)
  - Locks: `app/services/day_documents_service.py` → `_visualize_lock` (и reuse в `archive_service`)
- **Acceptance — инвентаризация (WP1):**
  - [x] Документирован артефакт `ai_docs/develop/architecture/plate-runtime-globals-inventory.md`: все call sites мутации `cfg.*` / `get_plate_mutable_runtime()` с пометкой hot/cold и владельцем миграции.
- **Acceptance — FastAPI hot paths (WP2):**
  - [x] Production/commercial endpoints, мутирующие plate state, получают `PlateOrderContext` через `Depends(get_plate_order_context)` или явный параметр от `request.state`.
  - [x] `day_documents_service` / `archive_service` visualization path выполняется внутри `ctx.bound()` контекста запроса (не создаёт orphan `fresh_empty()` без bind, если запрос уже изолирован middleware).
  - [x] Нет регрессии: существующие production API тесты зелёные.
- **Acceptance — optimization/config hot path (WP3):**
  - [x] Критический путь оптимизации (commercial preview, day documents generation) не зависит от неявного global default/demo order вне `bound()`.
  - [x] `run_in_order_context` / `PlateOrderContext.bound()` используется в сервисах, вызываемых из API, где ещё остаётся legacy `import cfg`.
  - [x] `core/` не импортирует `app/` (`tests/test_core_no_app_import.py` зелёный).
- **Acceptance — изоляция и locks (WP4):**
  - [x] Integration-тест: два параллельных HTTP-запроса с разным plate text → ответы не смешивают `PLATES_*` / diagnostics (`tests/test_plate_runtime_request_isolation.py`).
  - [x] `tests/test_plate_mutable_runtime_isolation.py` расширен request-level тестами.
  - [x] `_visualize_lock` **удалён** из hot path (`day_documents_service`, `archive_service`).
- **Verify:** новые/расширенные тесты изоляции; `pytest tests/ -q` зелёный; ручная проверка commercial preview + day documents под параллельными запросами (опционально).

---

## 6. Testing Strategy

| Находка | Тип теста | Файл |
|---------|-----------|------|
| S1 | Unit (settings + middleware + startup) | `tests/test_bot_auth.py` (расширить) |
| A3/S5 | Unit (context scopes) | `tests/test_plate_mutable_runtime_isolation.py` (расширить) |
| A3/S5 | Integration (HTTP parallel) | `tests/test_plate_runtime_request_isolation.py` (новый) |
| A3 | Service-level | `tests/test_production_day_documents_isolation.py` или внутри существующих production tests |
| Регрессия | Full suite | `pytest tests/ -q` |

Safety-net: тесты изоляции **до** удаления `_visualize_lock`.

---

## 7. Boundaries

### Always
- Сначала тест, фиксирующий поведение (особенно S1 fail-closed и A3 isolation), затем рефакторинг.
- Порядок этапов: **A (S1)** можно параллельно с **B0 (WP1 inventory)**; **B2–B4** строго после инвентаризации hot paths.
- `PlateOrderContext.bound()` на время мутации legacy `cfg` в request scope.
- `logging` + security audit events для отказов auth (`log_bot_security_event`).
- `core/` не импортирует `app/`.
- Бот остаётся deprecated; не расширять функциональность бота в P1.

### Ask first
- Изменение семантики `APP_ENV` / новых значений enum окружения.
- Удаление `_visualize_lock` если тесты изоляции нестабильны на CI — согласовать временный feature flag.
- Новые pip-зависимости (например `slowapi` для optional S2).
- Публичные изменения API-контрактов (не ожидаются в P1).

### Never
- Восстанавливать synthetic admin «для удобства» вне явного `APP_ENV=development`.
- Отключать `PlateMutableRuntimeIsolationMiddleware` в production paths.
- Удалять падающие тесты изоляции без согласования.
- Тащить S2/S3/bot decomposition в обязательный closure P1.
- Коммитить без явной просьбы пользователя.

---

## 8. Success Criteria (спринт «готово»)

- [x] Человек прочитал ASSUMPTIONS и подтвердил/поправил.
- [x] **Этап A закрыт:** `S1` по acceptance; `tests/test_bot_auth.py` зелёный; production misconfig → бот не стартует.
- [x] **Этап B закрыт:** `A3`/`S5` — hot paths на `PlateOrderContext`/DI; инвентаризация готова; тесты параллельной изоляции зелёные; `_visualize_lock` убран.
- [x] `pytest tests/ -q` зелёный (**744 passed, 12 skipped**); `tests/test_core_no_app_import.py` зелёный.
- [x] Health Score пересчитан в audit → critical **0** из исходных 4 (A1, A2, A3, S1); остаточный риск — cold legacy `cfg` proxy, high-priority backlog.
- [x] Бот по-прежнему deprecated; S1 закрыт для случая «случайный запуск с wrong env».

### Deferred / out of closure scope

- **S2** (rate limit login) — **deferred** (optional WP5, не реализован в P1).
- **S3** (object-level RBAC КП).
- Frontend reload при **409** `plan_version_conflict`.
- Полное удаление legacy `cfg.PLATES_*` proxy.

---

## 9. Risks & Mitigations

| Риск | Митигация |
|------|-----------|
| Strangler A3 оставляет скрытые call sites | WP1 inventory обязателен; grep-чеклист в PR |
| Удаление `_visualize_lock` выявит скрытую гонку | Тесты параллельных запросов до удаления lock; инкрементальный WP4 |
| S1 ломает локальную разработку с ботом | Явный `APP_ENV=development` + документация в `bot/README.md` |
| Рефакторинг optimization задевает PuLP/CBC поведение | Сохранить golden/fixture тесты optimization; не менять алгоритм |
| «Частичный S1» создаёт ложное чувство безопасности | Acceptance явно требует staging/unknown env fail-closed + `sys.exit` |
| Большой blast radius `config_and_data` | Hot paths first; cold paths (scripts, bot) — document only |

---

## 10. Decisions (предварительные — уточнить на ревью)

1. **Порядок спринта:** `S1` (WP0) стартует немедленно; `A3` inventory (WP1) параллельно; WP2→WP3→WP4 последовательно.
2. **Strangler:** legacy `import core.config_and_data as cfg` не удаляем в P1; новые изменения в hot paths — только через context.
3. **Synthetic admin:** допустим **только** `APP_ENV=development` + `BOT_AUTH_ENABLED=false`; все прочие комбинации — deny.
4. **Startup:** fatal misconfiguration → `sys.exit(1)` в `bot/bot_main.py`, не silent `return`.
5. **Инвентаризация:** артефакт в `ai_docs/develop/architecture/` (включено в WP1).
6. **S2:** optional WP5 — не блокирует закрытие P1.
7. **Бот:** deprecated; S1 — defense in depth, не возврат к активной поддержке.

---

## Closure summary (2026-06-19)

| WP | Результат |
|----|-----------|
| WP0 | S1 fail-closed bot auth, startup `sys.exit(1)` |
| WP1 | `plate-runtime-globals-inventory.md` |
| WP2 | FastAPI hot paths → `PlateOrderContext` via `Depends` |
| WP3 | `optimize()` под `ctx.bound()` |
| WP4 | Параллельные HTTP isolation-тесты; `_visualize_lock` удалён |
| WP5 | **Deferred** — S2 rate limit на login |

**Verify:** `pytest tests/ -q` → 744 passed, 12 skipped.

## Следующий шаг (после P1)

1. Спринт P2: **S2**, **S3**, **S4**, frontend 409 reload, расширение test coverage (Q5).
2. Backlog: полное удаление PEP 562 proxy `cfg.PLATES_*`; bot god-modules (A6).
3. Post-P1 секция в audit: [`../develop/audits/2026-06-19-full-project-audit.md`](../develop/audits/2026-06-19-full-project-audit.md).
