# PLAN: Стабилизация P1 — runtime isolation + bot auth (аудит 2026-06-19)

> **Фаза SDD:** PLAN → IMPLEMENT — **закрыт 2026-06-19**
> **Дата:** 2026-06-19
> **Orchestration:** `orch-2026-06-19-stabilizaciya-p1`
> **Спека:** [`../../specs/stabilizaciya-p1-runtime-security-2026-06-19.md`](../../specs/stabilizaciya-p1-runtime-security-2026-06-19.md)
> **Baseline:** [`../../specs/project-baseline.md`](../../specs/project-baseline.md)
> **Предшественник:** [`2026-06-19-stabilizaciya-p0.md`](./2026-06-19-stabilizaciya-p0.md) — **закрыт**
> **Источник:** [`../audits/2026-06-19-full-project-audit.md`](../audits/2026-06-19-full-project-audit.md) → Post-P0

---

## 0. Резюме плана

Два кластера, **WP0 можно стартовать сразу**; A3 — инкрементальный strangler:

- **Кластер Security:** `S1` — fail-closed bot auth, startup guard, ужесточение synthetic admin (только explicit development).
- **Кластер Runtime:** `A3` + `S5` — инвентаризация globals → request-scoped context на FastAPI hot paths → optimization/config path → тесты изоляции + снятие `_visualize_lock`.

**Optional:** `S2` rate limit login (WP5) — только при наличии времени, не блокирует closure P1.

**Health Score цель:** ~3/10 → **~5–6/10** (закрытие 2 remaining critical: `A3`, `S1`).

### Граф зависимостей

```
WP0 (S1: bot auth fail-closed)                    [старт сразу, независим]
        │
WP1 (A3: inventory globals)                      [параллельно с WP0]
        │
        ▼
WP2 (A3: request-scoped context, FastAPI hot paths)
        │
        ▼
WP3 (A3: optimization/config hot path off globals)
        │
        ▼
WP4 (A3: isolation tests + remove _visualize_lock)

WP5 (S2: rate limit login) — OPTIONAL, после WP0 или параллельно, не блокирует P1 closure
```

**Параллельно:** `WP0` и `WP1`.
**Строго последовательно:** `WP1 → WP2 → WP3 → WP4`.

---

## Прогресс (closure 2026-06-19)

| WP | Статус | Находка |
|----|--------|---------|
| WP0 | ✅ done | S1 — bot auth fail-closed |
| WP1 | ✅ done | A3 — inventory global mutations |
| WP2 | ✅ done | A3 — FastAPI hot paths + DI |
| WP3 | ✅ done | A3 — optimization/config hot path |
| WP4 | ✅ done | A3/S5 — isolation tests + locks |
| WP5 | ⏸ deferred | S2 — rate limit login (optional, вне closure P1) |

### Gates

| Gate | Статус |
|------|--------|
| G0 (WP0) | ✅ closed |
| G1 (WP1) | ✅ closed |
| G2 (WP2) | ✅ closed |
| G3 (WP3) | ✅ closed |
| **G4 (P1 closure)** | ✅ **closed** |
| G5 (WP5 optional) | ⏸ deferred |

### Completion summary

- **WP0:** synthetic admin только в `APP_ENV=development`; staging/unknown env → deny; `validate_bot_startup()` → `sys.exit(1)` при fatal misconfig; расширен `tests/test_bot_auth.py`.
- **WP1:** артефакт [`../architecture/plate-runtime-globals-inventory.md`](../architecture/plate-runtime-globals-inventory.md).
- **WP2:** production/commercial/archive hot paths получают `PlateOrderContext` через `Depends(get_plate_order_context)`; visualization в `ctx.bound()`.
- **WP3:** optimization/commercial preview и day documents под `ctx.bound()`; `tests/test_core_no_app_import.py` зелёный.
- **WP4:** `tests/test_plate_runtime_request_isolation.py` (параллельные HTTP); `_visualize_lock` удалён из `day_documents_service` / `archive_service`.

**Регрессия:** `pytest tests/ -q` → **744 passed, 12 skipped** (было ~726 на старте P1 после WP0).

**Health Score:** ~3/10 → **~6/10** (все 4 исходных critical закрыты: A1/A2 в P0, A3/S1 в P1).

---

## WP0 — S1: bot auth fail-closed

**Зачем:** закрыть critical `S1`. Частичный фикс уже есть (Pydantic + middleware deny в production), но synthetic admin в non-dev env и мягкий startup остаются.

**Текущее состояние:**
- `core/config/settings.py` — `validate_bot_telegram_auth` блокирует `production` + `BOT_AUTH_ENABLED=false`.
- `bot/middleware/auth.py` — в production при disabled auth handler не вызывается; в остальных env — synthetic admin.
- `bot/bot_main.py` — `validate_bot_startup()` возвращает `False`, но не `sys.exit(1)`.

**Работы:**
1. **`bot/middleware/auth.py`:**
   - Synthetic admin только при `app_env == "development"` **и** `not bot_auth_enabled`.
   - Иначе (`staging`, `test`, unknown, empty) → `log_bot_security_event("misconfiguration", ...)` + deny (не вызывать handler).
2. **`bot/bot_main.py`:**
   - `validate_bot_startup()` при fatal config → `sys.exit(1)` (или вызывающий `main` после `False`).
   - Сохранить deprecation warning + dev open-access warning.
3. **`core/config/settings.py`** (если нужно):
   - Уточнить `bot_auth_fail_closed_enabled` для non-production unknown env (опционально: treat unknown as production-like).
4. **`bot/README.md`:** документировать dev-only open access.
5. **Тесты:** расширить `tests/test_bot_auth.py`:
   - `test_auth_middleware_denies_disabled_auth_in_staging`
   - `test_auth_middleware_allows_synthetic_admin_only_in_development`
   - `test_bot_startup_exits_on_production_misconfig` (monkeypatch `sys.exit`)

**Files (~4–5):**
- `bot/middleware/auth.py`
- `bot/bot_main.py`
- `bot/README.md`
- `tests/test_bot_auth.py`
- (опционально) `core/config/settings.py`

**Verify:**
```powershell
pytest tests/test_bot_auth.py -q
```

**Зависимости:** нет. **Рекомендуется первым WP для implement.**

**Gate G0:** все тесты `test_bot_auth` зелёные; synthetic admin только в development; production misconfig → exit.

---

## WP1 — A3: инвентаризация global mutations

**Зачем:** полная карта call sites до рефакторинга; снижает риск пропущенных утечек (S5).

**Работы:**
1. Статический обход (ripgrep) по паттернам:
   - `import core.config_and_data`, `from core.config_and_data`
   - `get_plate_mutable_runtime`, `bind_plate_mutable_runtime`, `plate_mutable_runtime_scope`
   - `cfg.PLATES_`, `cfg.PLATE_`, `set_plate_lists_from_text`
   - `PlateOrderContext`, `run_in_order_context`, `request.state.plate_order_ctx`
2. Классификация каждого site: **hot** (FastAPI request path) / **warm** (bot, deprecated) / **cold** (scripts, tests).
3. Отметить уже изолированные: `app/middleware/plate_runtime_isolation.py`, `app/dependencies/plate_context.py`, частично `commercial.py` endpoint.
4. Отметить locks: `_visualize_lock` в `day_documents_service`, reuse в `archive_service`.
5. Сохранить артефакт: `ai_docs/develop/architecture/plate-runtime-globals-inventory.md` (таблица: файл, символ, тип мутации, hot/warm/cold, WP-миграция).

**Files (~1 doc + чеклист):**
- `ai_docs/develop/architecture/plate-runtime-globals-inventory.md` (new)
- (опционально) `scripts/audit_plate_runtime_usage.py` — grep-helper для CI

**Verify:**
- Документ покрывает ≥90% вхождений из grep (ручная сверка).
- Hot paths из audit явно перечислены: `day_documents_service`, `archive_service`, `commercial_workflow_service`, `commercial_service`, `core/optimization/*`.

**Зависимости:** нет (параллельно с WP0).

**Gate G1:** inventory approved; список файлов для WP2/WP3 зафиксирован.

---

## WP2 — A3: request-scoped context для FastAPI hot paths

**Зачем:** гарантировать, что production/commercial HTTP handlers используют `request.state.plate_order_ctx`, а не orphan globals.

**Работы:**
1. Проверить регистрацию `PlateMutableRuntimeIsolationMiddleware` в `app/main.py` (должна быть на всех API routes).
2. Расширить `app/dependencies/plate_context.py` — единая точка `get_plate_order_context`.
3. Подключить `Depends(get_plate_order_context)` к endpoints без контекста (production day view, documents, archive visualization — по inventory WP1).
4. **`day_documents_service` / `archive_service`:**
   - Принимать `PlateOrderContext` параметром от caller.
   - Внутри visualization: `with ctx.bound():` вместо ad-hoc `PlateOrderContext.fresh_empty()` без bind.
5. Убедиться, что middleware-created ctx — тот же объект, что уходит в сервисы.

**Files (~5–8, по inventory):**
- `app/main.py` (middleware order, если нужно)
- `app/dependencies/plate_context.py`
- `app/api/v1/endpoints/production.py` (и/или `archive.py`, `commercial.py`)
- `app/services/day_documents_service.py`
- `app/services/archive_service.py`
- `app/services/commercial_workflow_service.py` (частично уже использует ctx)
- тесты smoke для затронутых endpoints

**Verify:**
```powershell
pytest tests/test_archive_endpoints.py -q
pytest tests/test_production_planning_service.py -q
# + существующие production API tests
pytest tests/ -q --ignore=tests/test_plate_runtime_request_isolation.py
```

**Зависимости:** `WP1` (список hot paths).

**Gate G2:** все hot endpoints из inventory используют request ctx; регрессия API зелёная.

---

## WP3 — A3: миграция optimization/config hot path off globals

**Зачем:** критические пути оптимизации и парсинга не должны читать/писать demo default runtime вне `bound()`.

**Работы:**
1. Обернуть вызовы optimization в `commercial_service` / `commercial_workflow_service` в `ctx.bound()` (если ещё не везде).
2. **`day_documents_service`:** путь `_build_visualization_ctx` + optimize — только внутри переданного ctx; убрать неявную зависимость от global demo order (`factory_demo_order`).
3. Рассмотреть `run_in_order_context` для sync-участков, вызываемых из async handlers.
4. **Не трогать** (document only): `bot/handlers/*`, `scripts/`, bulk `viz_modules/` — cold/warm paths в inventory.
5. При необходимости — тонкий порт в `core/optimization/context.py` для явной передачи state без module globals (strangler, без big-bang).

**Files (~4–6):**
- `app/services/day_documents_service.py`
- `app/services/archive_service.py`
- `app/services/commercial_workflow_service.py`
- `app/services/commercial_service.py`
- `core/plate_order_context.py` (утилиты, если нужны)
- `core/optimization/context.py` (минимальные правки)

**Verify:**
```powershell
pytest tests/test_commercial_web_flow.py -q
pytest tests/test_plate_mutable_runtime_isolation.py -q
pytest tests/test_core_no_app_import.py -q
```

**Зависимости:** `WP2`.

**Gate G3:** optimization/commercial hot path не оставляет следов в global runtime после завершения запроса (проверка тестом WP4).

---

## WP4 — A3/S5: isolation tests + remove locks

**Зачем:** доказать отсутствие cross-request leakage (S5); убрать `_visualize_lock` как костыль сериализации.

**Работы:**
1. **Новый** `tests/test_plate_runtime_request_isolation.py`:
   - Два параллельных `TestClient` запроса (или `asyncio.gather` + ASGI) с разным plate text / plan context.
   - Assert: результаты не содержат данных друг друга (diagnostics, counts, preview fields).
2. Расширить `tests/test_plate_mutable_runtime_isolation.py` — edge cases (nested `bound()`, middleware + Depends).
3. **Удалить `_visualize_lock`** из `day_documents_service` и import в `archive_service` после зелёных isolation tests.
4. Если удаление lock нестабильно на CI — feature flag `PLATE_RUNTIME_SERIALIZE_VIS=1` (Ask first); предпочтение — удалить.

**Files (~3–5):**
- `tests/test_plate_runtime_request_isolation.py` (new)
- `tests/test_plate_mutable_runtime_isolation.py` (extend)
- `app/services/day_documents_service.py` (remove lock)
- `app/services/archive_service.py` (remove lock import/usage)

**Verify:**
```powershell
pytest tests/test_plate_runtime_request_isolation.py -q
pytest tests/test_plate_mutable_runtime_isolation.py -q
pytest tests/ -q
```

**Зависимости:** `WP3`.

**Gate G4 (P1 closure):** isolation tests зелёные; `_visualize_lock` удалён; full suite green.

---

## WP5 — OPTIONAL: S2 rate limiting на login

**Зачем:** high-priority `S2` из audit; не блокирует закрытие P1.

**Работы:**
1. Добавить rate limiter (например `slowapi` или lightweight in-process counter) на `POST /api/v1/auth/login`.
2. Лимит: **5 попыток/мин на IP** → **429** + `Retry-After`.
3. Тест: 6-й запрос → 429.

**Files (~2–4):**
- `app/api/v1/endpoints/auth.py`
- `app/main.py` (limiter init)
- `tests/test_auth_rate_limit.py` (new)
- `requirements.txt` / `pyproject.toml` (если `slowapi`) — **Ask first**

**Verify:**
```powershell
pytest tests/test_auth_rate_limit.py -q
```

**Зависимости:** желательно после `WP0` (security cluster); не зависит от A3.

---

## Риски и митигации (на уровне плана)

| Риск | Где | Митигация |
|------|-----|-----------|
| Пропущенный call site globals | WP2/WP3 | WP1 inventory + PR checklist |
| Flaky parallel HTTP test | WP4 | Достаточные `sleep`/barrier; детерминированные fixtures |
| S1 ломает dev workflow | WP0 | Explicit `APP_ENV=development` documented |
| Lock removal exposes race | WP4 | Тесты до удаления; revert flag Ask first |
| Scope creep в viz_modules | WP3 | Cold path — только inventory, не рефакторинг |
| Optional WP5 тянет зависимость | WP5 | In-process limiter без Redis для single instance |

---

## Контрольные точки верификации (gates)

| Gate | Условие перехода |
|------|------------------|
| **G0** (после WP0) | `test_bot_auth` зелёный; fail-closed вне development |
| **G1** (после WP1) | Inventory утверждён; hot file list для WP2 зафиксирован |
| **G2** (после WP2) | Hot API paths на `PlateOrderContext`; endpoint tests зелёные |
| **G3** (после WP3) | Optimization path under `bound()`; commercial/day tests зелёные |
| **G4 — P1 closure** (после WP4) | Parallel isolation tests зелёные; locks сняты; `pytest tests/ -q` green |
| **G5 — optional** (после WP5) | Login rate limit 429 test green |

---

## Оценка трудозатрат (ориентир)

| WP | Complexity | Оценка |
|----|------------|--------|
| WP0 | Simple | 0.5–1 день |
| WP1 | Simple | 0.5 день |
| WP2 | Moderate | 1–2 дня |
| WP3 | Moderate–Complex | 2–3 дня |
| WP4 | Moderate | 1–2 дня |
| WP5 (opt) | Simple | 0.5–1 день |

**Итого P1 (WP0–WP4):** ~5–8 рабочих дней.

---

## Следующий шаг (после P1)

1. **P2 security/quality:** S2 (rate limit login), S3 (object-level RBAC КП), S4 (npm CVE).
2. Frontend: reload плана при **409** `plan_version_conflict`.
3. Расширить integration-тесты production API (Q5).
4. Backlog strangler: полное удаление `cfg.PLATES_*` proxy; cold paths bot/scripts — по inventory.

**Deferred из P1 (не блокировало closure):**
- WP5 / **S2** — rate limiting на login.
- **S3**, bot decomposition (A6), frontend 409 UX.

---

## Deferred (явно вне этого плана)

- **S3** — object-level RBAC КП
- Frontend reload при **409** `plan_version_conflict`
- Bot god-module decomposition (`A6`)
- Полное удаление PEP 562 proxy в `config_and_data.py`
- **S4** npm CVE, **A4** bot→app inversion
