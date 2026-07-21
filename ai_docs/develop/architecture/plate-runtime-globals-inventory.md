# Инвентаризация: глобальное мутабельное plate/OPT runtime (A3, WP1)

> **Дата:** 2026-06-19  
> **Спека:** `ai_docs/specs/stabilizaciya-p1-runtime-security-2026-06-19.md`  
> **План:** WP1 → WP2 → WP3 → WP4  
> **Статус:** draft (инвентаризация, без изменений prod-кода)

## Краткий итог WP1

| Метрика | Значение |
|---------|----------|
| Файлов с `import core.config_and_data` / `cfg` | **~65** (prod + tests + scripts) |
| Строк инвентаря (символы/узлы) | **~42** |
| **Hot** (FastAPI request path) | **12** проблемных узлов |
| **Warm** (бот, deprecated) | **9** handlers |
| **Cold** (tests/scripts/viz) | **~30+** (document only) |
| Уже изолировано middleware | web + bot |
| Lock-костыль | `_visualize_lock` — 2 файла, 3 call site |

## Рекомендуемые точки входа WP2 (по приоритету)

1. `app/services/day_documents_service.py` — orphan `fresh_empty()`, `_visualize_lock`, production day docs
2. `app/services/archive_service.py` — то же + optimize без `bound()`
3. `app/services/commercial_service.py` — `optimize()` вне `ctx.bound()`; procurement только внутри
4. `app/services/commercial_workflow_service.py` — 5× `generate_preview()` без `plate_order_ctx`
5. `app/api/v1/endpoints/production.py` — day document endpoints без `Depends(get_plate_order_context)`
6. `app/api/v1/endpoints/commercial.py` — расширить DI на draft/calculate workflow (сейчас только `/generate-preview`)
7. `app/dependencies/plate_context.py` — уже есть; прокинуть в сервисы

## Рекомендуемые точки входа WP3

1. `app/services/optimization_service.py` — обернуть `optimize()` в `plate_ctx.bound()` / `bound_plate_order_context`
2. `core/production/planning.py` — `optimize_with_cascading_longitudinal_cuts` до `plate_ctx.bound()`
3. `core/visualization.py` + `viz_modules/*` — чтение/запись `cfg.PLATES_*`, `OPT_*`, `PLATE_MAX_REINFORCEMENT_MAP`
4. `core/parsing/plate_lists.py` — `set_plate_lists_from_text` (legacy global writer)
5. `core/optimization/legacy_width_plan.py` — пишет `OPT_PLAN` / `OPT_WIDTH_PRIORITY` из `cfg.PLATES_*`

---

## 1. Назначение

Карта всех call sites, где состояние заказа плит (`PLATES_*`, `PLATE_*`) и оптимизатора (`OPT_*`)
читается или мутируется через process-wide runtime. Цель P1 — strangler: явный
`PlateOrderContext` + `bound()` на hot paths FastAPI; бот и scripts — document only.

## 2. Инфраструктура (SSOT)

| Файл | Символ | Тип | Callers | Приоритет | Предлагаемый фикс |
|------|--------|-----|---------|-----------|-------------------|
| `core/plate_runtime_state.py` | `PlateMutableRuntime` | struct, in-place mutate | both | P0 infra | Оставить; единственный контейнер данных |
| `core/plate_runtime_state.py` | `_plate_cv`, `_tls` | storage | both | P0 infra | Без изменений |
| `core/plate_runtime_state.py` | `get_plate_mutable_runtime()` | read (resolve ctx/TLS) | both | P0 infra | Вызывать только внутри `bound()` на hot paths |
| `core/plate_runtime_state.py` | `bind_plate_mutable_runtime` / `reset_*` | write (ContextVar) | both | P0 infra | Через `plate_mutable_runtime_scope` |
| `core/plate_runtime_state.py` | `plate_mutable_runtime_scope` | bind/unbind | both | P0 infra | Используется в `PlateOrderContext.bound()` |
| `core/plate_runtime_state.py` | `fresh_plate_mutable_request_scope` | empty per request | web/bot | P0 infra | Дублирует middleware |
| `core/plate_runtime_state.py` | `MUTABLE_ATTR_MAP` (23 имени) | legacy alias map | both | backlog | Удаление после strangler |
| `core/plate_runtime_state.py` | `factory_demo_order` / `new_plate_mutable_runtime_from_demo` | seed demo data | cold | low | Не использовать на hot paths |
| `core/config_and_data.py` | `__getattr__` (PEP 562) | read proxy → runtime | both | backlog | Strangler; новый код — только ctx |
| `core/config_and_data.py` | `register_plate_metadata` | write `plate_metadata` | viz | WP3 | Параметр ctx или snapshot |
| `core/config_and_data.py` | `consume_plate_metadata` | read+delete metadata | viz | WP3 | То же |
| `core/config_and_data.py` | `clear_plate_metadata` | write clear | viz | WP3 | То же |
| `core/config_and_data.py` | `get_load_code_for_plate` / `get_exact_width` | read runtime | both | WP3 | Принимать `PlateMutableRuntime` |
| `core/parsing/plate_lists.py` | `set_plate_lists_from_text` | **write** lists+details | bot/cold/tests | WP3 | Парсить в `PlateOrder` / `ctx.plates` |
| `core/parsing/plate_lists.py` | `_clear_all_plate_lists` | write clear | via set_plate_lists | WP3 | — |
| `core/plate_order_context.py` | `PlateOrderContext` | SSOT plates+OPT dict | both | P0 infra | DI на все hot paths |
| `core/plate_order_context.py` | `fresh_empty()` | create isolated state | both | WP2 | **Не** создавать orphan внутри сервисов |
| `core/plate_order_context.py` | `bound()` | bind plates+OPT | both | WP2 | Обязателен на всю мутацию legacy |
| `core/plate_order_context.py` | `hydrate_from_order` | write plates | both | OK | — |
| `core/plate_order_context.py` | `load_optimization_snapshot` | write OPT dict | both | OK | — |
| `core/plate_order_context.py` | `run_in_order_context` | thread worker + bound | web/bot | WP2 | Использовать с request ctx, не fresh_empty |
| `core/optimization/context.py` | `OPT_*` proxies, `_tls`, `_opt_cv` | read/write OPT state | both | P0 infra | Уже ContextVar-aware |
| `core/optimization/context.py` | `optimization_context_scope` | bind OPT | both | OK | Вызывается в orchestrator + `bound()` |
| `core/optimization/orchestrator.py` | `optimize_with_cascading_*` | creates OPT scope | both | WP3 | Добавить plate `bound()` с caller ctx |
| `core/domain/plate_order.py` | `apply_to_globals` (deprecated) | write runtime | cold | backlog | `hydrate_from_order` |
| `core/domain/plate_order.py` | `get_current_plate_order` (deprecated) | read runtime | tests | backlog | `snapshot_core_order` |

## 3. Middleware и DI (уже есть)

| Файл | Символ | Тип | Callers | Приоритет | Предлагаемый фикс |
|------|--------|-----|---------|-----------|-------------------|
| `app/middleware/plate_runtime_isolation.py` | `PlateMutableRuntimeIsolationMiddleware` | per-request `fresh_empty` + `bound()` | web | ✅ done | Проверить порядок в `app/main.py` |
| `app/dependencies/plate_context.py` | `get_plate_order_context` | read `request.state` | web | WP2 | Подключить ко всем mutating endpoints |
| `bot/middleware/plate_runtime_isolation.py` | `PlateMutableRuntimeIsolationMiddleware` | per-update ctx | bot | warm | Document only (бот deprecated) |
| `bot/dependencies/plate_context.py` | `get_plate_order_context` | read from aiogram data | bot | warm | — |

## 4. Hot paths — FastAPI (WP2)

| Файл | Символ / паттерн | Тип | Endpoint / путь | Проблема | WP | Фикс |
|------|------------------|-----|-----------------|----------|-----|------|
| `app/api/v1/endpoints/commercial.py` | `Depends(get_plate_order_context)` | read ctx | `POST /generate-preview` | ✅ изолирован | — | — |
| `app/api/v1/endpoints/commercial.py` | draft/calculate routes | — | `/drafts/*/calculate`, wizard | ❌ нет DI ctx | WP2 | `Depends(get_plate_order_context)` → workflow |
| `app/services/commercial_service.py` | `optimization_service.optimize()` | write OPT (scope в orchestrator) | preview | ❌ вне `bound()` | WP2/WP3 | `with ctx.bound():` вокруг optimize+procurement |
| `app/services/commercial_service.py` | `ctx.bound()` + `build_*` | read cfg | preview | частично ✅ | WP2 | Передать request ctx, не `or fresh_empty()` |
| `app/services/commercial_workflow_service.py` | `generate_preview(text=...)` ×5 | orphan ctx | wizard/calculate | ❌ игнорирует middleware ctx | WP2 | `plate_order_ctx` от caller |
| `app/services/commercial_workflow_service.py` | `PlateOrderContext.fresh_empty()` | orphan | schema file gen | ❌ | WP2 | request ctx |
| `app/services/day_documents_service.py` | `_visualize_lock` | serialize | day docs | костыль S5 | WP4 | Удалить после isolation tests |
| `app/services/day_documents_service.py` | `_build_visualization_ctx` → `fresh_empty` | orphan ctx | `/production/days/*/documents/*` | ❌ не request ctx | WP2 | `Depends` → param |
| `app/services/day_documents_service.py` | `run_in_order_context(viz_ctx, ...)` | bound в thread | day docs | orphan ctx | WP2 | Тот же объект, что middleware |
| `app/services/archive_service.py` | import `_visualize_lock` | serialize | archive schema | костыль | WP4 | — |
| `app/services/archive_service.py` | `viz_ctx = fresh_empty()` | orphan | `download_archive_document` schema | ❌ | WP2 | request ctx |
| `app/services/archive_service.py` | `optimization_service.optimize()` | write OPT | archive schema | ❌ вне plate bound | WP3 | `bound_plate_order_context` |
| `app/services/file_generation_service.py` | `generate_visualization` + `ctx.bound()` | read/write globals | commercial schema | ✅ если ctx от caller | WP2 | Caller не должен давать fresh_empty |
| `app/services/optimization_service.py` | `optimize()` | OPT scope only | commercial/archive | plate runtime не bound | WP3 | `bound_plate_order_context` |
| `app/services/optimization_service.py` | `legacy_runtime()` (deprecated) | swap OPT globals | — | deprecated | WP3 | Удалить после миграции |
| `app/api/v1/endpoints/production.py` | `generate_day_*` | viz path | day documents | ❌ нет ctx DI | WP2 | `Depends(get_plate_order_context)` |
| `app/services/production_planning_service.py` | → `core.production.planning.optimize` | optimize+layout | `POST /plans/build` | partial bound | WP3 | optimize под `plate_ctx.bound()` |
| `core/production/planning.py` | `plate_ctx.fresh_empty()` локально | orphan per call | plan build | OK для sync pipeline | WP3 | Явный ctx param от service |
| `app/services/plate_parser_service.py` | `legacy_cfg.make_plate_name` | read-only fn | parse | ✅ без globals | — | — |

## 5. Warm paths — Telegram-бот (document only)

| Файл | Паттерн | Тип | Фикс |
|------|---------|-----|------|
| `bot/handlers/commercial.py` | `set_plate_lists_from_text` | **write** PLATES_* | `run_in_order_context` + migrate to PlateOrder |
| `bot/handlers/commercial.py` | `cfg.*` в pricing pipeline | read/write | уже `run_in_order_context` |
| `bot/handlers/optimize.py` | read `cfg.PLATES_*` | read | warm |
| `bot/handlers/production_execution.py` | optimize + viz | read/write | `run_in_order_context` (частично) |
| `bot/handlers/production_day_view.py` | `run_in_order_context` | bound | warm |
| `bot/handlers/kp.py` | `run_in_order_context` | bound | warm |
| `bot/handlers/production_completion.py` | `cfg.normalize_load_code` etc. | read-only helpers | cold |
| `bot/handlers/production_track_fill.py` | `cfg.normalize_load_code` | read-only | cold |

Middleware бота (`bot/handlers/__init__.py`) регистрирует `PlateMutableRuntimeIsolationMiddleware` — базовая изоляция есть.

## 6. Cold paths — core / viz / scripts (WP3 backlog, не блокер P1 closure)

| Область | Файлы (выборка) | Паттерн | Фикс |
|---------|-----------------|---------|------|
| Визуализация | `core/visualization.py` | **write** `PLATE_MAX_REINFORCEMENT_MAP`; read `OPT_*`, `cfg.PLATES_*` | snapshot / ctx.bound |
| Viz modules | `viz_modules/procurement/*.py`, `layout_sequence/*.py`, `visualization_drawing.py` | read `cfg.PLATE_LOAD_DETAILS`, `PLATES_*` | передавать snapshot |
| Optimization internals | `core/optimization/ilp_model.py`, `geometry.py`, `legacy_width_plan.py`, `optimize_2d/*` | read cfg; legacy_width **writes** OPT | layout_runtime_snapshot |
| Прочее core | `core/plan_commit.py`, `track_top_up.py`, `rescue_tracks.py`, `kp_db_nomenclature.py` | mostly read helpers | document |
| Preview XLSX | `core/plates_preview_xlsx.py` | `set_plate_lists_from_text` | cold |
| Scripts | `scripts/smoke_check.py`, `restore_price_db.py`, `fill_prices_from_xlsx.py` | global mutate | document |
| Tests | `tests/test_*.py` (~25 файлов) | fixture mutate cfg | оставить; расширить isolation tests в WP4 |

## 7. Locks (симптом globals, WP4)

| Файл | Символ | Причина | Фикс |
|------|--------|---------|------|
| `app/services/day_documents_service.py:44` | `_visualize_lock = asyncio.Lock()` | `visualize_plan` мутирует globals | Удалить после WP2+WP4 tests |
| `app/services/day_documents_service.py:145,172` | `async with _visualize_lock` | сериализация параллельных запросов | — |
| `app/services/archive_service.py:23,239` | reuse `_visualize_lock` | то же | — |

## 8. Сводка рисков S5 (cross-request leakage)

1. **Orphan `PlateOrderContext.fresh_empty()`** в сервисах — данные не утекают в соседний запрос (отдельный ctx), но **middleware ctx игнорируется**, а lock **сериализует** весь viz-путь.
2. **`commercial_service.optimize()` без plate `bound()`** — OPT изолирован (`optimization_context_scope`), plate lists в TLS middleware **могут** читаться optimization/viz кодом с неверным заказом.
3. **`commercial_workflow_service`** — 5 вызовов preview без request ctx; главный риск на wizard/calculate.
4. **Thread pool:** `run_in_order_context` → `asyncio.to_thread` + `ctx.bound()` — корректно **если** ctx тот же объект.

## 9. Чеклист grep для PR (WP2+)

```text
get_plate_mutable_runtime|PlateOrderContext\.fresh_empty|set_plate_lists_from_text
cfg\.PLATES_|cfg\.PLATE_|_visualize_lock
run_in_order_context|\.bound\(\)
import core\.config_and_data
```

## 10. Gate G1 — список файлов для WP2

1. `app/services/day_documents_service.py`
2. `app/services/archive_service.py`
3. `app/services/commercial_service.py`
4. `app/services/commercial_workflow_service.py`
5. `app/api/v1/endpoints/production.py`
6. `app/api/v1/endpoints/commercial.py` (draft workflow)
7. `app/services/file_generation_service.py` (caller contract)

## 11. Gate G1 — список файлов для WP3

1. `app/services/optimization_service.py`
2. `core/production/planning.py`
3. `core/visualization.py`
4. `core/optimization/legacy_width_plan.py`
5. `core/parsing/plate_lists.py`
6. `viz_modules/procurement/breakdown.py` (representative viz_modules)

---

## Ключевые архитектурные находки

**Уже работает:**

- `PlateMutableRuntimeIsolationMiddleware` на всех API routes (`app/main.py:52–56`)
- `PlateOrderContext.bound()` связывает plate runtime + OPT через ContextVar
- `optimize_with_cascading_longitudinal_cuts` оборачивает себя в `optimization_context_scope()`
- `POST /commercial/generate-preview` передаёт `plate_order_ctx` из middleware

**Главные gaps для WP2:**

- Production day documents и archive schema создают **свой** `fresh_empty()`, не `request.state.plate_order_ctx`
- Commercial wizard (`commercial_workflow_service`) **5 раз** вызывает preview без ctx
- `_visualize_lock` маскирует гонки, но **блокирует параллелизм** (S5 symptom)

**Для WP3:**

- `CommercialService.generate_preview`: строка 56 — `optimize()` до `ctx.bound()` (строки 58–69)
- `core/visualization.py` пишет в `cfg.PLATE_MAX_REINFORCEMENT_MAP` во время `visualize_plan`
