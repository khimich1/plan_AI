# ADR: граница core ↔ viz_modules (WP3 A1 + P4 WP1)

> **Статус:** accepted (slice 1, 2026-06-21; slice 2, 2026-06-21)  
> **Связано:** [`stabilizaciya-p3-architecture-2026-06-21.md`](../../specs/stabilizaciya-p3-architecture-2026-06-21.md) WP3 · [`stabilizaciya-p4-architecture-2026-06-21.md`](../../specs/stabilizaciya-p4-architecture-2026-06-21.md) WP1

## Контекст

`core/` — доменная и алгоритмическая логика (оптимизация, планирование, commit).  
`viz_modules/` — визуализация, procurement UI, layout sequence builder, drawing.

До P3 WP3 `core/visualization.py` и `core/production/planning.py` напрямую импортировали `viz_modules/*`, создавая циклическую связность (`viz_modules` уже импортирует `core`).

## Решение

**Ports & adapters (hexagonal):**

| Слой | Роль |
|------|------|
| `core/ports/visualization.py` | Protocol + registry + facades |
| `viz_modules/adapters/visualization_ports.py` | Default adapter — регистрирует реализации из viz |
| `app/adapters/visualization.py` | App startup wiring (`wire_visualization_ports()`) |

**Направление зависимостей:**

```
core/ports  ←  viz_modules/adapters  ←  viz_modules/*
     ↑
app/main.py (lifespan) + tests/conftest.py (pytest_configure)
```

`core/` **не импортирует** `viz_modules/` на runtime hot paths (slice 1 + slice 2 complete).

## Порты (slice 1)

| Port | Facade | Реализация |
|------|--------|------------|
| `BuildLayoutSequenceFn` | `core.ports.visualization.build_layout_sequence` | `viz_modules.layout_sequence.build_layout_sequence` |
| `LoadPriceTableFn` | `core.ports.visualization.load_price_table_from_xlsx` | `viz_modules.price_utils.load_price_table_from_xlsx` |

## Порты (slice 2 — P4 WP1)

| Port | Facade | Реализация |
|------|--------|------------|
| `BuildPriceRowsFn` | `build_price_rows` | `viz_modules.procurement.build_price_rows` |
| `BuildComponentBreakdownFn` | `build_component_breakdown` | `viz_modules.procurement.build_component_breakdown` |
| `GetOrdersFromOptPlanFn` | `get_orders_from_opt_plan` | `viz_modules.procurement.get_orders_from_opt_plan` |
| `BuildPriceRowsProductionFn` | `build_price_rows_production` | `viz_modules.procurement.build_price_rows_production` |
| `BuildComponentBreakdownProductionFn` | `build_component_breakdown_production` | `viz_modules.procurement.build_component_breakdown_production` |
| `DrawSegmentFn` | `draw_segment` | `viz_modules.visualization_drawing._draw_segment` |
| `DrawSplitPlateFn` | `draw_split_plate` | `viz_modules.visualization_drawing._draw_split_plate` |
| `DrawTransverseCutFn` | `draw_transverse_cut` | `viz_modules.visualization_drawing._draw_transverse_cut` |

## Мигрированные import paths

**Slice 1 (P3 WP3):**

1. `core/production/planning.py` — `build_layout_sequence` → port facade
2. `core/visualization.py` — `build_layout_sequence`, `load_price_table_from_xlsx` → port facade

**Slice 2 (P4 WP1):**

1. `core/visualization.py` — procurement breakdown + production breakdown → port facades (`build_price_rows`, `build_component_breakdown`, `get_orders_from_opt_plan`, production variants)
2. `core/visualization.py` — drawing helpers → port facades (`draw_segment`, `draw_split_plate`, `draw_transverse_cut`); removed top-level and lazy `from viz_modules.procurement import ...` / `from viz_modules.visualization_drawing import ...`

## Grep gate

```bash
rg "viz_modules" core/ --glob "*.py"
```

**Target (slice 2):** **0** runtime `import viz_modules` / `from viz_modules` in `core/*.py`.

**Allowed non-import mentions:** docstrings and comments only (e.g. `core/ports/visualization.py` module docstring, `core/plate_attribution.py` cross-reference comment).

**Hot-path modules verified:** `core/production/planning.py`, `core/visualization.py` — **0** `viz_modules` import statements.

CI/test gate: `tests/test_core_viz_import_boundary.py` (planning + visualization source scan; migrated port import assertions; pytest wiring smoke).

## Регистрация портов

- FastAPI: `app/main.py` → `lifespan` → `wire_visualization_ports()`
- pytest: `tests/conftest.py` → `pytest_configure`
- Standalone scripts: вызвать `wire_visualization_ports()` до использования port facades

## Последствия

- Тесты могут monkeypatch `core.production.planning.build_layout_sequence` или `core.ports.visualization.*` facades (через register).
- Новые вызовы viz из core — только через новые Protocol в `core/ports/`.
- Полный refactor `core/visualization.py` (~1242 LOC) — отдельные slices (A7); boundary contract A1 closed.
