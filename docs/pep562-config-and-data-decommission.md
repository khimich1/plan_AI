# PEP 562 decommission: `core/config_and_data.py`

Checklist для WP5 / WP6-B5. Канонический доступ к mutable plate state:

```python
from core.plate_runtime_state import get_plate_mutable_runtime
# или request-scoped: PlateOrderContext (app/core hot paths)
```

**Запрещено** в `app/` и `core/` (кроме allowlist):

```python
import core.config_and_data as cfg
cfg.PLATES_*   # proxy
```

Grep-gate: `tests/test_config_and_data_proxy_boundary.py`

---

## Phase 1 — Done

- [x] Grep-gate: нет `config_and_data as cfg` в `app/` / `core/`
- [x] Удалены proxy-имена: `LONGITUDINAL_CUTS`, `WASTE_AREA_M2`, `PLATE_METADATA`, `PLATE_NOMENCLATURE_CACHE`, `LAST_PARSE_DIAGNOSTICS`, `PLATES_1_5_TO_1_2` (`test_decommissioned_proxy_names_raise`)
- [x] FastAPI hot paths → `PlateOrderContext` + middleware (`test_plate_runtime_isolation.py`)
- [x] `viz_modules` только через `app/adapters/visualization.py`

---

## Phase 2 — Remaining (`MUTABLE_LEGACY_NAMES`)

Текущий allowlist в `core/plate_runtime_state.py`:

| Имя | Бывшие consumers | Действие |
|-----|------------------|----------|
| `PLATES_0_*` | `bot_archived/handlers/optimize.py` | Удалить с P6-B2 |
| `PLATES_1_2`, `PLATES_1_0`, `PLATES_1_08` | tests, semantics | Переписать тесты на runtime API |
| `PLATE_LOAD_DETAILS` | tests proxy contract | `get_plate_mutable_runtime().plate_load_details` |
| `PLATE_EXACT_WIDTHS` | tests | runtime |
| `PLATE_LENGTH_DM_RAW` | tests | runtime |

### Steps

1. [ ] `rg "MUTABLE_LEGACY_NAMES|config_and_data\.(PLATES_|PLATE_)" tests/ core/ app/` — список оставшихся call sites
2. [ ] Переписать каждый call site на `get_plate_mutable_runtime()` или explicit `plate_load_details` argument (см. `test_procurement_loads.py` A3 phase 2)
3. [ ] Обновить `scripts/smoke_check.py`: `set_plate_lists_from_text` остаётся публичным API модуля (не proxy)
4. [ ] Удалить `MUTABLE_LEGACY_NAMES` и `__getattr__` в `config_and_data.py` **или** оставить только `get_config()` + constants re-export
5. [ ] Удалить `DeprecationWarning` tests, если proxy удалён
6. [ ] `pytest tests/test_config_and_data_module_semantics.py tests/test_config_and_data_proxy_boundary.py -q`

---

## Phase 3 — Optional cleanup

- [ ] Переименовать модуль `config_and_data.py` → `plate_config.py` (breaking для внешних скриптов — только с ADR)
- [ ] Вынести `set_plate_lists_from_text` в `core/plate_text_pipeline.py`

---

## Verification

```bash
pytest tests/test_config_and_data_proxy_boundary.py tests/test_config_and_data_module_semantics.py tests/test_plate_runtime_isolation.py -q
rg "config_and_data as cfg" app core --glob "*.py"
```

---

## Not doing

- Удаление `get_config()` и constants (`TRACK_LENGTH_M`, etc.) — они не proxy
- Удаление `core/config_and_data.py` целиком в P6 — только proxy layer
