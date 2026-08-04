# Plan: Порядок армирования в layout (asc/desc)

**Created:** 2026-06-01
**Orchestration:** orch-2026-06-01-12-00-layout-reinf-order
**Status:** 🟢 Ready
**Goal:** Добавить настройку `LAYOUT_REINF_ORDER=asc|desc` и инвертировать сортировки/сравнения армирования в пайплайне layout при `desc` (max→min), сохранив текущее поведение по умолчанию (`asc`).

**Total Tasks:** 8
**Priority:** High
**Estimated Time:** ~4–5 hours

---

## Контекст

Текущий пайплайн: `build_layout_sequence` → `_build_sequence_from_plan` (`viz_modules/layout_sequence/from_plan.py`).

Сейчас армирование везде упорядочивается **по возрастанию** (min→max):

| Место | Текущая логика |
|-------|----------------|
| `solid_cuts.sort` | `(reinforcement ↑, length ↓)` |
| `cut_with_rest` | `(reinforcement ↑, width ↑, rest ↑)` |
| Greedy (`LAYOUT_GREEDY_REINF_MERGE`) | `rs < rg` → целая выигрывает (меньшее армирование) |
| `choose_best_separator` | tier с **min** reinforcement |

Существующие флаги: `LAYOUT_GREEDY_REINF_MERGE`, `LAYOUT_TRACK_REINF_PREFERENCE`, `LAYOUT_TRACK_START_REINF_RELAXATION`.

`LayoutSequenceCfgSlice` (`core/optimization/layout_runtime_snapshot.py`) прокидывает layout-флаги в `_build_sequence_from_plan`.

Тесты: `tests/test_layout_greedy_reinf_merge.py`.

---

## Architecture Decisions

1. **Новая настройка:** `LAYOUT_REINF_ORDER` со значениями `asc` | `desc`, default `asc` — обратная совместимость.
2. **Тип в коде:** `Literal["asc", "desc"]` или enum; валидация через Pydantic `field_validator`.
3. **Централизация:** helper `reinforcement_sort_key(reinf, secondary_keys, order)` и `prefer_lower_reinforcement(order)` в `helpers.py` — избежать дублирования инверсии в нескольких местах.
4. **`choose_best_separator`:** принимает `reinf_order: Literal["asc","desc"]`; при `desc` tier = **max** reinforcement (скачок к next_group — по-прежнему min distance).
5. **`split_sequence_into_tracks`:** **не менять** — выбор стартовой целой основан на близости к соседнему split (`_pick_track_starter_solid_index`), не на глобальном min/max порядке sequence. Задокументировать в plan/report.
6. **`builder.py` legacy-ветка** (строки ~273+, unreachable после early return ~267): не блокирует фичу; опционально синхронизировать или пометить dead code — см. LAYOUT-006.

---

## Tasks Overview

### LAYOUT-001: Настройка `LAYOUT_REINF_ORDER` в Settings
- **Priority:** Critical
- **Complexity:** Simple
- **Dependencies:** None
- **Files:** `core/config/settings.py`, `.env.example` (если есть секция layout)
- **Acceptance criteria:**
  - Поле `layout_reinf_order: Literal["asc", "desc"]` с alias `LAYOUT_REINF_ORDER`, default `"asc"`
  - Validator отклоняет прочие значения
  - `get_settings().layout_reinf_order == "asc"` без env

### LAYOUT-002: Прокинуть флаг в runtime snapshot
- **Priority:** Critical
- **Complexity:** Simple
- **Dependencies:** LAYOUT-001
- **Files:** `core/optimization/layout_runtime_snapshot.py`
- **Acceptance criteria:**
  - `LayoutSequenceCfgSlice.layout_reinf_order: Literal["asc", "desc"]`
  - `from_config_module(..., layout_reinf_order=...)` и override в тестах
  - `build_layout_runtime_snapshot` читает из `Settings`, передаёт в slice

### LAYOUT-003: Helpers для sort/compare по порядку
- **Priority:** High
- **Complexity:** Moderate
- **Dependencies:** LAYOUT-002
- **Files:** `viz_modules/layout_sequence/helpers.py`
- **Acceptance criteria:**
  - `reinf_sort_tuple(reinf: float, order: str) -> tuple` — для asc `(reinf, ...)`, для desc `(-reinf, ...)`
  - `greedy_prefers_solid(rs, rg, order) -> bool | None` — None при равенстве (tie → caller)
  - Unit-тесты helpers (можно в том же test-файле)

### LAYOUT-004: Сортировки и greedy в `from_plan.py`
- **Priority:** Critical
- **Complexity:** Moderate
- **Dependencies:** LAYOUT-003
- **Files:** `viz_modules/layout_sequence/from_plan.py`
- **Acceptance criteria:**
  - `solid_cuts.sort` использует order из `layout_cfg.layout_reinf_order`
  - `cut_with_rest` sorted с инвертированным reinforcement-ключом при desc
  - Greedy loop: сравнение через helper; log «мин.» → «макс.» при desc
  - `_append_inter_group_separator` передаёт order в `choose_best_separator`
  - При `asc` поведение byte-identical с текущим (реgression)

### LAYOUT-005: `choose_best_separator` для desc
- **Priority:** High
- **Complexity:** Simple
- **Dependencies:** LAYOUT-003
- **Files:** `viz_modules/layout_sequence/helpers.py`
- **Acceptance criteria:**
  - Параметр `reinf_order: Literal["asc","desc"] = "asc"`
  - asc: tier = min reinforcement (как сейчас)
  - desc: tier = max reinforcement
  - Log message отражает режим

### LAYOUT-006: Аудит `builder.py` (legacy duplicate path)
- **Priority:** Low
- **Complexity:** Simple
- **Dependencies:** LAYOUT-004
- **Files:** `viz_modules/layout_sequence/builder.py`
- **Acceptance criteria:**
  - Подтверждено: основной путь идёт через `_build_sequence_from_plan` (строки 152, 267)
  - Unreachable block ~273+ либо удалён/помечен, либо задокументирован как dead code без изменений
  - Если block всё ещё reachable в каких-то сценариях — применить тот же order helper

### LAYOUT-007: Тесты desc + regression asc
- **Priority:** High
- **Complexity:** Moderate
- **Dependencies:** LAYOUT-004, LAYOUT-005
- **Files:** `tests/test_layout_greedy_reinf_merge.py` (расширить) или `tests/test_layout_reinf_order.py`
- **Acceptance criteria:**
  - `desc`: solid_cuts и cut_groups идут max→min reinforcement
  - `desc` + greedy: тяжёлая целая вставляется перед лёгкими split-группами (инверсия `test_greedy_inserts_light_solid_before_splits`)
  - `choose_best_separator` desc выбирает max reinforcement
  - asc-тесты из существующего файла проходят без изменений ожиданий
  - `split_sequence_into_tracks` smoke test с desc sequence (item count integrity)

### LAYOUT-008: Верификация и документирование track splitter
- **Priority:** Medium
- **Complexity:** Simple
- **Dependencies:** LAYOUT-007
- **Files:** plan/report comment, опционально docstring в `core/visualization.py`
- **Acceptance criteria:**
  - Явная запись: `split_sequence_into_tracks` не зависит от `LAYOUT_REINF_ORDER` (proximity/cap logic)
  - `pytest tests/test_layout_greedy_reinf_merge.py` (+ новые) green
  - Релевантные layout-тесты проекта green

---

## Dependencies Graph

```
LAYOUT-001 → LAYOUT-002 → LAYOUT-003 → LAYOUT-004 → LAYOUT-006
                              ↓              ↓
                         LAYOUT-005 ──────────┘
                              ↓
                         LAYOUT-007 → LAYOUT-008
```

## Parallelization

- После LAYOUT-003: **LAYOUT-004** и **LAYOUT-005** можно выполнять параллельно (worker + worker).
- LAYOUT-006 — после LAYOUT-004.
- LAYOUT-007 — после LAYOUT-004 и LAYOUT-005.

## Progress

- ⏳ LAYOUT-001: Settings `LAYOUT_REINF_ORDER` (Pending)
- ⏳ LAYOUT-002: LayoutSequenceCfgSlice wiring (Pending)
- ⏳ LAYOUT-003: Reinforcement order helpers (Pending)
- ⏳ LAYOUT-004: from_plan sorts + greedy (Pending)
- ⏳ LAYOUT-005: choose_best_separator desc (Pending)
- ⏳ LAYOUT-006: builder.py audit (Pending)
- ⏳ LAYOUT-007: Tests desc + asc regression (Pending)
- ⏳ LAYOUT-008: Verification + track splitter doc (Pending)

## Implementation Notes

### Пример sort key

```python
def reinf_primary_key(reinf: float, order: str) -> float:
    r = float(reinf)
    return r if order == "asc" else -r
```

`solid_cuts.sort(key=lambda x: (reinf_primary_key(x.get("reinforcement", 999.0), order), -length))`

### Greedy inversion

```python
# asc: rs < rg → solid wins
# desc: rs > rg → solid wins  (equivalent: invert comparison)
```

### Test fixture extension

```python
def _layout_cfg(*, greedy: bool, reinf_order: str = "asc", ...):
    return LayoutSequenceCfgSlice.from_config_module(
        cfg,
        layout_greedy_reinf_merge=greedy,
        layout_reinf_order=reinf_order,
    )
```

## Risks

| Risk | Mitigation |
|------|------------|
| Default change breaks production layout | default `asc`, regression tests |
| Duplicate logic in builder dead branch | audit LAYOUT-006 |
| groupby order unchanged but list order changes | groupby follows sorted input — OK |
| Missing env in deployment | document in .env.example |

## Verification Commands

```bash
pytest tests/test_layout_greedy_reinf_merge.py -v
pytest tests/test_layout_reinf_order.py -v  # if new file
pytest tests/ -k layout -v
```
