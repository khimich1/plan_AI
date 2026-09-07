# Plan: Trim — соседние марки не воруют резы

**Created:** 2026-09-03  
**Спека:** [../../specs/trim-neighbor-sku-cuts.md](../../specs/trim-neighbor-sku-cuts.md)  
**Идея:** [../../ideas/trim-neighbor-sku-cuts.md](../../ideas/trim-neighbor-sku-cuts.md)  
**Goal:** 56-7 не забирает рез 56-7,2; допуск марки 10 мм (1 см); кромка 20 мм не трогаем.  
**Status:** IMPLEMENT ✅ шаг 1 (pytest 50 passed; шаг 2 не делали)  
**SDD:** SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅ шаг 1

## Overview

`_width_matches_cut` сейчас `abs(diff) <= 20` — это порог **кромки** (`MIN_BILLABLE_TRIM_MM`). Из‑за него `|700−720|=20` склеивает две марки. Документация завода: **1 см на плиту = 10 мм** — допуск **марки**.

Шаг 1: константа `PLATE_WIDTH_MATCH_TOLERANCE_MM = 10` и матч по ней. Шаг 2 (`target_order_key`) **не в этом срезе**.

## Locked decisions

| # | Решение |
|---|---|
| 1 | Матч SKU: **≤ 10 мм**. Не 0 и не 20. |
| 2 | Кромка: `MIN_BILLABLE_TRIM_MM = 20` и `_longitudinal_cuts_for_rest_secondary` **без изменений**. |
| 3 | 710/720 (Δ=10) — в допуске марки, не баг и не триггер шага 2. |
| 4 | Consume-once и snapshot-ключ — только шаг 2, отдельно. |
| 5 | Не коммитить без просьбы. Не убивать `./run+logs.sh`. ILP не трогать. |

## Architecture

Один вертикальный срез: красный тест на плане-клоне бага → константа + `_width_matches_cut` → матрица Δ=20 / Δ=10 / waste=20 → зелёный файл trim-тестов.

```
constants.py (10 мм)
    └── trim.py::_width_matches_cut
            └── test_procurement_trim_cuts.py
```

`price_rows` / `breakdown` не менять: оба читают trim.

## Risks

| Риск | Митигация |
|---|---|
| Старый тест опирался на склейку ±20 мм | Прогон всего `test_procurement_trim_cuts.py`; 725 vs 720 = 5 мм останется в допуске |
| `cuts` = ширина полосы 720 при заказе 700 | T-56-7 красный после фикса → стоп, шаг 2, не возвращать 20 |
| Кейс waste=20 начнёт давать 0 или 2 реза | Отдельный тест на `_longitudinal_cuts_for_rest_secondary`, код счётчика не трогаем |

## Tasks

TDD: failing test **до** правки `_width_matches_cut`.

### TRIM-001 — Красный план 56-7 / 56-7,2

- **Acceptance:** `_pb_56_7_and_72_plan()` как в драфте: prim 500@6,0 rest 700; sec 700 6,0→5,6 waste 0; sec 720 6,3→5,6 waste 160. На текущем коде 5,6×700 даёт `trans_cuts==2` и `long_cut_meterage==6.3` (документ бага). Ассерты **целевые**: 700 → `trans_cuts==1`, `long_cut_meterage==0`, remainder только `(0.4, 1)`; 720 → `trans_cuts==1`, meterage 6.3, remainder `(0.7, 1)`. Тест красный.
- **Verify:** `pytest tests/test_procurement_trim_cuts.py -q -k "56_7"` — fail на 700.
- **Files:** `tests/test_procurement_trim_cuts.py`

### TRIM-002 — Допуск 10 мм

- **Acceptance:** `PLATE_WIDTH_MATCH_TOLERANCE_MM = 10` в `core/config/constants.py`; реэкспорт в `core/config/__init__.py`. `_width_matches_cut` использует её, комментарий: 10 мм = марка, 20 мм = кромка. T-56-7 зелёный.
- **Verify:** `pytest tests/test_procurement_trim_cuts.py -q -k "56_7"`
- **Files:** `core/config/constants.py`, `core/config/__init__.py`, `viz_modules/procurement/trim.py`

### TRIM-003 — Матрица + кромка 20 мм

- **Acceptance:** Пары Δ=20 (700/720, 300/320, 480/500, 880/900) не делят `trans_cuts`/`long_cut_meterage`. 710/720 Δ=10 — матч допустим. `pieces=2`, `waste=20` → ровно 1 продольный (`_longitudinal_cuts_for_rest_secondary`). Каскад 665 без регрессии (уже в файле).
- **Verify:** `pytest tests/test_procurement_trim_cuts.py -q`
- **Files:** `tests/test_procurement_trim_cuts.py`

### TRIM-004 — Полный gate

- **Acceptance:** `pytest tests/test_procurement_trim_cuts.py tests/test_procurement_mixed_load_breakdown.py -q` зелёный. Шаг 2 не начинать.
- **Verify:** команда выше
- **Files:** нет, если 001–003 зелёные

## Out of this slice

Шаг 2, ILP, формула остатка в UI, архив, геометрия `tolerance_width=20`.
