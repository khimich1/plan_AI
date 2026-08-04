# Production Pricing Fallback

**Date:** 2026-05-08  
**Status:** Implemented  
**Related Code:**
- `viz_modules/procurement.py::_find_price_for_plate_production_fallback`
- `viz_modules/procurement.py::build_price_rows_production`
- `viz_modules/procurement.py::build_component_breakdown_production`
- `core/price_db.py::length_m_to_price_length_dm`

## Overview

Production smetae (`build_price_rows_production`, `build_component_breakdown_production`) now use a dedicated XLSX fallback function that aligns with database pricing logic: **rounding up to decimeters**.

Previously, fallback on XLSX used `round(length_m * 10)` which differed from the main DB path. This is now unified for the production path only.

## Key Changes

### Previous Behavior (Legacy)
- `find_price_for_plate()` used `round(length_m * 10)` as price table key.
- Example: `5.91 м → 59 дм` (banker's rounding).

### New Behavior (Production)
- `_find_price_for_plate_production_fallback()` uses `length_m_to_price_length_dm(length_m)`.
- Applies **ceiling rounding** aligned with DB `get_price(..., round_up=True)`.
- Example: `5.91 м → 59.1 дм → 60 дм`.

### Rounding Logic
The key transformation:
```python
def length_m_to_price_length_dm(length_m: float) -> int:
    raw_length_dm = length_m * 10
    return int(math.ceil(raw_length_dm - 1e-9))
```
- Subtracts `1e-9` epsilon to avoid floating-point errors.
- Applies `ceil()` to round up.
- Result is an integer decimeters key.

### Fallback Lookup Chain
When querying fallback XLSX table:
1. Try exact match: `price_table[length_dm_key][load_code]`.
2. If not found: search ±1 dm tolerance: `WHERE ABS(tbl_dm - length_dm_key) <= 1`.

## Scope

**Applies to:**
- Production pricing path only (`build_price_rows_production`, `build_component_breakdown_production`).

**Does NOT affect:**
- Commercial offer pricing (`build_price_rows` for regular КП).
- Legacy fallback behavior remains unchanged for backward compatibility.

## Testing

See `tests/test_procurement_production_fallback.py` for helper-level tests covering:
- Fallback lookup with exact and tolerance matches.
- Rounding edge cases (e.g., `5.91 м`, `6.37 м`).

## Related Documentation

- [Pricing Rules Overview](../prise_rules.md) — General pricing flow and DB logic.
- DB rounding: `core/price_db.py::get_price(..., round_up=True)`.
