# kp_db remediation — Stage 6: A2 completion orchestration

**Date:** 2026-06-04  
**Audit:** [2026-06-03-core-kp-db-audit.md](../audits/2026-06-03-core-kp-db-audit.md) — item **A2** (completion slice)

## Summary

Orchestration for `move_plates_to_completed` moved from [`core/kp_db_plates.py`](../../core/kp_db_plates.py) into [`app/services/plate_completion_service.py`](../../app/services/plate_completion_service.py). Persistence layer keeps SQL primitives and a backward-compatible facade.

## Architecture

| Module | Role |
|--------|------|
| `core/domain/plate_completion_types.py` | `UnmovedPlateInfo`, `CompletePlatesResult` |
| `core/domain/plate_completion_matching.py` | `find_kp_plate_row` (steps 0–7) |
| `core/kp_db_plates.py` | `_fetch_kp_plate_row_by_id`, `_record_plate_completion`, `_purge_zero_qty_plates`; `move_plates_to_completed` facade |
| `app/services/plate_completion_service.py` | `complete_plates_on_cursor`, `move_plates_to_completed` (conn lifecycle) |
| `app/services/production_completion_service.py` | Calls `PlateCompletionService` directly |

## Callers

- **Web:** `ProductionCompletionService.complete_day` → `PlateCompletionService.move_plates_to_completed(..., _external_conn=conn)`
- **Bot:** `bot/handlers/production_completion.py` → `kp_persistence.move_plates_to_completed` (re-export → facade → service)
- **Tests:** `kp_db.move_plates_to_completed` unchanged for golden regression tests

## Layer note

`kp_db_plates.move_plates_to_completed` uses a **lazy import** of `PlateCompletionService` (single transitional `core → app` shim). New code should import the service from `app.services`.

## Verification

```bash
pytest tests/test_kp_db_move_plates_to_completed.py \
  tests/test_plate_completion_service.py \
  tests/test_production_completion_service.py \
  tests/test_plate_audit.py \
  tests/test_plan_consistency.py -q
```

**Result:** 40 passed (3 new service tests + existing regression suite).

## Remaining A2 backlog

- `find_matching_rests` → `RestMatchingService`
- Debug NDJSON cleanup in completion hot path (A7 / S1)
