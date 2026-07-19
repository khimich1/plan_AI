# kp_db remediation — Stage 10: A3 production path boundaries

**Date:** 2026-06-04  
**Audit:** A3 (production slice)

## Summary

Production flow no longer imports monolithic `core.kp_db`:

| Module | Change |
|--------|--------|
| [`core/plan_commit.py`](../../core/plan_commit.py) | `kp_db_plates.mark_plates_as_planned`, `return_plan_plates_to_production` |
| [`app/services/production_planning_service.py`](../../app/services/production_planning_service.py) | `kp_db_plates`, `kp_db_schema`, `RestMatchingService` |
| [`app/services/production_completion_service.py`](../../app/services/production_completion_service.py) | `kp_db_plates`, `kp_db_rests`, `kp_db_schema`, `PlateCompletionService` |

## Guard

[`tests/test_production_services_kp_boundary.py`](../../tests/test_production_services_kp_boundary.py) — forbids `from core import kp_db` in `production_*.py`.

## Verification

```bash
pytest tests/test_production_services_kp_boundary.py tests/test_plan_consistency.py tests/test_production_completion_service.py -q
```
