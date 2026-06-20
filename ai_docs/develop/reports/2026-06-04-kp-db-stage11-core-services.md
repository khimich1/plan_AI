# kp_db remediation — Stage 11: core domain services (no core→app)

**Date:** 2026-06-04

## Summary

Domain orchestration lives under `core/`; `app/services/*` re-export for backward compatibility:

| Service | Location |
|---------|----------|
| Plate completion | [`core/plate_completion_service.py`](../../core/plate_completion_service.py) |
| Rest matching | [`core/rest_matching_service.py`](../../core/rest_matching_service.py) |
| KP save | [`core/kp_persistence_service.py`](../../core/kp_persistence_service.py) |

Facades in `kp_db_plates` / `kp_db_rests` / `kp_db_offers` delegate to core services (no `import app`).

## Guard

[`tests/test_core_no_app_import.py`](../../tests/test_core_no_app_import.py)

## Verification

```bash
pytest tests/test_core_no_app_import.py tests/test_kp_db_move_plates_to_completed.py -q
```
