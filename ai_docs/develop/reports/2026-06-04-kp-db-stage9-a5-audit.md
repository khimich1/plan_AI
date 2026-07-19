# kp_db remediation — Stage 9: A5 unified plate audit

**Date:** 2026-06-04  
**Audit:** A5

## Summary

Single INSERT implementation for `plate_status_log`:

| Module | Role |
|--------|------|
| [`core/kp_db_audit.py`](../../core/kp_db_audit.py) | `audit_append` — canonical SQL |
| [`app/repositories/plate_audit_repository.py`](../../app/repositories/plate_audit_repository.py) | App-layer facade over `audit_append` |
| [`core/kp_db_plates.py`](../../core/kp_db_plates.py) | Calls `audit_append` directly (same transaction) |
| [`core/kp_db_common.py`](../../core/kp_db_common.py) | `_audit_append` deprecated wrapper |

## Verification

```bash
pytest tests/test_plate_audit.py -q
```
