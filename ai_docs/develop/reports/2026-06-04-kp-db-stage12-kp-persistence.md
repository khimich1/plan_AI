# kp_db remediation — Stage 12: A2 `save_kp_to_db` orchestration

**Date:** 2026-06-04  
**Audit:** A2 (offers slice), M3/Q9 (partial — BLOB path validation unchanged)

## Summary

`save_kp_to_db` orchestration moved to [`core/kp_persistence_service.KpPersistenceService`](../../core/kp_persistence_service.py).

- [`core/kp_db_offers.save_kp_to_db`](../../core/kp_db_offers.py) — backward-compatible facade
- [`app/services/offers_service.py`](../../app/services/offers_service.py) — calls `KpPersistenceService` directly
- [`app/services/kp_persistence_service.py`](../../app/services/kp_persistence_service.py) — re-export

## Remaining (backlog)

- Split SQL primitives (`INSERT` offer / plate / file) out of service body
- Further commercial callers via dedicated app service only

## Verification

```bash
pytest tests/test_kp_persistence_service.py tests/test_kp_db_xlsx_path_validation.py tests/test_kp_db_split_preserves_identity.py -q
```
