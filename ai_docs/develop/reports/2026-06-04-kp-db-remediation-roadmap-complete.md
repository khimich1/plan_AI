# kp_db remediation roadmap — completed (2026-06-04)

Follow-up to [2026-06-03-core-kp-db-audit.md](../audits/2026-06-03-core-kp-db-audit.md).

## Summary

| Phase | Status | Deliverables |
|-------|--------|--------------|
| 0.1 | Done | `kp_plate_id` in bot `production_completion`; test `test_g_kp_plate_id_direct_lookup_wrong_name` |
| 0.2 | Done | `bot/handlers/debug_util.py`; gated NDJSON in bot handlers |
| 1.1 | Done | `core/kp_db_plates.py` + re-export from `kp_db` |
| 1.2 | Done | `core/kp_db_rests.py` + re-export |
| 2 | Done | `ensure_schema()` idempotent; startup in `app/main.py`, `bot/bot_main.py` |
| 3 | Done | `core/domain/plate_completion_matching.py`, `app/services/plate_completion_service.py` |
| 4 | Partial | A3 facade `bot/services/kp_persistence.py`; A5 audit via `kp_db_common._audit_append`; A6 indexes on `kp_plates` |
| 5 | Done | This report + audit status below |

## Architecture after split

```
core/kp_db_common.py       — DEFAULT_DB, _connect, _audit_append
core/kp_db.py              — schema, CRUD КП, managers, ensure_schema, re-exports
core/kp_db_plates.py       — lifecycle / completion
core/kp_db_rests.py        — plate_rests
core/domain/plate_completion_matching.py — find_kp_plate_row (steps 0–7)
app/services/plate_completion_service.py — domain facade for matching
bot/services/kp_persistence.py         — handler facade (commercial, completion)
```

## Health score (estimate)

- Before: **2.0 / 10**
- After first roadmap: **~4.5 / 10** (A1 partial, A2 partial, P0/P1 closed, hot-path `ensure_schema` cheap)

Remaining Critical: full A2 orchestration out of `move_plates_to_completed` transaction body; further A1 slices (`kp_offers` CRUD).

## Verification

```bash
pytest tests/test_kp_db_move_plates_to_completed.py \
  tests/test_kp_db_find_matching_rests.py \
  tests/test_kp_db_split_preserves_identity.py \
  tests/test_destructive_db_guard.py \
  tests/test_production_completion_service.py \
  tests/test_plate_audit.py \
  tests/test_kp_db_agent_debug_log.py -v
```
