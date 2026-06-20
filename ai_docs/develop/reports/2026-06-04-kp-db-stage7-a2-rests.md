# kp_db remediation — Stage 7: A2 rests (`RestMatchingService`)

**Date:** 2026-06-04  
**Audit:** [2026-06-03-core-kp-db-audit.md](../audits/2026-06-03-core-kp-db-audit.md) — A2 (rests slice)

## Summary

`find_matching_rests` orchestration moved from [`core/kp_db_rests.py`](../../core/kp_db_rests.py) into [`app/services/rest_matching_service.py`](../../app/services/rest_matching_service.py). Domain rules in [`core/domain/rest_matching.py`](../../core/domain/rest_matching.py).

## Architecture

| Module | Role |
|--------|------|
| `core/domain/rest_matching_types.py` | `RestMatch`, `MatchType` |
| `core/domain/rest_matching.py` | `classify_match_type`, `compute_cut_cost` |
| `core/kp_db_rests.py` | `_fetch_available_rests_candidates`; thin facade |
| `app/services/rest_matching_service.py` | `find_matching_rests_on_cursor`, `find_matching_rests` |
| `app/services/production_planning_service.py` | Calls `RestMatchingService` directly |

## Verification

```bash
pytest tests/test_kp_db_find_matching_rests.py tests/test_rest_matching_service.py -q
```

**Result:** 10 passed.

## Remaining

- Stage 8: debug NDJSON in completion hot path
- Transitional `core → app` shim in `kp_db_rests.find_matching_rests` (removed in Stage 11 pattern)
