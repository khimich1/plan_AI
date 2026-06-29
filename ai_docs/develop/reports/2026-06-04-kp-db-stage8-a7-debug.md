# kp_db remediation — Stage 8: A7/S1 completion hot path debug

**Date:** 2026-06-04  
**Audit:** A7, S1, Q14 (trace substrings removed)

## Summary

Removed unconditional agent NDJSON from completion matching and orchestration:

- [`core/domain/plate_completion_matching.py`](../../core/domain/plate_completion_matching.py) — no `append_agent_debug_log` / hardcoded plate traces
- [`core/plate_completion_service.py`](../../core/plate_completion_service.py) — lean write-off loop without target-plate debug blocks

Gating remains in [`core/debug_paths.append_agent_debug_log`](../../core/debug_paths.py) for other modules.

## Verification

```bash
pytest tests/test_kp_db_agent_debug_log.py -q
```

Extended test: `test_find_kp_plate_row_writes_no_log_when_debug_disabled`.
