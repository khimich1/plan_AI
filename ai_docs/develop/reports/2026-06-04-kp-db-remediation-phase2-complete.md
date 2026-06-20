# kp_db remediation phase 2 — complete (2026-06-04)

Follow-up to [stage 6](2026-06-04-kp-db-stage6-a2-completion.md) and [post-stage6 plan](.cursor/plans/kp-db_post-stage6_e8c16f14.plan.md).

## Stages delivered

| Stage | Report | Audit IDs |
|-------|--------|-----------|
| 7 | [stage7-a2-rests](2026-06-04-kp-db-stage7-a2-rests.md) | A2 rests |
| 8 | [stage8-a7-debug](2026-06-04-kp-db-stage8-a7-debug.md) | A7, S1 |
| 9 | [stage9-a5-audit](2026-06-04-kp-db-stage9-a5-audit.md) | A5 |
| 10 | [stage10-a3-production](2026-06-04-kp-db-stage10-a3-production.md) | A3 production |
| 11 | [stage11-core-services](2026-06-04-kp-db-stage11-core-services.md) | layer boundary |
| 12 | [stage12-kp-persistence](2026-06-04-kp-db-stage12-kp-persistence.md) | A2 save_kp |

## Architecture (target state)

```
core/kp_db_audit.py              — plate_status_log INSERT
core/plate_completion_service.py — completion orchestration
core/rest_matching_service.py    — rest matching orchestration
core/kp_persistence_service.py   — KP save orchestration
core/kp_db_plates.py             — SQL primitives + facades
core/kp_db_rests.py
core/kp_db_offers.py
core/kp_db.py                      — re-export shim
app/services/*                     — thin re-exports + production services
```

## Verification (baseline + phase 2)

```bash
pytest tests/test_kp_db_move_plates_to_completed.py \
  tests/test_plate_completion_service.py \
  tests/test_kp_db_find_matching_rests.py \
  tests/test_rest_matching_service.py \
  tests/test_kp_persistence_service.py \
  tests/test_production_completion_service.py \
  tests/test_plate_audit.py \
  tests/test_plan_consistency.py \
  tests/test_bot_kp_boundary.py \
  tests/test_production_services_kp_boundary.py \
  tests/test_core_no_app_import.py \
  tests/test_kp_db_agent_debug_log.py -q
```

**Result:** 94 passed (baseline was 80).

## Health score (re-estimate)

| Metric | Before audit | After phase 1 | After phase 2 |
|--------|--------------|---------------|---------------|
| Health Score | 2.0 | ~4.5 | **~6.2** |

Formula adjustment: Critical A2 largely addressed (3 slices); High A3/A5/A7 production path closed; remaining High items (A6, S6, full A3 commercial) keep score below 8.

## Backlog (phase 3)

- A6/Q2: simplify `find_kp_plate_row` steps / indexes
- A9: Unit of Work for multi-step transactions
- A12: TypedDict on plate dict boundaries
- Q11: bulk agent log in `production_execution.py`
- S6: SQLCipher / BLOB limits
- Wave 2 A3: `offers_service`, `admin_service`, `plan_manager` off monolith `kp_db`
