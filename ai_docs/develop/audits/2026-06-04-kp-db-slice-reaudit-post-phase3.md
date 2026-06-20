# Slice re-audit (post phase 3): `core/kp_db` boundaries

**Date:** 2026-06-04  
**Prior:** [2026-06-04-kp-db-slice-reaudit.md](2026-06-04-kp-db-slice-reaudit.md)  
**Remediation:** [2026-06-04-kp-db-stage13-20-phase3.md](../reports/2026-06-04-kp-db-stage13-20-phase3.md)

## Summary

Phase 3 closed **A1 Critical** (monolithic `kp_db_plates` decomposed), **A7/S1** hot-path debug in plates/plan_commit, **S8** write-path XLSX parity, **S4** runbook, **A3 wave 2** (app repository/admin), **A4** init_schema leaks in app + plate slices, **A6/Q3** matching refactor with per-step tests.

**pytest:** 118 passed (phase 2 baseline + `test_app_kp_boundary`, `test_plate_completion_matching_steps`).

## Health (estimated)

| Score | Value |
|-------|------:|
| Strict | **~5.6 / 10** |
| Pragmatic | **~7.0 / 10** |

## Remaining High (next wave)

- **A3:** `bot/services/kp_persistence.py`, `app/planning/plan_manager` (`core.kp_db` imports)  
- **A1 residual:** split `kp_db_plates_planning.py` (~792 lines)  
- **A2:** ports/adapters for domain matching without SQL  

## Closed in phase 3

- A1 god-slice `kp_db_plates` (replaced by shim + 4 modules)  
- A7/S1 plates + plan_commit agent NDJSON  
- S8 write XLSX validation  
- A4 init_schema in touched app paths and plate slices  
- A8 partial (`KpRepository` → `kp_db_offers`)  
- A6/Q3 matching structure + step tests  
