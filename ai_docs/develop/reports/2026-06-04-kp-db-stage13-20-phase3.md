# Phase 3 remediation: stages 13–20 (kp_db)

**Date:** 2026-06-04  
**Baseline:** re-audit strict Health 4.0, 94 pytest (phase 2)  
**Result:** **118 pytest passed** (phase 2 set + phase 3 additions)

## Stages delivered

| Stage | ID | Deliverable |
|-------|-----|-------------|
| 13 | A7/S1 | Removed agent NDJSON from `core/kp_db_plates` (pre-split), `core/plan_commit.py` |
| 14 | S8 | `resolve_kp_xlsx_path_for_write` / `resolve_kp_xlsx_output_path`; `get_xlsx_file` write guard; tests |
| 15 | S4 | Runbook `ai_docs/develop/guides/allow-cross-kp-runbook.md`; docstring in matching |
| 16 | A1 | Split: `kp_db_plates_common`, `_completion`, `_planning`, `_queries`; shim `kp_db_plates.py` (59 lines) |
| 16b | Q1/A13 | `insert_kp_plate_remainder_row` DRY; used in planning split/return paths |
| 17 | A3/A8 | `KpRepository`, `KpArchiveRepository`, `AdminService` → slice imports; `test_app_kp_boundary.py` |
| 18 | A4 | Removed `init_schema` from app repos/services, planning/completion/queries slices |
| 19 | A6/Q3 | Strategy-based `find_kp_plate_row`; tolerances as constants; `test_plate_completion_matching_steps.py` |
| 20 | — | This report; pytest 118 passed |

## Module sizes (A1)

| Module | Lines |
|--------|------:|
| `kp_db_plates.py` (shim) | 59 |
| `kp_db_plates_common.py` | 180 |
| `kp_db_plates_completion.py` | 94 |
| `kp_db_plates_planning.py` | 792 |
| `kp_db_plates_queries.py` | 249 |

`kp_db_plates_planning.py` remains large (plan/return flows); further split optional in phase 4.

## OPEN after phase 3

- **A1 (partial):** planning sub-slice still >500 lines  
- **A2:** domain SQL in matching (strategies only, no ports)  
- **A3:** `plan_manager`, `bot/services/kp_persistence.py` still use `core.kp_db` facade in places  
- **S6, A9, A12, A14:** backlog phase 4+

## Verification

```bash
pytest tests/test_kp_db_agent_debug_log.py tests/test_kp_db_xlsx_path_validation.py \
  tests/test_kp_db_move_plates_to_completed.py tests/test_plate_completion_service.py \
  tests/test_kp_db_find_matching_rests.py tests/test_rest_matching_service.py \
  tests/test_kp_persistence_service.py tests/test_production_completion_service.py \
  tests/test_plate_audit.py tests/test_plan_consistency.py tests/test_bot_kp_boundary.py \
  tests/test_production_services_kp_boundary.py tests/test_core_no_app_import.py \
  tests/test_kp_db_split_preserves_identity.py tests/test_destructive_db_guard.py \
  tests/test_kp_db_init_managers_seed.py tests/test_bot_auth.py \
  tests/test_app_kp_boundary.py tests/test_plate_completion_matching_steps.py -q
```

## Health estimate (post phase 3)

| Metric | Before | After (estimate) |
|--------|--------|------------------|
| Critical OPEN | 1 (god-slice) | **0** (shim + slices; planning size residual) |
| Strict Health | 4.0 | **~5.5–5.8** |
| Pragmatic Health | ~6.2 | **~7.0** |
