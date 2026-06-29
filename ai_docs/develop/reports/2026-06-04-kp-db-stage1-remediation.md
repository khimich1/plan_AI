# kp_db stage 1 remediation (2026-06-04)

Completed audit items Q3, Q1, S2, S5 from `2026-06-03-core-kp-db-audit.md`.

## Q3 — regression tests

- `tests/helpers/kp_db_fixtures.py` — isolated DB seed/snapshot helpers
- `tests/test_kp_db_move_plates_to_completed.py` — 12 golden cases (incl. G2.55 cross-KP documented)
- `tests/test_kp_db_find_matching_rests.py` — R1–R7
- `tests/test_kp_db_split_preserves_identity.py` — Q1 guard

## Q1 — split identity fields

- `mark_plates_as_planned`: SELECT + split INSERT copy `nomenclature_id`, `length_dm_raw`
- `return_plates_to_production`, `return_plate_rows_for_plan`, `recover_stuck_plates`: INSERT SELECT extended

## S2 — managers PII

- Removed hardcoded contacts from `init_default_managers`
- Seed: `data/managers_seed.json`, example `data/managers_seed.example.json`
- Env: `MANAGERS_SEED_PATH` (documented in `.env.example`)
- `scripts/init_managers.py --seed`

## S5 — XLSX path safety

- `core/kp_file_paths.py` — whitelist + realpath
- `save_kp_to_db`, `save_xlsx_file` use `resolve_kp_xlsx_path_for_read`

## Verification

`pytest` (7 kp_db modules + `test_plate_audit.py`): **43 passed**.

Not in scope: S4 (cross-KP fix), P0 debug NDJSON, A1/A2 decomposition.
