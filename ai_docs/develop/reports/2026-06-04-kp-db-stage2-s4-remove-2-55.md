# kp_db stage 2 — S4 remove step 2.55 (2026-06-04)

Closes audit item **S4** (cross-KP write-off in `find_one_row`) from `2026-06-03-core-kp-db-audit.md`.

## Change

- Removed step **2.55** in `core/kp_db.py` (`find_one_row`): no global search by length for 61,1/61,2 across all KPs when `allow_cross_kp=False`.
- Removed associated NDJSON debug (`H_61_cross_kp`).

**Preserved:**

- Step **2.5** — equivalent plate names (61,1 ↔ 61,2, 59,8 ↔ 59,9) within `prefer_kp_id`.
- Steps **5–6** — cross-KP only when `allow_cross_kp=True` (and `plan_ids` where applicable).
- `kp_plate_id` direct lookup in `move_plates_to_completed` (web P5 path).

## Tests

- `tests/test_kp_db_move_plates_to_completed.py::test_g2_55_cross_kp_61_1_61_2_when_disabled` — expects `count == 0`, stock in other KP unchanged.

## Verification

```
pytest tests/test_kp_db_move_plates_to_completed.py \
  tests/test_kp_db_split_preserves_identity.py \
  tests/test_kp_db_find_matching_rests.py \
  tests/test_production_completion_service.py \
  tests/test_plan_consistency.py -v
```

**41 passed.**

## Risk / follow-up

- **Bot** completion does not pass `kp_plate_id`; fuzzy `find_one_row` may return “not found” if plan `kp_id` does not match DB row (previously masked by 2.55 for 61,x).
- **Legacy** plans without `kp_plate_id` — same risk on length/name mismatch across KPs.

If production shows unmoved 61,x plates:

1. Pass `kp_plate_id` in bot complete_day payload (align with `production_completion_service`).
2. Fix optimizer `kp_id` attribution in `core/optimization/order_dispatch.py`.
3. If needed: narrow cross-KP by `plan_id IN plan_ids`, not global length scan.

## Not in scope (at stage 2)

- P0 debug NDJSON cleanup (S1/A7) — done in [stage 3](2026-06-04-kp-db-stage3-s7-s1.md)
- `clear_all_*` production guards (S7) — done in [stage 3](2026-06-04-kp-db-stage3-s7-s1.md)
- A1/A2 decomposition
