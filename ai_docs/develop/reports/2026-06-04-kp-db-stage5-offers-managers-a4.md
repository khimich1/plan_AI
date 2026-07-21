# kp_db stage 5 remediation (2026-06-04)

Follow-up to [2026-06-03-core-kp-db-audit.md](../audits/2026-06-03-core-kp-db-audit.md). Closes **A1** slice (offers + managers) and **A4** cleanup (schema entrypoints).

## Summary

| Phase | Status | Deliverables |
|-------|--------|--------------|
| 5a | Done | [`core/kp_db_schema.py`](../../core/kp_db_schema.py) — `_init_schema_impl`, idempotent `ensure_schema` |
| 5b | Done | [`core/kp_db_managers.py`](../../core/kp_db_managers.py) — CRUD + JSON seed |
| 5c | Done | [`core/kp_db_offers.py`](../../core/kp_db_offers.py) — KP CRUD, XLSX, clear, search, metrics |
| 5d | Done | Removed per-function `ensure_schema` / `_init_schema`; guard [`tests/test_kp_db_schema_boundary.py`](../../tests/test_kp_db_schema_boundary.py) |
| 5e | Done | Facade [`core/kp_db.py`](../../core/kp_db.py) (~145 lines); this report |

## Architecture after stage 5

```
core/kp_db_common.py    — DEFAULT_DB, _connect, _audit_append
core/kp_db_schema.py    — DDL, ensure_schema (once per db path)
core/kp_db_offers.py    — KP_offers, kp_meta, kp_files, destructive clear
core/kp_db_managers.py  — managers + init_default_managers
core/kp_db_plates.py    — plate lifecycle
core/kp_db_rests.py     — plate_rests
core/kp_db.py           — backward-compatible re-export facade
```

## A4 — schema entrypoints (contract)

| Location | Role |
|----------|------|
| `app/main.py` lifespan | `kp_db.ensure_schema(plita_db_path)` |
| `bot/bot_main.py` | same on bot startup |
| `app/repositories/kp_repository.py` | `init_schema` in `__init__` (safety net) |
| `tests/helpers/kp_db_fixtures.make_iso_db` | `init_schema` for isolated DB |
| `bot/handlers/production_execution.py` | `ensure_schema` when alternate `db_path` |

Persistence slices (`kp_db_offers`, `kp_db_managers`, `kp_db_plates`, `kp_db_rests`) **do not** call schema init (enforced by guard test).

## Scripts added

- `scripts/_extract_stage5.py` — one-off extraction helper
- `scripts/_cleanup_stage5_a4.py` — remove redundant `ensure_schema` calls

## Verification

```bash
pytest tests/test_kp_db_move_plates_to_completed.py \
  tests/test_kp_db_find_matching_rests.py \
  tests/test_kp_db_split_preserves_identity.py \
  tests/test_kp_db_init_managers_seed.py \
  tests/test_kp_db_xlsx_path_validation.py \
  tests/test_kp_db_search_by_customer.py \
  tests/test_kp_db_update_logistics_cost.py \
  tests/test_destructive_db_guard.py \
  tests/test_kp_db_agent_debug_log.py \
  tests/test_plate_audit.py \
  tests/test_production_completion_service.py \
  tests/test_bot_kp_boundary.py \
  tests/test_bot_debug_gated.py \
  tests/test_kp_db_schema_boundary.py -v
```

**97 passed** (2026-06-04 run).

## Not in scope (stage 6+)

- **A2** — domain orchestration still in `move_plates_to_completed` / `save_kp_to_db`
- **A5** — unify audit via `PlateAuditRepository`
- **A3 app** — `plan_commit` / services still import `core.kp_db` facade
- **Alembic** — runtime DDL remains in `kp_db_schema`

## Suggested next step

Stage 6: **A2** — move completion orchestration out of `kp_db_plates.move_plates_to_completed` transaction body into `PlateCompletionService`.
