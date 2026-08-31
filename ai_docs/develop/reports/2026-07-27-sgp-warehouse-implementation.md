# Report: Склад готовой продукции (СГП)

**Date:** 2026-07-27  
**Spec:** [`ai_docs/specs/sgp-warehouse.md`](../../specs/sgp-warehouse.md)  
**Plan:** [`ai_docs/develop/plans/2026-07-27-sgp-warehouse.md`](../plans/2026-07-27-sgp-warehouse.md)  
**Status:** ✅ Implemented (MVP)

## Summary

Разделены физический склад (`completed_plates`) и потребность КП (`kp_plates`). День уходит на СГП, есть вкладка склада с unlink/relink, wizard умеет «закрыть со склада», статус КП «На СГП» + бейдж N/M, экспорт XLSX «Со склада», qty-инварианты и plate_loss regression PASS.

## What was built

- Schema: nullable `completed_plates.kp_id`, `plan_id`, `kp_meta.ordered_qty`
- Enums: `ON_SGP`, reasons `sgp_*`
- `send_to_sgp` (+ alias `complete_day`), orphan pre-flight 422
- `SgpService`: list / unlink / relink / free / reserve / export
- API `/production/sgp/*`, `/plans/{id}/sgp-export`
- Frontend: вкладка СГП, DayDrawer labels, wizard badge+confirm, «с СГП» в day view, кнопка XLSX
- Archive: статус «На СГП» во «В производстве», бейдж N/M
- delete_plan: `plan_id=NULL` на СГП, qty не трогается

## Verification

- `pytest tests/test_sgp_*.py tests/test_production_completion_service.py` — green
- `scripts/run_plate_loss_regression.py` — PASS (orphan Σ=0)
- `cd frontend && npm run build` — green

## Out of scope (unchanged)

- Рез донора / отгрузка / откат дня / QR / Telegram
