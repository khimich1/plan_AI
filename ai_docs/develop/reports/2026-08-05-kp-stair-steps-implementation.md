# Report: КП на лестничные ступени (ЛС) MVP

**Date:** 2026-08-05  
**Spec:** [`ai_docs/specs/kp-stair-steps.md`](../../specs/kp-stair-steps.md)  
**Plan:** [`ai_docs/develop/plans/2026-08-05-kp-stair-steps.md`](../plans/2026-08-05-kp-stair-steps.md)  
**Idea:** [`ai_docs/ideas/kp-stair-steps.md`](../../ideas/kp-stair-steps.md)

## Summary

MVP «КП на лестничные ступени (ЛС)» реализован end-to-end по образцу свай **без** класса бетона: picker «Ступени», text/OCR/AI ingest, preview (марка|qty|цена), calculate → PDF/XLSX → save в `kp_steps`, архив с фильтром, production whitelist только `plates`. Прайс импортирован CLI → 42 марки в `step_prices`.

## AC checklist

| AC | Status | Tests / verify |
|----|--------|----------------|
| AC-1 picker | ✅ | `ProductTypePicker.tsx` + vitest |
| AC-2 input+OCR | ✅ | `StepInputStep.tsx`, OCR pipeline `steps`, `test_commercial_step_flow.py` |
| AC-3 parser | ✅ | `test_step_line_parser.py` |
| AC-4 preview no grade | ✅ | `KpStepPreviewPanel.tsx` + vitest |
| AC-5 import ≥42 | ✅ | `test_step_price_import.py`; CLI import 42 rows |
| AC-6 unpriced block | ✅ | `test_commercial_step_pricing.py`, step flow calculate 4xx |
| AC-7 client/result files | ✅ | export pdf+xlsx only; wizard hide schema/breakdown |
| AC-8 PDF/XLSX short mark | ✅ | `commercial_offer.py` / `xlsx` `is_step_order` branch |
| AC-9 save kp_steps | ✅ | `test_kp_persistence_steps.py`, step flow save |
| AC-10 archive badge/filter | ✅ | `test_archive_step.py`, archive UI |
| AC-11 regression | ✅ | pile flow 12/12; plate web 60/63 (3 pre-existing unrelated) |
| AC-12 production | ✅ | `test_production_step_exclusion.py`; whitelist plates |
| AC-13 kp_id series | ✅ | shared `KP_offers` insert |
| AC-14 drawer | ✅ | archive drawer without grade |
| AC-15 archive regen | ✅ | `test_archive_step.py` |
| AC-16 production disabled UI | ✅ | OfferDetailsDrawer «скоро» |
| AC-17 merge mark | ✅ | `test_step_line_parser.py` |
| AC-18 no grades API | ✅ | no `/steps/grades`; only piles grades remain |

## Key files

**Domain / core**
- `core/step_price_db.py`, `scripts/import_step_prices_from_xlsx.py`
- `core/step_line_parser.py`, `core/step_format_prompt.py`, `core/step_text_normalizer.py`
- `core/kp_db_schema.py` (`kp_steps`), `core/kp_persistence_service.py`, `core/kp_order_data.py`
- `core/commercial_pricing.py`, `core/commercial_offer.py`, `core/commercial_offer_xlsx.py`
- OCR: `core/ocr/pipeline.py` (`run_step_ocr_pipeline`), `recognition.py` (`apply_steps_with_ai`)

**API / services**
- `app/services/commercial_step_service.py`
- `app/services/commercial_workflow_service.py` (create/update/ai steps)
- `app/api/v1/endpoints/commercial.py` — `PATCH/POST .../steps`, `.../steps/ai`
- `app/schemas/commercial.py` — `ProductType=steps`, `WizardStepId.steps`, `ingest_steps`
- Archive + production: `archive_service.py`, `kp_repository.py` (plates whitelist)

**Frontend**
- `ProductTypePicker.tsx`, `StepInputStep.tsx`, `KpStepPreviewPanel.tsx`
- Wizard/archive wiring; types/API for steps

## Test results (final)

```
# Step foundation + flow + archive + production + pile regression
pytest tests/test_step_*.py tests/test_kp_steps_schema.py \
  tests/test_commercial_step_*.py tests/test_kp_persistence_steps.py \
  tests/test_production_step_exclusion.py tests/test_archive_step.py \
  tests/test_commercial_pile_flow.py tests/test_production_pile_exclusion.py -q
→ 66 passed

pytest tests/test_commercial_web_flow.py -q
→ 60 passed, 3 failed (pre-existing: schema context / get_next_kp_number / _build_offer_identity)

cd frontend && npm run typecheck && npm run test -- --run src/features/commercial-offer src/features/commercial-archive
→ typecheck OK; 62/62 vitest passed

cd frontend && npm run build
→ green (after removing unused useState in KpStepPreviewPanel)
```

## Price import (dev)

```
python scripts/import_step_prices_from_xlsx.py \
  "банк знаний/Прайс на лестничные ступени от 03.08.2026.xlsx" --sheet Прайс
→ 42 строк (42 уникальных марок ЛС) → pb.db
```

Marks stored normalized uppercase (e.g. `ЛС14-1ЛЕВ`); lookup/parser use the same normalize.

## Manual smoke (recommended)

1. `/commercial-offer/new` → Ступени → `ЛС11 2` → цена ≈ 1409.91 → PDF/XLSX → архив «Ступени»
2. Неизвестная марка → calculate blocked
3. Steps КП отсутствует в production wizard
4. Plate + pile smoke unchanged

## Out of scope (unchanged)

ЛМ, ФБС, grade UI, fuzzy-match, смешанное КП, производство/СГП для ступеней, generic multi-product framework, UI-импорт прайса.
