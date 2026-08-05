# Report: КП на лестничные марши (ЛМ) MVP

**Date:** 2026-08-05  
**Spec:** [`ai_docs/specs/kp-stair-marches.md`](../../specs/kp-stair-marches.md)  
**Plan:** [`ai_docs/develop/plans/2026-08-05-kp-stair-marches.md`](../plans/2026-08-05-kp-stair-marches.md)

## Summary

MVP «КП на лестничные марши (ЛМ)» реализован end-to-end: picker «Марши» → text/OCR/AI → после «Список верен» preview с **классом бетона** (как сваи) → client → PDF/XLSX → save в `kp_marches`, архив с бейджем/фильтром. Цены: матрица `march_prices` (7×5 = **35** строк) в `pb.db`. Production whitelist только `plates`.

Шаблоны: **piles** (grade matrix + `/marches/grades` + PDF колонка «Класс бетона»), **steps** (multi-product wiring + hide preview until confirm).

## AC checklist

| AC | Status | Tests / verify |
|----|--------|----------------|
| AC-1 picker «Марши» → `marches` | ✅ | `ProductTypePicker.tsx` + vitest |
| AC-2 input text/photo/AI | ✅ | `MarchInputStep.tsx`, OCR `marches`, flow test |
| AC-3 parser mark+grade+qty; merge; `2.8`→`2,8` | ✅ | `test_march_line_parser.py` |
| AC-4 preview grade dropdown + apply-all | ✅ | `KpMarchPreviewPanel.tsx`, `/marches/grades` |
| AC-5 CLI import ≥7×5 | ✅ | **35** rows imported; `test_march_price_import.py` |
| AC-6 missing mark+grade → block | ✅ | `test_commercial_march_pricing.py`, flow calculate |
| AC-7 client/result; pdf+xlsx only | ✅ | export service + wizard hide schema/breakdown |
| AC-8 PDF/XLSX: марка, класс, qty, цена, сумма | ✅ | pile-like columns; short mark |
| AC-9 save `kp_marches` + meta | ✅ | `test_kp_persistence_marches.py`, flow save |
| AC-10 archive badge + filter | ✅ | `test_archive_march.py`, archive UI |
| AC-11 plate/pile/step regression | ✅ | pile/step/wizard flows green (93 pytest matched) |
| AC-12 not in production | ✅ | `test_production_march_exclusion.py` |
| AC-13 shared `kp_id` | ✅ | shared `KP_offers` insert |
| AC-14–18 mirror piles | ✅ | drawer+grade, regen, production disabled UI, merge mark+grade, grades API |

## Price import

```bash
python scripts/import_march_prices_from_xlsx.py \
  "банк знаний/Прайс ЛМ от 03.08.2026.xlsx" --sheet Прайс
# → 35 строк (7 марок × 5 классов бетона) в pb.db
```

Canonical marks: `1ЛМ 27-11-14-4`, `1ЛМ 27-12-14-4`, `1ЛМ 30-11-15-4`, `1ЛМ 30-11-15-4 закладные справа`, `1ЛМ 30-12-15-4`, `ЛМ 2,8`, `ЛМ 2,9`.

## Test results

| Suite | Result |
|-------|--------|
| `pytest tests/ -k "march or commercial_pile or commercial_step or commercial_wizard"` | **93 passed**, 4 skipped |
| `pytest tests/test_commercial_wizard_step_service.py` | **20 passed** |
| Frontend typecheck | pass |
| Frontend vitest commercial-offer + archive | **65 passed** |
| Frontend build | pass |

## Key files

**Domain / core**
- `core/march_price_db.py`, `scripts/import_march_prices_from_xlsx.py`
- `core/march_line_parser.py`, `core/march_format_prompt.py`, `core/march_text_normalizer.py`
- `core/kp_db_schema.py` (`kp_marches`), persistence / order_data / offers_read
- `core/commercial_pricing.py` (`lookup_march_price`, `is_march_order`)
- `core/commercial_offer.py`, `commercial_offer_xlsx.py` (grade columns)
- OCR: `core/ocr/march_parser_gate.py`, pipeline/recognition/openai/gigachat branches

**App**
- `app/services/commercial_march_service.py`
- workflow/draft/wizard/calculation/export/archive services
- `app/api/v1/endpoints/commercial.py` — `/marches`, `/marches/ai`, `/marches/grades`
- schemas: `ProductType=marches`, `WizardStepId.marches`, `ingest_marches`

**Frontend**
- `ProductTypePicker` «Марши», `MarchInputStep`, `KpMarchPreviewPanel`
- wizard order / store / API / result (grade table, qty «маршей»)
- archive badge/filter/drawer with grade

**Tests**
- `tests/test_march_*`, `test_commercial_march_flow.py`, `test_kp_persistence_marches.py`, `test_archive_march.py`, `test_production_march_exclusion.py`
- FE: `buildMarchPreviewRows.test.ts`, picker/wizardStepOrder tests

## Preserved fixes

- Priced preview hidden while `pendingBatchReview` (until «Список верен»)
- Client-step validation errors not shown on product (marches) step

## Gaps / follow-ups

- No dedicated `test_march_ocr_pipeline.py` (OCR path wired; pile has a dedicated test)
- No component-level vitest for `MarchInputStep` / `KpMarchPreviewPanel` (lib + picker coverage)
- Generic multi-product framework still out of scope (next after ФБС consideration)
- ФБС — next product iteration

## Plan tasks

MARCH-001 → MARCH-604 completed (see plan checkboxes).
