# Report: КП на ФБС MVP

**Date:** 2026-08-05  
**Spec:** [`ai_docs/specs/kp-fbs.md`](../../specs/kp-fbs.md)  
**Plan:** [`ai_docs/develop/plans/2026-08-05-kp-fbs.md`](../plans/2026-08-05-kp-fbs.md)

## Summary

MVP «КП на ФБС» реализован end-to-end: picker «ФБС» → text/OCR/AI → после «Список верен» preview с классом бетона (B7_5 / B20 / B22_5 / B25, default **B25**) → client → PDF/XLSX → save в `kp_fbs`, архив с бейджем/фильтром. Цены: `fbs_prices` в `pb.db` (импорт только лист «Прайс»). Марка в КП/PDF — как ввёл менеджер. Bulk «класс ко всем» — apply where available, иначе skip + warning. Production whitelist только `plates`.

Шаблоны: **piles** (grade UX + PDF), **bridge_piles / marches** (multi-product wiring + hide preview until confirm). Без T/В alias groups (в прайсе нет).

## AC checklist

| AC | Status | Tests / verify |
|----|--------|----------------|
| AC-1 picker «ФБС» → `fbs` | ✅ | `ProductTypePicker` + vitest |
| AC-2 input text/photo/AI | ✅ | `FbsInputStep`, OCR fbs pipeline |
| AC-3 parser mark+grade+qty; merge | ✅ | `test_fbs_line_parser.py` |
| AC-4 preview available grades; hide until confirm | ✅ | `KpFbsPreviewPanel` |
| AC-5 CLI import лист «Прайс»; 14 marks × 4 grades | ✅ | 56 price rows; import tests |
| AC-6 missing mark+grade → block | ✅ | pricing + flow |
| AC-7 client/result; pdf+xlsx only | ✅ | export non-plate file types |
| AC-8 PDF/XLSX: mark as typed, grade, qty, price | ✅ | commercial_offer / xlsx fbs branch |
| AC-9 save `kp_fbs` + meta | ✅ | persistence + flow test |
| AC-10 archive badge + filter | ✅ | archive UI + types |
| AC-11 plate/pile/step/march/bridge_pile regression | ✅ | focused pytest 261 passed |
| AC-12 not in production | ✅ | `test_production_fbs_exclusion.py` |
| AC-13 shared `kp_id` | ✅ | shared `KP_offers` |
| AC-14–17 mirror piles | ✅ | grades bulk, drawer, regen |

## Price import

```bash
python scripts/import_fbs_prices_from_xlsx.py \
  "банк знаний/7_5 В прайс на ФБС  от 03.08.2026.xlsx" --sheet Прайс
# → 56 строк (14 марок × 4 классов бетона)
```

Grades: `B7_5`, `B20`, `B22_5`, `B25`. Default / preferred: **B25**. Lookup нормализует пробелы/регистр и `Т`↔`T`.

## Key files

**Backend:** `core/fbs_*`, `scripts/import_fbs_prices_from_xlsx.py`, `app/services/commercial_fbs_service.py`, endpoints `.../fbs`, OCR gate/pipeline.

**Frontend:** `ProductTypePicker`, `FbsInputStep`, `KpFbsPreviewPanel`, wizard/API/archive wiring.

**Tests:** `tests/test_fbs_*`, `tests/test_commercial_fbs_*`, `tests/test_production_fbs_exclusion.py`, frontend vitest.

## Verification run

```bash
pytest tests/ -k "fbs or bridge_pile or march or step or pile or wizard" -q
# → 261 passed, 4 skipped

cd frontend && npm run typecheck && npm run test -- --run \
  src/features/commercial-offer/components/ProductTypePicker.test.tsx \
  src/features/commercial-offer/lib/buildFbsPreviewRows.test.ts \
  src/features/commercial-offer/lib/wizardStepOrder.test.ts
# → green; full vitest has unrelated ShipmentDrawer timeout flake
npm run build  # → green
```

## Smoke test

1. Import prices (CLI above).
2. UI: picker «ФБС» → `ФБС 9.3.6-Т 2` → «Список верен».
3. Grade dropdown / bulk B25 → client → calculate → PDF → archive badge «ФБС».

## Remaining / next

- Next product: **TBD** (handoff).
- Discuss **generic multi-product framework** after this clone (duplication across piles/steps/marches/bridge_piles/fbs is intentional for now).
- Manual browser smoke on live stack after restart `./run+logs.sh`.

## Status

**Implemented** (2026-08-05). Plan tasks FBS-001…FBS-604 complete for MVP scope.
