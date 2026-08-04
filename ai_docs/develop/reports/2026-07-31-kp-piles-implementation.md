# Report: КП на сваи (MVP)

**Date:** 2026-07-31  
**Spec:** [`ai_docs/specs/kp-piles.md`](../../specs/kp-piles.md)  
**Plan:** [`ai_docs/develop/plans/2026-07-30-kp-piles.md`](../plans/2026-07-30-kp-piles.md)

## Summary

MVP «КП на цельные сваи» реализован end-to-end: wizard, API, save в `kp_piles`, PDF/XLSX, архив с фильтром, исключение из производства.

## AC checklist

| AC | Status | Tests / verify |
|----|--------|----------------|
| AC-1 | ✅ | `ProductTypePicker.tsx`, manual |
| AC-2 | ✅ | `PileInputStep.tsx`, `test_commercial_pile_flow.py` |
| AC-3 | ✅ | `test_pile_line_parser.py` |
| AC-4 | ✅ | `KpPilePreviewPanel.tsx` |
| AC-5 | ✅ | `test_update_pile_grades_bulk` |
| AC-6 | ✅ | `test_pile_draft_calculate_rejects_unknown_mark`, `test_wizard_state_piles_unpriced_blocks_proceed` |
| AC-7 | ✅ | `test_wizard_state_piles_no_wide_plates_gate` |
| AC-8 | ✅ | `test_pile_xlsx_contains_mark_and_grade` |
| AC-9 | ✅ | `test_pile_draft_save_to_archive`, `test_kp_persistence_piles.py` |
| AC-10 | ✅ | `test_archive_pile.py` (filter), frontend archive |
| AC-11 | ✅ | `test_commercial_web_flow.py` (59 passed) |
| AC-12 | ✅ | `test_production_pile_exclusion.py` |
| AC-13 | ✅ | `test_save_pile_kp_gets_next_kp_id_after_plates` |
| AC-14 | ✅ | `test_archive_detail_includes_piles`, `OfferDetailsDrawer.test.tsx` |
| AC-15 | ✅ | `test_archive_generate_pdf_for_saved_pile_kp`, `test_archive_download_pdf_http_for_pile_kp` |
| AC-16 | ✅ | `OfferDetailsDrawer.test.tsx` (disabled + «скоро») |
| AC-17 | ✅ | `test_pile_line_parser.py` merge tests |

## Test results (final)

```
pytest tests/ -k "pile or commercial_pile or archive_pile" -q  → 48 passed
pytest tests/test_commercial_web_flow.py -q                   → 59 passed
cd frontend && npm run test && npm run build                  → green
```

## Bug fixed during TDD closure

**Preview crash on unknown mark:** `get_draft_details` вызывал `calculate_total_cost(require_all_priced=True)` и падал при создании черновика с неизвестной маркой. Исправлено: preview totals с `require_all_priced=False`; блокировка на calculate/save через `ensure_order_priced`; `validation_errors` + `can_proceed_to=[]` в wizard.

## Manual smoke (recommended)

1. `/commercial-offer/new` → Сваи → текст `С120.35-12 B25 5` → save → архив
2. Неизвестная марка → ошибка в preview, calculate 422
3. OCR фото (если env настроен)

## Out of scope (unchanged)

Производство свай, fuzzy-match, смешанное КП, 1С, Telegram bot.
