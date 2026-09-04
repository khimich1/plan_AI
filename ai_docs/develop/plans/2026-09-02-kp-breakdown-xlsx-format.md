# Implementation Plan: breakdown.xlsx format

**Спека**: [../../specs/kp-breakdown-xlsx-format.md](../../specs/kp-breakdown-xlsx-format.md)  
**Идея**: [../../ideas/kp-breakdown-xlsx-format.md](../../ideas/kp-breakdown-xlsx-format.md)  
**Дата**: 2026-09-02  
**Статус**: PLAN ✅ · IMPLEMENT ✅

## Tasks

1. RED: `tests/test_save_breakdown_to_excel.py` (structure, labels, currency, widths).
2. GREEN: `save_breakdown_to_excel` — openpyxl widths + bold headers.
3. Verify: `pytest tests/test_save_breakdown_to_excel.py`.

## Files

- `core/commercial_offer.py` — `save_breakdown_to_excel`
- `tests/test_save_breakdown_to_excel.py`
