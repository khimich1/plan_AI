# Детальная разбивка цен — формат breakdown.xlsx

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-09-02  
**Leftover from**: [kp-append-preview-and-fresh-breakdown](./kp-append-preview-and-fresh-breakdown.md) (Out of Scope **E**)  
**Spec**: [../specs/kp-breakdown-xlsx-format.md](../specs/kp-breakdown-xlsx-format.md)  
**Plan**: [../develop/plans/2026-09-02-kp-breakdown-xlsx-format.md](../develop/plans/2026-09-02-kp-breakdown-xlsx-format.md)

## Problem Statement

`*_breakdown.xlsx` открывался с узкими колонками → визуально обрезанные «Базовая», «Попереч», «Продоль» и тесные формулы, хотя ячейки уже содержали полные подписи из `build_component_breakdown`.

## Root cause

`save_breakdown_to_excel` писал через pandas `to_excel` **без** `column_dimensions` → дефолт ~13 символов в LibreOffice/Excel.

## Done

- openpyxl writer + min widths (A≥42, B≥38, C≥18), bold header/product rows.
- Tests: headers, empty block separators, full labels, ` руб` suffix, column widths.
