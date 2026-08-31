# Production Day Archive Badge

**Date:** 2026-05-08  
**Status:** Implemented  
**Related Code:**
- `frontend/src/features/production/components/DayDrawer.tsx`
- `app/schemas/production.py` (field: `write_off_completed`)
- `app/services/day_view_service.py` (aggregation logic)
- `frontend/src/index.css` (styles: `.day-plate-badge--done`)

## Overview

Write-off completed snapshots (`write_off_completed=true`) now display a **`(ГОТОВО)`** badge instead of the previous "Списано" text. Badge uses production styling for consistency with the day view design system.

## Changes

### UI Label
- **Old:** Badge text "Списано" (written off).
- **New:** Badge text **(ГОТОВО)** (ready / done).
- **Styling:** Same CSS classes: `.day-plate-badge` + `.day-plate-badge--done`.

### Row Styling
- No change to row class (`.day-plates-table__row--written-off` remains).
- Badge styling indicates completion state clearly.

### Data Flow
1. Backend (`day_view_service`): Aggregates `DayPlateInfo` with `write_off_completed` field.
2. Schema (`app/schemas/production.py`): Field included in response.
3. Frontend (`DayDrawer.tsx`): Reads flag and renders badge with new text.

## CSS Classes

In `frontend/src/index.css`:
```css
.day-plate-badge--done
  /* Production styling for completed write-off badge */
```

Row class (unchanged):
```css
.day-plates-table__row--written-off
  /* Row styling for written-off plates */
```

## Behavior

- Rows with `write_off_completed=true` show badge **(ГОТОВО)**.
- Editing is blocked for these rows (braces/defect/completion fields are read-only).
- Intended to provide visual feedback on archive state.

## Testing

Related tests:
- `tests/test_day_view_service.py` — Aggregation of `write_off_completed`.
- `tests/test_production_completion_service.py::test_day_view_write_off_completed_false_before_complete_true_after_snapshot` — Lifecycle.
- Frontend: Vitest for badge rendering (if added).

## Related Documentation

- [CHANGELOG](../../changelog/CHANGELOG.md) — Feature entry under "Added".
