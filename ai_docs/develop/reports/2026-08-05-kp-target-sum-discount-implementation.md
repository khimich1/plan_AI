# KP target-sum discount — implementation report

**Date:** 2026-08-05

## Delivered

- Added shared, pure TypeScript target-sum/discount math with two-decimal rounding, bounds validation, product-base calculation, and the 16% approval predicate.
- Added one reusable confirmation dialog requiring the exact keyword `ПОДТВЕРЖДАЮ`.
- Added synchronized target-sum and discount drafts to both the commercial-offer result step and archive offer drawer. Both retain their existing discount persistence paths.
- Delivery is excluded from discounting: archive uses `delivery_service_total_rub`; wizard derives it from the current backend total, base products, and saved discount.

## Verification

- `npm run typecheck` — passed.
- `npm run test -- --run src/features/commercial-offer src/features/commercial-archive` — passed (19 files, 76 tests).
- `npm run build` — passed.
- `pytest tests/ -k "discount or commercial_pricing or archive" -q` — passed (97 tests). No Python code was added for this feature.

## Known limitation

The wizard has no explicit delivery total in its calculate response. The UI safely disables target-sum input if delivery inference is unavailable or materially negative. The formula follows the documented current-total inference; representative zero- and non-zero-logistics backend cases still require manual verification against a live draft.

The product rule requires two-decimal discount percentages and allows no residual line adjustment. For very large product bases, a 0.01% increment can mathematically move the total by more than 1 ₽; the client keeps the locked two-decimal percentage contract.
