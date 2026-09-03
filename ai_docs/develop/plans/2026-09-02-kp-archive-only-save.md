# Implementation Plan: КП create — archive-only save

**Спека**: [../../specs/kp-archive-only-save.md](../../specs/kp-archive-only-save.md)  
**Идея**: [../../ideas/kp-archive-only-save.md](../../ideas/kp-archive-only-save.md)  
**Дата**: 2026-09-02  
**Статус**: PLAN ✅ · IMPLEMENT ✅

## Tasks

1. RED: `SaveOfferSection.test.tsx` — нет срока/альтернатив; save → archive.
2. GREEN: упростить `SaveOfferSection`; `saveOfferSchema` archive-only; отвязать execution terms от Result step.
3. Verify: vitest SaveOfferSection + CalculationResultStep + MoveToProductionDialog; typecheck.

## Files

- `frontend/.../SaveOfferSection.tsx` (+ test)
- `frontend/.../schemas/commercialOffer.ts`
- `frontend/.../CalculationResultStep.tsx` (+ test props)
- `frontend/.../CommercialOfferWizard.tsx`
