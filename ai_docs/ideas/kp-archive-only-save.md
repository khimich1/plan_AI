# КП create: archive-only save

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-09-02  
**Leftover from**: [kp-append-preview-and-fresh-breakdown](./kp-append-preview-and-fresh-breakdown.md) (Out of Scope **D**)  
**Spec**: [../specs/kp-archive-only-save.md](../specs/kp-archive-only-save.md)  
**Plan**: [../develop/plans/2026-09-02-kp-archive-only-save.md](../develop/plans/2026-09-02-kp-archive-only-save.md)

## Problem Statement

На шаге результата мастера создания КП показывались «Срок изготовления» и «Другой вариант сохранения» (database / skip) — путь в производство из create. Нужно: create → только архив; производство — только из Archive (`MoveToProductionDialog`).

## Done

- UI: убраны срок изготовления и альтернативные режимы; кнопка «В архив».
- Schema: `saveOfferSchema` = `mode: "archive"` only.
- Archive → «В производство» не трогали.
