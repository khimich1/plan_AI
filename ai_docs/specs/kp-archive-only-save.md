# Spec: КП create — только архив

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-09-02  
**One-pager**: [../ideas/kp-archive-only-save.md](../ideas/kp-archive-only-save.md)  
**Related leftover D**: [kp-append-preview-and-fresh-breakdown.md](./kp-append-preview-and-fresh-breakdown.md)

## Objective

Сохранение из мастера создания КП — только в архив. Срок изготовления / перевод в производство остаются в Archive.

## Acceptance

- [x] Нет поля «Срок изготовления» и ссылки «Другой вариант сохранения» на Result create
- [x] Primary action «В архив» → `mode: "archive"`, пустой `executionTermsInput`
- [x] `MoveToProductionDialog` (Archive) без регрессий
- [x] FE tests на SaveOfferSection

## Out of scope

- Удаление backend `database` / `skip` modes (API совместимость)
- Изменение store `executionTermsInput` (можно оставить dormant)
