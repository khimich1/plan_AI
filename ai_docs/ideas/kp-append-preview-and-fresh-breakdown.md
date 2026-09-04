# КП: видимый состав при дописи + свежий breakdown

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-09-02  
**Spec**: [ai_docs/specs/kp-append-preview-and-fresh-breakdown.md](../specs/kp-append-preview-and-fresh-breakdown.md)  
**Plan**: [ai_docs/develop/plans/2026-09-02-kp-append-preview-and-fresh-breakdown.md](../develop/plans/2026-09-02-kp-append-preview-and-fresh-breakdown.md)

## Problem Statement

How might we сделать так, чтобы при дописи плит менеджер всегда видел уже забитый состав, а детальная разбивка/цена после правок не отдавала старую версию?

## Recommended Direction

Приоритет B+C: (1) предпросмотр на шаге плит = sealed ∪ текущий заход, «Добавить к списку» не прячет уже добавленное; (2) после изменения состава/конфигурации — инвалидация и пересчёт breakdown до скачивания.  
Попутно A: Drawer (i) слева + нумерация строк.

## Key Assumptions to Validate

- [x] Пустой предпросмотр — фильтр unsealed-only, draft цел
- [x] Старый breakdown — не обновляется после line edit / recalculate
- [x] Полный список на шаге 1 не мешает сверке текущего батча

## MVP Scope

**In:** B (полный состав в preview), C (свежий breakdown), A (Drawer left + №)  
**Out:** попап карандаша (D+E реализованы отдельно: [kp-archive-only-save](./kp-archive-only-save.md), [kp-breakdown-xlsx-format](./kp-breakdown-xlsx-format.md))

## Not Doing (and Why)

- ~~Archive-only save (D)~~ → [kp-archive-only-save](./kp-archive-only-save.md) **IMPLEMENT ✅**
- ~~Эталонный Excel разбивки (E)~~ → [kp-breakdown-xlsx-format](./kp-breakdown-xlsx-format.md) **IMPLEMENT ✅**
- Перенос (i) на Result — не просили

## Open Questions (LOCKED defaults for this implementation)

- Sealed + текущий заход: **одна таблица предпросмотра** со всеми строками типа (проще); при необходимости визуально отличить текущий батч — лёгкий бейдж, не два разных экрана.
- C: **авто-инвалидация/пересчёт** breakdown при изменении состава/конфига; не отдавать stale файл при скачивании (пересобрать или блокировать до готовности — предпочти пересборку).
