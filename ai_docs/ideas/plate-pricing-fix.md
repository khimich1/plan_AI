# Исправление расчёта стоимости плит (trim)

> Краткий one-pager ideation. **Полная спека:** [`../specs/plate-pricing-trim-bugs.md`](../specs/plate-pricing-trim-bugs.md)

## Problem Statement

Как сделать расчёт КП и производственной сметы детерминированным и совпадающим с ручным расчётом менеджера, без потери продольных резов и двойного учёта отходов?

## Recommended Direction

Системный аудит `_calc_trim_components` + регрессионная матрица + скрипт сверки с эталонами менеджера. Явное правило: **отход полосы — один раз на primary**, secondary только резы и остаток по длине.

## Key Assumptions to Validate

- [ ] Отход полосы только на primary — подтверждено менеджером
- [ ] 10,8 м = 1 продольный рез + factory strip waste
- [ ] Cross-cascade: отход не на secondary
- [ ] JSON эталоны 4 кейсов + новые заказы

## MVP Scope

**In:** 4 кейса → тесты → Fix-1, Fix-2 в trim.py → reconcile script  
**Out:** рефакторинг оптимизатора, пересчёт архива, core/pricing service

## Not Doing

- Полный граф раскроя — дорого для v1
- Автоисправление старых КП
- Изменение ILP без трассировки waste=240

## Open Questions

- Snapshot плана для кейса 4 (ПБ 56,3 + ПБ 46,4)
- Другие ширины с пропавшим продольным резом?
