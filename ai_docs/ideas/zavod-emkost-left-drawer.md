# Ёмкость завода: left drawer по кнопке

Дата: 2026-08-26. Статус: направление зафиксировано; спека готова.
Спека: [`ai_docs/specs/zavod-emkost-left-drawer.md`](../specs/zavod-emkost-left-drawer.md).
Родитель: [`zavod-emkost-vizual-gate.md`](zavod-emkost-vizual-gate.md) /  
[`../specs/zavod-emkost-vizual-gate.md`](../specs/zavod-emkost-vizual-gate.md).

## Problem Statement

Как убрать календарь завода из одного блока со строкой срока и открывать
его слева от края экрана только по нажатию «Ёмкость» — одинаково в
«В производство» и в графике поставок?

## Recommended Direction

Модалка снова узкая: оценка + срок + при red короткий hint + кнопки.
Кнопка **«Ёмкость»** открывает **drawer от левого края viewport** с
текущим `FactoryCapacityPanel` (календарь + нужно/свободно/Δ).

Закрытие drawer: ✕, Esc, **клик по backdrop**.

При red drawer **не** автооткрывается — hint остаётся в модалке; гейт
(disable submit / backend 4xx) без изменений.

Оба entry point: `MoveToProductionDialog`, `DeliveryScheduleDialog` /
Editor — один паттерн.

## Key Assumptions to Validate

- [ ] Кнопка «Ёмкость» без бейджа статуса находится достаточно быстро
- [ ] Hint в модалке достаточен, когда drawer закрыт
- [ ] z-index drawer выше backdrop модалки, focus/Esc не ломают закрытие модалки

## MVP Scope

**In:** кнопка «Ёмкость»; left-edge drawer; backdrop click + ✕ + Esc;
hint в модалке при red; оба entry points; убрать inline-панель из grid формы.

**Out:** авто-open при red; бейдж на кнопке; смена API/алгоритма гейта;
прилипание drawer к модалке (отвергнуто — край экрана).

## Not Doing (and Why)

- **Inline-календарь в форме** — отвергнуто скрином / UX-сессией
- **Авто-выезд при red** — выбран hint в модалке + ручная кнопка
- **Drawer к краю модалки** — выбран край экрана
- **Отдельный UX для графика** — оба места одинаково

## Open Questions

_Нет блокирующих._
