# КП: правка архива через конструктор

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-09-02  
**Spec**: [../specs/kp-archive-edit-in-constructor.md](../specs/kp-archive-edit-in-constructor.md)  
**Plan**: [../develop/plans/2026-09-02-kp-archive-edit-in-constructor.md](../develop/plans/2026-09-02-kp-archive-edit-in-constructor.md)  
**Related**: [kp-multi-nomenclature-append.md](./kp-multi-nomenclature-append.md), [kp-archive-only-save.md](./kp-archive-only-save.md), [kp-row-edit-delete-icons.md](./kp-row-edit-delete-icons.md)

## Problem Statement

How might we дать менеджеру править состав и итоги КП **со статусом «в архиве»** через уже готовый конструктор (допись / удаление / переименование строк), не дублируя редактор в drawer и не трогая КП в производстве?

## Recommended Direction

**Два входа из drawer архива + read-only сводка + переворот статус-гейта resume/save.**

| CTA | Landing |
|-----|---------|
| **(+ Добавить)** | Product type picker (`resume` + `start-append-cycle`) |
| **Редактировать** | Шаг 3 Результат (`resume` без append) |

Блок «Итоги» в drawer — только сводка (без скидки/рейса/целевой суммы).  
После сохранения КП остаётся **«в архиве»**, тот же `kp_id`.  
КП **«в работе»** (производство) — без правок. Старую кнопку «Добавить другое наименование» для «в работе» убрать.  
В шапке nav: **«Создать КП» → «Конструктор КП»**.

## Key Assumptions to Validate

- [x] Правки только для «в архиве»; производство не трогаем
- [x] Два landing’а нужнее одной кнопки «открыть конструктор»
- [x] Финансы в drawer больше не редактируем — только в конструкторе
- [x] Save из resume сохраняет статус «в архиве» и тот же номер
- [x] Старый resume для «в работе» больше не нужен (после archive-only-save)

## MVP Scope

**In:** две CTA в drawer «в архиве»; read-only Итоги; гейт hydrate/save → «в архиве»; nav rename  
**Out:** инлайн-редактор состава в drawer; правки в производстве; версионирование PDF; API-hardening discount PATCH (UI-only)

## Not Doing (and Why)

- Редактирование КП в производстве — бизнес-риск после запуска
- Финансы в drawer — дубль шага 3, путаница
- Одна кнопка вместо двух — менеджер хочет разный старт (допись vs правка строк)
- Rename/delete прямо в таблице архива — уже есть на Result

## Open Questions

_Нет — locked в ideation 2026-09-02._
