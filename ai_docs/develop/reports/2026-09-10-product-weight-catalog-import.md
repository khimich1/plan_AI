# Report: импорт справочника весов ФБС / ЛС / ЛМ

**Date:** 2026-09-10  
**Spec:** [`ai_docs/specs/kp-delivery-fbs-steps-marches.md`](../../specs/kp-delivery-fbs-steps-marches.md)  
**Plan:** [`ai_docs/develop/plans/2026-09-10-kp-delivery-fbs-steps-marches.md`](../plans/2026-09-10-kp-delivery-fbs-steps-marches.md)  
**Idea:** [`ai_docs/ideas/kp-delivery-fbs-steps-marches.md`](../../ideas/kp-delivery-fbs-steps-marches.md)

## Summary

В `plita.db` загружен справочник `product_weight_catalog` из трёх выгрузок 1С. Перед записью сделан бэкап `plita.db.bak-before-weight-catalog-20260910-132525`. Повторный прогон — 0 inserted, все строки updated (идемпотентность). `plita.db` в git не коммитить.

Итог в таблице: **14 ФБС + 18 ЛС + 9 ЛМ = 41**. Оценка «7 ЛМ» в Task 2 плана — это покрытие прайса 7/7, не число строк файла. В файле две дополнительные марки (`ЛМ 2,8`, `ЛМ 2,9`) с валидной плотностью.

## Backup

```
plita.db.bak-before-weight-catalog-20260910-132525
```

## Import commands

```bash
source venv/bin/activate
python scripts/import_product_weights_from_xlsx.py "банк знаний/Новая папка/Блоки.xlsx" --type fbs
python scripts/import_product_weights_from_xlsx.py "банк знаний/Новая папка/ЛС.xlsx" --type steps
python scripts/import_product_weights_from_xlsx.py "банк знаний/Новая папка/ЛМ и ЛП.xlsx" --type marches
```

## File reports

| Файл | Type | Imported | Quarantine | Skipped | First run | Re-run |
|------|------|----------|------------|---------|-----------|--------|
| Блоки.xlsx | fbs | 14 | 0 | 2 КБ | +14 / 0 upd | 0 / 14 upd |
| ЛС.xlsx | steps | 18 | 1 | 0 | +18 / 0 upd | 0 / 18 upd |
| ЛМ и ЛП.xlsx | marches | 9 | 3 | 2 ЛП | +9 / 0 upd | 0 / 9 upd |

### Quarantine (плотность вне 1500–3500 кг/м³)

- ЛС.xlsx: `Лестничные ступени ЛС-15-1 закл` (ожидаемый артефакт, плотность 251)
- ЛМ и ЛП.xlsx: `ЛП 28-15-5ш` (805 кг/м³)
- ЛМ и ЛП.xlsx: `1ЛП 28-15-5ш-1` (793 кг/м³)
- ЛМ и ЛП.xlsx: `ЛМ 30.13.15-5ш` (354 кг/м³)

Артефакты идеи «Сваи.xlsx / Перемычки.xlsx» в этот импорт не входят (вне scope D1).

### Skipped (вне scope)

- Кросс-Блок 50 (КБ-50), Кросс-Блок 100 (КБ-100)
- Лестничные площадки 2ЛП25.12в-4-к, 2ЛП25.12-4-к

## Catalog contents (plita.db)

`product_type` counts: fbs=14, steps=18, marches=9.

Дополнительные ЛМ относительно прайсового покрытия 7/7: `ЛМ 2,8` (975 кг), `ЛМ 2,9` (975 кг).

## Manual checklist (Task 10)

| Check | Status |
|-------|--------|
| Mono-ФБС КП в визарде: поле рейса + плашка рейсов | покрыто vitest; **в браузере не прогонялось** (нет browser/MCP tools в сессии) |
| Mixed плиты+ФБС: два котла, подписи «Рейс плит»/«Рейс свай» без регрессии | vitest + pytest mixed/export |
| Архивное старое КП без флага: карточка без плашки котла | vitest drawer; pytest D10 |
| Новое КП в архиве: PDF/XLSX со строкой котла | pytest export; **живые PDF/XLSX в UI не открывались** |

`npm run typecheck` зелёный. `npm run build` упал на **несвязанном** `PromisePeriodCalendar.tsx` (`knob` possibly undefined) — файл factory-capacity, в этой фиче не трогался.

## Notes

- Embed-XLSX (`xlsx_delivery_in_unit`) не менялся (A10 follow-up).
- Живую БД и бэкап в коммит не класть.
