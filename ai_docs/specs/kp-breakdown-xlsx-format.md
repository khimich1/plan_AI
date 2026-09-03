# Spec: breakdown.xlsx — эталонный layout

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-09-02  
**One-pager**: [../ideas/kp-breakdown-xlsx-format.md](../ideas/kp-breakdown-xlsx-format.md)  
**Related leftover E**: [kp-append-preview-and-fresh-breakdown.md](./kp-append-preview-and-fresh-breakdown.md)

## Objective

Экспорт «Детальная разбивка цен» читается как эталон: Компонент | Расчёт | Сумма, полные подписи, формулы, `N NNN,NN руб`, пустая строка между блоками марок.

## Acceptance

- [x] Headers: Компонент, Расчёт, Сумма
- [x] Product header + component rows + ИТОГО / Округлено / За N плит
- [x] Empty row between blocks
- [x] Full labels preserved in cells; column widths readable (not default ~13)
- [x] Sums end with ` руб`

## Notes

Данные строк уже строит `viz_modules/procurement/breakdown.py`; правка — слой Excel (`core/commercial_offer.save_breakdown_to_excel`).
