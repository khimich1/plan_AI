# Report: Рейсы свай и мостовых свай в КП

**Date:** 2026-09-01  
**Spec:** [`ai_docs/specs/pile-bridge-trip-calculation.md`](../../specs/pile-bridge-trip-calculation.md)  
**Plan:** [`ai_docs/develop/plans/2026-09-01-pile-bridge-trip-calculation.md`](../plans/2026-09-01-pile-bridge-trip-calculation.md)  
**Idea:** [`ai_docs/ideas/pile-bridge-trip-calculation.md`](../../ideas/pile-bridge-trip-calculation.md)

## Summary

КП считает доставку свай и мостовых отдельно от плит: гибрид `floor(qty/pcs)` + остатки / 19 800 кг, ручное N для марок без нормы (C18 и «-» в Excel). Плиты по-прежнему 18 600 кг. Два тарифа (`logistics_cost` / `pile_logistics_cost`), две строки в PDF/XLSX. Код отгрузок не менялся.

## AC checklist

| AC | Status | Tests / verify |
|----|--------|----------------|
| Импорт 44 марки; С160 pcs NULL; С140.40 = 5650 / 3 | ✅ | `tests/test_pile_catalog_import.py` |
| `C14-40T4` / `С14-40Т4` → `С140.40` | ✅ | тот же файл + `resolve_catalog_for_mark` |
| Тендер без C18 = 39 full + 3 remainder = 42 | ✅ | `tests/test_pile_trip_pricing.py` |
| C18 без N: ready=False, pile delivery 0; N=k → 42+k | ✅ | trip + logistics + PT-901 flow |
| Плиты 18600 без регрессии | ✅ | `tests/test_commercial_logistics_cost.py` |
| Mixed: две доставки в totals и XLSX | ✅ | logistics + `test_commercial_export_mixed.py` |
| Draft API: overrides, pending, trips | ✅ | PT-901 `test_commercial_bridge_pile_flow.py` |
| Save persist tariff + JSON | ✅ | `test_kp_persistence_piles.py`, PT-901 |
| PDF/XLSX: «Доставка плит» / «Доставка свай»; pending без строки свай | ✅ | export mixed |
| Result step: рейс без плит; mixed два поля; вопрос N; итог без 39+3 | ✅ | `CalculationResultStep.test.tsx` |
| Архив details + PATCH без авто-PDF | ✅ | archive service/endpoints + PT-901 |
| Нет diff `shipment*.py` | ✅ | `git diff --name-only` |

## Formula (locked fixture)

C14-40T4 52 → pcs 3 → 17 full; C9 19 → 2; C10 19 → 3; C13 19 → 3; C11 19 → 3; C15 45 → 11. Full **39**. Remainder 46 950 кг / 19 800 → **3**. Total **42**. C18-40T8 × 49 без каталога → pending.

Constants: plates `CARGO_DELIVERY_TRUCK_CAPACITY_KG = 18600`; piles `PILE_REMAINDER_TRUCK_CAPACITY_KG = 19800` in `core/pile_trip_pricing.py` only.

## Catalog import (ops)

```bash
python scripts/import_pile_catalog.py --xlsx "банк знаний/сваи вес и объем.xlsx" --sheet Лист1
```

Без `--sheet`: лист «Вес и объем», иначе «Лист1». «-» / пусто → `pcs_per_20t` NULL. Каталог в `plita.db`, не в `pb.db`.

## Key files

**Domain:** `core/pile_catalog.py`, `core/pile_trip_pricing.py`, `core/commercial_pricing.py`  
**Schema / persist:** `core/kp_db_schema.py`, `core/kp_persistence_service.py`, `core/kp/offers_write.py`  
**API:** `app/schemas/commercial.py`, `app/schemas/archive.py`, draft/calculate/archive services  
**Export:** `core/commercial_offer.py`, `core/commercial_offer_xlsx.py`  
**Frontend:** `CalculationResultStep.tsx`, `OfferDetailsDrawer.tsx`, commercial/archive API types

## Verification run

```
pytest tests/test_pile_catalog_import.py tests/test_pile_trip_pricing.py \
  tests/test_commercial_logistics_cost.py tests/test_commercial_calculation_service.py \
  tests/test_commercial_export_mixed.py tests/test_archive_service.py \
  tests/test_archive_endpoints.py tests/test_kp_persistence_piles.py \
  tests/test_commercial_bridge_pile_flow.py tests/test_kp_db_update_logistics_cost.py \
  tests/test_kp_db_schema_boundary.py -q
# → 192 passed

cd frontend && npm run typecheck && npm run test -- --run
# → tsc ok; 110 files / 728 tests passed
```

## Remaining / next

- На живой `plita.db` нужен CLI-импорт каталога (тесты сидят в изолированных БД).
- Архивный PATCH скидки по-прежнему собирает только плиты (`update_kp_discount`); смешанное КП лучше править рейсы через PATCH logistics, не через скидку.
- Браузерный smoke на `./run+logs.sh` после импорта каталога.

## Status

**Implemented** (2026-09-01). Plan tasks PT-001…PT-902 complete. Not committed.
