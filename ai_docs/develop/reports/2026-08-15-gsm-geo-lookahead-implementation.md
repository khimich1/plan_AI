# Report: GSM lookahead-генератор с географией

**Date:** 2026-08-15  
**Orchestration:** `orch-2026-08-15-gsm-geo-lookahead`  
**Status:** ✅ Completed  
**Spec:** [`../../specs/gsm-geo-lookahead-generator.md`](../../specs/gsm-geo-lookahead-generator.md)  
**Plan:** [`../plans/2026-08-15-gsm-geo-lookahead.md`](../plans/2026-08-15-gsm-geo-lookahead.md)  
**Acceptance:** [`2026-08-15-gsm-geo-lookahead-acceptance.md`](2026-08-15-gsm-geo-lookahead-acceptance.md)

## Summary

Генератор путевых листов v2: round-trip (2 плеча), lookahead на плотных заправках, мягкая гео-сортировка по направлению к следующей АЗС, частичная генерация вместо 422. Ночёвки сознательно не делались (фаза 2, R1).

Сессия прервалась перезагрузкой после T6/T8; resume с T7 (код T7 ещё не был на диске).

## What was built

- Data: геокодинг 27 станций (76/76 с координатами), `typical_station_ids` у 450/610 маршрутов (Palisade × КТК Магистральная id=1 — 116).
- Core: `core/gsm/geo.py`; generator — 2 плеча, lookahead, geo priority 1/2/3, `ProblematicDay`.
- Service/API: `max_daily_km` (default 700), `station_coords` в солвер, `POST /gsm/waybills/generate` → 200 + `problematic_days`.
- UI: коды `manual_intervention` / `balance_route`, красные дни, сводка частичной генерации.

## Completed tasks

1. ✅ T1 geocode stations  
2. ✅ T2 `geo.py` + тесты  
3. ✅ T3 link_route_stations  
4. ✅ T4 round-trip  
5. ✅ T5 lookahead  
6. ✅ T6 direction sort  
7. ✅ T7 partial generation  
8. ✅ T8 max_daily_km + coords  
9. ✅ T9 API problematic_days  
10. ✅ T10 warning badges  
11. ✅ T11 GsmPeriodView  
12. ✅ T12 Palisade май acceptance  
13. ✅ T13 docs  

## Metrics

- Palisade май 2026: 13 дней, **2** manual (08.05, 21.05), без 422.
- GSM pytest: все целевые сьюиты зелёные; frontend GSM **54 passed**.
- Полный `pytest tests/`: 1906 passed, 9 failed вне ГСМ (OCR/KP/commercial).

## Out of scope (фаза 2)

Ночёвки (`overnight_trip`, `return_leg`), составные дни из 3+ плеч, калибровка `max_daily_km` по машине.

## Known notes

- Nominatim для трассовых АЗС грубый (одна точка «Холмогоры») — для направления достаточно.
- `LibraryRoute.point_a/b` в runtime не заполняются (нет геокодинга из backend); направление опирается на `station_coords`.
- Коммитов по R10 не создавалось.
