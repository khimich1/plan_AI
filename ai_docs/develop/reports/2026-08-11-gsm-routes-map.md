# Report: Карта маршрутов ГСМ (Leaflet + OSRM)

**Date:** 2026-08-11  
**Orchestration:** `orch-2026-08-11-18-36-gsm-routes-map`  
**Status:** ✅ Completed  
**Spec:** [`ai_docs/specs/gsm-routes-map.md`](../../specs/gsm-routes-map.md)  
**Plan:** [`ai_docs/develop/plans/2026-08-11-gsm-routes-map.md`](../plans/2026-08-11-gsm-routes-map.md)  
**Idea:** [`ai_docs/ideas/gsm-routes-map.md`](../../ideas/gsm-routes-map.md)

## Summary

Из `ГСМ/пул_поездок.xlsx` (`routes_ab`) собран локальный HTML `ГСМ/карта_маршрутов.html`: уникальные маршруты A→B линиями по дорогам (OSRM), цвет и фильтр по машине, поиск адреса → маркер → топ-3 ближайших трека без жёсткого порога расстояния. Геокод и треки кэшируются на диске; повторный прогон может работать офлайн по кэшу.

## What Was Built

- **`scripts/build_gsm_routes_map.py`** — один пайплайн: Excel → Nominatim → OSRM → GeoJSON → HTML (Leaflet CDN, шаблон внутри скрипта).
- **`ГСМ/geo_cache/addresses.json`** — кэш геокода: **177 / 244** адресов с координатами; негативный кэш + `--force-geocode` / retry.
- **`ГСМ/geo_cache/routes.geojson`** — **434** LineString features (OSRM `overview=simplified` + Douglas-Peucker).
- **`ГСМ/карта_маршрутов.html`** — Leaflet: цвета по 4 машинам, фильтр «все / одна», попапы, поиск (встроенные известные адреса + Photon free-text).
- **`tests/test_build_gsm_routes_map.py`** — **35** тестов (чтение/дедуп, геокод с моками, OSRM, nearest point-to-segment, офлайн).
- **`requirements.txt`** — добавлены `xlrd`, `requests` (nearest в браузере; shapely не обязателен).

## Completed Tasks

1. ✅ **T1:** Read `routes_ab` → route list + unique addresses  
   - Files: `scripts/build_gsm_routes_map.py`, `tests/test_build_gsm_routes_map.py`

2. ✅ **T2:** Geocode with Nominatim + cache  
   - Files: скрипт, тесты, `ГСМ/geo_cache/addresses.json`  
   - Result: ~177/244 geocoded; упрощённые query; negative cache

3. ✅ **T3:** OSRM routes → `routes.geojson` cache  
   - Files: скрипт, тесты, `ГСМ/geo_cache/routes.geojson`  
   - Result: 434 features, simplified geometry

4. ✅ **T4:** Leaflet HTML — colors / filter / popup  
   - Files: шаблон в скрипте → `ГСМ/карта_маршрутов.html`

5. ✅ **T5:** Address search → top-3 nearest routes  
   - Files: HTML-логика + unit nearest; known addresses + Photon; point-to-segment; без порога

6. ✅ **T6:** requirements + final run + green tests  
   - Files: `requirements.txt` (`xlrd`, `requests`)  
   - Tests: 35 passing

## Technical Decisions

- **Nominatim** с упрощёнными запросами; negative cache и `--force-geocode` / retry для повторных попыток.
- **Публичный OSRM** (`overview=simplified`) + **Douglas-Peucker** для размера GeoJSON/HTML.
- **Поиск в браузере:** встроенный список известных адресов + Photon free-text; расстояние — point-to-segment до линии; **топ-3 без жёсткого порога**.
- **Один скрипт + дисковые кэши**; backend/frontend «Шишов» не трогали.
- Зависимости: `requests` (+ `xlrd` для экосистемы ГСМ); nearest на клиенте, без обязательного shapely.

## Metrics

| Metric | Value |
|--------|-------|
| Addresses in cache | 244 |
| Geocoded (coords) | ~177 |
| Route features | 434 |
| Tests | 35 passing |
| Tasks | 6 / 6 |

## How to run

```bash
.venv/bin/python scripts/build_gsm_routes_map.py \
  --xlsx "ГСМ/пул_поездок.xlsx" \
  --out "ГСМ/карта_маршрутов.html"

.venv/bin/python -m pytest tests/test_build_gsm_routes_map.py -q

xdg-open "ГСМ/карта_маршрутов.html"
```

## Known Issues / follow-ups

- ~67 адресов без координат — правка вручную в `addresses.json` или `--force-geocode` после улучшения query.
- Публичные Nominatim/OSRM/Photon — квоты и доступность; для продакшена личный OSRM опционален позже.
- Не покрыто: `rounds_aba`, мобильная геолокация, модуль в web «Шишов».

## Related Documentation

- Plan: [`2026-08-11-gsm-routes-map.md`](../plans/2026-08-11-gsm-routes-map.md)
- Spec: [`gsm-routes-map.md`](../../specs/gsm-routes-map.md)
- Idea: [`gsm-routes-map.md`](../../ideas/gsm-routes-map.md)
- Upstream pool: [`gsm-trip-pool.md`](../../ideas/gsm-trip-pool.md)

## Next Steps

- Точечно добить геокод «не распознано» в `addresses.json`.
- Ручной smoke: 5–10 известных заправок → топ-3 совпадает с ожиданием.
- При росте объёма — локальный OSRM вместо demo-сервера.
