# Карта маршрутов ГСМ

**Status:** ✅ Implemented  
**Date:** 2026-08-11  
**Report:** [`../reports/2026-08-11-gsm-routes-map.md`](../reports/2026-08-11-gsm-routes-map.md)  
**Spec:** [`../../specs/gsm-routes-map.md`](../../specs/gsm-routes-map.md)

## Description

Личная offline-карта уникальных поездок A→B из пула ГСМ: линии по дорогам, цвет/фильтр по машине, поиск адреса → топ-3 ближайших маршрута.

## How It Works

1. Скрипт читает `routes_ab` из `ГСМ/пул_поездок.xlsx`.
2. Адреса геокодируются через Nominatim → `ГСМ/geo_cache/addresses.json`.
3. Пары A→B роутятся через публичный OSRM → `ГСМ/geo_cache/routes.geojson`.
4. Генерируется `ГСМ/карта_маршрутов.html` (Leaflet CDN) со встроенными данными.
5. В браузере: фильтр машин; поиск (known addresses + Photon) → point-to-segment → топ-3.

## Usage

```bash
.venv/bin/python scripts/build_gsm_routes_map.py \
  --xlsx "ГСМ/пул_поездок.xlsx" \
  --out "ГСМ/карта_маршрутов.html"
xdg-open "ГСМ/карта_маршрутов.html"
```

## Artifacts

- `scripts/build_gsm_routes_map.py`
- `ГСМ/geo_cache/addresses.json` (~177/244 geocoded)
- `ГСМ/geo_cache/routes.geojson` (434 features)
- `ГСМ/карта_маршрутов.html`
- `tests/test_build_gsm_routes_map.py` (35 tests)

## Known Issues

- Часть адресов без координат — правка кэша вручную.
- Зависимость от публичных Nominatim/OSRM/Photon.
