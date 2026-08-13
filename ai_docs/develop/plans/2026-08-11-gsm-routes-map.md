# Plan: Карта маршрутов ГСМ (личная, по дорогам)

**Created:** 2026-08-11  
**Orchestration:** `orch-2026-08-11-18-36-gsm-routes-map`  
**Status:** ✅ Implemented (2026-08-11)  
**Spec:** [`ai_docs/specs/gsm-routes-map.md`](../../specs/gsm-routes-map.md) (✅ agreed in session)  
**Idea:** [`ai_docs/ideas/gsm-routes-map.md`](../../ideas/gsm-routes-map.md)  
**Report:** [`ai_docs/develop/reports/2026-08-11-gsm-routes-map.md`](../reports/2026-08-11-gsm-routes-map.md)

## Progress (updated by orchestrator)

- ✅ T1: Read routes_ab → route list + unique addresses (Done)
- ✅ T2: Geocode with Nominatim + cache (Done)
- ✅ T3: OSRM routes → routes.geojson cache (Done)
- ✅ T4: Leaflet HTML map with colors/filter/popup (Done)
- ✅ T5: Address search → top-3 nearest routes (Done)
- ✅ T6: requirements + final run + green tests (Done)

## Goal

Из `ГСМ/пул_поездок.xlsx` (`routes_ab`) собрать `ГСМ/карта_маршрутов.html`: все уникальные маршруты A→B линиями **по дорогам**, цвет по машине, фильтр «все/одна», ввод адреса → маркер → **топ-3 ближайших маршрута** с расстоянием.

## Current state

| Компонент | Сейчас |
|-----------|--------|
| Пул поездок | ✅ `scripts/build_gsm_trip_pool.py` → `ГСМ/пул_поездок.xlsx` |
| Карта маршрутов | ✅ `scripts/build_gsm_routes_map.py` → `ГСМ/карта_маршрутов.html` |
| Геокэш | ✅ `ГСМ/geo_cache/` (~177/244 адресов, 434 route features) |
| Зависимости | ✅ `requests`, `xlrd` в `requirements.txt` |
| Тесты | ✅ `tests/test_build_gsm_routes_map.py` (35) |

## Architecture decisions

1. **Один скрипт** `scripts/build_gsm_routes_map.py` — чтение Excel → геокод → OSRM → GeoJSON → HTML (шаблон внутри скрипта, Leaflet по CDN).
2. **Кэши** в `ГСМ/geo_cache/`: `addresses.json` (адрес→координаты, ключ = `normalize_address`) и `routes.geojson` (пара координат→LineString). Повторный прогон — без сети, где возможно; флаг `--offline`.
3. **Роутинг:** публичный OSRM demo; запросы только для пар, которых нет в кэше; пропуски (нет координат/ошибка) — в отчёт.
4. **Поиск адреса:** в браузере — геокод через Nominatim (JS fetch) → точка → nearest-расстояние до каждой линии (turf.js или простая point-to-segment на GeoJSON координатах) → топ-3 (машина, A, B, км).
5. **Рисуем все** уникальные A→B (~600), Leaflet справляется; цвет — фиксированная палитра на 4 машины.

## Phases & Tasks

### Phase 1 — данные и геокод
- [x] T1. ✅ Чтение `routes_ab` → список маршрутов (машина, A, B, км, частота) + уникальные адреса
  - Acceptance: на фикстуре получаем ожидаемое число строк/дедуп
  - Verify: `pytest -q`
  - Files: `scripts/build_gsm_routes_map.py`, `tests/test_build_gsm_routes_map.py`
- [x] T2. ✅ Геокод адресов (Nominatim) с кэшем и отчётом «не распознано»
  - Acceptance: `addresses.json` заполнен; `--offline` не ходит в сеть
  - Verify: тест с моком HTTP
  - Files: скрипт, тесты, `ГСМ/geo_cache/addresses.json`

### Phase 2 — треки
- [x] T3. ✅ OSRM route A→B → `routes.geojson` (кэш, пропуски логируются)
  - Acceptance: LineString для доступных пар; повтор без сети
  - Verify: тест мок OSRM
  - Files: скрипт, тесты, `ГСМ/geo_cache/routes.geojson`

### Phase 3 — карта
- [x] T4. ✅ HTML Leaflet: слои/цвета по машине, фильтр, попап (A, B, км, частота)
  - Acceptance: локальное открытие показывает карту и линии
  - Verify: ручной открыть в браузере
  - Files: скрипт (шаблон), `ГСМ/карта_маршрутов.html`
- [x] T5. ✅ Поиск адреса → маркер + топ-3 ближайших маршрута (расстояние км)
  - Acceptance: работает на известном адресе заправки
  - Verify: unit nearest-функции + ручная проверка
  - Files: скрипт, тесты, HTML-логика

### Phase 4 — завершение
- [x] T6. ✅ `requirements.txt` (+`xlrd`, `requests`; nearest в браузере — shapely не нужен), финальный прогон, тесты зелёные (35)
  - Acceptance: один запуск собирает карту; `pytest` зелёный
  - Verify: `pytest` + запуск скрипта
  - Files: `requirements.txt`, скрипт

## Verification checkpoints

- После Phase 1: адреса геокодированы, список «не распознано» пуст или осмыслен
- После Phase 2: треки покрывают большинство пар; кэш работает офлайн
- После Phase 3: карта открывается, поиск возвращает топ-3

## Out of scope

Мобильная геолокация, модуль в «Шишов», привязка к дню/ПЛ, ручная правка `Роману.xlsx` из карты.
