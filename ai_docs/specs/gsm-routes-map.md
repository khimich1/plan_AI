# Spec: Карта маршрутов ГСМ (личная, по дорогам)

> Связано: `ai_docs/ideas/gsm-routes-map.md`, `ai_docs/ideas/gsm-trip-pool.md`  
> Вход: `ГСМ/пул_поездок.xlsx`, лист `routes_ab`

## Objective

Из пула уникальных поездок собрать один локальный HTML с картой:

- все уникальные маршруты A→B линиями **по дорогам**;
- цвет по машине, фильтр «все / одна»;
- ввод адреса (например заправки) → маркер → **топ-3 ближайших маршрута** с расстоянием до линии (без жёсткого порога).

Пользователь: только Роман, ПК, браузер.  
Успех: ввёл адрес → видно, к какому маршруту точка «прилипает».

## Tech Stack

- Python 3.12 (`.venv`): openpyxl, requests, shapely
- Выход: GeoJSON + Leaflet 1.9 (CDN), тайлы OSM
- Роутинг: **публичный OSRM** (`router.project-osrm.org`), кэш треков
- Геокод: **Nominatim** (User-Agent, троттлинг, кэш с ручной правкой)
- Backend/frontend проекта не трогаем

## Commands

```bash
# Сборка данных и HTML
.venv/bin/python scripts/build_gsm_routes_map.py \
  --xlsx "ГСМ/пул_поездок.xlsx" \
  --out "ГСМ/карта_маршрутов.html"

# Тесты
.venv/bin/python -m pytest tests/test_build_gsm_routes_map.py -q

# Открыть
xdg-open "ГСМ/карта_маршрутов.html"
```

## Project Structure

```
scripts/build_gsm_routes_map.py     → сборка GeoJSON + HTML
ГСМ/пул_поездок.xlsx                → вход (routes_ab)
ГСМ/geo_cache/addresses.json        → геокод адресов (можно править руками)
ГСМ/geo_cache/routes.geojson        → треки по дорогам (кэш)
ГСМ/карта_маршрутов.html            → результат
tests/test_build_gsm_routes_map.py  → unit-тесты чистых функций
```

## Code Style

Как в `scripts/build_gsm_trip_pool.py`: argparse, dataclass, чистые функции, без ORM.  
Ключ адреса — `normalize_address()` из `build_gsm_trip_pool.py`.

## Testing Strategy

- pytest: нормализация/дедуп A→B; сборка GeoJSON; функция «точка → ближайший сегмент»; прогон без сети (кэш-фикстуры).
- Ручная проверка: 5–10 известных адресов заправок.

## Boundaries

- **Always:** кэш геокода/треков; User-Agent и троттлинг для Nominatim; тесты перед готовностью.
- **Ask first:** менять `Роману.xlsx`; тяжёлый JS-фреймворк вместо Leaflet; сервер/авторизация.
- **Never:** ключи/секреты в HTML; геокод без кэша; правки backend/frontend проекта.

## Success Criteria

- [x] В HTML отрисованы линии по дорогам для подавляющего большинства уникальных A→B (не прямые) — 434 features в `routes.geojson`
- [x] 4 цвета по машинам + фильтр «все / одна» работают
- [x] Ввод адреса → маркер → список **топ-3** маршрутов с расстоянием (км)
- [x] Повторный прогон не ходит в сеть по кэшированным адресам/трекам
- [x] Тесты зелёные: `.venv/bin/python -m pytest tests/test_build_gsm_routes_map.py -q` (35)

## Open Questions

Закрыты решениями сессии:
- Порог: не фиксировать, показывать топ-3.
- Роутер: публичный OSRM + кэш; локальный OSRM — опция позже.
- Рисуем все уникальные A→B.
