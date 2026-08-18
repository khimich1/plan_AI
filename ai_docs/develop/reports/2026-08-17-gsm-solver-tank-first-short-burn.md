# Report: GSM солвер — бак важнее группы АЗС + короткий дожиг

**Date:** 2026-08-17  
**Status:** ✅ Completed  
**Spec:** [`../../specs/gsm-solver-tank-first-short-burn.md`](../../specs/gsm-solver-tank-first-short-burn.md)  
**Plan:** [`../plans/2026-08-17-gsm-solver-tank-first-short-burn.md`](../plans/2026-08-17-gsm-solver-tank-first-short-burn.md)  
**Idea:** [`../../ideas/gsm-solver-tank-first-short-burn.md`](../../ideas/gsm-solver-tank-first-short-burn.md)

## Summary

Доработка генератора v2 в `core/gsm/generator.py`: пул дожига — все маршруты в `max_daily_km` (не сетка 150–250 км); при попадании в headroom — минимальный достаточный km; lookahead, если группа АЗС короткая, берёт минимальный достаточный маршрут из всей библиотеки машины + `balance_route`. API/UI/схема не менялись. `manual_intervention` остаётся предохранителем.

Порядок: сначала B (короткий дожиг), затем A (полная библиотека), иначе 06.05 уехал бы на Ковров до того, как четверг научился жечь ~10 л.

Приёмка Palisade май 2026 на **копии** `plita.db` (`/tmp/plita-gsm-accept.db`), без записи в рабочую БД.

## What changed

### B — короткий дожиг

- `_ordered_burn_routes`: фильтр `BURN_KM_MIN/MAX` удалён; константы удалены (больше нигде не использовались).
- `_plan_burn_in`: среди `reaching` сортировка `_daily_km` asc, frequency desc, `route_id`. Пока headroom не достигнут — по-прежнему max безопасный burn.

### A — fallback на всю библиотеку

- `_select_anchor_route_lookahead`: если `_pick_min_sufficient(group)` вернул `None`, повтор по всем `routes` (уже в капе). При fallback — warning `balance_route`. Группу не расширяем до проверки дожига.

## Palisade May (SC-G6')

| Параметр | Значение |
|---|---|
| Машина | Hyundai Palisade (`vehicle_id=1`) |
| Период | 2026-05-04 … 2026-05-31 |
| Старт | 28 л / одометр 128327 |
| `max_daily_km` | 700 (дефолт) |
| `force` | true (на копии) |
| БД | `/tmp/plita-gsm-accept.db` (копия; mtime рабочего `plita.db` не менялся) |

| Критерий | Факт | Статус |
|---|---|---|
| Без 422 | сервис вернул результат | PASS |
| `problematic_days == []` | **[]**, `manual_days=0` | PASS |
| Дней создано | **17** (было 13 в v2 — появились дни дожига) | — |
| 06.05 typical, не Ковров | Ростов 135×2 = 270 км, без `balance_route` | PASS |
| 07.05 короткий дожиг | Волгореченск 45×2 = 90 км, `fuel_end=7.29 ≥ 0` | PASS |
| 08.05 не manual | Ярославль 95×2, без красного | PASS |
| 20.05 из полной библиотеки | Переславль 200×2 + `balance_route` | PASS |
| 21.05 не manual | Якимово 175×2 + `balance_route` | PASS |

### `problematic_days`

Пусто. Красных дней нет; двухякорный lookahead не понадобился.

Жёлтый `balance_route` (норма удлинения): 04.05, 05.05, 14.05, 15.05, 20.05, 21.05.

### Маршруты (кратко)

- 04–05.05: Ковров 205×2 — lookahead (как v2).
- 06.05: Ростов 135×2 — typical, **не** удлинён (B: четверг дожигает).
- 07.05: Волгореченск 45×2 — короткий дожиг (~13 л).
- 08.05: Ярославль 95×2 — **не** manual (в v2 был красный).
- 20.05: Переславль 200×2 + `balance_route` — A, полная библиотека.
- 21.05: Якимово 175×2 + `balance_route` — **не** manual (в v2 был красный).
- 22.05: Ярославль 95×2.

## Tests

| Команда | Результат |
|---|---|
| `venv/bin/pytest tests/test_gsm_generator.py -q` | **34 passed** |
| `venv/bin/pytest tests/test_gsm_generator.py -q -k "lookahead or burn or direction or manual"` | **21 passed** |
| `venv/bin/pytest tests/test_gsm_geo.py tests/test_gsm_generation_service.py tests/test_gsm_generation_api.py -q` | **44 passed** |
| `venv/bin/pytest tests/test_gsm_generator.py tests/test_gsm_generation_service.py tests/test_gsm_generation_api.py -q` | **52 passed** |
| `venv/bin/pytest tests/test_gsm_*.py tests/test_geocode_gsm_stations.py tests/test_link_route_stations.py -q` | **205 passed** |

Новые кейсы: SC-B1 (короткий дожиг + min km), burn не уходит ниже 0, max-safe пока не в headroom, SC-A1 (полная библиотека), группа достаточна → библиотеку не трогать, true-impossible → `manual_intervention`.

## Success criteria

| ID | Статус |
|---|---|
| SC-B1 короткий дожиг | PASS |
| SC-A1 полная библиотека | PASS |
| SC-G6' Palisade май `problematic_days == []` | PASS |
| SC-R регрессия generator / direction / manual safety-net | PASS |

## Out of scope (не делалось)

Ночёвки, ILP, синтетические маршруты, UI/API, двухякорный lookahead, запись в рабочий `plita.db`, коммит.
