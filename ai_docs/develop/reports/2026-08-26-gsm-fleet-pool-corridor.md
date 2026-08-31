# Отчёт: GSM — коридор бака по пулу маршрутов парка

**Дата:** 2026-08-26  
**Спека:** [`../../specs/gsm-fleet-pool-corridor.md`](../../specs/gsm-fleet-pool-corridor.md)  
**План:** [`../plans/2026-08-26-gsm-fleet-pool-corridor.md`](../plans/2026-08-26-gsm-fleet-pool-corridor.md)  
**Статус:** T1–T4 выполнены. Checkpoint зелёный. Коммитов нет. Live `plita.db` не писалась.

## Что сделано

Системные красные дни июля (Monjaro 01.07 / 16.07, Tugella 952 27–28.07): солвер смотрел typical-группу или только `gsm_route` этой машины, хотя в парке был маршрут с нужным километражом в коридоре бака `[0…объём]`.

1. **Пул парка** (`LibraryRoute.vehicle_id`, `_load_routes()` без фильтра машины): якорь и дожиг выбирают из всех маршрутов с базой Кузнецкая 18Б.
2. **Каскад `_select_anchor_route_fleet`:** `in_corridor` → `S` (lookahead) → rank (свой важнее чужого). Не выбирать вне коридора, если коридор непустой. Чужой `route_id=None` в ногах.
3. **Жёлтый `borrowed_route`:** короткая подпись «Чужой пул», `SOFT_WARNING_CODES`, zip не hard-stop.
4. **Приёмка июля** на копии `/tmp/plita_fleet_pool.db`, `force=False`. Live не трогали.

## Задачи

| Task | Содержание | Статус |
|------|------------|--------|
| T1 | `LibraryRoute.vehicle_id`, `_city_key` / `_is_home_base` / `_norm_addr`, `generate(own_vehicle_id=0)`, `_load_routes()` без фильтра | ✅ |
| T2 | Каскад `_select_anchor_route_fleet`, rank own-first, persist `route_id=None` для донора, дожиг из того же пула + `borrowed_route`; `_select_anchor_route_legacy` для A/B без Кузнецкой | ✅ |
| T3 | Фронт: `borrowed_route` → «Чужой пул», `SOFT_WARNING_CODES`, не hard-stop | ✅ |
| Checkpoint | `tests/test_gsm_*.py` 248 passed; vitest T3 25 passed | ✅ |
| T4 | Июль Monjaro + 952 на копии `/tmp/plita_fleet_pool.db` | ✅ |

Чекбоксы плана T1–T4 и Checkpoint — все `[x]`.

## Success Criteria спеки

| # | Критерий | Доказательство | Вердикт |
|---|----------|----------------|---------|
| 1 | Мойка + пустая typical + свой короткий → день в коридоре, без `manual_intervention` | Unit T2 + копия Monjaro 01.07: 12 км, `fuel_end=6.07`, `balance_route` | ✅ |
| 2 | Нет своей ступени в окне + чужой в коридоре → бак ≥ 0, `borrowed_route` | Unit T2 (265 Кузнецкая) + копия 952 27.07: 520 км, `fuel_end=2.18`, `borrowed_route` | ✅ |
| 3 | Свой и чужой оба подходят → свой | Unit: свой `route_id` не None | ✅ |
| 4 | Дожиг берёт чужой короткий, если свой не в коридоре | Unit: круг 12, `borrowed_route` | ✅ |
| 5 | `tests/test_gsm_generator.py` и `tests/test_gsm_*.py` зелёные | **248 passed** | ✅ |
| 6 | Vitest warning / exportGate для `borrowed_route` | **25 passed** | ✅ |
| 7 | Копия, июль Monjaro: 01.07 и 16.07 без `manual_intervention`, бак ≥ 0 | 01.07 `fuel_end=6.07`; 16.07 `fuel_end=15.64` | ✅ |
| 8 | Копия, июль 952: 27.07 `fuel_end ≥ 0`, km ∈ [494, 547], не круг 560 | 27.07 `fuel_end=2.18`, km=520 | ✅ |

## Прогоны

```
.venv/bin/python -m pytest tests/test_gsm_*.py -q
  → 248 passed

cd frontend && npx vitest run \
  src/features/gsm/lib/waybillWarnings.test.ts \
  src/features/gsm/lib/exportGate.test.ts
  → 25 passed
```

Bulk API: пустой флот → `gsm_routes_required`; маршруты соседней машины позволяют generate (`test_bulk_sibling_routes_allow_generate`).

## Приёмка T4 (только копия)

**Метод:** `cp plita.db /tmp/plita_fleet_pool.db`, затем `GsmGenerationService.generate` на копии. Live uvicorn на `:8000` / рабочая `plita.db` не вызывались. `force=true` не использовался.

### Процедура копии

Первый `generate(..., force=False)` упёрся в `gsm_confirmed_conflict`: красные июльские дни были `exported`. На **копии** 24 строки периода машин 2 и 4 понижены `exported` → `draft` (якорь **30.06 не трогали**). Затем `force=False`.

### Vehicle 2 Monjaro

Было: 01.07 `fuel_end=-10.84` / 190 км / `manual_intervention`.

| date | fuel_end | km | warnings | route |
|------|----------|----|----------|-------|
| 2026-07-01 | 6.07 | 12 (6+6) | `balance_route` (нет `manual_intervention`) | Зеленая 1А ↔ Кузнецкая 18Б, `route_id=129` |
| 2026-07-02 | 7.26 | 190 | нет | Кузнецкая ↔ Ярославль, `route_id=3` |
| 2026-07-16 | 15.64 | 190 | нет | тот же Ярославль 95+95, `route_id=3` |

### Vehicle 4 Tugella 952

Было: 27.07 `fuel_end=-1.21` / 560 км.

| date | fuel_end | km | warnings | chosen |
|------|----------|----|----------|--------|
| 2026-07-27 | 2.18 | 520 ∈ [494, 547] | `borrowed_route` | плечо 260 Monjaro Мантурово; `route_id=null`; не 560 |
| 2026-07-28 | 3.3 | 520 | `borrowed_route` | тот же 260+260, `route_id=null` |

265 км СП машины 848 не выбран (rank: min km среди достаточных в коридоре).

### Якоря 2026-06-30 на месте

| vehicle | fuel_end | статус |
|---------|----------|--------|
| v2 Monjaro | 7.21 | exported/imported |
| v4 Tugella 952 | 17.02 | exported/imported |
| v1 Palisade | 15.84 | да |
| v3 Tugella 848 | 13.44 | да |

Все четыре якоря 30.06 на месте.

### Оставшиеся красные (`manual_intervention`) v2 и v4, 2026-07-01…2026-08-25

Нет. `red_count=0` у обеих машин.

### Live `plita.db`

**Не писалась.** Live по-прежнему содержит старые красные строки.

| Файл | md5 | mtime |
|------|-----|-------|
| `plita.db` до | `1d72ce7c2e4b49ce97e594e965ff5c71` | 2026-08-26 14:21:11 |
| `plita.db` после | `1d72ce7c2e4b49ce97e594e965ff5c71` | 2026-08-26 14:21:11 |
| `/tmp/plita_fleet_pool.db` | `5b5c5984b3bd77a3205dde36792db12e` | — |

## Reviewer nits (не чинить в этом срезе)

Approve-with-nits, follow-up:

1. **Лишние ключи rank** перед km: `preferred_ids` и `_direction_priority`. Спека: (1) свой/чужой, (2) тот же город, (3) `_daily_km`, (4) `-frequency`, `route_id`.
2. **Пустой коридор:** полный rank по fallback-пулу, а не `min(daily_km)` как в спеке.

Продакшен-путь совпадает со спекой на июльских кейсах (свои короткие / чужие достаточные в коридоре). Не блокирует.

## Файлы

Изменены:

- `core/gsm/generator.py` — `LibraryRoute.vehicle_id`; `_norm_addr` / `_is_home_base` / `_city_key`; `generate(own_vehicle_id=0)`; `_select_anchor_route_fleet` / `_select_anchor_route_legacy`; дожиг из пула; persist `route_id=None` для донора
- `core/gsm/models.py` — `RouteRef.route_id` / `LegPlan.route_id: int | None`
- `app/services/gsm_generation_service.py` — `_load_routes()` без фильтра машины; `own_vehicle_id` в `generate`
- `tests/test_gsm_generator.py` — ключи города/базы; каскад якоря; дожиг; свой vs чужой
- `tests/test_gsm_generation_api.py`, `tests/test_gsm_generate_bulk_api.py` — пустой флот → `gsm_routes_required`; sibling routes позволяют generate
- `frontend/src/features/gsm/lib/waybillWarnings.ts` — `borrowed_route` → «Чужой пул»
- `frontend/src/features/gsm/lib/exportGate.ts` — `SOFT_WARNING_CODES`
- `frontend/src/features/gsm/types/gsm.ts` — union `borrowed_route`
- соответствующие `.test.ts`

Документы:

- `ai_docs/develop/plans/2026-08-26-gsm-fleet-pool-corridor.md` — T1–T4 и Checkpoint `[x]`
- `ai_docs/develop/reports/2026-08-26-gsm-fleet-pool-corridor.md` — этот отчёт

Временная копия `/tmp/plita_fleet_pool.db` в репозиторий не входит.

## Границы

Соблюдено:

- нет схемы БД
- нет зажима `fuel_end=0`
- нет выдуманных городов
- нет `force=true` на 30.06
- нет generate на live `plita.db`
- uvicorn / `run+logs.sh` не останавливали
- нет коммита / push / git config

**Коммитов нет.**
