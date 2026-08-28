# Spec: GSM — ориентация круга с базы + домашний `route_id`

Дата: 2026-08-27. Статус: draft, план готов.
План: [`../develop/plans/2026-08-27-gsm-home-oriented-round-trip.md`](../develop/plans/2026-08-27-gsm-home-oriented-round-trip.md).
Идея: [`../ideas/gsm-home-oriented-round-trip.md`](../ideas/gsm-home-oriented-round-trip.md)
(2026-08-27, direction confirmed).
Предшественники: [`gsm-geo-lookahead-generator.md`](gsm-geo-lookahead-generator.md),
[`gsm-fleet-pool-corridor.md`](gsm-fleet-pool-corridor.md).

## ASSUMPTIONS I'M MAKING

1. **Схема БД и API generate не меняются.** Меняется только порядок ног и
   какой свой `route_id` пишется в JSON. `borrowed_route` по-прежнему
   `route_id=null`.
2. **Инвариант дня (этот срез):** если ровно один конец маршрута — база
   Кузнецкая 18Б, круг **база → объект → база**. Состояние «конец вчера»
   и ночёвки — не здесь.
3. **Ориентация после выбора км.** Rank / lookahead / corridor / warnings
   считают по выбранной солвером строке (59). Persist и `RouteRef` — уже
   ориентированная копия (ноги с базы, id зеркала 64 если есть).
4. **Зеркало** — та же `vehicle_id`, тот же `km`, концы переставлены по
   `_norm_addr`. Несколько кандидатов → меньший `route_id`. Нет пары →
   ноги перевёрнуты, id исходный (если свой). Не матчим по городу+km.
5. **Адреса в ногах** — с выбранной строки (текст 59), не подмена формулировки
   с зеркала 64. У зеркала берём только `route_id`.
6. **База** — существующий `_is_home_base` (кузнецкая + 18 после `_norm_addr`;
   «ул.Кузнецкая» без номера — тоже база). Все 4 машины парка на Кузнецкой.
7. **Drawer / ручной ПЛ** не переводим на 2 плеча в этом срезе: сегодня
   `libraryRouteToLegs` даёт **одну** ногу и `km` = плечо справочника.
   Фикс: `from` = база, `route_id` зеркала если есть. Круг 2×km — только
   автогенератор.
8. **Live `plita.db` не пишем.** Приёмка — юниты + vitest; перегенерация
   июля бухгалтером после выкладки, не в этом изменении.
9. **Хелпер ориентации на apply справочника — в этом срезе** (drawer и
   `ManualWaybillDialog`). Иначе выбор 59 откатит фикс.

→ Поправьте сейчас, иначе иду с этим в план/задачи.

## Objective

Убрать ложный выезд с объекта: сгенерированный (и выбранный из справочника)
маршрут читается как выезд **с Кузнецкой 18Б**, если маршрут касается базы.

**Пользователь:** бухгалтер (`accountant`), закрытие месяца на `/gsm`, правка
дня в drawer.

**Критерий успеха:** 13.07 Tugella 952 при тех же км/литрах/`balance_route`
стартует с Кузнецкой; в ногах `route_id=64`, если зеркало находится.

**Не цель:** непрерывность точки между днями; ночёвка; канонизация
`gsm_route`; смена бака/warnings; 2 плеча в drawer.

## Алгоритм

### `_is_home_base` / `_norm_addr`

Без изменений. Как в `core/gsm/generator.py` сейчас.

### `_find_home_twin(chosen, catalog) -> LibraryRoute | None`

Среди `catalog` с `vehicle_id == chosen.vehicle_id` и `km == chosen.km`:

```
_norm_addr(c.addr_a) == _norm_addr(chosen.addr_b)
and _norm_addr(c.addr_b) == _norm_addr(chosen.addr_a)
```

Если несколько — `min(route_id)`. Самого `chosen` не возвращать.

### `_orient_home_round_trip(chosen, *, catalog, own_vehicle_id) -> LibraryRoute`

Пусть `a_home = _is_home_base(chosen.addr_a)`, `b_home = _is_home_base(chosen.addr_b)`.

| a_home | b_home | Действие |
|--------|--------|----------|
| да | нет | уже дом-первым: `route_id` как `_persisted_route_id`; адреса как есть |
| нет | да | переворот: `addr_a, addr_b = chosen.addr_b, chosen.addr_a`; id = twin.id если twin и `chosen.vehicle_id == own_vehicle_id`, иначе `_persisted_route_id(chosen)` |
| да | да | как есть |
| нет | нет | как есть (legacy без Кузнецкой) |

Поля `km`, `frequency`, `typical_station_ids`, `vehicle_id`, координаты —
с `chosen`. Чужой пул: `route_id` остаётся `None` даже при найденном близнеце
донора.

Вызов: в `_emit_day` (якорь и дожиг), до `_round_trip_legs` / `_to_route_ref`.
`_round_trip_legs` по-прежнему `A→B` затем `B→A` от **уже ориентированного**
маршрута.

Warnings якоря считаются **до** ориентации, по солверному `chosen`.

### Frontend: `libraryRouteToLegs`

Сигнатура: `(route, catalog?: GsmRoute[])`.

1. Найти twin в `catalog` (тот же `km`, концы `_normAddr` переставлены,
   не тот же `id`).
2. Если ровно `addr_b` — база, а `addr_a` нет: `from`/`to` переставить;
   `route_id` = twin.id если twin есть, иначе исходный `route.id`.
3. Иначе одна нога как сейчас: `from=addr_a`, `to=addr_b`, `route_id=route.id`.

`isGsmHomeBase` / `normGsmAddr` в TS **повторяют** Python (кузнецкая + 18;
без цифры после «кузнецкая» — база). Drawer и ManualWaybillDialog передают
уже загруженный список маршрутов машины.

Число ног не меняем (по-прежнему 1). `setKm(route.km)` не меняем.

## Tech Stack

- Backend: Python 3, FastAPI, Pydantic v2, SQLite — без новых пакетов.
- Frontend: React + TypeScript; правка `libraryRouteToLegs` и vitest.
- `core/gsm` без I/O.

## Commands

```bash
.venv/bin/python -m pytest tests/test_gsm_generator.py -q
.venv/bin/python -m pytest tests/test_gsm_*.py -q
cd frontend && npx vitest run \
  src/features/gsm/lib/downstreamPreview.test.ts \
  src/features/gsm/lib/waybillWarnings.test.ts \
  src/features/gsm/components/WaybillDayDrawer.test.tsx
```

## Project Structure

```
core/gsm/generator.py
  CHANGED: _orient_home_round_trip, _find_home_twin; _emit_day принимает catalog
frontend/src/features/gsm/lib/downstreamPreview.ts
  CHANGED: libraryRouteToLegs(route, catalog); isGsmHomeBase / normGsmAddr
frontend/src/features/gsm/lib/gsmHome.ts   NEW опционально, если не тесним preview
frontend/src/features/gsm/components/WaybillDayDrawer.tsx
  CHANGED: catalog в libraryRouteToLegs
frontend/src/features/gsm/components/ManualWaybillDialog.tsx
  CHANGED: то же
tests/test_gsm_generator.py               CHANGED: ориентация + зеркало
frontend/.../downstreamPreview.test.ts    CHANGED: home-first / twin / no-twin
```

Схема, репозиторий, generate API — без изменений.

## Code Style

```python
def _orient_home_round_trip(
    chosen: LibraryRoute,
    *,
    catalog: Sequence[LibraryRoute],
    own_vehicle_id: int,
) -> LibraryRoute:
    a_home = _is_home_base(chosen.addr_a)
    b_home = _is_home_base(chosen.addr_b)
    if a_home == b_home:
        return chosen
    if a_home:
        return chosen
    twin = _find_home_twin(chosen, catalog)
    persist_id = (
        twin.route_id
        if twin is not None and chosen.vehicle_id == own_vehicle_id
        else (_persisted_route_id(chosen, own_vehicle_id) or 0)
    )
    # LibraryRoute.route_id is int; borrowed stays 0 then _persisted → None
    ...
```

- Один хелпер ориентации; не дублировать переворот в rank.
- `core/gsm` не импортирует `app.*`.
- TS-эвристика базы должна совпасть с `_is_home_base` по кейсам из
  `tests/test_gsm_generator.py` (`ул. Кузнецкая, д.18Б`, `ул.Кузнецкая`).

## Testing Strategy

| Кейс | Ожидание |
|------|----------|
| Свой 59 объект→дом + зеркало 64, km=225 | круг 450; `legs[0].addr_a` база; `legs[0].addr_b` объект 59; `route_id=64` на обеих ногах; бак/warnings как без ориентации |
| Только объект→дом, зеркала нет | старт с базы; `route_id` исходный (свой) |
| Чужой объект→дом, есть близнец донора | старт с базы; `route_id is None` |
| Оба конца база / ни одного (A/B без Кузнецкой) | порядок как в библиотеке |
| Дожиг с объектом-первым | тот же инвариант ног |
| Регрессия `tests/test_gsm_generator.py` | км, бак, коды warnings; меняется только порядок/id где есть база |
| `libraryRouteToLegs(59, catalog с 64)` | одна нога дом→Владимир, `route_id=64` |
| `libraryRouteToLegs(59, catalog без 64)` | одна нога дом→Владимир, `route_id=59` |
| Drawer: выбор 59 при загруженных маршрутах | превью `from` = Кузнецкая |

Coverage: новые ветки `_orient_home_round_trip` / `_find_home_twin` и
`libraryRouteToLegs` — юнитами выше.

## Boundaries

- **Always:** TDD; регрессия `tests/test_gsm_*.py`; ориентация в `_emit_day`
  (якорь и дожиг одним путём).
- **Ask first:** генерация на live `plita.db`; перевод drawer на 2 плеча;
  ослабление зеркала до `_city_key`.
- **Never:** схема БД; миграция `gsm_route`; ночёвки; зажим бака; коммит без
  просьбы.

## Success Criteria

1. Unit: объект→дом + зеркало → первая нога с `_is_home_base`, persist id зеркала.
2. Unit: объект→дом без зеркала → первая нога с базы, id исходный.
3. Unit: `borrowed_route` → ноги с базы, `route_id is None`.
4. Unit: фикстуры без Кузнецкой — поведение ног как сейчас.
5. `tests/test_gsm_*.py` зелёные; км и `fuel_end` регрессионных дней те же.
6. Vitest: `libraryRouteToLegs` ориентирует и подставляет twin id.
7. Drawer-тест: выбор библиотечной строки объект-первым даёт `from` базы.

## Open Questions

Закрыты в assumptions (1–9). План: T1 хелперы → T2 emit → T3/T4 UI.
