# Spec: GSM — коридор бака по пулу маршрутов всего парка

Дата: 2026-08-26. Статус: draft, план готов.
План: [`../develop/plans/2026-08-26-gsm-fleet-pool-corridor.md`](../develop/plans/2026-08-26-gsm-fleet-pool-corridor.md).
Идея: [`../ideas/gsm-fleet-pool-corridor.md`](../ideas/gsm-fleet-pool-corridor.md)
(2026-08-26, direction confirmed).
Предшественник: [`gsm-anchor-corridor-wash-qty.md`](gsm-anchor-corridor-wash-qty.md)
(коридор только внутри typical-группы — дыра осталась).

## ASSUMPTIONS I'M MAKING

1. **`core/gsm/generator.py` + загрузка маршрутов в сервисе.** Без генератора
   красные дни июля не лечатся. Схема БД не меняется.
2. **Пул парка и для якоря, и для дожига.** Чужие адреса в ПЛ допустимы
   (решение 2026-08-26). Свой маршрут важнее чужого при равном коридоре.
   Норма литров — этой машины.
3. **Круг по-прежнему `2×km`.** Ночёвка / соло-плечо — не этот срез.
4. **`borrowed_route` — жёлтый**, как `balance_route`: zip не hard-stop,
   confirm при экспорте. В drawer хватает короткой подписи, бейдж чужого
   госномера не делаем.
5. **«Тот же город»** — эвристика `_city_key(addr)` (см. ниже), не геокодер
   и не словарь всех населённых пунктов РФ.
6. **База** — адрес содержит «кузнецкая» и «18» после нормализации (пробелы,
   `ул.`/`улица`, ё/е). Все четыре машины парка на Кузнецкой 18Б.
7. **Чужой `route_id` в JSON ног не пишем** (`route_id=null`): drawer грузит
   справочник этой машины. Адреса и км в ногах — да.
8. **Приёмка июля** — на копии БД; якорь 30.06 confirmed не трогать
   (`force=false`). Live `plita.db` write-генерацией в этом срезе не гонять,
   пока копия зелёная.
9. **Справочник руками не наполняем** в MVP (ни Мантурово в 952, ни правка
   штампа СП 280→265). Солвер закрывает июль заимствованием на день.

→ Поправьте сейчас, иначе иду с этим в план/задачи.

## Objective

Убрать системные красные дни, когда в парке **есть** маршрут с нужным
километражом в коридоре бака `[0…объём]`, но солвер смотрит только typical-группу
или только `gsm_route` этой машины.

**Пользователь:** бухгалтер (`accountant`), закрытие месяца на `/gsm`.

**Не цель среза:** невозможный день (ни у кого в парке нет км в окне) — остаётся
красным. Не зажимать бак в 0. Не выдумывать города.

## Алгоритм якоря (замена `_select_anchor_route_lookahead`)

Вход: `own_vehicle_id`, `fuel_before`, `q_today`, `norm`, `tank_volume`,
`km_needed` (lookahead, 0 если не нужен), станция дня, `hooks`.

Кандидаты `fleet` = все загруженные `LibraryRoute` с `2×km ≤ max_daily_km`,
у которых **хотя бы один** конец `_is_home_base` (Кузнецкая 18Б).

```
in_corridor = { r in fleet | _fits_corridor(r, fuel_before, q_today, norm, tank) }

S = in_corridor
если km_needed > 0:
    S = { r in in_corridor | _daily_km(r) >= km_needed }

если S не пуст:
    chosen = rank(S)   # зелёный день (плюс жёлтые коды)
иначе если in_corridor не пуст:
    chosen = rank(in_corridor) по max _daily_km
    + manual_intervention  # сегодня бак ок, выжиг под Q_next не закрыт
иначе:
    chosen = min(fleet или typical-группа, key=daily_km)
    + manual_intervention  # сегодня минус; пул парка тоже пуст
```

**Запрещено:** если `in_corridor` непустой, выбирать маршрут вне коридора
(в т.ч. lookahead «560 км при окне 547»).

### Rank (меньше — лучше)

1. `vehicle_id != own` (0 свой, 1 чужой)
2. не `_same_city(r, hint)` — hint = город своего маршрута с km ближайшим
   к `km_needed` или к max в своей библиотеке, кто почти влез; если hint нет — 0
3. `_daily_km` (мойка / пустой lookahead: min; иначе min среди достаточных)
4. `-frequency`, `route_id`

Мойка (`q_today == 0`) не отдельный каскад: короткий выигрывает через п.3.

### Warnings

| Условие | Код |
|---------|-----|
| выбран не из typical-группы станции дня, но свой `vehicle_id` | `balance_route` |
| `vehicle_id != own` | `borrowed_route` |
| день вне коридора или lookahead-окно пусто при непустом коридоре | `manual_intervention` |

Можно оба жёлтых сразу (`balance_route` + `borrowed_route`).

### `_city_key` / `_is_home_base`

```python
def _norm_addr(s: str) -> str:
    t = s.casefold().replace("ё", "е")
    t = t.replace("улица", "ул").replace("ул.", "ул")
    return " ".join(t.split())

def _is_home_base(addr: str) -> bool:
    n = _norm_addr(addr)
    return "кузнецкая" in n and "18" in n

def _city_key(addr: str) -> str:
    n = _norm_addr(addr)
    for name in (
        "сергиев посад", "переславль-залесский", "переславль залесский",
        "нижний новгород", "н.новгород",
    ):
        if name in n:
            return name.replace("-", " ")
    # «г.Вологда, …» / «г. Мантурово, …»
    import re
    m = re.search(r"г\.\s*([^,]+)", n)
    if m:
        return m.group(1).strip()
    return n[:40]
```

Сравнение городов: ключ **небазового** конца (тот, что не home).

## Дожиг (`_plan_burn_in`)

Тот же пул, что якорь: `fleet ∩ home ∩ max_daily_km ∩ in_corridor` (на старте
дня дожига `q_today=0`). Политика км как сейчас (max безопасный выжиг, в день
попадания в коридор — min достаточный). Среди подходящих по км — **rank**:
свой важнее чужого. Если выбран `vehicle_id != own` — `borrowed_route` на
этом буднем дне, `route_id=None`.

## Tech Stack

- Backend: Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`).
- Frontend: React + TypeScript (только подпись warning).
- Новых зависимостей нет.

## Commands

```bash
.venv/bin/python -m pytest tests/test_gsm_generator.py -q
.venv/bin/python -m pytest tests/test_gsm_*.py -q
cd frontend && npx vitest run src/features/gsm/lib/waybillWarnings.test.ts src/features/gsm/lib/exportGate.test.ts

# Приёмка — только копия
cp plita.db /tmp/plita_fleet_pool.db
# generate vehicle 2 and 4, 2026-07-01..2026-08-25, force=false
```

## Project Structure

```
core/gsm/generator.py                 CHANGED: LibraryRoute.vehicle_id;
                                      generate(..., own_vehicle_id=);
                                      каскад якоря; _city_key / home base
app/services/gsm_generation_service.py CHANGED: _load_routes() — все машины;
                                      передать own_vehicle_id
app/repositories/gsm_repository.py    list_routes(vehicle_id=None) уже есть
frontend/src/features/gsm/lib/waybillWarnings.ts  CHANGED: borrowed_route
frontend/src/features/gsm/lib/exportGate.ts       CHANGED: soft set
frontend/src/features/gsm/types/gsm.ts            опционально явный union
tests/test_gsm_generator.py           CHANGED: три новых кейса + регрессия
```

## Code Style

```python
@dataclass(frozen=True, slots=True)
class LibraryRoute:
    route_id: int
    addr_a: str
    addr_b: str
    km: int
    frequency: int
    typical_station_ids: tuple[int, ...]
    vehicle_id: int = 0
    point_a: GeoPoint | None = None
    point_b: GeoPoint | None = None
```

- `core/gsm` без I/O и без SQL.
- `burn` всегда `norm_for` **этой** машины, не донора.
- Persist ног: `from`/`to`/`km`; `route_id` только если `chosen.vehicle_id == own`.

## Testing Strategy

| Кейс | Ожидание |
|------|----------|
| Мойка, typical-группа только 95 км, в своей библиотеке 6 км | круг 12 км, бак ≥ 0, не `manual_intervention`; допустим `balance_route` |
| Своё плечо 280 (круг 560, минус), в пуле чужое 265 с Кузнецкой | круг 530, бак ≥ 0, `borrowed_route`, свой важнее не применяется |
| Свой 6 км и чужой 6 км оба в коридоре | берётся **свой** |
| Дожиг: свой min 95 км уводит в минус, в пуле чужой 6 км | будний день 12 км, бак ≥ 0, `borrowed_route` |
| Lookahead окно 494–547, в пуле пусто, свой 280 вне коридора | `manual_intervention`; не выбирать 560 если есть любой `in_corridor` короче |
| Регрессия Palisade-мойка 24.08 / майские тесты generator | зелёные |

Vitest: `borrowed_route` → короткая подпись; в `SOFT_WARNING_CODES`; не в hard-stop.

## Boundaries

- **Always:** TDD; слои router → service → repository; регрессия `tests/test_gsm_*.py`.
- **Ask first:** ночёвка; запись чужого A→B в `gsm_route`; генерация на live `plita.db`.
- **Never:** схема БД; зажим `fuel_end=0`; синтетический адрес; `force` на confirmed 30.06 в приёмке; коммит без просьбы.

## Success Criteria

1. Unit: мойка + пустая typical + свой короткий → день в коридоре, без `manual_intervention`.
2. Unit: нет своей ступени в окне + чужой 265 Кузнецкая → бак ≥ 0, `borrowed_route`.
3. Unit: свой и чужой оба подходят → свой.
4. Unit: дожиг берёт чужой короткий, если свой не в коридоре.
5. `tests/test_gsm_generator.py` и `tests/test_gsm_*.py` зелёные.
6. Vitest warning/exportGate для `borrowed_route`.
7. Копия БД, июль Monjaro: 01.07 и 16.07 без `manual_intervention` (бак ≥ 0).
8. Копия БД, июль 952: 27.07 `fuel_end ≥ 0`, дневной km в [494, 547] (или эквивалент плеча 247–273), не круг 560 как единственный исход.

## Open Questions

Зафиксировано в assumptions (1–9). Дожиг = тот же пул парка, свой важнее.
