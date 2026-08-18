# Spec: GSM солвер — бак важнее группы АЗС + короткий дожиг

Дата: 2026-08-17. Статус: draft, на ревью.
Идея: [`../ideas/gsm-solver-tank-first-short-burn.md`](../ideas/gsm-solver-tank-first-short-burn.md).
Родительская спека v2: [`gsm-geo-lookahead-generator.md`](gsm-geo-lookahead-generator.md) (2026-08-15, реализована).
Базовый модуль: [`gsm-module-putevye-listy.md`](gsm-module-putevye-listy.md).

## ASSUMPTIONS I'M MAKING

1. Это **доработка солвера v2**, не новый модуль и не смена API. Меняется
   только `core/gsm/generator.py` (+ тесты). Схема БД, эндпоинты, UI бейджей
   не трогаем.
2. Каркас lookahead уже верный: (1) обычный якорь → (2) если будни дожгут
   остаток — не удлинять → (3) иначе удлинить сегодня. Ломаются **пулы**
   кандидатов, не эта политика.
3. **A:** если в `_anchor_route_group` нет `2×km ≥ km_needed`, искать по
   **всем** маршрутам машины в `max_daily_km`. Гео-сортировка остаётся мягкой
   после отбора по км (R3 родителя).
4. **B:** сетка `BURN_KM_MIN/MAX = 150…250` больше не единственный набор
   дожига. Кандидаты — все маршруты уже отфильтрованные `2×km ≤ max_daily_km`.
   `_plan_burn_in` и так отсекает уход в минус (`0 ≤ nxt ≤ tank`).
5. Нижнего пола «правдоподобия» нет (человек, 2026-08-17): минимальный
   **достаточный** км. Плечо 6 км возьмётся только если нужно сжечь ~2 л;
   на 07.05 (~10 л) возьмутся ~45–50 км.
6. Синтетических маршрутов нет. В библиотеке Palisade уже есть и ≥170 км
   (68 шт.), и ≤50 км (16 шт.).
7. `manual_intervention` остаётся предохранителем в core/API. Цель среза —
   чтобы на Palisade май 2026 он **не срабатывал**, не чтобы удалить поле.
8. Ночёвки, ILP, lookahead на два якоря вперёд — **не в этом срезе**.
   Двухякорный lookahead — только если после A+B 21.05 всё ещё красный.
9. Жёлтый `balance_route` на удлинённом якоре **оставляем** (не ошибка).
10. Коммиты/push — только по явной просьбе.

→ Поправьте сейчас, иначе иду с этим в план/задачи.

## Objective

Убрать ложные «ручные доработки» на плотных заправках, где решение уже есть
в `gsm_route`, но солвер его не видит.

**Пользователь:** бухгалтер (`accountant`). Job: нажал «Сгенерировать» →
каждый залив физически влезает в бак → подтвердил. Красный день — проигрыш
солвера, не сценарий работы.

**Две дыры v2 (замер Palisade, `vehicle_id=1`):**

| Кейс | Что ломается | Что должно произойти |
|---|---|---|
| 06→08.05, есть чт 07.05 | Дожиг только 150–250 км плеча (~43 л) из остатка 20 л нельзя | 06.05 остаётся «свой» Ростов 270 км; 07.05 короткий ~45 км, сжечь ~10 л |
| 20→21.05, 0 будней | Удлинение только в группе АЗС (95–100 км), нужно ~180 км плеча | 20.05 берёт минимальный достаточный из **всей** библиотеки + `balance_route` |

**Пользовательский успех:** Palisade май 2026 (04.05–31.05, старт 28 л /
128327 км) → `problematic_days == []`, без 422. Жёлтые «маршрут для баланса»
допустимы.

## Tech Stack

Без изменений: Python, чистый `core/gsm/*` (нет I/O, нет `app.*`). pytest
из корня через `venv/bin/pytest`.

## Commands

```bash
venv/bin/pytest tests/test_gsm_generator.py -q
venv/bin/pytest tests/test_gsm_generator.py -q -k "lookahead or burn or direction or manual"
venv/bin/pytest tests/test_gsm_geo.py tests/test_gsm_generation_service.py tests/test_gsm_generation_api.py -q
venv/bin/pytest tests/test_gsm_*.py tests/test_geocode_gsm_stations.py tests/test_link_route_stations.py -q
# Приёмка на копии БД (не писать в рабочий plita.db):
# GsmGenerationService.generate(vehicle_id=1, 2026-05-04..31, fuel_start=28, odometer_start=128327, force=True)
cd frontend && npm test -- --run src/features/gsm/
```

## Project Structure

```
core/gsm/generator.py              → CHANGED: пул lookahead (A), пул дожига (B)
tests/test_gsm_generator.py        → CHANGED: кейсы 06→08 и 20→21; регрессия
ai_docs/specs/                     → этот файл
ai_docs/develop/plans/             → план задач
```

Не меняем: `app/*`, frontend, схему SQLite, `scripts/`, `BURN_KM_*` как
публичный контракт API (константы можно удалить/не использовать внутри).

## Code Style

Чистые функции, frozen DTO, детерминизм (frequency / `route_id` / seed).

```python
# A: группа станции, затем вся библиотека
elongated = _pick_min_sufficient(group, km_needed=km_needed, ...)
if elongated is None:
    elongated = _pick_min_sufficient(routes, km_needed=km_needed, ...)  # уже в max_daily_km

# B: не фильтровать 150–250; route_list уже _routes_within_daily_cap
def _ordered_burn_routes(routes, *, seed: int) -> tuple[LibraryRoute, ...]:
    rng = Random(seed)
    tie = {r.route_id: rng.random() for r in routes}
    ordered = sorted(routes, key=lambda r: (-r.frequency, r.route_id, tie[r.route_id]))
    return tuple(ordered)

# Когда дожиг попадает в headroom — минимальный достаточный km (R4), не max burn
reaching.sort(key=lambda c: (_daily_km(c[0]), -c[0].frequency, c[0].route_id))
```

Пока не достигли headroom и впереди ещё будни — по-прежнему **максимальный
безопасный** burn за день (чтобы не растягивать зря). В день попадания в
коридор — минимальный km среди `nxt ≤ target_fuel_max`.

## Algorithm (diff от v2)

Нумерация шагов родителя сохраняется. Меняется только отбор кандидатов.

**Дожиг (`_plan_burn_in`).** Вход — все маршруты периода в `max_daily_km`,
не банда 150–250. Как сейчас: кандидат должен дать `fuel_end ∈ [0, tank]`.
Свободный будень, который может сжечь 8–15 л коротким рейсом, **считается
успешным дожигом** → якорь не удлиняется.

**Lookahead (`_select_anchor_route_lookahead`).** Считать `km_needed` с
новым пулом дожига. Если `km_needed > 0`:

1. `_pick_min_sufficient(group)` — как сейчас.
2. Если `None` — `_pick_min_sufficient(все routes в капе)` + `balance_route`.
3. Если снова `None` — оставить обычный `chosen` (дальше emit/manual как в v2).

Не расширять группу **до** проверки дожига: иначе 06.05 зря уедет на Ковров,
хотя четверг может дожечь 10 л.

## Testing Strategy

| Уровень | Что | Команда |
|---|---|---|
| Unit | B: между якорями есть будень; короткий дожиг сжигает 8–15 л; якорь остаётся «своим» (typical), без `manual_intervention` | `pytest tests/test_gsm_generator.py -q -k burn` |
| Unit | A: 0 будней, группа станции слишком короткая, полная библиотека даёт достаточный km → удлинение + `balance_route`, не manual | `-k lookahead` |
| Unit | Регрессия: dense Fri–Mon без коротких/длинных вне группы по-прежнему manual; `max_daily_km` режет; direction; round-trip | `test_gsm_generator.py -q` |
| Acceptance | Palisade май 2026 на копии БД: `problematic_days == []` | скрипт/сервис как 2026-08-15 |
| Регрессия GSM | существующие сьюиты зелёные | `pytest tests/test_gsm_*.py -q` |

TDD: красные тесты на A и B до изменения пулов.

## Boundaries

- **Always:** TDD; `core/gsm` без `app.*` и без I/O; детерминизм; не удалять
  падающие тесты без замены; не писать приёмку в рабочий `plita.db` без копии.
- **Ask first:** lookahead на два якоря, если май после A+B всё ещё с
  `problematic_days`; смена контракта generate; синтетический маршрут.
- **Never:** ночёвки; ILP; геокодинг из backend; смена схемы SQLite;
  удаление `problematic_days` / emit; коммит без просьбы; `bot_archived`.

## Success Criteria

1. **SC-B1 (короткий дожиг):** фикстура «ср заправка → чт свободен → пт большая
   заправка»; в группе якоря только короткие typical; в библиотеке есть плечо
   ~45 км. 06-е остаётся typical; 07-е — дожиг с `2×km` достаточным и
   `fuel_end ≥ 0`; пт без `manual_intervention`.
2. **SC-A1 (полная библиотека):** фикстура два якоря подряд, 0 будней; typical
   ≤100 км; в библиотеке есть плечо ≥180 км. Первый якорь удлинён из
   библиотеки, warning `balance_route`, второй не manual.
3. **SC-G6' (Palisade май):** 04.05–31.05, 28 л / 128327 км →
   `problematic_days == []`, 200, дни собраны. (Ужесточение SC-G6 родителя:
   было ≤3 manual.)
4. **SC-R (регрессия):** `tests/test_gsm_generator.py` зелёный, включая
   direction / round-trip / существующие lookahead / manual-предохранитель
   (когда в библиотеке правда нет решения).

## Open Questions

- Если после A+B на живом Palisade мае останется красный день — не угадывать
  в этом срезе: стоп, замер, отдельное решение (двухякорный lookahead).
- Сортировка `reaching` в `_plan_burn_in` сегодня frequency-first. Этот срез
  меняет на **min km** среди попавших в headroom (R4). Если регрессия v1
  завязана на частотный дожиг 190 км — правим тест под min sufficient.

## Related

- Idea: [`../ideas/gsm-solver-tank-first-short-burn.md`](../ideas/gsm-solver-tank-first-short-burn.md)
- Acceptance v2: [`../develop/reports/2026-08-15-gsm-geo-lookahead-acceptance.md`](../develop/reports/2026-08-15-gsm-geo-lookahead-acceptance.md)
- Код: `_anchor_route_group`, `_select_anchor_route_lookahead`,
  `_ordered_burn_routes`, `_plan_burn_in`
