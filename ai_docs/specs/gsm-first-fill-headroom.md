# Spec: ГСМ — запас бака смотрит на первую заправку дня, не на сумму

Дата: 2026-10-24. Статус: confirmed.
Базовый модуль: [`gsm-geo-lookahead-generator.md`](gsm-geo-lookahead-generator.md)
(lookahead и `manual_intervention`), [`gsm-anchor-corridor-wash-qty.md`](gsm-anchor-corridor-wash-qty.md)
(коридор `[0…tank]`), [`gsm-fleet-pool-corridor.md`](gsm-fleet-pool-corridor.md)
(жёлтый `balance_route` допустим).

## ASSUMPTIONS I'M MAKING

1. **Один лист на водителя на день.** Строки карты одного дня не режут ПЛ
   и не делят километры между заправками внутри дня. `max_daily_km`, бак и
   нормы не меняются.
2. **Одинаковая метка времени — одно событие заправки.** Литраж события —
   алгебраическая сумма `qty_liters` строк с этим `ts`, включая минус
   (возврат/сторно на той же секунде). События идут по возрастанию времени.
3. **Запас под заправку (`tank − Q`) и lookahead `q_next` берут первую
   заправку дня**, не сумму дня. Нет топлива — `Q = 0`, якорь мойки без
   литров по-прежнему не является `Q_next`.
4. **Сумма дня остаётся на листе и в коридоре.** `_fuel_issued`, поле
   `fuel_issued` путевого и `q_today` в `_fits_corridor` — алгебраическая
   сумма всех топливных строк дня.
5. **Одна нетто-заправка больше бака — по-прежнему красный день.**
   `target = tank − Q < 0` недостижим, день остаётся `manual_intervention`.
   Жёлтый `balance_route` не запрещаем.
6. **Живую SQLite не правим.** Ни схемы, ни транзакций, ни уже выписанных
   листов. Починка — пересчёт генератором после выкладки кода.
7. **Контракт `generate()` не меняется.** Новые хелперы рядом с
   `_fuel_issued`, без `app.*` и без доступа к БД.

## Objective

Убрать ложный красный `manual_intervention` на Geely Tugella О 848 ХР 44
за 2026-08-05 и 2026-08-06. Сейчас запас и lookahead подставляют сумму
топлива за день. 2026-08-06 на карте:

- 09:41:59 `+50.00` и `−15.28` — одна заправка, нетто 34.72 л;
- 17:01:49 `+10.00`;
- 17:06:42 `+13.91`.

Сумма дня 58.63 л при баке 55 л. Проверка `55 − 58.63 = −3.63` невозможна,
день красится красным, хотя к концу дня топливо уже внутри бака.

**Пользователь:** бухгалтер (`accountant`).

**Критерий успеха:** на цепочке 4–7 августа 2026 (бак 55, летняя норма 9.4,
суточный круг ≤ 700) дни 5, 6 и 7 августа не имеют `manual_intervention`;
на листе 6 августа `fuel_issued` равен сумме дня (58.63, до копейки).
Одна транзакция больше бака по-прежнему даёт `manual_intervention`.

## Tech Stack

- Backend: Python 3, чистый `core/gsm/generator.py` (без FastAPI и SQLite).
- Новых зависимостей нет.

## Commands

```bash
.venv/bin/python -m pytest tests/test_gsm_generator.py -q
.venv/bin/python -m pytest tests/test_gsm_generation_api.py tests/test_gsm_generation_service.py -q
```

## Project Structure

```
core/gsm/generator.py          → CHANGE: _fill_events, _first_fill_liters;
                                 запас и _next_fuel_anchor на первую заправку
tests/test_gsm_generator.py    → +цепочка 4–7.08.2026, заправка > бака,
                                 пара плюс/минус в одну секунду
ai_docs/specs/gsm-first-fill-headroom.md → этот документ
```

## Code Style

```python
def _fill_events(txs: Sequence[Transaction]) -> tuple[float, ...]:
    """Liters by ascending time; one timestamp is one algebraic sum."""


def _first_fill_liters(txs: Sequence[Transaction]) -> float:
    """First clustered fill of the day; 0 when the day has no fuel."""
```

Хелперы стоят сразу после `_fuel_issued`. Топливо — только
`service_type == "fuel"` и `qty_liters is not None`. Мойка в сумму не входит.

## Правило запаса

```
события = кластеры по ts ↑          # сумма qty, включая минус
Q_first = события[0] если есть иначе 0
Q_day   = Σ qty за день             # как сейчас _fuel_issued

до якоря, если Q_first > 0:
    target_max = tank − Q_first
    _plan_burn_in(..., target_fuel_max=target_max, fuel_to_issue=Q_first)

следующий якорь (_next_fuel_anchor):
    q_next = Q_first этого дня, если Q_first > 0

лист и коридор:
    fuel_issued = q_today = Q_day
```

Километры между внутридневными заправками не распределяются. Один
`WaybillDay` на дату.

## Testing Strategy

| Уровень | Что покрывает |
|---------|---------------|
| Unit | 4–7.08.2026, бак 55, норма 9.4, круг ≤ 700. 04.08: 22.97. 05.08: 28.25. 06.08: четыре строки выше. 07.08: 30.00 в 06:53 и 18.05 в 08:27. Нет `manual_intervention` на 05, 06 и 07; `fuel_issued` 06.08 ≈ 58.63 |
| Unit | Одна транзакция больше бака → `manual_intervention`. Нетто одной секунды больше бака — тоже |
| Unit | День из пары плюс/минус в одну секунду — одно событие; лист несёт нетто |
| Регрессия | `tests/test_gsm_generator.py`, `tests/test_gsm_generation_api.py`, `tests/test_gsm_generation_service.py` без ослабления старых проверок |

## Boundaries

- **Always:** минимальный diff; сумма дня на листе; первая заправка только
  в запасе и в `q_next`; существующие тесты не ослаблять.
- **Never:** не резать лист по заправкам; не менять `max_daily_km`, бак,
  нормы; не писать в продовую `plita.db`; не гасить жёлтый `balance_route`
  ради зелёного дня.

## Success Criteria

1. SC-1: 2026-08-05 и 2026-08-06 на описанной цепочке без
   `manual_intervention`.
2. SC-2: 2026-08-06 `fuel_issued` = 58.63 (допуск 0.01).
3. SC-3: 2026-08-07 без `manual_intervention`; его `q_next` для шестого
   числа — 30.00, не 48.05.
4. SC-4: одна заправка (или нетто одной секунды) больше бака остаётся
   `manual_intervention`.
5. SC-5: регрессия генератора, API и сервиса генерации зелёная.
