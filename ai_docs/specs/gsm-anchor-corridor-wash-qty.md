# Spec: GSM — коридор бака при выборе маршрута и честные литры моек

Дата: 2026-08-24. Статус: draft, на ревью.
Идея: [`../ideas/gsm-anchor-corridor-wash-qty.md`](../ideas/gsm-anchor-corridor-wash-qty.md) (2026-08-24, direction confirmed).
Базовый модуль: [`gsm-geo-lookahead-generator.md`](gsm-geo-lookahead-generator.md) (реализован).

## ASSUMPTIONS I'M MAKING

1. **`core/gsm/*` трогаем.** Это осознанное расширение границы против прошлого среза; без генератора красный день не лечится.
2. **Схема БД не меняется.** Правки: логика генератора + SQL-агрегат обзора + правило импорта.
3. **У мойки `qty_liters = None` уже задуман** в парсере и тестах. Старые записи с `1.0` в БД не мигрируем — после фильтра в обзоре они не влияют.
4. **Коридор считается по дневному кругу `2×km`**, как сейчас в `_emit_day`. Соло-плечо не вводим.
5. **«Мойка = короткий» — мягкое предпочтение**, а не жёсткое правило: если короткий маршрут уводит бак в минус, берём длинный из коридора.
6. **`liters_diff` в обзоре становится «топливо vs топливо»**: Σ транзакций только `service_type='fuel'` минус Σ `fuel_issued`. Бейдж скрыт при `wb_count=0`, как и раньше.
7. **Регрессия — существующие тесты генератора зелёные**; новый тест на две мойки подряд сначала красный.
8. **Приёмка на копии БД** (`cp plita.db /tmp/plita_accept.db`), рабочая БД write-операциями не трогается.

→ Поправьте сейчас, иначе иду с этим в план/задачи.

## Objective

Убрать системные причины красных дней и ложного расхождения литров, обнаруженные на Palisade в августе 2026:

- Якорный день (особенно мойка) не должен получать маршрут, выводящий бак из коридора `[0…tank]`, если в библиотеке есть подходящий.
- Σ литров в обзоре и журнале транзакций должна сравнивать топливо с топливом, а не все операции с топливом.

**Пользователь:** бухгалтер (`accountant`).

## Правила выбора маршрута: v1 → v2

### Сейчас (v1, актуальный код)

```
Шаг 1. Группа кандидатов
  Станция в typical? → typical
  Нет → крюк (hook) → hook
  Нет → вся библиотека → all

Шаг 2. Базовый выбор
  hook → минимальный крюк_км
  иначе → максимальная frequency (при равенстве — меньший route_id)

Шаг 3. Lookahead СВЕРХУ
  Следующая заправка Q_next
  Если fuel_after > tank − Q_next → нужно сжечь (km_needed)
  → ищем удлинённый маршрут ≥ km_needed
  → если нет — остаётся выбранный

Шаг 4. Эмиссия дня
  apply_day() → BalanceViolation?
  Да → manual_intervention, день в минус
  Нет → обычный draft
```

**Проблема:** шаг 3 смотрит только «не переполнится ли бак при следующей заправке». Он не спрашивает «не уйдёт ли бак в минус сегодня».

### После исправления (v2, предлагаемый)

```
Шаг 1. Группа кандидатов
  Без изменений: typical / hook / all

Шаг 2. ФИЛЬТР КОРИДОРА (новое)
  Для каждого кандидата:
    fuel_end = fuel_start + Q_today − burn(2×km, norm)
    0 ≤ fuel_end ≤ tank_volume ?
  Да → in_corridor
  Нет → вне коридора (отбрасываем)

Шаг 3. Базовый выбор ИЗ КОРИДОРА
  Если in_corridor пуст:
    → fallback: самый короткий из группы
    → дальше manual_intervention
  Иначе:
    мойка (Q=0) → минимальный km
    топливо    → максимальная frequency
    hook       → минимальный крюк_км

Шаг 4. Lookahead СВЕРХУ (как раньше)
  Может перебить базу: если для Q_next нужен выжиг —
  удлиняем из in_corridor или из всей группы, если в коридоре
  нет достаточно длинного

Шаг 5. Эмиссия дня (как раньше)
  manual_intervention только если даже самый короткий маршрут в минус
```

### Сравнение на случае Palisade 04.08

| Этап | v1 | v2 |
|:---|:---|:---|
| Группа | typical (6 км, 95 км) | typical (6 км, 95 км) |
| Базовый выбор | 95 км (freq=39) | 6 км (мойка → min km) |
| Коридор проверка | нет | 6 км ✅, 95 км ❌ |
| Lookahead | не нужен (13,58+0−27,55 ≤ 20) | не нужен (11,84 ≤ 20) |
| Результат | −13,97 → красный день | 11,84 → обычный день |

### Когда правило «мойка = короткий» перебивается

Только одним условием: **нужен выжиг для следующей заправки**.

```
Мойка 04.08, бак 60 л, следующая заправка 05.08 +50 л

Шаг 3 (мойка): сначала 6 км → бак после = 58,26
Шаг 4 (lookahead): 58,26 + 50 = 108,26 > 70 → нужно сжечь 38 л → 264 км
Итог: 6 км заменяется на маршрут ≥132 км плечо
```

## Tech Stack

- Backend: Python 3, FastAPI, Pydantic v2, SQLite.
- Новых зависимостей нет.

## Commands

```bash
# Backend tests
venv/bin/pytest tests/test_gsm_generator.py -q
venv/bin/pytest tests/test_gsm_overview_api.py -q
venv/bin/pytest tests/test_gsm_transaction_import.py -q
venv/bin/pytest tests/ -q

# Приёмка генерации — только на копии
cp plita.db /tmp/plita_accept.db
# ... генерация августа Palisade на копии
```

## Project Structure

```
core/gsm/generator.py              → CHANGED: коридор при выборе якоря,
                                     предпочтение мойки короткому маршруту
app/repositories/gsm_repository.py → CHANGED: tx_liters только fuel
app/services/gsm_transaction_service.py → CHANGED: wash → qty_liters=None
tests/test_gsm_generator.py        → CHANGED: +тест двух моек подряд
tests/test_gsm_overview_api.py     → CHANGED: Σ литров без моек
tests/test_gsm_transaction_import.py → CHANGED: wash qty=None даже при числе в файле
```

## Code Style

```python
# core/gsm/generator.py
def _fits_corridor(
    route: LibraryRoute,
    *,
    fuel_start: float,
    q_today: float,
    norm: float,
    tank_volume: float,
) -> bool:
    fuel_end = fuel_start + q_today - burn_for_km(_daily_km(route), norm)
    return 0 <= fuel_end <= tank_volume
```

- SQL только в репозитории.
- Суммы `round(..., 2)` на backend.
- `extra="forbid"` на запросах (не задето, но не ломаем).

## Testing Strategy

| Уровень | Что покрывает | Команда |
|---|---|---|
| Unit (pytest) | Две мойки подряд: второй день выбирает короткий маршрут, бак не уходит в минус | `venv/bin/pytest tests/test_gsm_generator.py -q` |
| Unit (pytest) | Коридор сверху по-прежнему работает (удлинение под выжиг) | существующие тесты generator |
| Unit (pytest) | Мойка с lookahead: короткий перебивается, если нужен выжиг | `test_gsm_generator.py` |
| Integration (pytest) | Обзор: `tx_liters` без моек, `liters_diff` топливо vs топливо | `test_gsm_overview_api.py` |
| Integration (pytest) | Импорт: wash с числом в файле → в БД `qty_liters IS NULL` | `test_gsm_transaction_import.py` |
| Regression | Все существующие тесты GSM зелёные | `venv/bin/pytest tests/test_gsm_*.py -q` |

## Boundaries

- **Always:** TDD (сначала красный тест); слои router → service → repository; регрессия `tests/test_gsm_*.py`.
- **Ask first:** новые зависимости; пагинация; изменение контрактов ответов сверх описанного.
- **Never:** схема БД; миграция старых `qty_liters`; соло-плечо; коммиты без явной просьбы.

## Success Criteria

1. Генерация августа Palisade на копии БД: 04.08 — обычный draft 12 км, без `manual_intervention`.
2. В обзоре августа Palisade `liters_diff` = 0.0 л (топливо vs топливо).
3. Импорт `.xls` с мойкой, у которой в файле число, сохраняет `qty_liters = None`.
4. Все существующие тесты GSM зелёные.
5. `red_days` в обзоре августа Palisade = 0.

## Open Questions

- Нет. Решения по мойке, `qty_liters` и регрессии зафиксированы.
