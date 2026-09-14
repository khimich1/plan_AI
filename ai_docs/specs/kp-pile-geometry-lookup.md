# Spec: Заводская марка сваи → норма загрузки в КП

> **Источник идеи:** [`ai_docs/ideas/kp-pile-geometry-lookup.md`](../ideas/kp-pile-geometry-lookup.md)  
> **План и задачи:** [`ai_docs/develop/plans/2026-09-14-kp-pile-geometry-lookup.md`](../develop/plans/2026-09-14-kp-pile-geometry-lookup.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS ✅ → IMPLEMENT ✅  
> **Статус:** implemented  
> **Дата:** 2026-09-14  
> **Связанные модули:** `core/pile_catalog.py`, `core/pile_trip_pricing.py`, `core/commercial_pricing.py`  
> **Расширяет:** [`pile-bridge-trip-calculation.md`](pile-bridge-trip-calculation.md) assumption 9 / D4 — геометрия уже есть для мостовых; заводской суффикс нагрузки на цельных сваях туда не попал.

---

## Assumptions I'm Making

1. **Web-only.** Черновик КП + архив; бот вне scope.
2. **Нет миграции БД и нет импорта.** 44 строки `pile_catalog` уже в `plita.db`. Каталог не расширяем полными марками.
3. **Нагрузка не логистика.** `-12` / `-13и` влияют на цену (`pile_prices`), не на `weight_kg` и не на `pcs_per_20t`. Одна геометрия = одна строка каталога.
4. **Тихий UI.** После матча блока «нет нормы» нет. Расшифровку «взяли С110.35, 6 шт.» не показываем (D16 родительского spec).
5. **`С160.*` честно pending.** После успешного резолва в `С160.35` `pcs_per_20t` всё ещё NULL → вопрос про машины. Не считаем `ceil(кг / 19800)` вместо нормы.
6. **Override-ключ не меняем.** В `pile_trip_overrides` остаётся отображаемая марка строки (`С160.35-12`), не канон каталога. Pending-список тоже показывает марку из КП.
7. **Отгрузки не трогаем.** `pile_weight_for_mark`, datalist, `search_pile_catalog` — вне MVP (D2 родителя).
8. **Прайс и парсер заказа не трогаем.** `pile_line_parser` / `_PILE_MARK_RE` по-прежнему отдают `С120.35-12` целиком — иначе сломается lookup цены.
9. **Без новых pip/npm зависимостей.** Без правок фронта, если бэкенд перестанет класть марку в `pile_trip_pending_marks`.
10. **Регресс родителя обязателен:** тендер 42 рейса, C18 pending, явный N=0.

→ Поправьте до PLAN, если что-то неверно.

---

## Decisions locked (ideation 2026-09-14)

| # | Тема | Решение |
|---|------|---------|
| D1 | Кто | Менеджер КП (итоги/скидка) + тот же расчёт в архиве |
| D2 | Успех | Тихий матч: норма есть → рейсы считаются, вопроса нет |
| D3 | Где чинить | Read-time lookup в `resolve_catalog_for_mark` / `parse_pile_mark` |
| D4 | Каталог | Не дублировать `-12` в `pile_catalog` |
| D5 | `С160.*` | Как сейчас: нет `pcs_per_20t` → спросить машины |
| D6 | Отгрузки | Не в этом релизе |
| D7 | Identity / `logistics_mark` в order_data | Не пишем; старые черновики лечатся тем же lookup |
| D8 | Хелпер суффикса | Только `core/pile_catalog.py`, не `pile_line_parser` |
| D9 | UI | Без новой плашки и без смены копирайта вопроса |

---

## Objective

Менеджер на шаге итогов КП с цельными сваями (`С110.35-12` и т.п.) видит посчитанные рейсы из заводской колонки «автомобильный г/п 20тн», а не ложный вопрос «нет нормы в справочнике».

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | ввёл `С110.35-12` 498 шт. | чтобы система взяла 6 шт./машину у `С110.35` | не считать фуры вручную |
| US-2 | в КП смесь марок с нормой | увидеть доставку свай после тарифа рейса | отдать клиенту логистику |
| US-3 | есть `С160.35-12` | чтобы **спросили** число машин | не выдумать загрузку 16 м на 13,6 м борт |
| US-4 | открыл архив того же КП | те же pending / те же рейсы | не пересчитывать |

### Reframed success criteria

| Требование | Критерий |
|------------|----------|
| Заводская марка с нормой | `resolve_catalog_for_mark("С110.35-12")` → запись `С110.35`, `pcs_per_20t=6` |
| `-13и` | `С120.35-13и` → `С120.35` |
| Канон без суффикса | `С110.35` по-прежнему exact match |
| Мостовые | `C14-40T4` / `C11-35T6` без регрессии |
| Quirk | `С137,5.40` без регрессии |
| C18 | `C18-40T8` pending (нет геометрии в каталоге) |
| 16 м | `С160.35-12` резолвится в `С160.35`, но `pcs_per_20t is None` → pending, пока нет override |
| Рейсы | `С110.35-12` × 12 при pcs=6 → `full_trips=2`, `ready=True`, не в `pending_marks` |
| UI | `pile_trip_pending_marks` не содержит марок, у которых в каталоге есть pcs ≥ 1 |
| Формула родителя | Тендер без C18 = 42; с C18 без N = `total_trips=0` |
| Отгрузки | Нет diff в `shipment_repository.py` / logistics UI |

---

## Tech Stack

| Слой | Стек |
|------|------|
| Domain | Python, `core/pile_catalog.py` |
| Расчёт | уже зовёт `resolve_catalog_for_mark` из `core/commercial_pricing.py` → `compute_pile_trips` |
| Tests | pytest `tests/test_pile_catalog_import.py`, `tests/test_pile_trip_pricing.py` |
| UI | без изменений, если pending-список правильный |

---

## Commands

```bash
source .venv/bin/activate

pytest tests/test_pile_catalog_import.py tests/test_pile_trip_pricing.py \
  tests/test_commercial_logistics_cost.py tests/test_commercial_pile_pricing.py -q

# Живая проверка резолва (не пишет БД):
.venv/bin/python -c "
from core.pile_catalog import load_pile_catalog, resolve_catalog_for_mark
e = load_pile_catalog('plita.db')
for m in ['С110.35-12','С120.35-13и','С160.35-12','C11-35T6']:
    r = resolve_catalog_for_mark(m, e)
    print(m, '->', None if r is None else (r.mark, r.pcs_per_20t))
"
```

---

## Project Structure

```
ai_docs/ideas/kp-pile-geometry-lookup.md     # идея
ai_docs/specs/kp-pile-geometry-lookup.md     # этот spec
core/pile_catalog.py                         # EXTEND — strip suffix + parse/resolve
tests/test_pile_catalog_import.py            # EXTEND — parse/resolve factory marks
tests/test_pile_trip_pricing.py              # EXTEND — trips ready for С120.35-12
```

Не ожидаем правок: `pile_line_parser.py`, `pile_trip_pricing.py` (логика pending уже верная), фронт, отгрузки, схема БД.

---

## Code Style

Один хелпер рядом с `parse_pile_mark`. Суффикс — только нагрузка цельной сваи в конце марки, не мостовой `C14-40T4`.

```python
_LOAD_SUFFIX_RE = re.compile(r"-\d+(?:[иИ])?$", re.UNICODE)


def strip_pile_load_suffix(mark: str) -> str:
    """«С110.35-12» / «С120.35-13и» → «С110.35». Мостовые и канон каталога без изменений."""
    text = (mark or "").strip()
    return _LOAD_SUFFIX_RE.sub("", text)


def parse_pile_mark(mark: str) -> tuple[Optional[float], Optional[int]]:
    # 1) отрезать нагрузку, 2) как сейчас: длина в дм, сечение в см после последней точки
    ...
```

`resolve_catalog_for_mark`:

1. Точный ключ исходной марки (`C↔С`).
2. Точный ключ после `strip_pile_load_suffix` (чтобы `С110.35-12` попал в `С110.35` без геометрии).
3. Как сейчас: `parse_bridge_pile_geometry`, иначе `parse_pile_mark`; матч `length_m` + `section_mm`; предпочесть строку с `pcs_per_20t ≥ 1`.

Мостовой regex **не** меняем. `C14-40T4` не должен терять `T4` на шаге 2: суффикс `-40T4` не матчит `-\d+(?:и)?$`.

---

## Testing Strategy

| Уровень | Где | Что |
|---------|-----|-----|
| Unit parse | `test_pile_catalog_import.py` | `С60.30`, `С137,5.40`, `С110.35-12` → (11.0, 350), `С120.35-13и` → (12.0, 350), `C14-40T4` parse_pile не обязан (мостовой парсер) |
| Unit resolve | тот же файл | factory → канон с pcs; `С160.35-12` → `С160.35` и pcs None; `C18-40T8` is None; мостовые без регрессии |
| Unit trips | `test_pile_trip_pricing.py` | `С120.35-12` qty кратно pcs → ready, не pending; `С160.35-12` pending без override |
| Регресс | существующие тесты 42 / C18 / plates ignored | зелёные без смены цифр |

Покрытие: не гонимся за %; обязательны позитив с `-12`, `-13и`, негатив 16 м и C18.

---

## Boundaries

- **Always:** тесты parse/resolve/trips зелёные до PLAN→IMPLEMENT; не менять ключ override; не писать в `pile_catalog`.
- **Ask first:** трогать отгрузки / datalist; менять формулу `С160` на вес; новый модуль identity; править `pile_line_parser`.
- **Never:** дублировать прайсовые марки в каталог; считать 16 м через 19800 «чтобы не спрашивать»; коммитить `plita.db`.

---

## Architecture

```
order_data.mark = «С110.35-12»     # цена, PDF, override-ключ
        │
        ▼
resolve_catalog_for_mark
  exact «C110.35-12»           → miss
  exact stripped «C110.35»     → hit С110.35 (pcs=6)
  else geometry L×a            → тот же hit
        │
        ▼
compute_pile_trips
  pcs≥1 → full/remainder
  pcs NULL (С160) / miss (C18) → pending_marks как сейчас
```

Фронт читает `totals.pile_trip_pending_marks`. Если бэкенд перестал класть `С110.35-12` — оранжевый блок пропадает сам.

---

## Success Criteria

- [x] `С110.35-12` / `С120.35-12` / `С140.35-12` / `С150.35-12` не в pending при живом каталоге.
- [x] `С160.35-12` в pending, пока нет N.
- [x] Тесты из Commands зелёные.
- [x] Diff не содержит logistics/shipment/parser цены.

---

## Open Questions

Нет блокирующих. Spec закрывает вопрос идеи про `pile_line_parser`: **не трогаем**. Assumptions подтверждены перед PLAN (2026-09-14).
