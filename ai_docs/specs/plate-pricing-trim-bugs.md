# Spec: Исправление багов расчёта стоимости плит (trim / отходы / продольный рез)

> **Тип:** bugfix + regression suite  
> **Фаза SDD:** SPECIFY → PLAN → TASKS → IMPLEMENT  
> **Дата:** 2026-07-19  
> **Статус:** на реализацию  
> **Связанные документы:** [`strategic-roadmap-pb-pricing-optimizer-1c.md`](./strategic-roadmap-pb-pricing-optimizer-1c.md) (Phase 1), [`../ideas/plate-pricing-fix.md`](../ideas/plate-pricing-fix.md)  
> **Эталон:** ручной расчёт менеджера завода

---

## Objective

Устранить периодические ошибки в расчёте стоимости плит в **коммерческом КП** и **производственной смете**, когда:

1. **теряется продольный рез** (строка «Продольный рез» отсутствует в breakdown);
2. **отход полосы учитывается дважды** — на primary и на secondary из того же слэба;
3. **появляются необъяснимые отходы** (например, 240 мм на дочерней плите без явного источника в раскладке).

**Пользователи:** менеджер (КП), производство (смета).  
**Успех:** все описанные кейсы + регрессионные тесты на каждый класс ошибки; `pytest` green; сверка с эталонами менеджера.

---

## Scope

### In scope

- `viz_modules/procurement/trim.py` — `_calc_trim_components`, `resolve_long_cut_pricing`, `apply_factory_strip_waste`
- `viz_modules/procurement/price_rows.py`, `breakdown.py` — оба потока (`build_price_rows` / `_production`)
- Регрессионные тесты в `tests/test_procurement_trim_cuts.py` (+ при необходимости `test_procurement_mixed_load_breakdown.py`)
- Скрипт сверки `scripts/reconcile_plate_prices.py` (новый, минимальный MVP)
- Документирование правил «владелец отхода» в комментариях к trim

### Out of scope

- Рефакторинг оптимизатора (ILP) — только трассировка `sec_cut['waste']`
- Пересчёт исторических КП в архиве
- Единый сервис `core/pricing/` (отдельная задача Phase 1 roadmap)
- UI breakdown (кроме отображения уже исправленных формул)

---

## Tech Stack

Без изменений: Python 3, pytest, `viz_modules/procurement/`, константы из `core/config/constants.py`:

| Константа | Значение | Смысл |
|-----------|----------|--------|
| `LONG_CUT_PRICE_PER_M` | 460 ₽/п.м. | Продольный рез |
| `TRANSVERSE_CUT_PRICE` | 1200 ₽/шт. | Поперечный рез |
| `MIN_BILLABLE_TRIM_MM` | 20 мм | Порог тарификации отхода/остатка |

---

## Commands

```bash
# Регрессия trim (основной gate)
pytest tests/test_procurement_trim_cuts.py tests/test_procurement_mixed_load_breakdown.py -q

# Полная регрессия procurement
pytest tests/ -q -k "procurement or trim"

# Сверка с эталонами менеджера (после реализации скрипта)
python scripts/reconcile_plate_prices.py --strict

# Smoke КП / breakdown
pytest tests/test_commercial_web_flow.py -q -k breakdown
```

---

## Project Structure

```
viz_modules/procurement/
  trim.py           ← основная логика (фиксы)
  price_rows.py     ← build_price_rows, build_price_rows_production
  breakdown.py      ← build_component_breakdown*

tests/
  test_procurement_trim_cuts.py   ← регрессия по кейсам
  fixtures/plate_pricing/         ← JSON эталоны менеджера (новое)

scripts/
  reconcile_plate_prices.py       ← diff факт vs эталон (новое)

ai_docs/specs/
  plate-pricing-trim-bugs.md      ← этот документ
```

---

## Бизнес-правила (источник истины — менеджер)

| # | Правило |
|---|---------|
| R1 | **Отход полосы по ширине** (160 мм, 560 мм, 120 мм factory strip) начисляется **один раз** — на **владельце primary-слэба**, с которого резали. |
| R2 | **Secondary-плита** (другая ширина с полосы rest) получает: базовая цена + продольный рез + поперечный рез + остаток по длине — **без** строки «Отходы (Nмм)», если отход уже на primary. |
| R3 | **Same-width cascade** (primary 530 + secondary 530 с rest) — отход на primary или в одной строке заказа; не дублировать (уже покрыто тестами). |
| R4 | **Cross-load cascade** (10п → 8п) — отход только на primary другой нагрузки; secondary — только резы (уже частично в `_apply_crossload_rest_secondaries`). |
| R5 | **Плиты 10,2–10,8 м** (1020–1080 мм): отход `(1200−W)/1200` через factory strip **и** **1 продольный рез** `460 × длина` на каждую плиту. |
| R6 | **Плиты 1,2 м** — без продольного реза (жёсткое правило в price_rows/breakdown). |
| R7 | **Итог за 1 плиту 10,8 м** = цена полной плиты 1,2 м (база пропорционально + отход = полная цена). |

---

## Найденные баги

### BUG-1: Пропадает продольный рез на плитах 10,8 м (и аналогичных 1020–1080 мм)

**Симптомы (кейсы 1–2):**

- Плиты ПБ 30,5-10,8-12п, ПБ 32,1-10,8-12п
- В breakdown: «Базовая цена» + «Отходы (120мм)» или «(120×2мм)»
- **Нет** строки «Продольный рез»
- Итог = цена полной плиты 1,2 м (корректно по сумме, но занижен по составу)

**Корневая причина:**

1. `apply_factory_strip_waste()` добавляет отход 120 мм для диапазона 1020–1080 мм (`trim.py:951–976`).
2. `resolve_long_cut_pricing()` при **наличии плана**, но **нулевом** `long_cut_meterage` из trim возвращает `(0, 0)` и **не использует fallback** (`trim.py:943–948`).
3. Primary match для width=1080 может не дать метраж реза (план / match / skip_primary_rest_cut).

**Исправление (Fix-1):**

В `resolve_long_cut_pricing`: если продольный рез не найден в trim, но:

- `1020 <= width_mm <= 1080` и `abs(width_m - 1.2) > 0.01`, **или**
- `apply_factory_strip_waste` добавил отход (или `waste_cost > 0` от factory strip),

→ начислить **1 продольный рез**: `LONG_CUT_PRICE_PER_M × length / qty` (на единицу).

Альтернатива (минимальнее): в `apply_factory_strip_waste` возвращать флаг `factory_strip_needs_long_cut=True`, потребители учитывают в `resolve_long_cut_pricing`.

**Файлы:** `viz_modules/procurement/trim.py`, возможно `price_rows.py`, `breakdown.py`.

---

### BUG-2: Двойной учёт отхода primary + secondary (cross-width cascade)

**Симптомы (кейс 3):**

- Primary: ПБ 63-7,2-8п → «Отходы (160мм)» ✓
- Secondary: ПБ 45-3,2-8п → **тоже** «Отходы (160мм)» ✗
- 160 мм должен быть **только** на ПБ 63-7,2-8п

**Корневая причина:**

- Primary path (`trim.py:674–746`) начисляет `unused_per_strip` на владельце слэба.
- Secondary path для `primary_matched=False` (`trim.py:833–881`) вызывает `_apply_secondary_cut(..., charge_strip_waste=True)` по умолчанию.
- Для same-width cascade уже `charge_strip_waste=False` в `_apply_cascade_secondary_for_primary` (`trim.py:405`), но **для другой ширины** (320 vs 720) — нет.

**Исправление (Fix-2):**

В блоке `if not primary_matched` перед `_apply_secondary_cut`:

```python
from_rest = _secondary_matches_primary_rest(sec_cut, _rest_groups_from_plan(current_plan))
charge = not from_rest  # отход только если НЕ с полосы primary rest
```

Или: всегда `charge_strip_waste=False`, если `_secondary_matches_primary_rest` **или** `_is_crossload_rest_secondary` (отход уже на primary другой нагрузки).

Вспомогательная функция `_rest_groups_from_plan(plan)` — собрать `rest_groups` из всех `primary_cuts` плана (не только matched по width строки заказа).

**Файлы:** `viz_modules/procurement/trim.py`.

---

### BUG-3: Двойной / необъяснимый отход 560 + 240 мм (multi-secondary на одном слэбе)

**Симптомы (кейс 4):**

- ПБ 56,3-3,2-8п: «Отходы (560мм)» — ожидаемо на primary
- ПБ 46,4-3,2-8п: «Отходы (240мм)» — **неясный источник** (вероятно `sec_cut['waste']` оптимизатора)

**Гипотеза:**

- Rest полосы ≈ 880 мм; продукты 320 + 320 = 640; **остаток/отход 240** попал в `secondary_cuts[].waste`.
- Fix-2 должен убрать начисление на secondary; 240 мм остаётся только в данных плана для визуализации, не в цене дочерней плиты.

**Дополнительно (Fix-3 — observability):**

- В DEBUG/breakdown при `waste_cost > 0` логировать: `waste_mm`, `owner=primary|secondary`, `primary_instance_id`, `sec_cut.id`.
- Тест: snapshot плана кейса 4 → assert waste только на primary.

**Файлы:** `trim.py`, `breakdown.py` (формула/метаданные), тест с fixture.

---

### BUG-4: «План есть, trim не нашёл резов» — silent zero (класс регрессий)

**Симптом:** периодически пропадает продольный рез на **любой** ширине, не только 10,8 м.

**Корневая причина:** `resolve_long_cut_pricing` строки 943–948 — WARNING в лог, цена реза = 0.

**Исправление (Fix-4):**

- Если `has_plan` и `not has_trim_cuts` и `width_m < 1.15` и `width_m != 1.2`:
  - использовать `fallback_long_cuts` из `build_procurement_items` (как при отсутствии плана);
  - добавить warning в metadata позиции (для UI), не silent.
- Мониторинг: считать WARNING в тестах / CI grep.

**Файлы:** `trim.py`, `price_rows.py`.

---

## Матрица регрессионных тестов

| ID | Тест | Кейс | Проверяет |
|----|------|------|-----------|
| T1 | `test_1080_factory_strip_includes_long_cut` | BUG-1, кейсы 1–2 | `long_cut_cost > 0`, `waste_cost` = 120/1200 × base, итог ≈ base_1_2m |
| T2 | `test_1080_qty2_waste_formula_120x2` | BUG-1 | waste_terms `[(120, 2)]`, long cut × 2 / qty |
| T3 | `test_waste_160_only_on_primary_720_320` | BUG-2, кейс 3 | primary waste > 0, secondary waste == 0 |
| T4 | `test_secondary_from_rest_no_strip_waste` | BUG-2 | generic cross-width plan |
| T5 | `test_waste_560_primary_secondary_240_zero` | BUG-3, кейс 4 | fixture snapshot; secondary без waste |
| T6 | `test_plan_exists_unmatched_width_uses_fallback_long_cut` | BUG-4 | расширение существующего теста |
| T7 | `test_price_rows_matches_breakdown_components` | интеграция | sum(breakdown) == unit_price |
| T8 | `test_production_matches_commercial_trim` | оба потока | одинаковый trim для одного заказа |

### Эталонные числа (кейс 3, для sanity-check)

При base_1_2m = 21 515 (63-7,2-8п) и 14 510 (45-3,2-8п):

- Отход 160 мм на **только** 63-7,2: `(160/1200) × 21 515 ≈ 2 868,67 ₽`
- На 45-3,2: **0 ₽** отхода полосы (резы и остаток по длине — отдельно)

---

## Invariants (property-тесты на любой план)

Для каждого `(plan, order_line)`:

1. `sum(w × n for w, n in all_waste_terms_across_all_lines) <= sum(primary_unused_strip_mm)` — отход не раздувается.
2. Каждая пара `(waste_mm, primary_instance_id)` встречается **не более одного раза** в начислениях.
3. Если `1020 <= width_mm <= 1080` и qty > 0 → `long_cut_meterage > 0` или явный warning.
4. `build_price_rows` и `build_price_rows_production` — одинаковый `trim` dict для одной позиции.

---

## Implementation Plan (порядок)

```
1. Тесты T3–T4 (BUG-2) — red
2. Fix-2 в trim.py — green
3. Тесты T1–T2 (BUG-1) — red
4. Fix-1 в resolve_long_cut_pricing — green
5. Fixture кейса 4 + T5 — red/green (Fix-2 часто достаточно)
6. Fix-4 + T6
7. scripts/reconcile_plate_prices.py + JSON эталоны от менеджера
8. T7–T8 интеграция
```

---

## Tasks

- [ ] **Task 1:** Добавить тесты T3, T4 (двойной отход cross-width)
  - Acceptance: падают на текущем main
  - Verify: `pytest tests/test_procurement_trim_cuts.py -k waste -q`
  - Files: `tests/test_procurement_trim_cuts.py`

- [ ] **Task 2:** Fix-2 — `charge_strip_waste=False` для secondary с primary rest
  - Acceptance: T3, T4 green; существующие тесты не ломаются
  - Verify: `pytest tests/test_procurement_trim_cuts.py -q`
  - Files: `viz_modules/procurement/trim.py`

- [ ] **Task 3:** Тесты T1, T2 (1080 мм + продольный рез)
  - Acceptance: падают до Fix-1
  - Files: `tests/test_procurement_trim_cuts.py`

- [ ] **Task 4:** Fix-1 — long cut для factory strip 1020–1080
  - Acceptance: T1, T2 green; R5, R7 выполнены
  - Files: `viz_modules/procurement/trim.py`, `price_rows.py`, `breakdown.py`

- [ ] **Task 5:** Snapshot fixture кейса 4 + T5
  - Acceptance: secondary ПБ 46,4 без waste line
  - Files: `tests/fixtures/plate_pricing/case4_cascade.json`, tests

- [ ] **Task 6:** Fix-4 — fallback long cut при plan-without-match
  - Acceptance: T6 green; WARNING остаётся в лог
  - Files: `trim.py`

- [ ] **Task 7:** `scripts/reconcile_plate_prices.py` + 4 JSON эталона
  - Acceptance: `--strict` exit 0 на эталонах
  - Files: `scripts/`, `tests/fixtures/plate_pricing/`

- [ ] **Task 8:** Интеграционные T7, T8
  - Verify: `pytest tests/ -q -k "procurement or breakdown"`

---

## Boundaries

**Always:**

- Менять trim через TDD: тест → фикс → green
- Прогонять `test_procurement_trim_cuts.py` перед merge
- Сохранять правило «1,2 м без продольного реза»

**Ask first:**

- Изменение формул оптимизатора (`sec_cut['waste']`)
- Изменение `LONG_CUT_PRICE_PER_M` / `TRANSVERSE_CUT_PRICE`
- Пересчёт архивных КП

**Never:**

- Удалять существующие тесты без замены
- Дублировать логику trim в `commercial_offer.py` вместо фикса в одном месте
- Silent fallback 4000×площадь

---

## Success Criteria

| # | Критерий | Проверка |
|---|----------|----------|
| 1 | 4 пользовательских кейса воспроизведены тестами T1–T5 | pytest |
| 2 | Нет двойного отхода на primary+secondary одного слэба | T3, T4, T5 + invariants |
| 3 | 10,8 м всегда с продольным резом в breakdown | T1, T2 |
| 4 | `build_price_rows` == `build_price_rows_production` по trim-компонентам | T8 |
| 5 | Менеджер подписал ≥4 JSON эталона | `reconcile_plate_prices.py --strict` |
| 6 | Существующий suite `test_procurement_trim_cuts.py` — all green | CI |

---

## Open Questions

1. **Кейс 4:** нужен JSON snapshot плана оптимизатора от пользователя — для точной проверки источника 240 мм.
2. **Пропавший продольный рез «без примера»:** собрать логи `[WARNING] План есть, но trim не нашёл резов` за N дней?
3. **240 мм в оптимизаторе:** баг данных (waste должен быть на primary) или корректное поле, которое trim не должен тарифицировать на secondary? → после Fix-2 сверить с менеджером.

---

## Changelog

| Дата | Изменение |
|------|-----------|
| 2026-07-19 | Первая версия: BUG-1..4, Fix-1..4, матрица тестов, tasks |
