# Spec: Соседние марки не воруют резы (trim width match)

> **Тип:** bugfix + regression suite  
> **Фаза SDD:** SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅ шаг 1 (2026-09-03)  
> **Дата:** 2026-09-03  
> **Статус:** шаг 1 в коде (2026-09-03); PLAN: [`../develop/plans/2026-09-03-trim-neighbor-sku-cuts.md`](../develop/plans/2026-09-03-trim-neighbor-sku-cuts.md)  
> **Идея:** [`../ideas/trim-neighbor-sku-cuts.md`](../ideas/trim-neighbor-sku-cuts.md)  
> **Связанные:** [`plate-pricing-trim-bugs.md`](./plate-pricing-trim-bugs.md) (R1–R6, порог 20 мм как отход), [`../ideas/plate-pricing-fix.md`](../ideas/plate-pricing-fix.md)  
> **Эталон:** схема раскладки + ручной расчёт менеджера (ПБ 65-9-8п vs ошибочная ПБ 56-7-8п)

---

## Assumptions (зафиксированы в ideation)

1. КП и производственная смета идут через один `_calc_trim_components` — одно правило на оба потока.
2. В `secondary_cuts[].cuts` оптимизатор пишет **ширину марки-выхода** (`output_width`), а не ширину полосы. Отход ≤ 20 мм живёт в `waste` / `pieces`, не в матчинге SKU.
3. Правило завода «два изделия с одного реза, кромка ≤ 20 мм не режется» уже реализовано в `MIN_BILLABLE_TRIM_MM` и `_longitudinal_cuts_for_rest_secondary`. Его **нельзя** убирать.
4. Допуск **марки** при матчинге `cuts` ↔ ширина строки: **10 мм (1 см на плиту, документация завода)**. Это не кромка 20 мм. `|700−720|=20` — разные марки. `|710−720|=10` — в допуске, шаг 2 из‑за одной этой пары **не** включаем.
5. Шаг 2 (`target_order_key`) — только если после шага 1 рез всё ещё садится на чужую марку с разницей **> 10 мм**, или один `secondary_instance_id` биллится дважды.
6. Исторические КП не пересчитываем.
7. Баг подписи «Остаток … (0,40м)» при сумме 0,40+0,70 — **вне скоупа**.
8. Consume-once в шаге 1 не делаем; в шаге 2 — обязательно.

→ Если пункт 2 неверен в проде — не возвращать допуск 20 мм в `_width_matches_cut`; это вход в шаг 2.

---

## Objective

Убрать ложную склейку **соседних заводских ширин** в расчёте резов: плита ширины W не наследует secondary с `|cuts−W| > 10 мм` (1 см на плиту). Пара 700/720 (20 мм) — разные марки. Кромка ≤ 20 мм по-прежнему не даёт второго продольного реза.

**Пользователи:** менеджер (разбивка КП), производство (смета).

**Симптом (КП 3 / draft `df9b2510…`, 2026-09-03):**

| Марка | Факт в смете (баг) | Должно быть (как 65-9) |
|---|---|---|
| ПБ 56-7-8п (5,6×700) | поперечный ×2, продольный `460 × 6,3 × 1`, остаток 2 147,86 (0,40+0,70) | 1 поперечный, остаток 0,40 м, **без** продольного (он на ПБ 60-5) |
| ПБ 56-7,2-8п (5,6×720) | те же **оба** реза зеркально | свой поперечный 6,3→5,6 и свой продольный 6,3 м при `waste=160` |
| ПБ 65-9-8п (6,5×900) | верно: поперечный ×1, без продольного | без регрессии |

**Почему:** `_width_matches_cut` считает ширины равными при `abs(diff) <= 20`. `|700−720|=20` попадает в этот порог (кромка), хотя допуск **марки** — 10 мм. `sec-5` (`cuts: [700]`) и `sec-9` (`cuts: [720]`) садятся на **обе** строки.

**Успех:** соседние марки не делят чужие `trans_cuts` / `long_cut_meterage`; правило «кромка ≤ 20 мм → нет второго продольного» сохраняется.

---

## Scope

### In scope (шаг 1)

- `viz_modules/procurement/trim.py` — только `_width_matches_cut` (+ комментарий, зачем 20 мм живёт в другом месте)
- Регрессия в `tests/test_procurement_trim_cuts.py`
- Матрица соседних пар и кейс «2 плиты + waste=20 → 1 продольный»

### In scope (шаг 2, условно)

- `viz_modules/procurement/plan_snapshot.py` — не выкидывать `target_order_key` / `secondary_instance_id`
- Матч secondary к строке сметы по ключу `(length, width, load)`; каждый `secondary_instance_id` — не больше одной строки
- Только если шаг 1 + матрица красные

### Out of scope

- `MIN_BILLABLE_TRIM_MM`, `GeometryConfig.tolerance_width`, ILP / `demand_tolerance_width`
- Раскладка, визуализация, UI формулы остатка
- Пересчёт архива, `core/pricing/`
- Потолок «число поперечных ≤ qty»
- Factory strip 1020–1080 (R5 в `plate-pricing-trim-bugs.md`)

---

## Tech Stack

Без новых зависимостей. Python 3, pytest, `viz_modules/procurement/`.

| Константа / допуск | Значение | Роль после фикса |
|---|---|---|
| `MIN_BILLABLE_TRIM_MM` | 20 мм | Отход/кромка: не тарифицировать, не добавлять второй продольный |
| `GeometryConfig.tolerance_width` | 20 мм | Геометрия: rest − target ≤ 20 → тип `transverse`, без добора по ширине |
| `_width_matches_cut` / `PLATE_WIDTH_MATCH_TOLERANCE_MM` | **10 мм** | Допуск марки: 1 см на плиту (документация). Не путать с кромкой 20 мм |
| ILP `demand_tolerance_width` | 10 мм | Не трогаем; совпадает по величине с допуском марки |

---

## Commands

```bash
# Gate шага 1
pytest tests/test_procurement_trim_cuts.py -q

# Соседний контур (не ломаем старые trim-баги)
pytest tests/test_procurement_trim_cuts.py tests/test_procurement_mixed_load_breakdown.py -q

# Узкий прогон по имени после добавления тестов
pytest tests/test_procurement_trim_cuts.py -q -k "neighbor or width_match or waste_20 or 56_7"
```

---

## Project Structure

```
viz_modules/procurement/
  trim.py              ← _width_matches_cut (шаг 1); матч по ключу (шаг 2)
  plan_snapshot.py     ← шаг 2: поля ключа реза
  price_rows.py        ← не менять, если trim достаточен
  breakdown.py         ← не менять (оба потока читают trim)

tests/
  test_procurement_trim_cuts.py   ← регрессия + матрица

ai_docs/ideas/trim-neighbor-sku-cuts.md
ai_docs/specs/trim-neighbor-sku-cuts.md   ← этот документ
```

---

## Code Style

Допуск 20 мм остаётся в **счётчике резов**, не в сравнении марок:

```python
# core/config/constants.py
MIN_BILLABLE_TRIM_MM = 20  # кромка: второй рез не делают
PLATE_WIDTH_MATCH_TOLERANCE_MM = 10  # марка: 1 см на плиту


def _width_matches_cut(width_mm: int, sec_cuts: list) -> bool:
    """Марка сметы ↔ выход secondary. Допуск 10 мм (1 см), не кромка 20 мм."""
    return any(
        abs(int(cut_width) - int(width_mm)) <= PLATE_WIDTH_MATCH_TOLERANCE_MM
        for cut_width in sec_cuts
    )


def _longitudinal_cuts_for_rest_secondary(sec_cut: dict, *, min_one_cut_per_op: bool = False) -> int:
    # без изменений: pieces=2, waste<=20 → один рез между плитами, без кромки
    waste_w_mm = float(sec_cut.get("waste", 0) or 0)
    internal_cuts = (kept_pieces - 1) + (
        1 if waste_w_mm > MIN_BILLABLE_TRIM_MM else 0
    )
```

Имена тестов: `test_<кейс>_<ожидание>` на русском в docstring, как в файле (`test_pb_422_665_321_530_no_double_long_cut`).

План для регрессии 56-7 — минимальный dict как в существующих `_pb_422_…_plan()` / `_foreign_sameload_rest_plan()`, не полный JSON драфта.

---

## Testing Strategy

**Фреймворк:** pytest, `tests/test_procurement_trim_cuts.py`.  
**Уровень:** unit на `_calc_trim_components` с явным `current_plan`. Браузер не нужен.

### Шаг 1 — обязательные кейсы

| ID | План (сжато) | Запрос | Ожидание |
|---|---|---|---|
| T-56-7 | prim 500@6,0 rest 700; sec `cuts[700]` 6,0→5,6 waste 0; sec `cuts[720]` 6,3→5,6 waste 160 | 5,6×700 qty=1 | `trans_cuts==1`, `long_cut_meterage==0`, remainder terms только `(0.4, 1)` |
| T-56-72 | тот же план | 5,6×720 qty=1 | `trans_cuts==1`, `long_cut_meterage==6.3`, remainder `(0.7, 1)`; **не** 0,40 от 700 |
| T-65-9 | prim 300@7,4 rest 900; sec `cuts[900]` 7,4→6,5 waste 0 | 6,5×900 qty=1 | `trans_cuts==1`, `long_cut_meterage==0` (без регрессии) |
| T-waste20 | rest 720, `pieces=2`, `cuts[350]`, `waste=20`, type `multiple` | 350 мм | `_longitudinal_cuts_for_rest_secondary` → 1; в trim не 2 продольных |
| T-665 | существующий каскад `cuts[665]`, rest 670, waste 5 | 2,8×665 | как сейчас: secondary/primary без пропажи реза |
| T-neigh | синтетика: две secondary одной длины, `cuts` 700 и 720 | по очереди 700 и 720 | каждая строка видит только свой `cuts` |

Матрица соседей (синтетика, одна длина, две марки):

| Пара | Δ мм | Ожидание |
|---|---|---|
| 700 / 720 | 20 | **не** делят резы (баг 56-7) |
| 300 / 320 | 20 | **не** делят |
| 480 / 500 | 20 | **не** делят |
| 880 / 900 | 20 | **не** делят |
| 710 / 720 | 10 | **могут** матчиться: в допуске 1 см; не триггер шага 2 |
| 665 / 670 | 5 | матч допустим (каскад / округление) |

Существующие тесты файла — зелёные (в т.ч. 725 мм, 665/530, foreign sameload).

### Шаг 2 — если после шага 1 рез с Δ>10 мм всё ещё чужой или один id биллится дважды

- Фикстура с `target_order_key`
- Матч по ключу; второй SKU не наследует рез
- `secondary_instance_id` не биллится дважды

---

## Boundaries

- **Always:** менять только матчинг марки; гонять `pytest tests/test_procurement_trim_cuts.py` до объявления шага 1 готовым; оставить 20 мм в тарификации отхода и счётчике продольных.
- **Ask first:** включать шаг 2; менять ILP / `demand_tolerance_width`; трогать `breakdown.py` (формула остатка).
- **Never:** возвращать `<= 20` в `_width_matches_cut`; удалять `MIN_BILLABLE_TRIM_MM`; править геометрию «чтобы смета сошлась»; пересчитывать архив; пропускать красные старые trim-тесты.

---

## Success Criteria

Шаг 1 готов, когда:

1. На плане-клоне бага: 5,6×700 → `trans_cuts == 1`, нет метража 6,3 м; 5,6×720 → свой 6,3 м и не забирает 0,40 м от 700.
2. Кейс двух плит с `waste=20` по-прежнему даёт **один** продольный рез, не два и не ноль.
3. Каскад 665/530 и прочий `test_procurement_trim_cuts.py` зелёный.
4. Матрица: пары с Δ=20 мм не делят резы; 710/720 (Δ=10) в допуске марки. Шаг 2 — только при краже с Δ>10 мм или двойном биллинге одного id.

Числа для 56-7 после шага 1 (база 18 745, qty=1): база 10 934,58 + поперечный 1 200 + остаток 0,40 м ≈ 781,04 → **≈ 12 915,62 ₽** (сейчас 18 380,45).

---

## Open Questions

Закрыты 2026-09-03: допуск марки **10 мм**; consume-once не в шаге 1; шаг 2 отдельно и только при краже Δ>10 мм.

---

## Next

PLAN: [`../develop/plans/2026-09-03-trim-neighbor-sku-cuts.md`](../develop/plans/2026-09-03-trim-neighbor-sku-cuts.md)
