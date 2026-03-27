---
date: 2026-03-27
topic: Корневые причины двух багов — 0.3 вместо 3.0 и утечка wide в regular
---

# Исследование: корневые причины в коде широких плит

## Резюме

Найдены **две независимые корневые причины**:
1. `make_plate_name` (config_and_data.py:986) special-case'ит 0.3м → "0.3" вместо стандартного дм формата "3".
2. После замены широких плит в хендлере, `plates_text` расширяется (больше строк), а `initial_user_plate_lines` остаётся оригинальным (меньше строк) → индексы `line_contributions[i]` не совпадают с `initial_user_plate_lines[i]`.

## Подробные находки

### Причина #1: почему 0.3 а не 3.0

**Расположение:** `core/config_and_data.py:986-989`

```python
if abs(width_m - 0.3) < 1e-6:
    width_str = '0.3'     # ← метры, а не дециметры!
elif abs(width_m - 0.2) < 1e-6:
    width_str = '0.2'
```

**Поток данных (для ПБ 74.15 → 1.5м → сплит 1.2+0.3):**

1. `canonicalize_plate_line("ПБ 74.15-8Вр1400-30 16")` → `"ПБ 74-15-8п 16"` (plate_text_normalizer.py:199)
2. `parse_pb_width_to_m("15")` → `15.0 / 10.0 = 1.5м` (config_and_data.py:102-115)
3. `add_items(width_m=1.5, ...)` → 1.45 ≤ 1.5 ≤ 1.55 → сплит (config_and_data.py:648)
4. `_record_contribution(line_idx, 7.4, 1.2, 8.0, '74')` (config_and_data.py:671)
5. `_record_contribution(line_idx, 7.4, 0.3, 8.0, '74')` (config_and_data.py:672)
   ↑ ширина записана как **0.3 метра** = 3 дм
6. `make_plate_name(7.4, 0.3, load_code=8)` → "Плиты ПБ 74-**0.3**-8п"
   ↑ спец-кейс на строке 986 → "0.3" вместо стандартного "3" (дм)

**Почему это баг:** Все остальные ширины используют дециметры: "12" = 12дм, "5,3" = 5.3дм, "7" = 7дм. Но 0.3м (= 3дм) выводится в метрах "0.3" вместо "3".

**Фикс:** Удалить спец-кейс для 0.3/0.2. Ширина 0.3м → `round(0.3 * 10, 2) = 3.0` → целое → "3". Обратный парсинг: `parse_pb_width_to_m("3")` → `3.0 / 10 = 0.3м` ✓. Совместимость с форматом "0.3" сохраняется через спец-кейс в `parse_pb_width_to_m`.

---

### Причина #2: утечка wide-сплитов в regular-блок

**Расположение:** `bot/handlers/commercial.py:568-583` и `:645-661`

**Поток данных:**

1. Пользователь присылает 15 строк (12 regular + 3 wide @ 15дм)
2. Бот парсит и нормализует → `plates_text` = 15 строк (config_and_data.py:570)
3. Бот обнаруживает широкие плиты, спрашивает пользователя
4. Пользователь подтверждает замену
5. `merged_lines` = 12 regular + 6 split = **18 строк** (commercial.py:537-542)
6. `final_plates_text` = нормализация merged_lines → **18 строк**
7. `await state.update_data(plates_text=final_plates_text, ...)` — **обновлён** (commercial.py:568-569)
8. `initial_user_plate_lines` — **НЕ обновлён**, по-прежнему 15 строк (commercial.py:579)
9. `build_plates_reconciliation_preview_xlsx(plates_text=final_plates_text, initial_user_plate_lines=...)` (commercial.py:580-583)

**В превью (plates_preview_xlsx.py:292-296):**
```
set_plate_lists_from_text(final_plates_text)  → n = 18 строк
initial_user_plate_lines                      → 15 строк
```

**Результат рассинхронизации:**

| i | initial_user_plate_lines[i] | line_contributions[i] (от final_plates_text) |
|---|---|---|
| 0..11 | ✓ совпадают | ✓ совпадают |
| 12 | ПБ 74.15... (wide) | (7.4, 1.2, 8, '74') — 1-й сплит 74.15 |
| 13 | ПБ 24.15... (wide) | (7.4, 0.3, 8, '74') — 2-й сплит **74.15** ← НЕСОВПАДЕНИЕ |
| 14 | ПБ 25.15... (wide) | (2.4, 1.2, 8, '24') — 1-й сплит **24.15** ← НЕСОВПАДЕНИЕ |
| 15 | ∅ (нет строки) | (2.4, 0.3, 8, '24') — 2-й сплит 24.15 |
| 16 | ∅ | (2.5, 1.2, 8, '25') — 1-й сплит 25.15 |
| 17 | ∅ | (2.5, 0.3, 8, '25') — 2-й сплит 25.15 |

Строки 15-17: `user_cell = ""`, `source_wide = False`, `_contrib_looks_like_wide_split = False` (только 1 вклад, не пара) → **классифицируются как regular** → УТЕЧКА.

Строки 13-14: `source_wide = True` (текст совпадает с wide), но `line_contributions[i]` содержит чужие данные → **неправильный B-столбец** в широком блоке.

**Фикс:** После замены обновлять `initial_user_plate_lines` в state, чтобы совпадали с `final_plates_text`.

---

## Ссылки на код

- `core/config_and_data.py:986-989` — спец-кейс "0.3" / "0.2" в make_plate_name
- `core/config_and_data.py:648-676` — сплит 1.5м → 1.2 + 0.3 в add_items
- `core/config_and_data.py:671-672` — запись вкладов 1.2 и 0.3 через _record_contribution
- `core/config_and_data.py:102-115` — parse_pb_width_to_m: спец-кейс для 0.3/0.2
- `core/plate_text_normalizer.py:199-233` — canonicalize_plate_line: нормализация каталожного формата
- `core/plate_text_normalizer.py:342-382` — normalize_order_text: построчная нормализация (сохраняет кол-во строк)
- `bot/handlers/commercial.py:537-542` — merged_lines: объединение regular + replacement (больше строк)
- `bot/handlers/commercial.py:568-569` — state.update_data: plates_text обновлён, initial — нет
- `bot/handlers/commercial.py:579-583` — call site 2: передаёт final_plates_text + старый initial
- `bot/handlers/commercial.py:645-661` — call site 3 (skip wide): аналогичная проблема
- `core/plates_preview_xlsx.py:292-311` — детект рассинхронизации (warning, но не исправление)
- `core/plates_preview_xlsx.py:99-108` — _format_preview_name (затычка для 0.3→3.0)

## Архитектурные наблюдения

1. `set_plate_lists_from_text` вызывает `normalize_order_text` внутри (строка 570), поэтому передача уже нормализованного текста → двойная нормализация (безвредная, т.к. каноническая форма не матчится `_CATALOG_CORE_RE`).
2. `add_items` при нестандартной ширине "снэпит" к ближайшему допустимому резу (строка 719-741). Например, 0.6м → 0.53м. Это объясняет "5,3" вместо "6,0" для ПБ 21.06.
3. `qty_for_contribution_key` (plates_preview_xlsx.py:116-150) имеет fallback для 1.2/0.3 → ищет запись с ~1.5м. Это корректно обрабатывает количества для сплитов.
4. `_contrib_looks_like_wide_split` (plates_preview_xlsx.py:79-96) — эвристика, которая работает только когда ОБА вклада (1.2 и 0.3) находятся в одной строке. После замены, когда каждый сплит — отдельная строка, эвристика бесполезна.
