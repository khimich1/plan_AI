---
date: 2026-03-02
topic: Номинал vs реальная длина в названии плиты (69 vs 68,8)
---

# Исследование: Где в коде в названии плиты появляется «69» (номинал) и где «68,8» (реальная длина в марке)

## Резюме

Часть длины в названии плиты (например, «69» или «68,8» в «Плиты ПБ 69-12-8п») формируется в `make_plate_name` в `core/config_and_data.py`. Если в вызов передаётся `length_dm_raw`, в марку подставляется он (номинал из ввода). Если не передаётся — строка длины вычисляется из `length_m` (реальная длина после правила −20 мм): при почти целом значении в дм выводится целое число, иначе — одно знаковое с запятой (например, 6.88 м → «68,8»). В текущем потоке КП смета и `order_data` строятся через `build_price_rows` и цикл по `procurement_items`, где `make_plate_name` вызывается без `length_dm_raw`; номинал «69» в итоговом названии может появиться только после обогащения из `prays_plity` (если в справочнике записано «Плиты ПБ 69-12-8п»).

## Подробные находки

### Правило перевода длины из марки в метры (length_dm_to_m)

**Расположение:** `core/config_and_data.py:39-74`

**Что делает:** Строка длины из марки (например, `"69"` или `"42,6"`) переводится в метры. Если в строке нет запятой/точки — считается номинал в мм, из него вычитается 20 мм, результат переводится в метры (`"69"` → 6.88 м). Если запятая/точка есть — строка разбирается как точное значение в дм и переводится в метры без вычета (`"42,6"` → 4.26 м).

**Следствие для номинала 69:** `length_dm_to_m("69")` возвращает `6.88` (реальная длина). Значение `6.90` для номинала 69 в этом методе не получается.

### Формирование строки длины в названии (make_plate_name)

**Расположение:** `core/config_and_data.py:832-876`

**Что делает:** Собирает строку наименования вида «Плиты ПБ {length_str}-{width_str}-{reinforcement}». Часть длины в марке (`length_str`) задаётся так:

- Если передан непустой `length_dm_raw`: `length_str = str(length_dm_raw).strip().replace('.', ',')` — в марке используется номинал/исходная строка из ввода (например, «69»).
- Если `length_dm_raw` не передан или пуст: `length_dm_val = length_m * 10`; если `abs(length_dm_val - round(length_dm_val)) < 0.01`, то `length_str = str(int(round(length_dm_val)))`, иначе `length_str = f'{length_dm_val:.1f}'.rstrip('0').rstrip('.').replace('.', ',')`.

**Пример для length_m = 6.88 (номинал 69):** `length_dm_val = 68.8`, `round(68.8) = 69`, `abs(68.8 - 69) = 0.2 > 0.01` → берётся ветка с одним знаком после запятой → «68,8». То есть при вызове без `length_dm_raw` в названии попадает реальная длина в марке («68,8»), а не номинал («69»).

**Пример для length_m = 6.90:** `length_dm_val = 69.0`, `abs(69.0 - 69) < 0.01` → `length_str = "69"`. То есть при реальной длине 6.90 в марке выведется «69».

### Вызовы make_plate_name без length_dm_raw (в названии — реальная длина)

**build_price_rows**  
**Расположение:** `viz_modules/procurement.py:336`  
**Что делает:** Для каждой позиции из `build_procurement_items()` вызывается `cfg.make_plate_name(L, W, load_code=load_code)`. Аргумент `length_dm_raw` не передаётся. В `items` из `build_procurement_items()` поля `length_dm_raw` нет (`viz_modules/procurement.py:127-135`, `151-157`). Итог: в строках сметы в названии используется длина, вычисленная из `length_m` (для 6.88 → «68,8»).

**build_production_breakdown (детальная разбивка)**  
**Расположение:** `viz_modules/procurement.py:979`  
**Что делает:** `name = cfg.make_plate_name(length, width_m, load_code=load_code)` без `length_dm_raw`. Аналогично — в названии реальная длина в марке.

**Другие вызовы в procurement.py:** строки 621, 628, 782, 1491, 1498, 1674 — везде `make_plate_name` вызывается только с длиной, шириной и нагрузкой, без `length_dm_raw`.

### Формирование order_data в хендлерах КП

**Расположение:** `bot/handlers/commercial.py:518-572` (основной поток генерации КП) и `bot/handlers/commercial.py:861-926` (альтернативный поток).

**Что делает:** Для каждого элемента из `build_procurement_items()` ищется совпадающая строка в `price_rows` (по длине/ширине/нагрузке). Если найдена — `name = matching_row[1]` (название из сметы). Если не найдена — `name = cfg.make_plate_name(length_m, width_m, load_code=load_code)` без `length_dm_raw`. В обоих случаях название строится без номинала из парсинга: в смете оно уже получено через `make_plate_name(L, W, load_code)` без `length_dm_raw`, а запасной вариант — тот же вызов без `length_dm_raw`. Поле `length_dm_raw` в `order_data` заполняется: при наличии `matching_row` — из названия по regex `r'ПБ\s+([\d,]+)-'` (`bot/handlers/commercial.py:556-557`, `910-911`), иначе — из `length_m` формулой `f'{length_m * 10:.1f}'.rstrip('0').rstrip('.').replace('.', ',')` (`bot/handlers/commercial.py:562`, `915`). Итог: до обогащения в названии в этом потоке — реальная длина в марке (например, «68,8»), а не номинал «69».

### Обогащение из prays_plity — источник номинала «69» в итоговом названии

**Расположение:** `core/kp_db.py:286-327` (`enrich_order_data_with_nomenclature`).

**Что делает:** Для каждого элемента `order_data` по полю `name` выполняется поиск в таблице `prays_plity` (база `pb.db`), колонка «Товар». При нахождении записи `item['name']` заменяется на значение «Товар» из справочника. Вызов выполняется в `bot/handlers/commercial.py:596-597` после формирования `order_data`. Итог: если в `prays_plity` записано «Плиты ПБ 69-12-8п», то после обогащения в названии будет номинал «69», а не «68,8». Если в справочнике записано «Плиты ПБ 68,8-12-8п» — в названии останется «68,8».

### Где сохраняется и откуда читается номинал (length_dm_raw)

**Парсинг ввода:** В `set_plate_lists_from_text` при разборе строки формата «ПБ L-W-8п qty» из марки извлекается `Ldm_str` (например, `"69"`). Он передаётся в `add_items(..., length_dm_raw=Ldm_str.strip())` (`core/config_and_data.py:785`). Значение попадает в `PLATE_LENGTH_DM_RAW` с ключом `(length_rounded, width_rounded, load_code)` (`core/config_and_data.py:584`, `667`).

**Использование в заказе:** В `to_orders_2d()` в структуру для оптимизатора и state попадает поле `length_dm_raw` из `plate_length_dm_raw` (`core/config_and_data.py:337-338`). В `build_procurement_items()` при формировании списка позиций из `PLATE_LOAD_DETAILS` или из плана оптимизации поле `length_dm_raw` в элементы не добавляется. Поэтому при построении сметы и названий в КП номинал из `PLATE_LENGTH_DM_RAW` в вызовы `make_plate_name` не попадает.

**Обратная засылка в БД:** В `core/kp_db.py:235-252` при миграции для записей в `kp_plates` с пустым `length_dm_raw` вызывается `extract_length_dm_raw_from_plate_name(plate_name)`: из сохранённого названия извлекается подстрока длины (например, «69» или «68,8») и записывается в колонку `length_dm_raw`. То есть в БД в `length_dm_raw` оказывается то, что было в названии на момент сохранения (номинал — если название пришло из prays_plity с «69», или реальная длина — если было «68,8»).

### Извлечение подстроки длины из названия (extract_length_dm_raw_from_plate_name)

**Расположение:** `core/config_and_data.py:928-944`

**Что делает:** По regex `r'(?:Плиты\s+)?П[БК]\s*([\d,\.]+)\s*-'` из строки названия извлекается первая числовая часть с запятой/точкой (длина в марке). Примеры: «Плиты ПБ 59,8-12-8п» → «59,8», «ПБ 78-12-8п» → «78». Используется при обратной засылке `length_dm_raw` в `kp_plates` (`core/kp_db.py:241-244`).

## Ссылки на код

- `core/config_and_data.py:39` — `length_dm_to_m`: правило −20 мм для целой длины в марке.
- `core/config_and_data.py:848-856` — в `make_plate_name` выбор между `length_dm_raw` и вычислением из `length_m`.
- `core/config_and_data.py:785` — при парсинге передача `Ldm_str` в `add_items` как `length_dm_raw`.
- `core/config_and_data.py:584`, `667` — запись в `PLATE_LENGTH_DM_RAW`.
- `core/config_and_data.py:337-338` — `length_dm_raw` в `to_orders_2d()`.
- `core/config_and_data.py:928-944` — `extract_length_dm_raw_from_plate_name`.
- `viz_modules/procurement.py:336` — `make_plate_name(L, W, load_code)` в `build_price_rows` без `length_dm_raw`.
- `viz_modules/procurement.py:979` — `make_plate_name` в разбивке без `length_dm_raw`.
- `bot/handlers/commercial.py:545`, `559` — имя в `order_data`: из `matching_row[1]` или `make_plate_name(..., без length_dm_raw)`.
- `bot/handlers/commercial.py:556-557`, `562` — заполнение `length_dm_raw` в `order_data` из названия или из `length_m`.
- `bot/handlers/commercial.py:596-597` — вызов `enrich_order_data_with_nomenclature`.
- `core/kp_db.py:317-318` — подстановка в `item['name']` названия из `prays_plity`.
- `core/kp_db.py:236-252` — обратная засылка `length_dm_raw` в `kp_plates` из `plate_name`.

## Архитектурные наблюдения

- Параметр `length_dm_raw` в `make_plate_name` предусмотрен для вывода номинала в марке, но в коде сметы и формирования `order_data` он нигде не передаётся: везде используется только `length_m`, поэтому в названиях до обогащения выводится реальная длина (например, «68,8»).
- Номинал из ввода («69») сохраняется в `PLATE_LENGTH_DM_RAW` и в `order_data.length_dm_raw`, но не участвует в формировании строки названия в основном потоке КП; он используется при списании для поиска по `length_dm_raw` в `kp_plates`.
- Единственное место, где в итоговом названии позиции может появиться номинал «69», — подстановка из справочника `prays_plity` в `enrich_order_data_with_nomenclature`: значение колонки «Товар» полностью заменяет текущее `item['name']`.

---

## Дополнение (2026-03-03): плита 57,1 записывалась как 57 в КП

**Симптом:** В заказе — «Плиты ПБ 57,1-12-8п» (19+4 шт), в КП — одна строка «Плиты ПБ 57-12-8п» (23 шт). Длина 5.71 м (57,1 дм) отображалась как 57 дм.

**Причина:** В fallback-ветке `build_procurement_items()` (`viz_modules/procurement.py`) при формировании позиций из списков `PLATES_1_2`, `PLATES_0_32` и т.д. использовалось округление **`round(L, 1)`**. В результате длина 5.71 м превращалась в 5.7 м; в `make_plate_name` при `length_m = 5.7` получается `length_dm_val = 57.0` → ветка `branch_001` → `length_str = "57"`.

**Исправление:** В том же fallback заменено `round(L, 1)` на `round(L, 3)` (через локальную функцию `_r`), чтобы сохранять 5.71 и выводить в марке «57,1».

**Ссылки:** `viz_modules/procurement.py:283-320` (fallback по PLATES_*); `core/config_and_data.py:879-885` (`make_plate_name`: условие branch_001 для целого числа дм).
