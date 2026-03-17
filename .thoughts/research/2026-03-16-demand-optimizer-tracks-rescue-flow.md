---
date: 2026-03-16
topic: Поток данных: спрос → оптимизатор → треки → missing → rescue
---

# Исследование: Поток данных спрос → оптимизатор → треки → missing → rescue

## Резюме

Спрос собирается в `orders_2d` в обработчике и передаётся в `optimize_with_cascading_longitudinal_cuts`. Оптимизатор строит `demand_2d`, добавляет ограничения по ключам спроса (primary+secondary источники), формирует `planned_primary_cuts` из переменных `x_prim`, затем `plate_assignments` (primary и secondary). Результат сохраняется в `OPT_CASCADING_PLAN`. Раскладка строится из `primary_cuts` и `secondary_cuts` плана; подсчёт «have» для rescue идёт по элементам треков с нормализацией ключа к заказам.

## Подробные находки

### 1. Формирование спроса и вызов оптимизатора

**Расположение:** `bot/handlers/production_execution.py:454-466`, `615-617`

**Что делает:** Из выбранных плит (`selected_plates`) собирается список словарей `orders_2d` с полями `length`, `width`, `qty`, `load_code` (через `cfg.normalize_load_code`), а также `reinforcement`, `kp_date`, `customer`, `plate_name`, `kp_id`, `length_dm_raw`. Затем в отдельном потоке вызывается `optimize_with_cascading_longitudinal_cuts(orders_2d=orders_2d)`.

**Ключевые зависимости:** `core.optimization.optimize_with_cascading_longitudinal_cuts`, `cfg.normalize_load_code`, `asyncio.to_thread`.

**Паттерны:** Спрос передаётся списком заказов; оптимизация выполняется в thread pool.

---

### 2. demand_2d и ограничения в оптимизаторе

**Расположение:** `core/optimization.py:455-461`, `850-939`, `997-1031`

**Что делает:** По каждому элементу `orders_2d` считается ключ `(order['length'], order['width'], load_code)`; `load_code` берётся как `order.get('load_code', 800)`. В `demand_2d` накапливается количество по ключу. Далее для каждого ключа спроса собираются источники в виде списка пар `(var_or_expr, assignment_key)`: для первичных опций `assignment_key = (opt['length'], lookup_width, load_code)`, для вторичных — `(output_length, output_width, opt_target_load)`. Источники группируются по `(frozenset(prim_ids), frozenset(sec_id_pieces))`; при совпадении ключа группы список `(var, assignment_key)` расширяется и в группу добавляется `(dk, qty)`. В цикле по группам для каждого `(dk, qty)` из группы отбираются переменные с тем же нормализованным `assignment_key`, что и `dk` (`_norm_key`), и в задачу добавляется ограничение `lpSum(sources_dk) >= qty`. Таким образом, ограничение по каждому ключу спроса выполняется за счёт суммы и первичных, и вторичных переменных.

**Ключевые зависимости:** `pulp`, `primary_options`, `secondary_options`, `x_prim`, `x_sec`, `demand_2d`.

**Паттерны:** Один constraint на каждый ключ спроса в группе; нормализация ключа через `_norm_key` (в т.ч. приведение load_code 8/800 к одному значению).

---

### 3. planned_primary_cuts и plate_assignments

**Расположение:** `core/optimization.py:1567-1588`, `1766-1797`, `1856-1868`

**Что делает:** После решения задачи по каждой первичной опции берётся `qty = int(round(value(x_prim[opt['id']])))`; для каждого `qty > 0` в `planned_primary_cuts` добавляется запись с `assignment_key = (opt['length'], lookup_width, lookup_load_code)`. Вторичные резы собираются из `x_sec` и `opt['pieces']` в `planned_secondary_cuts`. Далее `plate_assignments` формируется в два прохода: сначала по `result['primary_cuts']` (уже отсортированным) каждая плита даёт одну запись с `source: 'primary'` и полями из слотов (kp_id, plate_name и т.д.); затем по `result['secondary_cuts']` для каждого реза добавляется запись с `source: 'secondary'`, длиной и шириной из `cut['lengths'][0]` и `cut['cuts'][0]`.

**Ключевые зависимости:** `value(x_prim[...])`, `primary_cuts` (после сортировки и пост-коррекции), `slot_lists`/`_next_slot_info`, `secondary_cuts`.

**Паттерны:** Количество первичных плит по ключу определяется только значениями `x_prim`; вторичные в `plate_assignments` добавляются отдельно и тоже участвуют в общем счёте записей.

---

### 4. Сохранение плана и построение раскладки

**Расположение:** `bot/handlers/production_execution.py:816`, `860-862`, `917`  
**Расположение:** `viz_modules/layout_sequence.py:163`, `408-411`, `414`, `507-510`  
**Расположение:** `core/visualization.py:56` (split_sequence_into_tracks)

**Что делает:** После получения результата оптимизации выполняется `optimization.OPT_CASCADING_PLAN = optimization_result` и при наличии разбиения по нагрузкам — `OPT_CASCADING_PLAN_BY_LOAD = {'all': optimization_result}`. Далее вызывается `build_layout_sequence()` без аргументов; функция читает данные из `core.optimization.OPT_CASCADING_PLAN` (или из `OPT_CASCADING_PLAN_BY_LOAD`). Если в плане есть `plate_assignments`, используется флаг `use_2d_data`. Последовательность строится из `OPT_CASCADING_PLAN.get('primary_cuts', [])`: целые плиты и плиты с остатком обрабатываются отдельно; к элементам последовательности при необходимости подмешиваются вторичные резы из `secondary_cuts_info`. Результат `build_layout_sequence()` передаётся в `split_sequence_into_tracks(seq)`; возвращается список треков, каждый с полем `items`.

**Ключевые зависимости:** `OPT_CASCADING_PLAN`, `OPT_CASCADING_PLAN_BY_LOAD`, `primary_cuts`, `secondary_cuts`, при наличии — `plate_assignments`.

**Паттерны:** Раскладка и треки строятся из одного и того же плана, записанного в глобальные переменные модуля оптимизации.

---

### 5. order_counts, track_counts, missing_counts и rescue

**Расположение:** `bot/handlers/production_execution.py:1342-1367`, `1378-1386`, `1503-1508`  
**Расположение:** `bot/handlers/production_execution.py:1172-1238` (_count_tracks_for_rescue)

**Что делает:** По списку заказов (тот же, что шёл в оптимизатор) строится `raw_order_counts`: ключ `(L, W_canon, load_code)` с `L = round(float(order.get('length', 0)), 2)`, `W_canon = int(round(float(W)))`, `load_code = cfg.normalize_load_code(order.get('load_code', 8))`. Затем вызывается `_merge_to_canonical_order_keys(raw_order_counts, tol_len=0)` — получаются `order_counts` и `canonical_key_fn`; `order_keys = list(order_counts.keys())`. Функция `_count_tracks_for_rescue(all_tracks_list, order_keys, order_counts)` обходит все треки и все `items` в них; для каждого элемента по полям `length`/`target_length` (для transverse), `width`/`main_w`, `load_code` формируется ключ, который нормализуется через `_normalize_key_to_orders(..., order_keys, order_counts, remaining=remaining)`. Если нормализованный ключ есть в `remaining` и счётчик положительный, он уменьшается и увеличивается `counts[norm_key]`. Учитываются и вложенные `secondary_cuts` элемента. По итогу `track_counts` — это сколько раз каждый ключ заказа был «закрыт» элементами треков (с учётом порядка и конкурирующих ключей). Далее для каждого ключа из `order_counts` вычисляется `qty_have = track_counts.get(key, 0)`; если `qty_have < qty_need`, в `missing_counts[key]` пишется `qty_need - qty_have`. При непустом `missing_counts` вызывается `_create_rescue_tracks(missing_counts, info_map)` и полученные треки добавляются в `all_tracks_list`.

**Ключевые зависимости:** заказы из контекста обработчика, `all_tracks_list` (результат `split_sequence_into_tracks`), `_normalize_key_to_orders`, `_to_width_mm`, `cfg.normalize_load_code`.

**Паттерны:** «Have» для rescue — это не сырой счёт записей в plate_assignments, а счёт по элементам треков с приведением ключа к множеству заказов и с учётом оставшегося спроса (`remaining`).

---

### 6. Сравнение входа и выхода оптимизатора (только primary)

**Расположение:** `bot/handlers/production_execution.py:651-676`, `684-714`

**Что делает:** Вычисляется `input_plates = sum(p['qty'] for p in orders_2d)` и `output_plates_primary = sum(1 for p in all_assignments if p.get('source') == 'primary')`. Если они не совпадают, пишется ошибка в лог и строится разница между заказанным и произведённым по ключам только по первичным: `produced = Counter(_plate_key(p) for p in all_assignments if p.get('source') == 'primary')`, `missing = ordered - produced`. В лог H2_optimizer_output выводятся `missing_primary_keys` и `extra_primary_keys` — разница между заказанным и произведённым именно первичными резами.

**Ключевые зависимости:** `orders_2d`, `optimization_result.get('plate_assignments', [])`, `cfg.normalize_load_code`.

**Паттерны:** Контроль «вход/выход» завязан на совпадение общего числа первичных плит; вторичные в этом сравнении не участвуют.

---

## Ссылки на код

- `bot/handlers/production_execution.py:454-466` — формирование `orders_2d`
- `bot/handlers/production_execution.py:615-617` — вызов `optimize_with_cascading_longitudinal_cuts`
- `bot/handlers/production_execution.py:651-654` — подсчёт `output_plates_primary` только по `source == 'primary'`
- `bot/handlers/production_execution.py:816` — присвоение `OPT_CASCADING_PLAN = optimization_result`
- `bot/handlers/production_execution.py:860-862` — импорт и вызов `build_layout_sequence()`
- `bot/handlers/production_execution.py:917` — `all_tracks_list = split_sequence_into_tracks(seq)`
- `bot/handlers/production_execution.py:1172-1238` — `_count_tracks_for_rescue`: обход треков и items, нормализация ключа, учёт remaining и secondary_cuts
- `bot/handlers/production_execution.py:1342-1367` — построение `raw_order_counts`, `order_counts`, `order_keys`
- `bot/handlers/production_execution.py:1378` — `track_counts = _count_tracks_for_rescue(all_tracks_list, order_keys, order_counts)`
- `bot/handlers/production_execution.py:1382-1386` — расчёт `missing_counts` из `order_counts` и `track_counts`
- `bot/handlers/production_execution.py:1503-1508` — создание и добавление rescue-треков при непустом `missing_counts`
- `core/optimization.py:455-461` — построение `demand_2d` из `orders_2d`
- `core/optimization.py:850-939` — сбор источников с `assignment_key`, группировка в `sources_to_demands`
- `core/optimization.py:997-1031` — добавление ограничений по каждому `(dk, qty)` с фильтром `sources_dk` по `_norm_key(ak) == _norm_key(dk)`
- `core/optimization.py:1567-1588` — формирование `planned_primary_cuts` из `value(x_prim[...])`
- `core/optimization.py:1768-1797` — формирование `plate_assignments` из primary_cuts
- `core/optimization.py:1858-1868` — добавление в `plate_assignments` записей из secondary_cuts
- `core/optimization.py:63` — объявление `OPT_CASCADING_PLAN`
- `viz_modules/layout_sequence.py:163` — `build_layout_sequence()` без аргументов
- `viz_modules/layout_sequence.py:408-414` — чтение `OPT_CASCADING_PLAN`, проверка `primary_cuts` и `plate_assignments`
- `viz_modules/layout_sequence.py:507-510` — использование `all_primary_cuts = OPT_CASCADING_PLAN.get('primary_cuts', [])`
- `core/visualization.py:56` — определение `split_sequence_into_tracks`

## Архитектурные наблюдения

- Спрос в оптимизаторе задаётся одним ограничением на ключ: сумма (primary + secondary) переменных по этому ключу ≥ qty. Количество первичных резов по ключу нигде отдельно не ограничивается.
- В обработчике проверка «запрошено vs получено» и логирование недостачи (H2_optimizer_output) считают только первичные записи в `plate_assignments`; вторичные в этом сравнении не участвуют.
- Раскладка (sequence и треки) строится из `primary_cuts` и `secondary_cuts` плана; в треках учитываются как основные размеры элемента (length/width/load_code или target_length для transverse), так и вложенные `secondary_cuts`.
- Подсчёт «have» для rescue (`_count_tracks_for_rescue`) идёт по элементам треков с нормализацией ключа к множеству заказов и с учётом оставшегося спроса по ключам; один и тот же элемент может засчитаться только в один ключ заказа в рамках одного прохода.
