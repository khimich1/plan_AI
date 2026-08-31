# Правила формирования цены плит в КП

## Поток формирования цены

1. В `bot/handlers/commercial.py` вызывается расчёт сметы:
   - загружается прайс через `load_price_table_from_xlsx(...)`;
   - считается смета через `build_price_rows(...)` из `viz_modules/procurement.py`.
2. Из `price_rows` цена строки (`row[7]`) записывается в `order_data[i]["unit_price"]`.
3. В генераторах КП (`core/commercial_offer.py`, `core/commercial_offer_xlsx.py`) используется приоритет:
   - если есть `unit_price` -> берётся она;
   - если нет -> fallback на `get_plate_price(...)`.

Итого: фактическая цена в КП формируется в `viz_modules/procurement.py::build_price_rows`, а PDF/XLSX только отображают её.

## Источник базовой цены

В `build_price_rows(...)` базовая цена полной плиты 1.2 м (`base_price_1_2m`) берётся по цепочке:

1. `get_price(length, load_code, cfg.PRICE_DB_PATH)` из `core/price_db.py`;
2. если не найдено — `find_price_for_plate(price_table, length, load_code)` из `viz_modules/price_utils.py`;
3. если не найдено нигде — `0.0`.

Фактическая база под нужную ширину:

`base_price = base_price_1_2m * (W / 1.2)`

где `W` — ширина плиты в метрах.

### Как считается база для дробной длины

Для плит с дробной длиной (например, 5.98 м, 6.37 м) в этой версии используется **округление длины вверх** при запросе в БД:

1. В `viz_modules/procurement.py::build_price_rows` вызывается:
   - `get_price(L, load_code, cfg.PRICE_DB_PATH, round_up=True)`.
2. В `core/price_db.py::get_price` длина переводится в дм так:
   - `raw_length_dm = length_m * 10`
   - `length_dm = int(math.ceil(raw_length_dm - 1e-9))` при `round_up=True`.
3. После этого ищется точное совпадение `(length_dm, load_code)` в таблице `prices`.
4. Если точной строки нет, берётся ближайшая цена с допуском `±1 дм`:
   - `WHERE ABS(length_dm-?)<=1 ... ORDER BY ABS(length_dm-?) LIMIT 1`.

Примеры:

- `5.91 м -> 59.1 дм -> 60 дм` (округление вверх);
- `6.37 м -> 63.7 дм -> 64 дм`.

Важно: 
- Для **коммерческого КП** fallback на XLSX (`find_price_for_plate`) использует `round(length_m*10)` (банкирское округление).
- Для **производственной сметы** fallback на XLSX (`_find_price_for_plate_production_fallback`) использует `ceil(length_m*10)`, выравнивая со стороной БД.

Подробнее см. [Production Pricing Fallback](features/production-pricing-fallback.md).

## Формула цены 1 шт в КП

В `viz_modules/procurement.py::build_price_rows`:

`unit_price = base_price + long_cut_cost + trans_cut_cost + rest_cost + waste_cost`

Компоненты:

- `base_price` — база, пересчитанная по ширине;
- `long_cut_cost` — продольные резы;
- `trans_cut_cost` — поперечные резы;
- `rest_cost` — стоимость неиспользованных остатков;
- `waste_cost` — стоимость отходов.

## Константы тарифов резки

В `core/config_and_data.py`:

- `LONG_CUT_PRICE_PER_M = 460.0` руб/пог.м;
- `TRANSVERSE_CUT_PRICE = 1200.0` руб/рез.

## Как считаются резы и отходы

### Продольные резы

- Если есть подходящий план оптимизации (`OPT_CASCADING_PLAN_BY_LOAD`), берётся фактическое количество первичных/вторичных резов для этой плиты.
- Если плана нет — fallback:
  - `long_cut_cost = long_cuts * (LONG_CUT_PRICE_PER_M * L)`.
- Жёсткое правило: для ширины 1.2 м продольный рез = 0.

### Поперечные резы

- База: `trans_cut_cost = trans_cuts * TRANSVERSE_CUT_PRICE`.
- При вторичных операциях `trans_cuts` дорассчитывается из плана и пересчитывается стоимость.

### Остатки

Для неиспользованных остатков:

`rest_cost = (unused_rest_total_mm / 1200.0) * base_price_1_2m / qty`

### Отходы

Учитываются:

- отходы по ширине из вторичных резов;
- отходы по длине при поперечной резке;
- отдельное правило для ширин 1020–1080 мм:
  - `extra_waste_mm = 1200 - width_mm`;
  - добавляется стоимость как отход.

## Как цена попадает в документы КП

### PDF

`core/commercial_offer.py`:

- в `calculate_total_cost(...)` и при построении таблицы строк используется `item["unit_price"]` (если есть);
- fallback — `get_plate_price(...)`.

### XLSX

`core/commercial_offer_xlsx.py`:

- тот же приоритет `unit_price` -> fallback `get_plate_price(...)`.

## Скидка и НДС

В `calculate_total_cost(...)` (`core/commercial_offer.py` и `core/commercial_offer_xlsx.py`) используется:

1. `discounted_price = unit_price * (1 - discount_percent / 100)`
2. `item_cost = discounted_price * qty`
3. `total_cost_with_vat = Σ item_cost`
4. `subtotal = total_cost_with_vat / 1.22`
5. `vat_amount = total_cost_with_vat - subtotal`
6. `total_with_vat = total_cost_with_vat`

Т.е. код исходит из того, что `unit_price` уже включает НДС.

## Сохранение в БД

В `core/kp_db.py::save_kp_to_db`:

- суммы считаются через тот же `calculate_total_cost(order_data, discount_percent)`;
- в `KP_offers` сохраняются `subtotal`, `vat_amount`, `total_amount`;
- в `kp_plates` сохраняются `unit_price` и `discounted_price`.

Это выравнивает итог между сметой, документами КП и БД.

