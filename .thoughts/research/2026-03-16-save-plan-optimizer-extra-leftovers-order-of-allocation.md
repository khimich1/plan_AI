---
date: 2026-03-16
topic: Падение при сохранении плана — optimizer_extra (leftovers), порядок распределения primary/secondary
---

# Исследование: Падение при сохранении плана (optimizer_extra, порядок распределения)

## Резюме

При сохранении плана распределение плит по заказам выполняется по источникам в фиксированном порядке: сначала primary, затем secondary, затем rescue. Один общий массив `qty_to_mark_by_index` передаётся между вызовами. В воспроизведённом сценарии все primary-плиты закрывают спрос по identity; при распределении secondary по тем же identity остаточный спрос уже нулевой, все вторичные плиты попадают в leftovers. Проверка на непустой `optimizer_extra` выбрасывает исключение «В плане обнаружены плиты, не сопоставленные с заказами». Данные сессии зафиксированы в логе `debug-5b5324.log`.

## Подробные находки

### 1. Точка падения при сохранении

**Расположение:** `bot/handlers/production_export.py:979-1002`  
**Что делает:** Формируется `optimizer_leftovers` из `leftovers_by_source['primary']` и `leftovers_by_source['secondary']`. `optimizer_extra` — непустое подмножество (источники с ненулевыми leftovers). При непустом `optimizer_extra` логируется ошибка и выбрасывается `Exception("В плане обнаружены плиты, не сопоставленные с заказами")`.  
**Ключевые зависимости:** `leftovers_by_source` возвращается из `_distribute_assigned_plates_to_orders`.

### 2. Распределение по источникам и порядок вызовов

**Расположение:** `bot/handlers/production_export.py:395-434` (`_distribute_assigned_plates_to_orders`)  
**Что делает:** Инициализирует `qty_to_mark_by_index = [0] * len(orders_2d)`. В цикле по источникам в порядке `('primary', 'secondary', 'rescue')` для каждого источника вызывается `_allocate_assigned_counts_to_orders(orders_2d, source_counts, qty_to_mark_by_index)`. Возвращённый обновлённый `qty_to_mark_by_index` передаётся в следующий вызов. Результат по каждому источнику — `leftovers` — записывается в `leftovers_by_source[source]`.  
**Паттерны:** Один общий вектор «сколько уже распределено по строке заказа»; совпадение только по identity `(kp_id, plate_name)`.

### 3. Логика выделения «места» под плиты в заказах

**Расположение:** `bot/handlers/production_export.py:362-393` (`_allocate_assigned_counts_to_orders`)  
**Что делает:** Принимает `orders_2d`, счётчики по identity `source_counts` и текущий `qty_to_mark_by_index`. Для каждой строки заказа вычисляет `qty_missing = max(qty_ordered - qty_to_mark_by_index[idx], 0)`. Если `qty_missing <= 0`, строка пропускается. Иначе из `remaining_by_identity` для identity этой строки забирается не более `qty_missing` плит, `qty_to_mark_by_index[idx]` увеличивается на взятую величину. Остаток по identity возвращается как `leftovers`.  
**Паттерны:** Строка заказа получает плиты только по точному совпадению identity; после первичного прохода по primary последующие проходы видят уже обновлённый `qty_to_mark_by_index`.

### 4. Подсчёт плит по источникам до распределения

**Расположение:** `bot/handlers/production_export.py:292-359` (`_count_assigned_plates_for_save`)  
**Что делает:** По `optimization_result['plate_assignments']` считает плиты по полю `source`; identity — через `_make_order_identity(assignment)` (kp_id, plate_name). Записи без валидного identity попадают в `unmapped_assignments_by_source`. RESCUE считается отдельно из `all_tracks_list` по метке «РЕСКЬЮ». Возвращает `assigned_counts_by_source` (по источнику — словарь identity → количество) и `unmapped_assignments_by_source`.

### 5. Свидетельство из лога (сессия 5b5324)

**Источник:** файл `debug-5b5324.log` (NDJSON).

- **H_save_input:** при сохранении в state: `orders_2d_len`: 105, `plate_assignments_len`: 629, `all_tracks_list_len`: 34.
- **H_count:** после подсчёта: `assigned_totals_by_source`: primary 584, secondary 45; `unmapped_counts`: пусто по всем источникам.
- **H_distribute:** после распределения: `leftovers_primary_len`: 0, `leftovers_secondary_len`: 20; в `leftovers_secondary_serialized` перечислены identity (kp_id, plate_name) с ненулевым остатком (в сумме 45 плит по 20 разным identity).
- **H_checks / raise_reason:** запись с `"message": "raise: optimizer_extra (leftovers)"`, `"reason": "optimizer_extra"`; в `optimizer_extra` передан только источник `secondary` с сериализованным списком leftovers.

### 6. Соответствие identity в оптимизаторе и при сохранении

В логе оптимизатора (H_secondary): все 45 вторичных плит имеют `match_type`: "exact"; примеры identity совпадают с теми, что фигурируют в leftovers при сохранении (например, «Плиты ПБ 25,4-3,0-8п», «Плиты ПБ 23-6,65-8п»). То есть identity вторичных плит совпадают с identity строк в orders_2d; при этом при распределении они не находят «свободного» спроса по этим же identity.

## Ссылки на код

- `bot/handlers/production_export.py:287-289` — `_make_order_identity`
- `bot/handlers/production_export.py:292-359` — `_count_assigned_plates_for_save`
- `bot/handlers/production_export.py:362-393` — `_allocate_assigned_counts_to_orders`
- `bot/handlers/production_export.py:395-434` — `_distribute_assigned_plates_to_orders`, цикл по источникам и накопление `qty_to_mark_by_index`
- `bot/handlers/production_export.py:979-987` — формирование `optimizer_leftovers` и `optimizer_extra`
- `bot/handlers/production_export.py:989-1002` — проверка `if optimizer_extra` и выброс исключения

## Архитектурные наблюдения

- Распределение при сохранении — последовательное по источникам: один общий `qty_to_mark_by_index` обновляется сначала для primary, затем для secondary, затем для rescue.
- Primary и secondary в оптимизаторе могут выдавать плиты с одинаковым identity (один и тот же заказ/позиция); при сохранении спрос по identity закрывается в первую очередь primary, затем secondary получает остаток. Если по identity спрос уже полностью закрыт после primary, все плиты secondary по этому identity попадают в leftovers.
- Проверка при сохранении трактует любой непустой leftovers по primary или secondary как ошибку и блокирует сохранение исключением.
