---
date: 2026-03-27
topic: порядок плит шире 12 дм в превью и КП
---

# Исследование: Порядок плит шире 12 дм в списках (XLSX превью, КП PDF/XLSX)

## Резюме
Проверен текущий путь данных от ввода списка плит до формирования `xlsx` превью, `КП PDF` и `КП xlsx`. Отдельная логика определения широких плит (`ширина > 12 дм`) используется на этапе подтверждения/замены списка, после чего итоговый список снова парсится и идёт в общую сборку КП. В текущих генераторах порядок строк формируется общей сортировкой по ширине/длине или по порядку `order_data`; отдельного шага «переместить плиты шире 12 дм в конец» в этих точках не обнаружено.

## Подробные находки

**Расположение:** `core/plate_text_normalizer.py:236-339`  
**Что делает:** функция `get_wide_plate_lines()` выделяет строки заказа, где ширина в дм больше 12, поддерживая каталожные марки, канонические строки и формат `W×L`; возвращает пары `(исходная_строка, qty)`.  
**Ключевые зависимости:** `parse_catalog_mark`, регулярные выражения `_CANONICAL_L_W_RE`, `_DIM_RE`, `_QTY_END_RE` внутри модуля.  
**Паттерны:** критерий широких плит задан как `W_dm > 12` или `width_dm > 12.0`; обработка выполняется построчно после нормализации.

**Расположение:** `bot/handlers/commercial.py:214-361`, `bot/handlers/commercial.py:417-474`, `bot/handlers/commercial.py:485-666`  
**Что делает:** в сценарии КП список плит нормализуется, затем широкие строки ищутся через `get_wide_plate_lines`; при наличии широких строк пользователь переводится в шаг замены/пропуска; после замены итоговый `plates_text` сохраняется в state и отправляется в превью.  
**Ключевые зависимости:** `normalize_order_text`, `set_plate_lists_from_text`, `get_wide_plate_lines`, `build_plates_reconciliation_preview_xlsx`.  
**Паттерны:** FSM-поток `waiting_plates_list -> waiting_plates_confirm -> waiting_wide_plates_replacement`; итоговая строка заказа хранится в `plates_text` и далее используется для всех документов.

**Расположение:** `core/plates_preview_xlsx.py:214-342`  
**Что делает:** `build_plates_reconciliation_preview_xlsx()` строит лист `Превью списка` по вкладам строк; физические строки формируются через группировку вкладов, затем записываются в `A-E`.  
**Ключевые зависимости:** `set_plate_lists_from_text`, `cfg.PLATE_LOAD_DETAILS`, `preview_row_keyed_triples_for_contributions`, `make_plate_name`.  
**Паттерны:** порядок вкладов внутри строки задаётся `_sort_contribution_keys()` по `(length_m, width_m, length_dm_raw)` (`core/plates_preview_xlsx.py:93-100`); затем данные группируются `OrderedDict` по ключу `(LineContributionKey, kp_name, kp_qty)` и выводятся в порядке заполнения (`core/plates_preview_xlsx.py:286-305`).

**Расположение:** `viz_modules/procurement.py:91-183`  
**Что делает:** `build_procurement_items()` формирует позиции закупки, из которых затем строятся цены и `order_data` для КП; в ветках с `plan_orders` и `PLATE_LOAD_DETAILS` позиции проходят через `sorted(...)`.  
**Ключевые зависимости:** `get_orders_from_opt_plan`, `cfg.PLATE_LOAD_DETAILS`, `cfg.get_load_code_for_plate`, `cfg.PLATE_LENGTH_DM_RAW`.  
**Паттерны:** сортировка позиций выполняется ключом `(width, length, load)` (`viz_modules/procurement.py:132`, `viz_modules/procurement.py:157`, `viz_modules/procurement.py:222`), что задаёт порядок строк для последующих этапов.

**Расположение:** `bot/handlers/commercial.py:924-1051`, `bot/handlers/commercial.py:1656-1683`  
**Что делает:** обработчик КП получает `procurement_items = build_procurement_items()`, строит `order_data` в том же порядке прохода по `procurement_items`, затем по кнопке генерирует PDF/XLSX через `generate_commercial_offer_pdf` и `generate_commercial_offer_xlsx`, передавая этот `order_data`.  
**Ключевые зависимости:** `build_procurement_items`, `build_price_rows`, `generate_commercial_offer_pdf`, `generate_commercial_offer_xlsx`.  
**Паттерны:** источник порядка для КП — порядок элементов `order_data`, собранного в цикле `for item in procurement_items`.

**Расположение:** `core/commercial_offer.py:501-547`, `core/commercial_offer_xlsx.py:257-295`  
**Что делает:** генераторы КП PDF/XLSX формируют табличные строки по `for idx, item in enumerate(order_data, start=1)` без дополнительной сортировки внутри модулей.  
**Ключевые зависимости:** `order_data` как входной параметр обоих генераторов.  
**Паттерны:** порядок вывода строк в PDF и XLSX повторяет входной порядок `order_data`.

## Ссылки на код

- `core/plate_text_normalizer.py:236` — вход в определение широких плит (`get_wide_plate_lines`).
- `core/plate_text_normalizer.py:300` — условие широкой плиты по каталожному разбору (`W_dm > 12`).
- `core/plate_text_normalizer.py:330` — условие широкой плиты по размерному формату (`width_dm > 12.0`).
- `bot/handlers/commercial.py:320-323` — вычисление `wide_plate_lines` из нормализованного текста.
- `bot/handlers/commercial.py:432-445` — переход в шаг замены при наличии широких плит.
- `bot/handlers/commercial.py:523-545` — сборка итогового списка после замены строк.
- `bot/handlers/commercial.py:568-574` — сохранение итогового списка в state как `plates_text`/`raw_plate_lines`.
- `bot/handlers/commercial.py:92-104` — вызов генерации `xlsx` превью списка плит.
- `core/plates_preview_xlsx.py:93-100` — сортировка вкладов по длине/ширине/`length_dm_raw`.
- `core/plates_preview_xlsx.py:286-305` — группировка и формирование физических строк превью.
- `viz_modules/procurement.py:132` — сортировка `order_counter` по `(width, length, load)` в ветке `plan_orders`.
- `viz_modules/procurement.py:157` — сортировка `PLATE_LOAD_DETAILS` по `(width, length, load)`.
- `bot/handlers/commercial.py:947-1051` — формирование `order_data` из `build_procurement_items`.
- `core/commercial_offer.py:501-547` — вывод строк PDF по порядку `order_data`.
- `core/commercial_offer_xlsx.py:257-295` — вывод строк XLSX по порядку `order_data`.
- `bot/handlers/commercial.py:1659-1677` — передача `order_data` в генераторы PDF/XLSX.

## Архитектурные наблюдения

- Поток данных для КП в текущей реализации: `plates_text` из FSM -> `set_plate_lists_from_text`/`build_procurement_items` -> `order_data` -> `generate_commercial_offer_pdf` и `generate_commercial_offer_xlsx` (`bot/handlers/commercial.py:805-833`, `bot/handlers/commercial.py:947-1051`, `bot/handlers/commercial.py:1659-1677`).
- Поток данных для превью: `plates_text` + `initial_user_plate_lines` -> `build_plates_reconciliation_preview_xlsx` -> лист `Превью списка` (`bot/handlers/commercial.py:81-104`, `core/plates_preview_xlsx.py:214-342`).
- Для широких плит используется отдельный этап в FSM (обнаружение/замена/пропуск), а после завершения этого этапа общий пайплайн генерации документов остаётся единым для всех позиций (`bot/handlers/commercial.py:320-323`, `bot/handlers/commercial.py:432-445`, `bot/handlers/commercial.py:485-666`, `bot/handlers/commercial.py:797-833`).
