---
date: 2026-03-22
topic: список плит и XLSX-превью сверки
---

# Исследование: Список плит и XLSX-превью сверки

## Резюме
Список из сообщения и файл `Превью_списка_плит_...xlsx` формируются в сценарии создания КП в `bot/handlers/commercial.py`. Источник данных зависит от типа ввода: текст проходит через нормализацию и парсер заказа, фото сначала проходит OCR через `core/ocr_gpt.py`, после чего распознанный текст отправляется в тот же парсер (`bot/handlers/commercial.py:214-357`, `core/ocr_gpt.py:37-139`, `core/config_and_data.py:536-914`).

Для показа пользователю используется нормализованный текст заказа, а для XLSX-превью используется карта вкладов по строкам (`line_contributions`) и глобальная карта итоговых количеств `PLATE_LOAD_DETAILS`, которые создаёт `set_plate_lists_from_text()` (`bot/handlers/commercial.py:316-357`, `core/config_and_data.py:582-590`, `core/config_and_data.py:753-766`). Если ширина плиты больше 12 дм, сценарий переводит пользователя на шаг замены, а итоговый список после замены снова нормализуется, пересчитывается и отправляется в превью (`bot/handlers/commercial.py:432-470`, `bot/handlers/commercial.py:485-589`, `core/plate_text_normalizer.py:236-319`).

## Подробные находки

### 1. Точка входа сценария списка плит

**Расположение:** `bot/handlers/commercial.py:214-357`

**Что делает:**  
Обработчик `receive_plates_list()` принимает шаг 1 создания КП. Он поддерживает два источника данных:
- фото: вызывает OCR, получает текст и список плит из результата распознавания (`bot/handlers/commercial.py:226-259`);
- текст: берёт текст сообщения, разбивает его на строки и сразу парсит (`bot/handlers/commercial.py:277-301`).

После этого обработчик:
- нормализует текст через `normalize_order_text()` (`bot/handlers/commercial.py:316-318`);
- ищет строки с шириной больше 12 дм через `get_wide_plate_lines()` (`bot/handlers/commercial.py:320-322`);
- формирует текст сообщения для Telegram из `plates_text_to_store` (`bot/handlers/commercial.py:324-340`);
- сохраняет в FSM данные `plates_text`, `recognized_count`, `wide_plate_lines`, `initial_user_plate_lines`, `ocr_plates_snapshot`, `ocr_raw_text` и другие поля (`bot/handlers/commercial.py:342-354`).

**Ключевые зависимости:**  
`core.plate_text_normalizer.normalize_order_text`, `core.plate_text_normalizer.get_wide_plate_lines`, `core.config_and_data.set_plate_lists_from_text`, `core.config_and_data.get_current_plate_order`, `core.ocr_gpt.recognize_text_smart` (`bot/handlers/commercial.py:218`, `bot/handlers/commercial.py:238-239`, `bot/handlers/commercial.py:285-286`).

**Паттерны:**  
Сценарий хранит промежуточное состояние в FSM и всегда переводит пользовательский ввод в единый нормализованный текст до отображения и сохранения (`bot/handlers/commercial.py:317-318`, `bot/handlers/commercial.py:342-355`).

### 2. Ветка фото: как список появляется после OCR

**Расположение:** `core/ocr_gpt.py:37-139`, `bot/handlers/commercial.py:226-259`

**Что делает:**  
При фото обработчик вызывает `recognize_text_smart(..., force_gpt=True)`, то есть сразу использует GPT-4o Vision (`bot/handlers/commercial.py:238-239`). OCR возвращает структуру с полями:
- `text`: многострочный текст для парсера, где каждая строка имеет вид `"{name} {qty}"` (`core/ocr_gpt.py:117-119`);
- `plates`: список словарей `{"name": ..., "qty": ...}` (`core/ocr_gpt.py:56-62`, `core/ocr_gpt.py:126-132`).

Для GPT используется промпт, который требует:
- копировать марки посимвольно;
- не округлять длину и ширину;
- не группировать повторяющиеся строки;
- выводить JSON-массив по строкам сверху вниз (`core/ocr_gpt.py:230-299`).

После OCR обработчик:
- кладёт полный текст в `plates_text` (`bot/handlers/commercial.py:242-243`);
- делает снимок исходных OCR-строк в `initial_user_plate_lines` (`bot/handlers/commercial.py:244-248`);
- сохраняет снимок распознанных плит в `ocr_plates_snapshot` (`bot/handlers/commercial.py:249`);
- считает общее число плит как сумму `qty` по OCR-списку (`bot/handlers/commercial.py:251`).

**Ключевые зависимости:**  
`recognize_with_gpt_vision()`, `parse_gpt_response()` (`core/ocr_gpt.py:142-221`, `core/ocr_gpt.py:302-354`).

**Паттерны:**  
Для фото текст сначала приводится к формату строк для парсера, а затем обрабатывается тем же пайплайном, что и обычный текст (`core/ocr_gpt.py:117-119`, `bot/handlers/commercial.py:316-318`).

### 3. Нормализация текста перед парсером

**Расположение:** `core/plate_text_normalizer.py:199-363`, `core/config_and_data.py:565-581`

**Что делает:**  
`set_plate_lists_from_text()` сначала пытается прогнать вход через `normalize_order_text()` (`core/config_and_data.py:565-576`). Нормализатор:
- заменяет юникодные тире на обычный дефис и знаки умножения на `x` (`core/plate_text_normalizer.py:89-97`);
- исправляет OCR-ошибки в префиксе плит (`ПВ/ПГ/ПЕ` → `ПБ`) и добавляет пробел в `ПБ56` → `ПБ 56` (`core/plate_text_normalizer.py:100-114`);
- преобразует каталожный формат `ПБ 59.12-8Вр1400-25` в канонический `ПБ 59-12-8п` (`core/plate_text_normalizer.py:117-170`, `core/plate_text_normalizer.py:199-233`);
- возвращает итоговый нормализованный текст как набор строк (`core/plate_text_normalizer.py:322-363`).

**Ключевые зависимости:**  
`parse_catalog_mark()`, `canonicalize_plate_line()` (`core/plate_text_normalizer.py:117-170`, `core/plate_text_normalizer.py:199-233`).

**Паттерны:**  
Нормализатор работает построчно и сохраняет список `normalized_lines`, чтобы последующий парсер и превью использовали ту же структуру строк (`core/plate_text_normalizer.py:339-363`).

### 4. Основной парсер: из чего строится список

**Расположение:** `core/config_and_data.py:536-914`

**Что делает:**  
`set_plate_lists_from_text()` очищает глобальные структуры заказа (`core/config_and_data.py:563`, `core/config_and_data.py:476-501`), разбивает текст на строки (`core/config_and_data.py:579-584`) и затем для каждой строки пытается распознать один из двух форматов:

1. Размерный формат `ширина x длина [кол-во]`  
   Регулярное выражение ищет два числа и необязательное количество (`core/config_and_data.py:796-803`).  
   Если первое число похоже на длину, а второе на ширину, значения меняются местами (`core/config_and_data.py:804-809`).  
   Затем выполняется валидация размеров и количества (`core/config_and_data.py:771-790`, `core/config_and_data.py:811-817`).

2. Формат марки `ПБ/ПК L-W-Nп [qty]`  
   Сначала ищется вариант со словом `плита/плиты`, потом без него (`core/config_and_data.py:820-827`).  
   Длина берётся через `length_dm_to_m()`, ширина через `parse_pb_width_to_m()` (`core/config_and_data.py:831-836`).  
   Количество ищется после нагрузки или в конце строки (`core/config_and_data.py:841-856`).  
   Нагрузка извлекается отдельно и может быть дробной, например `12,5` (`core/config_and_data.py:858-876`).

Если строка не подошла ни под один шаблон, она попадает в `unparsed_lines` (`core/config_and_data.py:888-890`).

**Ключевые зависимости:**  
`length_dm_to_m()`, `parse_pb_width_to_m()`, `normalize_order_text()` (`core/config_and_data.py:39-72`, `core/config_and_data.py:102-115`, `core/config_and_data.py:565-576`).

**Паттерны:**  
Парсер одновременно заполняет:
- физические списки по ширинам `PLATES_*`;
- карту итогов `PLATE_LOAD_DETAILS`;
- карту вкладов строк `line_contributions`;
- построчные количества `line_plate_load_details` (`core/config_and_data.py:582-590`, `core/config_and_data.py:753-766`).

### 5. Правила разложения по ширинам

**Расположение:** `core/config_and_data.py:601-767`

**Что делает:**  
Функция `add_items()` решает, в какой список попадёт каждая позиция. Правила в коде такие:
- ширина `1.45..1.55` м считается плитой 1.5 м и раскладывается на две физические позиции: `1.2` м и `0.3` м (`core/config_and_data.py:647-676`);
- стандартные диапазоны раскладываются в списки `PLATES_1_2`, `PLATES_1_0`, `PLATES_1_08`, `PLATES_0_32`, `PLATES_0_46`, `PLATES_0_70`, `PLATES_0_72`, `PLATES_0_86`, `PLATES_0_48`, `PLATES_0_50`, `PLATES_0_88` (`core/config_and_data.py:681-718`);
- если ширина не попадает в диапазон, применяется правило "берём меньший рез": выбирается максимальная стандартная ширина, не превышающая фактическую, и `add_items()` вызывается повторно уже с этой шириной (`core/config_and_data.py:719-741`).

При добавлении:
- длина кладётся в нужный список по одной штуке на каждую плиту (`core/config_and_data.py:743-751`);
- точная ширина сохраняется в `PLATE_EXACT_WIDTHS` (`core/config_and_data.py:748-751`);
- нагрузка и исходный код длины сохраняются в `PLATE_LOAD_DETAILS` и `PLATE_LENGTH_DM_RAW` (`core/config_and_data.py:753-759`);
- вклад строки фиксируется в `line_contributions` и `line_plate_load_details` (`core/config_and_data.py:760-766`).

**Ключевые зависимости:**  
Глобальные структуры `PLATES_*`, `PLATE_LOAD_DETAILS`, `PLATE_EXACT_WIDTHS`, `PLATE_LENGTH_DM_RAW` (`core/config_and_data.py:119-190`).

**Паттерны:**  
Код разделяет "физические списки ширин" и "источник правды по итогам/нагрузкам", где главным словарём является `PLATE_LOAD_DETAILS` (`core/config_and_data.py:166-170`).

### 6. Почему в сообщении показывается именно такой список

**Расположение:** `bot/handlers/commercial.py:316-357`, `bot/handlers/commercial.py:547-577`

**Что делает:**  
После парсинга обработчик строит текст для пользователя не из исходного сообщения, а из `plates_text_to_store`, то есть из нормализованного текста (`bot/handlers/commercial.py:316-319`). Далее:
- если длина текста больше лимита Telegram, он обрезается (`bot/handlers/commercial.py:324-329`);
- для текстового ввода бот показывает:
  `Количество в заказе: X. Распознано: X. Одинаковое: да.` (`bot/handlers/commercial.py:336-340`);
- число `X` берётся как сумма значений `get_current_plate_order().plate_load_details.values()` после вызова `set_plate_lists_from_text()` (`bot/handlers/commercial.py:283-287`);
- после замены широких плит сообщение строится тем же способом, но уже по `final_plates_text` и `final_count` (`bot/handlers/commercial.py:542-566`, `bot/handlers/commercial.py:577`).

**Ключевые зависимости:**  
`get_current_plate_order()`, `PlateOrder.plate_load_details` (`bot/handlers/commercial.py:286`, `core/config_and_data.py:440-471`).

**Паттерны:**  
Текст сообщения опирается на уже нормализованный и, при необходимости, заменённый список, а не на исходный пользовательский ввод (`bot/handlers/commercial.py:316-319`, `bot/handlers/commercial.py:544-545`).

### 7. Проверка широких плит и шаг "Заменить"

**Расположение:** `core/plate_text_normalizer.py:236-319`, `bot/handlers/commercial.py:417-470`, `bot/handlers/commercial.py:485-589`

**Что делает:**  
`get_wide_plate_lines()` ищет строки, где ширина больше 12 дм. Он умеет распознавать:
- каталожные марки `ПБ L.W-load`;
- канонические марки `ПБ L-W-Nп`;
- размерный формат `W×L` (`core/plate_text_normalizer.py:236-319`).

Если такие строки найдены, `confirm_plates_list_callback()` не пускает дальше к выбору менеджера. Вместо этого:
- сохраняет флаг, что превью ещё не отправлено (`bot/handlers/commercial.py:432-433`);
- отправляет список широких строк и пример замены (`bot/handlers/commercial.py:434-444`);
- переводит FSM в `waiting_wide_plates_replacement` (`bot/handlers/commercial.py:436`).

На шаге замены:
- новый список замен валидируется тем же парсером (`bot/handlers/commercial.py:499-513`);
- замены нормализуются (`bot/handlers/commercial.py:519-521`);
- строки из исходного списка с шириной >12 дм заменяются строками из списка замен по порядку (`bot/handlers/commercial.py:523-540`);
- итог снова нормализуется, пересчитывается и отправляется в сообщение и XLSX-превью (`bot/handlers/commercial.py:542-584`).

**Ключевые зависимости:**  
`_build_wide_plates_replacement_example()`, `get_wide_plate_lines()` (`bot/handlers/commercial.py:372-414`, `core/plate_text_normalizer.py:236-319`).

**Паттерны:**  
Широкие плиты обрабатываются как отдельный сценарий перед переходом к следующему шагу FSM (`bot/handlers/commercial.py:432-445`, `bot/handlers/commercial.py:575-589`).

### 8. Когда и как создаётся файл `Превью_списка_плит_...xlsx`

**Расположение:** `bot/handlers/commercial.py:81-115`, `bot/handlers/commercial.py:447-470`, `bot/handlers/commercial.py:579-584`

**Что делает:**  
Файл создаётся функцией `_send_plates_preview_xlsx()`. Она:
- собирает имя файла `Превью_списка_плит_{user_id}_{timestamp}.xlsx` (`bot/handlers/commercial.py:94-98`);
- вызывает `build_plates_reconciliation_preview_xlsx()` в отдельном потоке (`bot/handlers/commercial.py:99-104`);
- отправляет файл пользователю с подписью `📊 Сверка строк: ввод → распознано → как в КП` (`bot/handlers/commercial.py:105-108`).

В обычном сценарии файл отправляется при первом нажатии `✅ Подтвердить`, если нет широких плит и превью ещё не отправлялось (`bot/handlers/commercial.py:447-470`). После сценария замены широких плит файл отправляется сразу же, без второго нажатия (`bot/handlers/commercial.py:579-584`).

**Ключевые зависимости:**  
`core.plates_preview_xlsx.build_plates_reconciliation_preview_xlsx()` (`bot/handlers/commercial.py:92`, `bot/handlers/commercial.py:100-104`).

**Паттерны:**  
Отправка превью управляется флагом `plates_preview_sent` в FSM (`bot/handlers/commercial.py:423`, `bot/handlers/commercial.py:433`, `bot/handlers/commercial.py:455`, `bot/handlers/commercial.py:472`).

### 9. Правила построения строк в XLSX-превью

**Расположение:** `core/plates_preview_xlsx.py:3-22`, `core/plates_preview_xlsx.py:214-341`

**Что делает:**  
Модульный докстринг описывает правила листа `Превью списка`:
- колонка A: "как прислал пользователь";
- B-C: "распознано" как наименование и количество по строке;
- D-E: "как в КП" как наименование и глобальное количество по всему заказу (`core/plates_preview_xlsx.py:5-21`).

Функция `build_plates_reconciliation_preview_xlsx()`:
- повторно парсит `plates_text` через `set_plate_lists_from_text()` и получает `line_contributions` и `line_plate_load_details` (`core/plates_preview_xlsx.py:232-233`);
- берёт глобальные итоги из `cfg.PLATE_LOAD_DETAILS` (`core/plates_preview_xlsx.py:233`);
- использует `split_plate_text_lines()` для нормализованных строк (`core/plates_preview_xlsx.py:235`, `core/reconciliation_xlsx.py:33-37`);
- работает по максимальной длине из нормализованных строк, вкладов и построчных словарей (`core/plates_preview_xlsx.py:236-243`);
- сопоставляет строки пользователя из `initial_user_plate_lines` с внутренними строками превью (`core/plates_preview_xlsx.py:245-264`).

Дальше функция:
- превращает вклад каждой строки в набор физических строк через `preview_row_keyed_triples_for_contributions()` (`core/plates_preview_xlsx.py:267-269`);
- группирует строки по `(key, kp_name, kp_qty)`, чтобы D-E показывались только в первой строке группы, а дальше оставались пустыми (`core/plates_preview_xlsx.py:286-305`);
- добавляет в конец строки без распознанного вклада, чтобы не потерять связь с вводом (`core/plates_preview_xlsx.py:306-308`);
- записывает заголовки и все физические строки в Excel (`core/plates_preview_xlsx.py:310-341`).

**Ключевые зависимости:**  
`preview_row_keyed_triples_for_contributions()`, `qty_for_contribution_key()`, `split_plate_text_lines()` (`core/plates_preview_xlsx.py:154-198`, `core/plates_preview_xlsx.py:56-90`, `core/reconciliation_xlsx.py:33-37`).

**Паттерны:**  
XLSX строится не по простому списку строк, а по ключам вкладов `LineContributionKey`, которые содержат длину, ширину, нагрузку и исходную строку длины из марки (`core/config_and_data.py:531-533`, `core/plates_preview_xlsx.py:154-198`).

### 10. Как вычисляются "распознано" и "как в КП" внутри превью

**Расположение:** `core/plates_preview_xlsx.py:56-90`, `core/plates_preview_xlsx.py:154-198`

**Что делает:**  
`qty_for_contribution_key()` ищет количество по ключу `(length_m, width_m, load_code, length_dm_raw)` в словаре количеств (`core/plates_preview_xlsx.py:56-76`).  
Отдельное правило есть для раскола 1.5 м: если ключ относится к ширине около `1.2` или `0.3`, а прямого совпадения нет, функция ищет исходную запись с шириной около `1.5` (`core/plates_preview_xlsx.py:78-88`).

`preview_row_keyed_triples_for_contributions()`:
- берёт уникальные ключи вклада;
- сортирует их по длине, ширине и `length_dm_raw` (`core/plates_preview_xlsx.py:162-170`, `core/plates_preview_xlsx.py:93-100`);
- строит каноническое имя через `make_plate_name()` (`core/plates_preview_xlsx.py:183-194`);
- считает `q_line` по построчному словарю и `q_global` по глобальному словарю (`core/plates_preview_xlsx.py:195-197`).

Именно эти значения попадают в B-C и D-E (`core/plates_preview_xlsx.py:274-304`).

**Ключевые зависимости:**  
`make_plate_name()`, `LineContributionKey`, `PLATE_LOAD_DETAILS` (`core/config_and_data.py:938-1000`, `core/config_and_data.py:531-533`, `core/plates_preview_xlsx.py:233`).

**Паттерны:**  
Одна логическая строка пользователя может стать несколькими физическими строками превью, если у неё несколько вкладов (`core/plates_preview_xlsx.py:223-225`, `core/plates_preview_xlsx.py:274-304`).

## Ссылки на код

- `bot/handlers/commercial.py:214` — вход в шаг приёма списка плит.
- `bot/handlers/commercial.py:226-259` — ветка обработки фото через OCR.
- `bot/handlers/commercial.py:277-301` — ветка обработки текстового списка.
- `bot/handlers/commercial.py:316-319` — выбор нормализованного текста для показа и хранения.
- `bot/handlers/commercial.py:336-340` — текст блока `Количество в заказе / Распознано / Одинаковое`.
- `bot/handlers/commercial.py:417-470` — логика кнопки `✅ Подтвердить`.
- `bot/handlers/commercial.py:485-589` — сценарий замены широких плит и построения итогового списка.
- `bot/handlers/commercial.py:81-115` — создание и отправка `Превью_списка_плит_...xlsx`.
- `bot/keyboards.py:618-627` — inline-кнопки `✅ Подтвердить` и `🔄 Заменить`.
- `core/ocr_gpt.py:117-119` — преобразование OCR-результата в текст для парсера.
- `core/ocr_gpt.py:230-299` — правила GPT OCR: копировать марки посимвольно и не группировать строки.
- `core/plate_text_normalizer.py:199-233` — нормализация одной строки заказа.
- `core/plate_text_normalizer.py:236-319` — поиск строк с шириной больше 12 дм.
- `core/plate_text_normalizer.py:322-363` — нормализация всего текста заказа.
- `core/config_and_data.py:536-914` — основной парсер заказа и заполнение структур списка.
- `core/config_and_data.py:601-767` — правила распределения плит по ширинам и особая обработка 1.5 м.
- `core/config_and_data.py:938-1000` — формирование канонического имени плиты.
- `core/reconciliation_xlsx.py:33-37` — разбиение текста на строки для сверки.
- `core/plates_preview_xlsx.py:3-22` — документированные правила колонок A-E в превью.
- `core/plates_preview_xlsx.py:214-341` — сборка листа `Превью списка`.
- `core/plates_preview_xlsx.py:56-90` — вычисление количества по ключу вклада.
- `core/plates_preview_xlsx.py:154-198` — преобразование ключей вклада в строки превью.

## Архитектурные наблюдения

- Сценарий списка плит построен как FSM-поток aiogram: приём списка → подтверждение → возможная замена широких плит → отправка превью → переход к выбору менеджера (`bot/handlers/commercial.py:151-152`, `bot/handlers/commercial.py:355`, `bot/handlers/commercial.py:436`, `bot/handlers/commercial.py:473`).
- Для всех источников ввода используется единый парсер `set_plate_lists_from_text()`, а фото лишь готовит текст для этого парсера (`bot/handlers/commercial.py:242-243`, `bot/handlers/commercial.py:285-286`, `core/ocr_gpt.py:117-119`).
- Источник правды по итоговому заказу — словарь `PLATE_LOAD_DETAILS`, а связь "какая строка дала какую позицию" хранится в `line_contributions` и `line_plate_load_details` (`core/config_and_data.py:166-170`, `core/config_and_data.py:548-551`, `core/config_and_data.py:582-590`).
- Превью XLSX показывает не только текстовые строки, но и разложение логической строки на физические позиции по ключам вклада, поэтому одна строка заказа может занять несколько строк в Excel (`core/plates_preview_xlsx.py:223-225`, `core/plates_preview_xlsx.py:286-305`).
