# MAIN Research

## 1. Scope
- Исследованы директории: `bot/`, `core/`, `viz_modules/`, `factory_cost/`, `tests/`, корневые entrypoints.
- Проверены ожидания из правила `codebase-researcher`: в текущем checkout отсутствуют `app/` и `frontend/`.
- Подробно исследован сценарий создания КП, включая клавиатуры-кнопки, FSM, генерацию файлов и сохранение в SQLite.
- Зафиксированы фактические связи между Telegram UI, обработчиками, core-модулями и БД.

## 2. Общая карта модулей
- **Backend (app):** отдельного слоя `app/` в текущем дереве нет; HTTP-часть представлена минимальным `main.py` с одним endpoint `GET /products/{product_id}`.
- **Frontend:** React/Vite SPA (`frontend/`) в текущем checkout отсутствует; интерфейс создания КП реализован как Telegram UI (reply/inline кнопки).
- **Bot:** основной пользовательский поток живет в `bot/handlers/commercial.py`; клавиатуры и callback-кнопки централизованы в `bot/keyboards.py`; FSM-этапы в `bot/states.py`; роутеры подключаются через `bot/handlers/__init__.py`.
- **Core/viz/factory_cost:** `core/*` содержит парсинг заказа, генерацию КП (PDF/XLSX), визуализацию и работу с БД `plita.db`; `viz_modules/*` содержит раскладку/закуп/ценовые расчеты; `factory_cost/*` — импорт и выдача заводской себестоимости.
- **Tests:** в `tests/` в основном скриптовые проверки парсинга, оптимизации, визуализации и генерации КП; прямое покрытие шага кнопочного сохранения КП в БД не найдено.

## 3. Создание КП: сквозной поток
- Пользователь запускает сценарий кнопкой `📝 Создать КП` или командой `/commercial_offer`; обработчик переводит FSM в `waiting_plates_list`.
- На этапе ввода списка плит выполняется парсинг `set_plate_lists_from_text(...)`; ошибки парсинга обрабатываются через `PlateParseError`.
- Кнопка `✅ Подтвердить` в шаге `confirm_plates_list` сначала отправляет XLSX-превью сверки (`_send_plates_preview_xlsx`), затем повторное подтверждение переводит к выбору менеджера.
- После заполнения клиента/скидки/транспорта/условий формируется summary и отправляется объединенная клавиатура `save_to_db_with_files_kb(...)` (кнопки файлов + сохранение).
- Кнопки `kp_file_*` вызывают ленивую генерацию и отправку PDF/XLSX/разбивки/схемы через `generate_commercial_offer_pdf`, `generate_commercial_offer_xlsx`, `save_breakdown_to_excel`, `visualize_plan`.
- Кнопка `save_kp_to_db` переводит в `waiting_execution_terms`, парсит срок и вызывает `kp_db.save_kp_to_db(..., status='в работе')`.
- Кнопка `save_kp_to_archive` сохраняет сразу в `kp_db.save_kp_to_db(..., status='в архиве', execution_terms=None)` без отдельного шага сроков.
- В `core/kp_db.py` сохранение включает: `KP_offers` -> `kp_plates` -> `kp_files` -> `kp_meta`, с возвратом `kp_id`.

Цепочка потока:
`Telegram UI (bot/keyboards.py)` -> `handler callbacks/messages (bot/handlers/commercial.py)` -> `core parsing/generation (core/config_and_data.py, core/commercial_offer*.py, core/visualization.py)` -> `SQLite save (core/kp_db.py)` -> `ответ в чат + файлы`.

## 4. Обертка кнопок при создании КП (фокус)

### 4.1 Где находится
- Главная обертка кнопок для сценария КП находится в `bot/keyboards.py` как функции-фабрики `InlineKeyboardMarkup`/`ReplyKeyboardMarkup`.
- Ключевая объединенная обертка этапа после генерации: `save_to_db_with_files_kb(...)`.
- Вызовы этих оберток сосредоточены в `bot/handlers/commercial.py` через параметр `reply_markup=...`.

### 4.2 Какие компоненты/функции участвуют
- `main_menu_kb()` — кнопка запуска сценария `📝 Создать КП`.
- `confirm_plates_list_kb()` — кнопки `confirm_plates_list`, `replace_plates_list`, `continue_kp_plates`.
- `wide_plates_actions_kb()` — шаг обработки широких плит (`skip_wide_plates`, `cancel_process`).
- `managers_selection_kb(managers_list)` — обертка выбора менеджера с callback `select_manager_{id}`.
- `transport_choice_kb()` и `conditions_choice_kb()` — ветки транспорта и условий.
- `save_to_db_kb()` и `save_to_db_with_files_kb(...)` — сохранение в БД/архив и скачивание файлов КП.

### 4.3 Как кнопки связаны с обработчиками и API
- Связка построена через `callback_data` и aiogram-декораторы `@router.callback_query(...)`.
- `confirm_plates_list` -> `confirm_plates_list_callback(...)` -> превью и переход к следующему шагу.
- `kp_file_*` -> `callback_kp_file_download(...)` -> генерация/отправка файлов.
- `save_kp_to_db` -> `callback_save_kp_to_db(...)` -> `receive_execution_terms(...)` -> `kp_db.save_kp_to_db(...)`.
- `save_kp_to_archive` -> `callback_save_kp_to_archive(...)` -> `kp_db.save_kp_to_db(..., status='в архиве')`.
- HTTP API для этого кнопочного потока в текущем проекте не используется: интеграция идет напрямую в core/SQLite.

### 4.4 Повторное использование и зависимости
- `save_to_db_with_files_kb(...)` переиспользуется как финальная панель действий после формирования КП (документы + сохранение).
- Одна и та же callback-схема (`kp_file_*`, `save_kp_to_db`, `save_kp_to_archive`, `skip_save_kp`) используется в едином обработчике.
- Обертки кнопок зависят от доступности данных в FSM (`state`): пути файлов, `order_data`, реквизиты менеджера/клиента, скидка, условия.
- Для рендера кнопок файлов применяется параметризация через `has_pdf`, `has_xlsx`, `has_breakdown`, `has_schema`, `has_schema_breakdown`.

## 5. Функции и зависимости по слоям
- **Entry points**
  - `run_bot.py` запускает `bot.bot_main.main()` через `asyncio.run`.
  - `bot/bot_main.py` валидирует токен, инициализирует БД, создает `Bot/Dispatcher`, подключает роутеры и запускает polling.
  - `main.py` поднимает минимальный FastAPI endpoint, не связанный со сценарием КП.
- **Bot handlers**
  - `bot/handlers/commercial.py` — основной orchestration сценария КП: FSM-переходы, парсинг ввода, генерация файлов, сохранение.
  - `bot/handlers/kp.py` — отдельный сценарий `/build_plan` через `visualize_plan`, не является этапом создания КП.
  - `bot/handlers/main.py` — `/start`, `/help`, `/stats`, `/cancel`.
  - `bot/handlers/__init__.py` — централизованное подключение всех router-модулей.
- **Core: парсинг/модель заказа/сохранение**
  - `core/config_and_data.py`: `PlateOrder`, `set_plate_lists_from_text(...)`, глобальные списки и карта нагрузок.
  - `core/plates_preview_xlsx.py`: `build_plates_reconciliation_preview_xlsx(...)` для preview-сверки перед переходом к менеджеру.
  - `core/reconciliation_xlsx.py`: формирование файла сверки с колонками "как прислал / распознано / как в КП".
  - `core/kp_db.py`: схема и операции `plita.db`, включая `save_kp_to_db(...)`, менеджеров и статусы.
  - `core/db_config.py`: единые пути `PB_DB_PATH` и `PLITA_DB_PATH`.
- **Core: генерация артефактов**
  - `core/commercial_offer.py`: `generate_commercial_offer_pdf(...)`, `save_breakdown_to_excel(...)`.
  - `core/commercial_offer_xlsx.py`: `generate_commercial_offer_xlsx(...)`, `calculate_total_cost(...)`.
  - `core/visualization.py`: `visualize_plan(...)` для схем/файлов по раскладке.
- **viz_modules**
  - `viz_modules/procurement.py`: `build_procurement_items(...)`, формирование закупочных позиций из плана/нагрузок.
  - `viz_modules/layout_sequence.py`, `viz_modules/visualization_drawing.py`, `viz_modules/price_utils.py`: построение последовательности/отрисовки/цен.
- **factory_cost**
  - `factory_cost/cost_engine.py`: получение себестоимости по имени и параметрам.
  - `factory_cost/excel_reader.py`, `factory_cost/import_from_xlsx.py`, `factory_cost/db_schema.py`: импорт и схема БД себестоимости.

## 6. Тестовое покрытие
- Сценарий генерации КП напрямую затрагивает `tests/test_kp_generation.py` (проверка `generate_commercial_offer_pdf` и запись `test_kp.pdf`).
- Парсинг входного текста и структуры заказа покрывается скриптами `tests/test_parse.py`, `tests/test_exact_widths.py`, `tests/test_load_codes.py`.
- Логика группировки по нагрузкам и закупке проверяется в `tests/test_procurement_loads.py`.
- Оптимизация/визуализация проверяются в `tests/test_order.py` и `tests/test_visualization.py`.
- Прямых тестов на callback-кнопки `save_kp_to_db`/`save_kp_to_archive` и FSM-цепочку в `bot/handlers/commercial.py` в каталоге `tests/` не обнаружено.

## 7. Ссылки на код
- `.cursor/rules/codebase-researcher.mdc:13-19` — перечислены ожидаемые слои (`app/`, `frontend/`, `bot/`, `core/`, `viz_modules/`, `factory_cost/`, `tests/`).
- `.cursor/rules/codebase-researcher.mdc:47-58` — стартовые точки для `frontend` и `app` в правиле.
- `main.py:1-8` — минимальный FastAPI endpoint без КП-потока.
- `run_bot.py:8-15` — запуск aiogram entrypoint.
- `bot/bot_main.py:44-76` — запуск polling и регистрация handlers.
- `bot/handlers/__init__.py:11-31` — централизованная регистрация роутеров, включая `commercial.router`.
- `bot/states.py:5-24` — FSM состояния `KPStates` для создания/сохранения КП.
- `bot/keyboards.py:5-26` — `main_menu_kb` с кнопкой `📝 Создать КП`.
- `bot/keyboards.py:40-48` — `conditions_choice_kb`.
- `bot/keyboards.py:51-59` — `transport_choice_kb`.
- `bot/keyboards.py:62-70` — `save_to_db_kb`.
- `bot/keyboards.py:73-106` — `save_to_db_with_files_kb` (объединенная кнопочная обертка).
- `bot/keyboards.py:604-627` — `managers_selection_kb` (`select_manager_{id}`).
- `bot/keyboards.py:639-651` — `confirm_plates_list_kb`.
- `bot/keyboards.py:654-661` — `wide_plates_actions_kb`.
- `bot/handlers/main.py:16-28` — `/start` и выдача `main_menu_kb`.
- `bot/handlers/commercial.py:45-47` — импорт кнопочных оберток и состояний.
- `bot/handlers/commercial.py:61-86` — переход к выбору менеджера и выдача `managers_selection_kb`.
- `bot/handlers/commercial.py:89-118` — `_send_plates_preview_xlsx` (отправка превью-документа).
- `bot/handlers/commercial.py:167-184` — старт создания КП (`📝 Создать КП` / `/commercial_offer`).
- `bot/handlers/commercial.py:246-251` — шаг ввода плит + подключение нормализатора.
- `bot/handlers/commercial.py:481-546` — обработчик `confirm_plates_list` и двухшаговое подтверждение с preview.
- `bot/handlers/commercial.py:1456-1468` — финальный экран с `save_to_db_with_files_kb(...)`.
- `bot/handlers/commercial.py:1895-1914` — callback `kp_file_*` (отправка готового файла).
- `bot/handlers/commercial.py:1937-1964` — ленивое создание PDF/XLSX по кнопкам.
- `bot/handlers/commercial.py:1981-2010` — генерация/отправка схемы и разбивки схемы.
- `bot/handlers/commercial.py:2019-2067` — callback `save_kp_to_db`, переход в `waiting_execution_terms`.
- `bot/handlers/commercial.py:2070-2128` — парсинг сроков выполнения.
- `bot/handlers/commercial.py:2182-2197` — вызов `kp_db.save_kp_to_db(..., status='в работе')`.
- `bot/handlers/commercial.py:2255-2266` — callback `save_kp_to_archive` (описание сценария).
- `bot/handlers/commercial.py:2331-2344` — вызов `kp_db.save_kp_to_db(..., status='в архиве')`.
- `bot/handlers/kp.py:24-32` — `/build_plan` и вызов `visualize_plan` (отдельный поток).
- `core/db_config.py:9-12` — пути к `pb.db` и `plita.db`.
- `core/config_and_data.py:208-244` — dataclass `PlateOrder` и сериализация.
- `core/config_and_data.py:453-455` — получение текущего заказа `get_current_plate_order()`.
- `core/config_and_data.py:554-573` — `set_plate_lists_from_text(...)` и контракт ошибок.
- `core/plates_preview_xlsx.py:262-283` — сборка превью XLSX и вызов парсера.
- `core/reconciliation_xlsx.py:168-178` — `build_reconciliation_xlsx(...)` для сверки.
- `core/commercial_offer.py:274-285` — сигнатура генерации PDF.
- `core/commercial_offer_xlsx.py:154-166` — сигнатура генерации XLSX.
- `core/commercial_offer_xlsx.py:113-118` — `calculate_total_cost(...)`.
- `core/kp_db.py:66-97` — создание `KP_offers`.
- `core/kp_db.py:532-546` — сигнатура `save_kp_to_db(...)`.
- `core/kp_db.py:580-594` — нормализация `order_data` и расчет итогов.
- `core/kp_db.py:614-624` — вставка записи в `KP_offers`.
- `core/kp_db.py:647-659` — вставка строк `kp_plates`.
- `core/kp_db.py:666-675` — вставка `kp_files`.
- `core/kp_db.py:678-684` — запись статуса в `kp_meta` и `commit`.
- `core/visualization.py:508-516` — `visualize_plan(...)`.
- `viz_modules/procurement.py:277-285` — `build_procurement_items(...)`.
- `factory_cost/cost_engine.py:20-33` — `get_cost_by_plate_name(...)`.
- `factory_cost/cost_engine.py:89-110` — `get_cost_by_params(...)`.
- `tests/test_kp_generation.py:13-41` — проверка генерации PDF КП.
- `tests/test_parse.py:22-41` — запуск `set_plate_lists_from_text(...)`.
- `tests/test_procurement_loads.py:15-18` — проверка `build_procurement_items/build_price_rows`.
- `tests/test_exact_widths.py:18-36` — тест точных ширин через парсер.
- `tests/test_load_codes.py:20-38` — тест нагрузок 8п/10п/12,5п.
- `tests/test_order.py:23-45` — тест оптимизации `optimize_with_cascading_longitudinal_cuts`.
- `tests/test_visualization.py:23-24` — импорт оптимизации и модуля visual sequence.
