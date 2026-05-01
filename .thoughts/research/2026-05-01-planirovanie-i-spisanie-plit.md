---
date: 2026-05-01
topic: Планирование и списание плит
scope:
  - app
  - bot
  - frontend
  - core
  - viz_modules
---

# Исследование: Планирование и списание плит

## Резюме

Исследованы web/API, Telegram-бот, общий слой `core`, модуль визуализации `viz_modules` и frontend-фича `production`. Web-цепочка строит план через `POST /api/v1/production/plans/build`, сохраняет JSON-план через `bot.handlers.plan_manager`, а плиты фиксирует в SQLite через `core.plan_commit.commit_plan_plates` и `core.kp_db.mark_plates_as_planned`. Списание дня в web идёт через `POST /api/v1/production/days/{target_date}/complete`, где `ProductionCompletionService` читает детальный день, переносит выполненные плиты в `completed_plates`, возвращает брак в `в производстве`, сохраняет часть вторичных остатков в `plate_rests` и помечает день выполненным в JSON-плане. Telegram-ветка имеет отдельные handlers для тех же сценариев, но использует те же plan-файлы и большую часть функций `core.kp_db`.

## Подробные находки

### FastAPI production endpoints

**Расположение:** `app/api/v1/endpoints/production.py:33-264`  
**Слой:** endpoint  
**Что делает:** создаёт `APIRouter(prefix="/production", tags=["production"])`, отдаёт список планов, создаёт план, строит план по фильтрам, отдаёт календарь/занятость/кандидатов/детальный день, завершает день, отдаёт документы и рабочий календарь.  
**Входы:** Pydantic-payload: `CreatePlanRequest`, `BuildPlanRequest`, `CompleteProductionDayRequest`, `SaveWorkCalendarRequest`; path/query параметры `plan_id`, `target_date`, `exclude_plan_id`, `limit`.  
**Выходы:** dict или response models `BuildPlanResponse`, `DayOccupancyResponse`, `KpCandidatesResponse`, `DayViewDetailResponse`, `DeletePlanResponse`; документы через `FileResponse`.  
**Ключевые зависимости:** `require_roles("admin", "production")`, `ProductionService`, `ProductionPlanBuildError`, `ProductionCompletionError`, `generate_day_schema`, `generate_day_breakdown`, `generate_day_formovka`.  
**Связи:** `build_plan_from_filters()` вызывает `ProductionService().build_plan_from_filters(...)` (`app/api/v1/endpoints/production.py:51-79`); `complete_day()` вызывает `ProductionService().complete_day(...)` и передаёт `rejected_plates` из payload (`app/api/v1/endpoints/production.py:150-168`).  
**Паттерны:** сервис инстанцируется прямо в endpoint; ошибки построения и завершения дня мапятся в HTTP 422, непредвиденная ошибка построения плана мапится в HTTP 500 (`app/api/v1/endpoints/production.py:71-78`, `app/api/v1/endpoints/production.py:164-168`).

### ProductionService

**Расположение:** `app/services/production_service.py:18-169`  
**Слой:** service  
**Что делает:** фасад над планами, кандидатами КП, календарём, построением плана, детальным днём, завершением дня и рабочим календарём.  
**Входы:** параметры методов: `start_date`, `tracks_count`, `filter_method`, `selected_kp_ids`, `selected_plate_ids`, `fill_targets`, `plan_id`, `target_date`, `rejected_plates`.  
**Выходы:** словари плана/календаря/кандидатов/результата завершения; `get_plan()` может вернуть `None`.  
**Ключевые зависимости:** `KpRepository`, `PlanRepository`, `WorkCalendarRepository`, `OptimizationService`, `ProductionPlanningService`, `ProductionCompletionService`, `build_day_view_detail`, `bot.handlers.plan_manager`.  
**Связи:** построение делегируется `ProductionPlanningService.build_plan()` (`app/services/production_service.py:69-90`); завершение делегируется `ProductionCompletionService.complete_day()`, затем день помечается выполненным через `plan_repository.mark_day_completed()` (`app/services/production_service.py:101-121`).  
**Паттерны:** сервис создаёт зависимости в конструкторе, если они не переданы (`app/services/production_service.py:18-37`).

### Web-сервис построения плана

**Расположение:** `app/services/production_planning_service.py:37-322`  
**Слой:** service  
**Что делает:** загружает КП и плиты из SQLite, строит `orders_2d`, запускает оптимизацию, строит последовательность раскладки, разбивает её на дорожки, распределяет дорожки по дням, коммитит плиты в БД и сохраняет план в файлы.  
**Входы:** `start_date`, `tracks_count`, `filter_method`, выбранные КП/плиты, `active_plan_id`, `plan_name`, `fill_targets` (`app/services/production_planning_service.py:51-62`).  
**Выходы:** `{"plan": safe_plan, "stats": stats, "summary": summary}` (`app/services/production_planning_service.py:305-322`).  
**Ключевые зависимости:** `sqlite3`, `kp_db.init_schema`, `get_reinforcement`, `OptimizationService`, `core.visualization.split_sequence_into_tracks`, `viz_modules.layout_sequence.build_layout_sequence`, `core.plan_commit.commit_plan_plates`, `plan_manager.add_tracks_to_plan`, `plan_manager.save_plan`, `plan_manager.update_plan_metadata`, `plan_manager.set_active_plan`.  
**Связи:** `_load_kp_list()` выбирает КП из `KP_offers JOIN kp_meta` со статусом `в работе` (`app/services/production_planning_service.py:349-393`); `_load_plates_for_kps()` выбирает `kp_plates` со статусом `в производстве`, опционально по выбранным `id` плит (`app/services/production_planning_service.py:404-487`); `_build_orders()` переносит в `orders_2d` идентичность `kp_id`, `plate_name`, `length_dm_raw` (`app/services/production_planning_service.py:501-570`).  
**Паттерны:** перед сохранением на диск сначала вызывается `commit_plan_plates`; если сохранение файла/metadata падает, код вызывает `kp_db.return_plan_plates_to_production(plan_id, self.plita_db_path)` для отката плит из статуса `в плане` (`app/services/production_planning_service.py:268-303`).

### Режим дозаполнения дней

**Расположение:** `app/services/production_planning_service.py:152-183`, `app/services/production_planning_service.py:823-917`  
**Слой:** service  
**Что делает:** при наличии `fill_targets` валидирует свободные слоты, обрезает список дорожек до суммарной ёмкости, раскладывает дорожки строго по указанным датам и создаёт новый план.  
**Входы:** список `{date, tracks}` из `BuildPlanRequest.fill_targets` (`app/schemas/production.py:80-95`).  
**Выходы:** `precomputed_tracks_by_day`, обновлённые `start_date` и `tracks_per_day_effective`, обрезанный `optimization_result`.  
**Ключевые зависимости:** `plan_manager.get_global_day_occupancy`, `plan_manager.MAX_TRACKS_PER_DAY`, `_trim_assignments_to_tracks`.  
**Связи:** `_validate_fill_targets()` проверяет формат даты, что `tracks >= 1`, и что запрошено не больше свободных слотов (`app/services/production_planning_service.py:823-850`); `_build_tracks_by_day_from_targets()` раскладывает срезы `kept_tracks` по датам без переноса на следующий рабочий день (`app/services/production_planning_service.py:852-871`); `_trim_assignments_to_tracks()` оставляет в `plate_assignments` не больше плит, чем попало в сохранённые дорожки (`app/services/production_planning_service.py:873-917`).  
**Паттерны:** при дозаполнении `active_plan_id` принудительно сбрасывается в `None`, поэтому создаётся новый план (`app/services/production_planning_service.py:177-182`).

### PlanRepository и plan_manager

**Расположение:** `app/repositories/plan_repository.py:7-61`, `bot/handlers/plan_manager.py:29-538`, `bot/handlers/plan_manager.py:705-735`, `bot/handlers/plan_manager.py:1050-1324`  
**Слой:** repository / bot utility  
**Что делает:** `PlanRepository` является тонкой обёрткой над `bot.handlers.plan_manager`; `plan_manager` хранит планы как JSON-файлы в `bot/data/plans`, metadata в `bot/data/plans_metadata.json`, распределяет дорожки по дням и собирает глобальный календарь.  
**Входы:** `plan_id`, `new_tracks_list`, `start_date`, `tracks_per_day`, lookup-таблицы, `orders_2d`, `optimization_result`, `precomputed_tracks_by_day`.  
**Выходы:** JSON-план, metadata, календарные структуры, `tracks` на дату.  
**Ключевые зависимости:** `core.kp_db`, `core.serialization.strip_plate_audit_from_plan`, `core.work_calendar.is_working_day/load_holidays/load_extra_workdays`.  
**Связи:** `save_plan()` пишет plan JSON (`bot/handlers/plan_manager.py:163-183`); `update_plan_metadata()` пишет summary в metadata (`bot/handlers/plan_manager.py:541-580`); `set_active_plan()` пишет `active_plan_id` (`bot/handlers/plan_manager.py:244-254`).  
**Паттерны:** `add_tracks_to_plan()` создаёт новый plan dict с `id`, `name`, `created_at`, `start_date`, `tracks_count`, `days`, lookup-таблицами, `orders_2d`, `optimization_result`, `completed_days` (`bot/handlers/plan_manager.py:361-442`); день хранится как `days[date_key] = {date, day_number, tracks, saved_tracks_count, total_tracks_count, completed}` (`bot/handlers/plan_manager.py:480-489`).

### Распределение дорожек и календарь

**Расположение:** `bot/handlers/plan_manager.py:287-358`, `bot/handlers/plan_manager.py:810-874`, `bot/handlers/plan_manager.py:1050-1175`, `bot/handlers/plan_manager.py:1178-1324`  
**Слой:** bot utility / shared plan storage  
**Что делает:** распределяет дорожки по рабочим дням, считает глобальную занятость, строит глобальный календарь и собирает дорожки конкретной даты из всех планов.  
**Входы:** список дорожек, дата старта, лимит дорожек в день, глобальная занятость, `exclude_plan_id`, `date_key`.  
**Выходы:** `tracks_by_day`, `occupancy`, `days_info`, bundle дня: `tracks`, lookup-таблицы, `orders_2d`, `optimization_result`, `source_plans`.  
**Ключевые зависимости:** `MAX_TRACKS_PER_DAY = 5` (`bot/handlers/plan_manager.py:34-35`), `count_day_tracks()`, `load_plans_metadata()`, `load_plan()`.  
**Связи:** `distribute_tracks_by_days()` пропускает нерабочие дни и дни без свободных слотов (`bot/handlers/plan_manager.py:287-358`); `get_global_day_occupancy()` суммирует фактическую длину `tracks` по всем планам (`bot/handlers/plan_manager.py:810-852`); `get_tracks_for_date_from_all_plans()` добавляет каждой дорожке `source_plan_id` и `source_plan_name` и объединяет lookup/заказы (`bot/handlers/plan_manager.py:1178-1324`).  
**Паттерны:** `count_day_tracks()` считает фактический `len(day_data["tracks"])`, а не только `saved_tracks_count` (`bot/handlers/plan_manager.py:791-807`).

### Коммит плана в SQLite

**Расположение:** `core/plan_commit.py:308-733`  
**Слой:** core  
**Что делает:** считает назначенные плиты по `plate_assignments`, распределяет их по `orders_2d`, валидирует unmapped/extra, вызывает `kp_db.mark_plates_as_planned()` и записывает `kp_plate_id` обратно в items дорожек.  
**Входы:** `plan_id`, `orders_2d`, `optimization_result`, `all_tracks_list`, `db_path`, `tracks_by_day`.  
**Выходы:** `CommitResult` со счётчиками `plates_marked`, `plates_skipped`, `plates_failed`, `plates_mismatched`, `lost_plates` (`core/plan_commit.py:45-62`).  
**Ключевые зависимости:** `core.kp_db`, `core.plate_name.canonical`.  
**Связи:** `count_assigned_plates()` считает identities `(kp_id, plate_name)` по источникам `primary`, `secondary`, `rescue` из `optimization_result["plate_assignments"]` (`core/plan_commit.py:69-136`); `distribute_assigned_plates_to_orders()` распределяет счётчики по строкам `orders_2d` (`core/plan_commit.py:164-210`); `_count_track_items_by_day()` считает root items и `secondary_cuts` по `production_day` (`core/plan_commit.py:278-305`).  
**Паттерны:** при наличии `tracks_by_day` `mark_plates_as_planned()` вызывается per-day с `day_number`; возвращённые `id_qty_pairs` используются для записи `kp_plate_id` в каждый root item и `secondary_cut` (`core/plan_commit.py:436-659`).

### Схема БД плит и audit

**Расположение:** `core/kp_db.py:106-319`  
**Слой:** core / SQLite  
**Что делает:** создаёт и мигрирует таблицы КП, плит, выполненных плит, остатков и журнала статусов.  
**Входы:** путь к SQLite.  
**Выходы:** таблицы и индексы в `plita.db`.  
**Ключевые зависимости:** `sqlite3`.  
**Связи:** `kp_plates` содержит `status`, `plan_id`, `length_dm_raw`, `nomenclature_id`; миграция добавляет `status`, `plan_id`, `day_number` и индекс `(plan_id, day_number)` (`core/kp_db.py:137-160`, `core/kp_db.py:285-319`).  
**Паттерны:** `completed_plates` хранит выполненные плиты с `completed_date` и `production_day` (`core/kp_db.py:198-214`); `plate_rests` хранит остатки с `status`, `created_date`, `used_date`, `production_day` (`core/kp_db.py:216-233`); `plate_status_log` фиксирует переходы `planned`, `completed`, `rejected`, `plan_rollback` с `plan_id`, `day_number`, `actor` (`core/kp_db.py:247-267`).

### Пометка плит как "в плане"

**Расположение:** `core/kp_db.py:2369-2590`  
**Слой:** core / SQLite  
**Что делает:** переводит нужное количество строк `kp_plates` из `в производстве` в `в плане`, записывает `plan_id`, `day_number`, audit и возвращает id затронутых строк.  
**Входы:** `kp_id`, `plate_name`, `qty_to_plan`, `plan_id`, `day_number`, `actor`.  
**Выходы:** dict с `success`, `requested_qty`, `available_qty`, `processed_count`, `remaining_unplanned`, `split_count`, `updated_ids`, `id_qty_pairs`.  
**Ключевые зависимости:** `_audit_append`, SQLite.  
**Связи:** функция выбирает все строки `kp_plates` по `kp_id`, `plate_name`, `status='в производстве'`, `qty > 0` (`core/kp_db.py:2416-2423`); при полном использовании строки обновляет `status`, `plan_id`, `day_number` (`core/kp_db.py:2478-2502`); при частичном использовании уменьшает текущую строку до `qty_for_plan` и создаёт новую строку остатка со статусом `в производстве` (`core/kp_db.py:2503-2542`).  
**Паттерны:** каждая пометка пишет audit-переход `from_status='в производстве'`, `to_status='в плане'`, `reason='planned'` (`core/kp_db.py:2485-2497`, `core/kp_db.py:2523-2535`).

### Web-списание дня

**Расположение:** `app/services/production_completion_service.py:16-309`  
**Слой:** service  
**Что делает:** завершает день плана: читает план и day-view, собирает выполненные/бракованные плиты, проверяет наличие строк в БД, переносит выполненные в `completed_plates`, возвращает брак в производство, сохраняет secondary-rests и проверяет автозавершение КП.  
**Входы:** `plan_id`, `target_date`, `rejected_plates`, `actor`.  
**Выходы:** dict со счётчиками `moved_plates`, `rejected_returned`, `completed_kps`, `affected_kps`, `day_number`, `planned_qty_total`, `completed_requested_qty`, `rejected_requested_qty`, `secondary_rests`.  
**Ключевые зависимости:** `build_day_view_detail`, `kp_db.move_plates_to_completed`, `kp_db.return_plates_to_production`, `kp_db.create_plate_rest`, `kp_db.check_and_update_kp_completion`.  
**Связи:** `_collect_plates_by_kp()` группирует плиты по КП и отделяет rejected по `(track_number, plate_index)` (`app/services/production_completion_service.py:456-571`); `_to_completed_plate_payload()` формирует payload для `kp_db.move_plates_to_completed()` и передаёт `kp_plate_id`, `length_dm_raw`, `is_secondary` (`app/services/production_completion_service.py:589-613`).  
**Паттерны:** вся цепочка `move → return_rejected → check_completion` выполняется в одной SQLite-транзакции с WAL и FK (`app/services/production_completion_service.py:117-124`); при несоответствиях вызывается `conn.rollback()` и поднимается `ProductionCompletionError` (`app/services/production_completion_service.py:173-197`, `app/services/production_completion_service.py:286-300`).

### Перенос плит в completed_plates

**Расположение:** `core/kp_db.py:1312-1833`  
**Слой:** core / SQLite  
**Что делает:** списывает плиты из `kp_plates`, вставляет записи в `completed_plates`, удаляет строки с `qty <= 0` и возвращает количество списанных плит.  
**Входы:** `kp_id`, список плит, `production_day`, `plan_ids`, `actor`, опциональное внешнее соединение.  
**Выходы:** число списанных плит или `(число, unmoved_plates)` при `return_unmoved=True`.  
**Ключевые зависимости:** вложенный `find_one_row()`, `_audit_append`, SQLite.  
**Связи:** если у плиты есть `kp_plate_id`, поиск идёт по `id`, `plan_id IN (...)`, `day_number`, `status IN ('в плане', 'в производстве')` (`core/kp_db.py:1644-1677`); иначе применяется поиск по `length_dm_raw`, `plate_name`, canonical-name, эквивалентным именам, длине/ширине/нагрузке (`core/kp_db.py:1382-1606`).  
**Паттерны:** для каждой найденной строки выполняется `UPDATE kp_plates SET qty = qty - ?`, затем `INSERT INTO completed_plates (...)`, затем audit `to_status='completed'`, `reason='completed'` (`core/kp_db.py:1691-1716`).

### Возврат брака и откат плана

**Расположение:** `app/services/production_completion_service.py:341-379`, `core/kp_db.py:2593-2798`  
**Слой:** service / core  
**Что делает:** возвращает бракованные позиции или все плиты плана обратно в статус `в производстве`.  
**Входы:** для брака `{kp_id, plate_name, qty}`, `actor`, `reason`; для отката `plan_id`.  
**Выходы:** количество возвращённого брака в service, количество строк при возврате плана.  
**Ключевые зависимости:** `kp_db.return_plates_to_production`, `kp_db.return_plan_plates_to_production`.  
**Связи:** `_return_rejected()` используется web-сервисом и Telegram-ботом как единая точка возврата брака (`app/services/production_completion_service.py:349-359`); `return_plates_to_production()` выбирает строки `status='в плане'`, переводит их в `в производстве`, очищает `plan_id`, либо делает частичный split (`core/kp_db.py:2639-2727`); `return_plan_plates_to_production()` обновляет все строки по `plan_id` (`core/kp_db.py:2742-2790`).  
**Паттерны:** возврат брака пишет audit `to_status='в производстве'`, `reason='rejected'` или переданный reason (`core/kp_db.py:2670-2682`, `core/kp_db.py:2706-2718`).

### Детальный вид дня для web

**Расположение:** `app/services/day_view_service.py:340-497`  
**Слой:** service  
**Что делает:** собирает день по всем планам на дату, группирует дорожки по `plan_id`, строит `plates_info` для UI и списания.  
**Входы:** `date_key`, опциональный `db_path`.  
**Выходы:** `{"date", "plans", "plans_count", "total_tracks"}` или `None`.  
**Ключевые зависимости:** `plan_manager.get_tracks_for_date_from_all_plans`, `_load_db_rows_for_plan_day`, `_aggregate_plates_for_track_from_db`, `_aggregate_plates_for_track`, `_build_smart_lookup`.  
**Связи:** если в track items есть `kp_plate_id`, сервис читает `kp_plates` по `plan_id + day_number + status='в плане'` (`app/services/day_view_service.py:217-255`, `app/services/day_view_service.py:404-424`); для legacy-планов используется lookup путь (`app/services/day_view_service.py:426-428`).  
**Паттерны:** новый путь помечает `is_legacy=False`, legacy путь `is_legacy=True` в track-блоке (`app/services/day_view_service.py:404-439`).

### Визуализация и раскладка

**Расположение:** `viz_modules/layout_sequence.py:163-404`, `viz_modules/layout_sequence.py:1040-1453`, `core/visualization.py:56-505`, `core/visualization.py:508-1088`, `viz_modules/visualization_drawing.py:14-200`  
**Слой:** viz / core  
**Что делает:** `build_layout_sequence()` строит последовательность плит по результату оптимизации, `split_sequence_into_tracks()` разбивает последовательность на дорожки, `visualize_plan()` рисует дорожки и сохраняет PNG/PDF/XLSX/CSV, `visualization_drawing` рисует сегменты и резы.  
**Входы:** глобали `core.optimization.OPT_CASCADING_PLAN`, `OPT_CASCADING_PLAN_BY_LOAD`, `core.config_and_data.PLATE_LOAD_DETAILS`; либо готовые `existing_tracks` для `visualize_plan`.  
**Выходы:** список sequence/grouped-sequence, список дорожек, файлы `Схема_*.png`, `Схема_*.pdf`, `Ведомость_*.csv/xlsx`, `Список_плит_*.xlsx`, `Детальная_разбивка_*.xlsx`.  
**Ключевые зависимости:** `core.reinforcement_db.get_reinforcement`, `core.optimization`, `viz_modules.procurement`, `matplotlib`.  
**Связи:** web-построение вызывает `build_layout_sequence()` и `split_sequence_into_tracks()` после `OptimizationService.optimize()` (`app/services/production_planning_service.py:668-674`); генерация документов вызывает `visualize_plan(existing_tracks=...)` (`app/services/day_documents_service.py:81-88`, `app/services/day_documents_service.py:152-192`).  
**Паттерны:** sequence может быть сгруппирован по нагрузкам и содержит `load_code`, `original_loads`, `sequence`, `label` (`viz_modules/layout_sequence.py:264-404`); `split_sequence_into_tracks()` возвращает дорожки с `items`, `length`, `load_code`, `label`, `max_reinforcement` (`core/visualization.py:56-78`, `core/visualization.py:261-328`); отрисовка использует зелёные solid/primary сегменты, синие secondary, серые остатки/отходы, красные поперечные резы (`viz_modules/visualization_drawing.py:14-200`, `core/visualization.py:839-879`).

### Web-документы дня

**Расположение:** `app/services/day_documents_service.py:1-240`  
**Слой:** service  
**Что делает:** генерирует PDF-схему, XLSX-разбивку и ZIP формовки для выбранной даты.  
**Входы:** `target_date`.  
**Выходы:** `(pdf_path, cleanup_dir)`, `(xlsx_path, cleanup_dir)`, `(zip_path, cleanup_dir)`.  
**Ключевые зависимости:** `plan_manager.get_tracks_for_date_from_all_plans`, `visualize_plan`, `create_formovka_files_for_tracks`, `OptimizationContext`, `_visualize_lock`.  
**Связи:** endpoints `/documents/schema`, `/documents/breakdown`, `/documents/formovka` вызывают эти функции и добавляют cleanup через `BackgroundTasks` (`app/api/v1/endpoints/production.py:171-240`).  
**Паттерны:** из-за глобалей `core.optimization` и `core.config_and_data` вызовы `visualize_plan` сериализуются через `asyncio.Lock` (`app/services/day_documents_service.py:10-12`, `app/services/day_documents_service.py:45-47`).

### Telegram: регистрация и старт меню

**Расположение:** `bot/handlers/__init__.py:3-31`, `bot/handlers/production_planning.py:15-30`, `bot/keyboards.py:98-116`  
**Слой:** bot handler / keyboards  
**Что делает:** production handlers подключены в общий dispatcher; кнопка "Планирование производства" открывает меню производства.  
**Входы:** Telegram message `F.text == "Планирование производства"`.  
**Выходы:** сообщение с inline-клавиатурой `production_menu_kb()`.  
**Ключевые зависимости:** `production_planning`, `production_plans_list`, `production_create`, `production_calendar`, `production_execution`, `production_day_view`, `production_export`, `production_completion`, `work_calendar_manager`.  
**Связи:** `register_all_handlers()` включает роутеры production в строках `20-28` (`bot/handlers/__init__.py:19-28`).  
**Паттерны:** production menu содержит кнопки "Календарный план", "Начать планирование", "Планы", "Производственный календарь", "Назад" (`bot/keyboards.py:98-116`).

### Telegram: построение плана

**Расположение:** `bot/handlers/production_execution.py:109-1455`  
**Слой:** bot handler  
**Что делает:** загружает КП по фильтрам из FSM, читает плиты из `kp_plates`, учитывает остатки, строит `orders_2d`, запускает оптимизацию, строит sequence/tracks, добавляет rescue, сохраняет данные в FSM и показывает кнопки дней.  
**Входы:** `FSMContext` с `tracks_count`, `filter_method`, `target_date`, `kp_ids`, `customer_name`, `kp_plate_ids`, `plan_start_date`.  
**Выходы:** state с `all_tracks_list`, `orders_2d`, lookup-таблицами, `optimization_result`, `days_info`; сообщение с `calendar_days_kb`.  
**Ключевые зависимости:** `kp_db`, `PB_DB_PATH`, `PLITA_DB_PATH`, `get_reinforcement`, `OptimizationService`, `build_layout_sequence`, `split_sequence_into_tracks`, `build_rescue_tracks`, `backfill_assignment_identity`, `backfill_track_items_identity`, `calendar_days_kb`.  
**Связи:** выбор КП делается по фильтрам `date`, `kp`, `all`, `customer` (`bot/handlers/production_execution.py:128-220`); плиты читаются из `kp_plates` со статусом `в производстве` (`bot/handlers/production_execution.py:249-301`); sequence/tracks формируются через `build_layout_sequence()` и `split_sequence_into_tracks()` (`bot/handlers/production_execution.py:895-986`); кнопки дней строятся через `calendar_days_kb(...)` (`bot/handlers/production_execution.py:1434-1445`).  
**Паттерны:** Telegram-план перед сохранением живёт в FSM, а пользователь отдельно нажимает кнопку сохранения из календарной клавиатуры (`bot/handlers/production_execution.py:1338-1351`, `bot/keyboards.py:305-312`).

### Telegram: просмотр дня и документы

**Расположение:** `bot/handlers/production_day_view.py:128-560`, `bot/handlers/production_day_view.py:563-1129`, `bot/keyboards.py:937-976`  
**Слой:** bot handler / keyboards  
**Что делает:** по callback `production_day_{day}` показывает состав дня, собирает дорожки из сохранённых планов или FSM, выводит список плит по дорожкам и даёт меню генерации документов/завершения дня.  
**Входы:** callback `production_day_*`, FSM state `from_saved_plan`, `plan_start_date`, `all_tracks_list`, lookup, `orders_2d`, `optimization_result`.  
**Выходы:** сообщения с составом дорожек, state с `current_day_tracks`, `current_day_orders_2d`, `current_day_optimization_result`, `current_day_source_plans`, клавиатура `day_documents_menu_kb`.  
**Ключевые зависимости:** `plan_manager.get_tracks_for_date_from_all_plans`, `visualize_plan`, `create_formovka_files_for_tracks`, `OptimizationContext`, `PlateOrder`.  
**Связи:** для сохранённых планов дата вычисляется по `plan_start_date + day_number - 1`, затем вызывается `get_tracks_for_date_from_all_plans()` (`bot/handlers/production_day_view.py:148-214`); документы генерируются через `visualize_plan(existing_tracks=day_tracks)` (`bot/handlers/production_day_view.py:639-648`, `bot/handlers/production_day_view.py:1097-1106`).  
**Паттерны:** меню дня содержит кнопки "Схема дорожек", "Детальная разбивка", "Файлы формовки", "День выполнен", "Назад к календарю" (`bot/keyboards.py:937-976`).

### Telegram: списание и кнопки брака

**Расположение:** `bot/handlers/production_completion.py:57-1501`, `bot/keyboards.py:420-556`  
**Слой:** bot handler / keyboards  
**Что делает:** стартует завершение дня по `complete_day_*`, строит список плит по дорожкам, показывает клавиатуру выбора брака, обрабатывает `+/-/reset`, подтверждает день и вызывает функции списания/возврата/остатков.  
**Входы:** callback `complete_day_{day}`, state с `current_day_tracks` или `all_tracks_list`, lookup, `current_day_source_plans`, `rejected_quantities`.  
**Выходы:** сообщения с отчётом, обновлённая календарная клавиатура, записи в SQLite и отметка дня выполненным в plan JSON.  
**Ключевые зависимости:** `kp_db.move_plates_to_completed`, `kp_db.check_and_update_kp_completion`, `kp_db.create_plate_rest`, `ProductionCompletionService._return_rejected`, `mark_day_completed`, `load_plan`, `calendar_days_kb`, `plates_completion_kb`.  
**Связи:** `start_day_completion()` собирает `day_plates_by_track` и показывает `plates_completion_kb()` (`bot/handlers/production_completion.py:57-743`); `confirm_day_completion()` делит позиции на выполненные и бракованные, группирует по КП и вызывает `move_plates_to_completed()` (`bot/handlers/production_completion.py:887-1079`); брак возвращается через `ProductionCompletionService._return_rejected()` (`bot/handlers/production_completion.py:1252-1288`).  
**Паттерны:** `plates_completion_kb()` рисует заголовки дорожек, кнопки плит, открываемые счётчики `− / Брак X/Y / + / Сбросить`, и кнопки "Подтвердить" / "Отмена" (`bot/keyboards.py:420-556`).

### Frontend: страница production и API

**Расположение:** `frontend/src/pages/production/ProductionPage.tsx:27-115`, `frontend/src/features/production/api/productionApi.ts:16-90`, `frontend/src/features/production/hooks/useProductionQueries.ts:18-153`, `frontend/src/features/production/types/production.ts:1-168`  
**Слой:** frontend page / feature api / hooks / types  
**Что делает:** страница `/production` переключает вкладки календаря, создания плана, списка планов и рабочего календаря; API-модуль вызывает backend; React Query hooks кешируют и инвалидируют данные.  
**Входы:** query param `tab`, state `basket`, `fillRequest`; API payload `BuildPlanRequest`, `RejectedPlateItem[]`.  
**Выходы:** React UI, запросы к `/api/v1/production/*`, скачивание документов.  
**Ключевые зависимости:** `ProductionTabs`, `GlobalCalendarView`, `CreatePlanWizard`, `PlansList`, `WorkCalendarEditor`, `httpClient`, TanStack Query.  
**Связи:** `buildPlan` вызывает `POST /api/v1/production/plans/build` (`frontend/src/features/production/api/productionApi.ts:32-37`); `completeDay` вызывает `POST /api/v1/production/days/{date}/complete` с `{plan_id, rejected_plates}` (`frontend/src/features/production/api/productionApi.ts:53-62`); после completion инвалидируются production и archive query keys (`frontend/src/features/production/hooks/useProductionQueries.ts:109-125`).  
**Паттерны:** корзина дозаполнения живёт на уровне `ProductionPage`, чтобы пережить переход из календаря в мастер (`frontend/src/pages/production/ProductionPage.tsx:31-74`).

### Frontend: календарь и дозаполнение

**Расположение:** `frontend/src/features/production/components/GlobalCalendarView.tsx:24-102`, `frontend/src/features/production/components/MonthCalendarGrid.tsx:55-182`, `frontend/src/features/production/components/DayDrawer.tsx:240-270`, `frontend/src/features/production/components/FillBasket.tsx:23-58`  
**Слой:** frontend feature components  
**Что делает:** календарь показывает загрузку дней, выбранный день открывает drawer, свободные слоты можно добавить в корзину дозаполнения, корзина переводит пользователя в мастер создания плана.  
**Входы:** `days_info`, `extra_holidays`, `extra_workdays`, `basket`, selected date.  
**Выходы:** календарные кнопки дней, drawer дня, chips корзины, callback `onProceed`.  
**Ключевые зависимости:** `useGlobalCalendarQuery`, `useWorkCalendarQuery`, `DayDrawer`, `FillBasket`, `Button`.  
**Связи:** `MonthCalendarGrid` определяет state ячейки: `completed`, `full`, `partial`, `holiday`, `empty`, `outside` (`frontend/src/features/production/components/MonthCalendarGrid.tsx:71-82`); кнопка дня показывает `occupied/max`, badge выполнения и отметку корзины (`frontend/src/features/production/components/MonthCalendarGrid.tsx:133-168`); `DayDrawer` показывает секцию "Добавить в дозаполнение" только если есть свободные слоты (`frontend/src/features/production/components/DayDrawer.tsx:88-123`, `frontend/src/features/production/components/DayDrawer.tsx:240-270`).  
**Паттерны:** `FillBasket` не редактирует количество дорожек напрямую; редактирование делается через `DayDrawer` (`frontend/src/features/production/components/FillBasket.tsx:16-22`).

### Frontend: мастер создания плана

**Расположение:** `frontend/src/features/production/components/CreatePlanWizard.tsx:54-649`  
**Слой:** frontend feature component  
**Что делает:** обычный режим ведёт по шагам дата → дорожки → выбор КП; режим дозаполнения сразу открывает шаг выбора КП и отправляет `fill_targets`.  
**Входы:** `fillRequest`, выбранная дата, количество дорожек, `filterMethod`, выбранные КП/плиты, `planName`.  
**Выходы:** `BuildPlanRequest` в `useBuildPlanMutation`.  
**Ключевые зависимости:** `useBuildPlanMutation`, `useDayOccupancyQuery`, `useGlobalCalendarQuery`, `useKpCandidatesQuery`, `MonthCalendarGrid`, `Button`.  
**Связи:** при `fillRequest` компонент устанавливает `fillTargets` и `step=3` (`frontend/src/features/production/components/CreatePlanWizard.tsx:100-107`); `handleSubmit()` формирует `selected_kp_ids`, частичный `selected_plate_ids`, `start_date`, `tracks_count`, `fill_targets` (`frontend/src/features/production/components/CreatePlanWizard.tsx:175-229`).  
**Паттерны:** в fill-mode `start_date` берётся как первая дата из `fillTargets`, а `tracks_count` как максимум по `tracks`, чтобы пройти базовую валидацию payload (`frontend/src/features/production/components/CreatePlanWizard.tsx:201-218`).

### Frontend: drawer дня, списание и кнопки

**Расположение:** `frontend/src/features/production/components/DayDrawer.tsx:73-492`  
**Слой:** frontend feature component  
**Что делает:** открывает детальный день, показывает документы, планы/дорожки/плиты, позволяет задать количество брака кнопками `-`, `+`, `Сброс` и отправить completion по конкретному плану.  
**Входы:** `date`, `summary`, `onAddToFillBasket`, `alreadyInBasketTracks`, данные `useDayViewQuery(date)`.  
**Выходы:** запросы документов, `completeDay(date, planId, rejectedPlates)`, локальный state `rejectedByPlate`.  
**Ключевые зависимости:** `useDayViewQuery`, `useCompleteDayMutation`, `useDayDocumentMutation`, `Button`, `Drawer`, `Alert`.  
**Связи:** `buildRejectedPlates()` превращает локальную карту брака в массив `{track_number, plate_index, qty}` (`frontend/src/features/production/components/DayDrawer.tsx:166-182`); `handleCompleteDay()` отправляет mutation и после успеха чистит rejected по плану (`frontend/src/features/production/components/DayDrawer.tsx:184-197`).  
**Паттерны:** для каждого плана есть кнопка `Отметить выполненным`, disabled если план уже completed или mutation pending (`frontend/src/features/production/components/DayDrawer.tsx:322-354`); таблица плит показывает qty, брак и выполнено, а `+/-/Сброс` управляют rejected count (`frontend/src/features/production/components/DayDrawer.tsx:381-475`).

## Поток данных

- `React SPA /production (ProductionPage)` → `GlobalCalendarView` → `productionApi.getCalendar()` → `GET /api/v1/production/calendar` → `ProductionService.get_calendar()` → `plan_manager.get_global_calendar_info()` → `bot/data/plans/*.json` + `plans_metadata.json` → `days_info` → `MonthCalendarGrid`.
- `React SPA /production CreatePlanWizard` → `productionApi.buildPlan()` → `POST /api/v1/production/plans/build` → `ProductionService.build_plan_from_filters()` → `ProductionPlanningService.build_plan()` → `KP_offers/kp_meta/kp_plates` → `OptimizationService` → `build_layout_sequence()` → `split_sequence_into_tracks()` → `plan_manager.add_tracks_to_plan()` → `commit_plan_plates()` → `kp_db.mark_plates_as_planned()` → `kp_plates.status='в плане', plan_id, day_number` + `plate_status_log` → `plan_manager.save_plan()`.
- `React SPA DayDrawer` → `productionApi.getDayView(date)` → `GET /api/v1/production/days/{date}` → `ProductionService.get_day_view_detailed()` → `build_day_view_detail()` → `plan_manager.get_tracks_for_date_from_all_plans()` → `kp_plates(plan_id, day_number, status='в плане')` при наличии `kp_plate_id` → `DayViewResponse`.
- `React SPA DayDrawer` → `completeDay(date, planId, rejectedPlates)` → `POST /api/v1/production/days/{date}/complete` → `ProductionService.complete_day()` → `ProductionCompletionService.complete_day()` → `build_day_view_detail()` → `kp_db.move_plates_to_completed()` → `completed_plates` + уменьшение/удаление `kp_plates` + `plate_status_log` → `_return_rejected()` → `kp_plates.status='в производстве'` для брака → `kp_db.check_and_update_kp_completion()` → `plan_manager.mark_day_completed()`.
- `React SPA DayDrawer` → download buttons → `GET /documents/schema|breakdown|formovka` → `day_documents_service` → `plan_manager.get_tracks_for_date_from_all_plans()` → `visualize_plan(existing_tracks=...)` / `create_formovka_files_for_tracks()` → PDF/XLSX/ZIP.
- `Telegram` → `production_menu_kb()` → `production_execution.load_and_plan_production()` → `KP_offers/kp_meta/kp_plates` → `OptimizationService` → `build_layout_sequence()` → `split_sequence_into_tracks()` → FSM state → `calendar_days_kb()` → `production_day_view.process_day_selection()` → `day_documents_menu_kb()` → `production_completion.start_day_completion()` → `plates_completion_kb()` → `production_completion.confirm_day_completion()` → `kp_db.move_plates_to_completed()` / `_return_rejected()` / `create_plate_rest()` / `mark_day_completed()`.

## Ссылки на код

- `app/api/v1/endpoints/production.py:33` — production router с prefix `/production`.
- `app/api/v1/endpoints/production.py:51-79` — endpoint построения плана по фильтрам.
- `app/api/v1/endpoints/production.py:150-168` — endpoint завершения дня.
- `app/services/production_service.py:69-90` — делегирование построения плана в `ProductionPlanningService`.
- `app/services/production_service.py:101-121` — завершение дня и отметка completed в плане.
- `app/services/production_planning_service.py:79-101` — загрузка КП/плит и сборка `orders_2d`.
- `app/services/production_planning_service.py:140-199` — оптимизация, occupancy, fill_targets и `plan_manager.add_tracks_to_plan`.
- `app/services/production_planning_service.py:268-303` — commit в БД до сохранения файла и rollback плит при ошибке сохранения.
- `core/plan_commit.py:308-733` — коммит плана в `kp_plates`.
- `core/kp_db.py:137-160` — таблица `kp_plates`.
- `core/kp_db.py:198-214` — таблица `completed_plates`.
- `core/kp_db.py:216-233` — таблица `plate_rests`.
- `core/kp_db.py:247-267` — таблица `plate_status_log`.
- `core/kp_db.py:2369-2590` — перевод плит в статус `в плане`.
- `core/kp_db.py:1312-1833` — перенос плит в `completed_plates`.
- `core/kp_db.py:2593-2798` — возврат брака и плит плана в производство.
- `app/services/day_view_service.py:340-497` — сборка детального view дня для web.
- `app/services/day_documents_service.py:152-232` — генерация схемы, разбивки и формовки дня.
- `bot/handlers/plan_manager.py:361-538` — создание/обновление JSON-плана.
- `bot/handlers/plan_manager.py:1050-1175` — глобальный календарь всех планов.
- `bot/handlers/plan_manager.py:1178-1324` — дорожки выбранной даты из всех планов.
- `bot/handlers/production_execution.py:109-1455` — Telegram-построение плана.
- `bot/handlers/production_day_view.py:128-560` — Telegram-просмотр состава дня и меню документов.
- `bot/handlers/production_completion.py:57-1501` — Telegram-завершение дня и списание.
- `bot/keyboards.py:176-322` — календарная клавиатура дней.
- `bot/keyboards.py:420-556` — клавиатура отметки брака по плитам.
- `bot/keyboards.py:937-976` — меню документов дня.
- `viz_modules/layout_sequence.py:163-404` — sequence по группам нагрузок.
- `core/visualization.py:56-505` — разбиение sequence на дорожки.
- `core/visualization.py:508-1088` — генерация визуализации и файлов.
- `viz_modules/visualization_drawing.py:14-200` — отрисовка solid/split/transverse плит.
- `frontend/src/pages/production/ProductionPage.tsx:27-115` — страница production и state корзины дозаполнения.
- `frontend/src/features/production/api/productionApi.ts:16-90` — frontend API к `/api/v1/production`.
- `frontend/src/features/production/hooks/useProductionQueries.ts:18-153` — React Query ключи, запросы и mutations.
- `frontend/src/features/production/components/MonthCalendarGrid.tsx:71-168` — состояние и кнопки дней календаря.
- `frontend/src/features/production/components/DayDrawer.tsx:166-197` — сбор payload брака и запуск completion.
- `frontend/src/features/production/components/DayDrawer.tsx:276-353` — кнопки документов и completion.
- `frontend/src/features/production/components/DayDrawer.tsx:381-475` — кнопки `-`, `+`, `Сброс` для брака.
- `frontend/src/features/production/components/CreatePlanWizard.tsx:175-229` — payload построения плана и `fill_targets`.
- `frontend/src/features/production/components/FillBasket.tsx:23-58` — UI корзины дозаполнения.

## Архитектурные наблюдения

- Web/API production и Telegram production используют общую файловую модель планов из `bot/handlers/plan_manager.py` и общие SQLite-функции `core.kp_db`.
- JSON-план хранится в `bot/data/plans/{plan_id}.json`, metadata и active plan — в `bot/data/plans_metadata.json`.
- SQLite-учёт плит хранится в `kp_plates`: планирование переводит строки в `status='в плане'` и записывает `plan_id/day_number`, списание уменьшает `qty` и переносит выполненное количество в `completed_plates`.
- `kp_plate_id` в item/secondary_cut связывает новый plan JSON с конкретной строкой `kp_plates`; `day_view_service` использует этот путь для чтения `plates_info` напрямую из БД.
- В web завершение дня атомарно оборачивает списание, возврат брака, создание остатков и проверку completion КП в одну транзакцию.
- В Telegram завершение дня построено как интерактивный FSM-flow: просмотр дня → клавиатура плит и брака → подтверждение → вызовы `kp_db` → обновление календаря.
- Визуализация строится из результатов оптимизации через `build_layout_sequence`, затем через `split_sequence_into_tracks`; документы дня используют готовые дорожки из сохранённых планов через `existing_tracks`.
- Frontend production построен вокруг вкладок `/production?tab=...`, React Query hooks и компонентов `GlobalCalendarView`, `CreatePlanWizard`, `DayDrawer`, `PlansList`.
