# Changelog

Все значимые изменения описываются здесь. Формат основан на [Keep a Changelog](https://keepachangelog.com/ru/).

## [Unreleased]

### Added

- **Ёмкость завода + жёсткий гейт:** мини-календарь загрузки и «нужно / свободно / Δ» в left drawer по кнопке «Ёмкость» («В производство» и график поставок); в модалке при red — короткий hint; симуляция с завтра; red блокирует сохранение (UI + backend). [Gate](../specs/zavod-emkost-vizual-gate.md) · [Left drawer UX](../specs/zavod-emkost-left-drawer.md) · [Idea](../ideas/zavod-emkost-left-drawer.md)
- **КП — несколько наименований (append loop):** один `kp_id` из нескольких заходов (в т.ч. повтор типа); sticky клиент/скидка; skip client со 2-го цикла; undo last batch / delete by `line_id`; логистика только по весу ПБ; unified PDF/XLSX при multi; append к сохранённому КП из архива (статус «в работе»); multi badges типов; production для `mixed` с плитами. [Report](../develop/reports/2026-08-12-kp-multi-nomenclature-append-implementation.md)
- **ГСМ: кнопка «Экспорт zip»** на вкладке Период — скачивание бланков без API/curl. Жёсткий стоп на `manual_intervention`/`unsolvable`; confirm на жёлтые warnings и повторный экспорт; ошибка LibreOffice с подсказкой админу. [Idea](../ideas/gsm-waybill-export-buttons.md) · [Report](../develop/reports/2026-08-17-gsm-waybill-export-buttons.md)
- **ГСМ солвер A+B (бак важнее группы АЗС + короткий дожиг):** дожиг из всех маршрутов в `max_daily_km` (не сетка 150–250 км); в headroom — min sufficient km; lookahead при короткой группе АЗС берёт маршрут из всей библиотеки + `balance_route`. Palisade май 2026: `problematic_days == []`. [Spec](../specs/gsm-solver-tank-first-short-burn.md) · [Report](../develop/reports/2026-08-17-gsm-solver-tank-first-short-burn.md)
- **ГСМ lookahead-генератор с географией:** round-trip (2 плеча), lookahead на плотных заправках, мягкая сортировка по направлению к следующей АЗС, частичная генерация вместо 422 (`problematic_days`, warning `manual_intervention` / `balance_route`). Настройка `max_daily_km` (дефолт 700). Скрипты `geocode_gsm_stations.py`, `link_route_stations.py`. Ночёвки — фаза 2. [Spec](../specs/gsm-geo-lookahead-generator.md) · [Acceptance](../develop/reports/2026-08-15-gsm-geo-lookahead-acceptance.md)
- **Модуль «ГСМ: путевые листы» (роль accountant):** импорт транзакций .xls, справочники, солвер якорей/дожигания, UI периода с правкой дня, экспорт zip бланков ОКУД 0345001 через LibreOffice. API `/api/v1/gsm/*` под `REQUIRE_ACCOUNTING`. Phase 0 blank PASS. Приёмка ≤20% ручных дней — pending. [Feature](../develop/features/gsm-module-putevye-listy.md) · [Report](../develop/reports/2026-08-14-gsm-module-implementation.md)
- **Карта маршрутов ГСМ:** скрипт `scripts/build_gsm_routes_map.py` собирает `ГСМ/карта_маршрутов.html` (Leaflet) из `пул_поездок.xlsx` — линии A→B по дорогам (OSRM), цвет/фильтр по машине, поиск адреса → топ-3 ближайших трека; кэш `ГСМ/geo_cache/` (~177/244 адресов, 434 features). Зависимости: `xlrd`, `requests`. Тесты: 35. [Report](../develop/reports/2026-08-11-gsm-routes-map.md)
- **График поставки (неделя 3 / MVP):** документ XLSX/PDF (`GET .../document`), frontend-редактор партий со светофором, импорт шаблона и скачивание из drawer архива КП. Автотесты: backend 95, frontend 239. Статус: READY_FOR_HUMAN_QA. [Report](../develop/reports/delivery-schedule-week3.md)
- **График поставки (неделя 2):** API GET/PUT `/commercial/archive/{kp_id}/delivery-schedule` с живым светофором; XLSX `/template` и `/import` (черновик без сохранения). [Report](../develop/reports/delivery-schedule-week2.md)
- **Порядок армирования в раскладке:** параметр `layout_reinforcement_order` (asc/desc) управляет алгоритмом выбора целых плит при построении раскладки плана. Режим `desc` (сильные первыми) автоматически применяет match-greedy стратегию. Добавлен в Settings, API endpoint, ProductionPlanningService, LayoutRuntimeSnapshot, helpers (reinforcement_order_key, should_pick_solid_greedy). Frontend: выпадающий список в CreatePlanWizard. Полностью обратно совместимо (default `asc`). [Report](../reports/2026-06-02-layout-reinf-order-implementation.md)
- Поле `write_off_completed` в схеме `DayPlateInfo` и в агрегате дня (`day_view_service`): отражает строки из снимка после списания дня; фронт показывает бейдж **`(ГОТОВО)`** и блокирует редактирование брака/выполнения для таких позиций.
- Тесты: `tests/test_day_view_service.py` (агрегация и Pydantic); расширение `tests/test_production_completion_service.py` (`test_day_view_write_off_completed_false_before_complete_true_after_snapshot`); Vitest `frontend/src/features/commercial-offer/store/wizardDraftStore.test.tsx` (слияние шага при `hydrate-draft`).
- Красная кнопка «Создать новое КП» на шаге результата (`CalculationResultStep`, вариант `danger`).
- **Производственная сметка:** функция `_find_price_for_plate_production_fallback` для XLSX fallback с округлением вверх (ceil дм), выравнивая со сторной БД. Тесты в `tests/test_procurement_production_fallback.py`.

### Fixed

- Визард КП: после гидрации черновика локальный шаг больше не «откатывается» назад относительно серверного — берётся максимум по `WIZARD_STEP_ORDER` (`mergeWizardStepWithServer` в `wizardDraftStore`). На шаге менеджера кнопка «Далее» безопасно вызывает `onNext`, если он асинхронный (`Promise.resolve`).

### Changed

- **График поставки — остаток в шаблоне:** «Скачать шаблон» из открытого редактора собирает XLSX на клиенте (exceljs) из текущего черновика: сверху «Уже в поставках», ниже «Остаток». `GET /template` — тот же макет от сохранённого графика. Парсер пропускает заголовки полос. [Spec](../specs/delivery-schedule-remainder-template.md) · [Plan](../develop/plans/2026-08-22-delivery-schedule-remainder-template.md)
- **График поставки:** «Скачать шаблон» отдаёт XLSX с марками и количествами из выбранного КП (партии и даты пустые). Незаполненные строки при импорте пропускаются, заполненные партии попадают в веб-таблицу для проверки.

- Стили дня производства: классы строки `.day-plates-table__row--written-off`, бейдж `.day-plate-badge--done` (`frontend/src/index.css`).
- Бейдж готовности в ДейДроуэр теперь показывает текст **(ГОТОВО)** вместо "Списано".
