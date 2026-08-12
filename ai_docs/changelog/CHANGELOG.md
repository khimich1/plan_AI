# Changelog

Все значимые изменения описываются здесь. Формат основан на [Keep a Changelog](https://keepachangelog.com/ru/).

## [Unreleased]

### Added

- **КП — несколько наименований (append loop):** один `kp_id` из нескольких заходов (в т.ч. повтор типа); sticky клиент/скидка; skip client со 2-го цикла; undo last batch / delete by `line_id`; логистика только по весу ПБ; unified PDF/XLSX при multi; append к сохранённому КП из архива (статус «в работе»); multi badges типов; production для `mixed` с плитами. [Report](../develop/reports/2026-08-12-kp-multi-nomenclature-append-implementation.md)
- **Порядок армирования в раскладке:** параметр `layout_reinforcement_order` (asc/desc) управляет алгоритмом выбора целых плит при построении раскладки плана. Режим `desc` (сильные первыми) автоматически применяет match-greedy стратегию. Добавлен в Settings, API endpoint, ProductionPlanningService, LayoutRuntimeSnapshot, helpers (reinforcement_order_key, should_pick_solid_greedy). Frontend: выпадающий список в CreatePlanWizard. Полностью обратно совместимо (default `asc`). [Report](../reports/2026-06-02-layout-reinf-order-implementation.md)
- Поле `write_off_completed` в схеме `DayPlateInfo` и в агрегате дня (`day_view_service`): отражает строки из снимка после списания дня; фронт показывает бейдж **`(ГОТОВО)`** и блокирует редактирование брака/выполнения для таких позиций.
- Тесты: `tests/test_day_view_service.py` (агрегация и Pydantic); расширение `tests/test_production_completion_service.py` (`test_day_view_write_off_completed_false_before_complete_true_after_snapshot`); Vitest `frontend/src/features/commercial-offer/store/wizardDraftStore.test.tsx` (слияние шага при `hydrate-draft`).
- Красная кнопка «Создать новое КП» на шаге результата (`CalculationResultStep`, вариант `danger`).
- **Производственная сметка:** функция `_find_price_for_plate_production_fallback` для XLSX fallback с округлением вверх (ceil дм), выравнивая со сторной БД. Тесты в `tests/test_procurement_production_fallback.py`.

### Fixed

- Визард КП: после гидрации черновика локальный шаг больше не «откатывается» назад относительно серверного — берётся максимум по `WIZARD_STEP_ORDER` (`mergeWizardStepWithServer` в `wizardDraftStore`). На шаге менеджера кнопка «Далее» безопасно вызывает `onNext`, если он асинхронный (`Promise.resolve`).

### Changed

- Стили дня производства: классы строки `.day-plates-table__row--written-off`, бейдж `.day-plate-badge--done` (`frontend/src/index.css`).
- Бейдж готовности в ДейДроуэр теперь показывает текст **(ГОТОВО)** вместо "Списано".
