---
date: 2026-05-03
topic: Оптимизатор и layout_sequence (потеря плит)
scope:
  - app
  - bot
  - core
  - viz_modules
  - tests
---

# Исследование: Оптимизатор и layout_sequence (потеря плит)

## Резюме
Исследован фактический путь данных от `orders_2d` в оптимизаторе до последовательности и дорожек: `core/optimization.py` → `viz_modules/layout_sequence.py` → `core/visualization.py` → вызовы из web/bot сервисов. В оптимизаторе покрытие спроса считается по `demand_2d` и проверяется через `verify_coverage`, затем дополняется post-correction и force-add перед финальной атрибуцией (`core/optimization.py:1655-1677`, `core/optimization.py:1539-1622`). В `layout_sequence` secondary-сегменты привязываются к primary сначала по `parent_instance_id`, затем по геометрическому ключу `(length, rest_mm)` (`viz_modules/layout_sequence.py:1431-1447`). После построения sequence целостность перехода в tracks дополнительно валидируется `validate_track_integrity(..., strict=True)` в web и bot сценариях (`app/services/production_planning_service.py:624-631`, `bot/handlers/production_execution.py:986-993`, `core/visualization.py:592-632`).

## Подробные находки

**Расположение:** `app/services/optimization_service.py:15-44`, `app/services/optimization_service.py:46-61`  
**Слой:** service  
**Что делает:** запускает оптимизацию через `optimize_with_cascading_longitudinal_cuts(orders_2d=...)`, формирует `OptimizationContext` и временно прокидывает результат в legacy-глобали (`OPT_CASCADING_PLAN`, `OPT_CASCADING_PLAN_BY_LOAD`, `LOAD_TO_REINFORCEMENT_MAP`) через context manager.  
**Входы:** `PlateOrder`, опционально `orders_2d`.  
**Выходы:** `OptimizationContext` с `optimization_result`, `plan_by_load`, `load_to_reinforcement_map`.  
**Ключевые зависимости:** `core.optimization`, `app.domain.models.optimization_context`.  
**Связи:** вызывается из `ProductionPlanningService._run_optimization_and_split` (`app/services/production_planning_service.py:594-599`, `app/services/production_planning_service.py:621-622`).  
**Паттерны:** основной результат хранится в контексте, а для старого pipeline временно дублируется в глобали.

**Расположение:** `core/optimization.py:558-1898`  
**Слой:** core  
**Что делает:** `_optimize_2d_with_lengths` строит ILP assignment-модель по ключам `(length, width, load_code)`, извлекает `primary_cuts`/`secondary_cuts`, рассчитывает покрытие, выполняет post-correction и формирует `plate_assignments`.  
**Входы:** `orders_2d`, `plate_width`, `min_useful_width`, `OptimizationConfig`.  
**Выходы:** `result` с `primary_cuts`, `secondary_cuts`, `plate_assignments`, `orders_requested`, `rests_created`, `rests_used`, `_coverage_summary`, `_plate_audit`.  
**Ключевые зависимости:** `pulp`, `core.config_and_data.canonical_plate_key`, `core.price_db.get_price`, `core.plate_audit.PlateAudit`.  
**Связи:** публично вызывается через `optimize_with_cascading_longitudinal_cuts` (`core/optimization.py:2238-2294`), далее читается `layout_sequence` и сервисами планирования.  
**Паттерны:** модель использует equality-покрытие спроса `sum(z_prim+z_sec+unmet)==qty` (`core/optimization.py:1089-1099`) и отдельный `unmet`-safety net.

**Расположение:** `core/optimization.py:603-675`, `core/optimization.py:454-513`, `core/optimization.py:1713-1833`  
**Слой:** core  
**Что делает:** строит proportional slot-ledger (`slot_lists`, `slot_cursors`) по спросу и затем расходует его при присвоении identity как для primary, так и для secondary.  
**Входы:** `orders_2d`, `demand_2d`, `assignment_key`/`target_order_key`.  
**Выходы:** `kp_id/customer/kp_date/plate_name/load_code/identity_match_type` внутри `primary_cuts`, `secondary_cuts`, `plate_assignments`; при исчерпании слотов ставит `slot_exhausted`/`secondary_unmapped`.  
**Ключевые зависимости:** `_build_proportional_slot_lists`, `_next_slot_info`.  
**Связи:** результаты затем backfill-ятся на сервисном уровне (`app/services/production_planning_service.py:600-607`).  
**Паттерны:** identity назначается постфактум поверх уже рассчитанных физических резов.

**Расположение:** `core/optimization.py:1525-1622`, `core/optimization.py:1655-1677`, `core/optimization.py:84-135`  
**Слой:** core  
**Что делает:** после solver-вывода пересчитывает покрытие по ключам и добирает недостающие позиции через post-correction и no-sources force-add; отдельно формирует summary `verify_coverage`.  
**Входы:** `demand_2d`, `result["primary_cuts"]`, `result["secondary_cuts"]`, `no_sources_keys`.  
**Выходы:** дополненные `primary_cuts`, обновлённый `total_plates`, `_coverage_summary`.  
**Ключевые зависимости:** `_norm_key`, `verify_coverage`.  
**Связи:** итоговые cuts идут в `build_layout_sequence`; `verify_coverage` используется и в baseline-тестах (`tests/test_optimization_baseline.py:27-31`, `tests/test_optimization_baseline.py:149-166`).  
**Паттерны:** покрытие проверяется после пост-коррекции, не только по solver-raw output.

**Расположение:** `viz_modules/layout_sequence.py:175-423`, `viz_modules/layout_sequence.py:1064-1550`  
**Слой:** viz  
**Что делает:** `build_layout_sequence()` берёт план из `OPT_CASCADING_PLAN_BY_LOAD`/`OPT_CASCADING_PLAN`, строит sequence либо через grouped path (`all_sequences`), либо через `_build_sequence_from_plan`. В `_build_sequence_from_plan` secondary связывается сначала по `parent_instance_id`, потом fallback по `_secondary_geom_cut_key(length, rest_mm)`.  
**Входы:** глобали оптимизатора, `cfg.PLATE_LOAD_DETAILS`, `plan["primary_cuts"]`, `plan["secondary_cuts"]`.  
**Выходы:** sequence в формате list root-items либо grouped list (`[{load_code, sequence, label}]`).  
**Ключевые зависимости:** `_get_reinforcement_from_map`, `_choose_best_separator`, `_split_group_into_subgroups`, `_secondary_geom_cut_key`.  
**Связи:** вызывается из web/bot planning и visualizer (`app/services/production_planning_service.py:622`, `bot/handlers/production_execution.py:907`, `core/visualization.py:746`).  
**Паттерны:** для каждого root-item сохраняется `layout_uid`/`unit_id` (`viz_modules/layout_sequence.py:102-111`, `viz_modules/layout_sequence.py:1538`), secondary привязка учитывает both parent-id и legacy-геометрию (`viz_modules/layout_sequence.py:1431-1447`).

**Расположение:** `core/visualization.py:137-160`, `core/visualization.py:178-413`, `core/visualization.py:568-632`  
**Слой:** core  
**Что делает:** `split_sequence_into_tracks` делит sequence на дорожки по ограничениям длины и правилу старта с целой плиты, затем сравнивает вход/выход и запускает `validate_track_integrity`.  
**Входы:** `sequence`, `max_track_length`, `min_track_length`, `strict_layout_integrity`.  
**Выходы:** `tracks` с `items`, `length`, `load_code`, `label`, `max_reinforcement`.  
**Ключевые зависимости:** `_iter_sequence_items`, `_ensure_layout_uid`, `validate_track_integrity`.  
**Связи:** strict-режим включается в web/bot build-пайплайне (`app/services/production_planning_service.py:624-627`, `bot/handlers/production_execution.py:987-990`).  
**Паттерны:** после split всегда считается `input_count` vs `output_count` и integrity-report (`core/visualization.py:568-579`, `core/visualization.py:592-627`).

**Расположение:** `app/services/production_planning_service.py:572-792`, `app/services/production_planning_service.py:794-866`, `app/services/production_planning_service.py:965-983`  
**Слой:** service  
**Что делает:** оркестрирует этапы optimize → build_layout_sequence → split_sequence_into_tracks, затем дополняет результат rescue-треками и fallback-треками по gap между `plate_assignments` и `tracks`, плюс backfill identity для track-items.  
**Входы:** `orders_2d`.  
**Выходы:** `all_tracks_list`, `optimization_result` (включая rescue assignments).  
**Ключевые зависимости:** `OptimizationService`, `build_rescue_tracks`, `backfill_track_items_identity`.  
**Связи:** используется в сценарии построения production plan; в случае integrity-ошибки бросает `ProductionPlanBuildError` (`app/services/production_planning_service.py:628-631`).  
**Паттерны:** источник правды для rescue — `plate_assignments`, а не tracks (`app/services/production_planning_service.py:725-737`).

**Расположение:** `bot/handlers/production_execution.py:854-861`, `bot/handlers/production_execution.py:898-900`, `bot/handlers/production_execution.py:907-993`  
**Слой:** bot handler  
**Что делает:** после оптимизации записывает результат в глобали оптимизатора, строит sequence и tracks в strict integrity режиме; при ошибке integrity останавливает сценарий и отправляет сообщение пользователю.  
**Входы:** `optimization_result`, `orders_2d`.  
**Выходы:** `seq`, `all_tracks_list` (или ранний return при integrity error).  
**Ключевые зависимости:** `build_layout_sequence`, `split_sequence_into_tracks`.  
**Связи:** зеркалит web-пайплайн по шагам layout/split.  
**Паттерны:** grouped-представление задаётся как `OPT_CASCADING_PLAN_BY_LOAD = {'all': optimization_result}`.

**Расположение:** `tests/test_optimization_baseline.py:149-166`, `tests/test_optimization_baseline.py:198-222`, `tests/test_layout_identity_integrity.py:25-134`, `tests/test_layout_sequence_unbound.py:50-68`  
**Слой:** test  
**Что делает:** фиксирует инварианты покрытия спроса, отсутствие потерь на конкурирующих длинах, корректность secondary-привязки по parent-id/геометрии и наличие `length`/`plate_uid` в sequence для 2D-плана.  
**Входы:** synthetic `orders_2d` и synthetic plan fixtures.  
**Выходы:** assertions по `verify_coverage`, `secondary_cuts` mapping, `validate_track_integrity`.  
**Ключевые зависимости:** `optimize_with_cascading_longitudinal_cuts`, `_build_sequence_from_plan`, `build_layout_sequence`.  
**Связи:** покрывают обе стороны цепочки — optimizer и layout mapping.  
**Паттерны:** регрессионные кейсы для нестабильных ключей описаны как отдельные параметризованные сценарии.

## Поток данных
- `Telegram (bot/handlers/production_execution.py)` → `OptimizationService/optimization_result` → запись в `OPT_CASCADING_PLAN*` (`bot/handlers/production_execution.py:854-861`) → `build_layout_sequence()` (`bot/handlers/production_execution.py:907`) → `split_sequence_into_tracks(..., strict_layout_integrity=True)` (`bot/handlers/production_execution.py:987-990`) → tracks или остановка при `LayoutIntegrityError` (`bot/handlers/production_execution.py:991-996`).
- `Web service (app/services/production_planning_service.py)` → `OptimizationService.optimize()` (`app/services/production_planning_service.py:594-599`) → `legacy_runtime` + `build_layout_sequence()` (`app/services/production_planning_service.py:621-623`) → `split_sequence_into_tracks(..., strict_layout_integrity=True)` (`app/services/production_planning_service.py:624-627`) → `build_rescue_tracks` по `plate_assignments` (`app/services/production_planning_service.py:733-750`) → `backfill_track_items_identity` (`app/services/production_planning_service.py:762-765`) → fallback треки при gap (`app/services/production_planning_service.py:774-787`).
- `Optimizer core` → `demand_2d`/assignment model (`core/optimization.py:603-609`, `core/optimization.py:1089-1099`) → извлечение primary/secondary (`core/optimization.py:1381-1489`) → `post_correction` и `force_added_no_sources` (`core/optimization.py:1539-1619`) → `verify_coverage` (`core/optimization.py:1655-1677`) → итоговый `plate_assignments` (`core/optimization.py:1716-1825`).
- `Layout mapper` (`viz_modules/layout_sequence.py`) → build по grouped plan (`viz_modules/layout_sequence.py:276-340`) или single plan (`viz_modules/layout_sequence.py:420-423`) → secondary attach by parent/geometry (`viz_modules/layout_sequence.py:1431-1447`) → sequence items с `layout_uid` (`viz_modules/layout_sequence.py:1538`).

## Ссылки на код
- `app/services/optimization_service.py:15-44` — запуск 2D-оптимизации и сбор `OptimizationContext`.
- `app/services/optimization_service.py:46-61` — временный перенос контекста в глобали `OPT_CASCADING_PLAN*`.
- `core/optimization.py:603-637` — построение `demand_2d`, `order_info_list`, `slot_lists`.
- `core/optimization.py:991-1134` — assignment-модель `z_prim/z_sec`, equality-покрытие и soft solid-priority.
- `core/optimization.py:1381-1412` — извлечение primary напрямую из `z_prim[(opt, dk)]`.
- `core/optimization.py:1525-1622` — post-correction + force-add для ключей без источников.
- `core/optimization.py:1655-1677` — финальная проверка покрытия `verify_coverage`.
- `core/optimization.py:1716-1735` — атрибуция primary через slot ledger.
- `core/optimization.py:1783-1811` — атрибуция secondary через slot ledger.
- `viz_modules/layout_sequence.py:175-257` — построение карты армирования и дополнение из плана.
- `viz_modules/layout_sequence.py:276-340` — grouped path (`OPT_CASCADING_PLAN_BY_LOAD`) и вызов `_build_sequence_from_plan`.
- `viz_modules/layout_sequence.py:420-423` — single-plan path с ранним `return sequence`.
- `viz_modules/layout_sequence.py:1064-1131` — подготовка `_build_sequence_from_plan`, режим 2D/legacy.
- `viz_modules/layout_sequence.py:1124-1186` — сбор secondary-variants и индекс `secondary_cuts_by_parent`.
- `viz_modules/layout_sequence.py:1431-1447` — выбор secondary-варианта: parent-id → fallback geometry.
- `viz_modules/layout_sequence.py:1538-1549` — `_ensure_sequence_layout_uid` и отчёт unmatched secondary.
- `core/visualization.py:137-160` — контракт `split_sequence_into_tracks`.
- `core/visualization.py:178-413` — split grouped sequence по дорожкам.
- `core/visualization.py:568-573` — счётчик `input_count` vs `output_count`.
- `core/visualization.py:592-632` — `validate_track_integrity` и строгий режим.
- `app/services/production_planning_service.py:618-631` — strict split и исключение `ProductionPlanBuildError`.
- `app/services/production_planning_service.py:731-750` — rescue на основе `plate_assignments`.
- `app/services/production_planning_service.py:762-771` — backfill identity для root/secondary track-items.
- `app/services/production_planning_service.py:795-866` — fallback tracks при разрыве assignments↔tracks.
- `bot/handlers/production_execution.py:854-861` — публикация результата в `OPT_CASCADING_PLAN*`.
- `bot/handlers/production_execution.py:907-920` — сравнение `sequence_total` vs `plate_assignments_count`.
- `bot/handlers/production_execution.py:987-993` — strict integrity split в bot-пайплайне.
- `tests/test_optimization_baseline.py:149-166` — инвариант полного покрытия спроса.
- `tests/test_optimization_baseline.py:198-222` — регрессия по конкурирующим длинам без потерь.
- `tests/test_layout_identity_integrity.py:25-134` — привязка secondary к parent/geometric fallback.
- `tests/test_layout_sequence_unbound.py:50-68` — build-layout regression в 2D режиме.

## Архитектурные наблюдения
- Оптимизатор формирует физические резы и identity-атрибуцию в одном модуле (`core/optimization.py`), включая post-correction и coverage-report перед возвратом результата.
- `layout_sequence` строит визуальную последовательность на глобалях оптимизатора (`OPT_CASCADING_PLAN*`) и поддерживает grouped/single режимы через один API `build_layout_sequence`.
- Проверка целостности `sequence → tracks` централизована в `core/visualization.py` и может работать в strict-режиме; web и bot используют strict.
- В production-planning после split дополнительно включены слои rescue/backfill/fallback для выравнивания `orders_2d`, `plate_assignments` и `tracks`.
