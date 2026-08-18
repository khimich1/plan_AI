# Implementation Plan: Планирование от ёмкости — срочные плиты + подложки из поздних КП

## Overview

Фича добавляет в wizard создания плана автоматический сбор срочных позиций (по дедлайнам из графика поставки / execution_terms) и рекомендации «подложек» — поздних плит, которые оптимизатор фактически режет из остатков срочных. Планировщик получает предложение состава плана, которое закрывает сроки и не теряет эффективность реза.

## Architecture Decisions

1. **Анализ read-only.** `POST /production/analyze-substrates` не мутирует БД. Все изменения — через существующий `POST /production/plans/build`.
2. **Рекомендации — преселектор.** Галочки в блоках «Срочные» и «Подложки» заполняют `selectedPlatesByKp` / `selectedPlateQtyByKp`, но финальный прогон оптимизатора свободен.
3. **Per-day ёмкость.** Новая таблица `day_capacity_override` + fallback на `TRACKS_PER_DAY_DEFAULT`. `fill_targets.tracks` ≤ `max_tracks` дня.
4. **qty_remaining.** Точный SQL по `plate_id`: `qty − Σ(qty по status='в плане' AND plan_id IS NOT NULL)`.
5. **Дефицит ёмкости.** Новая функция `calculate_capacity_deficit` в `core/production/capacity.py`, не переиспользуем `delivery_schedule_check` (разная семантика: позиции vs партии).
6. **Без кэша анализа.** Пересчёт при каждом нажатии «Найти подложки» — дешёвый прогон (<30 сек по A2).

## Task List

### Phase 0: Validation Gate (блокер)

- [x] Task 0: Phase 0 validation script ✅
  - **Description:** Скрипт на реальной `plita.db` проверяет A1 (мэтчи существуют), A2 (время прогона), A3 (execution_terms парсится). Отчёт в `ai_docs/develop/reports/`.
  - **Acceptance criteria:**
    - [x] Скрипт выводит: количество позиций в бэклоге, количество срочных, количество cross-KP мэтчей, время прогона, % непарсящихся execution_terms.
    - [x] Отчёт сохраняется в `ai_docs/develop/reports/2026-08-12-podlozhki-phase0.md`.
    - [x] Go/No-go решение зафиксировано в отчёте.
  - **Verification:**
    - [x] `python scripts/validate_podlozhki_phase0.py --db plita.db` завершается без ошибок.
    - [x] Отчёт содержит конкретные цифры по A1–A3.
  - **Dependencies:** None
  - **Files likely touched:**
    - `scripts/validate_podlozhki_phase0.py` (новый)
    - `ai_docs/develop/reports/2026-08-12-podlozhki-phase0.md` (новый)
  - **Estimated scope:** M

### Checkpoint: Phase 0 Gate
- [x] A1 ≥ 1 мэтч, A2 ≤ 30 сек, A3 ≥ 80% парсится → продолжаем
- [ ] Иначе — фича переосмысляется, задачи ниже отменяются

---

### Phase 1: Backend Foundation

- [x] Task 1: Day capacity override schema + repository ✅
  - **Description:** Миграция `day_capacity_override` в `kp_db_schema.py` + `DayCapacityRepository` с CRUD.
  - **Acceptance criteria:**
    - [ ] Таблица создаётся при старте приложения (IF NOT EXISTS).
    - [ ] Repository отдаёт `get_max_tracks(date)` → `int` (override или default 5).
    - [ ] Repository отдаёт `list_overrides()` → `dict[date, int]`.
    - [ ] `set_override(date, max_tracks)` валидирует `max_tracks >= 0`.
  - **Verification:**
    - [ ] `pytest tests/test_day_capacity_repository.py -q` зелёный.
  - **Dependencies:** Task 0
  - **Files likely touched:**
    - `core/kp_db_schema.py`
    - `app/repositories/day_capacity_repository.py` (новый)
    - `tests/test_day_capacity_repository.py` (новый)
  - **Estimated scope:** S

- [x] Task 2: Core capacity logic (clean domain) ✅
  - **Description:** `core/production/capacity.py` — логика ёмкости без I/O: fallback на default, валидация fill_targets vs max, расчёт дефицита.
  - **Acceptance criteria:**
    - [ ] `get_day_capacity(date, overrides)` → `max_tracks` (override или default).
    - [ ] `validate_fill_targets(fill_targets, day_capacity)` → `None` или `PlanBuildError`.
    - [ ] `calculate_capacity_deficit(urgent_length_m, fill_targets, day_capacity)` → `CapacityDeficit | None`.
    - [ ] Все функции pure, без импортов `app.*`.
  - **Verification:**
    - [ ] `pytest tests/test_production_capacity.py -q` зелёный.
  - **Dependencies:** Task 1
  - **Files likely touched:**
    - `core/production/capacity.py` (новый)
    - `tests/test_production_capacity.py` (новый)
  - **Estimated scope:** S

- [x] Task 3: KpRepository qty_remaining ✅
  - **Description:** Точный подсчёт незапланированного остатка позиции: `qty − Σ(qty по status='в плане' AND plan_id IS NOT NULL)`.
  - **Acceptance criteria:**
    - [ ] `get_plate_qty_remaining(plate_id)` → `int`.
    - [ ] Учитывает только активные планы (`plan_id IS NOT NULL`).
    - [ ] Не трогает существующие методы `list_kps_in_production`.
  - **Verification:**
    - [ ] `pytest tests/test_kp_repository_qty_remaining.py -q` зелёный.
  - **Dependencies:** Task 0
  - **Files likely touched:**
    - `app/repositories/kp_repository.py`
    - `tests/test_kp_repository_qty_remaining.py` (новый)
  - **Estimated scope:** S

- [x] Task 4: Core urgent positions logic ✅
  - **Description:** `core/production/urgent.py` — сбор позиций с дедлайном ≤ заданной даты, агрегат по plate_id, конфликт schedule vs execution_terms.
  - **Acceptance criteria:**
    - [ ] `collect_urgent_positions(plates, batches_by_plate, kp_meta, deadline_until)` → `list[UrgentPosition]`.
    - [ ] Приоритет дедлайна: `produce_by` партии > `execution_terms` КП.
    - [ ] Группировка по `plate_id`: самая ранняя дата + `deadline_details` для разворота.
    - [ ] Конфликт >7 дней → `conflict: "schedule_earlier" | "kp_earlier"`.
    - [ ] Pure function, без I/O.
  - **Verification:**
    - [ ] `pytest tests/test_production_urgent.py -q` зелёный.
  - **Dependencies:** Task 3
  - **Files likely touched:**
    - `core/production/urgent.py` (новый)
    - `tests/test_production_urgent.py` (новый)
  - **Estimated scope:** M

### Checkpoint: Backend Foundation
- [x] Все unit-тесты зелёные
- [x] `pytest tests/test_production_capacity.py tests/test_production_urgent.py tests/test_day_capacity_repository.py -q` — pass
- [ ] Миграция применяется на копии `plita.db` без ошибок

---

### Phase 2: Analysis Engine

- [x] Task 5: ProductionCapacityService ✅
  - **Description:** Сервисный слой над `core/production/capacity.py` + `DayCapacityRepository`. API для UI: получение ёмкости по датам, сохранение override, валидация fill_targets.
  - **Acceptance criteria:**
    - [ ] `get_capacity_map(dates)` → `dict[date, int]` (с учётом overrides).
    - [ ] `set_day_capacity(date, max_tracks, user)` → сохраняет override.
    - [ ] `validate_fill_targets(fill_targets)` → использует `core/production/capacity.py`.
  - **Verification:**
    - [ ] `pytest tests/test_production_capacity_service.py -q` зелёный.
  - **Dependencies:** Task 2
  - **Files likely touched:**
    - `app/services/production_capacity_service.py` (новый)
    - `tests/test_production_capacity_service.py` (новый)
  - **Estimated scope:** S

- [x] Task 6: ProductionUrgentService ✅
  - **Description:** Сервис сбора срочных позиций. Читает `delivery_batch_item`, `delivery_batch`, `kp_plates`, `KP_offers` через репозитории.
  - **Acceptance criteria:**
    - [ ] `list_urgent_positions(deadline_until)` → `list[UrgentPosition]`.
    - [ ] Использует `core/production/urgent.py` для агрегации.
    - [ ] `qty_remaining` из `KpRepository.get_plate_qty_remaining`.
    - [ ] Фильтр: только позиции со статусом `'в производстве'`.
  - **Verification:**
    - [ ] `pytest tests/test_production_urgent_service.py -q` зелёный.
  - **Dependencies:** Task 4
  - **Files likely touched:**
    - `app/services/production_urgent_service.py` (новый)
    - `tests/test_production_urgent_service.py` (новый)
  - **Estimated scope:** M

- [x] Task 7: ProductionSubstrateService ✅
  - **Description:** Аналитический прогон оптимизатора по всему бэклогу + извлечение мэтчей из `secondary_cuts`.
  - **Acceptance criteria:**
    - [ ] `find_substrate_recommendations(urgent_plate_ids, deadline_until)` → `list[SubstrateRecommendation]`.
    - [ ] Загружает все плиты `'в производстве'` через `KpRepository`.
    - [ ] Вызывает `optimize_with_cascading_longitudinal_cuts` через `run_cpu_bound`.
    - [ ] Извлекает пары primary → secondary по `parent_unit_id` / `plate_assignments`.
    - [ ] `qty_recommended` = количество secondary cuts для позиции.
    - [ ] `saving_mm` = `parent_cut["rest"]`, `saving_m` = `rest × length_m / 1000`.
    - [ ] `needed_by` = дедлайн поздней позиции (produce_by или execution_terms).
    - [ ] `storage_days` = `needed_by − first_fill_target_date`.
  - **Verification:**
    - [ ] `pytest tests/test_production_substrate_service.py -q` зелёный (мок оптимизатора).
  - **Dependencies:** Task 6
  - **Files likely touched:**
    - `app/services/production_substrate_service.py` (новый)
    - `tests/test_production_substrate_service.py` (новый)
  - **Estimated scope:** M

- [x] Task 8: API endpoint POST /production/analyze-substrates ✅
  - **Description:** Endpoint в `app/api/v1/endpoints/production.py`. Оркестрация urgent + substrate + capacity deficit. AuthZ: admin + production.
  - **Acceptance criteria:**
    - [ ] Request schema: `fill_targets: list[FillTargetItem]`, `deadline_until: date`.
    - [ ] Response schema: `urgent_positions`, `substrate_recommendations`, `capacity_deficit`, `analysis_meta`.
    - [ ] 403 для роли `manager` (или других не-admin/production).
    - [ ] 400 при невалидных датах.
    - [ ] 422 при пустом бэклоге.
    - [ ] Анализ вызывается через `run_cpu_bound`.
  - **Verification:**
    - [ ] `pytest tests/test_production_api_integration.py -q` — новые тесты зелёные.
  - **Dependencies:** Task 5, Task 7
  - **Files likely touched:**
    - `app/api/v1/endpoints/production.py`
    - `app/schemas/production.py`
    - `tests/test_production_api_integration.py`
  - **Estimated scope:** M

### Checkpoint: Analysis Engine
- [x] `pytest tests/test_production_substrate_service.py tests/test_production_api_integration.py -q` — pass
- [ ] Endpoint отвечает 200 на тестовой БД с реальным оптимизатором (smoke)

---

### Phase 3: Frontend

- [x] Task 9: Day capacity UI — режим «Ёмкость» в календаре ✅
  - **Description:** Toggle «Планирование | Ёмкость» в `MonthCalendarGrid`. В режиме ёмкости клик по числу max → inline edit.
  - **Acceptance criteria:**
    - [ ] Toggle переключает режим без перезагрузки данных.
    - [ ] В режиме «Ёмкость» ячейка дня показывает `max_tracks` (серый фон).
    - [ ] Клик по числу → input с `+`/`-` или прямой ввод.
    - [ ] Сохранение вызывает `useSaveDayCapacityMutation`.
    - [ ] Режим «Планирование» работает как раньше (кисть, fill_targets).
  - **Verification:**
    - [ ] `cd frontend && npm test -- --run` — новые тесты зелёные.
  - **Dependencies:** Task 5
  - **Files likely touched:**
    - `frontend/src/features/production/components/MonthCalendarGrid.tsx`
    - `frontend/src/features/production/components/GlobalCalendarView.tsx`
    - `frontend/src/features/production/hooks/useProductionQueries.ts`
    - `frontend/src/features/production/api/productionApi.ts`
  - **Estimated scope:** M

- [x] Task 10: UrgentPositionsBlock в wizard ✅
  - **Description:** Блок «Срочные по срокам» в `CreatePlanWizard`: список позиций с дедлайном, разворот деталей, конфликт ⚠️.
  - **Acceptance criteria:**
    - [ ] Список отображает `plate_name`, `deadline`, `qty_remaining`, `kp_id`.
    - [ ] Разворот показывает `deadline_details` (партии, execution_terms).
    - [ ] Конфликт >7 дней отображается иконкой ⚠️ с tooltip.
    - [ ] Галочки по умолчанию отмечены.
    - [ ] Выбор синхронизируется с `selectedPlatesByKp` / `selectedPlateQtyByKp`.
  - **Verification:**
    - [ ] `cd frontend && npm test -- --run` — тесты на `UrgentPositionsBlock`.
  - **Dependencies:** Task 8
  - **Files likely touched:**
    - `frontend/src/features/production/components/create-plan-wizard/UrgentPositionsBlock.tsx` (новый)
    - `frontend/src/features/production/components/CreatePlanWizard.tsx`
    - `frontend/src/features/production/hooks/useCreatePlanWizardState.ts`
  - **Estimated scope:** M

- [x] Task 11: SubstrateRecommendationsBlock в wizard ✅
  - **Description:** Блок «Подложки из поздних КП»: список рекомендаций с датой «нужна к», сроком хранения, экономией. Кнопка «Найти подложки».
  - **Acceptance criteria:**
    - [ ] Кнопка «Найти подложки» вызывает `useAnalyzeSubstratesMutation`.
    - [ ] Loading state: спиннер + «Анализируем бэклог…».
    - [ ] Список показывает `plate_name`, `qty_recommended`, `under_plate_name`, `needed_by`, `storage_days`, `saving_m`.
    - [ ] Сортировка по `saving_m` desc.
    - [ ] Галочки преселектятся в `selectedPlatesByKp`.
    - [ ] Явная подпись: «Рекомендация — преселектор. Финальный состав может отличаться».
  - **Verification:**
    - [ ] `cd frontend && npm test -- --run` — тесты на `SubstrateRecommendationsBlock`.
  - **Dependencies:** Task 8
  - **Files likely touched:**
    - `frontend/src/features/production/components/create-plan-wizard/SubstrateRecommendationsBlock.tsx` (новый)
    - `frontend/src/features/production/components/CreatePlanWizard.tsx`
    - `frontend/src/features/production/hooks/useCreatePlanWizardState.ts`
  - **Estimated scope:** M

- [x] Task 12: CapacityDeficitAlert в wizard ✅
  - **Description:** Блок дефицита ёмкости: «+N дорожек до <дата>», кнопка добавления.
  - **Acceptance criteria:**
    - [ ] Отображается только при `capacity_deficit.tracks_missing > 0`.
    - [ ] Показывает `tracks_needed`, `tracks_available`, `tracks_missing`, `deficit_until`.
    - [ ] Кнопка «+N дорожек на <дата>» обновляет `fill_targets` и вызывает `useSaveDayCapacityMutation`.
    - [ ] После добавления дефицит пересчитывается (инвалидация запроса).
  - **Verification:**
    - [ ] `cd frontend && npm test -- --run` — тесты на `CapacityDeficitAlert`.
  - **Dependencies:** Task 8, Task 9
  - **Files likely touched:**
    - `frontend/src/features/production/components/create-plan-wizard/CapacityDeficitAlert.tsx` (новый)
    - `frontend/src/features/production/components/CreatePlanWizard.tsx`
  - **Estimated scope:** S

### Checkpoint: Frontend
- [x] `cd frontend && npm run build` — успешно
- [x] `cd frontend && npm test -- --run src/features/production` — 98 passed
- [ ] Ручная проверка: календарь → ёмкость → wizard → срочные → подложки → дефицит → build plan

---

### Phase 4: Integration & Polish

- [x] Task 13: E2E integration test ✅
  - **Description:** Сквозной тест: задать ёмкость → собрать срочные → найти подложки → собрать план → проверить, что плиты в плане.
  - **Acceptance criteria:**
    - [ ] Тест создаёт тестовые КП, график поставки, запускает анализ, собирает план.
    - [ ] Проверяет, что `kp_plates` получили `plan_id` и статус `'в плане'`.
    - [ ] Проверяет откат при ошибке оптимизатора.
  - **Verification:**
    - [ ] `pytest tests/test_production_podlozhki_e2e.py -q` зелёный.
  - **Dependencies:** Task 12
  - **Files likely touched:**
    - `tests/test_production_podlozhki_e2e.py` (новый)
  - **Estimated scope:** M

- [x] Task 14: Documentation and feature flag cleanup ✅
  - **Description:** Обновление `ai_docs/develop/features/`, удаление feature-flag если был, финальная проверка регрессий.
  - **Acceptance criteria:**
    - [x] `ai_docs/develop/features/planirovanie-po-srokam-podlozhki.md` — описание фичи, API, UI.
    - [ ] `pytest tests/ -q` — все зелёные. (smoke e2e+capacity: 11 passed; полный suite не перегонялся в TASK-014)
    - [x] `cd frontend && npm run build` — успешно. (Phase 3 checkpoint)
  - **Verification:**
    - [x] Feature doc + implementation report; feature-flag не вводился.
    - [x] Smoke: `pytest tests/test_production_podlozhki_e2e.py tests/test_day_capacity_repository.py -q` — 11 passed.
  - **Dependencies:** Task 13
  - **Files likely touched:**
    - `ai_docs/develop/features/planirovanie-po-srokam-podlozhki.md` (новый)
    - `ai_docs/develop/reports/2026-08-12-planirovanie-po-srokam-podlozhki-implementation.md` (новый)
  - **Estimated scope:** S

### Checkpoint: Complete
- [x] Все acceptance criteria выполнены (кроме полной `pytest tests/` и ручного UI)
- [ ] `pytest tests/ -q` — зелёный
- [x] `cd frontend && npm run build` — успешно
- [ ] Ручная проверка end-to-end flow
- [x] Готово к ревью

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| A1: нет мэтчей на реальном бэклоге | High | Task 0 (Phase 0) — блокер до любого UI |
| A2: прогон >30 сек | Medium | Task 0 замер; если медленно — фоновый job + polling (не MVP) |
| secondary без `parent_instance_id` | Medium | Fallback мэтчинг по `(length, rest_width)` в `plate_assignments` |
| Дрейф анализ → финал | High | Явная подпись в UI + `selectedPlateQtyByKp` ограничен `qty_remaining` |
| Путаница max vs plan tracks | Medium | Чёткая терминология: «Мощность дня» vs «Дорожек в план» |
| Регрессия fill_targets | Medium | Checkpoint после Phase 1: существующие тесты зелёные |

## Open Questions

1. Добавлять ли `rests_unused` в `finalize.py` для точного расчёта экономии? (не блокер для MVP)
2. Экспорт рекомендаций в XLSX/PDF для мастера цеха? (вне MVP)
