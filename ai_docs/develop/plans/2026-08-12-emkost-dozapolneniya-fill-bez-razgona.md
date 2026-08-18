# Implementation Plan: Ёмкость дозаполнения — hard cap 5, без тихого разгона

## Overview

Чинит рассинхрон analyze ↔ build и неверную семантику кнопки дефицита: на заводе **максимум 5 дорожек/день**. Кнопка дефицита не поднимает override; предлагает bump fill или другой день; календарь и build считают `free` одинаково; `plans/build` отдаёт конкретный текст ошибки.

Спека: [`ai_docs/specs/emkost-dozapolneniya-fill-bez-razgona.md`](../../specs/emkost-dozapolneniya-fill-bez-razgona.md).

## Architecture Decisions

1. **Hard cap = 5, floor = 0.** `PUT /day-capacity` с `max_tracks > 5` → 400; `0…5` ок; при чтении >5 → clamp 5. `today` = Europe/Moscow. Options: `add_tracks = free`, max 10, порядок A→B→C.
2. **Две проверки сходятся:** `requested ≤ free`, `free = day_max − occupied`, `day_max ≤ 5` — в analyze deficit, persist/build, calendar `days_info.max`.
3. **Deficit suggestion в core** с callback/predicate для nearest day (work calendar + occupancy в service).
4. **`capacity_deficit.options[]`** (список): A выбранные с headroom → B предыдущие ≥today → C будущие. Человек выбирает; без авто-merge.
5. **Закрытый день** = после перевода на СГП (`completed`) — не в options.
6. **Детальный 422 detail только на `plans/build`.**
7. **Не трогать** `bot_archived/`, оптимизатор, логику подложек.

## Risks

| Risk | Mitigation |
|------|------------|
| Существующие тесты ждут max>5 / кнопку с PUT | Обновить тесты под hard cap (capacity_service, wizard addCapacityTracks) |
| Дубль `validate_fill_targets` в planning.py vs capacity.py | Persist перевести на free=day_max−occupied; не оставлять «запрошено > const 5» как единственную проверку без occupancy+cap |
| Фронт кэширует старый calendar | invalidate calendar + dayCapacity после PUT ёмкости (уже частично есть) |

## Task List

### Phase 1: Hard cap + core deficit

- [x] Task 1: Hard cap в capacity domain + repository/API
  - **Acceptance:** `get_day_capacity` / repo read clamp ≤5; `set_override` / PUT reject >5 (400); schema/pydantic `max_tracks ≤ 5`.
  - **Verify:** `pytest tests/test_production_capacity.py tests/test_day_capacity_repository.py tests/test_production_capacity_service.py tests/test_production_api_integration.py -q` (целевые кейсы cap).
  - **Files:** `core/production/capacity.py`, `core/production_capacity.py` (если нужен CONSTANT alias), `app/repositories/day_capacity_repository.py`, `app/services/production_capacity_service.py`, `app/schemas/production.py`, `app/api/v1/endpoints/production.py`, tests.
  - **Scope:** S

- [x] Task 2: `calculate_capacity_deficit` → `options[]` (A→B→C)
  - **Acceptance:** options: bump в выбранных (free); затем предыдущие ≥today не СГП; затем будущие; список не применяется сам; schema `options`; FE чекбоксы/выбор.
  - **Verify:** `pytest tests/test_production_capacity.py -q` (+ новые кейсы).
  - **Files:** `core/production/capacity.py`, `app/services/production_service.py`, `app/schemas/production.py`, `CapacityDeficitAlert.tsx`, tests.
  - **Deps:** Task 1
  - **Scope:** M

### Checkpoint A
- [x] Unit capacity: hard cap + deficit actions зелёные

### Phase 2: Build + calendar sync

- [x] Task 3: Persist/build validation = free (day_max≤5 − occupied)
  - **Acceptance:** сценарий «запрошено 6 при max 5» → PlanBuildError; override 3 → free считает от 3; build endpoint `detail=str(exc)`.
  - **Verify:** `pytest tests/test_core_production_planning.py tests/test_production_planning_service_fill_targets.py tests/test_production_api_integration.py -q`
  - **Files:** `core/production/planning.py`, `core/production/dto.py` (PersistConfig), `app/services/production_planning_service.py`, `app/api/v1/endpoints/production.py`, tests.
  - **Deps:** Task 1
  - **Scope:** M

- [x] Task 4: `days_info.max` из day-capacity (≤5)
  - **Acceptance:** calendar response `max` = capacity map; FE freeSlots совпадает с override≤5.
  - **Verify:** integration/unit на `get_global_calendar_info` + FE если есть тесты календаря.
  - **Files:** `app/services/plan_distribution_service.py`, `app/services/production_service.py`, tests.
  - **Deps:** Task 1
  - **Scope:** S

### Checkpoint B
- [x] Build с fill≤free проходит; >free — 422 с конкретным текстом; calendar max≤5

### Phase 3: Frontend deficit UX

- [x] Task 5: Wizard — убрать PUT из дефицита; bump_fill / confirm add_day
  - **Acceptance:** `addCapacityTracks` не зовёт `saveDayCapacity`; bump увеличивает fill; add_day только после confirm; типы `action`.
  - **Verify:** `cd frontend && npm test -- --run useCreatePlanWizardState CapacityDeficitAlert`
  - **Files:** `useCreatePlanWizardState.ts`, `CapacityDeficitAlert.tsx`, `types/production.ts`, tests.
  - **Deps:** Task 2 (контракт action)
  - **Scope:** M

- [x] Task 6: UI «Ёмкость» — нельзя ввести >5
  - **Acceptance:** stepper/input clamp или reject >5; согласовано с API 400.
  - **Verify:** `npm test -- --run MonthCalendarGrid` (или ручной чеклист в verify).
  - **Files:** `MonthCalendarGrid.tsx`, related tests.
  - **Deps:** Task 1
  - **Scope:** S

### Checkpoint C
- [x] FE tests зелёные; ручной сценарий: дефицит → bump/propose day → build без 422 из-за max=6

### Phase 4: Docs

- [x] Task 7: Обновить feature-doc / supersede SC-6 подложек
  - **Acceptance:** в `planirovanie-po-srokam-podlozhki` feature отмечено, что кнопка дефицита не поднимает max; hard cap 5; ссылка на эту спеку.
  - **Files:** `ai_docs/develop/features/planirovanie-po-srokam-podlozhki.md`, при необходимости idea.
  - **Deps:** Tasks 1–6
  - **Scope:** S

## Order / Parallelism

```
Task1 ──┬── Task2 ── Task5
        ├── Task3
        ├── Task4
        └── Task6
Task1–6 ── Task7
```

Task2 и Task3 можно частично параллелить после Task1; Task5 ждёт контракт action (Task2).

## Verification (done when)

```bash
.venv/bin/python -m pytest tests/test_production_capacity.py \
  tests/test_production_capacity_service.py \
  tests/test_day_capacity_repository.py \
  tests/test_core_production_planning.py \
  tests/test_production_planning_service_fill_targets.py \
  tests/test_production_api_integration.py -q

cd frontend && npm test -- --run src/features/production
cd frontend && npm run build
```

Ручной: корзина с дефицитом → кнопка не пишет 6 в day-capacity → build успешен или предлагает день → при ошибке виден конкретный текст.

## Gate

IMPLEMENT ✓ — Tasks 1–7 done; verification pytest + frontend tests/build green.
