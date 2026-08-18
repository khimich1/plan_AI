# Планирование от ёмкости: срочные плиты + подложки

**Status**: ✅ MVP implemented  
**Date**: 2026-08-12  
**Orchestration**: `orch-2026-08-12-podlozhki`  
**Report**: [`../reports/2026-08-12-planirovanie-po-srokam-podlozhki-implementation.md`](../reports/2026-08-12-planirovanie-po-srokam-podlozhki-implementation.md)

## Purpose

Планировщик задаёт мощность (дорожки) по дням и получает предложение состава плана, которое:

1. **Закрывает сроки** — собирает позиции с дедлайном ≤ последний выбранный день.
2. **Не теряет эффективность реза** — предлагает «подложки»: поздние плиты, которые оптимизатор режет из остатков срочных.
3. **Не прячет дефицит** — если срочные не влезают в ёмкость, показывает **список вариантов** дозаполнения (bump fill / другой день); человек выбирает. **Не поднимает max дня** через `PUT /day-capacity` (см. supersede SC-6 ниже и спеку ёмкости).

**Роли:** `admin`, `production` (как `build_plan`).

## User flow

1. В календаре производства при необходимости включить режим **«Ёмкость»** и задать `max_tracks` на дни (override или default 5).
2. Открыть wizard создания плана, выбрать дни / `fill_targets`.
3. Нажать **«Найти подложки»** → `POST /api/v1/production/analyze-substrates`.
4. Просмотреть блок **«Срочные по срокам»** (галочки по умолчанию), при конфликтах дедлайнов — ⚠️.
5. Просмотреть блок **«Подложки из поздних КП»** (экономия, «нужна к», срок хранения); галочки преселектят позиции в wizard.
6. При дефиците — **CapacityDeficitAlert**: список `options[]` (A→B→C); выбор применяет bump fill или добавляет день в корзину **без** `PUT /day-capacity`.
7. Собрать план через существующий `POST /production/plans/build` (финальный оптимизатор свободен).

Hard cap завода: **0…5** дорожек/день (`max_tracks > 5` → 400; чтение clamp ≤5).

## Backend

### Schema

- Таблица `day_capacity_override` (`core/kp_db_schema.py`) — per-date `max_tracks`.
- Fallback: `TRACKS_PER_DAY_DEFAULT` (5).
- Repository: `app/repositories/day_capacity_repository.py`.

### Domain / services

| Layer | Module | Role |
|-------|--------|------|
| Domain | `core/production/capacity.py` | Pure: day capacity, `validate_fill_targets`, `calculate_capacity_deficit` |
| Domain | `core/production/urgent.py` | Pure: сбор срочных, конфликт schedule vs `execution_terms` |
| Domain | `core/production/substrate.py` | Извлечение мэтчей primary→secondary |
| Service | `production_capacity_service.py` | Overrides + map + validation |
| Service | `production_urgent_service.py` | Urgent из БД + `qty_remaining` |
| Service | `production_substrate_service.py` | Аналитический прогон + рекомендации |
| Orchestration | `production_service.analyze_substrates` | Urgent + substrate + deficit |

`qty_remaining`: `qty − Σ(qty по status='в плане' AND plan_id IS NOT NULL)`.

### API

#### `GET /api/v1/production/day-capacity?from=&to=`

**Auth:** `admin`, `production`  
**Response:** `{ "capacity": { "YYYY-MM-DD": max_tracks, ... } }` (override или default).

#### `PUT /api/v1/production/day-capacity`

**Body:** `{ "date": "YYYY-MM-DD", "max_tracks": N }` (`N` в `0…5`; `>5` → 400)  
**Response:** `{ "date", "max_tracks" }`  
**Важно:** кнопка дефицита в wizard **не** вызывает этот endpoint.

#### `POST /api/v1/production/analyze-substrates`

**Auth:** `admin`, `production` (403 для прочих)  
**Body:** `{ "fill_targets": [...], "deadline_until": "YYYY-MM-DD" }`  
**Response:** `urgent_positions`, `substrate_recommendations`, `capacity_deficit | null`, `analysis_meta`  
**Errors:** 400 (даты), 422 (пустой бэклог), 500  
**Семантика:** read-only; CPU через `run_cpu_bound`. Мутации плана — только через `plans/build`.

## Frontend

| Component | Location | Role |
|-----------|----------|------|
| Режим «Ёмкость» | `MonthCalendarGrid` / `GlobalCalendarView` | Toggle Планирование \| Ёмкость; inline edit max |
| `UrgentPositionsBlock` | `create-plan-wizard/` | Срочные, разворот деталей, конфликт ⚠️ |
| `SubstrateRecommendationsBlock` | `create-plan-wizard/` | Кнопка анализа, список рекомендаций |
| `CapacityDeficitAlert` | `create-plan-wizard/` | Дефицит + список options (человек выбирает) |

Галочки синхронизируют `selectedPlatesByKp` / `selectedPlateQtyByKp` (преселектор).

## Key invariants

1. **Рекомендации = преселектор.** Финальный `build_plan` свободен; пары из подсказки могут дрейфовать.
2. **`analyze-substrates` read-only.** Не пишет в БД; только `day-capacity` PUT (режим «Ёмкость») и `plans/build` мутируют.
3. **`fill_targets.tracks` ≤ free дня** (`free = day_max − occupied`, `day_max ≤ 5`).
4. Дедлайн: `produce_by` партии > `execution_terms` КП; конфликт >7 дней → `conflict`.
5. Подпись в UI: «Рекомендация — преселектор. Финальный состав может отличаться».
6. **SC-6 superseded (ёмкость дозаполнения, 2026-08-12):** кнопка/список дефицита **не** поднимает `max_tracks` дня. Hard cap 5. Спека: [`../../specs/emkost-dozapolneniya-fill-bez-razgona.md`](../../specs/emkost-dozapolneniya-fill-bez-razgona.md).

## How to test

```bash
# Backend — ключевые
.venv/bin/python -m pytest tests/test_day_capacity_repository.py tests/test_production_capacity.py \
  tests/test_production_urgent.py tests/test_production_capacity_service.py \
  tests/test_production_urgent_service.py tests/test_production_substrate_service.py \
  tests/test_production_api_integration.py tests/test_production_podlozhki_e2e.py -q

# Phase 0 (реальная plita.db)
.venv/bin/python scripts/validate_podlozhki_phase0.py --db plita.db \
  --report ai_docs/develop/reports/2026-08-12-podlozhki-phase0.md

# Frontend
cd frontend && npm test -- --run src/features/production
cd frontend && npm run build
```

## Related docs

- Idea: [`../../ideas/planirovanie-po-srokam-podlozhki.md`](../../ideas/planirovanie-po-srokam-podlozhki.md)
- Spec: [`../../specs/planirovanie-po-srokam-podlozhki.md`](../../specs/planirovanie-po-srokam-podlozhki.md)
- Ёмкость / deficit (hard cap 5, без PUT из дефицита): [`../../specs/emkost-dozapolneniya-fill-bez-razgona.md`](../../specs/emkost-dozapolneniya-fill-bez-razgona.md)
- Plan: [`../plans/2026-08-12-planirovanie-po-srokam-podlozhki.md`](../plans/2026-08-12-planirovanie-po-srokam-podlozhki.md)
- Phase 0: [`../reports/2026-08-12-podlozhki-phase0.md`](../reports/2026-08-12-podlozhki-phase0.md)
- Implementation: [`../reports/2026-08-12-planirovanie-po-srokam-podlozhki-implementation.md`](../reports/2026-08-12-planirovanie-po-srokam-podlozhki-implementation.md)
- Соседняя фича: [`delivery-schedule.md`](./delivery-schedule.md)
