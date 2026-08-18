# Spec: Планирование от ёмкости — срочные плиты + подложки из поздних КП

Дата: 2026-08-12. Статус: draft, на ревью.
Идея: [`ai_docs/ideas/planirovanie-po-srokam-podlozhki.md`](../ideas/planirovanie-po-srokam-podlozhki.md) (2026-08-10).
Связанные спеки: [`calendar-first-planning.md`](./calendar-first-planning.md), [`delivery-schedule.md`](../../../docs/specs/delivery-schedule.md).

## ASSUMPTIONS I'M MAKING

1. Пользователь с ролью `admin` или `production` (как `build_plan`) видит кнопку «Найти подложки» и блоки срочных/подложек. `manager` — только просмотр готового плана, без wizard.
2. API — отдельный endpoint `POST /production/analyze-substrates`, не расширение существующего `POST /production/plans/build` (разная семантика: анализ vs мутация).
3. `qty_remaining` для анализа = `kp_plates.qty − Σ(qty по status='в плане' с plan_id IS NOT NULL)` (незапланированный остаток позиции).
4. Мощность дня (`max_tracks`) по умолчанию 5 (`TRACKS_PER_DAY_DEFAULT`), переопределения хранятся per-date, могут быть и >5 (ручное завышение при дефиците).
5. Аналитический прогон оптимизатора — CPU-bound, вызывается через `run_cpu_bound` (как `build_plan`), блокирует запрос на время выполнения (секунды–десятки секунд, подтверждается Phase 0).
6. Дата производства подложки для расчёта срока хранения = первый день `fill_targets` (начало плана).

## Objective

Планировщик производства, задавая мощность (дорожки) на дни, получает автоматическое предложение состава плана, которое:
1. **Закрывает сроки** — собирает позиции с дедлайном ≤ последний выбранный день (партии графика поставки приоритетно, `execution_terms` как fallback).
2. **Не теряет эффективность реза** — находит поздние плиты, которые оптимизатор фактически режет из остатков срочных, и предлагает их как «подложки».
3. **Не прячет дефицит** — если срочные не влезают в ёмкость, показывает «+N дорожек до <дата>» с возможностью увеличить мощность дня.

**Пользователь:** планировщик / мастер производства (роли `admin`, `production`).

**Критерий успеха MVP:** менеджер принимает или редактирует предложенный список (не игнорирует); снижение отхода полосы на измеримую величину (по сравнению с ручным планированием).

## Tech Stack

- Backend: Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`), PuLP/CBC (существующий оптимизатор).
- Frontend: React 18, TypeScript, Vite, TanStack Query, feature-структура `frontend/src/features/production/`.
- Новых внешних зависимостей нет.

## Commands

```bash
# Backend dev
uvicorn app.main:app --reload

# Backend tests
pytest tests/ -q
pytest tests/test_production_planning_service_fill_targets.py -q
pytest tests/test_delivery_schedule_service.py -q

# Frontend dev / test / build
cd frontend && npm run dev
cd frontend && npm test -- --run
cd frontend && npm run build

# Phase 0 validation (первый шаг до UI)
python scripts/validate_podlozhki_phase0.py --db plita.db --report ai_docs/develop/reports/2026-08-12-podlozhki-phase0.md
```

## Project Structure

```
app/
  api/v1/endpoints/production.py          → +POST /production/analyze-substrates
  services/
    production_service.py                 → +analyze_substrates() orchestration
    production_urgent_service.py          → NEW: сбор позиций с дедлайном
    production_substrate_service.py       → NEW: аналитический прогон + извлечение мэтчей
    production_capacity_service.py        → NEW: per-day ёмкость (overrides + fallback)
  repositories/
    day_capacity_repository.py            → NEW: CRUD переопределений ёмкости
    kp_repository.py                      → +qty_remaining для позиций
  schemas/production.py                   → +AnalyzeSubstratesRequest/Response, UrgentPosition, SubstrateRecommendation

core/
  production/
    capacity.py                           → NEW: логика ёмкости (clean core, без I/O)
    urgent.py                             → NEW: дедлайны позиций (produce_by → execution_terms)
  optimization/optimize_2d/finalize.py    → возможно +rests_unused (open question)
  delivery_schedule_check.py              → reuse для дефицита «+N дорожек»

frontend/src/features/production/
  components/
    MonthCalendarGrid.tsx                 → +режим «Ёмкость» (toggle)
    FillBasket.tsx                        → отображение ёмкости vs плана
    CreatePlanWizard.tsx                  → +UrgentPositionsBlock, +SubstrateRecommendationsBlock
    create-plan-wizard/
      UrgentPositionsBlock.tsx            → NEW: список срочных с разворотом
      SubstrateRecommendationsBlock.tsx   → NEW: список подложек с датой «нужна к» и сроком хранения
      CapacityDeficitAlert.tsx            → NEW: «+N дорожек»
  hooks/useProductionQueries.ts           → +useAnalyzeSubstratesMutation, +useDayCapacityQuery

scripts/
  validate_podlozhki_phase0.py            → NEW: валидация A1–A3 на реальной БД
```

## Code Style

```python
# app/services/production_urgent_service.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any

from app.repositories.kp_repository import KpRepository
from core.execution_terms import parse_execution_terms_to_datetime


@dataclass(frozen=True, slots=True)
class UrgentPosition:
    plate_id: int
    kp_id: int
    plate_name: str
    qty_remaining: int
    deadline: date
    deadline_source: str  # "delivery_batch" | "execution_terms"
    deadline_details: list[dict[str, Any]]  # для разворота в UI
    conflict: str | None  # "schedule_earlier" | "kp_earlier" | None


class ProductionUrgentService:
    """Собирает позиции с дедлайном ≤ заданной даты."""

    def __init__(self, *, kp_repository: KpRepository | None = None) -> None:
        self.kp_repository = kp_repository or KpRepository()

    def list_urgent_positions(self, *, deadline_until: date) -> list[UrgentPosition]:
        ...
```

Конвенции:
- Роутер → сервис → репозиторий; сервис не содержит SQL.
- `from __future__ import annotations`; dataclass `frozen=True, slots=True` для DTO.
- Чистый core (`core/production/urgent.py`) без импортов `app.*`.
- Pydantic-схемы в `app/schemas/production.py`, `model_config = ConfigDict(extra="forbid")` для запросов.

## Testing Strategy

| Уровень | Что покрывает | Команда |
|---------|---------------|---------|
| Unit (pytest) | `core/production/urgent.py` (парсинг дедлайнов, агрегат, конфликт), `core/production/capacity.py` (fallback, per-day max) | `pytest tests/test_production_urgent.py -q` |
| Unit (pytest) | `production_substrate_service.py` — извлечение мэтчей из мокового `optimize` результата | `pytest tests/test_production_substrate_service.py -q` |
| Integration (pytest) | `POST /production/analyze-substrates` — схема запроса/ответа, AuthZ, интеграция с репозиториями | `pytest tests/test_production_api_integration.py -q` |
| Unit (vitest) | `UrgentPositionsBlock`, `SubstrateRecommendationsBlock` — рендер, разворот, галочки, qty input | `cd frontend && npm test -- --run` |
| Smoke (pytest) | Существующие `fill_targets` тесты не ломаются | `pytest tests/test_production_planning_service_fill_targets_smoke.py -q` |

Coverage expectation: новые сервисы ≥80% покрытие ветвлений.

## Boundaries

- **Always:** TDD — тест на логику до реализации; `pytest` зелёный перед коммитом; минимальный diff — не трогать несвязанный код; `.cursor/` и `ai_docs/` в `.gitignore` — локальные артефакты агента.
- **Ask first:** новая таблица `day_capacity_overrides` (миграция `kp_db_schema.py`); изменение `core/optimization/optimize_2d/finalize.py` (добавление `rests_unused`); изменение существующих API-контрактов `KpCandidatesResponse`.
- **Never:** не коммитить без явной просьбы; не менять `MAX_TRACKS_PER_DAY` глобально (только per-day override); не ломать frozen bot paths (`bot_archived/`).

## Success Criteria

1. **SC-1 (A1 validated):** Phase 0 скрипт на реальной `plita.db` находит ≥1 cross-KP мэтча (secondary cut с `kp_id` ≠ primary `kp_id`). Если 0 — MVP не строится.
2. **SC-2 (A2 validated):** Время полного прогона оптимизатора на бэклоге «в производстве» ≤ 30 секунд (интерактивная кнопка). Если >120 сек — переход на фоновый режим (не MVP).
3. **SC-3 (A3 validated):** `execution_terms` парсится у ≥80% КП «в работе» без графика поставки. Иначе fallback-дедлайны ненадёжны.
4. **SC-4 (UX):** В wizard создания плана блок «Срочные по срокам» отображает позиции с дедлайном ≤ последний выбранный день; разворот показывает источники дат; конфликт >7 дней отмечен ⚠️.
5. **SC-5 (UX):** Блок «Подложки из поздних КП» по кнопке «Найти подложки» показывает рекомендации с датой «нужна к», сроком хранения и экономией мм; галочки преселектятся в `selectedPlatesByKp`.
6. **SC-6 (Capacity):** При дефиците ёмкости отображается «+N дорожек до <дата>»; кнопка увеличивает `max` и `fill_targets` на выбранный день.
7. **SC-7 (Regression):** Существующий flow calendar brush → fill_targets → build plan работает без изменений; все существующие тесты зелёные.

## Data Model

Новая таблица в `core/kp_db_schema.py` (стиль `CREATE TABLE IF NOT EXISTS`):

```sql
day_capacity_override (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  date TEXT NOT NULL UNIQUE,           -- ISO YYYY-MM-DD
  max_tracks INTEGER NOT NULL CHECK (max_tracks >= 0),
  updated_at TEXT NOT NULL,
  updated_by TEXT                      -- user login, optional
)
```

Fallback: дата вне таблицы → `TRACKS_PER_DAY_DEFAULT` из `core/production_capacity.py`.

**Унификация констант:** `plan_calendar.py` импортирует `TRACKS_PER_DAY_DEFAULT` из `core/production_capacity.py` вместо локального `MAX_TRACKS_PER_DAY` (фикс drift из аудита delivery-schedule A4).

## API

### POST /production/analyze-substrates

Запрос:

```json
{
  "fill_targets": [{"date": "2026-08-12", "tracks": 5}, {"date": "2026-08-13", "tracks": 3}],
  "deadline_until": "2026-08-15"
}
```

Ответ 200:

```json
{
  "urgent_positions": [
    {
      "plate_id": 123,
      "kp_id": 115,
      "plate_name": "ПБ 57-7,2 ×8п",
      "qty_remaining": 2,
      "deadline": "2026-08-15",
      "deadline_source": "delivery_batch",
      "deadline_details": [
        {"type": "delivery_batch", "batch_name": "3 этаж", "deadline": "2026-08-15", "qty": 1},
        {"type": "delivery_batch", "batch_name": "4 этаж", "deadline": "2026-08-20", "qty": 1},
        {"type": "execution_terms", "deadline": "2026-08-26", "qty": 2}
      ],
      "conflict": "schedule_earlier"
    }
  ],
  "substrate_recommendations": [
    {
      "plate_id": 456,
      "kp_id": 127,
      "plate_name": "ПБ 57-4,8 ×8п",
      "qty_recommended": 3,
      "under_plate_id": 123,
      "under_kp_id": 115,
      "under_plate_name": "ПБ 57-7,2 ×8п",
      "needed_by": "2026-09-05",
      "storage_days": 24,
      "saving_mm": 480
    }
  ],
  "capacity_deficit": {
    "tracks_needed": 20,
    "tracks_available": 18,
    "tracks_missing": 2,
    "deficit_until": "2026-08-15",
    "suggestion": {"date": "2026-08-14", "add_tracks": 2}
  },
  "analysis_meta": {
    "orders_count": 127,
    "analysis_duration_ms": 8300,
    "optimization_status": "ok"
  }
}
```

Ошибки:
- `400` — невалидные даты, `deadline_until` раньше первого `fill_targets.date`.
- `403` — роль не `admin`/`production`.
- `422` — нет позиций «в производстве» для анализа.

## Open Questions

1. Добавлять ли явный `rests_unused` в `optimize_2d/finalize.py` для точного расчёта «экономии метров» в UI, или достаточно `rests_created` − `rests_used`?
2. При повторном открытии wizard с теми же `fill_targets` — кэшировать результат анализа (TTL 5 мин) или пересчитывать каждый раз?
3. Нужен ли экспорт рекомендаций в XLSX/PDF для обсуждения с мастером цеха (вне MVP)?
