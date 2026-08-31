# Spec: Domain enums SSOT in core (audit A1 + A6)

> **Источник:** [`ai_docs/develop/audits/2026-08-02-full-project-audit.md`](../develop/audits/2026-08-02-full-project-audit.md) — findings **A1** (critical), **A6** (high)  
> **Фаза SDD:** SPECIFY ✅ → IMPLEMENT ✅  
> **Статус:** APPROVED  
> **Дата:** 2026-08-02  
> **Scope:** перенос всех 6 доменных enum в `core/domain/enums.py`; re-export в `app/domain/enums.py`; устранение строковых дублей в `core/kp_db_shipments.py`. Без миграции импортов в `app/services` и тестах; без правок audit FIXED; без фронтенда и A2–A5.

---

## Objective

Убрать нарушение направления зависимостей `core → app` и дублирование строковых констант статусов/типов отгрузки: единый источник правды (SSOT) для доменных enum — слой `core`.

**Проблема сейчас:**
- `core/kp_db_plates_common.py` и `core/kp_db_plates_completion.py` импортируют `app.domain.enums` (A1).
- `core/kp_db_shipments.py` дублирует значения `ShipmentItemType` / `ShipmentStatus` локальными строками (A6).

**Результат:** `core` не импортирует `app`; все 6 enum живут в `core/domain/enums.py`; `app.domain.enums` — только re-export с той же object identity.

### Enums in scope

| Enum | Назначение |
|------|------------|
| `PlateStatus` | жизненный цикл плиты |
| `KpStatus` | статус КП |
| `PlateTransitionReason` | причина перехода в audit |
| `ShipmentStatus` | статус рейса |
| `DeliveryType` | доставка / самовывоз |
| `ShipmentItemType` | plate / free в составе рейса |

### Out of scope (Not Doing)

- Массовая миграция `from app.domain.enums import …` в `app/services`, тестах, API
- Пометка A1/A6 как FIXED в audit-файле
- Frontend, A2–A5, новые enum-значения или смена строковых `.value`

---

## Commands

```bash
# Архитектурный гейт: core не импортирует app
rg 'from app\.|import app\.' core/ --glob '*.py'

# Identity + регрессия затронутых доменов
pytest tests/test_domain_enums_location.py \
  tests/test_sgp_schema.py \
  tests/test_sgp_service.py \
  tests/test_shipment_service.py \
  tests/test_kp_readiness_service.py -q
```

Success: `rg` — 0 совпадений; pytest — зелёный.

---

## Structure

```
core/domain/
  enums.py                 → NEW SSOT: все 6 enum (+ docstring as-is)
  __init__.py              → без обязательного re-export enums (submodule import OK)

app/domain/
  enums.py                 → re-export from core.domain.enums + __all__

core/kp_db_plates_common.py      → import from core.domain.enums
core/kp_db_plates_completion.py  → import from core.domain.enums
core/kp_db_shipments.py          → ShipmentItemType / ShipmentStatus; без _ITEM_* / _STATUS_*

tests/test_domain_enums_location.py → NEW: identity assert для всех 6
```

---

## Code Style

- Значения строк enum **не менять** (русские статусы КП/плит, латинские shipment).
- `app/domain/enums.py` — только re-export, без дублирующих `class`-определений (сохранить `is`-identity).
- В `core/kp_db_shipments.py` использовать `EnumMember.value` в SQL-параметрах.
- Существующие импорты `from app.domain.enums import …` в app/tests остаются валидны.

```python
# app/domain/enums.py (pattern)
"""Re-export. SSOT: core.domain.enums."""
from core.domain.enums import (
    DeliveryType,
    KpStatus,
    PlateStatus,
    PlateTransitionReason,
    ShipmentItemType,
    ShipmentStatus,
)

__all__ = [
    "DeliveryType",
    "KpStatus",
    "PlateStatus",
    "PlateTransitionReason",
    "ShipmentItemType",
    "ShipmentStatus",
]
```

---

## Testing

| Уровень | Что | Где |
|---------|-----|-----|
| Identity | `app.domain.enums.X is core.domain.enums.X` для всех 6 | `tests/test_domain_enums_location.py` |
| Regression | SGP schema/service, shipment, KP readiness | перечисленные pytest-файлы |
| Architecture | нет `from app` / `import app` в `core/**/*.py` | `rg` |

---

## Boundaries

- **Always:** не менять `.value` строк; re-export с той же identity; pytest + `rg` после правок
- **Ask first:** миграция импортов по всему `app/` / тестам; изменение audit FIXED-маркеров
- **Never:** дублировать class-определения enum в app; коммит без явной просьбы; трогать frontend / A2–A5

---

## Success Criteria

1. `rg 'from app\.|import app\.' core/ --glob '*.py'` → 0 matches.
2. Все 6 enum определены только в `core/domain/enums.py`; `app/domain/enums.py` — re-export.
3. `core/kp_db_shipments.py` без `_ITEM_PLATE` / `_STATUS_IN_WORK` / `_STATUS_DONE`.
4. Identity-тест зелёный; регрессионный набор pytest зелёный.
5. Audit-файл **не** помечен FIXED в этом изменении.

---

## Open Questions (resolved)

| # | Вопрос | Решение |
|---|--------|---------|
| Q1 | Мигрировать импорты в `app/services` и тестах на `core.domain.enums`? | **No** — оставляем `from app.domain.enums import …` |
| Q2 | Нужен identity-тест `app.X is core.X`? | **Yes** — `tests/test_domain_enums_location.py` |
| Q3 | Обновлять audit FIXED сейчас? | **No** — не трогать audit-маркеры в этом изменении |
| Q4 | Имя файла спеки? | **`ai_docs/specs/core-domain-enums-a1.md`** |
