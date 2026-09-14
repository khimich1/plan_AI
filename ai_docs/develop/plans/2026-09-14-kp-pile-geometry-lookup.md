# Implementation Plan: Заводская марка сваи → норма загрузки в КП

**Created:** 2026-09-14  
**Status:** PLAN ✅ · TASKS ✅ · IMPLEMENT ✅  
**Spec:** [`ai_docs/specs/kp-pile-geometry-lookup.md`](../../specs/kp-pile-geometry-lookup.md)  
**Idea:** [`ai_docs/ideas/kp-pile-geometry-lookup.md`](../../ideas/kp-pile-geometry-lookup.md)

## Overview

Lookup рейсов свай не узнаёт заводскую марку `С110.35-12`: каталог хранит геометрию `С110.35`. Чиним только `core/pile_catalog.py` (strip нагрузки + exact/geometry). `compute_pile_trips` и UI не меняем — pending пропадёт сам. Отгрузки, парсер цены, схема БД — вне среза.

Вертикальный срез: после Task 2 резолв уже зелёный; после Task 3 рейсы для заводской марки ready. Task 4 — регресс родителя и живой `plita.db`.

## Architecture Decisions

- **A1. Хелпер только в `pile_catalog`.** `strip_pile_load_suffix`: `-12` / `-13и` в конце марки. Не трогать `pile_line_parser` — прайсу нужна полная марка.
- **A2. Порядок `resolve_catalog_for_mark`:** (1) exact исходного ключа, (2) exact после strip, (3) bridge geometry, (4) `parse_pile_mark` (уже без суффикса) + length×section, предпочесть pcs ≥ 1.
- **A3. Мостовой regex не меняем.** `C14-40T4` не матчит `-\d+(?:и)?$`, шаг 2 его не калечит.
- **A4. Override / pending-ключ = марка строки КП**, не канон каталога.
- **A5. `С160.*`:** резолв в канон с `pcs_per_20t is None` → pending, как D13 родителя. Не считать 19800 вместо нормы.
- **A6. Нет diff отгрузок.** DoD: `shipment*.py` / logistics UI не в diff.

```
С110.35-12
    → strip → С110.35
    → pile_catalog (pcs=6)
    → compute_pile_trips (уже есть)
    → totals.pile_trip_pending_marks без этой марки
    → UI без оранжевого блока
```

## Implementation order

| Phase | Focus | Depends |
|-------|-------|---------|
| 1 | RED+GREEN parse/resolve | — |
| 2 | RED+GREEN trips на заводской марке | 1 |
| 3 | Регресс родителя + живой каталог | 2 |

## Task List

### Phase 1: Lookup

## Task 1: RED — parse и resolve заводской марки

**Description:** Зафиксировать контракт до кода. Расширить параметризацию `test_parse_pile_mark` и `test_resolve_catalog_exact_and_geometry` (или соседние тесты в том же файле): суффикс нагрузки, `-13и`, канон без изменений, мостовые без регрессии, 16 м резолвится с pcs NULL, C18 → None.

**Acceptance criteria:**
- [x] `parse_pile_mark("С110.35-12") == (11.0, 350)`
- [x] `parse_pile_mark("С120.35-13и") == (12.0, 350)`
- [x] `parse_pile_mark("С120.35")` и `С137,5.40` без регрессии
- [x] `strip_pile_load_suffix("C14-40T4") == "C14-40T4"` (не отрезать мостовую геометрию)
- [x] На фикстуре/реальном xlsx: `resolve_catalog_for_mark("С110.35-12").mark == "С110.35"` и `pcs_per_20t == 6` (если xlsx нет — fallback-каталог из `test_pile_trip_pricing._FALLBACK_CATALOG` / тот же ряд в import-тесте)
- [x] `resolve_catalog_for_mark("С160.35-12")` → mark `С160.35`, `pcs_per_20t is None`
- [x] `resolve_catalog_for_mark("C18-40T8") is None`
- [x] `C14-40T4` / `C9-35T6` без регрессии
- [x] Тесты красные: import/assert, не skip

**Verification:**
- [x] `pytest tests/test_pile_catalog_import.py -q` — fail на новых ассертах

**Dependencies:** None

**Files likely touched:**
- `tests/test_pile_catalog_import.py`

**Estimated scope:** S

## Task 2: GREEN — `strip_pile_load_suffix` + parse + resolve

**Description:** Реализовать хелпер и встроить его в `parse_pile_mark` (отрезать до split по точке) и в `resolve_catalog_for_mark` шаг 2 (exact stripped key). `parse_bridge_pile_geometry` не менять.

**Acceptance criteria:**
- [x] Все тесты Task 1 зелёные
- [x] Нет I/O, нет правок `pile_line_parser` / shipment
- [x] Канон `С60.30` и quirk запятой по-прежнему парсятся

**Verification:**
- [x] `pytest tests/test_pile_catalog_import.py -q`

**Dependencies:** Task 1

**Files likely touched:**
- `core/pile_catalog.py`
- `tests/test_pile_catalog_import.py` (только если всплыл edge — сверить со spec, не ослаблять)

**Estimated scope:** S

### Checkpoint: Lookup

- [x] Factory mark резолвится в канон с pcs
- [x] 16 м резолвится, pcs NULL
- [x] Мостовые и C18 без регрессии
- [x] Не переходить к trips, пока resolve красный/дырявый

---

### Phase 2: Рейсы (вертикальный срез «pending пропал»)

## Task 3: RED+GREEN — `compute_pile_trips` на заводской марке

**Description:** В `tests/test_pile_trip_pricing.py`: строка `С120.35-12` qty кратно pcs каталога → `ready`, `full_trips`, не в `pending_marks`. `С160.35-12` без override → pending, ключ pending = марка строки (не `С160.35`). Регресс `test_tender_without_c18_is_42_trips` и C18 не менять. Код `pile_trip_pricing.py` трогать только если тесты докажут дыру (не ожидаем).

**Acceptance criteria:**
- [x] `С120.35-12` × 10 при pcs=5 → `full_trips == 2`, `pending_marks == ()`, `ready is True`
- [x] `С160.35-12` × 4 без override → `ready is False`, `"С160.35-12" in pending_marks` (или нормализованный display той же строки)
- [x] `С160.35-12` override `{"С160.35-12": 2}` → ready, `override_trips == 2`
- [x] Тендер 42 и C18 pending зелёные

**Verification:**
- [x] `pytest tests/test_pile_trip_pricing.py tests/test_pile_catalog_import.py -q`

**Dependencies:** Task 2

**Files likely touched:**
- `tests/test_pile_trip_pricing.py`
- `core/pile_trip_pricing.py` — только если unexpectedly красный

**Estimated scope:** S

### Checkpoint: Core

- [x] Заводская марка с нормой не pending
- [x] 16 м с суффиксом всё ещё спрашивает N
- [x] UI не нужен: `pile_trip_pending_marks` кормится этим расчётом

---

### Phase 3: Регресс и живые данные

## Task 4: Регресс родителя + smoke на `plita.db`

**Description:** Прогнать соседние тесты логистики КП. Ручной readonly-скрипт из spec по живой `plita.db`. Diff не содержит shipment/logistics/parser цены.

**Acceptance criteria:**
- [x] `pytest tests/test_pile_catalog_import.py tests/test_pile_trip_pricing.py tests/test_commercial_logistics_cost.py tests/test_commercial_pile_pricing.py -q` зелёный
- [x] На `plita.db`: `С110.35-12` → `С110.35` pcs=6; `С120.35-13и` → `С120.35`; `С160.35-12` → `С160.35` pcs=None; `C11-35T6` → `С110.35`
- [x] `git diff --name-only` без `app/repositories/shipment_repository.py`, `app/services/shipment*.py`, `frontend/src/features/logistics/**`, `core/pile_line_parser.py`

**Verification:**
- [x] Команды выше
- [x] Команда из spec Commands (python `-c` resolve)

**Dependencies:** Task 3

**Files likely touched:**
- нет кода, если Task 2–3 полные; иначе точечный фикс `core/pile_catalog.py`

**Estimated scope:** XS

### Checkpoint: Complete

- [x] Success criteria spec закрыты тестами
- [x] Отгрузки не в diff
- [x] Готово к review; ручная проверка мастера КП — после merge/локального прогона визарда (не блокер плана)

---

## Risks and Mitigations

| Риск | Impact | Mitigation |
|------|--------|------------|
| Strip отрежет мостовую марку | High | Тест `strip_pile_load_suffix("C14-40T4")`; regex только `-\d+и?$` |
| Живые марки с другим дефисом | Med | Task 4 smoke; при промахе — ask first, не расширять regex наугад |
| Случайный diff отгрузок | Med | DoD Task 4 по `git diff --name-only` |
| Ослабить C18 / 16 м «чтобы не спрашивать» | High | Явные тесты pending; spec D5 |

## Parallelism

Всё последовательно: TDD lookup → trips → регресс. Параллелить нечего (один модуль).

## Open Questions

Нет. Assumptions spec подтверждены («ок делай план»).
