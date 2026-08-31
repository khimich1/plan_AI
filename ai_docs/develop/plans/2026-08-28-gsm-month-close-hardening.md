# Implementation Plan: ГСМ — hardening закрытия месяца

Дата: 2026-08-28. Статус: этап 1 IMPLEMENT выполнен (тесты зелёные); live UI и коммит — по просьбе.  
Спека: [`../../specs/gsm-month-close-hardening.md`](../../specs/gsm-month-close-hardening.md).
Отчёт: [`../reports/2026-08-28-gsm-month-close-hardening-stage1.md`](../reports/2026-08-28-gsm-month-close-hardening-stage1.md).

## Overview

Сначала **только этап 1**: сервер повторяет замки комплекта (хвост, цепь, красные) на usage/export/generate; бейдж не прячет красные; кнопка «Экспорт» не обходит `planKit` в обход сервера. Этапы 2–4 — после живой проверки июля/августа на `/gsm`. Солвер и live `plita.db` не трогаем.

## Architecture Decisions

- Один helper `evaluate_kit_vehicle` / список eligibility; report и export его зовут **до** zip и **до** flip `exported`.
- Generate: хвост чужого месяца → skip/4xx; `chain_broken` → generate идёт.
- Skip-причины при смешанном флоте — только UI (`planKit`). Сервер молча не включает больных в zip.
- Норма дня, импорт, soffice — **не в этом срезе кода**.

```
fleet_overview / _chain_broken / red в периоде
        │
        ▼
gsm_kit_gate.evaluate(...)
        │
        ├── POST /report/usage
        ├── POST /waybills/export
        └── generate / generate_bulk (хвост да, цепь нет)
```

## Task List — этап 1 (первый код)

### Task 1: Helper гейта + pytest на фикстурах overview

**Description:** Вынести правила спеки в чистую функцию над данными обзора (или один vehicle_id + period → eligibility). Красные / хвост / цепь как в `planKit` + таблица generate.

**Acceptance:**
- [x] Юнит: july tail + период август → `gsm_kit_tail`, не allowed
- [x] Юнит: период = `open_before_month` → хвост не блокирует комплект
- [x] Юнит: `chain_broken` → комплект forbidden, generate allowed
- [x] Юнит: red в периоде → комплект forbidden, generate allowed

**Verification:** `.venv/bin/python -m pytest tests/test_gsm_kit_gate.py -q`

**Dependencies:** нет  
**Files:** `app/services/gsm_kit_gate.py` (new), `tests/test_gsm_kit_gate.py` (new)  
**Scope:** S

### Task 2: Usage-report и прямой export применяют гейт

**Description:** `build_usage_zip` фильтрует машины до soffice и до `export_zip`. `export_zip` сам отсекает ineligible (прямой API). `vehicle_ids: null` — все активные, затем фильтр. Ноль прошедших → существующий `gsm_report_no_data` / `gsm_export_empty`. Одна запрошенная плохая → 4xx, статусы не `exported`.

**Acceptance:**
- [x] pytest: август, Monjaro хвост июля, Palisade чистая, `null` → Palisade exported, Monjaro нет, zip без Monjaro
- [x] pytest: `vehicle_ids: [monjaro]` → не 200 с её exported
- [x] pytest: прямой export августа Monjaro с хвостом — не flip
- [x] Регрессия мая 848 зелёная

**Verification:** `.venv/bin/python -m pytest tests/test_gsm_usage_report.py tests/test_gsm_export.py -q`

**Dependencies:** Task 1  
**Files:** `app/services/gsm_report_service.py`, `app/services/gsm_export_service.py`, `tests/test_gsm_usage_report.py`, `tests/test_gsm_export.py`  
**Scope:** M

### Task 3: Generate / bulk — хвост стоп, цепь нет

**Description:** `generate` и `generate_bulk`: open_before на другом месяце → ошибка/per-id skip с тем же текстом, что UI. `chain_broken` не стоп.

**Acceptance:**
- [x] pytest: generate(август) при july open_before → 4xx/skip, нет новых draft августа
- [x] pytest: generate(август) при только chain_broken → успех
- [x] bulk: Palisade ок, Monjaro хвост → per-id, Palisade не откатывается

**Verification:** `.venv/bin/python -m pytest tests/test_gsm_generation_api.py -q`

**Dependencies:** Task 1  
**Files:** `app/services/gsm_generation_service.py`, `tests/test_gsm_generation_api.py`  
**Scope:** S

### Task 4: Статус обзора + кнопка Экспорт

**Description:** `_status_of`: `red_days > 0` → `has_red_days` раньше `needs_generation`. `handleExportKit` не вызывает `runKit` в обход `planKit` для текущего периода; прыжок на хвост остаётся, отказ API показывается.

**Acceptance:**
- [x] pytest overview: red + tx позже последнего ПЛ → `has_red_days`
- [x] vitest: период-отчёт не шлёт id из excluded `planKit`
- [x] vitest: Export при хвосте идёт в границы хвоста; при 4xx — ошибка, не «скачан zip»

**Verification:**
```
.venv/bin/python -m pytest tests/test_gsm_overview_api.py -q
cd frontend && npx vitest run src/features/gsm/components/FleetOverviewView.test.tsx src/features/gsm/lib/exportGate.test.ts
```

**Dependencies:** Task 2 (контракт 4xx)  
**Files:** `app/services/gsm_overview_service.py`, `tests/test_gsm_overview_api.py`, `frontend/src/features/gsm/components/FleetOverviewView.tsx`, `frontend/.../FleetOverviewView.test.tsx`  
**Scope:** M

## Checkpoint: этап 1

- [x] `tests/test_gsm_*.py` релевантные зелёные; `test_gsm_generator.py` без сюрпризов
- [x] vitest обзора зелёный
- [ ] Live UI **по просьбе**: август Monjaro хвост / Palisade чистая
- [ ] Коммит только по просьбе; этапы 2–4 не начинать без «этап 1 ок»

## Этапы 2–4 (не в первом IMPLEMENT)

| Этап | Когда | Суть |
|------|--------|------|
| 2 | после живого гейта | `norm_l_per_100` на ПЛ; round; registry не глотает switches |
| 3 | после 2 или параллельно с 2 (другой код) | tx на файл; 10 / 20k / 100 МБ; 400 на не-xls |
| 4 | после боевого комплекта на сервере | пакетный soffice, затем job в RAM + опрос + 429 |

Отдельные планы-задачи, когда дойдём.

## Risks and Mitigations

| Риск | Impact | Mitigation |
|------|--------|------------|
| Два SQL обзора разъедутся с гейтом | High | Гейт читает те же поля, что overview |
| Generate на месяце хвоста должен проходить | High | Тест: periodYm == open_before_month → allowed |
| `handleExportKit` прыжок июля без данных июля на клиенте | Med | Сервер — замок; UI показывает 4xx |
| Слишком широкий diff | Med | Строго этап 1, без season/import/soffice |

## Open Questions

Нет. Владелец: 1А 2А 3А 4А, аудит репо отдельно.
