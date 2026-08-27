# Implementation Plan: Визуал ёмкости завода + жёсткий гейт

> **Spec:** [`ai_docs/specs/zavod-emkost-vizual-gate.md`](../../specs/zavod-emkost-vizual-gate.md)  
> **Idea:** [`ai_docs/ideas/zavod-emkost-vizual-gate.md`](../../ideas/zavod-emkost-vizual-gate.md)  
> **Handoff:** [`ai_docs/develop/handoffs/2026-08-25-zavod-emkost-vizual-gate.md`](../handoffs/2026-08-25-zavod-emkost-vizual-gate.md)  
> **Orchestration:** `orch-2026-08-25-16-50-zavod-emkost-gate`  
> **Дата:** 2026-08-25  
> **Статус:** PLAN ✅ — готов к `/orchestrate`  
> **Не коммитить** без явной просьбы

---

## Overview

Менеджер на ПК видит мини-календарь загрузки завода + «нужно / свободно / Δ» в двух
местах: «В производство» (архив) и редактор графика поставок. При **red** —
сохранить/отправить нельзя (UI + backend); снять блок можно только изменив срок
в **существующем** поле (строка дней / `produce_by`). Override-галочек и новых
таблиц БД нет. Занятость — только план (`days_info`). Симуляция с **завтра**.

---

## Architecture Decisions (locked)

| ID | Решение |
|----|---------|
| D1 | Занятость = `max − occupied` из плана. СГП out. |
| D2 | Старт симуляции / окна = **завтра**. `check_batches(..., start_date=)`. |
| D3 | Пороги = существующие (green slack 5, buffer 1.15, 101 м, 5 дор/день). |
| D4 | Yellow = warning; red = hard block. Нет override / audit table. |
| D5 | Multi-batch: блок если **любая** партия red. |
| D6 | Поле срока move **без изменений** (одна строка). Target date = результат
     существующего `parse_execution_terms` (как сейчас: календарные дни/дата).
     Не переписывать «N дней» на рабочие дни в этой фиче. |
| D7 | Календарь UI: месяцы от **текущего** до **месяца target** (max produce_by). |
| D8 | Гейт на backend обязателен. |
| D9 | Один алгоритм: `core/delivery_schedule_check.check_batches`. |
| D10 | Бронь дорожек не делаем. |

```
execution_terms / produce_by
        │
        ▼
target_date (parse as today)
        │
        ▼
check_batches(start=tomorrow, occupancy=days_info, …)
        │
        ├── status green/yellow → allow
        └── status red         → 4xx + hint; UI disable + FactoryCapacityPanel
```

---

## Components

| Компонент | Роль |
|-----------|------|
| `core/delivery_schedule_check.py` | `start_date` вместо жёсткого today |
| `app/services/capacity_gate_service.py` | snapshot + enforce |
| `GET .../capacity-snapshot` | данные виджета для move |
| `archive_service.move_to_production` | gate before commit |
| `delivery_schedule_service` PUT | gate if any red |
| `frontend/.../factory-capacity/` | Panel + MiniCalendar + api/hooks |
| `MoveToProductionDialog` | вставка панели, disable submit |
| `DeliveryScheduleDialog/Editor` | вставка панели, disable save |

---

## Task List

### Phase 0: Core math

### CAP-001 — `start_date` в `check_batches`

**type:** `feat-be` · **dependsOn:** [] · **pipeline:** explore → worker → test-writer → test-runner → reviewer  
**securitySensitive:** false · **needsExplore:** true

**Description:** Параметр `start_date` (ISO) в `check_batches` / симуляции окна;
дефолт = сегодняшний день для обратной совместимости вызовов, новые callers
передают завтра. Тесты на старт с tomorrow.

**Acceptance:**
- [ ] Существующие тесты check зелёные с дефолтом
- [ ] Новый кейс: start=tomorrow не учитывает ёмкость «сегодня»
- [ ] Сигнатура задокументирована в docstring модуля

**Verify:** `pytest tests/test_delivery_schedule_check.py -q`

**Files:** `core/delivery_schedule_check.py`, `tests/test_delivery_schedule_check.py`

**Scope:** S

---

### CAP-002 — Чистая сборка snapshot (tracks / free / status / Δ)

**type:** `feat-be` · **dependsOn:** [CAP-001] · **pipeline:** worker → test-writer → test-runner → reviewer

**Description:** Хелпер (в `core/` или тонкий модуль рядом) / начало
`capacity_gate_service`: из occupancy + workdays + items + target →
`tracks_needed`, `tracks_free_in_window`, `delta`, `status`, `hint`,
`calendar_from_month`, `calendar_to_month`. Без FastAPI.

**Acceptance:**
- [ ] Red при дефиците до target; yellow/green по порогам партий
- [ ] Месяцы: текущий … месяц target
- [ ] KP как одна виртуальная партия работает

**Verify:** `pytest tests/test_capacity_gate.py -q`

**Files:** `app/services/capacity_gate_service.py` и/или `core/…`, `tests/test_capacity_gate.py`

**Scope:** M

---

### Checkpoint: Core
- [ ] CAP-001, CAP-002 green

---

### Phase 1: Backend API + enforce

### CAP-003 — GET `capacity-snapshot`

**type:** `feat-be` · **dependsOn:** [CAP-002] · **pipeline:** worker → test-writer → test-runner → reviewer  
**securitySensitive:** true (roles admin/manager)

**Description:** `GET /api/v1/commercial/archive/{kp_id}/capacity-snapshot`
с query `target` (ISO) или выводом target из текущего срока КП. Ответ по
контракту спеки. Roles как у архива.

**Acceptance:**
- [ ] 200 + snapshot поля
- [ ] 404 unknown kp; 403 wrong role
- [ ] days_info только план

**Verify:** `pytest tests/test_archive_endpoints.py -q` (или новый `test_capacity_snapshot_endpoints.py`)

**Files:** `app/schemas/…`, `app/api/v1/endpoints/archive.py`, service, tests

**Scope:** M

---

### CAP-004 — Gate на move-to-production

**type:** `feat-be` · **dependsOn:** [CAP-002] · **pipeline:** worker → test-writer → test-runner → reviewer  
**securitySensitive:** true

**Description:** Перед `commit_move_to_production` пересчёт статуса по
`execution_terms`. Red → `ArchiveValidationError` / 4xx с hint. Без новых
полей в `MoveToProductionRequest`.

**Acceptance:**
- [ ] Red → не меняет статус КП; 4xx
- [ ] Yellow/green → как сейчас успех
- [ ] Сообщение содержит понятный hint

**Verify:** `pytest tests/test_archive_endpoints.py tests/test_capacity_gate.py -q`

**Files:** `app/services/archive_service.py`, endpoints/tests

**Scope:** M

---

### CAP-005 — Gate на PUT delivery-schedule

**type:** `feat-be` · **dependsOn:** [CAP-001, CAP-002] · **pipeline:** worker → test-writer → test-runner → reviewer  
**securitySensitive:** true

**Description:** PUT графика: `check_batches` со `start_date=tomorrow`; если
любой batch red → reject, график не сохраняется. GET/светофор партий используют
тот же start (согласованность).

**Acceptance:**
- [ ] Any red → 4xx, БД без изменений
- [ ] All ≤ yellow → save ok
- [ ] Чипы и gate не расходятся по статусу

**Verify:** `pytest tests/test_delivery_schedule_service.py tests/test_delivery_schedule_endpoints.py -q`

**Files:** `app/services/delivery_schedule_service.py`, tests

**Scope:** M

---

### Checkpoint: Backend enforce
- [ ] CAP-003…005 green; ручной curl red→4xx

---

### Phase 2: Frontend widget

### CAP-006 — Feature `factory-capacity`: types, api, hooks

**type:** `feat-fe` · **dependsOn:** [CAP-003] · **pipeline:** worker → test-writer → test-runner → reviewer

**Description:** Папка `frontend/src/features/factory-capacity/` — типы
`CapacitySnapshot`, `capacityApi.getSnapshot`, TanStack hook.

**Acceptance:**
- [ ] Типы совпадают со схемой API
- [ ] Hook не дергается при `kpId=null`

**Verify:** `cd frontend && npm run test -- src/features/factory-capacity` && `npm run typecheck`

**Files:** `factory-capacity/api|hooks|types/**`

**Scope:** S

---

### CAP-007 — `FactoryMiniCalendar` + `FactoryCapacityPanel`

**type:** `feat-fe` · **dependsOn:** [CAP-006] · **pipeline:** worker → test-writer → test-runner → reviewer  
**needsExplore:** true (MonthCalendarGrid reuse vs mini)

**Description:** Read-only мини-календарь (месяцы current…target) + шапка
нужно/свободно/Δ + Alert hint при red. Без редактирования ёмкости, без
DayDrawer. Ask first только если нужен крупный рефактор грида — иначе
отдельный compact.

**Acceptance:**
- [ ] Цвета empty/partial/full; выходные отличимы
- [ ] Навигация только в диапазоне месяцев
- [ ] Red → hint видим
- [ ] ПК layout (сайд ок)

**Verify:** vitest компонентов + typecheck

**Files:** `factory-capacity/components/*`

**Scope:** L

---

### Checkpoint: Widget
- [ ] CAP-006, CAP-007 green

---

### Phase 3: Wire both entry points

### CAP-008 — Вставка в `MoveToProductionDialog`

**type:** `feat-fe` · **dependsOn:** [CAP-004, CAP-007] · **pipeline:** worker → test-writer → test-runner → reviewer

**Description:** Панель рядом/слева от формы. Snapshot от строки срока
(debounce). Submit disabled при red. Поле срока без новых инпутов.

**Acceptance:**
- [ ] При red кнопка неактивна + hint
- [ ] Увеличение дней → пересчёт → можно отправить (если не red)
- [ ] Estimate alert сохраняется

**Verify:** `cd frontend && npm run test -- MoveToProduction` / commercial-archive + typecheck

**Files:** `MoveToProductionDialog.tsx`, tests

**Scope:** M

---

### CAP-009 — Вставка в график поставок

**type:** `feat-fe` · **dependsOn:** [CAP-005, CAP-007] · **pipeline:** worker → test-writer → test-runner → reviewer

**Description:** Тот же `FactoryCapacityPanel` в `DeliveryScheduleDialog` /
Editor. Save disabled при любом red (по данным светофора/summary). Live при
смене партий.

**Acceptance:**
- [ ] Виджет на этапе согласования поставок
- [ ] Save blocked на red; ок на yellow/green
- [ ] Цвета согласованы с `BatchStatusChip`

**Verify:** `cd frontend && npm run test -- src/features/delivery-schedule` && typecheck

**Files:** `DeliveryScheduleDialog.tsx`, `DeliveryScheduleEditor.tsx`, tests

**Scope:** M

---

### CAP-010 — Регрессия + changelog note

**type:** `chore` · **dependsOn:** [CAP-008, CAP-009] · **pipeline:** worker → test-runner → reviewer

**Description:** Полный прогон релевантных pytest/vitest; запись в
`ai_docs/changelog/CHANGELOG.md` (коротко). Documenter сделает report в конце orch.

**Acceptance:**
- [ ] Нет регрессий delivery_schedule_check / archive move / schedule PUT
- [ ] Changelog строка есть

**Verify:**
```bash
pytest tests/test_delivery_schedule_check.py tests/test_capacity_gate.py \
  tests/test_archive_endpoints.py tests/test_delivery_schedule_endpoints.py \
  tests/test_delivery_schedule_service.py -q
cd frontend && npm run test -- --run && npm run typecheck
```

**Files:** `ai_docs/changelog/CHANGELOG.md`

**Scope:** S

---

## Suggested agent split

| Wave | Tasks | Notes |
|------|-------|-------|
| 1 | CAP-001 | сначала core |
| 2 | CAP-002 | после 001 |
| 3 | CAP-003 ∥ CAP-004 ∥ CAP-005 | 004∥005 после 002; 003 после 002 |
| 4 | CAP-006 → CAP-007 | FE foundation |
| 5 | CAP-008 ∥ CAP-009 | оба entry points |
| 6 | CAP-010 | gate |

---

## Risks

| Risk | Mitigation |
|------|------------|
| Расхождение чипов партий и панели | Один `check_batches`, один `start_date=tomorrow` |
| «N дней» = календарные, не рабочие | D6: не менять парсер; target как сейчас |
| Рефактор MonthCalendarGrid | Prefer отдельный mini; Ask first на крупный рефактор |
| Ложное «свободно» (нет СГП) | Accepted MVP; document in UI hint if needed |

---

## Out of scope

- Override / audit table  
- СГП в формуле  
- Резерв дорожек  
- Сайдбар на весь архив  
- Mobile  
- Авто-подбор даты  

---

## Execute

```
/orchestrate execute orch-2026-08-25-16-50-zavod-emkost-gate
```

Начать с **CAP-001**. Не коммитить без просьбы пользователя.
