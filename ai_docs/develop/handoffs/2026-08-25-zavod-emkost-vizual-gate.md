# Handoff: Визуал ёмкости завода + жёсткий гейт → оркестратор

> **Дата:** 2026-08-25  
> **Ветка:** текущая рабочая  
> **Статус:** Idea ✅ · Spec ✅ · Plan ✅ · **Implementation 0/10** — готов к `/orchestrate`  
> **Цель файла:** открыть **новое окно** и сразу запустить оркестратор без потери ideation/SDD.  
> **Не коммитить** без явной просьбы пользователя.

---

## Как стартовать новую сессию (скопируй в первый промпт)

```
/orchestrate execute orch-2026-08-25-16-50-zavod-emkost-gate

Контекст: handoff ai_docs/develop/handoffs/2026-08-25-zavod-emkost-vizual-gate.md
План уже готов — не перепланировать. TDD на каждую задачу. Не коммитить без просьбы.
Начать с CAP-001.
```

### Чеклист агента в новом окне

1. Прочитать **этот** handoff целиком.
2. Прочитать `.cursor/skills/plan-web-context/SKILL.md`.
3. Прочитать `.cursor/skills/orchestration/SKILL.md` (координатор сам код **не** пишет).
4. Загрузить workspace:
   - `.cursor/workspace/active/orch-2026-08-25-16-50-zavod-emkost-gate/progress.json`
   - `tasks.json`, `links.json`
5. Источник задач: `ai_docs/develop/plans/2026-08-25-zavod-emkost-vizual-gate.md` (не выдумывать scope).
6. Спека: `ai_docs/specs/zavod-emkost-vizual-gate.md` — assumptions **locked**.
7. Запустить task loop с **CAP-001**, далее по DAG (см. plan «Suggested agent split»).

**Режим:** `/orchestrate`, не «просто multitask». На каждую CAP-*: Worker → Test-Writer → Test-Runner → Reviewer (как в `pipeline` задачи).

---

## Артефакты (source of truth)

| Артефакт | Путь |
|----------|------|
| Idea | [`ai_docs/ideas/zavod-emkost-vizual-gate.md`](../../ideas/zavod-emkost-vizual-gate.md) |
| Spec | [`ai_docs/specs/zavod-emkost-vizual-gate.md`](../../specs/zavod-emkost-vizual-gate.md) |
| Plan (10 tasks) | [`ai_docs/develop/plans/2026-08-25-zavod-emkost-vizual-gate.md`](../plans/2026-08-25-zavod-emkost-vizual-gate.md) |
| Orchestration ID | `orch-2026-08-25-16-50-zavod-emkost-gate` |
| Workspace | `.cursor/workspace/active/orch-2026-08-25-16-50-zavod-emkost-gate/` |
| Родительский домен | [`docs/specs/delivery-schedule.md`](../../../docs/specs/delivery-schedule.md), `core/delivery_schedule_check.py` |

### Состояние оркестрации (на момент handoff)

```
status: ready
phase: PLAN
tasksTotal: 10
tasksCompleted: 0
currentTask: null
```

`links.json` → plan / spec / idea прописаны. Report = null (создаст documenter в конце).

---

## Что уже решено (не переспрашивать)

| ID | Решение |
|----|---------|
| Кто | Менеджер (admin/manager), **только ПК** |
| Где | (1) архив → «В производство»; (2) этап **графика поставок** |
| Данные | Только **план** (`days_info`); СГП out |
| Старт окна | **Завтра** |
| UI | Мини-календарь обязателен; месяцы = текущий … месяц конца заказа |
| Шапка | нужно · свободно · Δ |
| Гейт | **Red → нельзя сохранить**; yellow = можно с warning |
| Обход | **Нет** checkbox / второй даты / audit table. Срок — **та же одна строка** (дни). Увеличил N → red снялся → можно |
| Парсер срока | **Не менять** `execution_terms` (календарные дни как сейчас) |
| Алгоритм | Тот же `check_batches` / пороги партий |
| Бронь | Не резервируем дорожки |
| Multi-batch | Блок по **худшей** (любой red) |

---

## Фазы плана (10 задач)

| Phase | Tasks | Суть |
|-------|-------|------|
| 0 Core | CAP-001, CAP-002 | `start_date`; snapshot math |
| 1 API/enforce | CAP-003, CAP-004, CAP-005 | GET snapshot; gate move; gate PUT schedule |
| 2 FE widget | CAP-006, CAP-007 | feature folder; Panel + MiniCalendar |
| 3 Wire | CAP-008, CAP-009 | MoveToProduction + DeliverySchedule |
| 4 Gate | CAP-010 | регрессия + changelog |

Каждая задача в плане: **Acceptance + Verify**. Оркестратор: TDD, потом код.

### DAG (кратко)

```
CAP-001 → CAP-002 → CAP-003
                 ├→ CAP-004 ──────────────┐
                 └→ CAP-005 ──────────────┤
CAP-003 → CAP-006 → CAP-007 → CAP-008 ────┤
                              CAP-009 ────┤
                                          └→ CAP-010
```

Параллель после CAP-002: **CAP-003 ∥ CAP-004 ∥ CAP-005** (004/005 не ждут 003).  
После CAP-007: **CAP-008 ∥ CAP-009**.

---

## Ключевые файлы (ориентир)

**Backend:** `core/delivery_schedule_check.py`, `core/production_capacity.py`, `core/work_calendar.py`,  
`app/services/archive_service.py`, `app/services/delivery_schedule_service.py`,  
новый `app/services/capacity_gate_service.py`, `app/api/v1/endpoints/archive.py`

**Frontend:** новый `frontend/src/features/factory-capacity/`,  
`MoveToProductionDialog.tsx`, `DeliveryScheduleDialog.tsx` / `DeliveryScheduleEditor.tsx`,  
паттерн цветов — `MonthCalendarGrid.tsx` (не обязательно рефакторить)

**Tests:** `tests/test_delivery_schedule_check.py`, новый `tests/test_capacity_gate.py`,  
archive + delivery-schedule endpoint/service tests, vitest feature

---

## Антипаттерны (стоп)

- Не добавлять «Клиент согласовал» / override / новую таблицу аудита  
- Не учитывать СГП в «свободно»  
- Не менять парсер «N дней» на рабочие дни «заодно»  
- Не делать мягкий warning на red без блока  
- Не писать код координатором оркестрации  
- Не коммитить без просьбы пользователя  

---

## Definition of Done (орх)

- [ ] Все CAP-001…010 `completed` в `tasks.json`
- [ ] Red блокирует move и PUT schedule (API + UI)
- [ ] Виджет в обоих entry points; календарь до месяца target
- [ ] Светофор партий и панель не противоречат
- [ ] Report от documenter + строка в CHANGELOG
- [ ] Workspace → `completed/` (или по политике orch)

---

**Первое действие в новом окне:**  
`/orchestrate execute orch-2026-08-25-16-50-zavod-emkost-gate`
