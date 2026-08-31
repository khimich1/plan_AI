# Handoff: Ёмкость завода — left drawer → оркестратор

> **Дата:** 2026-08-26  
> **Ветка:** текущая рабочая  
> **Статус:** Idea ✅ · Spec ✅ · Plan ✅ · **Agent verify/docs ✅** · Human QA live pending  
> **Цель файла:** контекст для ручной приёмки / коммита.  
> **Не коммитить** без явной просьбы пользователя.

---

## Как стартовать новую сессию (скопируй в первый промпт)

```
/orchestrate execute orch-2026-08-26-12-19-zavod-emkost-left-drawer

Контекст: handoff ai_docs/develop/handoffs/2026-08-26-zavod-emkost-left-drawer.md
План уже готов — не перепланировать. Не greenfield: код left drawer уже в WT.
Начать с LDR-001 (stale DeliverySchedule test). Не коммитить без просьбы.
Не трогать parent CAP algorithm/backend.
```

### Чеклист агента в новом окне

1. Прочитать **этот** handoff целиком.
2. Прочитать `.cursor/skills/plan-web-context/SKILL.md`.
3. Прочитать `.cursor/skills/orchestration/SKILL.md` (координатор сам код **не** пишет).
4. Загрузить workspace:
   - `.cursor/workspace/active/orch-2026-08-26-12-19-zavod-emkost-left-drawer/progress.json`
   - `tasks.json`, `links.json`
5. Источник задач: `ai_docs/develop/plans/2026-08-26-zavod-emkost-left-drawer.md`.
6. Спека: `ai_docs/specs/zavod-emkost-left-drawer.md` — assumptions **locked**.
7. Parent gate (`zavod-emkost-vizual-gate`) — **только read**; algorithm/API out of scope.
8. Task loop: **LDR-001** → LDR-002 → LDR-003 (human QA) → LDR-004.

**Режим:** `/orchestrate`, не «просто multitask».

---

## Артефакты (source of truth)

| Артефакт | Путь |
|----------|------|
| Idea | [`ai_docs/ideas/zavod-emkost-left-drawer.md`](../../ideas/zavod-emkost-left-drawer.md) |
| Spec | [`ai_docs/specs/zavod-emkost-left-drawer.md`](../../specs/zavod-emkost-left-drawer.md) |
| Plan (4 tasks) | [`ai_docs/develop/plans/2026-08-26-zavod-emkost-left-drawer.md`](../plans/2026-08-26-zavod-emkost-left-drawer.md) |
| Orchestration ID | `orch-2026-08-26-12-19-zavod-emkost-left-drawer` |
| Workspace | `.cursor/workspace/active/orch-2026-08-26-12-19-zavod-emkost-left-drawer/` |
| Parent CAP orch | `orch-2026-08-25-16-50-zavod-emkost-gate` (**completed**) |

### Состояние оркестрации (после LDR verify)

```
status: completed
phase: DONE
tasksTotal: 4
tasksCompleted: 4
currentTask: null
note: Human QA live sign-off still open (checklist in plan)
```

---

## Inventory snapshot (post-verify)

**Done:**
- `Drawer` `side="left"`, Esc capture, z-index 1100
- Both dialogs: button «Ёмкость», left drawer + panel, red hint in modal
- `MoveToProductionDialog.test.tsx` — click path + Esc closes drawer first
- `DeliveryScheduleDialog.test.tsx` — aligned (panel only after «Ёмкость»)
- Automated: 12/12 vitest + typecheck
- CHANGELOG links gate + left-drawer specs

**Remaining for human:**
- Live QA checklist in plan LDR-003 (`./run+logs.sh`, both entry points, ✕/Esc/backdrop)
- Spec AC «Human QA» stays unchecked until you sign off
- Commit only when you ask

---

## DAG

```
LDR-001 → LDR-002 → LDR-003 → LDR-004
```

---

## Не делать

- Переоткрывать product decisions (badge, auto-open, drawer-to-modal-edge, mobile)
- Менять `check_batches` / capacity-snapshot / backend gate
- Коммитить без явной просьбы пользователя
- Отмечать Human QA AC без живой приёмки