# Implementation Plan: Архив КП — «Редактировать» на шаг 3 (result)

**Created:** 2026-09-07  
**Спека:** [../../specs/kp-archive-edit-land-on-result.md](../../specs/kp-archive-edit-land-on-result.md)  
**Related:** [2026-09-02-kp-archive-edit-in-constructor.md](./2026-09-02-kp-archive-edit-in-constructor.md) (родительская фича; ARC-005 не ставил `set-step`)  
**Goal:** «Редактировать» в drawer архива открывает конструктор на `result` явным `set-step`, не ослабляя hydrate-merge живого мастера.  
**Total Tasks:** 3  
**Priority:** High  
**Status:** PLAN ✅ · IMPLEMENT ⏳  
**Тип:** hotfix, минимальный diff

SDD: SPECIFY ✅ · PLAN ✅. Код не писать, пока human не подтвердит план.

## Overview

Одна ветка в `openInConstructor`: при `landing === "result"` после успешного `hydrate-draft` вызвать `dispatch({ type: "set-step", step: "result" })`. «(+ Добавить)» не трогаем, кроме теста «нет `set-step: result`». `mergeWizardStepWithServer` не менять. `PlateInputStep` не переписывать.

TDD: сначала красный тест drawer, затем одна правка handler. Не коммитить. Не убивать `./run+logs.sh`. Без новых npm/pip.

## Architecture Decisions

- **Force-step в drawer, не в merge.** Pin input-step при hydrate — защита «Обработать». Archive edit форсирует шаг точечно, как `handleUndoLastBatch` (hydrate → `set-step: result`).
- **До `navigate`.** `useCommercialOfferWizard` повторно гидратит `draftQuery.data`. Второй hydrate безопасен, только если local уже `result`. Поэтому `set-step` в том же handler сразу после hydrate.
- **Не чистить `sourceText`.** На `result` шаг ввода не монтируется; leftover localStorage не виден. Append по-прежнему чистит через `start-append-cycle`.
- **Нет backend.** Resume уже отдаёт `current_step=result`; дыра только FE landing.

## Task List

### Phase 1 — TDD + fix

#### Task LAND-001: RED — drawer assert `set-step: result`

**Type:** `test`  
**Priority:** Critical  
**Complexity:** Simple  
**dependsOn:** []  
**parallelSafe:** false  
**Estimated scope:** S (1 file)

**Description:** В `OfferDetailsDrawer.test.tsx` расширить кейс «Редактировать resumes draft, hydrates to result, without append cycle»: после `hydrate-draft` должен быть `{ type: "set-step", step: "result" }`, индекс `set-step` > индекса hydrate, по-прежнему нет `start-append-cycle`. В кейсе «(+ Добавить)» — `not.toHaveBeenCalledWith({ type: "set-step", step: "result" })`. Не трогать `wizardDraftStore.test.tsx`.

**Acceptance criteria:**
- [ ] Тест «Редактировать» падает на отсутствии `set-step: result` (RED до LAND-002)
- [ ] Тест «(+ Добавить)» фиксирует отсутствие `set-step: result`
- [ ] Resume fail / pending без изменений ожиданий

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx
```
Ожидание до LAND-002: падает только новый assert `set-step`.

**Files likely touched:**
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`

---

#### Task LAND-002: GREEN — `set-step: result` в `openInConstructor`

**Type:** `feat-fe`  
**Priority:** Critical  
**Complexity:** Simple  
**dependsOn:** [LAND-001]  
**parallelSafe:** false  
**Estimated scope:** S (1 file)

**Description:** В `openInConstructor` после `hydrate-draft`: если `landing === "append"` — как сейчас `start-append-cycle`; иначе `dispatch({ type: "set-step", step: "result" })`. Короткий why-комментарий: hydrate не поднимает input-step на result. Не менять `wizardDraftStore`, wizard, `PlateInputStep`.

**Acceptance criteria:**
- [ ] S1/S4 спеки: result-ветка шлёт `set-step`; append — нет
- [ ] Resume error по-прежнему без dispatch (throw до hydrate)
- [ ] Тесты LAND-001 зелёные

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx
```

**Files likely touched:**
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`

---

### Checkpoint: Core fix

- [ ] Drawer-тесты зелёные
- [ ] Diff ≈ 2 файла (test + handler)
- [ ] Human: ок к verify

---

### Phase 2 — Verify (регресс merge)

#### Task LAND-003: Регресс hydrate-merge + typecheck

**Type:** `chore`  
**Priority:** High  
**Complexity:** Simple  
**dependsOn:** [LAND-002]  
**parallelSafe:** false  
**Estimated scope:** S (0 файлов кода, если всё зелёное)

**Description:** Прогнать hydrate-merge store (тест «не поднимает шаг с plates до result» **должен** остаться зелёным) и typecheck. Если красное — чинить только этот hotfix, не «чинить» merge, инвертируя тест. `PlateInputStep` не трогать. Не коммитить.

**Acceptance criteria:**
- [ ] S6/S8: hydrate-merge тест зелёный без правок store
- [ ] typecheck зелёный
- [ ] Статусы спеки/плана: PLAN ✅ · IMPLEMENT 🔄 после зелёного прогона (IMPLEMENT ✅ — после accept human)

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx
cd frontend && npm run test -- src/features/commercial-offer/store/wizardDraftStore.test.tsx
cd frontend && npm run typecheck
```

Ручной smoke (живой `./run+logs.sh`, не убивать):
- [ ] Archive → «в архиве» → **«Редактировать»** → шаг **3. Результат**, скидка, состав
- [ ] **«(+ Добавить)»** → picker / ввод, не result
- [ ] Живой конструктор: на шаге 1 «Обработать» по-прежнему сам ведёт дальше (hydrate не прыгает на result)

**Files likely touched:**
- `ai_docs/specs/kp-archive-edit-land-on-result.md` (статус после прогона)
- `ai_docs/develop/plans/2026-09-07-kp-archive-edit-land-on-result.md` (статус)

---

### Checkpoint: Complete

- [ ] Все acceptance спеки S1–S8 закрыты тестами или smoke
- [ ] Готово к ревью / коммиту по просьбе

## Dependencies Graph

```
LAND-001 (RED test) ──► LAND-002 (GREEN handler) ──► LAND-003 (verify merge + typecheck)
```

Последовательность обязательна. Параллелить нечего.

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Повторный `hydrate-draft` из `draftQuery` сбросит шаг обратно на plates | High | `set-step` до `navigate`; merge pin'ит только local **input**-step, не `result` |
| Ослабить merge «чтобы archive заработал» | High | D-merge: store не трогать; LAND-003 сторожит тест |
| Leftover `sourceText` всплывёт на result | Low | Result не монтирует шаг ввода; чистка out of scope |
| «Готово, далее» всё ещё мёртвая после landing | Med | Тогда follow-up, не раздувать этот hotfix (`PlateInputStep` / `can_proceed_to`) |
| Случайно `set-step: result` на «(+ Добавить)» | Med | Assert в LAND-001; picker + result конфликтуют |

## Out of scope (remind)

- Redesign мастера, скидка в drawer, новые кнопки/страницы
- Изменение `mergeWizardStepWithServer`
- Rewrite `PlateInputStep` / `canNavigateToStep`
- Backend resume

## Open Questions

_Нет._ Если smoke покажет ловушку на шаге 1 после LAND-002 — остановиться и уточнить, не чинить вслепую `PlateInputStep`.

## Implementation Notes for workers

- Скиллы: `plan-web-context`, `test-driven-development`, `incremental-implementation`.
- TDD: LAND-001 красный **до** LAND-002.
- ≤2 файла кода+теста на фикс; store не открывать на запись.
- Комментарий why у `set-step`, не «что делает dispatch».
- Не коммитить без просьбы; не убивать `./run+logs.sh`.
