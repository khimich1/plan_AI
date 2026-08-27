# Implementation Plan: Ёмкость завода — left drawer (UX delta)

> **Spec:** [`ai_docs/specs/zavod-emkost-left-drawer.md`](../../specs/zavod-emkost-left-drawer.md)  
> **Idea:** [`ai_docs/ideas/zavod-emkost-left-drawer.md`](../../ideas/zavod-emkost-left-drawer.md)  
> **Parent gate (algorithm/backend — do NOT reopen):** [`ai_docs/specs/zavod-emkost-vizual-gate.md`](../../specs/zavod-emkost-vizual-gate.md) · plan [`2026-08-25-zavod-emkost-vizual-gate.md`](2026-08-25-zavod-emkost-vizual-gate.md) · orch `orch-2026-08-25-16-50-zavod-emkost-gate` (**completed**)  
> **Handoff:** [`ai_docs/develop/handoffs/2026-08-26-zavod-emkost-left-drawer.md`](../handoffs/2026-08-26-zavod-emkost-left-drawer.md)  
> **Orchestration:** `orch-2026-08-26-12-19-zavod-emkost-left-drawer`  
> **Дата:** 2026-08-26  
> **Статус:** VERIFY ✅ automated · Human QA checklist ready (live sign-off pending)  
> **Не коммитить** без явной просьбы

---

## Overview

UX-дельта поверх уже завершённого CAP-гейта: календарь `FactoryCapacityPanel` убран
из grid модалки и открывается **left drawer** по кнопке **«Ёмкость»** в
`MoveToProductionDialog` и `DeliveryScheduleDialog`. Алгоритм / `capacity-snapshot` /
backend 4xx **не трогаем**.

**Честный статус на 2026-08-26 (inventory):** продуктовый wiring в working tree
уже соответствует спеке. План — **не greenfield**, а residual fix + automated verify
+ human QA + docs. Не переоткрывать locked decisions (кнопка, left viewport, hint
в модалке, no auto-open, Esc/backdrop/✕, оба entry points).

---

## Inventory: done vs remaining

### Already satisfied (code + MoveToProduction Vitest)

| Spec AC | Evidence |
|---------|----------|
| Нет inline `FactoryCapacityPanel` в grid формы | Оба диалога: панель только внутри `<Drawer>` |
| Кнопка «Ёмкость» → `Drawer side="left"` + panel | `MoveToProductionDialog.tsx`, `DeliveryScheduleDialog.tsx` |
| Red: hint в модалке при закрытом drawer; submit/save disabled | Alert + `capacityBlocked` на submit/save |
| Закрытие ✕ / Esc (drawer first) / backdrop | `Drawer`: backdrop `onClick={onClose}`, ✕, Esc capture + `stopImmediatePropagation`; dialogs: `handleModalClose` / `handleMainClose` закрывают drawer раньше `onClose` |
| Left viewport edge; z-index > modal | `.app-drawer` `z-index: 1100`; `Modal` inline `zIndex: 1000`; `.app-drawer--left` |
| Parent gate / `isCapacityRed` untouched in this delta | reuse hooks + `isCapacityRed`; no API/algorithm edits in delta files |
| Vitest MoveToProduction capacity-drawer | `MoveToProductionDialog.test.tsx` — hint visible, panel only after click |

### Remaining / residual

| Gap | Severity | Task | Status |
|-----|----------|------|--------|
| DeliverySchedule stale test | High | LDR-001 | ✅ aligned |
| Esc/backdrop drawer-first auto | Low | LDR-002 | ✅ Esc smoke added |
| Human QA на живых данных | Required | LDR-003 | Checklist ready; live pending |
| Spec/changelog links | Docs | LDR-004 | ✅ done; Human QA AC unchecked |

**Out of this plan:** CAP algorithm, backend gate, new APIs, badge on button, auto-open, mobile.

---

## Architecture Decisions (locked — do not reopen)

| ID | Decision |
|----|----------|
| D1 | Кнопка ровно «Ёмкость», без бейджа статуса/Δ |
| D2 | Один паттерн в MoveToProduction и DeliverySchedule |
| D3 | Red: короткий Alert в модалке; drawer **не** автооткрывается |
| D4 | Drawer к **левому краю viewport**, не к краю модалки |
| D5 | Закрытие: ✕, Esc, клик по backdrop |
| D6 | Esc при открытом drawer закрывает **сначала** drawer (capture + stopImmediatePropagation) |
| D7 | z-index drawer ≥ 1100, modal 1000 |
| D8 | Inline-grid с календарём убран; модалка одноколоночная по ёмкости |
| D9 | Backend / check_batches / red→4xx — **не трогаем** |
| D10 | ПК-only |

---

## Task List

### Phase 1: Residual fix

### LDR-001 — Align DeliveryScheduleDialog tests to left-drawer UX

**type:** `feat-fe` · **dependsOn:** [] · **Priority:** Critical · **Scope:** S  
**pipeline:** explore → worker → test-runner → reviewer  
**securitySensitive:** false · **needsExplore:** true · **parallelSafe:** false

**Description:** Обновить `DeliveryScheduleDialog.test.tsx` по образцу
`MoveToProductionDialog.test.tsx`: при red — hint / disabled Save **без** panel;
после клика «Ёмкость» — `factory-capacity-panel` (+ mini-calendar при наличии).
Не менять прод-код, если он уже соответствует спеке (только если тест вскроет баг).

**Acceptance criteria:**
- [x] Test не ожидает panel до клика «Ёмкость»
- [x] Red: Save disabled; hint/capacity block message виден в модалке
- [x] После «Ёмкость»: panel в документе

**Verification:**
```bash
cd frontend && npm run test -- --run \
  src/features/delivery-schedule/components/DeliveryScheduleDialog.test.tsx
```

**Dependencies:** None  
**Files likely touched:**
- `frontend/src/features/delivery-schedule/components/DeliveryScheduleDialog.test.tsx`
- (only if bug) `DeliveryScheduleDialog.tsx`

---

### Checkpoint: After LDR-001
- [x] Stale DeliverySchedule assertion gone; test green

---

### Phase 2: Automated verify + light polish

### LDR-002 — Automated FE verify (+ optional Esc smoke)

**type:** `chore` · **dependsOn:** [LDR-001] · **Priority:** High · **Scope:** S  
**pipeline:** worker → test-runner → reviewer  
**securitySensitive:** false · **needsExplore:** false · **parallelSafe:** false

**Description:** Прогнать целевые Vitest + typecheck. Опционально (если дёшево):
добавить smoke на Esc/backdrop closing drawer without closing modal — только если
нет регрессий и не раздувает scope. Зафиксировать в plan/handoff результат прогона.
Не трогать backend pytest, кроме желаемого smoke (не обязателен).

**Acceptance criteria:**
- [x] Targeted vitest green (MoveToProduction + DeliverySchedule + factory-capacity) — 12/12
- [x] `npm run typecheck` green
- [x] Inventory re-check: Drawer `side`, z-index, both dialogs still match D1–D8
- [x] (Optional) Esc smoke: `Esc closes capacity drawer without closing modal` in MoveToProductionDialog.test.tsx

**Verification:**
```bash
cd frontend && npm run test -- --run \
  src/features/commercial-archive/components/MoveToProductionDialog.test.tsx \
  src/features/delivery-schedule/components/DeliveryScheduleDialog.test.tsx \
  src/features/factory-capacity
cd frontend && npm run typecheck
```

**Dependencies:** LDR-001  
**Files likely touched:**
- possibly `MoveToProductionDialog.test.tsx` / Drawer-related test (optional)
- plan progress notes only otherwise

---

### Checkpoint: After LDR-002
- [x] Automated suite for delta green; no accidental algorithm/API edits

---

### Phase 3: Human QA + docs

### LDR-003 — Human QA both entry points

**type:** `chore` · **dependsOn:** [LDR-002] · **Priority:** High · **Scope:** S  
**pipeline:** worker → reviewer  
*(agent prepares checklist / may assist with `./run+logs.sh` smoke; human signs off)*  
**securitySensitive:** false · **needsExplore:** false · **parallelSafe:** false

**Description:** Ручная приёмка на живых данных (или локальном стенде): оба диалога.
Agent готовит чеклист ниже; **sign-off — human** (агент browser QA не засчитывает).

### Human QA checklist (run on `./run+logs.sh`)

Стенд: `./run+logs.sh`. Нужен КП с партиями / сроком, где ёмкость может быть red (или подкрутить короткий срок).

**A. «В производство» (архив КП → Перевести в производство)**
- [ ] Модалка узкая: нет inline-календаря в форме
- [ ] Red: короткий hint в модалке; submit disabled; drawer **не** автооткрывается
- [ ] Кнопка «Ёмкость» → left drawer от края экрана поверх модалки + `FactoryCapacityPanel`
- [ ] Закрытие drawer: ✕ · Esc (модалка остаётся) · клик по backdrop
- [ ] Повторный Esc после закрытого drawer закрывает модалку

**B. График поставок (DeliveryScheduleDialog)**
- [ ] Тот же паттерн: hint в модалке при red; Save disabled; нет auto-open
- [ ] «Ёмкость» → left drawer + panel
- [ ] ✕ / Esc / backdrop закрывают drawer; Esc не роняет модалку, пока drawer открыт

**C. Visual**
- [ ] Drawer у **левого края viewport**, не у края модалки
- [ ] Drawer визуально выше модалки (z-index)

> Agent status 2026-08-26: checklist documented; **live sign-off pending human**. Spec AC «Human QA» остаётся unchecked.

**Acceptance criteria (manual):**
- [ ] «В производство»: узкая форма; «Ёмкость» → left drawer + panel
- [ ] Red: hint в модалке, submit disabled, drawer не автооткрывается
- [ ] Закрытие drawer: ✕, Esc, backdrop; повторный Esc закрывает модалку
- [ ] График поставок: тот же паттерн; Save disabled на red
- [ ] Drawer от левого края экрана, поверх модалки

**Verification:** Manual on `./run+logs.sh` (or equivalent local stack)

**Dependencies:** LDR-002  
**Files:** checklist in this plan Progress / LDR-003 block; handoff note

---

### LDR-004 — Docs polish (spec AC + changelog links)

**type:** `docs` · **dependsOn:** [LDR-003] · **Priority:** Medium · **Scope:** XS  
**pipeline:** worker → reviewer  
**securitySensitive:** false · **needsExplore:** false · **parallelSafe:** false

**Description:** После automated verify: честный статус AC; ссылка на left-drawer spec
в changelog; Progress обновлён. Human QA checkbox в spec — **только** после live
sign-off (пока unchecked). Не коммитить.

**Acceptance criteria:**
- [x] Spec AC «Human QA» **оставлен unchecked** (live pending) — честно
- [x] Changelog ссылается на gate + left-drawer specs
- [x] Plan Progress section updated

**Verification:** Read-through of linked docs  
**Dependencies:** LDR-003  
**Files likely touched:**
- `ai_docs/specs/zavod-emkost-left-drawer.md`
- `ai_docs/changelog/CHANGELOG.md`
- this plan Progress block

---

### Checkpoint: Complete
- [x] LDR-001…004 agent work done (Human QA live sign-off still open)
- [x] Spec success criteria met for automated path; Human QA AC pending
- [ ] Ready for user-requested commit (not part of this orch)

---

## Suggested agent split / DAG

```
LDR-001 → LDR-002 → LDR-003 → LDR-004
```

| Wave | Tasks | Notes |
|------|-------|-------|
| 1 | LDR-001 | fix stale DeliverySchedule test |
| 2 | LDR-002 | vitest + typecheck (+ optional Esc) |
| 3 | LDR-003 | human QA gate |
| 4 | LDR-004 | docs after QA |

**Parallel:** none (`parallelSafe: false` — shared dialog/test area; tiny serial chain).

**Inject into workers:** `plan-web-context`. Do **not** reopen parent CAP tasks
(`orch-2026-08-25-16-50-zavod-emkost-gate` already completed).

---

## Relation to parent CAP orch

| | Parent CAP | This LDR |
|--|------------|----------|
| Orch | `orch-2026-08-25-16-50-zavod-emkost-gate` | `orch-2026-08-26-12-19-zavod-emkost-left-drawer` |
| Scope | Algorithm, API snapshot, gate, first UI (inline then evolved) | Placement UX only |
| Status | completed (10/10) | ready — verify/polish |
| Plan | `2026-08-25-zavod-emkost-vizual-gate.md` | this file |

Parent acceptance «виджет виден в диалоге» уточнён left-drawer спекой: виджет
в drawer по кнопке; hint при red всегда в модалке.

---

## Risks

| Risk | Impact | Mitigation |
|------|--------|------------|
| Agents reimplement greenfield UI | Med | Inventory above; touch tests/docs first |
| Esc stacking Modal vs Drawer | Low | Already capture + dialog close guards; optional smoke in LDR-002 |
| Human QA blocked without live KP data | Med | Use seed/local KP; checklist still required |
| Accidental backend edits | High | D9; reviewer rejects non-FE files |

---

## Progress (orchestrator updates)

- ✅ LDR-001: Align DeliveryScheduleDialog tests `(feat-fe)` — already matched MoveToProduction pattern; vitest green
- ✅ LDR-002: Automated FE verify `(chore)` — 12/12 vitest + typecheck; Esc smoke added
- ✅ LDR-003: Human QA checklist documented `(chore)` — **live browser sign-off pending human** (see checklist above)
- ✅ LDR-004: Docs polish `(docs)` — changelog links left-drawer spec; Human QA AC unchecked honestly

**Verify commands (2026-08-26):**
```bash
cd frontend && npm run test -- --run \
  src/features/commercial-archive/components/MoveToProductionDialog.test.tsx \
  src/features/delivery-schedule/components/DeliveryScheduleDialog.test.tsx \
  src/features/factory-capacity
# → 5 files, 12 tests passed
cd frontend && npm run typecheck
# → green
```

---

## Execute

```
/orchestrate execute orch-2026-08-26-12-19-zavod-emkost-left-drawer
```

Начать с **LDR-001**. Не коммитить без просьбы пользователя. Checkpoint: user
approves this plan (types + DAG) before execute unless `/orchestrate execute`.
