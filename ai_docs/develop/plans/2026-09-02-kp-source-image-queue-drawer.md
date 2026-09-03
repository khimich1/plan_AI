# Plan: КП — очередь исходных фото на шаге 1 (Drawer)

**Created:** 2026-09-02 14:33  
**Orchestration:** orch-2026-09-02-11-33-kp-source-image-queue  
**Спека:** [../../specs/kp-source-image-queue-drawer.md](../../specs/kp-source-image-queue-drawer.md)  
**Идея:** [../../ideas/kp-source-image-queue-drawer.md](../../ideas/kp-source-image-queue-drawer.md)  
**Goal:** После «Список верен» на шаге 1 менеджер открывает очередь исходных фото текущего захода (1..N) в боковом Drawer; очередь живёт до нового круга / archive save / «Создать новое КП». FE-only.  
**Total Tasks:** 8  
**Priority:** High  
**Status:** PLAN ✅ · IMPLEMENT ✅

## Overview

Сейчас после confirm wizard зовёт `multiPage.reset()` + `clearRecognizedImagePreview()` — blob URL revoke’ятся, исходник на шаге 1 пропадает. План: **promote-then-reset** — перед OCR-reset снять snapshot очереди в отдельный FE store (свежие object URL из `File` / single preview), затем существующий reset чистит OCR UI; clear+revoke очереди только на lifecycle-событиях. CTA «Исходные фото (N)» + `Drawer` (left) на всех `*InputStep`. TDD на каждый срез поведения.

## Architecture Decisions

- **Promote-then-reset (не keep-pages).** Держать `multiPage.pages` после confirm ломает `SourceInputCard` (галерея OCR остаётся). Snapshot в `sourceImageQueue` → затем текущий `multiPage.reset()` / `clearRecognizedImagePreview()` как сейчас. Новые URL из `File` (или transfer без revoke) — старые revoke’ятся reset’ом безопасно.
- **Отдельный FE store, не server / IndexedDB.** `lib/sourceImageQueue.ts` (pure) + `hooks/useSourceImageQueue.ts`. F5 = пусто — ok.
- **Единая очередь для single + multi.** Multi: все страницы захода; single: 1 элемент из `useRecognizedImagePreview` / файла. Text-only → queue empty → CTA нет.
- **Clear + revoke только:** `start-append-cycle` / `handleAddOtherNomenclature`, успешный archive save (`handleSave` с resume navigate), «Создать новое КП» / `handleCreateNewOffer`, и согласованный abandon image→text через `resetSource` / `handleSourceTextChange` (чтобы не оставлять «висячие» фото без CTA-контекста).
- **Не clear на confirm** и не на переход client↔input на том же круге.
- **Drawer:** `frontend/src/shared/ui/Drawer.tsx`, `side="left"`, упрощённый img + имя + open-in-tab + pager `k/N`. Без полного zoom-lightbox сверки.
- **CTA только шаг 1:** проп `sourceQueue` (+ open drawer) во все 6 input steps; не client / Result / archive.
- **Без новых npm/pip.** Не трогать `./run+logs.sh`. Коммиты — по просьбе.

## Tasks Overview

1. **IMG-001** Pure `sourceImageQueue` helpers `(feat-fe)` — dependsOn: []
2. **IMG-002** `useSourceImageQueue` hook `(feat-fe)` — dependsOn: [IMG-001]
3. **IMG-003** Keep queue after confirm (promote, no revoke-on-confirm) `(feat-fe)` — dependsOn: [IMG-002]
4. **IMG-004** Clear queue on append / archive save / create-new `(feat-fe)` — dependsOn: [IMG-003]
5. **IMG-005** `SourceImageQueueDrawer` UI + pager `(ui)` — dependsOn: [IMG-001]
6. **IMG-006** CTA on `PlateInputStep` (RTL) `(ui)` — dependsOn: [IMG-005]
7. **IMG-007** Wire all input steps + wizard `(feat-fe)` — dependsOn: [IMG-004, IMG-006]
8. **IMG-008** Focused verify + docs status `(chore)` — dependsOn: [IMG-007]

## Dependencies Graph

```
IMG-001 ──► IMG-002 ──► IMG-003 ──► IMG-004 ──┐
    │                                          ├──► IMG-007 ──► IMG-008
    └──► IMG-005 ──► IMG-006 ──────────────────┘
```

`IMG-005` ∥ `IMG-002`/`IMG-003`/`IMG-004` (разные файлы; `parallelSafe` между ветками).

---

## Task List

### Phase 1 — Queue foundation (TDD)

#### Task IMG-001: Pure `sourceImageQueue` helpers

**Type:** `feat-fe`  
**Priority:** Critical  
**Complexity:** Simple  
**dependsOn:** []  
**parallelSafe:** true  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** TDD-first pure module: типы `SourceImageQueueItem` (`id`, `url`, `name`), `buildQueueFromPages(pages, createUrl)`, `buildQueueFromPreview({ url, name } | File)`, `clearQueueItems(items, revoke)`, `clampQueueIndex` / `nextIndex` / `prevIndex`. Не React.

**Acceptance criteria:**
- [x] RED→GREEN: build from N `PageSource`-like items → length N, urls from factory
- [x] clear вызывает `revoke` ровно по числу url и возвращает `[]`
- [x] pager helpers: N=1 stay; N=2 wrap or clamp per выбранной политике (зафиксировать clamp 0..N-1 без wrap — проще для UI)

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-offer/lib/sourceImageQueue
```

**Files likely touched:**
- `frontend/src/features/commercial-offer/lib/sourceImageQueue.ts` (new)
- `frontend/src/features/commercial-offer/lib/sourceImageQueue.test.ts` (new)

---

#### Task IMG-002: `useSourceImageQueue` hook lifecycle

**Type:** `feat-fe`  
**Priority:** Critical  
**Complexity:** Simple  
**dependsOn:** [IMG-001]  
**parallelSafe:** false  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** TDD-first hook: `setFromPages` / `setFromSinglePreview` (replace queue; revoke previous), `clear` (revoke all), expose `items` / `length`. Unmount cleanup revokes remaining. Injectable `createObjectURL` / `revokeObjectURL` for tests.

**Acceptance criteria:**
- [x] setFromPages → length N; повторный set revoke’ит старые url
- [x] clear → length 0 + revoke called
- [x] unmount revokes leftover urls

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-offer/hooks/useSourceImageQueue
```

**Files likely touched:**
- `frontend/src/features/commercial-offer/hooks/useSourceImageQueue.ts` (new)
- `frontend/src/features/commercial-offer/hooks/useSourceImageQueue.test.ts` (new)

---

### Phase 2 — Lifecycle wiring (TDD)

#### Task IMG-003: Keep queue after «Список верен» (promote-then-reset)

**Type:** `feat-fe`  
**Priority:** Critical  
**Complexity:** Moderate  
**dependsOn:** [IMG-002]  
**parallelSafe:** false  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** TDD-first: вынести/протестировать promote-хелпер или wizard-level callback. В `CommercialOfferWizard` на успешном confirm (**multi** `allConfirmed` branch ~L446 и **single** path ~L482): **сначала** `sourceQueue.setFromPages(...)` / `setFromSinglePreview(...)`, **потом** `multiPage.reset()` / `clearRecognizedImagePreview()`. Не revoke’ить urls очереди на confirm. Критический риск: сегодняшний early reset — см. Risks.

**Acceptance criteria:**
- [x] После confirm queue.length ≥ 1 при наличии фото; urls валидны (не revoked)
- [x] OCR session сброшен (`multiPage.pages.length === 0`) — SourceInputCard без галереи захода
- [x] Text-only confirm → queue остаётся пустой
- [x] Существующие multiPage/confirm тесты не регрессируют

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-offer/hooks/useSourceImageQueue
cd frontend && npm run test -- src/features/commercial-offer/hooks/useMultiPageRecognize
# плюс точечный тест promote/wizard helper, если вынесен
```

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- optional helper under `lib/` + test
- `frontend/src/features/commercial-offer/hooks/useMultiPageRecognize.test.ts` (только если меняется API reset — предпочтительно **не** менять)

---

#### Task IMG-004: Clear queue on append / archive save / create-new

**Type:** `feat-fe`  
**Priority:** Critical  
**Complexity:** Moderate  
**dependsOn:** [IMG-003]  
**parallelSafe:** false  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** TDD-first: `sourceQueue.clear()` в `handleAddOtherNomenclature` (перед/вместе с `multiPage.reset`), успешный `handleSave` archive path, `handleCreateNewOffer`. Также в `resetSource` / text-abandon (`handleSourceTextChange` когда сбрасывает pages), чтобы не оставлять stale queue. Не clear при обычной навигации step.

**Acceptance criteria:**
- [x] start-append / «Добавить другое наименование» → queue length 0 + revoke
- [x] archive save success → queue cleared
- [x] «Создать новое КП» / wizard reset → queue cleared
- [x] Confirm **не** чистит queue (регресс-тест из IMG-003 зелёный)

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-offer/hooks/useSourceImageQueue
# wizard/store-focused tests if added; иначе helper unit + typecheck later
```

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- tests colocated with promote/clear helper or wizard draft tests

---

### Checkpoint: Lifecycle

- [x] Confirm keeps queue; append/save/create-new clear + revoke
- [x] multiPage OCR reset after confirm still empties upload gallery
- [x] No server / IndexedDB

---

### Phase 3 — Drawer + CTA (TDD)

#### Task IMG-005: `SourceImageQueueDrawer` + pager

**Type:** `ui`  
**Priority:** High  
**Complexity:** Moderate  
**dependsOn:** [IMG-001]  
**parallelSafe:** true (vs IMG-002..004)  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** TDD-first RTL: компонент на `Drawer` (`side="left"`), title «Исходные фото», текущий img, имя файла, «Открыть в новой вкладке», счётчик `k / N`, prev/next с `aria-label`. N=1 — стрелки hidden или disabled. Без новых deps.

**Acceptance criteria:**
- [x] open=true → dialog; Esc/close работает (через Drawer)
- [x] N=2: next/prev меняют картинку и счётчик
- [x] N=1: нет обязательных активных стрелок
- [x] пустой items → null / ничего не рендерит (защита)

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-offer/components/SourceImageQueueDrawer
```

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/SourceImageQueueDrawer.tsx` (new)
- `frontend/src/features/commercial-offer/components/SourceImageQueueDrawer.test.tsx` (new)
- reuse `frontend/src/shared/ui/Drawer.tsx` (no API change unless needed)

---

#### Task IMG-006: CTA «Исходные фото (N)» on PlateInputStep

**Type:** `ui`  
**Priority:** High  
**Complexity:** Moderate  
**dependsOn:** [IMG-005]  
**parallelSafe:** false  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** TDD-first в `PlateInputStep.test.tsx`: пропы `sourceQueue` / `onOpenSourceQueue` (или controlled `sourceDrawerOpen`). CTA secondary `Исходные фото (N)` только при `sourceQueue.length > 0` и после/вне batch-review на шаге 1 (достаточно: всегда когда queue non-empty на input step — queue пуст до confirm). Клик открывает Drawer.

**Acceptance criteria:**
- [x] queue=[] → кнопки нет
- [x] queue.length=2 → текст содержит «Исходные фото (2)»
- [x] click → Drawer open с картинкой
- [x] существующие PlateInputStep тесты зелёные

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-offer/components/steps/PlateInputStep
```

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.test.tsx`

---

#### Task IMG-007: Wire all `*InputStep` + wizard

**Type:** `feat-fe`  
**Priority:** High  
**Complexity:** Moderate  
**dependsOn:** [IMG-004, IMG-006]  
**parallelSafe:** false  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** Пробросить queue + open state / handlers из `CommercialOfferWizard` во все: Plate, Pile, Step, March, BridgePile, Fbs. Одинаковый CTA+Drawer паттерн (скопировать пропы из Plate; минимальный RTL на одном peer или shared smoke). Drawer state в wizard (один instance) или внутри step — предпочтительно **в step** (проще) при передаче `sourceQueue` items.

**Acceptance criteria:**
- [x] Все 6 input steps показывают CTA при non-empty queue
- [x] S1–S6 спеки покрыты тестами/ручным smoke checklist
- [x] Client / Result без CTA

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-offer/components/steps
cd frontend && npm run test -- src/features/commercial-offer/components/SourceImageQueueDrawer
```

**Files likely touched:**
- `CommercialOfferWizard.tsx`
- `PileInputStep.tsx`, `StepInputStep.tsx`, `MarchInputStep.tsx`, `BridgePileInputStep.tsx`, `FbsInputStep.tsx` (+ tests where exist)

---

### Checkpoint: UI

- [x] CTA + left Drawer on step 1; pager works
- [x] Empty queue → no CTA

---

### Phase 4 — Verify + docs

#### Task IMG-008: Focused verify + docs status

**Type:** `chore`  
**Priority:** Medium  
**Complexity:** Simple  
**dependsOn:** [IMG-007]  
**parallelSafe:** false  
**needsExplore:** false  
**securitySensitive:** false  
**pipeline:** ["test-runner", "documenter"]

**Description:** Прогнать команды из спеки; обновить idea/spec → IMPLEMENT ✅ (или оставить IMPLEMENT ⬜ до user accept — **после зелёных тестов** поставить IMPLEMENT ✅). Не коммитить без просьбы.

**Acceptance criteria:**
- [x] S7: focused vitest + typecheck зелёные
- [x] Spec/idea статусы согласованы с фактом

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-offer/hooks/useMultiPageRecognize
cd frontend && npm run test -- src/features/commercial-offer/lib/multiPageSource
cd frontend && npm run test -- src/features/commercial-offer/lib/sourceImageQueue
cd frontend && npm run test -- src/features/commercial-offer/hooks/useSourceImageQueue
cd frontend && npm run test -- src/features/commercial-offer/components/SourceImageQueueDrawer
cd frontend && npm run test -- src/features/commercial-offer/components/steps
cd frontend && npm run typecheck
```

**Files likely touched:**
- `ai_docs/specs/kp-source-image-queue-drawer.md`
- `ai_docs/ideas/kp-source-image-queue-drawer.md`
- this plan Progress section

---

## Progress (updated by orchestrator)

- ✅ IMG-001: Pure sourceImageQueue helpers `(feat-fe)` (Done)
- ✅ IMG-002: useSourceImageQueue hook `(feat-fe)` (Done)
- ✅ IMG-003: Keep queue after confirm `(feat-fe)` (Done)
- ✅ IMG-004: Clear on append/save/create-new `(feat-fe)` (Done)
- ✅ IMG-005: SourceImageQueueDrawer `(ui)` (Done)
- ✅ IMG-006: PlateInputStep CTA `(ui)` (Done)
- ✅ IMG-007: Wire all input steps `(feat-fe)` (Done)
- ✅ IMG-008: Focused verify + docs `(chore)` (Done)

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| **Early `multiPage.reset()` after confirm (~L446) revokes blob URLs** | High — Drawer broken / blank img | Promote **before** reset; queue owns **independent** URLs from `File` |
| **`clearRecognizedImagePreview()` on single confirm (~L482)** | High — single-page path loses image | Promote single preview/File into queue **before** clear |
| Double memory (old previewUrl + new queue URL briefly) | Low | Accept for MVP (≤12 pages); optional transfer API later |
| SourceInputCard still shows pages if keep-pages chosen | High UX | Prefer promote-then-reset; do **not** keep pages as queue without emptying card props |
| Forget clear on `resetSource` / text abandon | Med — stale CTA | Explicit clear in `resetSource` + append/save/create-new |
| Unmount wizard revokes mid-Drawer | Low | Hook unmount clear; close Drawer on clear |
| Peer input steps drift (CTA only on Plate) | Med | IMG-007 checklist all 6; shared prop names |

## Open Questions / Blockers

_None blocking._ Defaults locked:

- Drawer side: **left**
- Pager: clamp index (no wrap) unless implementer prefers wrap — document in Drawer test
- Zoom: optional; MVP = img + open-in-tab

## Out of Scope

- IndexedDB / F5 persist
- Server OCR file storage
- Photos on client / Result / archive resume
- Permanent split-view on composition preview
- New npm/pip packages

## Implementation Notes for workers

- Inject `plan-web-context` + this plan + spec.
- **Every behavior task: write failing vitest first, then implement.**
- Do not kill `./run+logs.sh`; do not commit unless asked.
- Prefer small vertical slices; after IMG-003 manually sanity-check confirm still reaches composition preview.
