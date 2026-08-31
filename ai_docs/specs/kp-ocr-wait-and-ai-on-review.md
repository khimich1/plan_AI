# Spec: Wait-баннер OCR + AI на экране сверки КП

**Статус**: SPECIFY ✅ · PLAN ✅ · IMPLEMENT Phase A.3 ✅ · Phase A.3.1 ✅ · **REVIEW 2026-08-31 → Approve** (R10–R11 closed)  
**Дата**: 2026-08-31  
**One-pager**: [ai_docs/ideas/kp-ocr-wait-and-ai-on-review.md](../ideas/kp-ocr-wait-and-ai-on-review.md)  
**План**: [ai_docs/develop/plans/2026-08-31-kp-ocr-wait-and-ai-on-review.md](../develop/plans/2026-08-31-kp-ocr-wait-and-ai-on-review.md)  
**Parent (multi-page)**: [kp-multi-page-screenshots.md](./kp-multi-page-screenshots.md) · план [2026-08-31-kp-multi-page-screenshots.md](../develop/plans/2026-08-31-kp-multi-page-screenshots.md)  
**Follow-on (Apply sync + red/yellow)**: [kp-review-apply-sync-and-highlights.md](./kp-review-apply-sync-and-highlights.md) · план [2026-08-31-kp-review-apply-sync-and-highlights.md](../develop/plans/2026-08-31-kp-review-apply-sync-and-highlights.md)

## Objective

**Проблема.** После «Распознать» менеджер может 1–2 минуты смотреть на «тишину», пока первая страница не станет `ready`. Типичные OCR-ошибки (суффикс нагрузки п↔н и похожие) хочется править **на том же экране сверки**, где уже есть editable list — вместе с AI-инструкцией, а не прятать AI в «Дополнительно» / следующем append.

**Цель.** (1) Явный wait UX до первой `ready`. (2) AI-блок на batch-review для всех типов изделий. (3) Без авто-правил в OCR pipeline и без redesign progressive/lightbox.

**Пользователь:** менеджер на шаге ввода изделий после старта multi-page (и single-page) OCR.

**Успех:** пока нет `ready` — виден баннер «Идёт распознавание…»; после первой `ready` — баннер исчез, сверка как сейчас; на review доступны правка текста + AI для plates/piles/marches/steps/bridge/fbs.

---

## ASSUMPTIONS (locked 2026-08-31)

1. **Wait только до первой `ready`.** Условие: `hasStarted &&` ни одна страница ещё не в статусе `ready` (и не `confirmed`). После появления первой `ready` баннер скрыт; progressive review без изменений.
2. **Копирайт wait:** «Идёт распознавание, подождите 1–2 минуты» + простой spinner. Без прогресс-бара по страницам в этом MVP (k/n уже есть в шапке сверки после ready).
3. **п/н и аналоги — только через UI.** Не добавлять post-OCR pipeline rule `н→п` / suffix rewrite. Исправление = editable text на review + AI instruction на том же экране.
4. **AI UI на batch-review** для **всех** product types: plates, piles, marches, steps, bridge piles, fbs (`*InputStep` + существующий apply-AI wiring).
5. **Текст на review остаётся редактируемым** (уже есть per-page / list editor) — не убирать, не заменять AI.
6. **Reuse существующего AI API** (`applyAi` / instruction) — новый endpoint не обязателен для MVP.
7. **Не трогаем:** lightbox (A.2), progressive review model, Phase B server job, OCR preprocess/verify/upscale.
8. **Коммиты агент не делает**, пока явно не попросите.

→ Assumptions approved / locked with ideation 2026-08-31.

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| **D-wait** | Silence until first OCR | Banner + simple spinner while `hasStarted && no page ready yet`; hide when first page → `ready` |
| **D-ai-on-review** | Where AI lives | Move AI instruction block onto **batch-review page** for **all** product types (not only «Дополнительно» / next-step append card) |
| **D-no-auto-suffix** | п↔н and similar | **Do not** add automatic post-OCR pipeline rule (e.g. н→п). Fix in review window via editable text + AI |

---

## User Stories

- Как **менеджер**, после «Распознать» я вижу понятное ожидание («1–2 минуты»), а не пустой экран, пока первая страница не готова.
- Как **менеджер**, когда страница `ready`, я правлю список руками на сверке.
- Как **менеджер**, на том же экране сверки я могу ввести AI-инструкцию (например, «замени н на п в суффиксе нагрузки») и применить её — для любого типа изделий.

---

## Success Criteria

| # | Критерий | Метод |
|---|----------|--------|
| W1 | `hasStarted &&` нет `ready` → виден баннер с текстом про ожидание 1–2 минуты + spinner | RTL |
| W2 | Первая страница → `ready` → баннер скрыт; progressive review / PageReviewNav без регресса | RTL |
| W3 | До `hasStarted` баннер не показывается | RTL |
| A1 | На batch-review для **каждого** из шести `*InputStep` доступен AI instruction UI (textarea + apply) | RTL / smoke per type или shared review chrome test |
| A2 | Текст списка на review по-прежнему редактируем | regress existing review tests |
| A3 | Нет нового auto suffix rewrite в OCR pipeline / Python OCR path | code review / нет diff в pipeline rules |
| A4 | `npm run test -- src/features/commercial-offer` + typecheck green | CI / локально |

---

## Boundaries

**Always**
- Hide wait chrome as soon as any page is `ready`
- Keep text editable on review
- AI optional (not mandatory per OCR)

**Ask first**
- Удаление AI из «Дополнительно» vs дубль на review + append
- Смена семантики applyAi (только active page vs весь draft)

**Never**
- Auto н→п (или similar) в OCR / post-OCR pipeline
- Phase B server job в этом скоупе
- Mandatory AI on every recognize
- Redesign lightbox or progressive review model

---

## Out of scope (Not Doing)

- Automatic suffix rewrite in OCR pipeline  
- Phase B server job  
- Mandatory AI on every OCR  
- Redesign lightbox / progressive review  
- New multi-image / batch AI endpoint (unless IMPLEMENT proves existing API insufficient — then ask first)

---

## Phased delivery

### Phase A.3 — Wait + AI on review (эта спека)

1. Wait banner + spinner + tests (W1–W3)  
2. Move / surface AI UI on batch-review for all `*InputStep` + tests (A1–A2)  
3. Confirm A3–A4 (no pipeline suffix rule; suite green)

Depends on multi-page Phase A / A.1 / A.2 (done).

---

## Open Questions

- ~~Дублировать ли AI в «Дополнительно»~~ → kept (append) + review.  
- ~~ApplyAi scope~~ → reuse draft-level; R10 меняет post-apply reset, не payload.

---

## Code review (2026-08-31) — Approve (Phase A.3 / A.3.1)

Источник: [Review multi-page quality](../develop/plans/2026-08-31-kp-ocr-wait-and-ai-on-review.md) · companion batch helpers R12–R13 в [kp-multi-page-screenshots.md](./kp-multi-page-screenshots.md).

| ID | Severity | Ось | Проблема | Исправление | Status |
|----|----------|-----|----------|-------------|--------|
| **R10** | Critical | Correctness | `handleApplyAi` → `resetSource()` / `multiPage.reset()` на batch-review стирает multi-сессию (pages, queue, per-page edits) | Если `multiPage.hasStarted`: **не** вызывать `resetSource()`. Синхронизировать результат AI в active page + store; сессия остаётся | ✅ S20 |
| **R11** | Important | Correctness / UX | `AiInstructionBlock disabled={isRecognizing}` — AI недоступен, пока хвост OCR `pending`/`running`, хотя page1 уже `ready` (ломает progressive + правку н→п) | Не гейтить AI полным queue-busy; достаточно `isAiProcessing` (+ draft). Тест: ready + pending → Apply enabled | ✅ S21 |

### Success criteria

| # | Критерий | Метод | Status |
|---|----------|--------|--------|
| S20 | Multi 2+ ready pages → Apply AI → `pages.length` / `hasStarted` intact | unit / RTL | ✅ |
| S21 | Page1 `ready` + page2 `pending` → AI controls enabled | RTL | ✅ |

### Phase A.3.1 — remediation (до merge)

1. R10 — multi-aware `handleApplyAi` (no reset) — ✅  
2. R11 — AI enabled during progressive tail OCR — ✅  
3. `npm run test -- src/features/commercial-offer` + typecheck — ✅  

### Nit / FYI (не блокируют)

| ID | Severity | Note |
|----|----------|------|
| N6 | Nit | AI-on-review RTL только plates + marches |
| N7 | FYI | Apply-AI draft-level (API не видит несохранённый editor text) — зафиксировано как reuse current |

### Follow-on (locked 2026-08-31)

Visible list sync after Apply + red/yellow review highlights via existing lint (no н→п heuristic, no pipeline auto suffix): see [kp-review-apply-sync-and-highlights.md](./kp-review-apply-sync-and-highlights.md).
