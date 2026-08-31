# Implementation Plan: Несколько скриншотов страниц на шаге ввода КП

**Спека**: [ai_docs/specs/kp-multi-page-screenshots.md](../../specs/kp-multi-page-screenshots.md)  
**Идея**: [ai_docs/ideas/kp-multi-page-screenshots.md](../../ideas/kp-multi-page-screenshots.md)  
**Дата**: 2026-08-31  
**Статус**: Phase A MVP · A.1 · A.2 · A.3 · A.4 done · **remediations closed: R10–R13** (2026-08-31)  

## Overview

Менеджер набирает до 12 скринов страниц большого КП, видит превью с ✕, жмёт «Распознать» один раз. OCR идёт **последовательно** существующим single-image API; первая `ready` сразу открывает сверку (фото + её список). Навигация переключает пару фото↔текст. «Список верен» — по странице, с автопереходом на следующую `ready`. Плохое фото: ✕ удалить + добавить хорошее **в конец**, остальные не трогаем. Append UI / AI почти не меняем. Server job — не в этом плане (фаза B спеки).

## Architecture Decisions

- **Состояние страниц живёт рядом с wizard**, не в шести `*InputStep`. Сейчас `CommercialOfferWizard` держит `selectedImage: File | null` + `recognizedImagePreview`; `wizardDraftStore` — `selectedImageName`. Заменяем на `pages: PageSource[]` + `activePageId` (store и/или хук `useMultiPageRecognize`).
- **Конвейер — чистая логика + хук.** Типы/хелперы в `lib/multiPageSource.ts` (soft-cap, canRemove, nextReadyAfterConfirm); оркестрация API в `hooks/useMultiPageRecognize.ts`. UI только диспатчит.
- **API без новых endpoint в MVP.** Первый pending → create/replace (как сейчас); каждый следующий pending → append того же product_type. Facade внутри хука, не в бэкенде.
- **Галерея общая** в `SourceInputCard` / `SourceImageGallery`. Шаги получают active page image URL + text из родителя; ←/→ можно в шаге сверки или в галерее (клик по превью).
- **Confirm:** пер-page mark `confirmed` + существующий confirm-batch path адаптировать: либо confirm только активного сегмента текста, либо после каждой страницы «вливать» правки в draft так же, как сейчас один batch — **важно не сломать** `pendingBatchReview` для одиночного файла. Предпочтение: multi-page держит `pendingBatchReview=true`, пока не все страницы `confirmed`; confirm страницы сохраняет её `batchReviewText` в draft (append-семантика уже на сервере после OCR) и переключает UI.
- **Не трогаем:** OCR pipeline Python, verify/upscale policy, AI-instruction UX, multi-nomenclature loop, параллельный OCR.

## Risk & mitigation

| Risk | Mitigation |
|------|------------|
| `pendingBatchReview` + один `batchReviewText` не тянут N страниц | Per-page `batchReviewText` в `PageSource`; при switch подставляем текст активной; confirm пишет текущую |
| Append после create меняет «весь» draft preview | Как сегодня append; сверка фильтрует через существующий `filterDraftForBatchReview` **по тексту текущей страницы** |
| Object URL leaks | `revoke` в remove/unmount/reset |
| Регресс одиночного файла / текста | Compat: 0–1 page ведёт себя как сейчас; lint path без картинок без изменений |
| Wizard раздут | Вынести конвейер в хук; wizard только wiring |

## Dependency Graph

```
lib/multiPageSource (types, soft-cap, remove/add rules)
    │
    ├── SourceImageGallery
    │       │
    │       └── SourceInputCard (галерея вместо Alert; multi file/paste)
    │
    └── useMultiPageRecognize (sequential OCR create→append…)
            │
            ├── commercialOfferApi (existing)
            │
            └── CommercialOfferWizard + wizardDraftStore
                    │
                    └── *InputStep (active page img+text, nav, progress k/n, confirm)
```

## Task List

### Phase 1: Domain + gallery UI

#### Task 1: `PageSource` lib + unit tests

**Description:** Типы `PageStatus` / `PageSource`, константа `MAX_PAGES=12`, функции: `canRemovePage`, `addFilesToPages` (хвост, soft-cap, clear text coupling на уровне контракта), `removePage`, `pickNextReadyAfterConfirm`, `countRecognizedProgress`. Без React.

**Acceptance:**
- [x] Soft-cap: 12-й ок, 13-й отвергнут с reason
- [x] `canRemove`: pending/error/ready = true; running/confirmed = false
- [x] remove не трогает соседей; add всегда в конец
- [x] nextReady: после confirm id → ближайшая следующая ready, иначе первая ready, иначе null

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/lib/multiPageSource.test.ts`

**Dependencies:** None  
**Files:** `frontend/src/features/commercial-offer/lib/multiPageSource.ts`, `.../multiPageSource.test.ts`  
**Estimated scope:** S

#### Task 2: `SourceImageGallery` + wire into `SourceInputCard`

**Description:** Горизонтальный ряд квадратных превью, active outline, ✕ на removable, клик = select. `SourceInputCard`: убрать Alert имени файла; `input multiple`; paste дописывает; пропсы `pages` / callbacks вместо scalar `selectedImageName` (или адаптер: `pages.length` → hasImage). Сохранить lint-гейт для текста.

**Acceptance:**
- [x] ≥2 превью; ✕ вызывает onRemove только если canRemove
- [x] Alert «Выбран файл» отсутствует
- [x] Photo-only recognize всё ещё enabled при pages>0 и пустом тексте
- [x] Существующие SourceInputCard tests обновлены и green

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/components/SourceInputCard.test.tsx src/features/commercial-offer/components/SourceImageGallery.test.tsx`

**Dependencies:** Task 1  
**Files:** `SourceImageGallery.tsx`, `SourceImageGallery.test.tsx`, `SourceInputCard.tsx`, `SourceInputCard.test.tsx`  
**Estimated scope:** M

### Checkpoint: Gallery

- [x] Галерея + soft-cap + lint regress green
- [x] Визуально: превью с ✕ на тёмном/светлом фоне карточки читаемы

### Phase 2: Recognize pipeline

#### Task 3: `useMultiPageRecognize` (sequential)

**Description:** Хук: `pages`, `activeId`, `start(productType/mode)`, `addFiles`, `remove`, `setActive`, `updatePageText`, `confirmActive`. `start`: по очереди pending → running → API (1-й create/replace, далее append) → ready|error; не ждать весь батч для UI. После удаления + add в хвост — runner подхватывает новые pending. `beforeunload` если running/pending после старта.

**Acceptance:**
- [x] До start API не зовётся
- [x] Mock: page1 resolve → UI может читать ready, пока page2 ещё pending/running
- [x] error на k → k+1 всё равно стартует
- [x] remove ready/error + add file → новый pending в хвосте; остальные status/text неизменны
- [x] Soft-cap соблюдён

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/hooks/useMultiPageRecognize.test.ts`

**Dependencies:** Task 1  
**Files:** `hooks/useMultiPageRecognize.ts`, `hooks/useMultiPageRecognize.test.ts` (api mocked)  
**Estimated scope:** L

#### Task 4: Wire wizard + store

**Description:** Заменить `selectedImage` / `selectedImageName` / одиночный preview на pages из хука (или store actions). `handleRecognize("replace")` на первом заходе → `pipeline.start`. Проброс в шесть `*InputStep`: active recognized URL/name, active batch text, progress `k/n`, nav handlers. Одиночный файл = pages length 1 — поведение как сейчас. Append UI path: по-прежнему один файл (не включать multi в «Дополнительно» в MVP).

**Acceptance:**
- [x] Create draft с 1 картинкой — регресс flow
- [x] 2+ картинки — два API вызова (create затем append)
- [x] `pendingBatchReview` true пока не все confirmed (после start)
- [x] typecheck green; wizard tests обновлены

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/hooks/useCommercialOfferWizard.test.ts src/features/commercial-offer/store/wizardDraftStore.test.tsx` + `npm run typecheck`

**Dependencies:** Task 2, Task 3  
**Files:** `CommercialOfferWizard.tsx`, `wizardDraftStore.tsx`, `useCommercialOfferWizard.ts` (если затрагивается), шесть `*InputStep.tsx` + их tests  
**Estimated scope:** L

### Checkpoint: Progressive OCR

- [x] Ручной/мок: 2 файла → первая ready видна до конца второй
- [x] Progress «Распознано k/n» виден

### Phase 3: Review UX + confirm

#### Task 5: Навигация и confirm по странице

**Description:** В режиме сверки: ←/→ и клик по превью меняют active (img+text вместе). Confirm («Список верен») на active `ready` → `confirmed`, автопереход на next ready (D7). Когда все confirmed — снять multi pending / как сегодняшний выход из batch review. Свободная навигация по ready до confirm (D8). Нельзя confirm `running`/`pending`/`error`.

**Acceptance:**
- [x] Switch page меняет img src и textarea value
- [x] Confirm page1 → focus next ready
- [x] Нельзя далее по wizard, пока есть не-confirmed после start
- [x] PlateInputStep (и smoke одного другого типа) tests покрывают nav/confirm

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/components/steps/PlateInputStep.test.tsx` (и точечно другой step при наличии)

**Dependencies:** Task 4  
**Files:** `*InputStep.tsx`, связанные tests, возможно тонкий `PageReviewNav.tsx`  
**Estimated scope:** M

#### Task 6: Regress + docs touch

**Description:** Прогнать commercial frontend suite; убедиться текстовый lint path жив; обновить статус спеки/плана. Backend pytest `-k commercial` smoke (без новых endpoint — только регресс).

**Acceptance:**
- [x] S1–S11 закрыты или явно отложены с причиной
- [x] `npm run test` + `npm run typecheck` green в frontend
- [x] `pytest tests/ -q -k commercial` green (или зафиксированный subset)

**Note (2026-08-31):** `pytest -k commercial` → **372 passed**, 4 skipped; **3 failed** in `test_commercial_web_flow.py` appear **pre-existing / unrelated** to this MVP (`get_next_kp_number`, `_build_offer_identity`, schema plate context) — no backend changes in Phase A.

**Verification:** команды выше  
**Dependencies:** Task 5  
**Files:** docs status lines only  
**Estimated scope:** S

### Checkpoint: Done (MVP)

- [x] Ручной сценарий из спеки (3 скрина, ✕ плохого, add в хвост, confirm по страницам) — покрыт unit/RTL; ручной smoke желателен
- [x] Phase B (server job) не начата

### Phase A.1 — Review remediation (блокирует merge)

Источник: секция **Code review (2026-08-31)** в спеке (R1–R7, S12–S17).

#### Task 7: R1 empty-after-start + R5 stale remove/add

**Description:** Если после start страниц 0 — `reset()` multi-сессии. `remove`/`addFiles` возвращают актуальный список (или wizard не читает stale `pages` в том же тике).

**Acceptance:**
- [x] S12: remove all after start → hasStarted false / нет вечного multiPendingReview
- [x] S15: store imageName согласован с реальным хвостом после remove/add

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/hooks/useMultiPageRecognize.test.ts`

**Dependencies:** Phase A done  
**Files:** `useMultiPageRecognize.ts`, wizard handlers, tests  
**Estimated scope:** S

#### Task 8: R2–R4 error visibility + status chrome

**Description:** Alert с `errorMessage`; бейдж статуса на thumbnail; hint про ✕ + add в хвост когда есть error.

**Acceptance:**
- [x] S13, S14 green
- [x] Страница error не «молчаливая»

**Verification:** gallery + step/nav RTL  
**Dependencies:** Task 7  
**Files:** `SourceImageGallery.tsx`, `PageReviewNav` / InputStep, tests  
**Estimated scope:** M

#### Task 9: R6 single object URL + R7 getBatchCount

**Description:** Не создавать второй blob URL на ready. Confirm (multi + ideally single) использует store `getBatchCount` / общий helper для fbs/bridge.

**Acceptance:**
- [x] S16, S17 green
- [x] commercial-offer suite + typecheck green

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer && npm run typecheck`  
**Dependencies:** Task 8  
**Files:** hook, wizard confirm, possibly export getBatchCount from store/lib  
**Estimated scope:** S

### Checkpoint: Merge-ready

- [x] R1–R7 закрыты
- [x] S12–S17 green
- [x] Spec verdict → Approve

### Phase A.2 — busy fix + lightbox before OCR

Источник: спека **Phase A.2** (R8/R9, S18/S19); one-pager [kp-multi-page-lightbox-and-busy-fix.md](../../ideas/kp-multi-page-lightbox-and-busy-fix.md).

#### Task 10: R8 — busy only after recognition started

**Description:** `isRecognizingMulti` / hook `isRecognizing` true только когда OCR реально стартовал: `isRunning || (hasStarted && hasPendingOrRunning)`. До start кнопка «Распознать фото», enabled при pages>0; после start «Распознавание...» пока очередь бежит.

**Acceptance:**
- [x] S18: add files without start → `isRecognizing === false`; button not stuck on «Распознавание...»
- [x] После `start` с pending — `isRecognizing === true` пока очередь не idle

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/hooks/useMultiPageRecognize.test.ts` (+ SourceInputCard / wizard wiring if needed)

**Dependencies:** Phase A.1  
**Files:** `useMultiPageRecognize.ts`, `CommercialOfferWizard.tsx`, tests  
**Estimated scope:** S

#### Task 11: R9 — lightbox before OCR

**Description:** До `hasStarted` клик по thumbnail открывает `SourceImageLightbox` (крупное превью). Esc / backdrop / close закрывают; ←/→ листают страницы. ✕ на thumb — remove (stopPropagation). После `hasStarted` — прежний select для review (PageReviewNav не ломать).

**Acceptance:**
- [x] S19: open on click before start; close on Esc; next/prev; after start onSelect still for review
- [x] Remove ✕ still works without opening lightbox

**Verification:** RTL on lightbox + gallery wiring  
**Dependencies:** Task 10 (можно параллельно после R8 green)  
**Files:** `SourceImageLightbox.tsx` (+ test), `SourceImageGallery.tsx` / `SourceInputCard.tsx`, tests  
**Estimated scope:** M

### Checkpoint: Phase A.2 done

- [x] R8/R9 + S18/S19 green
- [x] `npm run test -- src/features/commercial-offer` + typecheck green
- [x] Spec/plan checklists updated

### Phase A.3 — Wait + AI on review

Задачи: [2026-08-31-kp-ocr-wait-and-ai-on-review.md](./2026-08-31-kp-ocr-wait-and-ai-on-review.md). Companion R10–R11 + R12–R13 closed below.

### Phase A.4 — Batch helpers remediation (R12–R13)

Источник: Code review 2026-08-31 в спеке multi-page.

#### Task 12: R12 — getBatches / getCurrentBatchReviewText for fbs/bridge

**Description:** Расширить `getBatches` (и типы) как `getDraftBatchCount`. Per-page review text = last batch only, иначе multi confirm дублирует строки.

**Acceptance:**
- [x] S22: fbs/bridge 2-page → distinct texts; confirm без дублей

**Verification:** unit on batchReview helpers + flow if present  
**Files:** `lib/batchReview.ts`, types, tests  
**Estimated scope:** M

#### Task 13: R13 — confirm merge uses shared batches for all six types

**Description:** Single-path (и multi final) `handleConfirmBatch` merge не падает в `plate_batches` для fbs/bridge.

**Acceptance:**
- [x] S23 green

**Verification:** unit  
**Files:** `CommercialOfferWizard.tsx`, shared helper  
**Estimated scope:** S

### Checkpoint: Merge-ready (all remediations)

- [x] R10–R11 (companion plan)  
- [x] R12–R13  
- [x] Specs → Approve  

## Out of this plan

- Server-side OCR job / Redis  
- Multi-page в UI append / смене номенклатуры  
- Drag-reorder, parallel OCR, multi-image vision  
- In-place replace файла в том же слоте  
- N2 (OCR corrections per-page filter) — optional later  
- Progressive review redesign / permanent split-preview before OCR (A.2 out)  

## Open for implementer

- Точные подписи («Распознано 2/5» vs «Страница 2 из 5»)  
- Нужен ли отдельный `PageReviewNav` или стрелки внутри существующей card изображения  
- Минимальный ли diff store: хранить File только в хуке, в store — meta без File (Files не сериализуются — **предпочтительно** pages с File только в React state хука/wizard)
