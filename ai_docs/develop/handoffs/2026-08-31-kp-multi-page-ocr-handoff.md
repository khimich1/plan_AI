# Handoff: КП multi-page OCR → Phase A.4 (Apply sync + highlights)

> **Дата:** 2026-08-31  
> **Ветка:** текущая рабочая · **весь стек multi-page / wait / AI-on-review — в working tree, не закоммичен**  
> **Статус:** Multi-page Phase A…A.3.1 + A.4 remediations R10–R13 ✅ · **Phase A.4 (Apply→список + red/yellow) — IMPLEMENT ⏳**  
> **Цель файла:** открыть **новое окно в мультитаске** и сразу **implement** A.4 с тестами, без повторного ideation.  
> **Не коммитить** без явной просьбы пользователя.

---

## Как стартовать новую сессию (скопируй в первый промпт)

```
В мультитаске: реализуй Phase A.4 с тестами на каждый шаг. Не коммить.

Контекст: прочитай целиком
ai_docs/develop/handoffs/2026-08-31-kp-multi-page-ocr-handoff.md

Источник правды:
- ai_docs/specs/kp-review-apply-sync-and-highlights.md
- ai_docs/develop/plans/2026-08-31-kp-review-apply-sync-and-highlights.md

Стек Phase A…A.3.1 уже в working tree — не откатывай, не переideate.
Задача: (1) после «Применить инструкцию» список рядом с фото реально меняется + visible error при fail;
(2) на сверке red = parser reject, yellow = soft уже wired; parser-accepted (в т.ч. 8н) НЕ подсвечивать;
(3) без auto н→п в OCR pipeline.

TDD. В конце: npm run test -- src/features/commercial-offer && npm run typecheck.
Обнови чеклисты спеки/плана. Не коммить.
```

### Чеклист агента в новом окне

1. Прочитать **этот** handoff целиком.  
2. `.cursor/skills/plan-web-context/SKILL.md`.  
3. Спека + план A.4 (пути выше).  
4. Не запускать `/idea-refine` заново — решения locked.  
5. Implement Tasks 1→2 из плана A.4 с тестами S1–S7.  
6. Не трогать `test_kp.pdf` / посторонние diff (`tests/test_march_line_parser.py`), если не нужны задаче.  
7. Не коммитить без просьбы.

**Режим:** multitask + worker/implement · **не** ideation · **не** полный orchestration с нуля.

---

## Артефакты

| Артефакт | Путь | Статус |
|----------|------|--------|
| Idea multi-page | [`ai_docs/ideas/kp-multi-page-screenshots.md`](../../ideas/kp-multi-page-screenshots.md) | ✅ |
| Spec multi-page | [`ai_docs/specs/kp-multi-page-screenshots.md`](../../specs/kp-multi-page-screenshots.md) | ✅ A…A.2 · R12–R13 ✅ |
| Plan multi-page | [`ai_docs/develop/plans/2026-08-31-kp-multi-page-screenshots.md`](../plans/2026-08-31-kp-multi-page-screenshots.md) | ✅ |
| Idea busy+lightbox | [`ai_docs/ideas/kp-multi-page-lightbox-and-busy-fix.md`](../../ideas/kp-multi-page-lightbox-and-busy-fix.md) | ✅ R8–R9 |
| Idea wait+AI review | [`ai_docs/ideas/kp-ocr-wait-and-ai-on-review.md`](../../ideas/kp-ocr-wait-and-ai-on-review.md) | ✅ |
| Spec wait+AI | [`ai_docs/specs/kp-ocr-wait-and-ai-on-review.md`](../../specs/kp-ocr-wait-and-ai-on-review.md) | ✅ A.3 · R10–R11 ✅ |
| Plan wait+AI | [`ai_docs/develop/plans/2026-08-31-kp-ocr-wait-and-ai-on-review.md`](../plans/2026-08-31-kp-ocr-wait-and-ai-on-review.md) | ✅ |
| **Idea A.4 (эта работа)** | [`ai_docs/ideas/kp-review-apply-sync-and-highlights.md`](../../ideas/kp-review-apply-sync-and-highlights.md) | **locked** |
| **Spec A.4** | [`ai_docs/specs/kp-review-apply-sync-and-highlights.md`](../../specs/kp-review-apply-sync-and-highlights.md) | **IMPLEMENT ⏳** |
| **Plan A.4** | [`ai_docs/develop/plans/2026-08-31-kp-review-apply-sync-and-highlights.md`](../plans/2026-08-31-kp-review-apply-sync-and-highlights.md) | Tasks 1–2 open |
| Text lint (red pattern) | [`ai_docs/specs/unparsed-line-live-highlight.md`](../../specs/unparsed-line-live-highlight.md) | ✅ уже в main (`cb2241a`) |

---

## Что уже сделано (не переделывать)

### Multi-page OCR (Phase A → A.2)

- Галерея превью + ✕, soft-cap 12, paste/multi file  
- Sequential OCR create→append, progressive: первая `ready` → сверка  
- Confirm per page, плохое фото = ✕ + add в хвост  
- **R8:** busy только после `hasStarted` (`multiPage.isRecognizing`)  
- **R9:** lightbox по клику на thumb **до** OCR  
- Empty-after-start reset (R1), error chrome, single object URL (R6)

### Wait + AI on review (Phase A.3 → A.3.1)

- `OcrWaitBanner` до первой `ready`  
- `AiInstructionBlock` на batch-review для всех 6 типов (+ дубль в «Дополнительно» для append)  
- **R10:** `planApplyAiSessionSync` — при `hasStarted` **не** `resetSource()`; пишется `nextActivePageText`  
- **R11:** AI на review **не** `disabled={isRecognizing}` (хвост OCR не блокирует Apply)  
- **R12–R13:** `getBatches` / confirm для fbs/bridge

### Тесты (последний прогон агента)

- commercial-offer suite ~**208** passed + typecheck green (после R10–R13)

---

## Что чинить сейчас (Phase A.4) — пользовательский баг

На сверке менеджер пишет инструкцию («замени суффикс н на п») → **«Применить инструкцию»** → **список рядом с фото не меняется** (остаётся `8н`). Нужно:

1. **Apply success → UI sync** списка (editor + store / active page), не только hydrate draft.  
2. **Apply fail → visible error**, не тишина.  
3. **Подсветка на сверке:** красное = parser reject (как text lint); жёлтое = soft уже wired; **если парсер принял `8н` — не подсвечивать** (без эвристики н→п).

### Locked decisions (не переспрашивать)

| ID | Решение |
|----|---------|
| D-apply-sync | Success → hydrate **и** запись текста в review editor + store |
| D-apply-error | Visible error on fail |
| D-red | Parser reject / unparsed |
| D-yellow | Только soft already wired — **не** новая н-heuristic |
| D-no-false-yellow | Parser accept (`8н`) → no highlight |
| D-no-auto-suffix | Нет auto н→п в OCR pipeline |

### Подозрение по корню Apply (для implementer)

`handleApplyAi` уже использует `planApplyAiSessionSync` и при multi вызывает `multiPage.updatePageText(...)`, но UI может:

- брать `batchReviewText` из store, а не из page после sync;  
- показывать controlled value, который не перерисовывается;  
- API возвращает текст, который `getCurrentBatchReviewText` не отражает как «новый» last batch;  
- apply мутирует draft, но review editor bound к stale `state.batchReviewText` / `reviewBatchText` без dispatch.

Смотреть: `CommercialOfferWizard.tsx` → `handleApplyAi`, `applyAiSession.ts`, как `*InputStep` получает `batchReviewText` на review.

Highlights: переиспользовать `useSourceTextLint` / highlights map как в `SourceInputCard` / `unparsed-line-live-highlight` — на **batch-review** `PlateListEditor`.

---

## Ключевые файлы кода (уже в WT)

```
frontend/src/features/commercial-offer/
  hooks/useMultiPageRecognize.ts (+ test)
  lib/multiPageSource.ts (+ test)
  lib/multiPageStepProps.ts
  lib/applyAiSession.ts (+ test)          ← R10 helper
  lib/getDraftBatchCount.ts (+ test)
  lib/batchReview.ts (+ test)             ← R12 getBatches fbs/bridge
  components/SourceImageGallery.tsx (+ test)
  components/SourceImageLightbox.tsx (+ test)
  components/SourceInputCard.tsx (+ test)
  components/OcrWaitBanner.tsx (+ test)
  components/AiInstructionBlock.tsx
  components/PageReviewNav.tsx (+ test)
  components/CommercialOfferWizard.tsx    ← handleApplyAi, multi wiring
  components/steps/*InputStep.tsx         ← review AI + list
  store/wizardDraftStore.tsx
  types/commercialOffer.ts
```

Несвязанный шум в `git status` (не трогать без нужды): `test_kp.pdf`, `tests/test_march_line_parser.py`.

---

## Команды проверки

```bash
cd frontend && npm run test -- src/features/commercial-offer
cd frontend && npm run typecheck
# Dev уже может быть поднят: ./run+logs.sh
```

Ручной smoke после A.4:

1. Несколько скринов → Распознать → дождаться первой ready (баннер исчезает).  
2. На сверке инструкция «замени н на п» → Применить → **список справа меняется**.  
3. Строки, которые парсер не ест → красные; принятые (`8н` если парсер ок) → без ложной подсветки.  
4. Apply с ошибкой сети → видна ошибка.

---

## Out of scope этого handoff

- Auto н→п в OCR pipeline  
- Server job Phase B  
- Commit / PR без просьбы  
- Повторный idea-refine  
- Redesign lightbox / progressive model  

---

## Definition of done (новое окно)

- [ ] Plan Task 1 (Apply sync + fail UX) + S1–S2 (+ R10 intact)  
- [ ] Plan Task 2 (red/yellow on review) + S3–S6  
- [ ] S7 suite + typecheck green  
- [ ] Spec/plan checklists → IMPLEMENT ✅  
- [ ] Краткий отчёт пользователю; **без commit**
