# КП сверка: Apply AI → sync списка + red/yellow подсветка

> Ideation locked: 2026-08-31  
> Parent: [`kp-ocr-wait-and-ai-on-review.md`](./kp-ocr-wait-and-ai-on-review.md)  
> Спека: [`../specs/kp-review-apply-sync-and-highlights.md`](../specs/kp-review-apply-sync-and-highlights.md)  
> План: [`../develop/plans/2026-08-31-kp-review-apply-sync-and-highlights.md`](../develop/plans/2026-08-31-kp-review-apply-sync-and-highlights.md)  
> Related: [`unparsed-line-live-highlight.md`](./unparsed-line-live-highlight.md) (source lint language), [`kp-multi-page-screenshots.md`](./kp-multi-page-screenshots.md)

## Problem Statement

**How might we** after «Применить инструкцию» show the **updated** list text next to the photo, and on batch-review **highlight** lines the parser rejects in **red** (same as text lint), with other soft problem signals in **yellow** — without painting accepted lines (e.g. `8н`) yellow via a н→п heuristic?

## Recommended Direction

Follow-on to wait-banner + AI-on-review (Phase A.3 / A.3.1). Session no longer resets on Apply (R10), but the list beside the image can still lag the AI result; highlights on review must reuse existing lint, not invent new suffix heuristics.

1. **Fix Apply AI → UI sync.** On success: hydrate draft **and** write updated batch / active-page text into the review editor + store so the list next to the photo changes. On failure: **visible** error (not silence).
2. **Batch-review highlights via existing PlateListEditor / source lint.** **Red** = parser reject / unparsed. **Yellow** = other soft signals already in product (e.g. OCR corrections) **if already wired** — do **not** add a new load-suffix `н` heuristic. If the parser **accepts** a line (e.g. `8н`), do **not** highlight it.
3. **No OCR pipeline auto н→п.** Same boundary as parent: no auto suffix rewrite in preprocess/verify/upscale path.

Итог: после Apply менеджер сразу видит новый текст у фото; на сверке красное = «парсер не ест», жёлтое = уже существующие soft-сигналы; принятые строки (включая `8н`) без ложной подсветки.

## Key Assumptions to Validate

- [x] Apply success must update review editor + per-page / store text, not only draft hydrate
- [x] Apply failure must surface a visible error
- [x] Red = existing parser-reject / unparsed lint language (same as text path / PlateListEditor)
- [x] Yellow = only soft signals already in product; no new н→п / load-suffix heuristic
- [x] Parser-accepted lines (e.g. `8н`) stay unhighlighted
- [x] No OCR pipeline auto suffix rewrite

## MVP Scope

**In:**
- Apply AI → sync review editor + store (batch / active page) after success
- Visible error path when Apply fails
- Red/yellow highlights on batch-review list via **existing** PlateListEditor / source lint wiring + tests
- Unit/RTL for sync and highlight behavior

**Out:** auto suffix rewrite; highlight all lines; new AI model; layout redesign

## Not Doing (and Why)

- **Автоматическая замена суффикса н→п в OCR pipeline** — контроль остаётся у менеджера / AI-инструкции; не «тихие» правки
- **Жёлтая эвристика н→п на принятых строках** — парсер принял `8н` → не красить; иначе ложные тревоги
- **Подсветить все строки** — шум; только reject (red) и уже существующие soft signals (yellow)
- **Новая AI-модель / endpoint** — reuse текущего applyAi
- **Redesign layout сверки** — только sync + lint wiring

## Open Questions

- Какой именно soft-signal уже wired как yellow на review (OCR corrections и т.п.) — подтвердить при IMPLEMENT; не изобретать новые
- Apply sync: достаточно active page + draft, или нужен явный refresh всех ready pages после draft-level AI — уточнить по текущему API shape
