# Spec: Apply AI → sync списка на сверке + red/yellow highlights

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-08-31  
**One-pager**: [ai_docs/ideas/kp-review-apply-sync-and-highlights.md](../ideas/kp-review-apply-sync-and-highlights.md)  
**План**: [ai_docs/develop/plans/2026-08-31-kp-review-apply-sync-and-highlights.md](../develop/plans/2026-08-31-kp-review-apply-sync-and-highlights.md)  
**Parent (wait + AI on review)**: [kp-ocr-wait-and-ai-on-review.md](./kp-ocr-wait-and-ai-on-review.md) · план [2026-08-31-kp-ocr-wait-and-ai-on-review.md](../develop/plans/2026-08-31-kp-ocr-wait-and-ai-on-review.md)  
**Related lint**: [unparsed-line-live-highlight.md](./unparsed-line-live-highlight.md) · multi-page [kp-multi-page-screenshots.md](./kp-multi-page-screenshots.md)

## Objective

**Проблема.** После Phase A.3 / A.3.1 (AI на batch-review, R10 не сбрасывает multi-сессию) «Применить инструкцию» может обновить draft, но **текст списка рядом с фото** не всегда отражает результат — менеджер не видит применённую правку. На сверке также нужна та же семантика подсветки, что у text lint: **красное = парсер отверг**, **жёлтое = уже существующие soft-сигналы**, без жёлтой эвристики н→п на строках, которые парсер **принял** (например `8н`).

**Цель.** (1) После успешного Apply — hydrate draft **и** записать обновлённый текст в review editor + store (batch / active page), чтобы список у фото сменился. (2) Visible error при провале Apply. (3) Red/yellow на review через **существующий** PlateListEditor / source lint — без новых suffix-heuristics и без auto н→п в OCR pipeline.

**Пользователь:** менеджер на batch-review после OCR, правит список руками и/или через AI.

**Успех:** Apply success → список рядом с фото = новый текст; Apply fail → ошибка видна; reject-строки красные; soft already-wired — жёлтые; `8н` (parser accept) без подсветки.

---

## ASSUMPTIONS (locked 2026-08-31)

1. **Apply → UI sync обязателен.** Успех = hydrate draft **и** запись обновлённого batch / active-page текста в review editor + store. Список beside image должен измениться без ручного refresh.
2. **Apply failure не молчит.** Visible error (toast / inline / существующий error surface продукта) — не silent catch.
3. **Red = parser reject / unparsed.** Та же семантика, что у text lint / PlateListEditor «не распарсилось» — не отдельная цветовая схема.
4. **Yellow = только soft signals уже в продукте** (например OCR corrections), **если уже wired**. Не добавлять новую load-suffix `н` heuristic.
5. **Parser accept → no highlight.** Строка вроде `8н`, которую парсер принимает, **не** жёлтеет и **не** краснеет из-за суффикса.
6. **Нет auto н→п в OCR pipeline** — граница parent-спеки сохраняется.
7. **Reuse existing AI API** — без новой модели / endpoint.
8. **Коммиты агент не делает**, пока явно не попросите.

→ Assumptions approved / locked with ideation 2026-08-31.

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| **D-apply-sync** | After «Применить инструкцию» | On success: hydrate draft **and** write updated text into review editor + store (batch / active page) so list beside image updates |
| **D-apply-error** | Apply failure UX | Visible error — not silence |
| **D-red** | Hard highlight | **Red** = parser reject / unparsed (existing PlateListEditor / source lint) |
| **D-yellow** | Soft highlight | **Yellow** = other soft signals **already in product** if wired — **do not** add н→п / load-suffix heuristic |
| **D-no-false-yellow** | Accepted lines | If parser accepts (e.g. `8н`) → **do not** highlight |
| **D-no-auto-suffix** | OCR pipeline | **No** auto н→п / suffix rewrite (unchanged from parent) |

---

## User Stories

- Как **менеджер**, после успешного «Применить инструкцию» я сразу вижу **новый** текст списка рядом с фото страницы.
- Как **менеджер**, если Apply упал, я вижу ошибку и понимаю, что список не обновился.
- Как **менеджер**, на сверке строки, которые парсер не ест, красные — как в text lint.
- Как **менеджер**, уже существующие soft-проблемы (если wired) жёлтые; принятые строки вроде `8н` не подсвечены ложно.

---

## Success Criteria

| # | Критерий | Метод |
|---|----------|--------|
| S1 | Apply AI success → review editor + store text matches updated list; UI beside image shows new text | unit / RTL |
| S2 | Apply AI failure → visible error; list not silently stale-as-success | RTL / unit |
| S3 | Batch-review: parser-reject / unparsed lines → **red** (existing lint language) | RTL |
| S4 | Soft signals already in product → **yellow** if wired; no new н-suffix heuristic | RTL + code review |
| S5 | Parser-accepted line `8н` (or equivalent) → **not** highlighted | RTL |
| S6 | No OCR pipeline auto suffix rewrite | code review / no pipeline rule diff |
| S7 | `npm run test -- src/features/commercial-offer` + typecheck green | CI / локально |

---

## Boundaries

**Always**
- Sync review editor + store on Apply success
- Surface Apply errors visibly
- Reuse PlateListEditor / source lint for red/yellow semantics
- Leave parser-accepted lines unhighlighted

**Ask first**
- Extending yellow to new soft-signal types not already in product
- Changing ApplyAi payload scope (active page vs full draft) beyond syncing result into UI

**Never**
- Auto н→п (or similar) in OCR / post-OCR pipeline
- New yellow heuristic for load-suffix `н` on accepted lines
- Highlight all lines
- New AI model / redesign layout
- Commit unless explicitly asked

---

## Out of scope (Not Doing)

- Automatic suffix rewrite in OCR pipeline  
- Highlight every line  
- New AI model / new apply endpoint (unless IMPLEMENT proves impossible — ask first)  
- Layout redesign of review / lightbox / progressive model  
- New soft-signal categories beyond what product already wires  

---

## Phased delivery

### Phase A.4 — Apply sync + review highlights (эта спека)

1. Apply → editor/store sync + visible error + tests (S1–S2, S7) ✅  
2. Red/yellow on batch-review via existing lint + tests (S3–S6, S7) ✅

Depends on parent Phase A.3 / A.3.1 (AI on review, R10 no-reset) — done.

---

## Open Questions

- Exact yellow soft-signal set already wired on review — confirm at IMPLEMENT; do not invent  
- After draft-level Apply, whether all ready pages’ editors need refresh vs active page only — follow current API result shape  

---

## Relation to parent review notes

Parent [kp-ocr-wait-and-ai-on-review.md](./kp-ocr-wait-and-ai-on-review.md) closed R10 (no reset) / R11 (AI enabled during tail OCR). This spec is the **follow-on**: make Apply **visible in the list beside the photo** and align review highlights with existing lint — without false yellow on accepted `н` lines.
