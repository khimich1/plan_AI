# КП OCR: wait-баннер + AI на экране сверки

> Ideation locked: 2026-08-31  
> Parent: [`kp-multi-page-screenshots.md`](./kp-multi-page-screenshots.md)  
> Спека: [`../specs/kp-ocr-wait-and-ai-on-review.md`](../specs/kp-ocr-wait-and-ai-on-review.md)  
> План: [`../develop/plans/2026-08-31-kp-ocr-wait-and-ai-on-review.md`](../develop/plans/2026-08-31-kp-ocr-wait-and-ai-on-review.md)  
> Multi-page спека/план: [`../specs/kp-multi-page-screenshots.md`](../specs/kp-multi-page-screenshots.md), [`../develop/plans/2026-08-31-kp-multi-page-screenshots.md`](../develop/plans/2026-08-31-kp-multi-page-screenshots.md)  
> Follow-on (Apply sync + red/yellow): [`kp-review-apply-sync-and-highlights.md`](./kp-review-apply-sync-and-highlights.md)

## Problem Statement

**How might we** not leave the manager in silence while the first page OCR runs, and let them fix OCR mistakes (esp. load suffix п↔н) on the same review screen with editable text + AI?

## Recommended Direction

Два UX-улучшения поверх уже готового multi-page progressive review — без правок OCR-пайплайна и без Phase B.

1. **Wait UX.** Пока `hasStarted &&` ещё ни одна страница не `ready`, показать баннер + простой спиннер: «Идёт распознавание, подождите 1–2 минуты». Как только первая страница становится `ready` — баннер скрыть; прогрессивная сверка без изменений.
2. **п/н и похожие ошибки OCR.** Не добавлять автоматическое правило post-OCR (н→п и т.п.) в пайплайн. Исправлять **в окне сверки**: редактируемый текст (уже есть) + блок AI-инструкции на том же batch-review экране.
3. **Перенести AI-блок** с «Дополнительно» / карточки следующего append-захода на **batch-review page** для **всех** типов изделий (plates, piles, marches, steps, bridge, fbs). Текст на сверке остаётся редактируемым.

Итог: менеджер видит ожидание первой страницы, затем правит список руками и/или через AI на том же экране, где сверяет фото.

## Key Assumptions to Validate

- [x] Пользователь явно отказался от auto suffix rewrite в OCR pipeline — fix в UI сверки
- [x] Wait-баннер нужен только до первой `ready`; дальше progressive review as-is
- [x] AI на review нужен для всех product types с `*InputStep` / `SourceInputCard`
- [x] Существующий `applyAi` / instruction API достаточно переиспользовать с review (без нового endpoint)
- [x] Перенос AI-блока с «Дополнительно» не ломает append-сценарий (оставить там или дублировать осознанно — см. Open Questions)

## MVP Scope

**In:**
- Баннер + spinner при `hasStarted && no ready yet`; hide on first `ready`
- Перенос / показ AI instruction UI на batch-review для всех шести `*InputStep`
- Редактирование текста на review (уже есть — не ломать)
- Unit/RTL на wait-баннер и наличие AI UI на review

**Out:** auto н→п в pipeline; Phase B server job; mandatory AI на каждый OCR; redesign lightbox / progressive review

## Not Doing (and Why)

- **Автоматическая замена суффикса н→п (и аналоги) в OCR pipeline** — пользователь хочет контроль в окне сверки, не «тихие» правки модели/правил
- **Phase B server job** — отдельная фаза multi-page; не связана с wait/AI UX
- **Обязательный AI на каждый OCR** — инструкция опциональна; менеджер правит руками или жмёт «Применить», когда нужно
- **Redesign lightbox / progressive review** — A.2 и Phase A уже закрыты; трогаем только wait-gap и placement AI

## Open Questions

- Оставлять ли AI-блок в «Дополнительно» после переноса на review (дубль) или убрать оттуда для replace-захода?
- Копирайт баннера: фиксированные «1–2 минуты» vs динамика по числу страниц?
- AI на review применяется к **активной** странице / её тексту или ко всему draft batch — подтвердить при IMPLEMENT (предпочтение: как сейчас applyAi к draft, с фокусом на текущий editable list)

## Follow-on

Apply success must also refresh the list beside the photo; batch-review red/yellow via existing lint (no н→п yellow heuristic): [`kp-review-apply-sync-and-highlights.md`](./kp-review-apply-sync-and-highlights.md).
