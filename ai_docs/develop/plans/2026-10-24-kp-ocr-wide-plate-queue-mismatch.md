# Implementation Plan: Список замены широкой плиты остаётся на своём фото

**Created:** 2026-10-24  
**Spec:** [../../specs/kp-ocr-wide-plate-queue-mismatch.md](../../specs/kp-ocr-wide-plate-queue-mismatch.md)  
**Goal:** «Применить решения» на сверке нескольких фото пишет замену на ту же страницу, где снимок, и не показывает «Число списков не совпадает с числом фото».  
**Status:** IMPLEMENT done  
**SDD:** SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅

## Overview

Обычная сессия из нескольких фото уже хранит OCR страницы `i` в `PageSource.batchReviewText` и в `plate_batches[i]`. Пока сверка не подтверждена, `hydrate-draft` кладёт в `state.batchReviewText` текст последней партии. Фокус остаётся на текущем готовом снимке. `handleApplyWidePlates` подставляет этот store в слот активного фото, flush не равен `input_text`, `update` с `mode: "replace"` схлопывает партии в одну. Resolve возвращает одну партию, `syncWidePlateQueuePages` отказывается (`N !== 1`). К этому моменту `sync-after-wide-plates` уже очистил `widePlateActions`, карточка снова открывается на «Подтвердить».

Фикс на клиенте, без нового URL и без нового поля. Очередь фото не вызывает `update` `replace`. В flush активного снимка уходит текст этой страницы. После успеха страница `i` получает свою партию. Если менеджер правил OCR этой страницы, на ней остаётся её текст с полной заменой (2 или 4 строки), чужой список на неё не кладётся. Пока новый список на страницу не записан, выбор `replace` / `exclude` не становится `confirm`.

## Architecture Decisions

- Скриншотный путь (список не правили, store затёрт последней партией) лечится тем, что `update` `replace` перед resolve не вызывается. Контракт `update` и `wide-plates/resolve` не меняем (D9).
- Правку списка у фото нельзя положить в её партию существующим `update`: ветка не-append всегда ставит `batches = [одна партия]`, а `append` добавляет новую. Отличить этот вызов от «Список верен» и текстового ввода без нового поля нельзя, смысл `replace` для них не трогаем (D8). Поэтому очередь фото `replace` не шлёт вообще. Resolve по-прежнему переписывает каждую уже существующую партию на месте. `commercial_plate_resolve.py` не меняем.
- Если текст страницы совпал с её партией до resolve, после успеха страница получает `plate_batches[i]`. Если менеджер правил OCR, на страницу пишется её текст, а широкая строка заменяется целиком текстом решения (без потолка в две строки). Соседняя страница не затрагивается. Неширокие правки уезжают на сервер вместе с «Список верен», когда очередь уже сбрасывается и `replace` по-прежнему схлопывает список в один текст.
- `syncWidePlateQueuePages` при разной длине по-прежнему возвращает отказ и не придумывает пары (D6). Страницы `pending` / `error` этот отказ оставляют.
- Пока запись на страницу не случилась, store держит выбор (`widePlateChoicesHeld`). `sync-after-wide-plates` и `hydrate-draft` в этом состоянии не очищают `widePlateActions` и не подменяют текст сверки. Карточка не открывается заново из старого OCR как «Подтвердить» (D7).
- «Список верен» на очереди фото шлёт `PageSource.batchReviewText` активной страницы, а не store, если hydrate снова положил туда последнюю партию (D8).
- Суффикс «п» до apply не добавляем и не снимаем. `buildAutoSplitSuggestion` остаётся подсказкой. Verify OCR, invalid width и unpriced не трогаем.

## Dependency graph

```
T1 flush страницы и запрет replace на очереди фото
        |
        +---- T2 текст страницы после resolve (своя партия / своя правка)
        |
        +---- T3 выбор не падает в confirm, пока страница не обновлена
                    |
               T4 «Список верен» шлёт текст страницы
                    |
               Checkpoint
```

T2 и T3 можно писать после T1. T4 зависит от того, что успешный apply уже записал текст на страницу.

## Task 1: Flush активного фото — текст страницы, без replace

**Description:** Новая сборка flush для apply берёт `batchReviewText` каждой страницы, включая активную, и не подставляет `state.batchReviewText`, если там последняя партия. `flushThenResolveWidePlates` при `photoQueue: true` не вызывает `updateInput` с `mode: "replace"`, даже если flush отличается от `input_text`. Текстовый ввод без очереди фото по-прежнему вызывает replace, когда тексты различаются. Четыре строки `replacementText` уходят в decision целиком.

**Acceptance criteria:**
- [x] Две страницы `ready`, активна первая, store равен второй, persisted — склейка страниц: flush равен склейке страниц, `update` replace не вызывается
- [x] Одна страница и тот же текст, что `input_text`: replace не вызывается, resolve вызывается
- [x] Очередь фото с правкой страницы (flush ≠ persisted) тоже не вызывает replace
- [x] Без `photoQueue` старый тест «update до resolve» остаётся зелёным
- [x] `replacementText` из четырёх строк не обрезается

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/flushThenResolveWidePlates.test.ts`

**Dependencies:** None

**Files likely touched:**
- `frontend/src/features/commercial-offer/lib/flushThenResolveWidePlates.ts`
- `frontend/src/features/commercial-offer/lib/flushThenResolveWidePlates.test.ts`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- `frontend/src/features/commercial-offer/hooks/useCommercialOfferWizard.ts`

**Estimated scope:** Small

## Task 2: Своя партия на свою страницу, правка не схлопывается

**Description:** После resolve `syncWidePlateQueuePages` по-прежнему пишет `plate_batches[i]` на страницу `i` и отказывается, если длины не равны. Мастер для страницы, чей текст до кнопки совпал с её партией, берёт этот результат. Для страницы, которую правили, пишет её текст и подставляет в широкую строку полный `replacementText` (совпадение марки как у `normalizeLineKey`: `п` и «шт» не мешают). Вторая страница не получает список первой.

**Acceptance criteria:**
- [x] Две партии и две страницы: у первой четыре строки замены и нет `76-15-8`, вторая без изменений
- [x] Разная длина → `ok: false`, поля `pages` нет
- [x] Неправленная страница получает текст своей партии после resolve
- [x] Правленная страница сохраняет свою неширокую строку и все строки замены

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/syncWidePlateQueuePages.test.ts`

**Dependencies:** Task 1

**Files likely touched:**
- `frontend/src/features/commercial-offer/lib/syncWidePlateQueuePages.ts`
- `frontend/src/features/commercial-offer/lib/syncWidePlateQueuePages.test.ts`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`

**Estimated scope:** Small

## Task 3: Отказ zip не откатывает радио в «Подтвердить»

**Description:** На очереди фото `sync-after-wide-plates` и запись черновика в кэш откладываются до успешного zip. Пока страницы не обновлены, стоит `widePlateChoicesHeld`: и sync, и hydrate не очищают `widePlateActions` и не подменяют текст сверки текстом последней партии. Успешный zip снимает удержание и только тогда делает обычный sync. Ошибка resolve снимает удержание, но не стирает выбор. Текстовый путь без фото по-прежнему очищает выбор в `sync-after-wide-plates`.

**Acceptance criteria:**
- [x] Существующий тест sync по-прежнему очищает `widePlateActions` и обновляет `batchReviewText`
- [x] Пока удержание включено, sync с уже разрешённым черновиком не делает `replace` видом `confirm` и не затирает старую широкую строку в тексте сверки
- [x] Hydrate с `wide_plates_resolved: true` при удержании тоже не очищает выбор

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/store/wizardDraftStore.test.tsx`

**Dependencies:** Task 1

**Files likely touched:**
- `frontend/src/features/commercial-offer/store/wizardDraftStore.tsx`
- `frontend/src/features/commercial-offer/types/commercialOffer.ts`
- `frontend/src/features/commercial-offer/store/wizardDraftStore.test.tsx`
- `frontend/src/features/commercial-offer/store/draftStorage.test.ts`
- `frontend/src/features/commercial-offer/hooks/useCommercialOfferWizard.ts`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`

**Estimated scope:** Medium

## Task 4: «Список верен» шлёт уже записанный текст страницы

**Description:** После успешного apply мастер пишет текст страницы в `PageSource` и выставляет store текстом активной страницы, а не последней партии. Пока идёт сверка фото, эффект снова подтягивает store из активной страницы, если hydrate положил последнюю партию. «Список верен» для очереди фото берёт `active.batchReviewText`. Смысл `mode: "replace"` после сброса очереди и для текстового ввода не меняется.

**Acceptance criteria:**
- [x] Подтверждение активной страницы не читает отравленный store
- [x] Текстовый ввод без `multiPage.hasStarted` не меняет ветку replace
- [x] Успешный apply не оставляет на активной странице `ПБ 76-15-8`

**Verification:**
- [x] Целевые vitest из спеки и `cd /home/username/Code/plan_web/frontend && npm run typecheck`

**Dependencies:** Task 2, Task 3

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`

**Estimated scope:** Small

## Checkpoint: Complete

- [x] S1–S6 закрыты тестами Task 1–4. S7 сервер не трогает: `test_align_text_batch_normalized_uses_parser_text` не запускаем отдельно, файл resolve не меняется
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/syncWidePlateQueuePages.test.ts src/features/commercial-offer/lib/flushThenResolveWidePlates.test.ts src/features/commercial-offer/store/wizardDraftStore.test.tsx`
- [x] `cd /home/username/Code/plan_web/frontend && npm run typecheck`
- [x] Серверный pytest не нужен: `commercial_plate_resolve.py` и `product_draft_handler.py` не меняются

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| В очереди есть `pending` / `error`, либо партии не из этих страниц | Zip по-прежнему отказывает. Это задумано | Выбор остаётся на экране, тексты страниц не выдумываем |
| Неширокая правка OCR не попадает в `plate_batches` до «Список верен» | До подтверждения заказ на сервере считает исходный OCR плюс замену широкой строки | Страница показывает правку и полную замену. «Список верен» шлёт этот текст. Новое поле update не добавляем |
| Поздний hydrate после успеха снова кладёт в store последнюю партию | «Список верен» мог бы вернуть чужой список | Подтверждение читает текст страницы; эффект возвращает store к активной странице |
| Удержание выбора залипает | Карточка широких плит не гаснет | Снимаем удержание после успешного zip, при ошибке resolve, при reset и при новой сверке |

## Open Questions

Нет. Контракт не меняем: правку страницы нельзя записать в её партию отдельным полем, не сломав `replace` для текста и для «Список верен». Этот план такой путь не добавляет.
