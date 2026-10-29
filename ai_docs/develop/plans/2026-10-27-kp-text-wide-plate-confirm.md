# Implementation Plan: «Список верен» записывает замену на тексте

**Created:** 2026-10-27  
**Spec:** [../../specs/kp-text-wide-plate-confirm.md](../../specs/kp-text-wide-plate-confirm.md)  
**Continues:** [../../specs/kp-wide-plate-slot-text.md](../../specs/kp-wide-plate-slot-text.md)  
**Goal:** На сверке обычного текста «Список верен» сначала записывает открытые слоты широких плит и сохраняет этот текст. Широкая строка, которую заменили или исключили, в состав не попадает.  
**Status:** IMPLEMENT done  
**SDD:** SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅

## Overview

`commitWideSlots` уже пишет «Заменить», «Исключить» и «Подтвердить» сверху вниз и останавливается на пустом поле. «ОК» и «Применить решения» его вызывают. «Список верен» на тексте вызывает только сравнение `state.batchReviewText` с последней пачкой. Пока слот не принят, эти строки равны, запрос не уходит, сверка закрывается, широкие марки остаются.

План добавляет чистую `prepareTextReviewConfirm` и один вызов в текстовой ветке `handleConfirmBatch`. Сервер, `commitWideSlots` и ветка очереди фото не меняются.

Каждый шаг: сначала тест из колонки RED, он падает. Потом минимальный код, пока зелёный только этот прогон. Параллельно T1 и T2 не запускать: мастер вызывает функцию из T1.

## Architecture Decisions

- Новая функция живёт в `widePlateSlots.ts` и вызывает `commitWideSlots` без своего алгоритма замены. Так пустое поле, порядок сверху вниз и уже принятые слоты остаются прежними.
- Возврат говорит мастеру две вещи: закрывать ли сверку и какой текст слать. `uploadText === null` при `closeReview: true` — пачка уже совпадает, `update` не нужен, сверка закрывается.
- Мастер подставляет в `update` строку из возврата, в том же обработчике. После `dispatch` читать `state.batchReviewText` нельзя: это ещё текст до записи.
- Слоты на входе — `wideSlotsByPageRef` ключа `text-review`, тот же ref, куда радио пишет «Заменить». Текст на входе — `state.batchReviewText` на момент клика.
- Ошибка записи обновляет текст и слоты в store так же, как `applyLocalWideSlots`, ставит ту же ошибку на блок и выходит до `mutateAsync` и до `confirm-batch-review`.
- Вызов стоит в ветке `handleConfirmBatch`, которая начинается после `if (multiPage.hasStarted ...) return`. Ветка снимков не получает эту функцию.
- `mergeEditedBatchIntoFullText` остаётся: в `update` уходит записанный текст последней пачки, ранние пачки не склеиваются заново.

## Dependency graph

```
T1 prepareTextReviewConfirm и тесты сценариев 1–6
  → T2 текстовая ветка handleConfirmBatch сохраняет возвращённый текст
        → Checkpoint
```

## Task 1: Чистая подготовка текста перед «Список верен»

**Description:** RED в `widePlateSlots.test.ts`, блок `prepareTextReviewConfirm`. Экспортировать функцию с сигнатурой из спеки. Внутри один вызов `commitWideSlots(editorText.split(/\r?\n/), slots)` без списка id: принятые слоты функция уже пропускает.

Кейсы в этом блоке, ожидания не ослаблять:

- `writes open replacements and keeps the typed neighbor`. Строки `ПБ 76-15-8 13`, `ПБ 76-12-8п 18`. Один открытый слот `lineIndex: 0`, `action: "replace"`, `replacementText: "ПБ 76-12-8п 13\nПБ 76-3.0-8п 13"`. Результат: `closeReview: true`, `uploadText` равен `ПБ 76-12-8п 13\nПБ 76-3.0-8п 13\nПБ 76-12-8п 18`, в нём нет `76-15-8`. Пачка на сервере — исходные две строки.
- `drops an excluded line before comparing with the server batch`. Открытый `exclude` на единственной `ПБ 68-15-8 15`. `uploadText` пустой смысл не несёт: текст после записи пустой, `uploadText === null`, `closeReview: true`. Так существующее сравнение `if (editedText && editedText !== original)` тоже не шлёт пустую строку. Для непустого списка из двух строк, где исключена первая, `uploadText` — оставшаяся вторая.
- `keeps confirmed lines and skips upload when the batch already matches`. Один открытый `confirm`, текст равен `serverBatchText`. `uploadText === null`, `closeReview: true`, текст тот же.
- `does not apply an accepted replacement again`. Слот `accepted: true`, `action: "replace"`, поле из двух строк, редактор уже `ПБ 76-12-8п 13`. `uploadText` этой замены второй раз не содержит, текст редактора тот же.
- `stops on an empty replacement and leaves the review open`. Два открытых replace, у второго `replacementText: "  "`. Первый записан, `closeReview: false`, `error` равен `Укажите заменяющие позиции.`, второго поля в тексте нет.
- `places both suggestion lines and leaves the following neighbor`. Покрыто первым кейсом; отдельный тест не добавлять, если первый уже ждёт обе строки и соседа.

**Acceptance criteria:**
- [x] Открытая замена даёт `uploadText` без широкой марки и с соседом
- [x] Пустая замена даёт `closeReview: false` и не затирает слоты ниже
- [x] «Подтвердить» при том же тексте даёт `uploadText === null` и `closeReview: true`
- [x] Принятый слот не пишется второй раз

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/widePlateSlots.test.ts`

**Dependencies:** None

**Files:**
- `frontend/src/features/commercial-offer/lib/widePlateSlots.ts`
- `frontend/src/features/commercial-offer/lib/widePlateSlots.test.ts`

**Scope:** Small

## Task 2: Текстовая ветка «Список верен» сохраняет записанный текст

**Description:** В `handleConfirmBatch` после раннего `return` ветки `multiPage.hasStarted` вызвать `prepareTextReviewConfirm`. Слоты: `wideSlotsByPageRef.current[TEXT_REVIEW_SLOT_PAGE_ID] ?? []`. Текст: `state.batchReviewText`. Пачка: `lastBatch?.normalized_text ?? ""`.

Дальше как `applyLocalWideSlots`: записать `wideSlotTextRef` и `wideSlotsByPageRef` для `text-review`, `dispatch` `set-batch-review-text` и `set-page-wide-slots`. При `closeReview: false` вызвать `setWidePlateError(plan.error)` и `return`. `updateInputMutation` и `confirm-batch-review` в этой ветке не выполняются.

При `closeReview: true` ошибку широких плит сбросить. В существующее сравнение подставить `plan.uploadText`: если оно не `null`, `mergeEditedBatchIntoFullText(batches, plan.uploadText)` и тот же `update` с `mode: "replace"`, затем `hydrate-draft` и `confirm-batch-review`. Если `null` — только `confirm-batch-review`, без запроса. Ветку `multiPage.hasStarted` не редактировать. `applyLocalWideSlots` не переписывать.

**Acceptance criteria:**
- [x] Текстовый «Список верен» шлёт `plan.uploadText`, а не `state.batchReviewText` с момента до записи
- [x] Ошибка записи не закрывает сверку и не вызывает `update`
- [x] В файле мастера `prepareTextReviewConfirm` встречается только в ветке без очереди фото

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/widePlateSlots.test.ts`
- [x] `cd /home/username/Code/plan_web/frontend && npm run typecheck`
- [x] Поиск по `CommercialOfferWizard.tsx`: вызов стоит ниже блока, который делает `return` для `multiPage.hasStarted`

**Dependencies:** Task 1

**Files:**
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`

**Scope:** Small

## Checkpoint: After Tasks 1–2

- [x] `widePlateSlots.test.ts` зелёный
- [x] `npm run typecheck` зелёный
- [ ] На IMPLEMENT, если мастер запущен: вставить текст с `ПБ 76-15-8 13` и соседней строкой другого количества, на сверке выбрать «Заменить», не нажимать «Применить решения», нажать «Список верен». В «Составе КП» нет `76-15-8`, сосед на месте, карточка «N позиций требуют внимания» по этой марке не возвращается
- [x] Имплемент выполнен по просьбе

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| После `dispatch` мастер снова прочитает старый `batchReviewText` | Замена снова не уйдёт на сервер | В `update` передавать только `plan.uploadText` |
| Пустой результат «Исключить» последней строки отправит пустой `update` и сотрёт заказ | Высокий для списка из одной строки | `uploadText` остаётся `null`, когда после `trim` текст пустой. Это текущее условие `if (editedText && ...)` |
| Запись на очереди фото начнёт менять чужой снимок | Средний | Вызов только после `return` ветки `multiPage.hasStarted`. T2 это проверяет поиском |
| `commitWideSlots` начнут менять вместе с обёрткой | Сломаются «ОК» и низ | T1 вызывает его как есть. Алгоритм в этой задаче не трогать |

## Open Questions

- Автозапись на снимке в эту работу не входит. Отдельная просьба — отдельная спека.
