# Implementation Plan: Плашка и слот широкой плиты на фото и на тексте

**Created:** 2026-10-26  
**Spec:** [../../specs/kp-wide-plate-slot-text.md](../../specs/kp-wide-plate-slot-text.md)  
**Idea:** [../../ideas/wide-plate-slot.md](../../ideas/wide-plate-slot.md)  
**Goal:** Один блок на фото и на тексте без фото: свёрнутая плашка со счётом, маленькая «ОК» в ряду радио, локальная запись. `wide_plates_resolved` после «Подтвердить» шаг не держит.  
**Status:** IMPLEMENT done  
**SDD:** SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅

## Overview

Слот, перенос строки и `commitWideSlots` уже есть. Этот план их не переписывает. Меняется оболочка `WidePlateSlotSection` и то, какой экран её вызывает.

Сейчас секция всегда раскрыта, «ОК» стоит отдельной строкой сетки и растягивается на ширину карточки. Плашка «N позиций требуют внимания» осталась только у `WidePlatesInlineSection`. Текст без фото по-прежнему монтирует старый блок и уходит в `flushThenResolveWidePlates`, потому что `applyLocalWideSlots` выходит сразу, если нет `multiPage.activeId`.

Каждый шаг: сначала тест из колонки RED, он падает. Потом минимальный код, пока зелёный только этот прогон. `carryWideSlots` не менять. `WidePlatesInlineSection` не удалять: тесты шага, которые не передают `wideSlots`, продолжают видеть старый блок. Мастер на сверке фото и на тексте без фото старый блок больше не монтирует.

## Architecture Decisions

- N на плашке — число слотов с `accepted === false` после `uniqueWideSlotsByLine`. Это то же множество, которое держит «Готово, далее»: широкая строка без принятого слота.
- «ОК» — последний элемент flex-ряда радио, `variant="secondary"`, без `fullWidth` и без отдельной строки сетки. Иначе грид карточки снова растянет её в полосу.
- `routeWidePlateApply(false)` остаётся `"resolve"`. Так подтверждённая очередь фото не начинает писать локально. Открытая сверка текста — отдельный вход той же функции и возвращает `"local"`.
- Слоты текста лежат в `wideSlotsByPage["text-review"]`. Константа `TEXT_REVIEW_SLOT_PAGE_ID`. Id страницы фото туда не пишется.
- Текст списка для этого ключа — `state.batchReviewText`. Перенос при правке — уже существующий `carryWideSlots`.
- Гейт «Готово, далее» на тексте без фото вызывает `photoQueueWidePlateBlocksFinish` для одной страницы `text-review`. `wide_plates_resolved` в этом условии не читается. «Список верен» по-прежнему блокирует «Готово, далее», пока сверка не подтверждена.
- Смена ключа списка (`collapseKey`: id фото или `text-review`) снова ставит `expanded` в false. Раскрытие не переезжает на соседний снимок.
- Параллельно эти шаги не запускать. T1–T3 делят секцию, T4–T7 делят мастер и шаг.

## Dependency graph

```
T1 плашка свёрнута
  → T2 «ОК» в ряду радио
  → T3 смена снимка снова сворачивает
        → T4 маршрут записи: текст тоже local, подтверждённое фото всё ещё resolve
        → T5 шаг показывает секцию слотов без снимков
        → T6 гейт шага не смотрит на wide_plates_resolved
              → T7 мастер пишет text-review и не зовёт resolve
                    → Checkpoint
```

## Task 1: Плашка свёрнута и считает открытые слоты

**Description:** RED в `WidePlateSlotSection.test.tsx`: `starts collapsed and shows the attention count`. Один открытый слот. Есть кнопка `▸ 1 позиций требуют внимания` и абзац `Разверните блок, чтобы принять решение по каждой позиции, затем нажмите «Применить решения».` Радио, поле замены и «Применить решения» отсутствуют. Второй кейс в том же тесте: два открытых и один `accepted` дают `▸ 2 позиций требуют внимания`. Клик по плашке показывает `▾ 2 позиций требуют внимания` и радио. Уже существующие тесты секции, которые ищут радио, «ОК» или низ, после этого кликают плашку. Их утверждения про радио, поле и id слота не удалять.

**Acceptance criteria:**
- [x] До клика панелей нет, после клика панель есть
- [x] N не считает `accepted`
- [x] Ноль открытых слотов по-прежнему не рисует секцию

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/components/WidePlateSlotSection.test.tsx`

**Dependencies:** None

**Files:**
- `frontend/src/features/commercial-offer/components/WidePlateSlotSection.tsx`
- `frontend/src/features/commercial-offer/components/WidePlateSlotSection.test.tsx`

**Scope:** Small

## Task 2: «ОК» стоит в ряду радио

**Description:** RED. Тест `places OK in the radio row`. После раскрытия родитель радио «Исключить» содержит кнопку «ОК». Родитель поля «Заменяющие позиции» эту кнопку не содержит. Кнопка не `fullWidth`. Реализация: перенести `Button` в тот же flex-ряд, что радио, `variant="secondary"`, `type="button"`. Ряд с `alignItems: "center"`, чтобы высота кнопки совпала с переключателями.

**Acceptance criteria:**
- [x] «ОК» сразу после «Исключить» в одном ряду
- [x] При «Заменить» поле под рядом, а «ОК» остаётся в ряду
- [x] Клик «ОК» по-прежнему зовёт `onAccept` с id слота

**Verification:**
- [x] тот же прогон `WidePlateSlotSection.test.tsx`

**Dependencies:** Task 1

**Files:** те же два

**Scope:** Small

## Task 3: Смена списка снова сворачивает плашку

**Description:** RED. Тест `collapses again when collapseKey changes`. Рендер с `collapseKey="page-a"`, клик по плашке, радио видно. `rerender` с `collapseKey="page-b"` и теми же слотами: радио нет, на плашке снова `▸`. Тот же ключ при ререндере список не сворачивает.

**Acceptance criteria:**
- [x] Новый `collapseKey` ставит `expanded` в false
- [x] Тот же ключ оставляет раскрытие

**Verification:**
- [x] тот же прогон `WidePlateSlotSection.test.tsx`

**Dependencies:** Task 2

**Files:** те же два

**Scope:** Small

## Checkpoint: After Tasks 1–3

- [x] `WidePlateSlotSection.test.tsx` зелёный
- [x] Панели не видны, пока плашку не нажали
- [x] «ОК» в ряду радио

## Task 4: Открытый текст пишется локально, подтверждённое фото — нет

**Description:** RED в `widePlateSlots.test.ts`. Не менять тест `photo queue applies locally`: `routeWidePlateApply(false)` остаётся `"resolve"`. Новый тест `open text review applies locally and a confirmed photo queue still resolves`. Расширить вход, не ломая старый boolean: например второй аргумент `{ textReview: true }`. `routeWidePlateApply(false, { textReview: true })` равен `"local"`. `routeWidePlateApply(false)` и `routeWidePlateApply(true, { textReview: false })` при уже подтверждённой очереди не требуются; достаточно, что вызов без text review и с `photoQueue: false` остаётся `"resolve"`. Экспортировать `TEXT_REVIEW_SLOT_PAGE_ID = "text-review"`.

**Acceptance criteria:**
- [x] Старый вызов `routeWidePlateApply(false)` не стал `"local"`
- [x] Открытая сверка текста возвращает `"local"`
- [x] Константа ключа равна `text-review`

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/widePlateSlots.test.ts`

**Dependencies:** None. Писать после T3, не параллельно с мастером.

**Files:**
- `frontend/src/features/commercial-offer/lib/widePlateSlots.ts`
- `frontend/src/features/commercial-offer/lib/widePlateSlots.test.ts`

**Scope:** Small

## Task 5: Шаг без снимков показывает секцию слотов

**Description:** RED в `PlateInputStep.test.tsx`. Тест `shows the slot section for text review without photos`. `pages=[]`, `pendingBatchReview`, `wideSlots` с одним открытым слотом на `ПБ 76-15-8 13`. Видна плашка `▸ 1 позиций требуют внимания`. Нет текста `Количество:`. «Список верен» enabled. `collapseKey` для этого экрана мастер позже передаст как `text-review`; в тесте шага достаточно, что секция слотов смонтирована, когда `wideSlots` передан и снимков нет. Тесты, которые не передают `wideSlots` и кликают старую плашку (`shows wide card after batch review on post-confirm screen`, `shows wide card for Excel tab paste`, `shows the wide card on a photo whose known marks are not applied yet`), не переписывать.

**Acceptance criteria:**
- [x] Без страниц и с переданным `wideSlots` монтируется новая секция
- [x] Старые тесты шага без `wideSlots` остаются зелёными без смены ожиданий

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/components/steps/PlateInputStep.test.tsx`

**Dependencies:** Task 1

**Files:**
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.test.tsx`

**Scope:** Small

## Task 6: «Готово, далее» на тексте смотрит на слоты

**Description:** RED. Два теста шага при `pages=[]`, `pendingBatchReview={false}`, `batchReviewText="ПБ 76-15-8 13"`, `wide_plates_resolved: false`, остальные гейты ширины и цены закрыты. `enables Готово, далее when the text-review slot is accepted`: слот `text-review` с `accepted: true` — кнопка enabled. `holds Готово, далее when the text-review slot is open`: тот же слот с `accepted: false` — кнопка disabled и title `Сначала примите решение по позициям шире стандартной`. Широкая строка без слота тоже держит кнопку, даже если `wide_plates_resolved: true`. Условие в шаге для экрана без очереди фото: `photoQueueWidePlateBlocksFinish([{ text: batchReviewText, slots: wideSlotsByPage["text-review"] ?? [] }])`, когда `wideSlots` передан. Ветку `queueHasUndecidedWidePlates` и тесты invalid width / unpriced не менять.

**Acceptance criteria:**
- [x] Принятый слот не держит кнопку при `wide_plates_resolved === false`
- [x] Открытый слот или широкая строка без слота держит кнопку при `wide_plates_resolved === true`
- [x] Фото-тесты `does not hold Готово, далее for accepted photo slots` и `holds Готово, далее when a photo page has a wide line and no slot` остаются зелёными

**Verification:**
- [x] тот же прогон `PlateInputStep.test.tsx`

**Dependencies:** Task 4, Task 5

**Files:** те же, что в Task 5, плюс импорт гейта из `widePlateSlots.ts`

**Scope:** Small

## Checkpoint: After Tasks 4–6

- [x] `widePlateSlots.test.ts` и `PlateInputStep.test.tsx` зелёные
- [x] `routeWidePlateApply(false)` всё ещё `"resolve"`
- [x] Шаг без фото с переданными слотами не читает `wide_plates_resolved`

## Task 7: Мастер пишет текст без фото локально

**Description:** Подключить уже проверенные функции. Не рендерить весь мастер, если в репозитории нет его теста: проверка — вызовы в коде и зелёные тесты T4–T6. В `CommercialOfferWizard` для открытой сверки без страницы фото:

- слоты читать и писать в `wideSlotsByPage[TEXT_REVIEW_SLOT_PAGE_ID]`;
- `applyLocalWideSlots` не выходит по пустому `multiPage.activeId`: цель — `text-review`, строки — `state.batchReviewText`;
- `handleApplyWidePlates` при `routeWidePlateApply(false, { textReview: true })` делает локальную запись и `return` до `flushThenResolveWidePlates`;
- правка `batchReviewText` на этом экране зовёт `carryWideSlots` и сохраняет слоты `text-review`;
- первый показ сверки, если слотов ещё нет, сеет их через `carryWideSlots` от текущего текста, как эффект фото, и не затирает уже лежащие слоты повторным прогоном от пустой предыдущей строки;
- в `PlateInputStep` передать `wideSlots`, `onAcceptWideSlot`, `onApplyWideSlots`, `onWideSlotAction`, `onWideSlotReplacement`, `collapseKey={TEXT_REVIEW_SLOT_PAGE_ID}`;
- `hasUnresolvedWidePlates` мастера при `!multiPage.hasStarted` считает `photoQueueWidePlateBlocksFinish` по `text-review` и `batchReviewText`, не флаг `wide_plates_resolved`;
- на фото `collapseKey` — id активной страницы. `text-review` в слоты фото не писать.

Пустая замена, пачка сверху вниз и липкое поле уже покрыты `widePlateSlots.test.ts`. Новый алгоритм diff не писать.

**Acceptance criteria:**
- [x] «ОК» и «Применить решения» на тексте без фото не вызывают `resolveWidePlatesMutation`
- [x] Запись идёт в `batchReviewText` и в слоты `text-review`
- [x] Подтверждённая очередь фото по-прежнему попадает в ветку resolve: `routeWidePlateApply(false)` без `textReview`
- [x] Навигация мастера на тексте без фото отпускает шаг, когда слот `accepted`, даже если `wide_plates_resolved` ложен

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/widePlateSlots.test.ts src/features/commercial-offer/components/WidePlateSlotSection.test.tsx src/features/commercial-offer/components/steps/PlateInputStep.test.tsx`
- [x] `cd /home/username/Code/plan_web/frontend && npm run typecheck`
- [x] В `handleApplyWidePlates` локальная ветка текста стоит выше `flushThenResolveWidePlates` и выходит из функции

**Dependencies:** Task 3, Task 4, Task 6

**Files:**
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`

**Scope:** Medium

## Checkpoint: Complete

- [x] Команды из спеки зелёные
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run` — полный прогон. Уже известные таймауты `PileCostPanel.test.tsx` и `ShipmentDrawer.test.tsx`, если падают только они и по отдельности зелёные, не чинить в этом плане
- [x] `carryWideSlots` и тесты `flushThenResolveWidePlates` не переписаны
- [x] Коммит не делать, пока его не попросят

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| `routeWidePlateApply(false)` станет `"local"` | Подтверждённые фото перестанут звать resolve | T4 явно хранит старый boolean. Текст — второй аргумент |
| `applyLocalWideSlots` смотрит только `activeId` | «ОК» на тексте молча ничего не пишет | T7 подставляет `text-review`, пока нет страницы фото |
| Эффект сидирования гоняет `carry` от `""` | Радио сбрасывается в «Подтвердить» | Повторный прогон при том же тексте выходит сразу, как на фото |
| Гейт шага и гейт мастера разъедутся | Кнопка на экране и переход по шагам спорят | Оба зовут `photoQueueWidePlateBlocksFinish` по `text-review` |
| Старые тесты шага ждут `Количество:` | Ложное красное, если секция включится без `wideSlots` | Без переданного `wideSlots` старый блок остаётся. Эти тесты не переписывать |
| «ОК» снова в сетке карточки | Полоса на всю ширину | T2 держит кнопку внутри flex-ряда радио |

## Open Questions

Нет. Пункты 7 и 8 спеки остаются: смена фото сворачивает список, «ОК» «Подтвердить» не переживает полную перезагрузку. Если это неверно — остановить план до кода.

## Parallelization

Не параллелить. T1–T3 — один файл секции. T5–T7 — шаг и мастер. T4 можно было бы писать рядом с T1, но один проход по номеру проще ревьюить.
