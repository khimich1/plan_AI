# Implementation Plan: Слот широкой плиты на сверке фото

**Created:** 2026-10-26  
**Spec:** [../../specs/kp-ocr-wide-plate-slot.md](../../specs/kp-ocr-wide-plate-slot.md)  
**Idea:** [../../ideas/wide-plate-slot.md](../../ideas/wide-plate-slot.md)  
**Goal:** На сверке фото блок широкой плиты держится за строку списка. ОК и «Применить решения» пишут текст этой страницы одной функцией и не зовут `wide-plates/resolve`.  
**Status:** IMPLEMENT done.  
**SDD:** SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅

## Overview

Ширина по-прежнему считается в `liveWidePlateLines`: строка шире 12 дм. Идентичность слота — отдельное поле с стабильным `id`. Пока число строк то же, слот остаётся на том же индексе, даже если текст марки изменился. Если строк стало больше или меньше, слот переносится только по точному тексту в прежнем порядке; в одном разрыве между совпавшими строками ровно одна исчезнувшая и одна новая строка считаются правкой того же слота.

ОК и низ зовут `commitWideSlots`. «Подтвердить» не меняет текст и ставит `accepted`. «Заменить» и «Исключить» меняют текст, затем слоты пересчитываются. Пустая замена останавливает пачку: уже записанное выше остаётся, текущий слот и всё ниже в этот клик не пишутся.

Сервер на эти кнопки не вызывается. Текст без фото остаётся на `WidePlatesInlineSection` и `flushThenResolveWidePlates`. Для фото — новый блок `WidePlateSlotSection`.

Каждый шаг ниже: сначала тест из колонки RED, он падает, потому что функции ещё нет или поведение не то. Потом минимальный код, пока зелёный только этот прогон. Чужие тесты не подгонять.

## Architecture Decisions

- Детектор ширины не копировать. Слот строится только для индексов, которые `liveWidePlateLines` уже считает широкими. `id` вида `live-wide-0` слотом не является.
- Слоты фото лежат в `wizardDraftStore` как `wideSlotsByPage: Record<pageId, WideSlot[]>`. Полная перезагрузка черновика их не восстанавливает (пункт 7 спеки).
- `widePlateActions`, `normalizeLineKey` и `pruneWidePlateActions` для фото не читаются. Текстовый путь их не теряет.
- Поле замены заполняется один раз в `chooseWideSlotAction`, когда действие становится `replace`, а поле пустое. `carryWideSlots` поле не переписывает.
- «Готово, далее» на очереди фото смотрит только на открытые слоты (`accepted === false`). Флаг `wide_plates_resolved` эту кнопку на очереди фото больше не держит: иначе «Подтвердить» не откроет шаг, марка в тексте осталась, сервер по-прежнему `resolved: false`.
- Старые `undecidedWideKeys` / `photoQueueHasUndecidedWidePlate` не удалять и не переписывать. Фото просто перестаёт их звать.

## Dependency graph

```
T1 правка на том же индексе
  → T2 вставка выше
  → T3 две одинаковые марки
  → T4 слот гаснет, если уже не шире 12
  → T5 ОК снимается, если строку снова сделали широкой другой
        → T6–T9 запись одной строки
        → T10 пачка останавливается на ошибке
        → T11 подсказка один раз
              → T12 две страницы в store
              → T13 гейт «Готово, далее»
                    → T14–T15 новый блок UI
                    → T16 мастер: локальная запись, без resolve
                    → T17 кнопка шага
                          → Checkpoint
```

T6 можно писать сразу после T1: запись не зависит от переноса. В одном сеансе всё равно идти по номеру: так проще ревьюить красный тест.

## Task 1: Правка марки остаётся на том же индексе

**Description:** RED. В новом `widePlateSlots.test.ts` тест `keeps the slot when the only change is the text of that line`. Две строки, слот на индексе 1 с `action: "replace"`, своим `id` и непустым `replacementText`. Новый текст той же длины, вторая строка другая и всё ещё шире 12 дм. `carryWideSlots` возвращает тот же `id`, тот же action, тот же `replacementText`, `lineIndex === 1`.

**Acceptance criteria:**
- [x] Тест красный до функции, зелёный после
- [x] `id` не равен `live-wide-*`

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/widePlateSlots.test.ts`

**Dependencies:** None

**Files:**
- `frontend/src/features/commercial-offer/lib/widePlateSlots.ts`
- `frontend/src/features/commercial-offer/lib/widePlateSlots.test.ts`

**Scope:** Small

## Task 2: Вставка строки выше не крадёт слот

**Description:** RED. Тест `keeps the slot on the same plate when a line is inserted above`. Было `["ПБ 10-12-8 1", "ПБ 76-15-8 13"]`, слот на второй. Стало три строки: новая сверху, две прежние без изменений. Слот на индексе 2, тот же `id`.

**Acceptance criteria:**
- [x] Индекс слота сдвинулся вместе со строкой, а не остался на 1

**Verification:**
- [x] тот же прогон `widePlateSlots.test.ts`

**Dependencies:** Task 1

**Files:** те же два

**Scope:** Small

## Task 3: Две одинаковые марки — два слота

**Description:** RED. Тест `edits the first of two identical marks and leaves the second slot`. Две строки `ПБ 76-15-8 13`, два слота, у первого `replacementText` «первая», у второго «вторая». В новом тексте первая строка `ПБ 68-15-8 13`, вторая та же. Первый слот на индексе 0 с «первая», второй на индексе 1 с «вторая».

**Acceptance criteria:**
- [x] Слоты не склеиваются по тексту марки

**Verification:**
- [x] тот же прогон

**Dependencies:** Task 2

**Files:** те же два

**Scope:** Small

## Task 4: Узкая или удалённая строка забирает слот

**Description:** RED. Два теста. `drops the slot when the line is no longer wider than 12 dm`: та же длина, строка стала `ПБ 76-12-8 13`, слотов нет, в тексте нет подсказки замены. `drops the slot when its line is deleted`: строка исчезла, соседние строки на месте, слота нет.

**Acceptance criteria:**
- [x] Функция переноса только возвращает слоты. Она не дописывает `replacementText` в текст

**Verification:**
- [x] тот же прогон

**Dependencies:** Task 3

**Files:** те же два

**Scope:** Small

## Task 5: Повторная широкая правка снимает ОК

**Description:** RED. Тест `clears accepted when a confirmed line is edited into another wide mark`. Слот `accepted: true`, `action: "confirm"`, `replacementText` непустой. Та же длина, текст строки другой и шире 12 дм. Тот же `id`, `accepted: false`, action и поле прежние. Второй тест `keeps accepted when the confirmed line text is unchanged`.

**Acceptance criteria:**
- [x] Неизменённый текст не снимает ОК
- [x] Смена текста на том же индексе снимает только `accepted`

**Verification:**
- [x] тот же прогон

**Dependencies:** Task 4

**Files:** те же два

**Scope:** Small

## Checkpoint: After Tasks 1–5

- [x] `widePlateSlots.test.ts` зелёный
- [x] Перенос не пишет в текст и не смотрит на `normalizeLineKey`
- [x] Ревью правила разрыва: не одна старая и одна новая строка между совпавшими — старый слот гаснет, новые широкие получают новые id

## Task 6: «Заменить» подставляет строки поля

**Description:** RED. Тест `replaces that line with the non-empty field lines`. Индекс 1, поле из двух строк. В результате на этом месте две строки, исходной марки нет, строка 0 не изменилась.

**Acceptance criteria:**
- [x] Пустые строки поля отбрасываются
- [x] Соседи не сдвигаются иначе как из-за вставки замены

**Verification:**
- [x] тот же прогон

**Dependencies:** Task 1

**Files:** те же два

**Scope:** Small

## Task 7: «Исключить» удаляет строку по индексу

**Description:** RED. Тест `removes only the excluded line`. Три строки, исключается средняя. Края на месте.

**Acceptance criteria:**
- [x] Удаление по индексу, не по тексту марки

**Verification:**
- [x] тот же прогон

**Dependencies:** Task 6

**Files:** те же два

**Scope:** Small

## Task 8: «Подтвердить» не меняет текст

**Description:** RED. Тест `confirm returns the same lines`. `applyWideSlot` для `confirm` отдаёт тот же массив строк.

**Acceptance criteria:**
- [x] Текст побайтово тот же, включая соседние строки

**Verification:**
- [x] тот же прогон

**Dependencies:** Task 7

**Files:** те же два

**Scope:** Small

## Task 9: Пустая замена — отказ

**Description:** RED. Тест `rejects an empty replacement and leaves the lines`. Поле `""` и поле из пробелов и пустых строк. `{ ok: false, error: "Укажите заменяющие позиции." }`, строки те же.

**Acceptance criteria:**
- [x] Текст ошибки именно этот

**Verification:**
- [x] тот же прогон

**Dependencies:** Task 8

**Files:** те же два

**Scope:** Small

## Task 10: Пачка останавливается на первом отказе

**Description:** RED. Тест `commits slots in order and stops at the first empty replace`. Три слота: exclude, replace с пустым полем, confirm. `commitWideSlots` удаляет первую строку, вторую не трогает, третью не помечает `accepted`, `failedId` — id второго, `error` — та же строка. Успешный exclude уже в тексте.

**Acceptance criteria:**
- [x] Ниже отказа в этот вызов записи нет
- [x] Верхний успех не откатывается

**Verification:**
- [x] тот же прогон

**Dependencies:** Task 9

**Files:** те же два

**Scope:** Small

## Task 11: Подсказка заполняется один раз

**Description:** RED. Тест `fills the suggestion only when replace is chosen and the field is empty`. `chooseWideSlotAction` на пустом поле даёт текст `buildAutoSplitSuggestion` текущей строки. Второй тест `does not rewrite a field that already has text when replace is chosen again`. Третий: после `carryWideSlots` с новой маркой поле остаётся старым (это уже следствие Task 1, здесь отдельный тест `carry does not refresh the replacement field`).

**Acceptance criteria:**
- [x] Повторный выбор «Заменить» не затирает ручной текст
- [x] Перенос слота не зовёт подсказку

**Verification:**
- [x] тот же прогон

**Dependencies:** Task 5

**Files:** те же два, `widePlateSuggestion.ts` только читается

**Scope:** Small

## Checkpoint: After Tasks 6–11

- [x] `widePlateSlots.test.ts` зелёный целиком
- [x] В модуле нет импорта `flushThenResolveWidePlates` и нет `fetch`

## Task 12: Две страницы не делят слоты

**Description:** RED в `wizardDraftStore.test.tsx`. Тест `stores wide slots per page`. Диспатч `set-page-wide-slots` для страницы A, затем для B. В состоянии у A свои слоты, у B свои. Повторный диспатч A не очищает B.

Слоты сессионные: отдельный тест не требуется, гидратация черновика их не заполняет. Если сегодняшний `hydrate-draft` затирает неизвестные поля, тест `hydrate keeps wide slots already on the page` фиксирует, что поле переживает hydrate. Если нет — тест красный, и reducer при hydrate сохраняет `wideSlotsByPage`.

**Acceptance criteria:**
- [x] Ключ — id страницы очереди
- [x] Hydrate не обнуляет уже набранные слоты

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/store/wizardDraftStore.test.tsx`

**Dependencies:** Task 5

**Files:**
- `frontend/src/features/commercial-offer/store/wizardDraftStore.tsx`
- `frontend/src/features/commercial-offer/store/wizardDraftStore.test.tsx`
- `frontend/src/features/commercial-offer/types/commercialOffer.ts` — только тип состояния, если слоты лежат в нём

**Scope:** Small

## Task 13: Гейт шага смотрит на открытые слоты

**Description:** RED в `widePlateSlots.test.ts`. `photoQueueWidePlateBlocksFinish`. Страница с широкой строкой и слотом `accepted: false` — блок. Тот же текст, слот `accepted: true` — не блок. Вторая страница с открытым слотом — блок, даже если первая закрыта. Страница без слотов, но с широкой строкой в тексте — блок: слоты ещё не посеяны, широкая строка считается открытой. Текст без широких строк — не блок.

**Acceptance criteria:**
- [x] Подтверждённая широкая марка, которая осталась в тексте, шаг не держит
- [x] Функция не читает `wide_plates_resolved`

**Verification:**
- [x] тот же прогон `widePlateSlots.test.ts`

**Dependencies:** Task 5

**Files:** те же `widePlateSlots.ts` и тест

**Scope:** Small

## Checkpoint: After Tasks 12–13

- [x] Оба прогона зелёные
- [x] `widePlatePageGate.test.ts` не менялся и остаётся зелёным: `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/widePlatePageGate.test.ts`

## Task 14: Блок на слот и кнопка ОК

**Description:** Новый `WidePlateSlotSection`. RED в `WidePlateSlotSection.test.tsx`. Тест `shows one block per open slot and hides an accepted slot`. Два слота, второй `accepted`. На экране одна марка — текст первого слота. Радио «Заменить» отмечено, если action такой. Поле «Заменяющие позиции» видно только у «Заменить». У каждого открытого блока есть кнопка «ОК». Тест `hides the section when every slot is accepted`.

Старый `WidePlatesInlineSection.test.tsx` не трогать.

**Acceptance criteria:**
- [x] Принятый слот не рисуется
- [x] Поле замены не показано для «Подтвердить» и «Исключить»

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/components/WidePlateSlotSection.test.tsx`

**Dependencies:** Task 11

**Files:**
- `frontend/src/features/commercial-offer/components/WidePlateSlotSection.tsx`
- `frontend/src/features/commercial-offer/components/WidePlateSlotSection.test.tsx`

**Scope:** Small

## Task 15: ОК и низ сообщают id, а не пишут текст сами

**Description:** RED. Тест `OK reports that slot id`. Клик «ОК» на блоке зовёт `onAccept` с id этого слота один раз. Тест `bottom button accepts every open slot id in order`. «Применить решения» зовёт `onApply` со списком id открытых слотов сверху вниз. Смена радио зовёт `onAction` с id и новым действием. Правка поля зовёт `onReplacement` с id и новым текстом.

**Acceptance criteria:**
- [x] Секция не вызывает сеть и не меняет текст списка
- [x] Низ не подменяет действие на confirm внутри секции: наружу уходят id, действие уже лежит на слоте

**Verification:**
- [x] тот же прогон секции

**Dependencies:** Task 14

**Files:** те же компонент и тест

**Scope:** Small

## Task 16: Мастер пишет активную страницу и не зовёт resolve

**Description:** Вынести из `CommercialOfferWizard` развилку в чистую функцию `routeWidePlateApply(photoQueue: boolean): "local" | "resolve"` рядом со слотами, либо прямо в `widePlateSlots.ts`. RED. Тест `photo queue applies locally`. `true` → `"local"`. `false` → `"resolve"`.

В мастере ветка `local` вызывает `commitWideSlots` для текста активной страницы, пишет результат в `multiPage.updatePageText` и `set-batch-review-text`, кладёт новые слоты через `set-page-wide-slots` и выходит до `flushThenResolveWidePlates`. Ветка `resolve` — сегодняшний код без изменений.

Правка списка (`onBatchReviewTextChange`) для активной страницы фото вызывает `carryWideSlots` и снова кладёт слоты в store. Чужие страницы этот диспатч не получает.

**Acceptance criteria:**
- [x] Тест развилки зелёный
- [x] У `commitWideSlots` нет аргумента resolve
- [x] При `photoQueue === false` тело `handleApplyWidePlates` по-прежнему доходит до `flushThenResolveWidePlates`

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/widePlateSlots.test.ts`
- [x] `cd /home/username/Code/plan_web/frontend && npm run typecheck`

**Dependencies:** Tasks 10, 12, 15

**Files:**
- `frontend/src/features/commercial-offer/lib/widePlateSlots.ts`
- `frontend/src/features/commercial-offer/lib/widePlateSlots.test.ts`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx` — на сверке фото рисовать `WidePlateSlotSection`, иначе старую секцию

**Scope:** Medium (4 файла, одна развилка)

## Task 17: «Готово, далее» отпускает закрытые слоты

**Description:** RED, если в `PlateInputStep.test.tsx` уже есть кейс широкой плиты на сверке. Иначе RED в `widePlateSlots.test.ts` уже закрыт Task 13, а этот шаг только проводит результат в кнопку.

`hasUnresolvedWidePlates` на очереди фото равен `photoQueueWidePlateBlocksFinish` и не прибавляет `!wide_plates_resolved`. Вне очереди фото условие старое: живые `wide_plate_lines` и `wide_plates_resolved`.

Тест шага, если он есть и ломается: подтверждённые слоты при `wide_plates_resolved: false` и широкой строке в тексте не держат кнопку из-за широкой плиты. Не гасить кнопку по другим гейтам (заводской рез, без цены, «Список верен»).

**Acceptance criteria:**
- [x] Очередь фото с всеми `accepted: true` не даёт причину «Сначала примите решение по позициям шире стандартной»
- [x] Очередь без слотов и с широкой строкой эту причину даёт
- [x] Текстовый шаг без фото по-прежнему смотрит на `wide_plates_resolved`

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/components/steps/PlateInputStep.test.tsx src/features/commercial-offer/lib/widePlateSlots.test.ts`
- [x] `cd /home/username/Code/plan_web/frontend && npm run typecheck`

**Dependencies:** Tasks 13, 16

**Files:**
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.test.tsx` — только если текущий тест кнопки противоречит новому гейту

**Scope:** Small

## Checkpoint: Complete

- [x] Tasks 1–17 зелёные своими командами
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run` — 163 файла / 1360 тестов зелёные. Два чужих таймаута (`PileCostPanel`, `ShipmentDrawer`) падают только в полном прогоне и проходят отдельно; к слотам не относятся
- [x] `cd /home/username/Code/plan_web/frontend && npm run typecheck`
- [x] `flushThenResolveWidePlates.test.ts` и `widePlatePageGate.test.ts` зелёные без правки их ожиданий
- [ ] На сверке фото в сетевом журнале нет `wide-plates/resolve` после ОК. Это ручная проверка на уже запущенном `run+logs.sh`, не новый e2e

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Оверлей `liveWideDraft` снова рисует карточку из `wide_plate_lines`, хотя слот уже `accepted` | High | Фото рисует только `WidePlateSlotSection` из слотов store. Старая секция на очереди фото не монтируется |
| `wide_plates_resolved: false` держит «Готово, далее» после «Подтвердить» | High | Task 17. На очереди фото этот флаг в условие кнопки не входит |
| Разрыв «удалили одну и вставили две» гасит слот | Med | Так и задумано спекой. Отдельный тест не обязателен: в Checkpoint после T5 ревьюер смотрит правило. Не улучшать эвристику в этом плане |
| Hydrate обнуляет слоты при каждом ответе сервера | High | Task 12, тест на hydrate |
| Правка текста идёт мимо `onBatchReviewTextChange` | Med | Task 16 вешает `carryWideSlots` на этот единственный колбэк сверки |

## Open Questions

Пункт 7 спеки остаётся: ОК «Подтвердить» не переживает полную перезагрузку. План под это и сделан. Если подтверждение надо хранить на сервере — остановить план до нового решения, endpoint по-прежнему не добавлять молча.

## Parallelization

Последовательно. T6–T9 можно вести тем же файлом сразу за T1, но не параллельно с T2–T5: оба пишут `widePlateSlots.ts`.
