# Implementation Plan: После отказа zip «Заменить» остаётся на той же марке

**Created:** 2026-10-24  
**Spec:** [../../specs/kp-ocr-wide-plate-choice-hold.md](../../specs/kp-ocr-wide-plate-choice-hold.md)  
**Parent:** [../../specs/kp-ocr-wide-plate-queue-mismatch.md](../../specs/kp-ocr-wide-plate-queue-mismatch.md) — уже реализован, не переоткрываем  
**Goal:** После розовой ошибки «Число списков не совпадает с числом фото» выбранные «Заменить» или «Исключить» остаются на той же марке, а следующий черновик в мастере не пропадает.  
**Status:** IMPLEMENT done  
**SDD:** SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅

## Overview

Родитель уже не шлёт `update` с `mode: "replace"` перед resolve на очереди фото и пишет замену на свою страницу, когда длины совпали. Если zip всё же отказывает, `handleApplyWidePlates` выходит из `!layout.ok`, не сняв `widePlateChoicesHeld`. Пока флаг истинный, `hydrate-draft` делает `return state` и выбрасывает весь следующий черновик.

Дальше карточка читает overlay. Серверный список широких строк уже пуст, id строки становится `live-wide-0`, а выбор лежит под прежним `wide-0`. Переключатель падает в «Подтвердить». Повторное «Применить решения» видит пустой `wide_plate_lines`, выходит до resolve и снимает удержание, страницу не записав. Боковая колонка смотрит на сырой черновик: `wide_plates_resolved: true` открывает шаг клиента, пока на экране марка ещё есть.

Фикс на клиенте. Нового URL и нового поля нет. Сервер не меняем. Смысл `mode: "replace"` для текста и для «Список верен» после сброса очереди не меняем. Суффикс «п» не трогаем. Ветку `?? storeText` и неширокие правки OCR до «Список верен» не трогаем.

## Architecture Decisions

- Флаг `widePlateChoicesHeld` больше не глушит `hydrate-draft`. Следующий черновик сливается: шаг (по уже существующему `mergeWizardStepWithServer`), клиент, нормализованный текст, `lastDraft`. Явный `refreshBatchText` тоже применяется, поэтому ответ «Список верен» не выбрасывается. Автоподмена текста сверки последней партией при удержании не включается: её по-прежнему не просил вызывающий код, и экран сверки читает текст страницы.
- Пока удержание включено, `wide_plates_resolved: true` не стирает «Заменить» и «Исключить». Успешный zip по-прежнему снимает удержание и только потом делает `sync-after-wide-plates`, который очищает выбор после записи страницы. `sync-after-wide-plates` при удержании по-прежнему не затирает выбор.
- Переход с шага плит вперёд одним только `hydrate-draft` не открываем. Это правило родителя. Тест принимает шаг, который слияние и так принимает: локально уже «клиент», с сервера приходит «результат».
- Карточка и повторный resolve ищут «Заменить» / «Исключить» по марке (`normalizeLineKey`: `п`, «шт», пробелы). Прямой id строки побеждает, если менеджер переключил уже новый id. Марка без решения по-прежнему «Подтвердить». При выборе запоминаем строку и ключ марки, чтобы после смены id `wide-0` → `live-wide-0` решение находилось.
- Пустой серверный `wide_plate_lines` не считается «решать нечего», если в тексте страницы есть широкая марка и по ней уже выбран «Заменить» или «Исключить». Тогда resolve вызывается снова. Очередь фото по-прежнему не шлёт `update` replace. Если длин снова не совпало, ошибка остаётся, удержание не снимается, страница не выдумывается.
- Выбор марки снимается, когда в текстах страниц этой марки больше нет, когда менеджер переключил пункт, или когда zip записал страницу. Пустой корпус страниц не считается поводом стереть выбор.
- «Готово, далее» уже смотрит на overlay. Боковая колонка переходит на ту же сборку: `filterDraftForBatchReview` + `overlayDraftWithLiveWideLines` от текста, который виден на экране. `wide_plates_resolved: true` при старой марке на странице шаг клиента не открывает.

## Dependency graph

```
T1 красные тесты
        |
        +---- T2 hydrate сливает черновик и не стирает выбор
        |
        +---- T3 поиск по марке и повтор resolve
        |
        +---- T4 боковая колонка и снятие выбора, когда марки на страницах нет
                    |
               Checkpoint
```

T3 и T4 можно писать после T1. T2 не зависит от карточки. Общий прогон — после всех трёх.

## Task 1: Красные тесты на текущем коде

**Description:** Зафиксировать три разрыва спеки тестами, которые падают до правки, и проверить, есть ли уже тест «Готово, далее» при overlay. Разная длина партий и страниц и успешный zip родителя остаются как есть.

**Acceptance criteria:**
- [x] При удержании входящий черновик с другим шагом и `wide_plates_resolved: true` должен обновить шаг и `lastDraft`, не стирая «Заменить» по `ПБ 76-15-8 13` и четыре строки замены
- [x] Решение под `wide-0` находится у строки `live-wide-0` с той же маркой: переключатель «Заменить», поле не пустое; марка без решения — «Подтвердить»
- [x] Пустой `currentWideLines`, марка на странице, «Заменить» на ключе марки: `resolveWidePlates` вызывается с этими четырьмя строками, тихого return нет
- [x] Тест «Готово, далее» добавляется только если его ещё нет

**Verification:**
- [x] Новые тесты красные на текущем коде, старые тесты длины и успешного zip ещё зелёные

**Dependencies:** None

**Files likely touched:**
- `frontend/src/features/commercial-offer/store/wizardDraftStore.test.tsx`
- `frontend/src/features/commercial-offer/components/WidePlatesInlineSection.test.tsx`
- `frontend/src/features/commercial-offer/lib/flushThenResolveWidePlates.test.ts`
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.test.tsx`

**Estimated scope:** Small

## Task 2: Hydrate не выбрасывает черновик и не стирает выбор

**Description:** Убрать `return state` из `hydrate-draft`. Черновик сливается и при удержании. «Заменить» и «Исключить» не очищаются только потому, что у пришедшего черновика `wide_plates_resolved: true`. Без удержания прежняя очистка остаётся. `sync-after-wide-plates` при удержании по-прежнему ничего не затирает. На `!layout.ok` флаг не снимается: страница не записана.

**Acceptance criteria:**
- [x] Тест Task 1 про шаг и `lastDraft` зелёный, четыре строки на месте
- [x] Старый тест «не сбрасывает replace в confirm, пока список фото не записан» зелёный
- [x] Успешный `sync-after-wide-plates` без удержания по-прежнему очищает выбор

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/store/wizardDraftStore.test.tsx`

**Dependencies:** Task 1

**Files likely touched:**
- `frontend/src/features/commercial-offer/store/wizardDraftStore.tsx`
- `frontend/src/features/commercial-offer/types/commercialOffer.ts`

**Estimated scope:** Small

## Task 3: Поиск «Заменить» по марке и повтор resolve

**Description:** Карточка берёт решение по id, затем по ключу марки, затем по `sourceLine` с тем же `normalizeLineKey`. При переключении запоминается строка и ключ марки. `buildWidePlateResolveDecisions` при пустом серверном списке собирает «Заменить» / «Исключить» с живых строк страницы и не выходит молча. Подтверждение без решения по-прежнему не шлёт пустой resolve. Очередь фото не вызывает `update` replace.

**Acceptance criteria:**
- [x] Тест карточки зелёный: `live-wide-0` показывает «Заменить» и непустое поле, соседняя марка — «Подтвердить»
- [x] Тест flush зелёный: `resolveWidePlates` вызван с `action: "replace"` и четырьмя строками, `updateInput` не вызван
- [x] Разная длина по-прежнему `ok: false` без `pages`
- [x] Две готовые страницы, отравленный store: `update` replace не вызывается; успешный zip пишет первую страницу и не трогает вторую

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/flushThenResolveWidePlates.test.ts src/features/commercial-offer/lib/syncWidePlateQueuePages.test.ts`

**Dependencies:** Task 1

**Files likely touched:**
- `frontend/src/features/commercial-offer/lib/flushThenResolveWidePlates.ts`
- `frontend/src/features/commercial-offer/components/WidePlatesInlineSection.tsx`
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- `frontend/src/features/commercial-offer/store/wizardDraftStore.tsx`

**Estimated scope:** Medium

## Task 4: Боковая колонка как overlay, выбор гаснет вместе с маркой

**Description:** `hasUnresolvedWidePlates` в мастере считается от того же overlay, что карточка и «Готово, далее»: текст страницы на сверке, иначе нормализованный текст. Шаг клиента закрыт, пока overlay видит нерешённую широкую марку, даже если сервер уже прислал `wide_plates_resolved: true`. Когда в текстах страниц марки больше нет, соответствующий «Заменить» / «Исключить» снимается. Пустой список страниц выбор не стирает.

**Acceptance criteria:**
- [x] «Готово, далее» закрыта, пока overlay видит широкую марку при `wide_plates_resolved: true` у сырого черновика
- [x] Переход к клиенту использует тот же признак, что карточка
- [x] «Список верен» по-прежнему не зависит от широкой марки и его `hydrate-draft` не выбрасывается

**Verification:**
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/components/steps/PlateInputStep.test.tsx`
- [x] `cd /home/username/Code/plan_web/frontend && npm run typecheck`

**Dependencies:** Task 2, Task 3

**Files likely touched:**
- `frontend/src/features/commercial-offer/lib/batchReview.ts`
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- `frontend/src/features/commercial-offer/store/wizardDraftStore.tsx`

**Estimated scope:** Small

## Checkpoint: Complete

- [x] S1–S6 закрыты тестами Task 1–4
- [x] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/store/wizardDraftStore.test.tsx src/features/commercial-offer/lib/flushThenResolveWidePlates.test.ts src/features/commercial-offer/lib/syncWidePlateQueuePages.test.ts src/features/commercial-offer/components/steps/PlateInputStep.test.tsx`
- [x] `cd /home/username/Code/plan_web/frontend && npm run typecheck`
- [x] Сервер не меняется: pytest не нужен

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Слияние шага с плит сразу на клиента | Менеджер перескочит «Обработать» | `mergeWizardStepWithServer` не двигаем. Тест шага идёт с клиента на результат |
| Удержание залипает после отказа | Следующие hydrate снова пропадут | Hydrate при флаге больше не делает `return state`. Флаг только бережёт выбор и блокирует ранний sync |
| Повтор resolve при пустом серверном списке шлёт «Подтвердить» | Лишний resolve или тихий выход | Шлём только уже выбранные «Заменить» и «Исключить» |
| Текст без фото | `mode: "replace"` изменится | Ветка `!photoQueue` в flush не меняется |
| Неширокая правка OCR и `?? storeText` | Лишний diff | Эти места не открываем |

## Open Questions

Нет. Контракт `update` и `wide-plates/resolve` не меняем.
