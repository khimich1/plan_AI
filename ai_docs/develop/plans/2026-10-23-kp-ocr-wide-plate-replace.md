# Implementation Plan: Замена широкой плиты после OCR

**Created:** 2026-10-23  
**Spec:** [../../specs/kp-ocr-wide-plate-replace.md](../../specs/kp-ocr-wide-plate-replace.md)  
**Goal:** «Применить решения» на сверке фото заменяет широкую плиту и записывает новый список в ту же страницу очереди, где лежит снимок.  
**Status:** IMPLEMENT done  
**SDD:** SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅

## Overview

Суффикс нагрузки не меняем: до кнопки у фото остаётся текст распознавания. Парсер по-прежнему может держать канон с `п` внутри черновика. Фикс в двух местах. Сначала решение «Заменить» должно доехать до сервера, даже если на экране `ПБ 76-15-8 13`, а в `wide_plate_lines` — `ПБ 76-15-8п 13`. Затем ответ сервера записывается в `PageSource.batchReviewText` этой страницы, а не только в store мастера. Store повторяет текст активной страницы.

Свежая сессия фото создаёт партии в том же порядке, что и страницы очереди: первый снимок — `create`, следующие — `append`. После resolve `metadata.plate_batches[i].normalized_text` — новый список страницы `i`.

## Architecture Decisions

- Не вызывать `_align_text_batch_normalized` для `image`. Тест `test_align_text_batch_normalized_uses_parser_text` остаётся как есть.
- Ключ склейки на клиенте доводим до того, что сервер уже делает в `_normalize_wide_line_key`: без регистра, без префикса «Плиты» / «ПБ», без суффикса `п` у нагрузки, без «шт» и без количества на конце. Суффикс в видимом тексте не добавляем и не снимаем.
- Строка очереди без выбора по-прежнему уходит как `confirm`. Выбор `replace` или `exclude`, который ни с одной серверной строкой не склеился, не подменяется на `confirm`: resolve не вызывается, карточка показывает ошибку.
- Текст страницы после успеха берётся из переписанной партии, не из последнего batch и не из склеенного `normalized_text` всего черновика. Так «Список верен» отправляет уже применённый список и не возвращает широкую марку.
- Текстовый ввод без `multiPage.hasStarted` не трогаем: store по-прежнему обновляет `sync-after-wide-plates`.

## Dependency graph

```
P1 backend-тест image replace
P2 склейка решения на клиенте
        \                 /
         \               /
          P3 текст партий → страницы очереди
                    |
          P4 запись в страницу из мастера
                    |
               Checkpoint
```

P1 и P2 независимы. P3 не зависит от их кода, только от формы `plate_batches`. P4 вызывает функцию из P3.

## Task 1: Сервер принимает замену OCR-строки без `п`

**Description:** Тест на `resolve_wide_plates`: партия `source_type: image`, в её `normalized_text` и в `normalized_lines` строка `ПБ 76-15-8 13`, в `wide_plate_lines` — `ПБ 76-15-8п 13`. Решение `replace` с текстом `ПБ 76-12-8 13` и `ПБ 76-3.0-8 13`. Отдельный кейс `exclude` по той же паре. Сопоставление уже идёт через `_normalize_wide_line_key`, поэтому тест может быть зелёным без правки кода. Код `commercial_plate_resolve.py` менять только если тест красный.

**Acceptance criteria:**
- [ ] В сохранённом тексте и в `plate_batches[0].normalized_text` есть обе замены и нет `76-15-8`
- [ ] `exclude` удаляет марку и из `input_text`, и из текста image-партии
- [ ] `test_align_text_batch_normalized_uses_parser_text` зелёный без правок этой функции

**Verification:**
- [ ] `/home/username/Code/plan_web/.venv/bin/python -m pytest tests/test_commercial_wide_plates_resolve.py -q`

**Dependencies:** None

**Files likely touched:**
- `tests/test_commercial_wide_plates_resolve.py`
- `app/services/commercial_plate_resolve.py` — только если новый тест красный

**Estimated scope:** Small

## Task 2: Выбор «Заменить» не становится «Подтвердить»

**Description:** `normalizeLineKey` в `flushThenResolveWidePlates.ts` сравнивает марки так же, как серверный `_normalize_wide_line_key`: снимает `п`, «шт» и хвостовое количество, не меняя исходную строку. `buildWidePlateResolveDecisions` по этому ключу переносит `replace` с `live-wide-*` на id сервера. Если у live-строки действие `replace` или `exclude`, а серверной строки с тем же ключом нет, функция сообщает о промахе. `flushThenResolveWidePlates` в этом случае бросает ошибку до `resolveWidePlates`. Строка без выбора по-прежнему `confirm`.

**Acceptance criteria:**
- [ ] `ПБ 76-15-8 13` на `live-wide-0` и flushed `ПБ 76-15-8п 13` дают `action: "replace"` и тот же `replacementText`
- [ ] Строка с «шт» склеивается с каноном без «шт»
- [ ] Несведённый `replace` не уходит как `confirm`, resolve не вызывается
- [ ] Существующий кейс `34-15-10п 15` сохраняет действие

**Verification:**
- [ ] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/flushThenResolveWidePlates.test.ts`

**Dependencies:** None

**Files likely touched:**
- `frontend/src/features/commercial-offer/lib/flushThenResolveWidePlates.ts`
- `frontend/src/features/commercial-offer/lib/flushThenResolveWidePlates.test.ts`

**Estimated scope:** Small

## Checkpoint: After Tasks 1–2

- [ ] Оба целевых прогона зелёные
- [ ] Суффикс в тексте распознавания нигде не дописывается
- [ ] Ревью склейки до записи в очередь

## Task 3: Новый текст партии раскладывается по страницам очереди

**Description:** Чистая функция рядом с `applyAiSession.ts`. Вход: страницы `{ id, batchReviewText }`, id активной, черновик после resolve. Если число `metadata.plate_batches` равно числу страниц, страница `i` получает `plate_batches[i].normalized_text`. Иначе функция не угадывает порядок и возвращает отказ. Активной странице в результате отдаётся её новый текст, чтобы store не остался текстом последней партии.

**Acceptance criteria:**
- [ ] Две страницы и две партии: у каждой страницы текст своей партии
- [ ] Текст неактивной страницы не подменяется текстом активной
- [ ] Разная длина списков → отказ, тексты страниц не придуманы

**Verification:**
- [ ] `cd /home/username/Code/plan_web/frontend && npm test -- --run src/features/commercial-offer/lib/applyAiSession.test.ts`  
      Новый файл теста функции допустим рядом, если в `applyAiSession.test.ts` ей не место. Прогон — по фактическому пути теста.

**Dependencies:** None

**Files likely touched:**
- `frontend/src/features/commercial-offer/lib/applyAiSession.ts` или новый файл в том же `lib/`
- тест этой функции

**Estimated scope:** Small

## Task 4: Мастер пишет результат в очередь фото

**Description:** В `handleApplyWidePlates` после успешного `mutateAsync` разложить партии по страницам функцией из Task 3 и для каждой вызвать `multiPage.updatePageText`. Store активной страницы выставить тем же текстом через `set-batch-review-text`, чтобы он не остался последней партией из `sync-after-wide-plates`. Отказ раскладки (разная длина) показать в `setWidePlateError`, страницы очереди не затирать. Без сессии фото (`hasStarted === false`) очередь не трогать. Ошибка склейки из Task 2 уже попадает в `setWidePlateError` через существующий `catch`.

**Acceptance criteria:**
- [ ] После успеха `updatePageText` вызван для каждой страницы её новым текстом
- [ ] Store активной страницы равен её `batchReviewText`, а не тексту последней партии, если страницы две
- [ ] Текстовый путь без фото не вызывает `updatePageText`
- [ ] Ошибка resolve или промах склейки не меняет тексты страниц

**Verification:**
- [ ] Unit на функцию из Task 3 плюс точечный тест мастера, если в `CommercialOfferWizard` уже есть тест apply; иначе проверка ограничена чистой функцией и типом вызова в мастере
- [ ] `cd /home/username/Code/plan_web/frontend && npm run typecheck`

**Dependencies:** Task 2, Task 3

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- тест мастера только если он уже покрывает `handleApplyWidePlates`

**Estimated scope:** Medium (3 файла с Task 3)

## Checkpoint: Complete

- [ ] Сценарий спеки S1–S7 закрыт тестами Task 1–4
- [ ] `pytest tests/test_commercial_wide_plates_resolve.py -q` зелёный
- [ ] vitest файлов склейки и раскладки очереди зелёный
- [ ] `npm run typecheck` в `frontend/` зелёный
- [ ] Реализация не начата, пока этот план не принят

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Число партий не равно числу страниц (к сессии примешали текстовый добор) | Высокое: неверный список у чужого фото | Функция раскладки отказывается. Мастер показывает ошибку и не пишет в очередь |
| `sync-after-wide-plates` кладёт в store текст последней партии | Среднее: «Список верен» затрёт активную страницу текстом последней | Сразу после await выставить store текстом активной страницы |
| Серверный тест уже зелёный, и правка resolve «заодно» ломает invalid width | Среднее | В Task 1 не менять `_match_plate_resolve_item_to_line`, если тест прошёл |
| Ключ на клиенте станет строже и перестанет находить старые live-id | Среднее | Регрессия `34-15-10п 15` и кейс exclude без `п` в том же файле тестов |

## Open Questions

Нет. Суффикс нагрузки оставлен как есть. Реализация после ревью этого плана.
