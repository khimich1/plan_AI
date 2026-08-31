# Spec: Несколько скриншотов страниц на шаге ввода КП

**Статус**: SPECIFY ✅ · PLAN ✅ · IMPLEMENT Phase A MVP ✅ · Phase A.1 ✅ · Phase A.2 ✅ · Phase A.3 ✅ · Phase A.4 ✅ · **REVIEW 2026-08-31 → Approve** (R12–R13 closed; companion R10–R11 closed)  
**Дата**: 2026-08-31  
**One-pager**: [ai_docs/ideas/kp-multi-page-screenshots.md](../ideas/kp-multi-page-screenshots.md)  
**Delta (busy + lightbox)**: [ai_docs/ideas/kp-multi-page-lightbox-and-busy-fix.md](../ideas/kp-multi-page-lightbox-and-busy-fix.md)  
**Delta (wait + AI on review)**: [ai_docs/ideas/kp-ocr-wait-and-ai-on-review.md](../ideas/kp-ocr-wait-and-ai-on-review.md) · [kp-ocr-wait-and-ai-on-review.md](./kp-ocr-wait-and-ai-on-review.md)  
**План**: [ai_docs/develop/plans/2026-08-31-kp-multi-page-screenshots.md](../develop/plans/2026-08-31-kp-multi-page-screenshots.md)  
**Связано**: [ux-wizard-step-plates.md](./ux-wizard-step-plates.md), OCR-пайплайн (`ocr-verify-apply-and-upscale`, `ocr-small-screenshot-verify`), [kp-multi-nomenclature-append.md](./kp-multi-nomenclature-append.md) (не путать: там другие номенклатуры / заходы)

## Objective

**Проблема.** Большое КП приходит несколькими страницами (скрин/фото листа). Сейчас на шаге источника можно выбрать **один** файл; Alert «Выбран файл: …»; OCR и сверка заточены под один кадр. Менеджер либо гоняет N циклов, либо теряет контроль «какая страница → какой кусок списка».

**Цель.** Набрать несколько страниц до «Распознать», видеть превью с удалением, после старта **сразу работать с первой готовой** страницей (фото + её список переключаются вместе), не теряя per-image OCR-ухищрения. Append-режим («Добавить к списку») в MVP почти не меняем.

**Пользователь:** менеджер, который собирает КП из многостраничной таблицы / чертежа на шаге ввода изделий.

**Успех:** 3+ скрина → один клик «Распознать» → правка страницы 1, пока 2…N ещё в очереди; навигация не смешивает чужие тексты с чужими фото; итоговый draft = подтверждённые страницы по порядку.

---

## ASSUMPTIONS (locked 2026-08-31)

1. **Scope изделий — все типы с `SourceInputCard`** (плиты, сваи, марши, ступени, мостовые, ФБС). Общая карточка источника и общий page-state в wizard; шаги не копируют галерею шесть раз.
2. **Только фото-путь multi-page.** Текст по-прежнему один textarea; выбор/добавление картинок очищает текст (как сейчас файл чистит текст). Смесь «текст + пачка картинок» в одном заходе — out.
3. **Порядок страниц = порядок добавления** (file multi-select / последовательный paste). Drag-reorder в MVP нет.
4. **Старт OCR только по «Распознать».** Drop/paste только кладёт в галерею (`pending`), без вызова API.
5. **MVP = фронтовый конвейер (направление 2):** после «Распознать» фронт последовательно вызывает **существующий** single-image create/update OCR endpoint на каждый файл. Первый `ready` сразу включает режим сверки по этой странице. Server job / Redis (направление 1) — фаза 2, тот же UX-контракт страниц.
6. **Каждый файл = полный текущий OCR-пайплайн** (preprocess, extract, verify, upscale-policy и т.д.). Не multi-image vision, не склейка кадров.
7. **Модель UI: массив страниц**, не один `selectedImageName` / один `recognizedImageUrl`. Активная страница задаёт большое фото **и** редактируемый список этой страницы.
8. **«Список верен» — по текущей готовой странице.** Подтверждение фиксирует текст страницы и переключает на следующую `ready` (или ждёт, если хвост ещё `running`). Выход из batch-review / сборка в draft — когда **все** страницы подтверждены (или пользователь явно завершил — см. D5). Пока есть `pending`/`running`/`error`, «Готово, далее» по wizard как сейчас не пускаем дальше неподтверждённого батча.
9. **Первая страница create draft (`replace`), остальные — append того же product_type** внутренним конвейером (не путать с UI «Добавить к списку» / сменой номенклатуры). Снаружи — один заход multi-page. Существующий UI append / AI-инструкции **не перепроектируем**; multi-file input в карточке append можно не включать в MVP.
10. **Лимит страниц:** soft-cap **12** (константа на фронте + отказ с понятным сообщением). При необходимости поднять без смены модели.
11. **Ошибка / плохое фото:** конвейер **не валит** весь батч. Плохую страницу **удаляют** (✕); хорошее фото **добавляют в конец** очереди. Остальные страницы и их статусы/тексты не трогаем. In-place «заменить слот» нет.
12. **Закрытие вкладки mid-batch** в MVP допустимо теряет незавершённый хвост (нет job store). Предупреждение `beforeunload`, если есть `running`/`pending` после старта.
13. **✕ удаление:** разрешено для `pending` | `error` | `ready` (ещё не confirmed). Запрещено для `running` и `confirmed`. После confirm — только правка текста или «Начать заново».
14. **Новый файл после старта батча** всегда `pending` в хвосте; последовательный runner подхватывает хвост (если уже idle — стартует OCR только для новых `pending`, без пересоздания draft с нуля).
15. **Коммиты агент не делает**, пока явно не попросите.

→ A1–A15 + Q1–Q4 approved 2026-08-31.

---

## Decisions locked (из ideation 2026-08-31)

| # | Тема | Решение |
|---|------|---------|
| D1 | Результат UX | Скрин и текст **переключаются вместе** (не один общий список) |
| D2 | Прогрессивность | Можно работать с первой `ready`, не ждать весь батч |
| D3 | OCR | Per-image существующий пайплайн; не multi-image vision |
| D4 | Архитектура MVP | Фронтовый последовательный конвейер; server job — фаза 2 |
| D5 | Append UI | Почти не трогаем; multi-page на первом replace-заходе |
| D6 | Старт распознавания | Только по кнопке «Распознать» |
| D7 | Confirm → next | Автопереход на следующую `ready` (Q1) |
| D8 | Навигация | Свободно по любой `ready`; confirm по одной (Q2) |
| D9 | Soft-cap | 12 страниц (Q3) |
| D10 | Плохое фото | ✕ удалить + добавить хорошее **в конец** очереди; остальные не трогаем (Q4) |
| D11 | Прогресс в UI | Текст «Распознано k/n» в шапке сверки |
| D12 | API MVP | Первый файл — существующий create/replace; остальные — существующий append; facade на фронте |

---

## User Stories

- Как **менеджер**, я вставляю 5 скринов страниц, вижу 5 превью, крестиком убираю лишний, жму «Распознать».
- Как **менеджер**, когда готова 1-я страница, я сразу сверяю её фото со списком, не глядя на спиннер остальных.
- Как **менеджер**, стрелками / кликом по превью переключаю страницу: меняются и фото, и текст.
- Как **менеджер**, на каждой странице получаю то же качество OCR, что на одиночном фото.
- Как **менеджер** в режиме «Добавить к списку», я по-прежнему работаю как сейчас (один файл / текст) — multi-page туда не обязателен.

---

## Success Criteria

| # | Критерий | Метод |
|---|----------|--------|
| S1 | Можно добавить ≥2 изображения до recognize; превью с ✕; Alert имени файла нет | RTL |
| S2 | До «Распознать» API OCR не вызывается | unit / mock API |
| S3 | После ready страницы 1 UI сверки доступен, пока 2…N ещё running | RTL + mock delayed OCR |
| S4 | Смена активной страницы меняет и `img`, и текст редактора | RTL |
| S5 | Каждый файл уходит отдельным вызовом существующего OCR path (не новый multi-image endpoint в MVP) | API mock assertions |
| S6 | Confirm страницы фиксирует её текст; итоговый draft содержит позиции всех confirmed в порядке | integration / flow test |
| S7 | Ошибка на странице k не отменяет k+1 | unit конвейера |
| S8 | Soft-cap: 13-й файл не добавляется, есть сообщение | RTL |
| S9 | ✕ на `error`/`ready`/`pending` убирает только эту страницу; add нового файла → хвост; статусы остальных без изменений | unit + RTL |
| S10 | Текстовый путь и lint-гейт без регрессий | существующие SourceInputCard / lint tests green |
| S11 | `npm run test`, `npm run typecheck`, релевантные pytest commercial flows — green | CI / локально |

---

## Tech Stack

| Слой | Технология |
|------|------------|
| Frontend | React 19, TS, TanStack Query / существующие commercial API helpers, vitest |
| Backend MVP | Без нового job API; существующие create/update draft + OCR multipart |
| Backend фаза 2 (позже) | Optional: job id + status polling; тот же OCR service per file |
| OCR | `core/ocr/pipeline.py` + commercial draft services — **reuse as-is per file** |

---

## Commands

```bash
# Frontend
cd frontend && npm run test -- src/features/commercial-offer
cd frontend && npm run typecheck

# Backend (регрессия commercial; новых endpoint в MVP может не быть)
pytest tests/ -q -k "commercial"
# точечно, когда появятся тесты конвейера / контракта:
# pytest tests/test_commercial_web_flow.py -q

# Dev
./run_local.sh
```

Ручная проверка: мастер КП → источник → вставить 3 скрина → превью + ✕ → Распознать → править стр.1 при спиннере на 2–3 → ←/→ → Список верен по страницам → состав КП полный.

---

## Project Structure (ориентир MVP)

```
frontend/src/features/commercial-offer/
  components/SourceInputCard.tsx          → галерея превью вместо Alert; multi file/paste
  components/SourceImageGallery.tsx       → НОВЫЙ: thumbnails + ✕ + active
  components/steps/*InputStep.tsx         → навигация страниц в режиме сверки (тонко)
  hooks/useMultiPageRecognize.ts          → НОВЫЙ: очередь, статусы, sequential OCR
  lib/multiPageSource.ts                  → типы PageSource { id, file, previewUrl, status, text, imageUrl… }
  api/commercialOfferApi.ts               → без multi-image endpoint в MVP; N вызовов
  pages/ или wizard state                 → заменить scalar selectedImage* на pages[]

# Фаза 2 (не в MVP, задел в типах/комментариях ок):
app/api/v1/endpoints/commercial.py        → optional job create/status
app/services/commercial_ocr_job.py        → sequential server queue

ai_docs/ideas/kp-multi-page-screenshots.md
ai_docs/specs/kp-multi-page-screenshots.md
```

Не трогаем в MVP: AI-instruction UX, multi-nomenclature append loop, Telegram, смена OCR-промптов/verify policy.

---

## Code Style

- Страница — явная сущность (`PageSource`), не параллельные массивы `files[]` + `texts[]` без связи.
- Конвейер изолирован в хуке: UI только диспатчит `start`, `confirm`, `setActive`, `remove`, `addPages`.
- Первый OCR — `replace`/`create`; последующие в том же заходе — внутренний `append` API **без** раскрытия «Дополнительно» в UI.
- Плохое фото: `remove(id)` + `addPages([file])` в хвост — не `replaceFile(id)`.
- Object URLs: `revoke` при remove/unmount.

```ts
type PageStatus = "pending" | "running" | "ready" | "error" | "confirmed";

type PageSource = {
  id: string;
  file: File;
  name: string;
  previewUrl: string;
  status: PageStatus;
  errorMessage?: string;
  /** blob/remote URL большого превью после OCR, если отличается */
  recognizedImageUrl?: string | null;
  batchReviewText: string;
};
```

```tsx
// Галерея до recognize
<SourceImageGallery
  pages={pages}
  activeId={activeId}
  onSelect={setActiveId}
  onRemove={removePage} // pending | error | ready; не running/confirmed
  onAdd={addPages} // всегда в хвост
  disabled={/* running page can't be removed; gallery add ok if under cap */}
/>
```

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit | конвейер: порядок статусов, continue after error, soft-cap, confirm merge order |
| RTL | галерея ✕/multi-add; сверка: switch page меняет img+text; progressive ready |
| Flow | plates (и один non-plates smoke): 2 fake images → 2 OCR mocks → confirm → draft lines |
| Regress | SourceInputCard text lint / photo-only recognize; существующие InputStep tests |

---

## Boundaries

**Always**
- Reuse per-file OCR pipeline; не обходить verify/preprocess
- Прогрессивный UI: не блокировать сверку первой `ready` ожиданием хвоста
- Тесты на конвейер и галерею до merge

**Ask first**
- Введение server job API / новых зависимостей очереди
- Поднятие soft-cap или платный параллельный OCR (2+ одновременных вызова)
- Multi-page внутри UI append / смены номенклатуры
- Изменение семантики «Список верен» для одиночного файла (регресс)

**Never**
- Молча мержить все страницы в один textarea без навигации
- Запускать OCR на drop без кнопки
- Коммитить секреты / живые БД
- Ломать текстовый lint-гейт ради галереи

---

## Phased delivery

### Phase A — MVP (эта спека)

1. `PageSource` state + галерея превью с ✕  
2. Sequential recognize hook на существующих API  
3. Сверка: active page ↔ image + text; confirm per page  
4. Сборка draft из confirmed; soft-cap; error continue; remove + add-to-end  
5. Тесты S1–S11

### Phase B — server job (отдельная спека/дописка)

- Upload / job id, progress polling, устойчивость к reload  
- Тот же `PageSource` status machine на фронте

---

## Out of scope (Not Doing)

- Drag-and-drop reorder страниц  
- OCR при добавлении файла  
- Склейка изображений / один multi-image LLM call  
- Глубокий редизайн append / AI  
- Параллельный OCR всех страниц сразу (можно revisit ради latency, но бьёт по rate limit/cost)  
- Удаление уже `confirmed` страницы без полного reset  
- In-place «заменить фото в том же слоте» — вместо этого ✕ + add в хвост

---

## Open Questions

Закрыты (2026-08-31). Residual только на IMPLEMENT: точные подписи кнопок галереи и копирайт «Распознано k/n».

---

## Code review (2026-08-31) — Approve (after Phase A.1)

Пятиосевой разбор Phase A. Initial verdict was **Request changes**; remediation closed R1–R7 with S12–S17 green (2026-08-31).

### Context

- Цель Phase A достигнута по happy path: галерея, sequential OCR, прогресс, confirm по странице, тесты lib/hook/gallery green.
- Проблемы — в error/empty edge cases, видимости статусов и согласованности wizard state после remove — **исправлены в Phase A.1**.

### Findings → обязательные исправления

| ID | Severity | Ось | Проблема | Исправление (в scope до merge) | Status |
|----|----------|-----|----------|--------------------------------|--------|
| **R1** | Critical | Correctness | После `hasStarted`, если пользователь ✕ все страницы (`pages.length === 0`), `allConfirmed === false` → `multiPendingReview` навсегда true; wizard застревает в batch-review. | При `pages.length === 0`: сбросить multi-сессию (`reset()` или эквивалент: `hasStarted=false`, снять multi pending). Тест: start → remove all → `hasStarted===false` / `multiPendingReview` не блокирует. | ✅ S12 |
| **R2** | Critical | Correctness / UX | `page.errorMessage` пишется в хуке, но **нигде не рендерится**. OCR fail на странице незаметен; менеджер не понимает, что ✕ и add в хвост. | Показывать ошибку активной/`error` страницы (Alert под галереей или на превью). Минимум: `activePage.status==='error'` → Alert с `errorMessage`. Тест RTL. | ✅ S13 |
| **R3** | Important | Correctness / UX | В галерее нет визуального статуса `pending` / `running` / `error` / `confirmed` — только active border. | Бейдж или оверлей на thumbnail (спиннер/«!»/✓). `aria-label` включает статус. Тест: error page имеет доступное имя/текст статуса. | ✅ S14 |
| **R4** | Important | Correctness | `allConfirmed` требует `every === confirmed`. Страница `error` никогда не confirm → без ✕ выход невозможен (ок по A8), но без R2/R3 это выглядит как «кнопка сломалась». | Вместе с R2/R3: copy/hint «Удалите страницу с ошибкой и добавьте фото в конец». Опционально: `canFinalize` = нет pending/running/error и все оставшиеся confirmed. | ✅ hint |
| **R5** | Important | Correctness | `handleRemovePage` / `handleAddFiles` читают `multiPage.pages` **сразу после** `remove`/`addFiles` — это React state прошлого рендера (stale). `imageName` в store может врать. | Считать remaining/count из **возвращаемого** результата хелперов / синхронного `pagesRef`, либо `set-source` только по `name` из callback. Не читать `multiPage.pages` в том же тике после мутации. Тест на wizard/handler или хук API, возвращающий next pages. | ✅ S15 |
| **R6** | Important | Architecture | При `ready` хук делает второй `createObjectURL(file)` в `recognizedImageUrl`, хотя `previewUrl` уже есть → лишний blob, риск забыть revoke. | Не дублировать URL: `recognizedImageUrl = previewUrl` или не задавать отдельное поле для local File. Тест: после ready revoke вызывается один раз на page remove. | ✅ S16 |
| **R7** | Important | Correctness | Multi `handleConfirmBatch` при финале берёт `batchCount` только из plate/pile/step/march; **fbs / bridge_piles** падают в `plate_batches` (часто 0). `confirm-batch-review` всё равно снимает pending, но `confirmedBatchCount` врёт для следующего append. | Использовать тот же `getBatchCount(draft)` что в store (уже знает fbs/bridge), в multi и single confirm path. Тест на fbs или unit на helper. | ✅ S17 |

### Nit / Optional (не блокируют, но зафиксировать)

| ID | Severity | Проблема | Рекомендация |
|----|----------|----------|--------------|
| N1 | Nit | В wizard `Math.min(12, …)` вместо `MAX_PAGES` | Импорт константы — частично снято (handlers больше не считают cap локально) |
| N2 | Consider | Глобальный Alert OCR-corrections на весь draft при просмотре одной страницы | Фильтровать/помечать «по текущей странице» позже |
| N3 | FYI | Шесть `*InputStep` всё ещё копируют wiring; `multiPageStepProps` уже снижает боль | Не раздувать рефактор сейчас |
| N4 | FYI | Phase B server job — out | Без изменений |

### Verification (после фиксов)

| # | Критерий | Метод | Status |
|---|----------|--------|--------|
| S12 | Remove all pages after start → multi-сессия сброшена, нет вечного pending | unit hook | ✅ |
| S13 | OCR error на странице виден в UI | RTL | ✅ |
| S14 | Thumbnail отражает error/running | RTL | ✅ |
| S15 | Remove/add не опирается на stale `pages` для store imageName | unit/RTL | ✅ |
| S16 | Один object URL на page; revoke без двойного leak | unit hook | ✅ |
| S17 | fbs/bridge confirm ставит `confirmedBatchCount` из реальных `*_batches` | unit | ✅ |

### Verdict

- [x] **Approve** — R1–R7 + S12–S17 green (Phase A.1, 2026-08-31)
- [ ] **Request changes** — ~~см. таблицу выше~~ closed

### Phase A.1 — remediation (до merge)

1. R1: empty-after-start reset — ✅  
2. R2 + R3 + R4: error visibility + status chrome + hint — ✅  
3. R5: sync remove/add → store without stale read — ✅  
4. R6: drop duplicate object URL — ✅  
5. R7: shared `getDraftBatchCount` for confirm — ✅  
6. Прогон `npm run test -- src/features/commercial-offer` + typecheck — ✅ (178 tests, typecheck green)

---

## Phase A.2 — busy false positive + lightbox before OCR (2026-08-31)

Источник: user confirm — busy-fix + lightbox до OCR; progressive review без redesign.  
One-pager: [kp-multi-page-lightbox-and-busy-fix.md](../ideas/kp-multi-page-lightbox-and-busy-fix.md).

### Findings

| ID | Severity | Ось | Проблема | Исправление | Status |
|----|----------|-----|----------|-------------|--------|
| **R8** | Critical | Correctness / UX | `isRecognizingMulti = isRunning \|\| hasPendingOrRunning` — `hasPendingOrRunning` true для `pending` **до** `hasStarted`. После add files кнопка «Распознавание...» и disabled, OCR не стартовал. | Busy только когда recognition реально начат: `isRunning \|\| (hasStarted && hasPendingOrRunning)` или `isRecognizing` из хука. До start: «Распознать фото», enabled при pages>0. | ✅ S18 |
| **R9** | Important | UX | До OCR нет крупного просмотра страниц — только 72px thumbnails. | Клик thumbnail **до** `hasStarted` → lightbox (Esc / backdrop / close); ←/→ листают; ✕ на thumb всё ещё remove. После `hasStarted` — select/review как сейчас. | ✅ S19 |

### Success criteria (A.2)

| # | Критерий | Метод | Status |
|---|----------|--------|--------|
| S18 | Add files without start → not treating as recognizing; button «Распознать фото», enabled | unit hook / RTL | ✅ |
| S19 | Thumbnail click before start opens lightbox; Esc closes; ←/→ navigate; after start select behavior intact | RTL | ✅ |

### Phase A.2 — remediation checklist

1. R8: busy only after recognition started — ✅  
2. R9: lightbox before OCR (open / Esc / next-prev) — ✅  
3. Прогон `npm run test -- src/features/commercial-offer` + typecheck — ✅ (190 tests, typecheck green)  

### Out of A.2

- Progressive review redesign  
- Server job Phase B  
- Permanent split-preview pane before OCR  
- Commit (unless asked)

---

## Phase A.3 — Wait banner + AI on review

Companion: [kp-ocr-wait-and-ai-on-review.md](./kp-ocr-wait-and-ai-on-review.md) · plan [2026-08-31-kp-ocr-wait-and-ai-on-review.md](../develop/plans/2026-08-31-kp-ocr-wait-and-ai-on-review.md).

IMPLEMENT ✅ (wait + AI on review). Companion R10–R11 and batch-helper R12–R13 closed in Phase A.3.1 / A.4 (2026-08-31).

---

## Code review (2026-08-31) — Approve (batch helpers)

| ID | Severity | Ось | Проблема | Исправление | Status |
|----|----------|-----|----------|-------------|--------|
| **R12** | Critical | Correctness | `getBatches` / `getCurrentBatchReviewText` без fbs/bridge — multi OCR кладёт cumulative text; confirm merge **дублирует** строки | Расширить `getBatches` как `getDraftBatchCount`; per-page text = last batch only | ✅ S22 |
| **R13** | Important | Correctness | Single-path `handleConfirmBatch` merge для fbs/bridge всё ещё берёт `plate_batches` (R7 чинил только count) | Shared getBatches во всех confirm paths (merge + count) | ✅ S23 |

### Success criteria

| # | Критерий | Метод | Status |
|---|----------|--------|--------|
| S22 | fbs/bridge 2-page mock → distinct page texts; confirm без дублей | unit / flow | ✅ |
| S23 | fbs/bridge single edit-confirm merges correct `*_batches` | unit | ✅ |

### Phase A.1.x / A.4 — remediation (до merge)

1. R12 — getBatches + getCurrentBatchReviewText for fbs/bridge — ✅  
2. R13 — confirm merge uses shared helper — ✅  
3. commercial-offer tests + typecheck — ✅  

См. также R10–R11 в [kp-ocr-wait-and-ai-on-review.md](./kp-ocr-wait-and-ai-on-review.md).
