# Spec: КП — правка и удаление строки иконками

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-08-31  
**One-pager**: [ai_docs/ideas/kp-row-edit-delete-icons.md](../ideas/kp-row-edit-delete-icons.md)  
**Plan**: [ai_docs/develop/plans/2026-08-31-kp-row-edit-delete-icons.md](../develop/plans/2026-08-31-kp-row-edit-delete-icons.md)  
**Related**: [kp-multi-nomenclature-append.md](./kp-multi-nomenclature-append.md) (`DELETE .../lines/{line_id}`), [unparsed-line-live-highlight.md](./unparsed-line-live-highlight.md)

## Objective

**Проблема.** На шаге 1 таблица «Состав КП (предпросмотр)» только для чтения: ошибку в марке или количестве после появления цены приходится ловить на сверке или ждать шаг 3. На шаге 3 удаление уже есть, но это широкая кнопка «Удалить», без правки строки и без отмены этой операции (есть только откат **пакета**).

**Цель.** (1) Карандаш и мусорка в каждой строке состава на предпросмотре шага 1 (все 6 типов) и на шаге 3. (2) Qty → новая цена линии без полного ILP. Марка → парсер этой строки; 0/1/2+ позиций = **замена** этой строки. (3) Карточки wide / «нет в прайсе» **как сейчас**; карандаш не блокируем. (4) Тост «Отменить» последней операции строки, не путать с undo-last-batch.

**Пользователь:** менеджер в мастере КП после «Список верен» (предпросмотр с ценами) и на шаге «Результат».

**Успех:** в строке две иконки; карандаш меняет qty или марку и таблица/итоги обновляются; мусорка удаляет сразу; «Отменить» откатывает последнюю такую операцию; wide/unpriced не переделываем.

---

## ASSUMPTIONS I'M MAKING

1. **Экраны.** Иконки только в `Kp*PreviewPanel` (шаг 1 **после** сверки, не в `PlateListEditor` на batch-review) и в `CalculationResultStep` (шаг 3). На сверке список и так редактируется.
2. **Карандаш = qty и/или `source_text`.** Не свободное «наименование как в PDF». Поле «как в списке»: `ПБ 28-5,3-8п 2`. Парсер — **тип этой строки** (`line.product_type`), не всегда плиты.
3. **Qty-only без ILP.** Меняем `qty` существующей линии, пересчитываем сумму/веса/totals. `generate_preview` / раскладку всего КП не гоняем.
4. **Марка = preview фрагмента.** `source_text` прогоняем через тот же `spec.generate_preview` / stamp, что одна новая строка этого типа. Результат **0 / 1 / N линий** вставляем вместо одной (замена). ILP всего заказа не перезапускаем.
5. **1→0 / 1→2+.** Пустой разбор при непустом вводе = ошибка у строки, состав не трогаем. Пустой разбор + пользователь явно стёр смысл? Нет: 0 позиций от **успешного** parse/preview (модель вернула пустое) = строка исчезает. 2+ = две (и более) вместо одной. Wide split из preview фрагмента — принять.
6. **Wide / unpriced.** UI-карточки и их API **без изменений**. После PATCH/DELETE возвращаем полный draft; карточки читают metadata как сейчас. Карандаш **доступен**, даже если карточки открыты.
7. **Undo.** Один слот: последняя операция строки. Тост ~8 с с кнопкой «Отменить». Слот сбрасывается по таймауту, по следующей операции строки, или при уходе с шага мастера. **Не** серверный журнал, **не** undo-last-batch.
8. **Нет `line_id`.** Иконки не показываем (как сейчас delete на шаге 3).
9. **Без новых npm/pip зависимостей.** Иконки — inline SVG. Тост — существующий `Alert` / тонкая плашка над таблицей, не новая toast-библиотека.
10. **Коммиты агент не делает**, пока явно не попросите.

Assumptions **approved 2026-08-31** (user: «вперед переходим к работе»).

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| **D-where** | Где иконки | Предпросмотр шага 1 (6 панелей) + шаг 3. Не batch-review editor |
| **D-icons** | Вид | Карандаш в квадрате + мусорка; ghost, ~16–20px; `aria-label`. Шаг 3: текст «Удалить» убрать |
| **D-qty** | Количество | PATCH qty существующей линии; цена единицы та же; totals пересчитать |
| **D-mark** | Марка | Поле `source_text` + парсер/preview **этого** `product_type`; замена 1 → 0..N |
| **D-replace** | 0 / 2+ из preview | Принять замену строки. Невалидный ввод (парсер не ест) — ошибка, без mute-delete |
| **D-wide** | Wide / unpriced | Карточки as-is; карандаш не disable |
| **D-delete** | Удаление | Сразу; существующий `DELETE /drafts/{id}/lines/{line_id}` |
| **D-undo** | Отмена | Клиентский снимок inverse + API; тост ~8 с; не undo-last-batch |
| **D-no-ilp-full** | ILP | Не пересобирать весь заказ на qty и не на mark-фрагмент целиком |

---

## User Stories

- Как **менеджер**, в предпросмотре состава я меняю количество карандашом и сразу вижу новую цену строки и итоги.
- Как **менеджер**, я исправляю марку («как в списке»); если из одной строки вышло две — в таблице две вместо одной.
- Как **менеджер**, я удаляю строку мусоркой на шаге 1 и на шаге 3 (без слова «Удалить»).
- Как **менеджер**, если промахнулся, жму «Отменить» на тосте и строка/qty/марка возвращаются.
- Как **менеджер**, карточки «шире стандартной» и «нет в прайсе» работают как раньше; карандашом я могу поправить строку, из‑за которой они появились.

---

## Tech Stack

| Слой | Стек |
|------|------|
| Frontend | React 19, TS, Vite, Vitest + Testing Library, TanStack Query |
| Backend | FastAPI, существующий `CommercialDraftLifecycle` / `ProductDraftHandler` / `ProductDraftSpec` |
| API | REST `/api/v1/commercial/drafts/...` |

Новых пакетов нет.

## Commands

```
# Frontend
cd frontend && npm run test -- src/features/commercial-offer
cd frontend && npm run typecheck

# Backend (минимум контракт PATCH/restore + delete regress)
pytest tests/test_commercial_draft_append.py tests/test_commercial_web_flow.py -q
```

Dev: уже запущенный `./run+logs.sh` не убивать.

## Project Structure

```
frontend/src/features/commercial-offer/
  components/Kp*PreviewPanel.tsx     → колонка действий + line_id
  components/steps/CalculationResultStep.tsx
  components/LineRowActions.tsx      → общий карандаш/мусорка (новый)
  lib/buildKpPreviewRows.ts          → прокинуть line_id
  api/commercialOfferApi.ts          → PATCH line + restore
app/api/v1/endpoints/commercial.py
app/services/commercial_draft_lifecycle.py
app/schemas/commercial.py
tests/                               → pytest PATCH / replace / undo restore
ai_docs/specs/  ai_docs/ideas/
```

## Code Style

Иконки — `button` + SVG, не новая библиотека. Пример контракта UI:

```tsx
<button type="button" aria-label={`Изменить строку ${lineId}`} onClick={onEdit}>
  {/* pencil-in-square SVG */}
</button>
<button type="button" aria-label={`Удалить строку ${lineId}`} onClick={onDelete}>
  {/* trash SVG */}
</button>
```

Карандаш открывает inline-поля на этой строке (qty + `source_text`), не модалку. Сохранить / Esc-отмена редактирования **без** записи в undo-слот, пока запрос не успел.

Python: логика в lifecycle, роутер тонкий; русские `detail` как у delete (`Строка не найдена.`).

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit TS | `buildKpPreviewRows` отдаёт `lineId`; helper replace 1→N в тестовом виде если вынесен |
| RTL | Иконки на предпросмотре и шаге 3; нет текста «Удалить»; карандаш → поля; delete зовёт колбэк; тост undo |
| API pytest | PATCH qty; PATCH source_text 1→1 цена; 1→2 splice + новые `line_id`; невалидный текст 4xx состав тот же; restore после delete; append_batches.line_ids согласованы |
| Regress | Существующий `test_delete_line_*`, undo-last-batch **не** ломается |

## API (черновик контракта)

**PATCH** `/api/v1/commercial/drafts/{draft_id}/lines/{line_id}`  
Auth как у delete. Body:

```json
{ "qty": 90 }
{ "source_text": "ПБ 28-5,3-8п 2" }
{ "qty": 2, "source_text": "ПБ 28-5,3-8п 2" }
```

- Только `qty` — та же линия, тот же `line_id`.
- Есть `source_text` — preview фрагмента этого `product_type`, **замена** линии на 0..N новых stamped lines; старый `line_id` уходит из `order_data` и из `append_batches[].line_ids`; новые id в тот же batch, если старый там был.
- Невалидный `source_text` → 400, русское сообщение, draft без изменений.
- 404 если линии нет — как delete.

**POST** `/api/v1/commercial/drafts/{draft_id}/lines/restore`  
**Locked at PLAN:** dedicated restore (not inverse PATCH). Inverse PATCH cannot meet S7 after DELETE (404); after 1→N it would re-parse instead of restoring the snapshot. Qty-only undo uses inverse PATCH qty.  
Body: `{ "lines": [ { ...order line snapshot } ], "index": <number>, "replace_line_ids": [] }` — optionally remove `replace_line_ids` first, then splice `lines` at `index` (undo delete / undo replace).

Ответ обоих: `CommercialDraftDetails` (hydrate как после delete). Totals пересчитать существующим `compute_totals`.

**DELETE** — без изменений.

## Success Criteria

| # | Критерий | Метод |
|---|----------|--------|
| S1 | Предпросмотр (хотя бы плиты) + шаг 3: видны иконки изменить/удалить у строк с `line_id` | RTL | ✅ |
| S2 | Шаг 3: нет кнопки с текстом «Удалить» | RTL | ✅ |
| S3 | PATCH qty → qty и сумма строки/totals обновились; `line_id` тот же | pytest | ✅ |
| S4 | Валидный `source_text` 1→1: марка/цена обновились | pytest | ✅ |
| S5 | `source_text` → 2+ линий: старая снята, новые на её месте, новые `line_id` | pytest | ✅ |
| S6 | Невалидный `source_text` → 400, состав не изменился, в UI ошибка у строки | pytest + RTL | ✅ |
| S7 | Мусорка → существующий delete; тост; «Отменить» возвращает строку | RTL + pytest restore | ✅ |
| S8 | Undo-last-batch кнопка/API без регресса | pytest regress | ✅ |
| S9 | Wide/unpriced секции в коде панелей не выпилены; карандаш не `disabled` из‑за их флагов | RTL / code review | ✅ |
| S10 | `npm run test -- src/features/commercial-offer` + `npm run typecheck` + релевантный pytest зелёные | CI / локально | ✅ (line-edit + append; 3 pre-existing web_flow fails unrelated) |

## Boundaries

**Always**
- Общий `LineRowActions` на шаг 1 и шаг 3
- Прокинуть `line_id` в preview rows
- Русские ошибки API
- Карандаш при открытых wide/unpriced

**Ask first**
- ~~Тело restore, если inverse PATCH достаточен~~ → **locked PLAN:** POST restore (see API)
- Гнать `generate_preview` фрагмента для **плит**, если без него нельзя получить 1→2 — **yes, fragment only**
- Менять metadata wide/unpriced кроме того, что уже делает текущий preview-путь фрагмента — **do not rewrite** draft-level wide/unpriced on PATCH

**Never**
- Новая иконка-зависимость
- Ручная цена / пустая новая строка / confirm-dialog
- Подмена undo-last-batch
- Полный ILP всего КП на смену qty
- Disable карандаша из‑за wide/unpriced
- Commit без явной просьбы
- Правки `bot_archived`

## Out of scope (Not Doing)

- Ручной unit_price, добавление пустой строки, Excel-редактор
- Иконки на экране сверки с фото
- Confirm перед удалением
- Серверный undo-журнал / persist undo между сессиями
- Редизайн карточек wide / unpriced

## Phased delivery

Одна фаза после «спека ок» + PLAN: API qty/mark/replace → иконки UI → undo toast → тесты.

## Open Questions

- ~~Нужен ли отдельный `POST .../restore`?~~ **PLAN:** dedicated POST restore. Qty undo = inverse PATCH qty.
- Точное место тоста: **PLAN:** над таблицей состава (preview Card) и над «Позиции» на шаге 3 — не `stepError`.

## Risks

| Риск | Почему | Смягчение |
|------|--------|-----------|
| `generate_preview` плит на фрагмент всё же тяжёлый | ILP | Qty-path без preview; mark-path только текст этой строки |
| Замена 1→N разъезжается с `append_batches` | line_ids | Тест: id старой линии вычищен, новые в том же batch |
| Undo путают с «отменить добавление» | Две кнопки | Копирайт тоста: «Строка удалена» / «Количество изменено», не «добавление» |
