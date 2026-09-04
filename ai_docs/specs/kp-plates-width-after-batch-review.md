# Spec: Гейты ширины — только после сверки OCR (экран 2 шага «Плиты»)

> **Источник идеи:** ideation-сессия 2026-09-03 (idea-refine)  
> **Фаза SDD:** SPECIFY ✅ → PLAN (ожидает) → TASKS/IMPLEMENT (частично ✅ 2026-09-03)  
> **Parent:** [`kp-nevernaia-shirina.md`](./kp-nevernaia-shirina.md) (детекция + карточка + API resolve)  
> **Related:** [`kp-review-apply-sync-and-highlights.md`](./kp-review-apply-sync-and-highlights.md), [`ux-wizard-step-plates.md`](./ux-wizard-step-plates.md)  
> **Дата:** 2026-09-03  
> **Статус:** реализовано (wide на сверке; invalid/unpriced после «Список верен»; bugfix resolve), коммит — по просьбе

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| 1 | Где живут карточки ширины | **Wide (>12):** на экране сверки OCR (под списком). **Invalid (завод не режет) + unpriced:** только после «Список верен» (`pendingBatchReview = false`) |
| 2 | Шаг «2. Клиент» | Карточки ширины **не** переносим на шаг клиента |
| 3 | «Список верен» | Доступен **без** решения по ширине (wide / invalid / unpriced) |
| 4 | Подсветка на сверке | **Wide** (шире 12 дм) — **да**, красная подсветка строк; tooltip: «решение на следующем экране» |
| 5 | Подсветка invalid на сверке | **Нет** — решение только на экране 2 |
| 6 | Gate «Готово, далее» | Заблокирован, пока не решены wide → invalid width → unpriced (порядок гейтов без изменений) |
| 7 | Экран 2 (целевой UI) | Как скриншот 3: карточка «Нестандартная ширина» + «Состав КП (предпросмотр)» + «Применить» |
| 8 | Bugfix «Применить» | Resolve invalid width должен находить строку в `input_text` **по размерам**, если `item.line` = display name (`Плиты ПБ …`), а в списке — компактная марка (`ПБ … qty`) |

→ Assumptions approved в чате 2026-09-03.

---

## Assumptions

1. «Второй экран» = тот же wizard step `plates`, флаг `pendingBatchReview = false` после confirm batch (не отдельный sidebar step).
2. Менеджер — тот же человек на обоих экранах; отдельная роль «коммерческий» не нужна.
3. Мульти-страничный OCR: сверка постранично; гейты ширины — по **всему** накопленному draft после последнего «Список верен» по всем страницам (как сейчас).
4. Порядок карточек на экране 2: wide → invalid width → unpriced → предпросмотр КП (как в parent spec).
5. Backend API и metadata **не меняем** — только UX placement + matching fix в resolve.
6. Коммиты агент не делает без явной просьбы.

---

## Objective

**Проблема.** На экране сверки фото ↔ список карточки «Нестандартная ширина» отвлекают от главной задачи — подтвердить OCR. Кнопка «Список верен» блокировалась из‑за нерешённой ширины. При тестах «Применить» в карточке invalid width не менял состав КП: backend не находил строку в тексте, когда `invalid_width_lines[].line` хранил display name, а `normalized_lines` — компактную марку.

**Цель.** Разделить **сверку** и **коммерческие решения по ширине** на два последовательных экрана внутри шага «Плиты». После «Применить» список и предпросмотр обновляются.

**Пользователь:** менеджер, собирающий КП из OCR/текста.

**Успех:**
- OCR-сверка быстрая: фото + список + подсветка wide, без карточек решений.
- После «Список верен» — экран с карточками и предпросмотром (скрин 3).
- «Применить» переписывает строку (напр. `68-11-10п` → `68-12-10п`), карточка и алерт исчезают.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | менеджер на сверке OCR | видеть широкие строки подсвеченными, но без карточки решений | быстро нажать «Список верен» |
| US-2 | менеджер после сверки | решать wide / invalid / unpriced на одном экране с предпросмотром | видеть цены и состав КП |
| US-3 | менеджер | нажать «Применить» в invalid width | увидеть обновлённую марку в «Составе КП» |
| US-4 | менеджер | не уйти к клиенту с нерешённой шириной | «Готово, далее» заблокировано до resolve |

---

## Два экрана шага «Плиты»

### Экран 1 — сверка (`pendingBatchReview = true`)

```
┌─────────────────────────────────────────────────────────────────┐
│ Шаг 1. Плиты — «Сверьте распознанный список…»                   │
├──────────────────────────────┬──────────────────────────────────┤
│ Исходное фото                │ Список плит для расчёта          │
│                              │ (PlateListEditor + highlights)   │
│                              │  • wide → красная подсветка      │
│                              │  • invalid → без подсветки       │
├──────────────────────────────┴──────────────────────────────────┤
│ [Нестандартная ширина] — wide >12 (если есть)                   │
│ [Начать заново]                    [Список верен] [Готово, далее*]│
└─────────────────────────────────────────────────────────────────┘
* «Готово, далее» disabled, пока pendingBatchReview
```

**Показываем:** PageReviewNav, фото, редактор списка, AI-блок (если wired), OCR alerts, **WidePlatesInlineSection**.  
**Скрываем:** `InvalidWidthsInlineSection`, `UnpricedPlatesInlineSection`, `KpPlatePreviewPanel`.

### Экран 2 — решения + предпросмотр (`pendingBatchReview = false`)

```
┌─────────────────────────────────────────────────────────────────┐
│ Шаг 1. Плиты — «Добавьте ещё плиты или перейдите к клиенту»       │
├─────────────────────────────────────────────────────────────────┤
│ [Нестандартная ширина] — wide / invalid (если есть)             │
│ [Без цены в прайсе] — unpriced (если есть)                      │
│ [Состав КП (предпросмотр)] — KpPlatePreviewPanel                 │
│ SourceInputCard / «Добавить к списку»                           │
├─────────────────────────────────────────────────────────────────┤
│ [Начать заново]     [Готово, далее*] [Добавить к списку]        │
└─────────────────────────────────────────────────────────────────┘
* disabled, пока unresolved wide / invalid / unpriced
```

**Целевой вид:** как приложенный скриншот 3 (карточка с радио «10,8 / 12 / Исключить» + фиолетовая «Применить» + таблица предпросмотра).

---

## Tech Stack

| Слой | Стек |
|------|------|
| Frontend | React + TS, `PlateInputStep.tsx`, `plateLineHighlights.ts`, `useBatchReviewHighlights.ts` |
| Backend | Python, `CommercialPlateResolve._match_plate_resolve_item_to_line` |
| API | без изменений: `POST .../invalid-widths/resolve`, wide/unpriced resolve как раньше |
| Тесты | pytest `tests/test_invalid_width_resolve.py`; vitest `PlateInputStep.test.tsx`, `plateLineHighlights.test.ts` |

---

## Commands

```bash
# Backend (корень, .venv)
/home/username/Code/plan_web/.venv/bin/python -m pytest tests/test_invalid_width_resolve.py -q

# Frontend
cd frontend && npm test -- --run PlateInputStep.test.tsx plateLineHighlights.test.ts
cd frontend && npm run typecheck
```

---

## Project Structure (delta)

```
frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx
  → width cards: only !isBatchReviewMode
  → canConfirmBatch: без блокировки по wide
  → canFinishPlates: + invalid + unpriced gates

frontend/src/features/commercial-offer/lib/plateLineHighlights.ts
  → buildPlateLineHighlightMap(..., { batchReview?: boolean })
  → batchReview: wide only, другой tooltip

frontend/src/features/commercial-offer/hooks/useBatchReviewHighlights.ts
  → mergeReviewHighlights(..., { batchReview: true })

app/services/commercial_plate_resolve.py
  → _match_plate_resolve_item_to_line: exact → name substring → _dims_match

tests/test_invalid_width_resolve.py
  → test_resolve_invalid_widths_matches_by_dimensions_when_line_is_display_name
```

---

## Code Style (ключевые фрагменты)

**Условный рендер карточек (frontend):**

```tsx
{!isBatchReviewMode && onWidePlateDecisionChange && onApplyWidePlates && (
  <WidePlatesInlineSection draft={liveWideDraft ?? draft} ... />
)}
{!isBatchReviewMode && draft && onInvalidWidthDecisionChange && onApplyInvalidWidths && (
  <InvalidWidthsInlineSection draft={draft} ... />
)}
```

**Сопоставление строки при resolve (backend):**

```python
def _match_plate_resolve_item_to_line(line: str, items: list[dict]) -> dict | None:
    # 1) exact item.line
    # 2) item.name in line
    # 3) _dims_match(length_m, width_m, load_code, line)
```

---

## Testing Strategy

### pytest

| Тест | Что проверяет |
|------|----------------|
| `test_resolve_invalid_widths_matches_by_dimensions_when_line_is_display_name` | `line="Плиты ПБ 29-8-8п"`, input `ПБ 29-8-8п 1` → rewrite на 8,6 |
| Регрессия `test_resolve_invalid_widths_replace_to_86` | exact line match по-прежнему работает |

### vitest

| Тест | Что проверяет |
|------|----------------|
| `does not show wide card during batch review; confirm stays enabled` | карточки скрыты, «Список верен» enabled при unresolved wide |
| `shows wide card after batch review on post-confirm screen` | карточка visible при `pendingBatchReview={false}` |
| `skips invalid_width highlight on batch review screen` | только wide на экране 1 |

### Ручная проверка (Success Criteria)

1. OCR → экран сверки: wide подсвечены, карточек нет, «Список верен» активен.
2. «Список верен» → экран 2: карточка invalid + предпросмотр.
3. Выбор реза → «Применить» → в таблице новая ширина, алерт и карточка исчезли.
4. «Готово, далее» → клиент (если все гейты закрыты).

---

## Boundaries

**Always**
- Минимальный diff; не трогать шаг «Клиент» для width cards.
- Gate на выходе с «Плит», не на «Список верен».
- Dimensional match в resolve для invalid + unpriced (общий `_lookup_plate_resolve_decision`).

**Ask first**
- Отдельный sidebar step `wide-plates` (legacy).
- Снятие gate с «Готово, далее».
- Объединение трёх «Применить» в одну кнопку.

**Never**
- Карточки решений на экране сверки OCR (после этой спеки).
- Regex-only matching при resolve без fallback по размерам.

---

## Success Criteria

- [x] Карточки invalid / unpriced **не** рендерятся при `isBatchReviewMode`; wide **рендерится** на сверке.
- [x] «Список верен» **не** блокируется `hasUnresolvedWidePlates`.
- [x] На сверке: подсветка wide с текстом «решение на следующем экране»; invalid_width highlight отключён.
- [x] «Готово, далее» блокируется при unresolved wide / invalid / unpriced.
- [x] Resolve invalid width: display name в metadata + compact line в text → текст переписывается.
- [x] pytest `test_invalid_width_resolve.py` — 5 passed.
- [x] vitest `PlateInputStep` + `plateLineHighlights` — green.
- [ ] Полный `pytest tests/ -q` / browser e2e — вне scope этой спеки.

---

## Open Questions

| # | Вопрос | Статус |
|---|--------|--------|
| Q1 | Переносить ли `UnpricedPlatesInlineSection` вместе с invalid на экран 2? | **Да** (реализовано: все три карточки на экране 2) |
| Q2 | Live overlay wide lines в batch editor (`liveWideDraft`) — нужен ли на экране 2? | На экране 2 используется `draft` без live overlay; при необходимости — отдельная задача |
| Q3 | Обновить parent spec `kp-nevernaia-shirina.md` § Design «Порядок на шаге плит»? | Рекомендуется ссылкой на эту спеку |

---

## Not in this spec

- Изменение таблицы заводских резов / детекции invalid width (см. parent).
- Шаг «Клиент» и расчёт КП.
- Новый sidebar step между «Плиты» и «Клиент».
- E2E через Chrome DevTools MCP.

---

## Changelog

| Дата | Изменение |
|------|-----------|
| 2026-09-03 | SPECIFY: двухэкранный flow + bugfix dimensional match |
| 2026-09-03 | IMPLEMENT: `PlateInputStep`, `plateLineHighlights`, `commercial_plate_resolve`, tests |
