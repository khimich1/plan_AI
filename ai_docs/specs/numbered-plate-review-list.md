# Spec: Нумерация списка плит при сверке

> **Тип:** feature-spec (SDD Phase: SPECIFY ✅ → PLAN ✅ → TASKS ✅ → IMPLEMENT ✅)  
> **Дата:** 2026-08-22  
> **Статус:** код PN-101…106 влит в рабочее дерево; PN-107 (ручной S7) открыт  
> **Источник идеи:** [`ai_docs/ideas/numbered-plate-review-list.md`](../ideas/numbered-plate-review-list.md)  
> **План:** [`ai_docs/develop/plans/2026-08-22-numbered-plate-review-list.md`](../develop/plans/2026-08-22-numbered-plate-review-list.md)

---

## ASSUMPTIONS I'M MAKING

1. Канал — **веб wizard КП**, карточка «Список плит для расчёта» в режиме сверки текущего источника (`isBatchReviewMode`). Telegram-бот вне scope.
2. Номер — **только UI**. В `batchReviewText` / `value` / `normalized_text` / парсер / API / БД префиксов `1.` нет.
3. Нумерация — **1-based по непустым строкам**. `trim()` пустой → без номера, счётчик не растёт. «Позиция 3» = третья плита, не третья строка textarea.
4. Scope изделий — **только плиты** (`PlateInputStep` передаёт флаг). Общий `PlateListEditor` получает опциональный проп; default `false`, остальные шаги не меняются.
5. Номера видны **и с фото, и без** — тот же редактор, колонка фото опциональна.
6. После «Список верен» редактор сверки скрывается вместе с режимом. Номера в `KpPlatePreviewPanel` **не** добавляем.
7. Перенос длинной марки **не запрещаем**. Номер стоит в начале логической строки и выровнен по её верху (первая визуальная строка). `white-space` / `word-break` как сейчас.
8. Копирование из textarea **без номеров** — ожидаемо. Номера `aria-hidden` / `user-select: none`, клики проходят в textarea.
9. Подсветки (исправлено / не попало / шире / добор) и бейдж «↔ добор» сохраняют семантику. Фон подсветки может захватывать гуттер, как в редакторе кода.
10. Новых npm-зависимостей нет. Бэкенд не трогаем.

→ **A1–A10 approved 2026-08-22.** P1–P3 locked below.

---

## Decisions locked (из ideation 2026-08-22)

| # | Тема | Решение |
|---|------|---------|
| D1 | Форма номера | Визуальный гуттер слева, не префикс в тексте |
| D2 | Что нумеровать | Только непустые строки |
| D3 | Изделия | Сначала только плиты |
| D4 | Работа пользователя | Сопоставить строку с N-й позицией источника и не пропустить строку |

### Locked 2026-08-22 (P1–P3 + остаток)

| # | Тема | Решение |
|---|------|---------|
| P1 | Нет фото | Номера всё равно |
| P2 | После confirm | Номеров нет — редактор сверки скрывается; `KpPlatePreviewPanel` не трогаем |
| P3 | Wrap марки | Номер у логической строки, сверху; `nowrap` не вводим |
| P4 | Ширина гуттера | `max(2, цифр в maxNumber)` ch — 1–99 как `2ch`, 100+ как `3ch` |
| P5 | RTL подсветки в CI | Не обязателен; S7 — ручной прогон wrap + 15+ строк |

---

## Objective

Менеджер сверяет распознанный список плит с фото или текстом. Сейчас строки идут сплошным моноширинным текстом, а число в конце строки — **количество**, не порядковый номер. Нужен устойчивый индекс позиций, который не ломает правку и разбор марок.

### Пользователь

Менеджер по продажам на шаге ввода плит: смотрит исходное фото/текст слева (если есть) и список справа, правит строки, жмёт «Список верен».

### User Stories

- Как **менеджер**, я вижу `1`, `2`, `3` слева от марок и могу сказать «третья на фото = третья в списке».
- Как **менеджер**, я вставляю пустую строку между плитами — номера плит не сдвигаются из‑за пустой строки.
- Как **менеджер**, я копирую список из поля — в буфере марки и количества, без `1.`.
- Как **менеджер** на сваях / ФБС / маршах / ступенях, я **не** вижу номеров в этой версии (гипотезу проверяем на плитах).

### Success Criteria (измеримые)

| # | Критерий | Метод проверки |
|---|----------|----------------|
| S1 | Две непустые строки → номера `1` и `2` слева, не в `value` | unit + RTL |
| S2 | `"A\\n\\nB"` → номера `1` и `2` на `A` и `B`, пустая без номера | unit |
| S3 | Строка из пробелов считается пустой | unit |
| S4 | `onChange` / `value` не содержат `1.` / `2.` | RTL: textarea value |
| S5 | `showLineNumbers={false}` (default) — номеров в DOM нет | RTL |
| S6 | `PlateInputStep` передаёт `showLineNumbers`; прочие `*InputStep` — нет | grep / чтение вызова |
| S7 | Подсветка и добор-бейдж не ломаются (оверлей + padding синхронны) | RTL на фикстуре с highlight **или** ручной прогон |
| S8 | `npm run test` и `npm run typecheck` в `frontend/` — green | CI / локально |

---

## Tech Stack

| Компонент | Технология |
|-----------|------------|
| UI сверки | `PlateListEditor` + `AutoResizeTextarea` |
| Включение | `PlateInputStep` (batch review) |
| Логика номеров | чистая функция в `frontend/src/features/commercial-offer/lib/` |
| Тесты | vitest + Testing Library |
| Backend | без изменений |

Новых пакетов нет.

---

## Commands

```bash
cd frontend
npm run test -- src/features/commercial-offer/lib/plateListLineNumbers.test.ts
npm run test -- src/features/commercial-offer/components/PlateListEditor.test.tsx
npm run test
npm run typecheck
```

Ручная проверка (стек уже поднят через `./run+logs.sh`): мастер КП → плиты → распознать фото или вставить текст → в карточке «Список плит для расчёта» слева от марок стоят 1, 2, 3; правка строки не пишет номер в текст; «Список верен» убирает этот редактор.

---

## Project Structure

```
frontend/src/features/commercial-offer/lib/plateListLineNumbers.ts
frontend/src/features/commercial-offer/lib/plateListLineNumbers.test.ts
frontend/src/features/commercial-offer/components/PlateListEditor.tsx
frontend/src/features/commercial-offer/components/PlateListEditor.test.tsx   # новый
frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx
ai_docs/ideas/numbered-plate-review-list.md
ai_docs/specs/numbered-plate-review-list.md
```

Не трогаем: backend, парсер, `batchReview.ts` (кроме косвенного: текст по-прежнему сырой), `KpPlatePreviewPanel`, `PileInputStep` / `FbsInputStep` / `MarchInputStep` / `StepInputStep` / `BridgePileInputStep`.

---

## Code Style

Логика номеров — чистая функция, без React. Компонент только рисует.

```ts
/** 1-based index for non-empty (trim) lines; blank → null. */
export function assignNonEmptyLineNumbers(lines: readonly string[]): Array<number | null> {
  let n = 0;
  return lines.map((line) => {
    if (!line.trim()) {
      return null;
    }
    n += 1;
    return n;
  });
}
```

В `PlateListEditor`:

- новый проп `showLineNumbers?: boolean` (default `false`);
- ширина гуттера от числа цифр в максимальном номере (`2ch` / `3ch` + зазор), одинаковая у оверлея и у `padding-left` textarea;
- номер внутри существующего per-line `div` оверлея, `color` явный (родитель оверлея `color: transparent`);
- `pointer-events: none` на номерах, как у оверлея.

`PlateInputStep`:

```tsx
<PlateListEditor
  draft={batchReviewDraft}
  value={batchReviewText}
  onChange={onBatchReviewTextChange}
  minHeight={recognizedImageUrl ? 440 : undefined}
  showLineNumbers
/>
```

---

## Testing Strategy

| Уровень | Что | Где |
|---------|-----|-----|
| Unit | S1–S3, пустой ввод, только пустые, leading/trailing newline | `plateListLineNumbers.test.ts` |
| RTL | S4–S5, номера видны при флаге, `value` без префикса | `PlateListEditor.test.tsx` |
| Контракт | S6 — флаг только у плит | ревью diff / точечный тест не обязателен |
| Ручной | S7 на длинной марке (wrap) и на списке 15+ строк рядом с фото | локальный wizard |

Покрытие веток функции: 100%. Не мокаем OCR и не гоняем backend.

Фикстуру драфта брать по образцу `plateLineHighlights.test.ts`.

---

## Boundaries

### Always

- Номера считать с `value.split("\n")`, не писать обратно в `value`.
- Default `showLineNumbers = false`.
- Синхронизировать padding оверлея и textarea.
- Покрыть правило нумерации unit-тестами до merge.

### Ask first

- Включать флаг на сваях / ФБС / маршах / ступенях / мостовых сваях.
- Нумеровать все строки textarea, включая пустые.
- Добавлять номера в `KpPlatePreviewPanel` или в заголовок «N позиций».
- Менять перенос (`nowrap`) ради гуттера.
- Писать префиксы в нормализованный текст «для копирования».

### Never

- Префиксы `1.` в тексте, который уходит в confirm / merge / парсер.
- Менять семантику количества в конце строки.
- Backend, OCR, bbox на фото, чек-лист сверки.
- Новые зависимости.
- Коммит секретов / живых БД.

---

## Open Questions

Нет. P1–P5 закрыты 2026-08-22.

---

## Out of scope

- Префиксы в тексте, счётчик в заголовке, `×кол-во`, чекбоксы, номера на фото.
- Прочие типы изделий.
- Превью КП после сверки.
- Любые изменения расчёта / PDF / XLSX.
