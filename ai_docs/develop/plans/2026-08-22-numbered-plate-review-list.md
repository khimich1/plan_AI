# Plan: Нумерация списка плит при сверке

> **Spec:** [`ai_docs/specs/numbered-plate-review-list.md`](../../specs/numbered-plate-review-list.md)  
> **Идея:** [`ai_docs/ideas/numbered-plate-review-list.md`](../../ideas/numbered-plate-review-list.md)  
> **Дата:** 2026-08-22  
> **Статус:** IMPLEMENT ✅ (PN-101…106) — PN-107 manual S7 pending

---

## Approach

Не трогаем текст списка и бэкенд. Считаем 1-based номера по непустым строкам чистой функцией. `PlateListEditor` рисует их в оверлее подсветки; textarea остаётся источником правды. Флаг `showLineNumbers` включает только `PlateInputStep`.

```
batchReviewText
    │
    ├── value.split("\n")
    │
    ├── assignNonEmptyLineNumbers(lines)  → (number | null)[]
    │
    └── PlateListEditor(showLineNumbers?)
              │
              ├── overlay: номер в per-line div (явный color)
              └── textarea: тот же value, extra padding-left = гуттер
```

---

## Architecture Decisions

- **Номера не в `value`.** Иначе парсер и `mergeEditedBatchIntoFullText` увидят `1.`. Отклонённый вариант ideation.
- **Чистая функция в `lib/`.** Правило «непустая = trim» тестируется без React.
- **Номер внутри существующего per-line overlay `div`.** Высота логической (в т.ч. wrapped) строки уже есть у подсветки; отдельная колонка вне оверлея разъедется.
- **Проп, не форк компонента.** Default `false` — сваи/ФБС/марши/ступени без изменений. Включить позже = одна строка в их `*InputStep`.
- **Ширина гуттера:** `max(2, String(maxNumber).length)` ch. Не фиксируем `3ch` на трёх плитах.

---

## Components

| Компонент | Роль | Зависит от |
|-----------|------|------------|
| `assignNonEmptyLineNumbers` | правило S1–S3 | — |
| `lineNumberGutterCh` | ширина гуттера | max номер |
| `PlateListEditor` | гуттер + padding | функции выше |
| `PlateInputStep` | `showLineNumbers` | редактор |

Backend, OCR, `batchReview.ts`, `KpPlatePreviewPanel`, прочие `*InputStep` — не в этом плане.

---

## Implementation order

Последовательность обязательна (TDD). Параллелить нечего: один UI-слайс.

```
1. Tests (RED)     → plateListLineNumbers
2. Function        → GREEN
3. Tests (RED)     → PlateListEditor RTL
4. Editor + padding → GREEN
5. Wire plates     → showLineNumbers
6. Regression      → frontend test + typecheck
7. Manual S7       → не блокер merge кода
```

---

## Tasks

- [x] **PN-101:** RED — unit на нумерацию
  - Acceptance: S1–S3 описаны тестами и падают; кейсы: `""`, только пробелы, `"A\n\nB"`, leading/trailing newline, `"  A  \nB"`
  - Verify: `cd frontend && npm run test -- src/features/commercial-offer/lib/plateListLineNumbers.test.ts` — fail (модуля ещё нет или функция не экспортирована)
  - Files: `frontend/src/features/commercial-offer/lib/plateListLineNumbers.test.ts`
  - Scope: XS

- [x] **PN-102:** GREEN — `assignNonEmptyLineNumbers` + `lineNumberGutterCh`
  - Acceptance: S1–S3 green; `lineNumberGutterCh(9) === 2`, `lineNumberGutterCh(10) === 2`, `lineNumberGutterCh(100) === 3`
  - Verify: тот же test file — green
  - Files: `frontend/src/features/commercial-offer/lib/plateListLineNumbers.ts`, тест из PN-101
  - Dependencies: PN-101
  - Scope: XS

- [x] **PN-103:** RED — RTL `PlateListEditor`
  - Acceptance: при `showLineNumbers` в документе есть `1` и `2` для двух непустых строк; `textarea` value = исходный текст без `1.`; без флага номеров нет (`queryByText` по отдельным номерам в гуттере — через `data-testid="plate-line-number"` чтобы не ловить `2` из количества)
  - Verify: `cd frontend && npm run test -- src/features/commercial-offer/components/PlateListEditor.test.tsx` — fail
  - Files: `frontend/src/features/commercial-offer/components/PlateListEditor.test.tsx`
  - Dependencies: PN-102 (фикстура драфта как в `plateLineHighlights.test.ts`)
  - Scope: S

- [x] **PN-104:** GREEN — гуттер в `PlateListEditor`
  - Acceptance: S4–S5; проп `showLineNumbers?: boolean` default `false`; номер в overlay-строке, `color` явный, `aria-hidden` / `data-testid="plate-line-number"`; `padding-left` textarea = `0.9rem + gutterCh`; оверлей тот же сдвиг; подсветка/добор не регрессируют визуально
  - Verify: RTL из PN-103 — green
  - Files: `frontend/src/features/commercial-offer/components/PlateListEditor.tsx`, тест из PN-103
  - Dependencies: PN-103
  - Scope: S

- [x] **PN-105:** Включить только у плит
  - Acceptance: S6 — `PlateInputStep` передаёт `showLineNumbers`; grep по `*InputStep.tsx` — флаг только в плитах
  - Verify: чтение diff; регрессия `cd frontend && npm run test && npm run typecheck`
  - Files: `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`
  - Dependencies: PN-104
  - Scope: XS

- [x] **PN-106:** Regression pack (S8)
  - Acceptance: весь frontend test + typecheck green
  - Verify: `cd frontend && npm run test && npm run typecheck`
  - Files: нет, если зелёный
  - Dependencies: PN-105
  - Scope: XS

- [ ] **PN-107:** Manual S7 (после кода, не блокер merge)
  - Acceptance: wrap длинной марки — номер у первой визуальной строки; список 15+ рядом с фото; копирование без номеров; сваи без номеров
  - Verify: wizard через уже запущенный `./run+logs.sh`
  - Dependencies: PN-106

---

## Verification checkpoints

| After | Command | Expected |
|-------|---------|----------|
| PN-101 | `npm run test -- src/features/commercial-offer/lib/plateListLineNumbers.test.ts` | RED |
| PN-102 | тот же | green |
| PN-103 | `npm run test -- src/features/commercial-offer/components/PlateListEditor.test.tsx` | RED |
| PN-104 | тот же | green |
| PN-106 | `npm run test && npm run typecheck` | green |

---

## Risks

| Риск | Почему | Mitigation |
|------|--------|------------|
| Номер невидим | оверлей `color: transparent` | явный `color` на span номера |
| Текст и подсветка разъедутся | разный `padding-left` | одна константа padding + `gutterCh` на оверлей и textarea |
| RTL ловит `2` из «ПБ … 2» | количество в конце строки | `data-testid="plate-line-number"`, не `getByText("2")` по всему дереву |
| Wrap разъезжает номер | номер вне line-div | номер *внутри* per-line overlay div, `align-items: flex-start` |
| Случайно включить везде | общий компонент | default `false`; PN-105 только `PlateInputStep` |

---

## Out of scope

Префиксы в тексте, другие изделия, счётчик в заголовке, `×кол-во`, чек-лист, bbox на фото, превью КП, backend.
