# Implementation Plan: Подсказка формата заявки

**Created:** 2026-09-25  
**Status:** implemented  
**Orchestration:** `orch-2026-09-25-16-38-kp-format-hint`  
**Spec:** [`ai_docs/specs/kp-format-hint.md`](../../specs/kp-format-hint.md)  
**Idea:** [`ai_docs/ideas/podskazka-formata-zayavki.md`](../../ideas/podskazka-formata-zayavki.md)

> Implemented 2026-09-25: HINT-001…004 green (config texts, FormatHint, SourceInputCard mount, typecheck).

## Overview

На шаге ввода КП над полем списка (все 7 типов) появляется свёрнутое слово «▸ Подсказка». Клик раскрывает общий каркас, пояснение типа и образец (= текущий `labels.placeholder`). «Скопировать образец» пишет только в буфер; поле списка не меняется. Состояние не персистится. Бэкенд и парсеры не трогаем.

**Тексты:** пользователь перешёл к плану без правки черновика. Раздел спеки «Тексты» (общий каркас + семь пояснений) — **утверждённая копия для реализации**, пока ревью плана явно не попросит иное. Open Question в спеке про правку текстов до кода закрыт этим решением.

## Architecture Decisions (locked)

| ID | Решение |
|----|---------|
| D1 | Слово «Подсказка» **между** `listLabel` и `PlateListEditor`, внутри `FieldWrapper` как sibling-дети, **не** через `FieldWrapper.hint` (тот рендерится *под* children). |
| D2 | Старт свёрнут; клик toggle; `useState` локальный; нет localStorage. |
| D3 | Общий каркас — одна константа `FORMAT_HINT_SHARED` рядом с конфигом. Per-type: `explanation` + `sample`, где `sample` **читается** из `labels.placeholder`, не дублируется вручную. |
| D4 | Копирование: `navigator.clipboard.writeText(sample)` только. Успех → краткое «Скопировано»; отказ → «Не удалось скопировать»; `onTextChange` не вызывать. |
| D5 | Все 7 типов через существующий `SourceInputCard` / `productTypeConfig`. `PlateInputStep` / `SimpleProductInputStep` новых пропсов не получают. |
| D6 | Визуальный жест как у «▸ Дополнительно» / «▾ Дополнительно»: текстовая кнопка, `type="button"`, `aria-expanded`. |
| D7 | TDD: каждый behavior-task начинается с падающего теста, затем минимальная реализация. Нет «тесты потом». |

```
productTypeConfig (FORMAT_HINT_SHARED + formatHint per type)
        │
        ▼
FormatHint (toggle + sample + copy)
        │
        ▼
SourceInputCard
  FieldWrapper(label)
    ├── FormatHint          ← между label и редактором
    └── PlateListEditor
```

## Components and dependency order

| # | Компонент | Роль | Зависит от |
|---|-----------|------|------------|
| 1 | `productTypeConfig.ts` (+ test) | `FORMAT_HINT_SHARED`, `formatHint.explanation`, `sample` = `labels.placeholder` | — |
| 2 | `FormatHint.tsx` (+ test) | кнопка, раскрытие, образец, копирование | 1 (тексты как props) |
| 3 | `SourceInputCard.tsx` (+ test) | монтирует `FormatHint` между label и `PlateListEditor` | 1 + 2 |

**Sequential:** HINT-001 → HINT-002 → HINT-003 → HINT-004.  
**Parallel:** нет — общий конфиг и общие test files; `parallelSafe: false`.

## Risks

| Risk | Impact | Mitigation |
|------|--------|------------|
| Clipboard permission / API failure | Med | UI показывает «Не удалось скопировать»; образец остаётся читаемым в блоке; поле не меняется. Тест мокает `navigator.clipboard.writeText` (resolve + reject). |
| `FieldWrapper.hint` ставит текст **под** контрол | High | Не передавать `hint=`. Дети: `FormatHint` затем `PlateListEditor`. Проверка DOM-порядка в `SourceInputCard.test.tsx`. |
| Drift образца vs placeholder при ручном дубле `sample` | High | `sample` = `labels.placeholder` (getter или при сборке объекта). Тест: `formatHint.sample === labels.placeholder` для всех 7 типов. |

## Verification commands (canonical)

```bash
cd frontend && npm run test -- --run \
  src/features/commercial-offer/components/SourceInputCard.test.tsx \
  src/features/commercial-offer/components/FormatHint.test.tsx \
  src/features/commercial-offer/lib/productTypeConfig.test.ts

cd frontend && npm run typecheck
```

## Checkpoints

### After HINT-001
- [x] Config tests green: explanations non-empty; placeholders unchanged; sample ≡ placeholder
- [x] No UI yet — ok

### After HINT-002
- [x] FormatHint: collapsed → expand → collapse; copy does not need a field
- [x] Clipboard success/fail paths covered

### After HINT-003
- [x] Hint above editor for `plates` and `piles`; present when `sourceText` non-empty; click does not change editor value

### After HINT-004 (complete)
- [x] Canonical test command + typecheck green
- [x] Existing SourceInputCard / placeholder / lint tests still green
- [ ] Human UI spot-check optional; plan status stays until review accepts

## Task List

### Phase 1: Config texts

## HINT-001 — RED+GREEN: formatHint texts in productTypeConfig

**type:** `feat-fe` · **dependsOn:** [] · **pipeline:** explore → worker → test-writer → test-runner → reviewer  
**securitySensitive:** false · **needsExplore:** true · **parallelSafe:** false · **needsArchitectureReview:** false

**Description:** TDD. Сначала расширить `productTypeConfig.test.ts` падающими кейсами из Testing Strategy спеки: у каждого из 7 типов непустой `formatHint.explanation`; существующие exact placeholder-строки **не меняются**; `formatHint.sample === labels.placeholder`. Затем добавить `FORMAT_HINT_SHARED` и per-type `formatHint` с текстами из спеки «Тексты» (accepted as drafted).

**Acceptance:**
- [ ] `FORMAT_HINT_SHARED` = дословный каркас из спеки
- [ ] У всех 7 типов непустой `explanation` по таблице «Тексты»
- [ ] Exact placeholders из текущего теста без изменений
- [ ] `sample` не захардкожен отдельно от `labels.placeholder`

**Verify:**
```bash
cd frontend && npm run test -- --run src/features/commercial-offer/lib/productTypeConfig.test.ts
```

**Files:**
- `frontend/src/features/commercial-offer/lib/productTypeConfig.test.ts`
- `frontend/src/features/commercial-offer/lib/productTypeConfig.ts`

**Scope:** S

---

### Checkpoint: Config
- [ ] HINT-001 green
- [ ] Не начинать UI, пока sample drift не закрыт тестом

---

### Phase 2: FormatHint component

## HINT-002 — RED+GREEN: FormatHint collapse / expand / copy

**type:** `ui` · **dependsOn:** [HINT-001] · **pipeline:** explore → worker → test-writer → test-runner → reviewer  
**securitySensitive:** false · **needsExplore:** true · **parallelSafe:** false

**Description:** TDD. Создать `FormatHint.test.tsx` с падающими кейсами: старт свёрнут (раскрытый блок отсутствует в DOM); клик показывает shared + explanation + sample; второй клик прячет; `aria-expanded` совпадает; «Скопировать образец» вызывает `clipboard.writeText` ровно со `sample`; при mock-reject показывается ошибка копирования; компонент **не** принимает и не вызывает смену поля списка. Затем минимальный `FormatHint.tsx` (props: `explanation`, `sample`; shared из константы).

**Acceptance:**
- [ ] Старт: «▸ Подсказка», `aria-expanded={false}`, нет раскрытого блока
- [ ] Раскрытие: «▾ Подсказка», каркас + explanation + sample (`pre` / `pre-line`)
- [ ] Copy success → краткое «Скопировано», затем снова «Скопировать образец»
- [ ] Copy failure → «Не удалось скопировать»
- [ ] Нет связи с `onTextChange` / значением редактора

**Verify:**
```bash
cd frontend && npm run test -- --run \
  src/features/commercial-offer/components/FormatHint.test.tsx \
  src/features/commercial-offer/lib/productTypeConfig.test.ts
```

**Files:**
- `frontend/src/features/commercial-offer/components/FormatHint.test.tsx` (new)
- `frontend/src/features/commercial-offer/components/FormatHint.tsx` (new)
- (read-only import from `productTypeConfig.ts` for shared constant)

**Scope:** M (≤3 files written)

---

### Checkpoint: FormatHint
- [ ] HINT-002 green
- [ ] Clipboard failure path covered before mount

---

### Phase 3: Mount in SourceInputCard

## HINT-003 — RED+GREEN: FormatHint between label and PlateListEditor

**type:** `ui` · **dependsOn:** [HINT-002] · **pipeline:** explore → worker → test-writer → test-runner → reviewer  
**securitySensitive:** false · **needsExplore:** true · **parallelSafe:** false

**Description:** TDD. Расширить `SourceInputCard.test.tsx`: для `plates` и `piles` слово «Подсказка» есть над редактором (DOM: после label, до list editor); при непустом `sourceText` слово на месте; клик по подсказке / копированию **не** меняет значение редактора и не вызывает `onTextChange`. Затем монтировать `FormatHint` внутри `FieldWrapper` как первый child перед `PlateListEditor`, тексты из `PRODUCT_TYPE_CONFIG[productType]` (+ shared). Не использовать `hint=` у `FieldWrapper`.

**Acceptance:**
- [ ] Подсказка на `plates` и `piles` (прокси всех 7 — один SourceInputCard)
- [ ] Порядок: label → FormatHint → PlateListEditor
- [ ] Непустой `sourceText` не убирает «Подсказка»
- [ ] Toggle/copy не меняют `sourceText`
- [ ] Режим hasDraft / lint / «Дополнительно» без регрессий существующих тестов карточки

**Verify:**
```bash
cd frontend && npm run test -- --run \
  src/features/commercial-offer/components/SourceInputCard.test.tsx \
  src/features/commercial-offer/components/FormatHint.test.tsx \
  src/features/commercial-offer/lib/productTypeConfig.test.ts
```

**Files:**
- `frontend/src/features/commercial-offer/components/SourceInputCard.test.tsx`
- `frontend/src/features/commercial-offer/components/SourceInputCard.tsx`
- (uses existing `FormatHint.tsx`, `productTypeConfig.ts`)

**Scope:** M

---

### Checkpoint: Mount
- [ ] HINT-003 green
- [ ] Canonical three-file test command green

---

### Phase 4: Final verify

## HINT-004 — typecheck + full suite slice

**type:** `chore` · **dependsOn:** [HINT-003] · **pipeline:** test-runner → reviewer  
**securitySensitive:** false · **needsExplore:** false · **parallelSafe:** false

**Description:** Прогнать канонические команды. Починить только type/import ошибки, если всплывут в затронутых файлах. Код новой логики не добавлять.

**Acceptance:**
- [ ] Test command ниже exit 0
- [ ] `npm run typecheck` exit 0
- [ ] Нет изменений в `core/`, парсерах, backend

**Verify:**
```bash
cd frontend && npm run test -- --run \
  src/features/commercial-offer/components/SourceInputCard.test.tsx \
  src/features/commercial-offer/components/FormatHint.test.tsx \
  src/features/commercial-offer/lib/productTypeConfig.test.ts

cd frontend && npm run typecheck
```

**Files:** none planned (fix-only if typecheck fails on prior tasks)

**Scope:** XS

---

### Checkpoint: Complete
- [ ] All HINT-* green
- [ ] Plan remains «plan ready for review» until human accepts; then orchestrate execute
- [ ] Spec status → implemented **только после** кода и ревью (не сейчас)

## Not Doing (from spec)

- Шпаргалка, открытая по умолчанию
- Самооткрытие на неразобранной / красной строке
- Кнопка «вставить образец» в поле / автоподстановка
- Запоминание раскрытия между заходами
- Отдельные тексты для фото и OCR
- Админка текстов / правка без коммита в `productTypeConfig.ts`
- Разбор каждой битой строки своим абзацем
- Backend, parsers, admin, `PlateInputStep`/`SimpleProductInputStep` API changes
- Мобильный отдельный дизайн

## Open for plan review only

- Если ревьюер правит формулировки в «Тексты» — обновить только HINT-001 copy; структура задач не меняется.
- Иначе тексты из спеки идут в код as-is.

## Next Steps

**Checkpoint:** дождаться одобрения плана (типы + DAG + тексты), затем:

```text
/orchestrate execute orch-2026-09-25-16-38-kp-format-hint
```

Код до одобрения не писать. Не коммитить этот plan-only diff без явной просьбы.
