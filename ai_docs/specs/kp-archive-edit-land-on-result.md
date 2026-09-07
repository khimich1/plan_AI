# Spec: Архив КП — «Редактировать» сразу на шаг 3 (result)

**Статус**: SPECIFY ✅ · PLAN 🔄  
**Дата**: 2026-09-07  
**Тип**: hotfix (не redesign мастера)  
**Plan:** [../develop/plans/2026-09-07-kp-archive-edit-land-on-result.md](../develop/plans/2026-09-07-kp-archive-edit-land-on-result.md)  
**Related**: [kp-archive-edit-in-constructor.md](./kp-archive-edit-in-constructor.md) (родительская фича, IMPLEMENT ✅), [kp-multi-nomenclature-append.md](./kp-multi-nomenclature-append.md), [kp-archive-only-save.md](./kp-archive-only-save.md)

## Objective

**Проблема.** Из карточки архивного КП кнопка **«Редактировать»** должна открывать конструктор на шаге **3 / result** (скидка и итоги, состав уже загружен). Сейчас она вызывает `openInConstructor("result")`, делает `resume` + `hydrate-draft` и уходит на `/new?draft=…`, но **не ставит шаг мастера на `result`**. Локальный store после hydrate остаётся на шаге ввода (для плит — `plates`). Менеджер оказывается на шаге 1. Кнопка **«Готово, далее»** там «мёртвая»: UI на вводе, серверный wizard уже на `result`, `can_proceed_to` не даёт идти вперёд, сайдбар «3. Результат» тоже недоступен.

**Цель.** После **«Редактировать»** конструктор открывается на шаге `result` с данными КП (позиции, скидка, итоги). **«(+ Добавить)»** по-прежнему ведёт в цикл дописи (picker → ввод). Обе кнопки остаются. Глобальный merge шага при hydrate **не** меняем.

**Пользователь:** менеджер в Archive → drawer КП «в архиве» → конструктор.

**Успех:** клик «Редактировать» → шаг 3, скидка правится, состав/итоги на месте; пользователь не застревает на шаге 1.

---

## ASSUMPTIONS I'M MAKING (locked)

1. **Hotfix only.** Не переделываем мастер КП, не меняем `mergeWizardStepWithServer` глобально, не разрешаем hydrate промотировать `plates` → `result`.
2. **Две CTA как сейчас.** «Редактировать» → result. «(+ Добавить)» → `start-append-cycle` (picker / ввод). Без chooser-диалога, без инлайн-правки скидки в drawer.
3. **Точка фикса.** После успешного `resume` + `hydrate-draft` для `landing === "result"` явно `dispatch({ type: "set-step", step: "result" })` — тот же приём, что `handleUndoLastBatch` в `CommercialOfferWizard.tsx` (hydrate, затем force `result`).
4. **Вторичный симптом** («Готово, далее» не ведёт дальше) исчезает, когда landing верный. Переписывать `PlateInputStep` / `handleFinishPlates` **не** нужно, пока тест и ручной smoke это подтверждают.
5. **Типы изделий.** `set-step: result` общий для всех шести типов. Pin в merge срабатывает на любом input-step (`plates` / `piles` / `steps` / `marches` / `bridge_piles` / `fbs`).
6. **Без новых npm/pip.** Коммиты — только по просьбе. Не убивать `./run+logs.sh`.

→ Locked для PLAN (направление продукта уже решено).

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| **D-hotfix** | Объём | Только landing «Редактировать» → `result`. Не redesign мастера |
| **D-merge** | `mergeWizardStepWithServer` | Не трогать. Тест «hydrate с plates не прыгает на result» остаётся зелёным |
| **D-force** | Как ставить шаг | После hydrate: `dispatch({ type: "set-step", step: "result" })` только при `landing === "result"` |
| **D-append** | «(+ Добавить)» | Как сейчас: hydrate + `start-append-cycle`. **Без** `set-step: result` |
| **D-cta** | Кнопки | Обе остаются, подписи не меняем |
| **D-drawer** | Скидка в карточке | Не возвращаем редактор скидки в drawer |
| **D-plates** | `PlateInputStep` | Не переписывать как safety net в этом hotfix |

---

## Problem — current vs desired

### Точки входа (UI)

| Место | Элемент | `landing` | Ожидаемый экран |
|-------|---------|-----------|-----------------|
| Archive → drawer КП **«в архиве»** → секция **«Итоги»** | **«Редактировать»** | `"result"` | Конструктор, шаг **3. Результат** (`currentStep === "result"`), `CalculationResultStep` |
| Там же | **«(+ Добавить)»** | `"append"` | Picker номенклатуры / шаг 1 ввода (`start-append-cycle`) |

Оба зовут `openInConstructor` в `OfferDetailsDrawer.tsx`. Resume: `POST /api/v1/commercial/archive/{kp_id}/resume`. Навигация: `/new?draft={draft_id}`.

Шаги мастера (`WizardProgress`): **1.** ввод (`plates` или аналог) → **2. Клиент** (`client`, часто skip при resume) → **3. Результат** (`result`). Для пользователя «шаг 3» = `result` / скидка и итоги.

### Что происходит сегодня

```
«Редактировать»
  → archiveApi.resume(kp_id)
  → dispatch({ type: "hydrate-draft", payload: draft })   // wizard_state.current_step обычно "result"
  → landing === "result"  ⇒  только НЕ вызывать start-append-cycle
  → navigate /new?draft=…
```

Дыра: **нет** `{ type: "set-step", step: "result" }`.

`hydrate-draft` в `wizardDraftStore.tsx` считает шаг так:

```ts
currentStep: mergeWizardStepWithServer(
  state.currentStep,
  action.payload.wizard_state?.current_step,
  productType,
)
```

`mergeWizardStepWithServer`: если локальный шаг = product input (`plates` / `piles` / …) и сервер впереди — **оставить локальный**. Это защита живого конструктора: hydrate после «Обработать» не должен сам перекидывать на `result`. Store по умолчанию и после `draftStorage.load()` часто на `plates`. Итог archive-resume: **локально шаг 1, на сервере уже result**.

Дальше `useCommercialOfferWizard` при появлении `draftQuery.data` снова делает `hydrate-draft`. Второй hydrate безопасен **только если** локальный шаг уже `result` (merge берёт max / pin только на input-step). Поэтому force-step должен быть **в drawer до `navigate`**, в том же handler, сразу после hydrate.

### Почему «Готово, далее» мёртвая (вторичный симптом)

На шаге 1 кнопка зовёт `handleFinishPlates` → `draft.wizard_state.can_proceed_to[0]`. Resume-draft уже на `result`, массив пустой → ошибка / no-op, перехода нет.

`canNavigateToStep` в `CommercialOfferWizard.tsx`: шаг ввода всегда разрешён (`step === inputStep`); переход на `result` — только если индекс меньше серверного **или** `can_proceed_to` содержит `result`. При local=`plates`, server=`result` клик по сайдбару «3. Результат» **тоже** обычно недоступен. Пользователь заперт на шаге 1.

Это не отдельный баг `PlateInputStep`: при верном landing кнопки на шаге 1 нет на экране.

### Желаемое поведение

| Действие | После успеха |
|----------|----------------|
| **«Редактировать»** | hydrate → **`set-step: result`** → `/new?draft=…`. UI = `CalculationResultStep`. Скидка редактируется там. Состав и итоги из resume-draft. `start-append-cycle` **не** вызывается |
| **«(+ Добавить)»** | hydrate → `start-append-cycle` → picker. **Нет** `set-step: result` |
| Resume error | Сообщение в drawer, без dispatch шага, без navigate, CTA снова активны |
| Resume pending | Обе CTA disabled, текст «Открываем…» |

Существующий паттерн force-result после hydrate:

```1142:1144:frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx
      const draft = await undoLastAppendBatchMutation.mutateAsync(state.draftId);
      dispatch({ type: "hydrate-draft", payload: draft });
      dispatch({ type: "set-step", step: "result" });
```

Целевой код drawer (смысл, не обязательно байт-в-байт):

```ts
async function openInConstructor(landing: "append" | "result") {
  // … resume …
  dispatch({ type: "hydrate-draft", payload: draft });
  if (landing === "append") {
    dispatch({ type: "start-append-cycle" });
  } else {
    // hydrate не поднимает input-step → result (намеренно для живого мастера)
    dispatch({ type: "set-step", step: "result" });
  }
  navigate(`/new?draft=${encodeURIComponent(draft.draft_id)}`);
  onClose();
}
```

---

## User Stories

- Как **менеджер**, в drawer архивного КП жму **«Редактировать»** и сразу вижу шаг 3 с составом, суммами и полем скидки — не шаг ввода плит/свай/…
- Как **менеджер**, жму **«(+ Добавить)»** и по-прежнему попадаю в выбор номенклатуры, чтобы дописать позиции.
- Как **менеджер**, если resume не удался, остаюсь в drawer с ошибкой и могу повторить.

---

## Tech Stack

| Слой | Стек |
|------|------|
| Frontend | React 19, TypeScript, Vite, Vitest + Testing Library, React Router |
| Store | `wizardDraftStore` (`hydrate-draft`, `set-step`, `start-append-cycle`) |
| API | Существующий `POST /api/v1/commercial/archive/{kp_id}/resume` — **без** изменений контракта |

Backend не трогаем.

## Commands

```
cd frontend && npm run test -- src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx
cd frontend && npm run test -- src/features/commercial-offer/store/wizardDraftStore.test.tsx
cd frontend && npm run typecheck
```

Dev: не убивать `./run+logs.sh`. Ручной smoke: Archive → КП «в архиве» → «Редактировать» / «(+ Добавить)».

## Project Structure

```
frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx
  → openInConstructor: при landing "result" после hydrate — set-step result
frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx
  → «Редактировать» должна assert set-step result (порядок после hydrate)
frontend/src/features/commercial-offer/store/wizardDraftStore.tsx
  → mergeWizardStepWithServer — НЕ менять
frontend/src/features/commercial-offer/store/wizardDraftStore.test.tsx
  → «не поднимает шаг с plates до result при hydrate-draft» — НЕ инвертировать
frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx
  → canNavigateToStep / handleFinishPlates — не переписывать в этом hotfix
```

## Code Style

- Один handler `openInConstructor(landing)`, без копипасты двух resume.
- `set-step: result` **только** в ветке `landing === "result"`.
- Комментарий *why*: hydrate намеренно не промоутит input → result; archive edit обязан форсировать шаг.
- Русские сообщения ошибок — как сейчас (`getErrorMessage`).

## Testing Strategy

Проект: Vitest + Testing Library. TDD: сначала красный тест drawer, потом код.

| Уровень | Что менять |
|---------|------------|
| RTL `OfferDetailsDrawer.test.tsx` | **«Редактировать»**: после hydrate есть `{ type: "set-step", step: "result" }`; индекс `set-step` > индекса `hydrate-draft`; **нет** `start-append-cycle`; navigate + `onClose` как сейчас |
| RTL тот же файл | **«(+ Добавить)»**: hydrate + `start-append-cycle`; **нет** `set-step` с `result` |
| RTL тот же файл | Resume fail / pending — без регрессий (уже есть) |
| `wizardDraftStore.test.tsx` | Тест «hydrate с plates **не** прыгает на result» **оставить**. Не добавлять в hydrate глобальный promote |
| `PlateInputStep` / wizard E2E | Не обязаны в этом hotfix, если landing-тест зелёный |

`makeResumeDraft` уже отдаёт `wizard_state.current_step: "result"` — этого достаточно: тест ловит отсутствие `set-step`, а не «плохой» payload.

## Boundaries

- **Always:** landing «Редактировать» = `result`; «(+ Добавить)» = append cycle; merge hydrate как был; focused vitest зелёные.
- **Ask first:** менять `mergeWizardStepWithServer`; править `PlateInputStep` / `canNavigateToStep` как «страховку»; трогать backend resume.
- **Never:** redesign мастера; chooser-диалог; скидка в drawer; переименовать кнопки; отдельная страница «правка КП»; новые зависимости; коммит без просьбы; убивать `./run+logs.sh`; инвертировать hydrate-merge тест.

## Success Criteria

| # | Критерий |
|---|----------|
| S1 | «Редактировать» → dispatch `hydrate-draft`, затем `set-step: result`, navigate `/new?draft=…`, drawer закрыт |
| S2 | После этого `currentStep === "result"`: `CalculationResultStep`, скидка редактируется, состав/итоги из draft |
| S3 | Нет ловушки на шаге 1; «Готово, далее» не нужна на этом пути |
| S4 | «(+ Добавить)» = hydrate + `start-append-cycle`, без `set-step: result` |
| S5 | Обе CTA на месте; drawer «Итоги» по-прежнему read-only |
| S6 | Тест hydrate-merge (plates не → result) зелёный |
| S7 | Resume error / pending без регрессий |
| S8 | `mergeWizardStepWithServer` и глобальный hydrate не ослаблены |

## Edge cases

| Кейс | Ожидание |
|------|----------|
| **Resume failure** | Ошибка в drawer, `dispatch` не зовётся (как сейчас), нет navigate / `onClose` |
| **Pending resume** | Обе CTA → «Открываем…», disabled; второй клик не стартует параллельный resume (`resumePending`) |
| **Пустой состав** | Всё равно шаг `result` (пустая таблица / нулевые итоги). Не редирект на шаг 1 |
| **Другой product type** | Input-step другой (`piles` …), pin тот же. Force `result` общий. «(+ Добавить)» по-прежнему picker |
| **Локальный мастер на другом шаге** | `set-step: result` после hydrate перекрывает `client` / leftover `plates`. Если локально уже `result`, вызов идемпотентен |
| **Повторный hydrate из `draftQuery`** | После force local=`result`, merge(`result`, server `result`) = `result`. Поэтому `set-step` **до** `navigate` |
| **Leftover `sourceText` в localStorage** | `hydrate-draft` **не** чистит `sourceText`. На `result` `PlateInputStep` не монтируется — текст не виден. Чистить поле в этом hotfix **не** нужно. «(+ Добавить)» / append с result уже чистит через `start-append-cycle` |
| **Серверный GET draft с `current_step=plates`** | После force local=`result`: merge не pin'ит (local ≠ input), max → `result` |
| **`skipClient`** | Сайдбар может быть 1 + 3; целевой шаг всё равно `result` |

## Out of Scope

- Redesign мастера КП
- Глобально разрешить hydrate `plates` → `result`
- Правка скидки / рейса внутри карточки архива
- Переименование кнопок, третий CTA, chooser
- Отдельная страница «редактирование КП»
- Переписывание `PlateInputStep` / `handleFinishPlates` / `canNavigateToStep` как основной фикс
- Backend resume / статус-гейт «в архиве» (уже сделано 2026-09-02)
- Новые npm/pip

## Open Questions

_Нет блокирующих._ Продуктовое направление locked. Если после force-step «Готово, далее» всё ещё воспроизводится — тогда отдельный follow-up, не этот hotfix.

---

## Историческая ошибка родительской спеки

[kp-archive-edit-in-constructor.md](./kp-archive-edit-in-constructor.md) / ARC-005: «result → только hydrate (metadata уже `current_step=result`)». Это **неверно** при живом `mergeWizardStepWithServer`: серверный `result` не поднимает локальный input-step. Этот hotfix закрывает дыру, не отменяя остальную фичу (две CTA, read-only Итоги, save в тот же `kp_id`).
