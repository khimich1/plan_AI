# Implementation Plan: КП — прозрачный picker типов в мульти-append

**Спека**: [ai_docs/specs/kp-multi-type-picker-transparency.md](../../specs/kp-multi-type-picker-transparency.md)  
**Идея**: [ai_docs/ideas/multi-kp-transparent-type-picker.md](../../ideas/multi-kp-transparent-type-picker.md)  
**Related**: [kp-multi-nomenclature-append.md](../../specs/kp-multi-nomenclature-append.md)  
**Дата**: 2026-09-02  
**Статус**: PLAN ✅ · IMPLEMENT ✅

## Overview

Только frontend: `ProductTypePicker` в режиме append показывает полоску «Уже в КП», шапку менеджер+заказчик, для уже добавленных типов — ✓ + (+) + (i) с read-only `Drawer` (наименование · qty). Недобавленные типы — клик по всей плитке. «К результату» сбрасывает picking. API append без изменений.

## Architecture Decisions

- **Данные только из draft на клиенте.** `selectedProductTypes` = unique `product_type` из `order_data`. Строки Drawer фильтруем по типу. Нет новых endpoint’ов.
- **Два вида плитки.** Unselected: один `<button>` на карточку (как сейчас). Selected: `div`-карточка + два крупных `button` (+) / (i); клик по фону не вызывает `onSelect`.
- **Drawer.** `shared/ui/Drawer` right; колонки name/mark + qty; без цен.
- **Шапка append.** H1 заменяем на блок «Менеджер · Заказчик»; подзаголовок про выбор типа для дополнения. Дублирующий Alert «Клиент: …» в wizard при append убрать (скидку при желании оставить одной строкой или в подзаголовке — минимально: не дублировать клиента).
- **Cancel pick.** Новый action `cancel-append-pick` (или эквивалент): `isPickingProductType: false`, `currentStep: "result"`. `set-step` сам по себе picking не снимает — не полагаться только на него.
- **Иконки.** Inline SVG, hit-area ≥40px, `aria-label`. Без npm-icon packs.

## Task List

### Phase 1 — ProductTypePicker UI

#### Task 1: Props + create-mode copy + append strip/header

**Description:** Расширить `ProductTypePicker`: `mode`, `selectedProductTypes`, `orderLines`, `managerName`, `clientName`, `onBackToResult`. Create: убрать текст «только один тип», нейтральный подзаголовок. Append: шапка менеджер+заказчик, полоска «Уже в КП», кнопка «К результату».

**Acceptance:**
- [x] S5, S6, S7 (шапка), часть S1 (полоска)
- [x] Create без полоски / без менеджерской шапки
- [x] «К результату» зовёт `onBackToResult`

**Verification:** обновить/добавить RTL в `ProductTypePicker.test.tsx`  
**Dependencies:** None  
**Files:** `ProductTypePicker.tsx`, `ProductTypePicker.test.tsx`

#### Task 2: Плитки (+) / (i) + Drawer

**Description:** Для типов из `selectedProductTypes` — ✓, (+) → `onSelect`, (i) → Drawer со строками типа (name/mark + qty). Фон selected-плитки не вызывает select. Unselected — вся плитка как сейчас.

**Acceptance:**
- [x] S1–S4
- [x] Пустой список в Drawer, если строк нет (не должно открываться на unselected)
- [x] Существующие create-тесты зелёные

**Verification:** RTL (+) / (i) / Drawer / no select on background  
**Dependencies:** Task 1  
**Files:** `ProductTypePicker.tsx` (± `ProductTypeLinesDrawer.tsx`), `ProductTypePicker.test.tsx`, reuse `shared/ui/Drawer.tsx`

### Phase 2 — Wizard wiring

#### Task 3: Wire CommercialOfferWizard + store cancel

**Description:** При `isPickingProductType` передать в picker `mode="append"`, selected types и lines из draft, manager/client names, `onBackToResult`. Добавить `cancel-append-pick` в store (clear picking → result). Убрать дубль Alert клиента на append-экране (или оставить только скидку, если нужна).

**Acceptance:**
- [x] S8
- [x] Append select по-прежнему идёт в `handleAppendProductTypeSelect` / `startAppendCycle`
- [x] Create path (если picker используется вне append) не сломан

**Verification:** при наличии — store/wizard тест на cancel; иначе ручная проверка + typecheck  
**Dependencies:** Task 1–2  
**Files:** `CommercialOfferWizard.tsx`, `wizardDraftStore.tsx`, optionally `wizardDraftStore.test.tsx`

### Phase 3 — Verify

#### Task 4: Tests + typecheck

**Description:** Прогнать тесты commercial-offer и typecheck.

**Acceptance:**
- [x] S9

**Verification:**
```
cd frontend && npm run test -- src/features/commercial-offer/components/ProductTypePicker
cd frontend && npm run typecheck
```
**Dependencies:** Task 1–3

## Risks

| Risk | Mitigation |
|------|------------|
| `set-step` не снимает `isPickingProductType` | Явный `cancel-append-pick` |
| Имя менеджера только в draft metadata | Брать `lastDraft.metadata.manager_name` / managers lookup по `managerId` |
| Create-тесты ищут role button по всей карточке | Selected-плитки меняют структуру — не ломать create-кейсы |

## Out of Scope (plan)

Chips на Result, edit в Drawer, суммы/цены, API changes.
