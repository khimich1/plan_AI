# Spec: КП — прозрачный picker типов в мульти-append

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-09-02  
**One-pager**: [ai_docs/ideas/multi-kp-transparent-type-picker.md](../ideas/multi-kp-transparent-type-picker.md)  
**Plan**: [ai_docs/develop/plans/2026-09-02-kp-multi-type-picker-transparency.md](../develop/plans/2026-09-02-kp-multi-type-picker-transparency.md)  
**Related**: [kp-multi-nomenclature-append.md](./kp-multi-nomenclature-append.md) (`start-append-cycle`, `ProductTypePicker`, `isPickingProductType`)

## Objective

**Проблема.** При «Добавить другое наименование» менеджер снова видит экран выбора типа как при **новом** КП: устаревший текст «в одном КП только один тип», нет следов уже забитых типов. Путаются все — непонятно, что уже в КП.

**Цель.** В режиме append сделать picker **прозрачным**: полоска «Уже в КП», разное поведение плиток (новые vs уже добавленные), (+) для дописи, (i) → боковой Drawer со списком уже добавленных изделий этого типа. Cold start тоже поправить copy.

**Пользователь:** менеджер в мастере КП на экране выбора типа (первый заход и каждый append-круг).

**Успех:** на новом круге сразу видно состав типов; (+) дописывает; (i) показывает позиции типа; недобавленный тип выбирается кликом по всей плитке; append API без изменений.

---

## ASSUMPTIONS I'M MAKING

1. **Только frontend UX.** Backend `POST .../append/start`, metadata `append_batches`, sticky client — as-is. Новых endpoint’ов нет.
2. **Источник «уже в КП».** Множество `product_type` из `draft.order_data` (и/или `append_batches`). Тип «уже добавлен», если есть ≥1 строка этого типа.
3. **Cold start vs append.** `mode: "create" | "append"` (или эквивалент `isPickingProductType` + наличие строк). Create: вся плитка кликабельна, без ✓ / (+) / (i) / полоски состава. Append: поведение по таблице ниже.
4. **Плитка «не добавлен».** Клик по всей карточке → `onSelect(type)` → существующий `startAppendCycle` / set product type.
5. **Плитка «уже добавлен».** Фон карточки **не** вызывает select. Две крупные кнопки: **(+)** → `onSelect(type)` (допись); **(i)** → открыть Drawer по этому типу.
6. **Drawer — read-only.** Список строк `order_data` с `product_type === selected`: наименование (`name` / `mark`) и количество (`qty`). Без edit/delete. Закрытие: Esc / overlay / кнопка закрытия у `Drawer`.
7. **Компонент Drawer.** Переиспользуем `frontend/src/shared/ui/Drawer.tsx` (side right).
8. **Полоска состава.** Над сеткой в append: «Уже в КП: Плиты · Сваи · …» (лейблы как в `PRODUCT_TYPE_LABELS`). Если после undo строк не осталось — режим как create (или пустая полоска + все плитки «недобавленные»).
9. **Copy / шапка.** Убрать «в одном КП только один тип». Create: заголовок «Добавить к коммерческому предложению» не нужен на старте — нейтральный «Создание коммерческого предложения» + подзаголовок выбора типа. **Append (2-й+ круг):** вместо H1 — **имя менеджера и заказчик** (из draft/sticky state); подзаголовок про дополнение типов. Отдельный Alert «Клиент · скидка» можно упростить/убрать дубль клиента, если шапка уже показывает заказчика.
10. **Drawer колонки.** Только **наименование + количество** (без цены/суммы строки).
11. **«К результату».** На append-picker обязательна; сбрасывает `isPickingProductType`, шаг `result`.
12. **Без новых npm/pip.** Иконки (+) и (i) — inline SVG в `button`.
13. **A11y.** `aria-label` на (+) и (i).
14. **Коммиты агент не делает**, пока явно не попросите.

→ **Assumptions + Q&A approved 2026-09-02.**

---

## Decisions locked (from ideation)

| # | Тема | Решение |
|---|------|---------|
| **D-pain** | Боль | На новом круге непонятно, что уже в КП — у всех менеджеров |
| **D-success** | Успех | Снять путаницу + быстрее собирать смешанное КП |
| **D-reselect** | Повтор типа | Разрешён; только **дополнение** списка, не replace |
| **D-context** | Контекст | **B+**: полоска «Уже в КП» + галочки / действия на плитках |
| **D-scope** | Объём | Малый UX на picker (+ Drawer read-only); не chips на Result |
| **D-tile-new** | Недобавленный тип | Вся плитка активна |
| **D-tile-added** | Добавленный тип | ✓ + крупный **(+)** + крупный **(i)**; фон не кликабелен как select |
| **D-plus** | (+) | Старт append-цикла этого типа |
| **D-info** | (i) | Боковая панель со уже добавленными изделиями типа |
| **D-no-api** | API | Без изменения append/export/save |

### Locked from Q&A 2026-09-02

| # | Тема | Решение |
|---|------|---------|
| **D-copy** | Create-заголовок | «Создание коммерческого предложения» + нейтральный подзаголовок (без «только один тип») |
| **D-append-header** | Шапка 2-го+ круга | **Вместо H1** — менеджер + заказчик; подзаголовок: выбрать тип для дополнения КП |
| **D-drawer-cols** | Drawer | Только наименование · кол-во (без суммы/цены строки) |
| **D-back** | Назад | «К результату» — да |
| **D-check** | Галочка | Индикатор в углу добавленной плитки |

---

## User Stories

- Как **менеджер**, нажав «Добавить другое наименование», я вижу полоску «Уже в КП: …» и не думаю, что КП сбросилось.
- Как **менеджер**, для типа, которого ещё нет, я кликаю по всей плитке и иду во ввод.
- Как **менеджер**, для типа, который уже есть, я жму **(+)** и дописываю ещё позиции этого типа.
- Как **менеджер**, жму **(i)** и в боковой панели вижу, какие изделия этого типа уже в КП.
- Как **менеджер**, на первом создании КП по-прежнему выбираю тип одной плиткой; текст больше не врёт про «только один тип».

---

## Tech Stack

| Слой | Стек |
|------|------|
| Frontend | React 19, TypeScript, Vite, Vitest + Testing Library |
| UI | существующий `Card`, `Drawer`, wizard store |
| Backend | без изменений |

Новых пакетов нет.

## Commands

```
cd frontend && npm run test -- src/features/commercial-offer/components/ProductTypePicker
cd frontend && npm run test -- src/features/commercial-offer
cd frontend && npm run typecheck
```

Dev: не убивать уже запущенный `./run+logs.sh`.

Backend pytest не обязателен (нет API-изменений); при регрессии wizard:

```
pytest tests/test_commercial_multi_append_flow.py -q
```

## Project Structure

```
frontend/src/features/commercial-offer/
  components/ProductTypePicker.tsx          → mode, selectedTypes, lines, (+) / (i), полоска
  components/ProductTypePicker.test.tsx     → create vs append, (+) / (i), Drawer
  components/ProductTypeLinesDrawer.tsx     → optional extract: read-only list (или внутри picker)
  components/CommercialOfferWizard.tsx      → props из draft + «К результату»
  store/wizardDraftStore.tsx                → при необходимости action cancel-append-pick → result
shared/ui/Drawer.tsx                        → reuse
ai_docs/ideas/multi-kp-transparent-type-picker.md
ai_docs/specs/kp-multi-type-picker-transparency.md
```

## Code Style

Пример контракта props:

```tsx
type ProductTypePickerProps = {
  mode?: "create" | "append"; // default "create"
  selectedProductTypes?: ReadonlyArray<ProductType>;
  orderLines?: ReadonlyArray<Record<string, unknown>>;
  managerName?: string;
  clientName?: string;
  onSelect: (productType: ProductType) => void;
  onBackToResult?: () => void; // только append
};
```

- Добавленная плитка: не `<button>` на весь Card; карточка-`div` + два `button` (+) / (i).
- Недобавленная: как сейчас, один `button` на карточку.
- Иконки — `aria-label`, крупный hit-area (≥40×40 CSS px).

## Testing Strategy

| Уровень | Что |
|---------|-----|
| RTL `ProductTypePicker` | create: клик плитки → `onSelect`; нет полоски «Уже в КП»; нет (+) / (i) |
| RTL append | полоска с лейблами; у selected — (+) и (i), клик фона не зовёт `onSelect`; (+) → `onSelect`; (i) → Drawer с name/qty строк типа |
| RTL | тип без строк после фильтра — поведение create-плитки |
| Regress | существующие тесты create (`steps` / `marches` / …) зелёные |
| Optional | wizard: «К результату» снимает `isPickingProductType` и ставит step `result` |

## Boundaries

- **Always:** только draft-данные на клиенте; не ломать cold-start select; a11y labels на (+) / (i).
- **Ask first:** edit/delete из Drawer; chips на Result; смена API append.
- **Never:** новые зависимости ради иконок; запрет повторного выбора типа; копипаст второго Drawer-компонента вместо `shared/ui/Drawer`.

## Success Criteria

| # | Критерий |
|---|----------|
| S1 | Append-picker показывает «Уже в КП» со всеми типами, у которых есть строки |
| S2 | Недобавленный тип: клик по плитке → `onSelect` |
| S3 | Добавленный тип: (+) → `onSelect`; клик по фону карточки → нет `onSelect` |
| S4 | (i) открывает Drawer со строками только этого `product_type` (name/mark + qty) |
| S5 | Copy не содержит «в одном КП только один тип» |
| S6 | Create-mode: нет ✓ / (+) / (i) / полоски; select по плитке работает |
| S7 | Append-шапка показывает менеджера и заказчика вместо «Создание КП» |
| S8 | «К результату» возвращает на result без потери draft |
| S9 | Vitest по picker зелёный; typecheck зелёный |

## Out of Scope

- Выбор типа с экрана Result (chips)
- Мультивыбор типов за один заход
- Редактирование состава в Drawer
- Цены / суммы в Drawer
- Счётчики «N позиций» на карточках
- Любые изменения PDF / XLSX / save / undo-batch API

## Open Questions

_Нет — закрыты 2026-09-02._

---

**Next:** Done (IMPLEMENT ✅).
