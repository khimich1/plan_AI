# Конструктор КП — шаг 3: слияние input-шагов и preview-панелей

**Статус:** план-хэндофф для выполнения в новом окне.
**Дата:** 2026-09-08. **Основание:** полное ревью конструктора КП (2026-09-07) + шаг 2
(`ai_docs/develop/reports/kp-constructor-product-type-config-2026-09-07.md`).

---

## 1. Контекст для исполнителя

Модуль: `frontend/src/features/commercial-offer`. Стек: React 19, TS, TanStack Query, vitest.
Команды из `frontend/`: `npm run typecheck`, `npx vitest run src/features/commercial-offer`.

HEAD на момент хэндоффа: `aa10dd1` (отчёт шага 2). В дереве может лежать несвязанное `D test_kp.pdf` — не трогать.

**Шаг 2 уже сделал:** `lib/productTypeConfig.ts` — SSOT по изделию (`endpointSegment`, `batchesField`,
`supportsGrades`, `isSimpleKp`, `ingestAction`, `labels.nounPlural` / `nounGenitivePlural`).
Хук: 3 мутации (`updateInputMutation` / `applyAiMutation` / `updateGradesMutation`).
Визард: один `handleFinishInput`, сообщения из конфига. Тесты: 57 файлов / 406, typecheck чисто.

**Что осталось копипастой** (замерено в ревью 2026-09-07, размеры 2026-09-08):

| Что | Файлы | Строк | Идентичность |
|---|---|---|---|
| Input-шаги simple KP | `Pile/March/Fbs/BridgePile/StepInputStep.tsx` | ~490 × 5 | ~85–95% (diff 44–100 строк) |
| Input-шаг плит | `PlateInputStep.tsx` | 643 | тот же каркас + wide/unpriced/invalid секции |
| Preview с классом бетона | `KpPile/March/Fbs/BridgePilePreviewPanel.tsx` | ~220 × 4 | ~90%+; отличаются каталог грейдов и `build*PreviewRows` |
| Preview ступеней | `KpStepPreviewPanel.tsx` | 130 | тот же каркас без грейдов |
| Preview плит | `KpPlatePreviewPanel.tsx` | 226 | другие колонки + флаги wide/invalid |
| Рендер-блоки в визарде | `CommercialOfferWizard.tsx` ~1227–1460 | 6 × ~50 | ~95%; отличаются `onFinish*` имя и панель |
| `handleLineGradeChange` | тот же визард ~605–650 | ~50 | if по `isFbsFlow` / билдерам строк |

Итого ориентир ревью: **~−2500 строк**. `PlateInputStep` и `KpPlatePreviewPanel` **не** сливать с simple-веткой
вслепую: у плит уникальные resolve-гейты.

Конфиг сейчас **тонкий** — UI-копи (placeholder, stepTitle, …) в шагах всё ещё захардкожены.
`WizardProgress.stepTitles` и `ProductTypePicker.PRODUCT_TYPE_LABELS` — leftover шага 2.

## 2. Цель шага 3

Параметризованные компоненты вместо пер-издельских клонов:

- один `SimpleProductInputStep` на 5 simple-изделий;
- один `KpGradedPreviewPanel` на 4 панели с классом бетона;
- `KpStepPreviewPanel` либо тот же graded с `supportsGrades: false`, либо остаётся тонким (решить по diff);
- `PlateInputStep` / `KpPlatePreviewPanel` остаются отдельными (или слот/children для resolve-секций — не форсировать);
- визард рендерит **2 блока** (plates vs simple), не 6;
- `handleLineGradeChange` — lookup билдеров из конфига, флаги `isFbsFlow` уходят;
- опечатка `canFinishPiles` в `FbsInputStep.tsx:197` / `BridgePileInputStep.tsx:197` чинится сама при слиянии
  (`canFinish` / `canFinishInput`).

**Non-goals (не трогать в этом шаге):**

- Расчленение `CommercialOfferWizard` на `useLineUndo` / `usePlateResolutions` / `useAppendCycle` — шаг 4.
- Sync-кейсы `wizardDraftStore` — шаг 4.
- Типизация `order_data` (discriminated union) — шаг 5.
- Тесты на уровне визарда с нуля — шаг 6 (существующие тесты шагов/панелей переписать на новые имена — ок).
- REST-контракты, хук-мутации, `handleFinishInput`, редьюсер.
- Не «улучшать» UX и не править копипаст-плейсхолдеры, пока не решите явно (см. §5).

## 3. Предлагаемая форма

Расширить **существующий** `productTypeConfig.ts`, не плодить второй реестр.

```ts
// добавить в ProductTypeConfig.labels (сейчас только nounPlural / nounGenitivePlural):
labels: {
  nounPlural: string;
  nounGenitivePlural: string;
  stepTitle: string;        // "Шаг 1. Сваи"
  listLabel: string;        // "Список свай для расчёта"
  placeholder: string;      // точные текущие строки, включая косяки копипасты
  emptySubtitle: string;
  aiHint: string;
};

// опционально, когда дойдёте до preview / grade-change:
preview: {
  supportsGrades: boolean;          // уже есть на корне конфига — не дублировать
  gradeCodes: readonly string[];    // [] для plates/steps
  formatGradeLabel: (code: string) => string;
  buildRows: (draft: CommercialDraftDetails) => GradedPreviewRow[];
  buildLines: (rows: GradedPreviewRow[]) => string;
};
```

Точную форму `preview` исполнитель уточняет **после** замера diff панелей (инкремент 1b / doubt).
Критерий: в `SimpleProductInputStep` и `KpGradedPreviewPanel` нет `if (productType === ...)`.
Каталоги грейдов уже есть: `lib/pileGrades.ts`, `marchGrades.ts` (алиас pile), `fbsGrades.ts`
(B7_5…B25 + особый `formatFbsGradeLabel`), `bridgePileGrades.ts` (B25/B30).
У FBS в панели есть ветка `available_grades` на строке — сохранить.

`WizardProgress`: `"1. " + cfg.labels.nounPlural`. Client/result — как сейчас («2. Клиент», «3. Результат»).
`ProductTypePicker`: `title` из `nounPlural`; `description`/`emoji` в пикере оставить локально
(это не дубль шагов, это карточки выбора).

Визард после слияния:

```tsx
state.currentStep === "client" ? <ClientConditionsStep ... />
  : state.currentStep === "result" ? <CalculationResultStep ... />
  : productConfig.isSimpleKp ? <SimpleProductInputStep productType={productType} ... />
  : <PlateInputStep ... />
```

`onFinishPiles` / `onFinishFbs` / … → один `onFinishInput={handleFinishInput}`.

## 4. Инкременты (каждый — отдельный коммит, дерево зелёное)

1. **Labels в конфиг, без слияния файлов.** Добавить `stepTitle` / `listLabel` / `placeholder` /
   `emptySubtitle` / `aiHint` (точные текущие строки). Юнит-тесты полноты Record.
   Переключить 6 InputStep + `WizardProgress` + `ProductTypePicker` title на lookup.
   `MarchInputStep.test.tsx` остаётся зелёным (`MARCH_LIST_PLACEHOLDER` можно реэкспортировать из конфига).
2. **`SimpleProductInputStep`.** Вынести из `PileInputStep`, параметр `productType`.
   Переключить piles → marches → fbs/bridge → steps. Удалить 5 старых файлов.
   Тесты: `MarchInputStep.test.tsx` переписать на `SimpleProductInputStep` + `productType="marches"`
   с теми же ассертами. Визард: 5 simple-блоков → один (plates-блок пока отдельно).
3. **`KpGradedPreviewPanel`.** Замерить diff четырёх панелей (не верить ревью вслепую: тогда
   `KpMarch` vs `KpPile` казался большим — перепроверить). Слить 4 graded; steps — если diff
   маленький, тот же компонент с `supportsGrades: false`. `KpPlatePreviewPanel` не трогать.
   Существующие `KpPilePreviewPanel.test.tsx` / unparsed-тест — перецепить.
4. **`handleLineGradeChange` + добивание флагов.** Билдеры строк в конфиг (или рядом
   `productTypePreview.ts`, если тащить функции в конфиг-модуль неудобно из-за циклов).
   Удалить `isFbsFlow` / `isBridgePileFlow` / `isMarchFlow`. Grep этих имён — пусто
   (кроме исторических комментов).
5. **(Опционально, если остаётся запас)** `CalculationResultStep`: убрать 5 булевых `is*Draft`
   в пользу `getProductTypeConfig(draft.metadata.product_type)` (`supportsGrades` / `isSimpleKp` /
   `labels`). Хук перестаёт экспортировать `isPileDraft`…. Не смешивать с инкрементом 2.

Ориентир: −1500…−2500 production-строк. Если инкремент >500 строк правок — кодмод/один механический PR,
не ручная правка шести копий.

## 5. Грабли

- **Плейсхолдеры ФБС и мостовых свай — копия свай** (`"С120.35-12 B25 5\nС120.35-13и 3"` в
  `FbsInputStep` / `BridgePileInputStep`). Это баг копипасты, не домен. **По умолчанию сохранить
  дословно** (шаг 2 так и делал, пока заказчик не сказал иначе). Если правите — отдельный коммит
  и явное «да» в чате; для ФБС канонический пример из отчёта внедрения: `ФБС 9.3.6-Т 2`.
- **`canFinishPiles` в FBS/bridge** — не «сохранять имя». При слиянии назвать `canFinish`.
- **Plate vs simple — разные пропсы.** У `PlateInputStep` есть wide/unpriced/invalid. Не протаскивать
  их в `SimpleProductInputStep` «на всякий случай».
- **FBS `available_grades` на строке** — не выкинуть при слиянии панелей (у свай список глобальный).
- **Циклические импорты:** `productTypeConfig` не должен импортировать `buildPilePreviewRows` и т.п.
  Если `preview.buildRows` нужен — вынести в `lib/productTypePreview.ts`, который импортирует конфиг,
  не наоборот.
- **Референциальная стабильность** не актуальна для JSX-компонентов. Для `gradeCodes` — `as const`
  массивы на уровне модуля, как сейчас.
- **Существующие тесты не менять по ассертам.** Исключение — arrange, привязанный к удалённому
  имени файла/пропу `onFinishPiles`. Переписать на новое имя, те же ожидания.
- **`isPending` уже общий** (шаг 2, §5 того плана). Не возвращать пер-издельский pending.
- Не трогать `generateFilesMutation` `onSettled`, типы `CommercialWizardState`.

## 6. Приёмка

- [ ] `npm run typecheck` — чисто после каждого инкремента
- [ ] `npx vitest run src/features/commercial-offer` — все тесты зелёные; существующие ассерты поведения
      не ослаблять
- [ ] Grep `isFbsFlow|isPileFlow|isMarchFlow|isStepFlow|isBridgePileFlow` — **пусто** в `frontend/src`
- [ ] Grep `PileInputStep|MarchInputStep|FbsInputStep|BridgePileInputStep|StepInputStep` — только
      git-история / этот план / отчёт (файлов нет)
- [ ] Grep `canFinishPiles` — пусто
- [ ] Визард рендерит не больше двух input-веток (simple / plates)
- [ ] Diff production отрицательный, порядка сотен–тысяч строк
- [ ] Отчёт в `ai_docs/develop/reports/` (конвенция как у шага 2)

Поведение визуально то же: те же тексты, те же кнопки, те же грейд-дропдауны.
Браузер не обязателен (нет нового UI); если сомнение по вёрстке — один проход simple (сваи) + plates.

## 7. Рекомендованные скиллы (прочитать в новом окне до старта)

**Обязательные:**

| Скилл | Зачем |
|---|---|
| `plan-web-context` | Стек, пути, команды, конвенции ai_docs |
| `refactor-workflow` | Analyze → Refactor → Verify → Docs; тесты уже есть |
| `incremental-implementation` | 4–5 коммитов, keep-it-green, «заметил — не трогай» |
| `code-simplification` | Дубли → параметризация; правило 500 строк |
| `test-driven-development` | Сначала тесты новых label-полей конфига, потом слияние |
| `code-review-and-quality` | Само-ревью по 5 осям перед отчётом |

**Ситуативные:**

| Скилл | Когда |
|---|---|
| `doubt-driven-development` | Перед инкрементом 2 (граница Plate vs Simple) и перед инкрементом 3 (сливать ли March/Step в graded) |
| `frontend-ui-engineering` | Только если после слияния поедет вёрстка; нового UI нет |
| `git-workflow-and-versioning` | Атомарные коммиты |
| `documentation-and-adrs` | Отчёт в `ai_docs/develop/reports` |

**Не нужны:** `api-and-interface-design`, `security-*`, `orchestration`, `browser-testing-with-devtools`
(пока нет визуальных сомнений), `deprecation-and-migration` (внутренние файлы — удалять, не deprecate).

## 8. Стартовый промпт для нового окна

```
Выполни план ai_docs/develop/plans/2026-09-08-kp-constructor-merge-input-steps.md
(шаг 3 ревью конструктора КП: слияние input-шагов и preview-панелей).
Прочти скиллы из §7 плана.
Работай инкрементами по §4, после каждого — typecheck + тесты модуля.
Поведение не менять (плейсхолдеры ФБС/мостовых свай — дословно, см. §5).
По завершении — отчёт в ai_docs/develop/reports/.
```
