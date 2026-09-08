# Конструктор КП — шаг 2: конфиг по изделиям (productTypeConfig)

**Статус:** выполнен (2026-09-07). Отчёт: `ai_docs/develop/reports/kp-constructor-product-type-config-2026-09-07.md`.
**Дата:** 2026-09-07. **Основание:** полное ревью конструктора КП + фикс-пак 1
(`ai_docs/develop/reports/kp-constructor-fixpack-1-2026-09-07.md`).

---

## 1. Контекст для исполнителя

Модуль: `frontend/src/features/commercial-offer` (~14 300 строк без тестов, 60+ файлов).
Стек: React 19, TS, TanStack Query, vitest. Команды из `frontend/`:
`npm run typecheck`, `npx vitest run src/features/commercial-offer`.

Каждое изделие (plates, piles, steps, marches, bridge_piles, fbs) сейчас размножает код
копипастой. Карта дублей (замерено нормализованным diff на 2026-09-07):

| Где | Что | Объём |
|---|---|---|
| `hooks/useCommercialOfferWizard.ts` | 6 update-мутаций + 6 applyAi-мутаций + 4 grades-мутаций, идентичны кроме URL/типа | ~250 строк |
| `api/commercialOfferApi.ts` | 6 `updateDraft*` + 6 `applyAi*` + 4 `update*Grades`, отличается сегмент URL | ~200 строк |
| `components/CommercialOfferWizard.tsx` | вложенные тернарники выбора мутаций (стр. ~153–174); 6 хендлеров `handleFinish*` (~150 строк); тернарники сообщений в `handleRecognize`, `handleApplyAi`, `handleSidebarStepClick` | ~250 строк |
| `lib/wizardStepOrder.ts` | if-цепочки `getProductInputStep`, `fullWizardStepOrder`, `isSimpleKpProductType`, `resolveDraftProductType`; 6 констант `*_WIZARD_STEP_ORDER` | ~100 строк |
| `components/steps/CalculationResultStep.tsx` | локальный `PRODUCT_TYPE_LABELS` | ~15 строк |

Уже сделано (фикс-пак 1, в рабочем дереве на момент хэндоффа): `isProceeding={calculateMutation.isPending}`
в 6 шагах, bounds-check в `handleLineGradeChange`, try/catch в `draftStorage`.
Текущее состояние тестов: 56 файлов / 373 теста — зелёные, `tsc --noEmit` чисто.

## 2. Цель шага 2

Единый источник правды об изделии: `Record<ProductType, ProductTypeConfig>`.
Убрать пер-издельскую копипасту из **хука, API, визарда (хендлеры/тернарники), wizardStepOrder**.

**Non-goals (не трогать в этом шаге):**
- 6 step-компонентов (`PileInputStep`, `MarchInputStep`, `FbsInputStep`, `BridgePileInputStep`,
  `StepInputStep`) и 5 preview-панелей — это шаг 3 (слияние в параметризованные компоненты).
  Визард продолжает рендерить 6 блоков; меняются только источники мутаций и тексты.
- `CommercialOfferWizard.tsx` как god-component (расчленение — шаг 4).
- REST-контракты бэкенда — не меняются ни в одном месте.
- Редьюсер `wizardDraftStore` — пер-издельской логики там нет, не трогаем.
- Опечатку `canFinishPiles` в `FbsInputStep.tsx:197` / `BridgePileInputStep.tsx:197` — поправится в шаге 3.

## 3. Предлагаемая форма конфига

Новый файл `lib/productTypeConfig.ts` (+ `productTypeConfig.test.ts`):

```ts
export type ProductTypeConfig = {
  productType: ProductType;
  inputStep: WizardStepId;               // сейчас совпадает с productType
  endpointSegment: "plates" | "piles" | "steps" | "marches" | "bridge-piles" | "fbs";
  supportsGrades: boolean;               // piles/marches/bridge_piles/fbs
  isSimpleKp: boolean;                   // всё кроме plates
  ingestAction: WizardNextRequiredAction; // "ingest_plates" | "ingest_piles" | ...
  labels: {
    nounGenitivePlural: string;  // "свай" | "плит" | ... (для сообщений)
    listLabel: string;           // "Список свай"
    stepTitle: string;           // "Шаг 1. Сваи"
    placeholder: string;
    emptySubtitle: string;
    aiHint: string;
    unitName: string;            // "ступеней" — строка готовности в результате
  };
};
export const PRODUCT_TYPE_CONFIG: Record<ProductType, ProductTypeConfig>;
export const getProductTypeConfig = (type: ProductType | null | undefined): ProductTypeConfig; // fallback → plates
```

Точную форму полей исполнитель уточняет по фактическому потреблению — критерий: все
тернарники/иф-цепочки по изделию заменяются lookup'ом, новых тернарников не появляется.
Сообщения из `handleRecognize`/`handleApplyAi`/`handleSidebarStepClick` и `handleFinish*`
собираются из `labels` + общей карты `nextRequiredAction → текст`
(plates-специфичные `resolve_*`-ветки остаются plates-only через эту карту, не if'ами по изделию).

## 4. Инкременты (каждый — отдельный коммит, дерево зелёное после каждого)

1. **Config + тесты.** Новый файл, юнит-тесты: полнота `Record` (каждый `ProductType` покрыт),
   корректность `endpointSegment`, `getProductTypeConfig(undefined) === plates`.
   Потребителей пока не переключать.
2. **API-слой.** Добавить обобщённые `updateDraftInput(draftId, productType, payload)`,
   `applyAiInstruction(draftId, productType, payload)`, `updateGrades(draftId, productType, grade)`
   (сегмент URL из конфига). Переключить хук на них, удалить 16 старых методов.
   Сохранить семантику onSuccess каждого семейства (applyAi диспатчит `start-batch-review` — см. §5).
3. **Хук.** 16 мутаций → 3 фабричные/обобщённые. Публичный интерфейс хука меняется —
   единственный потребитель `CommercialOfferWizard.tsx`, переключается в этом же коммите.
4. **Визард.** Удалить тернарники выбора мутаций; 6 `handleFinish*` → один `handleFinishInput`;
   сообщения из конфига. Рендер-блоки шагов пока остаются (пропсы `isConfirmingBatch` и т.п.
   переписываются на обобщённые мутации — см. §5 про эквивалентность isPending).
5. **wizardStepOrder.** if-цепочки → lookup в конфиг; 6 констант порядков шагов свернуть
   к производным от конфига. Проверить использование устаревшего алиаса `WIZARD_STEP_ORDER`:
   если только тесты — удалить, иначе оставить с `@deprecated`.

Ориентир по итогу: −400…−600 строк, без изменения поведения.

## 5. Грабли (найдено при ревью — учесть обязательно)

- **onSuccess семейств различается:** `update*` не диспатчат, `applyAi*` диспатчат
  `start-batch-review`, `update*Grades` диспатчат `hydrate-draft` с `refreshBatchText: true`.
  Обобщённые мутации должны сохранить это поведение (параметр onSuccess-стратегии или
  отдельные хуки-обёртки).
- **isPending-эквивалентность:** сейчас шаг получает `isPending` своей пер-издельской мутации.
  После обобщения `updateInputMutation.isPending` охватывает все изделия. Эквивалент сохраняется,
  т.к. одновременно активно одно изделие; но при append-цикле смена типа во время pending
  покажет pending на новом типе — приемлемо, зафиксировать в отчёте.
- **`resolveDraftProductType` мапит неизвестное → "plates"** — `getProductTypeConfig` обязан
  сохранить этот fallback, иначе сломаются старые драфты/тесты.
- **`isSimpleKpDraft`** гейтит `breakdownQuery` и ветки в `CalculationResultStep` —
  семантика «всё кроме plates» не должна измениться.
- `generateFilesMutation`/`generateSchemaMutation` используют `onSettled`, а не `onSuccess` — не трогать.
- `CommercialWizardState`/`WizardNextRequiredAction` синхронизированы с
  `app.schemas.commercial` — типы не менять.

## 6. Приёмка

- [x] `npm run typecheck` — чисто после каждого инкремента
- [x] `npx vitest run src/features/commercial-offer` — все тесты зелёные **без модификации
      существующих тестов** (правка тестов = сигнал смены поведения; исключение — тесты хука/API,
      если их arrange привязан к удалённым методам: тогда тест переписывается на обобщённый метод
      с теми же ассертами поведения)
- [x] Ни одного тернарника/if-цепочки по `ProductType` вне `productTypeConfig.ts`
      (grep по `isFbsFlow|isPileFlow|isMarchFlow|isStepFlow|isBridgePileFlow` — только определение
      флагов для рендер-блоков, которые уйдут в шаге 3)
- [x] Diff суммарно отрицательный по строкам
- [x] Отчёт в `ai_docs/develop/reports/` по конвенции проекта

## 7. Рекомендованные скиллы (прочитать в новом окне до старта)

**Обязательные:**

| Скилл | Зачем |
|---|---|
| `plan-web-context` | Стек, пути, команды, конвенции ai_docs — читать первым |
| `refactor-workflow` | Это чистый рефакторинг: цикл Analyze → Refactor → Verify → Docs, правило «тесты до рефакторинга» (уже выполнено), сайзинг скоупа 6–10 файлов → разбить на сессии = наши инкременты |
| `incremental-implementation` | 4–5 инкрементов-коммитов, keep-it-green, scope discipline («заметил — не трогай») |
| `code-simplification` | Прямое попадание: nested ternaries → lookup objects, preserve behavior exactly, правило 500 строк (у нас ~650 — рассмотреть sed/кодмод для механических замен в хуке/API) |
| `test-driven-development` | Сначала тесты конфига (инкремент 1), потом переключение потребителей |
| `code-review-and-quality` | Финальное само-ревью по 5 осям + чек-лист перед отчётом |

**Ситуативные:**

| Скилл | Когда |
|---|---|
| `doubt-driven-development` | Перед инкрементом 2 — свежим взглядом проверить форму конфига (поля labels, стратегия onSuccess), дальше переделывать дороже |
| `git-workflow-and-versioning` | Формат атомарных коммитов по инкрементам |
| `documentation-and-adrs` | Короткий отчёт по завершении (конвенция `ai_docs/develop/reports`) |

**Не нужны:** `frontend-ui-engineering` (нет нового UI), `api-and-interface-design`
(REST-контракты не меняются), `security-*`, `performance-optimization`,
`browser-testing-with-devtools` (поведение не меняется), `deprecation-and-migration`
(внутренний модуль), `orchestration` (избыточен для одного шага плана).

## 8. Стартовый промпт для нового окна

```
Выполни план ai_docs/develop/plans/2026-09-07-kp-constructor-product-type-config.md
(шаг 2 ревью конструктора КП: productTypeConfig). Прочти скиллы из §7 плана.
Работай инкрементами по §4, после каждого — typecheck + тесты модуля.
Поведение не менять. По завершении — отчёт в ai_docs/develop/reports/.
```
