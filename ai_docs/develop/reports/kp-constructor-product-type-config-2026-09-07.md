# Конструктор КП — шаг 2: productTypeConfig

Дата: 2026-09-07. Основание: план
[`ai_docs/develop/plans/2026-09-07-kp-constructor-product-type-config.md`](../plans/2026-09-07-kp-constructor-product-type-config.md).
Ревью + фикс-пак 1: [`kp-constructor-fixpack-1-2026-09-07.md`](./kp-constructor-fixpack-1-2026-09-07.md).

Рефакторинг без смены REST-контрактов и без слияния step/preview-компонентов (это шаг 3).

## Что сделано

Единый источник правды об изделии: `frontend/src/features/commercial-offer/lib/productTypeConfig.ts`
(`Record<ProductType, ProductTypeConfig>` + `getProductTypeConfig` с fallback на `plates`).

| Инкремент | Коммит | Суть |
|---|---|---|
| 1 | `97459a7` + `733e811` | Конфиг + юнит-тесты; правки по doubt-ревью (`hasOwnProperty`, `RESOLVE_GATE_MESSAGES`) |
| 2 | `71a6604` | 16 API-методов → `updateDraftInput` / `applyAiInstruction` / `updateGrades` |
| 3 | `8cace03` | 16 мутаций хука → 3; визард переключён, `productType` в variables |
| 4 | `e905055` | 6 `handleFinish*` → `handleFinishInput`; сообщения из `labels` |
| 5 | `09f6959` | `wizardStepOrder` / `getBatches` / `getDraftBatchCount` → lookup |

## Метрики

Ориентир плана: −400…−600 строк. Факт по `frontend/src/features/commercial-offer`:

| Срез | Diff |
|---|---|
| Production (без `*.test.ts`) | **+290 / −632 (−342)** |
| С тестами конфига и API | **+491 / −633 (−142)** |

Потребители (хук/API/визард/wizardStepOrder) сжались сильнее; обратно добавляют новый конфиг (~130 строк) и его тесты (~136).

Размеры после:

| Файл | Строк (было → стало) |
|---|---|
| `useCommercialOfferWizard.ts` | 517 → 364 |
| `commercialOfferApi.ts` | 374 → 289 |
| `CommercialOfferWizard.tsx` | 1823 → 1638 |
| `wizardStepOrder.ts` | 143 → 128 |
| `productTypeConfig.ts` | — → 129 |

## Приёмка §6

- [x] `npm run typecheck` — чисто после каждого инкремента
- [x] `npx vitest run src/features/commercial-offer` — **57 файлов / 406 тестов**, все зелёные
  (было 56 / 373; +15 тестов конфига, +18 тестов generic API). Существующие тесты **не модифицировались**
- [x] Тернарники/if-цепочки по `ProductType` вне конфига: grep `isFbsFlow|isPileFlow|isMarchFlow|isStepFlow|isBridgePileFlow` — остались только флаги для `handleLineGradeChange` (выбор preview-билдера; уйдёт в шаге 3) и рендер-блоки 6 шагов
- [x] Diff суммарно отрицательный
- [x] Этот отчёт

## Поведение (сохранено / осознанно изменено)

Сохранено:

- REST URL (в т.ч. `bridge_piles` → `/bridge-piles`); `productType` **не** попадает в тело PATCH
- onSuccess семейств: `update*` — cache+invalidate; `applyAi*` — `start-batch-review`; `update*Grades` — `hydrate-draft` + `refreshBatchText: true`
- Fallback неизвестного типа → `plates` (`getProductTypeConfig` / `resolveDraftProductType`)
- `isSimpleKp` = всё кроме plates → `breakdownQuery` гейтится как раньше
- `handleFinishInput`: `validation_errors` (сервер) → ingest → resolve-гейты → fallback. На пустом драфте по-прежнему «Список плит пустой.» с бэкенда, не клиентский ingest-текст
- Порядок шагов — те же module-level массивы (стабильная ссылка, `useEffect` визарда не зацикливается)

Осознанно изменено (согласовано с заказчиком):

- No-draft сообщения AI и сайдбара для **ФБС / мостовых свай** больше не падают в «список плит». Теперь: «список ФБС» / «список мостовых свай»

Зафиксированные эквивалентности (не регрессии на достижимых путях):

- **isPending:** один `updateInputMutation.isPending` на все изделия. Одновременно активно одно; при append-смене типа во время pending спиннер покажется на новом шаге. Приемлемо по §5 плана
- **`updateGrades`:** generic метод строит `/plates/grades` и `/steps/grades`, если его вызвать. UI грейдов рендерится только у piles/marches/bridge_piles/fbs — путь недостижим. Раньше тернарник по умолчанию бил в `/piles/grades`

## Doubt-ревью формы конфига

Перед инкрементом 2 — адверсариальное ревью ([Reviewer](7d4963de-aa71-4691-9515-d14b80221d54)): 3 major / 7 minor. Кросс-модель пропущена по решению заказчика.

Учтено в коде:

1. `getProductTypeConfig`: `hasOwnProperty`, не `in` (ключ `"constructor"` не обходит fallback)
2. Карта сообщений сужена до `RESOLVE_GATE_MESSAGES`; ingest идёт через `cfg.ingestAction`
3. Порядки шагов — prebuilt module-level arrays, не `[]` на каждый вызов
4. `handleFinishInput` сохраняет приоритет `validation_errors`

Trade-off (задокументирован, не чинился): `resolve_*` в карте видны любому изделию; бэкенд выставляет их только при непустых plate-гейтах. На недостижимых путях текст мог бы стать plates-gate вместо «проверьте список свай» — инвариант BE.

## Не трогали (шаги 3–4 плана ревью)

- 6 input-step компонентов и 5 preview-панелей
- `handleLineGradeChange` — if по билдерам preview (`buildFbsPreviewRows` и т.д.)
- `CommercialOfferWizard` как god-component (расчленение — шаг 4)
- `wizardDraftStore`
- опечатка `canFinishPiles` в `FbsInputStep` / `BridgePileInputStep`
- REST бэкенда, `CommercialWizardState` / `WizardNextRequiredAction`

## Замечено, не трогали

- `ProductTypePicker` и `WizardProgress` всё ещё держат локальные `PRODUCT_TYPE_LABELS` / `stepTitles` — естественный потребитель `labels.nounPlural` в шаге 3
- `getDraftProductType` в `batchReview.ts` больше не используется `getBatches` (хелпер оставлен, публичный)
- `CalculationResultStep` по-прежнему принимает 5 булевых `is*Draft` — контракт шага, не этого рефакторинга

## Верификация

```bash
cd frontend && npm run typecheck
# чисто

cd frontend && npx vitest run src/features/commercial-offer
# Test Files  57 passed (57)
# Tests       406 passed (406)
```

## Статус

**Выполнен** (2026-09-07). Следующий шаг ревью: слияние input-шагов и preview-панелей (~−2500 строк дублей).
