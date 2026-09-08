# Конструктор КП — шаг 3: слияние input-шагов и preview-панелей

Дата: 2026-09-08. Основание: план
[`ai_docs/develop/plans/2026-09-08-kp-constructor-merge-input-steps.md`](../plans/2026-09-08-kp-constructor-merge-input-steps.md).
Предыдущий шаг: [`kp-constructor-product-type-config-2026-09-07.md`](./kp-constructor-product-type-config-2026-09-07.md).

Пять простых изделий (piles/steps/marches/bridge_piles/fbs) обслуживаются одним параметризованным
input-шагом и одной preview-панелью; плиты остаются отдельной веткой. REST-контракты не менялись.

## Что сделано

| Инкремент | Коммит | Суть |
|---|---|---|
| 1 | `0af1d52` | `labels` (12 полей) в `productTypeConfig`; 6 input-шагов, `WizardProgress`, `ProductTypePicker` — на lookup |
| 2 | `9628bd7` | `SimpleProductInputStep` вместо 5 клонов input-шага; визард рендерит 2 ветки (simple/plates) |
| 3 | `4add703` | `KpGradedPreviewPanel` вместо 5 клонов панели; `labels.previewSubtitle/previewUnpricedMessage`; реестр `productTypePreview.ts` подключён |
| 4 | `228d6a7` | `handleLineGradeChange` — lookup `getProductTypePreview` вместо 4 if-веток; удалены `isMarchFlow/isBridgePileFlow/isFbsFlow` |
| 5 (опц.) | `b2d7b65` | `CalculationResultStep` — один `draftProductType` вместо 5 булевых `is*Draft`; хук экспортирует resolved тип |

Форма `preview` от плана §3 уточнена: не поле конфига, а соседний модуль
`lib/productTypePreview.ts` (grade-каталоги + `buildRows/buildLines` на изделие). Причина — грабля
циклического импорта из §5: конфиг остаётся чистым data-модулем, направление импорта одностороннее
(`productTypePreview` → `productTypeConfig`, не наоборот).

## Метрики

Ориентир плана: ~−2500 строк дублей. Факт по `frontend/src` (diff `aa10dd1..HEAD`):

| Срез | Diff |
|---|---|
| Production (без `*.test.*`) | **+590 / −3334 (−2744)** |
| С тестами | **+1072 / −3635 (−2563)**, 27 файлов |

Размеры после:

| Файл | Строк (было → стало) |
|---|---|
| `CommercialOfferWizard.tsx` | 1638 → 1427 |
| 5 input-шагов + 5 preview-панелей | ~3200 → `SimpleProductInputStep` 504 + `KpGradedPreviewPanel` 229 |
| `productTypeConfig.ts` | 129 → 260 |
| `productTypePreview.ts` | — → 108 |

## Приёмка §6

- [x] `npm run typecheck` — чисто после каждого инкремента
- [x] `npx vitest run src/features/commercial-offer` — **55 файлов / 415 тестов**, зелёные
  (было 57 / 410 на handoff: −3 файла клонов, +1 `KpGradedPreviewPanel.test.tsx`; +2 теста конфига,
  +2 теста `available_grades`, +1 «steps без grade-UI даже с колбэками»)
- [x] Grep `isFbsFlow|isPileFlow|isMarchFlow|isStepFlow|isBridgePileFlow` по `src/` — **пусто**
- [x] Grep `PileInputStep|MarchInputStep|FbsInputStep|BridgePileInputStep|StepInputStep` по `src/` — **пусто**
- [x] `canFinishPiles` — **пусто** (ушло вместе с клонами шага)
- [x] Визард рендерит ровно 2 input-ветки: `SimpleProductInputStep` + `PlateInputStep`
- [x] Diff суммарно отрицательный, порядок тысяч строк
- [x] Этот отчёт

## Поведение (сохранено)

- Плейсхолдеры ФБС/мостовых свай — **дословно** свайные (`С120.35-12 B25 5\nС120.35-13и 3`),
  зафиксированы пин-тестами конфига (§5: осознанный copy-paste до решения заказчика)
- Per-row `available_grades` у ФБС/мостовых свай: селект строки ограничен списком строки,
  fallback на глобальный каталог только при отсутствии списка — покрыто двумя новыми тестами
- Sealed-семантика: bulk-контрол только при незапечатанных строках; re-ingest грейдов шлёт
  только unsealed-линии с `append`/`replace` как раньше
- Steps: нет колонки «Класс», нет bulk-блока, `minWidth 560`, row-key `mark-index` — дословно
  как в старой `KpStepPreviewPanel` (ветвление по `supportsGrades`, не по `productType`)
- Remount при смене типа изделия: `key={productType}` на панели восстанавливает семантику
  разных компонентов (состояние `bulkGrade` не протекает между изделиями)
- Fallback неизвестного типа → `plates`; `CalculationResultStep` получает resolved
  `draftProductType` с тем же `?? state.productType`, что давал хук — legacy-черновики
  без `metadata.product_type` рендерятся как раньше

Зафиксированные эквивалентности (недостижимые пути):

- `handleLineGradeChange` для steps/plates: колбэк передаётся только при `supportsGrades` —
  недостижимо и до, и после. Раньше steps упали бы в pile-ветку, теперь — в steps-запись реестра
- Реестр для steps: `gradeCodes: []`, `isGradeCode: () => false`, `buildLines` без sealed-фильтра —
  симметрия реестра, за `supportsGrades` не вызывается; тест «нет grade-UI даже с колбэками» сторожит
- Панель piles/marches теперь латентно читает `row.available_grades` — их билдеры поле не ставят,
  fallback идентичен старому «всегда глобальный каталог»

## Doubt-ревью инкремента 3

Адверсариальное ревью слияния панелей ([Reviewer](e51309ea-9332-4a10-a38b-208f45a9389f)),
артефакт — diff + контракт поведения. Кросс-модель пропущена по решению заказчика.

6 находок, классификация:

1. **Actionable:** потеря remount при смене `productType` (застойный `bulkGrade`) — исправлено `key={productType}`
2. **Actionable:** row-key steps `mark--index` вместо `mark-index` — исправлено дословным ветвлением
3. Trade-off: plates-значения `previewSubtitle/previewUnpricedMessage` в `labels` не читает
   `KpPlatePreviewPanel` (панель вне скоупа шага; поля — полнота `Record` и SSOT на будущий проход)
4. Trade-off: латентный `available_grades` у piles/marches (см. выше)
5. Шум на момент ревью: «мёртвые» `buildLines`/`getProductTypePreview` — подключены в инкременте 4
6. Шум: «footgun» пустого каталога steps — сторожится тестом и конфиг-пинами

## Не трогали (шаг 4+ плана ревью)

- `CommercialOfferWizard` как god-component (расчленение — шаг 4)
- `PlateInputStep` / `KpPlatePreviewPanel` — слияние с simple-веткой не планировалось
- `wizardDraftStore`, backend, REST-контракты
- Тесты уровня визарда (`handleLineGradeChange` и ко.) — не-цель, шаг 6
- `test_kp.pdf` (удалён в дереве до начала шага, вне задачи; в коммиты не включался)

## Верификация

```bash
cd frontend && npm run typecheck
# чисто

cd frontend && npx vitest run src/features/commercial-offer
# Test Files  55 passed (55)
# Tests       415 passed (415)
```

## Статус

**Выполнен** (2026-09-08), включая опциональный инкремент 5. Следующий шаг ревью: шаг 4 —
расчленение `CommercialOfferWizard` (god-component, ~1400 строк).
