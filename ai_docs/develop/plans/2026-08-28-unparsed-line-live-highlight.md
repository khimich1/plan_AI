# Implementation Plan: Живая подсветка нераспознанных строк в поле ввода КП

**Спека**: [ai_docs/specs/unparsed-line-live-highlight.md](../../specs/unparsed-line-live-highlight.md)  
**Идея**: [ai_docs/ideas/unparsed-line-live-highlight.md](../../ideas/unparsed-line-live-highlight.md)  
**Дата**: 28.08.2026  
**Статус**: план выполнен 28.08.2026 (T1–T9; браузер вручную не прогнан)

## Overview

Менеджер видит нераспознанные строки в том же поле, куда печатает. Проверка фоном через расширение `POST /commercial/parse`. Кнопка обработки/добавления серая, пока линт в полёте или есть красные. Мёртвая плашка с предпросмотра и шага 3 уходит. Шесть типов изделий, один хук и одна карточка источника.

## Architecture Decisions

- **Линт — отдельный модуль**, не поле в `ProductDraftSpec`. P1-config и так перегружен create/update/AI; диспетчер `product_type → parse_*_line` живёт в `app/services/commercial_line_lint.py`. Preview/ILP/DraftStore не вызываем.
- **Плиты на `/parse` — два прохода.** Старые поля по-прежнему из `CommercialService.parse` (`parse_plate_text`, без ILP). Новое `lines` — из построчного линта (`parse_line` + `validate_plate_values`). Расхождение «нормализатор всего текста vs строка» принимаем по спеке.
- **Остальные типы** на `/parse` отдают только `product_type`, `lines`, `unparsed_lines`. Без фейкового `order`.
- **Фронт:** `useSourceTextLint` + `SourceInputCard` + оверлей `PlateListEditor` с внешней картой подсветки. Шаги тонкие.
- **Не трогаем** `WizardNextRequiredAction`, «Список верен», грамматику парсеров, слэш-формат.

## Dependency Graph

```
commercial_line_lint (построчный разбор)
    │
    ├── schemas + POST /parse (compat плит + lines)
    │       │
    │       └── commercialOfferApi.parseSource
    │               │
    │               └── useSourceTextLint
    │                       │
    │                       ├── PlateListEditor (внешние highlights)
    │                       │       │
    │                       └── SourceInputCard + гейт кнопки
    │                               │
    │                               └── шесть *InputStep
    │
    └── filter unparsed warnings → Kp*PreviewPanel + CalculationResultStep
        (независимо от карточки, после или параллельно с волной UI)
```

Мёртвые плашки не зависят от линта по данным — можно снять после того, как гейт не пускает красное в текстовый submit. Порядок: сначала гейт, потом убрать плашки, чтобы не оставить дыру.

## Task List

### Phase 1: Backend lint + `/parse`

#### Task 1: Построчный lint-сервис

**Description:** Чистая функция/сервис: текст + `product_type` → список `LineLint` (index, text, empty, ok, reason_text). Плиты: `parse_line` + `validate_plate_values`. Остальные: `parse_*_line`. Пустые строки — `empty=True`, `ok=True`.

**Acceptance:**
- [x] Шесть типов: хотя бы одна ok и одна not ok
- [x] Слэш `ПБ 40,3/2,6-8п` → ok=false
- [x] Нет вызовов preview/optimize/DraftStore

**Verification:** `pytest tests/test_commercial_line_lint.py -q`

**Dependencies:** None  
**Files:** `app/services/commercial_line_lint.py`, `tests/test_commercial_line_lint.py`  
**Estimated scope:** S

#### Task 2: Расширить `POST /commercial/parse`

**Description:** `CommercialParseRequest.product_type` optional default `plates`. Ответ: для плит — прежние ключи + `product_type` + `lines`; для других — `product_type`, `lines`, `unparsed_lines`. Невалидный тип → 422. Роутер тонкий.

**Acceptance:**
- [x] Без `product_type` плитный текст — 200, старые ключи на месте, есть `lines`
- [x] Каждый product_type покрыт HTTP-тестом (ok + not ok)
- [x] RBAC без cookie как у текущего `/parse`
- [x] Существующие тесты `test_commercial_parse_*` зелёные

**Verification:** `pytest tests/test_commercial_web_flow.py tests/test_commercial_parse_lint.py tests/test_rbac_server_side.py -q -k parse`

**Dependencies:** Task 1  
**Files:** `app/schemas/commercial.py`, `app/api/v1/endpoints/commercial.py`, `tests/test_commercial_parse_lint.py` (+ точечная правка web_flow только если сломается compat)  
**Estimated scope:** M

### Checkpoint: Backend

- [x] `pytest tests/test_commercial_line_lint.py tests/test_commercial_parse_lint.py` + parse/RBAC `-k parse` зелёные
- [ ] Полный safety net `test_commercial_web_flow.py` — 3 падения identity/schema вне скоупа линта
- [x] Ревью контракта `lines` до фронта (по желанию)

### Phase 2: Фронт — линт в поле

#### Task 3: Клиент `/parse`

**Description:** Тип ответа + `commercialOfferApi.parseSource({ text, productType })`.

**Acceptance:**
- [x] Тест клиента: POST JSON с `product_type`, читает `lines`

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/api/commercialOfferApi.test.ts`

**Dependencies:** Task 2 (контракт)  
**Files:** `frontend/src/features/commercial-offer/api/commercialOfferApi.ts`, `frontend/src/features/commercial-offer/types/commercialOffer.ts`, `frontend/src/features/commercial-offer/api/commercialOfferApi.test.ts`  
**Estimated scope:** S

#### Task 4: Хук `useSourceTextLint`

**Description:** Дебаунс 500 ms, сразу `isPending` при изменении текста, AbortController/seq против stale, пустой текст и `enabled=false` — без fetch.

**Acceptance:**
- [x] Пустой текст — 0 запросов, не pending
- [x] Быстрый ввод — один запрос после паузы, старый ответ отброшен
- [x] `enabled=false` (фото) — нет запросов

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/hooks/useSourceTextLint.test.ts`

**Dependencies:** Task 3  
**Files:** `frontend/src/features/commercial-offer/hooks/useSourceTextLint.ts`, `frontend/src/features/commercial-offer/hooks/useSourceTextLint.test.ts`  
**Estimated scope:** S

#### Task 5: Оверлей без обязательного draft

**Description:** `PlateListEditor` принимает опциональную карту подсветки (index → unparsed + title). Если карты нет — текущее поведение от `draft`. Режим линта не требует полного черновика.

**Acceptance:**
- [x] Тест: внешняя карта красит строку «не попало в расчёт» / title с причиной
- [x] Старые тесты нумерации и draft-highlights зелёные

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/components/PlateListEditor.test.tsx`

**Dependencies:** None (можно параллельно с 3–4)  
**Files:** `frontend/src/features/commercial-offer/components/PlateListEditor.tsx`, `frontend/src/features/commercial-offer/components/PlateListEditor.test.tsx`, при необходимости `lib/plateLineHighlights.ts`  
**Estimated scope:** S

#### Task 6: `SourceInputCard` + гейт на шаге плит

**Description:** Вынести карточку источника из `PlateInputStep` в общий компонент. Подключить хук + оверлей. «Обработать текст» / «Добавить к списку»: disabled при pending или красных; `title` «Проверка списка…» / «Исправьте красные строки». Подпись кнопки не менять. Файл без текста — линт выключен.

**Acceptance:**
- [x] Тест: красные → кнопка disabled + title про исправление
- [x] Тест: pending → disabled + title про проверку
- [x] Тест: все непустые ok → кнопка живая (если есть текст)
- [x] Тест: только файл → линт не серит

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/components/SourceInputCard.test.tsx src/features/commercial-offer/components/steps/PlateInputStep.tsx`

**Dependencies:** Task 4, Task 5  
**Files:** `SourceInputCard.tsx`, `SourceInputCard.test.tsx`, `steps/PlateInputStep.tsx`  
**Estimated scope:** M

#### Task 7: Остальные пять шагов ввода

**Description:** Та же карточка в Pile/Step/March/Fbs/BridgePile `*InputStep`. Только `productType` и подписи полей отличаются.

**Acceptance:**
- [x] В каждом шаге больше нет локального голого `Textarea` источника (карточка общая)
- [x] `productType` прокинут верный (сваи не зовут plates)

**Verification:** `cd frontend && npm run test && npm run typecheck`

**Dependencies:** Task 6  
**Files:** пять `*InputStep.tsx`  
**Estimated scope:** M (ровно 5 файлов; поведение уже покрыто карточкой)

### Checkpoint: Линт в UI

- [x] Frontend test + typecheck зелёные
- [ ] Вручную (если dev поднят): ломаная строка на плитах — красная, кнопка серая; исправить — живая

### Phase 3: Убрать мёртвую плашку

#### Task 8: Хелпер + плиты + шаг 3

**Description:** Хелпер отфильтровывает warning «Не удалось распознать строк: N» (и эквивалент шага 3). `KpPlatePreviewPanel` не рендерит «Не попали в состав». `CalculationResultStep` не добавляет «Строки, не попавшие в расчёт». Прочие warnings оставить.

**Acceptance:**
- [x] Тесты панелей/шага 3: нет этих строк при `unparsed_lines.length > 0`
- [x] Warning про Д×Ш×H / нагрузку по умолчанию всё ещё виден, если есть

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/components/KpPlatePreviewPanel src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx`

**Dependencies:** Task 6 (гейт уже не пускает красное текстом)  
**Files:** новый `lib/compositionWarnings.ts` (+ тест), `KpPlatePreviewPanel.tsx`, `CalculationResultStep.tsx`  
**Estimated scope:** M

#### Task 9: Остальные preview-панели

**Description:** Тот же хелпер и скрытие «Не попали в состав» в pile/fbs/march/step/bridge-pile панелях.

**Acceptance:**
- [x] Ни одна `Kp*PreviewPanel` не показывает блок «Не попали в состав»
- [x] Баннер распознанных строк-счётчика нет

**Verification:** `cd frontend && npm run test -- src/features/commercial-offer/components/Kp`

**Dependencies:** Task 8  
**Files:** пять `Kp*PreviewPanel.tsx`  
**Estimated scope:** M

### Checkpoint: Complete

- [ ] Safety net pytest commercial (как в спеке Commands)
- [ ] `cd frontend && npm run test && npm run typecheck`
- [ ] Вручную: текст — подсветка + гейт; фото — «Распознать фото» живая; «Готово, далее» не серится из‑за линта; шаг 3 без мёртвой плашки; wide/unpriced на плитах живы
- [ ] Success criteria спеки отмечены

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Линт-строка ≠ то, что примет полный `parse_plate_text` после нормализации | Med | Спека это приняла; ложные красные чинят в поле. Не подключать generate_preview. |
| Шесть InputStep разъедутся при выносе карточки | High | Сначала плиты (T6), затем пять копипаст одним коммитом-волной (T7) |
| `/parse` для не-плит ломает клиента, ждавшего `order` | Low | Старый вызов без product_type — только плиты. Фронт линта — единственный новый потребитель. |
| Дебаунс + серая кнопка «навсегда» при ошибке сети | Med | Хук: ошибка запроса → не ok для гейта (кнопка серая) + можно повторить на следующий debounce; не глотать 401/500 как «все строки зелёные» |
| Снятие плашки до гейта | High | T8 строго после T6 |

## Parallelization

- T5 параллельно T3–T4
- T8–T9 не начинать до T6
- T7 только после стабильной карточки

## Open Questions

Нет. Допущения спеки 1–10 в силе. Коммиты — только по просьбе.
