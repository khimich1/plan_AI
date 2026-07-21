# Implementation Plan: UX менеджера — редизайн мастера КП

> Spec: [manager-ux-redesign.md](../specs/manager-ux-redesign.md)  
> Idea: [manager-ux-improvements.md](../ideas/manager-ux-improvements.md)

## Overview

Поэтапный редизайн веб-мастера коммерческих предложений: **quick wins** (можно отдельными PR) → **3-шаговый мастер v1** (backend + frontend) → **архив v2** (копирование КП). Каждая фаза оставляет систему в рабочем состоянии.

## Architecture Decisions

| Решение | Обоснование |
|---------|-------------|
| 3 канонических шага: `plates`, `client`, `result` | Баланс скорости и безопасности OCR; ~40% меньше кликов vs 5 шагов |
| Wide plates и manager — UI внутри шагов, не отдельные step id | Меньше cognitive load; сервер проверяет prerequisites, не номер экрана |
| Legacy aliases для `wide-plates`, `manager` | Черновики в localStorage и на сервере не ломаются |
| Default manager = `AuthUser.manager_id` | Поле уже в БД и API `/auth/me`; отдельный шаг не нужен |
| Quick wins отдельно от редизайна | Можно выкатить за 1 день до основного рефакторинга |
| Archive copy — v2 | Не блокирует v1 мастера; высокая ценность для повторных заказов |

## Dependency Graph

```
Phase 0: Quick wins (независимо)
    │
Phase 1: Backend 3-step model
    │
    ├── Phase 2a: Step 1 UI (plates + OCR review + inline wide)
    ├── Phase 2b: Step 2 UI (client + manager default)
    ├── Phase 2c: Step 3 UI (checklist)
    │
Phase 2.5: Batch OCR review (v1.1) — после 2.1
    │
Phase 3: Cleanup + migration + polish
    │
Phase 4 (v2): Archive copy КП
```

---

## Phase 0: Quick Wins

Можно выполнять параллельно или до Phase 1. Каждая задача — отдельный маленький PR.

### Task 0.1: Скрыть «Производство» для manager

**Description:** В `AppHeader.tsx` не рендерить NavLink «Производство», если `user.role === "manager"`. Admin и production — без изменений.

**Acceptance criteria:**
- [ ] Manager не видит пункт «Производство».
- [ ] Admin видит «Производство».
- [ ] Production user видит только production nav (как сейчас).
- [ ] Прямой URL `/production` для manager по-прежнему блокируется `RequireRole`.

**Verification:**
- [ ] Ручная проверка под тремя ролями.
- [ ] `npm run build`.

**Dependencies:** None

**Files:**
- `frontend/src/app/layout/AppHeader.tsx`

**Scope:** XS

---

### Task 0.2: Primary-кнопки на шаге ввода

**Description:** Упростить кнопки: primary «Распознать фото» / «Обработать текст». **«Добавить к списку»** — secondary на основном уровне (частый сценарий). **«Заменить» и «ИИ»** — в collapsible «Дополнительно» (редкие).

**Acceptance criteria:**
- [ ] До распознавания — одна primary-кнопка.
- [ ] «Добавить к списку» видна на основном уровне при наличии черновика.
- [ ] «Заменить» и «ИИ» — только в «Дополнительно».

**Verification:**
- [ ] Ручная проверка: фото-only, text-only.
- [ ] `npm run build`.

**Dependencies:** None (если только переименование; полный collapsible — Task 2.1)

**Files:**
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`

**Scope:** S

---

### Task 0.4: Save mode — дефолт «В работе»

**Description:** В `SaveOfferSection`: primary action — сохранение «В работе»; «В архив» и «Пропустить» — в collapsible «Другой вариант сохранения».

**Acceptance criteria:**
- [ ] По умолчанию выделен/предложен режим «В работе».
- [ ] Альтернативы доступны, но не конкурируют с primary.

**Verification:**
- [ ] Ручная проверка save flow.
- [ ] `npm run build`.

**Dependencies:** None

**Files:**
- `frontend/src/features/commercial-offer/components/SaveOfferSection.tsx`

**Scope:** S

---

### Task 0.3: Чеклист на шаге результата (минимальный)

**Description:** Добавить блок «Готовность КП» в `CalculationResultStep` до существующих секций: кол-во плит, клиент, сумма, список предупреждений.

**Acceptance criteria:**
- [ ] Отображаются ✓ N плит, ✓ клиент, ✓ сумма.
- [ ] При `warnings` / `unparsed_lines` — ⚠ с кратким текстом.

**Verification:**
- [ ] Ручная проверка на черновике с предупреждениями.
- [ ] `npm run build`.

**Dependencies:** None

**Files:**
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.tsx`

**Scope:** S

---

### Checkpoint: Phase 0
- [ ] Quick wins можно выкатить без изменения backend step model.
- [ ] Регрессий в auth/RBAC нет.

---

## Phase 1: Backend — 3-шаговая модель

### Task 1.1: Обновить WizardStepId и legacy aliases

**Description:** В `app/schemas/commercial.py` оставить канонические `plates`, `client`, `result`. В `_coerce_wizard_step_id` добавить aliases: `wide-plates` → `plates`, `manager` → `client`. Обновить `WizardNextRequiredAction` при необходимости (убрать/сохранить `select_manager` как internal prerequisite).

**Acceptance criteria:**
- [ ] Enum содержит 3 шага.
- [ ] Legacy значения корректно коэрсятся.
- [ ] OpenAPI / существующие clients не падают на старых черновиках.

**Verification:**
- [ ] `pytest tests/ -k wizard -q`
- [ ] Новые тесты на coercion.

**Dependencies:** None

**Files:**
- `app/schemas/commercial.py`
- `tests/test_commercial_wizard_step_service.py` (или новый файл)

**Scope:** M

---

### Task 1.2: Переписать CommercialWizardStepService под 3 шага

**Description:** Обновить `infer_wizard_current_step`, `infer_can_proceed_to`, `infer_next_required_action`:

- `plates`: can_proceed → `client` если есть `order_data` и wide plates resolved.
- `client`: can_proceed → `result` после successful calculate prerequisites.
- Wide plates blocking возвращает effective step `plates`, не `wide-plates`.
- Missing manager blocking на шаге `client`, не отдельном `manager`.

**Acceptance criteria:**
- [ ] Все существующие wizard step tests обновлены и green.
- [ ] `can_proceed_to` из plates не включает legacy steps.
- [ ] Черновик с неразрешёнными wide plates не переходит на client.

**Verification:**
- [ ] `pytest tests/test_commercial_wizard_step_service.py -q`
- [ ] `pytest tests/ -q` (полный прогон)

**Dependencies:** Task 1.1

**Files:**
- `app/services/commercial_wizard_step_service.py`
- `tests/test_commercial_wizard_step_service.py`

**Scope:** M

---

### Task 1.3: Frontend types и wizardStepOrder

**Description:** Синхронизировать `WizardStepId`, `WIZARD_STEP_ORDER`, `WizardProgress` labels. Добавить `mapLegacyWizardStep()` для hydrate из localStorage.

**Acceptance criteria:**
- [ ] TypeScript компилируется без `wide-plates` / `manager` в union (кроме legacy mapper).
- [ ] `wizardDraftStore` merge учитывает 3 шага.

**Verification:**
- [ ] `npm run build`
- [ ] `npm test -- wizardDraftStore`

**Dependencies:** Task 1.1 (контракт согласован)

**Files:**
- `frontend/src/features/commercial-offer/types/commercialOffer.ts`
- `frontend/src/features/commercial-offer/lib/wizardStepOrder.ts`
- `frontend/src/features/commercial-offer/store/wizardDraftStore.tsx`
- `frontend/src/features/commercial-offer/components/WizardProgress.tsx`

**Scope:** M

---

### Checkpoint: Phase 1
- [ ] Backend и frontend согласованы на 3 шага.
- [ ] `pytest` + `npm run build` green.
- [ ] Старые черновики открываются (legacy mapping).

---

## Phase 2: Frontend — UI редизайн шагов

### Task 2.1: Шаг 1 — режим сверки OCR (side-by-side)

**Description:** Рефакторинг `PlateInputStep`:
1. Два подрежима: **ввод** / **сверка** (после успешного recognize).
2. Сверка: фото слева (zoom, open in new tab), список справа — редактируемый `normalizedText` с подсветкой строк из `ocr_corrections`, `unparsed_lines`, `wide_plate_lines`.
3. Footer: «Список верен — далее» → вызывает resolve wide (если есть) + `handleProcess` → client.
4. Блок «Дополнительно»: append, replace, ИИ.

**Acceptance criteria:**
- [ ] Side-by-side при наличии `recognizedImageUrl`.
- [ ] Технические подписи заменены на бизнесовые (см. spec).
- [ ] Primary CTA — «Список верен — далее» в режиме сверки.

**Verification:**
- [ ] Ручная: фото → recognize → сверка → далее.
- [ ] `npm run build`

**Dependencies:** Task 1.2, 1.3

**Files:**
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx` (navigation)
- Возможно новый: `OcrReviewPanel.tsx`, `PlateListEditor.tsx`

**Scope:** L → разбить на 2.1a (layout) + 2.1b (highlighting) при необходимости

---

### Task 2.2: Inline wide plates в шаге 1

**Description:** Встроить логику `WidePlateReviewStep` в `PlateInputStep`: **collapsible, свёрнуто по умолчанию**, бейдж «N позиций требуют внимания». Удалить отдельный step `wide-plates`.

**Acceptance criteria:**
- [ ] При `wide_plate_lines.length > 0` — секция **свёрнута** с бейджем N; разворачивается по клику.
- [ ] «Далее» disabled пока wide plates не resolved (server validation).
- [ ] `WidePlateReviewStep.tsx` deprecated или удалён после переноса.

**Verification:**
- [ ] Ручная: черновик с широкими плитами.
- [ ] `npm run build`

**Dependencies:** Task 2.1

**Files:**
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`
- `frontend/src/features/commercial-offer/components/steps/WidePlateReviewStep.tsx` (удалить/реэкспорт)
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`

**Scope:** M

---

### Task 2.3: Шаг 2 — клиент + менеджер по умолчанию

**Description:** Расширить `ClientConditionsStep`:
- Автоподстановка `managerId` из `useAuth().user.manager_id` при mount (если менеджер есть в списке).
- Collapsed «Другой менеджер» с select.
- Убрать `ManagerStep` из wizard flow; объединить submit (manager + client + calculate).

**Acceptance criteria:**
- [ ] При `manager_id` в профиле — менеджер выбран, блок смены свёрнут.
- [ ] Без `manager_id` в профиле — блок выбора менеджера **раскрыт по умолчанию**, submit без выбора невозможен.
- [ ] `ManagerStep.tsx` не используется в wizard.

**Verification:**
- [ ] Unit test: default manager from auth mock.
- [ ] Ручная: submit → result.

**Dependencies:** Task 1.2, 1.3

**Files:**
- `frontend/src/features/commercial-offer/components/steps/ClientConditionsStep.tsx`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- `frontend/src/features/commercial-offer/components/steps/ManagerStep.tsx` (удалить)

**Scope:** M

---

### Task 2.4: Шаг 3 — polish чеклиста и кнопок файлов

**Description:** Доработать Task 0.3: контекстные кнопки PDF/Excel/схема (показывать по `draft.files` / metadata). Улучшить подсказки save mode.

**Acceptance criteria:**
- [ ] Не все кнопки генерации видны одновременно без необходимости.
- [ ] Save mode с кратким пояснением каждого варианта.

**Verification:**
- [ ] Ручная на полном flow.
- [ ] `npm run build`

**Dependencies:** Task 0.3, Phase 2 предыдущие

**Files:**
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.tsx`
- `frontend/src/features/commercial-offer/components/DownloadFilesSection.tsx`
- `frontend/src/features/commercial-offer/components/SaveOfferSection.tsx`

**Scope:** S

---

### Checkpoint: Phase 2
- [ ] E2E flow: фото → сверка → клиент → результат → PDF.
- [ ] 3 шага в сайдбаре.
- [ ] Нет мёртвого кода ManagerStep / WidePlateReviewStep (или помечен deprecated).

---

## Phase 2.5: Batch OCR review (v1.1)

> Spec: [Многофото: batch vs cumulative](../specs/manager-ux-redesign.md#многофото-batch-vs-cumulative-v11)  
> Зависит от Phase 2.1; не блокирует Phase 3.

### Task 2.5.1: Состояние pending batch review на фронте

**Description:** В `wizardDraftStore` — `pendingBatchReview`, `confirmedBatchCount`, `batchReviewText`. После recognize/append — `start-batch-review`; после «Список верен» — `confirm-batch-review`. Без backend `confirmed` flag (достаточно `plate_batches[]`).

**Acceptance criteria:**
- [ ] После OCR открывается сверка только последнего batch.
- [ ] «Список верен» закрывает сверку, остаёмся на шаге 1.
- [ ] Reload: незавершённый batch снова в сверке (`confirmedBatchCount` в localStorage).

**Verification:**
- [ ] `npm test -- batchReview`
- [ ] `npm test -- wizardDraftStore`

**Files:**
- `frontend/src/features/commercial-offer/store/wizardDraftStore.tsx`
- `frontend/src/features/commercial-offer/types/commercialOffer.ts`
- `frontend/src/features/commercial-offer/lib/batchReview.ts`

**Scope:** M

---

### Task 2.5.2: UI — две CTA и batch-only side-by-side

**Description:** `PlateInputStep`: сверка показывает `plate_batches[-1].normalized_text`; две кнопки — **«Список верен»** (primary в режиме сверки) и **«Готово, далее»** (disabled при pending review / wide plates). `KpPlatePreviewPanel` — суммарный `order_data` без изменений.

**Acceptance criteria:**
- [ ] Append: в textarea сверки только новые строки, не cumulative.
- [ ] Подсветка OCR — только строки текущего batch (`filterDraftForBatchReview`).
- [ ] «Готово, далее» disabled, пока не подтверждён текущий batch.
- [ ] Wide plates блокируют обе кнопки по всему черновику.

**Verification:**
- [ ] Ручная: 2 фото → сверка batch 2 only → «Список верен» → «Готово, далее».
- [ ] `npm run build`

**Files:**
- `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- `frontend/src/features/commercial-offer/components/PlateListEditor.tsx`
- `frontend/src/features/commercial-offer/lib/plateLineHighlights.ts` (через batch filter)
- `frontend/src/features/commercial-offer/components/KpPlatePreviewPanel.tsx`

**Scope:** M

---

### Task 2.5.3: Редактирование batch и тесты

**Description:** При правке в сверке — merge в full text + `updatePlates(replace)` на confirm. Unit-тесты `batchReview.test.ts`; обновить тексты alert в preview.

**Acceptance criteria:**
- [ ] Правка текста batch сохраняется по «Список верен».
- [ ] Single-photo flow без регрессии.
- [ ] `npm test` + `npm run build` green.

**Verification:**
- [ ] `npm test -- batchReview`
- [ ] `pytest tests/test_commercial_wizard_step_service.py -q` (если backend не трогали — smoke)

**Files:**
- `frontend/src/features/commercial-offer/lib/batchReview.ts`
- `frontend/src/features/commercial-offer/lib/batchReview.test.ts`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`

**Scope:** S

---

### Checkpoint: Phase 2.5 (v1.1)
- [ ] Многофото: batch-сверка + cumulative preview работают по spec.
- [ ] Backend без новых полей в `CommercialPlateBatch`.

---

## Phase 2.6: Синхронизация списка ↔ wide plates ↔ состав КП (v1.2)

> Spec: [Синхронизация v1.2](../specs/manager-ux-redesign.md#синхронизация-список-для-расчёта--wide-plates--состав-кп-v12)  
> Зависит от Phase 2.2 (inline wide plates) и 2.5 (batch review).

### Task 2.6.1: Backend — обновление `plate_batches` после `resolve_wide_plates`

**Description:** При `resolve_wide_plates` применять exclude/replace/confirm к `normalized_text` каждого batch, чтобы фронт мог показать актуальный «Список для расчёта».

**Acceptance criteria:**
- [ ] Исключение wide plate удаляет строку из `plate_batches[-1].normalized_text`.
- [ ] Замена подставляет новые строки в соответствующий batch.
- [ ] `order_data` пересчитывается (уже было).

**Verification:**
- [ ] `pytest tests/test_commercial_web_flow.py -k resolve_wide_plates -q`

**Files:**
- `app/services/commercial_workflow_service.py`
- `tests/test_commercial_web_flow.py`

**Scope:** S

---

### Task 2.6.2: Frontend — sync store после «Применить решения»

**Description:** После `resolve_wide_plates` — `sync-after-wide-plates`: обновить `batchReviewText`, `normalizedText`, сбросить `widePlateActions`. Улучшить `hydrate-draft` для refresh batch text при pending review.

**Acceptance criteria:**
- [ ] Textarea «Список для расчёта» сразу отражает exclude/replace.
- [ ] «Состав КП» показывает пересчитанный `order_data`.
- [ ] Ручные правки textarea учитываются на «Список верен» (`update_plates`).

**Verification:**
- [ ] `npm test -- wizardDraftStore`
- [ ] `npm test -- batchReview`
- [ ] `npm run build`

**Files:**
- `frontend/src/features/commercial-offer/store/wizardDraftStore.tsx`
- `frontend/src/features/commercial-offer/hooks/useCommercialOfferWizard.ts`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`

**Scope:** M

---

### Checkpoint: Phase 2.6 (v1.2)
- [ ] После «Применить решения» список и состав КП синхронны.
- [ ] Ручное удаление широкой строки в textarea ≡ «Исключить» (через «Список верен»).

---

## Phase 3: Cleanup, тексты, валидация

### Task 3.1: Человекочитаемые ошибки и предупреждения

**Description:** Пройти `validation_errors`, alerts в commercial-offer feature; заменить жаргон. Согласовать с backend messages в `commercial_calculation_service.py` где нужно.

**Acceptance criteria:**
- [ ] Нет «backend», «нормализованный» в user-facing текстах шага 1–3.
- [ ] Wide plate messages — рекомендация замены простым языком.

**Verification:**
- [ ] Grep по feature на запрещённые термины.
- [ ] `pytest` + `npm run build`

**Dependencies:** Phase 2 complete

**Files:**
- `PlateInputStep.tsx`, `ClientConditionsStep.tsx`, `CalculationResultStep.tsx`
- `app/services/commercial_calculation_service.py` (если server messages)

**Scope:** S

---

### Task 3.2: Обновить тесты и удалить legacy step handling в UI

**Description:** Финальная чистка: тесты wizard, удаление unused imports, обновление `wizardDraftStore.test.tsx`.

**Acceptance criteria:**
- [ ] `pytest tests/ -q` green.
- [ ] `npm test` green.
- [ ] `npm run build` green.

**Verification:** CI-эквивалент локально.

**Dependencies:** Task 3.1

**Scope:** S

---

### Task 3.3: Ручная приёмка с менеджером

**Description:** 5 реальных КП think-aloud; зафиксировать время, ошибки, feedback.

**Acceptance criteria:**
- [ ] Протокол в `ai_docs/develop/reports/manager-ux-v1-acceptance.md` (или устно — по договорённости).
- [ ] Критичные блокеры заведены как issues.

**Dependencies:** Task 3.2

**Scope:** — (human)

---

### Checkpoint: v1 Complete
- [ ] Все Success Criteria из spec v1 выполнены.
- [ ] Готово к релизу v1 мастера.

---

## Phase 4 (v2): Архив — «Создать КП на основе»

> После стабилизации v1. Spec-детали — кратко здесь; полный spec v2 — отдельный документ при старте фазы.

### Task 4.1: API — копирование позиций и условий из архива

**Description:** Endpoint duplicate: создаёт draft **только с позициями плит** из архива. Клиент, условия — пустые/стандартные; менеджер = текущий user.

**Acceptance criteria:**
- [ ] Новый draft с `wizard_state.current_step = plates` или `client` (если плиты валидны).
- [ ] Авторизация: только commercial roles.

**Verification:**
- [ ] pytest для endpoint.

**Dependencies:** v1 complete

**Scope:** M

---

### Task 4.2: UI — кнопка в архиве и переход в мастер

**Description:** В `OfferDetailsDrawer` / списке архива: «Создать КП на основе №…» → navigate `/new` с prefilled draft id.

**Acceptance criteria:**
- [ ] Копируются **только позиции плит**; шаг 2 заполняется заново.
- [ ] Менеджер = **текущий залогиненный user** (`user.manager_id`); если null — обязательный выбор на шаге «Клиент».

**Verification:**
- [ ] Ручная: archive → new wizard.

**Dependencies:** Task 4.1

**Scope:** M

---

### Task 4.3: Discoverability редактирования скидки в архиве

**Description:** Заметная кнопка **«Изменить»** (скидка/логистика) прямо в карточке/drawer архива — не спрятана в глубине.

**Dependencies:** None (можно параллельно 4.1)

**Scope:** S

---

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Рассинхрон frontend/backend step model | High | Phase 1 checkpoint; единый enum; тесты на обоих сторонах |
| Сломанные localStorage черновики | Med | `mapLegacyWizardStep`; не менять ключ storage без migration |
| `manager_id` null у всех users | Med | Graceful fallback: явный выбор; документировать настройку в admin |
| Scope creep (2 экрана, дашборд) | Med | Spec boundaries; v2/v3 явно out |
| Регрессия OCR flow | High | Ручная приёмка с фото; не менять OCR pipeline в v1 |

## Parallelization

| Параллельно | Последовательно |
|-------------|-----------------|
| Task 0.1, 0.2, 0.3 | Phase 1 → Phase 2 |
| Task 3.1 (тексты) частично с 2.4 | 2.1 → 2.2 (wide inline) |
| Task 4.3 (archive UX) с 4.1 | 4.1 → 4.2 |

## Estimated Effort (ориентир)

| Phase | Оценка |
|-------|--------|
| Phase 0 Quick wins | 0.5–1 день |
| Phase 1 Backend | 1–1.5 дня |
| Phase 2 Frontend | 2–3 дня |
| Phase 3 Polish | 0.5–1 день |
| Phase 4 v2 Archive | 1.5–2 дня |

**Итого v1:** ~4–6 дней focused work.  
**v1 + v2:** ~6–8 дней.

## Решения заказчика (2026-07-13)

| Вопрос | Решение |
|--------|---------|
| `manager_id = null` | Обязательный выбор менеджера (раскрытый блок) |
| Save mode | Дефолт «В работе»; остальное в «Другой вариант сохранения» |
| Порядок выката | **Phase 0 first** → Phase 1+ |
| OCR автопереход | Только «Список верен — далее» |
| v2 duplicate: что копировать | **Только плиты**; клиент и условия заново |
| Широкие плиты | Свёрнуто + бейдж N |
| Дополнительно шаг 1 | «Добавить» на уровне; «Заменить»/«ИИ» в collapsed |
| Список в сверке | Textarea |
| Archive скидка | Кнопка «Изменить» в drawer |

## Следующий шаг

1. ~~Подтвердить spec + план~~ — **подтверждено**.
2. Начать **Phase 0** (Tasks 0.1–0.3) отдельным PR.
3. После checkpoint Phase 0 — Phase 1 (backend 3-step model).
