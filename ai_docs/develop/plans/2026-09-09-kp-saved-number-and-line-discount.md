# Plan: КП — номер после сохранения и цены со скидкой на шаге 3

**Дата**: 2026-09-09  
**Spec**: [../../specs/kp-saved-number-and-line-discount.md](../../specs/kp-saved-number-and-line-discount.md)  
**Idea**: [../../ideas/kp-saved-number-and-line-discount.md](../../ideas/kp-saved-number-and-line-discount.md)  
**Статус**: PLAN ✅ · IMPLEMENT ✅

## Overview

Два связанных бага на шаге результата конструктора. (1) Identity всегда `WEB_{draft}` даже после записи в архив, потому что `build_offer_identity` не смотрит на `kp_id`, а `save_offer` генерирует XLSX *до* persist. (2) Таблица шага 3 рисует прайс, PDF — цену со скидкой. План: сначала канон номера (export service → порядок save → реген файлов), затем display-only скидка на фронте. Без предсказания номера, без новых таблиц, без кода в этой сессии.

## Architecture Decisions

- **A1.** Один метод identity: `CommercialExportService.build_offer_identity(draft_id, metadata=None)`. `kp_id` из `saved_offer.kp_id` или `resume_kp_id`. Иначе `WEB_{draft[:8].upper()}`. Не восстанавливать `workflow._build_offer_identity` и не звать `get_next_kp_number`.
- **A2.** Первый save: persist с `xlsx_path=None` → записать `saved_offer` в metadata → generate существующих видов с новым номером → `save_xlsx_file` / `update_offer_from_order_data` с путём XLSX. Resume: generate сразу с уже известным `kp_id`.
- **A3.** После save регенерировать только уже лежащие в `generated_files` виды плюс всегда `xlsx`. `breakdown`/`schema` сейчас skip-if-exists — на этом регене **форсировать** overwrite, иначе в UI останутся WEB-имена.
- **A4.** Скидка — чистый TS-хелпер + замена ячеек в `CalculationResultStep`. `order_data.unit_price` не трогать. API `discounted_price` на строках черновика не добавлять, пока не разъедется округление с PDF.
- **A5.** `SaveOfferSection` не хачить: плашка уже читает `result_card.offer_number`. Починить бэкенд.
- **A6.** Старые файлы в `outputs/` не удалять.

## Dependency graph

```
identity (export service)
    │
    ├── generate_files / get_draft_details
    │
    └── save_offer: persist → metadata → generate → xlsx_path
            │
            └── UI плашка / DownloadFilesSection (без своей логики)

discount display helper  (параллельно identity после Task 2, не раньше Checkpoint 1 если один агент)
    └── CalculationResultStep все ветки таблицы
```

Два среза можно параллелить **после** Task 2: save-порядок и скидка не зависят друг от друга. Identity должен быть первым.

## Task List

### Phase 1: Identity foundation

- [x] **Task 1: RED — identity WEB vs saved vs resume**
  - **Description:** Заменить сирот `test_build_offer_identity_uses_predicted_kp_number` и починить `test_build_offer_identity_prefers_saved_kp_id`: вызывать `export_service.build_offer_identity`, не `workflow._build_offer_identity`. Predicted-сценарий удалить. Добавить кейсы: нет kp_id → `WEB_` + stem `kp_{draft[:8]}_`; `saved_offer.kp_id=777` → `"777"` + `kp_777_`; только `resume_kp_id` → тот же контракт.
  - **Acceptance criteria:**
    - [ ] Нет ассерта на `predicted_kp_id` / `get_next_kp_number`
    - [ ] Три кейса (unsaved / saved / resume) красные, пока метод не принимает metadata
  - **Verification:**
    - [ ] `pytest tests/test_commercial_web_flow.py::test_build_offer_identity_uses_web_prefix_when_unsaved tests/test_commercial_web_flow.py::test_build_offer_identity_prefers_saved_kp_id tests/test_commercial_web_flow.py::test_build_offer_identity_uses_resume_kp_id -q`
  - **Dependencies:** None
  - **Files likely touched:**
    - `tests/test_commercial_web_flow.py`
  - **Estimated scope:** S

- [x] **Task 2: GREEN — `build_offer_identity` читает metadata**
  - **Description:** Сигнатура `(draft_id, metadata=None)`. Резолв `kp_id` как в спеке. Прокинуть metadata из `generate_files` (уже есть `payload["metadata"]`), `get_draft_details`, `build_offer_identity_payload`. Мок в `test_offer_counterparty_validation.py` (`lambda draft_id:`) расширить до `*args, **kwargs`, иначе CTR-006 упадёт.
  - **Acceptance criteria:**
    - [ ] Тесты Task 1 зелёные
    - [ ] `generate_files` без saved kp по-прежнему даёт `WEB_`
    - [ ] `get_draft_details` при `saved_offer.kp_id` возвращает `offer_number == str(kp_id)`
  - **Verification:**
    - [ ] `pytest tests/test_commercial_web_flow.py tests/test_offer_counterparty_validation.py -q`
  - **Dependencies:** Task 1
  - **Files likely touched:**
    - `app/services/commercial_export_service.py`
    - `app/services/commercial_draft_lifecycle.py` (только вызовы identity / details)
    - `tests/test_offer_counterparty_validation.py`
  - **Estimated scope:** M

### Checkpoint: Foundation

- [x] Identity-тесты зелёные; CTR-006 не сломан
- [x] Несохранённый черновик визуально всё ещё `WEB_*` (регрессия не сломана)
- [x] Не приступать к save-порядку, пока Checkpoint 1 не зелёный

### Phase 2: Save → generate with kp_id

- [x] **Task 3: RED — save persist-then-generate**
  - **Description:** Тест первого save: `result_card.offer_number == str(kp_id)`; в возвращённых/записанных `generated_files` pdf/xlsx нет `WEB_` и нет `kp_{draft[:8]}`; stem начинается с `kp_{kp_id}_`. Зафиксировать, что `generate_files` после save не зовут с пустым metadata. Resume (MNA-304) — тот же `kp_id`, не create. Если в suite уже есть интеграционный save — расширить ассертами identity, не плодить второй полный флоу без нужды.
  - **Acceptance criteria:**
    - [ ] Первый save: номер и имена файлов = `kp_id`
    - [ ] Повторный generate после save не возвращает `WEB_`
    - [ ] Resume update не создаёт новый id
  - **Verification:**
    - [ ] `pytest tests/test_commercial_web_flow.py tests/test_commercial_draft_append.py -k "save_offer or offer_identity or generate" -q`
  - **Dependencies:** Task 2
  - **Files likely touched:**
    - `tests/test_commercial_web_flow.py`
    - при необходимости `tests/test_commercial_draft_append.py`
  - **Estimated scope:** M

- [x] **Task 4: GREEN — порядок save + force-regen существующих видов**
  - **Description:** В `save_offer` убрать generate-xlsx до persist. Create: `save_offer(..., xlsx_path=None)` → `update_metadata(saved_offer)` → `generate_files` с видами `xlsx` + те, что уже были в `generated_files` (`pdf`, `breakdown`, `schema`) → проставить `xlsx_path` через существующий `save_xlsx_file` / `update_offer_from_order_data`. Resume: identity уже известен, generate можно до или после update, но не `WEB_*`. В `generate_files` для этого вызова отключить skip-if-exists у `breakdown`/`schema` (флаг `replace_existing=True` или локальная очистка kinds перед генерацией). `result_card` строить после metadata с `kp_id`.
  - **Acceptance criteria:**
    - [ ] Тесты Task 3 зелёные
    - [ ] MNA-304 / CTR-006 зелёные
    - [ ] После save `get_draft_details.files` не содержит WEB-имён для регененных видов
  - **Verification:**
    - [ ] `pytest tests/test_commercial_web_flow.py tests/test_commercial_draft_append.py tests/test_offer_counterparty_validation.py tests/test_commercial_export_mixed.py -q`
  - **Dependencies:** Task 3
  - **Files likely touched:**
    - `app/services/commercial_draft_lifecycle.py`
    - `app/services/commercial_export_service.py` (force replace)
  - **Estimated scope:** M

### Checkpoint: Номер после save

- [x] Create → «В архив» → плашка `в архиве: {n}`; скачать PDF/XLSX с `kp_{n}_` и шапкой `{n}`
- [x] Resume из архива: identity сразу `{n}`, без `WEB_*`
- [x] Скидку ещё не трогали — таблица по-прежнему с прайсом (ожидаемо)

### Phase 3: Display discount

- [x] **Task 5: RED/GREEN — `discountedUnitPrice`**
  - **Description:** Новый pure helper (предпочтительно `frontend/src/features/commercial-offer/lib/lineDiscountDisplay.ts`): формула как в `core/commercial_offer.py` (`price * (1 - clamp(percent,0,100)/100)`), без предварительного round; форматирование остаётся в `formatOfferNumber` / `formatOfferSum`.
  - **Acceptance criteria:**
    - [ ] 42508 × 10% → 38257.2
    - [ ] 0% = прайс; 100% = 0; невалидная цена → `null`
  - **Verification:**
    - [ ] `cd frontend && npm run test -- --run src/features/commercial-offer/lib/lineDiscountDisplay.ts`
  - **Dependencies:** None (можно параллельно с Phase 2 после Checkpoint 1)
  - **Files likely touched:**
    - `frontend/src/features/commercial-offer/lib/lineDiscountDisplay.ts`
    - `frontend/src/features/commercial-offer/lib/lineDiscountDisplay.test.ts`
  - **Estimated scope:** S

- [x] **Task 6: Таблица шага 3 показывает цену со скидкой**
  - **Description:** Во всех трёх ветках `CalculationResultStep` (ступени / grade-simple / плиты) Цена и Сумма из `discountedUnitPrice(item.unit_price, draft.metadata.discount_percent)`. Не менять `order_data`. Тест: `discount_percent: 10`, `unit_price: 10000`, qty 2 → в ячейках не «10 000» / «20 000», а 9 000 / 18 000. Регрессия: целевая сумма по-прежнему от прайса (существующие тесты `discountFromTargetSum` не трогать).
  - **Acceptance criteria:**
    - [ ] Плиты, grade-simple и ступени на шаге 3 со скидкой в Цена/Сумма
    - [ ] `unit_price` в пропе draft не мутируется
  - **Verification:**
    - [ ] `cd frontend && npm run test -- --run src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx src/features/commercial-offer/lib/discountFromTargetSum.ts`
    - [ ] `cd frontend && npm run typecheck`
  - **Dependencies:** Task 5
  - **Files likely touched:**
    - `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.tsx`
    - `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx`
  - **Estimated scope:** S

- [x] **Task 7: Плашка save показывает номер из result_card**
  - **Description:** Узкий тест `SaveOfferSection`: `lastSaveResult.result_card.offer_number = "1188"` → текст плашки содержит `1188`, не `WEB_`. Кода в компоненте, скорее всего, не нужно (A5).
  - **Acceptance criteria:**
    - [ ] Тест зелёный без хака маппинга WEB→номер на фронте
  - **Verification:**
    - [ ] `cd frontend && npm run test -- --run src/features/commercial-offer/components/SaveOfferSection.test.tsx`
  - **Dependencies:** Task 4 (контракт), Task 6 не обязателен
  - **Files likely touched:**
    - `frontend/src/features/commercial-offer/components/SaveOfferSection.test.tsx`
  - **Estimated scope:** XS

### Checkpoint: Complete

- [x] pytest: web_flow + draft_append + counterparty + export_mixed
- [x] frontend: commercial-offer tests + typecheck
- [ ] Ручной smoke (ниже)
- [x] Spec статус → IMPLEMENT ✅ после кода

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Persist без XLSX ломает `save_kp_to_db` | High | `xlsx_file_path` уже Optional; сразу после generate вызвать `save_xlsx_file(kp_id, path)`. Если create без файла падает — в Task 4 поймает RED save-тест |
| `breakdown`/`schema` skip-if-exists оставит WEB-имена | High | Явный force-replace только на post-save generate, не менять обычный generate до save |
| Моки `generate_files` / `build_offer_identity_payload(draft_id)` разъедутся по сигнатуре | Med | Task 2 сразу чинит CTR-006 mock; grep `build_offer_identity` |
| Двойной generate на resume (xlsx до update + после) | Low | Resume: один generate с известным kp_id; не генерировать «для persist» отдельным WEB-проходом |
| Округление UI vs PDF на копейку | Low | Без extra round, как спека Q3; follow-up только если smoke покажет расхождение |
| Скидка в превью шага 2 (KpPlatePreviewPanel) | Low | Out of scope; только шаг 3 |
| `save_offer` integration tests ждут generate до persist | Med | Task 3/4 правит ассерты порядка, не поведение архива |

## Open Questions

Нет блокирующих. Решения спеки D1–D9 в силе.

Если при Task 4 окажется, что `save_kp_to_db` реально требует существующий xlsx — не предсказывать номер; допустим временный файл, затем перегенерация с `kp_id` и `save_xlsx_file`. Не откатываться к generate-до-persist с WEB в шапке.

## Verification Commands

```bash
source venv/bin/activate
pytest tests/test_commercial_web_flow.py \
  tests/test_commercial_draft_append.py \
  tests/test_offer_counterparty_validation.py \
  tests/test_commercial_export_mixed.py -q

cd frontend && npm run test -- --run src/features/commercial-offer
cd frontend && npm run typecheck
```

Ручной smoke на `./run+logs.sh`:

1. Новое КП, шаг 3, скидка 10% → строка меняет Цену/Сумму; итог как в PDF.
2. Сформировать PDF/XLSX до save → имена `kp_{hex}_…`, номер в шапке `WEB_…`.
3. «В архив» → плашка `в архиве: {n}`; список файлов `kp_{n}_…`; скачать PDF — шапка `{n}`.
4. Повторно «Сформировать» — снова `{n}`, не WEB.
5. Архив → «Редактировать» → шаг 3 сразу с номером `{n}`.
6. Целевая сумма после скидки в UI — база от прайса, apply не ломается.

## Parallelization

| Можно параллельно | Только последовательно |
|-------------------|------------------------|
| Task 5–6 (скидка) после Checkpoint 1 | 1 → 2 → 3 → 4 |
| Task 7 после Task 4 | Checkpoint 2 до объявления «номер готов» |

Один агент: порядок 1–4, затем 5–7.
