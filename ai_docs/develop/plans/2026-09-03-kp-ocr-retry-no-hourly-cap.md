# Plan: КП — OCR retry без часового капа + сверка с live-текстом

**Created:** 2026-09-03 10:27  
**Orchestration:** orch-2026-09-03-10-27-kp-ocr-retry  
**Спека:** [../../specs/kp-ocr-retry-no-hourly-cap.md](../../specs/kp-ocr-retry-no-hourly-cap.md)  
**Идея:** [../../ideas/kp-ocr-retry-no-hourly-cap.md](../../ideas/kp-ocr-retry-no-hourly-cap.md)  
**Goal:** Снять часовой кап OCR (`0` = выкл.); кнопка «Перераспознать» заменяет **только текст выбранной страницы**; сверка (и wide у плит) считается от **текущего** текста редактора.  
**Total Tasks:** 10  
**Priority:** High  
**Status:** PLAN ✅ · IMPLEMENT 🔄 (code + focused tests green; awaiting user accept)

SDD: SPECIFY ✅ · PLAN ✅ (human: «ок»). Phase 3 IMPLEMENT locally complete; not marked IMPLEMENT ✅ until user accept.

## Overview

Три дыры одного экрана сверки КП (фото слева, список справа):

1. **Кап.** `COMMERCIAL_OCR_UPLOADS_PER_HOUR` сейчас `Field(default=10, ge=1)` в `core/config/settings.py` — 11-й кадр очереди (до 12) и retry после 10-го дают 429 из `check_commercial_ocr_rate_limit`. Default **0** = лимитер выключен; `ge=0`; при `lim <= 0` не пишем события. Логин и `COMMERCIAL_UPLOAD_MAX_BYTES` не трогаем. Тест с `env=2` остаётся 429.
2. **Retry страницы.** `ProductDraftHandler.update` умеет только `append` | `replace` всего ввода (`app/services/product_draft_handler.py`). Multi-page `recognizePage` в `CommercialOfferWizard.tsx` уже делает first=`replace`, rest=`append` — повторный вызов этого пути **дублирует** страницу или **стирает** остальные. Нужен `POST /api/v1/commercial/drafts/{id}/ocr-page`: тот же OCR (`resolve_source_input` / `extract_text_from_image`), **без persist**. FE `rerunPage` подставляет `normalized_text` только в эту страницу. Single-page (нет multi-сессии) **может** идти существующим `update` + `mode=replace` + тот же `File`.
3. **Live-сверка.** `WidePlatesInlineSection` и гейт «Список верен» читают замороженные `metadata.wide_plate_lines`. После правки `44-15-10п 5` → `34-15-10п 15` карточка всё ещё split’ит старую OCR-строку. Детекция wide на FE от live-строк редактора (плиты); на всех шести шагах — список/баннеры/гейт страницы от live-текста и **нового** OCR-ответа, не от hydrate с 2 строками.

**TDD:** каждый implementation-таск — failing test **до** кода. Не коммитить. Не убивать `./run+logs.sh`. Без новых npm/pip.

## Locked product decisions (must not drift)

| # | Решение |
|---|---------|
| 1 | Cap default **0** = limiter off; `ge=0`; skip `check` if `lim <= 0`. Login limiter и `COMMERCIAL_UPLOAD_MAX_BYTES` **не трогать**. `test_commercial_upload_rate_limit_returns_429` (`env=2`) по-прежнему 429. |
| 2 | Нет 4× upscale, нет смены 8 MiB / `COMMERCIAL_UPLOAD_MAX_BYTES`, нет смены `OCR_MAX_API_CALLS`. |
| 3 | Кнопка «Перераспознать» **под** выбранным фото (под картинкой / подсказкой зума, не в шапке зума) на **всех шести** `*InputStep`. Видна для `ready`, `error` **и** `confirmed`. Disabled **только** если эта страница `running`/`pending` или нет `File`. |
| 4 | Retry **заменяет текст этой страницы целиком**. Типичный кейс: OCR вернул 2 из 10 строк — новый полный список **вместо** двух, не append и не merge с ручными правками. |
| 5 | Retry `confirmed` → статус снова `ready` («Список верен» ещё раз). |
| 6 | Fail OCR → `error`, `File` остаётся; менеджер сам добавляет новое фото в **конец** очереди. Нет мастера «заменить файл». |
| 7 | Multi-page retry **не** `append` и **не** replace-all. Новый `POST .../ocr-page`. Single-page (нет `multiPage.hasStarted`) **может** `update` `mode=replace` + тот же `File`. |
| 8 | Live review от текста редактора на всех шести шагах. Wide-карточка **только у плит**. После retry баннеры/corrections страницы — с **нового** ответа, не stale 2-line hydrate. |
| 9 | Apply wide на сверке: **сначала flush** editor → draft (`updateInput` text, без картинки), затем `resolve_wide_plates` с **live** `sourceLine`. |
| 10 | Коммиты только по просьбе. Не убивать `./run+logs.sh`. Без новых npm/pip. |

**Путь API (open question спеки):** фиксируем `POST /api/v1/commercial/drafts/{draft_id}/ocr-page` как в контракте спеки. Product type **из черновика**, не из form.

## Architecture Decisions

- **Limiter skip, not delete.** `_CommercialOcrUploadLimiter` и `test_commercial_upload_rate_limit_returns_429` остаются. Меняем только default/`ge` и early-return в `check_commercial_ocr_rate_limit`.
- **ocr-page не пишет draft.** Обёртка вокруг `CommercialDraftService.resolve_source_input` / `extract_text_from_image` (как `ProductDraftHandler.update` L82–87, но **без** `_replace_preview`). Ответ узкий: `{ normalized_text, ocr_verify_failed, ocr_corrections }`. Upload path = существующий `prepare_commercial_ocr_upload` (кап при `>0`, magic JPEG/PNG/PDF, 8 MiB не трогаем — лимит байт как сейчас).
- **Retry ≠ pumpQueue.** `useMultiPageRecognize.pumpQueue` при `pending` зовёт `recognizePage` → wizard first=`replace` / rest=`append`. `rerunPage` **никогда** не ставит страницу в `pending` (иначе append-дубли). Отдельный callback; статус `running` только на этой странице, затем `ready` или `error`.
- **Per-page OCR signals.** Уже есть `PageSource.ocrVerifyFailed` (`multiPageSource.ts`, пишется в hook L134). Добавить `ocrCorrections?: OcrCorrection[]`. UI баннера — `resolveActivePageOcrCorrections` по аналогии с `resolveActivePageOcrVerifyFailed` (`lib/ocrVerifyFailed.ts`). Multi retry **не** делает `dispatch({ type: "hydrate-draft" })`.
- **Live wide — overlay, не порт `get_wide_plate_lines`.** MVP: канон/голые марки `ПБ? L-W-нагрузка qty`, ширина дм > 12. `buildAutoSplitSuggestion` сегодня требует префикс `ПБ` (`widePlateSuggestion.ts` L7) — расширить, чтобы `34-15-10п 15` давал split 12+3 с qty **15**. `WidePlatesInlineSection` кормить overlay-draft с live `wide_plate_lines` (синтетические id `live-wide-{i}`). WxL в метрах — out of scope.
- **Apply flush в multi.** Как `handleConfirmBatch` (wizard ~L429–448): слить тексты всех страниц (активная = editor), `update` text `mode=replace` **без** image, затем `resolve_wide_plates` с live `sourceLine`. Не слать только текст текущей страницы — это wipe остальных. После flush серверные id remap’ать по нормализованной строке (решения в `widePlateActions` сейчас ключ = metadata id).

## Tasks Overview

1. **OCR-001** Cap default 0 + skip if `lim<=0` `(feat-be)` — dependsOn: []
2. **OCR-002** `POST .../ocr-page` без persist `(api)` — dependsOn: [OCR-001]
3. **OCR-003** `commercialOfferApi.ocrPage` `(feat-fe)` — dependsOn: [OCR-002]
4. **OCR-004** `rerunPage` в `useMultiPageRecognize` `(feat-fe)` — dependsOn: [OCR-003]
5. **OCR-005** Кнопка «Перераспознать» на `PlateInputStep` `(ui)` — dependsOn: [OCR-004]
6. **OCR-006** `liveWidePlateLines` + suggestion + карточка/гейт плит `(feat-fe)` — dependsOn: [OCR-005]
7. **OCR-007** Live-review баннеры страницы (corrections с retry) `(feat-fe)` — dependsOn: [OCR-004, OCR-005]
8. **OCR-008** Кнопка + live-review на остальных пяти `*InputStep` `(ui)` — dependsOn: [OCR-005, OCR-007]
9. **OCR-009** `handleApplyWidePlates` flush-then-resolve `(feat-fe)` — dependsOn: [OCR-006]
10. **OCR-010** Focused verify + статусы спеки `(chore)` — dependsOn: [OCR-008, OCR-009]

## Dependencies Graph

```
OCR-001 ──► OCR-002 ──► OCR-003 ──► OCR-004 ──► OCR-005 ──► OCR-006 ──► OCR-009 ──┐
                                              │              │                     ├──► OCR-010
                                              │              └──► OCR-007 ──► OCR-008 ──┘
                                              └─────────────────────┘
```

`OCR-007` ∥ `OCR-006` после `OCR-005` (**разные файлы:** 006 = `liveWidePlateLines` + `WidePlatesInlineSection`/`PlateInputStep` wide-гейт; 007 = `ocrVerifyFailed`/corrections helper + баннер в `PlateInputStep`). Если оба правят `PlateInputStep.tsx` в одном срезе — **не параллелить**, сначала 006, потом 007 (или вынести overlay в lib в 006, баннер в 007). Предпочтительно: 006 не трогает блок Alert corrections; 007 не трогает `WidePlatesInlineSection`.

`OCR-006` **не** parallelSafe с `OCR-005` (оба `PlateInputStep.tsx`).

---

## Task List

### Phase 1 — Cap (TDD)

#### Task OCR-001: Cap default 0, skip limiter if `lim <= 0`

**Type:** `feat-be`  
**Priority:** Critical  
**Complexity:** Simple  
**dependsOn:** []  
**parallelSafe:** true  
**needsExplore:** true  
**securitySensitive:** true  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer", "security-auditor"]

**Description:** TDD-first. В `core/config/settings.py` поле `commercial_ocr_uploads_per_hour`: `default=0`, `ge=0` (сейчас `default=10`, `ge=1`, ~L150–154). В `check_commercial_ocr_rate_limit` (`app/services/commercial_upload_validation.py` ~L53–55): `if lim <= 0: return` **до** `_ocr_upload_limiter.check`. Не менять `auth_login_attempts_per_minute`, `COMMERCIAL_UPLOAD_MAX_BYTES`, `OCR_MAX_API_CALLS`. Документировать `0` в `.env.example` (комментарий у rate-limit блока) и строку default в `ai_docs/develop/architecture/rate-limiting.md`.

- **Test first:**
  - `test_commercial_ocr_uploads_per_hour_allows_zero` в `tests/test_settings_guards.py` — `Settings(..., commercial_ocr_uploads_per_hour=0)` не `ValidationError`; default без kwargs = 0; `-1` → `ValidationError`.
  - `test_check_commercial_ocr_rate_limit_skips_when_zero` — `COMMERCIAL_OCR_UPLOADS_PER_HOUR=0`, `get_settings.cache_clear()`, `reset_commercial_ocr_rate_limiter_for_tests()`, 11 вызовов `check_commercial_ocr_rate_limit(1)` без исключения.
  - `test_commercial_upload_rate_limit_zero_allows_eleven_uploads` в `tests/test_commercial_web_flow.py` (рядом с существующим) — env `0`, 11× `POST /api/v1/commercial/from-form` с `_MINIMAL_PNG_BYTES` не 429 (мок `create_draft_from_form` как в `test_commercial_upload_rate_limit_returns_429`).
  - Регресс: **не менять** `test_commercial_upload_rate_limit_returns_429` (env=`2`, третья загрузка 429, detail «Превышен лимит загрузок для распознавания. Попробуйте позже.»).
- **Acceptance:** `0` выключает кап; `>0` считает как сейчас; логин-лимитер и max-bytes не изменены.
- **Verify:**
  ```bash
  pytest tests/test_settings_guards.py tests/test_commercial_web_flow.py tests/test_commercial_ocr_policy.py -q -k "rate_limit or ocr_upload or OCR_UPLOADS or commercial_ocr_uploads"
  pytest tests/test_commercial_web_flow.py -q -k "test_commercial_upload_rate_limit_returns_429"
  ```
- **Files:** `core/config/settings.py`, `app/services/commercial_upload_validation.py`, `tests/test_settings_guards.py`, `tests/test_commercial_web_flow.py`, `.env.example`, `ai_docs/develop/architecture/rate-limiting.md`

---

### Checkpoint: Cap

- [x] Default 0; 11+ загрузок без 429
- [x] env=2 → 3-я загрузка 429 (регресс)
- [x] Login / max-bytes / `OCR_MAX_API_CALLS` не в diff

---

### Phase 2 — ocr-page API (TDD)

#### Task OCR-002: `POST /api/v1/commercial/drafts/{draft_id}/ocr-page`

**Type:** `api`  
**Priority:** Critical  
**Complexity:** Moderate  
**dependsOn:** [OCR-001]  
**parallelSafe:** false  
**needsExplore:** true  
**securitySensitive:** true  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer", "security-auditor"]

**Description:** TDD-first новый endpoint. Auth: `REQUIRE_ADMIN_OR_MANAGER` + `verify_draft_ownership` (как PATCH plates, `commercial.py` ~L319–328). Multipart `image` обязателен. Вызвать `prepare_commercial_ocr_upload`; если `(None, None)` → 400. Product type из `metadata.product_type` черновика (не form). Тонкая обёртка: `ProductDraftHandler.recognize_page` (имя на усмотрение, не `update`) → `resolve_source_input` → вернуть text + флаги из `source_metadata` (`ocr_verify_failed`, `ocr_corrections` как в `_map_ocr_result_metadata`). **Запрещено** `_replace_preview` / запись `input_text` / batches / `wide_plate_lines` / `order_data`. Pydantic: `CommercialOcrPageResponse` в `app/schemas/commercial.py`. Workflow: тонкий `recognize_draft_page` на `CommercialWorkflowService`. Ошибки: 400 нет/битый файл (`require_magic_image_or_pdf`); 404 draft; 413 размер (существующий cap); 503 OCR выкл (`ensure_external_ocr_enabled` внутри prepare); provider 502/503 как `_run_product_ai` (`ai_provider=True`). Мок OCR: `CommercialDraftService.extract_text_from_image` (паттерн `tests/test_commercial_pile_flow.py` `test_create_pile_draft_from_image_mock_ocr`).

- **Test first** (новый `tests/test_commercial_ocr_page.py`; PNG = `_MINIMAL_PNG_BYTES` из `test_commercial_web_flow.py`):
  - `test_ocr_page_returns_normalized_text_and_does_not_mutate_draft` — после create+update с известным `input_text`/batches снимок payload; POST ocr-page с мок-OCR «полный список»; 200 `{normalized_text, ocr_verify_failed, ocr_corrections}`; повторный GET/load: `input_text`, batches, `wide_plate_lines`, `order_data` **байт-в-байт те же**.
  - `test_ocr_page_uses_draft_product_type` — pile draft → `extract_text_from_image(..., product_type="piles")`.
  - `test_ocr_page_404_unknown_draft`
  - `test_ocr_page_400_missing_or_empty_image` / bad magic
  - `test_ocr_page_503_when_ocr_disabled` — `OCR_EXTERNAL_ENABLED=false`
  - `test_ocr_page_403_foreign_draft` — чужой `owner_user_id` (паттерн 403 в `test_commercial_web_flow.py` ~L692)
- **Acceptance:** OCR вызывается; draft не мутируется; контракт спеки; кап при `0` не считает (наследует OCR-001).
- **Verify:**
  ```bash
  pytest tests/test_commercial_ocr_page.py -q
  pytest tests/test_commercial_web_flow.py -q -k "test_commercial_upload_rate_limit_returns_429"
  ```
- **Files:** `app/schemas/commercial.py`, `app/services/product_draft_handler.py`, `app/services/commercial_workflow_service.py`, `app/api/v1/endpoints/commercial.py`, `tests/test_commercial_ocr_page.py`

---

### Checkpoint: API

- [x] ocr-page 200 + no persist
- [x] 400/403/404/503 покрыты
- [x] Кап env=2 регресс зелёный

---

### Phase 3 — FE retry (TDD)

#### Task OCR-003: `commercialOfferApi.ocrPage`

**Type:** `feat-fe`  
**Priority:** High  
**Complexity:** Simple  
**dependsOn:** [OCR-002]  
**parallelSafe:** true (vs OCR-006 lib-only; **не** vs OCR-004)  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** TDD-first. В `frontend/src/features/commercial-offer/api/commercialOfferApi.ts` метод `ocrPage(draftId, image)` → `POST /api/v1/commercial/drafts/${draftId}/ocr-page` с `FormData` (`image` only). Тип ответа: `{ normalized_text: string; ocr_verify_failed: boolean; ocr_corrections: OcrCorrection[] }` (не полный `CommercialDraftDetails`).

- **Test first:** `it("POSTs ocr-page multipart and returns page OCR payload")` в `commercialOfferApi.test.ts` — `httpClient.post` URL и FormData содержат `image`; ответ мапится 1:1.
- **Acceptance:** клиент не вызывает PATCH plates/piles.
- **Verify:**
  ```bash
  cd frontend && npm run test -- src/features/commercial-offer/api/commercialOfferApi
  ```
- **Files:** `frontend/src/features/commercial-offer/api/commercialOfferApi.ts`, `frontend/src/features/commercial-offer/api/commercialOfferApi.test.ts`, при необходимости тип в `types/commercialOffer.ts`

---

#### Task OCR-004: `rerunPage` in `useMultiPageRecognize`

**Type:** `feat-fe`  
**Priority:** Critical  
**Complexity:** Moderate  
**dependsOn:** [OCR-003]  
**parallelSafe:** false  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** TDD-first. Расширить `PageSource` (`multiPageSource.ts`): `ocrCorrections?: OcrCorrection[]`. Hook принимает `rerunPage: (args: { image: File; draftId: string; productType: ProductType }) => Promise<{ normalized_text; ocr_verify_failed; ocr_corrections }>`. Метод `rerunPage(pageId)`:

- Не ставит `pending` (иначе `pumpQueue` → `recognizePage` → **append**).
- Ставит **эту** страницу `running`, зовёт callback с `page.file` + `draftIdRef`.
- Success: `batchReviewText = normalized_text` (**replace**, не конкатенация), `ocrVerifyFailed`, `ocrCorrections`, `status: "ready"` (в т.ч. если было `confirmed` или `error`).
- Fail: `status: "error"`, `errorMessage = getErrorMessage(error)`, **тот же `File`**, прежний текст не обязателен чистить.
- Другие страницы не меняются; `lastDraft` / `draftId` не перезаписывать с «нового» draft (его нет).

Wizard: если `multiPage.hasStarted` — `commercialOfferApi.ocrPage`; иначе (нет multi-сессии) — существующий `updateInputMutation` `mode: "replace"` + тот же `File` и маппинг полей из draft. На multi-retry **не** `hydrate-draft`. Если активная страница — `dispatch({ type: "set-batch-review-text", text })`.

- **Test first** (дописывать `useMultiPageRecognize.test.ts`; мок `rerunPage` отдельно от `recognizePage`):
  - `rerunPage replaces only that page text (does not append)` — страница A `"line-a"`, B две строки; rerun B → полный `"full-b"`; A без изменений; `recognizePage` больше не вызывался.
  - `rerunPage of confirmed page returns status ready`
  - `rerunPage OCR failure keeps File and sets error`
  - `rerunPage must not enqueue pending` — после rerun нет `status==="pending"` у цели.
- **Acceptance:** S3, S4, S5, S9, S10 спеки на уровне hook.
- **Verify:**
  ```bash
  cd frontend && npm run test -- src/features/commercial-offer/hooks/useMultiPageRecognize
  cd frontend && npm run test -- src/features/commercial-offer/lib/multiPageSource
  ```
- **Files:** `frontend/src/features/commercial-offer/lib/multiPageSource.ts`, `frontend/src/features/commercial-offer/hooks/useMultiPageRecognize.ts`, `frontend/src/features/commercial-offer/hooks/useMultiPageRecognize.test.ts`, `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx` (callback wiring only; кнопку UI — OCR-005)

---

#### Task OCR-005: Button «Перераспознать» on `PlateInputStep`

**Type:** `ui`  
**Priority:** High  
**Complexity:** Moderate  
**dependsOn:** [OCR-004]  
**parallelSafe:** false  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** TDD-first RTL в `PlateInputStep.test.tsx`. Пропы: `onRerecognize`, `isRerecognizing` (или вывод из `activePage.status === "running"` + есть File). Кнопка **под** `<img>` / подсказкой «Ctrl + колёсико», ghost/secondary, текст «Перераспознать» / «Распознавание...» (как сниппет в спеке). Показывать при `recognizedImageUrl` и выбранной странице со статусом `ready` | `error` | `confirmed`. `disabled` если эта страница `running`/`pending` или `!page.file` (и при `isRerecognizing`). Не disable из‑за того, что **другая** страница в очереди `pending`. Нет отдельного file-picker.

- **Test first:**
  - `shows Перераспознать under selected photo when page is ready`
  - `shows button for error and confirmed pages`
  - `disables button when page is running or pending`
  - `click calls onRerecognize` (не `onRecognize`)
- **Acceptance:** S8 (plates); существующие PlateInputStep тесты зелёные.
- **Verify:**
  ```bash
  cd frontend && npm run test -- src/features/commercial-offer/components/steps/PlateInputStep
  ```
- **Files:** `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`, `frontend/src/features/commercial-offer/components/steps/PlateInputStep.test.tsx`, `CommercialOfferWizard.tsx` (проп `onRerecognize={() => void multiPage.rerunPage(activeId)}`)

---

### Checkpoint: Retry

- [x] Multi: rerun страницы 2 не дублирует и не трёт страницу 1
- [x] confirmed → ready; fail → error + File
- [x] Кнопка на Plate: ready/error/confirmed; disabled running/pending

---

### Phase 4 — Live wide / live review (TDD)

#### Task OCR-006: `liveWidePlateLines` + plate card/gate

**Type:** `feat-fe`  
**Priority:** Critical  
**Complexity:** Moderate  
**dependsOn:** [OCR-005]  
**parallelSafe:** false  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** TDD-first новый `frontend/src/features/commercial-offer/lib/liveWidePlateLines.ts`: разбор live-строк редактора (канон/`ПБ` и голые `L-W-нагрузка qty`; ширина дм > 12). Не порт WxL из `get_wide_plate_lines`. Overlay: синтетические `{ id: "live-wide-{i}", line, qty }` → подмена `metadata.wide_plate_lines` / `wide_plates_resolved` для `WidePlatesInlineSection` и `hasUnresolvedWidePlates` **на batch-review**. Расширить `buildAutoSplitSuggestion`: строка без `ПБ` (`34-15-10п 15`) → suggestion содержит **12**, **3** (или `3.0`) и **15**, не qty со старого OCR. Если в тексте нет ширины >12 — карточка не рендерится, «Список верен» не блокируется wide. Highlights `kind: "wide"` на сверке — от live keys, не от frozen metadata (`mergeReviewHighlights` / overlay draft в `PlateInputStep`).

- **Test first:**
  - `liveWidePlateLines.test.ts`: `"44-15-10п 5"` → wide qty 5; после замены текста на `"34-15-10п 15"` → wide qty 15; `"34-12-10п 15"` → не wide; `"ПБ 34-15-10п 15"` → wide.
  - `widePlateSuggestion`: `buildAutoSplitSuggestion("34-15-10п 15", 5)` содержит 12, 3, **15** (не использовать fallbackQty=5 как qty split).
  - RTL Plate: editor `34-15-10п 15` при stale `metadata.wide_plate_lines: [{ line: "44-15-10п 5", qty: 5 }]` → карточка/suggestion от **34-15 qty 15**; editor `34-12-10п 15` → нет «Нестандартная ширина», confirm не заблокирован wide.
- **Acceptance:** S6, S7 спеки.
- **Verify:**
  ```bash
  cd frontend && npm run test -- src/features/commercial-offer/lib/liveWidePlateLines
  cd frontend && npm run test -- src/features/commercial-offer/lib/widePlateSuggestion
  cd frontend && npm run test -- src/features/commercial-offer/components/steps/PlateInputStep
  cd frontend && npm run test -- src/features/commercial-offer/components/WidePlatesInlineSection
  ```
- **Files:** `frontend/src/features/commercial-offer/lib/liveWidePlateLines.ts` (new), `liveWidePlateLines.test.ts` (new), `widePlateSuggestion.ts` (+ colocated test **new** если файла теста нет), `PlateInputStep.tsx`, `WidePlatesInlineSection.tsx` только если без overlay не обойтись (предпочтительно не менять контракт секции)

---

#### Task OCR-007: Live-review banners from the new OCR response

**Type:** `feat-fe`  
**Priority:** High  
**Complexity:** Moderate  
**dependsOn:** [OCR-004, OCR-005]  
**parallelSafe:** true vs OCR-006 **если** не правит wide-блок в `PlateInputStep`  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** TDD-first. Сейчас баннер corrections = `draft?.metadata.ocr_corrections` (`PlateInputStep` ~L198). После ocr-page draft не обновляется → stale 2-line hydrate. Helper рядом с `ocrVerifyFailed.ts`: `resolveActivePageOcrCorrections(pages, activePageId, draftFallback)`. После retry страница показывает **новые** corrections (или пусто) и `ocrVerifyFailed` с ответа. Unparsed по-прежнему live lint (`useBatchReviewHighlights` / `useSourceTextLint`) — не перепроектировать. AI-инструкция / unpriced — не трогать.

- **Test first:**
  - `ocrVerifyFailed.test.ts` (или `ocrPageSignals.test.ts`): active page с `ocrCorrections: [{action:"replaced",...}]` побеждает `draft.metadata.ocr_corrections` от старого прохода.
  - RTL Plate: после пропов страницы с новым `ocrCorrections` / `ocrVerifyFailed` баннеры соответствуют странице, не draft с 2 строками.
- **Acceptance:** S8/S9 баннеры; `resolveActivePageOcrVerifyFailed` регресс зелёный.
- **Verify:**
  ```bash
  cd frontend && npm run test -- src/features/commercial-offer/lib/ocrVerifyFailed
  cd frontend && npm run test -- src/features/commercial-offer/components/steps/PlateInputStep
  ```
- **Files:** `frontend/src/features/commercial-offer/lib/ocrVerifyFailed.ts` (или новый `ocrPageSignals.ts`), `PlateInputStep.tsx` (источник `ocrCorrections`)

---

### Checkpoint: Live review (plates)

- [x] Wide/qty от editor `34-15-10п 15`
- [x] Ширина ≤12 дм — нет wide-гейта
- [x] Corrections/verify после retry с нового ответа

---

### Phase 5 — All six input steps (TDD)

#### Task OCR-008: Wire retry + live-review on remaining `*InputStep`

**Type:** `ui`  
**Priority:** High  
**Complexity:** Moderate  
**dependsOn:** [OCR-005, OCR-007]  
**parallelSafe:** false  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** Тот же паттерн кнопки (под фото) + per-page corrections/verify, что на Plate, во: `PileInputStep.tsx`, `StepInputStep.tsx`, `MarchInputStep.tsx`, `BridgePileInputStep.tsx`, `FbsInputStep.tsx`. **Не** добавлять `WidePlatesInlineSection` на не-плиты. Wizard уже рендерит все шесть шагов (~L1308–1525) — пробросить `onRerecognize` / flags одинаково. RTL минимум: существующий `MarchInputStep.test.tsx` + один не-плитный сценарий кнопки (ready/confirmed). Если общего компонента фото-карточки нет — копировать кнопку как CTA очереди в плане 2026-09-02, не рефакторить шесть карточек целиком.

- **Test first:**
  - `MarchInputStep.test.tsx`: кнопка «Перераспознать» на `ready` и `confirmed`; нет wide-карточки.
  - Plate регресс из OCR-005/006/007 зелёный.
- **Acceptance:** S8 — кнопка на всех шести; live-wide только плиты; live-текст сверки на всех (lint уже есть; corrections с страницы).
- **Verify:**
  ```bash
  cd frontend && npm run test -- src/features/commercial-offer/components/steps
  ```
- **Files:** пять `*InputStep.tsx` + `MarchInputStep.test.tsx`; `CommercialOfferWizard.tsx` props

---

### Phase 6 — Apply flush (TDD)

#### Task OCR-009: Flush editor then `resolve_wide_plates` with live source lines

**Type:** `feat-fe`  
**Priority:** Critical  
**Complexity:** Moderate  
**dependsOn:** [OCR-006]  
**parallelSafe:** false  
**needsExplore:** true  
**securitySensitive:** false  
**pipeline:** ["explore", "worker", "test-writer", "test-runner", "reviewer"]

**Description:** TDD-first вынести логику из `handleApplyWidePlates` (`CommercialOfferWizard.tsx` ~L794–812). Сейчас шлётся `sourceLine: item.line` из **замороженных** `currentDraft.metadata.wide_plate_lines`. Нужно:

1. Если multi-сессия: как confirm (~L429–448) — snapshot текстов страниц с live editor для active, `updateInput` `text=merged`, `image=null`, `mode="replace"`.
2. Single-page: `updateInput` текущим `batchReviewText` (полный ввод), без image, `mode="replace"`.
3. Затем `resolveWidePlatesMutation` с `sourceLine` = **live** строки (после flush сервер пересчитает `wide_plate_lines`; решения матчить по `normalizeLineKey`, не по старому id). Действие `confirm` (оставить ширину) по-прежнему валидно.

Не вызывать resolve, пока flush не успел (иначе id/line mismatch).

- **Test first:** unit helper (предпочтительно `lib/flushThenResolveWidePlates.ts` или рядом с `liveWidePlateLines.ts`):
  - `builds resolve payload from live lines after flush text` — editor `34-15-10п 15`, stale metadata line `44-15-10п 5` → `sourceLine` live, не OCR.
  - `multi-page merge does not drop other pages` — две страницы, apply на второй → merged содержит текст обеих.
  - Wizard/hook test если helper слишком тонкий: мок `updateInput` вызван **до** `resolveWidePlates`.
- **Acceptance:** S6 на Apply path; confirm 15 дм не ломается.
- **Verify:**
  ```bash
  cd frontend && npm run test -- src/features/commercial-offer/lib/liveWidePlateLines
  cd frontend && npm run test -- src/features/commercial-offer/lib/flushThenResolveWidePlates
  # или colocated wizard helper test
  cd frontend && npm run typecheck
  ```
- **Files:** `CommercialOfferWizard.tsx` `handleApplyWidePlates`; new helper + test under `frontend/src/features/commercial-offer/lib/`

---

### Checkpoint: Apply + six steps

- [x] Apply: flush затем resolve с live `sourceLine`
- [x] Кнопка retry на всех шести шагах
- [x] Wide-карточка только у плит

---

### Phase 7 — Verify + docs

#### Task OCR-010: Focused verify + spec/plan status

**Type:** `chore`  
**Priority:** Medium  
**Complexity:** Simple  
**dependsOn:** [OCR-008, OCR-009]  
**parallelSafe:** false  
**needsExplore:** false  
**securitySensitive:** false  
**pipeline:** ["test-runner", "documenter"]

**Description:** Прогнать команды Testing Strategy спеки. После зелёных тестов: idea/spec IMPLEMENT ⬜ до user accept; PLAN ✅ только если human уже approved план **и** код совпал. Не коммитить.

- **Test first:** n/a (verify-only). Не писать новый прод-код.
- **Acceptance:** S1–S11 команд из спеки зелёные локально; typecheck зелёный.
- **Verify:**
  ```bash
  pytest tests/test_commercial_web_flow.py tests/test_commercial_ocr_policy.py tests/test_settings_guards.py tests/test_commercial_ocr_page.py -q -k "rate_limit or ocr_upload or OCR_UPLOADS or ocr_page or commercial_ocr_uploads"
  pytest tests/test_commercial_ocr_page.py -q
  cd frontend && npm run test -- \
    src/features/commercial-offer/lib/widePlateSuggestion \
    src/features/commercial-offer/lib/liveWidePlateLines \
    src/features/commercial-offer/hooks/useMultiPageRecognize \
    src/features/commercial-offer/components/steps/PlateInputStep \
    src/features/commercial-offer/components/steps/PileInputStep \
    src/features/commercial-offer/components/WidePlatesInlineSection \
    src/features/commercial-offer/api/commercialOfferApi
  cd frontend && npm run typecheck
  ```
- **Files:** `ai_docs/specs/kp-ocr-retry-no-hourly-cap.md`, `ai_docs/ideas/kp-ocr-retry-no-hourly-cap.md`, this plan Progress

---

## Progress (updated by orchestrator)

- ✅ OCR-001: Cap default 0 `(feat-be)` (Completed)
- ✅ OCR-002: POST ocr-page `(api)` (Completed)
- ✅ OCR-003: commercialOfferApi.ocrPage `(feat-fe)` (Completed)
- ✅ OCR-004: rerunPage hook `(feat-fe)` (Completed)
- ✅ OCR-005: PlateInputStep button `(ui)` (Completed)
- ✅ OCR-006: liveWidePlateLines `(feat-fe)` (Completed)
- ✅ OCR-007: Live-review banners `(feat-fe)` (Completed)
- ✅ OCR-008: Wire five peer InputSteps `(ui)` (Completed)
- ✅ OCR-009: Apply flush-then-resolve `(feat-fe)` (Completed)
- ✅ OCR-010: Focused verify + docs `(chore)` (Completed)

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| **`rerunPage` ставит `pending` → `pumpQueue` → `recognizePage` append** | High — дубли позиций в draft | Никогда не `pending` на retry; отдельные тесты «recognizePage not called»; wizard multi path только `ocrPage` |
| **Multi retry через `mode=replace`** | High — wipe остальных страниц | Replace **только** если `!multiPage.hasStarted`; иначе ocr-page |
| **`hydrate-draft` после ocr-page** | High — editor/баннеры со старого last-batch (2 строки) | Multi retry: только `set-batch-review-text` + per-page flags |
| **confirmed → ready забыли** | Med — «Список верен» на устаревшем списке | Hook test: confirmed rerun → `ready`; `canConfirmActivePage` уже `status === "ready"` |
| **Stale wide metadata / id remap** | High — Apply ищет OCR-line `44-15-10п 5` | Live overlay на UI; Apply: flush merge, затем `sourceLine` live + match по нормализованной строке |
| **Flush только текущей страницы** | High — replace всего ввода одной страницей | Merge всех `batchReviewText` как `handleConfirmBatch` |
| **Параллельный retry vs pump другой страницы** | Low/Med — два OCR | Спека разрешает; disable только **этой** страницы `running`/`pending` |
| **`buildAutoSplitSuggestion` не матчит голую марку** | Med — fallback `ПБ 60-12-8п` | Расширить regex в OCR-006; тест `34-15-10п 15` |
| **Шесть InputStep разъедутся** | Med | OCR-008 checklist; одинаковые имена пропов |
| **Кап 0 в проде без Redis** | Accepted | Уже in-process; продукт явно просил выкл. Не трогать login limiter |

## Parallel vs sequential

- **Sequential critical path:** OCR-001 → 002 → 003 → 004 → 005 → 006 → 009 и 005 → 007 → 008 → 010.
- **Parallel-safe после OCR-005:** OCR-006 и OCR-007, если не пересекаются hunks `PlateInputStep` (wide vs banners). Иначе строго 006 затем 007.
- **OCR-003** можно готовить сразу после контракта OCR-002 (маленький клиент).
- Не параллелить два агента на `CommercialOfferWizard.tsx` (004 wiring, 005 props, 009 Apply).

## Verification checkpoints

1. After OCR-001: cap 0 / cap 2 регресс.
2. After OCR-002: ocr-page persist identity + error codes.
3. After OCR-004–005: retry без append/wipe; кнопка Plate.
4. After OCR-006–007: live wide + fresh banners на плитах.
5. After OCR-008–009: шесть шагов + Apply flush.
6. After OCR-010: команды спеки + typecheck.

## Out of scope (from spec)

- Апскейл 4×, снятие 8 MiB, смена `OCR_MAX_API_CALLS`
- Поднять cap до 40/200 вместо нуля
- Retry через существующий `append`
- Merge ручных правок с новым OCR / авто-retry
- In-place «заменить файл» в слоте; мастер ошибки OCR
- Порт полного `get_wide_plate_lines` (WxL метры) на FE
- Live-wide на не-плитах; unpriced live-sync
- Redis/shared rate limit; новый npm/pip
- Хранение фото на сервере / IndexedDB
- Переписывать жёлтые OCR-corrections по `row_index` как отдельную фичу (только подставить новый массив ответа)
- Коммит секретов / `plita.db`; трогать `./run+logs.sh`

## Implementation Notes for workers

- Inject `plan-web-context` + this plan + spec `kp-ocr-retry-no-hourly-cap.md`.
- **Every behavior task: write the failing test named above, run it (RED), then implement (GREEN).**
- Do not kill `./run+logs.sh`; do not commit unless asked; no new npm/pip.
- Reuse: `prepare_commercial_ocr_upload`, `resolve_source_input`, `getErrorMessage`, `resolveActivePageOcrVerifyFailed`, `filterDraftForBatchReview` (не ломать lint), `handleConfirmBatch` merge as flush template.
- User checkpoint after this plan before any code.

## Open Questions / Blockers

_None blocking._ Path locked: `/ocr-page`. PLAN ⏳ until human review.
