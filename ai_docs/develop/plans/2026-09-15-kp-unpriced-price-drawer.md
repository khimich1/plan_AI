# Implementation Plan: КП — прайс слева и договорная цена на «нет в прайсе»

**Created:** 2026-09-15 16:25  
**Спека:** [../../specs/kp-unpriced-price-drawer.md](../../specs/kp-unpriced-price-drawer.md)  
**Идея:** [../../ideas/kp-unpriced-price-drawer.md](../../ideas/kp-unpriced-price-drawer.md)  
**Goal:** На шаге 1 при «нет в прайсе» менеджер открывает справочник слева (поиск по группе) и ставит договорную ₽/шт только в этот черновик. Заводской прайс не пишется.  
**Total Tasks:** 8  
**Priority:** High  
**Status:** IMPLEMENT ✅ (UPD-001…008; браузер — ручной остаток)

## Overview

Два вертикальных среза: (1) read-only каталог `GET /price-catalog` + left Drawer; (2) `PATCH .../lines/{id}` с `unit_price` / `price_source=oneoff`, поле на красной строке, гейт wizard сам разблокируется. PDF/XLSX и скидка уже берут `unit_price` со строки — отдельный pricing-path не нужен.

## Architecture Decisions

- **Каталог — один модуль SELECT**, не пять `list_*` в `*_price_db.py`. `core/price_catalog_query.py`: whitelist `product_type` → таблица; `price > 0`; фильтр `q` в Python после нормализации C↔С/пробелы/регистр; кап 2000 → `ValueError` (роутер → 400).
- **GET без draft_id.** Справочник общий. Auth: `Depends(REQUIRE_ADMIN_OR_MANAGER)` как у `/parse`. Владение черновиком не проверяем.
- **`unit_price` в PATCH через `model_fields_set`.** Omit ≠ explicit `null`. Сброс договорной — только `{ "unit_price": null }` при уже `price_source=oneoff`.
- **`price_source` только на `order_data`.** Preview-builders не трогаем: панель и шаг 3 читают флаг с `draft.order_data` по `line_id`.
- **Кнопка и ввод — `KpGradedPreviewPanel`.** Плиты не получают UI. Drawer — новый `PriceCatalogDrawer`, `Drawer side="left"`.
- **Поиск при открытии:** пин дырок из draft + префилл `q` общим префиксом непрорасценённых марок (длина ≥ 4), иначе пустой `q`.
- **Oneoff не в карандаш и не в undo-тосте qty/марки.** Отдельное поле в ячейке «Цена». Сброс — то же поле (пустое → PATCH null).
- **Sealed** (`append_batch_id` непустой) → 400 на oneoff. Кнопка «Прайс» смотрит только незапечатанные `unit_price === null`.
- **Без записи в прайс-таблицы.** Тест S5: после oneoff SELECT count/imported_at не вырос.
- **Не убивать `./run+logs.sh`.** Коммиты — по просьбе. Новых npm/pip нет.

## Tasks Overview

1. **UPD-001** `price_catalog_query` SELECT + поиск + кап `(feat-be)` — dependsOn: []
2. **UPD-002** `GET /price-catalog` `(api)` — dependsOn: [UPD-001]
3. **UPD-003** PATCH `unit_price` oneoff `(feat-be)` — dependsOn: []
4. **UPD-004** FE API + типы каталога/PATCH `(feat-fe)` — dependsOn: [UPD-002, UPD-003]
5. **UPD-005** `PriceCatalogDrawer` + префилл поиска `(ui)` — dependsOn: [UPD-004]
6. **UPD-006** Кнопка «Прайс» + поле договорной в панели `(ui)` — dependsOn: [UPD-005]
7. **UPD-007** Wizard + шаг 3 «договорная» `(feat-fe)` — dependsOn: [UPD-006]
8. **UPD-008** Focused verify + статус спеки `(chore)` — dependsOn: [UPD-007]

## Dependencies Graph

```
UPD-001 ──► UPD-002 ──┐
                      ├──► UPD-004 ──► UPD-005 ──► UPD-006 ──► UPD-007 ──► UPD-008
UPD-003 ──────────────┘
```

`UPD-001` ∥ `UPD-003` (`parallelSafe` между ветками).

---

## Task List

### Phase 1 — Backend (TDD)

#### Task UPD-001: `core/price_catalog_query.py`

**Type:** `feat-be`  
**Priority:** Critical  
**Complexity:** Simple  
**dependsOn:** []  
**parallelSafe:** true  
**needsExplore:** true  
**securitySensitive:** false  

**Description:** TDD: `list_price_catalog(product_type, q, db_path) -> list[dict]`. Типы: `piles`, `bridge_piles`, `fbs`, `marches`, `steps`. Неизвестный тип → `ValueError`. Строки с `price <= 0` не отдаём. Пустой `q` — все; непустой — подстрока нормализованной марки. Поля: `mark`, `concrete_grade` (`None` у ступеней), `price`. Кап 2000.

**Acceptance criteria:**
- [x] Seed `pile_prices`: `q=C110` → есть `С110.30-9`, нет `С110.30-6`
- [x] Пустой `q` на маленькой фикстуре не падает
- [x] `steps` без колонки класса; неизвестный type → ошибка
- [x] Кап: monkeypatch лимита или >cap → `ValueError`

**Verification:**
```
pytest tests/test_price_catalog_query.py -q
```

**Files likely touched:**
- `core/price_catalog_query.py` (new)
- `tests/test_price_catalog_query.py` (new)

**Estimated scope:** S

---

#### Task UPD-002: GET `/api/v1/commercial/price-catalog`

**Type:** `api`  
**Priority:** Critical  
**Complexity:** Simple  
**dependsOn:** [UPD-001]  
**parallelSafe:** false  
**needsExplore:** true  
**securitySensitive:** false  

**Description:** Endpoint `product_type` + `q` (optional). Auth `REQUIRE_ADMIN_OR_MANAGER`. Ответ `{ items: [...] }`. `ValueError` каталога → 400, русский `detail`. `db_path` — штатный PRICE_DB (в тестах monkeypatch как в line-edit).

**Acceptance criteria:**
- [x] 200 + items на seeded piles `q=C110`
- [x] 401/403 без роли менеджера/админа (как соседние commercial GET)
- [x] 400 на неизвестный `product_type` и на кап
- [x] Роутер не пишет в SQLite (только SELECT)

**Verification:**
```
pytest tests/test_price_catalog_api.py -q
```

**Files likely touched:**
- `app/schemas/commercial.py`
- `app/api/v1/endpoints/commercial.py`
- `tests/test_price_catalog_api.py` (new)

**Estimated scope:** S

---

#### Task UPD-003: PATCH oneoff `unit_price`

**Type:** `feat-be`  
**Priority:** Critical  
**Complexity:** Moderate  
**dependsOn:** []  
**parallelSafe:** true  
**needsExplore:** true  
**securitySensitive:** false  

**Description:** Расширить `CommercialDraftLinePatchRequest` и `patch_order_line`. `unit_price` в body учитывается iff `"unit_price" in model_fields_set`. Правила спеки: `> 0`, ≤ 10_000_000, finite; только `unit_price is None` или уже `oneoff`; не sealed; не вместе с `source_text`; прайсовая строка → 400 «Цена из прайса. Для изменения используйте скидку.» Успех: `price_source=oneoff`, `line_total`, persist+totals. `null` на oneoff снимает цену и флаг. Гейт: `unpriced_position_labels` пуст после oneoff. Qty-only регресс без изменений.

**Acceptance criteria:**
- [x] Oneoff на null-строке сваи → 200, флаг, гейт пуст
- [x] Правка oneoff; сброс `null` → снова нет цены
- [x] 400: прайсовая строка; `source_text`+цена; sealed; qty=0 как сейчас
- [x] qty+oneoff в одном PATCH сохраняет оба
- [x] После oneoff count строк в `pile_prices` тот же

**Verification:**
```
pytest tests/test_commercial_draft_line_edit.py tests/test_commercial_draft_oneoff_price.py -q
```

**Files likely touched:**
- `app/schemas/commercial.py`
- `app/services/commercial_draft_lifecycle.py`
- `app/api/v1/endpoints/commercial.py`
- `app/services/commercial_workflow_service.py` (прокинуть kwargs, если нужно)
- `tests/test_commercial_draft_oneoff_price.py` (new)

**Estimated scope:** M

---

### Checkpoint: Backend

- [x] Каталог читается, oneoff пишется только в черновик
- [x] `pytest` точечные команды Phase 1 зелёные
- [x] Регресс qty/source_text зелёный

### Phase 2 — Frontend

#### Task UPD-004: FE client + типы

**Type:** `feat-fe`  
**Priority:** High  
**Complexity:** Simple  
**dependsOn:** [UPD-002, UPD-003]  
**parallelSafe:** false  

**Description:** `getPriceCatalog({ productType, q })`. `patchDraftLine` принимает `unit_price?: number | null`. Zod/типы ответа каталога. Тесты api.

**Acceptance criteria:**
- [x] GET query `product_type` + `q`
- [x] PATCH JSON может содержать `unit_price: null`
- [x] vitest `commercialOfferApi` зелёный

**Verification:**
```
cd frontend && npm run test -- src/features/commercial-offer/api/commercialOfferApi
```

**Files likely touched:**
- `frontend/src/features/nomenclature/...` — **не трогать**
- `frontend/src/features/commercial-offer/api/commercialOfferApi.ts`
- `frontend/src/features/commercial-offer/api/commercialOfferApi.test.ts`
- `frontend/src/features/commercial-offer/types/commercialOffer.ts`

**Estimated scope:** S

---

#### Task UPD-005: `PriceCatalogDrawer`

**Type:** `ui`  
**Priority:** High  
**Complexity:** Moderate  
**dependsOn:** [UPD-004]  
**parallelSafe:** false  

**Description:** Left `Drawer`, заголовок «Прайс», input поиска, блок пина (марки из пропса `missingMarks`), таблица `items` (марка, класс если есть, цена). Подсветка stem. Read-only. Pure helper `commonMarkSearchPrefix(marks)` (нормализация, общий префикс ≥ 4).

**Acceptance criteria:**
- [x] RTL: пин C110.30-6 при отсутствии в items
- [x] C110.30-9 в таблице с ценой
- [x] Поиск фильтрует (controlled `q` / onChange)
- [x] Esc/оверлей закрывают
- [x] `commonMarkSearchPrefix(["C110.30-6","C110.40-8.1"])` → `C110`

**Verification:**
```
cd frontend && npm run test -- src/features/commercial-offer/components/PriceCatalogDrawer src/features/commercial-offer/lib/commonMarkSearchPrefix
```

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/PriceCatalogDrawer.tsx` (new)
- `frontend/src/features/commercial-offer/components/PriceCatalogDrawer.test.tsx` (new)
- `frontend/src/features/commercial-offer/lib/commonMarkSearchPrefix.ts` (new)
- `frontend/src/features/commercial-offer/lib/commonMarkSearchPrefix.test.ts` (new)

**Estimated scope:** M

---

#### Task UPD-006: Панель состава

**Type:** `ui`  
**Priority:** Critical  
**Complexity:** Moderate  
**dependsOn:** [UPD-005]  
**parallelSafe:** false  

**Description:** В `KpGradedPreviewPanel`: `Card.actions` кнопка «Прайс» если есть незапечатанные `unit_price === null`. Drawer. Красная ячейка цены — input ₽/шт + сохранить (и сброс). После oneoff (`price_source` с draft по `lineId`) — число + «договорная», алерт `previewUnpricedMessage` скрыт если дырок не осталось. Пропсы колбэков, без wizard.

**Acceptance criteria:**
- [x] Нет красных незапечатанных → кнопки нет
- [x] Кнопка открывает Drawer с пином
- [x] Сохранение цены зовёт `onSetOneoffPrice(lineId, number)`
- [x] `price_source=oneoff` → не «нет в прайсе», алерт снят
- [x] Плиты не затронуты (`KpPlatePreviewPanel` без кнопки)

**Verification:**
```
cd frontend && npm run test -- src/features/commercial-offer/components/KpGradedPreviewPanel
```

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/KpGradedPreviewPanel.tsx`
- `frontend/src/features/commercial-offer/components/KpGradedPreviewPanel.test.tsx`

**Estimated scope:** M

---

#### Task UPD-007: Wizard + шаг результата

**Type:** `feat-fe`  
**Priority:** High  
**Complexity:** Moderate  
**dependsOn:** [UPD-006]  
**parallelSafe:** false  

**Description:** Прокинуть в `SimpleProductInputStep` → панель: загрузка каталога при open (`productType` шага, `q` из префилла/ввода), PATCH oneoff, hydrate draft, ошибка у строки. На `CalculationResultStep` у `price_source=oneoff` подпись «договорная» (скидка по-прежнему через `discountedUnitPrice`). Кнопки «Прайс» на шаге 3 нет.

**Acceptance criteria:**
- [x] Open Drawer → GET catalog
- [x] Submit oneoff → PATCH → hydrate, validation «Нет цен…» уходит
- [x] Шаг 3 показывает «договорная» только на oneoff
- [x] Карандаш qty/марки не сломан

**Verification:**
```
cd frontend && npm run test -- src/features/commercial-offer/components/steps/SimpleProductInputStep src/features/commercial-offer/components/steps/CalculationResultStep src/features/commercial-offer/hooks/useCommercialOfferWizard
cd frontend && npm run typecheck
```

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/steps/SimpleProductInputStep.tsx`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- `frontend/src/features/commercial-offer/hooks/useCommercialOfferWizard.ts`
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.tsx`
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx`

**Estimated scope:** M (чуть плотно; не раздувать wizard сверх прокидки)

---

### Checkpoint: Frontend

- [x] Срез C110: пин + сосед + договорная + алерт снят
- [x] typecheck зелёный

### Phase 3 — Verify

#### Task UPD-008: Focused verify + docs

**Type:** `chore`  
**Priority:** Medium  
**Complexity:** Simple  
**dependsOn:** [UPD-007]  
**parallelSafe:** false  

**Description:** Прогнать команды спеки. Статус спеки/плана → IMPLEMENT по факту зелёного прогона (не раньше). Не коммитить.

**Acceptance criteria:**
- [x] Команды ниже зелёные
- [x] S1–S10 спеки закрыты или явно отмечен ручной остаток (браузер)

**Verification:**
```
pytest tests/test_price_catalog_query.py tests/test_price_catalog_api.py tests/test_commercial_draft_oneoff_price.py tests/test_commercial_draft_line_edit.py -q
cd frontend && npm run test -- src/features/commercial-offer/components/KpGradedPreviewPanel src/features/commercial-offer/components/PriceCatalogDrawer src/features/commercial-offer/api/commercialOfferApi
cd frontend && npm run typecheck
```

**Files likely touched:**
- `ai_docs/specs/kp-unpriced-price-drawer.md`
- `ai_docs/develop/plans/2026-09-15-kp-unpriced-price-drawer.md`

**Estimated scope:** S

---

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| `model_fields_set` забыт, `null` не сбрасывает | Med | Тест explicit null vs omit |
| Oneoff сгорает на смене класса | Low | Зафиксировано в спеке; не «чинить» restore |
| Drawer 1590 строк тормозит | Low | Префилл C110; кап 2000 |
| UPD-007 раздует wizard | Med | Только прокидка catalog/oneoff; логика в панели |
| Случайно INSERT в прайс | High | S5 count; каталог — чистый SELECT |

## Open Questions

Нет. Скидка на oneoff как на прайс.

## Stop

Implement не начинать, пока план не подтвердят.
