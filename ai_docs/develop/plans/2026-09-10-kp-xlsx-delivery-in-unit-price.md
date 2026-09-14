# Implementation Plan: КП — XLSX с доставкой в цене изделия

**Created:** 2026-09-10  
**Status:** PLAN ✅ · IMPLEMENT done  
**Spec:** [`ai_docs/specs/kp-xlsx-delivery-in-unit-price.md`](../../specs/kp-xlsx-delivery-in-unit-price.md)  
**Idea:** [`ai_docs/ideas/kp-xlsx-delivery-in-unit-price.md`](../../ideas/kp-xlsx-delivery-in-unit-price.md)

## Overview

Менеджер в карточке архива скачивает второй Excel того же КП: без строк доставки, с той же суммой к оплате, доставка (уже с НДС, без скидки) равномерно на штуку после скидки. Обычные PDF/XLSX, карточка и БД не меняются. Сначала чистая арифметика, затем генератор, затем `GET .../files/xlsx_delivery_in_unit`, затем кнопка.

Это не горизонтальный «весь бэкенд потом весь UI»: после Task 4 уже есть проверяемый файл; после Task 5 менеджер может скачать его через API; Task 6 только открывает этот путь в drawer.

## Architecture Decisions

- **A1. Один helper — единственное место арифметики.** `core/embed_delivery_in_unit_price.py`. Генератор и архив не делят доставку сами. `calculate_total_cost` не трогаем: обычный НДС и рейсы остаются как есть.
- **A2. Два котла.** `plate_delivery_total` → `plates` (legacy без type = plates). `pile_delivery_total` → `piles` + `bridge_piles`. Тип строки — `line_product_type`. ФБС / ступени / марши надбавки не получают.
- **A3. Итог файла = `totals["total_with_vat"]`.** Доли в копейках (`divmod` + extra-штуки по порядку строк). Σ `line_sum` товарных вшитых строк + суммы оставленных строк доставки (если котёл не вшился) = `total_with_vat` (± 0,01 ₽). Если у котла delivery > 0 и `total_qty == 0` — котёл не вшивать, оставить обычную строку доставки.
- **A4. Колонка «Сумма» во вшитом режиме — число, не формула `qty*Цена`.** Сейчас openpyxl всегда перезаписывает «Сумма» формулой (`commercial_offer_xlsx.py` ~503–508). В embed-режиме эту перезапись пропустить, иначе копейки разъедутся.
- **A5. НДС вшитого листа.** Строк доставки нет (или только невшитый котёл) → существующая формула `SUM(товарные)*0.22` сама берёт налог с надбавки. Обычный XLSX по-прежнему вырезает delivery-строки из НДС.
- **A6. Точка входа только архив.** Флаг `embed_delivery_in_unit_price: bool = False` на `generate_commercial_offer_xlsx`. `FileGenerationService` / визард флаг не передают. Kind: `xlsx_delivery_in_unit`. Файл: `КП_{id}_с_доставкой_в_цене.xlsx`.
- **A7. Пояснение** под блоком НДС, только если вшили хотя бы один котёл.
- **A8. Кнопка disabled** при `delivery_service_total_rub === 0`. Обычный XLSX живой.
- **A9. Без схемы БД, без PDF, без визарда, без UPDATE цен.**

```
embed_delivery_in_unit_prices
        │
        └── generate_commercial_offer_xlsx(flag)
                    │
                    └── ArchiveService.generate_document(kind)
                                │
                                └── GET .../files/xlsx_delivery_in_unit
                                            │
                                            └── OfferDetailsDrawer button
```

Wizard export и `GET /offers/{id}/xlsx` в граф не входят.

## Implementation order

| Phase | Focus | Depends |
|-------|-------|---------|
| 1 | Pure math + pytest | — |
| 2 | Флаг генератора XLSX | 1 |
| 3 | Archive kind + HTTP | 2 |
| 4 | Кнопка в drawer | 3 (контракт kind уже в спеке — UI-тесты можно параллелить с 2–3) |
| 5 | Mixed / край / регрессия | 2–4 |

## Task List

### Phase 1: Foundation

## Task 1: RED — контракт `embed_delivery_in_unit_prices`

**Description:** Зафиксировать арифметику тестами до кода helper. Возврат: список строк + какие котлы вшиты. Скидка только на изделие; два котла независимы; extra-копейки; delivery 0; пустой котёл с delivery > 0 → котёл не вшит.

**Acceptance criteria:**
- [x] Скидка 50%, изделие 122, доставка 100, qty 1 → `line_sum == 161`, надбавка 100, не 50
- [x] Mixed: plate_delivery только на plates qty, pile_delivery только на piles/bridge_piles qty; Σ надбавок = сумма котлов
- [x] `divmod`: delivery 100.01 (10001 коп), qty 3 → суммы строк в копейках дают ровно 10001
- [x] delivery 0 → надбавка 0, `line_sum` = изделие после скидки
- [x] plate_delivery > 0 при нулевом qty плит → `embedded_plate is False`
- [x] Тесты красные: модуля ещё нет

**Verification:**
- [x] `pytest tests/test_embed_delivery_in_unit_price.py -q` — fail import / fail assert, не skip

**Dependencies:** None

**Files likely touched:**
- `tests/test_embed_delivery_in_unit_price.py`

**Estimated scope:** S

## Task 2: GREEN — helper

**Description:** Реализовать `core/embed_delivery_in_unit_price.py` по спеке: копейки, два котла, `line_product_type`-совместимые ключи (`plates` / `piles` / `bridge_piles`). Округление скидки на единицу: `discounted_unit_k = round(unit_price * (1 - d/100) * 100)` как в спеке; Σ `line_sum` вшитых строк котла = изделия_котла_после_скидки + delivery_котла (в копейках).

**Acceptance criteria:**
- [x] Все тесты Task 1 зелёные
- [x] Нет I/O, нет Excel, нет вызова `calculate_total_cost`

**Verification:**
- [x] `pytest tests/test_embed_delivery_in_unit_price.py -q`

**Dependencies:** Task 1

**Files likely touched:**
- `core/embed_delivery_in_unit_price.py`
- `tests/test_embed_delivery_in_unit_price.py` (правки ассертов только если всплыл round-edge — сверить со спекой, не с генератором)

**Estimated scope:** S

### Checkpoint: Foundation

- [x] Математика зелёная без XLSX
- [x] Обычные `test_commercial_logistics_cost.py` не запускались и не ломались — helper изолирован
- [x] Не переходить к генератору, пока extra-копейки и два котла не зелёные

---

### Phase 2: Core artifact (вертикальный срез «есть файл»)

## Task 3: RED — вшитый XLSX vs обычный

**Description:** Тесты на `generate_commercial_offer_xlsx(..., embed_delivery_in_unit_price=True)`: читать лист `КП` через pandas/openpyxl. Кейс плит со скидкой и доставкой (как `test_calculate_total_cost_applies_discount_only_to_products`). Кейс mixed с двумя доставками. Регрессия: вызов **без** флага по-прежнему содержит «доставк» в наименованиях.

**Acceptance criteria:**
- [x] Флаг on, котлы вшились: в «Наименование» нет «доставк»
- [x] Σ колонки «Сумма» по товарным строкам = `calculate_total_cost(...)["total_with_vat"]` (± 0,01)
- [x] Ячейка «в том числе НДС (22%)» = эта Σ × 0.22 (± 0,01) — читать значение формулы через openpyxl data_only **или** проверить формулу `SUM` по всем товарным строкам (без вычета delivery), не мокать
- [x] Есть текст «В стоимость изделий включена доставка» и «Скидка на доставку не распространяется»
- [x] Флаг off: строка доставки на месте (регрессия mixed/plates)
- [x] Тесты красные, пока генератор не принимает флаг

**Verification:**
- [x] `pytest tests/test_embed_delivery_in_unit_price.py tests/test_commercial_export_mixed.py::test_export_service_mixed_xlsx_includes_pb_only_delivery -q`

**Dependencies:** Task 2

**Files likely touched:**
- `tests/test_embed_delivery_in_unit_price.py` (xlsx-кейсы в том же модуле или `tests/test_commercial_xlsx_embed_delivery.py` если файл раздуется — не больше одного нового test-модуля)

**Estimated scope:** S

## Task 4: GREEN — флаг в генераторе

**Description:** `generate_commercial_offer_xlsx`: после `calculate_total_cost` при флаге вызвать helper; подставить Цена/Сумма; `kp_delivery_export_lines` фильтровать по невшитым котлам; в цикле формул «Сумма» не затирать число, если embed. Пояснение после НДС. Default флага `False` — байт-в-байт текущие тесты.

**Acceptance criteria:**
- [x] Тесты Task 3 зелёные
- [x] `test_commercial_export_mixed.py` и `test_commercial_logistics_cost.py` зелёные без правок ассертов (кроме случая, если тесты звали генератор позиционно — тогда только `**kwargs`)
- [x] Визард/FileGenerationService не передают флаг → поведение как сейчас

**Verification:**
- [x] `pytest tests/test_embed_delivery_in_unit_price.py tests/test_commercial_export_mixed.py tests/test_commercial_logistics_cost.py -q`

**Dependencies:** Task 3

**Files likely touched:**
- `core/commercial_offer_xlsx.py`
- `tests/test_embed_delivery_in_unit_price.py`

**Estimated scope:** M

### Checkpoint: Core artifact

- [x] Вшитый buffer.XLSX сходится по итогу и НДС; обычный XLSX со строкой доставки
- [x] `FileGenerationService.generate_offer_xlsx` без новых аргументов в вызовах
- [x] Review: цикл `sum_cell.value = =qty*price` не ломает обычный режим

---

### Phase 3: Download path (вертикальный срез «можно скачать»)

## Task 5: Archive kind `xlsx_delivery_in_unit`

**Description:** Расширить `ArchiveFileKind`. В `generate_document` ветка рядом с `xlsx`: тот же `generate_commercial_offer_xlsx` с флагом True, имя `КП_{kp_id}_с_доставкой_в_цене.xlsx`. HTTP `GET /api/v1/commercial/archive/{id}/files/xlsx_delivery_in_unit` — 200, spreadsheet content-type, Content-Disposition с суффиксом. Media type: текущий `else` уже spreadsheet (`kind in {pdf, schema}` → pdf). 404/пустое КП — как у xlsx. Auth не менять.

**Acceptance criteria:**
- [x] Скачивание kind нового возвращает xlsx; в файле нет строки доставки при ненулевой доставке плит
- [x] `files/xlsx` и `files/pdf` без регрессии
- [x] Неизвестный kind по-прежнему validation (существующая ветка else)
- [x] КП в БД после generate не меняет `unit_price` / `logistics_cost` / `discount_percent`

**Verification:**
- [x] `pytest tests/test_archive_endpoints.py tests/test_archive_pile.py tests/test_archive_service.py -q -k "xlsx or pdf or files or generate_document"`

**Dependencies:** Task 4

**Files likely touched:**
- `app/schemas/archive.py`
- `app/services/archive_service.py`
- `tests/test_archive_endpoints.py`
- при необходимости `tests/test_archive_service.py` (не оба test-файла плюс три прод-файла — если endpoints-тест достаточно, service-тест не трогать)

**Estimated scope:** M

`app/api/v1/endpoints/archive.py` менять только если media_type не покрывает новый kind (ожидание: не нужно). Если понадобится — вытеснить `test_archive_service.py` из списка, лимит 5 файлов.

### Checkpoint: Backend complete

- [x] `GET .../files/xlsx_delivery_in_unit` на сохранённом КП с доставкой отдаёт «вшитый» файл
- [x] Повторный `files/xlsx` — со строкой доставки
- [x] Не приступать к UI, пока имя файла и kind не совпадают со спекой (`xlsx_delivery_in_unit`, суффикс `_с_доставкой_в_цене`)

---

### Phase 4: UI

## Task 6: Кнопка в `OfferDetailsDrawer`

**Description:** Тип `ArchiveFileKind` + fallback имени в `archiveApi`. Кнопка «XLSX (доставка в цене)» сразу справа от «XLSX». `downloadFile(buildDocumentUrl(id, "xlsx_delivery_in_unit"))`. `disabled` + title, если `delivery_service_total_rub === 0`. Карточка «Итого» и таблица состава не пересчитывают цены.

**Acceptance criteria:**
- [x] Кнопка видна рядом с PDF/XLSX
- [x] При `delivery_service_total_rub === 0` — disabled; обычный XLSX enabled
- [x] При доставке > 0 — enabled, URL содержит `xlsx_delivery_in_unit`
- [x] `buildDocumentUrl` / `downloadDocument` fallback: `КП_{id}_с_доставкой_в_цене.xlsx`

**Verification:**
- [x] `cd frontend && npm run test -- --run src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx src/features/commercial-archive/api/archiveApi.test.ts`
- [x] `cd frontend && npm run typecheck`

**Dependencies:** Task 5 (для e2e); контракт kind достаточно для тестов кнопки — можно писать параллельно с Task 5

**Files likely touched:**
- `frontend/src/features/commercial-archive/types/archive.ts`
- `frontend/src/features/commercial-archive/api/archiveApi.ts`
- `frontend/src/features/commercial-archive/api/archiveApi.test.ts`
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`

**Estimated scope:** M (ровно 5 файлов — не добавлять хуки)

### Checkpoint: Complete

- [x] Все Commands из спеки зелёные (pytest focused + vitest archive + typecheck; `npm run build` / `./run+logs.sh` не гонялись)
- [ ] Ручной КП с доставкой: два Excel, разный НДС, один итог; карточка со строкой доставки
- [ ] Ручной самовывоз: кнопка серая
- [ ] Mixed КП: надбавки не смешаны (проверить две марки разных типов в файле)

---

### Phase 5: Polish (только если Checkpoint 2–3 оставили дыры)

## Task 7: Mixed archive + невшиваемый котёл + sweep

**Description:** Если Task 3–5 уже покрыли mixed и fallback пустого котла — задача пустая, закрыть чеклистом «уже сделано». Иначе: archive generate для mixed КП; кейс plate_delivery > 0 без plate-строк (оставить строку доставки плит); прогон регрессии без `-k`.

**Acceptance criteria:**
- [x] Mixed archive вшитый XLSX: нет «Доставка плит»/«Доставка свай», если оба котла вшились
- [x] Невшитый котёл → его строка доставки на месте, второй котёл вшит
- [x] Полный focused pytest/vitest из Verification зелёный

**Verification:**
- [x] `pytest tests/test_embed_delivery_in_unit_price.py tests/test_commercial_export_mixed.py tests/test_commercial_logistics_cost.py tests/test_archive_endpoints.py tests/test_archive_pile.py tests/test_archive_service.py -q`
- [x] `cd frontend && npm run test -- --run src/features/commercial-archive && npm run typecheck`

**Dependencies:** Task 5, Task 6

**Files likely touched:**
- `tests/test_archive_pile.py` или `tests/test_embed_delivery_in_unit_price.py` (только если дыра)
- точечные правки генератора — только по красным тестам

**Estimated scope:** S

---

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Формула `qty*Цена` разъезжается с копейками | High | A4: во вшитом режиме «Сумма» — число; тест Σ vs `total_with_vat` |
| `line_product_type` ≠ котлы `calculate_total_cost` | High | Тот же `line_product_type`; mixed-тест двух котлов в Task 1 и 3 |
| pandas `read_excel` не считает формулы НДС | Med | Сверять формулу ячейки и/или openpyxl; не требовать Excel-приложение |
| Случайно вшить доставку в визарде | High | Флаг только из archive kind; не прокидывать в `FileGenerationService` |
| Task 6 раздувается >5 файлов | Med | Хуки не трогать; `useArchiveQueries` не расширять, drawer качает по URL как PDF/XLSX |
| Округление скидки helper vs `item_cost` без построчного round | Med | Task 2: Σ копеек котла = round(изделия+доставка) с добивкой extra; Task 3 ловит расхождение с `total_with_vat` |

## Parallelization

| Можно параллельно | После |
|-------------------|--------|
| Task 6 UI-тесты (мок URL) ∥ Task 3–5 | Контракт kind зафиксирован в спеке (уже) |
| Task 7 пустая, если 3–5 закрыли mixed | Checkpoint 3 |

Один агент: строго 1→2→3→4→5→6→7.

## Open Questions

Нет. Продуктовые развилки закрыты в спеке 2026-09-10.

## DoD (не начинать IMPLEMENT без)

- [x] Каждая задача: acceptance + verify + ≤5 files + S/M
- [x] Чекпоинты после foundation / artifact / backend / UI
- [x] Граф зависимостей и риски
- [ ] Человек подтвердил план (этот файл)

Первая задача реализации: **Task 1 RED**.
