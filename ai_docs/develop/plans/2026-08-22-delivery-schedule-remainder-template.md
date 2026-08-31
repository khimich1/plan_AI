# Implementation Plan: График поставки — остаток в XLSX-шаблоне

> **Spec:** [`ai_docs/specs/delivery-schedule-remainder-template.md`](../../specs/delivery-schedule-remainder-template.md)  
> **Идея:** [`ai_docs/ideas/grafik-postavki-ostatok-v-shablone.md`](../../ideas/grafik-postavki-ostatok-v-shablone.md)  
> **Родитель:** [`docs/specs/delivery-schedule.md`](../../../docs/specs/delivery-schedule.md)  
> **Дата:** 2026-08-22  
> **Статус:** IMPLEMENT ✅  
> **Не коммитить** без явной просьбы

---

## Overview

Менеджер качает шаблон из открытого редактора и видит две полосы: сверху текущий черновик партий, ниже только невошедший остаток. Файл собирается **на клиенте** (exceljs) из `plates` + `batches` модалки — без `GET /template`. Тот же двухполосный макет отдаёт серверный `GET /template` от **сохранённого** графика (или только «Остаток» = все плиты КП, если графика нет). Импорт по-прежнему заменяет черновик; парсер пропускает строки-заголовки полос.

Не трогаем: светофор, PDF/XLSX-документ клиенту, схему БД, контракты PUT/import, сортировку колонки «Позиции КП», жёлтую плашку при урезании импорта.

---

## Architecture Decisions

Зафиксированы в спеке (D1–D10). Не переоткрывать:

- Шаблон из открытого редактора — **клиент**, не `POST /template`.
- `GET /template` — тот же макет от сохранённого графика (или remainder-only).
- Заголовки полос точно: `Уже в поставках`, `Остаток`. Парсер скипает, если любая из A–F после `strip` совпадает.
- Даты в файле — строки `ДД.ММ.ГГГГ`.
- Полосу и заголовок пишем только если в ней есть строки.
- `exceljs` (не SheetJS). Импорт заменяет черновик. Плашки нет.

Клиентская чистая функция рядов — источник истины для макета; серверный `build_template` повторяет те же правила для GET.

```
plates + batches (черновик модалки)
        │
        ▼
buildScheduleTemplateRows()     ← чистые ряды (vitest)
        │
        ▼
buildScheduleTemplateXlsx()     ← exceljs blob
        │
        ▼
saveBlobAs(delivery_schedule_template.xlsx)

GET /template
        │
        ▼
service: plates + saved batches (или [])
        │
        ▼
core.build_template(path, plates=, batches=)
        │
        ▼
parse_template: skip section titles → существующий матчинг
```

---

## Components

| Компонент | Роль | Зависит от |
|-----------|------|------------|
| `scheduleTemplateRows.ts` | Чистые ряды: полосы, остаток, даты ДД.ММ.ГГГГ | `OfferPlateForSchedule`, `BatchDraft` |
| `buildScheduleTemplateXlsx.ts` | exceljs: лист, merge A–F, фон разделителя | ряды |
| `ScheduleDocumentButtons` | «Скачать шаблон» из редактора = клиент; без GET /template | xlsx-builder, `saveBlobAs` |
| `DeliveryScheduleEditor` | Передаёт `plates` + `batches` в кнопки | — |
| `core/delivery_schedule_xlsx.py` | `build_template(..., batches=)`; `parse_template` skip titles | openpyxl (уже есть) |
| `delivery_schedule_service.build_template_bytes` | Передаёт сохранённые партии в `build_template` | core xlsx |
| `exceljs` в `frontend/package.json` | Клиентская сборка XLSX | — |

Вне скоупа: `delivery_schedule_pdf.py`, документ XLSX (`build_document`), PUT/import API, светофор, колонка позиций.

---

## Task List

### Phase 1: Client row builder (TDD)

### Task 1: Чистые ряды шаблона

**Description:** Функция `buildScheduleTemplateRows(plates, batches)` и константы заголовков полос. Без I/O и без exceljs.

**Acceptance criteria:**
- [x] Пустой черновик + qty=40 → только секция «Остаток» и строка с 40, пустые партия/даты
- [x] Черновик 10 из 40 → сверху секция с 10 и датами ДД.ММ.ГГГГ, снизу остаток 30
- [x] Две партии одной марки (10+15) → две верхние строки, остаток 15
- [x] Полностью разбито → только верхняя полоса
- [x] Позиция без марки не попадает никуда; item с `qty < 1` не вверху и не в сумме остатка

**Verification:**
- [x] `cd frontend && npm run test -- src/features/delivery-schedule/lib/scheduleTemplateRows.test.ts`

**Dependencies:** None

**Files likely touched:**
- `frontend/src/features/delivery-schedule/lib/scheduleTemplateRows.ts`
- `frontend/src/features/delivery-schedule/lib/scheduleTemplateRows.test.ts`

**Estimated scope:** S

### Checkpoint: Rows
- [x] Vitest рядов зелёный

### Phase 2: Client XLSX + editor wiring

### Task 2: exceljs + сборка файла

**Description:** Добавить зависимость `exceljs` (exact, как в package.json). Собрать workbook: лист «График поставки», заголовки колонок, merge A–F + фон на секциях, даты строками.

**Acceptance criteria:**
- [x] `npm install exceljs` в `frontend/`, зависимость в package.json
- [x] Smoke: ненулевой blob, лист называется «График поставки»
- [x] Разделитель — отдельная merged-строка с заголовком полосы

**Verification:**
- [x] `cd frontend && npm run test -- src/features/delivery-schedule/lib/buildScheduleTemplateXlsx.test.ts`

**Dependencies:** Task 1

**Files likely touched:**
- `frontend/package.json`, `frontend/package-lock.json`
- `frontend/src/features/delivery-schedule/lib/buildScheduleTemplateXlsx.ts`
- `frontend/src/features/delivery-schedule/lib/buildScheduleTemplateXlsx.test.ts`
- `frontend/vite.config.ts` — только если exceljs требует browser alias

**Estimated scope:** S

### Task 3: Кнопка «Скачать шаблон» из черновика

**Description:** В открытом редакторе кнопка собирает файл на клиенте из текущих `plates`+`batches` и качает `delivery_schedule_template.xlsx`. `useDownloadDeliveryScheduleTemplateMutation` / `GET /template` не вызывается.

**Acceptance criteria:**
- [x] Editor передаёт `plates` и `batches` в `ScheduleDocumentButtons`
- [x] Клик не вызывает template mutation
- [x] XLSX/PDF документа по-прежнему через GET /document

**Verification:**
- [x] `cd frontend && npm run test -- src/features/delivery-schedule`
- [x] `cd frontend && npm run typecheck`

**Dependencies:** Task 2

**Files likely touched:**
- `frontend/src/features/delivery-schedule/components/ScheduleDocumentButtons.tsx`
- `frontend/src/features/delivery-schedule/components/DeliveryScheduleEditor.tsx`
- `frontend/src/features/delivery-schedule/hooks/useDeliveryScheduleQueries.ts` (комментарий: mutation не для кнопки редактора)
- опционально тест кнопки

**Estimated scope:** S

### Checkpoint: Frontend
- [x] Тесты delivery-schedule зелёные
- [x] typecheck чистый

### Phase 3: Backend two-band + parser

### Task 4: Парсер пропускает заголовки полос (TDD)

**Description:** `parse_template` скипает строку, если любая из A–F после trim равна `Уже в поставках` или `Остаток`. Старые файлы без полос не ломаются. Заголовки не попадают в `unmatched_rows`.

**Acceptance criteria:**
- [x] Двухполосный файл: партии только из верхней полосы; пустые строки остатка не создают партии
- [x] Заголовки полос не в unmatched
- [x] Файл без полос (только марка+qty / заполненные партии) парсится как сейчас

**Verification:**
- [x] `pytest tests/test_delivery_schedule_xlsx.py -q`

**Dependencies:** None (параллельно Phase 1–2 по смыслу; делаем после фронта в одном прогоне)

**Files likely touched:**
- `core/delivery_schedule_xlsx.py`
- `tests/test_delivery_schedule_xlsx.py`

**Estimated scope:** S

### Task 5: `build_template(..., batches=)` двухполосный

**Description:** Тот же макет, что на клиенте. Без `batches` / пустые партии — только «Остаток» (плиты КП). С партиями — верх + остаток. Обновить существующие тесты префилла: строка 2 теперь заголовок «Остаток».

**Acceptance criteria:**
- [x] `build_template(path, plates=)` пишет «Остаток» и марки
- [x] `build_template(..., batches=)` пишет обе полосы и остаток
- [x] Даты в файле — `ДД.ММ.ГГГГ`
- [x] Round-trip: построенный файл → parse → верхние партии

**Verification:**
- [x] `pytest tests/test_delivery_schedule_xlsx.py -q`

**Dependencies:** Task 4

**Files likely touched:**
- `core/delivery_schedule_xlsx.py`
- `tests/test_delivery_schedule_xlsx.py`

**Estimated scope:** S

### Task 6: GET /template от сохранённого графика

**Description:** `build_template_bytes` загружает плиты КП и, если график есть, сохранённые партии; передаёт их в `build_template`. Нет графика — `batches=[]` (remainder-only). Контракт эндпоинта не меняется.

**Acceptance criteria:**
- [x] Нет графика: файл как remainder-only (марка+qty под «Остаток»)
- [x] Есть сохранённые партии: верх = они, низ = остаток
- [x] 404 на отсутствующее КП без изменений

**Verification:**
- [x] `pytest tests/test_delivery_schedule_xlsx.py tests/test_delivery_schedule_service.py -q`

**Dependencies:** Task 5

**Files likely touched:**
- `app/services/delivery_schedule_service.py`
- `tests/test_delivery_schedule_service.py`

**Estimated scope:** S

### Checkpoint: Complete
- [x] `pytest tests/test_delivery_schedule_xlsx.py tests/test_delivery_schedule_service.py -q`
- [x] `cd frontend && npm run test -- src/features/delivery-schedule`
- [x] `cd frontend && npm run typecheck`
- [x] Спека указывает на этот план; changelog Unreleased при привычке репо

---

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| exceljs тянет Node `fs`/`stream` в Vite | Med | Smoke-тест в vitest; при необходимости alias на `exceljs/dist/exceljs.min.js` |
| Существующие pytest ждут плиты со строки 2 | Med | Обновить префилл-ассерты под секцию «Остаток»; parse-тесты без полос оставить |
| Логика рядов разъедется клиент/сервер | Med | Одинаковые литералы и правила из спеки; зеркальные кейсы 40→10→30 |
| Партия с именем «Остаток» скипается | Low | Зарезервированные заголовки — осознанно в спеке |
| exceljs раздувает бандл | Low | Принято (D10); не SheetJS |

## Open Questions

Нет — решения залочены. Блокировать реализацию только при противоречии спеки коду, которое нельзя разрешить минимальным diff.
