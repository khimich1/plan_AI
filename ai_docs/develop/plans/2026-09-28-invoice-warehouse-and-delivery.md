# План: Склад и доставка в JSON счёта

**Дата:** 2026-09-28  
**Статус:** готов  
**Идея:** `ai_docs/ideas/invoice-warehouse-and-delivery.md`  
**Спека:** `ai_docs/specs/invoice-warehouse-and-delivery.md`

## Overview

Кнопка «Отправить в 1С» открывает пять складов. Выбранное имя сохраняется на КП и пишется в шапку JSON. В каждом файле есть `Самовывоз`. `Цена` и `ПроцентСкидки` всегда про товар до скидки. Доля доставки — `СуммаДоставки` на строке, ноль при самовывозе, и она прибавляется после скидки. `СуммаДокумента` и `СуммаНДС` сходятся с этими строками. Excel «доставка в цене» не меняется.

## Architecture Decisions

- **Сборщик не пишет файл.** `core/invoice_export.py` возвращает словарь. Файл пишет сервис архива после ворот, затем статус и имя склада.
- **Скидка не входит в доставку.** `Цена` остаётся ценой карточки до скидки. `ПроцентСкидки` остаётся в шапке и на строке. `СуммаДоставки` — копейки котла на всю строку, не цена штуки.
- **Четыре котла не смешиваются.** Плиты, сваи с мостовыми, ФБС со ступенями и маршами, каждая длина длинных свай. Делитель копеек тот же, что у плит и свай. Excel по-прежнему зовёт старую функцию и не получает котлы ФБС и длинных свай.
- **НДС файла — 22/122 по сумме строки.** Сумма строки = цена после скидки × количество + `СуммаДоставки`. Карточный `vat_amount` в файл не копируется.
- **Неготовый котёл — не самовывоз.** Есть марки без веса или без тарифа — файл не пишется.
- **Склад не в хеше снимка.** Исправление повторяет сохранённое имя. Пустая колонка у уже отправленного счёта один раз принимает имя в теле `update`.

## Dependency graph

```
копейки доставки по строке (4 котла)
    │
    └── сборщик JSON: Склад, Самовывоз, СуммаДоставки, НДС
            │
            ├── колонка invoice_warehouse
            │       │
            │       └── POST invoice-export / invoice-correction
            │               │
            │               └── список складов в кнопке
            │
            └── карточка отдаёт сохранённый склад
```

## Task List

### Phase 1 — Копейки и документ

- [x] **WD-001: Доля доставки по строке**
  - **Description:** Чистая функция отдаёт копейки доставки на каждую строку заказа: плиты, сваи и мостовые сваи, ФБС/ступени/марши, отдельно каждая длина длинных свай. Скидка в функцию не входит. Существующий Excel продолжает звать `embed_delivery_in_unit_prices` без этих котлов.
  - **Acceptance criteria:**
    - [x] Сумма копеек котла равна сумме по его строкам
    - [x] ФБС не получает плитные копейки, другая длина сваи не получает чужие
    - [x] Остаток копейки добивается на первые штуки, как у плит
    - [x] Тесты Excel «доставка в цене» не меняют ожидания
  - **Verification:**
    - [x] `pytest tests/test_embed_delivery_in_unit_price.py tests/test_commercial_offer_xlsx.py -q -k "embed or delivery_in_unit"`
  - **Dependencies:** нет
  - **Files likely touched:** `core/embed_delivery_in_unit_price.py`, `tests/test_embed_delivery_in_unit_price.py`
  - **Estimated scope:** M

- [x] **WD-002: Поля счёта**
  - **Description:** Сборщик пишет `Склад`, `Самовывоз`, на каждой строке `Цена`, `ПроцентСкидки` и `СуммаДоставки`. Сумма строки = цена после скидки × количество + доля доставки. `СуммаДокумента` — сумма строк. `СуммаНДС` — построчный 22/122. При нулевой доставке `Самовывоз` истина и `СуммаДоставки` равна 0.
  - **Acceptance criteria:**
    - [x] Скидка не умножает `СуммаДоставки`
    - [x] Отдельной строки «Доставка» в `Товары` нет
    - [x] `СуммаДокумента` равна сумме строк, `СуммаНДС` не копируется с карточки
  - **Verification:**
    - [x] `pytest tests/test_archive_invoice_export.py -q -k "invoice or pickup or delivery_share"`
  - **Dependencies:** WD-001
  - **Files likely touched:** `core/invoice_export.py`, `tests/test_archive_invoice_export.py`
  - **Estimated scope:** M

### Checkpoint: Phase 1

- [x] Сборщик и делитель копеек зелёные без схемы, без файла на диске и без HTTP

### Phase 2 — Запись файла

- [x] **WD-003: Колонка склада**
  - **Description:** `KP_offers.invoice_warehouse`. Повтор схемы не падает. Карточка архива отдаёт сохранённое имя или пусто.
  - **Acceptance criteria:**
    - [x] Пустая колонка читается как отсутствие склада
    - [x] Повторный запуск схемы не ломает существующие КП
  - **Verification:**
    - [x] `pytest tests/test_archive_endpoints.py tests/test_archive_service.py -q -k "schema or invoice_warehouse"`
  - **Dependencies:** нет
  - **Files likely touched:** `core/kp_db_schema.py`, `core/kp/offers_read.py`, `app/schemas/archive.py`, `frontend/src/features/commercial-archive/types/archive.ts`
  - **Estimated scope:** M

- [x] **WD-004: Отправка и исправление**
  - **Description:** `POST invoice-export` принимает `{ "warehouse": "…" }` из пяти строк. Чужое или пустое имя — отказ, файла нет. Неготовый котёл — отказ, это не самовывоз. После записи файла имя сохраняется и статус становится «на согласовании». Исправление без тела повторяет сохранённый склад и заново считает доставку. Пустая колонка на исправлении принимает то же тело один раз.
  - **Acceptance criteria:**
    - [x] Успешный create пишет `Склад` и `Самовывоз` в `invoice_create_*.json` и сохраняет имя
    - [x] Неготовый котёл, чужой склад и пустое тело create не меняют статус
    - [x] Correction со складом в базе не требует тела и пишет то же имя
    - [x] Correction с пустой колонкой без тела — отказ; с телом — пишет `update` и сохраняет имя
  - **Verification:**
    - [x] `pytest tests/test_archive_invoice_export.py tests/test_archive_endpoints.py -q -k "invoice_export or invoice_correction or warehouse"`
  - **Dependencies:** WD-002, WD-003
  - **Files likely touched:** `app/services/archive_service.py`, `app/api/v1/endpoints/archive.py`, `app/schemas/archive.py`, `tests/test_archive_invoice_export.py`, `tests/test_archive_endpoints.py`
  - **Estimated scope:** M

### Checkpoint: Phase 2

- [x] Файл на диске содержит склад, самовывоз и долю доставки. Статус меняется только после записи

### Phase 3 — Кнопка

- [x] **WD-005: Список складов**
  - **Description:** Фиолетовая «Отправить в 1С» открывает пять названий и «Отмена». Серая кнопка список не открывает. «Отмена» не вызывает POST. Выбор названия отправляет это имя. После успеха карточка показывает склад текстом. «Отправить исправление» открывает список только если склад ещё не сохранён.
  - **Acceptance criteria:**
    - [x] «Отмена» не шлёт запрос
    - [x] Выбор названия шлёт `warehouse` этой строкой
    - [x] Сохранённый склад виден на карточке и исправление уходит без списка
  - **Verification:**
    - [x] `cd frontend && npm run test -- src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
    - [x] `cd frontend && npm run typecheck`
  - **Dependencies:** WD-004
  - **Files likely touched:** `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`, `frontend/src/features/commercial-archive/api/archiveApi.ts`, `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
  - **Estimated scope:** M

### Checkpoint: готово

- [x] Пять критериев успеха спеки закрыты тестами фазы 1–3
- [x] Excel «доставка в цене» без новых ожиданий

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Долю доставки положат внутрь `Цены`, и скидка срежет её | Высокий | `СуммаДоставки` отдельным полем. Тест: скидка 10 % не меняет эту сумму |
| Неготовый котёл уйдёт как самовывоз | Высокий | Отказ до записи файла, если есть марки без веса или без тарифа |
| Цена штуки с доставкой потеряет копейку | Средний | В JSON сумма строки, не вторая цена штуки |
| Правка общего делителя сломает Excel | Средний | Новая функция. Старый `embed_delivery_in_unit_prices` и его тесты не трогать по ожиданиям |

## Open Questions

Блокирующих нет. Ключи `Склад`, `Самовывоз` и `СуммаДоставки` можно переименовать по просьбе 1С, не меняя расчёт.
