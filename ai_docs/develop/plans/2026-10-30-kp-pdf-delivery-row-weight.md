# Implementation Plan: Вес груза в строке доставки PDF

**Created:** 2026-10-30  
**Status:** IMPLEMENT  
**Spec:** [`ai_docs/specs/kp-pdf-delivery-row-weight.md`](../../specs/kp-pdf-delivery-row-weight.md)

## Overview

В PDF КП по плитам ячейка «Вес(кг)» строки «Услуга по доставке грузов» сейчас всегда `0,00`. Подвал уже считает килограммы через `document_cargo_weight_kg` и пишет их с тремя знаками. В ячейку нужно то же число, в формате плитных ячеек: `_format_money_ru` (два знака). Рейсы, цена, сумма, XLSX и другие раскладки не меняются.

## Architecture Decisions

- **A1. Одна формула.** Килограммы — тот же вызов `document_cargo_weight_kg(order_data, pile_catalog_db_path=..., weight_catalog_db_path=...)`, что уже стоит у подвала. Второй расчёт не вводим. Вызов один на документ, подвал берёт то же значение.
- **A2. Только ветка плит.** Меняется `else` в цикле `delivery_export` внутри `generate_commercial_offer_pdf`. Ветки unified, свай, моста, составных свай, ФБС, маршей и ступеней не пишут «Вес(кг)» и не трогаются.
- **A3. Формат ячейки не равен формату подвала.** Ячейка: `_format_money_ru` (`228 480,00`). Подвал: по-прежнему `:,.3f` (`228 480,000`).
- **A4. Деньги строки не из веса.** `trips`, `unit_price` и `amount` остаются из `kp_delivery_export_lines`. В ячейку не попадает 18 600 и не остаток рейса.

## Risks

| Риск | Что делаем |
|------|------------|
| Два знака случайно попадут в подвал | Подвал не переводим на `_format_money_ru`; тест проверяет фразу «общим весом» с тремя знаками |
| Тот же вес на строке плиты и в доставке | Для плит-only это ожидаемо: одна масса. Проверка смотрит строку доставки, не весь текст PDF |
| Нулевая доставка | Строка по-прежнему не добавляется, если `kp_delivery_export_lines` пуст |

## Implementation order

| Phase | Focus | Depends |
|-------|-------|---------|
| 1 | RED: PDF плит, вес строки доставки | — |
| 2 | GREEN: ячейка вместо `"0,00"` | 1 |

## Task List

### Phase 1: RED

## Task 1: Падающий тест веса строки доставки

**Description:** В `tests/test_commercial_offer_document_weight.py` заказ только из плит с ненулевой доставкой. Строка «Услуга по доставке грузов» содержит `document_cargo_weight_kg`, отформатированный как плиты (`_format_money_ru`), и не содержит константу `0,00`. Подвал содержит «общим весом» с тремя знаками.

**Acceptance criteria:**
- [x] Тест красный на текущем `"0,00"`

**Verification:**
- [x] `pytest tests/test_commercial_offer_document_weight.py -q` — новый тест падает

**Dependencies:** None

**Files likely touched:**
- `tests/test_commercial_offer_document_weight.py`

**Estimated scope:** S

### Phase 2: GREEN

## Task 2: Подставить вес груза в ячейку плит

**Description:** В `else` цикла доставки PDF заменить литерал `"0,00"` на `_format_money_ru(cargo_kg)`. `cargo_kg` посчитать один раз теми же путями каталогов, что уже передаются в `generate_commercial_offer_pdf`, и отдать это же число подвалу.

**Acceptance criteria:**
- [x] Новый тест зелёный
- [x] Существующие тесты подвала зелёные
- [x] XLSX и остальные ветки таблицы не изменены

**Verification:**
- [x] `pytest tests/test_commercial_offer_document_weight.py tests/test_kp_document_cargo_weight.py -q`

**Dependencies:** Task 1

**Files likely touched:**
- `core/commercial_offer.py`

**Estimated scope:** S
