# Implementation Plan: Доставка в JSON для 1С и в XLSX «доставка в цене»

**Created:** 2026-11-11
**Status:** IMPLEMENT. Тесты из спеки зелёные.
**Spec:** [`ai_docs/specs/1c-delivery-invoice-parity.md`](../../specs/1c-delivery-invoice-parity.md)
**Idea:** [`ai_docs/ideas/1c-delivery-invoice-parity.md`](../../ideas/1c-delivery-invoice-parity.md)

## Overview

При доставке JSON для 1С и XLSX «доставка в цене» считают одни и те же деньги. Прайс округляется до копеек, итог к оплате не меняется, ручная скидка документа равна базе минус этот итог. Плюс — скидка, минус — наценка. Сумма строки = база − скидка строки, доставка в сумму не входит. В JSON цена — округлённый прайс, в XLSX цена — сумма строки / количество, колонка «Сумма» — число, не формула.

Самовывоз, обычные PDF/XLSX, карточка и база не меняются. Новой кнопки нет: флаг `embed_delivery_in_unit_price` и сборка счёта уже есть.

Эталоны уже загружены в 1С: заказ 478 (скидка 17,33 %, итог 1 809 912,75) и заказ 481 (наценка −2,67 %, итог 2 247 756,15).

## Architecture Decisions

- **A1. Одна функция, без ввода-вывода.** `core/delivery_invoice_parity.py`, `delivery_invoice_lines`. Она не импортирует `invoice_export` и не считает доставку сама. Вызывающий передаёт цель — сегодняшние значения `_line_payable`.
- **A2. Знак не ветка.** Скидка документа = база − цель. Минус остаётся минусом в `ПроцентСкидки` и `СуммаСкидки`. `None` функция возвращает только когда база документа равна нулю.
- **A3. Цель не пересчитывается вторым способом.** `invoice_export` по-прежнему считает `_line_payable` и отдаёт эти числа в функцию. Итог документа равен их сумме.
- **A4. Самовывоз не заходит в функцию.** Все доли доставки нулевые — сегодняшние `_line_discount_amount` и коммерческий процент.
- **A5. XLSX берёт те же строки.** Доли доставки для цели — `allocate_line_delivery_kopecks` от тех котлов, которые обычный XLSX уже выводит строками. Кнопка больше не вызывает `embed_delivery_in_unit_prices`.
- **A6. Колонка «Сумма» во вшитом режиме остаётся числом.** Формулу `количество × цена` по-прежнему не писать: 9 188,58 × 200 = 1 837 716,00, а строка 1 837 716,92.
- **A7. НДС вшитого листа — число.** Сумма строчных `round(сумма × 22/122, 2)`, не `SUM × 0,22`. Порог `kp_id < 27` на этот файл не действует. Обычный XLSX порог сохраняет.
- **A8. Сноски «скидка на доставку не распространяется» нет.** Строк доставки во вшитом файле нет.
- **A9. База, хеш снимка, карточка, обычный PDF/XLSX не меняются.**

```
_line_payable (как сегодня)
        │
        └── delivery_invoice_lines
                │
                ├── build_invoice_document, только если сумма долей > 0
                │
                └── generate_commercial_offer_xlsx(embed_delivery_in_unit_price=True)
```

`ArchiveService` новые аргументы не получает: он уже передаёт `line_delivery_kopecks`.

## Implementation order

| Фаза | Что | Зависит |
|------|-----|---------|
| 1 | Тесты и функция денег | — |
| 2 | Подключение JSON | 1 |
| 3 | Подключение XLSX | 1 |
| 4 | Регрессия обоих путей | 2 и 3 |

Фазы 2 и 3 можно делать параллельно после фазы 1. Фаза 3 не ждёт фазу 2: оба вызывают одну функцию.

## Task List

### Phase 1: Foundation

## Task 1: RED — контракт `delivery_invoice_lines`

**Description:** Зафиксировать оба эталона и края тестами до кода функции. Возврат — список `ParityLine` или `None`, если база документа равна нулю.

**Acceptance criteria:**
- [x] Эталон 1 (скидка КП 25 %): процент 17,33, скидки 310 110,40 и 69 193,15, суммы 1 479 745,60 и 330 167,15, НДС 266 839,37 и 59 538,34, цены счёта 7 398,73 и 11 005,57
- [x] Эталон 2 (скидка КП 5 %): процент −2,67, скидки −47 860,92 и −10 678,93, суммы 1 837 716,92 и 410 039,23, НДС 331 391,58 и 73 941,50, цены счёта 9 188,58 и 13 667,97
- [x] Одна строка получает всю скидку документа, положительную и отрицательную
- [x] Строка с количеством 0 не входит в базу и не получает остаток
- [x] База документа 0 → `None`
- [x] Скидка документа 0 → цена счёта равна округлённому прайсу
- [x] Тесты красные: модуля ещё нет

**Verification:**
- [x] `pytest tests/test_delivery_invoice_parity.py -q` падает на импорте, не skip

**Dependencies:** None

**Files likely touched:**
- `tests/test_delivery_invoice_parity.py`

**Estimated scope:** S

## Task 2: GREEN — функция

**Description:** Реализовать `core/delivery_invoice_parity.py` по формуле спеки. `Decimal`, копейка, `ROUND_HALF_UP`. Процент один на все строки, со знаком. Остаток копеек — на последнюю строку с количеством больше нуля.

**Acceptance criteria:**
- [x] Все проверки Task 1 зелёные
- [x] Сумма `payable` равна сумме переданных целей
- [x] Сумма `discount_amount` равна базе минус эта сумма
- [x] Функция не читает файлы и не импортирует Excel или базу

**Verification:**
- [x] `pytest tests/test_delivery_invoice_parity.py -q`

**Dependencies:** Task 1

**Files likely touched:**
- `core/delivery_invoice_parity.py`
- `tests/test_delivery_invoice_parity.py`

**Estimated scope:** S

## Checkpoint: Foundation

- [x] `pytest tests/test_delivery_invoice_parity.py -q` зелёный
- [x] Существующие тесты счёта и XLSX не запускались с новой функцией и остаются зелёными

### Phase 2: JSON

## Task 3: Сборка счёта вызывает функцию при доставке

**Description:** В `invoice_export._goods_and_totals`, если сумма долей доставки больше нуля и функция вернула строки, подставить `ParityLine`. `СуммаДоставки` копировать из доли. Если функция вернула `None` или доставки нет — оставить сегодняшний расчёт. Хеш снимка не трогать: он считается из коммерческой скидки и прайса карточки.

**Acceptance criteria:**
- [x] Эталон 2, собранный через `build_invoice_document`, совпадает с `data/exchange_export/new_211_markup_5.json` по цене, проценту, суммам, НДС и доставке
- [x] Доставка и скидка 10 % из `test_invoice_delivery_share_is_not_discounted`: процент уже не 10, сумма строки равна прежней цели, доставка в сумму не прибавлена
- [x] Скидка 0 % и доставка дают минус в `ПроцентСкидки` и `СуммаСкидки`
- [x] Самовывоз (`test_invoice_header_totals_follow_lines_not_card_vat` и склад без доставки) не меняется: процент 13 и 10, цена без нового округления
- [x] Хеш двух документов с разным складом и одной доставкой по-прежнему совпадает

**Verification:**
- [x] `pytest tests/test_archive_invoice_export.py tests/test_delivery_invoice_parity.py -q`

**Dependencies:** Task 2

**Files likely touched:**
- `core/invoice_export.py`
- `tests/test_archive_invoice_export.py`
- `tests/test_delivery_invoice_parity.py`

**Estimated scope:** M

### Phase 3: XLSX

## Task 4: Флаг «доставка в цене» берёт ту же функцию

**Description:** При `embed_delivery_in_unit_price=True` посчитать доли через `allocate_line_delivery_kopecks` из котлов, которые обычный файл уже показывает строками. Цель строки — тот же `_line_payable`. В таблицу записать округлённую цену счёта и число суммы. Строки доставки не добавлять. Сноску не писать. НДС — число, сумма строчных 22/122. Без флага генератор не меняется.

**Acceptance criteria:**
- [x] С флагом нет строк «Доставка» и нет текста «Скидка на доставку не распространяется»
- [x] Сумма колонки «Сумма» равна сегодняшнему итогу к оплате
- [x] «Сумма» — число, не формула `количество × цена`
- [x] НДС равен сумме строчных 22/122, в том числе для `kp_id` меньше 27
- [x] На числах эталона 2 цены листа 9 188,58 и 13 667,97, суммы 1 837 716,92 и 410 039,23
- [x] Без флага строка доставки и прежний НДС на месте
- [x] `embed_delivery_in_unit_prices` из генератора больше не вызывается

**Verification:**
- [x] `pytest tests/test_embed_delivery_in_unit_price.py tests/test_commercial_offer_xlsx.py tests/test_commercial_offer_document_weight.py -q`

**Dependencies:** Task 2

**Files likely touched:**
- `core/commercial_offer_xlsx.py`
- `tests/test_embed_delivery_in_unit_price.py`
- `tests/test_commercial_offer_xlsx.py`

**Estimated scope:** M

## Checkpoint: оба выхода

- [x] JSON и XLSX на эталоне 2 дают один итог 2 247 756,15 и одни суммы строк
- [x] Самовывоз и обычный XLSX не изменились

### Phase 4: Regression

## Task 5: Прогон команды из спеки

**Description:** Прогнать список тестов спеки и поправить упавшие ожидания, которые описывали старую надбавку на штуку, сноску или НДС `× 0,22` во вшитом файле. Поведение вне флага и самовывоза не подгонять под новую формулу.

**Acceptance criteria:**
- [x] Команда из спеки зелёная
- [x] Кнопка архива по-прежнему выключена при доставке 0 и скачивает `xlsx_delivery_in_unit`
- [x] Карточка и `discount_percent` в базе не пишутся этим изменением

**Verification:**
- [x] `pytest tests/test_delivery_invoice_parity.py tests/test_archive_invoice_export.py tests/test_embed_delivery_in_unit_price.py tests/test_commercial_offer_xlsx.py tests/test_archive_endpoints.py -q`

**Dependencies:** Task 3, Task 4

**Files likely touched:**
- `tests/test_archive_endpoints.py` только если упало имя файла или флаг
- тесты из Task 3 и Task 4

**Estimated scope:** S

## Checkpoint: Complete

- [x] Семь критериев успеха спеки закрыты тестами
- [x] `new_211_markup_5.json` не переписан
- [x] Реализация готова к ревью

## Risks and Mitigations

| Риск | Влияние | Что делать |
|------|---------|------------|
| XLSX передаст в функцию не все котлы, которые JSON уже раскладывает (ФБС, длинные сваи) | Итог файла разойдётся со счётом | Брать доли из `allocate_line_delivery_kopecks` по тем суммам, которые обычный XLSX уже печатает строками доставки. Сумма строк вшитого файла равна прежнему итогу |
| Старые тесты вшитого XLSX ждут надбавку на штуку, сноску и `SUM × 0,22` | Ложные падения | Переписать ожидания флага. Тесты без флага не менять |
| `test_invoice_delivery_share_is_not_discounted` ждёт процент 10 при доставке | Падение на новом проценте | Оставить проверку итога и поля `СуммаДоставки`. Процент и `СуммаСкидки` сверить с функцией |
| Ограничение ручной наценки в соглашении 1С | Загрузка в проде может срезать минус | В план кода не входит. Это настройка 1С, отмечена в спеке |
| Остаток копеек на числе строк, отличном от двух | Расхождение на 1–2 копейки | Правило спеки не менять. Отдельной подстройки под 1С нет |

## Open Questions

Нет.
