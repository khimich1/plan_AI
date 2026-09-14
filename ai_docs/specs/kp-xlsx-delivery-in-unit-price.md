# Spec: КП — XLSX с доставкой в цене изделия

> **Источник идеи:** [`ai_docs/ideas/kp-xlsx-delivery-in-unit-price.md`](../ideas/kp-xlsx-delivery-in-unit-price.md)  
> **План и задачи:** [`ai_docs/develop/plans/2026-09-10-kp-xlsx-delivery-in-unit-price.md`](../develop/plans/2026-09-10-kp-xlsx-delivery-in-unit-price.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS ✅ → IMPLEMENT  
> **Статус:** ready for implementation  
> **Дата:** 2026-09-10  
> **Связанные модули:** `OfferDetailsDrawer`, `ArchiveFileKind`, `ArchiveService.generate_document`, `core/commercial_offer_xlsx.py`, `core/commercial_pricing.calculate_total_cost`

---

## Assumptions I'm Making

1. **Только архив.** Визард и `GET /api/v1/offers/{id}/xlsx` не получают этот режим.
2. **Экспорт, не мутация.** `unit_price`, `discount_percent`, `logistics_cost`, `pile_logistics_cost` в БД не меняются. «В производство» и обычные PDF/XLSX как сейчас.
3. **Тариф рейса уже с НДС.** 156 800 ₽ на скрине — сумма к оплате за доставку, не «нетто + налог сверху». Итог к оплате во «вшитом» файле **равен** `totals["total_with_vat"]` обычного КП.
4. **Обычный документ НДС с доставки не показывает.** `vat_amount = plates_after_discount × 0.22`, доставка плюсуется к итогу. Этот spec **не** чинит обычный PDF/XLSX.
5. **Во «вшитом» XLSX** НДС = сумма товарных строк × 0.22 (та же конвенция `VAT_RATE`, не 22/122). Строк доставки нет → налог берётся и с надбавки. Строка НДС будет **больше**, чем в карточке архива; «Сумма без НДС» (если появится в шапке/итогах через `total − vat`) — меньше.
6. **Скидка только на изделие:** `discounted = unit_price × (1 − d/100)`, затем надбавка доставки. Не `(unit_price + delivery) × (1 − d)`.
7. **«Равномерно» = на штуку**, не на вес и не на сумму строки. Два котла: `plate_delivery_total` → строки `plates` (legacy без `product_type` = plates); `pile_delivery_total` → `piles` + `bridge_piles`. ФБС / ступени / марши надбавки не получают (у них сейчас нет своей доставки в `calculate_total_cost`).
8. **Копейки.** Все доли в целых копейках. Σ(сумма строки) по товарным позициям = `plates_after_discount + delivery_embedded`. Допускается, что у последней затронутой строки цена за штуку отличается от «ровной» надбавки на доли копейки — лишь бы сумма строк сошлась.
9. **`pile_delivery_ready is False`:** свайная доставка уже 0 в totals — вшивать нечего. Плитную вшиваем, если она > 0.
10. **Доставка 0:** кнопка disabled (title: нет суммы доставки). Обычный XLSX доступен.
11. **Auth** как у текущего `GET .../files/{kind}`: `admin`, `manager`.
12. **Без новых pip/npm зависимостей.** Без колонок в SQLite.
13. **Web-only.** Бот вне scope.

---

## Decisions locked (ideation 2026-09-10)

| # | Тема | Решение |
|---|------|---------|
| D1 | Кто | Менеджер (и админ) в карточке архива |
| D2 | Формат MVP | Только XLSX. PDF с вшитой доставкой — не делаем |
| D3 | Рядом с чем | Вторая кнопка, обычные PDF/XLSX не заменяем |
| D4 | Базис надбавки | На штуку после скидки |
| D5 | НДС во вшитом файле | Со всей суммы строк (доставка уже с НДС) |
| D6 | НДС обычного КП | Не меняем |
| D7 | Скидка на доставку | Нет |
| D8 | Mixed | Два котла: плиты / сваи+мостовые |
| D9 | Хранение | КП в БД не трогаем |
| D10 | Пояснение в файле | Одна строка под итогами про включённую доставку |
| D11 | Доставка 0 | Кнопка неактивна |
| D12 | Имя файла | `КП_{kp_id}_с_доставкой_в_цене.xlsx` |

---

## Objective

Менеджер из архива скачивает Excel для заказчика, которому нельзя показывать отдельную услугу доставки: цены изделий уже «с доставкой на объект», итог к оплате тот же, НДС во вшитом файле считается и с доставки.

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | открыл карточку КП в архиве | вторую кнопку Excel | отдать заказчику файл без строки доставки |
| US-2 | заказчик просит доставку отдельной строкой | оставить обычный XLSX/PDF | не ломать текущий документ |
| US-3 | на КП скидка | чтобы доставка не скидывалась | не подарить логистику |
| US-4 | в КП плиты и сваи | чтобы фуры не смешались в одной надбавке | сваи не оплачивали доставку плит |
| US-5 | самовывоз / доставка 0 | не получить «другой» файл ниоткуда | не путать два одинаковых Excel |

### Reframed success criteria

| Требование | Критерий |
|------------|----------|
| Итог | Σ сумм товарных строк вшитого XLSX = `total_with_vat` обычного КП (± 0,01 ₽) |
| Нет строки доставки | В таблице нет «Услуга по доставке грузов» / «Доставка плит» / «Доставка свай» |
| Скидка | Надбавка считается от цены после скидки; `d` на доставку не действует |
| НДС вшитого | «в том числе НДС (22%)» = Σ товарных строк × 0,22 (включая надбавку) |
| Обычный файл | `GET .../files/xlsx` и PDF без регрессии (строка доставки на месте, НДС как сейчас) |
| Mixed | `plate_delivery` только на qty плит; `pile_delivery` только на qty свай/мостовых |
| Копейки | Σ долей доставки по строкам = `plate_delivery_total + pile_delivery_total` |
| UI | Кнопка рядом с XLSX; disabled при доставке 0; имя файла с суффиксом |
| Карточка | Плашка «Услуга по доставке» и состав заказа не меняются |

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2 |
| Domain | `core/commercial_pricing.py`, новый pure-helper, `core/commercial_offer_xlsx.py` |
| Archive | `ArchiveService.generate_document`, `ArchiveFileKind` |
| Frontend | React 19, TypeScript; `OfferDetailsDrawer` + `archiveApi` |
| Tests | pytest (математика + archive download); Vitest на кнопку |

---

## Commands

```bash
source venv/bin/activate
pytest tests/test_embed_delivery_in_unit_price.py tests/test_archive_endpoints.py tests/test_commercial_export_mixed.py tests/test_commercial_logistics_cost.py -q

cd frontend && npm run typecheck
cd frontend && npm run test -- --run src/features/commercial-archive
cd frontend && npm run build

./run+logs.sh
```

---

## Project Structure

```
ai_docs/ideas/kp-xlsx-delivery-in-unit-price.md
ai_docs/specs/kp-xlsx-delivery-in-unit-price.md          # этот spec
core/embed_delivery_in_unit_price.py                     # NEW: pure math
tests/test_embed_delivery_in_unit_price.py               # NEW
core/commercial_offer_xlsx.py                            # флаг embed_delivery_in_unit_price
core/commercial_pricing.py                               # без смены формулы НДС обычного КП
app/schemas/archive.py                                   # ArchiveFileKind
app/services/archive_service.py                          # kind → filename + flag
app/api/v1/endpoints/archive.py                          # media_type xlsx для нового kind
frontend/src/features/commercial-archive/
  types/archive.ts
  api/archiveApi.ts
  components/OfferDetailsDrawer.tsx
  components/OfferDetailsDrawer.test.tsx
```

---

## Code Style

Чистая функция, без I/O и без генерации Excel:

```python
from dataclasses import dataclass

PLATE_POOL = frozenset({"plates"})
PILE_POOL = frozenset({"piles", "bridge_piles"})


@dataclass(frozen=True)
class EmbeddedLine:
    index: int
    qty: int
    discounted_unit: float  # после скидки, до надбавки
    surcharge_unit: float   # надбавка на 1 шт
    unit_price: float       # discounted_unit + surcharge_unit
    line_sum: float         # точная сумма строки в рублях (2 знака)


def embed_delivery_in_unit_prices(
    *,
    qty_by_index: list[int],
    product_type_by_index: list[str],
    unit_price_by_index: list[float],
    discount_percent: float,
    plate_delivery_total: float,
    pile_delivery_total: float,
) -> list[EmbeddedLine]:
    """Скидка на изделие, затем доля доставки котла на штуку. Σ line_sum = products + deliveries."""
```

Алгоритм котла (плиты и сваи независимо):

1. Деньги в целых копейках.
2. `discounted_unit_k = round(unit_price * (1 - d/100) * 100)`.
3. Строки котла с `qty > 0`. Если доставка котла > 0, а `total_qty == 0` — **не вшивать этот котёл**: генератор оставляет для него обычную строку доставки (редкий край). Если доставка 0 — надбавка 0.
4. `per_piece_k, extra_k = divmod(delivery_kopecks, total_qty)`.
5. Раздать `extra_k` штукам по порядку строк: строка i забирает `min(qty_i, remaining_extra)` копеек сверх `qty_i * per_piece_k`.
6. `line_delivery_k = qty * per_piece_k + extras_on_line`.
7. `line_sum_k = discounted_unit_k * qty + line_delivery_k`.
8. `unit_price` для колонки «Цена» = `round_to_rub(line_sum_k / qty)` (2 знака).
9. Колонка «Сумма» во вшитом файле — **число** `line_sum_k / 100`, не формула `qty * Цена`, чтобы итог не разъехался из‑за округления цены.

Генератор XLSX:

```python
def generate_commercial_offer_xlsx(..., embed_delivery_in_unit_price: bool = False) -> io.BytesIO:
    totals = calculate_total_cost(...)  # как сейчас, с реальными тарифами
    if embed_delivery_in_unit_price:
        embedded = embed_delivery_in_unit_prices(...)
        # в строках таблицы: Цена/Сумма из embedded; delivery_export не добавлять
        # (кроме котла, который не удалось вшить — см. п.3)
        # НДС-формула: SUM по всем товарным строкам * VAT_RATE
        # после блока НДС — пояснение (D10)
    else:
        # текущее поведение
```

Пояснение (если вшили хотя бы один котёл):

```
В стоимость изделий включена доставка {N} рейс(ов) на сумму {X} ₽. Скидка на доставку не распространяется.
```

`N` = сумма рейсов вшитых котлов, `X` = сумма вшитых доставок. Для mixed с двумя котлами допустима одна фраза с общим N и X (не две строки).

Имя файла в архиве:

```python
filename = f"КП_{kp_id}_с_доставкой_в_цене.xlsx"  # kind xlsx_delivery_in_unit
```

Kind:

```python
ArchiveFileKind = Literal["pdf", "xlsx", "schema", "xlsx_delivery_in_unit"]
```

Media type — тот же spreadsheet, что у `xlsx`.

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Pytest unit | `embed_delivery_in_unit_prices`: скидка не на доставку; Σ сумм = изделия+доставка; две независимые надбавки mixed; extra-копейки; qty=0 пропускается; delivery 0 → надбавка 0 |
| Pytest XLSX | Генерация с флагом: нет строк доставки; Σ «Сумма» = `total_with_vat`; НДС-ячейка = Σ × 0.22; есть пояснение; обычный вызов без флага не изменился |
| Pytest archive | `GET .../files/xlsx_delivery_in_unit` → 200, filename suffix, spreadsheet content-type; `xlsx`/`pdf` без регрессии |
| Vitest | Кнопка есть рядом с XLSX; `disabled` при `delivery_service_total_rub === 0`; URL kind новый; обычные PDF/XLSX не тронуты |

Регрессия обязательна: `tests/test_commercial_logistics_cost.py`, mixed export (`test_commercial_export_mixed.py`), archive generate pdf/xlsx.

---

## Boundaries

**Always:**

- Считать totals тем же `calculate_total_cost`, не дублировать рейсы
- Вшивать только в копии для файла
- Тесты на pure math до правки UI
- Сохранить обычный XLSX/PDF

**Ask first:**

- PDF в том же режиме
- Менять `vat_amount` обычного КП
- Кнопка в визарде
- Новый endpoint вместо kind на существующем `.../files/{kind}`

**Never:**

- UPDATE цен/скидки/логистики в `KP_offers` / строках
- Надбавка по весу в v1
- Скидка на доставку
- Один котёл на mixed (плиты+сваи в кучу)
- Commit секретов / живых БД

---

## Success Criteria

1. В `OfferDetailsDrawer` есть кнопка «XLSX (доставка в цене)» рядом с «XLSX».
2. Скачивание не меняет данные КП; повторный обычный XLSX по-прежнему со строкой доставки.
3. Во вшитом файле нет строк доставки; сумма позиций = прежний итог к оплате; НДС = эта сумма × 22%.
4. Скидка 22% на скрине-кейсе уменьшает только изделие; 156 800 распределяется по штукам после скидки.
5. Mixed: две надбавки не пересекаются по типам строк.
6. При `delivery_service_total_rub === 0` кнопка disabled.
7. pytest/vitest/typecheck по Commands зелёные.

---

## Out of scope (Not Doing)

- PDF «с доставкой в цене»
- Пересчёт НДС в обычном КП/PDF/XLSX
- Визард
- Запись вшитых цен в БД
- Надбавка по весу / по сумме строки
- Telegram / бот
- Отдельный preview в карточке (цены в таблице UI не пересчитываем)

---

## Open Questions

Нет. Ниже — только implementation notes, не продуктовые развилки.

---

## PLAN ✅ / TASKS ✅

Полный план, граф, риски и задачи Task 1–7: [`2026-09-10-kp-xlsx-delivery-in-unit-price.md`](../develop/plans/2026-09-10-kp-xlsx-delivery-in-unit-price.md).

Порядок: RED/GREEN математика → RED/GREEN XLSX → archive kind → кнопка → polish mixed.

Первая задача реализации: **Task 1 RED** (`tests/test_embed_delivery_in_unit_price.py`).
