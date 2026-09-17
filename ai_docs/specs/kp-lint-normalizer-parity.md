# Spec: Подготовка строки заявки до парсера КП

**Статус**: SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅ (2026-09-15) — [2026-09-15-kp-lint-normalizer-parity.md](../develop/plans/2026-09-15-kp-lint-normalizer-parity.md)  
**Дата**: 2026-09-15  
**One-pager**: [ai_docs/ideas/kp-lint-normalizer-parity.md](../ideas/kp-lint-normalizer-parity.md)  
**Родитель**: [unparsed-line-live-highlight.md](./unparsed-line-live-highlight.md) — контракт линта (физическая `\n`-строка, гейт кнопки, не серить «Список верен») не меняем  
**Смежно**: [kp-bridge-piles.md](./kp-bridge-piles.md) Q4 — в КП/PDF марка как в заявке  

Эта спека — источник правды для implement. План дробит работу; идея объясняет «зачем».

---

## Objective

**Проблема.** На шаге 1 живая проверка бьёт в сырой `parse_*_line`. Расчёт чуть чистит текст иначе. Заявка с `шт`, номером списка или ГОСТ мостовых `C 14.35-T7` краснеет, кнопка «Обработать текст» серая, хотя позиция в прайсе есть.

**Цель.** Один построчный prepare копии строки перед парсером на lint, preview и OCR-gate.

- Общий мусор формата — все шесть типов.
- ГОСТ→заводской канон — только мостовые и только с явным `T`/`В`+номер.
- Поле не переписываем. Цену не угадываем. Тип изделия по строке не выбираем.

**Пользователь:** менеджер КП; текст из письма/Excel и фото таблицы в одном срезе.

**Успех (менеджер):**

| Шаг | Ввод | Ожидание |
|-----|------|----------|
| Мостовые | `C14-35T7 80` | зелёное (регресс) |
| Мостовые | `C14-35T7 80 шт` / `C14-35T7 80шт` | зелёное, qty 80, цена канона |
| Мостовые | `C 14.35-T7 80 шт` | зелёное; цена как `C14-35T7`; в PDF/КП написание заявки |
| Мостовые | `C 14.35 80` без класса | красное |
| Плиты | `ПБ 78-12-8п 5шт` | зелёное |
| Плиты | `ПБ 40,3/2,6-8п` | красное |
| Плиты | `C14-35T7 80шт` | красное (не тот тип) |
| Сваи | `С120.35-12 5шт` | зелёное |
| Сваи | `С110.30-9у` / `C 110.30-9у` | зелёное; display с `у`; цена как у `С110.30-9` |
| Сваи | `С120.35-13и` | зелёное; SKU `и`, не `13` |
| Сваи | `С110.30-6у` | зелёное (разобрано); без цены — unpriced, не unparsed |
| Сваи | `C 110.30-9.1у` / `C 110.40-8.1` | display с `.1`/`у`; цена `-9` / `-8` |
| ФБС / ЛС / ЛМ | канон марки + `2шт` | зелёное |

---

## ASSUMPTIONS I'M MAKING

Зафиксировано ideation 2026-09-15. Implement идёт с этим.

1. Текст и фото: тот же prepare в lint, `generate_preview` и OCR `parser_gate`.
2. Textarea не меняем. `/parse` `lines[].text` = как в поле.
3. Линт — физическая строка `\n`. Prepare внутри строки. Целый `normalize_*_order_text` (split `;`, пустые, добор) в lint не звать.
4. `шт`/`шт.`/`штук` после цифры qty, пробел необязателен. Нумерация `1.` / `1)` / `1,`. Юникодные тире. Не вырезать `шт` из середины марки.
5. Мостовой ГОСТ только с `Tn`/`Вn`. `14.35` = 14 м × 35 см → `14-35`, не дециметры обычных свай.
6. `order_data.name` и `.mark` мостовых — заявка после общей чистки без ГОСТ-rewrite. Lookup и геометрия — через `prepare_bridge_pile_mark`.
7. Плиты в lint: общая чистка + `parse_line` + `validate_plate_values`. Catalog/dobor — только preview.
8. Слэш плит не учим. Автотип изделия не делаем.
9. Не серим «Список верен» / «Готово, далее». Не зовём `generate_preview` с дебаунса.
10. Фронт не обязателен. Коммиты — по просьбе.
11. Если в фикстуре прайса нет `C14-35T7`, тесты цены строить на марке из импорта (`C13-35T7` / `C8-35T1`) и ГОСТ под неё (`C13.35-T7`). Зелёный lint при отсутствии SKU → существующий unpriced, не этот срез.
12. D11/D12: суффикс `у` и fallback `.1` — только обычные сваи. `и` не трогаем. Lookup не мапит соседнюю длину (`С110.30-6` ≠ `С100.30-6`).

→ Если что-то неверно — до implement, не в коде.

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| D1 | Подход | Prepare перед парсером, не вектор/ИИ/fuzzy |
| D2 | Чистка | Все шесть типов |
| D3 | Канон нотации | Только мостовые ГОСТ с явным классом |
| D4 | Без класса | Красное, без «похоже на» |
| D5 | Поле | Не переписывать |
| D6 | PDF/КП мостовых | Марка заявки; цена заводского канона |
| D7 | Линт | Per-line prepare, тот же `POST /parse` |
| D8 | Слитное `80шт` | Все типы; после цифры, не только `\s+шт` |
| D9 | Карточка типа | Не угадывать; мостовые на «Плитах» красные |
| D10 | API prepare | `prepare_source_line(raw, product_type)`, `prepare_bridge_pile_mark(mark)` в `core/line_prepare.py` |
| D11 | Сваи «у» | Только `product_type=piles`: хвостовая `у`/`У` после нагрузки — валидный суффикс (усиление / 1C `guid_1c_u`). Display/PDF оставляют `у`. Lookup срезает только хвостовую `у`/`У` (в `pile_prices` таких марок нет). `и` — другой SKU, lookup не срезает. `merge_pile_lines` не сливает `С110.30-9` и `С110.30-9у`. В `order_data` флаг `reinforced: true`. Прочие пять типов — без грамматики `у`. GUID в счёт не этот срез. |
| D12 | Сваи `.1` | Только сваи: если после среза `у` точной цены нет — повтор схлопнуть хвостовой `.\d` **только у числа нагрузки** (`-9.1`→`-9`). Display оставляет `.1`. Textarea не переписываем. Нет fuzzy `С110.30-6`→`С100.30-6`. Порядок lookup: точное → C↔С/пробелы → срез `у` → `.1`. |

---

## Tech Stack

- Backend: Python 3, FastAPI, Pydantic v2, pytest
- Frontend: без обязательных правок (`useSourceTextLint` уже есть)
- Парсеры не меняют канон-regex; меняется вход

---

## Commands

```bash
source venv/bin/activate

pytest tests/test_line_prepare.py tests/test_commercial_line_lint.py tests/test_commercial_parse_lint.py -q
pytest tests/test_bridge_pile_line_parser.py tests/test_commercial_bridge_pile_flow.py tests/test_ocr_parser_gate.py -q
pytest tests/test_commercial_bridge_pile_pricing.py tests/test_pile_catalog_import.py -q -k "bridge or geometry or lookup"

# только если задели фронт (по плану — нет)
cd frontend && npm run test -- --run src/features/commercial-offer && npm run typecheck
```

---

## Project Structure

```
core/line_prepare.py                 NEW
core/*_text_normalizer.py            per-line → prepare_source_line
core/bridge_pile_price_db.py         lookup через prepare_bridge_pile_mark
core/pile_catalog.py                 geometry через тот же mark-prepare
app/services/commercial_line_lint.py prepare затем parse; text = raw
app/services/commercial_bridge_pile_service.py  display vs lookup
core/ocr/*_parser_gate.py            prepare(candidate + qty)
tests/test_line_prepare.py           NEW
tests/test_commercial_line_lint.py   EXTEND
tests/test_ocr_parser_gate.py        EXTEND
tests/test_commercial_bridge_pile_flow.py  EXTEND
```

Роутер `/parse` и схемы HTTP не расширяем.

---

## Code Style

```python
prepared = prepare_source_line(raw, product_type="bridge_piles")
result = parse_bridge_pile_line(prepared)
# lint: ok = result.parsed; LineLint.text = raw
# preview: price(result.mark); name/mark = display из raw после shared cleanup only
```

```python
def prepare_source_line(raw: str, product_type: str) -> str: ...
def prepare_bridge_pile_mark(mark: str) -> str: ...
```

ГОСТ-rewrite узкий: `C`/`С` + длина + `.`/`,` + сечение + `-` + `T`/`Т`/`B`/`В` + номер. Идемпотентен на `C14-35T7`.

---

## Design

### Поток

```
физическая строка textarea (не меняем)
        │
        ├─ lint:   prepare(copy) → parse → ok; ответ.text = original
        ├─ preview: normalize (prepare per line) → parse → price
        └─ OCR:    prepare(candidate qty) → parse → reject или нет
```

### Общая чистка (все типы)

1. NBSP → пробел, юникодные тире → `-`
2. `^\d+[.)]\s+`
3. `(?<=\d)\s*шт(?:ук)?\.?\b`
4. Сжать пробелы

### Мостовые

После чистки, если нет заводского regex — ГОСТ-rewrite только с классом. Lookup: `normalize_bridge_pile_mark_for_lookup` → `prepare_bridge_pile_mark` → C↔С / пробелы.

### Прочие типы

Lint: общая чистка + нынешний парсер. «Сваи 90.30», catalog плит, добор — в своих нормализаторах preview, не в lint (кроме опционального per-line pile `_normalize_line`, если не меняет число строк).

### OCR

Gate не переписывает `raw_name`. Меняется только строка, которую ест парсер. `payload["text"]` — через `normalize_*` после того, как они зовут prepare.

---

## Testing Strategy

**Unit prepare** — таблица Objective + идемпотентность.

**Lint** — те же строки; `lines[].text` исходный; пустые ok; модуль без preview/DraftStore.

**HTTP** — `test_commercial_parse_lint.py` без смены контракта.

**Preview мостовых** — ГОСТ+шт → qty 80, цена канона на фикстуре, display с `14.35` и `T7`.

**OCR gate** — ГОСТ+T и `80шт` без `parser_rejected`.

**Регресс** — `C8-35T1 2`, `ПБ 78-12-8п 2`, `С120.35-12 5`, `strip_pile_load_suffix("C14-40T4")`, `С120.35-13и` ценой `13и` не `13`.

**Вручную (после implement):** мостовые `C14-35T7 80шт` и `C 14.35-T7 80 шт`; плиты + мостовая строка красная; слэш красный.

---

## Boundaries

- **Always:** per-line prepare; канон без регресса; ГОСТ без T красный; слэш красный; мостовые на плитах красные; lint без preview
- **Ask first:** слэш плит; серить «Список верен»; новый URL; fuzzy; перепись textarea; автотип
- **Never:** `generate_preview` с дебаунса; парсер на TypeScript; подставлять `T6`; коммит секретов

---

## Success Criteria

- [x] Lint: `C14-35T7 80 шт`, `C14-35T7 80шт`, `C 14.35-T7 80 шт` → ok на `bridge_piles`
- [x] Lint: `C 14.35 80` → не ok
- [x] Lint: `C14-35T7 80шт` на `plates` → не ok
- [x] Lint: `ПБ 40,3/2,6-8п` → не ok; `ПБ 78-12-8п 5шт` → ok
- [x] Lint: `С110.30-9у` / `C 110.30-9у` → ok; `С120.35-13и` → ok как и-SKU; `С110.30-6у` → ok (parsed)
- [x] `lines[].text` = поле; число записей = физические строки включая пустые
- [x] Preview мостовых: цена канона; display ГОСТ
- [x] Lookup/геометрия понимают ГОСТ-марку с T
- [x] OCR gate мостовых не reject ГОСТ+T и слитное шт
- [x] Сваи: display `у`/`.1`; lookup без `у` и с `.1`→целое; `reinforced` на линии; `и` не срезается
- [x] Перечисленные pytest зелёные

---

## Not in this spec

Векторный поиск, «это C14-35T7?», ИИ-парсер текста, новая грамматика парсера кроме суффикса `у` у обычных свай (D11), автовыбор типа, слэш/добор в lint, GUID `guid_1c_u` в счёт КП.

---

## Open Questions

Нет блокирующих. GUID усиленных свай в счёте КП — отдельный срез (`get_guid_for_invoice(..., reinforced=True)` уже есть).
