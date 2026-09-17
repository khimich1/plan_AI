# Implementation Plan: Подготовка строки заявки до парсера КП

**Спека**: [ai_docs/specs/kp-lint-normalizer-parity.md](../../specs/kp-lint-normalizer-parity.md)  
**Идея**: [ai_docs/ideas/kp-lint-normalizer-parity.md](../../ideas/kp-lint-normalizer-parity.md)  
**Дата**: 2026-09-15  
**Статус**: PLAN ✅ · IMPLEMENT ✅ (2026-09-15, включая D11/D12)

## Overview

Один построчный `prepare` перед `parse_*_line`: общая чистка (`шт` с пробелом и без, нумерация списка, тире) на шесть типов; ГОСТ→заводской канон только у мостовых и только с явным `T`/`В`. Lint, preview и OCR-gate вызывают ту же функцию. Textarea и HTTP `/parse` не меняем. Фронт не трогаем.

## Architecture Decisions

- **Один модуль `core/line_prepare.py`.** Не копипастить regex `шт` по шести нормализаторам. Публичные функции:
  - `prepare_source_line(raw, product_type) -> str` — копия строки для парсера;
  - `prepare_bridge_pile_mark(mark) -> str` — ГОСТ-rewrite марки без qty (для lookup/геометрии).
- **Линт:** на непустой физической `\n`-строке `parse(prepare(raw))`. В ответе `text` = исходный `raw`. Целый `normalize_*_order_text` не звать.
- **Preview:** `normalize_*_order_text` на каждой линии зовёт `prepare_source_line` (плюс уже существующие type-specific починки: «Сваи 90.30», catalog/dobor плит — как сейчас, не в lint).
- **Мостовые display:** `order_data.name`/`mark` = исходная марка после *только* общей чистки (без ГОСТ-rewrite). Lookup цены — `prepare_bridge_pile_mark` внутри `normalize_bridge_pile_mark_for_lookup`.
- **OCR gate:** `prepare_source_line` на `f"{candidate} {qty}"` до `parse_*_line`.
- **Не трогаем:** фронт, URL `/parse`, слэш плит, «Список верен», автовыбор типа изделия (D9). Грамматика парсера свай — только суффикс `у` (D11); `.1` только в lookup (D12).
- **Сваи D11/D12:** `_PILE_MARK_RE` допускает `[иИуУ]`; lookup `get_pile_price` — точное → C↔С → срез `у` → `.1` нагрузки; `order_data.reinforced`; GUID в счёт не этот срез.

## Dependency Graph

```
core/line_prepare.py  (чистка + ГОСТ марки)
    │
    ├── commercial_line_lint          ← зелёные шт/ГОСТ в поле
    ├── *_text_normalizer (per-line)  ← preview/OCR payload text
    ├── commercial_bridge_pile_service (display vs lookup)
    ├── bridge_pile_price_db + pile_catalog.geometry
    └── ocr/*_parser_gate
```

Порядок: prepare + тесты → lint (менеджер уже видит зелёное) → preview/display/lookup → OCR gate. Без фронта.

## Task List

### Phase 1: Prepare + lint

#### Task 1: Общая чистка строки

**Description:** `core/line_prepare.py`: NBSP/тире, срезать `^\d+[.)]\s+`, срезать qty-единицы `(?<=\d)\s*шт(?:ук)?\.?\b`, сжать пробелы. Пока без ГОСТ. `prepare_source_line` для любого `product_type` делает только это.

**Acceptance:**
- [x] `C14-35T7 80шт` → `C14-35T7 80`
- [x] `C14-35T7 80 шт` / `80 шт.` / `80 штук` → qty без единицы
- [x] `1. C14-35T7 80` → без `1.`
- [x] канон без мусора идемпотентен
- [x] `шт` в середине нерелевантной марки не вырезаем тестом на qty-хвост

**Verification:** `pytest tests/test_line_prepare.py -q`

**Dependencies:** None  
**Files:** `core/line_prepare.py`, `tests/test_line_prepare.py`  
**Estimated scope:** S

#### Task 2: ГОСТ-prepare марки мостовых

**Description:** В том же модуле: если после общей чистки марка ещё не заводская (`C14-35T7`), и есть длина`.`сечение`-`T/В+номер — переписать в `C14-35T7` (буквы как в источнике). Без суффикса — не трогать. `prepare_source_line(..., "bridge_piles")` применяет это к строке. `prepare_bridge_pile_mark` — только токен марки.

**Acceptance:**
- [x] `C 14.35-T7 80` → строка, которую ест текущий `parse_bridge_pile_line`
- [x] `C14-35T7 80` без изменений смысла
- [x] `C 14.35 80` без T — без rewrite
- [x] `prepare(prepare(x)) == prepare(x)`

**Verification:** `pytest tests/test_line_prepare.py tests/test_bridge_pile_line_parser.py -q`

**Dependencies:** Task 1  
**Files:** `core/line_prepare.py`, `tests/test_line_prepare.py`  
**Estimated scope:** S

#### Task 3: Lint вызывает prepare

**Description:** `lint_source_lines`: непустая строка → `prepare_source_line(raw, product_type)` → существующий parse. `LineLint.text` остаётся `raw`. Плиты: prepare, затем `parse_line` + `validate_plate_values`. Модуль по-прежнему без preview/DraftStore.

**Acceptance:**
- [x] Таблица приёмки спеки: `80шт` / `80 шт` / ГОСТ+T ok на мостовых; ГОСТ без T not ok; слэш плит not ok; `ПБ … 5шт` и `С120.35-12 5шт` ok
- [x] D9: `C14-35T7 80шт` на `plates` → not ok
- [x] ФБС / ЛС / ЛМ: канон + `2шт` → ok
- [x] Индексы и пустые строки как сейчас
- [x] HTTP `test_commercial_parse_lint.py` зелёный (текст в `lines` исходный)

**Verification:** `pytest tests/test_commercial_line_lint.py tests/test_commercial_parse_lint.py -q`

**Dependencies:** Task 2  
**Files:** `app/services/commercial_line_lint.py`, `tests/test_commercial_line_lint.py`  
**Estimated scope:** S

### Checkpoint: Lint

- [x] `pytest tests/test_line_prepare.py tests/test_commercial_line_lint.py tests/test_commercial_parse_lint.py -q`
- [x] Канонические марки без регресса
- [x] **Не** открывать UI на реализацию — только тесты. Ручная проверка мастера — после всех фаз, когда скажут implement и код уже есть.

### Phase 2: Preview, цена, OCR

#### Task 4: Нормализаторы preview зовут prepare

**Description:** Per-line в `bridge_pile` / `pile` / `fbs` / `step` / `march` text_normalizer: сначала `prepare_source_line`, затем нынешние type-specific правила (префикс «Сваи», «лестничные…»). Плиты: в `basic_text_cleanup` или начале per-line — общая чистка **без** catalog/dobor (они остаются в `normalize_order_text`).

**Acceptance:**
- [x] Старые тесты нормализаторов зелёные
- [x] `C8-35В4 1 шт` и `C8-35В4 1шт` сходятся к парсибельной строке у мостовых
- [x] Добор плит по-прежнему сплитуется только в полном `normalize_order_text`

**Verification:** `pytest tests/test_bridge_pile_format_prompt.py tests/test_pile_ocr_normalizer.py tests/test_plate_normalizer.py -q -k "normaliz"`

**Dependencies:** Task 2  
**Files (4a):** `core/bridge_pile_text_normalizer.py`, `core/pile_text_normalizer.py`, `core/fbs_text_normalizer.py`  
**Files (4b, сразу следом):** `core/step_text_normalizer.py`, `core/march_text_normalizer.py`, `core/plate_text_normalizer.py`  
**Estimated scope:** M (два прохода ≤3 файлов)

#### Task 5: Display vs lookup мостовых в preview

**Description:** `CommercialBridgePileService.generate_preview`: parse подготовленной строки (цена/qty с канона); в `order_data.name` и `.mark` — марка из исходной линии после общей чистки **без** ГОСТ-rewrite (пробелы можно сжать). Lookup `lookup_bridge_pile_price` по заводскому ключу.

**Acceptance:**
- [x] Ввод `C 14.35-T7 80 шт` → qty 80, цена как у `C14-35T7` на фикстуре
- [x] `name`/`mark` содержат `14.35` и `T7` (не обязаны быть `14-35`)
- [x] Канон `C8-35T1 2` — display и цена как сейчас

**Verification:** `pytest tests/test_commercial_bridge_pile_flow.py tests/test_commercial_bridge_pile_pricing.py -q`

**Dependencies:** Task 4  
**Files:** `app/services/commercial_bridge_pile_service.py`, `tests/test_commercial_bridge_pile_flow.py`  
**Estimated scope:** S

#### Task 6: Lookup и геометрия понимают ГОСТ-марку

**Description:** `normalize_bridge_pile_mark_for_lookup` вызывает `prepare_bridge_pile_mark` до C↔С/пробелов. `parse_bridge_pile_geometry` — либо тот же prepare, либо regex ГОСТ `14.35-T7` → 14 м × 350 мм. Регресс `C14-40T4` → 14 м × 400 мм.

**Acceptance:**
- [x] Lookup ГОСТ-марки находит ту же строку прайса, что заводской канон (A11: если нет `C14-35T7` в фикстуре — `C13-35T7` / `C8-35T1` и ГОСТ под неё)
- [x] `strip_pile_load_suffix("C14-40T4")` без регресса
- [x] ГОСТ без T геометрию не выдумывает

**Verification:** `pytest tests/test_pile_catalog_import.py tests/test_bridge_pile_price_import.py -q -k "bridge or geometry or lookup"`

**Dependencies:** Task 2  
**Files:** `core/bridge_pile_price_db.py`, `core/pile_catalog.py`, соответствующий test  
**Estimated scope:** S

#### Task 7: OCR parser-gate

**Description:** Все `*_parser_gate` (включая `parser_gate.py` плит): перед parse — `prepare_source_line` на `candidate + qty`. Не менять JSON `raw_name` в ответе OCR, только вход в parse.

**Acceptance:**
- [x] Мостовой candidate `C 14.35-T7` + qty 80 → нет `parser_rejected`
- [x] `C14-35T7` + `80шт` в candidate-строке → нет reject
- [x] Существующий `tests/test_ocr_parser_gate.py` зелёный + новые кейсы

**Verification:** `pytest tests/test_ocr_parser_gate.py -q`

**Dependencies:** Task 2  
**Files:** `core/ocr/bridge_pile_parser_gate.py`, `core/ocr/pile_parser_gate.py`, `core/ocr/parser_gate.py`, `core/ocr/fbs_parser_gate.py` (+ step/march если не влезает — 7b)  
**Estimated scope:** M

#### Task 8: Сваи «у» и fallback `.1`

**Description:** D11: парсер обычных свай принимает хвостовую `у`/`У` после нагрузки; OCR-repair «Сваи/Свай» тоже; `strip_pile_load_suffix` срезает `-6у`/`-9.1у`, но не мостовую `C14-40T4`. Display и merge сохраняют `у` отдельно от без-`у`. `order_data.reinforced=true`. Lookup не срезает `и`. D12: если цены нет — схлопнуть `.\d` только у нагрузки. Порядок в `get_pile_price`: точное → C↔С/пробелы → срез `у` → `.1`.

**Acceptance:**
- [x] Lint `С110.30-9у` / `C 110.30-9у` ok; `С120.35-13и` ok как и-SKU; `С110.30-6у` parsed
- [x] Preview: `C 110.30-9.1у` цена как `С110.30-9`, mark с `.1` и `у`, `reinforced: true`
- [x] `C 110.40-8.1` цена как `С110.40-8`; display с `.1`
- [x] `С120.35-13и` не схлопывается в `13`
- [x] `strip_pile_load_suffix("C14-40T4")` без изменений
- [x] `get_guid_for_invoice` из КП не вызывается

**Verification:** `pytest tests/test_pile_line_parser.py tests/test_pile_catalog_import.py tests/test_commercial_pile_pricing.py tests/test_commercial_pile_flow.py tests/test_pile_price_import.py -q`

**Dependencies:** Task 3  
**Files:** `core/pile_line_parser.py`, `core/pile_catalog.py`, `core/pile_text_normalizer.py`, `core/pile_price_db.py`, `app/services/commercial_pile_service.py`, соответствующие tests  
**Estimated scope:** M

### Checkpoint: Complete (после implement)

- [x] `pytest tests/test_line_prepare.py tests/test_commercial_line_lint.py tests/test_commercial_parse_lint.py tests/test_bridge_pile_line_parser.py tests/test_commercial_bridge_pile_flow.py tests/test_ocr_parser_gate.py -q`
- [ ] Вручную: шаг «Мостовые сваи» — `C14-35T7 80шт` и `C 14.35-T7 80 шт` не красные, состав с ценой; шаг «Плиты» + мостовая строка — красная; слэш плит — красный
- [x] Фронт `npm run test` / typecheck только если что-то всё же задели (по плану — нет)

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Lint через целый `normalize_*_order_text` съедет индексами | High | Только per-line prepare; пустые строки не выкидывать |
| ГОСТ без T станет «как попало» | High | Rewrite только с `T`/`В`+цифры |
| PDF покажет заводской канон | Med | Task 5: display из raw; lookup отдельно |
| `80шт` срежет не qty | Low | Якорь `(?<=\d)` и хвост строки / токен qty |
| OCR verify станет реже на кривом candidate | Low | Канон только узкий regex; цифры не округляем |

## Open Questions

Нет блокирующих. Если в фикстуре нет `C14-35T7` — тесты цены на марке из импорта (`C13-35T7` / `C8-35T1`), ГОСТ под неё. GUID `у` в счёт — не этот срез.

## Parallelization

После Task 2 параллельно: Task 3 (lint), Task 4 (нормализаторы), Task 6 (lookup), Task 7 (OCR). Task 5 после 4. Task 8 после Task 3.
