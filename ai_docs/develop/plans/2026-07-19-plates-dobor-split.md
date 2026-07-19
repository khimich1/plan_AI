# Spec: Плиты «+доб» (добор) — сплит на две позиции в КП

**Created:** 2026-07-19  
**Status:** APPROVED — PLAN/IMPLEMENT  
**Source:** idea-refine (направление α: parser-split + pair metadata + UI-связка)  
**Scope gate:** только шаг ввода плит в коммерческом предложении (не раскладка / СГП / отдельный прайс-модуль)

---

## ASSUMPTIONS I'M MAKING

Исправь сейчас, иначе идём с этим:

1. **Направление α:** детерминированный парсер сплитует строку; в metadata есть `dobor_pairs`; UI красит обе позиции и показывает связку «пара добора».
2. **Ширина добора** всегда `W_доб = 12 − W` (в дециметрах). Пример: `7,2 → 4,8`.
3. **Количество** одно на обе позиции: `… + доб 5-шт` → обе ×5.
4. **MVP-маркеры** (пока нет корпуса примеров 6D):  
   `+ доб`, `+доб`, `доб`, `добор` (регистр/пробелы не важны); qty: `5`, `5шт`, `5-шт`, `5 шт`.
5. **OCR / GigaChat не сплитует** — только желательно не выкидывать маркер `+доб` из текста (мягкое усиление промпта, не ядро).
6. **Сплит делается при нормализации/парсе** (`normalize_order_text` → дальше `PlateParserService`), результат попадает в `normalized_lines` / `order_data` как **две отдельные плиты**.
7. **Edge:** если `W ≥ 12` или `W_доб ≤ 0` — строку **не** сплитуем; оставляем как есть + warning (или unparsed — решить в ревью, см. Open Questions).
8. **Ручное редактирование:** если менеджер удалит/изменит одну из парных строк, metadata пары может устареть; UI не падает, подсветка просто пропадает для «осиротевших» строк.

→ Подтверди или поправь перед PLAN/IMPLEMENT.

---

## Objective

### Problem

В заказах встречается пометка **добор** (`+доб`): одна строка означает **две** плиты одной длины/нагрузки — основную ширину и остаток от формы **12 дм**. Сейчас строка либо не парсится, либо учитывается как одна позиция; менеджер правит вручную.

### User

Менеджер, собирающий КП на шаге «плиты» (текст / OCR → список позиций).

### Goal

Как сделать так, чтобы строка с `+доб` автоматически давала две корректные позиции и на вебе было видно, что это **пара добора**.

### User story

> Как менеджер КП, я хочу, чтобы `ПБ 57-7,2-8п + доб 5-шт` стало двумя позициями  
> `ПБ 57-7,2-8п 5` и `ПБ 57-4,8-8п 5`, обе подсвечены как пара добора,  
> чтобы я не считал остаток ширины и не дублировал строку руками.

### Canonical example

| Input | Output lines | Notes |
|-------|--------------|--------|
| `ПБ 57-7,2-8п + доб 5-шт` | `ПБ 57-7,2-8п 5`<br>`ПБ 57-4,8-8п 5` | `12 − 7,2 = 4,8`; qty=5 на обе; один `dobor_pair_id` |

---

## Tech Stack

| Layer | Tech | Touch |
|-------|------|--------|
| Domain parse | Python 3, `core/plate_text_normalizer.py`, `core/plate_line_parser.py` | сплит + форматирование марок |
| API/service | FastAPI, `app/services/plate_parser_service.py`, draft metadata | `dobor_pairs` в metadata |
| Schemas/types | Pydantic / TS | контракт metadata |
| Frontend | React 19, `frontend/src/features/commercial-offer/` | highlight kind `dobor` + связка |
| Tests | pytest (`tests/`), vitest (`frontend/`) | numeric examples обязательны |

---

## Commands

```bash
# Backend
source venv/bin/activate
pytest tests/test_plate_line_parser.py tests/test_plate_text_normalizer.py tests/test_commercial_web_flow.py -q
# (или новые tests/test_dobor_split.py)

# Frontend
cd frontend && npm run test -- --run src/features/commercial-offer/lib/plateLineHighlights.test.ts
cd frontend && npm run typecheck

# Full local (smoke)
./run+logs.sh
```

---

## Project Structure

```
core/
  plate_text_normalizer.py   # expand +доб → 2 canonical lines; emit pair meta
  plate_line_parser.py       # при необходимости: не ломать parse_line на «хвосте» +доб
app/
  services/plate_parser_service.py
  domain/models/parse_result.py
  schemas/…                  # если metadata типизирован в API
frontend/src/features/commercial-offer/
  types/commercialOffer.ts
  lib/plateLineHighlights.ts
  lib/plateLineHighlights.test.ts   # создать/расширить
  components/PlateListEditor.tsx
tests/
  test_dobor_split.py        # предпочтительно новый focused suite
ai_docs/develop/plans/
  2026-07-19-plates-dobor-split.md   # этот документ
```

**Не трогаем в MVP:** раскладка (`viz_modules/`, `core/optimization/`), СГП, архивный бот, отдельный прайс «новые цены» (это другой пункт Task).

---

## Code Style

Стиль — как в существующем нормализаторе: чистые функции, русские docstring где принято, numeric examples в тестах.

Пример желаемого API (ориентир, не финальная сигнатура):

```python
@dataclass(frozen=True)
class DoborPair:
    pair_id: str
    primary_line: str   # "ПБ 57-7,2-8п 5"
    complement_line: str  # "ПБ 57-4,8-8п 5"
    source_line: str    # исходник с +доб

def expand_dobor_line(line: str) -> tuple[list[str], DoborPair | None]:
    """
    «ПБ 57-7,2-8п + доб 5-шт» →
      (["ПБ 57-7,2-8п 5", "ПБ 57-4,8-8п 5"], DoborPair(...))
    Без маркера добора → ([canonical_or_cleaned], None)
    """
    ...
```

Форматирование ширины добора: сохранить представление в дм с запятой как в каноне парсера (`7,2`, `4,8`, `12`), согласованно с `_parse_pb_width_to_m` / существующим выводом `canonicalize_plate_line`.

---

## Behaviour (нормативная часть)

### Detection

Строка считается добором, если после базовой чистки в **хвосте** (после марки нагрузки) есть маркер из MVP-набора и опционально qty.

Псевдо-паттерн (логика, не единственный regex):

```
<марка ПБ/ПК L-W-Nп> <опц. qty> <маркер доб> <опц. qty>
```

Правило qty:

- Если qty указан только у маркера добора → он для **обеих** позиций.
- Если qty указан у марки и у добора — **MVP: брать qty у добора** (и warning); либо reject — см. Open Questions.
- Если qty нигде нет → `1` на обе.

### Expansion

1. Распарсить основную марку (длина, ширина дм, нагрузка, префикс).
2. `W_c = round(12.0 - W, 1)` (одна цифра после запятой в дм; уточнить в тестах float).
3. Собрать две канонические строки с одинаковым L, load, qty.
4. Назначить `pair_id` (стабильный в рамках одного parse-pass, напр. `dobor-{n}` или uuid4 hex short).
5. Подставить **две** строки вместо одной в поток нормализации.

### Metadata contract

Добавить в draft metadata (рядом с `wide_plate_lines` / `unparsed_lines`):

```ts
dobor_pairs: Array<{
  id: string;
  source_line: string;
  primary_line: string;
  complement_line: string;
}>;
```

`ParseResult` / `build_preview_metadata` прокидывают это поле без потери при append-batch.

### UI

- Новый kind подсветки: `dobor` (отдельный цвет, не пересекается с `wide` / `unparsed` / `correction`).
- Обе строки пары подсвечиваются.
- Связка: легенда «Пара добора» + title на строке вроде `Добор: пара с «ПБ 57-4,8-8п 5»` (или общий badge на обеих).
- Приоритет highlight, если строка одновременно unparsed/wide: **wide > unparsed > dobor > correction** (зафиксировать в коде; можно поменять на ревью).

### OCR (опционально, non-blocking)

В `core/ocr/prompts.py` / `core/plate_format_prompt.py`: одна фраза — сохранять фрагмент `+доб` / `добор` в `raw_name` / тексте, **не** разбивать на две строки на стороне LLM.

---

## Testing Strategy

| Level | Where | What |
|-------|--------|------|
| Unit | `tests/test_dobor_split.py` | expand: пример 57-7,2; маркеры `+доб`/`добор`; qty; edge W=12; нет маркера |
| Unit | existing normalizer/parser tests | регрессия без `+доб` |
| Service | `tests/test_commercial_web_flow.py` или focused | metadata `dobor_pairs` после parse preview |
| Frontend | vitest highlights | две строки с одним `id` → kind `dobor`; легенда |

**Coverage expectation:** каждый numeric example из Success Criteria — отдельный assert; без «примерных» тестов только на regex match без проверки второй марки.

---

## Boundaries

**Always:**

- Писать pytest с числовым примером до/после сплита.
- Прогонять затронутые тесты перед завершением задачи.
- Держать сплит детерминированным (без LLM в hot path).
- Валидировать ширины через существующий `validate_plate_values` после expand.

**Ask first:**

- Менять правило `12 − W` или вводить таблицу пар.
- Расширять скоуп за пределы шага плит КП (цена отходов, раскладка).
- Менять поведение qty при конфликте двух чисел в строке.
- Добавлять зависимости npm/pip.

**Never:**

- Сплитовать в GigaChat как единственный источник истины.
- Коммитить секреты / живые БД.
- Ломать существующие wide-plate / unparsed highlights без миграции legend.

---

## Success Criteria

Конкретные, проверяемые:

1. **Parse:** вход `ПБ 57-7,2-8п + доб 5-шт` → в `normalized_lines` / order ровно две позиции: ширина **0.72 м** ×5 и **0.48 м** ×5, та же длина **5.7 м**, нагрузка **8**.
2. **Имя:** канонические марки содержат `7,2` и `4,8` (или эквивалент, который уже принят парсером для этих ширин).
3. **Metadata:** после preview/parse в draft есть `dobor_pairs` с одним элементом, `primary_line` / `complement_line` совпадают с двумя строками списка.
4. **UI:** обе строки в `PlateListEditor` имеют highlight `dobor`; в легенде есть «Пара добора»; title/связка указывает на напарника.
5. **Regression:** строки **без** `+доб` парсятся как раньше (снимки существующих тестов зелёные).
6. **OCR soft:** отсутствие маркера в OCR не падает пайплайн; при наличии `+доб` в тексте сплит срабатывает на парсере.

---

## Open Questions

1. При `W = 12` + `+доб`: warning и одна строка, или unparsed?
2. Конфликт qty (`ПБ … 3 + доб 5`): брать 5 / брать 3 / unparsed?
3. Нужно ли **переписывать** `input_text` в editor двумя строками (β) или только normalized/order + overlay (α как сейчас в спеке)?
4. Когда появятся реальные OCR-строки (6D) — расширяем ли маркеры без новой спеки или через patch этого файла?

---

## Out of Scope (Not Doing)

| Не делаем | Почему |
|-----------|--------|
| Сплит в GigaChat как ядро | Хрупко; выбран парсер (5A) |
| Раскладка / СГП / двойной учёт отходов | Другие пункты Task; скоуп 1A |
| Таблица пар ширин ≠ `12−W` | Зафиксировано 2A |
| Полноценный «граф связей» при любом edit | Достаточно best-effort metadata |
| Отдельный цвет в PDF КП | Только веб-редактор списка плит |

---

## Related modules

- `core/plate_text_normalizer.py` — `normalize_order_text`, `canonicalize_plate_line`
- `core/plate_line_parser.py` — `parse_line`, qty regex
- `app/services/plate_parser_service.py` — сбор `ParseResult`
- `app/services/commercial_workflow_service.py` / draft metadata builders
- `frontend/.../plateLineHighlights.ts`, `PlateListEditor.tsx`
- OCR (secondary): `core/ocr/prompts.py`, `core/plate_format_prompt.py`

---

## Proposed task breakdown (после approve спеки)

Не выполнять, пока спека не подтверждена:

- [x] **DOB-001** — `expand_dobor_line` + unit tests (canonical example)
- [x] **DOB-002** — встроить expand в `normalize_order_text` / parser pipeline; `dobor_pairs` в `ParseResult` + metadata
- [x] **DOB-003** — API/TS types: `dobor_pairs` на draft
- [x] **DOB-004** — UI highlight `dobor` + легенда + связка
- [x] **DOB-005** — soft OCR prompt note (optional)
- [x] **DOB-006** — commercial flow regression + manual smoke в wizard

---

## Approval gate

Ответь коротко:

- **A)** Спека ок → можно PLAN/IMPLEMENT (`/orchestrate` или `/implement`)
- **B)** Ок с правками: … (вставь Open Questions ответы)
- **C)** Переделать направление (например β: expand в textarea)

---

## Implementation Plan

**Created:** 2026-07-19  
**Approach:** Direction α — parser-split in `normalize_order_text`, `dobor_pairs` metadata, UI highlight kind `dobor`.

### Tasks

| ID | Task | Acceptance criteria | Verification |
|----|------|---------------------|--------------|
| DOB-001 | `expand_dobor_line` in `core/dobor_split.py` + unit tests | Canonical `ПБ 57-7,2-8п + доб 5-шт` → two lines; MVP markers; qty rules; W≥12 warning | `pytest tests/test_dobor_split.py -q` |
| DOB-002 | Integrate into `normalize_order_text`; `dobor_pairs` on `ParseResult` + draft metadata | Two normalized lines; metadata serialized via `CommercialDraftService.serialize_dobor_pairs` | `pytest tests/test_dobor_split.py::test_parse_plate_text_dobor_produces_two_positions -q` |
| DOB-003 | API/TS types for `dobor_pairs` | `CommercialDoborPair` schema; `CommercialDraftMetadata.dobor_pairs`; TS `DoborPair` type | `cd frontend && npm run typecheck` |
| DOB-004 | UI highlight `dobor` + legend + pair connection | Both lines highlighted; legend «Пара добора»; title shows partner; priority wide > unparsed > dobor > correction | `npm run test -- --run plateLineHighlights.test.ts` |
| DOB-005 | Soft OCR prompt note | Prompt mentions preserving `+доб` without LLM split | `pytest tests/test_plate_format_prompt.py -q` |
| DOB-006 | Commercial flow regression | Existing parser/normalizer tests green; manual smoke in wizard | Full verification commands below |

### Checkpoints

1. **After DOB-001:** `expand_dobor_line` passes all marker/qty/edge tests.
2. **After DOB-002:** `PlateParserService.parse_plate_text` returns 2 lines + `dobor_pairs`.
3. **After DOB-004:** Frontend vitest green; both lines show `dobor` highlight in `PlateListEditor`.
4. **After DOB-006:** Full backend + frontend test suite for touched modules.

### Verification commands

```bash
source venv/bin/activate
pytest tests/test_dobor_split.py tests/test_plate_line_parser.py tests/test_commercial_web_flow.py -q
cd frontend && npm run test -- --run src/features/commercial-offer/lib/plateLineHighlights.test.ts
cd frontend && npm run typecheck
```

### Open decisions (resolved for MVP)

| Question | Decision |
|----------|----------|
| W = 12 + доб | Warning + no split (single line) |
| Qty conflict (mark vs dobor) | Use dobor qty + warning |
| Rewrite `input_text` | No — only `normalized_lines` + overlay (α) |
| Orphaned pairs after edit | UI best-effort; no crash |

### Task status

- [x] DOB-001 — `expand_dobor_line` + unit tests
- [x] DOB-002 — normalize pipeline + ParseResult + metadata
- [x] DOB-003 — API/TS types
- [x] DOB-004 — UI highlight + legend + pair connection
- [x] DOB-005 — OCR soft prompt
- [x] DOB-006 — regression tests (3 pre-existing failures in `test_commercial_web_flow.py` unrelated to dobor)
