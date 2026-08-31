# Spec: КП на лестничные марши (ЛМ)

> **Источник:** прайс `банк знаний/Прайс ЛМ от 03.08.2026.xlsx` (лист «Прайс», **7 SKU**)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS ✅ → IMPLEMENT ✅  
> **Статус:** implemented (2026-08-05)  
> **Report:** [`ai_docs/develop/reports/2026-08-05-kp-stair-marches-implementation.md`](../develop/reports/2026-08-05-kp-stair-marches-implementation.md)  
> **Предшественники:** [`kp-stair-steps.md`](kp-stair-steps.md) (`steps`), [`kp-piles.md`](kp-piles.md) (`piles`)  
> **Идея rollout:** ЛС → **ЛМ** → ФБС  

---

## Decisions locked (2026-08-05)

| # | Тема | Решение |
|---|------|---------|
| Q1 | `product_type` | **`marches`** (UI: «Марши») |
| Q2 | Модель цены | **mark + concrete_grade** (как сваи); UI dropdown + bulk; default **B25** |
| Q3 | Нормализация `ЛМ 2,8` | Парсер принимает `2,8` и `2.8`; **канон ключа/вывода = как в прайсе** (`ЛМ 2,8`, запятая). Регистр/пробелы нормализуем как у других продуктов. |
| Q4 | «закладные справа» | **Отдельный SKU** — в прайсе отдельная строка с тем же рядом цен; ключ включает суффикс |
| Q5 | PDF/XLSX name | Короткая марка **без** «Лестничные марши» |
| Q6 | Preview до «Список верен» | **Скрывать** (как у ЛС после фикса) |
| Q7 | OCR | Да, в MVP |
| Q8 | Таблицы | `kp_marches` + `march_prices` |
| Q9 | Архив | Бейдж «Марши» + фильтр |
| Q10 | Нумерация | Общая серия `kp_id` |
| Q11 | Dedup | mark+grade → merge qty |
| Q12 | PDF колонки | № \| Марка \| Класс бетона \| Кол-во \| Цена \| Сумма |
| Q13 | Production | Whitelist plates (уже есть) |

---

## Assumptions I'm Making

1. **Web-only** — React-мастер КП; бот вне scope.
2. **Отдельный `product_type`** (предложение: **`marches`**, UI «Марши» / «Лестничные марши») — не смешивать с `steps`.
3. **Прайс — матрица классов**, как у свай: колонки `15 | 20 | 22.5 | 25 | 30 на граните` → коды `B15`, `B20`, `B22_5`, `B25`, `B30_granite` (reuse `GRADE_CODES` / `grade_code_from_value` из pile path).
4. **UX = clone свай** (не ступеней): ввод текст/фото/ИИ + **dropdown класса бетона** + «применить ко всем».
5. **Ключ цены = короткая марка + grade.** Примеры ключей марок из прайса:
   - `1ЛМ 27-11-14-4`
   - `1ЛМ 27-12-14-4`
   - `1ЛМ 30-11-15-4`
   - `1ЛМ 30-11-15-4 закладные справа`
   - `1ЛМ 30-12-15-4`
   - `ЛМ 2,8`
   - `ЛМ 2,9`  
   Без префикса «Лестничные марши» в PDF/lookup.
6. **Default grade** при отсутствии в тексте — **B25** (как у свай), пока не скажете иначе.
7. **Одно КП = один тип**; ЛМ не мешать с ЛС/плитами/сваями в одном документе.
8. **OCR в MVP** — отдельный system prompt под марки ЛМ (как у свай/ступеней).
9. **Production** — whitelist `plates` уже есть; марши не попадают в производство.
10. **Импорт прайса** — CLI → `march_prices(mark, concrete_grade, price)` в `pb.db`; UI-импорт out.
11. **Точечный clone** pile/step-ветки; **не** делаем generic product framework в этом релизе (но после ЛМ стоит запланировать).
12. **7 SKU** текущего файла — полный каталог MVP.

---

## Objective

Дать менеджеру **отдельное КП на лестничные марши (ЛМ)** с UX как у **свай**: ввод текста/фото, класс бетона per-line, цена из матрицы прайса, PDF/XLSX, архив.

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | создаю КП | выбрать «Марши» | не путать с ступенями/сваями |
| US-2 | есть заявка | вставить текст/фото | не искать цены в Excel |
| US-3 | вижу preview | указать класс бетона | цена из нужной колонки |
| US-4 | опечатка / нет в прайсе | ошибка до расчёта | не отдать неверную цену |
| US-5 | заполнил клиента | PDF/XLSX + архив | закрыть сделку |
| US-6 | смотрю архив | бейдж «Марши» | отличать тип |

### Reframed success criteria

| Требование | Критерий |
|------------|----------|
| «Как сваи» | 3 шага; text/photo/AI; grade UI; calculate → PDF/XLSX → save |
| «100% из прайса» | `get_march_price(mark, grade)`; block on missing |
| «Не как ступени» | Есть `concrete_grade` в order_data / `kp_marches` / PDF |
| «Без производства» | `product_type=marches` не в production candidates |

---

## Tech Stack

Как у steps/piles: FastAPI + React wizard + SQLite (`pb.db` / `plita.db`) + pytest/vitest.  
Шаблон кода: **pile-ветка** (из‑за grade), структура файлов как у steps (`march_*`).

---

## Commands

```bash
source venv/bin/activate
uvicorn app.main:app --reload

python scripts/import_march_prices_from_xlsx.py \
  "банк знаний/Прайс ЛМ от 03.08.2026.xlsx" --sheet Прайс

pytest tests/ -k "march or commercial" -q
pytest tests/test_commercial_pile_flow.py tests/test_commercial_step_flow.py -q

cd frontend && npm run typecheck && npm run test && npm run build
```

---

## Project Structure (proposed)

```
core/
  march_price_db.py           # NEW — matrix import + get_march_price
  march_line_parser.py        # NEW — 1ЛМ… / ЛМ 2,8 + optional grade + qty
  march_format_prompt.py      # NEW — OCR/AI
  march_text_normalizer.py    # NEW
  commercial_pricing.py       # EXTEND
  commercial_offer*.py        # EXTEND — pile-like columns (with grade)
  kp_db_schema.py             # EXTEND — kp_marches
  kp_persistence_service.py / kp_order_data.py / offers_read.py

scripts/import_march_prices_from_xlsx.py

app/
  services/commercial_march_service.py
  schemas + endpoints .../marches + .../marches/ai + .../marches/grades
  archive / production whitelist (already plates-only)

frontend/
  ProductTypePicker + «Марши»
  MarchInputStep, KpMarchPreviewPanel (with grade)
  archive badge/filter/drawer

tests/
  test_march_*, test_commercial_march_flow.py, …
```

---

## Architecture (sketch)

```
Picker → marches
  → MarchInputStep (text/photo/AI)
  → Preview: Марка | Класс ▼ | Кол-во | Цена
  → Client → Result (pdf+xlsx) → Archive
```

**Parsing examples (MVP):**

```
1ЛМ 27-11-14-4 2
1ЛМ 27-11-14-4 B25 2
ЛМ 2,8 5
1ЛМ 30-11-15-4 закладные справа B22.5 1
```

**`march_prices`:**

```sql
CREATE TABLE IF NOT EXISTS march_prices (
    mark TEXT NOT NULL,
    concrete_grade TEXT NOT NULL,
    price REAL NOT NULL,
    display_name TEXT,
    price_list_date TEXT,
    imported_at TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (mark, concrete_grade)
);
```

**`kp_marches`:** как `kp_piles` (mark, concrete_grade, qty, unit_price, discounted_price).

---

## Code Style

Зеркало `pile_*` / `Pile*`; имена `march_*` / `March*`.  
`product_kind: "march"`.  
Не рефакторить plates/piles/steps beyond точек расширения.

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit | parser (вкл. `ЛМ 2,8`, «закладные справа»), import 7×5 grades |
| Unit | pricing / unpriced block |
| Integration | create → grades → calculate → save |
| Archive / production | badge; excluded from production |
| Regression | plate + pile + step flows green |

---

## Boundaries

### Always
- pytest/vitest before merge
- Immutable `product_type` after draft create
- Block calculate/save on unknown mark+grade
- Production = plates only

### Ask first
- SQLite schema, new deps, OCR provider changes
- Generic multi-product framework
- Adding ФБС in same PR

### Never
- Mix product types in one KP
- Write marches into `kp_steps` / `kp_piles`
- Fuzzy-match without agreement
- Commit secrets / live DBs

---

## Acceptance Criteria (MVP)

- [x] AC-1 Picker «Марши» → `product_type=marches`
- [x] AC-2 Input text/photo/AI
- [x] AC-3 Parser mark + optional grade + qty; merge mark+grade
- [x] AC-4 Preview with grade dropdown + apply-all
- [x] AC-5 CLI import ≥7 marks × 5 grades
- [x] AC-6 Missing mark+grade → block calculate
- [x] AC-7 Client/result; files pdf+xlsx only
- [x] AC-8 PDF/XLSX: марка, класс, кол-во, цена, сумма
- [x] AC-9 Save `kp_marches` + meta
- [x] AC-10 Archive badge + filter
- [x] AC-11 Plate/pile/step regression
- [x] AC-12 Not in production
- [x] AC-13 Shared `kp_id`
- [x] AC-14–18 Mirror piles (drawer, regen, disabled production, merge, grades API)

---

## Not Doing

| Item | Why |
|------|-----|
| ФБС | Следующая итерация |
| Смешанное КП ЛС+ЛМ | Продуктовое правило |
| Generic framework | После 2–3 клонов |
| UI-импорт прайса | CLI |
| Производство/СГП | Как у свай/ступеней |

---

## Open Questions — resolved 2026-08-05

| # | Вопрос | Решение |
|---|--------|---------|
| OQ-1 | `product_type` | **`marches`**, UI «Марши» |
| OQ-2 | Бетон UI / default | Как сваи; default **B25** |
| OQ-3 | `ЛМ 2,8` vs `2.8` | Распознавать оба; канон вывода/ключа = **`ЛМ 2,8`** (как в прайсе) |
| OQ-4 | «закладные справа» | Отдельная строка прайса → **отдельная марка** (не опция UI) |
| OQ-5 | PDF name | Только короткая марка |
| OQ-6 | Preview | Скрывать до «Список верен» |

---

## Next

IMPLEMENT complete — see report. Optional: dedicated OCR pipeline test; ФБС as next product.
