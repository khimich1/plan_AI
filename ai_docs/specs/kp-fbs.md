# Spec: КП на ФБС

> **Источник идеи:** [`ai_docs/ideas/kp-bridge-piles-and-fbs.md`](../ideas/kp-bridge-piles-and-fbs.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ (in plan) → TASKS ✅ → IMPLEMENT  
> **Статус:** implemented (2026-08-05)  
> **Plan:** [`ai_docs/develop/plans/2026-08-05-kp-fbs.md`](../develop/plans/2026-08-05-kp-fbs.md)  
> **Образец UX:** [`kp-piles.md`](kp-piles.md) — grade matrix как у свай (не ЛС)  
> **Шаблон кода:** [`kp-bridge-piles.md`](kp-bridge-piles.md) / marches — multi-product wiring  
> **Прайс:** `банк знаний/7_5 В прайс на ФБС  от 03.08.2026.xlsx`, лист «Прайс» (**14** SKU)  
> **Связанные модули:** `CommercialOfferWizard`, pile/bridge_pile ветки, `commercial.py`, `KpPersistenceService`

---

## Decisions locked (2026-08-05)

| # | Тема | Решение |
|---|------|---------|
| Q1 | `product_type` / UI | **`fbs`**, UI «ФБС» |
| Q2 | Таблицы | `fbs_prices` (pb.db), `kp_fbs` (plita.db) |
| Q3 | Модель цены | **mark + concrete_grade** (clone piles / bridge_piles / marches) |
| Q4 | Grades | Колонки `7.5 \| 20 \| 22.5 \| 25` → коды **`B7_5`, `B20`, `B22_5`, `B25`** |
| Q5 | Default grade | **`B25`** (как piles/marches; имя файла «7_5» — не default) |
| Q6 | Матрица | **Плотная** (у каждой марки все 4 класса с ценой); available-only dropdown всё равно |
| Q7 | Алиасы T/В | **Нет** (в отличие от bridge piles) |
| Q8 | Импорт | CLI; **только лист «Прайс»**; skip empty/zero cells |
| Q9 | Mark canon | Как в Excel (`ФБС 9.3.6-Т`); normalize spaces/case для lookup; **display = как ввёл/распознал менеджер** |
| Q10 | Preview | Скрывать до «Список верен» |
| Q11 | Client-step errors | Не на шаге ввода |
| Q12 | Bulk «класс ко всем» | Как piles (обычно все grades доступны); если недоступен — **skip + warning** (как bridge piles A) |
| Q13 | OCR | В MVP — dedicated prompt |
| Q14 | Production | Whitelist **plates** only; shared `kp_id` |
| Q15 | Scope | Точечный clone; **не** generic framework; не смешивать продукты |

---

## Assumptions I'm Making

1. **Web-only** — React-мастер КП; бот вне scope.
2. **`product_type = "fbs"`**, UI-лейбл **«ФБС»**.
3. **Не писать** в `pile_prices` / `bridge_pile_prices` / `kp_piles` / `kp_bridge_piles`.
4. **14 SKU** текущего файла — полный каталог MVP (`ФБС 9.3.6-Т` … `ФБС 24.6.6-Т`).
5. **Нормализация lookup:** collapse spaces, case-insensitive, `Т`↔`T` в суффиксе; display без подмены на канон прайса.
6. **OCR** — отдельный system prompt под марки ФБС.
7. **Точечный clone** bridge_pile/pile ветки; без generic multi-product framework.
8. **Shared `kp_id`**; production candidates exclude `fbs`.

→ Если что-то неверно — поправьте **до** перехода к PLAN (решения уже locked пользователем).

---

## Objective

Дать менеджеру **отдельное КП на ФБС** с UX как у свай: ввод текста/фото/ИИ, класс бетона per-line, цена из прайса, PDF/XLSX, архив — без производства.

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | создаю КП | выбрать «ФБС» | не перепутать с плитами/сваями |
| US-2 | есть заявка | вставить текст/фото | не искать цены в Excel |
| US-3 | вижу preview | указать класс бетона; марка как в заявке | цена верная |
| US-4 | опечатка / нет в прайсе | ошибка до расчёта | не отдать неверную цену |
| US-5 | заполнил клиента | PDF/XLSX + архив | закрыть сделку |
| US-6 | смотрю архив | бейдж «ФБС» | отличать тип |

### Reframed success criteria

| Требование | Критерий |
|------------|----------|
| «Как сваи» | 3 шага; text/photo/AI; grade UI; calculate → PDF/XLSX → save |
| «100% из прайса» | `get_fbs_price(mark, grade)`; block on missing |
| «Не другие продукты» | Отдельный `product_type`, таблицы, picker |
| «Без производства» | `fbs` не в production candidates |

---

## Tech Stack

Как у piles/bridge_piles: FastAPI + React wizard + SQLite (`pb.db` / `plita.db`) + pytest/vitest.  
Шаблон кода: **bridge_pile-ветка** (`bridge_pile_*` → `fbs_*` / `Fbs*` / `FBS*`).

---

## Commands

```bash
source venv/bin/activate
uvicorn app.main:app --reload

python scripts/import_fbs_prices_from_xlsx.py \
  "банк знаний/7_5 В прайс на ФБС  от 03.08.2026.xlsx" --sheet Прайс

pytest tests/ -k "fbs or bridge_pile or march or step or pile or wizard" -q

cd frontend && npm run typecheck && npm run test && npm run build
```

---

## Project Structure (proposed)

```
core/
  fbs_price_db.py
  fbs_line_parser.py
  fbs_format_prompt.py
  fbs_text_normalizer.py
  commercial_pricing.py         # EXTEND
  commercial_offer*.py          # EXTEND
  kp_db_schema.py               # EXTEND — kp_fbs
  kp_persistence_service.py / kp_order_data.py / offers_read.py
  ocr/fbs_parser_gate.py

scripts/import_fbs_prices_from_xlsx.py

app/
  services/commercial_fbs_service.py
  schemas + endpoints .../fbs + .../ai + .../grades
  archive badge/filter; production whitelist unchanged

frontend/
  ProductTypePicker + «ФБС»
  FbsInputStep, KpFbsPreviewPanel
  archive badge/filter/drawer

tests/
  test_fbs_*, test_commercial_fbs_flow.py, …
```

---

## Architecture (sketch)

```
Picker → fbs
  → FbsInputStep (text/photo/AI)
  → Preview (после «Список верен»): Марка | Класс ▼ | Кол-во | Цена
  → Client → Result (pdf+xlsx) → Archive
```

**Parsing examples (MVP):**

```
ФБС 9.3.6-Т 2
ФБС 9.3.6-Т B25 2
ФБС 12.4.6-Т 25 1
фбс 24.6.6-т B7.5 3
```

**`fbs_prices`:**

```sql
CREATE TABLE IF NOT EXISTS fbs_prices (
    mark TEXT NOT NULL,
    concrete_grade TEXT NOT NULL,
    price REAL NOT NULL,
    display_name TEXT,
    price_list_date TEXT,
    imported_at TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (mark, concrete_grade)
);
-- Import only sheet «Прайс»; skip zero/empty cells.
-- No variant_group (no T/В aliases).
```

**`kp_fbs`:** mark, concrete_grade, qty, unit_price, discounted_price (mirror `kp_piles`).  
`mark` = as recognized/typed (light whitespace normalize), not forced Excel canon.

**Grades UX:**
- Options = grades with price for mark (subset of `{B7_5, B20, B22_5, B25}`).
- Default: prefer `B25` if priced; else single available; else first.
- Bulk apply: apply where grade exists; **skip + warning** if somehow unavailable.

---

## Code Style

Зеркало `bridge_pile_*` / `BridgePile*`; имена `fbs_*` / `Fbs*` / UI «ФБС».  
`product_kind: "fbs"`.  
Не рефакторить другие продукты beyond точек расширения.

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit | parser (ФБС marks, grades 7.5/20/22.5/25); import 14 marks × 4 grades |
| Unit | pricing; unpriced block; available-grades |
| Integration | create → grades → calculate → save |
| Archive / production | badge; excluded from production |
| Regression | plate + pile + step + march + bridge_pile flows |

---

## Boundaries

### Always
- pytest/vitest before merge
- Immutable `product_type` after draft create
- Block calculate/save on unknown mark+grade
- Production = plates only
- Не писать ФБС в `kp_piles` / `kp_bridge_piles` / `pile_prices` / `bridge_pile_prices`

### Ask first
- SQLite schema beyond `kp_fbs` / `fbs_prices`, new deps, OCR provider changes
- Generic multi-product framework
- Changing default grade away from B25

### Never
- Mix product types in one KP
- Treat FBS as piles/bridge_piles
- Silent overwrite of manager-typed mark with price-list canon
- Fuzzy-match without agreement
- Commit secrets / live DBs
- Import sheets other than «Прайс»

---

## Acceptance Criteria (MVP)

- [x] AC-1 Picker «ФБС» → `product_type=fbs`
- [x] AC-2 Input text/photo/AI
- [x] AC-3 Parser mark + optional grade + qty; merge mark+grade
- [x] AC-4 Preview: grade dropdown (available-only); hide until confirm
- [x] AC-5 CLI import only «Прайс»; 14 marks; grades B7_5/B20/B22_5/B25; skip zeros
- [x] AC-6 Missing mark+grade → block calculate
- [x] AC-7 Client/result; files pdf+xlsx only
- [x] AC-8 PDF/XLSX: марка as typed, класс, кол-во, цена, сумма
- [x] AC-9 Save `kp_fbs` + meta
- [x] AC-10 Archive badge + filter
- [x] AC-11 Plate/pile/step/march/bridge_pile regression
- [x] AC-12 Not in production
- [x] AC-13 Shared `kp_id`
- [x] AC-14–17 Mirror piles (drawer, regen, disabled production, merge, grades API + bulk skip+warning)

---

## Not Doing

| Item | Why |
|------|-----|
| Generic framework | После клонов — отдельное обсуждение |
| Mixing products | Одно КП = один тип |
| T/В alias groups | Нет в прайсе ФБС |
| UI-импорт прайса | CLI |
| Производство/СГП | plates only |
| Default B7.5 | Locked: B25 |

---

## Next

SPECIFY approved → PLAN: `ai_docs/develop/plans/2026-08-05-kp-fbs.md` → IMPLEMENT (TDD, clone bridge_pile/pile ветки).
