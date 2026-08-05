# Spec: КП на мостовые сваи

> **Источник идеи:** [`ai_docs/ideas/kp-bridge-piles-and-fbs.md`](../ideas/kp-bridge-piles-and-fbs.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ (review) → TASKS ✅ (in plan) → IMPLEMENT  
> **Статус:** implemented (2026-08-05)  
> **Plan:** [`ai_docs/develop/plans/2026-08-05-kp-bridge-piles.md`](../develop/plans/2026-08-05-kp-bridge-piles.md)  
> **Образец:** [`kp-piles.md`](kp-piles.md) (`product_type=piles`) — UX/архитектура **как у свай**, не как ЛМ  
> **Прайс:** `банк знаний/Прайс на мостовые сваи от 03.08.2026.xlsx`, лист «Прайс» (**64** SKU)  
> **Следом (не этот spec):** ФБС — `банк знаний/7_5 В прайс на ФБС  от 03.08.2026.xlsx`  
> **Связанные модули:** `CommercialOfferWizard`, pile-ветка (`core/pile_*`, `PileInputStep`, `KpPilePreviewPanel`), `commercial.py`, `KpPersistenceService`

---

## Decisions locked (ideation 2026-08-05)

| # | Тема | Решение |
|---|------|---------|
| Q1 | Кто | Только менеджеры КП |
| Q2 | Порядок продуктов | Сначала **мостовые сваи**, потом ФБС |
| Q3 | UI vs обычные сваи | **Отдельная** карточка picker (не подтип `piles`) |
| Q4 | Алиасы `T` / `В` | Lookup по синонимам группы (одна цена). **В КП/PDF — как вписал менеджер** на шаге ввода (без принудительной смены на «канон») |
| Q5 | UX / образец | Clone **свай** (`kp-piles`), не ЛМ |
| Q6 | Одно КП | Один `product_type`; смешанных позиций нет |
| Q7 | Production | Whitelist `plates` only; мостовые не в СГП |
| Q8 | Нумерация | Общая серия `kp_id` |
| Q9 | Импорт | CLI → `pb.db`; **только лист «Прайс»** (остальные листы Excel игнор) |
| Q10 | `product_type` / лейбл | **`bridge_piles`**, UI «Мостовые сваи» |
| Q12 | Bulk «класс ко всем» | **A:** skip строк без этого класса в прайсе + warning |

---

## Assumptions I'm Making

1. **Web-only** — React-мастер КП; бот вне scope.
2. **`product_type = "bridge_piles"`**, UI-лейбл **«Мостовые сваи»** (не путать с «Сваи»).
3. **Не писать** в `pile_prices` / `kp_piles` — отдельные таблицы `bridge_pile_prices` / `kp_bridge_piles`.
4. **Модель цены = mark + concrete_grade**, как у свай.
5. **Grade codes для мостовых:** только **`B25`** и **`B30`** (колонки прайса `25` и `30`). Не использовать `B30_granite` и не тащить полный `GRADE_CODES` обычных свай в dropdown.
6. **Разреженная матрица:** у марки цена обычно только в одной колонке. Dropdown grades показывает **только классы с ненулевой ценой** для этой марки. Если доступен ровно один класс — подставлять его (даже если это B30).
7. **Default grade** при отсутствии в тексте: если у марки один доступный класс → он; иначе **`B25`** (если есть цена), иначе единственный оставшийся / ошибка unpriced.
8. **Алиасы импорта:** строка `C8-35T4; C8-35В4` → группа синонимов с одной ценой по grades; lookup находит цену по любой форме группы. **Display / PDF = текст марки, как ввёл менеджер** (после нормализации пробелов/регистра, без подмены на другую форму из группы).
9. **Нормализация:** лат. `C` ↔ кир. `С`; `B` ↔ `В` в суффиксе для lookup; trim пробелов. Не переписывать ввод менеджера на «левую часть до `;`».
10. **OCR в MVP** — отдельный system prompt под марки мостовых свай (clone pile OCR).
11. **Preview** скрывать до «Список верен» (как у ЛС/ЛМ после фикса; тот же паттерн wizard).
12. **Client-step errors** не показывать на шаге ввода (как у piles/steps/marches).
13. **Точечный clone** pile-ветки; **не** generic multi-product framework в этом релизе.
14. **64 SKU** текущего файла — полный каталог MVP; источник — **только лист «Прайс»**.
15. **ФБС** — отдельный spec/PR после приёмки мостовых.
16. **Variant dropdown** из ранней ideation — **снят** (см. Q4): при одинаковой цене достаточно того, что вписал менеджер.

→ Если что-то неверно — поправьте **до** перехода к PLAN.

---

## Objective

Дать менеджеру **отдельное КП на мостовые сваи** с UX как у обычных свай: ввод текста/фото, класс бетона per-line, выбор вариации марки при нескольких формах, цена из прайса, PDF/XLSX, архив — без производства.

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | создаю КП | выбрать «Мостовые сваи» отдельно от «Сваи» | не перепутать каталоги |
| US-2 | есть заявка | вставить текст/фото | не искать цены в Excel |
| US-3 | вижу preview | указать класс бетона; марка как в заявке | цена верная, в КП то же написание |
| US-4 | опечатка / нет в прайсе | ошибка до расчёта | не отдать неверную цену |
| US-5 | заполнил клиента | PDF/XLSX + архив | закрыть сделку |
| US-6 | смотрю архив | бейдж «Мостовые сваи» | отличать тип |

### Reframed success criteria

| Требование | Критерий |
|------------|----------|
| «Как сваи» | 3 шага; text/photo/AI; grade UI; calculate → PDF/XLSX → save |
| «100% из прайса» | `get_bridge_pile_price(mark, grade)`; block on missing |
| «Алиасы» | Lookup по группе синонимов; PDF/UI показывают марку как ввёл менеджер |
| «Не обычные сваи» | Отдельный `product_type`, таблицы, picker |
| «Без производства» | `bridge_piles` не в production candidates |

---

## Tech Stack

Как у piles: FastAPI + React wizard + SQLite (`pb.db` / `plita.db`) + pytest/vitest.  
Шаблон кода: **pile-ветка** (`pile_*` → `bridge_pile_*` / `BridgePile*`).

---

## Commands

```bash
source venv/bin/activate
uvicorn app.main:app --reload

python scripts/import_bridge_pile_prices_from_xlsx.py \
  "банк знаний/Прайс на мостовые сваи от 03.08.2026.xlsx" --sheet Прайс

pytest tests/ -k "bridge_pile or pile or commercial" -q
pytest tests/test_commercial_pile_flow.py -q

cd frontend && npm run typecheck && npm run test && npm run build
```

---

## Project Structure (proposed)

```
core/
  bridge_pile_price_db.py       # NEW — import + get_bridge_pile_price + variant groups
  bridge_pile_line_parser.py    # NEW — C8-35T1 / aliases + grade + qty
  bridge_pile_format_prompt.py  # NEW — OCR/AI
  bridge_pile_text_normalizer.py
  commercial_pricing.py         # EXTEND
  commercial_offer*.py          # EXTEND — pile-like columns
  kp_db_schema.py               # EXTEND — kp_bridge_piles
  kp_persistence_service.py / kp_order_data.py / offers_read.py

scripts/import_bridge_pile_prices_from_xlsx.py

app/
  services/commercial_bridge_pile_service.py
  schemas + endpoints .../bridge-piles + .../ai + .../grades
  archive badge/filter; production whitelist unchanged (plates)

frontend/
  ProductTypePicker + «Мостовые сваи»
  BridgePileInputStep, KpBridgePilePreviewPanel (grade + variant select)
  archive badge/filter/drawer

tests/
  test_bridge_pile_*, test_commercial_bridge_pile_flow.py, …
```

---

## Architecture (sketch)

```
Picker → bridge_piles
  → BridgePileInputStep (text/photo/AI)
  → Preview: Марка | Класс ▼ | Кол-во | Цена
  → Client → Result (pdf+xlsx) → Archive
```

**Parsing examples (MVP):**

```
C8-35T1 2
C8-35T1 B25 2
C8-35В4 25 1
C13-35T4 B30 3
с8-35t1 5
```

**`bridge_pile_prices` (предложение):**

```sql
CREATE TABLE IF NOT EXISTS bridge_pile_prices (
    mark TEXT NOT NULL,
    concrete_grade TEXT NOT NULL,
    price REAL NOT NULL,
    variant_group TEXT,          -- общий id для алиасов одной строки прайса (lookup)
    display_name TEXT,
    price_list_date TEXT,
    imported_at TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (mark, concrete_grade)
);

-- Импорт только с листа «Прайс».
-- Для каждой части до/после ';' — отдельная строка mark + один variant_group;
-- price только для grades с ненулевой ячейкой.
```

**`kp_bridge_piles`:** как `kp_piles` (mark, concrete_grade, qty, unit_price, discounted_price).  
Поле `mark` = как распознано/введено менеджером (после лёгкой нормализации пробелов), не принудительный канон из прайса.

**Alias / lookup:**
- Цена: любая форма из `variant_group` → та же цена (через synonym → canonical key или прямой multi-row import).
- UI/PDF: **без** dropdown «Вариант»; оставляем написание менеджера.

**Grades UX:**
- Список опций = grades с ценой для **найденной** mark (подмножество `{B25, B30}`).
- «Применить класс ко всем» (как у обычных свай): строки, у которых выбранного grade нет в прайсе — **skip + warning** (решение A / Q12). Не молча ставить unpriced класс.

---

## Code Style

Зеркало `pile_*` / `Pile*`; имена `bridge_pile_*` / `BridgePile*`.  
`product_kind: "bridge_pile"`.  
Не рефакторить plates/piles/steps/marches beyond точек расширения.  
Не смешивать с `GRADE_CODES` обычных свай без явного adapter.

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit | parser (алиасы, C/С, B/В, grade 25/30); import 64 rows + variant groups |
| Unit | pricing; unpriced block; available-grades per mark |
| Integration | create → grades/variant → calculate → save |
| Archive / production | badge; excluded from production |
| Regression | plate + pile + step + march flows green |

---

## Boundaries

### Always
- pytest/vitest before merge
- Immutable `product_type` after draft create
- Block calculate/save on unknown mark+grade
- Production = plates only
- Не писать мостовые позиции в `kp_piles` / `pile_prices`

### Ask first
- SQLite schema, new deps, OCR provider changes
- Generic multi-product framework
- Добавление ФБС в тот же PR
- Изменение семантики bulk-grade при недоступном классе

### Never
- Mix product types in one KP
- Treat bridge piles as `product_type=piles`
- Silent overwrite of manager-typed mark with price-list canonical alias
- Fuzzy-match without agreement
- Commit secrets / live DBs
- Import sheets other than «Прайс» as commercial prices

---

## Acceptance Criteria (MVP)

- [x] AC-1 Picker «Мостовые сваи» → `product_type=bridge_piles`
- [x] AC-2 Input text/photo/AI
- [x] AC-3 Parser mark + optional grade + qty; merge mark+grade
- [x] AC-4 Preview: grade dropdown (только доступные для марки); **без** variant picker
- [x] AC-5 CLI import только лист «Прайс»; ≥64 marks; алиасы → synonym groups; нулевые ячейки не как цена
- [x] AC-6 Missing mark+grade → block calculate
- [x] AC-7 Client/result; files pdf+xlsx only
- [x] AC-8 PDF/XLSX: марка **как ввёл менеджер**, класс, кол-во, цена, сумма
- [x] AC-9 Save `kp_bridge_piles` + meta
- [x] AC-10 Archive badge + filter
- [x] AC-11 Plate/pile/step/march regression
- [x] AC-12 Not in production
- [x] AC-13 Shared `kp_id`
- [x] AC-14–17 Mirror piles (drawer, regen, disabled production, merge, grades API)

---

## Not Doing

| Item | Why |
|------|-----|
| ФБС | Следующая итерация |
| Смешение с обычными сваями | Другой каталог |
| Подмена марки на канон прайса в PDF | Решение Q4: оставлять ввод менеджера |
| Variant dropdown | Снят после ревью (одинаковая цена) |
| Generic framework | После 2+ клонов |
| UI-импорт прайса | CLI |
| Производство/СГП | Как у свай |
| B15 / B22.5 / B30_granite в UI | Нет в прайсе мостовых |

---

## Open Questions — статус

| # | Вопрос | Статус |
|---|--------|--------|
| OQ-1 | `product_type` / лейбл | ✅ **`bridge_piles`**, «Мостовые сваи» |
| OQ-2 | Источник / grade 30 | ✅ Только лист **«Прайс»**; колонка 30 → **`B30`** |
| OQ-3 | PDF/UI при алиасах | ✅ **Как вписал менеджер** на шаге ввода |
| OQ-4 | «Применить класс ко всем» на строку без класса | ✅ **A:** skip + warning |
| OQ-5 | Нормализация `С`/`C` | ✅ Lookup insensitive; display = ввод менеджера |

---

## Next

SPECIFY approved → PLAN: `ai_docs/develop/plans/2026-08-05-kp-bridge-piles.md` → TASKS → IMPLEMENT (TDD, clone pile-ветки).  
ФБС — отдельный spec после приёмки.
