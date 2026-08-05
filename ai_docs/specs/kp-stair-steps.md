# Spec: КП на лестничные ступени (ЛС)

> **Источник идеи:** [`ai_docs/ideas/kp-stair-steps.md`](../ideas/kp-stair-steps.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → IMPLEMENT ✅  
> **Статус:** implemented (2026-08-05)  
> **Аналог:** [`ai_docs/specs/kp-piles.md`](kp-piles.md) (`product_type=piles`)  
> **Прайс:** `банк знаний/Прайс на лестничные ступени от 03.08.2026.xlsx`, лист «Прайс» (42 SKU)  
> **Связанные модули:** `CommercialOfferWizard`, `app/api/v1/endpoints/commercial.py`, `CommercialWorkflowService`, `core/commercial_offer*.py`, `KpPersistenceService`, pile-ветка как шаблон

---

## Decisions locked (из ideation + ниже)

| # | Тема | Решение |
|---|------|---------|
| Q1 | Продукт MVP | Только **лестничные ступени (ЛС)**; ЛМ / ФБС — следующие релизы |
| Q2 | UX | **Clone свай** минус класс бетона |
| Q3 | `product_type` | **`steps`** (внутренний код API/БД; UI-лейбл «Ступени») |
| Q18b | PDF наименование | Только короткая марка `ЛС…` (без префикса «Лестничные ступени») |
| Q6b | OCR | Полный clone свай в том же MVP PR |
| Q10b | Production filter | Whitelist: `COALESCE(product_type,'plates') = 'plates'` |
| Q4 | Цена | Одна цена на марку: `get_step_price(mark)` |
| Q5 | UI бетона | **Нет** dropdown / bulk grade / поле `concrete_grade` |
| Q6 | Ввод | Текст + фото + ИИ (как сваи) |
| Q7 | КП scope | Одно КП = один тип; смешанных позиций нет |
| Q8 | Архив | Бейдж «Ступени» + фильтр «Все / Плиты / Сваи / Ступени» |
| Q9 | Save / статусы | Как у плит/свай: «В архив» / «В работе» |
| Q10 | Производство | **Не MVP**; step КП исключены из production / СГП; кнопка disabled «скоро» |
| Q11 | Логистика | Те же поля metadata/UI; без калькулятора веса/объёма |
| Q12 | БД позиций | Отдельная таблица `kp_steps` + `kp_meta.product_type='steps'` |
| Q13 | Нумерация | Единая серия `kp_id` |
| Q14 | Дубликаты марок | Одинаковая `mark` → слить, qty суммируется |
| Q15 | Неизвестная марка | Strict match; блокировка calculate/save |
| Q16 | Импорт прайса | CLI (как сваи); UI-импорт — out |
| Q17 | PDF/XLSX из архива | Обязательно в MVP (реген) |
| Q18 | Колонки PDF/XLSX | № \| Наименование (`ЛС…`) \| Кол-во \| Цена \| Сумма (**без** класса бетона) |
| Q19 | Скидка / НДС | Как у плит (скидка %, НДС 22%) |
| Q20 | Редактирование | Правка normalized text → пересчёт (как плиты/сваи) |

---

## Assumptions I'm Making

1. **Web-only** — React-мастер КП; Telegram-бот вне scope.
2. **`product_type = "steps"`** — короткий стабильный код; UI-лейбл «Ступени» / «Лестничные ступени».
3. **Ключ цены = короткая марка** (`ЛС11`, `ЛС14-1лев`, `ЛС11-Б-1`) после нормализации регистра и пробелов. Полное имя из Excel (`Лестничные ступени ЛС11`) храним опционально как `display_name` при импорте, но lookup — по короткой марке.
4. **Импорт листа «Прайс»:** колонки «Наименование» + одна числовая цена (заголовок «15» игнорируем как grade-matrix). Из наименования извлекаем `ЛС…` regex; лишние пробелы в полном имени (`ЛС14-Б`) не ломают ключ.
5. **42 SKU** текущего прайса — полный каталог MVP (семейства ЛС11/12/14/15/18/22 с суффиксами; у ЛС22 в прайсе нет `-Б*`).
6. **OCR в MVP** — полный clone свай (отдельный system prompt). Если OCR env не настроен — text path обязан работать.
7. **Нет `concrete_grade`** в `order_data`, `kp_steps`, API grades-эндпоинтов.
8. **Client/result без изменений** — менеджер, скидка, доставка/оплата как у плит.
9. **SQLite schema change допустима** — `kp_steps` + расширение enum/`Literal` `product_type`.
10. **Не делаем generic product framework** в этом релизе — точечный clone pile-ветки.
11. **Прайс загружен до demo** через CLI; пустой `step_prices` → все позиции unpriced.

→ Если что-то неверно — поправьте **до** перехода к PLAN.

---

## Objective

Дать менеджеру **отдельное КП на лестничные ступени (ЛС)** с UX как у свай: ввод текста/фото, таблица марка+кол-во+цена из `step_prices`, PDF/XLSX, архив — **без** выбора класса бетона.

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | создаю КП | выбрать «Ступени» наряду с плитами/сваями | не путать продукты |
| US-2 | есть заявка списком | вставить текст или фото | не искать цены в Excel |
| US-3 | вижу preview | видеть марку, qty, цену **без** класса бетона | не тратить время на лишний UI |
| US-4 | опечатка в марке | ошибка до расчёта | не отдать клиенту неверную цену |
| US-5 | заполнил клиента | PDF/XLSX + архив | закрыть сделку как со сваями |
| US-6 | смотрю архив | бейдж «Ступени» + фильтр | отличать тип КП |

### Reframed success criteria

| Требование | Измеримый критерий |
|------------|-------------------|
| «Как сваи» | Те же 3 шага мастера; text/photo/AI; calculate → PDF/XLSX → save |
| «100% цены из прайса» | `unit_price = get_step_price(mark)`; `PriceNotFoundError` блокирует calculate |
| «Без бетона» | Нет grade в UI, API, PDF, `kp_steps` |
| «Без производства» | `product_type=steps` исключён из production wizard / СГП |
| «Только ЛС» | ЛМ/ФБС не в picker и не в прайсе этого MVP |

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2, SQLite (`pb.db`, `plita.db`) |
| Domain | NEW `core/step_*`; EXTEND commercial pricing/offer/persistence по образцу piles |
| API | `app/api/v1/endpoints/commercial.py` |
| Frontend | React 19, Vite, TypeScript, TanStack Query, Vitest |
| OCR/ИИ | Существующий pipeline + step system prompt |
| Tests | pytest (`tests/`), Vitest (`frontend/src/features/commercial-offer/`) |

---

## Commands

```bash
# Backend
source venv/bin/activate   # or .venv
uvicorn app.main:app --reload

# Import price list (после реализации)
python scripts/import_step_prices_from_xlsx.py \
  "банк знаний/Прайс на лестничные ступени от 03.08.2026.xlsx" \
  --sheet Прайс

# Tests
pytest tests/ -k "step or stair or commercial" -q
pytest tests/test_commercial_pile_flow.py tests/test_commercial_web_flow.py -q

# Frontend
cd frontend && npm run typecheck && npm run test && npm run build
```

---

## Project Structure

```
core/
  step_price_db.py            # NEW — step_prices CRUD + xlsx import
  step_line_parser.py         # NEW — parse ЛС… qty
  step_format_prompt.py       # NEW — OCR/AI system prompt
  step_text_normalizer.py     # NEW — optional, mirror pile_text_normalizer
  commercial_pricing.py       # EXTEND — lookup_step_price / is_step_order
  commercial_offer.py         # EXTEND — PDF branch steps (no grade col)
  commercial_offer_xlsx.py    # EXTEND — XLSX branch
  kp_db_schema.py             # EXTEND — kp_steps
  kp_persistence_service.py   # EXTEND — save steps
  kp/offers_read.py           # EXTEND — load kp_steps
  kp_order_data.py            # EXTEND — order_data_from_kp_steps

scripts/
  import_step_prices_from_xlsx.py   # NEW

app/
  api/v1/endpoints/commercial.py    # EXTEND — /steps endpoints
  schemas/commercial.py             # EXTEND — ProductType += steps
  services/commercial_workflow_service.py
  services/commercial_step_service.py   # NEW — mirror CommercialPileService
  services/commercial_wizard_step_service.py
  services/commercial_export_service.py
  services/archive_service.py
  repositories/kp_repository.py     # EXTEND — exclude steps from production

frontend/src/features/commercial-offer/
  components/ProductTypePicker.tsx      # EXTEND — + Ступени
  components/steps/StepInputStep.tsx    # NEW
  components/KpStepPreviewPanel.tsx     # NEW — no grade
  components/CommercialOfferWizard.tsx  # EXTEND
  api/commercialOfferApi.ts             # EXTEND
  types/commercialOffer.ts              # EXTEND
  lib/wizardStepOrder.ts                # EXTEND

frontend/src/features/commercial-archive/  # EXTEND badge + filter + drawer

tests/
  test_step_line_parser.py
  test_step_price_import.py
  test_commercial_step_flow.py
  test_commercial_step_pricing.py
  test_kp_steps_schema.py
  test_kp_persistence_steps.py
  test_production_step_exclusion.py
  # + archive / wizard step service extensions

ai_docs/
  ideas/kp-stair-steps.md
  specs/kp-stair-steps.md          # этот документ
  develop/plans/…                  # после approve SPECIFY
```

---

## Architecture

### Flow

```
/new → ProductTypePicker (plates | piles | steps)
         ↓
   steps: StepInputStep (text | image | ai)
         ↓ normalize + get_step_price(mark)
   KpStepPreviewPanel (mark, qty, unit_price)   # NO grade
         ↓
   ClientConditionsStep (reuse)
         ↓
   CalculationResultStep (pdf+xlsx only; no breakdown/schema)
         ↓
   Save → KP_offers + kp_steps + kp_meta(product_type=steps)
         ↓
   Archive badge «Ступени»
```

### Wizard step model

| product_type | Step 1 | Step 2 | Step 3 |
|--------------|--------|--------|--------|
| `plates` | `plates` | `client` | `result` |
| `piles` | `piles` | `client` | `result` |
| `steps` | `steps` | `client` | `result` |

- `WizardStepId` += `steps`
- `WizardNextRequiredAction` += `ingest_steps`
- Для `steps`: без wide-plates gate; без pile grades endpoints

### API

| Method | Path | Назначение |
|--------|------|------------|
| `POST` | `/commercial/drafts` | `product_type`: `plates` \| `piles` \| `steps` |
| `PATCH` | `/commercial/drafts/{id}/steps` | append/replace step lines (text/image) |
| `POST` | `/commercial/drafts/{id}/steps/ai` | ИИ-правка списка |
| existing | `meta`, `calculate`, `generate-files`, `save` | ветка по `product_type` |
| **нет** | `.../steps/grades` | grade UI отсутствует |

Create/update: multipart как у piles (`text`, `image`, `mode`, `product_type`).  
`product_type` **нельзя** сменить после create draft.

### Data model

**Draft metadata:**

```python
product_type: Literal["plates", "piles", "steps"] = "plates"
step_batches: list[CommercialStepBatch] = []
# NO default_concrete_grade for steps
```

**order_data item (step):**

```python
{
    "product_kind": "step",
    "name": "ЛС14-1лев",       # display / PDF
    "mark": "ЛС14-1лев",       # ключ step_prices
    "qty": 10,
    "unit_price": 1815.59,
    "line_total": 18155.90,    # optional, computed
}
```

**`step_prices` (pb.db):**

```sql
CREATE TABLE IF NOT EXISTS step_prices (
    mark TEXT NOT NULL PRIMARY KEY,
    price REAL NOT NULL,
    display_name TEXT,
    price_list_date TEXT,
    imported_at TEXT DEFAULT (datetime('now'))
);
```

**`kp_steps` (plita.db):**

```sql
CREATE TABLE IF NOT EXISTS kp_steps (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kp_id INTEGER NOT NULL,
    position_number INTEGER NOT NULL,
    mark TEXT NOT NULL,
    qty INTEGER NOT NULL,
    unit_price REAL NOT NULL,
    discounted_price REAL NOT NULL,
    FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_kp_id_steps ON kp_steps(kp_id);
```

`kp_meta.product_type` уже существует (TEXT); допустимые значения расширяются до `steps`.  
КП steps: `KP_offers` + `kp_steps` + `kp_meta(product_type='steps')`. **Не** писать в `kp_plates` / `kp_piles`.

### Parsing (`core/step_line_parser.py`)

Поддерживаемые форматы (MVP):

```
ЛС11 10
ЛС14-1лев 5
ЛС11-Б-1 2
лс12-2лев 3          # → upper/normalize
Лестничные ступени ЛС15-1 4   # извлечь ЛС15-1
```

Не парсить как grade числа бетона (`B15` / `15`) — qty только явное целое в конце / отдельным токеном.  
`StepLineParseResult`: `{ parsed, mark, qty, reason_code, reason_text }`.

Дубликаты: одинаковая `mark` → merge qty (Q14).

### Pricing

```python
def lookup_step_price(mark: str, *, db_path: str) -> float:
    price = get_step_price(normalize_step_mark(mark), db_path)
    if price is None:
        raise PriceNotFoundError(f"Ступень не найдена в прайсе: {mark}")
    return price
```

`ensure_order_priced` / `collect_unpriced_positions` — ветка `product_kind == "step"`.

### File generation

- PDF/XLSX: фирменный макет; колонки **без** класса бетона; наименование = короткая марка `ЛС…` (без «Лестничные ступени»).
- Не генерировать `breakdown` / `schema`.
- `generate-files` / archive regen: `order_data_from_kp_steps`.

### Frontend UX

**ProductTypePicker:** три карточки — Плиты / Сваи / Ступени.

**StepInputStep:** паттерн `PileInputStep`; placeholder:

```
ЛС11 10
ЛС14-1лев 5
ЛС11-Б-1 2
```

**KpStepPreviewPanel:** Марка | Кол-во | Цена | Сумма; error highlight без цены; **нет** grade toolbar.

**Archive:** badge «Ступени»; фильтр + drawer колонки без бетона; «В производство» disabled.

### Production exclusion

Расширить фильтры, где сейчас `!= 'piles'`, на исключение и `steps` (и будущих non-plate типов — минимально: `NOT IN ('piles','steps')` или `= 'plates'` only для production candidates — **предпочтительно whitelist plates**, чтобы ЛМ/ФБС не просочились случайно).

Рекомендация спеки: production candidates = `COALESCE(product_type,'plates') = 'plates'`.

---

## Code Style

Зеркало pile-паттернов; имена `step_*` / `Step*`.

```python
@dataclass
class StepLineParseResult:
    parsed: bool
    mark: str = ""
    qty: int = 1
    reason_code: str = ""
    reason_text: str = ""


def parse_step_line(line: str) -> StepLineParseResult:
    ...
```

```typescript
export type ProductType = "plates" | "piles" | "steps";
export type WizardStepId = "plates" | "piles" | "steps" | "client" | "result";

export type StepOrderLine = {
  mark: string;
  name: string;
  qty: number;
  unit_price: number | null;
};
```

- Не импортировать `app` из `core/`.
- Минимальный diff: не рефакторить plates/piles beyond необходимых точек расширения.

---

## Testing Strategy

| Уровень | Что | Где |
|---------|-----|-----|
| Unit | parse_step_line, mark extract, merge | `tests/test_step_line_parser.py` |
| Unit | import xlsx → 42 marks; lookup | `tests/test_step_price_import.py` |
| Unit | pricing / unpriced block | `tests/test_commercial_step_pricing.py` |
| Integration | create → update → calculate → save | `tests/test_commercial_step_flow.py` |
| Schema/persist | `kp_steps`, product_type | `tests/test_kp_steps_schema.py`, `test_kp_persistence_steps.py` |
| Production | steps excluded | `tests/test_production_step_exclusion.py` |
| Frontend | step order, preview without grade | `frontend/.../*.test.ts(x)` |
| Regression | plates + piles flows green | existing commercial/pile tests |

**Manual smoke:**
1. Импорт прайса → 42 марки в `step_prices`.
2. КП из текста `ЛС11 2` → цена ≈ 1409.91.
3. Неизвестная марка → calculate blocked.
4. PDF/XLSX без колонки бетона; архив «Ступени».
5. Steps КП нет в production wizard.
6. Plate и pile regression smoke.

---

## Boundaries

### Always do
- Релевантные pytest/vitest перед merge
- Валидировать `product_type` на create; immutable после create
- Блокировать calculate/save при неизвестной марке
- Исключать non-plates из production (`product_type = 'plates'` whitelist)
- Писать step lines только в `kp_steps`

### Ask first
- Изменения схемы SQLite
- Новые зависимости
- OCR provider / rate limits
- Generic multi-product refactor
- Добавление ЛМ/ФБС в тот же PR

### Never do
- Смешивать типы в одном КП
- Писать steps в `kp_plates` / `kp_piles`
- UI/API класса бетона для steps
- Fuzzy-match без согласования
- Подключать steps к производству/СГП в MVP
- Коммитить секреты / live DB
- Ломать plate/pile tests ради steps

---

## Acceptance Criteria (MVP)

- [x] **AC-1** Picker: Плиты / Сваи / Ступени; выбор → `metadata.product_type=steps`
- [x] **AC-2** `StepInputStep`: textarea, upload, paste, recognize, AI
- [x] **AC-3** Парсер: mark + qty; извлекает `ЛС…` из полного наименования
- [x] **AC-4** Preview: mark, qty, unit_price, sum — **без** grade UI
- [x] **AC-5** CLI import листа «Прайс» → ≥42 уникальных `mark` в `step_prices`
- [x] **AC-6** Неизвестная mark → UI error + calculate 4xx
- [x] **AC-7** Client/result без wide-plates; files = pdf+xlsx only
- [x] **AC-8** PDF/XLSX: марка, кол-во, цена, сумма (без бетона)
- [x] **AC-9** Save: `KP_offers` + `kp_steps` + `kp_meta(product_type=steps)`
- [x] **AC-10** Архив: badge «Ступени» + фильтр со steps
- [x] **AC-11** Plate и pile flows без регрессий
- [x] **AC-12** Steps КП не в production wizard / plan candidates
- [x] **AC-13** Общая серия `kp_id`
- [x] **AC-14** Archive drawer: колонки mark/qty/price
- [x] **AC-15** PDF/XLSX regen из архива
- [x] **AC-16** «В производство» disabled + «скоро»
- [x] **AC-17** Дубликат mark → merge qty
- [x] **AC-18** Нет endpoint/UI `steps/grades`

---

## Not Doing (MVP)

| Item | Reason |
|------|--------|
| ЛМ, ФБС | Следующие итерации |
| Класс бетона UI/API | Одна цена в прайсе |
| Смешанное КП | Продуктовое решение |
| Fuzzy-match | Риск цены |
| Производство / СГП | Как у свай |
| Generic product framework | Premature |
| UI-импорт прайса | CLI достаточно |
| Калькулятор веса/объёма | Нет требований |
| Отдельный счётчик КП | Одна серия `kp_id` |
| Telegram-бот | Archived |

---

## Open Questions — resolved 2026-08-05

| # | Вопрос | Решение |
|---|--------|---------|
| OQ-1 | Код `product_type` | **`steps`** — внутренний id в API/БД; на UI «Ступени». (Это не название для клиента.) |
| OQ-2 | PDF/XLSX наименование | Короткая марка **`ЛС…`** без префикса «Лестничные ступени» |
| OQ-3 | OCR | **Как у свай** — полный pipeline + step prompt в том же MVP |
| OQ-4 | Production | Whitelist **`COALESCE(product_type,'plates') = 'plates'`** |

---

## Success Criteria (Definition of Done)

1. AC-1…AC-18 выполнены и покрыты тестами / manual checklist.  
2. `pytest` step+commercial и pile/plate regression — green.  
3. `npm run typecheck && npm run test && npm run build` — green.  
4. Прайс ЛС импортирован в dev `pb.db`; smoke text → PDF пройден.  
5. Spec/plan/tasks обновлены после IMPLEMENT.

---

## Next Step (SDD)

PLAN: [`ai_docs/develop/plans/2026-08-05-kp-stair-steps.md`](../develop/plans/2026-08-05-kp-stair-steps.md) → после ревью плана → TASKS → IMPLEMENT.
