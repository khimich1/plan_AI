# Spec: КП на сваи

> **Источник идеи:** [`ai_docs/ideas/kp-piles.md`](../ideas/kp-piles.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → IMPLEMENT ✅  
> **Report:** [`ai_docs/develop/reports/2026-07-31-kp-piles-implementation.md`](../develop/reports/2026-07-31-kp-piles-implementation.md)  
> **Статус:** implemented (MVP)
> **Связанные модули:** `CommercialOfferWizard`, `app/api/v1/endpoints/commercial.py`, `CommercialWorkflowService`, `core/pile_price_db.py`, `core/plate_line_parser.py`, `core/commercial_offer*.py`, `KpPersistenceService`

---

## Decisions locked (Q1–Q16)

| # | Тема | Решение |
|---|------|---------|
| Q1 | Колонки PDF/XLSX | **Минимум:** марка, класс бетона, кол-во, цена, сумма |
| Q2 | Архив | **Бейдж + фильтр** «Все / Плиты / Сваи» — в MVP |
| Q3 | Статус save | **Как у плит:** «В архив» → `в архиве`, «В работе» → `в работе` |
| Q3+ | Производство свай | **Фаза 2** (не MVP); pile КП исключены из production wizard |
| Q4 | Логистика | **Те же поля**, что у плит (доставка, logistics_cost) |
| Q5 | БД позиций | **Отдельная таблица `kp_piles`** + `kp_meta.product_type` |
| Q6 | Fixtures парсера | **Тестовые марки из прайса** (`С120.35-12` и т.п.) |
| Q7 | OCR | **Тот же pipeline**, отдельный system prompt для свай |
| Q8 | Редактирование | **Как у плит:** правка normalized text → пересчёт |
| Q9 | Дубликаты марки | **Разрешены** разные строки с разным классом бетона |
| Q10 | Скидка / НДС | **Как у плит** (скидка %, НДС 22%) |
| Q11 | Spec approved | Переход к PLAN |
| Q12 | Нумерация КП | **Единая сквозная серия** для плит и свай: следующий номер = max(`kp_id`) + 1. Пример: КП на плиты №2 → следующее КП на сваи №3 |
| Q13 | Карточка в архиве | **Полная как у плит**, колонки свай (марка, класс, кол-во, цена); без активного «В производство» |
| Q14 | PDF/XLSX из архива | **Обязательно** в MVP (перескачивание / перегенерация) |
| Q15 | «В производство» | Кнопка **disabled** + подсказка «скоро» (фаза 2) |
| Q16 | Дубликат марка+класс | **Слить в одну строку**, qty суммируется (рекомендация; разный класс — отдельные строки) |

---

## Assumptions I'm Making

1. **Web-only MVP** — только React-мастер КП; Telegram-бот (`bot_archived/`) вне scope.
2. **Отдельное КП** — в одном документе либо плиты, либо сваи; смешанных позиций нет.
3. **Прайс уже загружен** — таблица `pile_prices` в `pb.db` заполнена через `import_pile_price_from_xlsx`; импорт прайса в UI не входит в MVP.
4. **Strict match марок** — цена только при точном совпадении `mark` + `concrete_grade`; fuzzy-match не делаем.
5. **Классы бетона** — фиксированный набор из `pile_price_db.GRADE_CODES`: `B15`, `B20`, `B22_5`, `B25`, `B30_granite`; default при отсутствии в тексте — `B25`.
6. **Клиентский шаг без изменений** — менеджер, скидка, условия доставки/оплаты те же, что у плит.
7. **Без производства в MVP** — статусы save как у плит; wizard производства / СГП для свай — **фаза 2** (после MVP).
8. **Нумерация КП** — **единая сквозная** для плит и свай через общую таблицу `KP_offers` (`kp_id` AUTOINCREMENT). Отображаемый номер после save = `kp_id` (как в архиве: «КП №3»). Отдельные префиксы/счётчики для свай **не** вводим.
9. **SQLite schema change допустима** — добавим `product_type` и таблицу `kp_piles` (см. § Data Model); согласовать перед merge.
10. **Логистика** — те же поля metadata и UI, что у плит; отдельный калькулятор веса свай не делаем в MVP.

→ Если что-то из этого неверно — поправьте до перехода к PLAN.

---

## Objective

Дать менеджеру возможность создать **отдельное коммерческое предложение на цельные сваи** с UX, аналогичным плитам: ввод текста/фото, таблица позиций с классом бетона, расчёт из `pile_prices`, PDF/XLSX и сохранение в архив.

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | нажимаю «Создать КП» | выбрать «Плиты» или «Сваи» | не путать два разных продукта |
| US-2 | создаю КП на сваи | вставить список или загрузить фото таблицы | не набирать позиции в Excel вручную |
| US-3 | вижу распознанный список | указать/исправить класс бетона по каждой свае | цена бралась из правильной колонки прайса |
| US-4 | в заявке опечатка в марке | увидеть ошибку до расчёта | не отправить клиенту КП с неверной ценой |
| US-5 | заполнил клиента и условия | получить PDF/XLSX и сохранить в архив | закрыть сделку как с плитами |
| US-6 | смотрю архив | видеть, что КП — на сваи | не открывать каждый документ |

### Reframed success criteria

| Требование | Измеримый критерий |
|------------|-------------------|
| «Быстрее Excel» | Полный цикл text → PDF ≤ 5 мин для типовой заявки (≤20 позиций) без ручного поиска в прайсе |
| «100% цены из прайса» | Каждая позиция с `unit_price = get_pile_price(mark, grade)`; расчёт блокируется при `PriceNotFoundError` |
| «Тот же UX» | 3 шага мастера, те же паттерны: textarea, file upload, Ctrl+V, «Распознать», «Обработать», client, result, save |
| «Класс per-line» | Класс парсится из текста (`B25`, `22.5`, `15`…); в таблице — dropdown + «Применить ко всем» |
| «Без производства» | КП с `product_type=piles` не появляется в wizard производства / СГП |

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2, SQLite (`pb.db`, `plita.db`) |
| Domain | `core/pile_line_parser.py` (новый), `core/pile_price_db.py`, `core/commercial_pricing.py` (расширение) |
| API | `app/api/v1/endpoints/commercial.py` |
| Services | `CommercialWorkflowService`, `CommercialCalculationService`, `CommercialWizardStepService` |
| Frontend | React 19, Vite, TypeScript, TanStack Query, Vitest |
| OCR/ИИ | Существующий pipeline (`prepare_commercial_ocr_upload`, OpenAI/GigaChat) с pile-промптом |
| Tests | pytest (`tests/`), Vitest (`frontend/src/features/commercial-offer/`) |

---

## Commands

```bash
# Backend (из корня репозитория)
source .venv/bin/activate
uvicorn app.main:app --reload

# Тесты — backend (релевантные наборы)
pytest tests/test_pile_price_import.py -q
pytest tests/test_commercial_web_flow.py -q
pytest tests/test_commercial_wizard_step_service.py -q
pytest tests/ -k "pile or commercial" -q

# Frontend (из frontend/)
npm run dev
npm run typecheck
npm run test
npm run build
```

---

## Project Structure

```
core/
  pile_price_db.py          # уже есть — lookup цен
  pile_line_parser.py       # NEW — разбор строк свай
  pile_format_prompt.py     # NEW — OCR/ИИ system prompt
  pile_text_normalizer.py   # NEW — нормализация списка (опционально, по аналогии с plate_text_normalizer)
  commercial_pricing.py     # EXTEND — pile pricing path
  commercial_offer.py       # EXTEND — PDF для свай
  commercial_offer_xlsx.py  # EXTEND — XLSX для свай
  kp_db_schema.py           # EXTEND — kp_piles, kp_meta.product_type
  kp_persistence_service.py # EXTEND — ветка сохранения свай

app/
  api/v1/endpoints/commercial.py      # EXTEND — /piles endpoints, product_type при create
  schemas/commercial.py               # EXTEND — ProductType, pile batch, wizard actions
  services/commercial_workflow_service.py
  services/commercial_calculation_service.py
  services/commercial_wizard_step_service.py

frontend/src/features/commercial-offer/
  components/
    ProductTypePicker.tsx             # NEW — шаг 0
    steps/PileInputStep.tsx           # NEW
    KpPilePreviewPanel.tsx            # NEW — таблица с grade dropdown
    CommercialOfferWizard.tsx         # EXTEND — ветвление по product_type
  api/commercialOfferApi.ts           # EXTEND — piles mutations
  hooks/useCommercialOfferWizard.ts   # EXTEND
  types/commercialOffer.ts            # EXTEND
  lib/wizardStepOrder.ts              # EXTEND — piles first step

tests/
  test_pile_line_parser.py            # NEW
  test_commercial_pile_flow.py        # NEW — e2e API flow
  test_commercial_pile_pricing.py     # NEW

ai_docs/
  ideas/kp-piles.md                   # ideation one-pager
  specs/kp-piles.md                   # этот документ
  develop/plans/2026-07-30-kp-piles.md  # PLAN (следующая фаза)
```

---

## Architecture

### Flow

```
/new → ProductTypePicker (plates | piles)
         ↓
   piles: PileInputStep (text | image | ai)
         ↓ normalize + price lookup per line
   KpPilePreviewPanel (mark, grade▼, qty, unit_price)
         ↓ [Обработать] — все марки в прайсе
   ClientConditionsStep (reuse)
         ↓
   CalculationResultStep (reuse, без schema/breakdown плит)
         ↓
   Save → archive (badge «Сваи»)
```

### Wizard step model

| product_type | Step 1 id | Step 2 | Step 3 |
|--------------|-----------|--------|--------|
| `plates` | `plates` | `client` | `result` |
| `piles` | `piles` | `client` | `result` |

- Расширить `WizardStepId` enum: добавить `piles`.
- Расширить `WizardNextRequiredAction`: `ingest_piles`, убрать `resolve_wide_plates` для pile drafts.
- `CommercialWizardStepService`: для `product_type=piles` не вызывать wide-plates gate.

### API (новые / изменённые endpoints)

| Method | Path | Назначение |
|--------|------|------------|
| `POST` | `/commercial/drafts` | + form field `product_type` (`plates` \| `piles`, default `plates`) |
| `PATCH` | `/commercial/drafts/{id}/piles` | append/replace pile lines (text/image) |
| `POST` | `/commercial/drafts/{id}/piles/ai` | ИИ-обработка списка свай |
| `PATCH` | `/commercial/drafts/{id}/piles/grades` | bulk update `concrete_grade` по позициям (опционально, если не через meta normalized edit) |
| existing | `.../meta`, `.../calculate`, `.../generate-files`, `.../save` | ветвление внутри сервисов по `product_type` |

Контракт create/update — зеркало plates: multipart (`text`, `image`, `mode`, `product_type`).

### Data model

**Draft metadata** (`CommercialDraftMetadata`):

```python
product_type: Literal["plates", "piles"] = "plates"
pile_batches: list[CommercialPileBatch] = []  # аналог plate_batches
default_concrete_grade: str = "B25"
```

**order_data item (pile):**

```python
{
    "product_kind": "pile",
    "name": "С120.35-12",           # display / file column
    "mark": "С120.35-12",           # ключ pile_prices
    "concrete_grade": "B25",
    "qty": 5,
    "unit_price": 44634.03,
    "line_total": 223170.15,        # optional, computed
}
```

**Нумерация (Q12):**

| Этап | Номер | Пример |
|------|-------|--------|
| Черновик (до save) | Временный `WEB_{draft_id…}` | `WEB_01ABCD12` — только в wizard |
| После save | **`kp_id`** из `KP_offers` | Плиты №1, №2 → сваи №3 |

- Плиты и сваи пишутся в **одну** таблицу `KP_offers` → один AUTOINCREMENT.
- PDF/XLSX и архив после save: `offer_number = str(kp_id)` (как сейчас у плит в `archive_service`).
- **Запрещено:** отдельный счётчик, префикс «С-», второй INSERT-path с другой нумерацией.

**Persistence (новое — Ask first):**

```sql
-- kp_meta: product_type для архива и фильтрации производства
ALTER TABLE kp_meta ADD COLUMN product_type TEXT DEFAULT 'plates';

-- kp_piles: позиции КП на сваи (не смешивать с kp_plates)
CREATE TABLE IF NOT EXISTS kp_piles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kp_id INTEGER NOT NULL,
    position_number INTEGER NOT NULL,
    mark TEXT NOT NULL,
    concrete_grade TEXT NOT NULL,
    qty INTEGER NOT NULL,
    unit_price REAL NOT NULL,
    discounted_price REAL NOT NULL,
    FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
);
```

КП на сваи сохраняется в `KP_offers` + `kp_piles` + `kp_meta(product_type='piles', status='в архиве'|'в работе')`. Таблица `kp_plates` для pile КП **не используется**.

### Parsing (`core/pile_line_parser.py`)

Поддерживаемые форматы строк (MVP):

```
С120.35-12 5
С120.35-12 B25 5
С120.35-12 25 5          # grade как число → B25
С120.35-13и B30 2
с 120.35-12 b25 5        # нормализация регистра/пробелов
```

`LineParseResult`: `{ parsed, mark, concrete_grade | None, qty, reason_code }`.

Grade detection — reuse логики `_grade_code_from_header` из `pile_price_db.py` (вынести в shared helper при необходимости).

### Pricing

```python
from core.pile_price_db import get_pile_price, GRADE_CODES

def lookup_pile_price(mark: str, concrete_grade: str, *, db_path: str) -> float:
    price = get_pile_price(mark, concrete_grade, db_path)
    if price is None:
        raise PriceNotFoundError(f"Свая не найдена в прайсе: {mark}, {concrete_grade}")
    return price
```

`ensure_order_priced` / `collect_unpriced_positions` — ветка по `product_kind`.

### File generation

- **PDF/XLSX:** тот же фирменный макет (`commercial_offer.py`, `commercial_offer_xlsx.py`), таблица позиций:
  - № | Наименование (марка) | Класс бетона | Кол-во, шт | Цена за шт | Сумма
- **Не генерировать** для piles: `breakdown`, `schema` (раскладка плит).
- `generate-files` API: фильтровать `kind` по `product_type`.

### Frontend UX details

**ProductTypePicker:**
- Две карточки: «Плиты» / «Сваи»; выбор сохраняется в local wizard store до reset.
- После выбора — монтируется `CommercialOfferWizard` с `productType` prop.

**PileInputStep** — копия паттернов `PlateInputStep`:
- Textarea placeholder: `С120.35-12 B25 5\nС120.35-13и 3`
- File upload + Ctrl+V
- Фазы A/B: «Распознать» → preview → «Обработать»
- Без `WidePlatesInlineSection`

**KpPilePreviewPanel:**
- Колонки: Марка | Класс ▼ | Кол-во | Цена | Сумма
- Toolbar: «Применить класс ко всем: [B25 ▼]»
- Строки без цены — highlight error, блок «Обработать» / переход на client

**Archive:**
- Badge «Сваи» / «Плиты» рядом с номером КП
- Фильтр **«Все / Плиты / Сваи»** в MVP (AC-10)

---

## Code Style

Следовать существующим паттернам плит: dataclass для parse result, Pydantic v2 enums, сервисная оркестрация через `CommercialWorkflowService`.

```python
# core/pile_line_parser.py — образец стиля
@dataclass
class PileLineParseResult:
    parsed: bool
    mark: str = ""
    concrete_grade: str | None = None
    qty: int = 1
    reason_code: str = ""
    reason_text: str = ""


def parse_pile_line(line: str, *, default_grade: str = "B25") -> PileLineParseResult:
    ...
```

```typescript
// types/commercialOffer.ts
export type ProductType = "plates" | "piles";
export type WizardStepId = "plates" | "piles" | "client" | "result";

export type PileOrderLine = {
  mark: string;
  name: string;
  concrete_grade: string;
  qty: number;
  unit_price: number | null;
};
```

- Имена файлов: `pile_*` в `core/`, `Pile*` компоненты во frontend.
- Не импортировать `app` из `core/` (граница проекта).
- Минимальный diff: не рефакторить plate flow без необходимости.

---

## Testing Strategy

| Уровень | Что покрываем | Где |
|---------|---------------|-----|
| Unit | `parse_pile_line`, grade extraction, qty edge cases | `tests/test_pile_line_parser.py` |
| Unit | pile pricing, unpriced blocking | `tests/test_commercial_pile_pricing.py` |
| Integration | API: create pile draft → update → calculate → save | `tests/test_commercial_pile_flow.py` |
| Integration | wizard step service для piles (no wide-plates) | extend `tests/test_commercial_wizard_step_service.py` |
| Frontend unit | grade bulk apply, wizard step order | `frontend/src/features/commercial-offer/**/*.test.ts(x)` |
| Regression | plate flow не сломан | `tests/test_commercial_web_flow.py` (existing) |

**Coverage expectation:** все acceptance criteria ниже имеют хотя бы один автоматический тест или явный manual verify step в TASKS.

**Manual smoke (перед merge):**
1. Создать КП на сваи из текста с B25.
2. Создать из фото (если OCR env настроен).
3. Сменить класс одной позиции — цена обновилась.
4. Несуществующая марка — расчёт заблокирован.
5. PDF/XLSX скачиваются; архив показывает «Сваи».

---

## Boundaries

### Always do
- Запускать релевантные pytest/vitest перед merge
- Валидировать `product_type` на create draft; нельзя сменить тип после создания черновика
- Блокировать calculate/save при неизвестной марке или grade без цены
- Хранить `concrete_grade` в order_data и в `kp_piles` при save
- Фильтровать pile КП из производственного wizard (`product_type != 'piles'`)

### Ask first
- Изменения схемы SQLite (`kp_piles`, `kp_meta.product_type`)
- Новые npm/pip dependencies
- Изменения OCR provider config / rate limits
- Общий рефакторинг `CommercialWorkflowService` beyond pile branch

### Never do
- Смешивать плиты и сваи в одном КП (MVP)
- Писать pile lines в `kp_plates`
- Fuzzy-match марок без явного согласования
- Подключать pile КП к производству / СГП в этом MVP
- Коммитить секреты (`.env`, API keys)
- Удалять/ослаблять failing plate tests

---

## Acceptance Criteria (MVP)

- [x] **AC-1** Экран выбора «Плиты / Сваи» перед мастером; выбор фиксируется в draft metadata
- [x] **AC-2** `PileInputStep`: textarea, upload, paste, recognize (text + image), AI instruction
- [x] **AC-3** Парсер извлекает mark, qty; grade из текста если есть, иначе default `B25`
- [x] **AC-4** Preview-таблица: mark, grade dropdown (`GRADE_CODES`), qty, unit_price, line total
- [x] **AC-5** «Применить класс ко всем» обновляет все строки и пересчитывает цены
- [x] **AC-6** Неизвестная mark+grade → ошибка в UI + `validation_errors`; calculate возвращает 4xx
- [x] **AC-7** Шаг client/result работает без wide-plates gate
- [x] **AC-8** PDF и XLSX генерируются с колонками: марка, класс, кол-во, цена, сумма
- [x] **AC-9** Save в архив создаёт `KP_offers` + `kp_piles` + `kp_meta(product_type=piles)`
- [x] **AC-10** Архив: badge «Сваи» + фильтр «Все / Плиты / Сваи»
- [x] **AC-11** Существующий plate flow (create → save) проходит тесты без регрессий
- [x] **AC-12** Pile КП не отображается в wizard производства / кандидатах плана
- [x] **AC-13** После save КП на сваи получает **следующий** `kp_id` в общей серии
- [x] **AC-14** Карточка архива для свай: полный drawer, колонки марка/класс/кол-во/цена
- [x] **AC-15** PDF/XLSX скачиваются из архива для pile КП (перегенерация через `archive_service`)
- [x] **AC-16** «В производство» для свай: кнопка disabled + tooltip «скоро»
- [x] **AC-17** При ingest одинаковые mark+grade сливаются (qty суммируется)

---

## Not Doing (MVP)

| Item | Reason |
|------|--------|
| Смешанное КП (плиты + сваи) | Продуктовое решение ideation |
| Fuzzy-match марок | Риск неверной цены |
| Производство, СГП, оптимизация | Out of scope |
| Составные / шнековые сваи | Нет в текущем `pile_prices` |
| Excel-импорт заявки | Text + photo priority |
| Отдельный счётчик / префикс для свай | Q12: одна серия `kp_id`; интеграция с 1С — позже |
| Telegram-бот | Archived bot out of scope |
| Производство свай | Фаза 2; в MVP pile КП исключены из production wizard |
| Калькулятор веса / доставки для свай | Нет данных; логистика as-is или 0 |

---

## Success Criteria (Definition of Done)

Спека считается реализованной, когда:

1. Все AC-1…AC-17 выполнены и покрыты тестами или manual checklist.
2. `pytest tests/ -k "pile or commercial"` и `npm run test && npm run build` — green.
3. Plate regression (`test_commercial_web_flow.py`) — green.
4. Spec/plan/tasks документы обновлены; решения по Q1–Q5 зафиксированы.

---

## Next Step (SDD Phase 3)

Plan: [`ai_docs/develop/plans/2026-07-30-kp-piles.md`](../develop/plans/2026-07-30-kp-piles.md) — после ревью плана → TASKS → IMPLEMENT.
