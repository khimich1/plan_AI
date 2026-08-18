# Spec: КП — несколько наименований (append loop)

> **Источник идеи:** [`ai_docs/ideas/kp-multi-nomenclature-append.md`](../ideas/kp-multi-nomenclature-append.md)  
> **План:** [`ai_docs/develop/plans/2026-08-12-kp-multi-nomenclature-append.md`](../develop/plans/2026-08-12-kp-multi-nomenclature-append.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → IMPLEMENT ✅  
> **Статус:** Implemented 2026-08-12 (`orch-2026-08-12-14-05-kp-multi-append`)  
> **Report:** [`ai_docs/develop/reports/2026-08-12-kp-multi-nomenclature-append-implementation.md`](../develop/reports/2026-08-12-kp-multi-nomenclature-append-implementation.md)  
> **Связанные модули:** `CommercialOfferWizard`, `CalculationResultStep`, `ProductTypePicker`, draft/metadata, archive drawer, `commercial_pricing`, `commercial_offer` / `_xlsx`, `KpPersistenceService`, offers read/write  
> **Дата:** 2026-08-12  
> **Orchestration:** `orch-2026-08-12-14-05-kp-multi-append`  
> **Handoff:** [`ai_docs/develop/handoffs/2026-08-12-kp-multi-nomenclature-append.md`](../develop/handoffs/2026-08-12-kp-multi-nomenclature-append.md)

---

## Assumptions I'm Making

1. **Web-only** — React-мастер КП; `bot_archived` мёртвый код, вне scope. ✅ approved 2026-08-12
2. **Один draft / один `kp_id`** после save — не склейка нескольких КП. ✅ approved 2026-08-12
3. **Каждая строка `order_data` несёт `product_type`**. ✅ approved 2026-08-12
4. **`metadata.product_type`** = тип **текущего** цикла ввода, не «тип всего КП». ✅ approved 2026-08-12
5. **`append_batches`** в metadata draft для undo; в PDF/XLSX сегментов нет. ✅ approved 2026-08-12
6. **Со 2-го захода** шаг клиента **полностью пропускаем**. Скидку меняют на result. ✅ approved 2026-08-12
7. **Моно-КП без append** — PDF/логистика/save как сейчас (без регрессии). ✅ approved 2026-08-12 (R3)
8. **Логистика = только ПБ (plates).** Вес для рейсов — **только** строки `product_type=plates` (`resolve_kp_line_weight_kg`). Non-plates в массу рейса **не** входят (пока нет их «цены»/веса для доставки). Поле «стоимость рейса» **активно**, если в КП есть хотя бы одна plate-строка; иначе — неактивно / без строки доставки (как у mono non-plates). ✅ locked via Q2 2026-08-12
9. **Save:** один `KP_offers` + строки в несколько `kp_*` по `line.product_type`, сквозной `position_number`. `kp_meta.product_type = mixed` при >1 типе. ✅ approved 2026-08-12
10. **Архив / resume:** **Q1 = C** — можно дописывать **уже сохранённое** КП из архива («Добавить другое наименование» → тот же loop → пересчёт → новый PDF/XLSX). Несколько бейджей типов (Q3). ✅ locked Q1/Q3
11. **Delete line + undo last batch** → всегда на **result**. ✅ approved 2026-08-12
12. **Колонка «Тип»** в result и multi-документе. ✅ approved 2026-08-12
13. **Unified колонки multi-документа (Q4):**  
    `№ | Тип | Наименование | Кол-во | Цена | Сумма`  
    Класс бетона и пр. тип-специфика — **в тексте наименования** (напр. `С30.15-3 (B25)`). Отдельной колонки «Класс» / «Вес» нет. ✅ locked Q4
14. **Production / СГП:** только plate-строки `kp_id`. ✅ approved 2026-08-12
15. **Без лимита заходов;** повтор типа ок. ✅ approved 2026-08-12
16. **Без новых npm/pip зависимостей.** ✅ approved 2026-08-12
17. **`line_id` на каждую строку** (стабильный id при создании строки). ✅ locked Q5

→ **All A1–A17 approved 2026-08-12.** Residuals R1–R3 locked below → PLAN.

---

## Decisions locked

### Из ideation

| # | Тема | Решение |
|---|------|---------|
| D1 | UX | «Добавить другое наименование» на result → picker → цикл → append |
| D2 | Порядок | Хронологический append, без группировки по типу |
| D3 | Клиент со 2-го цикла | Skip |
| D4 | Скидка | Одна на всё КП, sticky |
| D5 | Колонка типа | Да |
| D7 | Удаление | Строка + undo последнего захода |
| D9 | Лимит заходов | Нет |

### Из Q&A 2026-08-12

| # | Тема | Решение |
|---|------|---------|
| Q1 | Resume | **C — дописывать уже сохранённое КП** из архива (тот же `kp_id`, пересчёт, новый export) |
| Q2 / D6 | Логистика | **Считать только по ПБ (plates).** Non-plates не дают вес/«цену» доставки. Не полная заглушка. |
| Q3 | Бейджи архива | **Несколько бейджей** — по одному на каждый `product_type`, присутствующий в КП |
| Q4 | Класс / колонки | **В имени** + **унифицированные колонки** (без отдельной колонки класса/веса) |
| Q5 | Идентичность строк | **`line_id` на каждую строку** |

### Residuals locked 2026-08-12 (defaults after «все ок»)

| # | Тема | Решение |
|---|------|---------|
| R1 | Версии PDF после C | **Без истории версий в MVP** — последний export = актуальный (overwrite) |
| R2 | Статусы для append | **Только `status = «в работе»`** (узкий safe default; не СГП / выполнено / архив) |
| R3 | Mono vs unified | **Mono без append** — текущий PDF/XLSX шаблон без регрессии; multi / post-append → unified |

**Снято:** прежняя «логистика-заглушка для всего multi» — заменена правилом «рейсы только от веса ПБ».

---

## Objective

Менеджер собирает **одно КП** из нескольких заходов (часто >10, с повтором номенклатур), в т.ч. **дописывает уже сохранённое КП**, с общей скидкой, логистикой по весу ПБ и одним актуальным PDF/XLSX.

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | на result | «Добавить другое наименование» | дописать заказ в то же КП |
| US-2 | выбираю тип | тот же OCR/ввод | не учить новый UI |
| US-3 | клиент и скидка уже есть | skip client; скидка наследуется | быстрее закрыть чертеж |
| US-4 | смотрю result / PDF | порядок добавления + тип; единые колонки | читать один документ |
| US-5 | ошибся в OCR | undo заход / удалить строки по `line_id` | не начинать с нуля |
| US-6 | в КП есть ПБ | задать стоимость рейса; рейсы от веса **только плит** | доставка не врала из-за свай без веса |
| US-7 | КП уже в архиве | снова добавить наименование | дособрать КП без нового номера |
| US-8 | сохраняю / скачиваю | один `kp_id`, актуальные файлы | отправить клиенту |

### Reframed success criteria

| Требование | Измеримый критерий |
|------------|-------------------|
| Append loop | ≥2 захода (в т.ч. повтор типа) → один `kp_id` |
| Порядок | `order_data` / PDF / XLSX = порядок append |
| Скидка | один `%`; пересчёт всех строк |
| Skip client | со 2-го захода нет client step |
| Колонки | multi: `№ \| Тип \| Наименование \| Кол-во \| Цена \| Сумма`; grade в имени |
| line_id | у каждой строки; delete/undo по id |
| Логистика ПБ | `cargo_kg` = сумма весов только `plates`; delivery = trip × ceil(kg/18600); non-plates не влияют |
| Нет ПБ | нет строки доставки / поле рейса неактивно |
| Archive C | из карточки сохранённого КП (`в работе`) → append → totals/files обновлены на том же `kp_id` |
| Бейджи | в списке архива N бейджей по типам в КП |
| Mono | один заход одного типа — без регрессии |
| Production | только plate-строки |

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2 |
| Domain | `core/` (pricing, KP persistence, PDF/XLSX) |
| Frontend | React 19, TypeScript, Vite, TanStack Query |
| Data | SQLite `plita.db` / `pb.db` |
| Tests | pytest, vitest + Testing Library |

---

## Commands

```bash
./run+logs.sh

source venv/bin/activate
uvicorn app.main:app --reload
pytest tests/ -k "commercial or wizard or kp_ or archive" -q
pytest tests/test_commercial_web_flow.py tests/test_commercial_wizard_step_service.py -q

cd frontend
npm run typecheck
npm run test
npm run build
```

UI: `http://localhost:5173/commercial-offer/new` + архив КП.

---

## Project Structure (touch map)

```
ai_docs/ideas/kp-multi-nomenclature-append.md
ai_docs/specs/kp-multi-nomenclature-append.md
ai_docs/develop/plans/2026-08-12-kp-multi-nomenclature-append.md

app/schemas/commercial.py                          # line_id, append_batches, mixed
app/services/commercial_draft_service.py           # append / undo / delete
app/services/commercial_wizard_step_service.py
app/services/commercial_calculation_service.py     # PB-only cargo weight
app/services/commercial_export_service.py
app/services/archive_service.py / offers_service   # resume saved KP → edit/append
app/api/v1/endpoints/commercial.py
app/api/v1/endpoints/archive.py

core/cargo_delivery_pricing.py                     # EXTEND — weight filter plates-only option
core/commercial_pricing.py
core/commercial_offer.py / commercial_offer_xlsx.py
core/commercial_line_format.py                     # format_line_name (planned)
core/kp_persistence_service.py                     # multi-table + update existing kp_id
core/kp/offers_read.py / offers_write.py
core/kp_db_schema.py                               # mixed; line_id on line tables

frontend/.../CalculationResultStep.tsx
frontend/.../CommercialOfferWizard.tsx
frontend/.../ProductTypePicker.tsx
frontend/.../commercial-archive/*                  # CTA append, multi badges
frontend/.../wizardDraftStore.tsx
frontend/.../wizardStepOrder.ts
frontend/.../types + api
```

---

## Code Style

- Явный `product_type` + `line_id` на каждой строке.
- Не полагаться на `is_pile_order(entire_order)` для mixed — ветвление **per line**.
- Логистика: `total_order_cargo_weight_kg(order_data, product_types={"plates"})` (или эквивалент).
- Display name для export: единая функция `format_line_name(item)` (марка + grade в скобках и т.д.).
- Пример строки:

```python
{
    "line_id": "ln_01HZX…",
    "product_type": "piles",
    "append_batch_id": "b3",
    "mark": "С30.15-3",
    "concrete_grade": "B25",
    "qty": 12,
    "unit_price": 15200.0,
}
# в PDF наименование: "С30.15-3 (B25)"
```

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit | append/undo/delete by `line_id`; PB-only cargo; name formatting with grade; discount on mixed |
| API flow | Плиты→Сваи→Плиты create; **append to saved kp_id**; export; mono regression |
| Persistence | mixed multi-table; position_number order; update existing KP lines (sync by line_id) |
| Archive UI | multi badges; CTA «добавить наименование» only «в работе» |
| Manual | дописать архивное КП, скачать PDF, проверить рейсы только от плит |

**PLAN rule:** каждый шаг обкладывать тестами (TDD: failing test → implement → green). См. Verify в плане.

---

## Boundaries

**Always**
- Релевантные pytest + frontend typecheck/test перед done
- Mono без регрессии
- `product_type` + `line_id` на строках
- Рейсы считать **только** от веса plates; non-plates не добавлять в cargo_kg
- Append/update saved KP только при статусе «в работе»

**Ask first**
- Миграции сверх `mixed` / хранения `line_id` в line-таблицах (уже в PLAN)
- Веса/доставка non-plates
- Production whitelist changes beyond mixed-with-plates
- Append для других статусов

**Never**
- Склейка нескольких `kp_id` в один PDF
- Сегменты «Этаж N» в MVP
- Группировка строк по типу в export
- Bot path
- Секреты / live DB в git
- Подмешивать вес non-plates в рейсы «тихо»
- История версий PDF в MVP

---

## Behaviour detail

### Flow (wizard)

```
[Picker] → [Input] → [Client]* → [Result]
                                    ├─ Добавить наименование → [Picker] → [Input] → [Result]
                                    ├─ Undo last batch / Delete line → [Result]
                                    └─ Save / Download
* только первый заход новой сессии создания (header ещё пуст).
```

### Flow (archive — Q1 C)

```
[Архив → карточка КП, status=«в работе»] → «Добавить другое наименование»
  → загрузка order_data + header в wizard (тот же kp_id)
  → [Picker] → [Input] → [Result] (client skip)
  → Save обновляет тот же kp_id, перегенерирует файлы (overwrite, без version history)
```

### Логистика (Q2)

```
plates_kg = sum(weight(line) for line in order_data if line.product_type == "plates")
trips = ceil(plates_kg / 18600)   # если plates_kg > 0
delivery = trip_cost * trips     # если trip_cost > 0 и trips > 0
```

- Target-sum: T = сумма продукции со скидкой **+** delivery (как сейчас), где delivery только от ПБ.
- Нет plate-строк → delivery = 0, UI рейса disabled.
- Плиты→Плиты (multi same type): логистика **живая** (весь вес плит).

### Export (Q4 + R3)

- Mono один тип / один заход: **текущие шаблоны** (без регрессии).
- Multi / post-append / C-edit: unified `№ | Тип | Наименование | Кол-во | Цена | Сумма`; grade в наименовании; строка доставки — только если delivery > 0 по правилу ПБ.

### Persistence

- Create или **update** `kp_id` (C): **sync by `line_id`** (не wipe `kp_plates` с production state); сквозной `position_number`.
- `kp_meta.product_type = mixed` при >1 типе.
- Read: все `kp_*` → merge sort by `position_number`.

### Archive badges (Q3)

- Уникальные типы в КП → бейдж на каждый (Плиты, Сваи, …).
- Фильтр архива: КП показывается, если содержит выбранный тип (mixed с плитами проходит фильтр «Плиты»).

---

## Success Criteria

- [x] SC-1: Плиты→Сваи→Плиты → один `kp_id`
- [x] SC-2: Unified PDF/XLSX: порядок, тип, grade в имени, одна скидка/сумма
- [x] SC-3: Skip client со 2-го захода
- [x] SC-4: Undo/delete по `line_id` / batch
- [x] SC-5: Delivery только от веса plates; без plates — без доставки
- [x] SC-6: Append к сохранённому КП из архива (Q1 C) на том же `kp_id` (только «в работе»)
- [x] SC-7: Несколько бейджей типов в архиве
- [x] SC-8: Mono без append — без регрессии
- [x] SC-9: Tests green; production только plates (mixed-with-plates OK)

Evidence: [`2026-08-12-kp-multi-nomenclature-append-implementation.md`](../develop/reports/2026-08-12-kp-multi-nomenclature-append-implementation.md).

---

## Open Questions

**None blocking PLAN.** Residuals R1–R3 locked 2026-08-12.

---

## Out of scope

- Вес/доставка non-plates (сваи и др. в cargo_kg)
- Сегменты / этажи
- Разные скидки по заходам
- История версий PDF (R1)
- Append вне статуса «в работе» (R2)
- Production/СГП для non-plates
- Большой generic multi-product framework сверх нужного для append

---

**Gate:** SPECIFY ✅ → PLAN ✅ → IMPLEMENT ✅. Report: [`2026-08-12-kp-multi-nomenclature-append-implementation.md`](../develop/reports/2026-08-12-kp-multi-nomenclature-append-implementation.md).
