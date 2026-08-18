# Handoff: КП — несколько наименований (append loop) → оркестратор

> **Дата:** 2026-08-12  
> **Ветка:** `aleksey_web`  
> **Статус:** Idea ✅ · Spec ✅ · Plan ✅ · **Implementation 0/20** — готов к `/orchestrate`  
> **Цель файла:** открыть **новое окно** и сразу запустить оркестратор без потери контекста ideation/SDD.  
> **Не коммитить** без явной просьбы пользователя.

---

## Как стартовать новую сессию (скопируй в первый промпт)

```
/orchestrate execute orch-2026-08-12-14-05-kp-multi-append

Контекст: handoff ai_docs/develop/handoffs/2026-08-12-kp-multi-nomenclature-append.md
План уже готов — не перепланировать. TDD на каждую задачу. Не коммитить без просьбы.
Начать с MNA-001 ∥ MNA-002 (phase 0).
```

### Чеклист агента в новом окне

1. Прочитать **этот** handoff.
2. Прочитать `.cursor/skills/plan-web-context/SKILL.md`.
3. Прочитать `.cursor/skills/orchestration/SKILL.md` (координатор сам код **не** пишет).
4. Загрузить workspace:
   - `.cursor/workspace/active/orch-2026-08-12-14-05-kp-multi-append/progress.json`
   - `tasks.json`, `links.json`
5. Источник задач: `ai_docs/develop/plans/2026-08-12-kp-multi-nomenclature-append.md` (не выдумывать scope).
6. Спека: `ai_docs/specs/kp-multi-nomenclature-append.md` — все A1–A17 и R1–R3 **locked**.
7. Запустить task loop с **MNA-001** (можно параллельно **MNA-002** / **MNA-301** — см. plan «Suggested agent split»).

**Режим:** `/orchestrate`, не «просто multitask». Multitask можно как оболочку фона, но цикл обязан быть Worker → Test-Writer → Test-Runner → Reviewer на каждую MNA-*.

---

## Артефакты (source of truth)

| Артефакт | Путь |
|----------|------|
| Idea | [`ai_docs/ideas/kp-multi-nomenclature-append.md`](../../ideas/kp-multi-nomenclature-append.md) |
| Spec | [`ai_docs/specs/kp-multi-nomenclature-append.md`](../../specs/kp-multi-nomenclature-append.md) |
| Plan (20 tasks) | [`ai_docs/develop/plans/2026-08-12-kp-multi-nomenclature-append.md`](../plans/2026-08-12-kp-multi-nomenclature-append.md) |
| Orchestration ID | `orch-2026-08-12-14-05-kp-multi-append` |
| Workspace | `.cursor/workspace/active/orch-2026-08-12-14-05-kp-multi-append/` |
| Предыдущий handoff (линейка продуктов) | [`2026-08-05-kp-multiproduct-ls-lm-to-fbs.md`](2026-08-05-kp-multiproduct-ls-lm-to-fbs.md) |

### Состояние оркестрации (на момент handoff)

```
status: ready
phase: PLAN
tasksTotal: 20
tasksCompleted: 0
currentTask: null
```

`links.json` → plan / spec / idea уже прописаны. Report = null (создаст documenter в конце).

---

## Что уже решено (не переспрашивать)

| ID | Решение |
|----|---------|
| UX | Кнопка «Добавить другое наименование» на result → picker → OCR/ввод → append в конец |
| Порядок | Хронологический; **без** группировки по типу в PDF |
| Client | Со 2-го цикла **skip**; sticky client/manager/conditions/discount |
| Скидка | Одна на всё КП |
| Колонки multi | `№ \| Тип \| Наименование \| Кол-во \| Цена \| Сумма`; **класс в имени** |
| Identity | `line_id` на каждую строку; `append_batches` для undo |
| Удаление | Delete line + «Отменить последний заход» → снова result |
| **Q1** | **C** — дописывать **сохранённое** КП из архива, тот же `kp_id` |
| **Q2** | Логистика **живая**, рейсы **только от веса ПБ (plates)**; non-plates ∉ cargo_kg |
| **Q3** | Несколько бейджей типов в архиве; фильтр «содержит тип» |
| **R1** | Без истории версий PDF — последний export = актуальный |
| **R2** | Append **только** при `status = «в работе»` (не СГП / выполнено / …) |
| **R3** | Mono без append — **текущие** PDF/XLSX шаблоны, без регрессии |
| Production | Только plate-строки; bot out of scope |
| Persist update | Sync **by `line_id`** — не wipe `kp_plates` (production state) |

Ломает старый инвариант handoff 2026-08-05 «одно КП = один `product_type`» → для mixed `kp_meta.product_type = mixed`.

---

## Фазы плана (20 задач)

| Phase | Tasks | Суть |
|-------|-------|------|
| 0 Domain | MNA-001, MNA-002 | PB-only cargo filter; `format_line_name` |
| 1 Draft | MNA-101…104 | schemas, line_id, append/undo/delete API, skip client |
| 2 Calc | MNA-201, MNA-202 | plates-only logistics + mixed discount |
| 3 Persist | MNA-301…304 | migration line_id, multi-table create/read, update+status gate |
| 4 Export | MNA-401, MNA-402 | unified multi; mono R3 |
| 5 Wizard FE | MNA-501, MNA-502 | UI Тип/CTA/undo; loop sticky |
| 6 Archive | MNA-601, MNA-602 | hydrate from KP (Q1=C); badges + filter |
| 7 Gate | MNA-701, MNA-702 | production mixed-with-plates; E2E + mono regression |

Каждая задача в плане имеет **Acceptance + Verify** (pytest/vitest/typecheck). Оркестратор: сначала тесты (TDD), потом код.

**Стартовый split:** `MNA-001 ∥ MNA-002 ∥ MNA-301` → затем `MNA-101 → MNA-102 → MNA-103`.

---

## Ключевые модули (куда лезть)

```
core/cargo_delivery_pricing.py      # PB-only weight filter
core/commercial_pricing.py          # calculate_total_cost, per-line types
core/commercial_offer.py / _xlsx.py # unified export branch
core/kp_persistence_service.py      # multi-table create + sync-by-line_id update
core/kp/offers_read.py / offers_write.py
core/kp_db_schema.py                # line_id on kp_*; mixed meta

app/schemas/commercial.py
app/services/commercial_draft_service.py
app/services/commercial_wizard_step_service.py
app/services/commercial_calculation_service.py
app/services/commercial_export_service.py
app/services/archive_service.py / offers_service
app/api/v1/endpoints/commercial.py (+ archive/offers)

frontend/src/features/commercial-offer/
  CalculationResultStep, CommercialOfferWizard, ProductTypePicker
  wizardDraftStore, wizardStepOrder, types, api
frontend/src/features/commercial-archive/  # badges, CTA append
```

UI smoke: `./run+logs.sh` → `http://localhost:5173/commercial-offer/new` + архив.

---

## Команды проверки (из плана / спеки)

```bash
source venv/bin/activate
pytest tests/ -k "commercial or wizard or kp_ or archive or cargo or logistics" -q

cd frontend && npm run typecheck && npm run test && npm run build
```

Точные `-k` / vitest-пути — в блоке Verify каждой MNA-задачи в plan.md.

---

## Out of scope (не делать в этом orch)

- Вес/доставка non-plates в cargo  
- Сегменты / «Этаж N» / группировка по типу в PDF  
- Разные скидки по заходам  
- История версий PDF  
- Append вне статуса «в работе»  
- Production для non-plates  
- Bot / `bot_archived`  
- Generic multi-product framework «на будущее» сверх append  
- Новые npm/pip зависимости  
- Git commit/push без явной просьбы  

---

## Риски (из плана — не игнорировать)

1. Wipe `kp_plates` при update → потеря production — только sync by `line_id`.  
2. `is_*_order(entire_order)` ломает mixed — ветвление **per line**.  
3. Append API не должен replace весь `order_data` чужого типа.  
4. Фильтр архива `product_type = ?` промахивает `mixed` — «содержит тип».  
5. FE step order завязан на один `productType` — отделить cycle type от sticky header.  
6. Mono PDF regression — R3 тесты обязательны до/вместе с unified.

---

## Связь с прошлой линейкой

Предыдущий handoff закрывал отдельные `product_type` (ЛС→ЛМ→мостовые→ФБС) при инварианте «1 КП = 1 тип».  
Эта работа — **горизонтальный** append нескольких типов (и повторов) в одно КП + дозапись из архива. ФБС и прочие mono-потоки не выкидывать; расширять.

---

## Definition of done (оркестрация)

- [ ] Все MNA-001…702 `completed` в `tasks.json`  
- [ ] SC из спеки закрыты (в т.ч. Q1=C, PB-only logistics, multi badges, mono R3)  
- [ ] Documenter: report в `ai_docs/develop/reports/` + обновление спеки PLAN ✅ → IMPLEMENT ✅  
- [ ] Пользователь решает про commit отдельно  

---

**Первое действие в новом окне:** `/orchestrate execute orch-2026-08-12-14-05-kp-multi-append`
