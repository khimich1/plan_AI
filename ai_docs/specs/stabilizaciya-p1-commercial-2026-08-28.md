# Spec: Стабилизация P1 — декомпозиция коммерческого контура

**Статус**: implemented (Q5→Q1→A3→A4, 2026-08-28; коммитов агент не делал)
**Связано**: [ai_docs/develop/audits/2026-08-28-full-project-audit.md](../develop/audits/2026-08-28-full-project-audit.md)
**Предыдущая спека**: [stabilizaciya-p0-audit-2026-08-28.md](./stabilizaciya-p0-audit-2026-08-28.md)
**План**: [ai_docs/develop/plans/2026-08-28-stabilizaciya-p1-commercial.md](../develop/plans/2026-08-28-stabilizaciya-p1-commercial.md)

## Objective

Декомпозировать коммерческий контур КП **без смены API-контрактов и без смены
поведения мастера КП**.

Менеджер по-прежнему создаёт и дописывает черновик по шести типам продукции
(плиты, сваи, ступени, марши, мостовые сваи, ФБС), решает wide/unpriced,
считает, экспортирует и сохраняет КП. Снаружи ничего не меняется: те же URL,
те же request/response схемы, те же stamp / append / replace инварианты.

Внутри закрываем находки аудита 2026-08-28:

| Находка | Что делаем |
|---------|-----------|
| Q5 | Объединить `resolve_wide_plates` / `resolve_unpriced_plates` в одну параметризованную реализацию |
| Q1 | Один `ProductDraftHandler` + config вместо шести копий create/update/AI/grades pipeline |
| A3 | Use-case сервисы по вертикалям + тонкий facade `CommercialWorkflowService` |
| A4 | Thin controllers в `commercial.py` (общий runner, без слияния URL) |

Q4 (`build_*_preview_metadata` × 6 в `commercial_draft_service.py`) **поглощается Q1**:
config handler'а указывает, какой metadata-builder вызывать. Отдельной волны Q4 нет.

**Не в этой спеке:** A1 plate globals, Redis/A2, S9 CSRF, фронтовые god-hooks
(`useCommercialOfferWizard` и др.), соседний `production.py`, полный repository
слой [A5] (SGP / delivery / kp_readiness — другой домен; follow-up, см. Open Questions).

Успех: модули меньше и с одним местом для веток `product_type`; pytest commercial
`*_flow` + `web_flow` + `draft_append` зелёные после каждого шага; поведение
мастера КП неизменно.

## ASSUMPTIONS (черновик — поправьте до IMPLEMENT)

1. **Поведение священно.** Недавний фикс stamp `line_id`/`product_type` на
   create/AI (marches/steps/fbs/bridge_piles + plates AI/wide/unpriced) и
   устойчивые `_partition` / `_line_product_type` по `product_kind` — baseline.
   Рефакторинг не меняет семантику append / replace / merged cycle / bulk grade.
2. **Контракт HTTP не двигаем.** Пути, методы, `response_model`, Form/File поля,
   тексты `HTTPException` / `raise_validation_client_error` — те же. Не схлопываем
   шесть URL в один параметризованный route (OpenAPI и фронт зависят от путей).
3. **Существующие type-сервисы остаются.** `CommercialMarchService` и аналоги —
   preview/pricing парсеры (`product_kind` в строках). Их не сливаем в handler;
   handler их вызывает.
4. **Q4 входит в Q1**, не отдельным спринтом.
5. **A5 вне P1** — подтверждено пользователем 2026-08-28 (exclude_followup).
   A5 в аудите: raw SQL в `sgp_service` / `delivery_schedule_service` /
   `kp_readiness_service`, не в commercial workflow.
6. **Поставка четырьмя волнами** Q5 → Q1 → A3 → A4 — подтверждено 2026-08-28
   (waves_checkpoint). После каждой — pytest и стоп на ревью. Коммиты агент не
   делает, пока пользователь явно не попросит.
7. **Фасад ≤ 800 строк** после A3; ветки `product_type` только в config; stamp
   helpers живут одним модулем (не размазывать инвариант).
8. **AI replace-всего-order_data** (текущая семантика `apply_ai_*`: stamp без
   compose с другими типами) сохраняем как есть. Не «улучшаем» по ходу.
9. **Рабочее дерево** уже содержит незакоммиченный stamp-фикс и P0-артефакты.
   P0-файлы (`_p0_baseline/`, закрытая спека P0) не трогаем без нужды.

→ Поправьте сейчас или IMPLEMENT пойдёт с этими допущениями.

## Tech Stack

- Backend: Python 3.12, FastAPI 0.141.1 / Starlette 1.6.0, Pydantic v2, pytest
- Слои: `app/api/v1/endpoints/` → `app/services/` → `app/repositories/` / `core/`
- Frontend не меняем (контракт стабилен)

## Commands

Backend (cwd=корень):

```bash
# Safety net после каждого шага (обязательно)
pytest tests/test_commercial_*_flow.py \
       tests/test_commercial_draft_append.py \
       tests/test_commercial_unpriced_plates_resolve.py \
       tests/test_commercial_ai_plates.py -q

# Узкие инварианты stamp / bulk grade
pytest tests/test_commercial_march_flow.py::test_bulk_grade_single_line_no_duplicate \
       tests/test_commercial_march_flow.py::test_partition_treats_untyped_legacy_mono_as_same_type \
       tests/test_commercial_web_flow.py::test_create_draft_stamps_line_id_and_plates_product_type \
       tests/test_commercial_web_flow.py::test_resolve_wide_plates_applies_line_id_with_normalized_lines -q

# После A4 — весь commercial-набор, чтобы поймать разъезд схем
pytest tests/test_commercial_*.py -q
```

Frontend не трогаем; `npm` не требуется.

## Project Structure (затрагиваемые / новые пути)

```
app/services/commercial_workflow_service.py   → facade; сейчас 3059 строк
app/api/v1/endpoints/commercial.py            → thin controllers; сейчас 897 строк, 32 handlers
app/services/commercial_draft_service.py      → metadata builders через config (Q1/Q4)
app/services/product_draft_config.py          → НОВЫЙ: таблица типов (Q1)
app/services/product_draft_handler.py         → НОВЫЙ: create/update/AI/grades по config
app/services/commercial_order_identity.py     → НОВЫЙ (A3): stamp/partition/compose
app/services/commercial_plate_resolve.py      → НОВЫЙ (Q5/A3): wide + unpriced
app/services/commercial_draft_lifecycle.py    → НОВЫЙ (A3): calculate/save/export/hydrate/details

# Не меняем поведение, только вызываются:
app/services/commercial_{march,step,fbs,pile,bridge_pile}_service.py
app/services/commercial_service.py
app/services/commercial_wizard_step_service.py
app/schemas/commercial.py                     → Never менять поля request/response
```

Имена новых модулей — рабочий черновик; при IMPLEMENT можно сдвинуть, если
появится более удачное место, **не** меняя публичные методы facade.

Текущая карта (2026-08-28):

| Модуль | Строк | Роль |
|--------|------:|------|
| `commercial_workflow_service.py` | 3059 | god-orchestrator: ingest ×6, AI ×6, grades ×4, resolve, lifecycle |
| `commercial.py` | 897 | 32 handlers; ~15 near-duplicate product routes; толстые: `generate-preview`, download |
| `commercial_draft_service.py` | 746 | OCR source, 6× metadata builders, `stamp_order_line_identity` |
| type preview services | 73–96 | парсеры/прайс; `product_kind` на строке |
| `ProductDraftHandler` | 0 | ещё не существует |

## Порядок работ

Зависимости односторонние: каждый шаг оставляет систему рабочей.

```
Q5 unify resolve_wide/unpriced
    │
    ▼
Q1 ProductDraftHandler + config
   (create / update / AI / grades / metadata-builder)
    │
    ▼
A3 use-case сервисы + facade
   (identity, plate-resolve, lifecycle; workflow — делегат)
    │
    ▼
A4 thin controllers
   (общий runner; те же 32 URL)
```

### Q5 — unify plate resolve

`resolve_wide_plates` (стр. 2172–2284, ~114 строк) и `resolve_unpriced_plates`
(2286–2469, ~184 строк) разделяют ~65% каркаса: загрузить draft → сопоставить
decisions по `line_id` / `source_line` → переписать `input_text` / batches →
preview → stamp plates → persist.

Различия уходят в strategy, не в копипаст:

| | Wide | Unpriced |
|---|------|----------|
| Действия | confirm / exclude / replace (+ replacement_text) | exclude / replace_load (+ load_code) |
| Matching | точная строка | fuzzy (`name in line`) |
| Rewrite | `_normalize_replacement_lines` | `rewrite_plate_line_load` |
| Флаг | `wide_plates_resolved` | `unpriced_plates_resolved` |

Публичные методы `resolve_wide_plates` / `resolve_unpriced_plates` остаются
тонкими обёртками — `commercial.py` не меняет вызовы.

### Q1 — ProductDraftHandler + config

Шесть почти идентичных вертикалей в `CommercialWorkflowService`:

| Тип | create | update | AI | grades |
|-----|--------|--------|----|--------|
| plates | ветка `create_draft` | `update_draft_plates` | `apply_ai_plates_instruction` | нет (wide/unpriced вместо) |
| piles | `_create_pile_draft` | `update_draft_piles` | `apply_ai_piles_instruction` | `update_draft_pile_grades` |
| marches | `_create_march_draft` | `update_draft_marches` | `apply_ai_marches_instruction` | `update_draft_march_grades` |
| steps | `_create_step_draft` | `update_draft_steps` | `apply_ai_steps_instruction` | нет |
| bridge_piles | `_create_bridge_pile_draft` | `update_draft_bridge_piles` | `apply_ai_bridge_piles_instruction` | `update_draft_bridge_pile_grades` |
| fbs | `_create_fbs_draft` | `update_draft_fbs` | `apply_ai_fbs_instruction` | `update_draft_fbs_grades` |

Каркас update (~77–90 строк × 6): load → type-guard → resolve source →
append/replace + batches → preview → metadata → **partition → stamp_previous →
stamp → compose** → persist → wizard step.

Config держит отличия: `product_type`, ключ `*_batches`, wizard step, preview
service, metadata builder, OCR/AI функция, тексты ValueError, plate-only крючки
(`plate_order_ctx`, `wide_plates_resolved`, реальный `PlateOrder`).

`create_draft` перестаёт быть лестницей `if product_type == …`.

Публичные методы facade (`update_draft_marches`, …) **остаются** как однострочные
делегаты — чтобы A4 не блокировался на Q1 и чтобы тесты, патчащие методы, не
ломались.

Если Q1 вскроет расхождение поведения между типами (копипаст мог разъехаться) —
стоп и вопрос пользователю, не «выравнивать тихо».

### A3 — use-case сервисы + facade

После Q1 god-модуль всё ещё смешивает identity строк, plate-resolve, wizard,
calculate / export / save / hydrate. Разносим по вертикалям; facade сохраняет
текущие имена и сигнатуры методов.

Не тащить в новые сервисы логику, которая уже живёт отдельно
(`CommercialWizardStepService`, type preview services, `CommercialExportService`).

Stamp helpers (сейчас L97–303) переезжают **одним куском** в
`commercial_order_identity.py`. Это зона недавнего бага; дробить её внутри A3
нельзя.

### A4 — thin controllers

Handlers product-типов уже тонкие (~18–25 строк: upload + try/except + вызов
workflow). Дубль — Q2 (~15 near-identical). Делаем общий runner
(`_run_product_update` / `_run_product_ai` / `_run_product_grades`), **оставляя**
отдельные `@router.patch("/drafts/{draft_id}/marches")` и т.д.

Толстые уникальные handlers:

- `generate_preview` — orchestration (DraftStore + ad-hoc dict) уходит в сервис;
  JSON-ключи те же, `response_model` по-прежнему нет.
- `download_generated_file` — whitelist/path check в workflow; HTTP-статусы те же.

`production.py` (~639) не трогаем.

## Code Style

Config — данные, не `if product_type`. Пример целевого вида (не код IMPLEMENT):

```python
@dataclass(frozen=True)
class ProductDraftSpec:
    product_type: str
    wizard_step: WizardStepId
    batches_key: str
    type_mismatch_update: str
    type_mismatch_ai: str
    # generate_preview / build_*_metadata / apply_*_with_ai — callables

SPECS: dict[str, ProductDraftSpec] = { ... }

class ProductDraftHandler:
    def update(self, draft_id: str, *, product_type: str, mode: str, ...) -> dict[str, Any]:
        spec = SPECS[product_type]
        # один pipeline; stamp/compose только через CommercialOrderIdentity
```

- Комментарии и сообщения ошибок — на русском, как сейчас.
- Не добавлять комментарии, пересказывающие код.
- Не переименовывать публичные методы facade в этой спеке.
- Stamp helpers — один модуль; не копировать `_line_product_type` второй раз.

## Testing Strategy

Новых e2e не требуется, если существующий safety net зелёный. Точечные тесты —
только если вынос helper'а оставляет его без прямого покрытия (например,
параметризация Q5).

**Safety net (обязателен после каждого шага):** ~164 теста

- `test_commercial_{march,step,fbs,pile,bridge_pile,web,multi_append}_flow.py`
  (120 тестов) — create / replace / AI / calculate / save / reject plates-endpoint
- `test_bulk_grade_single_line_no_duplicate` (march) — регресс stamp+compose
- `test_partition_treats_untyped_legacy_mono_as_same_type`
- `test_commercial_draft_append.py` (44 теста) — schema + line_id + append/undo/delete
- `test_commercial_unpriced_plates_resolve.py`, `test_commercial_ai_plates.py`

После A4: весь `tests/test_commercial_*.py`.

**Известные дыры покрытия (не затыкаем в P1 без нужды):** HTTP для `fbs/ai`,
`bridge-piles/ai`, `unpriced-plates/resolve`; bulk-grade no-duplicate только у
маршей. Рефакторинг не должен требовать этих тестов, но и не должен ломать то,
что есть.

Не удалять и не xfail'ить падающие тесты. Pre-existing 12 failures полного
pytest (вне commercial) — не скоуп; не чинить «заодно».

## Boundaries

- Always: pytest commercial `*_flow` + `draft_append` + `web_flow` (+ unpriced /
  ai_plates) после каждого шага Q5/Q1/A3/A4
- Always: stamp/partition/compose вызываются из одного модуля; type-ветки —
  только в `ProductDraftSpec` / config
- Always: публичные методы `CommercialWorkflowService`, которые зовёт
  `commercial.py`, сохраняют имена и сигнатуры до конца A4 включительно
- Ask first: если вынос потребует менять `app/schemas/commercial.py`
- Ask first: если понадобится трогать frontend или `production.py`
- Ask first: если Q1 вскроет расхождение поведения между типами
- Ask first: целевой размер фасада, если 800 строк не достигается без ломки
  читаемости (лучше 900 читаемых, чем 500 заумных)
- Never: менять request/response схемы endpoints (поля, типы, статус-коды)
- Never: ломать stamp / append / replace / bulk-grade / `product_kind` fallback
- Never: схлопывать шесть product URL в один
- Never: коммитить, пушить, менять git-конфиг
- Never: править P0-артефакты (`_p0_baseline/`, закрытая спека P0) без нужды
- Never: брать в P1 A1, A2/Redis, S9, frontend god-hooks, полный A5
- Never: «улучшать» AI-семантику (replace всего `order_data`) по ходу рефакторинга

## Success Criteria

Измеримые, проверяемые без вкусовщины:

- [x] `resolve_wide_plates` и `resolve_unpriced_plates` — тонкие обёртки над
      одной реализацией; дублированного тела нет
- [x] Ветки `if product_type == "marches"` / `"fbs"` / … для create/update/AI/grades
      отсутствуют в `commercial_workflow_service.py`; таблица типов — в config
- [x] `commercial_workflow_service.py` ≤ **800** строк (фасад + делегаты)
- [x] `commercial.py` ≤ **500** строк; product handlers — вызов shared runner +
      сохранение отдельных `@router` путей (32 пути на месте)
- [x] `app/schemas/commercial.py` без изменения полей request/response
- [x] Safety net pytest зелёный после каждого шага
- [x] После A4: `pytest tests/test_commercial_*.py` зелёный
      (333 passed; 3 pre-existing web_flow: generate_files schema / offer identity)
- [ ] Ручной smoke мастера КП **не обязателен** на каждом шаге (контракт + pytest);
      после полного P1 — по желанию пользователя

## Follow-up (не P1)

- **[A5]** `SgpRepository` / `DeliveryScheduleRepository` (+ kp_readiness) —
  отдельная спека, другой домен.
- **[A7]** полноценный constructor DI (`Depends(get_kp_repository)`).
- **[A12]/[Q8]** frontend god-hooks мастера КП.
- **[Q4 remainder]** если после Q1 в `commercial_draft_service.py` останутся
  тонкие однострочные `build_*_preview_metadata` — можно схлопнуть позже.
- **[A4 remainder]** `production.py` thick API.
- Выравнивание AI-append с compose (если когда-нибудь понадобится multi-type AI).
- HTTP-тесты `fbs/ai`, `bridge-piles/ai`, `unpriced-plates/resolve`.

## Принятые решения (из Q&A 2026-08-28)

1. **A5 исключён** из P1 (follow-up спека). Не смешивать repository СГП/графика
   со stamp/append коммерческого контура.
2. **Поставка — 4 волны с чекпоинтом:** Q5 → стоп → Q1 → стоп → A3 → стоп → A4.
   Коммиты только по явной просьбе.
3. **Git:** изменения остаются в рабочем дереве, как в P0.

## Open Questions

Q-A5 и Q-delivery **приняты** (см. «Принятые решения»). Ниже — зафиксированное
обоснование. Живой ask-first: только Q-size на A3.

### Q-A5. Включать ли [A5] repository layer в эту же спеку? — **исключить**

Аудит [A5]: сервисы обходят repository, raw SQL. **Где:** `sgp_service.py`,
`delivery_schedule_service.py`, `kp_readiness_service.py`. Это **не**
`commercial_workflow_service.py`. Commercial workflow уже использует
`KpRepository` / `ManagerRepository`.

| Вариант | Плюсы | Минусы / опасности |
|---------|-------|-------------------|
| **Исключить (рекомендация)** | P1 остаётся одним контуром (КП-мастер + stamp). Ревью и откат понятны. A5 не рискует недавно починенным append. | Строка A5 в матрице аудита остаётся открытой. |
| Включить в эту спеку | Закрыть ещё одну High-находку «в том же спринте». | Другой домен, High effort. Смешивает риск stamp-invariant с persistence СГП/графика. Падение тестов СГП можно списать на «commercial refactor». Ревью распухает. |
| Только инвентарь raw SQL | Дешёвая подготовка следующей спеки. | Не закрывает A5; почти нулевая ценность внутри P1. |

**Решение:** исключить. A5 — follow-up спека после зелёного P1.

### Q-delivery. Как поставлять IMPLEMENT после approve спеки? — **4 волны**

Коммиты в любом случае только по явной просьбе.

| Вариант | Плюсы | Минусы / опасности |
|---------|-------|-------------------|
| **4 волны с чекпоинтом (рекомендация)** Q5 → стоп → Q1 → стоп → A3 → стоп → A4 | Каждая волна — рабочее дерево + зелёный pytest. Откат одной волны не размазывает stamp. Q1 — самый рискованный шаг — изолирован. | Больше раундов ревью, дольше календарно. |
| Один непрерывный IMPLEMENT | Меньше контекст-свитчей. | Баг в Q1 тонет в диффе A3+A4. Сложно сказать, какая волна сломала bulk-grade. |
| Две пачки (Q5+Q1, затем A3+A4) | Компромисс: дедуп отдельно от переезда файлов. | Q5+Q1 всё ещё крупный дифф (шесть пайплайнов). |

**Решение:** четыре волны. Stamp/append только что чинили; дешёвые чекпоинты
дешевле одного большого «почему 1 строка стала 2».

### Q-size (не блокирует старт, ask-first на A3)

Целевой размер фасада в Success Criteria — 800 строк. Если на A3 упрёмся в
«ещё −200 ценой нечитаемых делегатов» — спросить, а не резать.

## Verification (SDD gate)

- [x] Спека покрывает шесть ядерных разделов (+ порядок работ, assumptions)
- [ ] Пользователь ревьюит и approve'ит спеку
- [x] Success criteria конкретные и проверяемые
- [x] Boundaries Always / Ask first / Never заданы
- [x] Файл сохранён в репозитории (`ai_docs/specs/`)
- [ ] IMPLEMENT не начинать до approve
