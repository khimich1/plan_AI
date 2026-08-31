# Implementation Plan: Стабилизация P1 — декомпозиция коммерческого контура

**Спека**: [ai_docs/specs/stabilizaciya-p1-commercial-2026-08-28.md](../../specs/stabilizaciya-p1-commercial-2026-08-28.md)
**Дата**: 2026-08-28
**Статус**: implemented (волны Q5→Q1→A3→A4 завершены 2026-08-28)

## Overview

Четыре последовательные волны внутри коммерческого контура КП: сначала схлопнуть
дубли resolve wide/unpriced (Q5), затем один config-driven pipeline вместо шести
копий (Q1), затем разрезать god-модуль на use-case сервисы с фасадом (A3), затем
убрать копипаст HTTP-handlers без смены URL (A4). Поведение мастера КП и
HTTP-контракт не меняются. Коммитов агент не делает.

Safety net после каждой задачи:

```bash
pytest tests/test_commercial_*_flow.py \
       tests/test_commercial_draft_append.py \
       tests/test_commercial_unpriced_plates_resolve.py \
       tests/test_commercial_ai_plates.py -q
```

## Architecture Decisions

- **Поведение > структура.** Stamp/partition/compose (workflow L97–303 +
  `stamp_order_line_identity`) переезжают одним куском. AI-пути по-прежнему
  replace всего `order_data` (без compose) — документированный статус-кво.
- **Facade стабилен до конца A4.** `commercial.py` продолжает звать
  `update_draft_marches` и т.д.; внутри — делегат в handler. Переименование
  публичных методов — вне скоупа.
- **Type preview services не трогаем.** `CommercialMarchService` и аналоги —
  парсеры/прайс. Handler их вызывает, не поглощает.
- **URL не схлопываем.** 32 пути остаются; A4 — shared runner, не один
  параметризованный route.
- **A5 вне плана.** Raw SQL СГП/графика — другая спека.
- **Волны с чекпоинтом.** После каждой фазы — pytest + стоп на ревью
  пользователя (рекомендация; см. Open Questions спеки).
- **Без коммитов.** Рабочее дерево для ревью пользователем.

## Dependency Graph

```
Q5 plate-resolve unify
    │
    ▼
Q1 ProductDraftSpec + ProductDraftHandler
   (create / update / AI / grades; Q4 metadata via config)
    │
    ├── stamp helpers still in workflow until A3
    │
    ▼
A3 extract identity + plate-resolve + lifecycle; workflow = facade ≤800
    │
    ▼
A4 shared HTTP runners; commercial.py ≤500; schemas untouched
```

Параллелить волны нельзя: Q1 опирается на нетронутый stamp; A3 двигает файлы
после того, как type-ветки уже в config; A4 тончает API, когда facade стабилен.

## Task List

### Phase Q5: Unify plate resolve

- [x] **Task 1: Параметризованный resolve wide/unpriced**
  - **Description:** Вынести общий каркас `resolve_wide_plates` /
    `resolve_unpriced_plates` и `_apply_*_decisions_to_batches` в одну
    реализацию с strategy (actions, matching, rewrite, metadata flags).
    Публичные методы остаются тонкими обёртками с теми же сигнатурами.
    На этом шаге код может жить ещё в `commercial_workflow_service.py`
    (переезд файла — Task 7 / A3), чтобы дифф Q5 был про поведение, не про
    импорты.
  - **Acceptance:**
    - [ ] Дублированного тела двух resolve нет
    - [ ] Обёртки `resolve_wide_plates` / `resolve_unpriced_plates` сохраняют
          сигнатуры и сообщения ошибок
    - [ ] Stamp plates на выходе тот же (`product_type=plates`, `line_id`)
  - **Verification:**
    - [ ] `pytest tests/test_commercial_web_flow.py -k "resolve_wide_plates" tests/test_commercial_unpriced_plates_resolve.py -q`
    - [ ] полный safety net (команда Overview)
  - **Files:** `app/services/commercial_workflow_service.py` (возможно новый
    helper в том же файле)
  - **Estimated scope:** M
  - **Dependencies:** None

### Checkpoint: после Task 1 (Q5)

- [ ] Safety net зелёный
- [ ] Пользователь может ревьюить дифф Q5 до Q1
- [ ] Стоп, если wide/unpriced разошлись сильнее, чем ожидали (~65% общего)

### Phase Q1: ProductDraftHandler + config

- [x] **Task 2: Config шести типов**
  - **Description:** Новый `product_draft_config.py`: `ProductDraftSpec` на
    plates / piles / marches / steps / bridge_piles / fbs. Поля: wizard step,
    `batches_key`, preview callable, metadata-builder, AI callable, тексты
    type-mismatch, флаги `needs_plate_ctx` / `has_grades`. Поведение не
    менять — только таблица, которую ещё никто не вызывает, либо вызвать из
    одного самого простого пути после Task 3. Предпочтительно сразу Task 3,
    если таблица без вызова бесполезна: тогда Tasks 2–3 делаются вместе, но
    ревьюитcя как «сначала таблица видна».
  - **Acceptance:**
    - [ ] Шесть спек; steps без grades; plates с `plate_order_ctx` и wide flag
    - [ ] Нет живых `if product_type ==` в config-файле
  - **Verification:** импорт модуля без ошибок; safety net ещё не обязан
    меняться, если Task 3 в той же сессии — прогон после Task 3
  - **Files:** `app/services/product_draft_config.py` (новый)
  - **Estimated scope:** S
  - **Dependencies:** Task 1

- [x] **Task 3: Handler для update (шесть ingest-пайплайнов)**
  - **Description:** `ProductDraftHandler.update` — канонический 12-шаговый
    pipeline. Шесть `update_draft_*` в workflow становятся делегатами
    `handler.update(draft_id, product_type=...)`. Stamp/partition/compose
    пока вызываются с workflow (или через тонкий delegate), **не копируются**.
  - **Acceptance:**
    - [ ] Тела шести `update_draft_*` не содержат копипаста каркаса
    - [ ] append / replace / merged_cycle_text / stamp reuse — как сейчас
    - [ ] plates по-прежнему отвергают piles/marches/steps drafts теми же текстами
  - **Verification:** safety net; особенно
    `test_update_*_draft_replace`, `test_commercial_draft_append.py`,
    `test_commercial_multi_append_flow.py`
  - **Files:** `app/services/product_draft_handler.py` (новый),
    `app/services/commercial_workflow_service.py`
  - **Estimated scope:** M (граница 5 файлов: config уже есть; не трогать
    schemas / commercial.py)
  - **Dependencies:** Task 2

- [x] **Task 4: Handler для create / AI / grades**
  - **Description:** Тот же config: `create_draft` без лестницы if;
    `_create_*_draft` удаляются; шесть `apply_ai_*` делегируют в
    `handler.apply_ai` (**без** partition/compose — полный replace order_data);
    четыре `update_draft_*_grades` — в `handler.update_grades` через
    `_current_cycle_lines` + stamp/compose `mode=append`, `merged_cycle_text=True`.
    Metadata builders draft_service подключаются из spec (Q4 absorbed).
  - **Acceptance:**
    - [ ] В workflow нет type-лестниц create/AI/grades
    - [ ] AI по-прежнему не compose'ит с другими типами
    - [ ] `test_bulk_grade_single_line_no_duplicate` зелёный
  - **Verification:** safety net +
    `pytest tests/test_commercial_march_flow.py::test_bulk_grade_single_line_no_duplicate tests/test_commercial_ai_plates.py -q`
  - **Files:** `product_draft_handler.py`, `product_draft_config.py`,
    `commercial_workflow_service.py`, при необходимости тонкие правки
    `commercial_draft_service.py` (без смены сигнатур stamp)
  - **Estimated scope:** M
  - **Dependencies:** Task 3

### Checkpoint: после Tasks 2–4 (Q1)

- [ ] Safety net зелёный
- [ ] Type-ветки create/update/AI/grades только в config
- [ ] Стоп на ревью: это самый рискованный дифф (stamp на всех путях)
- [ ] Ask first, если вскрылось расхождение типов

### Phase A3: Use-case сервисы + facade

- [x] **Task 5: Вынести order identity одним куском**
  - **Description:** Перенести `_stamp_order_data`, `_line_product_type`,
    `_line_is_sealed`, `_partition_order_by_product_type`,
    `_stamp_previous_for_product_update`, `_compose_order_data_for_product_update`,
    `_current_cycle_lines` (и при необходимости `_seal_unbatched_lines`) в
    `commercial_order_identity.py`. Handler и workflow вызывают этот модуль.
    Никакой новой семантики.
  - **Acceptance:**
    - [ ] В workflow нет второй копии `_line_product_type`
    - [ ] `test_partition_treats_untyped_legacy_mono_as_same_type` зелёный
  - **Verification:** safety net, особенно march partition + draft_append MNA-102
  - **Files:** `app/services/commercial_order_identity.py` (новый),
    `product_draft_handler.py`, `commercial_workflow_service.py`
  - **Estimated scope:** M
  - **Dependencies:** Task 4

- [x] **Task 6: Вынести draft lifecycle**
  - **Description:** `calculate_draft`, `update_draft_meta`, `get_draft_details`,
    `get_draft_breakdown`, `hydrate_draft_from_saved_kp`, `generate_files`,
    `save_offer`, `save_draft` → `commercial_draft_lifecycle.py`. Facade
    оставляет те же публичные методы. Append-cycle (`start_append_cycle`,
    undo, delete line) можно оставить на facade или унести вместе с
    `_seal_unbatched_lines` — выбрать меньший дифф, не дробить seal.
  - **Acceptance:**
    - [ ] Facade не содержит тел calculate/save/export
    - [ ] Сигнатуры публичных методов неизменны
  - **Verification:** safety net + calculate/save тесты в каждом `*_flow.py`
  - **Files:** `app/services/commercial_draft_lifecycle.py` (новый),
    `commercial_workflow_service.py`
  - **Estimated scope:** M
  - **Dependencies:** Task 5

- [x] **Task 7: Вынести plate-resolve модуль + подчистить фасад**
  - **Description:** Перенести unified resolve из Q5 в
    `commercial_plate_resolve.py`. Workflow — делегат. Замерить `wc -l`
    facade: цель ≤ 800. Если больше 800 из-за делегатов и wizard glue —
    ask first, не сжимать читаемость.
  - **Acceptance:**
    - [ ] `commercial_workflow_service.py` ≤ 800 **или** зафиксировано
          исключение в спеке после ask-first
    - [ ] Нет type-веток ingest в facade
  - **Verification:** safety net; `wc -l app/services/commercial_workflow_service.py`
  - **Files:** `app/services/commercial_plate_resolve.py` (новый),
    `commercial_workflow_service.py`
  - **Estimated scope:** S–M
  - **Dependencies:** Task 1, Task 6

### Checkpoint: после Tasks 5–7 (A3)

- [ ] Safety net зелёный
- [ ] Фасад тонкий; identity в одном модуле
- [ ] Стоп на ревью структуры файлов до правки HTTP-слоя

### Phase A4: Thin controllers

- [x] **Task 8: Shared runners для product HTTP**
  - **Description:** В `commercial.py` — `_run_product_update` / `_run_product_ai`
    / `_run_product_grades` (OCR prep, type-mismatch plates extra
    `PlateParseError`, одинаковый 404/400/500). Каждый `@router.*` путь
    остаётся. Не мержить четыре grades request schema в OpenAPI.
  - **Acceptance:**
    - [ ] 32 пути (method + path) на месте
    - [ ] `commercial.py` ≤ 500 строк
    - [ ] `app/schemas/commercial.py` без изменения полей
  - **Verification:**
    - [ ] `pytest tests/test_commercial_*.py -q`
    - [ ] `wc -l app/api/v1/endpoints/commercial.py`
  - **Files:** `app/api/v1/endpoints/commercial.py`
  - **Estimated scope:** M
  - **Dependencies:** Task 7

- [x] **Task 9: Унести orchestration из толстых handlers**
  - **Description:** `generate_preview` — persist draft + сборка dict в сервис
    (те же ключи JSON, без нового `response_model`). `download_generated_file` —
    path whitelist в workflow, те же 400/404. Не расширять скоуп на `/parse`,
    если он уже достаточно тонкий.
  - **Acceptance:**
    - [ ] В `commercial.py` нет прямой записи `DraftStore().save_preview` в
          generate-preview
    - [ ] Тесты download / from-form / web_flow зелёные
  - **Verification:** `pytest tests/test_commercial_web_flow.py tests/test_commercial_*.py -q`
  - **Files:** `commercial.py`, `commercial_workflow_service.py` и/или
    `commercial_service.py` / lifecycle
  - **Estimated scope:** S
  - **Dependencies:** Task 8

- [x] **Task 10: Статусы в отчёте аудита**
  - **Description:** Пометить в `2026-08-28-full-project-audit.md` Q5/Q1/A3/A4
    как resolved со ссылкой на спеку. A5 оставить открытым (follow-up).
    P0-статусы не переписывать. Финальное сообщение: список файлов, команда
    pytest, напоминание про коммит.
  - **Acceptance:**
    - [ ] Матрица аудита обновлена только по P1-находкам
    - [ ] Success criteria спеки отмечены
  - **Verification:** прочитать обновлённые секции отчёта
  - **Files:** `ai_docs/develop/audits/2026-08-28-full-project-audit.md`,
    спека (чекбоксы Success Criteria)
  - **Estimated scope:** XS
  - **Dependencies:** Task 9

### Checkpoint: Complete

- [x] Все Success Criteria спеки выполнены
      (ручной smoke — опционально; 3 pre-existing web_flow не в скоупе P1)
- [x] `pytest tests/test_commercial_*.py` — 333 passed, 3 pre-existing
- [x] Рабочее дерево готово к ревью; агент не коммитил

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Q1 сломает stamp/append/bulk-grade (строка дублируется) | High | Волны; `test_bulk_grade_single_line_no_duplicate` + partition + draft_append после каждого шага Q1 |
| Копипаст типов уже разъехался; config «выровняет» поведение | High | Boundary ask first; не молча унифицировать сообщения/guards |
| AI случайно начнёт compose с другими типами | High | Явный тест-якорь текущей семантики; handler.apply_ai без partition |
| Вынос identity разорвёт `self.` и циклические импорты | Med | identity — чистые функции / маленький класс без workflow import |
| Фасад не ужимается до 800 без каши делегатов | Low | Ask first; лучше 900 читаемых |
| A4 схлопнет OpenAPI grades schemas | Med | Четыре отдельных request class и пути; только общий runner |
| Pre-existing 12 pytest failures примут за регрессию P1 | Low | Гонять commercial-набор, не «чинить заодно» весь pytest |
| Незакоммиченный stamp-фикс смешается с P1 в одном дереве | Med | Волны; в сообщении чекпоинта явно: «stamp-фикс уже был, это дифф волны N» |

## Parallelization

Волны **строго последовательны**. Внутри волны не параллелить handler и
identity. Документацию аудита (Task 10) — только в конце.

## Open Questions

Q-A5 и Q-delivery закрыты пользователем 2026-08-28:

1. **A5** — исключён (follow-up спека).
2. **Поставка** — 4 волны с чекпоинтом после Q5, Q1, A3, A4.
3. **Целевой размер фасада** — 800; отклонение только через ask first.

IMPLEMENT не стартует, пока пользователь не approve'ит спеку.
