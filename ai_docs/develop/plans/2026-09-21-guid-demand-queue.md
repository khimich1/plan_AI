# План: Очередь спроса GUID из архивного КП

**Дата:** 2026-09-21  
**Статус:** реализовано (IMPLEMENT ✅)  
**Идея:** `ai_docs/ideas/guid-demand-queue.md`  
**Спека:** `ai_docs/specs/guid-demand-queue.md`  
**Родитель:** раунд 2 `ai_docs/develop/plans/2026-09-17-guid-price-sync-round2.md`

## Overview

При архиве КП резолвер уже считает `missing`. Сейчас это уходит только в колокольчик, а `/prices` рисует весь справочник дырок. Нужна таблица спроса `guid_demand` в `pb.db`: писать при архиве, читать в «Завести в 1С» / «Дубли GUID», закрывать sweep-ом после импорта 1С и выбора GUID. «Ввести цену» и Excel-сводка не трогаем.

## Architecture Decisions

- **Спрос ≠ каталог.** `collect_guid_queue` остаётся для `scripts/build_guid_queue.py`. Веб `list_tasks` читает `guid_demand`.
- **Один вызов на архив.** `notify_kp_guid_missing` после `check_invoice_guids` пишет спрос (commit в `pb.db`) и потом колокольчик. Экономистов нет — спрос всё равно есть.
- **PK `(product_kind, mark, field)`.** `mark` = `OrderLine.mark` из КП. Смена create↔duplicate — не UPDATE PK: закрыть старую строку, upsert новую с тем же `kp_ids`.
- **Кандидаты дубля** по-прежнему из `duplicate_candidates` / `prays_plity`; фильтр по open-спросу `field=duplicate`. Свая «у»: ключ каталога = `strip_trailing_u_suffix(mark)`. Нет кандидатов — карточку не рисуем, ждём sweep.
- **Sweep** после commit `import-1c` и после успешного `duplicates/resolve`. По каждой open-строке снова `check_invoice_guids([OrderLine])`.
- **`kp_ids` аддитивно** в существующий envelope tasks/duplicates. Без новых роутов.
- **`ensure_schema` при record/list/sweep** — как у соседних таблиц; отдельной миграции живого `pb.db` нет.
- Не убивать `./run+logs.sh`. Коммиты — по просьбе. Новых npm/pip нет.

## Tasks Overview

1. **GDQ-001** таблица `guid_demand` + record/list `(feat-be)` — dependsOn: []
2. **GDQ-002** sweep `(feat-be)` — dependsOn: [GDQ-001]
3. **GDQ-003** запись спроса в `notify_kp_guid_missing` `(feat-be)` — dependsOn: [GDQ-001]
4. **GDQ-004** `GET /tasks` и `/duplicates` из спроса + `kp_ids` `(api)` — dependsOn: [GDQ-001]
5. **GDQ-005** sweep на импорт 1С и resolve дубля `(feat-be)` — dependsOn: [GDQ-002, GDQ-004]
6. **GDQ-006** UI «КП №…» `(ui)` — dependsOn: [GDQ-004]
7. **GDQ-007** focused verify + статус спеки `(chore)` — dependsOn: [GDQ-005, GDQ-006]

## Dependencies Graph

```
GDQ-001 ──► GDQ-002 ──┐
     │                ├──► GDQ-005 ──┐
     ├──► GDQ-003     │              ├──► GDQ-007
     └──► GDQ-004 ────┘              │
            └──► GDQ-006 ────────────┘
```

`GDQ-002` ∥ `GDQ-003` ∥ `GDQ-004` после 001.  
`GDQ-006` ∥ `GDQ-005` после 004.

---

## Task List

### Phase 1 — Хранилище

- [x] **GDQ-001: Таблица `guid_demand` + record**
  - **Description:** `core/guid_demand.py`: `ensure_schema()`, dataclass строки, `record_guid_demand(conn, kp_id, missing)`, `list_open()`, маппинг reason→field (`missing_u`→`guid_1c_u`, `missing`→`guid_1c`, dup/ambiguous→`duplicate`). Повтор того же КП не дублирует `kp_id`. Resolved + снова missing → `open`, `resolved_at=NULL`.
  - **Acceptance:**
    - [x] Повторный `ensure_schema` не падает
    - [x] Record двух КП на одну марку → одна open-строка, `kp_ids=[12,18]`
    - [x] Reopen после resolved
  - **Verify:** `pytest tests/test_guid_demand.py -q`
  - **Dependencies:** —
  - **Files:** `core/guid_demand.py`, `tests/test_guid_demand.py`
  - **Scope:** S

- [x] **GDQ-002: Sweep**
  - **Description:** `sweep_guid_demand(conn)`: для каждой open — `check_invoice_guids([OrderLine(kind, mark)])`. Пустой missing → `resolved`. Смена категории create↔duplicate → resolve старый `(kind, mark, field)`, upsert новый field, те же `kp_ids`. Чужие open не трогать.
  - **Acceptance:**
    - [x] После upsert GUID в `nomenclature_guid` sweep закрывает create
    - [x] Create→dup обновляет field, state=open
    - [x] Соседняя open-марка остаётся
  - **Verify:** `pytest tests/test_guid_demand.py -q`
  - **Dependencies:** GDQ-001
  - **Files:** `core/guid_demand.py`, `tests/test_guid_demand.py`
  - **Scope:** S

### Checkpoint: Foundation

- [x] `pytest tests/test_guid_demand.py -q` зелёный
- [x] Каталог `tests/test_guid_queue.py` не сломан (модуль ещё не подключён к вебу)

### Phase 2 — Проводки

- [x] **GDQ-003: Архив пишет спрос**
  - **Description:** В `notify_kp_guid_missing` после gate, до закрытия `pb` conn: `record_guid_demand` + `commit`. Нет экономистов — спрос есть, колокольчик как сейчас (warning). Дырок нет — record не зовём / пустой missing no-op.
  - **Acceptance:**
    - [x] Архив с дыркой → open в `guid_demand` даже без economist
    - [x] Архив без дырок → таблица пуста
    - [x] Колокольчик A5 без регресса
  - **Verify:** `pytest tests/test_notifications.py -q`
  - **Dependencies:** GDQ-001
  - **Files:** `app/services/kp_guid_notify.py`, `tests/test_notifications.py`
  - **Scope:** S

- [x] **GDQ-004: API очередей из спроса**
  - **Description:** `Create1cTaskOut` / `DuplicateTaskOut` + `kp_ids: list[int]`. `NomenclatureQueueService.list_tasks`: `to_create_1c` из open `field!=duplicate`; `duplicates` = каталожные кандидаты ∩ open `field=duplicate` (сваи «у» — strip суффикса); `to_price` без изменений. `GET /duplicates` тот же фильтр. Переписать `test_guid_gate_api.test_tasks_*`: се pal missing `С30.30-3` **не** попадает в `to_create_1c` без demand (S2).
  - **Acceptance:**
    - [x] Пустой спрос + дырки справочника → `to_create_1c=[]`, `duplicates=[]`, `to_price` жив
    - [x] Open create с `kp_ids` виден economist, manager 403
    - [x] Open dup без кандидатов → не в `duplicates`
  - **Verify:** `pytest tests/test_guid_gate_api.py tests/test_guid_queue.py -q`
  - **Dependencies:** GDQ-001
  - **Files:** `app/schemas/nomenclature.py`, `app/services/nomenclature_queue_service.py`, `tests/test_guid_gate_api.py`, плюс хелпер списка в `core/guid_demand.py` если раздувается сервис
  - **Scope:** M

### Checkpoint: Core write/read

- [x] Архив (через notify) → `GET /tasks` показывает марку и `kp_ids`
- [x] Справочник без спроса → пустые 🏭/⚠️
- [x] `test_guid_queue.py` зелёный

### Phase 3 — Закрытие и UI

- [x] **GDQ-005: Sweep на импорт и resolve**
  - **Description:** После `conn.commit()` в `_sync_and_commit` (тот же conn, до close — sweep затем commit). После успешного `resolve_duplicate` — sweep. Тест: demand open → partial-импорт GUID → tasks без марки. Resolve дубля → dup-спрос resolved.
  - **Acceptance:**
    - [x] Partial import закрывает одну марку, соседний demand open
    - [x] `POST /duplicates/resolve` закрывает спрос этой пары
    - [x] 409 «кандидат исчез» не режет чужой спрос
  - **Verify:** `pytest tests/test_nomenclature_import_api.py tests/test_duplicate_resolve.py -q`
  - **Dependencies:** GDQ-002, GDQ-004
  - **Files:** `app/services/nomenclature_import_service.py`, `app/services/nomenclature_queue_service.py`, `tests/test_nomenclature_import_api.py`, `tests/test_duplicate_resolve.py`
  - **Scope:** M

- [x] **GDQ-006: Подпись КП в UI**
  - **Description:** Тип `kp_ids?: number[]`. `TasksSection`: `… — КП №12, №18` после hint. `DuplicatesSection`: номера на карточке. Пустой спрос → «задач нет» / «дублей нет».
  - **Acceptance:**
    - [x] Vitest: строка create с `kp_ids` содержит «КП №12»
    - [x] Карточка дубля с `kp_ids` содержит «КП №12»
    - [x] Мок пустого спроса при непустом каталоге не рисует список заведения
  - **Verify:** `cd frontend && npm run test -- src/features/price-desk`
  - **Dependencies:** GDQ-004
  - **Files:** `frontend/src/features/nomenclature/types/nomenclature.ts`, `frontend/src/features/price-desk/components/TasksSection.tsx`, `frontend/src/features/price-desk/components/DuplicatesSection.tsx`, `frontend/src/features/price-desk/components/PricesView.test.tsx`
  - **Scope:** M

### Checkpoint: Complete

- [x] **GDQ-007: Focused verify**
  - **Description:** Прогон команд из спеки. Спека: SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅. Идея: статус «к реализации закрыта». Не коммитить.
  - **Acceptance:**
    - [x] pytest-набор из спеки зелёный
    - [x] vitest price-desk + nomenclature зелёный
    - [x] `npm run typecheck` зелёный
  - **Verify:** команды из `ai_docs/specs/guid-demand-queue.md` § Commands
  - **Dependencies:** GDQ-005, GDQ-006
  - **Files:** `ai_docs/specs/guid-demand-queue.md`, `ai_docs/ideas/guid-demand-queue.md`
  - **Scope:** S

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| `test_guid_gate_api` сейчас ждёт каталожную дырку в `/tasks` | High | Явно переписать в GDQ-004 (это и есть S2) |
| Смена `field` ломает PK | Med | Resolve старой + upsert новой, не UPDATE PK |
| Свая «у» не матчится с `duplicate_candidates.mark_norm` | Med | strip суффикса только для join кандидатов |
| Sweep на полном импорте по всем open | Low | Open-спрос маленький (только архивные дырки) |
| Живой `/prices` после выката опустеет | Low | Ожидаемо (D-backfill). Новые архивы наполнят |

## Out of this plan

- «Ввести цену» по спросу, backfill архива, один inbox, пинг менеджеру, ворота счёта, Excel→спрос.

## Open Questions

Нет. Допущения спеки locked.
