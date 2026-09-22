# Spec: Очередь спроса GUID из архивного КП

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-09-21  
**One-pager**: [../ideas/guid-demand-queue.md](../ideas/guid-demand-queue.md)  
**Plan**: [../develop/plans/2026-09-21-guid-demand-queue.md](../develop/plans/2026-09-21-guid-demand-queue.md)  
**Родитель**: [guid-price-sync-round2.md](./guid-price-sync-round2.md) (колокольчик при архиве есть; список на `/prices` — инвентаризация каталога)

## Objective

**Проблема.** На вкладке «Прайсы и 1С» секции «Завести в 1С» и «Дубли GUID» показывают весь справочник дырок: все сваи без карточки «у», все пары GUID. Экономист не видит, что *заказали*. Архив КП шлёт колокольчик, но задачу в список не ставит. Новая марка (нет строки в `nomenclature_guid`) в списке не появляется.

**Цель.** Список на `/prices` — только спрос из КП, сохранённых в архив: нет GUID / нет «у» / невыбранная пара. Колокольчик как сейчас. Каталог дырок на вкладке больше не показывается.

**Пользователь:** экономист (и admin) на `/prices`. Триггер — менеджер, который сохраняет КП в архив.

**Успех:** после архива КП с дыркой экономист видит колокольчик и короткую строку этой марки (с номерами КП); сваи без «у», которых никто не заказывал, в списке нет; после выгрузки 1С или выбора GUID строка сгорает.

---

## ASSUMPTIONS I'M MAKING

1. **Порог спроса — архив**, не «в работе» и не шаг 3. Уже так у колокольчика.
2. **Старый архив не сканируем.** Задачи только с новых сохранений в архив после выката.
3. **Таблица в `pb.db`**, рядом с GUID. Уведомления остаются в `plita.db`.
4. **Повторный архив:** дырка снова есть → задача `open` (идемпотентный upsert, дописать `kp_id`). Дырка закрыта → строку спроса не трогаем.
5. **Плиты входят.** Резолвер уже кладёт их в `missing`.
6. **«Ввести цену» не трогаем.** Это GUID без заводской цены, не заведение карточки.
7. **Два раздела UI**, не один inbox. Завести = файл 1С; пара = клик на месте.
8. **Excel-скрипт `scripts/build_guid_queue.py`** остаётся каталожной сводкой (`collect_guid_queue`). Живой веб читает спрос.
9. **Закрытие** — повторный прогон `check_invoice_guids` по открытым строкам спроса после импорта 1С и после `duplicates/resolve`. Ручной «готово» нет.
10. **Выбор GUID по паре — на все будущие счета**, не «только это КП».
11. **Без новых npm/pip.** Коммиты — по просьбе. `./run+logs.sh` не убивать.

→ Correct me now or these are locked for PLAN.

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| **D-demand** | Источник списка | Только открытые строки `guid_demand`, не итерация `nomenclature_guid` |
| **D-trigger** | Когда писать | Сохранение КП со статусом «в архиве» (мастер + `create_offer`) |
| **D-backfill** | Старый архив | Нет |
| **D-key** | Ключ задачи | `(product_kind, mark, field)` где `mark` = `OrderLine.mark` как в КП |
| **D-field** | `field` | `guid_1c` / `guid_1c_u` → «Завести в 1С»; `duplicate` → «Дубли GUID» |
| **D-kp** | Номера КП | JSON-массив `kp_ids`, в UI: `Свая · С 70.35-9у — КП №12, №18` |
| **D-notify** | Колокольчик | Без смены контракта `kp_guid_missing`; запись спроса в том же вызове, даже если экономистов нет |
| **D-close** | Сгорание | `sweep`: по каждой open-строке `check_invoice_guids([OrderLine])`; пустой `missing` → `resolved` |
| **D-reason-shift** | Была «нет GUID», стала пара | Не закрывать: обновить `field`/`reason`/`hint`, оставить `open` |
| **D-reopen** | GUID пропал, снова архив | `state=open`, `resolved_at=NULL` |
| **D-price** | «Ввести цену» | Как сейчас, каталог `price_queue` |
| **D-excel** | `build_guid_queue.py` | Не переключать на спрос |
| **D-unknown** | Марки нет в справочнике | Строка спроса всё равно есть (сейчас только колокольчик) |

---

## User Stories

- Как **менеджер**, архивирую КП со сваей без GUID — мастер не блокируется (как сейчас).
- Как **экономист**, получаю «КП №12: изделия без GUID — С 70.35-9у», открываю `/prices` и вижу **только эту** строку в «Завести в 1С», а не все сваи без «у».
- Как **экономист**, вижу пару GUID только если она встретилась в архивном КП; выбираю GUID — строка сгорает.
- Как **экономист**, загружаю универсальный отчёт с отбором на заведённую карточку — задача сгорает, остальные открытые не трогаются.
- Как **экономист**, если задач нет — «задач нет» / «дублей нет», даже когда справочник полный дырок.

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python, FastAPI, SQLite `pb.db` + `plita.db` |
| Domain | `core/guid_gate.py` (резолвер), новый `core/guid_demand.py` |
| API | существующие `GET /nomenclature/tasks`, `GET /nomenclature/duplicates` — тот же envelope, плюс `kp_ids` |
| Frontend | `TasksSection`, `DuplicatesSection` — подпись КП |
| Тесты | pytest `tests/`, vitest `frontend/src/features/price-desk` |

Новых пакетов нет.

## Commands

```
# Backend
pytest tests/test_guid_demand.py tests/test_notifications.py \
  tests/test_guid_gate_api.py tests/test_nomenclature_import_api.py \
  tests/test_duplicate_resolve.py tests/test_offers_service.py -q

# Frontend
cd frontend && npm run test -- src/features/price-desk src/features/nomenclature
cd frontend && npm run typecheck

# Dev: не убивать ./run+logs.sh
```

## Project Structure

```
core/guid_demand.py                         → таблица guid_demand, record + sweep
core/guid_gate.py                           → без смены контракта резолвера
app/services/kp_guid_notify.py              → после gate: record demand, затем колокольчик
app/services/nomenclature_queue_service.py  → list_tasks: create/dup из спроса; to_price как сейчас
app/services/nomenclature_import_service.py → после sync: sweep_guid_demand
app/schemas/nomenclature.py                 → kp_ids на Create1cTaskOut / DuplicateTaskOut
frontend/.../nomenclature.ts                → kp_ids?: number[]
frontend/.../TasksSection.tsx               → «КП №…» в строке
frontend/.../DuplicatesSection.tsx          → «КП №…» на карточке
core/guid_queue.py                          → каталог; веб его больше не использует для 🏭/⚠️
scripts/build_guid_queue.py                 → без изменений (каталог)
tests/test_guid_demand.py                   → unit record/sweep/list
```

## Code Style

Слои: SQL в `core/guid_demand.py` (`ensure_schema` идемпотентен, как `nomenclature_guid`). Оркестрация в уже существующих сервисах. Русские `detail`. Минимальный diff.

Схема:

```sql
CREATE TABLE IF NOT EXISTS guid_demand (
    product_kind TEXT NOT NULL,
    mark TEXT NOT NULL,
    field TEXT NOT NULL,          -- guid_1c | guid_1c_u | duplicate
    reason TEXT NOT NULL,
    hint TEXT NOT NULL,
    state TEXT NOT NULL,          -- open | resolved
    kp_ids_json TEXT NOT NULL,    -- [12, 18]
    opened_at TEXT NOT NULL,
    resolved_at TEXT,
    PRIMARY KEY (product_kind, mark, field)
);
```

Запись из `GateReport.missing`:

```
REASON_MISSING_U          → field=guid_1c_u, секция «Завести»
REASON_MISSING            → field=guid_1c,   секция «Завести»
REASON_DUP / AMBIGUOUS    → field=duplicate, секция «Дубли»
```

`list_tasks.to_create_1c` — open, `field != duplicate`.  
`list_tasks.duplicates` / `GET /duplicates` — open `field=duplicate`, кандидаты как сейчас из `duplicate_candidates` / `prays_plity`, фильтр по `(product_kind, key)`; для сваи «у» ключ каталога — марка без суффикса (`strip_trailing_u_suffix`). Нет кандидатов (уже сняли дубль в 1С, а sweep ещё не закрыл) — карточку не рисуем, строка ждёт sweep.

Пример API (аддитивно):

```python
Create1cTaskOut(
    product_kind="pile",
    mark="С 70.35-9у",
    hint="заведите карточку «у» в 1С и загрузите отчёт",
    field="guid_1c_u",
    kp_ids=[12, 18],
)
```

## Design

### Триггер

`notify_kp_guid_missing` уже открывает `pb.db` и гоняет `check_invoice_guids`. Сразу после отчёта: `record_guid_demand(conn, kp_id, report.missing)`. Потом фан-аут колокольчика. Нет экономистов → warning как сейчас, спрос всё равно записан.

Точки входа не плодим: мастер (`commercial_draft_lifecycle`, статус «в архиве») и `offers_service.create_offer` уже зовут notify.

### Sweep

После успешного `import-1c` (commit sync) и после успешного `duplicates/resolve`:

```
for row in list_open(conn):
    missing = check_invoice_guids([OrderLine(row.product_kind, row.mark)], conn).missing
    if not missing:
        set_resolved(row)
    elif category(missing[0]) != category(row.field):
        update field/reason/hint  # create ↔ duplicate
    # иначе оставить
```

Импорт с отбором закрывает только совпавшие марки.

### Повтор и дедуп

- Тот же КП, та же марка, та же `field`: `kp_ids` без дубля номера, `state` остаётся/становится `open`.
- Другое КП, та же марка: дописать `kp_id`.
- Колокольчик: дедуп (КП, марка) без изменений (A5 раунда 2).

### UI

- «Завести в 1С»: `{вид} · {марка} — {hint} — КП №{ids}`. Read-only, как сейчас.
- «Дубли GUID»: заголовок марки + КП №; радио как сейчас.
- «Ввести цену», загрузка прайса, выгрузка 1С — без изменений.
- Пустые тексты те же.

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit demand | ensure_schema идемпотентен; record из GateReport; повтор КП не дублирует kp_id; reopen после resolved |
| Unit sweep | GUID появился → resolved; причина сменилась create→dup → field обновлён, open; несвязанные open не трогаем |
| Unit list | справочник полный дырок, demand пуст → `to_create_1c=[]`, `duplicates=[]`; `to_price` как был |
| Unit unknown | марки нет в `nomenclature_guid` → строка create с kp_ids |
| Service | архив с дыркой → demand + notify; архив без дырок → тихо и без demand |
| API | `GET /tasks` отдаёт только спрос; manager 403; `kp_ids` в JSON |
| Import | partial xlsx закрывает одну марку спроса |
| RTL | TasksSection/DuplicatesSection рендерят «КП №12»; пустой спрос → «задач нет» при непустом каталоге в моке |
| Regress | `test_guid_queue.py` (каталог) зелёный; колокольчик A5; resolve дубля 409 |

## Boundaries

- **Always:** реюз `check_invoice_guids`; идемпотентный `ensure_schema`; русские сообщения; тесты на record/sweep/list; спрос пишется даже без экономистов.
- **Ask first:** бэкап/миграция живого `pb.db` вне `ensure_schema`; смена `collect_guid_queue` / Excel-скрипта на спрос; трогать «Ввести цену»; backfill архива.
- **Never:** блок архива КП; автовыбор GUID; новые зависимости; коммит без просьбы; ГСМ / production layout; ручной ввод GUID; сканирование старого архива при старте.

## Success Criteria

| # | Критерий |
|---|----------|
| S1 | Архив КП с маркой без GUID → open-строка спроса + колокольчик economist; вкладка показывает эту марку и `kp_ids` |
| S2 | Справочник свай без «у» при пустом спросе → «Завести в 1С» пуст |
| S3 | Дубль в каталоге, но не в архивном КП → «Дубли GUID» пуст |
| S4 | Архив КП с невыбранным дублем → строка в «Дубли GUID» с кандидатами и `kp_ids`; выбор закрывает спрос |
| S5 | Неизвестная марке (нет `nomenclature_guid`) → строка в «Завести в 1С» |
| S6 | Partial-импорт 1С, GUID дописан → эта задача `resolved`, список без неё |
| S7 | Повторный архив того же КП не плодит строки и не дублирует колокольчик (A5); новое КП с той же маркой дописывает номер |
| S8 | «Ввести цену», импорт .xls/.xlsx, роли вкладки — без регресса |
| S9 | Focused pytest + vitest + typecheck зелёные |

## Out of Scope

- «Ввести цену» по спросу из КП
- Backfill текущего архива
- Один inbox вместо двух секций
- Пинг менеджеру «можно выставлять счёт»
- Кнопка «В 1С» / сущность счёта
- E-mail / Telegram
- Перевод Excel-очереди на спрос

## Open Questions

Внутренних нет — закрыты в one-pager.
