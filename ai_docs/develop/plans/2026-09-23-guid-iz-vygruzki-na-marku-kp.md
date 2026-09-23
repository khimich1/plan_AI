# План: GUID из выгрузки 1С на марку КП

**Дата:** 2026-09-23  
**Статус:** выполнен  
**Идея:** `ai_docs/ideas/guid-iz-vygruzki-na-marku-kp.md`  
**Спека:** `ai_docs/specs/guid-iz-vygruzki-na-marku-kp.md`  
**Родитель:** `ai_docs/develop/plans/2026-09-21-guid-demand-queue.md`

## Overview

Загрузка универсального отчёта пишет GUID только в марки, которые уже лежат в `nomenclature_guid`. Три сваи из файла экономиста есть в `pile_prices` и в открытом `guid_demand`, но строки справочника нет, поэтому они уходят в «Ввести цену». Нужно тем же сопоставлением имени создать строку и записать единственный GUID, не трогая прайс. Договорная марка (в прайсе нет, спрос из архива есть) идёт той же дверью. Два GUID не выбираем. Плашка шага 3 должна погаснуть сразу после импорта.

## Architecture Decisions

- **Врезка в `sync_pricelist`, не новый эндпоинт.** Контракт `POST /import-1c` не меняется. `partial` и `disappeared` остаются как есть.
- **Индекс прайса и спроса строится теми же кандидатами**, что `_candidates_for_mark` / `_build_name_index`. «Сваи С 120.35-12» → `С120.35-12`, `C` ≡ `С`, суффикс «у» → поле `guid_1c_u`. Сосед `С120.35-13` не кандидат этого имени.
- **Источники допуска, по группе файла:** distinct `mark` из `pile_prices` / `bridge_pile_prices` / `fbs_prices` / `march_prices` / `step_prices`, плюс марки открытого `guid_demand` с `field` `guid_1c` или `guid_1c_u`. Черновик не читаем. Плиты в `nomenclature_guid` не создаём.
- **Одно имя → одна марка.** Несколько разных марок на одно нормализованное имя — как дубль: GUID не пишем.
- **Применение через существующий `_apply_mark`.** Сейчас цикл идёт только по строкам справочника, поэтому план на новую марку никуда не пишется. Марку из прайса или спроса, которой нет в `db_rows`, добавляем как пустую строку в этот цикл. Один GUID → `upsert` со статусом `auto`. Два GUID → статус `ambiguous`, `guid_1c` пустой, кандидаты в `duplicate_candidates` на эту марку. Дальше уже написанный `sweep_guid_demand` закрывает «Завести» или переводит в «Дубли».
- **Не класть в `unmatched_1c`.** Иначе `nomenclature_import_service._fill_price_queue` откроет «Ввести цену» при живой цене в прайсе. Чужое имя (нет ни в прайсе, ни в спросе) по-прежнему unmatched.
- **`manual` не затирать.** Это уже делает `_apply_mark`. Новую строку создаём только когда строки не было.
- **Плашка шага 3.** `useImport1cMutation` кроме `nomenclature` сбрасывает `["commercial", "guid-check"]`. Отдельной записи в черновик нет: проверка читает справочник. `staleTime` 15 с на этот сброс не полагаемся.
- Не убивать `./run+logs.sh`. Коммиты по просьбе. Новых пакетов нет.

## Risks

| Риск | Как гасим |
|------|-----------|
| GUID прилип к соседней марке | Тест: в прайсе есть `С120.35-12` и `С120.35-13`, в файле только первая |
| Договорная цена уехала в прайс | Тест: число строк прайс-таблицы до и после одинаковое |
| «Ввести цену» на марку, которая уже в прайсе | Тест сервиса: после импорта `price_queue` по этому GUID пуст |
| Два GUID тихо записались | Тест: справочник без нового GUID, спрос после sweep в `field=duplicate` |
| Пустая строка справочника на имя вне прайса и вне КП | Тест: такой строки нет, есть unmatched и `price_queue` |

## Tasks Overview

1. **GKP-001** единственный GUID на марку из прайса `(feat-be)` — dependsOn: []
2. **GKP-002** договорная марка из открытого спроса `(feat-be)` — dependsOn: [GKP-001]
3. **GKP-003** два GUID → дубль, чужое имя → «Ввести цену» `(feat-be)` — dependsOn: [GKP-001]
4. **GKP-004** сброс проверки GUID на шаге 3 `(ui)` — dependsOn: []
5. **GKP-005** focused verify + статус спеки `(chore)` — dependsOn: [GKP-002, GKP-003, GKP-004]

## Dependencies Graph

```
GKP-001 ──► GKP-002 ──┐
     └──► GKP-003 ────┼──► GKP-005
GKP-004 ──────────────┘
```

`GKP-004` параллельно с бэкендом. `GKP-002` и `GKP-003` оба после 001, между собой независимы.

---

## Task List

### Phase 1 — Запись GUID

- [x] **GKP-001: Единственный GUID на марку из прайса**
  - **Description:** В `sync_pricelist`, если у группы строк файла один GUID и в справочнике цели нет, найти марку в прайс-таблице группы тем же индексом кандидатов. Пустую марку прогнать через `_apply_mark` (статус `auto`, нужное поле GUID). В `unmatched` не класть. `matched_marks` пополнить, чтобы частичный отчёт не считал её исчезнувшей. Прайс-таблицу не писать.
  - **Acceptance:**
    - [x] `С120.35-12` есть в `pile_prices`, в справочнике нет, файл «Сваи С 120.35-12» → строка `auto` с этим GUID
    - [x] `С120.35-13` в файле нет → строки справочника для неё нет
    - [x] Число строк `pile_prices` не изменилось
    - [x] Импорт через сервис не открывает `price_queue` на этот GUID
    - [x] Повтор того же файла не создаёт вторую строку; `manual` с другим GUID не затирается
  - **Verify:** `pytest tests/test_nomenclature_import.py tests/test_nomenclature_import_api.py -q`
  - **Dependencies:** —
  - **Files:** `core/nomenclature_sync.py`, `tests/test_nomenclature_import.py`, `tests/test_nomenclature_import_api.py`
  - **Scope:** M

- [x] **GKP-002: Договорная марка из архивного спроса**
  - **Description:** Тот же индекс принимает марку, которой нет в прайсе, если она есть в открытом `guid_demand` этой группы (`guid_1c` / `guid_1c_u`). GUID пишется так же. Прайс не растёт, `price_queue` не открывается. После импорта существующий `sweep_guid_demand` закрывает «Завести в 1С».
  - **Acceptance:**
    - [x] Марки нет в `pile_prices`, open-спрос есть, один GUID в файле → справочник `auto`, прайс не вырос
    - [x] `price_queue` по этому GUID пуст
    - [x] Sweep переводит спрос в `resolved`
    - [x] Черновик без строки спроса марку не создаёт
  - **Verify:** `pytest tests/test_nomenclature_import.py tests/test_guid_demand.py -q`
  - **Dependencies:** GKP-001
  - **Files:** `core/nomenclature_sync.py`, `tests/test_nomenclature_import.py`
  - **Scope:** S

### Checkpoint: Backend doors

- [x] `pytest tests/test_nomenclature_import.py tests/test_nomenclature_import_api.py tests/test_guid_demand.py tests/test_commercial_draft_oneoff_price.py -q`
- [x] Прайс-таблицы в этих тестах только читаются

### Phase 2 — Дубли и плашка

- [x] **GKP-003: Два GUID и чужое имя**
  - **Description:** Если на одно имя два GUID и оно сходится с одной маркой прайса или открытого спроса, GUID не писать. В `_record_ambiguous` передать эту марку, чтобы `duplicate_candidates` сели на неё, а `_apply_mark` поставил `match_status=ambiguous` без GUID. Sweep переводит open-спрос в `field=duplicate`. Имя вне прайса и вне спроса остаётся `unmatched_1c` и открывает `price_queue`. Файл с «Отбор:» по-прежнему не наполняет `disappeared`.
  - **Acceptance:**
    - [x] Два GUID → в справочнике нет нового `guid_1c`, статус `ambiguous`, спрос после sweep `duplicate`
    - [x] Чужое имя → нет новой строки справочника, есть открытая `price_queue`
    - [x] Частичный файл не добавляет `disappeared` по маркам, которых в файле не было
  - **Verify:** `pytest tests/test_nomenclature_import.py tests/test_guid_demand.py -q`
  - **Dependencies:** GKP-001
  - **Files:** `core/nomenclature_sync.py`, `tests/test_nomenclature_import.py`
  - **Scope:** M

- [x] **GKP-004: Сброс плашки шага 3**
  - **Description:** В `onSuccess` у `useImport1cMutation` инвалидировать `guidCheckKeys.all` рядом с `nomenclatureKeys.all`.
  - **Acceptance:**
    - [x] Успешный импорт инвалидирует и `nomenclature`, и `commercial/guid-check`
  - **Verify:** `cd frontend && npm run test -- src/features/nomenclature/hooks/useNomenclatureQueries.ts`
  - **Dependencies:** —
  - **Files:** `frontend/src/features/nomenclature/hooks/useNomenclatureQueries.ts`, тест хука, если его ещё нет — рядом с `nomenclatureApi.test.ts` или новый узкий тест
  - **Scope:** S

### Checkpoint: Done

- [x] **GKP-005: Сводка проверок**
  - **Description:** Прогнать команды ниже. В спеке статус `PLAN ✅ · IMPLEMENT ✅` только после зелёного прогона. Код тестовых мостовых свай и `ЛС11` не чистить.
  - **Verify:**
    - [x] `pytest tests/test_nomenclature_import.py tests/test_nomenclature_import_api.py tests/test_guid_demand.py tests/test_commercial_draft_oneoff_price.py -q`
    - [x] `cd frontend && npm run test -- src/features/nomenclature/hooks/useNomenclatureQueries.ts src/features/nomenclature/components/Import1cDialog.test.tsx`
    - [x] `cd frontend && npm run typecheck`
  - **Dependencies:** GKP-002, GKP-003, GKP-004
  - **Files:** `ai_docs/specs/guid-iz-vygruzki-na-marku-kp.md` (строка статуса)
  - **Scope:** S

## Out of Scope

- Новая цена из файла и ручной ввод GUID
- Полная заливка всех марок прайса, у которых нет строки справочника
- Плиты в `nomenclature_guid`
- Удаление тестовых строк мостовых свай и `ЛС11` из живой `guid_demand`
