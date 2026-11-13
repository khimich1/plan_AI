# Implementation Plan: Почему нет GUID — сверка очереди, причина и срез, замена марки

**Created:** 2026-11-12
**Status:** 🚧 IMPLEMENT: срезы 0–3 (NG-001…NG-304) выполнены локально 2026-11-12, не закоммичены; NG-401/402/403/404 — «на деплой».
**Spec:** [`ai_docs/specs/pochemu-net-guid.md`](../../specs/pochemu-net-guid.md) (SPECIFY ✅ 2026-11-12)
**Idea:** [`ai_docs/ideas/pochemu-net-guid.md`](../../ideas/pochemu-net-guid.md)
**Родитель:** [`plate-load-12-guid-alias.md`](../../specs/plate-load-12-guid-alias.md) (пересмотрен D-load12, D-name125)
**Формат:** план и задачи в одном файле (как соседние планы в `ai_docs/develop/plans/`, путь из `.cursor/config.json` → `documentation.paths.plans`). `tasks.json` в `.cursor/workspace/` создаёт оркестратор, не этот план.

## Overview

Четыре среза по D-order: **0** сверка очереди `guid_demand` с КП → **1** причина, «та же карточка», срез, поиск без записи в КП + правило «12 → 12,5п» (D-load12) и имя/нагрузка (D-name125) → **2** замена экономиста на «Завести в 1С» → **3** замена менеджера на шаге 3. Перед релизом: сухой прогон сверки на копии прода (показать пользователю), диф ворот по всем КП копии, один тестовый счёт в 1С, починка КП 206 и 210.

---

## 1. Что проверено по коду (вопросы, оставленные спекой «на план»)

Всё ниже — чтение репозитория; прод читался один раз, только `mode=ro` (пересчёт имён «12п»).

### C1. Нагрузка из имени карточки: 12,5 → 1250

- Готовой функции нет. `parse_load_code_from_name` (`core/config_and_data.py:241-278`, `int(val + 0.5)` в стр. 276) даёт 13. Нельзя.
- Есть кусочки: регэксп `_LOAD_SUFFIX_RE` с группой `load` (`core/guid_gate.py:52-55`) уже ловит `12`, `12.5`, `12,50`; `normalize_load_code` (`core/domain/plate_order.py:22-44`) знает 12.5 и 1250 ↔ 12.5. В обратную сторону есть `format_reinforcement_from_load_code` (`core/config_and_data.py:67-87`, 12.5 → «12,5п»).
- **Нужна одна маленькая новая функция** в `core/kp_db_nomenclature.py` (рядом с переехавшим `_LOAD_SUFFIX_RE`): `load_class_from_plate_name(name) -> int | None`: `float(load.replace(",", "."))`, `int(round(v * 100))` → 1250 / 1200 / 800; нет хвоста `…п` → `None`. `normalize_load_code` из `core.domain` не импортировать в `kp_db_nomenclature` (риск цикла; для `*100` он не нужен).
- Где подставлять (сегодня `load_class` считается из набранного `load_code`):
  - `app/services/commercial_service.py:211` `(normalize_load_code(load_code) or 8) * 100` при том, что имя перезаписывается канонической карточкой в стр. 191-193 → при «12п → 12,5п» останется 1200. Правка: если `canonical_name` применён и из него читается нагрузка — брать её.
  - `core/kp_db_nomenclature.py:157-163` (`enrich_order_data_with_nomenclature` переписывает `name`, `load_class` не трогает).
  - Новая замена (срез 2/3) — тем же хелпером.
- **Скрытая дыра в коде:** `app/services/plate_parser_service.py:211-216` строит имя для кэша через `load_code=int(float(load_code))` — набранное `12,5п` ищется как `…12п`. Это сломает «правило одностороннее» (набрано 12,5п — на 12п не откатываемся) и в вебе, и после D-load12 (поиск пойдёт по ключу «12п»). `fill_plate_nomenclature_cache` (`core/kp_db_nomenclature.py:123-127`) так не делает. Правка входит в NG-102.
- `parse_load_code_from_name` в `_build_order_data` **остаётся** (`commercial_service.py:164`, подбор строки прайса через `load_code_for_price_match`: 12 и 12,5 → 12). Поэтому тест спеки «нет `parse_load_code_from_name`» писать на хелпер, замену и `enrich`, а не на весь `_build_order_data` (иначе ложный красный).
- Производство: `core/production/planning.py:587` `normalize_load_code(load_class // 100)` → `1250 // 100 = 12` — как в спеке. Цена: `core/commercial_pricing.py:254` `int(load_class) // 100` = 12 — 1250 считается как 12, как и сегодня для набранного `12,5п`. Мелочь: `commercial_pricing.py:383-384` подпись `…-{load_class // 100}п` покажет «12п» (текст ошибки, не запись).

### C2. Как показать причину «несколько кандидатов по LIKE»

- Сегодня `LIKE` вообще не считается: `count_nomenclature_by_plate_name` (`kp_db_nomenclature.py:47-65`) смотрит только точное имя и варианты; `lookup_nomenclature_by_plate_name` (стр. 68-99) берёт первую строку по GUID (`_fetch_plate_row`, стр. 27-34) и только **её** проверяет на длину (стр. 91-96) — если первая не подошла по длине, остальные не смотрятся.
- Решение: **новая причина** `REASON_VARIANTS = "стандартной карточки нет, есть особые"` в `core/guid_gate.py` (рядом с `REASON_*`, стр. 45-48), а не подсказка к `REASON_MISSING`. Почему: (1) её видно в `check_invoice_guids`, в ворота счёта (`archive_service.py:612`, `:672`, текст блокировки) и в `/guid-check` (шаг 3, `CalculationResultStep.tsx:464-477` печатает `item.reason`) без фронтового кода; (2) не `REASON_DUP`: там одно имя и два GUID, здесь разные имена; (3) `field_for_reason` (`core/guid_demand.py:119-126`) всё незнакомое отправляет в `FIELD_GUID` — ключ очереди `(kind, mark, field)`, `FIELDS` и тип `field` на фронте не меняются.
- В очередь уходит через `hint` (колонка `guid_demand.hint` уже есть и динамическая): `«…особые: Пв-7.1, Пв-7.2 — заведите карточку в 1С и загрузите отчёт»` (до 3 имён и «и ещё N»). `Create1cTaskOut` (`app/schemas/nomenclature.py:39-44`) отдаёт `hint`, `field`, `kp_ids`; `TasksSection.tsx:61` печатает `item.hint` — **фронт очереди для текста не меняется**.
- Отдельно: `sweep_guid_demand` (`core/guid_demand.py:260-285`) при той же `field` ничего не обновляет (`continue`, стр. 272-273), reason/hint остаются старыми. Для смены «нет GUID» → «есть особые» нужно обновлять reason/hint на месте (NG-103).
- `MissingItem` (frozen dataclass, `guid_gate.py:70-74`) получает необязательное `candidates: tuple[str, ...] = ()` — совместимо.
- Список особых карточек для менеджера и экономиста — не в причине, а в срезе (NG-104/NG-107, `explain`).

### C3. Тесты, которые упадут или потребуют правки

Если сохранить сигнатуры (`lookup_nomenclature_by_plate_name` → 3-кортеж, `count_nomenclature_by_plate_name` → int; новое поведение — новая функция `find_plate_card`, старые — тонкие обёртки), **ни один существующий тест не должен упасть**. Это целевой вариант плана (минимальный diff). Если реализатор всё же поменяет возврат на «два имени» (как в спеке), красными станут:

| Файл | Тесты | Почему |
|------|-------|--------|
| `tests/test_kp_db_nomenclature.py` | `test_lookup_exact_match` (стр. 41-53), `test_lookup_duplicate_names_returns_stable_guid` (56-83), `test_lookup_prays_width_variant` (86-100), `test_lookup_trailing_decimal_zero_matches_catalog_mark` (103-114), `test_lookup_prays_length_variant_bridges_40_and_40_0` (117-131), `test_lookup_like_rejects_length_mismatch` (134-148), `test_lookup_like_accepts_matching_fractional_length` (151-165) | распаковка `name, nom_id, match_type = …` в строках 47, 77, 94, 110, 126, 143, 158 |
| `tests/test_kp_db_nomenclature.py` | `test_enrich_order_data_*` (168-200) | `enrich_order_data_with_nomenclature` вызывает lookup (стр. 157) |

- Оба `test_lookup_like_*` сегодня про **одного** кандидата. Спека пишет «переписаны под это» — по коду они остаются зелёными (после правки длины по всем кандидатам, не только по первому). Нужны **новые** тесты: два кандидата → ничего не выбрано; один кандидат + другая длина → не выбран.
- `tests/test_guid_gate.py:226-306` (`test_plate_load_12_*`, 8 тестов) работают только через публичный `check_invoice_guids`; порядок «алиас до LIKE» их не ломает: прошёл по каждому вручную (единственная `12,5п`; `12.5п`, `12,50п`; ширина `7 → 7,0`; дубль без выбора → `REASON_DUP`; дубль с выбором по имени карточки; своя `12п` выигрывает; `8п` и не-плиты не алиасятся). Остаются **регрессионным щитом**: их нельзя править «под зелёное».
- `tests/test_guid_demand.py` (стр. 178, 198, 237 — `sweep_guid_demand`): остаются зелёными, если обновление reason/hint в `sweep` срабатывает только при смене текста (NG-103). Новые тесты — туда же.
- `tests/test_notifications.py` (`notify_kp_guid_missing`, стр. 93-231), `tests/test_guid_gate_api.py`, `tests/test_duplicate_resolve.py`, `tests/test_plate_guid_choice.py`: не должны меняться. Сверку **не** встраивать внутрь `notify_kp_guid_missing` (иначе эти тесты полезут во вторую базу).
- `tests/test_commercial_service_order_data.py`: тестов на канонику и `load_class` нет — добавить (NG-102). `tests/test_app_plate_parser_service.py` (стр. 22, `65,6-12-12,5  2`): перечитать перед NG-102 (кэш номенклатуры по `int(float())`).
- Новые файлы: `tests/test_guid_demand_reconcile.py` (сверка), `tests/test_guid_explain.py`, `tests/test_guid_replace.py` (+ vitest на фронте).

### C4. Куда вешать сверку и как не сломать сохранение

Источники событий:

| Событие | Точка | Что делаем |
|---------|-------|------------|
| Сохранение КП из конструктора (новое и правка архивного) | `app/services/commercial_draft_lifecycle.py:1051-1066` (блок `notify_kp_guid_missing` внутри `if persist_status == "в архиве"`, все ошибки уже в `try/except` + `logger.exception`) | **после** блока, **вне** условия по статусу (правка «на согласовании» тоже меняет состав): `sync_guid_queue_with_kp(...)` |
| Создание КП через API офферов | `app/services/offers_service.py:117` (`_notify_guid_missing`, стр. 233-248, `try/except`) | то же, рядом |
| Удаление КП | `app/services/archive_service.py:368-376` `delete_offer` → `repository.delete` → `kp_archive_repository.py:117` → `core/kp/offers_write.py:614-651` `delete_kp_by_id` (ловит любые исключения, возвращает `False`) | после успешного `repository.delete` в `ArchiveService.delete_offer`; **не** внутрь `delete_kp_by_id` (core не должен знать про pb.db; `bot_archived/*` — мёртвый код) |
| Открытие «Завести в 1С» | `app/api/v1/endpoints/nomenclature.py:82` `GET /nomenclature/tasks` → `NomenclatureQueueService.list_tasks` (`nomenclature_queue_service.py:40-70`) | перед `collect_guid_queue`; GET получает запись — писать **только если сверка что-то изменила** |
| Разовый прогон | `scripts/reconcile_guid_demand.py` | dry-run по умолчанию, `--apply` |

Как не ломать сохранение КП, если `pb.db` недоступна (запись в две базы не атомарна):

1. КП уже закоммичен в `plita.db` раньше (`update_offer_from_order_data` коммитит сам; `KpPersistenceService.update_kp_from_order_data`, `core/kp_persistence_service.py:655`). Сверка идёт строго **после** этого.
2. Весь вызов — в `app/services/guid_queue_sync.py::sync_guid_queue_with_kp(pb_db_path, plita_db_path)`: проверка `Path(pb).is_file()`, соединение с `timeout=3` (залипший писатель не вешает запрос), `plita` только `mode=ro`, `try/except Exception → logger.exception → return None`. Наружу не бросает никогда.
3. Обрыв оставляет лишнюю строку очереди, задачу не теряет; следующая сверка (любое сохранение, `list_tasks`, скрипт) исправляет. Флаг `queue_synced` не нужен (допущение 17).
4. Тест: подмена сверки на `raise` и `pb.db` отсутствует → сохранение КП успешно, ответ тот же.
5. Для замены в нескольких КП (срез 2) — отдельный порядок «все КП одним соединением и одним `commit`», затем сверка; см. NG-205/NG-206. Сейчас `update_kp_from_order_data` (`kp_persistence_service.py:357-661`) открывает своё соединение и коммитит сам; `_external_conn` в нём нет. Образцы: `core/kp/offers_write.py:439-460`, `:937-956`.

### C5. «Факты с прода» и разовый скрипт

- В репозитории **скрипта, который эти факты выдаёт, нет**: числа 59/52/7 получены чтением прода вручную. `scripts/build_guid_queue.py` и `_tmp_guid_match.py` — другое (Excel-книги сопоставления, латинская `C`).
- Нужен `scripts/reconcile_guid_demand.py` (тонкая обёртка над `core/guid_demand_reconcile.py`, как остальные `scripts/*.py`: `PROJECT_ROOT` в `sys.path`, `argparse`):
  - `--pb-db PATH`, `--plita-db PATH` (по умолчанию из `get_settings()`), `--apply` (без него dry-run), `--report PATH`;
  - dry-run = полный прогон в транзакции и `rollback()` (один путь кода для dry-run и apply, нет двойной логики). **Только на копии:** `check_invoice_guids` вызывает `ensure_schema` (CREATE IF NOT EXISTS) на соединении pb;
  - отчёт на русском: по каждой строке «было → стало» (закрыта / сужена КП / перепривязана / открыта новая), сводка и «не изменилось: N». Формат повторяет таблицу «Ожидаемый результат» из спеки, чтобы пользователь сверил глазами;
  - второй запуск подряд печатает «изменений нет» (критерий идемпотентности).
- Копия прода — только с согласия пользователя (NG-401).

### C6. Прочее, найденное при чтении

- **Выбор дубля по варианту.** `_lookup_plate` (`guid_gate.py:117-141`) при дубле возвращает `matched_name` — **набранное** имя, а `get_choice` (стр. 176) ищет выбор по нему; выбор из «Дубли GUID» записывается по имени карточки (`resolve_duplicate`, `nomenclature_queue_service.py:149-162`). Для `22,80` при дубле на `22,8` выбор не находится. Единый `find_plate_card`, который возвращает `canonical_name` карточки, лечит это попутно. Это и есть реальный смысл «двух имён» из спеки: имя в КП = `canonical_name`, а набранное у вызывающего уже есть — отдельного второго поля не нужно.
- **«Отправить в 1С недоступна»** реализовано без фронта: `_invoice_export_block` (`archive_service.py:645-680`) возвращает текст `InvoiceGuidBlockError`. Новая причина автоматически попадёт в этот текст.
- **Риск включения D-variants для старых КП:** строки, которые сегодня находят карточку молча через `LIKE` (первая по GUID), после правки получат `REASON_VARIANTS` и заблокируют «Отправить в 1С» у этих КП. Нужен диф ворот по всем КП копии прода до выкладки (NG-402).
- **Плиты в плане.** `KpPersistenceService._update_line` для плит (`kp_persistence_service.py:889-947`) пишет имя, размеры, `load_class` без проверки статуса плиты; защита `_PROTECTED_PLATE_STATUSES` (`:52`) применяется только к удалению (`:566-573`). Замена размеров у плиты «в плане»/«в производстве» разведёт КП и план — см. вопрос Q2.
- **История «кто».** `update_kp_from_order_data` не получает пользователя (в строке КП его нет — факт спеки). Для D-history нужен необязательный `changed_by` по цепочке `lifecycle → kp_repository.update_offer_from_order_data (kp_repository.py:85-121) → offers_write.update_kp_from_order_data (offers_write.py:82-115) → KpPersistenceService`.
- **Пересчёт «девять» из спеки.** Чтение прода (`/data/plita.db`, ro, 2026-11-12): в `kp_plates` 23 строки с `…12п` (не `12,5п`) в КП 132, 167, 172, 206, 210, **11 различных имён** (8 через алиас + 2 дубля + 1 `51-12-12п Пв-7.1`), а не «девять»; 2 из 23 строк — с `Пв-7.1` (КП 206, 210), остальные 21 — как в спеке. Всего плит 1414, `nomenclature_id` пуст у 677; строк `12,5п` 93 в 12 КП. Расхождение с цифрой в спеке — в «Замечаниях к спеке».

---

## 2. Architecture Decisions

- **A1. Один разбор плиты — `find_plate_card(cur, name)`** в `core/kp_db_nomenclature.py`. Возвращает `PlateCardMatch(status ∈ {found, dup, variants, none}, canonical_name, guid, guid_count, candidates, match_type)`. Порядок строго D-load12: точное → варианты (`plate_name_to_prays_variants`) → «12 → 12,5п» точным и теми же вариантами → `LIKE` (длина проверяется по **всем** кандидатам; 1 → found, >1 → variants, 0 → none). Точное имя с двумя GUID остаётся «первый по GUID» в `lookup_…` (тест `…stable_guid`), а `dup` его видит через `guid_count`.
- **A2. Старые функции — обёртки.** `lookup_nomenclature_by_plate_name` (3-кортеж, `variants` → `(None, None, None)`) и `count_nomenclature_by_plate_name` (int) вызывают `find_plate_card`. Вызывающие (`kp_db_nomenclature.py:128,157`, `plate_parser_service.py:217`, `kp_db.py`) и 7 тестов не меняются.
- **A3. `_plate_load_12_alias` и `_LOAD_SUFFIX_RE` переезжают** в `kp_db_nomenclature` (публичное имя `plate_load_12_alias`); `guid_gate` их импортирует, как и остальное (спека, «Project Structure»).
- **A4. Ворота плит упрощаются**: `_resolve_plate`/`_lookup_plate`/`_finish_plate_lookup` (`guid_gate.py:117-199`) становятся: сохранённый выбор по набранному имени → `find_plate_card` → `dup` + `get_choice(canonical_name)` → `variants` → `REASON_VARIANTS` → `none` → `REASON_MISSING`.
- **A5. Сверка — отдельный чистый модуль** `core/guid_demand_reconcile.py` (две базы передаются соединениями, ничего не знает про FastAPI). Тонкая обёртка с `try/except` — `app/services/guid_queue_sync.py`. `sweep_guid_demand` не меняется, кроме обновления reason/hint (NG-103).
- **A6. Сверка вызывается после уведомления, не внутри него**: `notify_kp_guid_missing` (колокольчик, допущение 12) не трогаем.
- **A7. Один разбор причин на все виды** — `core/guid_explain.py::explain_missing_guid(line, conn) -> GuidExplain`, роль (менеджер/экономист) определяет эндпоинт, не клиент. Менеджерский ответ не содержит `selectable=true` при `guid_count != 1`; **повторная проверка на сервере при записи**, клиенту не верим.
- **A8. Замена: чистый построитель строки** (`core/guid_replace.py`: `build_replacement(line, target, discount_percent, kind)`) + расчёт денег для окна (`preview_replacement`). Один и тот же для менеджера (черновик) и экономиста (КП).
- **A9. Запись замены экономиста:** все КП одним соединением и одним `commit` (через `_external_conn`), затем `recalc_on_composition_change` каждого КП (после коммита, `recalc_promise_after_edit` ловит свои ошибки), затем сверка (`try/except`). Ответ успешен, даже если шаг очереди упал.
- **A10. История** — таблица рядом с КП в `plita.db`, пишется внутри `update_kp_from_order_data` в той же транзакции при любом изменении марки строки (по `line_id`), поле `source` (`replace_economist` / `replace_manager` / `edit`).
- **A11. Без новых пакетов; `./run+logs.sh` не останавливать; коммит только по просьбе.**

---

## 3. Зависимости и порядок

```
Срез 0 (сверка)
 NG-001 ядро ──┬─ NG-002 вызовы после сохранения/удаления ── NG-003 «Завести в 1С»
               └─ NG-004 скрипт
                          │ чекпоинт 0
Срез 1 (причина и срез)   ▼
 NG-101 find_plate_card ──┬─ NG-102 имя/нагрузка (12→12,5п)
                          ├─ NG-103 REASON_VARIANTS + sweep
                          └─ NG-104 explain: плиты ── NG-105 explain: сваи/мостовые/составные
                                                   └─ NG-106 explain: ФБС/ступени/марши
                              NG-107 API (explain, search) ── NG-108 панель UI ── NG-109 подключение
                          │ чекпоинт 1
Срез 2 (экономист)        ▼
 NG-205 внешнее соединение ── NG-201 история ── NG-202 история на карточке КП
 NG-203 построитель ── NG-204 деньги (окно)         (203/204 параллельно 205/201)
 NG-206 сервис замены (нужны 201,203,204,205) ── NG-207 API ── NG-208 UI
                          │ чекпоинт 2
Срез 3 (менеджер)         ▼
 NG-301 сервис черновика ── NG-302 API ── NG-303 UI ── NG-304 регрессия
                          │ чекпоинт 3
Проверка до выкладки: NG-401 … NG-404
```

**Строго подряд (по D-order и общим файлам):**
- Срезы по сборке 0 → 1 → 2 → 3, релиз срезом 0 раньше остальных (вопрос Q1).
- NG-101 раньше NG-102/103/104 (общий `kp_db_nomenclature.py`, `guid_gate.py`); NG-102 после NG-101 (оба в `kp_db_nomenclature.py`).
- NG-205 раньше NG-201 (оба правят `core/kp_persistence_service.py`); NG-201 раньше NG-206; NG-206 раньше NG-207/208.
- Миграция схемы истории (NG-201) — один раз, до замены.
- NG-107 (контракт ответов) раньше NG-108; NG-203/204 раньше срезов 2 и 3 (их использует и менеджер).
- Фаза 4 только после чекпоинта 3 (кроме NG-401, её можно сделать сразу после чекпоинта 0).

**Можно параллельно (разные файлы):**
- NG-004 (скрипт) ∥ NG-002/NG-003 после NG-001.
- NG-102 ∥ NG-103 после NG-101.
- NG-105 ∥ NG-106 (разные виды в одном `core/guid_explain.py` — разнести по функциям видов, конфликт слияния маловероятен; если сессии две, лучше разные модули `core/guid_explain_*.py`).
- NG-108 (UI на моках по схемам NG-107) ∥ NG-105/106.
- NG-203 ∥ NG-205/NG-201; NG-202 ∥ NG-206.
- NG-208 (UI) на моках ∥ NG-206.
- Техническая подготовка по D-order: NG-101 формально после среза 0, файлы у них разные — писать RED-тесты и ветку можно параллельно, вливать после чекпоинта 0.

**Правила выполнения каждой задачи:** TDD (сначала красный тест), затем минимальная реализация; тесты на фикстурах, без живых `pb.db`/`plita.db`; ruff/pytest по затронутым файлам после правки; при красном в чужом тесте — стоп, не подгонять.

---

## 4. Task List

### Phase 0: Сверка очереди с КП

- [x] NG-001 Ядро сверки `(type: feat-be, M)`
- [x] NG-002 Вызов после сохранения и удаления КП `(type: feat-be, M, dependsOn: NG-001)`
- [x] NG-003 Сверка при открытии «Завести в 1С» `(type: feat-be, S, dependsOn: NG-002)`
- [x] NG-004 Скрипт разового прогона (dry-run / apply) `(type: chore, S, dependsOn: NG-001)`

### Phase 1: Причина, срез, правило 12 → 12,5п

- [x] NG-101 `find_plate_card`: порядок D-load12, `LIKE` не выбирает сам `(type: feat-be, M)`
- [x] NG-102 Имя КП и `load_class` из карточки `(type: feat-be, M, dependsOn: NG-101)`
- [x] NG-103 `REASON_VARIANTS` и обновление reason/hint в `sweep` `(type: feat-be, S, dependsOn: NG-101)`
- [x] NG-104 `explain_missing_guid`: плиты, срез, подпись различия `(type: feat-be, M, dependsOn: NG-101, NG-103)`
- [x] NG-105 `explain`: сваи, мостовые, составные сваи `(type: feat-be, M, dependsOn: NG-104)`
- [x] NG-106 `explain`: ФБС, ступени, марши `(type: feat-be, M, dependsOn: NG-104)`
- [x] NG-107 API: объяснение и поиск (менеджер и экономист) `(type: api, M, dependsOn: NG-104)`
- [x] NG-108 UI: панель причины, «той же карточки», среза и поиска `(type: feat-fe, M, dependsOn: NG-107)`
- [x] NG-109 Подключение панели к шагу 3 и к «Завести в 1С» `(type: feat-fe, S, dependsOn: NG-108)`

### Phase 2: Замена экономиста

- [x] NG-205 Путь `update_kp_from_order_data` через внешнее соединение `(type: refactor, M)`
- [x] NG-201 История изменения марки `(type: feat-be, M, dependsOn: NG-205)`
- [x] NG-202 Запись о замене на карточке КП `(type: feat-fe, S, dependsOn: NG-201)`
- [x] NG-203 Построитель строки замены (все виды) `(type: feat-be, M)`
- [x] NG-204 Деньги для окна «было / станет» `(type: feat-be, M, dependsOn: NG-203)`
- [x] NG-206 Сервис замены в КП строки очереди `(type: feat-be, M, dependsOn: NG-201, NG-203, NG-204, NG-205, NG-002)`
- [x] NG-207 API замены экономиста `(type: api, S, dependsOn: NG-206)`
- [x] NG-208 UI замены на «Завести в 1С» `(type: feat-fe, M, dependsOn: NG-207, NG-109)`

### Phase 3: Замена менеджера на шаге 3

- [x] NG-301 Сервис замены в черновике `(type: feat-be, M, dependsOn: NG-203, NG-204)`
- [x] NG-302 API замены менеджера `(type: api, S, dependsOn: NG-301)`
- [x] NG-303 UI замены на шаге 3 `(type: feat-fe, M, dependsOn: NG-302, NG-109)`
- [x] NG-304 Регрессия: очередь не трогается, «Отправить в 1С» заблокирована `(type: chore, S, dependsOn: NG-303)`

### Phase 4: До выкладки

- [ ] NG-401 Сухой прогон сверки на копии прода и показ пользователю `(type: chore, S)`
- [ ] NG-402 Диф ворот по всем КП копии прода, до и после `(type: chore, S, dependsOn: NG-103)`
- [ ] NG-403 Один тестовый счёт в 1С `(type: chore, S, dependsOn: NG-102)`
- [ ] NG-404 Полная регрессия и починка КП 206 и 210 `(type: chore, S, dependsOn: NG-304)`

---

### Phase 0 — подробно

#### NG-001: Ядро сверки `core/guid_demand_reconcile.py`

**Цель.** `reconcile_guid_demand(pb_conn, plita_conn) -> ReconcileReport`: для каждой открытой строки `guid_demand` список КП заменяется актуальным (D-reconcile).

**Файлы:** `core/guid_demand_reconcile.py` (новый); `core/guid_demand.py` (`_set_kp_ids`, `_set_resolved` публичнее; без смены существующих функций); `tests/test_guid_demand_reconcile.py` (новый). Всего 3.

**Как.**
1. Индекс «что сейчас в КП» за один проход: `(kind, mark) → [kp_id]` из `kp_plates.plate_name` (kind `plate`) и `.mark` в `kp_piles` (`pile`), `kp_steps` (`stair_step`), `kp_marches` (`stair_flight`), `kp_bridge_piles` (`bridge_pile`), `kp_composite_piles`, `kp_fbs` (`core/kp_db_schema.py:53, 391-500`). Совпадение точное; любой статус КП; `КП` без `KP_offers` не бывает (CASCADE).
2. Для строки: актуальный список = индекс; равен — не трогаем; непустой — `kp_ids_json` заменяется (сужение, перепривязка); пустой — `state=resolved`, `resolved_at`.
3. Выпавшие КП (были в строке, в актуальном списке нет): для каждого резолвер `check_invoice_guids` по его текущим строкам из индекса, `record_guid_demand` открывает новую строку. Номер у новой марки есть → ничего не открывается (`ЛС11 → ЛС11-1`).
4. Отчёт: закрыты / сужены / перепривязаны / открыты / без изменений. Запись только если что-то изменилось.
5. Транзакции: функция не коммитит; вызывающий решает (dry-run = rollback).

**Приёмка:**
- [ ] Тесты из спеки («Сверка очереди с КП»): сужение; закрытие при удалении КП; закрытие при правке марки; перепривязка `С120.35-12` 138 → 9, строка открыта; новая строка, если у новой марки нет номера; не открывается, если номер есть; повторный прогон — пусто.
- [ ] Строка с `field = guid_1c_u` и «у»-маркой сверяется по точной марке с «у»
- [ ] Закрытая строка, которая снова нужна, открывается через `record_guid_demand` (существующий путь)

**Тесты (RED первыми):** фикстуры двух in-memory баз: `plita` с 7 таблицами (минимальные столбцы), `pb` с `guid_demand`, `prays_plity`, `nomenclature_guid`; без живых баз. Строки данных повторяют таблицу «Разъехавшиеся строки» спеки (КП 124, 126, 206, 59/55/57, 138/9).
**Размер:** M. **Риски:** (1) `mark` в `guid_demand` и в КП должны совпадать буква в букву — на проде 52 из 59 совпадают (факт); тест на «у» и кириллицу; (2) резолвер вызывает `ensure_schema` (DDL) — не вызывать на `plita`; (3) закрытая и снова нужная строка получает новый `opened_at`? Сохранить прежний (существующее поведение `_upsert_open`, стр. 177-178).

#### NG-002: Вызовы после сохранения и удаления КП

**Цель.** Очередь следует за КП (допущение 15) без риска для сохранения (C4).

**Файлы:** `app/services/guid_queue_sync.py` (новый); `app/services/commercial_draft_lifecycle.py` (после блока `:1051-1066`); `app/services/offers_service.py` (рядом с `:117`/`:233`); `app/services/archive_service.py` (`delete_offer`, `:368-376`); `tests/test_guid_demand_reconcile.py` (+ тесты вызовов). 5 файлов.

**Приёмка:**
- [ ] `sync_guid_queue_with_kp` не бросает: нет `pb.db`, залипшая база (`sqlite3.OperationalError`), ошибка внутри сверки — ответ сохранения тот же, в логе `logger.exception`
- [ ] После сохранения КП с исправленной маркой старая строка очереди закрыта
- [ ] После `delete_offer` строка по удалённому КП закрыта/перепривязана
- [ ] Коммит в `pb.db` только если сверка что-то изменила
- [ ] Колокольчик (`notify_kp_guid_missing`) работает как раньше; его тесты `test_notifications.py` не менялись

**Тесты:** unit на обёртку (monkeypatch на `raise`, отсутствующий путь); по одному сквозному на lifecycle и на delete с подменой путей баз (существующие фикстуры `tests/test_commercial_draft_append.py`, `tests/test_archive_*`).
**Размер:** M. **Риски:** сверка внутри запроса сохранения добавляет задержку (59 строк, 7 SELECT — миллисекунды; `timeout=3`); гонка двух сохранений (SQLite сериализует, повтор — идемпотентен).

#### NG-003: Сверка при открытии «Завести в 1С»

**Файлы:** `app/services/nomenclature_queue_service.py` (`list_tasks`, стр. 40-70; принять путь `plita`), `app/api/v1/endpoints/nomenclature.py` (`/tasks`, стр. 82; DI путь), `tests/test_guid_demand_reconcile.py` (+ тест сервиса). 3 файла.

**Приёмка:**
- [ ] `GET /nomenclature/tasks` перед выдачей приводит очередь в порядок; при падении сверки отвечает 200 с текущей очередью
- [ ] Строк, про которые КП давно нет, в ответе нет
- [ ] Повторные GET ничего не пишут (проверка: нет изменения файла/`resolved_at`)

**Размер:** S. **Риски:** GET с записью — кэш/refetch React Query (`useGuidTasksQuery`) может дёргать часто; запись только при изменении, сверка дешёвая. Мок-пути `plita` в тестах `tests/test_guid_gate_api.py` не должен ломать существующие тесты (параметр необязателен).

#### NG-004: Скрипт разового прогона

**Файлы:** `scripts/reconcile_guid_demand.py` (новый), `core/guid_demand_reconcile.py` (`format_report`), `tests/test_guid_demand_reconcile.py` (тест формата). 3 файла.

**Приёмка:**
- [ ] Без `--apply` — dry-run, база не меняется (проверка хешем файла копии в тесте)
- [ ] `--apply` меняет; второй запуск — «изменений нет», код возврата 0
- [ ] Отчёт русский, по строке: марка, КП было → стало, действие; сверяется с таблицей «Ожидаемый результат» спеки
- [ ] В начале отчёта и в `--help` печатаются пути обеих баз и пометка «запускать только на копии прода» (отдельного защитного флага не вводим)

**Размер:** S. **Риски:** случайный `--apply` на проде → спека требует показать первый прогон пользователю; в `--help` и в начале отчёта печатать пути баз.

**Checkpoint 0**
- [ ] `venv/bin/pytest tests/test_guid_demand_reconcile.py tests/test_guid_demand.py tests/test_notifications.py tests/test_guid_gate_api.py -q` зелёный
- [ ] Показ пользователю: NG-401 (сухой прогон на копии) — до включения вызовов на проде

---

### Phase 1 — подробно

#### NG-101: `find_plate_card`: порядок D-load12

**Цель.** Один разбор плиты в трёх местах (ворота, разбор строки, подсчёт дублей); `LIKE` с несколькими кандидатами сам не выбирает.

**Файлы:** `core/kp_db_nomenclature.py`; `core/guid_gate.py` (`:52-55`, `:117-199`, импорт); `tests/test_kp_db_nomenclature.py`; `tests/test_guid_gate.py`. 4 файла.

**Приёмка:**
- [ ] `51-12-12п` при `12,5п`, `12п Пв-7.1`, `12п Пв-7.2` → `canonical_name = Плиты ПБ 51-12-12,5п`, GUID `0de0f4a7-dc8e-11ea-8f69-04d9f5387845`; одинаково в `lookup`, `count` (=1), `check_invoice_guids`
- [ ] Своя чистая `12п` выигрывает; два GUID у `12,5п` → `dup`, `REASON_DUP`, выбор по `canonical_name` закрывает строку (в т. ч. при вариантах `22,80 → 22,8`)
- [ ] Набрано `12,5п`, карточки `12,5п` нет, есть `12п` с хвостом → **не** откатывается на `12п`
- [ ] `LIKE` с ≥2 кандидатами → `status = variants`, `lookup` возвращает `(None, None, None)`; 1 кандидат → как раньше; длина проверяется у **всех** кандидатов
- [ ] `8п`, `10п`, не-плиты — без алиаса; 7 существующих тестов `test_kp_db_nomenclature` и `test_plate_load_12_*` зелёные без правок
- [ ] Строка `REASON_VARIANTS` в воротах пока может быть заглушкой `REASON_MISSING` до NG-103 (но `variants` в `PlateCardMatch` уже есть)

**Тесты (RED):** добавить в `test_kp_db_nomenclature.py`: два кандидата `LIKE`; один подходящий по длине из двух; алиас в `lookup/count`; одностороннее правило. В `test_guid_gate.py`: `51-12-12п` с тремя карточками; `22,80` + дубль + выбор по карточке.
**Размер:** M. **Риски:** (1) самый рискованный срез: меняет, что считается найденным у старых КП (см. NG-402); (2) `plate_name_to_prays_variants` для алиаса применяется к имени `…12,5п` — тест на ширину `7-12п` (уже в `test_guid_gate.py:248`); (3) циклический импорт — функции переезжают вниз, `guid_gate` уже импортирует `kp_db_nomenclature`.
**Чекпоинт внутри:** `pytest tests/test_guid_gate.py tests/test_kp_db_nomenclature.py tests/test_guid_demand.py -q`.

#### NG-102: Имя КП и `load_class` из карточки (D-name125)

**Файлы:** `core/kp_db_nomenclature.py` (`load_class_from_plate_name`, `enrich_order_data_with_nomenclature`); `app/services/commercial_service.py` (`:191-193`, `:211`); `app/services/plate_parser_service.py` (`:211-216`); `tests/test_commercial_service_order_data.py`; `tests/test_kp_db_nomenclature.py`. 5 файлов.

**Приёмка:**
- [ ] Набрано `Плиты ПБ 51-12-12п`, карточка `12,5п` → строка КП: `name = …12,5п`, `load_class = 1250`, `nomenclature_id = 0de0f4a7-…`; `Пв-7.1` нет
- [ ] `load_class_from_plate_name`: `…12,5п` → 1250; `…12п` → 1200; `…8п` → 800; `…12.5п`, `…12,50п` → 1250; без хвоста → `None`
- [ ] Хелпер, `enrich` и замена **не** вызывают `parse_load_code_from_name` (тест на них; `_build_order_data` вызывает его законно для прайса, `commercial_service.py:164`)
- [ ] Набрано `12,5п` → в кэше номенклатуры ищется `…12,5п` (убрано `int(float())`)
- [ ] Производство: `normalize_load_code(1250 // 100) == 12`, не 13
- [ ] Цена строки с `load_class = 1250` та же, что для 12 (`load_code_for_price_match`)
- [ ] Уже сохранённые строки «12п» массово не переписываются (никаких миграций)

**Размер:** M. **Риски:** (1) `weight`, `length_dm_raw`, сортировка и `cache_key` строятся по `load_code`; менять только `name` и `load_class`; (2) `plate_parser_service` кэш — ключи `plate_load_details` содержат float 12.5; (3) где-то в запасных ветках 13 (`commercial_offer.py:326`, `unpriced_plate_replacements.py:105,160`, `invalid_width_lines.py:78`) остаётся — не трогаем, читают `load_class` первым.

#### NG-103: `REASON_VARIANTS` и обновление reason/hint в `sweep`

**Файлы:** `core/guid_gate.py` (`REASON_VARIANTS`, `HINT_VARIANTS`, `MissingItem.candidates`); `core/guid_demand.py` (`sweep_guid_demand`); `tests/test_guid_gate.py`; `tests/test_guid_demand.py`. 4 файла.

**Приёмка:**
- [ ] `51-12-12п` при только `Пв-7.1` и `Пв-7.2` (стандартной нет) → `REASON_VARIANTS`, `hint` с именами особых (до 3 + «и ещё N»), `candidates` заполнены
- [ ] `field_for_reason(REASON_VARIANTS) == FIELD_GUID` (ключ очереди и тип `field` не меняются)
- [ ] Очередь: строка открыта с этим reason/hint; `GET /nomenclature/tasks` отдаёт hint с «особые: …» (фронт не меняется)
- [ ] `sweep` при той же `field` обновляет reason/hint, если они изменились (было «нет GUID 1С» → стало «особые»), и не трогает, если совпали; существующие `test_guid_demand` (178, 198, 237) зелёные
- [ ] `REASON_DUP` не используется для variants

**Размер:** S. **Риски:** `Create1cTaskOut` без `reason` — текст только в `hint`; следить, чтобы hint не стал длиннее ~200 символов.

#### NG-104: `explain_missing_guid` для плит

**Цель.** Причина, «та же карточка», число номеров, срез (D-slice, допущение 3), подпись различия (D-label).

**Файлы:** `core/guid_explain.py` (новый: `GuidExplain`, плиты); `tests/test_guid_explain.py` (новый). Читать до кода: `core/commercial_pricing.py:246-280` (`lookup_plate_price`), `core/config_and_data.py:210-239` (`plate_name_to_prays_variants`), `load_code_for_price_match` (`:317-330`). 2 файла (+ при необходимости 1 helper).

**Приёмка (из «Testing Strategy» спеки, плиты):**
- [ ] `22,80-12-8п` → та же карточка `22,8-12-8п`, один номер, подпись «лишний ноль»
- [ ] `29-5,25-8п` → в срезе `29-5,3-8п`; ничего не подставлено
- [ ] `49,5-15-8п` → той же нет; срез: ширины 12 и 10,2; ширины 15 нет
- [ ] `66,1-12-12п` → срез: 4 «с выбр», каждая со своим числом GUID (≤30 строк, ближние к набранному выше)
- [ ] Подпись только при одном отличии (буква, хвостовой ноль, одна ширина, одна нагрузка); иначе две строки без подписи
- [ ] `guid_count` и `selectable` не отдаются как «можно» при `guid_count != 1` в менеджерском представлении

**Размер:** M. **Риски:** «та же карточка» = правило поиска цены этого вида (допущение 2), а срез — из `prays_plity`; не смешать прайс (`prices`) и справочник GUID (`prays_plity`): марка есть в прайсе, а GUID нет → «номера нет» (`guid_count = 0`).

#### NG-105: `explain` для свай, мостовых, составных

**Файлы:** `core/guid_explain.py` (или `core/guid_explain_piles.py`); `tests/test_guid_explain.py`. Читать: `core/pile_price_db.py` (`strip_trailing_u_suffix`), `core/nomenclature_guid.py` (`get_by_mark`, `guid_1c_u`), `core/pile_catalog.py`, прайсы мостовых/составных.

**Приёмка:**
- [ ] `C9-35T7` → «в прайсе есть `С9-35T7`, номера нет» (`guid_count = 0`), менеджеру не выбор, экономисту выбор
- [ ] `C9-35T6` не «та же карточка» для `С9-35T7`, виден в срезе
- [ ] `C15-35T7`: две карточки (`guid_count = 2`), менеджеру нельзя
- [ ] Свая без «у»/с «у»: две строки в срезе; для «у» считается только `guid_1c_u`
- [ ] Срез: мостовая и цельная свая — то же сечение

**Размер:** M. **Риски:** кириллица/латиница (`C`/`С`, `T`/`Т`) — единая нормализация из существующих функций, не новая.

#### NG-106: `explain` для ФБС, ступеней, маршей

**Файлы:** `core/guid_explain.py` (+ модуль видов, если разнесено); `tests/test_guid_explain.py`.

**Приёмка:** ФБС — та же ширина и высота (`фбс 24.4.6-т` ↔ `ФБС 24.4.6-T`); ступени и марши — остальные марки своей серии (`ЛС11` ↔ `ЛС11-1`); поиск и срез по тем же правилам; марка не разобралась → среза нет, остаётся поиск.
**Размер:** M. **Риски:** различия написаний ступеней/маршей бывают в серии, а не в хвосте — D-label подписывает только очевидное, иначе две строки.

#### NG-107: API объяснения и поиска

**Цель.** Один сервис, два входа: менеджер (черновик) и экономист (очередь).

**Файлы:** `app/schemas/guid_explain.py` (новый); `app/services/guid_explain_service.py` (новый); `app/api/v1/endpoints/commercial.py` (рядом с `/drafts/{id}/guid-check`, `:681-712`); `app/api/v1/endpoints/nomenclature.py` (рядом с `/tasks`, `:82`); `tests/test_guid_gate_api.py`. 5 файлов.

**Эндпоинты:**
- `GET /api/v1/commercial/drafts/{draft_id}/guid-explain?product_kind=&mark=` (`REQUIRE_ADMIN_OR_MANAGER` + `verify_draft_ownership`) — менеджерское представление
- `GET /api/v1/commercial/drafts/{draft_id}/guid-search?product_kind=&q=`
- `GET /api/v1/nomenclature/guid-explain?product_kind=&mark=` (`REQUIRE_PRICES`) — экономист
- `GET /api/v1/nomenclature/guid-search?product_kind=&q=`

**Приёмка:**
- [ ] Менеджерский ответ не содержит `selectable=true` при `guid_count != 1`; экономистский — для 0 и 2 с выбором строки/карточки
- [ ] Поиск: от трёх знаков, только свой вид, ≤30 строк, плиты — `prays_plity`, прочие — прайс вида и `nomenclature_guid`
- [ ] Чужой черновик → 403/404 как у соседних; менеджеру в экономистский путь нельзя
- [ ] `guid-check` (`GuidCheckResponse`) не меняется

**Размер:** M. **Риски:** роль определяется эндпоинтом; не принимать `role` из параметров.

#### NG-108: UI панель

**Файлы:** `frontend/src/features/nomenclature/components/GuidExplainPanel.tsx` (общая, новая); `frontend/src/features/nomenclature/types/nomenclature.ts`; `frontend/src/features/nomenclature/api/nomenclatureApi.ts`; `frontend/src/features/commercial-offer/api/commercialOfferApi.ts` + `hooks`; `GuidExplainPanel.test.tsx`. 5-6 файлов (граница M).

**Приёмка:** причина русская; «та же карточка» с числом номеров всегда; срез ≤30 строк, не нажимается пока нет замены (кнопки — в NG-208/NG-303); поиск от трёх знаков с дебаунсом; строка с `guid_count = 0` видна, помечена «номера нет».
**Размер:** M. **Риски:** `.cursor` правила UI/стили проекта — смотреть соседний `Alert`/`Button`.

#### NG-109: Подключение панели

**Файлы:** `frontend/.../CalculationResultStep.tsx` (`:464-477`, `item.reason` → раскрывающаяся панель), `frontend/.../price-desk/components/TasksSection.tsx` (`:53-66`, строка «Завести в 1С»); их `.test.tsx`. 4 файла.
**Приёмка:** предупреждение шага 3 по-прежнему видно и предупреждает «Счёт в 1С не уйдёт…»; для строк очереди hint остаётся; панель раскрывается; существующие vitest (`CalculationResultStep.test.tsx:971`, `PricesView.test.tsx`) зелёные.
**Размер:** S.

**Checkpoint 1**
- [ ] `venv/bin/pytest tests/test_guid_gate.py tests/test_kp_db_nomenclature.py tests/test_guid_demand.py tests/test_guid_explain.py tests/test_guid_gate_api.py tests/test_commercial_service_order_data.py -q`
- [ ] `cd frontend && npm run test -- src/features/commercial-offer src/features/price-desk src/features/nomenclature && npm run typecheck`
- [ ] Показ пользователю панели на локальных данных (без записи в КП)

---

### Phase 2 — подробно

#### NG-205: `update_kp_from_order_data` через внешнее соединение

**Файлы:** `core/kp_persistence_service.py` (`:330-661`: параметр `_external_conn`, владение соединением, без `commit/rollback/close` при внешнем); `core/kp/offers_write.py` (`:82-115` проброс); `app/repositories/kp_repository.py` (`:85-121`); `tests/test_kp_persistence_mixed.py`. 4 файла.
**Приёмка:**
- [ ] Без `_external_conn` поведение и тесты (`test_kp_persistence_mixed.py`, `test_kp_persistence_piles.py`) не меняются
- [ ] С `_external_conn`: два КП одним соединением и одним `commit`; ошибка на втором — откат обоих (ничего не записано)
- [ ] Статус «в архиве» или «на согласовании» проверяется как раньше
**Размер:** M. **Риски:** `rollback()` в `except` (`:657-658`) не должен откатывать чужую транзакцию при внешнем соединении — только пробросить.

#### NG-201: История изменения марки (D-history)

**Файлы:** `core/kp_db_schema.py` (таблица `kp_mark_history`: `id, kp_id, line_id, product_kind, mark_old, mark_new, unit_price_old, unit_price_new, changed_by, changed_at, source`; CREATE IF NOT EXISTS по образцу соседних); `core/kp_persistence_service.py` (запись в `update_kp_from_order_data` при совпавшем `line_id` и изменившейся марке/имени; `_load_existing_lines_by_line_id`, `:664-681`, нужен `mark` и цена); `core/kp/offers_write.py`, `app/repositories/kp_repository.py`, `app/services/commercial_draft_lifecycle.py` (необязательный `changed_by`); `tests/test_kp_persistence_mixed.py` (+ новый тест истории). 6 файлов — разбить на два коммита: схема+запись, затем `changed_by`.
**Приёмка:**
- [ ] Любое изменение марки при обычном сохранении КП (по `line_id`) → запись: было, стало, цена была и стала (если есть), кто, когда
- [ ] Количество, цена без смены марки — записи нет
- [ ] Запись в той же транзакции (откат КП → записи нет)
- [ ] Для замены на шаге 3 — запись тоже появляется (при сохранении черновика в архив, источник `edit`) — см. допущение 16
- [ ] Миграция аддитивная, старые базы поднимаются
**Размер:** M. **Риски:** «кто» берётся из текущего пользователя запроса (сейчас не пробрасывается — C6); если нет — пишем `NULL` и `source`, не падаем.

#### NG-202: Запись о замене на карточке КП

**Файлы:** `app/schemas/archive.py` (`:139`, `:198`), `app/services/archive_service.py` (детали КП, последняя запись), `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx` (+ `types/archive.ts`), `OfferDetailsDrawer.test.tsx`. 5 файлов.
**Приёмка:** на карточке КП видна запись «марка X заменена на Y, дата, кто»; отсутствие записей — блока нет; PDF/XLSX без отдельной пометки (Open Question 1).
**Размер:** S (граница M).

#### NG-203: Построитель строки замены

**Файлы:** `core/guid_replace.py` (новый: `build_replacement`), `tests/test_guid_replace.py` (новый). 2 файла.
**Приёмка (D-replace, допущения 1, 4, 7, 8):**
- [ ] Плита: `name` (то, что выбрал человек), `load_class` через `load_class_from_plate_name` (1250 для `12,5п`), `length_m`, `width_m`, `length_dm_raw`, `weight`, `nomenclature_id`, `unit_price` (прайс), `discounted_price = unit_price × (1 − discount_percent / 100)`; `qty`, `line_id`, `status`, `plan_id` те же
- [ ] Остальные виды: `mark`, `unit_price`, `discounted_price`; класс бетона не меняется; нет цены выбранной марки в классе бетона строки → отказ с русским `detail` (плита и ступень от класса не зависят)
- [ ] Свая «у»: в КП ложится выбранное написание
- [ ] Договорная строка (`price_source = oneoff`): менеджеру отказ; экономисту можно с флагом «договорная заменяется прайсом»
- [ ] Скидка не применяется дважды
- [ ] Нет вызова `parse_load_code_from_name`
**Размер:** M. **Риски:** вес плиты — из какого источника (`resolve_kp_line_weight_kg`, `commercial_service.py` ~`:197-201`); использовать тот же, что при разборе.

#### NG-204: Деньги для окна «было / станет»

**Файлы:** `core/guid_replace.py` или `core/guid_replace_money.py`; `tests/test_guid_replace.py`. 2 файла. Читать до кода: как `update_kp_from_order_data` считает `subtotal/vat/total` (`kp_persistence_service.py` ~`:480-600`), `compute_totals_from_metadata` (`commercial_draft_lifecycle.py:~1093`).
**Приёмка:**
- [ ] `preview_replacement(...)` по каждому КП: цена за штуку со скидкой, сумма строки, итог КП, было и станет; строка доставки только если сумма изменилась
- [ ] Два КП с разной скидкой дают два итога; количество то же
- [ ] Числа окна совпадают с тем, что запишет `update_kp_from_order_data` (один тест сравнивает)
**Размер:** M. **Риски:** расхождение округлений окна и записи — одна функция расчёта итога для обоих.

#### NG-206: Сервис замены в КП строки очереди

**Файлы:** `app/services/guid_replace_service.py` (новый); `app/services/guid_queue_sync.py`; `tests/test_guid_replace.py`. 3 файла.
**Порядок (спека, «Две базы»):** (1) прочитать строку очереди и её `kp_ids`; (2) для каждого КП: статус не «в архиве» → пропуск с русским текстом; марки нет → пропуск; плита в плане/производстве → пропуск (Q2); (3) собрать новый `order_data` через NG-203; (4) один `sqlite3.Connection` на `plita`, `update_kp_from_order_data(..., _external_conn=conn)` для каждого; один `commit`; (5) `recalc_promise_after_edit` по каждому записанному КП (D-promise, после коммита); (6) сверка через `sync_guid_queue_with_kp` (`try/except`, «очередь обновится при открытии страницы»).
**Приёмка:**
- [ ] `C14-40T4` → `С14-40T4` в КП 47 и 162: оба записаны одним `commit`, скидки свои, итоги разные; старая строка закрыта, новая открыта (номера нет)
- [ ] Один КП «на согласовании» пропущен с текстом, остальные записаны; строка очереди остаётся, если марка осталась в КП вне «в архиве»
- [ ] Выбор строки без номера/дубль: экономисту можно; у `C15-35T7` выбор одной из двух, задача закрыта сразу
- [ ] Ошибка шага очереди (`pb.db` недоступна): ответ успешный, в логе ошибка
- [ ] Менеджер вызвать не может (роль на уровне API, NG-207)
- [ ] Каждая замена попала в историю (NG-201) с `source = replace_economist`
**Размер:** M. **Риски:** (1) не атомарно между двумя базами — закрыто сверкой; (2) `recalc_on_composition_change` после коммита — если упадёт, КП уже записан (логируется самим `recalc_promise_after_edit`); (3) очень большой список КП — лимит не нужен (десятки).

#### NG-207: API замены экономиста

**Файлы:** `app/schemas/guid_explain.py` (запрос/ответ), `app/api/v1/endpoints/nomenclature.py` (`POST /nomenclature/guid-replace` + `…/guid-replace/preview`, `REQUIRE_PRICES`), `tests/test_guid_gate_api.py`. 3 файла.
**Приёмка:** preview возвращает деньги по КП и список пропущенных с причинами; запись возвращает записанные/пропущенные; сервер сам перепроверяет `guid_count` цели и цену в классе бетона; 403 для менеджера.
**Размер:** S.

#### NG-208: UI замены на «Завести в 1С»

**Файлы:** `TasksSection.tsx`, `GuidExplainPanel.tsx`, `nomenclatureApi.ts`, `useNomenclatureQueries.ts` (`guidCheckKeys` инвалидация уже есть, `:24`), `types/nomenclature.ts`, тесты. 6 файлов (граница M).
**Приёмка:** кнопка «Заменить» на строке среза/поиска/«карточек две»; окно: цена за штуку со скидкой, сумма строки, итог КП, было/станет по строке на КП; договорная — отдельная строка «договорная цена будет заменена ценой прайса»; после записи список обновлён, колокольчик не гасится.
**Размер:** M.

**Checkpoint 2**
- [ ] `venv/bin/pytest tests/test_kp_persistence_mixed.py tests/test_kp_persistence_piles.py tests/test_guid_replace.py tests/test_guid_demand_reconcile.py tests/test_guid_gate.py tests/test_archive_invoice_export.py -q`
- [ ] vitest `price-desk`, `nomenclature`, `commercial-archive`; `npm run typecheck`
- [ ] Сквозной локальный сценарий (одна копия БД): `C14-40T4` → `С14-40T4` в двух КП; показать пользователю окно «было / станет»

---

### Phase 3 — подробно

#### NG-301: Сервис замены в черновике

**Файлы:** `app/services/commercial_draft_lifecycle.py` (новый метод рядом с `patch_order_line`, `:350-412`, и `persist_order_and_metadata`); `app/services/commercial_workflow_service.py` (проброс, рядом `:209`); `tests/test_commercial_draft_line_edit.py` (+ тесты). 3 файла.
**Приёмка:**
- [ ] Заменяются **все** строки этой марки черновика (допущение 10); запись сразу на сервер; несохранённое количество формы не подхватывается
- [ ] Сервер отказывает, если `guid_count != 1`, договорная цена, нет цены в классе бетона
- [ ] `22,80-12-8п` → `22,8-12-8п`: плашка гаснет, цена = прайс со скидкой черновика
- [ ] `29-5,25-8п` → `29-5,3-8п`: новые ширина, вес, `nomenclature_id`, цена
- [ ] **Очередь не трогается** и `recalc_on_composition_change` не вызывается (черновик)
- [ ] `_invalidate_breakdown(metadata)` как у других правок
**Размер:** M. **Риски:** размерная строка `plate_order_ctx` — черновик плит хранит и структуры оптимизации; как `patch_order_line` с `source_text` обновляет их — перечитать `_replace_line_from_source_text` до кода и решить: заменять текстом (как сегодня) или прямой записью (спека D-replace: «без повторного `LIKE` и `canonical_name`»).

#### NG-302: API замены менеджера

**Файлы:** `app/api/v1/endpoints/commercial.py` (`POST /drafts/{id}/guid-replace` и `…/preview`; `REQUIRE_ADMIN_OR_MANAGER`, `verify_draft_ownership`), `app/schemas/guid_explain.py`, `tests/test_guid_gate_api.py`. 3 файла.
**Приёмка:** preview — деньги для одного КП (черновика); запись — обновлённый черновик; чужой черновик нельзя; цель с `guid_count != 1` → 4xx с русским `detail`.
**Размер:** S.

#### NG-303: UI замены на шаге 3

**Файлы:** `CalculationResultStep.tsx`, `GuidExplainPanel.tsx`, `commercialOfferApi.ts`, `hooks/useGuidCheckQuery.ts` (инвалидация `guidCheckKeys.draft`), тесты. 5-6 файлов.
**Приёмка:** менеджер видит «Заменить» только на строках с одним номером; окно «было/станет» до записи; без подтверждения марка не меняется; после — плашка гаснет, пересчитаны итоги; `С9-35T7` без номера не нажимается; «карточек две» не нажимается; «у»: две строки.
**Размер:** M.

#### NG-304: Регрессия

**Файлы:** `tests/test_archive_invoice_export.py`, `tests/test_commercial_draft_oneoff_price.py`, `tests/test_guid_demand_reconcile.py`. 3 файла.
**Приёмка:** замена менеджера очередь не трогает; «Отправить в 1С» недоступна при одной строке без номера (`_invoice_export_block`); договорная цена у менеджера не меняется; тест «ничего не выбрали — КП не меняется» (D-gap).
**Размер:** S.

**Checkpoint 3**
- [ ] Команды из «Commands» спеки зелёные (см. NG-404)
- [ ] Показ пользователю сквозного сценария менеджера и экономиста

---

### Phase 4 — до выкладки

#### NG-401: Сухой прогон сверки на копии прода и показ пользователю (D-reconcile-run)

1. **С согласия пользователя** снять копию (действие на проде, только чтение источника): `docker exec plan_zavod-backend-1 python3 -c "…sqlite3.connect('file:/data/pb.db?mode=ro',uri=True).backup(…)"` в `/tmp` контейнера, `docker cp` на локаль, затем удалить временные файлы в контейнере. WAL-файлы руками не копировать.
2. Локально: `.venv/bin/python scripts/reconcile_guid_demand.py --pb-db <копия> --plita-db <копия> --report /tmp/reconcile-dryrun.txt` (dry-run).
3. Сверить с таблицей спеки: `97-10,2-8п` (КП 124) закрыта; `ЛС11` (КП 126) закрыта, новая не открыта; `85-5,2-8п` (КП 206) закрыта; три плиты сужены до КП 55 и 57; `С120.35-12` перепривязана 138 → 9, открыта; остальные 52 без изменений.
4. `--apply` на копии, затем dry-run ещё раз: «изменений нет».
5. Показать отчёт пользователю; **до его «да» вызовы NG-002/003 на проде не включаются** (Q1).

**Приёмка:** [ ] отчёт совпадает с таблицей; [ ] повторный прогон пуст; [ ] пользователь подтвердил. **Размер:** S.

#### NG-402: Диф ворот по всем КП копии

Скрипт-одноразовка (в `/tmp`, в репозиторий не класть): для каждой строки каждого КП копии `check_invoice_guids` на старом коде (git stash/worktree) и на новом; список КП, у которых «найдено → не найдено» (особенно `LIKE` → `REASON_VARIANTS`). Эти КП пользователь видит до выкладки: они перестанут отправляться, пока менеджер не выберет карточку.
**Приёмка:** [ ] список приложен к отчёту; [ ] `51-12-12п Пв-7.1` в КП 206 и 210 по-прежнему найдены (точное имя карточки). **Размер:** S.

#### NG-403: Один тестовый счёт в 1С (Open Question 4)

КП с `Плиты ПБ 51-12-12,5п` и GUID `0de0f4a7-dc8e-11ea-8f69-04d9f5387845` (текст = имя карточки). Отправляет пользователь/бухгалтер в тестовый контур 1С; смотрим, как 1С разбирает счёт. Без этого релиз D-name125 не выпускать. Для 21 сохранённой строки «12п» пара «текст 12п + GUID 12,5п» остаётся (в 1С она ещё не уходила).
**Размер:** S.

#### NG-404: Полная регрессия и починка КП 206 и 210

- Команды спеки: `venv/bin/pytest tests/test_guid_gate.py tests/test_kp_db_nomenclature.py tests/test_guid_demand.py tests/test_nomenclature_import.py tests/test_commercial_draft_oneoff_price.py tests/test_archive_invoice_export.py -q` + новые файлы; `cd frontend && npm run test -- src/features/commercial-offer src/features/price-desk && npm run typecheck && npm run build`.
- После выкладки (не код): менеджер заново вводит `51-12-12п` в КП 206 (поз. 10) и 210 через «Редактировать» либо заменяет по срезу → строка `…12,5п`, GUID `0de0f4a7-…`, 1250, `nomenclature_id` заполнен (D-repair-206-210). Проверка — чтением (ro).
- Обновить `ai_docs/changelog/CHANGELOG.md` и статус спеки (PLAN ✅/IMPLEMENT) — только после подтверждения пользователя и отдельной просьбой (документатор).
**Размер:** S.

---

## 5. Risks and Mitigations

| Риск | Влияние | Что делать |
|------|---------|------------|
| Правка `LIKE` блокирует «Отправить в 1С» у КП, которые сегодня находили карточку молча | Высокое | NG-402 до выкладки; список КП пользователю |
| Сверка закроет строку, которую бухгалтер уже начал заводить | Среднее | Не удаляет, а закрывает (`state resolved`); новая открывается, если марка вернётся; сухой прогон на копии |
| Запись в две базы не атомарна | Среднее | КП первыми одним `commit`; очередь — сверкой; все вызовы очереди в `try/except` (C4) |
| `GET /nomenclature/tasks` пишет | Низкое | Запись только при изменении; тест «повторные GET не пишут» |
| Замена размеров плиты «в плане» | Высокое | Отказ с русским текстом (Q2) |
| Окно и запись считают деньги по-разному | Среднее | Одна функция итога; сравнительный тест (NG-204) |
| Менеджер или клиент обходит правило «нужный номер один» | Высокое | Повторная проверка на сервере в NG-301/NG-206 |
| 1С по-разному отнесётся к тексту `12,5п` и GUID | Среднее | NG-403 до выкладки |
| `plate_parser_service` держит `int(float())` | Среднее | Правка в NG-102 с тестом на одностороннее правило |
| История без «кто» | Низкое | `changed_by` необязателен, пишем `NULL` |
| Цифры в спеке разошлись с проверкой («девять» → 11 имён) | Низкое | «Замечания к спеке»; на планирование не влияет |

---

## 6. Замечания к спеке (спеку не правил)

1. **«Девять» имён «12п»** — по чтению прода 2026-11-12: 11 различных имён и 23 строки (21 + 2 с `Пв-7.1`) в КП 132, 167, 172, 206, 210; плит в базе 1414 (спека: 1413). На решения не влияет.
2. **«Разбор строки возвращает два имени»** — достаточно `canonical_name` карточки и набранного имени у вызывающего; отдельное второе поле не нужно (C6). Совместимо со спекой по смыслу.
3. **«Прежние `test_lookup_like_*` переписаны»** — оба сегодня про одного кандидата и остаются зелёными; нужны новые тесты, не переписывание (C3).
4. **Тест «нет `parse_load_code_from_name`»** — не на весь `_build_order_data` (он вызывает его для прайса, `commercial_service.py:164`), а на хелпер, замену и `enrich` (C1).
5. **`plate_parser_service.py:214` `int(float(load_code))`** противоречит одностороннему правилу D-load12 («набрано 12,5п — на 12п не откатываемся»); исправление добавлено в NG-102.
6. **D-send «кнопка недоступна»** уже реализована как `_invoice_export_block` (текст блокировки), доработок нет (C6).
7. **Защита плит «в плане/в производстве»** при замене спекой не оговорена (Q2).

---

## 7. Вопросы к пользователю

Только то, чего нет ни в коде, ни в спеке. У каждого — рекомендация.

**Q1. Как выкладывать срез 0 (сверку)?** Спека: первый прогон на проде показывается пользователю «до включения», но как «включается», не сказано. Вызовы после сохранения/удаления и в `/nomenclature/tasks` (NG-002/003) сами сделают первый прогон на проде.
- **Рекомендую:** выкладывать срез 0 **отдельным релизом** раньше остальных, только после вашего «да» по отчёту сухого прогона на копии (NG-401). Без флага в настройках (новых механизмов нет); срезы 1–3 выкладывать после. Почему: так D-order и D-reconcile-run соблюдаются без лишнего кода.

**Q2. Плита в плане или в производстве: разрешать замену марки?** Код (`kp_persistence_service.py:889-947`) молча перепишет имя, размеры и `load_class` строки, а план останется на старых размерах. Спека требует сохранять статус и `plan_id`, но про размеры молчит.
- **Рекомендую:** **запретить** для такой строки (пропуск КП/строки с русским текстом «плита уже в плане, замените вручную»), как уже запрещено удаление. Почему: расхождение КП и плана хуже, чем отказ; такие случаи редки (замена идёт у КП «в архиве» до производства).

**Q3 (из спеки, п. 1). «Документ устарел».** Подтвердить: флага нет, на карточке КП запись о замене с датой (PDF/XLSX и так собираются при каждом скачивании).
- **Рекомендую:** подтвердить без флага. Почему: хранимого PDF нет; флаг пришлось бы привязывать к факту скачивания, которого программа не видит.

**Q4 (из спеки, п. 2). Колокольчик по прежней марке не гасится.**
- **Рекомендую:** подтвердить. Почему: уведомление — разовая весть о событии; закрытие задачи видно в очереди; гасить — отдельная работа с дедупом.

**Q5 (из спеки, п. 3). Скрипты сопоставления с латинской `C`** (`scripts/build_guid_queue.py`, `_tmp_guid_match.py`; книга «Сопоставление GUID 1С и программа.xlsx»).
- **Рекомендую:** отдельная задача после среза 3 (в этот план не входит). Почему: спека прямо исключает правку импорта и скриптов («Сначала спроси»), а срез 0–3 от них не зависит.

(Тестовый счёт в 1С — не вопрос, а пункт проверки NG-403.)

**Решения пользователя (2026-11-12):**
- Q1: сверка выкладывается отдельным релизом раньше остальных срезов, после показа отчёта сухого прогона.
- Q2: замену плиты «в плане»/«в производстве» **запретить** с русским текстом (понадобится в срезе 2; пока только записано).
- Q3: без флага «документ устарел».
- Q4: колокольчик по прежней марке не гасим.
- Q5: скрипты сопоставления с латинской `C` — отдельной задачей после среза 3.

**Отметки реализации (отклонения от плана/спеки):**
- Старые функции `lookup_nomenclature_by_plate_name` / `count_nomenclature_by_plate_name` остались тонкими обёртками над `find_plate_card` (сигнатуры не менялись); `count…` считает и найденные через `LIKE`.
- В `NomenclatureQueueService` сверка не запускается, если передан явный `db_path` без `plita_db_path` (тесты и скрипты).
- В `hint` для `REASON_VARIANTS` — полные имена карточек (первые три и «и ещё N»).
- `GuidExplain` отдаёт `same_cards` (список), а не одну карточку: у сваи «та же карточка» — две записи (без «у» и с «у»).
- Панель на фронте грузит данные после раскрытия. Кнопки «Заменить» — у экономиста на любой строке среза, у менеджера только при одном номере.
- Если у выбранной плиты ровно один GUID, сервер пишет его в строку сам: клиент номер не присылает.
- Пользователь 2026-11-12: когда у имени две карточки, экономист выбирает их двумя кнопками с именами карточек (как в 1С), без поля GUID. Если имена совпадают, на кнопке хвост GUID — последние 8 символов. Клик передаёт GUID этой карточки в уже существующую замену. Объяснение, срез и поиск отдают `cards: [{name, guid}]`.
- После замены экономиста обновляются задачи и объяснение. Проверка GUID черновика (колокольчик) не сбрасывается.
- Окно менеджера считает черновик без номера КП (`kp_id = 0`) и пишет «Количество», не «КП №0». Договорная строка отклоняется раньше проверки «один номер».
