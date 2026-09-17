# Spec: Контур GUID и цен — раунд 2: ворота счёта, предупреждения, вкладка «Прайсы и 1С»

**Статус:** согласовано к реализации (решения 17.09.2026)
**Дата:** 2026-09-17
**Идея:** `ai_docs/ideas/guid-price-sync-loop.md`
**План:** `ai_docs/develop/plans/2026-09-17-guid-price-sync-round2.md`
**Связанные:** `docs/specs/1c-integration-tz-v2.md` (Q3–Q7, Q13); план раунда 1 `ai_docs/develop/plans/2026-09-15-guid-price-sync.md` (GPS-001…009 ✅, GPS-010 ⏳ → детализирован здесь); `ai_docs/specs/kp-lint-normalizer-parity.md` (D11, «у»-суффикс); `ai_docs/specs/kp-unpriced-price-drawer.md` (oneoff-цены)

## Objective

Сегодня программа уже хранит GUID не-плит (`nomenclature_guid`, 271 позиция) и плит
(`prays_plity`, 57 тыс.), но **ничто не мешает выставить счёт без GUID** — блок GPS-010
не реализован: функция `get_guid_for_invoice` никем не вызывается, сущности счёта нет.
Параллельно два контура живут в Excel: дубли GUID (58 пар ждут бухгалтерию) и очередь
«💰 ввести цену» (26 позиций) — без живых уведомлений и ввода.

Цель раунда 2:

1. **Ворота счёта:** единый резолвер «у позиции есть GUID для счёта?» по всем 6 видам
   продукции и жёсткая функция-блокировка (подключится к кнопке «В 1С», когда та появится).
2. **Контур предупреждений:** мягкое предупреждение на шаге 3 мастера КП, уведомление
   экономисту при архиве КП, задачи на вкладке до закрытия.
3. **Вкладка «Прайсы и 1С»** (`/prices`, admin + economist): загрузка прайса + загрузка
   выгрузки 1С + живые задачи 🏭/💰 + разрешение дублей — одно место вместо Excel-карусели.
4. **Новый формат 1С:** «Универсальный отчёт с УИД номенклатуры» (.xlsx, с отборами)
   рядом со старым «Прайс-лист» .xls.

Success: счёт без GUID технически невозможен; о пробелах узнают на шаге 3, а не от 1С-ника;
экономист закрывает 🏭/💰/⚠️ задачи в вебе без Excel.

## ASSUMPTIONS (defaults, зафиксированы 17.09)

| # | Допущение |
|---|-----------|
| A1 | Выбранный GUID дубля исчез из 1С → 409 при резолве, задача открывается снова |
| A2 | У дублей цены обычно одинаковые → в карточке дубля цена 1С показывается справочно |
| A3 | После ввода цены 💰 GUID сразу пишется в `nomenclature_guid` со статусом `auto` |
| A4 | Уведомление `kp_guid_missing` — активным пользователям роли `economist`; если ни одного — warning в лог, задача всё равно на вкладке |
| A5 | Колокольчик — один раз на пару (КП, марка); задача на вкладке не сгорает |
| A6 | 💰-очередь и ⚠️-дубли не-плит персистятся в БД при каждом импорте 1С; Excel-очередь остаётся fallback |
| A7 | Во вкладке две кнопки загрузки: «Прайс-лист (.xls)» и «Универсальный отчёт (.xlsx)»; вид продукции детектится по содержимому, при неоднозначности — спрашиваем |
| A8 | Вес/объём из универсального отчёта пишем в справочник весов — бонус, не цель |

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| **D-gate** | Ворота счёта | Одна позиция без GUID блокирует весь счёт; ошибка со списком «что завести» |
| **D-custom** | Custom-позиции (11 марок) | Блокируем без исключений |
| **D-oneoff** | Договорная цена | Не освобождает от GUID |
| **D-u** | «у»-сваи | Нужен `guid_1c_u`; срез по суффиксу «у» той же функцией, что линт (D11) |
| **D-plates** | Плиты | GUID из `prays_plity` с учётом override `plate_guid_choice`; прайс не правим |
| **D-shorts** | Коротыши 4м/5м | Матрицу не расширяем; договорная; блок без GUID как все |
| **D-warn3** | Шаг 3 мастера | Неблокирующий Alert со списком марок без GUID |
| **D-arch** | Архив КП | Уведомление `kp_guid_missing` → economist (A4, A5) |
| **D-tab** | Вкладка «Прайсы и 1С» | `/prices`: прайс + выгрузка 1С + задачи 🏭/💰 + дубли; admin + economist |
| **D-dup** | Дубли | Единый UX/API; решение — бухгалтерия офлайн, клик — экономист; не-плиты → `manual` + аудит; плиты → `plate_guid_choice` |
| **D-fmt** | Форматы 1С | .xls «Прайс-лист» + .xlsx «Универсальный отчёт»; частичная выгрузка → без `disappeared` |
| **D-price** | Ввод цены 💰 | Гибрид: масса — «Загрузка прайса», точечно — вкладка, Excel — fallback; после ввода GUID → `nomenclature_guid` (auto) |
| **D-flow** | Кто первый | 1С первая по времени (карточка из КП); наши марки без файла 1С → 🏭 |
| **D-office** | Заведение карточек | Офис (экономист); спецификация на заводе — вне контура |

## User Stories

- Как **менеджер**, на шаге 3 вижу жёлтый Alert: «Свая С70.35-9у — нет GUID, счёт в 1С
  не уйдёт; экономист будет уведомлён при сохранении в архив» — и спокойно сохраняю КП.
- Как **экономист**, получаю уведомление «КП №12: 2 изделия без GUID», перехожу по ссылке
  на «Прайсы и 1С», вижу их в 🏭-очереди; завожу карточки в 1С, гружу универсальный
  отчёт (хоть с отбором на одну карточку) — задачи закрываются сами.
- Как **экономист**, в 💰-очереди ввожу цену за штуку прямо на вкладке — GUID и цена
  связаны, позиция готова к счёту.
- Как **экономист**, в секции «Дубли GUID» по решению бухгалтерии (почта/звонок) выбираю
  один из двух GUID радиокнопкой, жму «Запомнить выбор» — программа использует его в счётах.
- Как **офис**, при нажатии «Выставить счёт» (когда появится) получаю отказ со списком
  «что завести в 1С», если хоть одна позиция без GUID.

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | FastAPI, SQLite (`pb.db`), openpyxl/xlrd, существующие `nomenclature_guid.py`, `nomenclature_sync.py`, `price_import_queue.py`, `kp_db_nomenclature.py` |
| Frontend | React 19, TS, Vitest, TanStack Query, `Alert`, `PricesView`, `Import1cDialog`, колокольчик уведомлений |
| API | REST `/api/v1/...` |

Новых пакетов нет.

## Commands

```
# Backend
pytest tests/test_guid_gate.py tests/test_universal_report_1c_parser.py \
  tests/test_nomenclature_sync.py tests/test_price_queue.py -q

# Frontend
cd frontend && npm run test -- src/features/commercial-offer src/features/price-desk src/features/nomenclature
cd frontend && npm run typecheck

# Dev: не убивать ./run+logs.sh
```

## Project Structure

```
core/guid_gate.py                    → резолвер GUID по составу КП + assert-ворота (новый)
core/plate_guid_choice.py            → override-таблица выбора GUID для плит (новый)
core/duplicate_candidates.py         → персистентные кандидаты дублей не-плит (новый)
core/price_queue_db.py               → персистентная 💰-очередь (новый)
core/universal_report_1c_parser.py   → парсер «Универсального отчёта» .xlsx (новый)
core/nomenclature_sync.py            → запись duplicate_candidates; partial-режим
core/nomenclature_guid.py            → set_manual_choice (audited)
core/price_import_queue.py           → вынос apply-логики цены в переиспользуемую функцию
core/guid_queue.py                   → общий код очередей 🏭/💰 (из scripts/build_guid_queue.py)
app/api/v1/endpoints/nomenclature.py → tasks / duplicates / resolve; формат-детект импорта
app/services/nomenclature_import_service.py → выбор парсера, partial, веса
app/services/offers_service.py       → уведомление при архиве
frontend/src/features/price-desk/components/PricesView.tsx → секции вкладки
frontend/src/features/commercial-offer/components/steps/CalculationResultStep.tsx → Alert
frontend/src/features/nomenclature/components/Import1cDialog.tsx → переезд + подсказки
tests/                               → unit + API по каждой задаче
```

## Code Style

Слои: роутер тонкий, SQL/матчинг в `core/`, оркестрация в `app/services/`.
Русские `detail`/сообщения. `ensure_schema()` по образцу `nomenclature_guid.ensure_schema`.
Минимальный diff; реюз существующих lookup (`lookup_nomenclature_by_plate_name`,
`get_guid_for_invoice`, `strip_pile_load_suffix`).

Пример резолвера:

```python
def check_invoice_guids(order_lines: list[OrderLine], conn: sqlite3.Connection) -> GateReport:
    """Возвращает GateReport(ready, missing); missing блокирует счёт целиком."""
```

## Design

### Резолвер (`core/guid_gate.py`)

Вход — строки заказа `(product_kind, mark|plate_name, qty)`; выход — `GateReport`:

- `ready: [{line, guid}]`, `missing: [{line, reason, action_hint}]`.
- Не-плиты: `get_guid_for_invoice` по `nomenclature_guid` со статусом `auto|manual`;
  суффикс «у» → требуется `guid_1c_u` (срез — той же функцией, что D11-линт).
- Плиты: `plate_guid_choice` → иначе `lookup_nomenclature_by_plate_name` (`prays_plity`);
  дубль без выбора → missing с reason «выберите GUID (дубль 1С)».
- Custom (11 марок) — обычная проверка, без исключений (D-custom).
- `ambiguous`/`missing` → missing с reason для UI и текста уведомления.

### Ворота

`assert_invoice_guids(order_lines, conn)` → бросает `InvoiceGuidBlockError(report)`
если `missing` не пуст. Подключение — к будущей кнопке «В 1С»; в этом раунде только
функция + тесты + docstring «куда вставить».

### Хранилища (все в `pb.db`, `ensure_schema` идемпотентен)

- `plate_guid_choice(plate_name_norm PK, chosen_guid, chosen_name, decided_by, decided_at, note)`.
- `duplicate_candidates(kind, mark_norm, guid, name, price, first_seen, active, PK(kind, mark_norm, guid))` — пишется из `sync_pricelist` при ambiguous; чистится при резолве.
- `price_queue(guid PK, mark_norm, kind, name, state open|resolved, resolved_by, resolved_price, resolved_at, first_seen)` — пишется из `unmatched_1c` при импорте (идемпотентно).

### Универсальный отчёт (.xlsx)

Структура подтверждена образцом 17.09.2026 (`Пример Универсального отчета с УИД
номенклатуры.xlsx`): строки 2–5 — блок «Параметры:», строка «Отбор:» с фильтром
(если есть), пустая строка, шапка, данные, строка «Итого».

- Шапка плавающая: ищем строку с ячейками «Номенклатура» + «УИД»; колонки маппим
  по именам: «УИД» (GUID), «Код», «Номенклатура», «Вес (числитель)», «Объем
  (числитель)» — по префиксу, т.к. возможны «(знаменатель)»-колонки и объединённые
  ячейки (имя A:C, вес F:G; читаем якорные ячейки).
- Колонки «Количество» и цены в компоновке нет — GUID-синк и вес/объём только.
- Строка «Отбор:» над шапкой → `partial=True`: не считаем `disappeared`, не трогаем
  позиции вне файла (образец — выгрузка с отбором на одну позицию, это норма).
- Строка «Итого» — пропускаем. «Код» — текст с ведущими нулями. GUID — lowercase без
  скобок → нормализуем к upper, как в `pricelist_1c_parser`.
- Не «тот» файл (нет шапки «Номенклатура»+«УИД») → 400 «Похоже, это не универсальный
  отчёт 1С».
- Вид продукции: детект по маркам (regex семейств); неоднозначно → 400 с вопросом (A7).
  Позиции ⏸️/🚫-групп (площадки, перемычки, материалы) синк пропускает как «не ведётся».

### API

| Метод | Назначение | Роль |
|-------|-----------|------|
| `GET /commercial/drafts/{id}/guid-check` | missing-список для Alert шага 3 | роли мастера |
| `GET /nomenclature/tasks` | живые 🏭/💰/⚠️ очереди (общий код с `build_guid_queue.py`) | admin + economist |
| `POST /nomenclature/price-queue/resolve` | ввод цены 💰 → прайс-таблицы (реюз apply из `price_import_queue`) + `nomenclature_guid` auto + закрытие задачи | admin + economist |
| `GET /nomenclature/duplicates` | дубли: не-плиты (ambiguous) + плиты (GROUP BY guid) | admin + economist |
| `POST /nomenclature/duplicates/resolve` | выбор GUID: не-плиты → `manual`+аудит; плиты → `plate_guid_choice`; 409 если кандидат исчез | admin + economist |
| `POST /nomenclature/import-1c` | + автодетект формата (.xls/.xlsx), partial-режим, ответ с `mode` | admin + economist |

### Вкладка «Прайсы и 1С» (`/prices`)

Секции: **Загрузка прайса завода** (как сейчас, 3 типа) → **Выгрузка из 1С** (переезд
`Import1cDialog` из шапки; подсказка про частичные выгрузки) → **Задачи по изделиям**
(🏭 список «завести в 1С» read-only; 💰 список с inline-полем цены) → **Дубли GUID**
(кандидаты радиокнопками + заметка + «Запомнить выбор»). Кнопку «Выгрузка 1С» из шапки
убираем. Менеджер вкладку не видит (как сейчас).

### Уведомления

- Триггер: сохранение КП в архив (`offers_service.create_offer`); резолвер по составу.
- `kp_guid_missing`, payload `{kp_id, seq, marks[]}`, ссылка `/prices`.
- Фан-аут активным `economist` (A4); дедуп по паре (КП, марка) (A5).

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit gate | 6 видов; «у» без `guid_1c_u` → missing; плита с override; дубль без выбора → missing; custom → missing |
| Unit parser | синтетический .xlsx: полный/с отбором/без шапки → ошибка; веса парсятся |
| Unit store | ensure_schema идемпотентен; queue upsert; resolve чистит кандидатов |
| Sync | partial-режим не пишет `disappeared`; ambiguous пишет кандидатов |
| API | tasks/duplicates/resolve: RBAC (manager 403), 409 на исчезнувший GUID, цена ≤ 0 → 400 |
| Service | архив с дыркой → уведомление economist; без дырок → тихо; повторный архив не дублирует (A5) |
| RTL | Alert на шаге 3 (есть/нет), секции вкладки, ввод цены, выбор дубля |
| Regress | старый .xls импорт, price-desk, линт-парити, существующие wizard-тесты |

## Boundaries

- **Always:** реюз существующих lookup/парсеров; идемпотентные ensure_schema; русские
  сообщения; тесты на каждую задачу; бэкап `pb.db` перед миграциями-сидированием.
- **Ask first:** правка `prays_plity`; авто-dedup; смена ролей вкладки; удаление Excel-скриптов.
- **Never:** автовыбор GUID при дубле; запись цены в 1С; блок архива КП; новые зависимости;
  коммит без просьбы; трогать ГСМ/production layout.

## Success Criteria

| # | Критерий |
|---|----------|
| S1 | `assert_invoice_guids` падает со списком, если хоть одна позиция без GUID (все 6 видов, «у»-срез) |
| S2 | Шаг 3 мастера: желтый Alert со списком без GUID; архив не блокируется |
| S3 | Архив КП с дыркой → уведомление economist со списком и ссылкой на `/prices`; один раз на КП-марку |
| S4 | Вкладка `/prices`: 4 секции; Import1cDialog живёт там; шапка без кнопки |
| S5 | Загрузка универсального отчёта .xlsx (полного и с отбором) → GUID дописываются, `disappeared` не срабатывает при отборе |
| S6 | Старый .xls «Прайс-лист» работает как раньше |
| S7 | 💰-задача закрывается вводом цены на вкладке: цена в прайс-таблице, GUID в `nomenclature_guid` (auto) |
| S8 | Дубль не-плиты резолвится кликом → `match_status='manual'`, аудит; дубль плиты → `plate_guid_choice`; счётный резолвер использует выбор |
| S9 | Выбранный GUID исчез из 1С → 409 при повторном резолве, задача снова открыта |
| S10 | Manager получает 403 на tasks/duplicates/resolve; admin+economist — 200 |
| S11 | Вес/объём из отчёта попадают в справочник весов (где мэтч) |
| S12 | Focused pytest + vitest + typecheck зелёные |

## Out of Scope

- Кнопка «В 1С»/сущность счёта — только функция-ворота и точка подключения.
- Офлайн-реестр ПИ/коротышей у экономиста.
- Автоматическое слияние/удаление дублей в 1С (ждёт бухгалтерию).
- E-mail/телеграм-рассылки (только внутренний колокольчик).
- Изменения на стороне 1С; дельта-синк по ТЗ п.2.
- Новые прайс-группы (⏸️ 232) и материалы (🚫).

## Open Questions

_Внутренних нет. Внешние (не блокируют):_
1. Умеет ли «Универсальный отчёт» полную выгрузку без отборов (тогда .xls можно будет выводить)?
2. Ритм reexport из 1С (после каждого заведения / раз в день)?
3. Нужен ли где-то «Код» 1С, кроме GUID?

## Risks

| Риск | Почему | Смягчение |
|------|--------|-----------|
| «у»-срез резолвера разъедется с линтом | Два места с regex | Общая функция среза (D11), тесты на обеих |
| Частичная выгрузка примет vanished за disappeared | Отбор не отличить от полного файла без маркера | Детект блока отборов; при сомнении — partial; в ответе `mode` |
| Экономистов нет/неактивны | fan-out в пустоту | Warning в лог + задача на вкладке не сгорает (A4, A5) |
| Экономист кликнет не того кандидата | Человеческий фактор | Заметка note обязательна UI-подсказкой «по решению бухгалтерии»; аудит chosen_by |
| Дубли плит: override устарел | Карточку удалили в 1С | 409 при резолве; задача открывается снова (A1) |
| 💰-цена по классам свай | 5 классов в прайс-таблице | Резолв идёт через ту же apply-функцию, что Excel-импорт (одинаковая семантика) |
