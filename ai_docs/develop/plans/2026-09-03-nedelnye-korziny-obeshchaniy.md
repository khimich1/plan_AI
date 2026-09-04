# Implementation Plan: Недельные корзины обещаний (подбор срока КП)

> **Спека:** [`ai_docs/specs/nedelnye-korziny-obeshchaniy.md`](../../specs/nedelnye-korziny-obeshchaniy.md) (assumptions locked 2026-09-03)
> **Идея:** [`ai_docs/ideas/nedelnye-korziny-obeshchaniy.md`](../../ideas/nedelnye-korziny-obeshchaniy.md)
> **Дата:** 2026-09-03
> **Статус:** draft — НЕ выполнять до approve. Task 0 (калибровка) решает buffer до любого UI.

## Overview

Недельные корзины обещаний поверх диалога «В производство»: котировка срока из
«план + уже обещанное», двухъярусная бронь (холд на день → жёсткое обещание),
блок «Обещано на эту неделю» в wizard планировщика (уровень 2: причина +
уведомление), погашение при коммите плана. Оптимизатор и подневной гейт
графика поставок не трогаем.

## Architecture Decisions

- **Новый чистый модуль** `core/production/promise_buckets.py` (no I/O) — недельная
  математика проще и честнее подневной симуляции `check_batches`; форк не делаем,
  `check_batches` остаётся для графика поставок.
- **Отдельный журнал** (`kp_promise` + `kp_promise_alloc`), НЕ фантомная занятость в
  `days_info` — план не знает про обещания, котировка инжектит их сама.
- **Расширение существующих транзакций**, не параллельные пути: promise пишется в
  `commit_move_to_production` (паттерн `_external_conn`, offers_write.py:629),
  погашение — в `commit_plan_plates` (plan_commit.py:308).
- **Fail-closed котировка**: если occupancy плана недоступна — ошибка, а не
  «всё свободно» (существующий `_load_occupancy` глотает исключения → ложный
  зелёный; для нового пути это неприемлемо).
- **Гейт по пересчёту**: move-to-production пересчитывает корзины в момент
  перевода, а не доверяет устаревшей котировке; дата раньше promised_date → 4xx
  + ближайшая возможная. Оба пути (`POST /archive`, `PATCH /offers`) через один
  сервис — закрываем существующую дыру.
- **Ручка** `promise_tracks_per_day` в новой таблице `kp_setting` (образец —
  `gsm_setting`, kp_db_schema.py:612), аудит по образцу `day_capacity_override`
  (:641); default 3, cap `TRACKS_PER_DAY_HARD_CAP=5`; применяется только к новым
  расчётам.
- **Редактирование КП с активным promise** (open question спеки → решено):
  пересчёт tracks и окна; при сдвиге promised_date — in-web уведомление менеджеру.

## Task List

### Phase 0: Валидация (fail fast)

- [ ] **Task 0: Калибровочный скрипт буферов (A1)**
  - **Description:** `scripts/validate_promise_buffers.py` по образцу
    `validate_podlozhki_phase0.py`: для 3–5 прошлых заказов из `plita.db`
    сравнивает оценку дорожек (с 1.15 и без) с фактически занятыми дорожками в
    планах, куда эти КП попали. Отчёт с рекомендацией buffer.
  - **Acceptance:** отчёт содержит таблицу «КП → оценка vs факт» и вывод
    «buffer = X»; вывод записан в спеку (assumption 2).
  - **Verify:** `python scripts/validate_promise_buffers.py --db plita.db --report ai_docs/develop/reports/2026-09-XX-promise-buffers.md` отрабатывает без ошибок.
  - **Dependencies:** None
  - **Files:** `scripts/validate_promise_buffers.py` (NEW), отчёт (NEW)
  - **Scope:** S

### Checkpoint: Phase 0
- [ ] buffer зафиксирован в спеке; **human review перед продолжением** —
  если калибровка показывает, что сумма буферов сильно завышает сроки,
  модель корректируется до постройки UI.

### Phase 1: Котировка (read-path)

- [ ] **Task 1: Схема БД — журнал обещаний, уведомления, настройка**
  - **Description:** В `core/kp_db_schema.py` (ensure_schema, идемпотентно):
    `kp_promise` (id, kp_id, tracks_total, promised_date, kind hold|promise,
    status active|consumed|released|expired, created_by, created_at, expires_at),
    `kp_promise_alloc` (id, promise_id, week_start, tracks, status
    active|consumed|overdue), `kp_promise_exclusion` (id, kp_id, plan_id,
    week_start, reason, excluded_by, created_at), `notifications` (id, user_id,
    kind, payload_json, read_at, created_at), `kp_setting` (key, value,
    updated_by, updated_at). Дефолт `promise_tracks_per_day=3` при первом запуске.
  - **Acceptance:** свежая БД и существующая `plita.db` обе проходят
    ensure_schema без ошибок; таблицы и дефолт на месте.
  - **Verify:** `pytest tests/ -k "schema or kp_db" -q` зелёные; smoke — ensure_schema на tmp-БД.
  - **Dependencies:** None (∥ Task 2)
  - **Files:** `core/kp_db_schema.py`, `tests/test_promise_schema.py` (NEW)
  - **Scope:** S

- [ ] **Task 2: Чистое ядро корзин `promise_buckets`**
  - **Description:** `core/production/promise_buckets.py`: `WeekBucket`,
    `PromiseWindow`, `allocate(tracks, weeks)` — целиком в первую неделю с
    `free >= tracks`; если tracks > ёмкости целой корзины — жадное окно от
    первой недели с `free > 0`; `promised_date` = последний рабочий день
    последней недели окна. Построение недель от завтра (частичная первая
    неделя), праздники/extra через `core/work_calendar.py`. Без I/O и `app.*`.
  - **Acceptance:** покрыты кейсы: частичная неделя, праздничная неделя,
    whole-only при фрагментации (tracks=8, free 5/5/15 → третья неделя),
    крупное КП (20 → окно 15+5), `free=max(0,…)` при planned>capacity.
  - **Verify:** `pytest tests/test_promise_buckets.py -q` зелёные.
  - **Dependencies:** None (∥ Task 1)
  - **Files:** `core/production/promise_buckets.py` (NEW), `tests/test_promise_buckets.py` (NEW)
  - **Scope:** M

- [ ] **Task 3: Сервис + endpoint котировки `GET promise-quote`**
  - **Description:** `app/repositories/promise_repository.py` (NEW: чтение
    активных promise/hold, настройка), `app/services/promise_service.py` (NEW:
    сборка котировки — tracks КП, occupancy из `PlanDistributionService`
    (fail-closed!), рабочие дни, ручка, аллокация; соло-дата и соло+конец
    недели), `GET /api/v1/commercial/archive/{kp_id}/promise-quote` в
    `endpoints/archive.py` (roles admin, manager), схемы в `app/schemas/archive.py`.
  - **Acceptance:** ответ по контракту спеки (tracks, solo_*, earliest_start_week,
    window, weeks[{week_start, workdays, capacity, planned, promised, held, free}],
    knob); при недоступной occupancy — 503, не «всё свободно»; второй quote
    учитывает promise первого (после ручной вставки строки в журнал).
  - **Verify:** `pytest tests/test_promise_service.py tests/test_archive_endpoints.py -q`.
  - **Dependencies:** T1, T2
  - **Files:** до 5: repository (NEW), service (NEW), endpoints/archive.py, schemas/archive.py, tests (NEW+edit)
  - **Scope:** M

- [ ] **Task 4: Диалог «В производство» — котировка и полоса недель (UI)**
  - **Description:** `PromiseQuoteBlock.tsx` (NEW: первично «обещать к <дата»,
    вторично начало / «если только его» / «соло + до конца недели»),
    `PromiseWeekStrip.tsx` (NEW: недели план/обещано/холды/свободно) в
    `factory-capacity`; `MoveToProductionDialog.tsx` переключается с
    `useCapacitySnapshotQuery` на quote; drawer «Ёмкость» показывает полосу
    недель вместо мини-календаря (сам `FactoryMiniCalendar` и endpoint
    `capacity-snapshot` остаются для графика поставок — не удаляем).
    Старый red-гейт пока на месте (transitional, снимается в T6).
  - **Acceptance:** диалог показывает 4 числа и полосу; `ProductionEstimateAlert`
    заменён данными quote (tracks из одного источника); vitest на отображение.
  - **Verify:** `cd frontend && npm run test -- --run src/features/factory-capacity src/features/commercial-archive && npm run typecheck`.
  - **Dependencies:** T3
  - **Files:** до 5: 2 компонента (NEW), MoveToProductionDialog.tsx, api/types, тесты
  - **Scope:** M

### Checkpoint: Phase 1
- [ ] `pytest tests/ -q` зелёные; `npm run test -- --run` и `npm run build` чистые
- [ ] Ручная проверка: диалог на живом КП показывает котировку (журнал пуст → promised=0)
- [ ] Review с человеком перед write-path

### Phase 2: Обещание (write-path менеджера)

- [ ] **Task 5: Холды — backend**
  - **Description:** `POST/DELETE /api/v1/commercial/archive/{kp_id}/promise-hold`:
    создание по свежему пересчёту корзин (не по кэшу клиента), `expires_at` =
    конец текущего дня (локальное время), ленивый expire при чтении,
    снять может владелец или admin. Холд пишет аллокации как promise, но
    `kind=hold` и в `free` не вычитается — только в счётчик `held`.
  - **Acceptance:** холд виден в чужих котировках как `held`, не влияет на `free`;
    после полуночи (expires_at < now) читается как `expired`; повторный холд того
    же КП замещает старый.
  - **Verify:** `pytest tests/test_promise_service.py -q` (hold-кейсы).
  - **Dependencies:** T3
  - **Files:** endpoints/archive.py, services/promise_service.py, repositories/promise_repository.py, tests
  - **Scope:** M

- [ ] **Task 6: Гейт и атомарная запись promise при переводе (оба пути)**
  - **Description:** `commit_move_to_production` (offers_write.py:629) принимает
    опциональный promise-payload и пишет `kp_promise`+alloc в той же транзакции
    (паттерн `_external_conn`); конвертация холда — тот же путь (hold→promise
    в одной tx). `archive_service.move_to_production`: гейт = пересчёт корзин,
    дата < promised_date → `ArchiveValidationError` с ближайшей датой; старый
    `_enforce_capacity_gate_for_terms` с этого пути уходит (остаётся в графике
    поставок). `offers_service.move_to_production` получает тот же гейт —
    закрываем дыру PATCH-пути.
  - **Acceptance:** ранняя дата → 4xx + earliest на обоих путях; при ошибке
    записи promise — ROLLBACK всего (статус/срок/freeze не применяются);
    тест atomicity расширен promise-строками.
  - **Verify:** `pytest tests/test_move_to_production_atomicity.py tests/test_archive_service.py tests/test_archive_endpoints.py -q`.
  - **Dependencies:** T5
  - **Files:** до 5: offers_write.py, archive_service.py, offers_service.py, promise_service.py, tests
  - **Scope:** M

- [ ] **Task 7: Холд в UI — кнопка «Закрепить срок» и бейдж**
  - **Description:** В `MoveToProductionDialog`: кнопка «Закрепить срок»
    (создаёт холд по показанной дате), состояние «срок закреплён до сегодня»,
    «В производство» из холда без повторного ввода. Бейдж холда в карточке
    (`OfferDetailsDrawer`) и в строке списка архива; поимённая видимость
    (кто закрепил) в тултипе/деталях. Инвалидация quote после hold/convert.
  - **Acceptance:** сценарий quote → закрепить → бейдж → перевести одной
    кнопкой проходит; чужой холд виден как «холды N» в полосе недель.
  - **Verify:** vitest диалога/бейджа; `npm run typecheck`.
  - **Dependencies:** T4, T5, T6
  - **Files:** до 5: MoveToProductionDialog.tsx, OfferDetailsDrawer.tsx, список архива, hooks, тесты
  - **Scope:** M

### Checkpoint: Phase 2
- [ ] `pytest tests/ -q` зелёные; frontend tests+build чистые
- [ ] E2E вручную: два КП подряд — второй видит обещание первого (накопление);
  холд сгорает на следующий день; ранняя дата → 4xx с подсказкой
- [ ] **Менеджерский контур закрыт** — уже даёт ценность без Phase 3

### Phase 3: Сторона плана (планировщик, уровень 2)

- [ ] **Task 8: Погашение и overdue в `commit_plan_plates`**
  - **Description:** В `core/plan_commit.py` после успешной пометки: для КП
    плана, у которых не осталось незапланированных позиций, — аллокации
    покрытых недель → `consumed`; promise → `consumed`, когда все аллокации
    consumed. Аллокации недель, полностью покрытых коммитом, по невошедшим
    обещанным КП → `overdue` (не исчезают молча). В той же транзакции коммита.
  - **Acceptance:** после коммита недели свободное место корзин не занижено
    (нет двойного счёта); невошедшее обещанное КП → overdue, читается сервисом.
  - **Verify:** `pytest tests/test_promise_service.py tests/test_plan_commit*.py -q`.
  - **Dependencies:** T6
  - **Files:** core/plan_commit.py, services/promise_service.py, repositories/promise_repository.py, tests
  - **Scope:** M

- [ ] **Task 9: Данные «Обещано на неделю» для wizard**
  - **Description:** Расширить `GET /production/kp-candidates` (production.py:329)
    или соседний ответ: по каждому кандидату — признак/мета обещания
    (promised_date, неделя, состояние active/overdue); сводка по неделям для
    выбранного диапазона дней. Схемы в `app/schemas/production.py`.
  - **Acceptance:** wizard получает обещанные КП выбранной недели одним
    запросом вместе с кандидатами; overdue помечены.
  - **Verify:** `pytest tests/ -k "kp_candidates or promised" -q`.
  - **Dependencies:** T8
  - **Files:** endpoints/production.py, service слой production, schemas/production.py, tests
  - **Scope:** M

- [ ] **Task 10: `PromisedWeekBlock` в wizard — предвыбор и причина**
  - **Description:** `create-plan-wizard/PromisedWeekBlock.tsx` (NEW): обещанные
    КП недели предвыбраны; снятие галочки (КП целиком или части позиций
    обещанного) требует причины (модалка/инлайн); overdue-блок красным сверху.
    Причины уходят в payload `POST /production/plans/build` (поле exclusions).
    Хук `useCreatePlanWizardState` — предвыбор и сбор причин.
  - **Acceptance:** нельзя снять обещанное без причины; причины включены в
    payload; overdue визуально отличён.
  - **Verify:** vitest блока и wizard-state; `npm run typecheck`.
  - **Dependencies:** T9
  - **Files:** до 5: PromisedWeekBlock (NEW), useCreatePlanWizardState.ts, CreatePlanWizard.tsx, types, тесты
  - **Scope:** M

- [ ] **Task 11: Запись исключений и уведомлений при коммите плана**
  - **Description:** `POST /production/plans/build` принимает `exclusions`;
    при коммите: строки в `kp_promise_exclusion` (kp_id, plan_id, неделя,
    причина, кто) + `notifications` менеджеру-владельцу обещания
    (kind=`promise_excluded`, payload: kp_id, неделя, причина).
  - **Acceptance:** исключение без причины отклоняется 4xx; уведомление
    создано в той же tx, что коммит; журнал читается.
  - **Verify:** `pytest tests/ -k "exclusion or notification" -q`.
  - **Dependencies:** T8, T10
  - **Files:** endpoints/production.py, schemas/production.py, promise_service/plan_commit hook, tests
  - **Scope:** M

### Checkpoint: Phase 3
- [ ] `pytest tests/ -q` зелёные; frontend чистый
- [ ] E2E вручную: обещанное КП предвыбрано в wizard; снятие → причина →
  уведомление в БД; коммит гасит аллокации; пропущенное → overdue красным
- [ ] Review с человеком (уровень 2 — ключевая ставка A2)

### Phase 4: Уведомления, ручка, жизненный цикл

- [ ] **Task 12: In-web уведомления — endpoints + бейдж в шапке**
  - **Description:** `app/api/v1/endpoints/notifications.py` (NEW: GET список
    непрочитанных/всех, POST `/{id}/read`; roles — все авторизованные,
    фильтр по user_id), регистрация в `app/api/v1/router.py`.
    `frontend/src/features/notifications/` (NEW): бейдж-счётчик в шапке,
    поповер со списком, переход к КП.
  - **Acceptance:** менеджер видит уведомление об исключении своего КП;
    прочтение сбрасывает счётчик.
  - **Verify:** pytest endpoints; vitest бейджа; `npm run build`.
  - **Dependencies:** T11
  - **Files:** до 5: endpoints (NEW), router.py, репозиторий, frontend feature (NEW), тесты
  - **Scope:** M

- [ ] **Task 13: Ручка `promise_tracks_per_day` — endpoint + UI в drawer**
  - **Description:** `GET/PUT /api/v1/commercial/settings/promise-tracks-per-day`
    в `endpoints/archive.py` (roles admin, manager): чтение из `kp_setting`,
    запись с updated_by/updated_at, валидация 1..5. В drawer «Ёмкость» —
    кнопка настройки с инлайн-формой и подтверждением; подпись «влияет только
    на новые расчёты». После смены — инвалидация quote.
  - **Acceptance:** PUT меняет котировки новых расчётов, активные обещания не
    тронуты; изменение залогировано (кто/когда).
  - **Verify:** pytest settings-кейсы; vitest формы.
  - **Dependencies:** T4
  - **Files:** endpoints/archive.py, promise_service, drawer-компонент, тесты
  - **Scope:** S

- [ ] **Task 14: Жизненный цикл — удаление и редактирование КП**
  - **Description:** `archive_service.delete_offer` → release активных
    promise/hold (status=released в той же tx). Редактирование состава КП
    (путь edit-in-constructor) → пересчёт tracks/окна активного promise;
    при сдвиге promised_date — уведомление менеджеру.
  - **Acceptance:** удалённое КП освобождает корзины; увеличение состава,
    не влезающее в окно, → новое окно + уведомление; уменьшение — пересчёт без
    уведомления.
  - **Verify:** `pytest tests/test_promise_service.py tests/test_archive_service.py -q`.
  - **Dependencies:** T6, T12
  - **Files:** archive_service.py, promise_service.py, путь редактирования КП, tests
  - **Scope:** M

- [ ] **Task 15: «~N дорожек» на финальном шаге мастера КП**
  - **Description:** В результате мастера commercial-offer показать оценку
    дорожек (reuse `estimateFromLengthM` / общий `ProductionEstimateAlert`) —
    менеджер понимает масштаб заказа до архива. Только плиты.
  - **Acceptance:** финальный шаг показывает «~N дорожек» для плиточных КП;
    простые продукты (сваи/ступени/ФБС) — без оценки.
  - **Verify:** vitest шага; `npm run typecheck`.
  - **Dependencies:** None (формула готова)
  - **Files:** шаг результата мастера, тест
  - **Scope:** S

### Checkpoint: Complete
- [ ] Все acceptance criteria спеки выполнены; `pytest tests/ -q`,
      `npm run test -- --run`, `npm run build` зелёные
- [ ] Регресс: `test_capacity_gate.py`, `test_delivery_schedule*` — график
      поставок не тронут
- [ ] E2E полный цикл: мастер → архив → котировка → холд → перевод →
      wizard (предвыбор) → коммит (погашение) → уведомления
- [ ] Human QA + review перед merge

## Risks and Mitigations

| Риск | Impact | Mitigation |
|------|--------|------------|
| Калибровка: сумма буферов завышает сроки → потеря продаж по скорости | High | Task 0 первым; buffer — параметр, не константа; human gate после Phase 0 |
| Двойной счёт (обещание + план) → корзины врут | High | Погашение в той же tx коммита (T8); тест «нет двойного счёта» в acceptance |
| occupancy недоступна → ложное «всё свободно» | High | Fail-closed в T3 (ошибка вместо пустой занятости) — исправляет существующее поведение `_load_occupancy` для нового пути |
| Гонка двух менеджеров за последнее место недели | Med | Пересчёт+запись в одной tx (T6); SQLite single-writer; проигравший получает 4xx со свежей датой |
| Планировщик привыкает игнорировать причину (A2) | Med | Метрика доли снятий после запуска; при высокой — структурный уровень (фикс-костяк), отдельной дельтой |
| Спекулятивные холды (A3) | Low/Med | TTL + поимённая видимость + счётчик; лимит — только при факте |
| Объём (16 задач) | Med | Фазы автономны: после Phase 2 менеджерский контур уже работает; Phase 3–4 можно катить отдельно |

## Parallelization Opportunities

- **T1 ∥ T2** — схема и чистое ядро не пересекаются.
- **T15** — независим, можно в любой момент после Phase 1.
- **T12 ∥ T13 ∥ T14** — после своих зависимостей, разные файлы.
- Контракты для параллельной работы: схема quote (T3) — до любого frontend;
  поле `exclusions` в build-payload (T11) — согласовать до T10.

## Open Questions

- К пользователю — нет (все продуктовые решения в спеке locked).
- Технические, решаемые в полёте: точная форма payload `exclusions`;
  дедупликация уведомлений при повторных исключениях; часовой пояс TTL холда
  (локальное время сервера).
