# Implementation Plan: ГСМ — обзор флота и журналы (UX-редизайн)

## Overview

Спека: [`../../specs/gsm-fleet-overview-ux.md`](../../specs/gsm-fleet-overview-ux.md)
(утверждена 2026-08-24). Идея: [`../../ideas/gsm-fleet-overview-ux.md`](../../ideas/gsm-fleet-overview-ux.md).

Главный экран `/gsm` становится сводной таблицей машин за период со статусами
и раскрытием строк (журнал ПЛ + лента дней + drawer). Вкладка «Транзакции»
получает журнал. Массовые действия: generate-bulk и export по чекбоксам.
Причина красных дней персистентна (`warning_details`). Индикатор незакрытых
ПЛ до периода защищает цепочку показаний.

Масштаб из БД: 4 активные машины, пик 19 ПЛ/мес, ~170 транзакций/мес —
пагинации нет, всё клиентское.

## Architecture Decisions

1. **Backend-first, контракт из спеки.** Сначала репозиторий → схемы →
   сервисы → эндпоинты (checkpoint A: API доступно через Swagger).
   Frontend стартует только после зелёного checkpoint A.
2. **Агрегат сводки — один SQL в репозитории** (`fleet_overview`), статус —
   чистая функция `_status_of` в сервисе. Солвер не вызывается.
3. **`warning_details` аддитивно.** `WaybillOut.warnings` (`list[str]`) не
   меняется; детали — новое optional-поле. Хранение `warnings_json`
   расширяется до объектов `{code, detail}` только для проблемных дней;
   парсер принимает оба формата (старые строки не ломаются).
4. **Bulk = цикл существующего `generate()`** в одном запросе, per-vehicle
   отчёт `{ok, result | error{code,message}}`; HTTP 200 всегда, кроме
   невалидного периода (400). Первая генерация машины падает с
   `gsm_start_required` → UI ведёт в индивидуальный диалог.
5. **`GsmPeriodView` удаляется последним** — после того как журнал,
   диалог генерации и bulk-бар покрыли его функции (open question спеки).
   `VehiclePeriodStrip` и `WaybillDayDrawer` переиспользуются без изменений.
6. **DI как у соседей:** `get_gsm_overview_service()` в
   `app/dependencies/services.py` по образцу существующих фабрик.
7. **Приёмка генерации — на копии `plita.db`** (конвенция проекта);
   UI-приёмка (чтение) — на рабочей.

## Task List

### Backend

- [x] Task 1: Репозиторий — агрегаты для сводки и списков
  - **Description:** `GsmRepository`:
    `fleet_overview(period_from, period_to)` — по каждой активной машине:
    tx_count/tx_liters/tx_amount/tx_last_date (join card), wb_count,
    wb_km = Σ(odometer_end−odometer_start), wb_fuel_issued, wb_last_date,
    fuel_end_last, draft/confirmed/exported counts, red_days
    (`warnings_json LIKE '%manual_intervention%'`), open_before
    (`status IN ('draft','confirmed') AND date < period_from`);
    `list_transactions` — вариант без `vehicle_id` (все машины, join
    `gsm_fuel_card` для card_number/vehicle_id); `list_waybills` —
    `vehicle_id: int | None` (None = все машины, `ORDER BY vehicle_id, date`).
  - **Acceptance:**
    - [x] Юнит-тесты агрегата на фикстурах: 2 машины, красный день считается,
          open_before только до периода, Σ км = Σ одометрных дельт.
    - [x] Машина без транзакций и ПЛ присутствует в выдаче с нулями.
    - [x] Непривязанная карта (`vehicle_id NULL`) попадает в
          транзакции-всех-машин, не ломает агрегат.
  - **Verification:** `venv/bin/pytest tests/test_gsm_repository.py -q`
  - **Dependencies:** None
  - **Files:** `app/repositories/gsm_repository.py`, `tests/test_gsm_repository.py`
  - **Estimated scope:** M

- [x] Task 2: Схемы + `GET /gsm/transactions`
  - **Description:** `TransactionOut` (ts, card_number, vehicle_id|null,
    service_type, fuel_grade, qty_liters, amount, station_id, address);
    ответ — конверт `TransactionListResponse {rows, total_count,
    sum_liters, sum_amount}` (итоги по отфильтрованному набору — на
    backend); `GsmTransactionService.list_transactions(vehicle_id?, from,
    to, service_type?)`; эндпоинт под `REQUIRE_ACCOUNTING`.
  - **Acceptance:**
    - [x] Фильтры vehicle_id / service_type / период работают по отдельности
          и вместе; без фильтров — весь флот за период.
    - [x] Итоги в конверте совпадают с Σ строк; пустой набор → нули.
    - [x] Сортировка `ts, id`; машина `null` у непривязанной карты.
    - [x] 403 для `manager`/`production`.
  - **Verification:** `venv/bin/pytest tests/test_gsm_transactions_list_api.py -q`
  - **Dependencies:** Task 1 (list_transactions)
  - **Files:** `app/schemas/gsm.py`, `app/services/gsm_transaction_service.py`,
    `app/api/v1/endpoints/gsm.py`, `tests/test_gsm_transactions_list_api.py`
  - **Estimated scope:** S

- [x] Task 3: `GsmOverviewService` + `GET /gsm/overview`
  - **Description:** Статусная модель `_status_of` (6 веток по спеке,
    приоритеты, граница `tx_last_date == wb_last_date` → НЕ needs_generation);
    `liters_diff = round(tx_liters − wb_fuel_issued, 2)`; сборка
    `FleetOverviewRow`; DI-фабрика; эндпоинт.
  - **Acceptance:**
    - [x] Все 6 статусов покрыты unit-тестами, включая приоритет
          red_days > drafts и no_data при пустой машине.
    - [x] 403 для чужих ролей; 400 при `from > to`.
    - [x] На фикстуре: машина с tx и без ПЛ → needs_generation, liters_diff
          скрыт логикой `wb_count == 0` (поле есть, UI решает).
  - **Verification:** `venv/bin/pytest tests/test_gsm_overview_api.py -q`
  - **Dependencies:** Task 1
  - **Files:** `app/services/gsm_overview_service.py`, `app/schemas/gsm.py`,
    `app/dependencies/services.py`, `app/api/v1/endpoints/gsm.py`,
    `tests/test_gsm_overview_api.py`
  - **Estimated scope:** M

- [x] Task 4: `GET /gsm/waybills` — `vehicle_id` опционален
  - **Description:** Query-параметр `vehicle_id: int | None = None`;
    сервис/репозиторий без машины → все ПЛ периода (`vehicle_id, date`).
    Регрессия существующего вызова с vehicle_id.
  - **Acceptance:**
    - [x] Без vehicle_id: ПЛ всех машин, сортировка `vehicle_id, date`.
    - [x] С vehicle_id: поведение не изменилось (существующие тесты зелёные).
  - **Verification:** `venv/bin/pytest tests/test_gsm_waybills_list_all.py tests/test_gsm_generation_api.py -q`
  - **Dependencies:** Task 1
  - **Files:** `app/api/v1/endpoints/gsm.py`,
    `app/services/gsm_generation_service.py`,
    `tests/test_gsm_waybills_list_all.py`
  - **Estimated scope:** S

- [x] Task 5: `POST /gsm/waybills/generate-bulk`
  - **Description:** `WaybillBulkGenerateRequest{vehicle_ids, period_from,
    period_to, force}`; сервис циклом вызывает `generate()`; на машину —
    `{vehicle_id, ok, result | error{code, message}}`; `GsmGenerationError`
    ловится в per-vehicle error; 400 только при невалидном периоде.
  - **Acceptance:**
    - [x] 2 машины: одна без маршрутов (`gsm_routes_required`) → её error,
          вторая сгенерирована; ответ 200.
    - [x] `gsm_start_required` (нет confirmed-истории, нет override) —
          per-vehicle error, не HTTP 422.
    - [x] force поверх confirmed работает как в одиночной генерации.
    - [x] Пустой `vehicle_ids` → 200 с пустым results.
  - **Verification:** `venv/bin/pytest tests/test_gsm_generate_bulk_api.py -q`
  - **Dependencies:** None (существующий generate)
  - **Files:** `app/schemas/gsm.py`, `app/services/gsm_generation_service.py`,
    `app/api/v1/endpoints/gsm.py`, `tests/test_gsm_generate_bulk_api.py`
  - **Estimated scope:** M

- [x] Task 6: Перстистентная причина красных дней (`warning_details`)
  - **Description:** При сохранении дня в `generate()` warnings пишутся
    объектами для проблемных дат (detail из `result.problematic_days`);
    `_parse_warnings_json` → `(codes, details)`; `WaybillOut +=
    warning_details: list[WaybillWarningDetail] | None`.
    `_waybill_out` наполняет оба поля. Старые строковые записи →
    details = [].
  - **Acceptance:**
    - [x] Новая генерация с красным днем: `GET /waybills` отдаёт и код, и
          detail; повторный запрос (после «перезагрузки») — detail на месте.
    - [x] Строковые `warnings_json` из прошлых генераций парсятся,
          существующие тесты зелёные без правок фикстур.
    - [x] Дни без проблем: `warning_details` пуст/`None`.
  - **Verification:** `venv/bin/pytest tests/test_gsm_generation_service.py tests/test_gsm_waybill_edit.py -q`
  - **Dependencies:** None (но после Task 5 меньше конфликтов в файле)
  - **Files:** `app/schemas/gsm.py`, `app/services/gsm_generation_service.py`,
    `tests/test_gsm_generation_service.py`
  - **Estimated scope:** M

**Checkpoint A (гейт):** `venv/bin/pytest tests/test_gsm_*.py -q` и
`venv/bin/pytest tests/ -q` зелёные; Swagger показывает 3 новых эндпоинта;
ручной `curl GET /overview` на рабочей БД возвращает осмысленные агрегаты.

### Frontend

- [x] Task 7: Типы + API + хуки (фундамент)
  - **Description:** `types/gsm.ts`: `GsmTransaction`, `FleetOverviewRow`,
    `VehiclePeriodStatus`, `WaybillWarningDetail`, `BulkGenerateResult`;
    `WaybillListParams.vehicleId` опционален. `gsmApi.ts`:
    `listTransactions`, `getOverview`, `generateWaybillsBulk`.
    `useGsmQueries.ts`: `useGsmOverviewQuery`, `useGsmTransactionsQuery`,
    `useBulkGenerateMutation` (инвалидация overview+waybills+transactions).
  - **Acceptance:**
    - [x] `gsmApi.test.ts` покрывает query-string новых методов и
          опциональность vehicleId в listWaybills.
    - [x] `tsc`/`npm run build` чистые.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/api`
  - **Dependencies:** Checkpoint A (контракт стабилен)
  - **Files:** `frontend/src/features/gsm/{types/gsm.ts, api/gsmApi.ts,
    hooks/useGsmQueries.ts, api/gsmApi.test.ts}`
  - **Estimated scope:** S

- [x] Task 8: `fleetStatus.ts` + `FleetOverviewView` + `FleetOverviewTable`
  - **Description:** Маппинг статусов (ярлык/тон); таблица: чекбоксы,
    колонки по спеке (статус, транзакции, ПЛ дни/км, красные, бак,
    бейдж liters_diff — скрыт при wb_count=0), раскрытие строк
    (lazy), бейдж `open_before` в строке; индикатор хвостов в шапке
    (Σ open_before > 0). Период-пикер, дефолт — текущий месяц.
  - **Acceptance (vitest):**
    - [x] Статусы рендерятся ярлыками по fleetStatusMeta; порядок тонов.
    - [x] Бейдж расхождения: 🟢 при |Δ| ≤ 0.01, 🔴 иначе; скрыт при wb_count=0.
    - [x] Чекбокс выбирает/снимает строку; «выбрать все» не трогает
          no_data строки для bulk-действий.
    - [x] Индикатор хвостов виден только при Σ open_before > 0.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/`
  - **Dependencies:** Task 7
  - **Files:** `frontend/src/features/gsm/lib/fleetStatus.ts`,
    `components/FleetOverviewView.tsx`, `components/FleetOverviewTable.tsx`,
    тесты
  - **Estimated scope:** L

- [x] Task 9: `VehicleWaybillJournal` + `VehicleGenerateDialog`
  - **Description:** Раскрытая строка: шапка с кнопкой «+ Ручной ПЛ»
    (переиспользует существующий `ManualWaybillDialog` — функция уходит
    из удаляемого GsmPeriodView сюда); таблица ПЛ (дата, водитель, маршрут,
    км, бак нач/выд/кон, одометр, статус, warnings с detail из
    warning_details), итоговая строка Σ км/Σ выдано; `VehiclePeriodStrip`
    над таблицей; клик по строке → существующий `WaybillDayDrawer`.
    Диалог генерации: период предзаполнен, fuel/odometer override, force;
    показывает `problematic_days` результата.
  - **Acceptance:**
    - [x] Журнал показывает км из `WaybillOut.km`; итоги из ответа.
    - [x] Красный день показывает detail (mock warning_details).
    - [x] Диалог без override при наличии confirmed-истории → генерация ок;
          без истории → показывает `gsm_start_required` и подсвечивает поля.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/`
  - **Dependencies:** Task 8 (контейнер раскрытия)
  - **Files:** `components/VehicleWaybillJournal.tsx`,
    `components/VehicleGenerateDialog.tsx`, `lib/waybillWarnings.ts`, тесты
  - **Estimated scope:** M

- [x] Task 10: `TransactionsJournalView`
  - **Description:** Таблица транзакций (дата/время, машина, карта, услуга,
    литры, сумма, АЗС), фильтры машина/тип/период, итоги в подвале
    (Σ литров, Σ суммы — из ответа, не считать на клиенте), подсветка
    непривязанной карты (vehicle null) со ссылкой в Справочники→Карты.
    Кнопка «Импорт» остаётся.
  - **Acceptance:**
    - [x] Фильтры меняют query params запроса; итоги из backend-полей.
    - [x] Непривязанная карта подсвечена.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/`
  - **Dependencies:** Task 7
  - **Files:** `components/TransactionsJournalView.tsx`, тест;
    `pages/gsm/GsmPage.tsx` (вкладка)
  - **Estimated scope:** M

- [x] Task 11: Bulk-бар: generate-bulk + экспортный гейт по выбору
  - **Description:** «Сгенерировать выбранные» → `generateWaybillsBulk` →
    отчёт по машинам (Alert со списком ok/error); `gsm_start_required` →
    кнопка открывает `VehicleGenerateDialog` этой машины. «Экспорт zip
    выбранных» → машины с `red_days > 0` исключаются с причиной (данных
    overview достаточно, журналы не догружаем); confirm-текст сводный по
    счётчикам («M из N уже экспортировались»), без per-day деталей;
    один zip на чистые. Порядок строк сводки — реестровый (vehicle id),
    стабильный; индикатор хвостов кликабелен → переключает период на
    предыдущий месяц.
  - **Acceptance:**
    - [x] Ошибка одной машины видна в отчёте, остальные сгенерированы.
    - [x] Красная машина в выборе → исключена из zip, причина показана;
          zip по остальным скачивается.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/`
  - **Dependencies:** Task 8, Task 9
  - **Files:** `components/FleetOverviewView.tsx`, `lib/exportGate.ts`
    (переиспользование), тесты
  - **Estimated scope:** M

- [x] Task 12: Переключение страницы: «Обзор» дефолт, удаление «Период»
  - **Description:** `GsmTabs`: +«Обзор» (дефолт), −«Период»; `GsmPage`
    на вкладки Обзор/Транзакции/Справочники; `GsmPeriodView.tsx` и его
    тест удаляются (функции перенесены в Task 8/9/11); `roleRoutes` и
    `AppRouter` не меняются (тот же `/gsm`).
  - **Acceptance:**
    - [x] `/gsm` открывается на «Обзоре»; вкладки «Период» нет.
    - [x] Все функции старой вкладки доступны (генерация, ручной ПЛ,
          экспорт, правка дня) — чеклист вручную.
    - [x] `GsmPage.test.tsx`, `GsmTabs.test.tsx` зелёные.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm/ && npm run build`
  - **Dependencies:** Task 8, 9, 10, 11 (всё перенесено)
  - **Files:** `components/GsmTabs.tsx`, `pages/gsm/GsmPage.tsx`,
    удаление `components/GsmPeriodView.tsx` + его теста
  - **Estimated scope:** S

- [x] Task 13: Регрессия + приёмка + отчёт
  - **Description:** Полный pytest и vitest, `npm run build`. Приёмка
    генерации на **копии** `plita.db`; UI-приёмка на рабочей: чеклист из
    Success Criteria спеки (статусы 4 машин, раскрытие, журнал транзакций,
    bulk-генерация 2 машин, экспорт zip 2 машин, деталь красного дня
    после перезагрузки, индикатор хвостов). Отчёт в
    `ai_docs/develop/reports/2026-08-24-gsm-fleet-overview-ux.md`.
  - **Acceptance:** все пункты Success Criteria спеки отмечены; отчёт.
  - **Verification:** команды спеки; ручной чеклист
  - **Dependencies:** Task 12
  - **Files:** `ai_docs/develop/reports/...`
  - **Estimated scope:** S

## Dependency Graph

```
Backend:
  T1 (repo) ──→ T2 (transactions list) ──→ Checkpoint A
           ──→ T3 (overview)          ──↗
           ──→ T4 (waybills all)      ──↗
  T5 (bulk generate) ───────────────────↗
  T6 (warning details) ─────────────────↗

Frontend:
  Checkpoint A → T7 (types/api/hooks) ──→ T8 (overview table) ──→ T9 (journal+dialog)
                                      ──→ T10 (tx journal)    ──→ T11 (bulk bar) ← T8,T9
                                                              ──→ T12 (page switch) ← T8..T11
                                                              ──→ T13 (acceptance) ← T12
```

**Параллелизуемо:** T2/T3/T4 (после T1), T5/T6 — независимы, но все трогают
`endpoints/gsm.py` и `schemas/gsm.py` → при последовательной работе одного
агента конфликтов нет; при параллели — мержить по очереди. Во frontend
T10 независима от T8/T9.

## Risks

| Риск | Митигация |
|---|---|
| Агрегатный SQL неверен (red_days по LIKE, open_before граница) | Фикстурные unit-тесты репозитория до сервисов (T1) |
| Граница `tx_last_date == wb_last_date` ложно даёт needs_generation | Явный тест границы в T3 (== → НЕ needs_generation) |
| Старые строковые `warnings_json` ломают парсер | T6: парсер принимает оба формата; регрессия без правки фикстур |
| Удаление `GsmPeriodView` роняет тесты страницы/табов | T12 последним; strip/drawer переиспользуются; чеклист переноса |
| `liters_diff` «врёт» на ручных ПЛ с fuel_issued override | Задокументировать семантику в отчёте; тест с ручным ПЛ в T3 |
| Bulk-генерация 4 машин — таймаут запроса | Солвер мс на машину; при росте — только тогда про async, не сейчас |
| UI-стили: в проекте inline-стили (CSSProperties), таблицы новые | Держать тот же подход, без новых зависимостей |

## Out of Scope

- Схема БД, `core/gsm/*`, экспорт бланка, контракты generate/patch/confirm.
- Календарная сетка, auto-preview (dry-run), пагинация, inline-правки в
  таблицах, массовый confirm.
- Async bulk-генерация (очередь/прогресс) — только если замер покажет проблему.
