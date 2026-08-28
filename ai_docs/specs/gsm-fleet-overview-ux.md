# Spec: ГСМ — обзор флота и журналы (UX-редизайн управления)

Дата: 2026-08-24. Статус: draft, на ревью.
Идея: [`../ideas/gsm-fleet-overview-ux.md`](../ideas/gsm-fleet-overview-ux.md) (2026-08-24, direction confirmed).
Базовый модуль: [`gsm-module-putevye-listy.md`](gsm-module-putevye-listy.md) (реализован).

## ASSUMPTIONS I'M MAKING

1. **Схема БД не меняется.** Сводка — агрегирующие SQL к существующим `gsm_*`
   таблицам; новых таблиц/колонок/миграций нет.
2. **Роли без изменений:** `REQUIRE_ACCOUNTING` на все новые эндпоинты.
3. **Солвер (`core/gsm/*`), экспорт бланка, контракты generate/patch/confirm —
   не трогаем.** Срез — только чтение (новые GET) + один bulk-эндпоинт + UI.
4. **Массовая генерация** — новый `POST /gsm/waybills/generate-bulk`:
   последовательный вызов существующего `GsmGenerationService.generate()` по
   машинам в одном HTTP-запросе; ошибка одной машины (404/422) не откатывает
   соседей — отчёт по каждой. Ответ всегда 200, кроме глобальных ошибок
   (невалидный период → 400).
5. **Массовый экспорт** — существующий `POST /gsm/waybills/export`
   (`vehicle_ids: list`) без изменений; работа — вызов из UI с чекбоксами.
6. **Статус машины в сводке** — вычисляемое поле на backend (см. «Статусная
   модель»), в БД не хранится.
7. **Журнал транзакций read-only.** Редактирования/удаления транзакций нет.
8. **«Обзор» — дефолтная вкладка.** Вкладка «Период» уходит: её функции
   поглощаются раскрытием строки (лента + журнал + drawer) и диалогом
   генерации с override `fuel_start`/`odometer_start`.
9. **Пагинации нет.** Масштаб из БД (2026-08-24): 4 машины, пик 19 ПЛ/мес
   на машину, ~170 транзакций/мес на флот. Клиентская сортировка/фильтры.
10. **Раскрытие строки** — ленивая подгрузка ПЛ машины через существующий
    `GET /gsm/waybills?vehicle_id=` (кеш TanStack Query), не мега-запрос.
11. **км в ответах уже есть:** `WaybillOut.km` = сумма плеч `route_json`.
    Σ км сводки = Σ `(odometer_end - odometer_start)` — колоночно, без JSON.
12. **Бейдж расхождения литров скрыт при `wb_count = 0`** — строка
    «требуется генерация» не должна кричать «−450 л» до первой генерации.
13. **Массовый экспорт:** машины с красными днями (гейт `exportGate.ts`)
    исключаются из zip с причиной в строке; чистые выгружаются. Не
    «одна красная блокирует весь zip».
14. **Хвосты до периода видны (решение 2026-08-24):** в `/overview` каждая
    строка имеет `open_before` — число ПЛ в `draft`/`confirmed` до начала
    периода. Защита цепочки показаний: `_resolve_start` берёт старт из
    последнего `confirmed`/`exported` до периода
    (`get_last_confirmed_waybill`, `status IN ('confirmed','exported')`) и
    **молча перепрыгивает** незакрытые черновики хвоста — без индикатора
    показания теряются незаметно.
15. **Причина красного дня персистентна (расширение скоупа, согласовано
    2026-08-24):** `gsm_generation_service` при сохранении дня пишет в
    `warnings_json` объекты `{"code", "detail"}` для проблемных дней
    (детали уже есть в `result.problematic_days`; сейчас причина живёт
    только в ответе generate и теряется после перезагрузки). `core/gsm/*`
    не трогаем: перевод код → объект на app-слое. **Контракт ответа не
    ломается:** `WaybillOut.warnings` остаётся `list[str]` (коды, как
    раньше); добавляется опциональное `warning_details:
    list[{code, detail}]` — чисто аддитивно. `_parse_warnings_json`
    принимает оба формата хранения — старые строковые записи дают
    `warning_details = []`.

→ Поправьте сейчас, иначе иду с этим в план/задачи.

## Objective

Бухгалтер (`accountant`) закрывает месяц по ГСМ, не держа состояние флота
в голове:

1. **Обзор («что делать»).** Главный экран — сводная таблица машин за период
   (по умолчанию — текущий календарный месяц): статус машины, Σ литров по
   транзакциям, Σ км и дней по ПЛ, красные дни, бейдж расхождения литров,
   чекбоксы для массовых действий (генерация / экспорт zip).
2. **Журналы («что есть»).** Раскрытие строки машины — её журнал ПЛ
   (дата, маршрут, км, бак, одометр, статус, warnings) + лента дней;
   вкладка «Транзакции» — полный журнал с фильтрами и итогами.
3. **Управление из статуса.** Кнопка «Сгенерировать» живёт в строке машины,
   которой она нужна; слепая форма исчезает как точка входа.
4. **Диагностика красных дней.** Красный день показывает причину
   (`detail` от солвера, персистентно в `warnings_json`) и путь исправления
   (drawer правки км/маршрута → downstream-пересчёт → confirm). Экспортный
   гейт — страховка, а не рабочий процесс.
5. **Целостность цепочки между периодами.** Индикатор в шапке обзора:
   «до периода незакрыто N ПЛ по M машинам» (`open_before` по строкам) —
   незакрытые черновики хвоста не дают цепочке показаний перепрыгнуть
   месяц незаметно.

**Пользователь:** бухгалтер (`accountant`), администратор (`admin`).

**Критерий успеха MVP:** бухгалтер открывает `/gsm` и без единого клика
видит по всем 4 машинам: что загружено, что сгенерировано, где красные дни,
что готово к экспорту. Закрытие месяца (импорт → генерация → разбор красных
→ экспорт) проходит без переключения «машина × период» в форме.

### Статусная модель машины в периоде

Вычисляется на backend из агрегатов. Приоритет сверху вниз (первое
сработавшее — статус строки):

| # | Статус | Условие | Действие в строке |
|---|--------|---------|-------------------|
| 1 | `no_data` | нет транзакций И нет ПЛ в периоде | — (строка серая) |
| 2 | `needs_generation` | есть транзакции И (нет ПЛ ИЛИ `max(tx_date) > max(wb_date)`) | «Сгенерировать» |
| 3 | `has_red_days` | `red_days > 0` (`warnings_json` содержит `manual_intervention`) | раскрыть → drawer |
| 4 | `drafts_pending` | есть ПЛ в `draft` (красных нет) | «Экспорт» (черновики выгружаются напрямую) |
| 5 | `pending_export` | нет draft, есть `confirmed`, но не все `exported` | «Экспорт» |

Примечание (проверено в коде 2026-08-24): экспорт сам ставит `exported`
(`gsm_export_service.py`), а кнопки confirm в UI нет — `POST
/waybills/{id}/confirm` существует только в API. Реальный путь закрытия:
`draft → exported`; `confirmed` — опциональный промежуточный статус,
в этом срезе UI для него не добавляем.
| 6 | `ready` | есть ПЛ, все `exported`, цепочка покрывает транзакции | — (зелёный) |

Решение зафиксировано (2026-08-24): **«готово» = все ПЛ периода `exported`**.
Расхождение литров — **бейдж** в строке: зелёный `0.0 л` при
`|Σ_tx_liters − Σ_wb_fuel_issued| ≤ 0.01`, иначе красный `+Δ / −Δ л`.
Период по умолчанию — **текущий календарный месяц**.

## Tech Stack

- Backend: Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`). Новых
  зависимостей нет.
- Frontend: React 18, TypeScript, Vite, TanStack Query, feature-структура
  `frontend/src/features/gsm/`.

## Commands

```bash
# Backend dev / tests
uvicorn app.main:app --reload
venv/bin/pytest tests/test_gsm_*.py -q
venv/bin/pytest tests/ -q

# Frontend dev / test / build
cd frontend && npm run dev
cd frontend && npm test -- --run src/features/gsm/
cd frontend && npm run build
```

## Project Structure

```
app/
  api/v1/endpoints/gsm.py              → CHANGED: +GET /transactions,
                                          +GET /overview, +POST /waybills/generate-bulk;
                                          GET /waybills: vehicle_id опционален
  repositories/gsm_repository.py       → CHANGED: +list_transactions_all/по флоту,
                                          +fleet_overview (агрегат), list_waybills
                                          с vehicle_id=None
  services/
    gsm_transaction_service.py         → CHANGED: +list_transactions (фильтры)
    gsm_generation_service.py          → CHANGED: list_waybills без vehicle_id;
                                          +generate_bulk (цикл + per-vehicle отчёт);
                                          запись warnings_json: объекты
                                          {code, detail} для проблемных дней;
                                          _parse_warnings_json принимает оба формата
    gsm_overview_service.py            → NEW: сборка сводки + статусная модель
  schemas/gsm.py                       → CHANGED: +TransactionOut,
                                          +TransactionListResponse,
                                          +FleetOverviewRow/VehiclePeriodStatus,
                                          +WaybillBulkGenerateRequest/Result,
                                          +WaybillWarningDetail;
                                          WaybillOut += warning_details (optional)

core/                                  → НЕ ТРОГАЕМ (солвер, баланс, geo, blank)

frontend/src/
  features/gsm/
    api/gsmApi.ts                      → CHANGED: +listTransactions, +getOverview,
                                        +generateWaybillsBulk; listWaybills —
                                        vehicleId опционален
    types/gsm.ts                       → CHANGED: +GsmTransaction, FleetOverviewRow,
                                        VehiclePeriodStatus, BulkGenerateResult
    hooks/useGsmQueries.ts             → CHANGED: +useGsmOverviewQuery,
                                        +useGsmTransactionsQuery,
                                        +useBulkGenerateMutation
    lib/fleetStatus.ts                 → NEW: маппинг статуса → ярлык/тон/действие
    lib/waybillWarnings.ts             → CHANGED: чтение warning_details
                                          (detail в бейдже/drawer); warnings
                                          по-прежнему кодами — старый UI не задет
    components/
      FleetOverviewView.tsx            → NEW: период + сводная таблица + bulk-бар
      FleetOverviewTable.tsx           → NEW: строки с раскрытием (чекбоксы, бейджи)
      VehicleWaybillJournal.tsx        → NEW: раскрытая строка — таблица ПЛ
                                          (км, бак, одометр) + VehiclePeriodStrip
      VehicleGenerateDialog.tsx        → NEW: генерация по машине (период
                                          предзаполнен; fuel/odometer override; force)
      TransactionsJournalView.tsx      → NEW: заменяет заглушку во вкладке
                                          «Транзакции» (таблица + фильтры + итоги)
      GsmTabs.tsx                      → CHANGED: +«Обзор» (дефолт), −«Период»
  pages/gsm/GsmPage.tsx                → CHANGED: дефолт «Обзор»; GsmPeriodView
                                        и его форма удаляются со страницы
    (components/GsmPeriodView.tsx      → DELETED после переноса функций;
     VehiclePeriodStrip.tsx            → REUSED внутри VehicleWaybillJournal;
     WaybillDayDrawer.tsx              → REUSED без изменений;
     TransactionsImportDialog.tsx      → REUSED, кнопка остаётся во вкладке)

tests/
  test_gsm_overview_api.py             → NEW: сводка, статусы, бейдж-агрегаты, 403
  test_gsm_transactions_list_api.py    → NEW: GET /transactions фильтры, all-vehicles
  test_gsm_waybills_list_all.py        → NEW: GET /waybills без vehicle_id
  test_gsm_generate_bulk_api.py        → NEW: bulk-генерация, per-vehicle ошибки
  frontend: FleetOverviewView.test.tsx, TransactionsJournalView.test.tsx,
            fleetStatus.test.ts, gsmApi.test.ts (новые методы)
```

## Code Style

Backend — как в модуле (роутер → сервис → репозиторий; сервис без SQL;
`extra="forbid"` на запросах):

```python
# app/services/gsm_overview_service.py
class GsmOverviewService:
    """Fleet overview aggregates + per-vehicle period status (read-only)."""

    def overview(self, *, period_from: date, period_to: date) -> list[FleetOverviewRow]:
        rows = self._repo.fleet_overview(period_from=period_from, period_to=period_to)
        return [_to_row(r, status=_status_of(r)) for r in rows]


def _status_of(agg: dict[str, Any]) -> VehiclePeriodStatus:
    if agg["tx_count"] == 0 and agg["wb_count"] == 0:
        return "no_data"
    if agg["tx_count"] > 0 and (
        agg["wb_count"] == 0 or agg["tx_last_date"] > (agg["wb_last_date"] or "")
    ):
        return "needs_generation"
    if agg["red_days"] > 0:
        return "has_red_days"
    if agg["draft_count"] > 0:
        return "drafts_pending"
    if agg["exported_count"] < agg["wb_count"]:
        return "pending_export"
    return "ready"
```

Frontend — как в `features/gsm` (TanStack Query, типы рядом, форматеры в
`lib/`):

```tsx
// features/gsm/lib/fleetStatus.ts
export const fleetStatusMeta = (status: VehiclePeriodStatus) =>
  ({
    no_data: { label: "Нет данных", tone: "muted" },
    needs_generation: { label: "Требуется генерация", tone: "warning" },
    has_red_days: { label: "Есть красные дни", tone: "danger" },
    drafts_pending: { label: "Черновики", tone: "warning" },
    pending_export: { label: "Готово к экспорту", tone: "info" },
    ready: { label: "Выгружено", tone: "success" },
  })[status];
```

Конвенции:
- Даты в API — ISO (`YYYY-MM-DD`), сравнение дат строковое (SQLite `TEXT`).
- Все денежные/литровые суммы — `round(..., 2)` на backend; UI не считает Σ.
- Бейдж расхождения считает backend (`liters_diff`), UI только красит.
- Раскрытие строк — `useGsmWaybillsQuery({vehicleId, from, to})` с `enabled:
  expanded` (лениво, кеш на повторное раскрытие).

## API

Prefix `/api/v1/gsm`, auth `REQUIRE_ACCOUNTING`.

| Method | Path | Изменение |
|--------|------|-----------|
| GET | `/transactions?vehicle_id&from&to&service_type` | NEW. `vehicle_id`/`service_type` опциональны; ответ-конверт `TransactionListResponse {rows: list[TransactionOut], total_count, sum_liters, sum_amount}` — итоги считает backend по отфильтрованному набору |
| GET | `/overview?from&to` | NEW. `list[FleetOverviewRow]`: vehicle, tx_count/tx_liters/tx_amount/tx_last_date, wb_count/wb_km/wb_fuel_issued/wb_last_date, red_days, draft/confirmed/exported counts, fuel_end_last, liters_diff, open_before, status |
| POST | `/waybills/generate-bulk` | NEW. `{vehicle_ids, period_from, period_to, force}` → 200 `WaybillBulkGenerateResult{results: per-vehicle {vehicle_id, ok, result\|error{code,message}}}` |
| GET | `/waybills?vehicle_id&from&to` | CHANGED: `vehicle_id` опционален (без него — все машины, сортировка `vehicle_id, date`) |
| POST | `/waybills/export` | без изменений (уже `vehicle_ids: list`) |

Ошибки: `400` невалидный период; `403` чужая роль; `404` машина не найдена
(в bulk — per-vehicle error, не HTTP).

## Testing Strategy

| Уровень | Что покрывает | Команда |
|---------|---------------|---------|
| Unit (pytest) | Статусная модель `_status_of`: все 6 веток, приоритеты, граница `tx_last_date == wb_last_date` | `venv/bin/pytest tests/test_gsm_overview_api.py -q` |
| Integration (pytest) | `GET /overview`: агрегаты Σ литров/км, red_days, liters_diff, статус; 403 для manager | `test_gsm_overview_api.py` |
| Integration (pytest) | `GET /transactions`: фильтры vehicle_id/service_type/период, машина `null` у непривязанной карты | `test_gsm_transactions_list_api.py` |
| Integration (pytest) | `GET /waybills` без vehicle_id: все машины, сортировка; с vehicle_id — регрессия | `test_gsm_waybills_list_all.py` |
| Integration (pytest) | bulk: 2 машины (одна без маршрутов → per-vehicle error, вторая ок), ответ 200; force поверх confirmed | `test_gsm_generate_bulk_api.py` |
| Unit (vitest) | `fleetStatus.ts` маппинги; `FleetOverviewView` (чекбоксы, bulk-бар, раскрытие); `TransactionsJournalView` (итоги из ответа, фильтры) | `cd frontend && npm test -- --run src/features/gsm/` |
| Smoke | Существующие GSM-сьюиты зелёные | `venv/bin/pytest tests/test_gsm_*.py -q` |

Приёмка (ручная, на рабочей БД): открыть `/gsm` → дефолт «Обзор», текущий
месяц; по 4 машинам статусы соответствуют данным; раскрытие строки → журнал
с км; импорт → вкладка «Транзакции» показывает строки и итоги; выбрать 2
машины → «Сгенерировать выбранные» → отчёт; «Экспорт zip выбранных» качает
zip по 2 машинам.

## Boundaries

- **Always:** TDD на новые эндпоинты; роутер → сервис → репозиторий;
  `extra="forbid"`; Σ только на backend; даты ISO; регрессия
  `tests/test_gsm_*.py` и `npm test` зелёные.
- **Ask first:** изменение контрактов существующих эндпоинтов кроме
  опциональности `vehicle_id` и аддитивного `warning_details`; удаление
  `GsmPeriodView` до переноса всех функций; новые зависимости; пагинация.
- **Never:** схема БД; изменение `core/gsm/*`; редактирование транзакций;
  массовый confirm (только по одному через журнал); коммит без явной просьбы.

## Success Criteria

1. `/gsm` открывается на «Обзоре»: без кликов видны статус, Σ литров, Σ км,
   красные дни и бейдж расхождения по каждой активной машине за текущий месяц.
2. Раскрытие строки показывает журнал ПЛ машины с колонкой км и лентой;
   drawer правки работает как раньше (downstream-пересчёт не задет).
3. Вкладка «Транзакции» показывает журнал всех транзакций с фильтрами
   (машина/тип/период) и итогами Σ литров/Σ суммы; импорт остаётся там.
4. Чекбоксы → «Сгенерировать выбранные»: per-vehicle отчёт, ошибка одной
   машины не блокирует остальные; «Экспорт zip выбранных» — один zip.
5. `GET /gsm/waybills` без `vehicle_id` возвращает ПЛ всех машин.
6. Красный день после перезагрузки страницы показывает причину (`detail`),
   а не только код; drawer открывается из журнала кликом по строке.
7. При наличии незакрытых ПЛ до периода шапка обзора показывает индикатор
   с числом ПЛ и машин; при отсутствии — индикатора нет.
8. Все существующие тесты GSM зелёные (включая парсинг старых строковых
   `warnings_json`); новые сьюиты покрывают статусы, фильтры, bulk,
   двухформатность warnings.

## Open Questions

- Удаление `GsmPeriodView.tsx` — после переноса функций подтвердить,
  что override fuel/odometer в диалоге генерации покрывает первый период
  машины (`gsm_start_required` из bulk → индивидуальный диалог).
