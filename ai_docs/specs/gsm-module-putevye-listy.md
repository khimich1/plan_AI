# Spec: Модуль «ГСМ: путевые листы» (роль accountant)

Дата: 2026-08-14. Статус: draft, на ревью.
Идея: [`ai_docs/ideas/gsm-module-putevye-listy.md`](../ideas/gsm-module-putevye-listy.md) (2026-08-14).
Связанные идеи: [`poezdki-zapravki-moyki-lenta.md`](../ideas/poezdki-zapravki-moyki-lenta.md) (парсер транзакций, ключ `(карта, дата)`), [`marshruty-cherez-azs.md`](../ideas/marshruty-cherez-azs.md) (`крюк_км`, `routes_via`).

## ASSUMPTIONS I'M MAKING

1. Роль `accountant` — новое значение `app_users.role` (колонка `role TEXT`, миграция не нужна). Доступ: `REQUIRE_ACCOUNTING = require_roles("admin", "accountant")` в `app/dependencies/auth.py`.
2. БД — та же `plita.db`; новые таблицы с префиксом `gsm_` добавляются в `core/kp_db_schema.py` стилем `CREATE TABLE IF NOT EXISTS`.
3. Бланк ПЛ: существующий `.xls` (ОКУД 0345001, 19 формул) конвертируется в `.xlsx`-шаблон через `soffice` один раз; заполнение — openpyxl (формулы шаблона не трогаем); финальный экспорт в `.xls` с пересчётом — `soffice --headless`. **Подтверждено Phase 0 (2026-08-14, PASS):** `ai_docs/develop/reports/2026-08-14-gsm-blank-phase0.md`. Норма расхода захардкожена в формуле `BS41` шаблона — при экспорте патчим её под норму машины/сезона.
4. Транзакции загружаются пачкой .xls-файлов (мульти-файл, файл = выгрузка по одной карте); парсер переносится из `scripts/build_gsm_trip_feed.py` (шапка на 3-й строке, отброс «Итоги:», сверка сумм с итогами файла).
5. Период генерации = min/max дат загруженных транзакций, а не календарный месяц.
6. Сезон норм (лето/зима) — ручной переключатель в настройках модуля (дата начала зимы), не автомат по календарю.
7. Новая АЗС: геокодинг через существующий кэш `ГСМ/geo_cache/addresses.json` (Nominatim), подбор маршрута по минимальному `крюк_км` (механизм `routes_via`), порог 13 км; крюк > порога → день в ручной режим с подсветкой.
8. Карты и водители не удаляются физически — архивируются (связь с историей транзакций).
9. Оригиналы в `ГСМ/` (ПЛ, транзакции, xlsx) — read-only источники для одноразового импорта; после импорта источник правды — БД.

## Objective

Бухгалтер (роль `accountant`) загружает выгрузку транзакций по топливным картам и получает комплект путевых листов за период, в котором:

1. **Каждая транзакция — якорный день.** Заправка ИЛИ мойка = ПЛ на этот день, АЗС/мойка — точка маршрута дня (для налоговой отчётности).
2. **Списание сходится с литрами.** Полный баланс бака: остаток ведётся сквозной по дням в коридоре `[0 … объём бака]`, одометр — нарастающим итогом. Между якорями генератор добавляет дни-дожигания (будни, 150–250 км/день) из библиотеки маршрутов, чтобы сжечь заправленное. Количество дней — результат решения: `дней ≈ литры_периода / (норма × дневной_км)`.
3. **Маршруты правдоподобны.** Библиотека — 610 реальных направлений с типовыми АЗС (из `пул_поездок.xlsx`); для новой АЗС — подбор по минимальному крюку.

**Пользователь:** бухгалтер (`accountant`), администратор (`admin`). Остальные роли — 403.

**Критерий успеха MVP:** бухгалтер закрывает период по ГСМ за один сеанс без Excel; по итогам приёмки ручная переделка затрагивает ≤20% сгенерированных дней.

## Tech Stack

- Backend: Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`), xlrd (парсинг .xls — уже в venv), openpyxl (заполнение шаблона — уже в venv), LibreOffice headless (`/usr/bin/soffice` — системный, конвертация/пересчёт).
- Frontend: React 18, TypeScript, Vite, TanStack Query, feature-структура `frontend/src/features/gsm/` (по образцу `features/logistics/`).
- Геокодинг/OSRM для крюка новых АЗС — как в `scripts/build_gsm_routes_map.py`.
- Новых внешних pip-зависимостей нет.

## Commands

```bash
# Backend dev / tests
uvicorn app.main:app --reload
pytest tests/ -q
pytest tests/test_gsm_balance.py tests/test_gsm_generator.py -q

# Frontend dev / test / build
cd frontend && npm run dev
cd frontend && npm test -- --run
cd frontend && npm run build

# Phase 0 (блокер, первый шаг): round-trip бланка
python scripts/validate_gsm_blank_phase0.py --template "ГСМ/Geely Monjaro/2025 год/Апрель 2025/ПЛ 03.04.25.xls" --report ai_docs/develop/reports/2026-08-14-gsm-blank-phase0.md

# Одноразовый импорт истории (после Phase 0)
python scripts/import_gsm_history.py --db plita.db --gsm-dir "ГСМ"
```

## Project Structure

```
app/
  api/v1/endpoints/gsm.py                 → NEW: роутер модуля (регистрация в app/api/v1/router.py)
  dependencies/auth.py                    → +REQUIRE_ACCOUNTING
  services/
    gsm_registry_service.py               → NEW: CRUD справочников (машины/водители/карты/АЗС)
    gsm_transaction_service.py            → NEW: импорт xls-пачки, сверка итогов, дедупликация
    gsm_generation_service.py             → NEW: оркестрация генерации, пересчёт downstream
    gsm_export_service.py                 → NEW: заполнение бланка, zip-экспорт, вызов soffice
  repositories/
    gsm_repository.py                     → NEW: gsm_* таблицы
  schemas/gsm.py                          → NEW: Pydantic, ConfigDict(extra="forbid") для запросов

core/
  gsm/
    __init__.py
    models.py                             → DTO (frozen dataclasses): Transaction, Anchor, WaybillDay, TankBalance
    balance.py                            → чистая логика: коридор бака, перенос остатка, одометр
    generator.py                          → чистая логика: якоря + дожигание, будни, сезон, подбор маршрута
    blank.py                              → чистая логика: маппинг полей ПЛ → ячейки шаблона
  kp_db_schema.py                         → +gsm_* таблицы

frontend/src/
  features/gsm/
    api/gsmApi.ts                         → NEW
    types/gsm.ts                          → NEW
    hooks/useGsmQueries.ts                → NEW
    components/
      GsmPeriodView.tsx                   → NEW: период × машина, таймлайн остатка/одометра
      WaybillDayDrawer.tsx                → NEW: правка дня (маршрут, водитель, км) + downstream-пересчёт
      TransactionsImportDialog.tsx        → NEW: мульти-файл загрузка, отчёт сверки итогов
      CardsRegistryView.tsx               → NEW: привязка/архивация карт
      DriversRegistryView.tsx             → NEW: справочник водителей
      ManualWaybillDialog.tsx             → NEW: ручной конструктор ПЛ
  pages/gsm/GsmPage.tsx                   → NEW
  app/router/AppRouter.tsx                → +RequireRole ["admin","accountant"], path "gsm"
  app/layout/AppHeader.tsx                → +пункт «ГСМ» для accountant/admin

scripts/
  validate_gsm_blank_phase0.py            → NEW: round-trip бланка (блокер)
  import_gsm_history.py                   → NEW: Роману.xlsx + пул_поездок.xlsx + geo_cache → БД
```

## Code Style

```python
# core/gsm/balance.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass(frozen=True, slots=True)
class TankState:
    date: date
    fuel_start: float      # остаток при выезде, л
    fuel_issued: float     # выдано по заправочному листу, л
    fuel_end: float        # остаток при возвращении, л
    km: int
    odometer_start: int
    odometer_end: int


def burn_for_km(km: int, norm_per_100km: float) -> float:
    """Расход по норме за день: км × норма / 100."""
    return round(km * norm_per_100km / 100, 2)
```

Конвенции:
- Роутер → сервис → репозиторий; сервис не содержит SQL; `core/gsm/*` — без импортов `app.*` и без I/O.
- `from __future__ import annotations`; DTO — `@dataclass(frozen=True, slots=True)`.
- Солвер (`generator.py`, `balance.py`) — детерминированный: один вход → один выход (seed при tie-break).
- Вызов `soffice` — только из `gsm_export_service.py` через subprocess с таймаутом и изоляцией в tempdir.

## Testing Strategy

| Уровень | Что покрывает | Команда |
|---------|---------------|---------|
| Phase 0 (pytest + скрипт) | Round-trip бланка: конвертация .xls→.xlsx, fill, экспорт→.xls; формулы пересчитаны, значения на месте | `python scripts/validate_gsm_blank_phase0.py ...` |
| Unit (pytest) | `core/gsm/balance.py`: коридор `[0…бак]`, перенос остатка, переполнение → флаг | `pytest tests/test_gsm_balance.py -q` |
| Unit (pytest) | `core/gsm/generator.py`: якоря (заправка/мойка), дожигание, только будни, сезон норм, детерминизм | `pytest tests/test_gsm_generator.py -q` |
| Unit (pytest) | Парсер транзакций: шапка, «Итоги:», сверка сумм, дедупликация повторной загрузки | `pytest tests/test_gsm_transaction_service.py -q` |
| Integration (pytest) | `POST /gsm/transactions/import`, `POST /gsm/waybills/generate`, AuthZ 403 для manager/production | `pytest tests/test_gsm_api_integration.py -q` |
| Unit (vitest) | `GsmPeriodView`, `WaybillDayDrawer`, `TransactionsImportDialog` | `cd frontend && npm test -- --run` |
| Smoke (pytest) | Существующие тесты зелёные | `pytest tests/ -q` |

Coverage expectation: `core/gsm/*` ≥80% покрытие ветвлений (солвер — критичная логика).

## Boundaries

- **Always:** TDD — тест на логику солвера до реализации; `pytest` зелёный перед коммитом; минимальный diff; оригиналы `ГСМ/**` только на чтение.
- **Ask first:** новые таблицы `gsm_*` (миграция `kp_db_schema.py`); subprocess-вызов `soffice` из backend; сетевые вызовы Nominatim/OSRM из backend (геокод новых АЗС); изменение существующих API-контрактов.
- **Never:** не коммитить без явной просьбы; не удалять физически карты/водителей с историей (только архивация); не хранить транзакции вне БД после импорта; не ломать frozen bot paths (`bot_archived/`).

## Data Model

Новые таблицы в `core/kp_db_schema.py` (стиль `CREATE TABLE IF NOT EXISTS`):

```sql
gsm_vehicle (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,                    -- 'Geely Tugella 848'
  plate_number TEXT NOT NULL,            -- 'О 848 ХР 44'
  tank_volume_liters REAL NOT NULL,      -- 55
  norm_summer REAL NOT NULL,             -- 9.4
  norm_winter REAL NOT NULL,             -- 10.3
  primary_driver_id INTEGER REFERENCES gsm_driver(id),
  is_active INTEGER NOT NULL DEFAULT 1
);

gsm_driver (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  full_name TEXT NOT NULL,               -- 'Кулигин Никита Валерьевич'
  license_number TEXT NOT NULL,          -- '44 21 846315'
  license_issued_at TEXT,                -- '30.07.2015'
  personnel_number TEXT,                 -- табельный № '143'
  snils TEXT,                            -- для строки медосмотра
  is_active INTEGER NOT NULL DEFAULT 1
);

gsm_fuel_card (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  card_number TEXT NOT NULL UNIQUE,      -- '3005454268'
  vehicle_id INTEGER NOT NULL REFERENCES gsm_vehicle(id),
  assigned_at TEXT NOT NULL,
  archived_at TEXT                       -- NULL = активна; архивация вместо DELETE
);

gsm_station (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  address TEXT NOT NULL UNIQUE,          -- нормализованный 'Адрес ТО'
  brand TEXT,                            -- 'TATNEFT' / 'Газпромнефть' / 'ТНК'
  lat REAL, lon REAL,                    -- NULL до геокода
  geocode_source TEXT                    -- 'cache' | 'nominatim' | 'manual'
);

gsm_import_batch (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  filename TEXT NOT NULL,
  period_from TEXT, period_to TEXT,
  rows_total INTEGER, sum_liters REAL, sum_amount REAL,
  uploaded_by TEXT, uploaded_at TEXT NOT NULL
);

gsm_transaction (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  card_id INTEGER NOT NULL REFERENCES gsm_fuel_card(id),
  ts TEXT NOT NULL,                      -- ISO datetime транзакции
  service_type TEXT NOT NULL,            -- 'fuel' | 'wash' | 'other'
  fuel_grade TEXT,                       -- 'АИ-95'
  qty_liters REAL,                       -- NULL для мойки
  amount REAL NOT NULL,
  station_id INTEGER REFERENCES gsm_station(id),
  raw_address TEXT NOT NULL,
  batch_id INTEGER NOT NULL REFERENCES gsm_import_batch(id),
  UNIQUE(card_id, ts, qty_liters, amount)  -- дедупликация повторной загрузки
);

gsm_route (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  vehicle_id INTEGER NOT NULL REFERENCES gsm_vehicle(id),
  addr_a TEXT NOT NULL, addr_b TEXT NOT NULL,
  km INTEGER NOT NULL,                   -- утверждённое историей расстояние
  frequency INTEGER NOT NULL DEFAULT 1,
  typical_station_ids TEXT               -- JSON array of gsm_station.id
);

gsm_waybill (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  vehicle_id INTEGER NOT NULL REFERENCES gsm_vehicle(id),
  date TEXT NOT NULL,                    -- ISO YYYY-MM-DD
  driver_id INTEGER NOT NULL REFERENCES gsm_driver(id),
  status TEXT NOT NULL DEFAULT 'draft',  -- 'draft' | 'confirmed' | 'exported'
  source TEXT NOT NULL DEFAULT 'auto',   -- 'auto' | 'manual'
  odometer_start INTEGER, odometer_end INTEGER,
  fuel_start REAL, fuel_issued REAL, fuel_end REAL,
  route_json TEXT NOT NULL,              -- плечи: [{from,to,dep_time,arr_time,km,station_id?}]
  warnings_json TEXT,                    -- ['weekend_anchor','hook_above_threshold',...]
  UNIQUE(vehicle_id, date)
);

gsm_setting (
  key TEXT PRIMARY KEY,                  -- 'winter_start' | 'hook_threshold_km' | ...
  value TEXT NOT NULL
);
```

## Algorithm (core/gsm/generator.py + balance.py)

**Вход:** транзакции периода по машине, библиотека маршрутов машины, нормы + сезон, стартовые `остаток/одометр` (первый запуск — ввод бухгалтером; далее — из последнего подтверждённого ПЛ).

1. **Якоря:** каждая транзакция → день-якорь. Несколько транзакций в день = один ПЛ, все точки в маршруте. Мойка = якорь без литров. Транзакция в выходной/праздничный = якорь с warning `weekend_anchor` (в истории: выходных заправок 1/509, ПЛ на гос. праздники 0/1852 — праздники РФ считаются нерабочими наравне с выходными).
2. **Маршрут якоря:** маршрут библиотеки, где станция типовая; новая станция → минимальный `крюк_км`; крюк > `hook_threshold_km` (default 13) → warning `hook_above_threshold`, день помечается для ручного режима.
3. **Дожигание:** между якорями добавляются будни с маршрутами по убыванию частоты, пока `остаток` не вернётся в целевой коридор к следующей заправке: перед заправкой `Q` л остаток должен быть `≤ бак − Q` (иначе переполнение → дожигание усиливается задним числом).
4. **Баланс дня:** `fuel_end = fuel_start + fuel_issued − burn_for_km(km, норма_сезона)`; инвариант `0 ≤ fuel_end ≤ бак`; `odometer_end = odometer_start + km`.
5. **Выход:** черновики `gsm_waybill` + warnings. Правка дня через API → день фиксируется, downstream-дни пересчитываются.

## API

| Метод | Путь | Назначение |
|-------|------|-----------|
| POST | `/gsm/transactions/import` | Мульти-файл загрузка xls; ответ: отчёт по файлам (строк, литры, суммы vs итоги) |
| GET | `/gsm/transactions?from&to&vehicle_id` | Журнал транзакций |
| GET/POST/PATCH | `/gsm/vehicles`, `/gsm/drivers`, `/gsm/cards` | Справочники (PATCH cards — привязка/архивация) |
| POST | `/gsm/waybills/generate` | `{vehicle_id?, period_from, period_to}` → черновики + warnings |
| GET | `/gsm/waybills?vehicle_id&from&to` | Дни с таймлайном баланса/одометра |
| PATCH | `/gsm/waybills/{id}` | Правка маршрута/водителя/км → downstream-пересчёт |
| POST | `/gsm/waybills` | Ручной конструктор ПЛ |
| POST | `/gsm/waybills/export` | `{vehicle_ids, from, to}` → zip с «ПЛ DD.MM.YY.xls» |

AuthZ везде: `REQUIRE_ACCOUNTING`. Ошибки: `400` невалидный период, `409` генерация поверх `confirmed` без флага force, `422` переполнение бака неразрешимо (недостаточно будней между якорями).

## Success Criteria

1. **SC-0 (Phase 0, блокер):** round-trip бланка на реальном ПЛ: после `.xls→.xlsx→fill→.xls` формулы пересчитаны (остаток/расход сходятся с оригиналом ±0,01 л), вёрстка не сломана. Если нет — стоп, пересмотр экспорта (пересборка бланка).
2. **SC-1 (Импорт):** 9 файлов `ГСМ/транзакции/` → 509 транзакций (450 заправок + 59 прочих), суммы по каждому файлу сходятся с его «Итоги:»; повторная загрузка не дублирует.
3. **SC-2 (Солвер):** на историческом периоде генератор даёт комплект, где каждая заправка — в дне с ПЛ и станция в маршруте дня; `Σ fuel_issued − Σ burn = fuel_end_последний − fuel_start_первый`; ежедневный остаток в `[0…бак]`.
4. **SC-3 (Правки):** смена маршрута/км дня пересчитывает остаток и одометр всех последующих дней периода; подтверждённый (`confirmed`) день не перезаписывается генерацией.
5. **SC-4 (Экспорт):** zip содержит «ПЛ DD.MM.YY.xls» на каждый день; файл открывается, формулы считают, оборотная сторона содержит плечи с АЗС.
6. **SC-5 (AuthZ):** `manager`/`production`/`logistics` получают 403 на все `/gsm/*`; пункт меню «ГСМ» виден только `accountant`/`admin`.
7. **SC-6 (Regression):** существующие `pytest` и `vitest` зелёные.

## Этапы

- **Phase 0 (блокер):** `validate_gsm_blank_phase0.py` — round-trip бланка.
- **Phase 1:** роль + справочники (машины/водители/карты/АЗС) + импорт транзакций + `import_gsm_history.py`.
- **Phase 2:** солвер (`balance.py`, `generator.py`) + `generate`/`waybills` API.
- **Phase 3:** экран проверки (`GsmPeriodView`, `WaybillDayDrawer`) + ручной конструктор.
- **Phase 4:** экспорт zip + приёмка с бухгалтером на реальном периоде.

## Open Questions

1. Порог крюка 13 км — пересмотреть экспертно после первого отработанного периода (калибровка 2026-08-12 показала не-бимодальное распределение).
2. Период генерации = период выгрузки — подтвердить у бухгалтера на первом запуске.
3. Регенерация периода: перезапись `draft`, сохранение `confirmed` (предлагаемый дефолт) — или полное версионирование комплектов?
4. «Виртуальный клиент» для дальних новых АЗС (поиск правдоподобного адресата по тематике ЖБИ) — фаза 2, кандидат на отдельную идею после MVP.
