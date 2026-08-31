# Implementation Plan: Модуль «ГСМ: путевые листы» (роль accountant)

## Overview

Модуль превращает выгрузку транзакций по топливным картам в комплект путевых
листов за период: каждая транзакция (заправка/мойка) — якорный день с АЗС в
маршруте, между якорями генератор добавляет дни-дожигания из библиотеки
маршрутов, удерживая остаток бака в `[0 … объём бака]` и одометр нарастающим
итогом. Результат — zip бланков «ПЛ DD.MM.YY.xls» (ОКУД 0345001).

Спека: [`ai_docs/specs/gsm-module-putevye-listy.md`](../../specs/gsm-module-putevye-listy.md).
Идея: [`ai_docs/ideas/gsm-module-putevye-listy.md`](../../ideas/gsm-module-putevye-listy.md).

## Architecture Decisions

1. **Солвер — чистый core.** `core/gsm/balance.py` и `core/gsm/generator.py` —
   pure functions без I/O и импортов `app.*`: детерминированные, покрытые
   unit-тестами. Сервисный слой только читает/пишет БД и зовёт core.
2. **БД — источник правды.** Оригиналы `ГСМ/**` — read-only; история
   импортируется один раз (`import_gsm_history.py`), дальше всё живёт в
   `gsm_*` таблицах `plita.db`.
3. **Бланк = шаблон .xlsx + патч нормы.** Подтверждено Phase 0 (PASS):
   конвертация soffice, fill через openpyxl, экспорт в .xls с пересчётом.
   Норма расхода патчится в формуле `BS41` под машину/сезон при экспорте.
4. **Архивация вместо удаления.** Карты и водители с историей не удаляются
   (`archived_at` / `is_active=0`).
5. **Регенерация перезаписывает только `draft`.** Подтверждённые
   (`confirmed`/`exported`) ПЛ не трогаются без явного `force`.
6. **Правка дня фиксирует его и пересчитывает downstream.** Изменённый
   вручную день становится опорой; последующие draft-дни пересчитываются
   солвером от него.
7. **Новая АЗС → крюк, не синтез.** v1: геокод + минимальный `крюк_км` к
   библиотеке; крюк > `hook_threshold_km` (default 13) → warning + ручной
   режим. «Виртуальный клиент» — фаза 2, отдельной идеей.

## Task List

### Phase 0: Validation Gate (блокер)

- [x] Task 0: Round-trip бланка ПЛ ✅
  - **Description:** Скрипт `validate_gsm_blank_phase0.py`: .xls→.xlsx→fill→.xls, сценарий A (passthrough, 180 ячеек) + сценарий B (модификация км плеча, цепочка из 4 формул).
  - **Acceptance criteria:**
    - [x] Вердикт PASS: значения сохраняются, формулы пересчитываются.
    - [x] Отчёт: `ai_docs/develop/reports/2026-08-14-gsm-blank-phase0.md`.
  - **Verification:** `python scripts/validate_gsm_blank_phase0.py --template "ГСМ/Geely Monjaro/2025 год/Апрель 2025/ПЛ 03.04.25.xls"` → exit 0.
  - **Dependencies:** None
  - **Files:** `scripts/validate_gsm_blank_phase0.py`, отчёт
  - **Estimated scope:** M (фактически: выполнена, найдено 2 quirk'а LibreOffice — задокументированы)

### Checkpoint: Phase 0 Gate
- [x] PASS → продолжаем. Норма в `BS41` патчится при экспорте (Task 14).

---

### Phase 1: Foundation (справочники + импорт)

- [x] Task 1: Схема `gsm_*` + GsmRepository ✅
  - **Description:** 9 таблиц из спеки (`gsm_vehicle`, `gsm_driver`, `gsm_fuel_card`, `gsm_station`, `gsm_import_batch`, `gsm_transaction`, `gsm_route`, `gsm_waybill`, `gsm_setting`) в `core/kp_db_schema.py` + репозиторий с CRUD по каждой.
  - **Acceptance criteria:**
    - [ ] Таблицы создаются при старте (IF NOT EXISTS), индексы на `(card_id, ts)`, `(vehicle_id, date)`.
    - [ ] `UNIQUE(card_id, ts, qty_liters, amount)` на `gsm_transaction` — дедупликация.
    - [ ] Repository: CRUD справочников, `archive_card()`, `list_transactions(vehicle, period)`, `upsert_waybill()`.
  - **Verification:** `pytest tests/test_gsm_repository.py -q` зелёный.
  - **Dependencies:** Task 0
  - **Files:** `core/kp_db_schema.py`, `app/repositories/gsm_repository.py` (новый), `tests/test_gsm_repository.py` (новый)
  - **Estimated scope:** M

- [x] Task 2: Роль `accountant` + AuthZ ✅
  - **Description:** `DEFAULT_ACCOUNTANT_ROLE`, `REQUIRE_ACCOUNTING` в `app/dependencies/auth.py`; роль доступна в admin-UI создания пользователя.
  - **Acceptance criteria:**
    - [ ] `require_roles("admin", "accountant")` работает; прочие роли → 403.
    - [ ] Админ может создать пользователя с ролью `accountant` через существующий admin endpoint/UI.
  - **Verification:** `pytest tests/test_gsm_auth.py -q` зелёный (403/200 матрица ролей).
  - **Dependencies:** None (параллельно Task 1)
  - **Files:** `app/core/constants.py`, `app/dependencies/auth.py`, `tests/test_gsm_auth.py` (новый)
  - **Estimated scope:** S

- [x] Task 3: Импорт транзакций (парсер + сервис + endpoint) ✅
  - **Description:** `core/gsm/transactions.py` — pure-парсер .xls выгрузки (шапка на 3-й строке, отброс «Итоги:», классификация fuel/wash/other); `gsm_transaction_service` — мульти-файл импорт, сверка сумм с итогами, дедупликация, мэтчинг станции по `raw_address`; `POST /gsm/transactions/import`.
  - **Acceptance criteria:**
    - [ ] 9 файлов `ГСМ/транзакции/` → 509 транзакций; суммы по файлу сходятся с «Итоги:» (расхождение → warning в отчёте импорта, не стоп).
    - [ ] Повторная загрузка того же файла → 0 новых строк (дедупликация).
    - [ ] Неизвестная карта → транзакция принимается с флагом `unmatched_card` (карта заводится позже).
    - [ ] Ответ endpoint: отчёт по каждому файлу (строк, литров, сумма vs итог).
  - **Verification:** `pytest tests/test_gsm_transaction_import.py -q` зелёный (фикстура из реального xls).
  - **Dependencies:** Task 1, Task 2
  - **Files:** `core/gsm/__init__.py` (новый), `core/gsm/transactions.py` (новый), `app/services/gsm_transaction_service.py` (новый), `app/api/v1/endpoints/gsm.py` (новый, +`router.py`), `app/schemas/gsm.py` (новый), `tests/test_gsm_transaction_import.py` (новый)
  - **Estimated scope:** M

- [x] Task 4: `import_gsm_history.py` — одноразовый импорт ✅
  - **Description:** `Роману.xlsx` → машины/карты/нормы; `пул_поездок.xlsx` (лист `маршруты`) → `gsm_route` + типовые АЗС; `geo_cache/stations.geojson` → координаты станций; водители — авто-сбор из ПЛ (ФИО, удостоверение+дата, СНИЛС/табельный где заполнены; нормализация имён — «Cкрябин» лат./кирил. = один человек). **Крайний ПЛ каждой машины** → `gsm_waybill` со `status='confirmed', source='imported'` — стартовые одометр/остаток для первой генерации (решение 2026-08-14: полная история не импортируется).
  - **Acceptance criteria:**
    - [ ] 4 машины, 6 карт, ровно 8 водителей (после нормализации имён), 610 маршрутов, станции с координатами импортированы.
    - [ ] По каждой машине ровно 1 confirmed-ПЛ (крайняя дата) с одометром и остатком из файла.
    - [ ] Скрипт идемпотентен (повторный запуск не дублирует).
    - [ ] Отчёт конфликтов для ручного решения: удостоверение Лоншаковой (2835052 vs 283502), недостающие СНИЛС/табельные (5 водителей).
  - **Verification:** `python scripts/import_gsm_history.py --db tmp/test.db --gsm-dir "ГСМ"` на копии + сверка счётчиков.
  - **Dependencies:** Task 1
  - **Files:** `scripts/import_gsm_history.py` (новый), `tests/test_import_gsm_history.py` (новый)
  - **Estimated scope:** M

- [x] Task 5: Registry API (справочники CRUD) ✅
  - **Description:** `gsm_registry_service` + endpoints: `GET/POST/PATCH /gsm/vehicles|drivers|cards|stations`; `PATCH /gsm/cards/{id}` — привязка к машине / архивация; `GET/PUT /gsm/settings` (сезон, порог крюка).
  - **Acceptance criteria:**
    - [ ] Карта архивируется, не удаляется; архивная карта не предлагается в UI, но транзакции видны.
    - [ ] Валидация: нормы > 0, бак > 0, `card_number` уникален.
    - [ ] Все endpoints под `REQUIRE_ACCOUNTING`.
  - **Verification:** `pytest tests/test_gsm_registry.py -q` зелёный.
  - **Dependencies:** Task 1, Task 2
  - **Files:** `app/services/gsm_registry_service.py` (новый), `app/api/v1/endpoints/gsm.py`, `app/schemas/gsm.py`, `tests/test_gsm_registry.py` (новый)
  - **Estimated scope:** M

- [x] Task 6: Frontend foundation + навигация ✅
  - **Description:** `features/gsm/`: api, types, hooks; `GsmPage` с табами («Период», «Транзакции», «Справочники»); роут `/gsm` под `RequireRole ["admin","accountant"]`; пункт «ГСМ» в `AppHeader`.
  - **Acceptance criteria:**
    - [ ] Пункт меню виден только accountant/admin.
    - [ ] `/gsm` открывается, табы переключаются, данные справочников грузятся.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm` зелёный; `npm run build` успешен.
  - **Dependencies:** Task 5
  - **Files:** `frontend/src/features/gsm/api/gsmApi.ts` (новый), `.../types/gsm.ts` (новый), `.../hooks/useGsmQueries.ts` (новый), `frontend/src/pages/gsm/GsmPage.tsx` (новый), `frontend/src/app/router/AppRouter.tsx`, `frontend/src/app/layout/AppHeader.tsx`
  - **Estimated scope:** M

- [x] Task 7: Справочники UI + импорт транзакций UI ✅
  - **Description:** `DriversRegistryView`, `CardsRegistryView` (привязка/архивация), `VehiclesCard` (нормы/бак), `TransactionsImportDialog` (мульти-файл, отчёт сверки итогов).
  - **Acceptance criteria:**
    - [x] Карта привязывается к машине и архивируется из UI без перезагрузки.
    - [x] Импорт показывает отчёт по файлам; расхождения итогов подсвечены.
    - [x] Ошибки валидации API отображаются человекочитаемо.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm` зелёный; `npm run build` успешен.
  - **Dependencies:** Task 6
  - **Files:** `.../components/DriversRegistryView.tsx` (новый), `.../components/CardsRegistryView.tsx` (новый), `.../components/VehiclesCard.tsx` (новый), `.../components/TransactionsImportDialog.tsx` (новый), тесты
  - **Estimated scope:** M

### Checkpoint: Foundation
- [x] GSM foundation pytest suites зелёные (полный `pytest tests/` — см. T15 report / residuals)
- [x] Импорт реальных 9 файлов в dev-БД: 509 транзакций, отчёт без неожиданных расхождений
- [x] `import_gsm_history.py` на копии `plita.db`: счётчики сошлись

---

### Phase 2: Solver (ядро)

- [x] Task 8: `core/gsm/models.py` + `balance.py` ✅
  - **Description:** DTO (`Transaction`, `Anchor`, `WaybillDay`, `TankState`, `RouteRef`) и чистая логика баланса: `burn_for_km`, цепочка дня (`fuel_end = fuel_start + issued − burn`), инвариант `0 ≤ fuel_end ≤ бак`, одометр.
  - **Acceptance criteria:**
    - [ ] Все функции pure; детерминизм (один вход → один выход).
    - [ ] Переполнение и отрицательный остаток детектируются явно (`BalanceViolation`).
    - [ ] Покрытие ветвлений ≥80%.
  - **Verification:** `pytest tests/test_gsm_balance.py -q` зелёный.
  - **Dependencies:** None (pure core, параллельно Phase 1)
  - **Files:** `core/gsm/models.py` (новый), `core/gsm/balance.py` (новый), `tests/test_gsm_balance.py` (новый)
  - **Estimated scope:** M

- [x] Task 9: `core/gsm/generator.py` — солвер ✅
  - **Description:** Якоря (заправка/мойка; выходной/праздничный → warning), дожигание между якорями (только будни без гос. праздников РФ, 150–250 км, маршруты по частоте), защита от переполнения задним числом (перед заправкой Q: остаток ≤ бак − Q), сезон норм по `winter_start`, подбор маршрута якоря (типовая АЗС → min крюк → warning `hook_above_threshold`).
  - **Acceptance criteria:**
    - [ ] На синтетическом периоде: Σ issued − Σ burn = Δостатка; каждый день инвариант бака.
    - [ ] Каждая транзакция — в дне с ПЛ; станция в `route_json` дня.
    - [ ] Неразрешимое переполнение → явный результат `unsolvable` с причиной (не молча).
    - [ ] Детерминизм: два прогона → идентичный комплект.
  - **Verification:** `pytest tests/test_gsm_generator.py -q` зелёный.
  - **Dependencies:** Task 8
  - **Files:** `core/gsm/generator.py` (новый), `tests/test_gsm_generator.py` (новый)
  - **Estimated scope:** L

- [x] Task 10: Generation API ✅
  - **Description:** `gsm_generation_service` (оркестрация: транзакции + маршруты + настройки → солвер → draft waybills) + `POST /gsm/waybills/generate`, `GET /gsm/waybills` (таймлайн баланса/одометра + warnings).
  - **Acceptance criteria:**
    - [ ] Generate создаёт draft-комплект; повторный generate перезаписывает draft, не трогает confirmed (409 без `force`).
    - [ ] Стартовые остаток/одометр: из последнего confirmed ПЛ машины, иначе из ввода бухгалтера (поле запроса).
    - [ ] `GET` отдаёт по дням: маршрут, км, водитель, fuel_start/issued/end, одометр, warnings.
  - **Verification:** `pytest tests/test_gsm_generation_api.py -q` зелёный.
  - **Dependencies:** Task 9, Task 3
  - **Files:** `app/services/gsm_generation_service.py` (новый), `app/api/v1/endpoints/gsm.py`, `app/schemas/gsm.py`, `tests/test_gsm_generation_api.py` (новый)
  - **Estimated scope:** M

- [x] Task 11: Правка дня + ручной ПЛ (downstream-пересчёт) ✅
  - **Description:** `PATCH /gsm/waybills/{id}` (маршрут/водитель/км → день фиксируется, downstream draft пересчитывается); `POST /gsm/waybills` (ручной конструктор с авторасчётом полей горючего); `POST /gsm/waybills/{id}/confirm`.
  - **Acceptance criteria:**
    - [ ] После правки км дня N остаток/одометр дней N+1… пересчитаны и сходятся.
    - [ ] Ручной ПЛ участвует в балансе периода наравне с auto.
    - [ ] Confirmed-дни не пересчитываются.
  - **Verification:** `pytest tests/test_gsm_waybill_edit.py -q` зелёный.
  - **Dependencies:** Task 10
  - **Files:** `app/services/gsm_generation_service.py`, `app/api/v1/endpoints/gsm.py`, `app/schemas/gsm.py`, `tests/test_gsm_waybill_edit.py` (новый)
  - **Estimated scope:** M

### Checkpoint: Solver
- [x] Автотесты солвера/API зелёные (`test_gsm_balance`, `test_gsm_generator`, `test_gsm_generation_api`, `test_gsm_waybill_edit`) — инварианты на синтетике/API
- [ ] Полный прогон генерации на историческом периоде vs эталонные ПЛ (литры/каждый день) — ручная/боевая приёмка
- [x] GSM pytest suites зелёные; полный `pytest tests/` — ⚠️ см. отчёт (sandbox/sqlite noise; plus_code)

---

### Phase 3: Review UI

- [x] Task 12: `GsmPeriodView` — экран проверки ✅
  - **Description:** Сетка «период × машина»: дни с маршрутом/км/водителем, таймлайн остатка бака и одометра (мини-график или числовая лента), бейджи warnings (выходной якорь, крюк > порога, неразрешимое переполнение), кнопка «Сгенерировать».
  - **Acceptance criteria:**
    - [x] Видно весь период по машине: каждый день — км, АЗС, остаток на вечер.
    - [x] Warnings визуально отличимы и кликабельны (подсказка причины).
    - [x] Дни с транзакциями помечены (якоря).
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm` зелёный.
  - **Dependencies:** Task 10, Task 6
  - **Files:** `.../components/GsmPeriodView.tsx` (новый), `.../components/VehiclePeriodStrip.tsx` (новый), `.../hooks/useGsmQueries.ts`, тесты
  - **Estimated scope:** L

- [x] Task 13: `WaybillDayDrawer` + `ManualWaybillDialog` ✅
  - **Description:** Drawer правки дня: маршрут из библиотеки машины (фильтр по АЗС), водитель из справочника, км; превью пересчёта downstream до сохранения. Ручной конструктор ПЛ с нуля.
  - **Acceptance criteria:**
    - [x] Смена маршрута показывает, как изменятся остаток/одометр последующих дней (превью до commit).
    - [x] Сохранение → PATCH → сетка обновляется (инвалидация запроса).
    - [x] Ручной ПЛ: выбор машины/даты/маршрута/водителя, поля горючего автозаполнены и редактируемы.
  - **Verification:** `cd frontend && npm test -- --run src/features/gsm` зелёный (50 passed); `npm run build` OK.
  - **Dependencies:** Task 11, Task 12
  - **Files:** `.../components/WaybillDayDrawer.tsx` (новый), `.../components/ManualWaybillDialog.tsx` (новый), `.../components/GsmPeriodView.tsx`, тесты
  - **Estimated scope:** M

### Checkpoint: Review UI
- [x] `npm run build` успешен; GSM vitest зелёный
- [ ] Ручной e2e в UI: импорт → генерация → правка дня → пересчёт (часть SC-7 / опытной эксплуатации)

---

### Phase 4: Export & Acceptance

- [x] Task 14: Экспорт бланков (zip) ✅
  - **Description:** `core/gsm/blank.py` — маппинг `WaybillDay` → ячейки шаблона (дата, водитель+реквизиты, плечи, км, выдано; патч нормы в `BS41` под машину/сезон); `gsm_export_service` — openpyxl fill → soffice → «ПЛ DD.MM.YY.xls» → zip; `POST /gsm/waybills/export`.
  - **Acceptance criteria:**
    - [x] Zip содержит по файлу на каждый confirmed/draft день периода; имя `ПЛ DD.MM.YY.xls`.
    - [x] В файле: формулы пересчитаны (остаток/расход/одометр сходятся с данными БД ±0,01), оборотная сторона содержит плечи с АЗС.
    - [x] Норма в формуле соответствует машине и сезону даты ПЛ.
    - [x] soffice вызывается с изолированным профилем и таймаутом; падение → 500 с понятной ошибкой.
  - **Verification:** `pytest tests/test_gsm_export.py -q` зелёный (включая round-trip сверку значений через xlrd).
  - **Dependencies:** Task 11, Task 0
  - **Files:** `core/gsm/blank.py` (новый), `app/services/gsm_export_service.py` (новый), `app/api/v1/endpoints/gsm.py`, `tests/test_gsm_export.py` (новый)
  - **Estimated scope:** M

- [x] Task 15: Приёмка + документация ✅ (docs); ⏳ accountant sign-off
  - **Description:** Прогон реального периода с бухгалтером; фиксация % ручных переделок (цель ≤20%); feature doc + implementation report; проверка `soffice` в docker-образе деплоя.
  - **Acceptance criteria:**
    - [ ] Бухгалтер закрыла период в модуле без Excel; переделано ≤20% дней. **→ pending user sign-off (SC-7)**
    - [x] `ai_docs/develop/features/gsm-module-putevye-listy.md` + `ai_docs/develop/reports/2026-08-14-gsm-module-implementation.md`
    - [x] GSM pytest + `npm run build` зелёные (полный `pytest tests/` — см. residuals в отчёте); LibreOffice **отсутствует** в Dockerfile — **задокументирован** шаг `apt install libreoffice-calc-nogui` (образ не менялся в T15).
  - **Verification:** docs готовы; чек-лист приёмки бухгалтером — **не подписан**.
  - **Dependencies:** Task 13, Task 14
  - **Files:** `ai_docs/develop/features/gsm-module-putevye-listy.md`, `ai_docs/develop/reports/2026-08-14-gsm-module-implementation.md`
  - **Estimated scope:** S

### Checkpoint: Complete
- [x] SC-0…SC-5 — выполнены по автотестам/Phase 0 (см. отчёт)
- [ ] SC-6 — GSM suites OK; полный regression pytest — перепроверить вне sandbox
- [ ] SC-7 — приёмка бухгалтером ≤20% переделок — **pending user sign-off**
- [x] `npm run build` успешен (оркестрация T13+)
- [ ] Готово к опытной эксплуатации после SC-7 + soffice в deploy-образе

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Солвер выдаёт неправдоподобные дни | High | Экран проверки + правка дня (Task 12–13); критерий приёмки ≤20% переделок (Task 15) |
| Неразрешимое переполнение бака (плотные заправки) | Medium | Явный `unsolvable` + ручной режим; дожигание задним числом (Task 9) |
| Норма захардкожена в `BS41` шаблона | Medium | Патч формулы при экспорте (Task 14), подтверждено Phase 0 |
| `soffice` отсутствует/иное поведение в деплой-образе | Medium | Проверка в Task 15; quirk'и задокументированы в Phase 0 отчёте |
| Качество библиотеки маршрутов (редкие направления без типовых АЗС) | Medium | Fallback: min крюк → warning → ручной режим; библиотека пополняется подтверждёнными днями |
| Дрейф «период выгрузки» vs «месяц» у бухгалтера | Low | Open question — подтвердить на первом запуске; генератор берёт min/max дат загрузки |
| Регрессия существующих модулей | Low | Только новые файлы + точечные правки (`router.py`, `AppHeader`, `auth.py`); полный прогон тестов на каждом checkpoint |

## Параллельные треки

- **Данные (пользователь):** НЕ блокируют разработку — реквизиты водителей авто-собираются из ПЛ в Task 4 (СНИЛС есть у 3/8, табельные нестабильны). Ручное дополнение — к первому боевому экспорту (Task 14–15): СНИЛС/табельные 5 водителей (через UI справочника), удостоверение Лоншаковой (сверить с документом), ответственный за подпись ПЛ (в шаблоне «Прохоров Д.Д.» — подтвердить). Стартовые одометр/остаток — из крайних ПЛ автоматически.
- **Код:** Task 2 ∥ Task 1; Task 8–9 (pure core) ∥ Phase 1; Task 12 может стартовать после Task 10, не дожидаясь Task 11.

## Open Questions

1. Порог крюка 13 км — пересмотреть после первого периода (калибровка 2026-08-12 не бимодальна).
2. Период генерации = период выгрузки — подтвердить у бухгалтера на первом запуске.
3. Регенерация: перезапись draft / сохранение confirmed (принятый дефолт) — если нужно версионирование комплектов, сказать до Task 10.
4. «Виртуальный клиент» для дальних новых АЗС — фаза 2, отдельная идея после MVP.
