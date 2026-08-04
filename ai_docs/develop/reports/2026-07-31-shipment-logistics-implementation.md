# Report: Логистика и отгрузка (MVP-1 — реестр рейсов и списание СГП)

**Date:** 2026-07-31  
**Spec:** [`ai_docs/specs/shipment-logistics.md`](../../specs/shipment-logistics.md)  
**Plan:** [`ai_docs/develop/plans/2026-07-31-shipment-logistics.md`](../plans/2026-07-31-shipment-logistics.md)  
**Status:** ✅ Implemented (MVP-1), gates PASS — готово к пилоту

## Summary

Раздел «Логистика»: рейс собирается propose→confirm из реальных позиций СГП, выезд атомарно списывает `completed_plates` (audit `sgp_ship` с `shipment_id`), прогресс «отгружено X/M» в архиве, `KpStatus.DONE` только при полной отгрузке, справочник перевозчиков с дедупом/merge, каталог свай с автовесом, печатная форма XLSX, событие `shipment_completed` за feature-flag (off). Qty-инвариант расширен членом «отгружено» — plate_loss regression PASS.

## What was built (по задачам плана)

- **SHIP-000/001 Foundation:** enums (`ShipmentStatus`, `DeliveryType`, `ShipmentItemType`, `PlateStatus.SHIPPED`, `PlateTransitionReason.SGP_SHIP`); 5 таблиц (`shipments`, `shipment_orders`, `shipment_items`, `carriers`, `pile_catalog`) + `plate_status_log.shipment_id` + partial unique index на `carriers.name_normalized`; `audit_append(..., shipment_id=None)`; settings `VEHICLE_CLASS_LIMITS_KG` (t20=20000, t30plus=30000), `SHIPMENT_EVENTS_ENABLED` (default false), `EXCHANGE_EXPORT_DIR`; роль `logistics` + guard `REQUIRE_LOGISTICS` (admin+logistics) на всех `/logistics/*`.
- **SHIP-100/101 Справочники:** `scripts/import_pile_catalog.py` (upsert по mark, quirk `С137,5.40`, «-»→NULL); `scripts/import_carriers.py` (нормализация ОПФ/кавычек/ё + авто-дедуп + отчёт); `CarrierService` list/merge (`carrier_merge_conflict` 422).
- **SHIP-200…205 ShipmentService:** CRUD + `shipment_orders` (prefill ЯР из последнего рейса КП); availability как вычислимый резерв (`qty − Σ open-рейсов`); propose FIFO по `completed_date,id` с лимитом класса ТС и секцией `not_fit`, `propose_snapshot` для метрики; PUT items = полная замена состава с серверным пересчётом весов (плиты — формула, сваи — `pile_catalog.weight_kg × qty`, ручная правка веса); complete — одна транзакция: pre-flight (ЯР у всех orders, availability, ≥1 item) → списание → audit → `done` → DONE-check по каждому КП; cancel только из `in_work`; SGP guard `sgp_row_allocated` на unlink/relink зарезервированного.
- **SHIP-300/301 API:** `app/api/v1/endpoints/logistics.py` (все endpoint'ы из спеки + фильтры реестра дата/заказ/перевозчик/тип/«без УПД»/«внимание»); archive DTO += `shipped_progress {x, m}`.
- **SHIP-400…403 Frontend:** `features/logistics/` (registry + фильтры, ShipmentDrawer propose→confirm с редактором состава и предупреждением перегруза, CarriersView + merge-dialog, CarrierAutocomplete); route `/logistics`, `/logistics/carriers` + nav по роли; бейдж «отгружено X/M» в архиве рядом с «N/M на СГП».
- **SHIP-500:** `GET /logistics/shipments/{id}/sheet.xlsx` — «Лист отгрузки» (позиции в `sort_order`, шапка рейса), `Content-Disposition: attachment; filename="shipment_{id}_sheet.xlsx"`.
- **SHIP-600/601:** `tests/test_shipment_qty_balance.py` (Σ(kp_plates)+Σ(completed)+Σ(shipped done) = const по complete/cancel/мульти-КП); полный gate (см. Verification).
- **SHIP-603:** `scripts/shipment_propose_hitrate.py` (метрика пилота), этот отчёт, чеклист пилота ниже.

## Deviations (зафиксированы, осознанные)

1. **`completed_plates` обнуляются, не удаляются.** При полном списании строки она остаётся с `qty=0`: done-рейсы ссылаются на неё (`shipment_items.completed_plate_id`, карточка рейса, событие 1С). Частичное списание — декремент той же строки (remainder in place), отдельная split-строка не создаётся. DONE-check и qty-инвариант работают по `SUM(qty)`, поэтому нулевые строки безопасны.
2. **`sort_order: None` → индекс в payload.** Клиент может не присылать `sort_order`; сервер подставляет позицию строки в запросе (`item.sort_order if … is not None else index`).
3. **List endpoint — limit/offset.** `GET /shipments` имеет `limit=200` (max 1000), `offset=0`; фронт пагинацию не использует (для пилота достаточно).
4. **DONE-guard `ordered_qty <= 0`.** КП без известного M (freeze не дал числа) никогда не уходит в «выполнено» автоматически — защита от ложного DONE.
5. **Cancel удаляет строку рейса.** `DELETE FROM shipments` + `ON DELETE CASCADE` (orders/items); audit не пишется, availability восстанавливается автоматически; повторный GET → 404. Отмена `done` → 422 `shipment_not_in_work` (редактирование прошлого — фаза 2).
6. **Propose — `vehicle_class` query-параметр**, не JSON-body (спека описывала body); фронт выровнен под фактический контракт (см. Contract alignment).

## Contract alignment round (frontend ← backend, источник истины — backend)

Найдено и исправлено в `frontend/src/features/logistics/` (только follow-through, без изменений backend):

- `GET /shipments`, `/carriers`, `/pile-catalog` возвращают конверт `{items, count}` — фронт ожидал голый массив (runtime-баг: `.map` по объекту). Исправлено: unwrap `.items` в `api/logisticsApi.ts`.
- `POST /propose` читает `vehicle_class` из query-string — фронт слал JSON-body (класс ТС молча игнорировался). Исправлено: `?vehicle_class=…`.
- Типы ответов `PUT /items` (возвращает карточку, не `ShipmentItem[]`), `POST /complete`/`/cancel` (`ShipmentMutationResponse {ok, shipment_id, status, message}`, не `ShipmentDetails`) — приведены к фактическим.
- `ProposedItem` приведён к backend-схеме (`available_qty`, без `mark`/`note`/`sort_order`); `ProposeResponse += vehicle_class?`; `Carrier += note?, merged_into_id?`; мок `ShipmentDrawer.test.tsx` обновлён под реальный wire-format.
- Проверено без правок: карточка `available_by_kp: [{kp_id, plates: [...]}]`, archive `shipped_progress {x, m}`, Content-Disposition sheet.xlsx, 422 detail `{code, message, details}` (парсится `parseApiErrorPayload`).

## Verification (Phase 6 evidence)

**Backend gate (SHIP-600/601):**

- `pytest tests/ -q` — **1187 passed, 6 failed, 8 skipped** (95s). Все 6 падений — pre-existing, доказано прогоном на чистом worktree HEAD (acb4e76, фича не коммитилась):
  - `test_admin_service::test_reset_kp_only_keeps_completed_plates_and_rests` — на HEAD падает с идентичной причиной (`assert 2 == 1` по `completed_plates` после reset).
  - `test_core_no_app_import` — на HEAD падает идентично; нарушители — pre-feature `core/kp_db_plates_common.py`, `core/kp_db_plates_completion.py`; новый `core/kp_db_shipments.py` в списке отсутствует.
  - `test_plate_audit::test_audit_log_records_completion_and_rejection` — на HEAD падает идентично (нет reason `completed`).
  - `test_ocr_gigachat_provider` + 2× `test_recognition_pipeline` — на HEAD-worktree (без корневого `.env`) **проходят**; в основном дереве падают из-за GIGACHAT-креденшелов в `.env` (live-вызовы: «DID NOT RAISE ValueError», «None is not None»). К фиче отношения не имеют.
  - Новых падений от логистики нет.
- Targeted: `pytest tests/test_shipment_service.py (32) tests/test_shipment_qty_balance.py (3) tests/test_logistics_api.py (9) tests/test_sgp_service.py (4) -q` — **48 passed**.
- Фича-тесты справочников: `test_carrier_import.py (17)`, `test_pile_catalog_import.py (9)` — зелёные (в составе полного прогона).
- `./.venv/bin/python scripts/run_plate_loss_regression.py` — **PASS** (баланс OK, orphan Σ=0; вердикт «asc не прошёл, desc OK» — штатный паттерн).

**Frontend gate (SHIP-602):** `npm test -- --run` — **28 files / 164 tests passed**; `npm run build` — PASS (tsc clean; warning о chunk-size — pre-existing).

**E2E API smoke (временная БД, TestClient):** **23/23 PASS** — реальный login (logistics+admin через `/auth/login`) → search КП → create (delivery, 1 КП) → propose (`t20`, FIFO, `propose_snapshot` сохранён) → PUT items (plate 3 из 5 + free «С60.30»×14 → автовес **19 320 кг**, spec SC-7) → PATCH полей → sheet.xlsx (200, content-type, точный filename) → complete → проверки: `completed_plates` 5→2, audit `sgp_ship` (qty=3, shipment_id, actor, on_sgp→shipped), КП не DONE («На СГП»), archive `shipped_progress {x:3, m:5}`, повторный complete → 422 `shipment_not_in_work` с `{code, message, details}`, событие `shipment_completed` записано в `EXCHANGE_EXPORT_DIR` (flag=1) → cancel второго рейса: строка удалена, availability восстановлена (=2), `completed_plates` не тронуты, audit не добавлен. Скрипт: `/tmp/shipment_e2e_smoke.py` (throwaway, не коммитится).

**E2E импорта свай (реальный `import_pile_catalog.py` → temp db):** **3/3 PASS** — синтетический прайс (лист «Вес и объем», 44 марки с quirk `С137,5.40`) импортирован скриптом; `GET /logistics/pile-catalog?q=С60` возвращает импортированный вес; free-строка «С60.30»×14 через API → **19 320 кг** из импортированного каталога. Скрипт: `/tmp/pile_import_e2e.py`. Файл реального прайса в репозитории отсутствует (ПДн/цены) — пилотный запуск: `--xlsx` с путём к файлу (см. Pilot checklist).

**Hit-rate script:** синтетическая проверка бакетов (match/edited/no_snapshot/invalid, in_work исключён) — OK; на копии реальной `plita.db` — 0 done-рейсов (пилот впереди), корректный вывод «н/д».

## Pilot checklist (передача логисту)

Gate-критерии из спеки (Success Criteria «Пилот»): **hit-rate ≥ 50%**, **расхождения склада = 0 или задокументированы**, **buy-in логиста**.

1. **Шаг 0 — инвентаризация (до старта):** сверка «в системе = на складе» по каждому КП: физический остаток плит против `completed_plates`. Расхождения устранить до первого рейса (без этого propose и availability врут).
2. **Наполнение справочников (однократно):**
   - `./.venv/bin/python scripts/import_pile_catalog.py --xlsx "Прайс на цельные сваи от 27.07.2026.xlsx"` (44 марки; повторный запуск безопасен — upsert);
   - `./.venv/bin/python scripts/import_carriers.py --xlsx "Реестр отгрузок от 12082021.xlsx"` (импорт с авто-дедупом; отчёт о схлопнутых дублях в stdout; near-дубли — merge в UI «Перевозчики → Слить с…»).
3. **Пользователь логиста:** создать через admin UI (или `/auth/register` админом) с ролью `logistics`; доступ к разделу есть только у `logistics`/`admin`, остальные роли → 403.
4. **Cut-over неделя:** один тип выдачи (доставка ИЛИ самовывоз — выбрать) ведётся только в системе; **Excel-реестр — read-only** (новые строки не вбиваем). Каждый рейс: create → «Предложить состав» → правка при необходимости → «Утвердить состав» → поля (ЯР обязателен к выезду) → «Выезд» в день фактического выезда.
5. **Метрики в конце недели:**
   - `./.venv/bin/python scripts/shipment_propose_hitrate.py --verbose` — доля рейсов без ручной правки propose (gate ≥ 50%; ниже — пересмотр алгоритма propose);
   - повторная сверка «в системе = на складе» — расхождения 0 или каждое задокументировано (причина);
   - лист подтверждения логиста (buy-in) + список критичных замечаний завести в трекер.
6. **После пилота:** решение о полном cut-over (оба типа выдачи) и о включении `SHIPMENT_EVENTS_ENABLED` при старте интеграции G (1С).

## Operational notes

- **Настройки (env):** `VEHICLE_CLASS_LIMITS_KG` (JSON, default `{"t20": 20000, "t30plus": 30000}`, net-вес груза; уточнить у логиста — P2); `SHIPMENT_EVENTS_ENABLED` (default `false`, включать только с интеграцией 1С — Ask first); `EXCHANGE_EXPORT_DIR` (папка обмена, куда пишется `shipment_completed_{id}_{ts}.json`).
- **Событие 1С:** пишется после COMMIT закрытия рейса; ошибка записи — лог, закрытие не откатывается; повторная выгрузка — ручная (Ask first).
- **Вес — ориентир**, не учётный (учётный вес для УПД — в 1С). Перегруз класса ТС — предупреждение, закрытие не блокирует.
- **ЯР-номер** обязателен на закрытии (422 `shipment_missing_ya_order`), не при создании; prefill из последнего рейса того же КП.
- **Admin reset:** partial reset (kp/plans) не трогает logistics-таблицы; `reset_full` чистит рейсы, но сохраняет `carriers` и `pile_catalog` (справочники).
- **История Excel (17k строк) не мигрируется** — остаётся read-only архивом снаружи системы.

## Out of scope (unchanged, MVP-2/позже)

- Договоры-заявки (генерация DOCX, номер `N/MM/YY`), полная карточка перевозчика (реквизиты, налоговый режим), справочники водителей/ТС, ПДн
- Возврат/откат отгрузки (редактирование done-рейса), факт-стоимость/НДС/маржа, склад свай как полноценный учёт
- Оптимизатор укладки, услуги техники (жёлтые строки), ЭДО/ЭТрН/ГИС ЭПД, реальная отправка события в 1С
