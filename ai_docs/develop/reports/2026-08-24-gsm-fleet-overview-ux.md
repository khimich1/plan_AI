# Отчёт: ГСМ — обзор флота и журналы (UX-редизайн)

**Дата:** 2026-08-24  
**Спека:** [`../../specs/gsm-fleet-overview-ux.md`](../../specs/gsm-fleet-overview-ux.md)  
**План:** [`../plans/2026-08-24-gsm-fleet-overview-ux.md`](../plans/2026-08-24-gsm-fleet-overview-ux.md)  
**Статус:** T1–T13 выполнены. Коммитов нет.

## Что сделано

Главный экран `/gsm` — сводная таблица машин за период со статусами, раскрытием строки (журнал ПЛ + лента + drawer), массовой генерацией и zip-экспортом. Вкладка «Транзакции» — журнал с фильтрами и итогами с backend. Причина красных дней персистентна (`warning_details`). Индикатор незакрытых ПЛ до периода (`open_before`). Вкладка «Период» и `GsmPeriodView` удалены.

Слои: роутер → сервис → репозиторий. `core/gsm/*` и схема БД не менялись.

## Задачи

| Task | Содержание | Статус |
|------|------------|--------|
| T1 | Репозиторий: `fleet_overview`, списки транзакций/ПЛ без машины | ✅ |
| T2 | `GET /gsm/transactions` + конверт итогов | ✅ |
| T3 | `GsmOverviewService` + `GET /gsm/overview` | ✅ |
| T4 | `GET /gsm/waybills` — `vehicle_id` опционален | ✅ |
| T5 | `POST /gsm/waybills/generate-bulk` | ✅ |
| T6 | Персистентные `warning_details` | ✅ |
| T7 | Типы, API, хуки | ✅ |
| T8 | Сводная таблица обзора | ✅ |
| T9 | Журнал ПЛ + диалог генерации | ✅ |
| T10 | Журнал транзакций | ✅ |
| T11 | Bulk-бар: generate-bulk + zip-гейт | ✅ |
| T12 | Вкладка «Обзор», удаление «Период» | ✅ |
| T13 | Регрессия, приёмка, отчёт | ✅ |

## Success Criteria спеки

| # | Критерий | Доказательство | Вердикт |
|---|----------|----------------|---------|
| 1 | `/gsm` открывается на «Обзоре»: статусы, Σ литров, Σ км, красные дни, бейдж Δ по активным машинам за текущий месяц | `GsmPage`/`GsmTabs` дефолт `overview`; vitest страницы; live `GET /overview` за август 2026: 4 машины (Palisade `has_red_days`, Monjaro `ready`, Tugella `no_data`, Tugella `needs_generation`) | ✅ |
| 2 | Раскрытие строки: журнал ПЛ с км и лентой; drawer как раньше | `VehicleWaybillJournal` + существующие strip/drawer; vitest журнала | ✅ (компонентные тесты; клик в браузере — см. отклонения) |
| 3 | Журнал транзакций: фильтры, итоги с backend, импорт на месте | `TransactionsJournalView` + вкладка; live: 12 строк, Σ 361.35 л / 41416.01 ₽ | ✅ |
| 4 | Чекбоксы → generate-bulk с per-vehicle отчётом; zip один на чистые | vitest bulk-бара; generate-bulk на копии БД: машина 4 — 3 дня, машина 3 — 0 дней (нет транзакций) | ✅ API/тесты; live zip не качался (мутирует статусы) |
| 5 | `GET /waybills` без `vehicle_id` — все машины | pytest `test_gsm_waybills_list_all`; live август: 9 ПЛ, vehicle_id ∈ {1, 2} | ✅ |
| 6 | Красный день после перезагрузки показывает `detail` | На **копии** после `generate(force)` у Palisade 2026-08-04: `warning_details=[{code: manual_intervention, detail: "anchor day 2026-08-04 left corridor"}]`; повторное чтение совпало. На рабочей БД запись ещё строковая `["manual_intervention"]` → `details=[]` (контракт старого формата) | ✅ на копии / новых генерациях; ⚠️ live август — старый JSON |
| 7 | Индикатор хвостов при Σ `open_before` > 0 | vitest баннера; live: Monjaro 9, Tugella×2 по 1 | ✅ |
| 8 | Существующие GSM-тесты зелёные; новые сьюиты статусов/фильтров/bulk/двухформатности warnings | `venv/bin/pytest tests/test_gsm_*.py -q` → **214 passed** | ✅ |

## Прогоны

```
venv/bin/pytest tests/test_gsm_*.py -q     → 214 passed
venv/bin/pytest tests/ -q                  → 11 failed, 2240 passed, 8 skipped
cd frontend && npm test -- --run src/features/gsm/  → 19 files, 93 passed
cd frontend && npm test -- --run           → 78 files, 448 passed
cd frontend && npm run build               → tsc + vite OK
```

Полный pytest: те же 11 падений, что на Checkpoint A, **не из GSM**:

- `test_admin_service.py::test_reset_kp_only_keeps_completed_plates_and_rests`
- commercial flow (FBS/march/step/web identity)
- `test_kp_plates_resolve.py::test_order_data_from_completed_plates_omits_zero_unit_price`
- `test_ocr_gigachat_provider.py::test_require_gigachat_client_missing_credentials`
- `test_plate_audit.py::test_audit_log_records_completion_and_rejection`

## Приёмка генерации (только копия)

`cp plita.db /tmp/plita_accept.db` — рабочая `plita.db` генерацией не писалась.

- Сводка августа совпадает с live API.
- `generate_bulk([4, 3])`: id=4 ок (3 дня), id=3 ок (0 дней, нет транзакций).
- `generate(vehicle_id=1, force=True)`: красный день 2026-08-04 получил `warning_details.detail`; повторный `list_waybills` — тот же detail.

## Отклонения

1. **Полный pytest не зелёный** — 11 падений вне модуля ГСМ, как на Checkpoint A. Не чинились (вне скоупа).
2. **`liters_diff` Palisade = 2.0** при `wb_count=7` — поле считает backend (`tx_liters − wb_fuel_issued`); расхождение из ручных/красных дней, UI только красит. При `wb_count=0` бейдж скрыт (Tugella #4).
3. **Live `warning_details` пусты** у августовского красного дня Palisade: в БД старый строковый `warnings_json`. После новой генерации (проверено на копии) объекты `{code, detail}` сохраняются.
4. **UI в браузере:** Chrome DevTools MCP в сессии нет; SPA на `http://localhost:5173/commercial-offer/gsm` требует логин и JS. Клики обзора/zip на рабочей БД не гонялись, чтобы не менять статусы экспорта. Поведение закрыто vitest + live GET.
5. **`PlateListEditor.tsx`:** одна правка `reduce<number>` — иначе `tsc -b` падал на выводе аккумулятора (`number | null`). К ГСМ не относится, без неё DoD `npm run build` красный.
6. **Open question спеки** (удаление `GsmPeriodView`): файл удалён в T12. `gsm_start_required` из bulk открывает `VehicleGenerateDialog` той машины (кнопка «Указать старт»).

## Границы, которые соблюдены

Нет миграций БД, нет правок `core/gsm/*`, нет редактирования транзакций, нет массового confirm, нет новых npm/pip зависимостей, нет пагинации, нет коммита/push.
