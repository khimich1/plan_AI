# ГСМ: путевые листы

**Status**: ✅ Implemented (MVP code complete; accountant sign-off pending)  
**Date**: 2026-08-14  
**Orchestration**: `orch-2026-08-14-gsm-module`  
**Spec**: [`../../specs/gsm-module-putevye-listy.md`](../../specs/gsm-module-putevye-listy.md)  
**Plan**: [`../plans/2026-08-14-gsm-module.md`](../plans/2026-08-14-gsm-module.md)  
**Report**: [`../reports/2026-08-14-gsm-module-implementation.md`](../reports/2026-08-14-gsm-module-implementation.md)  
**Phase 0**: [`../reports/2026-08-14-gsm-blank-phase0.md`](../reports/2026-08-14-gsm-blank-phase0.md)

## Description

Бухгалтер загружает выгрузки транзакций по топливным картам (.xls) и получает комплект путевых листов (ОКУД 0345001) за период:

1. Каждая заправка/мойка — якорный день с АЗС/мойкой в маршруте.
2. Между якорями генератор добавляет дни-дожигания (будни), удерживая остаток бака в `[0 … объём]` и одометр нарастающим итогом.
3. Экспорт — zip «ПЛ DD.MM.YY.xls» (шаблон .xlsx + openpyxl + LibreOffice `soffice`).

## Roles

| Role | Access |
|------|--------|
| `accountant`, `admin` | Полный доступ к `/gsm` и API `/api/v1/gsm/*` (`REQUIRE_ACCOUNTING`) |
| `manager`, `production`, `logistics`, прочие | 403 на API; пункт меню «ГСМ» скрыт |

Роут UI: `/gsm` (`RequireRole ["admin","accountant"]`). Меню: `AppHeader` → «ГСМ».

## How to use

1. **Одноразово:** `python scripts/import_gsm_history.py --db plita.db --gsm-dir "ГСМ"` — машины, карты, водители, маршруты, станции, стартовый confirmed-ПЛ по машине.
2. Открыть **ГСМ** → вкладка **Транзакции** → **Импорт** (мульти-файл .xls). Проверить отчёт сверки с «Итоги:».
3. Вкладка **Период**: выбрать машину и даты → **Сгенерировать**.
4. Проверить дни (якоря, warnings, остаток/одометр). При необходимости — drawer правки или **Ручной ПЛ**.
5. **Экспорт zip** в той же форме: скачивает бланки. Красные дни (`manual_intervention` / `unsolvable`) блокируют кнопку; жёлтые warnings и повторный экспорт — confirm-диалог. Отдельный confirm дня в UI нет (`exported` закрывает цепочку бака/одометра).
6. **Справочники**: машины/нормы, водители, карты (привязка/архивация), настройки сезона (`winter_start`) и порога крюка.

Период генерации задаётся явно (`from`/`to`); продуктово обычно совпадает с min/max дат загруженных транзакций.

## Architecture (shipped)

| Layer | Path | Role |
|-------|------|------|
| Domain | `core/gsm/{models,balance,generator,geo,blank,transactions}.py` | Pure: баланс, солвер v2 (lookahead/round-trip/geo), маппинг бланка, парсер xls |
| Repository | `app/repositories/gsm_repository.py` | `gsm_*` таблицы |
| Services | `gsm_{registry,transaction,generation,export}_service.py` | Оркестрация + I/O |
| API | `app/api/v1/endpoints/gsm.py` | REST под `REQUIRE_ACCOUNTING` |
| UI | `frontend/src/features/gsm/`, `pages/gsm/GsmPage.tsx` | Период / импорт / справочники |
| Template | `core/gsm/templates/waybill_blank.xlsx` | Шаблон после Phase 0 |

БД: `plita.db`, таблицы `gsm_*` через `CREATE TABLE IF NOT EXISTS` в `core/kp_db_schema.py`.

## Key API endpoints

Prefix: `/api/v1/gsm`. Auth: `REQUIRE_ACCOUNTING`.

| Method | Path | Purpose |
|--------|------|---------|
| POST | `/transactions/import` | Мульти-файл .xls; отчёт сверки / дедуп |
| GET/POST/PATCH | `/vehicles`, `/drivers`, `/cards`, `/stations` | Справочники (карты — архивация) |
| GET | `/routes?vehicle_id=` | Библиотека маршрутов машины |
| GET/PUT | `/settings` | `winter_start`, `hook_threshold_km`, `max_daily_km` (default 700) |
| POST | `/waybills/generate` | 200: черновики + `problematic_days` / `manual_days` (`force` для confirmed). 422 только конфигурация (нет машины/маршрутов/водителя), не баланс |
| GET | `/waybills?vehicle_id&from&to` | Таймлайн дней |
| POST | `/waybills` | Ручной ПЛ |
| PATCH | `/waybills/{id}` | Правка → downstream-пересчёт draft |
| POST | `/waybills/{id}/confirm` | Статус `confirmed` |
| POST | `/waybills/export` | Zip «ПЛ DD.MM.YY.xls» |

Ошибки: `400` невалидный период / (желательно) битый xls; `403` чужая роль; `409` generate поверх confirmed без `force`; `422` только конфигурация / валидация (не нерешаемость бака); `500` сбой `soffice`.

## Lookahead и география (2026-08-15)

Генератор v2 (`orch-2026-08-15-gsm-geo-lookahead`):

1. **Round-trip:** день = 2 плеча (туда + обратно), дневной km = `2×` library km.
2. **Lookahead:** на плотных якорях без свободных будней выбирается минимальный достаточный маршрут, чтобы освободить бак под следующую заправку (`max_daily_km`, дефолт 700).
3. **География:** после отбора по баку — мягкая сортировка (станция на маршруте + направление к следующей АЗС, `angle_diff ≤ 90°`).
4. **Частичная генерация:** нерешаемый якорь → draft `manual_intervention`, период собирается целиком, ответ 200 с `problematic_days`. UI: красные дни, сводка «N дней, M требуют ручной доработки».

Data-миграции (не схема): `scripts/geocode_gsm_stations.py`, `scripts/link_route_stations.py`. Геокодинг только из `scripts/`, не из backend runtime.

**Фаза 2 (не в этом срезе):** ночёвки / `overnight_trip` / `return_leg` — для правдоподобия длинных дней, не для сходимости бака.

Приёмка Palisade май 2026 (v2, 2026-08-15): 13 дней, 2 manual (08.05, 21.05), без 422. [Acceptance](../reports/2026-08-15-gsm-geo-lookahead-acceptance.md).

## Солвер A+B (2026-08-17)

Генератор: бак важнее узкой группы АЗС + короткий дожиг ([spec](../../specs/gsm-solver-tank-first-short-burn.md)):

1. **Дожиг (B):** кандидаты — все маршруты с `2×km ≤ max_daily_km`, не сетка 150–250 км. Пока headroom не достигнут — max безопасный burn; в день попадания в коридор — минимальный достаточный km.
2. **Lookahead (A):** если в группе станции нет `2×km ≥ km_needed` — `_pick_min_sufficient` по всей библиотеке машины (уже в капе) + жёлтый `balance_route`. Группу не расширяем до проверки, что будни дожгут остаток.
3. **`manual_intervention`** остаётся предохранителем; на Palisade мае 2026 не срабатывает.

Приёмка Palisade 04.05–31.05 (копия БД, 28 л / 128327): 17 дней, `problematic_days == []`. 06.05 Ростов typical; 07.05 короткий дожиг 45 км; 20.05 Переславль из полной библиотеки. [Report](../reports/2026-08-17-gsm-solver-tank-first-short-burn.md).

## Frontend components

| Component | Role |
|-----------|------|
| `GsmPeriodView` / `VehiclePeriodStrip` | Период × машина, генерация, **Экспорт zip**, warnings |
| `WaybillDayDrawer` | Правка дня + превью downstream |
| `ManualWaybillDialog` | Ручной конструктор |
| `TransactionsImportDialog` | Мульти-файл импорт |
| `CardsRegistryView` / `DriversRegistryView` / `VehiclesCard` | Справочники |
| `GsmTabs` / `GsmRegistriesView` | Навигация вкладок |

## Deploy note (LibreOffice)

Экспорт ПЛ требует `/usr/bin/soffice` (LibreOffice headless). В корневом `Dockerfile` и `docker/backend/Dockerfile` пакет **не установлен** — см. implementation report (шаг установки).

## Known residuals

- Приёмка бухгалтером (≤20% ручных переделок) — **pending user sign-off**.
- [`ISS-002`](../issues/ISS-002-gsm-import-xls-parse-errors.md): `XLRDError` → 400 + allowlist .xls.
- `typical_station_ids`: заполнены скриптом `link_route_stations.py` (450/610; Palisade × станция id=1 — 116 маршрутов). Повторный прогон не затирает непустые.
- `openlocationcode` в `requirements.txt`, но может отсутствовать в venv → `test_plus_code_reference_uses_full_tail`.
- Журнал `GET /gsm/transactions` в UI пока не выведен (импорт есть).
- Confirm дня в drawer — не в UI (фаза 1.1; `exported` достаточен для следующего периода).

## How to test

```bash
# GSM backend suites
pytest tests/test_gsm_balance.py tests/test_gsm_generator.py \
  tests/test_gsm_repository.py tests/test_gsm_auth.py \
  tests/test_gsm_transaction_import.py tests/test_gsm_registry.py \
  tests/test_gsm_generation_api.py tests/test_gsm_waybill_edit.py \
  tests/test_gsm_export.py tests/test_import_gsm_history.py -q

# Frontend
cd frontend && npm test -- --run src/features/gsm && npm run build

# Phase 0 blank (при наличии шаблона и soffice)
python scripts/validate_gsm_blank_phase0.py \
  --template "ГСМ/Geely Monjaro/2025 год/Апрель 2025/ПЛ 03.04.25.xls"
```

## Related

- Idea: [`../../ideas/gsm-module-putevye-listy.md`](../../ideas/gsm-module-putevye-listy.md)
- Export UI: [`../../ideas/gsm-waybill-export-buttons.md`](../../ideas/gsm-waybill-export-buttons.md)
- Related ideas: trip feed / routes via AZS
- Issue: [`ISS-002`](../issues/ISS-002-gsm-import-xls-parse-errors.md)
