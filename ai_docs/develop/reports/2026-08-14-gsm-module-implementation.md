# Report: Модуль «ГСМ: путевые листы» — implementation

**Date:** 2026-08-14  
**Orchestration:** `orch-2026-08-14-gsm-module`  
**Status:** ✅ Code complete (T0–T14); T15 docs done; **accountant acceptance pending**  
**Commits:** нет (по brief / без явной просьбы)

**Spec:** [`../../specs/gsm-module-putevye-listy.md`](../../specs/gsm-module-putevye-listy.md)  
**Plan:** [`../plans/2026-08-14-gsm-module.md`](../plans/2026-08-14-gsm-module.md)  
**Feature:** [`../features/gsm-module-putevye-listy.md`](../features/gsm-module-putevye-listy.md)  
**Phase 0:** [`2026-08-14-gsm-blank-phase0.md`](2026-08-14-gsm-blank-phase0.md)

## Summary

Реализован MVP модуля ГСМ для роли `accountant`: импорт транзакций, справочники, солвер баланса/дожигания, UI проверки периода с правкой дня, экспорт zip бланков ПЛ через LibreOffice. Phase 0 blank round-trip — PASS. Ручная приёмка бухгалтером (цель ≤20% переделок дней) **не подписана** — отдельный user sign-off.

## Completed tasks (T0–T15)

| Task | Name | Status |
|------|------|--------|
| T0 | Round-trip бланка (Phase 0) | ✅ PASS + отчёт |
| T1 | Схема `gsm_*` + `GsmRepository` | ✅ |
| T2 | Роль `accountant` + AuthZ | ✅ |
| T3 | Импорт транзакций (парсер + API) | ✅ |
| T4 | `import_gsm_history.py` | ✅ |
| T5 | Registry API | ✅ |
| T6 | Frontend foundation + навигация | ✅ |
| T7 | Справочники UI + импорт UI | ✅ |
| T8 | `core/gsm` models + balance | ✅ |
| T9 | `generator.py` солвер | ✅ |
| T10 | Generation API | ✅ |
| T11 | Правка дня + ручной ПЛ | ✅ |
| T12 | `GsmPeriodView` | ✅ |
| T13 | `WaybillDayDrawer` + `ManualWaybillDialog` | ✅ |
| T14 | Экспорт zip бланков | ✅ |
| T15 | Приёмка + documentation | ✅ docs; ⏳ accountant sign-off |

## Success criteria checklist

Спека нумерует **SC-0…SC-6**; план/приёмка трактуют **SC-7** как ручную приёмку (≤20% переделок).

| ID | Criterion | Evidence | Verdict |
|----|-----------|----------|---------|
| SC-0 | Blank round-trip .xls→xlsx→fill→xls | [`2026-08-14-gsm-blank-phase0.md`](2026-08-14-gsm-blank-phase0.md) PASS | ✅ |
| SC-1 | Импорт 9 файлов / дедуп / сверка итогов | `tests/test_gsm_transaction_import.py`; T3 API | ✅ (авто) |
| SC-2 | Солвер: якоря, баланс, инвариант бака | `tests/test_gsm_balance.py`, `test_gsm_generator.py`, generation API | ✅ (unit/API; полный исторический e2e с эталонными ПЛ — частично checkpoint) |
| SC-3 | Правка дня → downstream; confirmed защищён | `tests/test_gsm_waybill_edit.py` | ✅ |
| SC-4 | Export zip «ПЛ DD.MM.YY.xls», формулы/норма | `tests/test_gsm_export.py`; `gsm_export_service` + `blank.py` | ✅ |
| SC-5 | AuthZ 403 чужим ролям; меню только accountant/admin | `tests/test_gsm_auth.py`; FE role routes / header | ✅ |
| SC-6 | Regression pytest + vitest | GSM suites зелёные; full `pytest tests/` — см. residuals (env noise) | ⚠️ partial |
| SC-7 | Бухгалтер закрыла период; ≤20% ручных дней | Нет подписанного чек-листа | ⏳ pending user |

## Verification commands

```bash
# GSM backend
pytest tests/test_gsm_balance.py tests/test_gsm_generator.py \
  tests/test_gsm_repository.py tests/test_gsm_auth.py \
  tests/test_gsm_transaction_import.py tests/test_gsm_registry.py \
  tests/test_gsm_generation_api.py tests/test_gsm_waybill_edit.py \
  tests/test_gsm_export.py tests/test_import_gsm_history.py -q

# Related trip-feed (optional; needs openlocationcode in venv)
pytest tests/test_build_gsm_trip_feed.py -q

# Frontend
cd frontend && npm test -- --run src/features/gsm && npm run build
```

**Зафиксированные прогоны (оркестрация 2026-08-14):**

- GSM API/registry suite (`test_gsm_auth`, generation, registry, import, waybill_edit, …): **~71 passed**.
- FE `npm test -- --run src/features/gsm`: **50 passed** (после T13); `npm run build` OK.
- Full `pytest tests/` в ограниченном окружении: **1740 passed, 89 failed, 4 skipped, 15 errors** (после ignore collection-broken `test_order` / `test_visualization`) — массово `sqlite3.OperationalError` / read-only FS; не атрибутировать регрессии GSM без чистого прогона вне sandbox.

## Deploy: LibreOffice / soffice

**Проверено:** в корневом `Dockerfile` и `docker/backend/Dockerfile` пакеты `libreoffice` / `soffice` **отсутствуют** (apt: ca-certificates, coinor-cbc, fonts-dejavu-core, gosu). На dev-хосте `/usr/bin/soffice` есть.

Экспорт ПЛ (`gsm_export_service`) вызывает headless LibreOffice. Без `soffice` в образе `POST /gsm/waybills/export` → 500.

**Рекомендуемый шаг установки (документировать / добавить в runtime image при деплое ГСМ):**

```dockerfile
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-calc-nogui \
        # или meta: libreoffice-writer-nogui libreoffice-calc-nogui
    && rm -rf /var/lib/apt/lists/*
```

Проверка в контейнере: `soffice --version` (ожидается `/usr/bin/soffice`). Quirk Phase 0: `--convert-to xls` без фильтра `"MS Excel 97"` — см. phase0 отчёт.  
**В этом срезе Dockerfile не менялся** (предпочитаем явное согласование размера образа).

## Known residuals / follow-ups

1. **SC-7 / приёмка:** прогон реального периода с бухгалтером; % ручных переделок ≤20% — **pending user sign-off**.
2. **`openlocationcode`:** в `requirements.txt`, часто не установлен в venv → FAIL `tests/test_build_gsm_trip_feed.py::test_plus_code_reference_uses_full_tail`. Fix: `pip install openlocationcode` (или `pip install -r requirements.txt`).
3. **Non-GSM / env failures (известные при checkpoint):** collection errors `tests/test_order.py`, `tests/test_visualization.py` (`unable to open database file`); прочие sqlite OperationalError в sandbox — перепроверить с write access.
4. **Review follow-ups:**
   - **T4:** `import_gsm_history` не заполняет `typical_station_ids` на `gsm_route` → солвер чаще уходит в крюк/warning вместо типовой АЗС.
   - **ISS-002:** catch `xlrd.XLRDError` → HTTP 400 + allowlist .xls на импорте ([`../issues/ISS-002-gsm-import-xls-parse-errors.md`](../issues/ISS-002-gsm-import-xls-parse-errors.md)).
   - T3: non-atomic per-row commits / broad IntegrityError→duplicate — hygiene, не блокер.
5. **UI:** журнал транзакций (GET list) — placeholder; импорт работает.
6. **Open questions спеки:** порог крюка 13 км; период = выгрузка vs месяц; версионирование комплектов; «виртуальный клиент» — вне MVP.

## Technical decisions (locked)

- Pure `core/gsm/*` без `app.*` / I/O.
- БД = источник правды после импорта; `ГСМ/**` read-only.
- Регенерация перезаписывает только `draft`; confirmed/exported без `force` → 409.
- Норма в формуле `BS41` патчится при экспорте под машину/сезон.
- Архивация карт/водителей вместо DELETE.

## Next steps

1. User: приёмка периода + фиксация % переделок (SC-7).
2. Deploy: добавить LibreOffice в runtime Dockerfile перед боевым экспортом.
3. Hardening: ISS-002; заполнение `typical_station_ids` при history import.
4. `pip install openlocationcode` в рабочих venv / CI.
5. Чистый full `pytest` + `npm test` вне sandbox перед merge.
