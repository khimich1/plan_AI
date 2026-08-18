# Implementation Plan: GSM lookahead-генератор с географией (round-trip, направление к следующей заправке)

## Overview

Доработка генератора путевых листов: устранение массовых падений (422) на
плотных участках заправок. Генератор v2 смотрит на следующую заправку
(lookahead), выбирает маршрут по направлению к ней (география), строит день
как round-trip (2 плеча) и при локальной нерешаемости сохраняет проблемный
день как `draft` с `manual_intervention` вместо 422 на весь период.

Спека: [`ai_docs/specs/gsm-geo-lookahead-generator.md`](../../specs/gsm-geo-lookahead-generator.md).
Идея: [`ai_docs/ideas/gsm-geo-lookahead-generator.md`](../../ideas/gsm-geo-lookahead-generator.md).
Базовый модуль: [`2026-08-14-gsm-module.md`](2026-08-14-gsm-module.md).

## Architecture Decisions

1. **Ночёвки — фаза 2, не MVP.** Проверено на истории: при `max_daily_km=700`
   ни у одной машины нет жёстких случаев, требующих ночёвки для сходимости
   баланса. Плотные заправки почти всегда рядом с базой (Кострома, 0–5 км);
   заправки «в пути» (Ярославль 63 км, Реутов 300 км) решаются длинным
   round-trip. Ночёвки — «улучшатель правдоподобия» длинных поездок (фаза 2).
2. **Бак в приоритете над географией.** Сначала отбор маршрутов по сходимости
   баланса (km достаточный), затем мягкая сортировка по направлению к
   следующей АЗС. 422 блокирует всё — хуже, чем «неидеальное» направление.
3. **Минимальный достаточный km.** Среди решающих баланс маршрутов берём
   минимальный (не жжём лишнего «в никуда») — экономия пробега и правдоподобие.
4. **Частичная генерация вместо 422.** Локально нерешаемый якорь → draft
   `manual_intervention`, период собирается целиком. Контракт `generate`
   меняется: 200 с `problematic_days` вместо 422 на нерешаемость.
5. **Чистый core.** `core/gsm/geo.py` — чистая математика (haversine, bearing),
   без I/O и `app.*`. Геокодинг — только из `scripts/`, не из backend runtime.
6. **Детерминизм.** Один вход → один выход; tie-break через частоту,
   route_id, seed. Геометрия детерминирована по координатам.

## Task List

### Phase 0: Data (геоданные + привязка станций)

- [x] Task 1: `scripts/geocode_gsm_stations.py` — геокодинг 27 станций ✅ Completed
  - **Description:** Для каждой `gsm_station` с `lat IS NULL` — запрос к Nominatim через существующий кэш `ГСМ/geo_cache/addresses.json`, запись `lat/lon`, `geocode_source='nominatim'`. Rate-limit 1 req/s. Трассовые станции («М8, 87 км») — по трассе.
  - **Acceptance criteria:**
    - [ ] Все 76 станций имеют `lat/lon` (или помечены `geocode_source='failed'` с логом).
    - [ ] Повторный запуск не дублирует запросы (кэш), не трогает уже геокодированные.
    - [ ] Только UPDATE NULL-полей; существующие координаты не затираются.
  - **Verification:** `python scripts/geocode_gsm_stations.py --db plita.db` → отчёт «27/27 геокодировано»; SQL-запрос показывает 0 станций с `lat IS NULL` (или осмысленный список failed).
  - **Dependencies:** None
  - **Files:** `scripts/geocode_gsm_stations.py` (новый)
  - **Estimated scope:** S
  - **Note:** требует сети к nominatim (только scripts, не backend runtime).

- [x] Task 2: `core/gsm/geo.py` + тесты ✅ Completed
  - **Description:** `GeoPoint`, `haversine_km`, `bearing_deg`, `angle_diff_deg`, `point_to_segment_km` (отклонение станции от линии A→B). Чистая математика, frozen DTO.
  - **Acceptance criteria:**
    - [ ] Контрольные пары: Кострома→Ярославль ~63,5 км / азимут ~257°; Кострома→Москва ~300 км / ~224° (±5%).
    - [ ] `angle_diff_deg` корректен на границах (0/360, 179/181).
    - [ ] `point_to_segment_km`: станция на отрезке → ~0; далеко → большое значение.
  - **Verification:** `pytest tests/test_gsm_geo.py -q` зелёный.
  - **Dependencies:** None (параллельно Task 1)
  - **Files:** `core/gsm/geo.py` (новый), `tests/test_gsm_geo.py` (новый)
  - **Estimated scope:** S

- [x] Task 3: `scripts/link_route_stations.py` — заполнение `typical_station_ids` ✅ Completed
  - **Description:** (а) из 4 исторических ПЛ — станция заправки дня → маршрут дня; (б) по географии: геокодинг адресов A/B маршрутов (кэш), станция «на маршруте» если `point_to_segment_km < 15`. Запись JSON `typical_station_ids`.
  - **Acceptance criteria:**
    - [ ] У маршрутов, проходящих мимо известных АЗС, `typical_station_ids` заполнены.
    - [ ] Palisade имеет маршруты через станцию id=1 (КТК Магистральная) — раньше было 0.
    - [ ] Станция без координат пропускается с логом (не падает).
  - **Verification:** `python scripts/link_route_stations.py --db plita.db` → отчёт покрытия; SQL: `SELECT COUNT(*) FROM gsm_route WHERE typical_station_ids IS NOT NULL AND typical_station_ids != '[]'` > 0.
  - **Dependencies:** Task 1 (координаты станций), Task 2 (point_to_segment)
  - **Files:** `scripts/link_route_stations.py` (новый)
  - **Estimated scope:** M

### Checkpoint: Data Gate
- [x] Все станции с координатами; `typical_station_ids` заполнены; Palisade май имеет маршруты через свои АЗС.

---

### Phase 1: Core generator v2

- [x] Task 4: Модели + round-trip (2 плеча) ✅ Completed
  - **Description:** `core/gsm/models.py` — LegPlan/WaybillDay поддерживает 2 плеча. `generator.py`: день = round-trip (туда + обратно), дневной km = `2×km`, сжигается `burn(2×km)`. Регрессия v1 сохранена.
  - **Acceptance criteria:**
    - [ ] Якорь и дожигание формируют 2 плеча в `route_json`.
    - [ ] Баланс считается от суммарного дневного km.
    - [ ] Существующие тесты генератора адаптированы/зелёные.
  - **Verification:** `pytest tests/test_gsm_generator.py -q` зелёный.
  - **Dependencies:** None
  - **Files:** `core/gsm/models.py`, `core/gsm/generator.py`, `tests/test_gsm_generator.py`
  - **Estimated scope:** M

- [x] Task 5: Lookahead на якоре ✅ Completed
  - **Description:** Для каждого якоря — `next_anchor` (дата + `fuel_issued`). Если свободных будней между якорями нет/не хватает: `burn_needed = fuel_after − (tank − Q_next)`, `km_needed = burn_needed/норма×100`. Выбор маршрута с `2×km ≥ km_needed`, минимальный достаточный, `2×km ≤ max_daily_km`.
  - **Acceptance criteria:**
    - [ ] Плотные якоря подряд (04–06.05) выбирают удлинённый маршрут.
    - [ ] При `burn_needed ≤ 0` — обычный выбор (без lookahead).
    - [ ] `2×km > max_daily_km` — маршрут отклоняется.
  - **Verification:** `pytest tests/test_gsm_generator.py -q -k lookahead` зелёный.
  - **Dependencies:** Task 4
  - **Files:** `core/gsm/generator.py`, `tests/test_gsm_generator.py`
  - **Estimated scope:** M

- [x] Task 6: География — мягкая сортировка по направлению ✅ Completed
  - **Description:** 3 уровня приоритета: (1) станция на маршруте И направление к след. АЗС (`angle_diff ≤ 90°`); (2) станция на маршруте; (3) подходящий km (крюк). Внутри уровня — минимальный достаточный km, tie-break частота/route_id/seed.
  - **Acceptance criteria:**
    - [ ] 20.05 Ярославль → 21.05 Москва: выбирается маршрут в сторону Москвы.
    - [ ] При отсутствии координат/направления — fallback на km.
    - [ ] Бак в приоритете: сначала отбор по сходимости, потом география.
  - **Verification:** `pytest tests/test_gsm_generator.py -q -k direction` зелёный.
  - **Dependencies:** Task 2, Task 3 (typical_station_ids), Task 5
  - **Files:** `core/gsm/generator.py`, `core/gsm/geo.py`, `tests/test_gsm_generator.py`
  - **Estimated scope:** M

- [x] Task 7: Частичная генерация (manual_intervention вместо 422) ✅ Completed
  - **Description:** Нерешаемый якорь → draft `manual_intervention` (лучший доступный маршрут, баланс помечен нарушенным), генерация продолжается от фактического `fuel_end`. `GenerateResult` + `problematic_days`.
  - **Acceptance criteria:**
    - [ ] Нерешаемый якорь не валит период; остальные дни генерируются.
    - [ ] `problematic_days` содержит date, reason, detail, fuel_before, fuel_to_issue, tank_volume.
    - [ ] Баланс после manual-дня пересчитывается от фактического остатка.
  - **Verification:** `pytest tests/test_gsm_generator.py -q -k manual` зелёный.
  - **Dependencies:** Task 5, Task 6
  - **Files:** `core/gsm/generator.py`, `core/gsm/models.py`, `tests/test_gsm_generator.py`
  - **Estimated scope:** M

### Checkpoint: Core Gate
- [x] `pytest tests/test_gsm_geo.py tests/test_gsm_generator.py -q` зелёный; май Palisade собирается без 422 в unit-симуляции.

---

### Phase 2: Service + API

- [x] Task 8: `gsm_setting.max_daily_km` + сервис ✅ Completed
  - **Description:** Чтение `max_daily_km` (default 700) из `gsm_setting`; передача в генератор. `gsm_generation_service` передаёт настройки и координаты станций.
  - **Acceptance criteria:**
    - [ ] Дефолт 700 при отсутствии настройки.
    - [ ] Настройка переопределяет дефолт.
  - **Verification:** `pytest tests/test_gsm_generation_service.py -q` зелёный.
  - **Dependencies:** Task 5
  - **Files:** `app/services/gsm_generation_service.py`, `app/services/gsm_registry_service.py`, `tests/`
  - **Estimated scope:** S

- [x] Task 9: API — частичная генерация, `problematic_days` ✅ Completed
  - **Description:** `POST /gsm/waybills/generate` → 200 с `problematic_days` вместо 422 на нерешаемость. `WaybillGenerateResult` + `problematic_days`, `manual_days`. 422 только для конфигурационных ошибок (нет машины/маршрутов/водителя).
  - **Acceptance criteria:**
    - [ ] Нерешаемый якорь → 200 с заполненным `problematic_days`.
    - [ ] `409` (confirmed без force) сохраняется.
    - [ ] AuthZ `REQUIRE_ACCOUNTING` без изменений.
  - **Verification:** `pytest tests/test_gsm_api_integration.py -q` зелёный.
  - **Dependencies:** Task 7, Task 8
  - **Files:** `app/api/v1/endpoints/gsm.py`, `app/schemas/gsm.py`, `tests/test_gsm_api_integration.py`
  - **Estimated scope:** M

### Checkpoint: API Gate
- [x] `POST /gsm/waybills/generate` на май Palisade возвращает 200 с днями + `problematic_days`.

---

### Phase 3: Frontend

- [x] Task 10: Warning-коды + бейджи ✅ Completed
  - **Description:** `waybillWarnings.ts` + `manual_intervention`, `balance_route`. `VehiclePeriodStrip.tsx` — бейдж «маршрут для баланса», красная подсветка проблемных дней, клик → drawer.
  - **Acceptance criteria:**
    - [ ] Новые коды имеют short/reason.
    - [ ] Проблемный день визуально отличим и кликабелен.
  - **Verification:** `cd frontend && npm test -- --run` зелёный.
  - **Dependencies:** Task 9
  - **Files:** `frontend/src/features/gsm/lib/waybillWarnings.ts`, `frontend/src/features/gsm/components/VehiclePeriodStrip.tsx`, тесты
  - **Estimated scope:** S

- [x] Task 11: `GsmPeriodView` — отображение частичной генерации ✅ Completed
  - **Description:** После generate — «Создано N дней, M требуют ручной доработки» со списком дат; убрана обработка 422 как фатальной.
  - **Acceptance criteria:**
    - [ ] Показывается сводка частичной генерации.
    - [ ] 422 обрабатывается только для конфигурационных ошибок.
  - **Verification:** `cd frontend && npm test -- --run` зелёный.
  - **Dependencies:** Task 9, Task 10
  - **Files:** `frontend/src/features/gsm/components/GsmPeriodView.tsx`, тесты
  - **Estimated scope:** S

---

### Phase 4: Acceptance + docs

- [x] Task 12: Acceptance Palisade май 2026 ✅ Completed
  - **Description:** Прогон 04.05–31.05 (старт 28 л / 128327 км) через UI/скрипт. Проверка: без 422, ≤3 manual days, маршруты следуют заправкам.
  - **Acceptance criteria:**
    - [ ] Генерация завершается 200 без 422.
    - [ ] ≤3 дней `manual_intervention` (группа 20–22.05 допустима).
    - [ ] Маршруты правдоподобны (следуют логистике заправок).
    - [ ] Все существующие тесты зелёные (регрессия v1).
  - **Verification:** отчёт о прогоне; `pytest tests/ -q` и `cd frontend && npm test -- --run` зелёные.
  - **Dependencies:** Task 1–11
  - **Files:** `ai_docs/develop/reports/2026-08-15-gsm-geo-lookahead-acceptance.md`
  - **Estimated scope:** S

- [x] Task 13: Документация ✅ Completed
  - **Description:** Обновить feature-doc, changelog; отметить фазу 2 (ночёвки) как future work.
  - **Acceptance criteria:**
    - [ ] `ai_docs/develop/features/gsm-module-putevye-listy.md` — секция про lookahead/географию.
    - [ ] CHANGELOG обновлён.
  - **Verification:** документы обновлены.
  - **Dependencies:** Task 12
  - **Files:** `ai_docs/develop/features/gsm-module-putevye-listy.md`, `ai_docs/changelog/CHANGELOG.md`
  - **Estimated scope:** S

## Dependency Graph

```
Phase 0 (Data):        Task 1 (geocode) ──┐
                       Task 2 (geo.py) ───┼──> Task 3 (link stations)
                                          │
Phase 1 (Core):        Task 4 (round-trip)│
                          └──> Task 5 (lookahead) ──> Task 6 (direction) ──> Task 7 (partial)
Phase 2 (Service/API): Task 8 (settings) ──> Task 9 (API)
Phase 3 (Frontend):    Task 10 (badges) ──> Task 11 (period view)
Phase 4 (Acceptance):  Task 12 (acceptance) ──> Task 13 (docs)
```

Параллельно: Task 1 ∥ Task 2; Phase 3 (frontend) ∥ после Task 9.

## Risks

| Риск | Митигация |
|---|---|
| Геокодинг трассовых станций неточен («М8, 87 км») | Для направления достаточно; point_to_segment с порогом 15 км пересмотреть после acceptance |
| `typical_station_ids` покрытие неполное (мало истории) | Fallback: приоритет 3 (km + крюк) работает и без привязки |
| День 350–410 км часто — правдоподобие для налоговой | Acceptance у бухгалтера (SC-G6); фаза 2 (ночёвки) как улучшение |
| Изменение контракта `generate` ломает существующий UI | Task 11 обновляет `GsmPeriodView`; интеграционные тесты на новый контракт |
| Регрессия v1 на простых периодах | Регрессионные тесты в Task 4/5 |

## Out of Scope (фаза 2)

- Дальний рейс с ночёвкой (`overnight_trip`, `return_leg`, бейдж «ночёвка»).
- Составные дни из 3+ плеч (полный солвер).
- Калибровка `max_daily_km` по машине.
