# Implementation Plan: GSM — ориентация круга с базы + домашний `route_id`

Дата: 2026-08-27. Статус: done.
Спека: [`../../specs/gsm-home-oriented-round-trip.md`](../../specs/gsm-home-oriented-round-trip.md).
Идея: [`../../ideas/gsm-home-oriented-round-trip.md`](../../ideas/gsm-home-oriented-round-trip.md).

## Overview

После выбора км солвер (и выбор строки справочника в UI) рисует день
**Кузнецкая 18Б → объект → …**. Если у своей машины есть зеркало с тем же km
и переставленными концами — в ноги пишем его `route_id`. Бак, rank и warnings
не меняем. Live `plita.db` не пишем.

## Architecture Decisions

- **Ориентация в `_emit_day`**, не в rank. Якорь и дожиг одним путём.
  `generate` / `_emit_maybe_manual` прокидывают `catalog=route_list`.
- **`dataclasses.replace` на frozen `LibraryRoute`:** переставляем `addr_a`/
  `addr_b`; при своём зеркале — ещё `route_id`. Остальное с `chosen`
  (текст адресов выбранной строки, не зеркала).
- **Чужой пул:** `_persisted_route_id` по-прежнему `None`; twin донора не
  пишем. Ноги всё равно с базы.
- **Фронт:** те же правила в `gsmHome.ts`; `libraryRouteToLegs(route, catalog)`
  остаётся **одной** ногой. Drawer на 2 плеча не переводим.
- **TDD:** красные тесты ориентации → хелперы → emit → UI.

```
_find_home_twin / _orient_home_round_trip
        │
        ├── _emit_day (якорь + дожиг)
        │
        └── TS: isGsmHomeBase + libraryRouteToLegs
                    │
                    └── Drawer / ManualWaybillDialog передают catalog
```

## Task List

### Phase 1: Солвер (TDD)

- [x] **Task 1: `_find_home_twin` + `_orient_home_round_trip`**
  - **Description:** Чистые хелперы по спеке. `_norm_addr` / `_is_home_base`
    не трогать. Юниты вызывают хелперы напрямую (как `_is_home_base`).
  - **Acceptance:**
    - [x] Объект→дом + зеркало той же `vehicle_id`/km → `addr_a` база,
          `addr_b` текст объекта с chosen, `route_id` зеркала.
    - [x] Без зеркала → переворот, `route_id` исходный.
    - [x] Несколько зеркал → меньший `route_id`.
    - [x] `chosen.vehicle_id != own` → переворот, `route_id` у LibraryRoute
          можно не менять; persist потом `None`.
    - [x] Дом-первым / оба конца база / ни одного — identity.
    - [x] Сам `chosen` не возвращается как twin.
  - **Verification:**
    `.venv/bin/python -m pytest tests/test_gsm_generator.py -k "orient or home_twin" -q`
  - **Dependencies:** нет
  - **Files:** `core/gsm/generator.py`, `tests/test_gsm_generator.py`
  - **Scope:** S

- [x] **Task 2: Прокинуть catalog в emit + generate-кейсы**
  - **Description:** `_emit_day` сначала `_orient_home_round_trip(route,
    catalog=..., own_vehicle_id=...)`. `_emit_maybe_manual` и оба вызова
    в `generate` (дожиг и якорь) передают `route_list`. Rank/warnings —
    по неориентированному `chosen`.
  - **Acceptance:**
    - [x] Generate: свой 59 объект→дом + 64 зеркало, km=225 → круг 450,
          `legs[0].addr_a` база, `route_id=64`, `balance_route` если был
          бы без ориентации (warnings не из-за смены id).
    - [x] Без зеркала → старт с базы, свой исходный `route_id`.
    - [x] Чужой объект→дом → старт с базы, `route_id is None`.
    - [x] Дожиг с объектом-первым → первая нога с базы.
    - [x] Фикстуры A/B без Кузнецкой — ноги как сейчас.
    - [x] Регрессия: км/`fuel_end` существующих тестов не плывут.
  - **Verification:**
    `.venv/bin/python -m pytest tests/test_gsm_generator.py tests/test_gsm_*.py -q`
  - **Dependencies:** Task 1
  - **Files:** `core/gsm/generator.py`, `tests/test_gsm_generator.py`
  - **Scope:** S

### Checkpoint: солвер

- [x] `tests/test_gsm_*.py` зелёные
- [x] Новый кейс «59/64» проходит без правки rank

### Phase 2: UI

- [x] **Task 3: `isGsmHomeBase` + ориентация `libraryRouteToLegs`**
  - **Description:** `frontend/src/features/gsm/lib/gsmHome.ts` — `normGsmAddr`
    / `isGsmHomeBase` как Python (кузнецкая+18; без цифры — база).
    `libraryRouteToLegs(route, catalog?)`: одна нога; twin по km +
    переставленным `_normAddr`; объект-первым → `from` база, `route_id` twin
    или исходный. Кейс без catalog / без Кузнецкой — как сейчас (тест
    «Завод→Объект» жив).
  - **Acceptance:**
    - [x] `isGsmHomeBase` истинно для «ул. Кузнецкая, д.18Б» и «ул.Кузнецкая».
    - [x] `libraryRouteToLegs(59, catalog с 64)` → дом→объект, `route_id=64`.
    - [x] Без 64 в catalog → дом→объект, `route_id=59`.
    - [x] Без базы в адресах → без переворота.
  - **Verification:**
    `cd frontend && npx vitest run src/features/gsm/lib/downstreamPreview.test.ts`
  - **Dependencies:** нет (правила из спеки; лучше после T1, чтобы не разъехаться)
  - **Files:** `frontend/src/features/gsm/lib/gsmHome.ts` (new),
    `downstreamPreview.ts`, `downstreamPreview.test.ts`
  - **Scope:** S

- [x] **Task 4: Catalog в drawer и ручном ПЛ**
  - **Description:** `WaybillDayDrawer.applyLibraryRoute` и
    `ManualWaybillDialog` передают загруженный список маршрутов машины
    в `libraryRouteToLegs`. Тест drawer: в mock ROUTES объект-первым +
    зеркало; после выбора в PATCH `route[0].from` — база, `route_id` зеркала.
    Существующий тест выбора id=5 (дом-первым) не ломать.
  - **Acceptance:**
    - [x] Выбор 59 при 64 в списке → payload `from` Кузнецкая, `route_id=64`.
    - [x] Ручной диалог тоже передаёт catalog (тест или тот же хелпер).
  - **Verification:**
    `cd frontend && npx vitest run src/features/gsm/lib/downstreamPreview.test.ts src/features/gsm/components/WaybillDayDrawer.test.tsx src/features/gsm/components/ManualWaybillDialog.test.tsx`
  - **Dependencies:** Task 3
  - **Files:** `WaybillDayDrawer.tsx`, `WaybillDayDrawer.test.tsx`,
    `ManualWaybillDialog.tsx`, при необходимости `ManualWaybillDialog.test.tsx`
  - **Scope:** S

### Checkpoint: Complete

- [x] `.venv/bin/python -m pytest tests/test_gsm_*.py -q`
- [x] vitest T3+T4 зелёный
- [x] Спека Success Criteria 1–7 закрыты
- [x] Live `plita.db` не писали

## Risks and Mitigations

| Риск | Вероятность | Митигация |
|------|-------------|-----------|
| `_norm_addr` не склеивает 59↔64 (пр-т / пр-кт) | Средняя | юнит на реальных строках из идеи; нет пары → переворот без смены id (допустимо по спеке) |
| Существующий generate-тест с объектом-первым и Кузнецкой | Низкая | grep `addr_b=_HOME`; поправить ожидание ног, не км |
| TS-эвристика разъедется с Python | Средняя | одни и те же строки в `test_is_home_base_kuznetskaya` и vitest |
| Drawer ставит `km` = плечо, не 2×km | Уже так | не чинить в этом срезе |
| `replace(route_id=twin)` меняет typical/warnings | Нет | warnings до emit; typical с chosen не используются после persist |

## Open Questions

Нет. Приёмка на копии БД (13.07 952) в срез не входит — бухгалтер
перегенерирует после выкладки.

---

**Следующий шаг после апрува:** Task 1 (красные тесты хелперов).
