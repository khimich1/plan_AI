# Implementation Plan: GSM — коридор бака по пулу маршрутов парка

Дата: 2026-08-26. Статус: draft, к реализации.
Спека: [`../../specs/gsm-fleet-pool-corridor.md`](../../specs/gsm-fleet-pool-corridor.md).
Идея: [`../../ideas/gsm-fleet-pool-corridor.md`](../../ideas/gsm-fleet-pool-corridor.md).

## Overview

Солвер якоря и дожига выбирает маршрут из `gsm_route` **всего парка**
(база Кузнецкая), строго в коридоре бака; свой важнее чужого.
Жёлтый `borrowed_route`. Приёмка июля Monjaro/952 на копии БД.

## Architecture Decisions

- **`LibraryRoute.vehicle_id` + `generate(..., own_vehicle_id=)`.** Один пул
  на якорь и дожиг; среди подходящих по км — свой важнее чужого.
- **Каскад в `_select_anchor_route_lookahead`** по спеке: `S` / `in_corridor` /
  fallback + `manual_intervention`. Не выбирать вне коридора, если
  `in_corridor` непустой.
- **`LegPlan.route_id: int | None`.** Чужой маршрут: `route_id=None` в ногах;
  адреса и km сохраняем.
- **Старые тесты generator** без `vehicle_id`: default `0`, в тесте
  `own_vehicle_id=0` — все маршруты «свои» (регрессия 24.08 жива).
- Фронт: только подпись + soft export; API контракт не меняется.

## Task List

### Phase 1: Ядро солвера (TDD)

- [x] **Task 1: Поля пула + home/city helpers + загрузка всех маршрутов**
  - **Description:** `LibraryRoute.vehicle_id`. `_norm_addr`, `_is_home_base`,
    `_city_key` (+ unit-тесты ключей). `generate(own_vehicle_id=)`.
    `_load_routes` без фильтра машины. `burn_routes` = тот же fleet∩home
    (не резать `own`). Тесты generator без `own_vehicle_id` — default `0`.
  - **Acceptance:**
    - [x] `_city_key("г.Сергиев Посад, ул.Маслиева…")` == ключ СП.
    - [x] `_is_home_base` истинно для «ул. Кузнецкая, д.18Б» и «ул.Кузнецкая».
    - [x] Существующие тесты generator зелёные после прокидки default.
  - **Verification:** `.venv/bin/python -m pytest tests/test_gsm_generator.py -q`
  - **Dependencies:** нет
  - **Files:** `core/gsm/generator.py`, `core/gsm/models.py` (если `route_id`
    optional сразу — можно отложить в T2), `app/services/gsm_generation_service.py`,
    `tests/test_gsm_generator.py`
  - **Scope:** S

- [x] **Task 2: Каскад якоря + persist без чужого route_id**
  - **Description:** Переписать `_select_anchor_route_lookahead` по спеке
    (fleet ∩ home, `in_corridor`, `S`, rank, warnings). `_round_trip_legs` /
    `_to_route_ref`: `route_id=None` если `chosen.vehicle_id != own`.
    `LegPlan`/`RouteRef.route_id: int | None`. Сериализация в сервисе уже
    пишет `leg.route_id` — уйдёт `null` в JSON.
  - **Acceptance:**
    - [x] Мойка, typical только 95 км, в списке свой 6 км → круг 12, бак ≥ 0,
          нет `manual_intervention`.
    - [x] Свой 280 вне коридора, чужой 265 Кузнецкая в коридоре → круг 530,
          бак ≥ 0, `borrowed_route`.
    - [x] Свой и чужой 6 км оба в коридоре → свой `route_id` не None.
    - [x] Lookahead не выбирает 280 (560 круг), если 265 в `S`.
    - [x] Дожиг: свой 95 км вне коридора, чужой 6 км в коридоре → круг 12,
          `borrowed_route`.
  - **Verification:**
    `.venv/bin/python -m pytest tests/test_gsm_generator.py -q`
  - **Dependencies:** Task 1
  - **Files:** `core/gsm/generator.py`, `core/gsm/models.py`,
    `tests/test_gsm_generator.py`, вызовы `RouteRef`/`LegPlan` по сьюиту
  - **Scope:** M

### Phase 2: UI warning

- [x] **Task 3: `borrowed_route` в подписи и exportGate**
  - **Description:** Мета «Чужой пул» / текст про маршрут другой машины.
    Добавить в `SOFT_WARNING_CODES`. Не в hard-stop.
  - **Acceptance:**
    - [x] `warningMeta("borrowed_route").short` задан.
    - [x] День только с `borrowed_route` не блокирует zip; с
          `manual_intervention` — блокирует.
  - **Verification:**
    `cd frontend && npx vitest run src/features/gsm/lib/waybillWarnings.test.ts src/features/gsm/lib/exportGate.test.ts`
  - **Dependencies:** нет (можно параллельно с T1)
  - **Files:** `frontend/src/features/gsm/lib/waybillWarnings.ts`,
    `exportGate.ts`, соответствующие `.test.ts`, опционально `types/gsm.ts`
  - **Scope:** XS

### Checkpoint

- [x] `.venv/bin/python -m pytest tests/test_gsm_*.py -q` зелёный
- [x] vitest T3 зелёный

### Phase 3: Приёмка (копия БД)

- [x] **Task 4: Июль Monjaro + 952 на `/tmp/plita_fleet_pool.db`**
  - **Description:** `cp plita.db /tmp/plita_fleet_pool.db`.
    `GsmGenerationService.generate` 2026-07-01…2026-08-25, `force=False`,
    vehicle_id 2 и 4. Live `plita.db` не писать.
  - **Acceptance:**
    - [x] Monjaro 01.07 и 16.07: `fuel_end ≥ 0`, нет `manual_intervention`.
    - [x] 952 27.07: `fuel_end ≥ 0`, дневной km ∈ [494, 547] (плечо 247–273),
          не исход «только круг 560».
    - [x] Якоря 30.06 confirmed на месте.
  - **Verification:** скрипт/pytest на копии или ручной вызов сервиса; зафиксировать
    даты/fuel_end/km в отчёте реализации.
  - **Dependencies:** Task 2
  - **Files:** нет кода (приёмка); при необходимости короткий скрипт не коммитить
  - **Scope:** S

## Risks and Mitigations

| Риск | Вероятность | Митигация |
|------|-------------|-----------|
| Регрессия тестов generator: все маршруты `vehicle_id=0` | Средняя | default 0 + `own_vehicle_id=0` в `generate` |
| Дожиг чаще берёт чужой объект | Низкая (ок продуктово) | rank: свой важнее при том же коридоре |
| `_city_key` не склеивает «Сергиев Посад» / «Сергиев-Посад» | Низкая | список составных имён в спеке + тест |
| Приёмка 952: 265 не проходит home-filter | Низкая | оба конца 848-маршрута проверить Кузнецкая |
| `RouteRef.route_id: int \| None` ломает сериализацию | Низкая | omit/null в JSON как в существующем optional |

## Open Questions

Нет. Дожиг = тот же пул парка (решение 2026-08-26: чужие маршруты допустимы).

---

## Как запустить в отдельном окне

1. В Cursor: **новый чат** (чистый контекст).
2. **Agent mode** + **Multitask Mode** (оркестратор сам запускает работников).
3. Вставить промпт ниже целиком и отправить.

---

# Промпт оркестратора (скопировать целиком в новое окно)

```
# Роль

Ты — оркестратор реализации «ГСМ: коридор бака по пулу маршрутов всего парка»
в проекте «Шишов» (FastAPI + SQLite + React/TS). Работаешь в Multitask Mode:
сам запускаешь фоновых работников (worker / test-writer / test-runner /
debugger / reviewer / documenter).

Спека и план УЖЕ утверждены. Не запускай planner. Не переписывай скоуп.
Не «улучшай» дизайн. Не задавай продуктовых вопросов — решения закрыты.

# Контекст — прочитай ДО любой правки кода

1. `.cursor/skills/project-shishov/SKILL.md`
2. `ai_docs/specs/gsm-fleet-pool-corridor.md` — ИСТОЧНИК ИСТИНЫ (алгоритм,
   rank, warnings, дожиг, `_city_key`, Never/Ask first)
3. `ai_docs/develop/plans/2026-08-26-gsm-fleet-pool-corridor.md` — этот план T1–T4
4. `ai_docs/ideas/gsm-fleet-pool-corridor.md` — зачем (июль Monjaro/952)
5. `.cursor/skills/orchestration/SKILL.md`
6. `.cursor/skills/test-driven-development/SKILL.md`
7. `.cursor/skills/incremental-implementation/SKILL.md`

# Продуктовые решения (не обсуждать, не откатывать)

- Кандидаты якоря И дожига = все `gsm_route` парка с базой Кузнецкая 18Б.
- Rank: свой важнее чужого; тот же город важнее; мойка → min km; lookahead →
  min достаточный в коридоре.
- Норма литров — ЭТОЙ машины, не донора.
- Чужой A→B в `gsm_route` этой машины НЕ копировать.
- Persist ног: адреса + km; `route_id=None`, если `chosen.vehicle_id != own`.
- Жёлтый `borrowed_route` (как `balance_route`): zip не hard-stop.
- Если в пуле никого нет в коридоре — день остаётся красным (`manual_intervention`).
- НЕ зажимать `fuel_end=0`. НЕ выдумывать города. Круг = `2×km` (ночёвка — не срез).
- Приёмка июля — только копия БД; якоря 30.06 confirmed не трогать (`force=false`).

# Почему это нужно (якоря июля, уже в live DB)

После сброса июля 2026:
- Monjaro (id=2): красные 01.07, 16.07 — мойка, typical-группа min 95 км
  (круг 190 ≈ 18 л), бак 7.21 л; свой 6 км есть, но `typical_station_ids=NULL`,
  старый каскад брал `min(group)`. 02.07 — каскад от 01.07.
- 952 (id=4): красные 27–28.07 — бак 1.43+50, lookahead окно ~494–547 км/день
  (плечо 247–273); своя сетка прыгает 225→280 (круг 560, Вологда/СП);
  у 848 есть 265 км (Сергиев Посад Маслиева), у Monjaro 260 (Мантурово).
  Солвер грузит только `list_routes(vehicle_id=current)`.

# Миссия

Выполнить T1–T4 плана. После каждой задачи:
- Verification-команда зелёная;
- отметь задачу `[x]` в файле плана (и acceptance-пункты).

Порядок:
- T3 (фронт warning) можно ПАРАЛЛЕЛЬНО с T1 (не пересекаются по файлам).
- T2 зависит от T1. T4 зависит от T2.
- Checkpoint после T1+T2+T3: `tests/test_gsm_*.py` и vitest T3.
- Затем T4 на копии БД.

TDD: сначала красный тест, потом код. Для T2 — три+ кейса из спеки до правки
`_select_anchor_route_lookahead` / `_plan_burn_in`.

# Алгоритм (кратко; детали в спеке)

Замена `_select_anchor_route_lookahead` в `core/gsm/generator.py`:

```
fleet = LibraryRoute с 2×km ≤ max_daily_km и хотя бы один конец _is_home_base
in_corridor = { r in fleet | _fits_corridor(...) }
S = in_corridor; если km_needed>0: S = { r | _daily_km(r) >= km_needed }

если S непуст: chosen = rank(S)
иначе если in_corridor непуст: chosen = rank(in_corridor) по max daily_km
     + manual_intervention
иначе: min(fleet или typical-группа) + manual_intervention
```

ЗАПРЕЩЕНО: если `in_corridor` непустой, выбирать маршрут вне коридора
(в т.ч. 560 км при окне 547). Дыра 24.08: пустая typical + fallback min(group)
— не повторять, если в fleet есть короткий свой.

Rank (меньше лучше): (1) чужой=1 свой=0 (2) не тот же город (3) daily_km
(4) -frequency, route_id.

Дожиг `_plan_burn_in`: тот же fleet∩home∩corridor; политика км как сейчас;
среди подходящих — rank (свой важнее); чужой → `borrowed_route`, `route_id=None`.

`LibraryRoute.vehicle_id` default 0. `generate(..., own_vehicle_id=0)` —
старые тесты без vehicle_id остаются «все свои».

`_load_routes` в `gsm_generation_service.py`: `list_routes()` БЕЗ фильтра
машины; передать `own_vehicle_id` в `generate`.

# Задачи

## T1 — поля пула + home/city + загрузка всех маршрутов
Files: `core/gsm/generator.py`, `core/gsm/models.py` (route_id optional можно
в T2), `app/services/gsm_generation_service.py`, `tests/test_gsm_generator.py`
Acceptance: `_city_key("г.Сергиев Посад, ул.Маслиева…")` == ключ СП;
`_is_home_base` для «ул. Кузнецкая, д.18Б» и «ул.Кузнецкая»; существующие
тесты generator зелёные с default 0.
Verify: `.venv/bin/python -m pytest tests/test_gsm_generator.py -q`

## T2 — каскад якоря + persist без чужого route_id + дожиг из пула
Files: `core/gsm/generator.py`, `core/gsm/models.py`, `tests/test_gsm_generator.py`
Acceptance (unit):
1. Мойка, typical только 95 км, свой 6 км → круг 12, бак ≥ 0, нет
   `manual_intervention` (допустим `balance_route`).
2. Свой 280 вне коридора, чужой 265 Кузнецкая в коридоре → круг 530, бак ≥ 0,
   `borrowed_route`.
3. Свой и чужой 6 км оба в коридоре → берётся свой, `route_id` не None.
4. Lookahead не выбирает 280 (круг 560), если 265 в S.
5. Дожиг: свой min 95 вне коридора, чужой 6 в коридоре → будний круг 12,
   `borrowed_route`.
Verify: `.venv/bin/python -m pytest tests/test_gsm_generator.py -q`

## T3 — warning `borrowed_route` (параллельно с T1)
Files: `frontend/src/features/gsm/lib/waybillWarnings.ts`, `exportGate.ts`,
их `.test.ts`, опционально `types/gsm.ts`
Acceptance: `warningMeta("borrowed_route").short` задан (напр. «Чужой пул»);
день только с `borrowed_route` не блокирует zip; с `manual_intervention` —
блокирует. В `SOFT_WARNING_CODES`. Не в hard-stop. Бейдж госномера донора
не делать.
Verify: `cd frontend && npx vitest run src/features/gsm/lib/waybillWarnings.test.ts src/features/gsm/lib/exportGate.test.ts`

## Checkpoint
`.venv/bin/python -m pytest tests/test_gsm_*.py -q`
vitest T3 зелёный.

## T4 — приёмка на КОПИИ, не на live
`cp plita.db /tmp/plita_fleet_pool.db`
`GsmGenerationService.generate` 2026-07-01…2026-08-25, `force=False`,
vehicle_id 2 (Monjaro) и 4 (Tugella 952).
Acceptance:
- Monjaro 01.07 и 16.07: `fuel_end ≥ 0`, нет `manual_intervention`.
- 952 27.07: `fuel_end ≥ 0`, дневной km ∈ [494, 547] (плечо 247–273),
  не исход «только круг 560».
- Якоря 30.06 confirmed на месте.
Скрипт приёмки не коммитить. Даты/fuel_end/km — в отчёт documenter.
Verify: вызов сервиса на копии; зафиксировать числа.

# Ограничения — Never / Ask first

Never:
- схема БД
- зажим `fuel_end=0`
- синтетический адрес / выдуманный город
- `force=true` на confirmed 30.06
- запись generate в live `plita.db`
- коммит / push / git config
- останавливать `run+logs.sh` / uvicorn :8000
- откатывать чужие незакоммиченные изменения (usage report, calendar, fleet UX)

Ask first (стоп и спроси пользователя, не делай сам):
- ночёвка / соло-плечо
- запись чужого A→B в `gsm_route`
- генерация на live `plita.db` после зелёной копии

Only files from the plan. Не чинить падения pytest вне `tests/test_gsm_*.py`.

# Окружение

- Репозиторий: `/home/roman/project/Шишов`
- Backend: `.venv/bin/python -m pytest tests/test_gsm_generator.py -q`
  затем `tests/test_gsm_*.py -q`
- Frontend: `cd frontend && npx vitest run src/features/gsm/lib/waybillWarnings.test.ts src/features/gsm/lib/exportGate.test.ts`
- Live `plita.db` — только чтение. Писать generate — только `/tmp/plita_fleet_pool.db`.

# Машины (для T4 и фикстур)

| id | Машина      | Госномер     | Бак | Лето | Водитель    |
|----|-------------|--------------|-----|------|-------------|
| 1  | Palisade    | О 521 УХ 44  | 70  | 14.5 | Шишов       |
| 2  | Monjaro     | О 165 ХУ 44  | 60  | 9.5  | Лоншакова   |
| 3  | Tugella 848 | О 848 ХР 44  | 55  | 9.4  | Кулигин     |
| 4  | Tugella 952 | О 952 ХР 44  | 55  | 9.4  | Кулигин     |

База всех: Кострома, Кузнецкая 18Б.
`list_routes(vehicle_id=None)` в `app/repositories/gsm_repository.py` уже есть.

# После всех задач

1. Чекбоксы плана T1–T4 `[x]`.
2. Documenter: `ai_docs/develop/reports/2026-08-26-gsm-fleet-pool-corridor.md`
   (что сделано, числа T4, какие дни остались красными).
3. Короткий доклад пользователю: тесты, результат копии Monjaro 01/16.07 и
   952 27.07, live не трогали. Коммит не делать.
```
