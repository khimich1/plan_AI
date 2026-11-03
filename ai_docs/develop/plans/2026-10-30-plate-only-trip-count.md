# Implementation Plan: Рейсы только у плит, правка только в конструкторе

**Created:** 2026-10-30  
**Status:** IMPLEMENT  
**Spec:** [`ai_docs/specs/plate-only-trip-count.md`](../../specs/plate-only-trip-count.md)  
**Idea:** [`ai_docs/ideas/plate-only-trip-count.md`](../../ideas/plate-only-trip-count.md)

## Overview

Ручное число рейсов остаётся у плит и только на шаге «Итоги и скидка». Сваи и ФБС снова считаются формулой, даже если в `trip_counts_json` уже лежит accepted. Архив при открытии не сбрасывает плиты по массе и не пишет JSON. Поле «Кол. рейсов» и плашка «было / стало» из карточки архива уходят. Строка «Услуга по доставке» остаётся и показывает рейсы своей суммы.

## Architecture Decisions

- **A1. Сваи и ФБС не читают accepted.** В `resolve_trip_counts` для `pile` и `fbs_lm` в сумму идёт `formula`. Запись в JSON для этих котлов можно оставить формулой, но читать её как ручное число нельзя. Ключи не удаляем.
- **A2. Сброс плит по массе — только конструктор.** Новый аргумент `reset_plate_on_mass_change`. Черновик вызывает расчёт с `True`. Архивный GET — с `False`: берётся сохранённый `plate.accepted`, notice плит пустой.
- **A3. GET архива не пишет файл.** Блок `save_trip_counts_json` в сборке деталей КП удаляется. Запись остаётся в конструкторе (`get_draft_details`) и в PATCH, которым конструктор сохраняет OK плит.
- **A4. `pile_trip_overrides` не путать с accepted.** Вопрос «сколько машин» у марки без нормы остаётся. Тесты сброса ручных рейсов свай переписываются. Тесты overrides не ослабляются.
- **A5. UI.** В `CalculationResultStep` поле и плашка только у плит. В `OfferDetailsDrawer` нет `TripCountField` и нет `TripCountNotices` ни в одной ветке. Плашки «Рейсов свай» / «Рейсов ФБС» остаются. Подпись строки доставки берёт рейсы из итога, не `ceil` от веса карточки.
- **A6. API можно не ломать.** Поля `pile_trips` и `fbs_lm_trips` в PATCH могут прийти, но расчёт их не применяет. Фронт архива их больше не шлёт.

```
конструктор  reset_plate_on_mass_change=True  → пишет JSON
архив GET    reset_plate_on_mass_change=False → JSON не пишет
сваи / ФБС   всегда formula
```

## Risks

| Риск | Что делаем |
|------|------------|
| КП с ручными рейсами свай изменит итог | Это решение идеи: формула снова главная. В отчёте после прогона назвать, что итог может сдвинуться |
| Строка «(49 рейсов)» считается от веса карточки, а деньги от 57 | Подпись собрать из `plate_trips + pile_trips + fbs_lm_trips` ответа, где плиты — сохранённые |
| Тесты логистики ждут accepted свай и notice свай | Переписать ожидание на формулу, не удалять сами кейсы |

## Implementation order

| Phase | Focus | Depends |
|-------|-------|---------|
| 1 | Расчёт: сваи и ФБС — формула, архив не сбрасывает плиты | — |
| 2 | GET архива не пишет JSON | 1 |
| 3 | Убрать поле и плашку с лишних экранов | 1 |

## Task List

### Phase 1: Расчёт

## Task 1: RED — формула свай и ФБС, архив не сбрасывает плиты

**Description:** Тесты до кода. `trip_counts.pile.accepted = 125` при формуле 2 даёт `pile_trips == 2` и доставку `pile_logistics_cost × 2`. То же для ФБС против `ceil(кг / 20000)`. Плиты: accepted 57 при той же массе остаётся 57. С `reset_plate_on_mass_change=False` масса, которая дала бы формулу 49, оставляет 57 и `plate_trips_notice is None`. С `True` та же масса пишет 49 и notice «было 57, стало 49». `pile_trip_overrides` по марке без нормы по-прежнему задаёт машины.

**Acceptance criteria:**
- [x] Новые проверки красные
- [x] Кейсы «ручные рейсы свай сохраняются» переписаны и тоже красные, пока код старый

**Verification:**
- [x] `pytest tests/test_commercial_logistics_cost.py tests/test_commercial_fbs_lm_delivery.py -q`

**Dependencies:** None

**Files likely touched:**
- `tests/test_commercial_logistics_cost.py`
- `tests/test_commercial_fbs_lm_delivery.py`

**Estimated scope:** M

## Task 2: GREEN — `resolve_trip_counts` и флаг архива

**Description:** Сваи и ФБС всегда берут формулу. Плиты сбрасываются по массе только если флаг включён. `calculate_total_cost` прокидывает флаг. Конструктор его не выключает.

**Acceptance criteria:**
- [x] Тесты Task 1 зелёные
- [x] Длинные сваи и ёмкости 18600 / 20000 не изменены

**Verification:**
- [x] `pytest tests/test_commercial_logistics_cost.py tests/test_commercial_fbs_lm_delivery.py tests/test_pile_trip_pricing.py -q`

**Dependencies:** Task 1

**Files likely touched:**
- `core/trip_counts.py`
- `core/commercial_pricing.py`
- `app/services/commercial_calculation_service.py`

**Estimated scope:** M

### Phase 2: Архив не пишет

## Task 3: GET деталей не сохраняет JSON

**Description:** Убрать `save_trip_counts_json` из сборки карточки. Вызов `calculate_total_cost` для этой карточки идёт с `reset_plate_on_mass_change=False`. Тест: КП с accepted плит 57 и весом на формулу 49 после GET всё ещё хранит 57; повторный GET тоже.

**Acceptance criteria:**
- [x] `trip_counts_json` до и после GET совпадает
- [x] В ответе `plate_trips == 57`, notice плит пустой
- [x] Конструктор по-прежнему пишет JSON при OK и при смене веса

**Verification:**
- [x] `pytest tests/test_archive_endpoints.py tests/test_commercial_logistics_cost.py -q`

**Dependencies:** Task 2

**Files likely touched:**
- `app/services/archive_service.py`
- `tests/test_archive_endpoints.py`

**Estimated scope:** S

### Phase 3: Экраны

## Task 4: RED — поле только у плит, в архиве его нет

**Description:** Vitest. Шаг итогов: у плит есть «Кол. рейсов» и плашка, у свай и ФБС поля нет. Drawer статуса «в архиве»: нет «Кол. рейсов», нет «пересчитаны», строка «Услуга по доставке» есть. Скобки строки равны рейсам суммы, не `ceil` веса.

**Acceptance criteria:**
- [x] Тесты красные на текущей вёрстке

**Verification:**
- [x] `cd frontend && npm run test -- --run src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`

**Dependencies:** None. Можно писать параллельно с Task 1.

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx`
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`

**Estimated scope:** S

## Task 5: GREEN — убрать лишние поля

**Description:** В шаге итогов оставить `TripCountField` и notice только для плит. В drawer убрать оба блока `tripEditors` / поля рейсов и `TripCountNotices`. Плашки «Рейсов свай» и «Рейсов ФБС» не удалять. Подпись доставки собрать из рейсов итога.

**Acceptance criteria:**
- [x] Тесты Task 4 зелёные
- [x] `npm run typecheck` без ошибок

**Verification:**
- [x] Команды из Task 4 и `cd frontend && npm run typecheck`
- [x] `pytest tests/test_commercial_export_mixed.py -q` — рейсы в файлах совпадают с суммой

**Dependencies:** Task 4. Для скобок нужен ответ Task 2, но вёрстку прятать можно раньше.

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.tsx`
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`
- тесты этих двух файлов

**Estimated scope:** M

## Out of scope

- Тариф `pile_logistics_cost` и привязка поля стоимости рейса.
- Вес свай на карточке «Общий вес груза».
- Удаление старых ключей из `trip_counts_json`.
- Длинные сваи и их тариф по длине.
- Вопрос «сколько машин» и `pile_trip_overrides`.
