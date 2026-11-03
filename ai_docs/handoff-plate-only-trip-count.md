# Переход: рейсы только у плит

Ветка `aleksey_web`. Последний коммит `c19af39` — `feat(kp): ручное число рейсов в сумме и в документах`. Пуш не делался. Всё ниже **не закоммичено**.

## Что уже решено

Ручное число рейсов остаётся только у плит и только в конструкторе, шаг «Итоги и скидка» (туда попадают через «Редактировать»).

- Пока вес плит тот же, вписанное число живёт. Вес изменился — поле становится формулой `ceil(кг / 18 600)`, на шаге висит «рейсы плит пересчитаны: было N, стало M», пока не закроют.
- Сумма, PDF и XLSX по плитам берут это число.
- Сваи, ФБС/ЛС/ЛМ и длинные сваи снова считаются формулой. Ключи `pile` и `fbs_lm` в `trip_counts_json` на сумму не влияют. Сами ключи из JSON не удаляли.
- Вопрос «сколько машин» (`pile_trip_overrides`) у сваи без нормы не трогали.
- Архив: нет поля «Кол. рейсов» и нет плашки. Строка «Услуга по доставке» остаётся. В скобках рейсы суммы: сохранённые плиты + формула свай + формула ФБС. Длинные сваи в скобки не входят.
- Открытие карточки архива не сбрасывает рейсы плит и не пишет `trip_counts_json`. Пример: сохранены 57, формула по весу 49 — после GET остаётся 57.

## Что в коде

- `core/trip_counts.py` — сваи и ФБС всегда берут формулу. Сброс плит по массе только при `reset_plate_on_mass_change=True` (так вызывает конструктор).
- `app/services/archive_service.py` — просмотр, файлы архива и пересчёт доставки идут с `reset_plate_on_mass_change=False`. `save_trip_counts_json` при GET убран.
- `CalculationResultStep.tsx` — поле и плашка только у плит.
- `OfferDetailsDrawer.tsx` — поля и плашки нет. Архив больше не шлёт `pile_trips` / `fbs_lm_trips`.

## Проверено

- `pytest tests/test_commercial_logistics_cost.py tests/test_commercial_fbs_lm_delivery.py tests/test_pile_trip_pricing.py tests/test_archive_endpoints.py tests/test_commercial_export_mixed.py -q` — 197 passed
- vitest `CalculationResultStep.test.tsx` и `OfferDetailsDrawer.test.tsx` — 111 passed
- `cd frontend && npm run typecheck` — чисто

Живой экран после этой правки заново не кликали.

## Дыра

Сохранение скидки или тарифа из архива всё ещё считает с флагом по умолчанию (`True`) в `core/kp/offers_write.py` (`calculate_total_cost` около строк 289 и 426) и может переписать рейсы плит, если вес уже не совпадает с сохранённым базисом. Само открытие карточки этого не делает.

КП, где у свай или ФБС уже было вписано своё число, пересчитаются по формуле, итог изменится.

## Не делать заново

- Не возвращать ручное число свай и ФБС.
- Не показывать «Кол. рейсов» и плашку в архиве.
- Не прятать строку доставки.
- Не чинить в этой задаче привязку тарифа свай (`pile_logistics_cost` vs `logistics_cost`). Этот баг откатили вместе с несохранённой правкой. Спека `ai_docs/specs/pile-trip-cost-roundtrip.md` удалена `git clean`.
- Не менять 18 600 / 20 000, `pile_catalog`, вес свай на карточке «Общий вес груза».

## Файлы

Неотслеживаемые:

- `ai_docs/ideas/plate-only-trip-count.md`
- `ai_docs/specs/plate-only-trip-count.md`
- `ai_docs/develop/plans/2026-10-30-plate-only-trip-count.md`

Изменённые, не в коммите: `core/trip_counts.py`, `core/commercial_pricing.py`, `core/commercial_offer.py`, `core/commercial_offer_xlsx.py`, `app/services/archive_service.py`, `CalculationResultStep.tsx` и его тест, `OfferDetailsDrawer.tsx` и его тест, `archiveApi.ts`, `useArchiveQueries.ts`, `tests/test_commercial_logistics_cost.py`, `tests/test_commercial_fbs_lm_delivery.py`, `tests/test_archive_endpoints.py`.

Черновики незанесённых КП: `.app_data/drafts/{uuid}.json`. В базу `KP_offers` они попадают только при сохранении.
