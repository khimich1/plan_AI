# Acceptance: GSM lookahead-генератор, Palisade май 2026

**Date:** 2026-08-15  
**Orchestration:** `orch-2026-08-15-gsm-geo-lookahead`  
**Spec:** [`../../specs/gsm-geo-lookahead-generator.md`](../../specs/gsm-geo-lookahead-generator.md)

Прогон через `GsmGenerationService.generate()` на копии `plita.db` (`/tmp/plita-gsm-accept.db`), без записи в рабочую БД.

## Вход

| Параметр | Значение |
|---|---|
| Машина | Hyundai Palisade (`vehicle_id=1`) |
| Период | 2026-05-04 … 2026-05-31 |
| Старт | 28 л / одометр 128327 |
| `max_daily_km` | 700 (дефолт) |
| `force` | true (на копии) |

## Результат

| Критерий | Факт | Статус |
|---|---|---|
| Без 422 (сервис вернул результат) | 200-эквивалент, период сохранён | PASS |
| ≤3 `manual_intervention` | **2** (08.05, 21.05) | PASS |
| Дней создано | 13 | — |
| Round-trip 2 плеча | все дни | PASS |
| Lookahead на плотных якорях 04–06 | 04–05: 410 км (Ковров), `balance_route` | PASS |
| 20.05 Ярославль → 21.05 Москва | 21.05: 630 км через Королёв (МО) | PASS (правдоподобно) |

### `problematic_days`

1. **2026-05-08** — заправка Опти (Городец / НН) 59.5 л при остатке 20.3; свободных будней нет, max round-trip не освобождает бак. Draft + `manual_intervention`.
2. **2026-05-21** — Реутов 54.6 л, `free_weekdays=0` после 20.05 Ярославль. День всё же собран: 630 км Королёв (направление к Москве), баланс помечен нарушенным (`fuel_end=2.9`).

Группа 20–22.05 допустима спекой (≤3 manual). 22.05 Кострома собрался от фактического остатка 21.05.

### Маршруты (кратко)

- 04–05.05: удлинённый Ковров 205×2 — lookahead.
- 06.05: Ростов 135×2.
- 08.05: Ярославль 95×2, manual.
- 20.05: Ярославль 95×2 (заправка в Ярославле).
- 21.05: Королёв 315×2 + `balance_route` + `manual_intervention`.
- 22.05: Иваново 110×2.

## Регрессия

| Сьюит | Результат |
|---|---|
| `tests/test_gsm_*.py` + geocode/link scripts | все зелёные |
| `cd frontend && npm test -- --run src/features/gsm/` | **54 passed** (15 files) |
| `venv/bin/pytest tests/ -q` | **1906 passed**, 8 skipped, **9 failed** вне ГСМ |

Внешние падения (не этот срез): `test_admin_service` (reset KP), `test_commercial_web_flow` (3), `test_kp_plates_resolve`, `test_ocr_gigachat_provider`, `test_plate_audit`, `test_recognition_pipeline` (2). OCR/KP/commercial — не файлы T1–T11.

## SC-G1…SC-G6

| ID | Статус |
|---|---|
| SC-G1 Гео (76 станций, контрольные пары) | PASS (T1+T2) |
| SC-G2 typical_station_ids, Palisade × id=1 | PASS (450/610, 116 маршрутов Palisade с id=1) |
| SC-G3 Lookahead плотные якоря | PASS |
| SC-G4 Дальний round-trip 20–22.05 | PASS (630 км МО, без ночёвки) |
| SC-G5 Частичная генерация | PASS (2 manual, период 200) |
| SC-G6 Acceptance май Palisade | PASS по автоматическим критериям; визуальная приёмка бухгалтером — отдельно |
