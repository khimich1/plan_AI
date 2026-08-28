# Отчёт: GSM — коридор бака и честные литры моек

**Дата:** 2026-08-24  
**Спека:** [`../../specs/gsm-anchor-corridor-wash-qty.md`](../../specs/gsm-anchor-corridor-wash-qty.md)  
**План:** [`../plans/2026-08-24-gsm-anchor-corridor-wash-qty.md`](../plans/2026-08-24-gsm-anchor-corridor-wash-qty.md)  
**Статус:** T1–T4 выполнены. Checkpoint GSM-сьюита зелёный. Полный `tests/` — 11 падений вне GSM (не регрессия этого среза). Коммитов нет.

## Что сделано

Системные причины красного дня Palisade 04.08 и ложного Δ литров в обзоре:

1. **Коридор бака при выборе якоря** (`core/gsm/generator.py`): перед базовым выбором отбрасываются маршруты, у которых `fuel_end` вне `[0…tank]`. Мойка (`q_today == 0`) берёт минимальный km из коридора; топливо — max frequency; пустой коридор → fallback на самый короткий из группы. Lookahead сверху может удлинить.
2. **Σ литров обзора — только fuel** (`fleet_overview`): `SUM(CASE WHEN t.service_type = 'fuel' THEN t.qty_liters ELSE 0 END)`. `liters_diff` = топливо транзакций − `fuel_issued`. Журнал транзакций не трогали.
3. **Импорт моек:** при `service_type == "wash"` сервис пишет `qty_liters=None`, даже если в файле число. Миграции старых строк нет.
4. **Приёмка** августа Palisade на копии `/tmp/plita_accept.db` через `GsmGenerationService` (не через live `:8000`).

Фронтенд не трогали. Рабочая `plita.db` write-операциями не писалась.

## Задачи

| Task | Содержание | Статус |
|------|------------|--------|
| T1 | Коридор бака + мойка выбирает min km; lookahead может перебить | ✅ |
| T2 | `tx_liters` в обзоре только `service_type='fuel'` | ✅ |
| T3 | Импорт wash → `qty_liters=None` | ✅ |
| Checkpoint | `tests/test_gsm_*.py` зелёный; полный `tests/` | ⚠️ GSM ✅ / полный сьюит 11 вне GSM |
| T4 | Генерация августа Palisade на копии БД | ✅ |

## Success Criteria спеки

| # | Критерий | Доказательство | Вердикт |
|---|----------|----------------|---------|
| 1 | Генерация августа Palisade на копии: 04.08 — обычный draft 12 км, без `manual_intervention` | Копия, `GsmGenerationService.generate(vehicle_id=1, 2026-08-01…31, force=True)`: 04.08 `status=draft`, `km=12` (плечо 6+6, `route_id=497`), `warnings=[]`, `warning_details=[]`, `fuel_end=37.65 ≥ 0` | ✅ |
| 2 | В обзоре августа Palisade `liters_diff` = 0.0 л (топливо vs топливо) | Обзор копии: `tx_liters=194.35`, `wb_fuel_issued=194.35`, `liters_diff=0.0` | ✅ |
| 3 | Импорт `.xls` с мойкой, у которой в файле число, сохраняет `qty_liters = None` | `test_import_wash_with_numeric_qty_persists_null_liters` + сервис `qty_liters = None if row.service_type == "wash"` | ✅ |
| 4 | Все существующие тесты GSM зелёные | `venv/bin/pytest tests/test_gsm_*.py -q` → **218 passed** | ✅ |
| 5 | `red_days` в обзоре августа Palisade = 0 | Копия: Palisade `red_days=0`, `status=drafts_pending` | ✅ |

## Приёмка T4 (только копия)

**Метод:** `cp plita.db /tmp/plita_accept.db`, затем `GsmGenerationService` + `GsmRepository(db_path="/tmp/plita_accept.db")`. Live uvicorn на `:8000` (рабочая `plita.db`) не вызывался. Временный сервер не поднимался.

**Машина:** Hyundai Palisade, `vehicle_id=1`, пластина `О 521 УХ 44`.  
**Старт периода:** последний confirmed/exported до 01.08 — `2026-05-29`, `fuel_end=41.13`, `odometer_end=132517`.  
**force:** true (на копии; в августе были draft, в т.ч. старый красный 04.08).

| Поле | Было на копии до generate | После generate |
|------|---------------------------|----------------|
| 04.08 | draft, `fuel_end=-13.97`, `manual_intervention` | draft, 12 км, `fuel_end=37.65`, без warning |
| days | 7 | 10 (`days_created=10`, `manual_days=0`) |
| overview Palisade | — | `status=drafts_pending`, `red_days=0`, `liters_diff=0.0` |

04.08 маршрут: Кострома пер.Инженерный ↔ ул.Кузнецкая, 6 км плечо × 2.

**Доказательство, что рабочая БД не писалась:**

| Файл | md5 | mtime | size |
|------|-----|-------|------|
| `plita.db` до | `0a96be7d64bd2ca27f8bac3e7a4c8c8c` | 2026-08-24 18:39:01 | 4022272 |
| `plita.db` после | `0a96be7d64bd2ca27f8bac3e7a4c8c8c` | 2026-08-24 18:39:01 | 4022272 |
| `/tmp/plita_accept.db` после generate | `8892d2feca9707e46d484acdca6ff58e` | 2026-08-24 18:50:27 | 4022272 |

`cmp plita.db /tmp/plita_accept.db` → файлы различаются (запись только в копию).

## Прогоны

```
venv/bin/pytest tests/test_gsm_*.py -q     → 218 passed, 92 warnings in 40.68s
venv/bin/pytest tests/ -q                  → 11 failed, 2244 passed, 8 skipped, 3593 warnings in 276.94s
```

11 падений **вне GSM**, тот же набор, что на приёмке `gsm-fleet-overview-ux`:

- `test_admin_service.py::test_reset_kp_only_keeps_completed_plates_and_rests`
- `test_commercial_fbs_flow.py::test_bulk_grade_applies_to_all`
- `test_commercial_march_flow.py::test_update_march_draft_replace`
- `test_commercial_march_flow.py::test_update_march_grades_bulk`
- `test_commercial_step_flow.py::test_update_step_draft_replace`
- `test_commercial_web_flow.py::test_generate_files_returns_schema_when_requested`
- `test_commercial_web_flow.py::test_build_offer_identity_uses_predicted_kp_number`
- `test_commercial_web_flow.py::test_build_offer_identity_prefers_saved_kp_id`
- `test_kp_plates_resolve.py::test_order_data_from_completed_plates_omits_zero_unit_price`
- `test_ocr_gigachat_provider.py::test_require_gigachat_client_missing_credentials`
- `test_plate_audit.py::test_audit_log_records_completion_and_rejection`

Не чинились: скоуп среза — GSM T1–T3; это не их регрессия. Checkpoint-бокс полного сьюита в плане оставлен `[ ]`.

## Отклонения

1. **Мойка + lookahead: 140 км плечо вместо 132 км в спеке.** Спека в изолированном примере считает «сжечь 38 л → 264 км круг → ≥132 км плечо» от бака после короткого маршрута (58,26 л). Реальный lookahead считает `km_needed` от `fuel_before + Q_next` **до** применения короткого выжига: `40 л / 14.5 × 100 ≈ 275.86` км круга → плечо ≥ 138 км. Тест `test_wash_lookahead_replaces_short_route_when_next_refill_overflows` ставит 140 км (круг 280). Поведение спеки («короткий перебивается, если нужен выжиг») выполняется; цифра 132 км — иллюстрация, не порог кода. Спеку не меняли.
2. **Живой Palisade 04.08 `fuel_end=37.65`, а не 11,84 из unit-теста.** Unit-тест T1 фиксирует 03.08 на частый 95 км typical (станция длинного маршрута), 04.08 — мойка 6 км → 11,84 л. На копии оба дня — мойки с typical, покрывающим 6 км, поэтому 03.08 тоже короткий (12 км), старт 04.08 = 39,39 л. Критерий спеки (draft 12 км, бак ≥ 0, без `manual_intervention`) выполняется.
3. **Полный pytest не зелёный** — 11 падений вне модуля ГСМ, как на предыдущем срезе. Checkpoint полного сьюита не закрыт.
4. **На копии появилось 10 draft-дней вместо прежних 7.** `force=True` пересобрал период новым генератором (добавились 05.08, 10.08, 12.08). `manual_days=0`. Не баг.
5. **Live `:8000` / рабочая БД** по-прежнему содержат старый красный 04.08 (`fuel_end=-13.97`). Исправление видно только после генерации; на рабочей БД generate не гоняли.

## Файлы

Изменены (T1–T3):

- `core/gsm/generator.py` — `_fits_corridor`, фильтр `in_corridor`, мойка → min km
- `tests/test_gsm_generator.py` — две мойки подряд; мойка + lookahead
- `app/repositories/gsm_repository.py` — `tx_liters` только fuel
- `tests/test_gsm_overview_api.py` — `test_overview_tx_liters_excludes_wash_qty`
- `app/services/gsm_transaction_service.py` — wash → `qty_liters=None`
- `tests/test_gsm_transaction_import.py` — импорт мойки с числом в файле

Обновлены / созданы документами этого шага:

- `ai_docs/develop/plans/2026-08-24-gsm-anchor-corridor-wash-qty.md` — T4 и GSM-checkpoint `[x]`; полный pytest `[ ]`
- `ai_docs/develop/reports/2026-08-24-gsm-anchor-corridor-wash-qty.md` — этот отчёт

Фронтенд не менялся. Временная копия `/tmp/plita_accept.db` в репозиторий не входит.

## Границы

Нет миграций БД, нет соло-плеча, нет правок фронтенда, нет новых зависимостей, нет коммита/push/git config.

**Коммитов нет.**
