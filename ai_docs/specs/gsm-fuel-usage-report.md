# Spec: ГСМ — отчёт об использовании ГСМ за период (этап 2: бланк + zip)

Дата: 2026-08-25. Статус: approved for implementation.
Идея: [`../ideas/gsm-fuel-usage-report.md`](../ideas/gsm-fuel-usage-report.md)
(источник истины по домену). План:
[`../develop/plans/2026-08-25-gsm-usage-report-orchestrate.md`](../develop/plans/2026-08-25-gsm-usage-report-orchestrate.md).
Этап 0 закрыт. Этап 1 (актуализация БД) — на паузе. Этапы 3–4 — отдельные
спеки; в этом срезе **не** реализовывать.

## ASSUMPTIONS

1. **Схема БД не меняется.** `gsm_manual_fuel` — этап 3; в этапе 2 не создавать.
2. **Роль:** `REQUIRE_ACCOUNTING` на `POST /gsm/report/usage`.
3. **Источник строк:** только ПЛ со статусом `confirmed` (и при необходимости
   уже `exported` как confirmed-цепочка) в `[period_from, period_to]`.
   Draft не входят в отчёт.
4. **1 строка отчёта = 1 ПЛ.** Склейка многодневных — этап 3.
5. **«По факту» ≡ «по норме»** в этапе 2 (копия значения).
6. **«Получено»** = сумма транзакций топливных карт машины, привязанных к
   ближайшей строке (±1 календарный день). Ручные чеки — этап 3.
7. **Пробег** = нормативный из ПЛ (`km` / сумма плеч), не телематика.
8. **Цепочка остатков/одометра:** от последнего confirmed/exported якоря
   **до** `period_from` (как в генераторе/экспорте) через
   `core.gsm.balance.apply_day_chain` / значения самих ПЛ; эталон —
   май 2026, Tugella О848ХР44.
9. **Сезонная норма** — `core.gsm.season.norm_for` + журнал
   `season_switches` из настроек.
10. **Конвейер xlsx→soffice→xls→zip** — переиспользовать паттерны
    `GsmExportService` (`run_soffice`, `convert_with_soffice`, изолированный
    LO-профиль). В тестах soffice мокировать как в `tests/test_gsm_export.py`.
11. **Не трогать** рабочую `plita.db` руками; acceptance — фикстуры/тестовая БД.
12. **Не коммитить.** Существующие незакоммиченные правки ГСМ не откатывать.

## Objective

Бухгалтер нажимает «Отчёт за период», выбирает диапазон (и опционально
машины) и скачивает zip:

- на каждую машину: `Отчет по использованию ГСМ <госномер>.xls` по бланку
  образца (лист «Образец» из `ГСМ/Отчет по использованию ГСМ 2024.xls`);
- плюс путевые листы за тот же период (как в существующем
  `POST /gsm/waybills/export`).

**Критерий приёмки (acceptance):** отчёт за 01.05.2026–31.05.2026 по
Geely Tugella О 848 ХР 44 совпадает с блоком «май 2026» их файла:

| Показатель | Значение |
|---|---|
| Остаток на начало | 16.25 |
| Σ расход по норме = Σ по факту | 339.34 |
| Σ получено | 335.60 |
| Остаток на конец | 12.51 |
| Одометр | 71 514 → 75 124 |
| Число строк | 10 |
| Даты строк | 04, 07, 12, 14, 19, 20, 22, 26, 28, 29 мая |

## API Contract

```
POST /api/v1/gsm/report/usage
Auth: REQUIRE_ACCOUNTING
Body (JSON, extra=forbid):
{
  "period_from": "YYYY-MM-DD",   # inclusive
  "period_to":   "YYYY-MM-DD",   # inclusive
  "vehicle_ids": [int, ...] | null   # null = все активные машины
}
→ 200 application/zip
  Content-Disposition: attachment; filename="gsm_usage_report_YYYY-MM-DD_YYYY-MM-DD.zip"
→ 4xx {"detail": {"code": "<machine_code>", "message": "..."}}
```

### Коды ошибок (минимальный набор)

| code | HTTP | Когда |
|---|---|---|
| `gsm_report_invalid_period` | 400 | `period_from > period_to` или невалидные даты |
| `gsm_vehicle_not_found` | 404 | указан несуществующий `vehicle_id` (как в других gsm) |
| `gsm_report_no_data` | 404/422 | нет confirmed ПЛ ни по одной выбранной машине за период |
| `gsm_export_soffice_*` | 500 | сбой конвертации (те же коды, что у экспорта ПЛ) |

Zip-содержимое:

1. `Отчет по использованию ГСМ <госномер>.xls` — по одной на машину с данными.
2. Путевые листы `.xls` за период — тем же именованием/набором, что
   `GsmExportService.export_zip` для тех же `vehicle_ids` + периода
   (можно вызвать существующий сервис и смержить байты в один zip, либо
   общий helper сборки zip).

## Domain rules (бланк)

### Шапка / подвал (из шаблона)

- «УТВЕРЖДАЮ / Директор ООО "ЖБК СТАРТ" ____ Шишов А.В.»
- Дата утверждения = **последний день периода** (`period_to`).
- Заголовок блока: «за *месяц год*» (рус. месяц).
- Подвал: «Бухгалтер ____ Никифорова Е.А.»
- Подписанты в этапе 2 зашиты в шаблон (настройки — этап 3 preview).

### Строка таблицы «1. Расход топлива»

| Колонка | Источник |
|---|---|
| № п/п | порядковый в блоке месяца |
| Марка а/м | из справочника машины |
| Госномер | plate |
| ФИО водителя | из ПЛ |
| Марка бензина | из настроек/машины (как в образце) |
| Остаток на начало | `fuel_start` ПЛ |
| Спидометр нач/кон | `odometer_start` / `odometer_end` |
| Пробег | нормативный км ПЛ |
| Норма л/100км | `norm_for(day, …)` |
| Расход по норме | burn по норме |
| Расход по факту | = по норме |
| Получено по смарт-карте | сумма tx, привязанных к строке |
| Остаток на конец | `fuel_end` ПЛ |
| Примечание | дата вида «04 мая» |
| Назначение | пункт назначения маршрута ПЛ |

После строк месяца: **ИТОГО** + дубль-строка по марке бензина (как в образце).

### Блоки месяцев

Если период пересекает несколько календарных месяцев — блоки идут **вниз
по одному листу** (год = имя листа по образцу, либо один лист периода —
копия «Образец» с несколькими блоками вниз). Не плодить лишние листы сверх
логики образца.

### Привязка транзакций

Для каждой tx с `qty_liters` по картам машины в окне периода ±1 день:
назначить ближайшей строке-ПЛ по дате (при равной дистанции — предпочтение
той же дате, иначе более ранней строке). Tx без подходящей строки в
±1 день — в этапе 2 **не** попадают в «Получено» (валидация покрытия —
этап 4).

## Project structure (новые/затрагиваемые файлы)

```
core/gsm/templates/gsm_usage_report.xlsx   # 2.1 из листа «Образец»
core/gsm/usage_report.py                   # optional pure fill helpers
app/services/gsm_report_service.py         # 2.2–2.5 сборка + zip
app/schemas/gsm.py                         # UsageReportRequest
app/api/v1/endpoints/gsm.py                # POST /report/usage
app/dependencies/services.py               # DI
tests/test_gsm_usage_report.py             # 2.8 acceptance + API
```

Frontend (этап 2.7 — параллельный работник, контракт выше):

```
frontend/src/features/gsm/api/gsmApi.ts
frontend/src/features/gsm/hooks/useGsmQueries.ts
frontend/src/features/gsm/components/...   # кнопка + диалог диапазона
frontend/src/pages/gsm/GsmPage.tsx         # точка входа при необходимости
```

## Commands

```bash
.venv/bin/python -m pytest tests/test_gsm_usage_report.py -q
.venv/bin/python -m pytest tests/ -k gsm -q
cd frontend && npx vitest run src/features/gsm
cd frontend && npx tsc --noEmit
```

## Testing strategy

1. **TDD:** красный acceptance-тест на числах мая 848 → реализация.
2. Фикстура: минимальный набор confirmed ПЛ + транзакций для 848 мая 2026
   (или загрузка из тестовой БД), **не** запись в prod `plita.db`.
3. Soffice: monkeypatch `convert_with_soffice` / `run_soffice` как в
   `test_gsm_export.py` — проверять состав zip и числа в xlsx **до**
   конвертации или в подменённом xls.
4. API-тест: 200 zip + имя файла; 400 на инвертированный период.

## Boundaries

**Always:** TDD; только файлы из плана/этой спеки; мок soffice в тестах.
**Ask first:** изменение схемы БД, правки солвера `core/gsm/generator.py`.
**Never:** импорт бумажных ПЛ / правка конфликтов Palisade/Monjaro;
создание `gsm_manual_fuel`; коммит/push; остановка `run+logs.sh`;
ручной экспорт на рабочей `plita.db`.

## Out of scope (этап 2)

- Предпросмотр JSON, ручные чеки, move-fuel, merge-days → этап 3.
- Валидация цепочки / покрытие tx / блокировка zip → этап 4.
- Актуализация БД (этап 1).
