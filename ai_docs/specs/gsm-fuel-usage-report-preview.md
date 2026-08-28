# Spec: ГСМ — предпросмотр отчёта об использовании (этап 3)

Дата: 2026-08-25. Статус: draft (реализация **после** этапа 2).
Идея: [`../ideas/gsm-fuel-usage-report.md`](../ideas/gsm-fuel-usage-report.md).
Зависит от: [`gsm-fuel-usage-report.md`](gsm-fuel-usage-report.md) (этап 2).
План: [`../develop/plans/2026-08-25-gsm-usage-report-orchestrate.md`](../develop/plans/2026-08-25-gsm-usage-report-orchestrate.md).

## ASSUMPTIONS

1. Этап 2 уже отдаёт корректный zip по confirmed ПЛ + карточным tx.
2. Без предпросмотра колонка «Получено» не совпадёт с бухгалтерской
   практикой (чеки вне выписок, сдвиг даты ±1 день, склейка «21–22 апр»).
3. Правки должны **жить в БД**, иначе регенерация ломает следующий месяц.
4. Схема БД **меняется** только здесь: таблица `gsm_manual_fuel`.

## Objective

Перед выгрузкой бухгалтер видит таблицы по машинам/месяцам и может:

1. Добавить чек-заправку (ручные литры).
2. Перенести заправку (карточную или ручную) на соседнюю строку.
3. Склеить многодневную поездку в одну строку («21–22 апр»).
4. Поправить дату утверждения и подписантов.
5. Нажать «Сформировать zip» → тот же контракт этапа 2, но с учётом правок.

**Критерий успеха:** правки в БД → повторный preview/zip даёт тот же
результат (воспроизводимость).

## Data model

```sql
CREATE TABLE gsm_manual_fuel (
  id INTEGER PRIMARY KEY,
  vehicle_id INTEGER NOT NULL REFERENCES gsm_vehicle(id),
  fuel_date TEXT NOT NULL,          -- YYYY-MM-DD
  qty_liters REAL NOT NULL,
  comment TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT
);
-- index (vehicle_id, fuel_date)
```

Ручные заправки участвуют в колонке «Получено» и в цепочке баланса
наравне с карточными транзакциями (согласовано в идее).

## API (черновик контракта)

```
GET  /api/v1/gsm/report/usage/preview
     ?period_from=&period_to=&vehicle_ids=1,2
→ 200 JSON:
{
  "period_from": "...",
  "period_to": "...",
  "approval_date": "...",          # default = period_to
  "signatories": {"director": "Шишов А.В.", "accountant": "Никифорова Е.А."},
  "vehicles": [
    {
      "vehicle_id": 1,
      "plate": "...",
      "months": [
        {
          "year": 2026, "month": 5,
          "rows": [ /* те же поля, что в бланке + fuel_allocations */ ],
          "totals": { "norm": ..., "fact": ..., "received": ..., ... }
        }
      ]
    }
  ]
}

POST /api/v1/gsm/manual-fuel
Body: { vehicle_id, fuel_date, qty_liters, comment? }
→ 201 ManualFuelOut

POST /api/v1/gsm/report/usage/preview/move-fuel
Body: { vehicle_id, source: "card"|"manual", source_id|tx_key, to_row_date }
→ 200 updated preview fragment / 4xx

POST /api/v1/gsm/report/usage/preview/merge-days
Body: { vehicle_id, dates: ["YYYY-MM-DD", ...] }  # consecutive PL days
→ 200 merged row preview / 4xx
```

Zip после правок: существующий `POST /gsm/report/usage` читает manual_fuel
и назначения/склейки из БД (детали хранения merge — уточнить при
реализации: отдельная таблица или JSON в настройках периода; **не**
решать в этапе 2).

## UI

- Экран/drawer предпросмотра на странице ГСМ после выбора периода.
- Редактируемая колонка «Получено», кнопки «Добавить чек», «Склеить дни».
- Поля даты утверждения / ФИО подписантов.
- CTA «Сформировать zip».

## Boundaries

**Always:** правки персистентны; TDD на воспроизводимость.
**Never:** начинать этап 3 до зелёного acceptance этапа 2 (май 848);
ломать контракт zip этапа 2 без миграции клиентов.

## Out of scope

- Валидация и блокировка zip → этап 4.
- Импорт бумажных ПЛ / разрешение конфликтов → этап 1.
