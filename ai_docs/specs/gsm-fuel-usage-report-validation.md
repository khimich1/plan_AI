# Spec: ГСМ — валидация отчёта об использовании (этап 4)

Дата: 2026-08-25. Статус: draft (реализация **после** этапа 3).
Идея: [`../ideas/gsm-fuel-usage-report.md`](../ideas/gsm-fuel-usage-report.md).
Зависит от: этап 2 ([`gsm-fuel-usage-report.md`](gsm-fuel-usage-report.md)),
этап 3 ([`gsm-fuel-usage-report-preview.md`](gsm-fuel-usage-report-preview.md)).
План: [`../develop/plans/2026-08-25-gsm-usage-report-orchestrate.md`](../develop/plans/2026-08-25-gsm-usage-report-orchestrate.md).

## ASSUMPTIONS

1. Preview (этап 3) уже показывает строки и «Получено» с ручными правками.
2. Критические ошибки **блокируют** zip; предупреждения — только в UI.
3. Сезонная норма сверяется с `core.gsm.season` + `season_switches`.

## Objective

До выгрузки бухгалтер видит список проблем; критические не дают скачать
битый отчёт. Май 848 остаётся зелёным (regression).

## Checks

| # | Проверка | Severity | Сообщение (смысл) |
|---|---|---|---|
| 1 | Стык остатков/одометра между соседними confirmed ПЛ (и с якорем до периода) | critical | разрыв на дате D: ожидалось X, в ПЛ Y |
| 2 | `0 ≤ fuel_start/fuel_end ≤ tank_volume` | critical | бак вне коридора на дате D |
| 3 | Все транзакции периода (карты машины) распределены по строкам | critical | tx на дате T не привязана |
| 4 | Норма строки = `norm_for(day)` при текущих season_switches | warning→critical* | норма дня ≠ сезону |
| 5 | Дата утверждения ≥ `period_to` | warning | дата утверждения раньше конца периода |
| 6 | Draft ПЛ в периоде при запросе отчёта | warning | есть незакрытые черновики |

\* Норма↔сезон: в MVP этапа 4 — **critical**, если расхождение > ε (0.01).

## API / UX

- `GET .../preview` (этап 3) расширяется полем:

```json
"validation": {
  "ok": false,
  "issues": [
    {"severity": "critical", "code": "gsm_chain_break", "day": "2026-05-12",
     "message": "...", "vehicle_id": 1}
  ]
}
```

- `POST /gsm/report/usage` при наличии critical → **409**
  `{"detail": {"code": "gsm_report_validation_failed", "message": "...",
    "issues": [...]}}` (или тот же паттерн `detail` что у gsm).
- Preview UI: список ошибок; кнопка zip disabled при critical.

## Testing

1. Май 848 — validation ok, zip 200 (regression этапа 2).
2. Разрыв цепочки (подмена `fuel_start` соседнего дня в фикстуре) →
   critical + блокировка zip с указанием дня.
3. «Потерянная» tx в периоде → critical «не распределена».

## Boundaries

**Always:** не ослаблять acceptance мая 848.
**Never:** молча «чинять» данные в БД из валидатора; только report + block.

## Out of scope

- Авто-рассылка, реестр версий отчётов, аналитика перерасхода (Not Doing
  в идее).
- Этап 1 (актуализация БД).
