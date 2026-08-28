# Spec: ГСМ — один месяц, один комплект

Дата: 2026-08-26. Статус: draft, план готов.
Идея: [`../ideas/gsm-month-close-kit.md`](../ideas/gsm-month-close-kit.md)
(2026-08-26, direction confirmed).
План: [`../develop/plans/2026-08-26-gsm-month-close-kit.md`](../develop/plans/2026-08-26-gsm-month-close-kit.md).
Предшественники: [`gsm-fleet-overview-ux.md`](gsm-fleet-overview-ux.md),
[`gsm-fuel-usage-report.md`](gsm-fuel-usage-report.md)
(assumption 3 «draft не в отчёте» **снимается** этим срезом).

## ASSUMPTIONS I'M MAKING

1. **Схема БД не меняется.** Статусы ПЛ по-прежнему `draft` / `confirmed` /
   `exported`. Новых таблиц нет.
2. **Роли без изменений.** `REQUIRE_ACCOUNTING` на overview / generate /
   `POST /gsm/report/usage`.
3. **Солвер (`core/gsm/generator.py`) и `_resolve_start` не трогаем.**
   Старт периода — последний `confirmed`/`exported` до `period_from`.
   Август **этой машины** не генерируют, пока её июльский хвост открыт;
   сосед без хвоста не блокируется.
4. **`POST /gsm/waybills/export` остаётся в API.** С экрана обзора кнопку
   «Экспорт zip выбранных» убираем. Комплект идёт только через
   `POST /gsm/report/usage`.
5. **Один набор дней.** Строки сводки = бланки ПЛ = дни, которым после
   успеха ставят `exported`. В набор входят `draft`, `confirmed`,
   `exported`. Красные (`manual_intervention`) в набор не входят.
6. **Отчёт без галочек `vehicle_ids: null`** — все активные машины
   (как сейчас). Машины с хвостом/разрывом цепи в текущем периоде
   исключаются из zip с текстом. Повторно «Выгружено» — confirm.
7. **Генерация только по галочкам.** «Сгенерировать выбранные» при нуле
   галок **ничего не делает** (кнопка disabled). В bulk — только
   отмеченные id. Строка / журнал «Сгенерировать» = эта одна машина.
   Не генерировать флот, если никто не выбран.
8. **Гейт красных — и фронт, и бэкенд.** Фронт не шлёт красные id и
   показывает даты. Бэкенд красную машину не кладёт в zip и **не**
   переводит её дни в `exported` (защита прямого API).
9. **Жёлтые warnings** (`borrowed_route`, `balance_route`,
   `hook_above_threshold`, `weekend_anchor`) — confirm, не стоп.
10. **Хвост `open_before`** — как сейчас: count `draft`+`confirmed` с
    `date < period_from`. Подпись — месяц **последней** такой даты
    (ближайший хвост к текущему периоду), не «до периода».
11. **Разрыв цепи** — сравнение последней ПЛ до периода (любой статус,
    max date) с первой ПЛ в периоде: `|Δл| > 0.01` или одометры не равны.
    Не `_rechain` старых км: только generate («Пересчитать»).
12. **Стрелки журнала** переписывают «С/По» на границы календарного
    месяца. Произвольный диапазон в полях дат — исключение; следующая
    стрелка снова схлопывает в месяц.
13. **Не коммитить / не писать live `plita.db`** в этом срезе до явной
    просьбы. Приёмка — vitest + pytest на фикстурах; live-проверка
    глазами после кода.

→ Поправьте сейчас, иначе иду с этим в план/задачи.

## Objective

Бухгалтер (`accountant`) и тот, кто перегенерирует месяц, работают с
**одним** экраном `/gsm` → Обзор:

1. Видят ровно один рабочий месяц (таблица = журнал).
2. Понимают хвост: «Июль не выгружен: 6 ПЛ», а не «до периода».
3. Закрывают месяц одной кнопкой «Отчёт за период» (сводка + путевые),
   черновики входят в сводку и становятся `exported`.
4. Не могут сгенерировать или выгрузить **текущий** месяц **у конкретной
   машины**, пока у неё открыт хвост или бак не стыкуется с концом
   предыдущего. Соседей без хвоста это не блокирует. Bulk-генерация —
   только с галочками.

**Пользователь:** `accountant`, `admin`. Не цель: папка 1С, автогенерация
следующего месяца, смена колонок бланка сводки.

## Поток закрытия

```
открыт август, у Monjaro open_before > 0, Palisade чистая
  → bulk generate: 0 галок → noop; галка Palisade → generate Palisade;
    галка Monjaro → skip + «сначала июль»
  → «Отчёт за период» без галок: Palisade в zip, Monjaro исключена
    с текстом «сначала выгрузите июль»
  → баннер / «Экспорт» в строке Monjaro: на июль + комплект июля этой машины

комплект июля Monjaro (нет красных)
  → zip: сводка + бланки; дни июля exported

если август Monjaro есть и цепь порвана
  → комплект/generate августа этой машины нельзя, CTA «Пересчитать август»
  → generate(август) от последнего exported июля (нужна галка или кнопка в строке)
Palisade не ждёт Monjaro
```

## UX

### Период

- Дефолт — текущий календарный месяц (`currentMonthBounds`), как сейчас.
- Стрелки `VehicleMonthCalendar` вызывают `onMonthChange`, родитель
  ставит `periodFrom`/`periodTo` = `monthBounds(новый YYYY-MM)`.
- `VehicleWaybillJournal` **не** держит независимый `month`: сетка =
  `periodFrom.slice(0, 7)` (если «С» не 1-е число — сетка месяца «С»,
  дни вне `[from, to]` не кликабельны / пустые слоты как сейчас за
  границами).
- Баннер хвоста: `periodFrom/To` = границы месяца
  `open_before_month` (см. API).

### Копирайт

| Было | Стало |
|---|---|
| «до периода незакрыто N ПЛ» | «{месяц} не выгружен: N ПЛ» (род. падеж: «Июль не выгружен») |
| Баннер: «До периода незакрыто N ПЛ по M машинам. Открыть предыдущий месяц.» | «{месяц} не выгружен: N ПЛ по M машинам. Открыть {месяц}.» |
| «Экспорт zip выбранных» | убрать |
| «Отчёт за период» | оставить; под кнопкой: «сводка и путевые» |
| Синее «Экспорт» в строке (`<span>`) | кнопка: комплект хвоста этой машины, если `open_before > 0`; иначе комплект **текущего** периода этой машины (те же гейты) |

### Генерация (галочки)

- **«Сгенерировать выбранные»** disabled, если `selectedIds.size === 0`.
  Текст «Выбрано: 0». Никакого POST.
- В запросе только id с галкой. `no_data` по-прежнему не выбирается
  select-all (как сейчас).
- По каждой отмеченной машине: если `open_before > 0` или `chain_broken`,
  **и** текущий период не является месяцем её хвоста → эту машину не
  генерировать, в отчёте bulk: «{имя}: сначала выгрузите {месяц}» /
  «пересчитайте {месяц}». Остальные отмеченные — генерировать.
- Если период = месяц хвоста этой машины — generate разрешён (закрываем
  / пересобираем хвост; force в диалоге как сейчас).
- «Сгенерировать» в строке / журнале — только эта машина, те же гейты
  (не требует галки; это явное действие по строке).

### Hard stop комплекта (по машине, не по флоту)

«Отчёт за период» **не** гасится из‑за чужого хвоста.

В комплект не входят машины с `open_before > 0` или `chain_broken`,
**кроме** случая: текущий период = месяц её хвоста (тогда как раз
закрываем хвост). Исключённые — alert с причиной. Чистые соседи едут.

Строка «Экспорт» / «Пересчитать»: только эта машина, те же правила.

Подсказки: «Сначала выгрузите {месяц}» / «Пересчитайте {месяц}: бак не
сходится с предыдущим».

### Красные / жёлтые (комплект)

Как `planBulkExport` + `exportConfirmMessages`, но цель — отчёт:

- `red_days > 0` → машина не в запросе, alert «{имя}: исправьте дни …».
- Только жёлтые у входящих → confirm «есть предупреждения».
- `exported_count == wb_count && wb_count > 0` у входящих → confirm
  «уже выгружалось. Скачать снова?».
- Все выбранные красные / некого слать → не POST, только alert.

### «Пересчитать {месяц}»

Видна у строки / в шапке журнала, если `chain_broken`. Открывает
существующий `VehicleGenerateDialog` на текущий `periodFrom`/`periodTo`
(force по умолчанию **false**; если в периоде есть exported — диалог как
сейчас потребует «Перезаписать»).

Не вызывать `_rechain_downstream` вместо generate.

## API

### `GET /gsm/overview`

Аддитивно к `FleetOverviewRow`:

```
open_before_month: str | null   # "YYYY-MM" max(date) среди
                                # draft|confirmed с date < from; null если 0
chain_broken: bool              # см. правило ниже
```

`open_before` не меняет смысл.

**`chain_broken`:** взять последнюю ПЛ машины с `date < period_from`
(любой статус, `ORDER BY date DESC, id DESC`) и первую с
`date >= period_from AND date <= period_to` (тот же порядок ASC).
Нет одной из двух → `false`. Иначе `true`, если
`abs(prev.fuel_end - first.fuel_start) > 0.01` **или**
`prev.odometer_end != first.odometer_start`.

Баннер флота: `openBeforeSummary` как сейчас + месяц =
max(`open_before_month`) по строкам с `open_before > 0` (ближайший хвост).

### `POST /gsm/report/usage`

Контракт тела/роль **без изменений**. Поведение:

1. Резолв машин: `null` → все активные; иначе переданный список.
2. Выкинуть машины с ≥1 ПЛ в периоде, у которых `warnings` содержит
   `manual_intervention`. Не 422 на весь запрос.
3. По оставшимся: ПЛ периода со статусом в
   `{draft, confirmed, exported}`. Пустые после фильтра — пропуск машины.
4. Если после 2–3 никого нет → `gsm_report_no_data` (как сейчас, но
   формулировка: нет ПЛ к комплекту / все красные).
5. Сводка: те же хелперы `usage_report.py`, источник строк = набор из п.3
   (не `CONFIRMED_STATUSES`).
6. Бланки: существующий `GsmExportService.export_zip` **только по id
   машин из п.3** (он уже берёт все статусы и ставит `exported`).
   Не вызывать export по красным.
7. Zip как сейчас: файлы сводки + файлы ПЛ.

Эталон мая 848: дни уже `exported` → набор совпадает со старым, цифры
блока «май 2026» не должны разъехаться.

Новый кейс: только `draft`, без красных → сводка с N строками = N бланков,
после ответа дни `exported`.

## Project Structure

```
app/schemas/gsm.py                      → FleetOverviewRow += поля
app/repositories/gsm_repository.py      → open_before_month; данные для chain
app/services/gsm_overview_service.py    → chain_broken
app/services/gsm_report_service.py      → набор дней = draft|confirmed|exported;
                                          skip red vehicles
core/gsm/usage_report.py                → CONFIRMED_STATUSES расширить
                                          или KIT_STATUSES
frontend/.../FleetOverviewView.tsx      → одна кнопка; хинт; hard stop; баннер
frontend/.../FleetOverviewTable.tsx     → копирайт хвоста; кнопка Экспорт
frontend/.../VehicleWaybillJournal.tsx  → месяц = период родителя
frontend/.../VehicleMonthCalendar.tsx   → onMonthChange поднимает период
frontend/.../lib/fleetStatus.ts         → подпись месяца; previousMonth →
                                          open_before_month
frontend/.../lib/exportGate.ts          → гейт комплекта (red / yellow /
                                          already exported); chain_broken
tests/test_gsm_overview_api.py
tests/test_gsm_usage_report.py
frontend/.../FleetOverviewView.test.tsx
frontend/.../VehicleWaybillJournal.test.tsx
frontend/.../exportGate.test.ts
```

## Code Style

Имена: `open_before_month`, `chain_broken`, `KIT_STATUSES`. Подпись
месяца — `toLocaleDateString("ru-RU", { month: "long" })` с заглавной
для «Июль не выгружен». Порог литров — тот же `litersDiffOk` (0.01).

Не вводить новый HTTP для комплекта. Не менять шаблон xlsx сводки.

## Testing Strategy

- **pytest:** overview отдаёт `open_before_month` / `chain_broken` на
  фикстуре (июль draft, август с другим fuel_start → true; стык → false;
  нет ПЛ в периоде → false).
- **pytest:** `report/usage` с одними draft — в ответе zip, в БД
  `exported`; машина с `manual_intervention` не exported и не в сводке;
  соседняя чистая — в сводке. Регрессия мая 848 (цифры блока).
- **vitest:** журнал не уезжает на другой месяц без смены `periodFrom`;
  стрелка меняет фильтры; баннер текст с месяцем; bulk generate при 0
  галок не вызывается; на августе отчёт без галок исключает машину с
  july-хвостом и оставляет чистую; `chain_broken` показывает
  «Пересчитать»; синее Экспорт — `button`.
- **Не** гонять generate на live `plita.db` в задачах этого среза.

## Boundaries

- **Always:** гейт красных на бэкенде отчёта; один набор дней; стрелки
  синхронизируют период; тесты overview + report + vitest журнала.
- **Ask first:** удаление `POST /waybills/export`; смена `_resolve_start`
  на draft; автоgenerate после комплекта; писать live БД.
- **Never:** зажимать `fuel_end` в 0; `_rechain` вместо generate при
  разрыве; молча класть красные дни в `exported`; коммитить без просьбы.

## Success Criteria

1. Обзор августа + 6 july draft Monjaro, Palisade без хвоста: баннер
   «Июль не выгружен: …»; журнал = август. Bulk generate без галок —
   noop. Галка только Palisade — generate Palisade. Галка Monjaro —
   skip «сначала июль». Отчёт без галок — Palisade в zip, Monjaro в
   alert, не в zip.
2. Стрелка «предыдущий месяц» ставит 01.07–31.07 в фильтрах и сетку июля.
3. Комплект июля по Monjaro без красных: N строк сводки = N бланков = N
   draft; после запроса статус `exported`; бейдж хвоста на августе сходит
   (если 952 тоже закрыт / не выбран — по факту выбранных).
4. Красный день на машине: её нет в zip, alert с датой, статус остаётся
   draft.
5. Monjaro цепь порвана: на августе `chain_broken`, комплект disabled,
   «Пересчитать август» открывает generate. После generate стык →
   `chain_broken false`.
6. Май 848 usage-report acceptance зелёный.
7. Кнопки «Экспорт zip выбранных» нет. «Экспорт» в строке — button.

## Open Questions

Нет продуктовых. Ниты реализации (манифест исключений в zip vs только
UI alert) — в плане, не блокер спеки: UI alert обязателен, бэкенд skip
обязателен.
