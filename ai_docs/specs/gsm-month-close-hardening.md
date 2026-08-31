# Spec: ГСМ — hardening закрытия месяца

Дата: 2026-08-28. Статус: **SPECIFY approved** (решения зафиксированы с владельцем).  
Дописка к [`gsm-month-close-kit.md`](./gsm-month-close-kit.md): UX комплекта не меняем; закрываем дыры API, цифры drawer и выгрузку на общем сервере.  
Аудит: [`../develop/audits/2026-08-26-gsm-audit.md`](../develop/audits/2026-08-26-gsm-audit.md).  
Ревью 28.08: гейты хвоста/цепи только на фронте; `POST /report/usage` ставит `exported` без них.

## ASSUMPTIONS I'M MAKING (зафиксированы)

1. **Схема БД не меняется.** Статусы `draft` / `confirmed` / `exported`. Job комплекта (этап 4) — в памяти процесса.
2. **Роли без изменений.** `REQUIRE_ACCOUNTING` на generate / export / report.
3. **Солвер не трогаем** (`core/gsm/generator.py`).
4. **Прод: один uvicorn worker, SQLite, несколько ПК и ролей** (бухгалтер, менеджер, производство, логист). Не `--workers N`.
5. **`POST /gsm/waybills/export` не удаляем;** на него тот же гейт машин, что на комплект (он тоже пишет `exported`).
6. **Generate при `chain_broken` разрешён** («Пересчитать»). Generate августа при открытом июле этой машины — запрещён.
7. **Если `red_days > 0` → статус `has_red_days`**, даже при более новой заправке без ПЛ.
8. **«Экспорт» в строке** прыгает на месяц хвоста; гейты считаются **за тот период, который в zip**. Источник правды — сервер.
9. **Частичный успех комплекта (5А):** чистых закрываем (zip + `exported`); больных нет в zip и не `exported`; текст skip на экране; файла `исключения.txt` нет. Если после фильтра никого — 4xx / `gsm_report_no_data`.
10. **Норма дня с сервера** (`norm_l_per_100` на ПЛ из журнала `season_switches`, тот же `norm_for`, что бланк). Drawer не угадывает сезон по «последнему переключению». Переключатель в справочниках — источник журнала. Округление литров — как CPython `round` (banker's). Без preview-API.
11. **Битый `season_switches`:** везде `gsm_settings_invalid`, registry не глотает в `[]`.
12. **Импорт:** транзакция **на файл** (хорошие файлы остаются); внутри файла — атомарно. Не-xls / `XLRDError` → **400**. Автосоздание карт/станций не отключаем.
13. **LibreOffice (этап 4):** сначала **один** `soffice` на папку xlsx, затем **фон + опрос + один слот**. Не Celery. Zip из xlsx без `.xls` — не в этом документе (нужно отдельное «да» бухгалтерии).
14. **Implement-порядок:** сначала **только этап 1**, живая проверка, затем 2→3→4. Live `plita.db` и git commit — только по просьбе.
15. **Skip-текст при 200:** только UI (`planKit`). Сервер не шлёт список исключений в заголовке.
16. **Job этапа 4:** в памяти процесса, не файл и не таблица SQLite.
17. **Лимиты импорта:** ≤10 файлов, ≤20 000 строк суммарно, суммарный байт ≤100 МБ (плюс уже существующий cap на файл).
18. **Полный аудит 2026-08-28** (КП/IDOR и т.д.) — **не** этот документ.

→ Решения 1А, 2А, 3А, 4А+сервер, 5А, 6 норма на ПЛ, 7А, 8Б, 9А, 10 = пакетный soffice затем in-process job.

## Objective

Бухгалтер закрывает календарный месяц на `/gsm` так, что:

1. печать `exported` и состав zip **нельзя** получить в обход хвоста / разрыва цепи / красных дней (в том числе прямым API);
2. литры в карточке дня совпадают с бланком и учитывают **переключатель сезона в справочниках**;
3. импорт заправок не оставляет полфайла в БД;
4. на общем сервере комплект не держит HTTP минутами на каждый `soffice` (после этапа 4).

**Пользователь:** `accountant` / `admin` на Обзоре; менеджер и другие роли на том же backend с других ПК.

**Не цель:** нарезка `generator.py`, CRUD маршрутов, codegen типов, маскирование СНИЛС, Celery, `--workers > 1`, отказ от `.xls`.

## Locked decisions (Q&A)

| # | Решение |
|---|---------|
| Export API | Тот же гейт, что usage-report |
| Цепь vs generate | Выгрузка запрещена; generate («Пересчитать») разрешён |
| Бейдж | Красные важнее «нужна генерация» |
| Кнопка «Экспорт» | Прыжок на хвост + серверный гейт июля |
| Частичный флот | 200 + zip чистых; 4xx если закрывать некого |
| Сезон | `norm_l_per_100` с сервера из журнала кнопки справочника |
| Битый журнал | Ошибка везде |
| Импорт | Commit на файл; 400 на мусор |
| soffice | Пакетная конвертация папки, затем фон/опрос/один слот |
| Первый код | Только этап 1, затем живая проверка |
| Skip в 200 | Только UI, без заголовка API |
| Job zip | Память процесса |
| Лимиты import | 10 файлов / 20k строк / 100 МБ суммарно |
| Аудит 28.08 всего репо | Вне этой спеки |

## Tech Stack

Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`); React + Vite + TypeScript; pytest + vitest. LibreOffice `soffice` как сейчас. Без Redis/Celery в этапе 4.

## Commands

```bash
.venv/bin/python -m pytest tests/test_gsm_usage_report.py tests/test_gsm_overview_api.py tests/test_gsm_generation_api.py tests/test_gsm_season.py tests/test_gsm_balance.py tests/test_gsm_export.py -q

cd frontend && npx vitest run src/features/gsm/lib/exportGate.test.ts src/features/gsm/lib/downstreamPreview.test.ts src/features/gsm/components/FleetOverviewView.test.tsx

.venv/bin/python -m pytest tests/test_gsm_generator.py -q
```

Пакетный soffice (этап 4) — отдельный тест с моком convert «один вызов на N xlsx», не обязателен live LO в CI.

## Project Structure

```
app/services/gsm_kit_gate.py              NEW: eligibility машины за период
app/services/gsm_report_service.py        гейт; этап 4: batch convert + job
app/services/gsm_export_service.py        гейт; convert_with_soffice → папка
app/services/gsm_generation_service.py    гейт generate / bulk
app/services/gsm_overview_service.py      приоритет has_red_days
app/services/gsm_transaction_service.py   tx на файл + лимиты
app/services/gsm_registry_service.py      не глотать битый season_switches
app/repositories/gsm_repository.py        batch insert в одной tx
app/api/v1/endpoints/gsm.py               400 parse; этап 4: start/status/download
app/schemas/gsm.py                        norm на WaybillOut; коды kit; job dto
core/gsm/season.py                        источник истины журнала
core/gsm/balance.py                       эталон round
core/gsm/transactions.py                  magic / XLRDError
frontend/.../lib/exportGate.ts            handleExportKit не в обход planKit
frontend/.../lib/downstreamPreview.ts     burn от waybill.norm, round как Python
frontend/.../hooks/useGsmQueries.ts       invalidate waybills после сезона
frontend/.../types/gsm.ts
tests/test_gsm_usage_report.py
tests/test_gsm_kit_gate.py                optional
frontend/.../FleetOverviewView.test.tsx
```

`core/gsm/generator.py` и шаблоны xlsx бланка — не в diff гейтов.

## Этап 1 — серверный гейт

Правила как `planKit` / `planBulkGenerate`, период = `period_from`…`period_to`.

| Условие | Комплект (usage + прямой export) | Generate |
|---------|----------------------------------|----------|
| в периоде ПЛ с `manual_intervention` | skip машины, не `exported` | не стоп (можно пересчитать) |
| `open_before > 0` и месяц периода ≠ `open_before_month` | skip «сначала выгрузите {месяц}» | skip / 4xx, та же причина |
| `chain_broken` | skip «пересчитайте {месяц}» | **разрешён** |
| иначе | в zip; после успеха → `exported` | ок |

`vehicle_ids: null` = все активные, затем фильтр.  
Одна плохая машина в запросе `[id]` → 4xx, без zip.  
Смесь чистых и плохих → 200, zip только чистых; UI показывает skip (как сейчас `planKit`). Сервер не кладёт плохих в zip и не ставит им `exported`.

Прыжок «Экспорт» на июль: гейт июля. Отказ сервера → `formatGsmError`, не успешный zip.

Коды: `gsm_kit_tail` | `gsm_kit_chain` | `gsm_kit_red` (или маппинг в существующий `gsm_kit_gate`). Порог литров **0.01**. Не дублировать SQL обзора: reuse `fleet_overview` / `_chain_broken`.

### Code style (гейт)

```python
@dataclass(frozen=True)
class KitEligibility:
    vehicle_id: int
    allowed: bool
    code: str | None
    message: str | None
```

## Этап 2 — норма и округление

Переключатель справочников (`POST /gsm/settings/season`) пишет журнал `season_switches`. `GET` ПЛ (и ответ PATCH) отдаёт `norm_l_per_100 = norm_for(date, vehicle, journal)` — живая цифра, не снимок на generate.

Drawer: `burn = round2(km * waybill.norm_l_per_100 / 100)` без `normForDate` по `season_mode` + инверсии.

После смены сезона инвалидировать не только settings, но waybills/overview.

`round2` на фронте = CPython `round(x, 2)` на фикстурах `1.725`, `2.675`. GET settings по-прежнему показывает «зима с …» для человека; литры из этого поля не считать.

Битый журнал: `gsm_settings_invalid` в registry/generate/export/report.

## Этап 3 — импорт

- Лимиты: ≤10 файлов, ≤20 000 строк суммарно, ≤100 МБ суммарно (+ cap на файл как сейчас).
- Транзакция на **файл**; пачка из пяти `.xls` — битый третий не откатывает первый и второй.
- BIFF / расширение; иначе 400 «нужен Excel 97–2003 (.xls)».
- Filename в batch — basename.

## Этап 4 — LibreOffice (2 затем 1)

Порядок внедрения **внутри этапа 4**, после живых гейтов:

**4.1 Пакетный convert.** Заполнить все `.xlsx` комплекта в temp, затем **один** вызов:

`soffice --headless --convert-to xls --outdir out <dir>`

не N запусков на файл. Прогнать на 2–3 бланках: не использовать фильтр `xls:"MS Excel 97"` (уже ломал Phase 0). Пока 4.1 в том же HTTP — менеджер всё ещё может ждать короче, но CPU занят.

**4.2 Фон + опрос + один слот.** `POST` комплекта → сразу job id (202 или эквивалент). Сборка zip (уже пакетный soffice) в потоке того же процесса. `GET` статуса; `GET` download когда ready. Второй комплект, пока слот занят → 429. `exported` только после успешного zip. Рестарт uvicorn теряет job — допустимо на одном инстансе; предупреждение в UI.

Не в этапе 4: Celery, zip-только-xlsx, несколько web-workers.

## Testing Strategy

| Этап | Доказательство |
|------|----------------|
| 1 | usage августа, Monjaro хвост июля, Palisade чистая, `vehicle_ids: null` → Palisade `exported`, Monjaro нет; `[monjaro]` → 4xx; `chain_broken` → usage не flip, generate проходит; red + новая tx → `has_red_days`; vitest: Export не шлёт id, который planKit исключил |
| 2 | три переключения в журнале; июльский ПЛ после «зима с сегодня» → летняя норма; `burnForKm` = `burn_for_km` на `.xx5`; битый JSON → ошибка GET settings |
| 3 | raise на второй строке файла → 0 строк этого файла; PDF → 400 |
| 4 | мок: convert вызван **один раз** на N xlsx; второй POST комплекта → 429 пока job pending |
| регрессия | май 848; `test_gsm_generator.py` без смены ожиданий солвера |

## Boundaries

- **Always:** гейт хвоста и цепи на usage **и** прямом export; generate не пишет август при открытом июле; красные не `exported`; норма дня из журнала переключателя; тесты этапа, который внедряем.
- **Ask first:** таблица jobs в SQLite; очередь вне процесса; zip без `.xls`; писать live БД; коммит; менять приоритет статуса.
- **Never:** skipped машина в zip/`exported`; `_rechain` вместо generate при разрыве; трогать солвер «заодно»; `--workers > 1` ради soffice; коммитить секреты и `plita.db`.

## Success Criteria

1. Прямой `POST /report/usage` и `POST /waybills/export` не закрывают август машины с открытым июлем.
2. Generate августа при хвосте июля — отказ; generate при только `chain_broken` — успех.
3. «Экспорт» с хвостом качает/закрывает **июль**, если июль проходит гейт; иначе ошибка, не `exported`.
4. Palisade в смешанном флоте закрывается; Monjaro нет; на экране текст skip; без `исключения.txt`.
5. Карточка дня после переключателя сезона в справочниках показывает ту же норму, что бланк на эту дату.
6. Частичный `.xls` не остаётся в БД; мусор — 400.
7. После этапа 4: один `soffice` на комплект; HTTP комплекта не держит менеджера на всём convert; второй комплект ждёт слот.

## Open Questions

Закрыты 2026-08-28: первый код = этап 1; skip только UI; job в RAM; лимиты import 10 / 20k / 100 МБ; полный аудит репо отдельно.

## SDD gate

SPECIFY закрыт. План: [`../develop/plans/2026-08-28-gsm-month-close-hardening.md`](../develop/plans/2026-08-28-gsm-month-close-hardening.md). IMPLEMENT — с этапа 1 после апрува плана.
