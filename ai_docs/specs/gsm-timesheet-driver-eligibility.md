# Spec: ГСМ — табель Т-13 решает, кого можно подписать в путевой

Дата: 2026-10-12. Статус: draft, на ревью.
Идея: [`../ideas/gsm-timesheet-driver-eligibility.md`](../ideas/gsm-timesheet-driver-eligibility.md)
(2026-10-12, direction confirmed).
Предшественники: [`gsm-module-putevye-listy.md`](gsm-module-putevye-listy.md)
(водитель = `primary_driver_id`), [`gsm-month-close-kit.md`](gsm-month-close-kit.md)
(замки комплекта), [`gsm-fleet-pool-corridor.md`](gsm-fleet-pool-corridor.md)
(жёлтые/красные коды дня).

## ASSUMPTIONS I'M MAKING

1. **Табель — форма Т-13 (.xlsx) как присланные № 44/47.** Строка человека:
   имя с `\n` в колонке C, табельный `СЭхх-ххххх` в колонке E. Сетка дней:
   колонки `I,K,M,N,P,R,T,V,X,Z,AB,AD,AF,AH,AK` строки `r` (1–15) и `r+2`
   (16–31; день 31 в колонке `AL` строки `r+2`). Строки часов (`r+1`, `r+3`)
   игнорируем. Цвет заливки не читаем.
2. **Период табеля — из имени файла** (регулярка `от DD.MM.YYYY` → месяц).
   Не распознали — бухгалтер выбирает месяц в диалоге загрузки вручную.
   Дни за пределами месяца не импортируем (февраль: колонки 29–31 отсутствуют).
3. **Разбор табельного:** берём цифры после дефиса, убираем ведущие нули:
   `СЭ00-00470 → 470`, `СЭ0В-00001 → 1` (буква в `СЭ0В` — кириллица, отрезается
   вместе с серией). Стыковка только по номеру с `gsm_driver.personnel_number`.
   Строки без совпадения — в отчёт импорта, дни их не сохраняются.
4. **Правило допуска:** в ПЛ на дату можно ставить только водителя с кодом
   **Я или В** в этот день и не занятого другим ПЛ на эту дату. Всё остальное
   (`ОТ`, `Б`, `К`, `ПК`, любой неизвестный код) — жёсткий запрет. Код `В`
   даёт жёлтый `weekend_driver`, не красный.
5. **Ротация запасных.** Кандидат = свободный с **меньшим счётчиком
   подстановок за период** (лист не на своей машине; seed из БД + дни текущего
   прогона); равенство → меньший табельный. Группы: без своей машины, потом
   чужие основные. Эффект на паре Полякова/Скрябин — чередование (П, С, П, С…);
   отпуск одного даёт серию другому. Счётчик — данные БД, не хардкод имён.
6. **Без табеля за период generate запрещён** (`gsm_timesheet_missing`),
   покрытие = объединение периодов загруженных батчей. Комплект месяца без
   табеля тоже не закрывается (export-валидация).
7. **Красный день «нужен водитель»** (`no_available_driver`) сохраняется как
   draft с `primary_driver_id` машины + warning; для комплекта — hard-stop,
   как `manual_intervention`. Бухгалтер чинит в drawer.
8. **Жёлтый `weekend_driver`** — мягкий: confirm при экспорте, рядом с
   `balance_route` / `borrowed_route`. Подстановка без пометки (plain):
   никаких бейджей «подставлено».
9. **История не переписывается.** Якорь 30.06 (Кулигин на обеих Tugella,
   `exported`) не трогаем; UNIQUE `(driver_id, date)` в SQLite в этом срезе
   **не ставим** — стопы на уровне сервисов. Старые месяцы валидацией не
   догоняем: export-валидация применяется к комплекту, который выгружают
   сейчас (июль уже выгружен — его не трогаем).
10. **Шишов 241 → 438** — ручная правка в справочнике до загрузки августа,
    не миграция и не подсказка импорта.
11. **Повторная загрузка месяца** заменяет дни батчей за пересекающийся
    период (перезалил исправленный Т-13 — явки обновились). Батчи-история
    хранится, дни — актуальные.
12. **Пул дня считается на момент генерации каждой машины.** Bulk generate
    идёт по машинам последовательно в заданном порядке; свободные для даты =
    допущенные по табелю минус уже стоящие в `gsm_waybill` на эту дату
    (любой машины). Свои дни этой машины внутри периода при подсчёте
    исключаются (они будут перезаписаны).

→ Поправьте сейчас, иначе иду с этим в план/задачи.

## Objective

Бухгалтер загружает Т-13 за месяц и генерирует путевые, в которых:

1. **Один человек — один ПЛ в день** по всему парку (не только по машине).
2. **Отсутствующий по табелю не подписывается**: `ОТ`/больничный/командировка
   — жёсткий запрет; выходной — жёлтое предупреждение.
3. **Подстановка автоматическая**: Кулигин занят 848-й → 952 получает
   Полякову/Скрябина по порядку; ручная разводка — только красные дни
   («нужен водитель», свободных нет).

**Пользователь:** бухгалтер (`accountant`), админ (`admin`).

**Критерий успеха MVP:** на табелях июля и августа 2026 generate+bulk не даёт
ни одной пары `(driver, date)` в двух ПЛ, ни одного листа на человека в
`ОТ`; месяц закрывается без ручной правки водителей, кроме дней, где свободных
нет физически (в присланных табелях таких будних нет; суббота/воскресенье —
`weekend_driver`).

## Tech Stack

- Backend: Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`), openpyxl
  (чтение .xlsx — уже в venv).
- Frontend: React 18, TypeScript, TanStack Query, `frontend/src/features/gsm/`.
- Новых внешних зависимостей нет.

## Commands

```bash
# Backend
uvicorn app.main:app --reload
pytest tests/test_gsm_timesheet.py tests/test_gsm_generator.py \
  tests/test_gsm_generation_api.py tests/test_gsm_export.py \
  tests/test_gsm_waybill_edit.py -q
pytest tests/ -q

# Frontend
cd frontend && npm test -- --run src/features/gsm && npm run build
```

## Project Structure

```
core/
  gsm/
    timesheet.py                → NEW: pure parser Т-13 (openpyxl → rows),
                                  разбор табельного, сетка кодов
    driver_pool.py              → NEW: pure: допуск по коду, ротация по
                                  счётчику подстановок, выбор водителя дня
    generator.py                → CHANGE: driver_id → выбор из пула дня
  kp_db_schema.py               → +gsm_timesheet_batch, gsm_timesheet_day

app/
  api/v1/endpoints/gsm.py       → +POST /timesheets/import, GET /timesheets
  schemas/gsm.py                → +TimesheetImportReport, TimesheetGrid
  services/
    gsm_timesheet_service.py    → NEW: импорт, замена периода, отчёт
    gsm_generation_service.py   → CHANGE: гейт табеля, пул дня, bulk-занятость
    gsm_export_service.py       → CHANGE: валидация водителей комплекта
  repositories/gsm_repository.py → +timesheet CRUD, занятость водителя по дате

frontend/src/features/gsm/
  api/gsmApi.ts                 → +importTimesheet, getTimesheet
  components/
    TimesheetImportDialog.tsx   → NEW: загрузка + отчёт (строки без водителя)
    TimesheetGridView.tsx       → NEW: сетка Я/В/ОТ по жёлтым за месяц
    GsmRegistriesView.tsx       → +карточка «Табель рабочего времени»
    WaybillDayDrawer.tsx        → тексты 409 по водителю
  lib/waybillWarnings.ts        → +weekend_driver (soft), no_available_driver
  lib/exportGate.ts             → HARD += no_available_driver,
                                  SOFT += weekend_driver

tests/
  test_gsm_timesheet.py         → NEW: парсер, номера, допуск, замена периода
  test_gsm_driver_pool.py       → NEW: порядок кандидатов, занятость, выходной
  fixtures/gsm/                 → +t13_july_small.xlsx (openpyxl, 2 машины)
```

## Code Style

```python
# core/gsm/timesheet.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import date

ELIGIBLE_CODES = frozenset({"Я", "В"})


@dataclass(frozen=True, slots=True)
class TimesheetRow:
    """One person's attendance grid parsed from a Т-13 sheet."""

    full_name: str
    personnel_number: str          # разобранный: '470', не 'СЭ00-00470'
    codes: tuple[tuple[date, str], ...]  # (день, код) за месяц


def parse_personnel_number(raw: str) -> str | None:
    """'СЭ00-00470' → '470'; 'СЭ0В-00001' → '1'; мусор → None."""
```

Конвенции (как в модуле): роутер → сервис → репозиторий; `core/gsm/*` без
`app.*` и без доступа к БД; DTO `@dataclass(frozen=True, slots=True)`;
солвер и пул детерминированы.

## Алгоритм выбора водителя дня

Вход на дату D машины V: `primary(V)`, табельные коды дня, занятость
(кто уже в `gsm_waybill` на D, кроме дней V внутри периода).

```
eligible = { d | code(d, D) in {Я, В} }
free     = eligible − busy(D)
ранг внутри группы = (счётчик_подстановок_за_период, табельный) ↑
candidates =
    [primary(V)]            если primary(V) in free
  + без-машинные(free)      по рангу
  + чужие-основные(free)    по рангу

если candidates: driver = head(candidates)
    driver != primary(V)    → счётчик driver +1 (локально, дальше по периоду)
    code(driver, D) == 'В'  → warning weekend_driver
иначе: driver = primary(V)  → warning no_available_driver (hard-stop комплекта)
```

Подстановка применяется к **каждому** дню (якорь и дожигание). `WaybillDay`
несёт выбранного `driver_id`; `no_available_driver` — draft + hard-stop,
чинится в drawer.

### Проверки вне generate

- `POST /waybills` (ручной ПЛ), `PATCH /waybills/{id}` (смена водителя):
  код не Я/В → 409 `gsm_driver_absent`; уже занят другой машиной →
  409 `gsm_driver_busy` (с датой и госномером в details). Если за месяц
  табель не загружен — правило по коду не применяется (generate в такой
  месяц и так запрещён), занятость проверяется всегда.
- `POST /waybills/export` (комплект): перед вызовом soffice валидация —
  нет пары `(driver, date)` на двух машинах, нет листа с кодом не Я/В;
  нарушение → 409 `gsm_driver_rule_violation` со списком дат. Табель за
  период отсутствует → 409 `gsm_timesheet_missing`.

### Таблицы (стиль `CREATE TABLE IF NOT EXISTS`)

```sql
gsm_timesheet_batch (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  filename TEXT NOT NULL,
  period_from TEXT NOT NULL, period_to TEXT NOT NULL,
  rows_total INTEGER, matched INTEGER,
  uploaded_by TEXT, uploaded_at TEXT NOT NULL
);

gsm_timesheet_day (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  batch_id INTEGER NOT NULL REFERENCES gsm_timesheet_batch(id),
  driver_id INTEGER NOT NULL REFERENCES gsm_driver(id),
  date TEXT NOT NULL,
  code TEXT NOT NULL
);
-- индекс (driver_id, date); импорт периода удаляет старые дни периода
```

## API

Prefix `/api/v1/gsm`, auth `REQUIRE_ACCOUNTING`.

| Method | Path | Purpose |
|--------|------|---------|
| POST | `/timesheets/import` | .xlsx Т-13 → отчёт: период, сопоставленные, строки без водителя, коды-итоги |
| GET | `/timesheets?period=YYYY-MM` | Сетка кодов по сопоставленным водителям за месяц |

Ошибки: `400` не Т-13/не xlsx/месяц не распознан и не выбран; `403` чужая
роль; `409` водительские правила и занятость; `422` generate без табеля
(`gsm_timesheet_missing` — конфигурация, не нерешаемость).

## Testing Strategy

| Уровень | Что покрывает |
|---------|---------------|
| Unit (pytest) | `parse_personnel_number`: `СЭ00-00143→143`, `СЭ0В-00001→1`, мусор→None |
| Unit (pytest) | Парсер Т-13: сетка кодов июля, 31 день из `AL`, строки часов пропущены, неизвестный код сохраняется и считается запретом |
| Unit (pytest) | Пул: ротация (0/0 → Полякова; после её дня → Скрябин; серия в отпуск Скрябина), `В` даёт weekend_driver, `ОТ` исключает, пусто → no_available_driver |
| Unit (pytest) | Генератор: две машины в bulk — Кулигин только на одной в каждый день; подстановка на якорь и на дожигание; запасные чередуются П, С; отпуск Лоншаковой 27.07–09.08 → её нет в листах Monjaro |
| Integration | Гейт `gsm_timesheet_missing`; create/patch 409 busy/absent; export-валидация блокирует пару и отпускника; повторный импорт заменяет дни периода |
| Unit (vitest) | Диалог импорта (отчёт), сетка, тексты warnings, exportGate HARD/SOFT |
| Приёмка | Копия БД: загрузить Т-13 июль+август, bulk generate обоих месяцев по 4 машинам → ноль пар, ноль `ОТ` в листах, `no_available_driver` = 0, июль/август закрываются комплектом |

## Boundaries

- **Always:** TDD на `core/gsm/timesheet.py` и `driver_pool.py`; `pytest`
  зелёный; минимальный diff; табель-файлы из `Загрузки/` — только чтение,
  фикстуры в `tests/fixtures/gsm/` (без персональных данных сверх 5 жёлтых).
- **Ask first:** новые таблицы `gsm_timesheet_*` (миграция схемы);
  UNIQUE `(driver_id, date)` в БД — отдельное решение после якоря 30.06;
  изменение контракта `generate()` в `core/gsm/generator.py` (сигнатура
  `driver_id` → пул) — зафиксировать в плане.
- **Never:** не править якорь 30.06 и exported-листы; не хранить табель
  вне БД после импорта; не подставлять по цвету/ФИО; не коммитить без
  явной просьбы.

## Success Criteria

1. SC-1: generate любого месяца без табеля → 422 `gsm_timesheet_missing`.
2. SC-2: bulk 848+952 на июль (Т-13 № 44): на каждый день Кулигин максимум
   в одном ПЛ; у второй Tugella запасные чередуются (Полякова/Скрябин по
   счётчику подстановок), в отпуск Скрябина 1–9.07 — серия Поляковой.
3. SC-3: в ПЛ Monjaro на 27.07–09.08 Лоншаковой нет; в ПЛ Palisade на
   18–31.08 Шишова нет.
4. SC-4: субботний якорь подписан основным/запасным с `weekend_driver`;
   комплект закрывается через confirm.
5. SC-5: drawer-правка «посадить занятого/отпускника» → 409 с понятной
   фразой; экспорт комплекта с такой парой → 409 до soffice.
6. SC-6: регрессия `pytest tests/` и `npm test -- --run` зелёные.

## Open Questions

1. UNIQUE `(driver_id, date)` в SQLite — после ручного разведения якоря
   30.06; в этом срезе только сервисные стопы.
2. Фильтрованный список водителей в drawer (только допущенные на дату) —
   follow-up; MVP — 409 с текстом.
3. Коды `Б`/`К`/`ПК` в живых табелях ещё не встречались — правило «не Я и
   не В» их уже покрывает; при появлении сверить расшифровку с бухгалтером.
4. Окно счётчика ротации = генерируемый период (месяц). Скользящее окно
   (30 дней назад) — если бухгалтеру важно равенство «по году», follow-up.
