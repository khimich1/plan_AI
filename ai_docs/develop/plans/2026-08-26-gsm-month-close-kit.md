# Implementation Plan: GSM — один месяц, один комплект

**Orchestration:** `orch-2026-08-26-gsm-month-close-kit`.
Дата: 2026-08-26. Статус: ✅ implementation completed (без коммита).
Отчёт: [`../reports/2026-08-26-gsm-month-close-kit.md`](../reports/2026-08-26-gsm-month-close-kit.md).
Спека: [`../../specs/gsm-month-close-kit.md`](../../specs/gsm-month-close-kit.md).
Идея: [`../../ideas/gsm-month-close-kit.md`](../../ideas/gsm-month-close-kit.md).

## Overview

Обзор ГСМ: один календарный месяц на экране; хвост «Июль не выгружен»;
одна кнопка комплекта (сводка + бланки, включая `draft` → `exported`);
генерация только по галочкам; стоп текушего месяца **по машине** (хвост
или разрыв цепи). Солвер и `_resolve_start` не трогаем. Live `plita.db`
не пишем.

## Architecture Decisions

- **Поля обзора, не схема БД.** `open_before_month` и `chain_broken`
  считаются в репозитории/сервисе из существующих `gsm_waybill`.
- **Комплект = существующий `POST /report/usage`.** Меняем набор дней
  (`KIT_STATUSES`) и skip красных на бэкенде. Хвост/цепь режет **фронт**
  (не шлёт id); бэкенд красных не закрывает даже при прямом API.
- **Гейт комплекта** — расширение `exportGate.ts` (`planKit`), не новый
  HTTP. UI alert обязателен; манифест в zip не делаем.
- **Одни часы.** Журнал не хранит свой `month`; стрелки пишут `periodFrom`/`To`
  родителя через `monthBounds`.
- **Generate bulk** уже disabled при 0 галок — сохранить и покрыть
  тестом, что POST нет. Skip хвоста/цепи — в том же bulk-отчёте по
  машинам.

## Task List

### Phase 1: Backend (TDD)

- [x] **Task 1: Overview — `open_before_month` и `chain_broken`**
  - **Description:** Расширить SQL/сервис обзора. `open_before_month` =
    `strftime('%Y-%m', max(date))` по `draft|confirmed` с `date < from`.
    `chain_broken` по правилу спеки (последняя ПЛ до from vs первая в
    периоде; Δл > 0.01 или одометры ≠). Схема Pydantic + TS-тип позже
    в T3, здесь Python.
  - **Acceptance:**
    - [x] Нет хвоста → `open_before_month is null`, `open_before == 0`.
    - [x] Draft в июле, from=01.08 → `open_before_month == "2026-07"`.
    - [x] Июль end 27.59 / 62846, август start 18.09 / 62946 →
          `chain_broken true`.
    - [x] Стык бака и одометра → `false`. Нет ПЛ в периоде → `false`.
  - **Verification:**
    `.venv/bin/python -m pytest tests/test_gsm_overview_api.py tests/test_gsm_repository.py -q -k overview`
  - **Dependencies:** нет
  - **Files:** `app/schemas/gsm.py`, `app/repositories/gsm_repository.py`,
    `app/services/gsm_overview_service.py`, `tests/test_gsm_overview_api.py`
    (при необходимости кусок в `test_gsm_repository.py`)
  - **Scope:** M

- [x] **Task 2: Отчёт — KIT_STATUSES, skip красных, draft → exported**
  - **Description:** Строки сводки из `{draft, confirmed, exported}`.
    Машина с `manual_intervention` в периоде не в zip и дни не
    `exported`. Сосед без красных — в сводке. `export_zip` только по
    прошедшим id. Регрессия мая 848.
  - **Acceptance:**
    - [x] Только draft, без красных: N строк сводки, после запроса
          статус `exported`.
    - [x] Красная машина не exported; чистая соседка в zip.
    - [x] Эталон мая 848 (цифры блока) зелёный.
  - **Verification:**
    `.venv/bin/python -m pytest tests/test_gsm_usage_report.py -q`
  - **Dependencies:** нет (параллельно T1)
  - **Files:** `core/gsm/usage_report.py`,
    `app/services/gsm_report_service.py`, `tests/test_gsm_usage_report.py`
  - **Scope:** M

### Checkpoint: Backend

- [x] `tests/test_gsm_overview_api.py` + `test_gsm_usage_report.py` зелёные
- [x] Контракт GET overview аддитивен (старые поля на месте)

### Phase 2: Фронт — контракт и гейты

- [x] **Task 3: Типы + подпись месяца + `planKit`**
  - **Description:** `FleetOverviewRow` += `open_before_month`,
    `chain_broken`. Хелпер подписи «Июль не выгружен». `planKit(rows,
    selectedIds | null, periodFrom)`: исключить red / хвост / цепь,
    кроме периода = месяц хвоста; confirm yellow / already exported.
    Bulk generate: 0 id → не вызывать mutate (тест).
  - **Acceptance:**
    - [x] `openBeforeMonthLabel("2026-07")` → текст с «июл».
    - [x] Август + open_before июля: машина не в `cleanIds` отчёта.
    - [x] Тот же ряд при periodFrom июля — в `cleanIds` (если не red).
    - [x] `chain_broken` на текущем периоде — не в clean, не red-only.
    - [x] `planKit` при `selectedIds=null` смотрит все строки.
  - **Verification:**
    `cd frontend && npx vitest run src/features/gsm/lib/fleetStatus.test.ts src/features/gsm/lib/exportGate.test.ts`
  - **Dependencies:** контракт T1 (поля)
  - **Files:** `frontend/src/features/gsm/types/gsm.ts`,
    `lib/fleetStatus.ts`, `lib/exportGate.ts`, их `.test.ts`
  - **Scope:** M

### Phase 3: UI

- [x] **Task 4: Один календарный месяц (журнал = фильтры)**
  - **Description:** Убрать локальный `month` из журнала как источник
    истины: сетка и запросы ПЛ/tx по `periodFrom`/`periodTo`. Стрелка
    календаря → родитель ставит `monthBounds`. Тест «после paging
    generate-диалог с месяцем журнала» заменить: paging меняет обзорный
    период.
  - **Acceptance:**
    - [x] Фильтры 01.08–31.08 → календарь «Август 2026», не июль.
    - [x] «Предыдущий месяц» → 01.07–31.07 в инпутах и сетке.
  - **Verification:**
    `cd frontend && npx vitest run src/features/gsm/components/VehicleWaybillJournal.test.tsx src/features/gsm/components/FleetOverviewView.test.tsx`
  - **Dependencies:** нет (можно сразу после T3 или параллельно с T3
    по журналу)
  - **Files:** `VehicleWaybillJournal.tsx`, `FleetOverviewView.tsx`,
    соответствующие `.test.tsx`
  - **Scope:** M

- [x] **Task 5: Обзор — комплект, баннер, строка, пересчёт**
  - **Description:** Хинт «сводка и путевые». Убрать «Экспорт zip
    выбранных» и модалку zip-только-ПЛ. Баннер с именем месяца,
    клик → границы `open_before_month`. «Отчёт за период» через
    `planKit` + alert исключений; 0 clean → не POST. Синее «Экспорт» —
    `button`: хвост этой машины или текущий период. CTA «Пересчитать
    {месяц}» при `chain_broken` → `VehicleGenerateDialog`. Bulk generate
    фильтрует skip-машины в список результатов без POST по ним.
  - **Acceptance:**
    - [x] Кнопки zip-ПЛ нет.
    - [x] Баннер «Июль не выгружен».
    - [x] Отчёт без галок: чистая в mutate, хвост в exclusions.
    - [x] Экспорт в строке — button, не span.
    - [x] `chain_broken` → «Пересчитать август» открывает диалог.
    - [x] Bulk generate 0 галок — нет mutate.
  - **Verification:**
    `cd frontend && npx vitest run src/features/gsm/components/FleetOverviewView.test.tsx src/features/gsm/components/FleetOverviewTable.test.tsx`
  - **Dependencies:** T3, T4
  - **Files:** `FleetOverviewView.tsx`, `FleetOverviewTable.tsx`,
    тесты; чуть `VehicleWaybillJournal.tsx` если CTA в шапке журнала
  - **Scope:** M

### Checkpoint

- [x] `.venv/bin/python -m pytest tests/test_gsm_overview_api.py tests/test_gsm_usage_report.py tests/test_gsm_repository.py -q`
- [x] `cd frontend && npx vitest run src/features/gsm/`
- [ ] Глазами на UI (без generate live): август, баннер, стрелка, 0 галок
      generate disabled. Live `plita.db` не писать.

## Risks and Mitigations

| Риск | Вероятность | Митигация |
|------|-------------|-----------|
| Май 848 usage-report падает из-за KIT_STATUSES | Средняя | Тот же набор дней (уже exported); гонять эталон в T2 первым |
| `chain_broken` ложные срабатывания float | Низкая | порог 0.01 как `litersDiffOk` |
| SQL `open_before_month` vs Python max date | Низкая | один источник в репозитории, unit на ISO |
| T5 раздувается | Средняя | CTA пересчёта только в таблице, не дублировать три места |
| Журнал generate обходит bulk-гейты | Средняя | те же `open_before`/`chain_broken` на кнопке журнала в T5 |

## Open Questions

Нет. Манифест в zip — не делаем (решение плана: UI alert + backend skip
красных).

## Порядок

```
T1 (overview) ──┐
T2 (report)  ───┴─ checkpoint backend ─ T3 (гейты) ─┐
T4 (часы) параллельно T3 ──────────────────────────┴─ T5 (обзор) ─ checkpoint
```
