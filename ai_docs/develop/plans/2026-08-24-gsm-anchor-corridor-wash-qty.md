# Implementation Plan: GSM — коридор бака и честные литры моек

Дата: 2026-08-24. Статус: draft, на ревью.
Спека: [`../../specs/gsm-anchor-corridor-wash-qty.md`](../../specs/gsm-anchor-corridor-wash-qty.md) (2026-08-24, direction confirmed).
Контекст: обнаружено при приёмке `gsm-fleet-overview-ux` на живых данных Palisade (август 2026).

## Overview

Мини-срез из 4 задач: коридор бака в генераторе, фильтр Σ литров по топливу в обзоре, обнуление `qty_liters` у моек при импорте, приёмка на копии БД. Без фронтенда.

## Architecture Decisions

- **Коридор как фильтр перед выбором**, а не после. Минимальное вмешательство в поток `group → choose → lookahead → emit`.
- **Мойка = минимальный km из коридора**, не «самый частый». Lookahead сверху может перебить.
- **Fallback при пустом коридоре**: самый короткий из группы → если и он в минус, `_emit_day` даёт `manual_intervention`. Новых warning-кодов не вводим.
- **Σ литров только fuel**: изменение SQL-агрегата `fleet_overview`, не трогаем `TransactionListResponse` (там итоги по фильтру, включая wash — это отдельный журнал).

## Task List

### Phase 1: Генератор (высший риск, первая)

- [x] **Task 1: Коридор бака при выборе маршрута якоря**
  - **Description:** В `_select_anchor_route_lookahead` добавить фильтр `in_corridor`. Мойка (`q_today == 0`) выбирает `min(in_corridor, key=km)`. Топливо — `max(in_corridor, key=frequency)`. Fallback: `min(group, key=km)`. Lookahead работает поверх, может удлинить.
  - **Acceptance criteria:**
    - [x] Две мойки подряд (03.08, 04.08) при баке 41,13 л и норме 14,5 → 04.08 получает 6 км (12 км круг), бак 11,84 л, без `manual_intervention`.
    - [x] Существующие тесты генератора зелёные.
    - [x] Тест «мойка + lookahead»: короткий маршрут заменяется на длинный, если нужен выжиг.
  - **Verification:**
    - `venv/bin/pytest tests/test_gsm_generator.py -q` — зелёный.
    - `venv/bin/pytest tests/test_gsm_generator.py -k "two_washes" -q` — новый тест зелёный.
  - **Dependencies:** None
  - **Files:** `core/gsm/generator.py`, `tests/test_gsm_generator.py`
  - **Scope:** S (1-2 файла)

### Phase 2: SQL-агрегат и импорт (параллельны после Task 1)

- [x] **Task 2: Σ литров в обзоре — только топливо**
  - **Description:** В `fleet_overview` изменить `SUM(t.qty_liters)` на `SUM(CASE WHEN t.service_type = 'fuel' THEN t.qty_liters ELSE 0 END)`. `liters_diff` становится топливо vs топливо.
  - **Acceptance criteria:**
    - [x] `test_gsm_overview_api.py` с мойкой в периоде: `tx_liters` не включает мойку, `liters_diff` = топливо − выданное.
    - [x] `red_days` и статусы не задеты.
  - **Verification:** `venv/bin/pytest tests/test_gsm_overview_api.py -q`
  - **Dependencies:** Task 1 (чтобы `red_days` был 0 для чистого теста)
  - **Files:** `app/repositories/gsm_repository.py`, `tests/test_gsm_overview_api.py`
  - **Scope:** XS (1-2 файла)

- [x] **Task 3: Импорт моек — `qty_liters = None`**
  - **Description:** В `GsmTransactionService.import_files` при `service_type == "wash"` передавать `qty_liters=None`, даже если парсер вернул число.
  - **Acceptance criteria:**
    - [x] Импорт `.xls` с мойкой и числом в файле → в БД `qty_liters IS NULL`.
    - [x] Существующие тесты импорта зелёные (парсер по-прежнему возвращает `None` для wash; сервис не ломает).
  - **Verification:** `venv/bin/pytest tests/test_gsm_transaction_import.py -q`
  - **Dependencies:** None
  - **Files:** `app/services/gsm_transaction_service.py`, `tests/test_gsm_transaction_import.py`
  - **Scope:** XS (1-2 файла)

### Checkpoint: Backend готов

- [x] `venv/bin/pytest tests/test_gsm_*.py -q` — зелёный (218 passed)
- [ ] `venv/bin/pytest tests/ -q` — не зелёный: 11 failed, 2244 passed, 8 skipped (те же 11 вне GSM, что на `gsm-fleet-overview-ux`; не регрессия T1–T3)

### Phase 3: Приёмка на живых данных (копия)

- [x] **Task 4: Генерация августа Palisade на копии БД**
  - **Description:** Скопировать `plita.db` → `/tmp/plita_accept.db`. Сгенерировать август 2026 для Palisade через API или сервис. Проверить, что 04.08 — draft 12 км, `red_days = 0`, `liters_diff = 0.0`.
  - **Acceptance criteria:**
    - [x] `GET /overview?from=2026-08-01&to=2026-08-31` на копии: Palisade `status = "ready"` или `"drafts_pending"`, `red_days = 0`, `liters_diff = 0.0`.
    - [x] Журнал Palisade: 04.08 — маршрут 6 км (12 км круг), бак ≥ 0.
  - **Verification:** curl на копии + ручная проверка
  - **Dependencies:** Task 1, Task 2
  - **Files:** нет (приёмка)
  - **Scope:** XS

## Risks and Mitigations

| Риск | Вероятность | Митигация |
|:---|:---|:---|
| Существующие тесты генератора сломаются из-за фильтра коридора | Средняя | TDD: сначала красный тест, потом минимальная правка; если старый тест ожидает длинный маршрут — проверить, был ли он в коридоре |
| Мойка с lookahead: короткий маршрут не даёт выжига | Низкая | Тест «мойка + lookahead» в Task 1 |
| `FILTER (WHERE ...)` не поддерживается SQLite | Низкая | SQLite ≥ 3.30 поддерживает; fallback — `SUM(CASE WHEN ...)` |
| Незакоммиченные изменения от прошлого среза мешают diff | Низкая | Правки не пересекаются по строкам; при коммите разделить по файлам |

## Open Questions

- Нет. Решения зафиксированы в спеке.
