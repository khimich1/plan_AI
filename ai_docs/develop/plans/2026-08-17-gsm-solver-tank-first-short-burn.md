# Implementation Plan: GSM солвер — бак важнее группы АЗС + короткий дожиг

## Overview

Доработка генератора v2: ложные `manual_intervention` на Palisade мае 2026
из‑за узкой группы АЗС (A) и сетки дожига 150–250 км (B). Каркас lookahead
не меняется. Спека: [`../../specs/gsm-solver-tank-first-short-burn.md`](../../specs/gsm-solver-tank-first-short-burn.md).
Идея: [`../../ideas/gsm-solver-tank-first-short-burn.md`](../../ideas/gsm-solver-tank-first-short-burn.md).

## Architecture Decisions

1. **Сначала B, потом A.** Новый пул дожига меняет `_required_lookahead_km`
   (шаг 2: «будни справятся»). Если сделать A первым, 06.05 уедет на Ковров
   ещё до того, как четверг научится жечь 10 л.
2. **Не расширять группу до проверки дожига.** A — fallback только если
   `km_needed > 0` после нового дожига и группа не набрала km.
3. **R4 на попадании в headroom:** среди дожигов, которые уже в коридоре к
   `Q_next`, брать минимальный km, не max burn и не max frequency.
4. **Emit / API / UI не трогаем.** Предохранитель остаётся; цель — чтобы он
   не стрелял на мае Palisade.
5. **Приёмка на копии БД**, не в рабочий `plita.db`.

## Task List

- [x] Task 1: Короткий дожиг (B) + тесты
  - **Description:** Убрать фильтр `BURN_KM_MIN/MAX` в `_ordered_burn_routes`;
    пул = `route_list` (уже `2×km ≤ max_daily_km`). В `_plan_burn_in` для
    `reaching` сортировать по `_daily_km` asc, затем frequency desc, `route_id`.
    Пока не достигли headroom — max безопасный burn как сейчас.
  - **Acceptance:**
    - [x] Якорь ср. + свободный чт. + пт. с большим Q: typical якоря короткие;
          в библиотеке есть плечо ~45 км. Якорь остаётся typical; чт. — дожиг
          с `fuel_end ≥ 0` и `fuel_end ≤ tank − Q_пт`; пт. не manual.
    - [x] Дожиг, который уводит бак ниже 0, не выбирается.
    - [x] Существующие lookahead-тесты, где дожиг 150–250 был частью сценария,
          зелёные или явно адаптированы.
  - **Verification:** `venv/bin/pytest tests/test_gsm_generator.py -q -k "burn or lookahead"`
  - **Dependencies:** None
  - **Files:** `core/gsm/generator.py`, `tests/test_gsm_generator.py`
  - **Estimated scope:** M

- [x] Task 2: Lookahead fallback на всю библиотеку (A) + тесты
  - **Description:** В `_select_anchor_route_lookahead`: если
    `_pick_min_sufficient(group)` вернул `None`, повторить по `routes`
    (полный кап). При смене маршрута — `balance_route`.
  - **Acceptance:**
    - [x] Два якоря подряд, 0 будней, typical ≤100 км, в библиотеке ≥180 км
          плеча → день 1 удлинён, `balance_route`, день 2 не manual.
    - [x] Если достаточный km есть в группе — полную библиотеку не трогать
          (гео/typical сохраняются).
    - [x] Нет решения ни в группе, ни в библиотеке → поведение v2 (manual).
  - **Verification:** `venv/bin/pytest tests/test_gsm_generator.py -q`
  - **Dependencies:** Task 1 (чтобы A не перехватывал кейс с буднем)
  - **Files:** `core/gsm/generator.py`, `tests/test_gsm_generator.py`
  - **Estimated scope:** S

- [x] Task 3: Регрессия GSM + приёмка Palisade май
  - **Description:** Полный generator suite + соседние GSM pytest. Прогон
    Palisade 04.05–31.05 на **копии** `plita.db` (28 л / 128327, force).
    Зафиксировать `problematic_days`. Если не пусто — стоп, не угадывать
    двухякорный lookahead в этом срезе.
  - **Acceptance:**
    - [x] `pytest tests/test_gsm_generator.py tests/test_gsm_generation_service.py tests/test_gsm_generation_api.py -q` зелёный
    - [x] Palisade май: `problematic_days == []`
    - [x] Короткий отчёт в `ai_docs/develop/reports/`
  - **Verification:** команды из спеки; отчёт
  - **Dependencies:** Task 1, Task 2
  - **Files:** `ai_docs/develop/reports/2026-08-17-gsm-solver-tank-first-short-burn.md` (после прогона)
  - **Estimated scope:** S

## Dependency Graph

```
Task 1 (B short burn) → Task 2 (A full library) → Task 3 (acceptance)
```

Параллелить A и B нельзя: порядок влияет на 06.05.

## Risks

| Риск | Митигация |
|---|---|
| Frequency-first дожиг в старых тестах станет min-km | Адаптировать asserts; зафиксировать R4 в тесте |
| Короткий дожиг 6 км на крошечном burn | По спеке ок; на 07.05 не выберется (недостаточный km) |
| A размазывает географию | Fallback только если группа не набрала km; внутри полной библиотеки та же `_direction_priority` |
| Май 21.05 всё ещё красный | Task 3 стоп; не тащить двухякорный lookahead молча |

## Out of Scope

- Ночёвки, ILP, синтетические маршруты, UI, API `problematic_days`,
  двухякорный lookahead, запись в рабочий `plita.db`.
