# Spec: Стабилизация модуля Производство — аудит 2026-08-12

> **Тип:** remediation feature-spec (стабилизационный срез)  
> **Фаза SDD:** SPECIFY ✅ → PLAN (ревью) → TASKS ✅ в плане → IMPLEMENT ⏸  
> **Дата:** 2026-08-14  
> **Статус:** PLAN READY FOR REVIEW  
> **План:** [`../develop/plans/2026-08-14-stabilizaciya-production-audit.md`](../develop/plans/2026-08-14-stabilizaciya-production-audit.md)  
> **Источник:** [`../develop/audits/2026-08-12-production-module-audit.md`](../develop/audits/2026-08-12-production-module-audit.md)  
> **Разбор приоритетов:** диалог 2026-08-14 (что чинить обязательно vs долг)  
> **Baseline:** [`project-baseline.md`](./project-baseline.md)  
> **Связанные:** [`planirovanie-po-srokam-podlozhki.md`](./planirovanie-po-srokam-podlozhki.md), [`calendar-first-planning.md`](./calendar-first-planning.md), [`sgp-warehouse.md`](./sgp-warehouse.md), [`move-to-production-atomicity-q1-q2.md`](./move-to-production-atomicity-q1-q2.md)

---

## Стратегия (одной фразой)

> Закрыть рассинхрон учёта и плана (два валидатора ёмкости, фейковый optimistic lock, неатомарные complete_day и build+SGP) и дешёвые дыры на входе API — без нарезки god-модулей и без очереди CPU-воркеров.

---

## ASSUMPTIONS I'M MAKING

1. **Scope = обязательные пункты** из разбора аудита, не все 48 findings. Три волны в одной спеке, реализация — вертикальными срезами.
2. **Дубликаты считаем одним тикетом:** A1=Q1, A3=S3, S1=Q7, Q4=A8 (Q4 только piggyback на Q5).
3. **План и КП живут в одном `plita.db`.** Complete_day — одна sqlite-транзакция. Build+SGP — компенсация (delete плана + возврат плит), не outbox.
4. **Повторный complete уже завершённого дня → 409** (`day_already_completed`), не идемпотентный 200. Snapshot с `write_off_completed` не списываем повторно.
5. **`expected_version` обязателен на mutating complete / remove_track**, если клиент его прислал. Если не прислал — поведение как сейчас (взять текущую version), чтобы не ломать старых клиентов/скрипты.
6. **Лимит диапазона дат = 366 календарных дней** включительно. `FillTargetItem.date` — ISO `YYYY-MM-DD`; span min..max fill ≤ 366; `min_fill` не дальше today+366. **Пола «сегодня−30» нет** (ломает фикстуры и дозаполнение прошлых дней).
7. **Формула оценки дорожек/дней = `ceil`** (как backend `math.ceil`). FE `Math.round(x + 0.5)` и закреплённый тест 38.3м → 2 дня — правим.
8. **A4 узкий:** `delivery_schedule_service` читает тот же calendar, что production (`PlanDistributionService.get_global_calendar_info` / PlanRepository + day_capacity). Полный deprecate `plan_calendar.py` и бот — **out of scope**.
9. **God-модули не режем:** `planning.py`, `production_completion_service.py`, `useCreatePlanWizardState.ts`, толстый router, CPU-worker, rate limit — **out of scope**.
10. Audit-файл **не** помечаем FIXED в этом изменении (отдельный follow-up после IMPLEMENT).
11. Поставка — **три PR** (волны). A3 в этом срезе — **компенсация** (delete плана + возврат плит), не протаскивание conn через `PlanPersistPort`.
12. **Текущий `plita.db` — тестовые данные, не боевой учёт.** Не пишем миграции/dual-read старых payload планов ради совместимости с рядами в БД. Pytest-фикстуры и контракт живого фронта — по-прежнему обязательны.

→ Решения D1–D9 зафиксированы. IMPLEMENT — после «ок» по плану.

---

## 1. Objective

### Что строим и зачем

После аудита модуля Производство (2026-08-12, здоровье 4/10) закрыть **реальный риск рассинхрона учёта и плана** и **дешёвые входные дыры**, из‑за которых планировщик видит «OK», а склад/КП уже в другом состоянии.

**Пользователи / акторы:**
- `production` / `admin` — визард плана, календарь ёмкости, завершение смены, удаление дорожки
- менеджер (график поставки) — сроки должны совпадать с ёмкостью цеха
- система учёта КП/СГП — списание и резерв не должны расходиться с JSON плана

### Problem statement

| ID | Severity | Суть (сейчас в коде) |
|----|----------|----------------------|
| A1/Q1 | Critical | `capacity.validate_fill_targets` без occupancy; `planning.validate_fill_targets` со свободными слотами. Analyze «OK», build 422 |
| S1/Q7 | High | FE шлёт `expected_version`; `CompleteProductionDayRequest` поля нет; DELETE track query не принимается. Lock мёртвый |
| S2 | High | `send_to_sgp` без guard `day.completed`; `_collect_plates_by_kp` не фильтрует `write_off_completed`; КП-списание и `mark_day_completed` — разные соединения |
| A3/S3 | High | `build_plan` коммитит план+плиты, затем отдельный conn на `reserve_on_conn`; при ошибке СГП план остаётся |
| S4 | High | `_date_range_inclusive` без max span; analyze `while cursor < min_fill` при далёкой дате; `FillTargetItem.date: str` |
| Q5 | High | schema `le=50` vs hard cap 5; `SaveDayCapacityRequest.max_tracks` без `le` |
| S9 | Medium | `GET /candidates?limit=` без верхней границы |
| S8 | Medium | `target_date` / `date` path без ISO |
| Q3 | High | FE `round(x+0.5)` vs BE `ceil` — ложные сроки в UI |
| A11 | Medium | occupancy API отдаёт глобальный `MAX_TRACKS_PER_DAY`, игнорируя override дня |
| Q9 | Medium | `ProductionSubstrateError` → пустые рекомендации без текста ошибки |
| A4 | High | график поставки тянет legacy `plan_calendar` с жёсткой пятёркой |

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| U1 | планировщик | чтобы analyze и build считали одну и ту же свободную ёмкость | не получать 422 после зелёного экрана |
| U2 | мастер смены | чтобы повторное / параллельное завершение дня не списывало плиты дважды | учёт КП совпадал с календарём |
| U3 | планировщик | чтобы при ошибке резерва СГП план не сохранялся «наполовину» | склад и план не расходились |
| U4 | двое в двух вкладках | чтобы устаревшая `version` не применяла КП-мутации | не терять чужие правки |
| U5 | любой клиент API | чтобы огромный диапазон дат не клал сервер | build/analyze оставались доступны |
| U6 | менеджер | чтобы график поставки видел ту же ёмкость дня, что цех | сроки не считались от «всегда 5» |

### Reframe → success criteria

| Требование | Конкретный критерий |
|------------|---------------------|
| «Analyze = persist» | День с occupancy 3 при max 5: fill `tracks=3` → analyze 200; `tracks=4` → analyze **и** build 422 с текстом про свободные слоты |
| «Нет второго списания» | Второй `POST .../complete` на уже completed день → 409, `completed_plates` не растут |
| «Обрыв после КП, до флага дня» | Искусственный fail `mark_day_completed` после успешного move → ROLLBACK КП; день не completed; retry успешен |
| «SGP fail не оставляет план» | `reserve_on_conn` бросает → плана нет в repo, плиты не «в плане», резервов СГП нет |
| «Version на complete» | `expected_version` stale → 409, КП не списаны |
| «Version на delete track» | query `expected_version` stale → 409, дорожка на месте |
| «Диапазон дат» | `GET /day-capacity?from=0001-01-01&to=9999-12-31` → 422 |
| «Fill ISO + горизонт» | `fill_targets.date=not-a-date` или `9999-12-31` → 422 |
| «Schema cap» | `tracks=6` / `max_tracks=6` → 422 на границе Pydantic |
| «Candidates limit» | `limit=10**9` → 422 |
| «Path date» | `/days/foo/complete` → 422 |
| «Оценка дней» | 38.3 м, 1 дорожка/день → `estimated_days === 1` на FE (ceil) |
| «Occupancy max» | день с override 3 → `max_per_day` в occupancy **или** per-day map отражает 3, не 5 |
| «Ошибка подложек» | `ProductionSubstrateError` → `analysis_meta.optimization_status="error"` + непустой `error_message`; UI показывает предупреждение |
| «График поставки» | override дня 0 или 3 виден в calendar-данных, которые использует delivery_schedule (не hard-coded 5) |

---

## 2. Tech Stack

Без изменений относительно baseline:

- Backend: Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`)
- Frontend: React + Vite + TypeScript (wizard analyze error, estimate formula, API `expected_version`)
- Транзакции: существующий паттерн `_external_conn` (`commit_plan_plates` / `PlateCompletionService` / `SgpService.reserve_on_conn`)
- Тесты: `pytest tests/`, vitest на FE estimate / wizard

---

## 3. Commands

```bash
# Целевой backend-набор (после реализации; имена файлов могут появиться новые)
.venv/bin/pytest \
  tests/test_production_capacity.py \
  tests/test_production_capacity_service.py \
  tests/test_core_production_planning.py \
  tests/test_production_api_integration.py \
  tests/test_production_completion_service.py \
  tests/test_plan_consistency.py \
  tests/test_plan_repository.py \
  tests/test_production_planning_service.py \
  -q

# Frontend
cd frontend && npm test -- --run \
  src/features/production/lib/productionEstimate.test.ts \
  src/features/production/hooks/useCreatePlanWizardState.test.ts \
  src/features/production/components/create-plan-wizard/SubstrateRecommendationsBlock.test.tsx

# Dev
uvicorn app.main:app --reload
cd frontend && npm run dev
```

---

## 4. Project Structure (затрагиваемые пути)

```
core/production/capacity.py              → единый validate_fill_targets(+ occupancy);
                                           occupancy-aware free slots; опционально max span helpers
core/production/planning.py              → удалить свою копию validate_fill_targets;
                                           persist вызывает shared; при SGP — не здесь, а в app service
core/production/__init__.py              → реэкспорт одного валидатора

app/schemas/production.py                → FillTargetItem.date ISO + горизонт;
                                           tracks/tracks_count/max_tracks le=TRACKS_PER_DAY_HARD_CAP;
                                           CompleteProductionDayRequest.expected_version;
                                           AnalysisMetaItem.error_message;
                                           path-даты через date / validator
app/api/v1/endpoints/production.py       → max span day-capacity; Query(ge,le) candidates;
                                           expected_version query на DELETE track;
                                           ISO path dates; не detail=str(exc) на новых ветках
                                           (существующий S5 целиком — out of scope)

app/services/production_service.py       → analyze: только shared validate (с occupancy);
                                           не глотать SubstrateError без error_message;
                                           occupancy max из capacity map (A11);
                                           complete_day: version + одна транзакция;
                                           build_plan_from_filters: одна транзакция с SGP
app/services/production_capacity_service.py → прокинуть occupancy в core validate
app/services/production_completion_service.py → guard completed; skip write_off_completed;
                                           mark_day_completed в той же conn
app/repositories/plan_repository.py      → mark_day_completed / save с _external_conn
app/services/plan_distribution_service.py → прокинуть expected_version с API (уже есть параметр)
app/services/delivery_schedule_service.py → calendar из production port, не plan_calendar
app/planning/plan_calendar.py            → не удаляем; delivery больше не единственный живой клиент

frontend/src/features/production/api/productionApi.ts     → уже шлёт version; сверить query/body
frontend/src/features/production/types/production.ts      → error_message; CompleteDay expected_version
frontend/src/features/production/lib/productionEstimate.ts → ceil; поправить тест
frontend/src/features/production/components/create-plan-wizard/*
                                                           → предупреждение при optimization_status=error
frontend/src/features/production/components/MonthCalendarGrid.tsx
                                                           → cap 5 оставить; не расходиться с schema

tests/test_production_fill_integrity.py  → NEW: analyze vs build occupancy; complete atomic;
                                           build+SGP rollback; expected_version 409
```

Не трогаем: `core/production/planning.py` split на load/optimize/persist, wizard hook split, CPU worker, rate limit, `POST /plans` legacy payload, bot_archived.

---

## 5. Code Style

Существующий паттерн транзакции:

```python
own_conn = _external_conn is None
if own_conn:
    conn = self._connect()
else:
    conn = _external_conn
try:
    # mutations
    if own_conn:
        conn.commit()
except Exception:
    if own_conn:
        conn.rollback()
    raise
finally:
    if own_conn:
        conn.close()
```

- Доменные ошибки: `PlanBuildError` / `ProductionCompletionError` / `ProductionAnalyzeBadRequest` / `PlanVersionConflict` — маппинг на HTTP как сейчас (422 / 409).
- Сообщения пользователю на русском, со свободными слотами («На 2026-09-10 свободно 2 дорожки, запрошено 3»).
- Hard cap импортировать из `core.production.capacity.TRACKS_PER_DAY_HARD_CAP`, не копировать `5` в schema.
- Не логировать PII; `logger.exception` при rollback complete/build.

---

## 6. Testing Strategy

| Уровень | Что | Где |
|---------|-----|-----|
| Unit | `validate_fill_targets` с occupancy; ceil estimate; schema 422 на cap/date/span | `tests/test_production_capacity.py`, pydantic/API, vitest estimate |
| Integration (tmp sqlite) | analyze OK / build fail при занятости; complete 409 повтор; rollback KP если mark_day падает; SGP fail откатывает план; stale version не списывает | NEW `tests/test_production_fill_integrity.py` + доп. к `test_production_api_integration.py` |
| Регрессия | существующие complete_day / planning / capacity тесты зелёные | список в Commands |
| FE | пустые подложки + `error_message` → Alert, не «нет рекомендаций» | wizard / SubstrateRecommendationsBlock test |

**Q6 из аудита закрывается этими тестами до рефакторинга валидаторов** — сначала красные тесты на текущее поведение A1 (analyze OK / build 422), затем правка валидатора.

Покрытие % не цель. Цель: три critical path не регрессируют.

---

## 7. Boundaries

**Always:**
- Тесты из Commands зелёные перед «готово»
- Один валидатор fill, occupancy на analyze и persist
- Одна sqlite-транзакция на complete (КП + флаг дня) и на build+SGP
- Входные лимиты на даты/cap/limit на schema/Query слое
- Минимальный diff: не «причесывать» соседние сервисы

**Ask first:**
- Менять схему SQLite (не планируется)
- Делать `expected_version` обязательным (сейчас optional)
- Идемпотентный 200 вместо 409 на повторный complete
- Расширять горизонт fill / span не 366
- Трогать `POST /plans` (legacy create)
- Помечать findings FIXED в audit-файле

**Never:**
- Нарезать `planning.py` / completion / wizard hook «заодно»
- Выносить optimize в job/worker
- Rate limit / Redis
- Менять `bot_archived`
- Коммитить секреты
- Удалять падающие тесты без замены

---

## Waves (порядок реализации после PLAN)

Спека одна; реализация — три вертикальных среза. Не начинать волну 2, пока волна 1 не зелёная.

### Wave 1 — вход и один валидатор (низкий риск)

S4, S8, S9, Q5 (+ Q4 piggyback schema), A1/Q1, Q6 тесты A1.

### Wave 2 — целостность учёта

S1/Q7 (`expected_version` на complete + remove_track), S2 (guard + skip snapshot + одна tx), A3/S3 (одна tx или компенсация, если `_external_conn` на commit_plan_plates не влезает в лимит файлов — **предпочтение: одна tx**).

### Wave 3 — правда в UI/смежных модулях

Q3 (ceil), A11 (occupancy max), Q9 (`error_message` + Alert), A4 узкий (delivery_schedule → production calendar).

---

## Decisions to lock (предлагаю)

| # | Решение | Альтернатива (отклоняем, пока не скажете иначе) |
|---|---------|------------------------------------------------|
| D1 | Scope = волны 1–3 выше | **Зафиксировано 2026-08-14:** пользователь выбрал полный обязательный срез |
| D2 | Повторный complete → **409** | **Зафиксировано в плане** |
| D3 | `expected_version` **optional**, но если передан — соблюдаем | **Зафиксировано в плане** |
| D4 | S2 = одна sqlite tx; A3 = compensating delete плана+плит | Одна tx на persist+SGP — follow-up (XL через PlanPersistPort) |
| D5 | ISO + span 366 + min_fill ≤ today+366; **без пола −30** | **Зафиксировано в плане** |
| D6 | Оценка = **ceil** | **Зафиксировано в плане** |
| D7 | A4 = смена импорта в delivery_schedule | **Зафиксировано в плане** |
| D8 | Три последовательных PR по волнам | Одна ветка до конца |
| D9 | Текущий `plita.db` — тестовые данные, не боевой учёт | Dual-read / миграция старых payload планов |

---

## Open Questions

Блокирующих нет. IMPLEMENT после ревью плана. Если нужна одна sqlite-транзакция и на build+SGP — сказать до Task 10.

---

## Out of scope (Not Doing)

Из аудита **намеренно не берём:**

- A2, A5, A7, A9, A10, A12 — god-модули / DI / router split  
- A6, S7 — CPU worker, rate limit  
- S5 целиком (массовая замена `str(exc)`), S6 (legacy POST /plans), S10 (work-calendar max items)  
- S11–S14, Q8, Q10–Q14, A13–A15, Q15–Q19  
- Пометка audit FIXED  
- Новые фичи планирования / подложек  
- Миграция/сохранение рядов текущего `plita.db` (тестовые данные)  

---

## Success Criteria (сводка «готово»)

1. Нет второй функции `validate_fill_targets` в `planning.py`; occupancy участвует в analyze и persist.
2. Complete дня: guard completed, skip snapshot, КП+флаг в одной транзакции, stale version → 409 без списания.
3. Build с `sgp_reservations`: ошибка резерва не оставляет план/пометку плит.
4. API отвергает гигантские диапазоны дат, не-ISO path, `tracks>5`, безлимитный `limit`.
5. FE оценка дней совпадает с `ceil`; ошибка анализа подложек видна; occupancy/delivery видят override ёмкости.
6. Команды из §3 зелёные. God-файлы не разрезаны «заодно».
