# Отчёт аудита: модуль Производство

**Дата**: 2026-08-12  
**Область**: модуль Производство (production) — `app/services/production*`, `app/api/v1/endpoints/production.py`, `core/production/`, `frontend/src/features/production/`, schemas, пересечения с planning  
**Аудиторы**: senior-reviewer + security-auditor + reviewer  
**Модели**: cursor-grok-4.5-high-fast

---

## Краткое резюме

**Оценка здоровья**: **4.0 / 10**

Расчёт: старт 10; Critical 1 → −2; High 16 × 0.5 (потолок −3); Medium 19 × 0.1 (потолок −1) → **4.0**

| Серьёзность | Архитектура | Безопасность | Качество кода | Итого |
|-------------|-------------|--------------|---------------|-------|
| Critical    | 1           | 0            | 0             | **1** |
| High        | 5           | 4            | 7             | **16** |
| Medium      | 6           | 6            | 7             | **19** |
| Low         | 3           | 4            | 5             | **12** |

**Рекомендация**: устранить Critical **A1** и High security **S1–S4** до следующего релиза production-критичных фич; остальное — sprint/backlog.

Модуль производства несёт реальный риск рассинхрона учёта и плана: два контракта валидации ёмкости (analyze «OK» / build 422), мёртвый optimistic lock на клиенте, неатомарные цепочки build+SGP и complete_day, а также DoS через неограниченные диапазоны дат на аутентифицированных CPU-эндпоинтах. Параллельно копится архитектурный долг — god-модули `planning.py` / completion / wizard hook и дрейф констант hard-cap.

**Дубликаты (консолидированы ниже):**
- **A1 ≈ Q1** — два `validate_fill_targets` (архитектура + именование/контракт качества)
- **A3 ≈ S3** — неатомарный build + SGP reserve (архитектура + целостность данных)
- **S1 ≈ Q7** — `expected_version` игнорируется (безопасность + drift контракта FE/BE)
- **A8 ≈ Q4** — дрейф констант hard-cap / длины дорожки

---

## Критические проблемы

### [A1] / [Q1] Dual `validate_fill_targets`: capacity vs planning free slots

**Категория**: Архитектура + Качество кода (одна корневая проблема)  
**Где**: `core/production/capacity.py` (`validate_fill_targets`, лимит = `day_max`, без occupancy) vs `core/production/planning.py` (лимит = `day_max − occupied`); analyze → `ProductionCapacityService` / `production_service.analyze_substrates`; persist → `planning.persist`; hard cap ещё раз в `clamp_day_max` / `TRACKS_PER_DAY_HARD_CAP` / FE `MonthCalendarGrid`  
**Влияние**: analyze/UI могут сказать «OK», а `build` упасть на занятости; одинаковое имя скрывает разный контракт; при дрейфе констант снова возможен обход hard cap на одном из путей.  
**Исправление**: один доменный валидатор в `core.production.capacity` с явными режимами (`against_day_max` / `against_free_slots`) или переименование (`validate_fill_against_day_max` / `validate_fill_against_free_slots`); analyze и persist вызывают его с occupancy; planning-копию удалить; hard cap — только `TRACKS_PER_DAY_HARD_CAP` (+ FE из API/shared const).

---

## Высокий приоритет

### [S1] / [Q7] `expected_version` заявлен на клиенте, сервер игнорирует (TOCTOU)

**Категория**: Безопасность + Качество (контракт FE/BE)  
**Где**: `frontend/.../productionApi.ts` (`expected_version` в body/query) vs `CompleteProductionDayRequest` (поля нет → Pydantic drop); `DELETE .../tracks/{track_index}` (query не принимается); `remove_track_from_plan(..., expected_version=...)` с API не вызывается  
**Влияние**: параллельные mutating-запросы читают одну `version`; побочные эффекты (возврат/списание в KP) уже применены, а `save` даёт 409 или last-write-wins. Клиент думает, что optimistic lock есть — его нет.  
**Исправление**: принять `expected_version` в schema/query; прокинуть в `mark_day_completed` / `remove_track`; при конфликте **не** коммитить KP-мутации (одна транзакция / saga с компенсацией); либо временно убрать поле с FE до фикса.

### [S2] `complete_day`: гонка / нет запрета повторного завершения + неатомарность KP ↔ plan

**Категория**: Безопасность (связано с A5)  
**Где**: `production_completion_service.send_to_sgp` (нет проверки `day.completed`); `_collect_plates_by_kp` не фильтрует `write_off_completed`; `production_service.complete_day` → списание, затем отдельно `plan_repository.mark_day_completed`  
**Влияние**: гонка двух complete / сбой между шагами → списанные плиты при «незавершённом» дне, повторные попытки на snapshot-позициях, рассинхрон учёта и плана.  
**Исправление**: guard `if day.completed → 409/422`; пропускать уже списанные snapshot; одна транзакция (или outbox) для KP + флаг дня; идемпотентный ключ.

### [A3] / [S3] Неатомарный build plan + SGP reserve

**Категория**: Архитектура + Безопасность (одна корневая проблема)  
**Где**: `ProductionService.build_plan_from_filters` — `planning_service.build_plan` (commit плит + save плана), затем отдельный `sqlite3` + `SgpService.reserve_on_conn`, затем снова `plan_repository.save_plan`  
**Влияние**: при ошибке SGP остаётся сохранённый план без/с частичными резервами; склад и план расходятся («потерянные» объёмы); компенсирующего отката плана нет.  
**Исправление**: единый use-case / транзакционный saga: резерв в той же транзакции или явный compensating delete/rollback; `save_plan` с `expected_version`.

### [S4] DoS через несвязанные даты / диапазоны (аутентифицированный)

**Категория**: Безопасность  
**Где**: `GET /day-capacity` (`_date_range_inclusive` без max span); `analyze_substrates` цикл `while cursor < min_fill` при далёком `fill_targets.date`; `FillTargetItem.date: str` без ISO/горизонта  
**Влияние**: `from=0001-01-01&to=9999-12-31` или fill на 9999 → миллионы дат, память/CPU, блокировка semaphore (`CPU_BOUND_MAX_CONCURRENT=2`) для всех.  
**Исправление**: max span (например ≤366 дней); ISO + горизонт на `FillTargetItem.date`; лимит длины `fill_targets` / `capacity_dates`.

### [A2] God-модуль pipeline в `core/production/planning.py` (~826 строк)

**Категория**: Архитектура  
**Где**: `core/production/planning.py` — `load` / `optimize` / `persist` + KP-load, layout/rescue/fallback, fill_targets, commit/rollback  
**Влияние**: любой change ёмкости/раскладки/persist рискует регрессией всего build; плохо тестируется и ревьюится.  
**Исправление**: разрезать на `load.py` / `optimize.py` / `persist.py` (+ helpers fill_targets); `planning` — тонкий фасад.

### [A4] Два источника «правды» календаря/ёмкости дня

**Категория**: Архитектура  
**Где**: production UI — `PlanDistributionService.get_global_calendar_info` (читает `ProductionCapacityService`); legacy `app/planning/plan_calendar.py` / `plan_manager` — жёсткий `MAX_TRACKS_PER_DAY`, без overrides; `delivery_schedule_service` тянет legacy  
**Влияние**: день с override 3/0 в production и «5» в графике поставки/legacy — рассинхрон планирования и сроков.  
**Исправление**: один calendar port на `PlanRepository` + day_capacity; legacy `plan_calendar` deprecated/прокси.

### [A5] `ProductionCompletionService` — god + прямой sqlite

**Категория**: Архитектура  
**Где**: `app/services/production_completion_service.py` (~669), прямой `sqlite3`, SQL, СГП/остатки/списание  
**Влияние**: обход repository-контрактов, сложные транзакции в application-слое, высокий риск partial updates (усиливает S2).  
**Исправление**: SQL → repositories; сервис оркестрирует use-case; разделить complete-day / rests / SGP write.

### [A6] Масштабирование CPU-bound только семафором `run_cpu_bound`

**Категория**: Архитектура  
**Где**: `production.py` — `build` / `analyze-substrates` / `create_plan`; внутри — полный optimize (`planning.optimize`, cascading cuts)  
**Влияние**: под нагрузкой очередь в event-loop threads, длинные timeouts, analyze и build дерутся за один пул; горизонтально не масштабируется.  
**Исправление**: оставить `run_cpu_bound` для API; вынести тяжёлый optimize в job/worker при росте concurrency; кэш/инвалидация для UI-тиков.

### [Q2] God-hook визарда (~715 строк)

**Категория**: Качество кода  
**Где**: `frontend/src/features/production/hooks/useCreatePlanWizardState.ts`  
**Влияние**: selection/qty/SGP/analyze/capacity options/`handleSubmit`/basket в одном хуке; высокий coupling и хрупкие тесты (отдельный FE-hotspot от A2/A5).  
**Исправление**: разрезать на `useFillTargets`, `usePlateSelection`, `useAnalyzeSubstrates`, `useBuildPlanSubmit`.

### [Q3] FE/BE формулы оценки дорожек расходятся (ceil vs round)

**Категория**: Качество кода  
**Где**: `frontend/.../productionEstimate.ts` (`Math.round(x + 0.5)`) vs `core/production/capacity.py` (`math.ceil`); баг закреплён тестом (`38.3m → estimated_days: 2` при 1 track/day)  
**Влияние**: систематическое завышение оценки на FE относительно BE; ложные ожидания по срокам в UI.  
**Исправление**: единый `ceil`; FE — та же семантика; поправить тесты.

### [Q4] / [A8] Константы hard-cap / длины дорожки размазаны

**Категория**: Качество + Архитектура  
**Где**: `MonthCalendarGrid.tsx`, `core/production/capacity.py` (`TRACKS_PER_DAY_HARD_CAP`), `core/production/dto.py`, `app/planning/plan_storage.py` (`MAX_TRACKS_PER_DAY`), `productionEstimate.ts`, `core/production_capacity.py`  
**Влияние**: следующий hard-cap fix снова разъедется между FE/BE/legacy.  
**Исправление**: один int source of truth в core; app/FE — импорт или контракт API (`max_per_day` / `track_length_m`).

### [Q5] Schema `le=50` vs hard cap `5`

**Категория**: Качество кода  
**Где**: `app/schemas/production.py` (`tracks`/`tracks_count` до 50); `SaveDayCapacityRequest.max_tracks: ge=0` без `le`  
**Влияние**: Pydantic принимает завышенные значения; отсечка позже в core/repo → отложенная/непонятная 422, путаница клиентов.  
**Исправление**: `le=TRACKS_PER_DAY_HARD_CAP` на schema-слое.

### [Q6] Пробелы тестов на critical paths

**Категория**: Качество кода  
**Где**: capacity options → build; analyze OK / build fail при занятых слотах; `complete_day` + `expected_version` e2e  
**Влияние**: регрессии A1/S1/S2 легко проходят незамеченными; рефакторинг валидаторов без страховочной сетки.  
**Исправление**: 2–3 интеграционных теста на эти ветки до рефакторинга.

---

## Средний приоритет

### Архитектура

#### [A7] `ProductionService` как широкая фасад-оркестрация + ad-hoc DI

**Где**: `production_service.py` — plans/calendar/analyze/build/SGP/complete/remove_track; внутри `analyze_substrates` создаёт capacity/urgent/substrate вручную  
**Влияние**: нарушение DIP/тестируемости, скрытые зависимости.  
**Исправление**: инжектировать capacity/urgent/substrate в конструктор; фасад оставить тонким.

#### [A9] `PlanDistributionService` смешивает distribution, calendar, gantt, ports

**Где**: `plan_distribution_service.py` + адаптеры `PlanLoadAdapter`/`PlanPersistAdapter`  
**Влияние**: calendar capacity-логика рядом с layout persist; сложнее эволюционировать ёмкость отдельно.  
**Исправление**: вынести calendar/gantt и adapters из distribution.

#### [A10] FE god-state/UI в wizard/calendar

**Где**: `useCreatePlanWizardState.ts` (~715), `DayDrawer.tsx` (~555), `MonthCalendarGrid.tsx` (~453)  
**Влияние**: высокий coupling (пересекается с Q2).  
**Исправление**: разделить hooks: selection, analyze/capacity options, SGP close-from-warehouse.

#### [A11] `get_day_occupancy` отдаёт глобальный `max_per_day`, игнорируя overrides

**Где**: `ProductionService.get_day_occupancy` → `MAX_TRACKS_PER_DAY`  
**Влияние**: клиенты occupancy API видят hard default, не фактическую ёмкость дня.  
**Исправление**: per-day `max` из capacity map (как в calendar).

#### [A12] Толстый router production (~616 строк)

**Где**: `app/api/v1/endpoints/production.py` — plans, analyze, complete, documents, SGP, work calendar, day-capacity  
**Влияние**: смешение bounded contexts на одном API-модуле.  
**Исправление**: под-роутеры `plans` / `calendar-capacity` / `sgp` / `day-docs` при сохранении prefix.

### Безопасность

#### [S5] Утечка internals в `detail=str(exc)`

**Где**: `production.py` build/analyze/`save_day_capacity`; `message=str(exc)` для SGP/REST  
**Влияние**: сырые доменные тексты на клиенте (разведка); XSS низкий.  
**Исправление**: стабильные user-facing сообщения + `code`; детали только в лог.

#### [S6] `POST /plans`: неограниченный нетипизированный payload

**Где**: `CreatePlanRequest` — `all_tracks_list`, `orders_2d`, `optimization_result` без размера/схемы  
**Влияние**: DoS по JSON/CPU; произвольная структура в `production_plans.payload_json`.  
**Исправление**: жёсткая схема + `max_length`/`max_items`; либо deprecate в пользу `/plans/build`.

#### [S7] Нет rate limit на mutating/CPU endpoints

**Где**: `/plans/build`, `/analyze-substrates`, документы дня (см. A6)  
**Влияние**: роль `production` может исчерпать CPU/IO.  
**Исправление**: per-user rate limit / очередь задач.

#### [S8] `target_date` / `date` path без строгой валидации

**Где**: `/days/{target_date}`, documents, `remove_track`; `FileResponse` filename; `mkdtemp` prefix  
**Влияние**: не ISO → странные имена файлов / риск header/prefix abuse.  
**Исправление**: `date.fromisoformat` → 422; filename только из нормализованного ISO.

#### [S9] `GET /candidates?limit=` без верхней границы

**Где**: `production.py` `limit: int = 500`  
**Влияние**: `limit=10**9` → тяжёлый SELECT/ответ.  
**Исправление**: `Query(ge=1, le=500)`.

#### [S10] `PUT /work-calendar` без валидации дат и размера

**Где**: `SaveWorkCalendarRequest` → `save_raw` пишет JSON на диск  
**Влияние**: мусор/огромные списки → порча календаря, DoS на диск.  
**Исправление**: ISO-даты, max items, reject unknown keys.

### Качество кода

#### [Q8] Фрагментированная иерархия ошибок + дубли mapping в endpoints

**Где**: `production_*_service.py`, `endpoints/production.py`  
**Влияние**: копипаста 409/documents handlers; нет единого adapter-слоя.  
**Исправление**: `map_production_error(exc)` + один helper для day-documents.

#### [Q9] `ProductionSubstrateError` глотается в analyze

**Где**: `production_service.py` (~191–193)  
**Влияние**: `substrates=[]`, `optimization_status="error"` без явного сигнала UI — «нет рекомендаций» выглядит как успех.  
**Исправление**: логировать `exc`, отдавать `error_message` в `analysis_meta`, предупреждение на FE.

#### [Q10] Слабая типизация plan payloads (TS + Pydantic)

**Где**: `types/production.ts`, `schemas/production.py` (`CreatePlanRequest` с сырыми dict/list), `PlanDetailResponse` `extra="allow"`  
**Влияние**: ошибки только в runtime; ломает рефакторинг plan shape.  
**Исправление**: явные DTO; сузить `CreatePlanRequest` или пометить legacy.

#### [Q11] Новые complexity hotspots

**Где**: `calculate_capacity_deficit` (~115), `extract_substrate_recommendations` (~141), `collect_urgent_positions` (~92), `DayDrawer` / `SgpWarehouseView`  
**Влияние**: сложность растёт вне уже отмеченных god-модулей.  
**Исправление**: extract steps A/B/C options; разбить drawer/warehouse.

#### [Q12] DRY дат/календаря на FE

**Где**: `MonthCalendarGrid.tsx`, `GlobalCalendarView.tsx`, `calendarRange.ts`  
**Влияние**: дубли `formatISO`/`startOfMonth`; локальный рассинхрон workday-логики с BE (см. A4).  
**Исправление**: один `lib/calendarDates.ts`.

#### [Q13] Дубли `_to_date` / `_to_iso_date`

**Где**: `production_service.py`, `production_capacity_service.py`, `urgent.py`, `capacity.py`, `day_capacity_repository.py`  
**Влияние**: пять парсеров с разным поведением при invalid.  
**Исправление**: `core/production/dates.py`.

#### [Q14] Повторная валидация fill_targets в analyze

**Где**: `production_service.py` (~138–161)  
**Влияние**: ручной парсинг + снова `capacity_service.validate_fill_targets` → расхождение сообщений.  
**Исправление**: только вызов shared validate.

---

## Низкий приоритет / предложения

### Архитектура

- **[A13]** Lazy-import `ProductionCapacityService` в `ProductionPlanningService.build_plan` без реального цикла — шум, скрытая зависимость. Fix: обычный import + инъекция.
- **[A14]** Лишняя прослойка ошибок (`PlanBuildError` → `ProductionPlanBuildError` / …) — дубли маппинга в endpoint. Fix: единый app-level error mapping у HTTP-границы.
- **[A15]** Premature abstraction почти нет — ports уместны; thin capacity/urgent/substrate — нормальный SRP. Риск в недорезанных god-модулях (A2/A5/A9), не в лишних абстракциях.

### Безопасность

- **[S11]** «IDOR» в пределах роли — нет per-plan ACL (для цеха обычно ок). Fix: аудит-лог actor; при необходимости object-level authz.
- **[S12]** AuthZ на фронте не замена серверу — API уже с `require_roles`; риска обхода сервера нет. Fix: без изменений; не полагаться только на UI.
- **[S13]** XSS в production UI — `dangerouslySetInnerHTML` не найден; text-only. Fix: сохранять text-only.
- **[S14]** Секреты / SQLi — параметризованный SQL; секретов в модуле нет. Fix: —.

### Качество кода

- **[Q15]** Deprecated `onSelectDate` всё ещё wired в `MonthCalendarGrid` — удалить после grep.
- **[Q16]** Смешение EN/RU в user-facing errors (`"Plan not found"` vs русские) — единый язык (RU) + коды.
- **[Q17]** Устаревший комментарий «A10 / WP3» в module docstring planning service — описать роль без ticket IDs.
- **[Q18]** `CreatePlanRequest` / legacy create — слабо типизированный «мешок»; `@deprecated` в OpenAPI или удалить.
- **[Q19]** `eslint-disable` exhaustive-deps на analyze effect — стабилизировать через `useEffectEvent` / явный deps.

---

## Матрица приоритетов

| ID | Проблема | Severity | Effort | Priority |
|----|----------|----------|--------|----------|
| A1/Q1 | Dual `validate_fill_targets` (capacity vs free slots) | Critical | Medium | **P0** — сейчас |
| S1/Q7 | `expected_version` игнорируется (TOCTOU) | High | Medium | **P0** — до релиза |
| S2 | `complete_day` race / nonatomic KP↔plan | High | High | **P0** — до релиза |
| A3/S3 | Неатомарный build + SGP reserve | High | High | **P0** — до релиза |
| S4 | DoS date ranges / fill horizon | High | Low | **P0** — до релиза |
| A2 | God `planning.py` | High | High | **P1** — sprint |
| A4 | Dual calendar / capacity truth | High | Medium | **P1** — sprint |
| A5 | Completion god + sqlite | High | High | **P1** — sprint |
| A6 | CPU только semaphore | High | Medium | **P1** — sprint |
| Q2 | God wizard hook | High | Medium | **P1** — sprint |
| Q3 | FE/BE estimate ceil vs round | High | Low | **P1** — sprint |
| Q4/A8 | Drift констант hard-cap | High | Low | **P1** — sprint |
| Q5 | Schema `le=50` vs hardcap 5 | High | Low | **P1** — sprint |
| Q6 | Test gaps critical paths | High | Medium | **P1** — sprint |
| A7–A12 | Фасад DI, distribution, occupancy, router… | Medium | Medium–High | **P2** — next sprint |
| S5–S10 | Error leak, payload, rate limit, dates, candidates, calendar | Medium | Low–Medium | **P2** — next sprint |
| Q8–Q14 | Errors, swallow, types, hotspots, DRY dates | Medium | Low–Medium | **P2** — next sprint |
| A13–A15, S11–S14, Q15–Q19 | Low / hygiene / notes | Low | Low | Backlog |

---

## Следующие шаги

1. **Немедленно (P0, до production-критичных фич)**  
   - Унифицировать/переименовать валидаторы fill (**A1/Q1**).  
   - Прокинуть и соблюдать `expected_version` (**S1/Q7**).  
   - Сделать идемпотентным и атомарным `complete_day` (**S2**).  
   - Атомарный saga build+SGP или компенсация (**A3/S3**).  
   - Ограничить span дат / горизонт fill (**S4**).

2. **Этот спринт (P1)**  
   - Разрезать `planning.py` и completion (**A2**, **A5**); единый calendar port (**A4**).  
   - Выровнять формулы/константы/schema (**Q3–Q5**, **A8**); разрезать wizard hook (**Q2**).  
   - Интеграционные тесты на critical paths (**Q6**); план по CPU/worker (**A6**).

3. **Следующий спринт (P2)**  
   - Medium: A7–A12, S5–S10, Q8–Q14 (валидация входов, rate limit, DRY, error mapping).

4. **Backlog**  
   - Low: A13–A15, S11–S14, Q15–Q19.

---

## Примечание по remediation (Phase 5)

Автоматическое исправление (Phase 5 audit-workflow) **не запускалось**.  
Запускать только после явного подтверждения пользователя («Start remediation? y/n»).  
Код в рамках этого аудита **не менялся** — только отчёт.
