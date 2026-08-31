# Brief для оркестратора: стабилизация Производство (аудит 2026-08-12)

> Вставь это сообщение **целиком** в новое окно агента (Agent mode).  
> Спека и план уже утверждены человеком (2026-08-14). **Не** начинать SDD/interview заново.

---

```
/orchestrate execute без подтверждения плана

Задача: реализовать утверждённый срез аудита модуля Производство.

## Источники правды (читать первым делом)

1. Спека: ai_docs/specs/stabilizaciya-production-audit-2026-08-12.md
2. План задач: ai_docs/develop/plans/2026-08-14-stabilizaciya-production-audit.md
3. Аудит (контекст, не scope): ai_docs/develop/audits/2026-08-12-production-module-audit.md
4. Скиллы: project-shishov, orchestration, incremental-implementation, test-driven-development
5. Backend: pytest tests/ из корня с .venv; frontend: npm test / tsc из frontend/

План УЖЕ содержит T1–T15 с acceptance, files, dependsOn, checkpoints. Не вызывать planner заново. Не переписывать спеку/план, кроме статусов задач (pending → completed) и короткого implementation report в конце.

Workspace: создай .cursor/workspace/active/orch-2026-08-14-production-audit/ с progress.json, tasks.json, links.json, указывающими на план выше. Resume с первой невыполненной задачи.

## Решения D1–D9 — не переоткрывать

- D1: волны 1–3 целиком (не только P0)
- D2: повторный complete → HTTP 409 day_already_completed (не идемпотентный 200). На FE показать «день уже завершён», не «ошибка списания»
- D3: expected_version optional; если передан — соблюдать (stale → 409, КП не списывать)
- D4: complete_day = одна sqlite-транзакция (КП + mark_day_completed через _external_conn). Build+SGP = компенсация: return_plan_plates_to_production + delete плана. НЕ протаскивать conn через PlanPersistPort/persist
- D5: ISO-даты; span from..to и min..max fill ≤ 366 дней; min_fill ≤ today+366. НЕТ пола «сегодня−30»
- D6: FE estimate = Math.ceil (тест 38.3 м / 1 дорожка → 1 день, не 2)
- D7: delivery_schedule._load_occupancy → PlanDistributionService.get_global_calendar_info. plan_calendar.py НЕ удалять
- D8: логически три волны/PR. В этой сессии реализуй ВСЕ волны на текущей ветке последовательно. Коммиты/gh pr create — только если пользователь явно попросит; по умолчанию НЕ коммитить
- D9: текущий plita.db = тестовые данные. Нет миграций dual-read старых payload. Pytest-фикстуры и контракт FE обязательны

## Порядок (DAG)

Волна 1 / checkpoint, затем волна 2, затем волна 3. Не начинать волну N+1, пока pytest/vitest checkpoint N красный.

T1 red tests W1
 → T2 API guards (параллельно T3 после T1, разные файлы)
 → T3 one validator occupancy
 → T4 analyze + occupancy
Checkpoint W1: analyze/build один контракт свободных слотов; span/cap/limit/ISO → 422

T5 red tests W2
 → T6 expected_version (∥ T7, T8, T10)
 → T7 complete guards
 → T8 PlanRepository _external_conn
 → T9 complete one tx (после T6+T7+T8)
 → T10 SGP compensate (после T5; не зависит от T8)
Checkpoint W2: повторный complete и stale version не портят КП; fail SGP не оставляет план

T11 ceil (∥ T12, T13, T14)
T12 occupancy max_by_day
T13 substrate error_message + wizard Alert
T14 delivery calendar
T15 regression gate spec §3

После каждой задачи — verification из плана. Красные тесты T1/T5 писать ДО фикса соответствующего кода (TDD).

## Never

- Нарезать planning.py / production_completion_service.py / useCreatePlanWizardState.ts / толстый router «заодно»
- CPU worker, rate limit, Redis
- Трогать bot_archived, POST /plans legacy payload целиком
- Удалять plan_calendar.py
- Делать expected_version обязательным
- Пол дат сегодня−30
- Одна sqlite tx через persist+SGP (это follow-up, не этот срез)
- Помечать audit FIXED
- Коммитить секреты; git commit / push / pr без явной просьбы
- Удалять падающие тесты без замены
- Менять схему SQLite

## Команды верификации (минимум)

Волна 1:
.venv/bin/pytest tests/test_production_fill_integrity.py tests/test_production_capacity.py tests/test_production_capacity_service.py tests/test_core_production_planning.py tests/test_production_api_integration.py tests/test_production_podlozhki_e2e.py -q

Волна 2: плюс
.venv/bin/pytest tests/test_production_completion_service.py tests/test_plan_consistency.py tests/test_plan_repository.py tests/test_production_planning_service.py -q

Волна 3: плюс
.venv/bin/pytest tests/test_delivery_schedule_service.py tests/test_delivery_schedule_endpoints.py -q
cd frontend && npm test -- --run src/features/production/lib/productionEstimate.test.ts src/features/production/hooks/useCreatePlanWizardState.test.ts src/features/production/components/create-plan-wizard/SubstrateRecommendationsBlock.test.tsx

Финал T15: весь набор из спеки §3.

## Definition of done

Критерии спеки «Success Criteria». God-файлы не разрезаны. Краткий report: ai_docs/develop/reports/2026-08-14-stabilizaciya-production-audit-implementation.md

Если задача упирается в противоречие с D1–D9 — остановись и напиши человеку. Не угадывай продукт заново.
```

---

## Как запустить в другом окне

1. Новое окно Cursor Chat, **Agent mode** (не Ask).
2. Прикрепи `@ai_docs/develop/plans/2026-08-14-orchestrator-brief.md` и `@ai_docs/develop/plans/2026-08-14-stabilizaciya-production-audit.md` и `@ai_docs/specs/stabilizaciya-production-audit-2026-08-12.md`.
3. Вставь блок между тройными кавычками выше (от `/orchestrate execute` до конца).
4. Не отвечай «ок» в том окне повторно на решения D1–D9 — они уже закрыты.

Это окно можно не держать открытым для реализации. Сюда имеет смысл вернуться, только если оркестратор остановится на противоречии или красном checkpoint.
