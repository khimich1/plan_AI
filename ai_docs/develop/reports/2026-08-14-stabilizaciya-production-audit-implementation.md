# Report: Стабилизация модуля Производство (аудит 2026-08-12)

**Date:** 2026-08-14  
**Orchestration:** `orch-2026-08-14-production-audit`  
**Status:** ✅ Completed (волны 1–3, T1–T15)  
**Audit FIXED:** ❌ нет (отдельный follow-up)  
**Commits:** нет (по brief)

**Spec:** [`../../specs/stabilizaciya-production-audit-2026-08-12.md`](../../specs/stabilizaciya-production-audit-2026-08-12.md)  
**Plan:** [`../plans/2026-08-14-stabilizaciya-production-audit.md`](../plans/2026-08-14-stabilizaciya-production-audit.md)  
**Audit (контекст):** [`../audits/2026-08-12-production-module-audit.md`](../audits/2026-08-12-production-module-audit.md)  
**Brief:** [`../plans/2026-08-14-orchestrator-brief.md`](../plans/2026-08-14-orchestrator-brief.md)

## Summary

Реализованы волны 1–3 стабилизации по утверждённой спеке/плану: один occupancy-aware валидатор fill, входные лимиты API, живой `expected_version`, атомарный `complete_day`, компенсация build+СГП, правда UI/графика поставки. Решения **D1–D9 не переоткрывались**.

## Decisions (locked)

D1–D9 из плана соблюдены: scope = волны 1–3; повторный complete → 409; `expected_version` optional; S2 = одна sqlite tx / A3 = compensating delete; ISO + span ≤366 без пола −30; FE estimate = ceil; delivery через `PlanDistributionService` (`plan_calendar` не удалён); три логических волны без коммитов; без dual-read миграций.

**Уточнение к тексту плана (опечатка):** свободные слоты = `max − occupied`. При occupancy=3, max=5 допустимо `tracks ≤ 2`, не `tracks=3` (в Task 1 acceptance ранее фигурировало «tracks=3 OK» — это неверно относительно формулы `day_free`).

## По волнам

### Волна 1 — вход и один валидатор (T1–T4)

- Красные тесты fill/API guards → `tests/test_production_fill_integrity.py`
- Schema/Query: ISO-даты, span ≤366, `tracks`/`max_tracks` ≤5, `limit` ≤500 → `app/schemas/production.py`, `app/api/v1/endpoints/production.py`
- Единый `validate_fill_targets(+occupancy)` в `core/production/capacity.py`; копия в `planning.py` удалена
- Analyze прокидывает occupancy → `production_capacity_service.py`, `production_service.py`

**Checkpoint W1:** analyze/build один контракт свободных слотов; span/cap/limit/ISO → 422.

### Волна 2 — целостность учёта (T5–T10)

- `expected_version` на complete и DELETE track → schemas / endpoints / `production_service.py`
- Guard `day.completed` + skip `write_off_completed` → `production_completion_service.py`
- `PlanRepository.save/create/mark_day_completed` + `_external_conn` → `plan_repository.py`
- `complete_day`: КП + mark в одной транзакции
- Fail SGP после build: `return_plan_plates_to_production` + `delete` плана (D4, без persist port)

**Checkpoint W2:** повторный complete / stale version не портят КП; fail SGP не оставляет план.

### Волна 3 — правда UI и график поставки (T11–T15)

- FE `estimateFromLengthM` = `Math.ceil` → `productionEstimate.ts` (+ тест 38.3→1)
- Occupancy API: `max_by_day` из capacity map → `production_service.py`, schemas, FE types
- `analysis_meta.error_message` + Alert в wizard → `production_service.py`, `useCreatePlanWizardState.ts`, `SubstrateRecommendationsBlock`
- `delivery_schedule._load_occupancy` → `PlanDistributionService.get_global_calendar_info` (`plan_calendar.py` на месте)
- Регрессионный gate (T15)

**Checkpoint W3 / Complete:** критерии спеки по срезу закрыты; audit не FIXED.

## Verification

| Suite | Result |
|-------|--------|
| Backend (spec §3 + fill_integrity / delivery / related) | **179 passed** |
| FE: estimate + wizard + SubstrateRecommendationsBlock | **26 passed** |
| `tsc -p tsconfig.app.json --noEmit` | clean |

## Out of scope (напоминание)

- Нарезка god-модулей (`planning.py`, completion, wizard, толстый router)
- CPU worker / rate limit / Redis
- Удаление `plan_calendar.py`
- Dual-read / миграции старых payload планов
- Протаскивание `_external_conn` через `PlanPersistPort` (follow-up A3)
- Пометка audit **FIXED**
- Git commit / PR (только по явной просьбе)

## Next steps

1. Human review / три PR по волнам (если нужно разнести историю).
2. После merge и smoke на стенде — отдельный follow-up: пометить findings в audit FIXED.
3. При инцидентах окна компенсации A3 — рассмотреть одну tx через persist port (вне этого среза).

## Review follow-up (2026-08-14)

**Critical (fixed):** `mark_day_completed` → `False` больше не приводит к `conn.commit()` списания КП. В `send_to_sgp` при `not day_completed` — `ProductionCompletionError` → rollback. Тест: `test_complete_day_mark_returns_false_rolls_back_kp`.
