# Report: Планирование от ёмкости — срочные + подложки (MVP)

**Date:** 2026-08-12  
**Orchestration:** `orch-2026-08-12-podlozhki`  
**Status:** ✅ Completed (Tasks 0–14)  
**Feature doc:** [`../features/planirovanie-po-srokam-podlozhki.md`](../features/planirovanie-po-srokam-podlozhki.md)  
**Plan:** [`../plans/2026-08-12-planirovanie-po-srokam-podlozhki.md`](../plans/2026-08-12-planirovanie-po-srokam-podlozhki.md)  
**Spec:** [`../../specs/planirovanie-po-srokam-podlozhki.md`](../../specs/planirovanie-po-srokam-podlozhki.md)  
**Phase 0:** [`./2026-08-12-podlozhki-phase0.md`](./2026-08-12-podlozhki-phase0.md)

## Summary

MVP «срочные по срокам + подложки из поздних КП» реализован end-to-end: per-day ёмкость, read-only анализ бэклога, UI в календаре и wizard, E2E-тест. Phase 0 (A1–A3) — go; на текущем бэклоге фактических cross-KP мэтчей 0, потенциал 1.

## What shipped

- **Ёмкость:** `day_capacity_override` + `GET/PUT /api/v1/production/day-capacity` + режим «Ёмкость» в календаре.
- **Анализ:** `POST /api/v1/production/analyze-substrates` → urgent + substrate recommendations + capacity deficit (read-only).
- **Wizard:** `UrgentPositionsBlock`, `SubstrateRecommendationsBlock`, `CapacityDeficitAlert`.
- **Инварианты:** рекомендации = преселектор; `fill_targets ≤ max_tracks`; финальный план через `plans/build`.

## Tasks 0–14

| Task | Name | Status | Notes |
|------|------|--------|-------|
| 0 | Phase 0 validation | ✅ | A1–A3 PASS; report phase0 |
| 1 | Day capacity schema + repo | ✅ | 9 tests |
| 2 | Core capacity domain | ✅ | 15 tests |
| 3 | `qty_remaining` | ✅ | 6 tests |
| 4 | Core urgent | ✅ | 15 tests |
| 5 | ProductionCapacityService | ✅ | 9 tests |
| 6 | ProductionUrgentService | ✅ | 9 tests |
| 7 | ProductionSubstrateService | ✅ | 11 tests (+ `core/production/substrate.py`) |
| 8 | API analyze-substrates | ✅ | AuthZ admin/production; security PASS_WITH_NITS |
| 9 | Calendar capacity mode | ✅ | FE |
| 10 | UrgentPositionsBlock | ✅ | FE |
| 11 | SubstrateRecommendationsBlock | ✅ | FE |
| 12 | CapacityDeficitAlert | ✅ | FE |
| 13 | E2E podlozhki | ✅ | 2 scenarios |
| 14 | Documentation | ✅ | Feature doc + this report |

**Checkpoints:** Phase 1–3 passed in orchestration (`phase3FrontendTestsPassed: 98`, build ok).

## Smoke (finalization)

```text
.venv/bin/python -m pytest tests/test_production_podlozhki_e2e.py tests/test_day_capacity_repository.py -q
→ 11 passed
```

Full `pytest tests/ -q` and human E2E UI walkthrough — not re-run in TASK-014; Phase 3 build/tests were green.

## Technical decisions

1. Analyze endpoint отделён от `plans/build` (read-only vs mutation).
2. Дефицит — `calculate_capacity_deficit` в `core/production/capacity.py`, не `delivery_schedule_check`.
3. Без кэша анализа; пересчёт по кнопке.
4. Feature-flag не вводился — cleanup не требуется.

## Known nits / follow-ups

- Ручная проверка UI (календарь → ёмкость → wizard → build) ещё не отмечена.
- Smoke `analyze-substrates` на живой `plita.db` с реальным оптимизатором — checkpoint Phase 2 остаётся открытым.
- На текущем бэклоге нет фактических secondary cross-KP мэтчей (Phase 0); ценность проявится на неполноширинных плитах.
- Многие ревью `APPROVE_WITH_NITS` — косметика/стиль, не блокеры.
- Вне MVP: `rests_unused` в finalize, экспорт рекомендаций XLSX/PDF, фоновый job при медленном прогоне, метрика A4 (доля принятых подложек).

## Key files

**Backend:** `core/kp_db_schema.py`, `core/production/{capacity,urgent,substrate}.py`, `app/repositories/day_capacity_repository.py`, `app/services/production_{capacity,urgent,substrate}_service.py`, `app/services/production_service.py`, `app/api/v1/endpoints/production.py`, `app/schemas/production.py`

**Frontend:** `MonthCalendarGrid.tsx`, `UrgentPositionsBlock.tsx`, `SubstrateRecommendationsBlock.tsx`, `CapacityDeficitAlert.tsx`, wizard state/API hooks

**Tests:** `tests/test_*capacity*`, `test_*urgent*`, `test_*substrate*`, `test_production_api_integration.py`, `test_production_podlozhki_e2e.py`; FE vitest under `src/features/production`

## Next steps

1. Human QA end-to-end в UI.
2. При появлении неполноширинных позиций — повторный Phase 0 / smoke анализа.
3. Пост-запуск: доля принятых подложек (A4).
