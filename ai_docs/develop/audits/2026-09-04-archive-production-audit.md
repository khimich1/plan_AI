# Аудит: архив и производство

**Дата**: 2026-09-04
**Скоуп**: архив КП + производство (kp subset) — backend `archive.py`, `production.py`, `delivery_schedule.py`; сервисы `archive_service`, `production_*`, `sgp_service`, `delivery_schedule_service`, `kp_readiness_service`; `kp_archive_repository.py`; schemas archive/production/delivery_schedule/sgp; `offer_access.py`; frontend `features/commercial-archive/`, `features/production/`, `pages/production/`, `pages/commercial-offer-archive/`. **Не входит**: конструктор КП / OCR / drafts.
**Реестр**: [FINDINGS.md](./FINDINGS.md)
**Прогон 0**: [2026-09-04-kp-audit.md](./2026-09-04-kp-audit.md) (конструктор КП). Более ранний production-only: [2026-08-12-production-module-audit.md](./2026-08-12-production-module-audit.md) (другая схема ID, без Health Score).

## Дельта

| ID | Было | Стало | Суть |
|----|------|-------|------|
| A4 | open | open | Толстый `production.py` 698 строк, 26 routes; day-documents и SGP CRUD минуют ProductionService |
| A5 | open | open | Raw SQL в sgp / delivery_schedule / kp_readiness, репозиториев нет |
| A11 | open | open | ArchiveService 868 строк, lazy-coupling к workflow / readiness / SGP |
| A15 | by-design | by-design | `owner_user_id` вне policy доступа |
| A22 | — | open | **NEW** — God-сервис `ProductionCompletionService` (~702 строк) + sqlite3 в application-слое |
| S1 | by-design | by-design | Общий архив — не IDOR |
| S8 | open | open | `str(exc)` в HTTP: production, archive, плюс delivery_schedule |
| Q8 | open | open | `useCreatePlanWizardState` 997 строк (было 724) |
| Q9 | open | open | `OfferDetailsDrawer` 1091 строка (было 1066) |
| Q10 | open | open | 9 JSON-маршрутов production без `response_model` |
| Q12 | open | open | ArchiveService глотает сбои enrichment; `move_to_production` re-raise на месте |
| Q18 | open | open | `exhaustive-deps` в MoveToProduction, DeliverySchedule, wizard (2 suppressions) |
| Q23 | — | open | **NEW** — Archive `round` vs production FE `ceil` — на 101 м: 2 дорожки vs 1 |
| Q24 | — | open | **NEW** — Copy-paste в day-documents: мёртвый `raise` после `raise_not_found` |

- Закрыто: 0
- Открыто (подтверждено): 9
- Новое: 3 (A22, Q23, Q24)
- Не воспроизвелось: 0

**Открытые P0 / P1 в скоупе**: 0 / 3 (A4, A5, A22)

## Действия (максимум 8)

1. **[A4]** Вынести из `production.py` в сервисы: `build_plan_from_filters`, day-documents, SGP CRUD, `_date_range_inclusive`.
2. **[A5]** Repository-слой для sgp / delivery_schedule / kp_readiness вместо `cur.execute`.
3. **[A22]** Разрезать `ProductionCompletionService` и убрать прямой sqlite3 — complete / write-off / mark_day через repos.
4. **[A11]** Разделить ArchiveService: Read / Mutation / Capacity; убрать lazy-coupling.
5. **[Q8]** Разрезать `useCreatePlanWizardState` (997 строк): selection / analyze-build / SGP-capacity.
6. **[Q12]** Не глотать исключения enrichment в архиве (`except Exception → None`).
7. **[Q23]** Одна формула оценки дорожек (archive BE и production FE расходятся на границе 101 м).
8. **[Q9]** Декомпозировать `OfferDetailsDrawer` (1091 строка).

## Открытые P0 / P1

### [A4] Толстый API-слой production.py
**Статус**: open
**Улика**: `app/api/v1/endpoints/production.py` wc=698; `build_plan_from_filters` `:104–180`; day-documents `:434–522` вызывают `generate_day_*` напрямую; SGP CRUD `:535–608`; helpers `_record_build_exclusions` / `_date_range_inclusive` `:628–678`
**Зачем**: Роутер смешивает HTTP, валидацию и доменную логику; day-documents и SGP обходят ProductionService; усложняет тестирование и повторное использование.

### [A5] Сервисы обходят repository, raw SQL
**Статус**: open
**Улика**: `app/services/sgp_service.py:78–89`; `app/services/delivery_schedule_service.py:494–500`; `app/services/kp_readiness_service.py:49–114`; нет *sgp* / *delivery* / *readiness* repos
**Зачем**: Raw SQL в сервисах обходит слой репозиториев, дублирует запросы, усложняет миграцию схемы и unit-тесты.

### [A22] ProductionCompletionService god + raw SQL
**Статус**: open
**Улика**: `app/services/production_completion_service.py` wc=702; sqlite3 `:4`, `:81–85`, `:125–128`, `:309`, `:515`
**Зачем**: God-сервис завершения производства (~702 строк) с прямым sqlite3 в application-слое; complete / write-off / mark_day не проходят через repository, усложняют транзакции и тесты.

## Приложение

### Архитектура (Medium)

**[A11] ArchiveService god-orchestrator** — `archive_service.py` 868 строк, 30 методов. Lazy-coupling: `:124–146` CommercialWorkflowService, `:140–146` / `:588–592` KpReadinessService, `:491–494` SgpService. Предложение: ArchiveReadService / ArchiveMutationService / ArchiveCapacityService.

**[A6] layout (вне скоупа)** — `core/production/planning.py:39–44` по-прежнему импортирует visualization at load.

**[A15] / [S1] by-design** — `offer_access.py:26–32`, ADR offer-access-policy.md; общий архив КП для менеджеров — не IDOR.

### Безопасность

Новых Critical/High не выявлено.

**[S8] Утечка деталей ошибок** — `str(exc)` клиенту через `raise_unprocessable_client_error`: `production.py:141`, `:174`, `:203`, `:209`, `:696`, `:147`, `:406`, `:428`, `:580`, `:605`; `archive.py:323`, `:347`, `:373`, `:407`, `:413`, `:491`, `:537`; `delivery_schedule.py:53`, `:76`, `:92`, `:128`, `:156`.

**S3** — in-process rate limit; kp-touchpoint вне этого скоупа.

**S5 / S9** — вне скоупа archive+production; новых улик нет.

Raw SQL параметризован — injection не выявлен.

**Logistics без финансов КП** — `tests/test_logistics_api.py:454–528`; RBAC `test_rbac_server_side.py`.

### Качество

**[Q8]** God-hook — `useCreatePlanWizardState.ts` wc=997 (было 724); `eslint-disable` `:256`, `:368`.

**[Q9]** God-component — `OfferDetailsDrawer.tsx` 1091 строка (было 1066).

**[Q10]** 9 routes без `response_model`: `production.py:87`, `:95`, `:220`, `:266`, `:321`, `:385`, `:525`, `:611`, `:619`; schemas есть `app/schemas/production.py:153–164`.

**[Q12]** ArchiveService скрывает сбои — `:487–500`, `:555–556`, `:593–594`, `:789–791` swallow; `move_to_production` `:247–253` re-raises.

**[Q18]** Suppressions exhaustive-deps — `MoveToProductionDialog.tsx:87`, `DeliveryScheduleDialog.tsx:92`, `useCreatePlanWizardState.ts:256`, `:368`.

**[Q23] NEW — расхождение оценки дорожек** — `archive_service.py:289–290` (`round`) vs `frontend/src/features/production/lib/productionEstimate.ts:24–25` (`ceil`); при 101 м BE tracks=2, FE tracks=1.

**[Q24] NEW — copy-paste day-document handlers** — `production.py:480–489` (breakdown), `:507–516` (formovka) vs schema `:447–458`; мёртвый `raise` после `raise_not_found`.

**useProductionEstimateQuery / archiveApi.getProductionEstimate** — на FE не используется (отдельный ID не заводился).

### Resolved 2026-08-12 (не в текущем реестре, без новых ID)

- dual validate_fill_targets — unified `core/production/capacity.py`
- expected_version на BE — `app/schemas/production.py:64`; tests `test_production_fill_integrity.py:235–267`
- complete_day races — guard + one tx
- date-range DoS — cap 366 дней
- dual calendar — delivery_schedule использует PlanDistributionService
- build+SGP — compensating delete `production_service.py:488–502`
