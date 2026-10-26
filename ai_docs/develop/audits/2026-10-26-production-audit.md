# Аудит: производство (плиты → план → СГП)

**Дата**: 2026-10-26
**Скоуп**: производство — менеджеры сохраняют плиты, уже стоящие в производстве, в план, и план переходит на СГП. Файлы: `app/api/v1/endpoints/production.py`; сервисы `production_*`, `sgp_service`, `day_view_service`, `plan_distribution_service`; `app/schemas/production.py`, `sgp.py`; `app/repositories/plan_repository.py`; `app/planning/` (включая `plan_storage.py`, `plan_calendar.py`); `core/production/`; `frontend/src/features/production/`; `frontend/src/pages/production/ProductionPage.tsx`. Только для сверки улик: `core/kp_db_plates_planning.py`, `core/visualization/__init__.py`, `app/repositories/auth_repository.py`, `app/dependencies/auth.py`, `core/production_capacity.py`, `app/services/promise_service.py` (образец актора). Вне скоупа: бот, GSM, конструктор КП, UI архива, график поставки, ILP `core/optimization/`.
**Реестр**: [FINDINGS.md](./FINDINGS.md)
**Прогон 0**: [2026-10-19-production-audit.md](./2026-10-19-production-audit.md)
**Модель**: родительская (Grok 4.7). `composer-2-fast` в сеансе не было; субагенты шли на той же модели, что и отчёт 2026-10-19.

## Дельта

| ID | Было | Стало | Суть |
|----|------|-------|------|
| A4 | open High P1 | open High P1 | `production.py` 699 строк, 26 маршрутов. Журнал `:168-179` после успешной сборки глотает сбой. Документы дня `:434-523`, `:628-647` из HTTP, минуя `ProductionService`. |
| A5 | open High P1 | open High P1 | `sgp_service.py:5`, `:78-89`, `:717-832`; `day_view_service.py:231`, `:269-277`. delivery/readiness не перепроверялись — ID не сужать. |
| A6 | open High P1 | open High P1 | `planning.py:41-46` тянет визуализацию; `visualization/__init__.py:16-18` — matplotlib Agg на импорте. В восьмёрку действий не входит. |
| A8 | open Medium | open Medium | Два календаря: `plan_calendar.py:136` и `plan_distribution_service.py:135`, `:197`. Живой HTTP: `production_service.py:511`. |
| A22 | open High P1 | open High P1 | `production_completion_service.py` ~702 строки, sqlite3 `:4`, `:54-66`, `:125-128`. Транзакция `:230-245` монолит не снимает. |
| A29 | open High P1 | open High P1 | Сборка уже закоммитила план; резерв СГП другим соединением. `save_plan` пишет JSON с `sgp_reservations` до `conn.commit()` резерва. Компенсация глотается. |
| S8 | open Medium | open Medium | `production.py` отдаёт `str(exc)`: `:141`, `:147`, `:174`, `:203`, `:209`, `:406`, `:428`, `:580`, `:605`, `:696`. archive/DS не перепроверялись. Severity не повышать. |
| S20 | open High P1 | open High P1 | Клиент шлёт `expected_version`; у `BuildPlanRequest` поля нет (`schemas/production.py:222`, только `CompleteProductionDayRequest` `:64`). `persist` берёт версию из БД. `fill_targets` обнуляет `active_plan_id`. |
| S21 | open High P1 | open High P1 | `DELETE /plans/{plan_id}` — два commit, без `return_plan_plates_to_production`. `kp_plates` остаются «в плане». Тест проверяет только 404. |
| Q6 | open Medium | open Medium | Две `get_global_calendar_info`: `plan_calendar.py:70`, max `:136` против `plan_distribution_service.py:135`, ёмкость `:197-200`. Вызовы delivery_schedule и archive не открывались. Partial. |
| Q8 | open Medium | open Medium | `useCreatePlanWizardState.ts` 998 строк (`:69–998`). В тело P1 не входит. |
| Q10 | open Medium | open Medium | Те же 9 JSON-маршрутов без `response_model`: `production.py:87`, `:95`, `:220`, `:266`, `:321`, `:385`, `:525`, `:611`, `:619`. |
| Q18 | open Low | open Low | Подавления `exhaustive-deps` на `useCreatePlanWizardState.ts:256` и `:368`. Новых Low нет. Глобально не закрывать. |
| Q23 | open Medium | open Medium | В скоупе `productionEstimate.ts:3`, `:24` — `Math.ceil(length/101)`, совпадает с `capacity.py:158` и `MAX_TRACK_LENGTH_M = 101.0`. Archive `round` не перепроверялся. Partial. |
| Q24 | open Medium | open Medium | Схема без мёртвого raise: `production.py:448–458`. Мёртвый raise в разбивке `:486–489` и формовке `:513–516`. Partial. |
| Q40 | open High P1 | open High P1 | Актор пустой: ищут email/login/user_id; сессия — id, username, role, manager_id, is_active, session_version, created_at. `updated_by` ёмкости дня и `reserve_on_conn` без actor. |
| A30 | — | open High P1 **NEW** | Пометка плит «в плане» и документ плана — разные соединения; ошибка `return_plan_plates_to_production` глотается. |

- Закрыто: 0
- Открыто (подтверждено): 16
- Новое: 1 (A30)
- Не воспроизвелось: 0

**Открытые P0 / P1 в скоупе**: 0 / 9

P1: A4, A5, A6, A22, A29, A30, S20, S21, Q40.

## Действия (максимум 8)

1. **[A30]** Пометка плит «в плане» и запись документа плана — один commit; ошибку `return_plan_plates_to_production` не глотать.
2. **[A29]** Сборка плана и резерв СГП — один commit; ошибку компенсации не глотать.
3. **[S21]** HTTP-удаление плана должно возвращать плиты в производство так же, как `plan_storage.delete_plan`, одним commit.
4. **[S20]** Принять `expected_version` в `BuildPlanRequest` и в `persist` либо перестать отправлять поле с клиента.
5. **[Q40]** Актор аудита брать из `username`/`id` (как исключения плана), включая `reserve_on_conn` и `updated_by` ёмкости дня.
6. **[A5]** `SgpService` (и сырой SQL дневного вида) через репозиторий.
7. **[A22]** Разрезать `ProductionCompletionService` и убрать прямой sqlite3 в репозитории.
8. **[A4]** Вынести из `production.py` журнал исключений и документы дня в сервисы.

A6 остаётся open P1 и в эту восьмёрку не входит: слоты заняты целостностью «плиты → план → СГП».

## Открытые P0 / P1

### [A30] Пометка плит и документ плана разными соединениями
**Статус**: open
**Улика**: `core/production/planning.py:337-377`; `core/kp_db_plates_planning.py:175` (`conn.commit()` в `mark_plates_as_planned`).
**Зачем**: Переход «плиты в производстве → в плане» не атомарен: `persist` сначала коммитит пометку плит, потом отдельным соединением пишет документ плана. Если запись плана падает, ошибка `return_plan_plates_to_production` логируется и не пробрасывается — плиты могут остаться «в плане» без строки плана.

### [A29] Сборка плана и резерв СГП разными соединениями
**Статус**: open
**Улика**: `app/services/production_service.py:466-505`; `app/repositories/plan_repository.py:147-176`, `:275-286`.
**Зачем**: Сборка уже закоммитила план, резерв СГП идёт другим соединением. `save_plan` коммитит JSON с `sgp_reservations` до `conn.commit()` резерва. Ошибка компенсации логируется и глотается. Тест `test_sgp_reserve_fail_after_build_compensates_no_plan_plates_not_in_plan` бьёт только исключение внутри `reserve_on_conn`, до `save_plan`.

### [S21] HTTP-удаление плана не возвращает плиты
**Статус**: open
**Улика**: `app/api/v1/endpoints/production.py:255`, `app/services/production_service.py:156`, `app/services/sgp_service.py:854`, `app/repositories/plan_repository.py:189`, `app/planning/plan_storage.py:166`, `core/kp_db_plates_planning.py:563`.
**Зачем**: `DELETE /plans/{plan_id}` снимает ссылку СГП и удаляет документ плана двумя commit и не вызывает `return_plan_plates_to_production`. `kp_plates` остаются «в плане». `plan_storage.delete_plan` возврат вызывает. Тест `tests/test_production_api_integration.py::test_activate_and_delete_plan` проверяет только 404 плана.

### [S20] `expected_version` сборки плана отбрасывается
**Статус**: open
**Улика**: `frontend/src/features/production/hooks/useCreatePlanWizardState.ts:902`, `frontend/src/features/production/types/production.ts:228`, `app/schemas/production.py:222` (поле отсутствует; `expected_version` только на `CompleteProductionDayRequest`, `:64`), `app/api/v1/endpoints/production.py:112`, `core/production/planning.py:307`, `core/production/planning.py:359`.
**Зачем**: Клиент шлёт `expected_version`, `BuildPlanRequest` поля не содержит. `persist` берёт версию из БД. При `fill_targets` `active_plan_id` обнуляется до записи.

### [Q40] Пустой актор аудита производства
**Статус**: open
**Улика**: `app/api/v1/endpoints/production.py:290`, `:392`, `:573`, `:593`, `:641`, `:691`; `app/repositories/auth_repository.py:84`; `app/dependencies/auth.py:21`; `app/services/production_completion_service.py:167`; `app/services/sgp_service.py:524`, `:693`, `:847`; `app/services/production_service.py:476`; `app/services/production_capacity_service.py:28`; `app/services/promise_service.py:285`.
**Зачем**: Актор пустой: ищут email/login/user_id, сессия содержит только id, username, role, manager_id, is_active, session_version, created_at. Тот же набор обнуляет `updated_by` ёмкости дня. `reserve_on_conn` без actor. Рабочий образец читает username/id. Обхода ролей нет (admin и production на всех 26 маршрутах) — это не новый S.

### [A5] Сырой SQL в обход репозитория
**Статус**: open
**Улика**: `app/services/sgp_service.py:5`, `:78-89`, `:717-832`; `app/services/day_view_service.py:231`, `:269-277`.
**Зачем**: `SgpService` и дневной вид ходят в SQLite сами, мимо репозитория. delivery/readiness по-прежнему не перепроверялись — ID не сужать.

### [A22] God-сервис завершения дня и прямой sqlite3
**Статус**: open
**Улика**: `app/services/production_completion_service.py:4`, `:54-66`, `:125-128`, `:230-245`.
**Зачем**: `ProductionCompletionService` ~702 строки и прямой sqlite3. Одна транзакция списания и флага дня монолит не снимает.

### [A4] Толстый роутер производства
**Статус**: open
**Улика**: `app/api/v1/endpoints/production.py:168-179`, `:434-523`, `:628-647`.
**Зачем**: Журнал исключений и документы дня по-прежнему в `production.py` (699 строк, 26 маршрутов). Журнал пишется после успешной сборки и глотает сбой. Документы дня вызывают генерацию прямо из HTTP, минуя `ProductionService`.

### [A6] Matplotlib на загрузке планирования
**Статус**: open
**Улика**: `core/production/planning.py:41-46`; `core/visualization/__init__.py:16-18`.
**Зачем**: Загрузка планирования тянет пакет визуализации, matplotlib Agg на импорте. Загрузка matplotlib остаётся P1, в восьмёрку действий этого прогона не вошла, потому что слоты заняты целостностью «плиты → план → СГП».

## Приложение

Подробности Medium и Low. В тело P0/P1 не входят. Закрытых находок нет. Целиком невоспроизведённых ID нет.

### Medium

**[A8]** Два календаря: `app/planning/plan_calendar.py:136` и `app/services/plan_distribution_service.py:135`, `:197`. Живой HTTP: `production_service.py:511`. Действия на эту неделю нет.

**[S8]** `production.py` отдаёт `str(exc)`: `:141`, `:147`, `:174`, `:203`, `:209`, `:406`, `:428`, `:580`, `:605`, `:696`. Severity не повышать. `archive.py` и `delivery_schedule.py` не перепроверялись.

**[Q6]** Две `get_global_calendar_info`: `plan_calendar.py:70` и max на `:136` против `plan_distribution_service.py:135` и ёмкость `:197-200`. Вызовы delivery_schedule и archive не открывались. Partial: половина в скоупе открыта, ID не закрывать.

**[Q8]** `useCreatePlanWizardState.ts` 998 строк, хук `:69–998`. Разрез: плиты `:126–164` и `:450–604`; ёмкость `:166–296`; обещания `:308–444` и `:727–758`; СГП `:697–815`; сборка `:817–923`. В тело P1 не входит.

**[Q10]** Те же 9 JSON-маршрутов без `response_model`: `production.py:87`, `:95`, `:220`, `:266`, `:321`, `:385`, `:525`, `:611`, `:619`.

**[Q23]** В скоупе `productionEstimate.ts:3`, `:24` — `Math.ceil(length/101)`, совпадает с `core/production/capacity.py:158` и `MAX_TRACK_LENGTH_M = 101.0` в `core/production_capacity.py:11`. Archive `round` не перепроверялся, ID не закрывать.

**[Q24]** Схема без мёртвого raise: `production.py:448–458`. Мёртвый raise после `raise_not_found` в разбивке `:486–489` и формовке `:513–516`. Partial, ID открыт.

### Low

**[Q18]** Подавления `exhaustive-deps` на `useCreatePlanWizardState.ts:256` и `:368`. Новых Low нет. Глобально не закрывать.

**[S1]** by-design, в этом прогоне не переоткрывался.
