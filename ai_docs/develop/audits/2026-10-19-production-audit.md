# Аудит: производство

**Дата**: 2026-10-19
**Скоуп**: модуль производство — `app/api/v1/endpoints/production.py`; сервисы `production_*`, `sgp_service`, `day_view_service`, `plan_distribution_service`; `app/schemas/production.py`, `sgp.py`; `app/repositories/plan_repository.py`; `app/planning/`; `core/production/`; `frontend/src/features/production/`; `frontend/src/pages/production/ProductionPage.tsx`. Вне скоупа: бот, GSM, конструктор КП, архив, график поставки, ILP `core/optimization/`.
**Реестр**: [FINDINGS.md](./FINDINGS.md)
**Прогон 0**: [2026-09-04-archive-production-audit.md](./2026-09-04-archive-production-audit.md); более ранний локальный (чужие ID): [2026-08-12-production-module-audit.md](./2026-08-12-production-module-audit.md)
**Модель**: inherit (Grok 4.7). `composer-2-fast` в сеансе не было.

## Дельта

| ID | Было | Стало | Суть |
|----|------|-------|------|
| A4 | open High P1 | open High P1 | `production.py` 699 строк, 26 routes. Журнал исключений и документы дня в роутере: `:168-169`, `:434-523`, `:628-647`. |
| A5 | open High P1 | open High P1 | Сырой SQL СГП подтверждён: `sgp_service.py:5`, `:78-89`; также `day_view_service.py:231`. delivery/readiness не перепроверялись. |
| A6 | open High P1 | open High P1 | `planning.py:41-45` импортирует `core.visualization`; `visualization/__init__.py:16-18` вызывает `matplotlib.use('Agg')` на импорте. |
| A8 | open Medium | open Medium | Два календаря: `plan_calendar.py:136` и `plan_distribution_service.py:183-200`. Живой HTTP: `production_service.py:511`. Не P1. |
| A22 | open High P1 | open High P1 | `production_completion_service.py` ~703 строки, sqlite3 `:4`, `:125`. Одна транзакция списания и флага дня `:233-245` находку не закрывает. |
| S8 | open Medium | open Medium | `production.py` отдаёт `str(exc)`: `:141`, `:147`, `:174`, `:203`, `:209`, `:406`, `:428`, `:580`, `:605`, `:696`. archive/DS не перепроверялись. Severity не повышать. |
| Q6 | open Medium | open Medium | Две `get_global_calendar_info`: `plan_calendar.py:70`, `:136` vs `plan_distribution_service.py:135`, `:197-200`. Вызовы delivery_schedule и archive не перепроверялись. |
| Q8 | open Medium | open Medium | `useCreatePlanWizardState.ts` 997 строк (`:69–997`). Разрез по плитам, ёмкости, обещаниям, СГП и сборке на месте. |
| Q10 | open Medium | open Medium | Те же 9 JSON-маршрутов без `response_model`: `production.py:87,95,220,266,321,385,525,611,619`. |
| Q18 | open Low | open Low | Подавления в мастере на месте: `useCreatePlanWizardState.ts:256`, `:368`. Глобально не закрывать. |
| Q23 | open Medium | open Medium | FE `Math.ceil(length/101)` совпал с `capacity.py:158` и `MAX_TRACK_LENGTH_M = 101.0`. Archive `round` не перепроверялся. |
| Q24 | open Medium | open Medium | Схема исправлена (`production.py:447–458`); мёртвый raise в разбивке `:480–489` и формовке `:507–516`. Partial. |
| A29 | — | open High P1 **NEW** | Сборка плана и резерв СГП — разные соединения; ошибка компенсации глотается. |
| S20 | — | open High P1 **NEW** | Клиент шлёт `expected_version`, `BuildPlanRequest` его отбрасывает; `persist` берёт version из БД. |
| S21 | — | open High P1 **NEW** | HTTP `DELETE /plans/{plan_id}` не возвращает плиты в производство; `kp_plates` остаются «в плане». |
| Q40 | — | open High P1 **NEW** | Актор аудита пустой: код ищет email/login/user_id, в сессии только username и id. |

- Закрыто: 0
- Открыто (подтверждено): 12
- Новое: 4 (A29, S20, S21, Q40)
- Не воспроизвелось: 0

**Открытые P0 / P1 в скоупе**: 0 / 8

## Действия (максимум 8)

1. **[S21]** HTTP-удаление плана должно возвращать плиты в производство так же, как `plan_storage.delete_plan`, одним commit.
2. **[A29]** Сборка плана и резерв СГП — один commit; ошибку компенсации не глотать.
3. **[S20]** Принять `expected_version` в `BuildPlanRequest` и в `persist` либо перестать отправлять поле с клиента.
4. **[Q40]** Актор аудита брать из `username`/`id` (как исключения плана), включая `reserve_on_conn`.
5. **[A22]** Разрезать `ProductionCompletionService` и убрать прямой sqlite3 в репозитории.
6. **[A4]** Вынести из `production.py` журнал исключений и документы дня в сервисы.
7. **[A6]** Импортировать раскладку из подмодуля, не из пакета `core.visualization` (matplotlib на загрузке).
8. **[A5]** `SgpService` (и сырой SQL дневного вида) через репозиторий.

## Открытые P0 / P1

### [A4] Толстый роутер производства
**Статус**: open
**Улика**: `app/api/v1/endpoints/production.py` — 699 строк, 26 routes. Журнал исключений `:168-169`, документы дня `:434-523`, `:628-647`.
**Зачем**: Журнал исключений и документы дня живут в HTTP-слое рядом с остальными маршрутами. Эту логику нельзя вызвать и проверить отдельно от роутера.

### [A5] Сырой SQL в обход репозитория
**Статус**: open
**Улика**: `app/services/sgp_service.py:5`, `:78-89`; `day_view_service.py:231`. `delivery_schedule_service.py` и `kp_readiness_service.py` в этом прогоне не смотрелись.
**Зачем**: СГП и дневной вид ходят в SQLite из сервиса. Прежние точки в графике поставки и готовности КП не перепроверялись, поэтому находка остаётся открытой целиком.

### [A6] Matplotlib на загрузке планирования
**Статус**: open
**Улика**: `core/production/planning.py:41-45` импортирует `core.visualization`; `core/visualization/__init__.py:16-18` вызывает `matplotlib.use('Agg')` на импорте.
**Зачем**: Загрузка планирования тянет пакет визуализации и backend Agg, хотя планированию нужна только раскладка.

### [A22] God-сервис завершения дня и прямой sqlite3
**Статус**: open
**Улика**: `production_completion_service.py` ~703 строки, sqlite3 `:4`, `:125`. Списание и флаг дня уже одна транзакция `:233-245`.
**Зачем**: Одна транзакция списания не снимает монолит и прямой доступ к SQLite. Сервис по-прежнему обходит репозиторий.

### [A29] Сборка плана и резерв СГП разными соединениями
**Статус**: open
**Улика**: `production_service.py:466-504`, `plan_repository.py:147-176`. Тест `test_sgp_reserve_fail_after_build_compensates_no_plan_plates_not_in_plan` закрывает только исключение до `save_plan`.
**Зачем**: План и резерв СГП коммитятся разными соединениями, ошибка компенсации глотается. Сбой после `save_plan` этим тестом не пойман.

### [S20] `expected_version` сборки плана отбрасывается
**Статус**: open
**Улика**: `useCreatePlanWizardState.ts:902`, `types/production.ts:227`, `schemas/production.py:222`, `production.py:112`, `planning.py:307`, `planning.py:359`. Версию проверяют complete и DELETE track: `test_complete_day_stale_expected_version_returns_409_kp_not_written_off`, `test_delete_track_stale_expected_version_returns_409_track_remains`.
**Зачем**: Клиент шлёт `expected_version`, схема `BuildPlanRequest` поле отбрасывает, `persist` берёт version из БД. Живой мастер с `fill_targets` обнуляет `active_plan_id` и создаёт новый план — токен всё равно не действует на дозаписи.

### [S21] HTTP-удаление плана не возвращает плиты
**Статус**: open
**Улика**: `production.py:255`, `production_service.py:156`, `plan_storage.py:166`, `kp_repository.py:342`, `core/kp_db_plates_planning.py:563`.
**Зачем**: `plan_storage.delete_plan` вызывает `return_plan_plates_to_production`, HTTP-путь — нет, и это два commit. `kp_plates` остаются «в плане» с нулевым остатком.

### [Q40] Пустой актор аудита производства
**Статус**: open
**Улика**: `production.py:290` (снятие дорожки), `:392` → `production_completion_service.py:167` (завершение дня), `:573`, `:593` → `sgp_service.py:524`, `:693` (СГП). Рабочий образец: `production.py:641` и `promise_service.py:285–290` читают username/id. `reserve_on_conn` из сборки без actor (`production_service.py:476–483`). Сессия: `auth_repository.py:84`.
**Зачем**: Код ищет email, login и user_id, в сессии только username и id, поэтому актор пустой. Резерв СГП при сборке плана actor не передаёт.

## Приложение

### Medium

**[A8]** Два календаря: `app/planning/plan_calendar.py:136` (`MAX_TRACKS_PER_DAY`) и `app/services/plan_distribution_service.py:183-200`. Живой HTTP календаря: `production_service.py:511`. В тело P0/P1 не входит, действия на эту неделю нет.

**[S8]** `production.py` по-прежнему отдаёт `str(exc)`: `:141`, `:147`, `:174`, `:203`, `:209`, `:406`, `:428`, `:580`, `:605`, `:696`. Строки `archive.py` и `delivery_schedule.py` не перепроверялись. Severity не повышать.

**[Q6]** Две `get_global_calendar_info`: `plan_calendar.py:70`, `:136` vs `plan_distribution_service.py:135`, `:197-200`. Вызовы delivery_schedule и archive не перепроверялись.

**[Q8]** `useCreatePlanWizardState.ts` 997 строк (`:69–997`). Разрез: плиты `:126–164`, `:450–604`; анализ/ёмкость `:166–296`; обещания `:308–444`, `:727–758`; СГП `:697–815`; сборка `:817–921`.

**[Q10]** Те же 9 JSON-маршрутов без `response_model`: `production.py:87,95,220,266,321,385,525,611,619`.

**[Q23]** Фронт перемерено: `productionEstimate.ts:3`, `:24–25` — `Math.ceil(length/101)`, совпадает с `capacity.py:158` и `MAX_TRACK_LENGTH_M = 101.0` в `core/production_capacity.py:11`. Архив `archive_service.py` не перепроверялся. 2026-10-19: FE ceil совпал с capacity; archive round не перепроверялся.

**[Q24]** Схема исправлена (`production.py:447–458`); мёртвый raise остался в разбивке `:480–489` и формовке `:507–516`. Статус open (partial).

### Low

**[Q18]** Подавления в мастере на месте: `useCreatePlanWizardState.ts:256`, `:368`. Глобально не закрывать. Новых Low нет.

**[S1]** by-design, в этом прогоне не переоткрывался.

Локальные пункты 2026-08-12, которых нет в реестре и которые в этом прогоне не открыты заново: один `validate_fill_targets` (`capacity.py:92-132`); `complete_day` — флаг завершённого дня, пропуск уже списанных, одна транзакция, 409 на старую версию; диапазон дат ≤366 (`production.py:362`, `:672`, `schemas/production.py:11`).
