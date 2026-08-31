# Аудит: --full (сверка реестра)

**Дата**: 2026-08-28  
**Скоуп**: полный проект (--full): backend FastAPI, frontend React, коммерческий/производственный контур, layout pipeline, GSM, auth/platform; bot исключён как живой продукт  
**Реестр**: [FINDINGS.md](./FINDINGS.md)  
**Прогон 0**: [2026-08-28-full-project-audit.md](./2026-08-28-full-project-audit.md)

---

## Дельта

| ID | Было | Стало | Суть |
|----|------|-------|------|
| A3 | open (прогон 0) | resolved | Facade `commercial_workflow_service.py` ≤710 строк; identity / plate-resolve / lifecycle / ProductDraftHandler вынесены |
| S2 | open | resolved | `fastapi==0.141.1`, `starlette==1.6.0`; pip-audit по HTTP-стеку чист |
| S4 | open | resolved | `npm audit:ci` high=0 |
| Q1, Q2, Q3, Q5 | open | resolved | product_draft_config/handler; A4 runners; layout builder; commercial_plate_resolve |
| S1, S11, A15 | open / High | by-design | Общий архив КП — осознанная политика ([offer-access-policy.md](../architecture/offer-access-policy.md)); CSRF double-submit |
| A21 | — | open (Medium) | **NEW**: `gsm_kit_gate` импортирует private `_chain_broken` из overview |
| Q21 | — | open (Low) | **NEW**: две реализации `formatLiters` на фронте GSM |
| A1, A2, A4–A20, S3, S5–S10, S12–S16, Q4, Q6–Q20 | open | open (подтверждено) | Улики обновлены; частичный прогресс у A1, A2, A16, Q20 |

**Счётчики прогона**

| Метрика | Значение |
|---------|----------|
| Закрыто в этом прогоне | 0 (новых закрытий) |
| Уже resolved / by-design подтверждено | A3, S2, S4, Q1, Q2, Q3, Q5, S1, S11, A15 |
| Открыто, подтверждено с уликами | 47 ID |
| Новое | 2 (A21, Q21) |
| Не воспроизведено | 0 |

**Открытые P0 / P1 в скоупе**: **2 / 13**

---

## Действия (максимум 8)

1. **[A1]** Завершить миграцию `PlateOrderContext` — убрать PEP 562 globals в `config_and_data.py`; покрыть BackgroundTasks/CLI.
2. **[A2] / [S3]** Shared store (Redis/DB) для rate limit и drafts **или** жёсткий single-worker contract во всех окружениях.
3. **[A18]** Вынести LibreOffice/`soffice` из HTTP-потока (очередь/worker); лимит параллельных конвертаций.
4. **[Q19]** Унифицировать округление литров: banker's round на фронте (`downstreamPreview.ts`) как в `balance.py`.
5. **[Q20]** Единая политика битого `season_switches`: throw в `gsm_registry_service.get_settings` как в generation/export.
6. **[S16]** Лимит числа файлов импорта GSM (`gsm.py:147–166`); per-file 50MB уже есть.
7. **[A16] / [A21]** Декомпозиция `gsm_generation_service.py` + публичный helper вместо `_chain_broken` coupling.
8. **[A4] remainder** — thin controller для `production.py` (~639 строк, `build_plan_from_filters`).

---

## Открытые P0 / P1

### P0 — Critical

#### [A1] Неявное мутабельное глобальное состояние заказа плит

| Поле | Значение |
|------|----------|
| **Статус** | open (частичный прогресс) |
| **Улика** | `core/config_and_data.py:383–411` (PEP 562 `PLATES_*`), `core/domain/plate_order.py:331–383` (`apply_to_globals`), `viz_modules/layout_sequence/from_plan.py:194`; изоляция HTTP: `app/middleware/plate_runtime_isolation.py:12–17`; CPU pool: `app/concurrency/cpu_bound.py:38–41` |
| **Зачем** | Middleware и CPU-pool частично изолируют HTTP, но globals живы — риск утечки состояния в BackgroundTasks, CLI и код без явного контекста |

#### [A2] In-process state блокирует горизонтальное масштабирование

| Поле | Значение |
|------|----------|
| **Статус** | open (частичный прогресс) |
| **Улика** | `app/security/login_rate_limit.py:84–85`, `app/services/draft_store.py:35–40`, `app/main.py:46–49` (`enforce_single_instance_workers`) |
| **Зачем** | Rate limit, drafts, counters in-memory; fail-fast только production+single_instance — несколько workers/replicas ломают консистентность |

---

### P1 — High

#### [A4] Толстый API-слой (remainder production.py)

| Поле | Значение |
|------|----------|
| **Статус** | open |
| **Улика** | `app/api/v1/endpoints/production.py` ~639 строк; `build_plan_from_filters` :101–165; `commercial.py` ~479 строк (прогресс) |
| **Зачем** | Presentation дублирует orchestration; сложно тестировать и эволюционировать API |

#### [A5] Сервисы обходят repository, raw SQL

| Поле | Значение |
|------|----------|
| **Статус** | open |
| **Улика** | `app/services/sgp_service.py:71–78`, `delivery_schedule_service.py:493–494`, `kp_readiness_service.py:48–49` |
| **Зачем** | SQL размазан; нет единой границы persistence |

#### [A6] Planning импортирует visualization / matplotlib at load

| Поле | Значение |
|------|----------|
| **Статус** | open |
| **Улика** | `core/production/planning.py:39–44`, `core/visualization/__init__.py:16–18` |
| **Зачем** | Домен планирования тянет тяжёлый viz-стек при импорте |

#### [A7] Неполный DI в FastAPI

| Поле | Значение |
|------|----------|
| **Статус** | open |
| **Улика** | `app/dependencies/services.py:41–156` vs `get_auth_service` :159–162 |
| **Зачем** | Скрытая связность; граф зависимостей неединообразен и слабо тестируем |

#### [A16] God-сервис жизненного цикла ПЛ GSM

| Поле | Значение |
|------|----------|
| **Статус** | open (частичный прогресс) |
| **Улика** | `app/services/gsm_generation_service.py` ~1030 строк; `gsm_kit_gate.py` вынесен, generation остаётся god-module |
| **Зачем** | CRUD + confirm + rechain в одном модуле — широкие регрессии при изменениях GSM |

#### [A17] God-модуль `core/gsm/generator.py`

| Поле | Значение |
|------|----------|
| **Статус** | open |
| **Улика** | `core/gsm/generator.py` ~1324 строки |
| **Зачем** | Монолит генерации ПЛ; сложно тестировать изолированно |

#### [A18] Блокирующий LibreOffice/`soffice` в HTTP

| Поле | Значение |
|------|----------|
| **Статус** | open |
| **Улика** | `app/services/gsm_export_service.py:64–70`, `app/api/v1/endpoints/gsm.py:620–662` |
| **Зачем** | Синхронная конвертация в request thread — DoS и таймауты при нагрузке |

#### [A19] Сезонная логика дублируется на фронте

| Поле | Значение |
|------|----------|
| **Статус** | open |
| **Улика** | `core/gsm/season.py:35–43` vs `frontend/.../downstreamPreview.ts:28–45` |
| **Зачем** | Расхождение контракта сезонных переключателей между API и UI |

#### [A20] Импорт транзакций без unit-of-work

| Поле | Значение |
|------|----------|
| **Статус** | open |
| **Улика** | `app/services/gsm_transaction_service.py:82–114`, `app/repositories/gsm_repository.py:391–424` (per-row commit) |
| **Зачем** | Частичный импорт при сбое; нет атомарности batch |

#### [S3] Rate limiting in-process

| Поле | Значение |
|------|----------|
| **Статус** | open |
| **Улика** | `app/security/login_rate_limit.py` (связано с A2) |
| **Зачем** | Brute-force лимиты обходятся при workers>1 |

#### [S16] Импорт GSM без лимита числа файлов

| Поле | Значение |
|------|----------|
| **Статус** | open |
| **Улика** | `app/api/v1/endpoints/gsm.py:147–166` — нет cap на количество файлов; per-file 50MB есть |
| **Зачем** | DoS через множество мелких upload |

#### [Q19] Расхождение округления литров

| Поле | Значение |
|------|----------|
| **Статус** | open |
| **Улика** | `frontend/.../downstreamPreview.ts:23–26` (`Math.round`) vs `core/gsm/balance.py:16–17` (banker's round); burn 25km 10.1l → 2.53 vs 2.52 |
| **Зачем** | Preview и отчёт расходятся на граничных значениях |

#### [Q20] Несогласованная обработка битого `season_switches`

| Поле | Значение |
|------|----------|
| **Статус** | open (частичный прогресс) |
| **Улика** | `gsm_registry_service.py:356–361` глотает corrupt JSON; `gsm_generation_service.py:703–711` бросает |
| **Зачем** | Разное поведение registry vs generation/export при повреждённых настройках |

---

## Приложение

### Medium (action this week)

| ID | Scope | Summary | Улика |
|----|-------|---------|-------|
| A8 | layout | Параллельные подсистемы планирования | `planning.py`, `plan_manager.py`, `plan_distribution*` |
| A9 | platform | Пустые app-сервисы-реэкспорты | `kp_persistence_service.py`, `rest_matching_service.py` |
| A10 | bot | Legacy bot paths в persistence | `plan_storage.py`; не живой контур |
| A11 | kp | ArchiveService god-orchestrator | `archive_service.py` ~812 строк |
| A12 | kp | Frontend god-hook мастера КП | `useCommercialOfferWizard.ts` ~453 строк |
| A21 | gsm | **NEW** Coupling kit_gate → private `_chain_broken` | `gsm_kit_gate.py:10`, `:120` |
| S5 | kp | OCR во внешние LLM | policy, не сюрприз |
| S7 | auth | CSP Report-Only + unsafe-inline | `security_headers.py` |
| S8 | platform | Утечка деталей ошибок в HTTP | `http_errors.py`, endpoints |
| S9 | auth | CSRF парсит multipart до токена | `csrf.py` |
| S10 | auth | Сессия 12ч без refresh-ротации | `session.py` |
| Q4 | kp | Пять копий `build_*_preview_metadata` | `commercial_draft_service.py` |
| Q6 | layout | Две реализации `get_global_calendar_info` | `plan_calendar.py` / `plan_distribution_service.py` |
| Q7 | kp | Product-type duplication на фронте | `commercialOfferApi.ts` + wizard |
| Q8 | kp | God-hook `useCreatePlanWizardState` | ~724 строк |
| Q9 | kp | God-component `OfferDetailsDrawer` | ~901 строк |
| Q10 | kp | Слабая типизация production API | production endpoints |
| Q11 | kp | preview: `Any` / `dict[str, Any]` | preview schemas |
| Q12 | kp | ArchiveService скрывает частичные сбои | `except Exception → None` |
| Q13 | layout | Нет прямых тестов `get_global_calendar_info` | после Q6 |

### Low

| ID | Scope | Summary | Улика |
|----|-------|---------|-------|
| A14 | layout | Монолит `core/visualization`, matplotlib at import | `visualization/__init__.py` |
| S6 | platform | SQLite без шифрования at rest | `plita.db`, `pb.db` |
| S12 | auth | Password policy messages на английском | auth schemas |
| S13 | auth | `/health` метаданные вне production | health endpoint |
| S14 | bot | Legacy bot auth bypass | `bot_archived/` |
| S15 | kp | Draft в sessionStorage (XSS-вектор) | `draftStorage.ts` |
| Q14 | platform | Однострочные delegate-обёртки | services |
| Q15 | layout | Имя `_merge_plate_texts` вводит в заблуждение | layout utils |
| Q16 | kp | `/parse` без `response_model` | commercial endpoint |
| Q17 | gsm | `GsmGenerationError` messages на английском | generation service |
| Q18 | kp | Подавление `react-hooks/exhaustive-deps` (9 файлов) | frontend |
| Q21 | gsm | **NEW** Две реализации `formatLiters` | `importReport.ts:28–29` vs `waybillWarnings.ts:80–84` |

### Подтверждённые resolved / by-design

| ID | Статус | Суть |
|----|--------|------|
| A3 | resolved | Facade commercial workflow ≤710 строк |
| S2 | resolved | fastapi 0.141.1, starlette 1.6.0 |
| S4 | resolved | npm audit:ci high=0 |
| Q1, Q2, Q3, Q5 | resolved | product pipeline, A4 runners, layout builder, plate resolve |
| S1, A15 | by-design | Общий архив КП — не IDOR для исправления |
| S11 | by-design | CSRF-cookie не HttpOnly (double-submit) |

### Частичный прогресс (still open)

- **A1**: HTTP middleware + CPU pool isolation добавлены; globals через PEP 562 и `apply_to_globals` всё ещё живы.
- **A2**: Документировано + fail-fast при production и workers≠1; shared store не внедрён.
- **A16**: `gsm_kit_gate` извлечён; `gsm_generation_service.py` остаётся god-module (~1030 строк).
- **Q20**: generation/export бросают на corrupt JSON; `gsm_registry_service.get_settings` по-прежнему глотает.

### Positive signals

- Draft ownership и path traversal guard на черновиках
- `destructive_db_guard` на опасные операции БД
- ACL-тесты logistics/financial и GSM `REQUIRE_ACCOUNTING`
- HttpOnly session cookies; CSRF double-submit (S11 by-design)
- Ports/adapters для viz boundary (`core/ports/visualization.py`)
- GSM kit gate с unit-тестами (`test_gsm_kit_gate.py`) — база для hardening month-close
