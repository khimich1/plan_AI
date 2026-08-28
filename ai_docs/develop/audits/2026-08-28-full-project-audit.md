# Отчёт по аудиту проекта

**Дата**: 2026-08-28  
**Область**: весь проект (критичные 20–30%: backend FastAPI, frontend React, коммерческий/производственный контур, auth, layout pipeline)  
**Аудиторы**: senior-reviewer + security-auditor + reviewer

---

## Краткое резюме

**Общая оценка здоровья**: 2.0/10

| Серьёзность | Архитектура | Безопасность | Качество кода | Итого |
|-------------|-------------|--------------|---------------|-------|
| Critical    | 2           | 0            | 0             | **2** |
| High        | 5           | 4            | 3             | **12** |
| Medium      | 5           | 6            | 10            | **21** |
| Low         | 2           | 5            | 5             | **12** |

**Расчёт Health Score**:

- Старт: **10**
- Critical: 2 × −2 = **−4** (потолок −6)
- High: 12 × −0.5 = −6, потолок **−3**
- Medium: 21 × −0.1 = −2.1, потолок **−1**
- Low: не учитываются
- **Итого = 10 − 4 − 3 − 1 = 2.0/10**

**Сильные стороны**: корректный паттерн Ports для viz_modules ([A16] — `core/ports/visualization.py`, adapters, `tests/test_core_viz_import_boundary.py`); HttpOnly session cookies, CSRF, защита draft path traversal, destructive DB guard, OCR magic bytes, доверенные XFF proxies, draft ownership.

**Рекомендация**: немедленно устранить 2 критические архитектурные проблемы ([A1], [A2]); в этом же спринте закрыть High-находки по безопасности [S1]–[S4] (особенно IDOR [S1] и уязвимости Starlette [S2]); параллельно начать декомпозицию god-модулей коммерческого контура ([A3], [A4], [Q1]–[Q3]).

**Ремедиация P0 (2026-08-28)**: [Q3], [S2], [S4] закрыты кодом; [S1]/[A15] закрыты как by design (ADR). Спека: [stabilizaciya-p0-audit-2026-08-28.md](../../specs/stabilizaciya-p0-audit-2026-08-28.md).

---

## Критические проблемы (исправить немедленно)

### [A1] Неявное мутабельное глобальное состояние заказа плит

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `core/plate_runtime_state.py` (PlateMutableRuntime, get_plate_mutable_runtime), `core/config_and_data.py` (PEP 562 __getattr__ → PLATES_*), `core/domain/plate_order.py` (apply_to_globals, get_current_plate_order), `app/middleware/plate_runtime_isolation.py` |
| **Влияние** | Бизнес-логика опирается на thread-local/ContextVar, а не на явно передаваемый контекст. Middleware изолирует HTTP, но BackgroundTasks, CPU-pool (`app/concurrency/cpu_bound.py`), CLI и код без `PlateOrderContext.bound()` рискуют утечкой состояния между запросами и задачами |
| **Исправление** | Завершить миграцию A1-002 — передавать `PlateOrderContext` явно, убрать legacy через `config_and_data` |

### [A2] In-process state блокирует горизонтальное масштабирование API

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `app/security/login_rate_limit.py` (_SlidingWindowRateLimiter), `app/services/draft_store.py`, `app/main.py` (enforce_single_instance_workers) |
| **Влияние** | Rate limiting, OCR-лимиты in-memory; при workers>1 защита ослабевается; черновики и счётчики не разделяются между процессами. Горизонтальное масштабирование API невозможно без потери консистентности |
| **Исправление** | Redis/DB shared store или формализовать single-worker как hard requirement при старте |

---

## Проблемы высокого приоритета (исправить скоро)

### Архитектура

#### [A3] God-module коммерческого контура

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `app/services/commercial_workflow_service.py` (~3027 строк, CommercialWorkflowService) |
| **Влияние** | OCR/AI для 6 типов, drafts, wizard, расчёт, экспорт в одном модуле. Нарушает SRP; любое изменение коммерческого контура рискует широкими регрессиями; сложно тестировать изолированно |
| **Исправление** | Use-case сервисы по вертикалям + тонкий facade |

#### [A4] Толстый API-слой КП

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `app/api/v1/endpoints/commercial.py` (~897 строк, 30+ handlers); соседний `production.py` ~639 строк |
| **Влияние** | Presentation дублирует orchestration; endpoints содержат бизнес-логику вместо thin controllers |
| **Исправление** | Thin controllers; делегирование в сервисы |

#### [A5] Сервисы обходят repository, raw SQL

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `app/services/sgp_service.py`, `delivery_schedule_service.py`, `kp_readiness_service.py` |
| **Влияние** | SQL размазан по сервисам; нет единой границы persistence; сложнее тестировать и эволюционировать схему |
| **Исправление** | `SgpRepository`, `DeliveryScheduleRepository` |

#### [A6] Planning зависит от visualization

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `core/production/planning.py` импортирует `core.visualization` / ports; `core/visualization/__init__.py` — matplotlib at load |
| **Влияние** | Домен планирования тянет тяжёлый viz-стек при импорте; side-effect import matplotlib замедляет и усложняет тестирование |
| **Исправление** | Вынести `split_sequence_into_tracks` в `core/layout` без matplotlib |

#### [A7] Неполный DI в FastAPI

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `app/dependencies/services.py` — большинство фабрик без constructor injection |
| **Влияние** | Скрытая связность; сервисы самоконструируются; неединообразный wiring; часть графа зависимостей не тестируема |
| **Исправление** | `Depends(get_kp_repository)` + конструкторы |

### Безопасность

#### [S1] IDOR: любой менеджер видит и изменяет чужие КП

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `app/security/offer_access.py:16–22`; `offers_service`, `archive_service`, `delivery_schedule_service`. Подтверждено `tests/test_archive_authorization.py`, `test_offers_production_authorization.py` |
| **Влияние** | Любой пользователь с ролью manager получает доступ ко всем коммерческим предложениям, включая чтение и изменение чужих данных |
| **Исправление** | Фильтр `owner_user_id` или явно задокументировать общий доступ как бизнес-политику |
| **Статус** | **documented (by design)** — 2026-08-28. Политика «общий архив» осознанна: [offer-access-policy.md](../architecture/offer-access-policy.md). Код доступа не менялся. |

#### [S2] Уязвимости Starlette 0.37.2 (транзитивно FastAPI 0.111.0)

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `requirements.txt` — `fastapi==0.111.0`; pip-audit: PYSEC-2026-161, PYSEC-2026-1943, PYSEC-2026-249 и др. |
| **Влияние** | Известные CVE в транзитивных зависимостях; риск эксплуатации через HTTP-стек |
| **Исправление** | Обновить FastAPI/Starlette (Starlette ≥ 0.47.2 / 1.3.1+), pip-audit, регрессия |
| **Статус** | **resolved** — 2026-08-28. `fastapi==0.141.1`, starlette 1.6.0, `pydantic==2.13.4`. pip-audit по fastapi/starlette чист. Полный pytest: 12 failed / 2324 passed — те же 12 pre-existing, регрессий нет. |

#### [S3] Rate limiting in-process

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `login_rate_limit.py`, `commercial_upload_validation.py`. Связано с [A2] |
| **Влияние** | Лимиты brute-force и злоупотребления загрузками обходятся при нескольких workers/репликах |
| **Исправление** | Redis, `UVICORN_WORKERS=1`, мониторить `/health` |

#### [S4] npm-зависимости frontend high

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `frontend/package.json`: nanoid, undici, uuid/exceljs, postcss |
| **Влияние** | Известные уязвимости в frontend-зависимостях |
| **Исправление** | `npm audit fix` |
| **Статус** | **resolved** — 2026-08-28. high 2→0; `npm run audit:ci` exit 0. Транзитивно: nanoid 3.3.18, undici 7.29.0, postcss 8.5.26. Остаток: uuid@8.3.2 через exceljs (moderate; фикс — breaking-даунгрейд exceljs, отложен). |

### Качество кода

#### [Q1] Шесть копий product-type pipeline

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `commercial_workflow_service.py` |
| **Влияние** | Шесть почти идентичных веток обработки типов продуктов; правка в одной не попадает в остальные |
| **Исправление** | `ProductDraftHandler` + config |

#### [Q2] Copy-paste HTTP-обработчиков

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `commercial.py` |
| **Влияние** | Дублированные паттерны обработки запросов; рост объёма и риск расхождения поведения |
| **Исправление** | Decorator + factory |

#### [Q3] ~720 строк мёртвого кода в build_layout_sequence

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `viz_modules/layout_sequence/builder.py` — после `return sequence` ~строка 271, DEPRECATED/UNREACHABLE |
| **Влияние** | Мёртвый код маскирует живую логику; затрудняет рефакторинг и review |
| **Исправление** | Удалить unreachable, разбить живую логику |
| **Статус** | **resolved** — 2026-08-28. Файл 991→~260 строк; sha256 sequence до/после идентичен (`77c686bd…`); layout-тесты 51 passed / 8 skipped. PDF `_p0_baseline/schema_{before,after}.pdf` — визуальный контроль пользователя. |

---

## Проблемы среднего приоритета (запланировать на следующий спринт)

### Архитектура

#### [A8] Параллельные подсистемы планирования

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `core/production/planning.py`, `production_planning_service.py`, `app/planning/plan_manager.py`, `plan_distribution.py`, `plan_distribution_service.py` |
| **Влияние** | Несколько путей планирования с пересекающейся логикой; риск расхождения поведения |
| **Исправление** | Единственный путь `ProductionPlanningService` + `core/production/planning` |

#### [A9] Пустые app-сервисы-реэкспорты

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `kp_persistence_service.py`, `rest_matching_service.py` |
| **Влияние** | Несколько canonical import path; путаница при навигации по кодовой базе |
| **Исправление** | Один canonical import path |

#### [A10] Legacy-пути бота в persistence планов

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `app/planning/plan_storage.py` (BOT_DIR, PLANS_DIR), `run_bot.py` stub, `bot_archived` |
| **Влияние** | Устаревшие пути хранения; смешение web-centric и bot-centric логики |
| **Исправление** | Web-centric paths, удалить `bot_archived` |

#### [A11] ArchiveService god-orchestrator

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `app/services/archive_service.py` ~812 строк |
| **Влияние** | Один сервис объединяет query, export, production readiness; сложно тестировать и изменять |
| **Исправление** | `ArchiveQueryService`, `ArchiveExportService`, `ArchiveProductionReadinessService` |

#### [A12] Frontend god-hook мастера КП

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `frontend/src/features/commercial-offer/hooks/useCommercialOfferWizard.ts` ~453 строк |
| **Влияние** | Вся логика wizard в одном hook; сложно тестировать и расширять |
| **Исправление** | `useDraftLifecycle`, `useProductTypeBranch`, `useWizardMutations` |

### Безопасность

#### [S5] OCR отправляет изображения во внешние LLM

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `core/ocr/pipeline.py`, `commercial_upload_validation.py`, settings `OCR_EXTERNAL_ENABLED` |
| **Влияние** | Коммерческие документы клиентов уходят к стороннему LLM; риски резидентности и конфиденциальности |
| **Исправление** | Default off, политика, self-hosted |

#### [S6] SQLite без шифрования at rest

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `plita.db`, `pb.db` |
| **Влияние** | Данные at rest не защищены при компрометации файловой системы |
| **Исправление** | SQLCipher/OS encryption, encrypted backups |

#### [S7] CSP Report-Only + unsafe-inline

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `app/middleware/security_headers.py` |
| **Влияние** | CSP не блокирует XSS; `unsafe-inline` ослабляет защиту |
| **Исправление** | Enforcing CSP, nonce/hash |

#### [S8] Утечка деталей ошибок в HTTP

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `delivery_schedule.py`, `commercial.py`, `http_errors.py` |
| **Влияние** | Внутренние детали исключений попадают в HTTP-ответы; информационная утечка |
| **Исправление** | Generic detail + traceback только в логах |

#### [S9] CSRF парсит multipart до проверки токена

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `app/middleware/csrf.py:36–47`. Связка с [S2] |
| **Влияние** | DoS-вектор через большие multipart до CSRF-проверки |
| **Исправление** | Token из header, body limits |

#### [S10] Длинная сессия 12ч без refresh-ротации

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `SESSION_COOKIE_MAX_AGE`, `app/security/session.py` |
| **Влияние** | Украденная сессия действует до 12 часов; нет ротации при активности |
| **Исправление** | Короче TTL, sliding expiration |

### Качество кода

#### [Q4] Пять копий build_*_preview_metadata

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `commercial_draft_service.py:262–592` |
| **Влияние** | Дублирование логики preview metadata по типам продуктов |
| **Исправление** | Обобщить через config/handler |

#### [Q5] Дублирование resolve_wide_plates / resolve_unpriced_plates

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `commercial_workflow_service.py` |
| **Влияние** | Почти идентичные функции; риск расхождения |
| **Исправление** | Единая функция с параметрами |

#### [Q6] Две реализации get_global_calendar_info

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `plan_calendar.py` vs `plan_distribution_service.py` |
| **Влияние** | Дублирование; риск расхождения календарной логики |
| **Исправление** | Одна canonical реализация |

#### [Q7] Product-type duplication на фронте

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `commercialOfferApi.ts` + `useCommercialOfferWizard.ts` |
| **Влияние** | Типы продуктов описаны в двух местах |
| **Исправление** | Единый source of truth |

#### [Q8] God-hook useCreatePlanWizardState

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | 724 строки |
| **Влияние** | Вся логика создания плана в одном hook |
| **Исправление** | Декомпозиция на специализированные hooks |

#### [Q9] God-component OfferDetailsDrawer

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | 901 строка, deep ternaries, `as unknown as` |
| **Влияние** | Сложность поддержки; слабая типобезопасность |
| **Исправление** | Разбить на sub-components; убрать type casts |

#### [Q10] Слабая типизация production API

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `production.ts` — `Record<string, unknown>` |
| **Влияние** | Ошибки типов обнаруживаются только в runtime |
| **Исправление** | Строгие TypeScript-интерфейсы |

#### [Q11] preview: Any и dict[str, Any]

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `commercial_draft_service` / workflow |
| **Влияние** | Слабая типизация preview-данных |
| **Исправление** | TypedDict / Pydantic models |

#### [Q12] ArchiveService скрывает частичные сбои

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `except Exception → None` |
| **Влияние** | Частичные ошибки не видны вызывающему коду |
| **Исправление** | Явная обработка/логирование; partial result pattern |

#### [Q13] Нет прямых тестов get_global_calendar_info

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `plan_calendar.py` / `plan_distribution_service.py` |
| **Влияние** | Календарная логика не покрыта unit-тестами |
| **Исправление** | Добавить тесты после консолидации [Q6] |

---

## Низкий приоритет / предложения

### Архитектура

#### [A14] Монолит core/visualization с side-effect import

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `core/visualization/__init__.py` — matplotlib `use('Agg')` at import, ~662 строк |
| **Влияние** | Тяжёлый import-time side effect |
| **Исправление** | Split + lazy import |

#### [A15] Модель авторизации КП не использует owner_user_id

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `app/security/offer_access.py` — только role check; `core/kp/offers_read.py` поддерживает `owner_user_id`. Связано с [S1] |
| **Влияние** | Расхождение модели данных и policy доступа |
| **Исправление** | Задокументировать policy или добавить owner-check |
| **Статус** | **documented** — 2026-08-28. [offer-access-policy.md](../architecture/offer-access-policy.md): `owner_user_id` зарезервирован, в policy не используется. |

### Безопасность

#### [S11] CSRF-cookie не HttpOnly

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `app/security/csrf.py` — double-submit |
| **Влияние** | Ожидаемо для паттерна double-submit cookie |
| **Исправление** | Документировать как осознанный trade-off |

#### [S12] Password policy messages на английском

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `password_policy.py` |
| **Влияние** | UX: сообщения не локализованы |
| **Исправление** | Русские сообщения или i18n |

#### [S13] /health раскрывает метаданные вне production

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `health.py` |
| **Влияние** | Информационная утечка в dev/staging |
| **Исправление** | Минимальный ответ в non-prod или auth |

#### [S14] Legacy bot auth bypass при BOT_AUTH_ENABLED=false

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `bot_archived/middleware/auth.py` |
| **Влияние** | Legacy code path; риск при случайном включении |
| **Исправление** | Удалить вместе с [A10] |

#### [S15] Draft в sessionStorage

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `draftStorage.ts` |
| **Влияние** | Draft доступен из JS; XSS-вектор для данных черновика |
| **Исправление** | Оценить риск; server-side draft preferred |

### Качество кода

#### [Q14] Однострочные delegate-обёртки

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | Различные сервисы |
| **Влияние** | Лишний indirection без пользы |
| **Исправление** | Inline или удалить |

#### [Q15] Имя _merge_plate_texts вводит в заблуждение

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | Соответствующий модуль |
| **Влияние** | Неверное имя затрудняет понимание |
| **Исправление** | Переименовать |

#### [Q16] /parse без response_model

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | Соответствующий endpoint |
| **Влияние** | Нет автодокументации OpenAPI |
| **Исправление** | Добавить response_model |

#### [Q17] GsmGenerationError messages на английском

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | GSM-модуль |
| **Влияние** | Нелокализованные сообщения об ошибках |
| **Исправление** | Русские сообщения |

#### [Q18] Подавление react-hooks/exhaustive-deps в 9 файлах

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | 9 frontend-файлов |
| **Влияние** | Скрытые stale-closure bugs |
| **Исправление** | Исправить deps или refactor |

---

## Матрица приоритетов

| ID | Проблема | Серьёзность | Effort | Приоритет | Статус |
|----|----------|-------------|--------|-----------|
| A1 | Неявное мутабельное глобальное состояние заказа плит | Critical | High | P0 |
| A2 | In-process state блокирует масштабирование | Critical | High | P0 |
| S1 | IDOR: любой менеджер видит чужие КП | High | Low | P0 | **documented (by design)** |
| S2 | Уязвимости Starlette/FastAPI | High | Medium | P0 | **resolved** (fastapi 0.141.1 / starlette 1.6.0) |
| S3 | Rate limiting in-process | High | Medium | P1 |
| S4 | npm-зависимости frontend high | High | Low | P1 | **resolved** (high→0; uuid/exceljs moderate отложен) |
| A3 | God-module CommercialWorkflowService | High | High | P1 |
| A4 | Толстый API-слой КП | High | Medium | P1 |
| A5 | Сервисы обходят repository | High | High | P1 |
| Q1 | Шесть копий product-type pipeline | High | Medium | P1 |
| Q2 | Copy-paste HTTP-обработчиков | High | Low | P1 |
| Q3 | ~720 строк мёртвого кода build_layout_sequence | High | Low | P1 | **resolved** |
| A6 | Planning зависит от visualization | High | Medium | P2 |
| A7 | Неполный DI в FastAPI | High | Medium | P2 |
| A8 | Параллельные подсистемы планирования | Medium | High | P2 |
| A11 | ArchiveService god-orchestrator | Medium | Medium | P2 |
| A12 | Frontend god-hook мастера КП | Medium | Medium | P2 |
| S5 | OCR во внешние LLM | Medium | Low | P2 |
| S8 | Утечка деталей ошибок в HTTP | Medium | Low | P2 |
| S9 | CSRF парсит multipart до проверки | Medium | Medium | P2 |
| Q8 | God-hook useCreatePlanWizardState | Medium | Medium | P2 |
| Q9 | God-component OfferDetailsDrawer | Medium | Medium | P2 |
| A9 | Пустые app-сервисы-реэкспорты | Medium | Low | P3 |
| A10 | Legacy-пути бота в persistence | Medium | Medium | P3 |
| S6 | SQLite без шифрования at rest | Medium | High | P3 |
| S7 | CSP Report-Only + unsafe-inline | Medium | Medium | P3 |
| S10 | Длинная сессия 12ч | Medium | Low | P3 |
| Q4–Q7 | Дублирование preview/plates/calendar/product-type | Medium | Medium | P3 |
| Q10–Q13 | Слабая типизация, скрытые сбои, нет тестов | Medium | Low–Med | P3 |
| A14–A15 | Viz monolith, owner_user_id policy | Low | Low–Med | P4 | A15 **documented** |
| S11–S15 | CSRF cookie, i18n, health, bot, sessionStorage | Low | Low | P4 |
| Q14–Q18 | Delegate wrappers, naming, response_model, i18n, eslint | Low | Low | P4 |

---

## Следующие шаги

### Немедленно (до следующего релиза)

- **[A1]** — завершить миграцию A1-002: явный `PlateOrderContext`, убрать legacy `config_and_data` (`/refactor core/plate_runtime_state.py`)
- **[A2]** — enforce single-worker при старте или Redis shared store (`/implement fix A2 rate-limit store`)
- ~~**[S1]**~~ documented (by design) — [offer-access-policy.md](../architecture/offer-access-policy.md)
- ~~**[S2]**~~ resolved — fastapi 0.141.1 / starlette 1.6.0

### Этот спринт

- **[S3]** — rate limits (S4 закрыт: npm audit high→0)
- **[A3]**, **[A4]**, **[Q1]**, **[Q2]** — декомпозиция коммерческого контура (`/refactor commercial_workflow_service.py`, `/refactor commercial.py`)
- **[A5]**, **[A7]** — repository layer + DI
- ~~**[Q3]**~~ resolved — мёртвый код `builder.py` удалён

### Следующий спринт

- **[A6]**, **[A8]** — развязать planning от visualization; единый путь планирования
- **[A11]**, **[A12]**, **[Q8]**, **[Q9]** — декомпозиция archive и frontend god-components/hooks
- **[S5]–[S10]** — security hardening: OCR policy, CSP, error sanitization, CSRF, session TTL
- **[Q4]–[Q13]** — дедупликация, типизация, тесты календаря

### Бэклог

- **[A9]**, **[A10]**, **[A14]** — cleanup re-exports, bot legacy, viz split
- ~~**[A15]**~~ documented — [offer-access-policy.md](../architecture/offer-access-policy.md)
- **[S11]–[S15]**, **[Q14]–[Q18]** — i18n, health metadata, eslint deps, naming, response_model

Для структурных проблем: `/refactor [file]`.  
Для feature-level security-фиксов: `/implement [fix]`.

---

## Связанные документы

- Предыдущий полный аудит: [2026-08-03-full-project-audit.ru.md](./2026-08-03-full-project-audit.ru.md)
- ADR по деплою: [deployment-single-instance.md](../architecture/deployment-single-instance.md)
- ADR политики доступа к КП: [offer-access-policy.md](../architecture/offer-access-policy.md)
- Спека P0: [stabilizaciya-p0-audit-2026-08-28.md](../../specs/stabilizaciya-p0-audit-2026-08-28.md)
- План P0: [2026-08-28-stabilizaciya-p0-audit.md](../plans/2026-08-28-stabilizaciya-p0-audit.md)
