# Отчёт об аудите проекта

**Дата**: 2026-08-02  
**Область**: весь проект (`app/`, `core/`, `frontend/src/`, `bot/`, `viz_modules/`)  
**Аудиторы**: senior-reviewer + security-auditor + reviewer

---

## Краткое резюме

**Общая оценка здоровья**: 2.0/10

| Severity | Architecture | Security | Code Quality | Total |
|----------|:------------:|:--------:|:------------:|:-----:|
| Critical | 2 | 0 | 0 | **2** |
| High | 6 | 3 | 4 | **13** |
| Medium | 8 | 8 | 12 | **28** |
| Low | 4 | 5 | 5 | **14** |
| **Всего** | **20** | **16** | **21** | **57** |

**Формула Health Score:** старт 10 → Critical −2 каждый (макс. −6) → −4 (2 crit); High −0.5 каждый (макс. −3) → −3 (13 high, cap); Medium −0.1 каждый (макс. −1) → −1 (28 med, cap); Low игнорируются → **10 − 4 − 3 − 1 = 2.0**

**Рекомендация**: Устранить 2 критических архитектурных проблемы до следующего релиза; параллельно закрыть High security (rate limit, OCR data egress, npm CVE) и High quality (дублирование move_to_production, пробелы тестов).

### Контекст

Проект «Шишов» (FastAPI + React + Telegram-бот + viz_modules) демонстрирует зрелые решения в аутентификации, параметризованном SQL и изоляции домена упаковки (`core/shipment_packing/`). Однако совокупность 57 находок указывает на системные проблемы зрелости: god-сервисы без repository-слоя, частичный DI, архитектура хранения состояния, не рассчитанная на multi-instance, и нарастающая связность bounded context «Логистика» с commercial archive. Health Score 2.0 отражает концентрацию критических и высокоприоритетных архитектурных и операционных рисков, требующих внимания перед production-релизом.

---

## Критические проблемы (исправить немедленно)

### [A1] God-модуль ShipmentService

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Расположение** | `app/services/shipment_service.py` (~1464 строк, 40+ методов, 54 обращения к БД) |
| **Impact** | Нарушение SRP: один модуль объединяет CRUD рейсов, propose/confirm, упаковку (`core/shipment_packing`), complete/cancel, экспорт XLSX, SQL-хелперы и маппинг в Pydantic-схемы. Невозможность изолированного тестирования, высокий риск регрессий при любых изменениях логистики |
| **Исправление** | Разделить на `ShipmentRepository`, `ShipmentProposeService`, `ShipmentCompletionService`, `ShipmentExportService`; оставить тонкий orchestrator |

---

### [A2] Архитектура хранения состояния не рассчитана на multi-instance

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Расположение** | `core/kp_db_common.py`, `settings.plita_db_path`; `app/services/draft_store.py`, `APP_STORAGE_LAYOUT=single_instance` |
| **Impact** | Единый SQLite + файловые черновики КП работают только в single-instance deployment. При горизонтальном масштабировании (N workers) — рассинхронизация данных, потеря черновиков, гонки транзакций |
| **Исправление** | Зафиксировать deployment-модель в ADR; для prod — shared FS / Redis / внешняя БД для drafts; добавить startup guard, блокирующий multi-instance при `single_instance` layout |

---

## Высокий приоритет (исправить скоро)

### Архитектура

#### [A3] God-модуль CommercialWorkflowService с Service Locator

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Расположение** | `app/services/commercial_workflow_service.py:29–49` |
| **Impact** | Конструктор создаёт 10+ зависимостей inline — скрытая связность, невозможность подмены в тестах, нарушение DIP |
| **Исправление** | Декомпозиция по use-case; constructor injection через `dependencies/services.py` |

#### [A4] Сервисы application-слоя возвращают Pydantic HTTP-схемы

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Расположение** | `shipment_service.py`, `sgp_service.py`, `carrier_service.py`, `archive_service.py` |
| **Impact** | Смешение transport-слоя (HTTP-схемы) с domain/application-слоем; изменение API-контракта затрагивает бизнес-логику |
| **Исправление** | Domain models в сервисах; маппинг в Pydantic-схемы только в endpoints |

#### [A5] SQL и транзакции в сервисах без repository-слоя

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Расположение** | ShipmentService, SgpService, CarrierService |
| **Impact** | SQL размазан по сервисам; дублирование transaction boilerplate; сложность миграции БД |
| **Исправление** | Ввести repositories в `app/repositories/`; сервисы работают только с domain models |

#### [A6] Скрытая связность через lazy import

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Расположение** | `plan_storage.py:187–190`; `archive_service.py:406–408, 477–479` |
| **Impact** | Lazy import маскирует циклические зависимости; runtime-ошибки при рефакторинге; невозможность статического анализа |
| **Исправление** | Protocols + DI; явные зависимости через constructor |

#### [A7] Частичный Dependency Injection

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Расположение** | `app/dependencies/services.py` — только AuthService получает repository через Depends |
| **Impact** | Несогласованный паттерн DI; остальные сервисы создаются inline или как singletons; сложность тестирования |
| **Исправление** | Единообразные factories для всех сервисов + `dependency_overrides` в тестах |

#### [A8] Bounded context «Логистика» зависит от commercial archive API

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Расположение** | `CreateShipmentDialog` → `archiveApi.search`; backend `app/api/v1/endpoints/archive.py` |
| **Impact** | Coupling между commercial и logistics; изменения archive API ломают логистику; нарушение bounded context |
| **Исправление** | Ввести `/api/v1/logistics/kp-search` — slim read-model с минимальным набором полей |

### Безопасность

#### [S2] Rate limiting только in-process

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Расположение** | `app/security/login_rate_limit.py`, `commercial_upload_validation.py` |
| **Impact** | При N workers лимит ≈ N× — атакующий может обойти rate limit, распределяя запросы между workers |
| **Исправление** | Redis / shared store для счётчиков; или явно зафиксировать `workers=1` в deployment-модели |

#### [S3] OCR отправляет коммерческие документы во внешние LLM

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Расположение** | `core/ocr/*`, commercial endpoints; флаг `OCR_EXTERNAL_ENABLED` |
| **Impact** | Утечка коммерчески конфиденциальных документов (цены, условия, контрагенты) во внешние LLM-провайдеры без явного согласия |
| **Исправление** | Data classification; opt-in per tenant; audit trail отправок; on-prem OCR по умолчанию |

#### [S4] npm high CVE в react-router-dom

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Расположение** | `react-router-dom@7.18.0` — GHSA-qwww-vcr4-c8h2 |
| **Impact** | Известная уязвимость в frontend-зависимости; потенциальный XSS или redirect-атака |
| **Исправление** | Обновить `react-router-dom` до patched-версии; добавить `npm run audit:ci` в CI pipeline |

### Качество кода

#### [Q1] Дублирование move_to_production с расходящимся поведением

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Расположение** | `archive_service.py` (strict) vs `offers_service.py` (default_if_empty) |
| **Impact** | Разное поведение одной операции в зависимости от entry point; silent data loss или unexpected errors |
| **Исправление** | Единая реализация в domain-сервисе; оба entry point делегируют одному методу |

#### [Q2] Критический UI-путь логистики без тестов

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Расположение** | `ShipmentItemsSection.tsx` (~644 строк), `draftItems.ts` |
| **Impact** | Основной пользовательский flow создания/редактирования позиций рейса не покрыт тестами; высокий риск регрессий |
| **Исправление** | Unit-тесты для `draftItems.ts`; component-тесты для ShipmentItemsSection (propose, confirm, edit flows) |

#### [Q3] Недостаточное покрытие SGP-мутаций

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Расположение** | `sgp_service.py` — unlink/relink/send_to_sgp; `test_sgp_service.py` — только 5 тестов |
| **Impact** | Критические warehouse-операции без тестового покрытия; риск потери связей плит с КП |
| **Исправление** | Добавить тесты для unlink, relink, send_to_sgp, edge cases (пустой склад, дубликаты) |

#### [Q4] Расхождение обогащения прогресса КП

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Расположение** | `offers_service.py` vs `archive_service.py` — разные статусы для `completion_percentage` |
| **Impact** | Пользователь видит разный прогресс одного КП в offers и archive; путаница и потеря доверия к UI |
| **Исправление** | Единая функция расчёта completion в domain-слое; оба сервиса делегируют ей |

---

## Средний приоритет (следующий спринт)

### Архитектура

| ID | Проблема | Расположение | Исправление |
|----|----------|--------------|-------------|
| **A9** | SGP под REST namespace `/production` | `app/api/v1/endpoints/production.py` | Перенести в `/api/v1/sgp/*` или `/warehouse/sgp/*` |
| **A10** | Frontend cross-feature coupling logistics → commercial-archive | `CreateShipmentDialog` imports `archiveApi` | `logisticsApi.searchKp()` |
| **A11** | God React-компонент ShipmentItemsSection (~644 строк) | `frontend/src/features/logistics/components/ShipmentItemsSection.tsx` | Split на list / propose / editor |
| **A12** | Широкая инвалидация React Query | `useLogisticsQueries` инвалидирует archive/production/sgp keys | Scoped invalidation по затронутым ключам |
| **A13** | Параллельные слои планирования | `app/planning/`, `production_planning_service.py`, `core/production/` | Единый orchestrator |
| **A14** | ArchiveService god orchestrator (~591 строк) | `app/services/archive_service.py` | Query / Export / Mutation split |
| **A15** | core/kp_db.py god-facade | `core/kp_db.py` | Импорт из submodules; deprecate facade |
| **A16** | Legacy PlateMutableRuntime — implicit shared mutable state | middleware isolation | Mandatory context на всех entry points; migrate to explicit context |

### Безопасность

| ID | Проблема | Расположение | Исправление |
|----|----------|--------------|-------------|
| **S5** | CSP Report-Only с unsafe-inline | `security_headers.py` | Enforcing CSP + nonce |
| **S6** | Нет rate limiting на мутирующих business API | business endpoints | Per-endpoint или per-user rate limits |
| **S7** | Данные at rest без шифрования | SQLite + drafts plaintext | Encryption at rest для prod |
| **S8** | Длинная сессия 12ч без idle timeout | `session.py` | Idle timeout 30–60 мин |
| **S9** | APP_DEBUG не блокируется в production | `main.py` / `settings.py` | Startup guard: reject if DEBUG in prod |
| **S10** | CSRF-cookie httponly=False | `csrf.py` | httponly=True (double-submit pattern сохраняется) |
| **S11** | Нет автоматического dependency scanning в CI | CI pipeline | npm audit + pip-audit в CI |
| **S12** | Logistics видит КП всех менеджеров через /archive/search | `/archive/search` — enumeration kp_id + customer | Отдельный endpoint с scoped ACL |

### Качество кода

| ID | Проблема | Расположение | Исправление |
|----|----------|--------------|-------------|
| **Q5** | DRY: SQLite transaction boilerplate ~9× | `shipment_service.py` (+ sgp, carrier) | Transaction context manager / decorator |
| **Q6** | DRY: двойная проверка доступности СГП | `shipment_service.py` | Единая функция availability check |
| **Q7** | DRY: SQL kp_plates matching 4× | `sgp_service.py` | Extract shared query helper |
| **Q8** | Магические строки статусов | multiple services | `KpStatus` / `PlateStatus` enums |
| **Q9** | Несогласованная обработка ошибок | domain errors vs ValueError strings в offers | Unified error hierarchy |
| **Q10** | Тихое проглатывание ошибок | `ArchiveService._to_list_item` | Log + propagate или skip с метрикой |
| **Q11** | list_shipments.count = размер страницы, не total | `shipment_service.py` | Отдельное поле total_count |
| **Q12** | DRY: formatDims дублируется на фронте | frontend logistics + production | Shared utility |
| **Q13** | Пробелы frontend-тестов commercial-archive | MoveToProductionDialog, ArchiveOfferList, DiscountEditDialog | Component tests |
| **Q14** | Сложность _propose_v2_packing ~110 строк | `shipment_service.py` | Extract sub-functions |
| **Q15** | Untyped auth context user: dict | frontend auth context | Typed User interface |
| **Q16** | eslint-disable exhaustive-deps | logistics UI components | Fix deps или extract stable callbacks |

---

## Низкий приоритет / предложения

### Архитектура

| ID | Проблема | Расположение | Исправление |
|----|----------|--------------|-------------|
| **A17** | Re-export wrappers без ценности | plate_completion_service, rest_matching_service, kp_persistence_service | Inline imports или удалить wrappers |
| **A18** | sys.path.insert hack | `plan_manager.py:16–18` | Proper package structure |
| **A19** | bot_archived — параллельный presentation-слой | `bot_archived/` | Archive or merge into active bot |
| **A20** | Ручное дублирование API-контрактов frontend ↔ backend | types/logistics.ts vs schemas/logistics.py | OpenAPI codegen или shared types |

### Безопасность

| ID | Проблема | Расположение | Исправление |
|----|----------|--------------|-------------|
| **S13** | CORS allow_headers=["*"] | `main.py:66` | Explicit allowed headers |
| **S14** | session_version в публичных auth-ответах | auth endpoints | Remove from public response |
| **S15** | Нет Permissions-Policy | security headers | Add Permissions-Policy header |
| **S16** | Роль production имеет доступ к /api/v1/managers | RBAC config | Restrict to admin/manager roles |
| **S17** | Legacy POST /web/login без CSRF в форме | web login template | Add CSRF token to form |

### Качество кода

| ID | Проблема | Расположение | Исправление |
|----|----------|--------------|-------------|
| **Q17** | Неиспользуемый параметр actor | `cancel()` in shipment_service | Remove or use for audit |
| **Q18** | Module-level draftKeyCounter | `draftItems.ts` | Move inside hook/factory |
| **Q19** | DRY generate_pdf/generate_xlsx | `offers_service.py` | Shared document generation helper |
| **Q20** | Минимальные auth-тесты (2 теста) | auth test suite | Expand coverage |
| **Q21** | Мёртвый файл _tmp_old.py | project root | Delete |

---

## Матрица приоритетов

| ID | Проблема | Severity | Effort | Priority |
|----|----------|----------|--------|----------|
| A1 | God-модуль ShipmentService | Critical | High | **P0 — немедленно** |
| A2 | Single-instance storage architecture | Critical | High | **P0 — немедленно** |
| S2 | In-process rate limiting | High | Medium | **P1 — этот спринт** |
| S3 | OCR data egress to external LLM | High | Medium | **P1 — этот спринт** |
| S4 | npm CVE react-router-dom | High | Low | **P1 — этот спринт** |
| Q1 | Дублирование move_to_production | High | Medium | **P1 — этот спринт** |
| Q2 | UI логистики без тестов | High | Medium | **P1 — этот спринт** |
| Q3 | SGP-мутации без тестов | High | Medium | **P1 — этот спринт** |
| Q4 | Расхождение completion_percentage | High | Low | **P1 — этот спринт** |
| A3 | CommercialWorkflowService god module | High | High | **P1 — этот спринт** |
| A4 | Pydantic schemas в сервисах | High | Medium | **P1 — этот спринт** |
| A5 | SQL без repository-слоя | High | High | **P1 — этот спринт** |
| A6 | Lazy import coupling | High | Medium | **P1 — этот спринт** |
| A7 | Частичный DI | High | Medium | **P1 — этот спринт** |
| A8 | Logistics → archive API coupling | High | Medium | **P1 — этот спринт** |
| A9–A16 | Architecture medium findings | Medium | Mixed | **P2 — следующий спринт** |
| S5–S12 | Security medium findings | Medium | Mixed | **P2 — следующий спринт** |
| Q5–Q16 | Quality medium findings | Medium | Mixed | **P2 — следующий спринт** |
| A17–A20 | Architecture low | Low | Low | **P3 — бэклог** |
| S13–S17 | Security low | Low | Low | **P3 — бэклог** |
| Q17–Q21 | Quality low | Low | Low | **P3 — бэклог** |

---

## Следующие шаги

### 1. Немедленно (до следующего релиза)

- **[A1]** Начать декомпозицию `ShipmentService` — выделить `ShipmentRepository` как первый шаг (`/refactor app/services/shipment_service.py`)
- **[A2]** Зафиксировать deployment-модель в ADR: single-instance vs multi-instance; добавить startup guard при несовместимой конфигурации

### 2. Этот спринт

- **[S4]** Обновить `react-router-dom`; добавить `npm run audit:ci` в CI (`/implement security/npm-audit-ci`)
- **[S2]** Rate limiting через Redis или явный `workers=1` в deployment (`/implement security/rate-limit-redis`)
- **[S3]** OCR: opt-in + audit trail + on-prem default (`/implement security/ocr-data-classification`)
- **[Q1]** Унифицировать `move_to_production` — единая domain-реализация (`/implement fix/move-to-production-unify`)
- **[Q2], [Q3]** Тесты для ShipmentItemsSection, draftItems, SGP-мутаций (`/orchestrate test-coverage-logistics-sgp`)
- **[Q4]** Единый расчёт completion_percentage (`/implement fix/completion-percentage`)
- **[A8]** Slim endpoint `/api/v1/logistics/kp-search` (`/implement logistics/kp-search-endpoint`)

### 3. Следующий спринт

- **[A3–A7]** Repository-слой, единый DI, декомпозиция god-сервисов (`/refactor` + `/orchestrate architecture-layering`)
- **[A9–A16]** Namespace SGP, split React-комponents, scoped React Query invalidation
- **[S5–S12]** CSP enforcing, business API rate limits, session idle timeout, CI dependency scanning
- **[Q5–Q16]** DRY рефакторинг, enum-статусы, frontend test gaps

### 4. Бэклог

- **[A17–A20]** Cleanup wrappers, sys.path hack, bot_archived, API codegen
- **[S13–S17]** CORS headers, Permissions-Policy, RBAC tightening
- **[Q17–Q21]** Minor cleanup, auth test expansion, delete `_tmp_old.py`

---

## Маршрутизация исправлений

| Тип проблемы | Команда | Примеры |
|--------------|---------|---------|
| Структурные / архитектурные | `/refactor [file]` | A1 ShipmentService, A3 CommercialWorkflow, A14 ArchiveService |
| Security / behavioral fixes | `/implement [fix]` или `/orchestrate` | S2 rate limit, S3 OCR, S4 npm CVE |
| Тестовое покрытие | `/orchestrate test-coverage-*` | Q2, Q3, Q13 |
| ADR / deployment decisions | Документировать в `ai_docs/develop/architecture/` | A2 deployment model |

---

## Связанная документация

- Конфигурация аудитов: `.cursor/config.json` → `documentation.paths.audits`
- Сравнение с предыдущим аудитом: `ai_docs/develop/audits/2026-08-02-audit-comparison.md`
- Спецификации логистики: `ai_docs/specs/shipment-logistics.md`, `ai_docs/specs/shipment-propose-v2.md`
- План стабилизации: `ai_docs/specs/stabilizaciya-p0-p1-audit-2026-08-02.md`

---

*Отчёт сгенерирован documenter-агентом по результатам аудита senior-reviewer + security-auditor + reviewer. Health Score рассчитан по формуле audit-workflow skill.*
