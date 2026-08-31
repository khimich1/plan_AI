# Отчёт об аудите проекта

**Дата:** 2026-07-13  
**Область:** полный проект (`app/`, `core/`, `frontend/src/`, `viz_modules/`, `bot_archived/`)  
**Аудиторы:** senior-reviewer, security-auditor, reviewer  
**Проект:** «Шишов» — FastAPI + React, оптимизация раскладки ЖБ плит, коммерческие предложения (КП), производственное планирование

---

## Краткое резюме

**Общий Health Score: 2.0 / 10**

Расчёт: `10 − 2×2 (critical) − 3 (high cap) − 1 (medium cap) = 2.0`

| Серьёзность | Архитектура | Безопасность | Качество кода | Итого |
|-------------|-------------|--------------|---------------|-------|
| Critical    | 2           | 0            | 0             | **2** |
| High        | 6           | 5            | 4             | **15** |
| Medium      | 7           | 9            | 10            | **26** |
| Low         | 5           | 8            | 4             | **17** |
| **Всего**   | **20**      | **22**       | **18**        | **60** |

### Общая оценка

Проект функционально зрелый и решает сложную предметную задачу, но архитектурно и операционно находится в переходном состоянии: legacy-слой, глобальное мутабельное состояние, однопоточное хранилище и фрагментированный pipeline планирования создают риски для масштабирования, параллельной работы и безопасности. Критические проблемы связаны с **глобальным состоянием заказа плит** и **одноузловой персистентностью** — обе требуют немедленного внимания до расширения нагрузки или multi-worker деплоя.

### Ключевая рекомендация

Приоритизировать устранение **A1** (глобальное мутабельное состояние) и **A2** (гонки SQLite/JSON), затем консолидировать слой доступа к данным (**A3**) и декомпозировать god-сервисы (**A4**, **Q1**, **Q11**). Параллельно закрыть security high: rate limiting (**S1**, **S10**), OCR-утечки (**S3**, **S14**), legacy attack surface (**S5**).

---

## Критические проблемы (исправить немедленно)

### [A1] Неявное глобальное мутабельное состояние заказа плит

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Расположение** | `core/plate_runtime_state.py`, `core/config_and_data.py`, `core/visualization/__init__.py`, `viz_modules/layout_sequence/from_plan.py`, `viz_modules/procurement/load_context.py` |
| **Связанные находки** | ↔ **S9** (утечка plate runtime в cpu_bound путях) |

**Описание.** Функция `get_plate_mutable_runtime()` предоставляет глобальное мутабельное состояние заказа плит, которое используется по всей цепочке визуализации и закупки. Если при обработке запроса не выполнена корректная привязка (binding) контекста, состояние одного запроса может «протечь» в другой.

**Влияние.**
- Порча данных между параллельными HTTP-запросами
- Недетерминированные результаты оптимизации и визуализации
- Критический блокер для multi-worker / multi-user деплоя

**Рекомендуемое исправление.**
1. Заменить глобальный singleton на **request-scoped контекст** (FastAPI `Depends`, contextvars или явная передача объекта состояния).
2. Убрать неявные side-effect'ы из `get_plate_mutable_runtime()` — все мутаторы должны работать с явно переданным контекстом.
3. Добавить интеграционные тесты на параллельные запросы с разными заказами.
4. Аудит всех call-site'ов в `viz_modules/` и `core/visualization/`.

---

### [A2] Одноузловая персистентность (JSON-черновики + SQLite)

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Расположение** | `app/services/draft_store.py`, `app/repositories/plan_repository.py`, `core/kp_db_common.py` |
| **Связанные находки** | ↔ **S2** (гонки SQLite + JSON), ↔ **S12** (admin destructive ops), ↔ **S13** (черновики без шифрования) |

**Описание.** Черновики КП хранятся как JSON-файлы на диске, планы и данные плит — в SQLite (`plita.db`). Нет транзакционной координации между хранилищами, нет блокировок при concurrent write.

**Влияние.**
- Гонки при одновременной записи планов и черновиков
- Потеря или повреждение данных при multi-worker uvicorn/gunicorn
- Невозможность горизонтального масштабирования без внешнего хранилища

**Рекомендуемое исправление.**
1. Краткосрочно: file locking (fcntl/msvcrt) для JSON-черновиков, WAL mode + busy_timeout для SQLite.
2. Среднесрочно: миграция черновиков в SQLite или PostgreSQL с единой транзакционной моделью.
3. Долгосрочно: внешнее хранилище (PostgreSQL + Redis для сессий/rate limit) для production.

---

## Проблемы высокого приоритета

### Архитектура

#### [A3] Двойная несогласованная модель доступа к данным

| Поле | Значение |
|------|----------|
| **Расположение** | `core/kp_db_*` vs `app/repositories/*` |
| **Связанные находки** | ↔ **A5** (обход repository в production_completion) |

**Описание.** Два параллельных слоя доступа к БД: legacy-модули `core/kp_db_*` и новый слой `app/repositories/*`. Логика дублируется, контракты расходятся, миграция на repository/ports не завершена (**A14**).

**Влияние.** Сложность сопровождения, риск расхождения данных, затруднённое тестирование.

**Исправление.** Определить единый data-access layer; legacy `kp_db_*` пометить deprecated и постепенно переносить в repositories.

---

#### [A4] God-service коммерческого workflow

| Поле | Значение |
|------|----------|
| **Расположение** | `app/services/commercial_workflow_service.py` (~714–766 строк) |
| **Связанные находки** | ↔ **Q11** (766-line façade) |

**Описание.** Единый сервис агрегирует OCR, расчёт, сохранение, экспорт, валидацию и переходы wizard — нарушение SRP.

**Влияние.** Сложность тестирования, высокий риск регрессий при изменениях, затруднённый onboarding.

**Исправление.** Декомпозиция на: `CommercialCalculationService`, `CommercialExportService`, `CommercialDraftService`, `CommercialWizardOrchestrator`.

---

#### [A5] Production completion обходит repository

| Поле | Значение |
|------|----------|
| **Расположение** | `app/services/production_completion_service.py` |
| **Связанные находки** | ↔ **A3**, ↔ **A14** |

**Описание.** Сервис завершения производства использует raw `sqlite3` вместо repository-слоя.

**Влияние.** Обход абстракций, дублирование SQL, риск SQL-инъекций (хотя parameterized queries используются в других местах), несогласованность транзакций.

**Исправление.** Перенести SQL в `PlanRepository` или dedicated `ProductionCompletionRepository`.

---

#### [A6] Фрагментированный pipeline планирования

| Поле | Значение |
|------|----------|
| **Расположение** | `core/production/planning.py`, `app/planning/plan_*.py`, `plan_distribution_service`, `production_planning_service` |

**Описание.** Логика планирования размазана между core и app, между несколькими сервисами и модулями `plan_*`.

**Влияние.** Дублирование, неочевидный flow данных, сложность отладки.

**Исправление.** Выделить единый `PlanningPipeline` с явными этапами: aggregate → distribute → persist → notify.

---

#### [A7] Слабая dependency injection

| Поле | Значение |
|------|----------|
| **Расположение** | `app/dependencies/services.py` |

**Описание.** DI-контейнер минимален; сервисы часто создаются inline или через прямые импорты singleton'ов.

**Влияние.** Затруднённое unit-тестирование, скрытые зависимости, tight coupling.

**Исправление.** Расширить FastAPI `Depends`-граф; внедрять repositories через конструкторы сервисов.

---

#### [A8] Бот выведен из эксплуатации, но всё ещё referenced

| Поле | Значение |
|------|----------|
| **Расположение** | `bot_archived/`, `run_bot.py` (stub), `core/config/settings.py` (загрузка `bot/bot.env`) |
| **Связанные находки** | ↔ **S4** (bot.env при старте) |

**Описание.** Telegram-бот архивирован, но код, конфигурация и env-файлы остаются в проекте и загружаются при старте приложения.

**Влияние.** Мёртвый код, путаница, потенциальная утечка секретов из `bot.env`.

**Исправление.** Удалить загрузку `bot.env` из settings; переместить `bot_archived/` в отдельный repo или archive branch; удалить `run_bot.py` stub.

---

### Безопасность

#### [S1] In-process rate limiting не shared между workers

| Поле | Значение |
|------|----------|
| **Расположение** | `app/security/login_rate_limit.py`, `commercial_upload_validation.py` |
| **Связанные находки** | ↔ **S10** (нет global API rate limiting) |

**Описание.** Rate limiter хранит счётчики в памяти процесса. При multi-worker каждый worker имеет свой лимит.

**Влияние.** Brute-force login и upload abuse обходят лимиты.

**Исправление.** Redis-based rate limiter или nginx/traefik rate limiting на edge.

---

#### [S2] Гонки SQLite + JSON при concurrent access

| Поле | Значение |
|------|----------|
| **Расположение** | `app/services/draft_store.py`, `app/repositories/plan_repository.py`, `core/kp_db_common.py` |
| **Связанные находки** | ↔ **A2** |

**Описание.** Одновременная запись несколькими запросами может повредить файлы черновиков или SQLite DB.

**Влияние.** Потеря данных, corrupted JSON, SQLITE_BUSY errors.

**Исправление.** См. рекомендации A2; добавить retry logic с exponential backoff.

---

#### [S3] OCR отправляет изображения внешним провайдерам

| Поле | Значение |
|------|----------|
| **Расположение** | OCR pipeline (при `OCR_EXTERNAL_ENABLED=true`) |
| **Связанные находки** | ↔ **S14** (GPT snippets в stdout) |

**Описание.** При включённом внешнем OCR коммерческие документы (содержащие цены, контрагентов) отправляются сторонним API.

**Влияние.** Утечка коммерчески чувствительных данных; compliance risk.

**Исправление.** Документировать политику; добавить opt-in на уровне пользователя; рассмотреть on-prem OCR; redact перед отправкой.

---

#### [S4] Decommissioned bot.env загружается при старте

| Поле | Значение |
|------|----------|
| **Расположение** | `core/config/settings.py` |
| **Связанные находки** | ↔ **A8** |

**Описание.** Settings loader читает `bot/bot.env` даже когда бот не используется.

**Влияние.** Потенциальная загрузка устаревших/утёкших секретов; расширенная attack surface.

**Исправление.** Удалить загрузку bot.env; ротация секретов если файл когда-либо был в git.

---

#### [S5] Legacy web routes расширяют attack surface

| Поле | Значение |
|------|----------|
| **Расположение** | `app/web/legacy_routes.py` |
| **Связанные находки** | ↔ **A12** |

**Описание.** Устаревшие web-маршруты остаются активными параллельно с REST API.

**Влияние.** Дополнительные endpoints без современных security controls; потенциально слабее auth.

**Исправление.** Инвентаризация legacy routes; deprecate и удалить неиспользуемые; redirect на React frontend.

---

### Качество кода

#### [Q1] Mega-functions 100–965 строк

| Поле | Значение |
|------|----------|
| **Расположение** | `viz_modules/.../builder.py`, `from_plan.py`, `finalize.py`, `core/visualization/__init__.py`, `breakdown.py` |
| **Связанные находки** | ↔ **A9** (viz_modules god modules) |

**Описание.** Ключевые функции визуализации и оптимизации превышают 100 строк, некоторые — до 965.

**Влияние.** Невозможность unit-тестирования отдельных этапов; высокая cognitive load.

**Исправление.** Extract function refactoring; выделить pipeline stages с явными входами/выходами.

---

#### [Q2] Frontend god components

| Поле | Значение |
|------|----------|
| **Расположение** | `CommercialOfferWizard`, `PlateInputStep`, `DayDrawer`, `useCreatePlanWizardState` |
| **Связанные находки** | ↔ **A13** |

**Описание.** Крупные React-компоненты и hooks совмещают UI, бизнес-логику, API-вызовы и state management.

**Влияние.** Сложность тестирования и рефакторинга; дублирование логики.

**Исправление.** Разделить на container/presentational; вынести hooks для data fetching; использовать composition.

---

#### [Q3] Тонкое покрытие frontend-тестами

| Поле | Значение |
|------|----------|
| **Расположение** | `frontend/src/` (~13 test files vs 100+ sources) |

**Описание.** Менее 15% исходных файлов имеют тесты; критические wizard-компоненты не покрыты.

**Влияние.** Регрессии при рефакторинге wizard и commercial flow.

**Исправление.** Приоритет: wizard steps, plate input validation, API contract tests.

---

#### [Q4] day_documents_service.py без dedicated tests

| Поле | Значение |
|------|----------|
| **Расположение** | `app/services/day_documents_service.py` |
| **Связанные находки** | ↔ **Q7** (duplicate day document generators) |

**Описание.** Сервис генерации дневных документов не имеет unit/integration тестов.

**Влияние.** Риск некорректных производственных документов.

**Исправление.** Добавить тесты с fixture PDF/XLSX output validation.

---

## Проблемы среднего приоритета

### Архитектура

| ID | Проблема | Расположение | Связи |
|----|----------|--------------|-------|
| **A9** | God modules в viz_modules | `from_plan.py`, `builder.py`, `trim.py`, `breakdown.py` | ↔ Q1, Q5 |
| **A10** | Wizard step machine продублирована в 3 местах | `app/schemas/commercial.py`, `commercial_wizard_step_service.py`, `frontend/.../wizardStepOrder.ts` | ↔ Q17 |
| **A11** | Sham re-exports в app-layer | `kp_persistence_service`, `rest_matching_service`, `plate_completion_service` | — |
| **A12** | Параллельная legacy web surface | `app/web/legacy_routes.py` | ↔ S5 |
| **A13** | Frontend god components | `CommercialOfferWizard`, `PlateInputStep`, `DayDrawer`, `useCreatePlanWizardState` | ↔ Q2 |
| **A14** | Repository/ports pattern частично внедрён | `app/repositories/*` vs direct SQL | ↔ A3, A5 |
| **A15** | Lazy imports маскируют circular coupling | `core/kp_db_rests.py` | — |

### Безопасность

| ID | Проблема | Расположение | Связи |
|----|----------|--------------|-------|
| **S6** | Нет max_length на commercial text payloads | `app/schemas/commercial.py` | — |
| **S7** | CSP Report-Only с unsafe-inline | `app/middleware/security_headers.py` | — |
| **S8** | APP_DEBUG=true раскрывает stack traces | `app/main.py` | — |
| **S9** | Residual plate runtime leakage в cpu_bound | `app/concurrency/cpu_bound.py` | ↔ A1 |
| **S10** | Нет global API rate limiting | API layer | ↔ S1 |
| **S11** | PBKDF2 вместо bcrypt/argon2 | `app/repositories/auth_repository.py` | — |
| **S12** | Admin destructive ops гоняются с active writes | Admin endpoints | ↔ A2 |
| **S13** | Commercial drafts не зашифрованы на диске | `app/services/draft_store.py` | ↔ A2 |
| **S14** | OCR печатает GPT snippets в stdout при ошибках | OCR error handlers | ↔ S3 |

### Качество кода

| ID | Проблема | Расположение | Связи |
|----|----------|--------------|-------|
| **Q5** | Duplicate procurement breakdown builders | `viz_modules/procurement/` | ↔ A9 |
| **Q6** | Track-item width extraction скопирован 4+ раз | Multiple viz/plan modules | — |
| **Q7** | Day document generators дублируют flow | Day document services | ↔ Q4 |
| **Q8** | TypeScript `Record<string, unknown>` размывает API contracts | `frontend/src/` API types | — |
| **Q9** | `print()` в hot paths оптимизации/визуализации | `core/`, `viz_modules/` | — |
| **Q10** | Silent error swallowing | `plan_aggregation.py` | — |
| **Q11** | CommercialWorkflowService 766-line façade | `commercial_workflow_service.py` | ↔ A4 |
| **Q12** | offers_service.py без unit tests | `app/services/offers_service.py` | — |
| **Q13** | day_view_service.py minimal coverage | `app/services/day_view_service.py` | — |
| **Q14** | Duplicated completion-percentage fetch | `archive_service.py` | — |

---

## Низкий приоритет / рекомендации

### Архитектура

| ID | Проблема | Расположение |
|----|----------|--------------|
| **A16** | Legacy compatibility layer | `core/config_and_data.py` |
| **A17** | Split domain model core vs app | `plate_order.py` |
| **A18** | Inconsistent service packaging | `day_documents_service.py` |
| **A19** | Presentation leaks into core types в API endpoints | API layer |
| **A20** | CPU-bound concurrency cap непрозрачен | `app/concurrency/cpu_bound.py` |

### Безопасность

| ID | Проблема | Расположение | Примечание |
|----|----------|--------------|------------|
| **S15** | CSRF cookie non-HttpOnly (double-submit pattern) | Auth middleware | By design для double-submit |
| **S16** | Session token — base64 JSON, не encrypted | Session handling | Acceptable при HTTPS |
| **S17** | Wizard state в sessionStorage | Frontend | XSS risk mitigated CSP |
| **S18** | npm audit clean; pip-audit не запускался на Windows | Dependencies | Запустить в CI Linux |
| **S19** | `/health` public с extra metadata в non-prod | `app/main.py` | OK для dev |
| **S20** | Frontend role gating — UI-only mismatch | Frontend auth | Backend must enforce |
| **S21** | Auth error messages via `detail=str(exc)` | Auth endpoints | Acceptable |
| **S22** | offer_access edge case для NULL owner_user_id | Access control | Edge case |

### Качество кода

| ID | Проблема | Расположение |
|----|----------|--------------|
| **Q15** | Repeated exception boilerplate | `commercial.py` endpoints |
| **Q16** | Bare `except: pass` в debug regions | `optimize_2d` |
| **Q17** | Inconsistent wizard step typing | Frontend wizard |
| **Q18** | Deprecated `legacy_runtime()` | `optimization_service.py` |

---

## Положительные моменты

Проект демонстрирует зрелые security practices в ряде ключевых областей:

| Контроль | Описание |
|----------|----------|
| **Cookie sessions HttpOnly/Secure** | Session cookies защищены от XSS exfiltration |
| **CSRF double-submit** | Реализован double-submit pattern для state-changing requests |
| **CORS allowlist** | CORS настроен на явный список origins, не wildcard |
| **Upload validation** | Коммерческие загрузки проходят валидацию типа/размера |
| **File download auth** | Скачивание файлов требует аутентификации |
| **Admin DB reset blocked in prod** | Деструктивные admin-операции заблокированы в production |
| **Parameterized SQL** | SQL-запросы используют параметризацию, SQL injection risk минимален |
| **No dangerouslySetInnerHTML** | Frontend не использует опасный HTML injection |
| **npm audit clean** | Frontend dependencies без известных CVE |

---

## Матрица приоритетов

| ID | Проблема | Серьёзность | Усилия | Приоритет |
|----|----------|-------------|--------|-----------|
| A1 | Глобальное мутабельное состояние plate runtime | Critical | L | **P0** |
| A2 | Одноузловая персистентность JSON+SQLite | Critical | L | **P0** |
| S2 | Гонки SQLite + JSON | High | M | **P0** |
| S9 | Plate runtime leakage в cpu_bound | High | M | **P0** |
| A3 | Двойная модель доступа к данным | High | L | **P1** |
| A4 | God-service commercial workflow | High | L | **P1** |
| A5 | Production completion bypass repository | High | S | **P1** |
| A6 | Фрагментированный planning pipeline | High | L | **P1** |
| A7 | Слабая dependency injection | High | M | **P1** |
| A8 | Bot decommissioned но referenced | High | S | **P1** |
| S1 | In-process rate limiting | High | M | **P1** |
| S3 | OCR external providers | High | M | **P1** |
| S4 | bot.env loaded at startup | High | S | **P1** |
| S5 | Legacy web routes attack surface | High | M | **P1** |
| Q1 | Mega-functions 100–965 lines | High | L | **P1** |
| Q2 | Frontend god components | High | L | **P1** |
| Q3 | Thin frontend test coverage | High | L | **P1** |
| Q4 | day_documents_service no tests | High | M | **P1** |
| S6 | No max_length on payloads | Medium | S | **P2** |
| S7 | CSP Report-Only unsafe-inline | Medium | M | **P2** |
| S8 | APP_DEBUG stack traces | Medium | S | **P2** |
| S10 | No global API rate limiting | Medium | M | **P2** |
| S11 | PBKDF2 vs bcrypt/argon2 | Medium | M | **P2** |
| S12 | Admin destructive ops race | Medium | M | **P2** |
| S13 | Drafts unencrypted on disk | Medium | M | **P2** |
| S14 | OCR GPT snippets to stdout | Medium | S | **P2** |
| A9–A15 | Architecture medium findings | Medium | M | **P2** |
| Q5–Q14 | Code quality medium findings | Medium | M | **P2** |
| A16–A20 | Architecture low | Low | S | **P3** |
| S15–S22 | Security low | Low | S | **P3** |
| Q15–Q18 | Code quality low | Low | S | **P3** |

**Легенда усилий:** S = Small (1–2 дня), M = Medium (3–5 дней), L = Large (1–2 недели)

---

## Следующие шаги

### 1. Немедленно (P0 — эта неделя)

- **[A1]** Внедрить request-scoped plate runtime context; убрать глобальный mutable singleton
- **[A2] + [S2]** Добавить file locking для JSON-черновиков и WAL mode для SQLite
- **[S9]** Аудит cpu_bound paths на утечку plate runtime между запросами
- **[S4] + [A8]** Удалить загрузку `bot.env` и references на decommissioned bot

### 2. Этот спринт (P1 — 2–3 недели)

- **[A3] + [A5] + [A14]** Консолидация data-access layer; миграция raw SQL в repositories
- **[A4] + [Q11]** Декомпозиция `CommercialWorkflowService` на 3–4 сервиса
- **[S1] + [S10]** Redis-based rate limiting для login и API
- **[S5] + [A12]** Инвентаризация и удаление legacy web routes
- **[S3] + [S14]** OCR data policy; убрать GPT snippets из stdout
- **[Q1] + [A9]** Начать рефакторинг mega-functions в viz_modules (extract pipeline stages)
- **[Q3]** Добавить frontend tests для wizard critical path

### 3. Следующий спринт (P2 — 4–6 недель)

- **[A6]** Unified planning pipeline
- **[A7]** Расширить DI через FastAPI Depends
- **[A10] + [Q17]** Single source of truth для wizard steps (backend → frontend codegen)
- **[S6]–[S8], [S11]–[S13]** Security hardening batch
- **[Q2] + [A13]** Frontend component decomposition
- **[Q4]–[Q14]** Test coverage expansion для services

### 4. Бэклог (P3)

- Legacy cleanup: **A16**, **A17**, **Q18**
- CSP enforcement (убрать Report-Only): **S7**
- Password hashing upgrade: **S11**
- Frontend role gating alignment: **S20**
- pip-audit в CI (Linux): **S18**
- Code style: **Q15**, **Q16**, **Q9** (заменить print на logging)

---

## Приложение: карта перекрёстных ссылок

```
A1 (global state) ←→ S9 (cpu_bound leakage)
A2 (single-node)  ←→ S2 (races) ←→ S12 (admin ops) ←→ S13 (unencrypted drafts)
A3 (dual DAL)     ←→ A5 (bypass repo) ←→ A14 (partial ports)
A4 (god service)  ←→ Q11 (766-line façade)
A8 (bot legacy)   ←→ S4 (bot.env)
A9 (viz god)      ←→ Q1 (mega-functions) ←→ Q5 (duplicate builders)
A10 (wizard x3)   ←→ Q17 (inconsistent typing)
A12 (legacy web)  ←→ S5 (attack surface)
A13 (FE god)      ←→ Q2 (same components)
S1 (in-proc RL)   ←→ S10 (no global RL)
S3 (OCR external) ←→ S14 (stdout leak)
```

---

*Отчёт сформирован на основе аудита senior-reviewer, security-auditor и reviewer.  
Следующий полный аудит рекомендуется после закрытия P0 и P1 items.*
