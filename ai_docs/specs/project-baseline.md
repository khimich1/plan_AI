# Spec: Project Baseline — Система «Шишов»

> **Тип:** системная базовая спека (living document)  
> **Фаза SDD:** SPECIFY (утверждение baseline перед feature-spec)  
> **Дата:** 2026-06-18  
> **Обновлено:** 2026-06-20 (remediation trail после P0/P1/P2)  
> **Статус:** черновик на ревью

---

## ASSUMPTIONS I'M MAKING

1. Целевая платформа — **веб-приложение** (React SPA); **Telegram-бот deprecated** (код в репо, не целевой канал в проде).
2. Авторизация web — **HMAC session cookie** (`app_session`), не JWT и не OAuth.
3. Основное хранилище — **SQLite** (`plita.db`, `pb.db`), не PostgreSQL.
4. Деплой по умолчанию — **single instance** (`APP_STORAGE_LAYOUT=single_instance`), горизонтальное масштабирование без shared volume не поддерживается для черновиков.
5. Пользователи — **внутренние сотрудники завода** (менеджеры, производство, админ), не публичный SaaS.
6. OCR КП — опционально через **OpenAI GPT Vision**; без ключа распознавание отключено в коде.
7. Оптимизация раскладки — **PuLP + CBC**; solver должен быть доступен в runtime (Docker: `coinor-cbc`).
8. Документация агента (`ai_docs/`) — **локальная**, в `.gitignore`, не коммитится.
9. В этом запросе **не указана конкретная фича** — данный документ фиксирует текущее состояние системы как baseline для будущих feature-spec.

→ Поправь допущения сейчас или подтверди, что baseline корректен.

---

## Objective

Зафиксировать единый источник правды о системе автоматизации завода ЖБ изделий «Шишов»: домен, архитектура, API, роли, границы слоёв, команды и критерии качества. Документ служит основой для всех последующих feature-spec в `ai_docs/specs/<slug>.md`.

### Пользователи

| Роль | Описание | Основные сценарии |
|------|----------|-------------------|
| `manager` | Менеджер по продажам | Создание КП, архив, скидки |
| `production` | Производство | Планы, календарь, завершение смен, документы дня |
| `admin` | Администратор | Всё выше + сброс БД, восстановление плит, статистика |

### User Stories (текущая система)

- Как **менеджер**, я создаю КП через wizard (текст/OCR/AI), чтобы получить расчёт, PDF/XLSX и сохранить в БД/архив.
- Как **менеджер**, я ищу КП в архиве по номеру или заказчику, редактирую скидку и логистику.
- Как **менеджер**, я переношу КП в производство, чтобы плиты стали доступны для планирования.
- Как **производство**, я создаю и активирую план, заполняю треки, завершаю день и скачиваю схему/формовку/разбивку.
- Как **admin**, я сбрасываю части БД или восстанавливаю «застрявшие» плиты.
- ~~Как **менеджер в Telegram**~~ — **legacy/deprecated**; основной канал — React SPA (бот заморожен с P0).

---

## Tech Stack

### Уже в проекте

| Слой | Технологии |
|------|------------|
| Backend | Python 3, FastAPI 0.111, Pydantic v2, pydantic-settings, uvicorn |
| Persistence | SQLite (`plita.db`, `pb.db`), raw SQL в `core/kp_db_*` |
| Оптимизация | PuLP ≥2.6, CBC solver |
| Документы | pandas, openpyxl, python-docx, reportlab, matplotlib |
| OCR | openai ≥1.0 (GPT Vision) |
| Concurrency | filelock (DraftStore), thread-local plate runtime |
| Frontend | React, Vite, TypeScript, TanStack Query, react-hook-form, zod |
| Frontend tests | Vitest, Testing Library |
| Bot | aiogram (`requirements-bot.txt`) — **deprecated**, fail-closed auth (P1) |
| Security | `offer_access.py` (object RBAC), `login_rate_limit.py` (REST login) |
| Backend tests | pytest, FastAPI TestClient |
| Deploy | Docker, docker-compose |

### Не используется

- ORM (SQLAlchemy) — нет
- JWT / OAuth — нет
- Redis/Postgres в production — только опциональные env-поля, не основной путь

---

## Commands

```powershell
# Корень проекта (Windows)
Set-Location "c:\Users\Роман\Desktop\Шишов"
.\.venv\Scripts\Activate.ps1

# Backend (dev)
uvicorn app.main:app --reload

# Health
curl http://localhost:8000/health
curl http://localhost:8000/api/v1/health

# Тесты backend (все)
pytest tests/ -q

# Тесты backend (примеры по домену)
pytest tests/test_commercial_web_flow.py -q
pytest tests/test_archive_endpoints.py -q
pytest tests/test_production_planning_service.py -q
pytest tests/test_core_no_app_import.py -q

# Frontend
Set-Location frontend
npm run dev          # http://localhost:5173
npm run build        # tsc -b && vite build
npm run typecheck    # tsc --noEmit
npm run test         # vitest run

# Telegram-бот
Set-Location "c:\Users\Роман\Desktop\Шишов"
python run_bot.py

# Docker
docker compose build
docker compose up -d
```

### Переменные окружения (ключевые)

| Переменная | Назначение |
|------------|------------|
| `APP_SECRET_KEY` | HMAC подпись session cookie (мин. 32 символа) |
| `PLITA_DB_PATH` | Путь к `plita.db` |
| `PB_DB_PATH` | Путь к `pb.db` |
| `BACKEND_CORS_ALLOWED_ORIGINS` | CORS для SPA |
| `OPENAI_API_KEY` | OCR через GPT Vision |
| `BOT_TOKEN` | Telegram-бот |
| `DRAFTS_DIR`, `OUTPUTS_DIR`, `PLANS_DIR` | Файловое хранилище |
| `APP_STORAGE_LAYOUT` | `single_instance` \| `shared_volume` |

---

## Project Structure

```
app/                          # Web-слой FastAPI (может импортировать core)
  main.py                     # create_app(), lifespan, CORS, middleware
  api/v1/
    router.py                 # Сборка роутеров
    endpoints/                # auth, commercial, archive, production, admin, ...
  services/                   # Бизнес-логика web (21 сервис)
  repositories/               # Тонкие обёртки над core/kp_db
  schemas/                    # Pydantic request/response
  dependencies/               # auth, plate_context, commercial_draft
  security/session.py         # HMAC cookies
  middleware/                 # PlateMutableRuntimeIsolationMiddleware

core/                         # Домен (НЕ импортирует app/)
  kp_db.py                    # Фасад persistence КП
  kp_db_schema.py             # DDL plita.db
  kp_db_offers.py             # CRUD КП
  kp_db_plates.py             # Жизненный цикл плит
  optimization/               # ILP, geometry, orchestrator
  plate_line_parser.py        # Парсинг строк заказа
  commercial_offer_xlsx.py    # Генерация XLSX
  plan_commit.py              # Коммит плана в БД
  config/settings.py          # Settings (env-backed)

bot/                          # Telegram (aiogram) — DEPRECATED, не в проде
  handlers/                   # legacy; не цель консолидации
  middleware/                 # fail-closed auth (P1)
  services/                   # Тонкий слой над core

frontend/src/
  app/                        # router, layout, providers
  pages/                      # login, commercial-offer-create, archive, production
  features/                   # commercial-offer, commercial-archive, production, auth, admin
  shared/                     # api/httpClient, ui, lib

tests/                        # pytest из корня
viz_modules/                  # Визуализация раскладки, закупки
factory_cost/                 # Расчёт себестоимости
docker/                       # seed БД, entrypoint
ai_docs/specs/                # Спеки (локально)
```

### Архитектурный поток данных

```
frontend/ ──HTTP+cookie──► app/api/v1/endpoints/
                                │
                                ▼
                          app/services/
                                │
                    ┌───────────┴───────────┐
                    ▼                       ▼
            app/repositories/           core/
                    │                       │
                    └───────────┬───────────┘
                                ▼
                          plita.db / pb.db
                                │
                          файлы (drafts, outputs)
                          production_plans (SQLite, authoritative для web)

bot/handlers/ ──► bot/services/ ──► core/ ──► legacy JSON plans (deprecated path)
```

### Инварианты слоёв

- `core/` **не импортирует** `app/` — проверка: `tests/test_core_no_app_import.py`
- Роутер: валидация (Pydantic) → сервис → схема ответа; без бизнес-логики в endpoint
- `PlateOrderContext` + middleware изоляции — отдельное mutable-состояние на каждый HTTP-запрос (P1)
- **Web-планы:** единственный authority — SQLite `production_plans` + `PlanRepository` + optimistic `version` (P0)
- **Планирование web:** `core/production/planning.py` — доменный pipeline; `production_planning_service` — тонкий адаптер (P0)
- Object-level RBAC КП: `offer_access.py`, колонка `owner_user_id` в `kp_meta` (P2); role `production` — gap, см. P3 spec
- Bot deprecated — не цель паритета; fail-closed при misconfiguration (P1)

### Remediation trail (стабилизация 2026-06-19 — 2026-06-20)

| Спринт | Спека | Статус | Ключевой результат |
|--------|-------|--------|-------------------|
| P0 | `stabilizaciya-p0-audit-2026-06-19.md` | closed | SQLite планы, pipeline в core, Q1/Q3, бот deprecated |
| P1 | `stabilizaciya-p1-runtime-security-2026-06-19.md` | closed | PlateOrderContext hot paths, bot fail-closed |
| P2 | `bezopasnost-p2-audit-2026-06-19.md` | closed | Login rate limit, manager RBAC, FE-409 |
| P3 | `stabilizaciya-p3-audit-2026-06-20.md` | **approved** — WP0 legacy login, WP1 production RBAC, WP2 destructive+SQLite |

Аудиты: [`2026-06-19`](../develop/audits/2026-06-19-full-project-audit.md) (Post-P0/P1/P2), [`2026-06-20`](../develop/audits/2026-06-20-full-project-audit.md) (новый backlog).

---

## Domain Model (plita.db — ключевые таблицы)

| Таблица | Назначение |
|---------|------------|
| `KP_offers` | КП: заказчик, менеджер, суммы, условия, logistics_cost |
| `kp_plates` | Позиции плит: размеры, qty, status, plan_id, nomenclature_id |
| `kp_files` | XLSX BLOB и пути |
| `managers` | Справочник менеджеров |
| `users` | Web-пользователи (auth) |
| `plate_rests` | Остатки плит |
| `audit_log` | Аудит операций |
| `production_plans` | Производственные планы: `payload_json`, `version`, `is_active` (P0) |
| `kp_meta` | Метаданные КП incl. `owner_user_id` для object-level RBAC (P2) |

### Статусы плит (kp_plates.status)

- `в производстве` — доступна для планирования
- `в плане` — привязана к plan_id

### Черновик КП (DraftStore)

- Файловое хранилище в `DRAFTS_DIR` (по умолчанию `.app_data/drafts/`)
- Блокировка через `filelock`; timeout → HTTP 503
- Содержит: order, optimization, metadata, wizard_state, files, totals

### Wizard КП — шаги

| Step ID | Описание |
|---------|----------|
| `plates` | Ввод/парсинг/OCR плит |
| `wide-plates` | Разрешение широких плит |
| `manager` | Выбор менеджера |
| `client` | Заказчик, скидка, условия, логистика |
| `result` | Расчёт, preview, генерация файлов, сохранение |

`WizardNextRequiredAction`: `ingest_plates`, `resolve_wide_plates`, `select_manager`, `complete_client_terms`, `post_calculate`.

---

## API Contract (текущее состояние)

Базовый префикс: `/api/v1`

### Auth — `/auth`

| Method | Path | Auth | Описание |
|--------|------|------|----------|
| POST | `/login` | public | Login → Set-Cookie `app_session` |
| POST | `/logout` | public | Очистка cookie |
| GET | `/me` | session | Текущий пользователь |

### Managers — `/managers`

| Method | Path | Roles |
|--------|------|-------|
| GET | `/` | admin, manager |

### Commercial — `/commercial`

| Method | Path | Roles | Описание |
|--------|------|-------|----------|
| POST | `/parse` | admin, manager | Парсинг текста заказа |
| POST | `/drafts` | admin, manager | Создание черновика (text + image OCR) |
| GET | `/drafts/{id}` | admin, manager | Детали черновика |
| PATCH | `/drafts/{id}/plates` | admin, manager | Обновление плит |
| POST | `/drafts/{id}/plates/ai` | admin, manager | AI-распознавание |
| POST | `/drafts/{id}/wide-plates/resolve` | admin, manager | Решения по широким плитам |
| PATCH | `/drafts/{id}/meta` | admin, manager | Метаданные (менеджер, клиент, условия) |
| POST | `/drafts/{id}/calculate` | admin, manager | Расчёт/оптимизация |
| POST | `/generate-preview` | admin, manager | Preview без черновика |
| POST | `/from-form` | admin, manager | Создание из формы |
| POST | `/drafts/{id}/generate-files` | admin, manager | PDF/XLSX/breakdown/schema |
| POST | `/drafts/{id}/save` | admin, manager | Сохранение в БД/архив |
| GET | `/drafts/{id}/breakdown` | admin, manager | Таблица разбивки |
| GET | `/files/{filename}` | admin, manager | Скачивание сгенерированного файла |

### Archive — `/commercial/archive`

| Method | Path | Roles | Описание |
|--------|------|-------|----------|
| GET | `/` | admin, manager | Список (section: archived/production/...) |
| GET | `/search` | admin, manager | Поиск по kp_id или customer (≥2 символа) |
| GET | `/{kp_id}` | admin, manager | Детали КП |
| GET | `/{kp_id}/files/{kind}` | admin, manager | pdf/xlsx/breakdown/schema |
| PATCH | `/{kp_id}/discount` | admin, manager | Скидка |
| PATCH | `/{kp_id}/logistics-cost` | admin, manager | Логистика |
| DELETE | `/{kp_id}` | admin, manager | Удаление |
| POST | `/{kp_id}/move-to-production` | admin, manager | В производство |
| GET | `/{kp_id}/production-estimate` | admin, manager | Оценка для плана |
| GET | `/current-plan/gantt` | admin, manager | Gantt текущего плана |

### Production — `/production`

| Method | Path | Roles | Описание |
|--------|------|-------|----------|
| GET | `/plans` | admin, production | Список планов |
| POST | `/plans` | admin, production | Создать план |
| POST | `/plans/build` | admin, production | Построить план (ILP) |
| GET | `/plans/{id}` | admin, production | Детали плана |
| DELETE | `/plans/{id}` | admin, production | Удалить план |
| POST | `/plans/{id}/activate` | admin, production | Активировать |
| DELETE | `/plans/{id}/tracks/{track_id}` | admin, production | Удалить трек |
| GET | `/calendar` | admin, production | Календарь |
| GET | `/day-occupancy` | admin, production | Занятость дня |
| GET | `/kp-candidates` | admin, production | Кандидаты КП |
| GET | `/days/{date}` | admin, production | Детали дня |
| POST | `/days/{date}/complete` | admin, production | Завершить день |
| GET | `/days/{date}/documents/{schema\|breakdown\|formovka}` | admin, production | Документы дня |
| GET | `/candidates` | admin, production | Кандидаты (legacy alias) |
| GET/PUT | `/work-calendar` | admin, production | Рабочий календарь |

### Admin — `/admin`

| Method | Path | Roles | Описание |
|--------|------|-------|----------|
| GET | `/db/stats` | admin | Статистика БД |
| POST | `/db/reset/full` | admin | Полный сброс |
| POST | `/db/reset/kp-only` | admin | Только КП |
| POST | `/db/reset/plans-only` | admin | Только планы |
| POST | `/db/reset/calendar-only` | admin | Только календарь |
| POST | `/db/recover-plates` | admin | Восстановление плит |

### Health

| Method | Path | Описание |
|--------|------|----------|
| GET | `/health` | Root health |
| GET | `/api/v1/health` | API health |

---

## Frontend Routes

| Path | Страница | Роли |
|------|----------|------|
| `/login` | LoginPage | public |
| `/new` | CommercialOfferCreatePage (wizard) | admin, manager |
| `/archive` | CommercialOfferArchivePage | admin, manager |
| `/production` | ProductionPage | admin, production |

Защита: `ProtectedRoute` + cookie session. API: `shared/api/httpClient.ts` с `credentials: include`.

### Feature-модули

| Feature | Ключевые файлы |
|---------|----------------|
| `commercial-offer` | wizard, `commercialOfferApi.ts`, `wizardDraftStore` |
| `commercial-archive` | поиск, drawer, скидки, move-to-production |
| `production` | календарь, CreatePlanWizard, FillBasket, DayDrawer |
| `auth` | AuthProvider, login, ProtectedRoute |
| `admin` | DbManagementModal, reset dialogs |

---

## Code Style

### Backend endpoint (эталон)

```python
@router.post("/drafts/{draft_id}/save", response_model=CommercialSaveOfferResponse)
def save_commercial_draft(
    draft_id: str,
    payload: CommercialSaveDraftRequest,
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
) -> CommercialSaveOfferResponse:
    workflow = CommercialWorkflowService()
    try:
        return workflow.save_draft(draft_id, payload, owner_user_id=int(user["id"]))
    except DraftNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except DomainValidationError as exc:
        raise_validation_client_error(exc, where="save_commercial_draft")
```

**Конвенции:**
- `from __future__ import annotations`
- Типы на всех публичных функциях
- Ошибки домена → `raise_validation_client_error` / `raise_parse_client_error` / HTTPException
- Shared deps: `REQUIRE_ADMIN_OR_MANAGER = require_roles("admin", "manager")`
- Схемы в `app/schemas/`, сервисы инстанцируются в endpoint (или Depends при необходимости)

### Frontend API (эталон)

```typescript
export const saveDraft = (draftId: string, payload: SaveDraftPayload) =>
  httpClient.post<CommercialSaveResult>(
    `/api/v1/commercial/drafts/${encodeURIComponent(draftId)}/save`,
    payload,
  );
```

**Конвенции:**
- Path alias `@/` → `frontend/src/`
- Типы в `features/<domain>/types/`
- React Query hooks в `hooks/use*Queries.ts`
- Colocated tests: `*.test.ts` / `*.test.tsx`

---

## Testing Strategy

| Уровень | Расположение | Примеры | Когда писать |
|---------|--------------|---------|--------------|
| Unit | `tests/test_*.py` | `test_plate_line_parser.py`, `test_optimization_config.py` | Парсинг, оптимизация, чистые функции |
| Service | `tests/test_*_service.py` | `test_archive_service.py`, `test_production_planning_service.py` | Бизнес-правила |
| Integration API | `tests/test_*_endpoints.py`, `test_commercial_web_flow.py` | TestClient + auth cookies | Новые/изменённые endpoints |
| Boundary | `tests/test_*_boundary.py`, `test_core_no_app_import.py` | Слои, импорты | Любое изменение границ |
| Frontend unit | `frontend/src/**/*.test.ts` | `buildKpPreviewRows.test.ts` | Утилиты, store, hooks |
| Manual | — | Wizard UI, бот, OCR | Сложный UX, визуал |

### Фикстуры

- `tests/conftest.py`: `APP_SECRET_KEY` для pytest, сброс кэша `get_settings()`
- Auth в тестах: `create_session_token` + cookies
- OCR rate limiter: `reset_commercial_ocr_rate_limiter_for_tests`

### Минимальная верификация после изменений

| Область | Команда |
|---------|---------|
| Backend | `pytest tests/test_<релевантный>.py -q` |
| Границы | `pytest tests/test_core_no_app_import.py -q` |
| Frontend (существенные) | `npm run build` + `npm run test` |
| API contract | Проверить `app/schemas/` и OpenAPI `/docs` |

---

## Boundaries

### Always

- Минимальный diff; не трогать несвязанный код
- `router → service → repository`; схемы отдельно
- `core/` не импортирует `app/`
- Валидация входа: Pydantic (backend), zod (frontend forms)
- Секреты в `.env` / `bot.env` — не в коде, не в git
- Windows: `Set-Location`, не `&&` в PowerShell
- Русские сообщения об ошибках для пользователя где уже принято в домене

### Ask first

- Изменение DDL в `core/kp_db_schema.py`
- Новые pip/npm зависимости
- Изменение ролей или модели auth
- Изменение `core/optimization/result_contract.py`
- Деструктивные операции БД (`core/destructive_db_guard.py`)
- Переход на `shared_volume` / multi-instance
- Изменение формата session cookie

### Never

- Коммитить без явной просьбы пользователя
- Хранить секреты в коде
- Смешивать ORM, Pydantic и бизнес-логику в одном файле
- Удалять падающие тесты без согласования
- Редактировать `docker/seed/*.db` без понимания последствий
- Дублировать доменную логику в `bot/handlers/` вместо `core/`

---

## Success Criteria (baseline «готово»)

- [ ] Человек прочитал ASSUMPTIONS и подтвердил или поправил их
- [ ] Документ покрывает 6 core areas SDD: Objective, Commands, Structure, Code Style, Testing, Boundaries
- [ ] API-таблицы соответствуют `app/api/v1/endpoints/*.py` на дату спеки
- [ ] Архитектурные инварианты (`core ↛ app`) задокументированы
- [ ] Baseline принят как основа для feature-spec

---

## Risks & Mitigations

| Риск | Митигация |
|------|-----------|
| SQLite + filelock не масштабируется горизонтально | `APP_STORAGE_LAYOUT=shared_volume` + общий том; иначе single instance |
| Session invalidation при смене `APP_SECRET_KEY` | Документировано; future: server-side sessions / JWT с kid |
| OCR зависит от внешнего API | Rate limit, graceful degrade без ключа |
| Сложность оптимизации | Контракт `result_contract.py`, обширные тесты в `tests/test_optimization_*` |
| Дрейф spec ↔ code | Обновлять baseline при крупных рефакторах; feature-spec ссылается на baseline |

---

## Open Questions

1. **Конкретная фича не указана** — какую feature-spec писать следующей? (заполни блок ЗАДАЧА)
2. Нужно ли версионировать baseline в git отдельно от `ai_docs/` (сейчас в `.gitignore`)?
3. Есть ли планы миграции с SQLite на PostgreSQL?
4. ~~Нужна ли отдельная spec для Telegram-бота vs web parity matrix?~~ — **решено:** бот deprecated (P0); revival — отдельное согласование.
5. Какой целевой deployment: только Docker или также bare-metal Windows на заводе?

---

## Out of Scope (для baseline)

- Детальная спека алгоритма оптимизации (см. `core/optimization/`, отдельная spec при рефакторинге)
- Полная схема `pb.db` (прайсы)
- UI design system / Figma
- Процедуры backup/restore production data
- CI/CD pipeline (если не настроен в репо)

---

## Как создать feature-spec (следующий шаг)

1. Скопировать шаблон из `spec-driven-development` skill
2. Заполнить секцию **ЗАДАЧА**: название, слои, роли, критерии успеха
3. Сохранить в `ai_docs/specs/<slug-фичи>.md`
4. Сослаться на этот baseline для архитектурных ограничений
5. После ревью spec → фаза **PLAN** → **TASKS** → **IMPLEMENT**

### Шаблон slug

`фильтр-архива-по-датам` → `ai_docs/specs/filtr-arhiva-po-datam.md`
