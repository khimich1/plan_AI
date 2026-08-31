# Implementation Plan: P6 Legacy Decommission

**Spec:** [`docs/specs/p6-legacy-decommission.md`](../../../docs/specs/p6-legacy-decommission.md)  
**Created:** 2026-07-06  
**Status:** Ready for review (Phase A); Phase B blocked on grace period + Open Questions

## Overview

Завершить миграцию на единый стек `React SPA → /api/v1/* → SQLite`, убрав Telegram-бот, HTML `/web/*`, JSON-планы на диске и PEP 562 proxy. План разбит на **вертикальные PR-срезы** по спеке: сначала подготовка (пути, аудит, backfill, observability), затем hard delete после 30 дней нулевого `/web/*` трафика.

**WP6-A5 (documentation sync)** в спеке уже закрыт — в Phase A не дублируем, только обновляем план/отчёты по мере выполнения.

## Architecture Decisions

| # | Решение | Rationale |
|---|---------|-----------|
| D1 | Локальные defaults `data/`, Docker остаётся на `/data/*` через env | `docker-compose.yml` уже задаёт `PLANS_DIR=/data/plans`; меняем только dev-defaults в `Settings` |
| D2 | `work_calendar.py` читает путь из `Settings`, не хардкод `bot/data` | Единый источник правды; избегаем рассинхрона с `WORK_CALENDAR_PATH` |
| D3 | Grace period 30 дней перед B1 | Assumption A2 в спеке; блокер для удаления `legacy_routes.py` |
| D4 | Backfill (A4) до удаления legacy read path (B4) | Assumption A4; fuzzy-match нельзя резать до данных |
| D5 | PEP 562 (B5) после удаления бота (B2) | `MUTABLE_LEGACY_NAMES` включает `PLATES_0_*` из `bot_archived` |
| D6 | nginx `/web/`: решение Q3 откладывается; в B1 подготовить оба варианта в плане PR | Ask first в спеке для prod nginx |

## Dependency Graph

```
Task 1 (Settings paths)
    ├── Task 2 (work_calendar path)
    ├── Task 3 (admin/schemas docstrings)
    └── Task 4 (script help text)
            │
            ▼
    [Checkpoint 1: local dev + pytest]
            │
    Task 5 (prod file copy — manual) ──► Task 6 (deploy A2)
            │
            ▼
    Task 7 (prod JSON audit) ──► Task 8 (migrate if needed)
            │
            ▼
    Task 9 (backfill kp_plate_id / day_number)
            │
    [Checkpoint 2: plan data integrity]
            │
    Task 10 (runbook + grace start date) ── parallel with 1–9
            │
            ▼
    [WAIT: 30 calendar days, Q1 answered]
            │
    Task 11 (remove /web/*) ──► Task 12 (delete bot_archived)
            │                         │
            │                         ▼
            │                   Task 13 (admin JSON cleanup)
            │
    Task 9 done ──► Task 14 (legacy plan read path)
            │
            ▼
    Task 15 (PEP 562 phase 2)
            │
    [Checkpoint 3: DoD P6]
```

---

## Task List

### Phase 1: Data paths (WP6-A2) — PR #1

#### Task 1: Перенести defaults `Settings` с `bot/data` на `data/`

**Description:** Обновить дефолтные пути `plans_dir`, `plans_metadata_path`, `current_plan_path`, `work_calendar_path` в `core/config/settings.py` с `PROJECT_ROOT / "bot" / "data" / ...` на `PROJECT_ROOT / "data" / ...`. Убедиться, что `ensure_runtime_dirs()` создаёт `data/` при старте.

**Acceptance criteria:**
- [ ] Все четыре поля указывают на `data/`, не `bot/data`
- [ ] Локальный dev без `.env` использует `data/work_calendar.json` и т.д.
- [ ] Docker compose не требует изменений (env `/data/*` уже переопределяет defaults)

**Verification:**
- [ ] `pytest tests/test_production_planning_service.py tests/test_plan_repository.py -q`
- [ ] `rg "bot/data" core/config/settings.py` — нет совпадений

**Dependencies:** None

**Files likely touched:**
- `core/config/settings.py`
- `tests/` — только если есть assert на старые пути

**Estimated scope:** S (1–2 files)

---

#### Task 2: Убрать хардкод `bot/data` из `work_calendar.py`

**Description:** Заменить модуль-level `CALENDAR_PATH = .../bot/data/work_calendar.json` на чтение из `get_settings().work_calendar_path` (lazy) или функцию `_calendar_path()`. Сохранить fallback на пустой календарь при отсутствии файла.

**Acceptance criteria:**
- [ ] `CALENDAR_PATH` не ссылается на `bot/data`
- [ ] Поведение load/save календаря не регрессирует

**Verification:**
- [ ] `pytest tests/test_production_planning_service.py tests/test_production_completion_service.py -q`
- [ ] `rg "bot/data" core/work_calendar.py` — нет совпадений

**Dependencies:** Task 1

**Files likely touched:**
- `core/work_calendar.py`
- Возможно тесты с monkeypatch `CALENDAR_PATH`

**Estimated scope:** S

---

#### Task 3: Обновить упоминания `bot/data` в admin и schemas

**Description:** Убрать устаревшие ссылки на бота в docstrings и описаниях API: `app/services/admin_service.py` (docstring про «остановите бота»), `app/schemas/admin.py` (`plans_count` description).

**Acceptance criteria:**
- [ ] Docstring `_clear_all_plans` не упоминает бота
- [ ] OpenAPI descriptions не говорят `bot/data/plans`

**Verification:**
- [ ] `pytest tests/test_admin_service.py tests/test_admin_destructive_guard.py -q`
- [ ] `rg "bot/data" app/services/admin_service.py app/schemas/admin.py` — нет

**Dependencies:** Task 1

**Files likely touched:**
- `app/services/admin_service.py`
- `app/schemas/admin.py`

**Estimated scope:** S

---

#### Task 4: Обновить help text в migration-скриптах

**Description:** Привести docstrings/argparse help в `scripts/migrate_plans_to_sqlite.py` и `scripts/backfill_day_number.py` к новым путям (`data/plans` по умолчанию).

**Acceptance criteria:**
- [ ] Default paths в help соответствуют `Settings`
- [ ] Скрипты по-прежнему принимают `--plans-dir` override

**Verification:**
- [ ] `python scripts/migrate_plans_to_sqlite.py --help`
- [ ] `python scripts/backfill_day_number.py --help`

**Dependencies:** Task 1

**Files likely touched:**
- `scripts/migrate_plans_to_sqlite.py`
- `scripts/backfill_day_number.py`

**Estimated scope:** S

---

### Checkpoint 1: After Tasks 1–4

- [ ] `pytest tests/ -q` зелёный
- [ ] `rg "bot/data" core/config/settings.py core/work_calendar.py app/ --glob "*.py"` — только допустимые (archived, comments) или ноль
- [ ] Локально: `data/work_calendar.json` читается (скопировать из `bot/data/` если нужно для dev)
- [ ] **Human review** перед merge PR #1

---

### Phase 2: Prod deploy + data audit (WP6-A2 ops, WP6-A3) — PR #2

#### Task 5: Prod pre-deploy — копирование файлов данных (manual)

**Description:** На VPS **до** деплоя PR #1: скопировать `work_calendar.json`, `plans_metadata.json`, `current_plan.json` (если есть) из старых путей в `/data/` (или подтвердить, что они уже там). Сделать бэкап `/data` (snapshot тома).

**Acceptance criteria:**
- [ ] `WORK_CALENDAR_PATH` на prod указывает на существующий файл после деплоя
- [ ] Бэкап `/data` сделан и задокументирован (дата, путь)

**Verification:**
- [ ] SSH: `ls -la /data/work_calendar.json` (или env `WORK_CALENDAR_PATH`)
- [ ] Smoke: календарь в SPA production view загружается

**Dependencies:** Checkpoint 1

**Files likely touched:** None (ops only)

**Estimated scope:** XS (ops runbook entry)

---

#### Task 6: Prod deploy PR #1

**Description:** Задеплоить изменения paths на `zavodstart.ru`, проверить health и production calendar.

**Acceptance criteria:**
- [ ] Backend стартует без ошибок
- [ ] `pytest` на CI зелёный (pre-merge)

**Verification:**
- [ ] `curl /health` или аналог
- [ ] Логи без ошибок загрузки календаря

**Dependencies:** Task 5

**Estimated scope:** XS (ops)

---

#### Task 7: Аудит JSON-планов на prod (WP6-A3)

**Description:** Проверить том prod: есть ли `*.json` в `/data/plans`, `bot/data/plans`, или только SQLite. Заполнить Q2 в спеке. Сохранить краткий отчёт в `ai_docs/develop/reports/p6-prod-data-audit.md`.

**Acceptance criteria:**
- [ ] Q2 в спеке заполнен
- [ ] Список файлов вне SQLite задокументирован (или «только SQLite»)
- [ ] `current_plan.json` / `plans_metadata.json` — статус: obsolete / in use

**Verification:**
- [ ] `python scripts/check_plan_vs_db.py` на prod/staging (если применимо)
- [ ] Отчёт создан

**Dependencies:** Task 6

**Files likely touched:**
- `ai_docs/develop/reports/p6-prod-data-audit.md` (new)
- `docs/specs/p6-legacy-decommission.md` (Q2 answer only)

**Estimated scope:** S

---

#### Task 8: Миграция оставшихся JSON → SQLite (если Task 7 нашёл файлы)

**Description:** Прогнать `scripts/migrate_plans_to_sqlite.py` на prod или staging с бэкапом. Пропустить, если Task 7 = «только SQLite».

**Acceptance criteria:**
- [ ] Все релевантные JSON-планы в `production_plans` SQLite
- [ ] Исходные JSON архивированы или удалены по согласованию

**Verification:**
- [ ] Повторный `check_plan_vs_db.py` — расхождений нет
- [ ] `pytest tests/test_migrate_plans_to_sqlite.py -q`

**Dependencies:** Task 7

**Files likely touched:** None (ops) или отчёт

**Estimated scope:** S (conditional)

---

### Phase 3: Active plans backfill (WP6-A4) — часть PR #2

#### Task 9: Backfill `kp_plate_id` и `day_number` для активных планов

**Description:** На prod (с бэкапом БД): `scripts/backfill_day_number.py` и при необходимости SQL/скрипт для `kp_plate_id`. Задокументировать критерий «архивный legacy» — только completed/read-only планы.

**Acceptance criteria:**
- [ ] Все **активные** планы: items с `kp_plate_id` где применимо
- [ ] Smoke-план в day_view: `is_legacy=false`
- [ ] Критерий архивного legacy описан в отчёте или спеке

**Verification:**
- [ ] `pytest tests/test_plan_consistency.py tests/test_day_view_service.py -q`
- [ ] Ручная проверка day_view на staging/prod для текущего активного плана

**Dependencies:** Task 8 (или Task 7 if no JSON)

**Files likely touched:**
- `ai_docs/develop/reports/p6-prod-data-audit.md` (append)
- Возможно доработка `scripts/backfill_day_number.py` если SQLite-only gap

**Estimated scope:** M (ops + возможный скрипт)

---

### Checkpoint 2: After Tasks 5–9

- [ ] Q2 заполнен; миграция JSON выполнена или N/A
- [ ] Активные планы без `is_legacy=true` на smoke
- [ ] `pytest tests/ -q` зелёный
- [ ] **Human review** перед merge PR #2 / prod backfill sign-off

---

### Phase 4: Observability (WP6-A1) — PR #3

#### Task 10: Runbook legacy `/web/*` traffic

**Description:** Создать `docs/runbooks/legacy-web-traffic.md`: как считать hits (`grep "Legacy web route" /data/logs/backend.log`), что делать при ненулевом трафике, как продлить grace period.

**Acceptance criteria:**
- [ ] Документирован способ подсчёта hits
- [ ] Runbook: notify users / extend grace / когда можно начинать B1
- [ ] Ссылка из спеки P6 (опционально одна строка в Related)

**Verification:**
- [ ] Ручной grep на prod/staging возвращает ожидаемый формат лога
- [ ] Peer review runbook

**Dependencies:** None (параллельно с Phase 1–3)

**Files likely touched:**
- `docs/runbooks/legacy-web-traffic.md` (new)

**Estimated scope:** S

---

#### Task 11: Зафиксировать дату старта grace period

**Description:** После деплоя observability (логи уже пишутся): записать дату старта 30-дневного окна в спеку (Q4) и в этот план.

**Acceptance criteria:**
- [ ] Q4 в `docs/specs/p6-legacy-decommission.md` заполнен
- [ ] Baseline hits за день 0 задокументирован

**Verification:**
- [ ] Grep логов за день 0 выполнен; число hits записано

**Dependencies:** Task 10, Task 6 (prod logging active)

**Files likely touched:**
- `docs/specs/p6-legacy-decommission.md` (Q4 only)
- `ai_docs/develop/plans/p6-legacy-decommission-plan.md` (status section)

**Estimated scope:** XS

---

### GATE: Phase B prerequisites

**Не начинать Tasks 12–15 пока не выполнено:**

- [ ] 30 календарных дней подряд: ноль `Legacy web route` в prod logs (Q1 = «нет трафика»)
- [ ] Q3 решён: nginx — удалить `/web/` или вечный 301
- [ ] Checkpoint 2 пройден (A4 на prod)
- [ ] Human approval на hard delete

**Ориентировочная дата earliest B1:** Q4 + 30 days

---

### Phase 5: Remove `/web/*` (WP6-B1) — PR #4

#### Task 12a: Удалить backend legacy routes

**Description:** Удалить `app/web/legacy_routes.py`, убрать include из `app/web/router.py`. Удалить или сократить `tests/test_web_legacy_deprecation.py`. Проверить `app/web/shell.py` — удалить `page()` если не используется.

**Acceptance criteria:**
- [ ] Нет `legacy_routes` в router
- [ ] `rg "/web/" app/ --glob "*.py"` — только тесты nginx/доки или ноль
- [ ] `page()` удалён или обоснованно оставлен

**Verification:**
- [ ] `pytest tests/ -q`
- [ ] `npm run build` в `frontend/` (DoD P6)

**Dependencies:** GATE

**Files likely touched:**
- `app/web/legacy_routes.py` (delete)
- `app/web/router.py`
- `app/web/shell.py`
- `tests/test_web_legacy_deprecation.py`

**Estimated scope:** M

---

#### Task 12b: Обновить nginx для `/web/`

**Description:** По решению Q3: удалить `location /web/` из `deploy/nginx-docker-ssl.conf` **или** заменить на `return 301 /commercial-offer/...`. Задеплоить с окном обслуживания (Ask first).

**Acceptance criteria:**
- [ ] Решение Q3 задокументировано в спеке
- [ ] Старые закладки `/web/*` ведут на SPA или 404 предсказуемо

**Verification:**
- [ ] `curl -I https://zavodstart.ru/web/login` — 301 или 404 по выбранной политике
- [ ] SPA и `/api/v1/*` работают

**Dependencies:** Task 12a

**Files likely touched:**
- `deploy/nginx-docker-ssl.conf`
- `frontend/nginx.app.locations.conf` (если используется)

**Estimated scope:** S

---

### Phase 6: Hard delete bot (WP6-B2) — PR #5

#### Task 13: Удалить `bot_archived/` и bot entrypoints

**Description:** Удалить `bot_archived/`, `run_bot.py`, `requirements-bot.txt`, `tests/archived/test_bot_*` (5 файлов). Обновить `scripts/smoke_check.py`, `requirements.txt` если ссылаются на бот.

**Acceptance criteria:**
- [ ] `test ! -d bot_archived`
- [ ] `run_bot.py` удалён
- [ ] Нет `tests/archived/test_bot_*`

**Verification:**
- [ ] `pytest tests/ -q`
- [ ] `rg "bot_archived" app/ core/ tests/ --glob "*.py"` — ноль в active paths

**Dependencies:** Task 12a (рекомендуется; не строго — можно параллельно после GATE)

**Files likely touched:**
- `bot_archived/**` (delete)
- `run_bot.py`, `requirements-bot.txt`
- `tests/archived/*`

**Estimated scope:** M (много удалений, мало логики)

---

#### Task 14: Убрать bot settings из `Settings`

**Description:** Удалить или no-op: `bot_token`, `bot_auth_*`, `BOT_DIR`, `load_dotenv(BOT_DIR/bot.env)`, validators для bot allowlist. Минимальный diff — не трогать unrelated settings.

**Acceptance criteria:**
- [ ] Active code paths не читают `BOT_TOKEN`
- [ ] `pytest tests/ -q` зелёный

**Verification:**
- [ ] `pytest tests/ -q`
- [ ] `rg "bot_token|BOT_TOKEN" app/ core/ --glob "*.py"` — ноль в hot paths

**Dependencies:** Task 13

**Files likely touched:**
- `core/config/settings.py`
- Тесты auth если ссылаются на bot settings

**Estimated scope:** M

---

### Phase 7: Admin cleanup (WP6-B3) — PR #5 (same PR as 13–14)

#### Task 15: Упростить `AdminService._clear_all_plans`

**Description:** Убрать unlink `plans_dir/*.json` из hot path; `_clear_all_plans` только SQLite (+ obsolete metadata если ещё нужны). `get_stats().current_plan_present` не зависит от `current_plan_path` если файл obsolete.

**Acceptance criteria:**
- [ ] Нет перебора `plans_dir.glob("*.json")` в reset
- [ ] Stats отражают SQLite authority

**Verification:**
- [ ] `pytest tests/test_admin_service.py -q`

**Dependencies:** Task 13

**Files likely touched:**
- `app/services/admin_service.py`
- `tests/test_admin_service.py`

**Estimated scope:** S

---

### Phase 8: Legacy plan read path (WP6-B4) — PR #6

#### Task 16: Ограничить или удалить fuzzy / `legacy_identity_qty` hot paths

**Description:** В `day_view_service.py`, `plan_repository.py`, `plan_manager.py`, `plan_distribution_service.py` — удалить legacy branch или пометить `# legacy-archive-only` + `pytest.mark` для архивных сценариев. Добавить/обновить grep-gate: новые планы не получают `is_legacy=true`.

**Acceptance criteria:**
- [ ] `day_view` для новых планов: `is_legacy=false`
- [ ] `test_commit_warns_when_legacy_branch_used_with_tracks_by_day` — legacy недостижим для new plans

**Verification:**
- [ ] `pytest tests/test_plan_commit.py tests/test_plan_consistency.py tests/test_production_completion_service.py -q`

**Dependencies:** Task 9 (A4 on prod), Task 13

**Files likely touched:**
- `app/services/day_view_service.py`
- `app/repositories/plan_repository.py`
- `app/planning/plan_manager.py`
- `app/services/plan_distribution_service.py`
- `tests/test_plan_commit.py`

**Estimated scope:** M

---

### Phase 9: PEP 562 phase 2 (WP6-B5) — PR #6 (same PR as 16)

#### Task 17a: Переписать tests с proxy на runtime API

**Description:** По чеклисту `docs/pep562-config-and-data-decommission.md` Phase 2: `rg` call sites, переписать на `get_plate_mutable_runtime()`.

**Acceptance criteria:**
- [ ] Нет `config_and_data.PLATES_*` в active tests
- [ ] `test_config_and_data_module_semantics.py` зелёный

**Verification:**
- [ ] `pytest tests/test_config_and_data_proxy_boundary.py tests/test_config_and_data_module_semantics.py -q`

**Dependencies:** Task 13 (bot PLATES_0_* consumers gone)

**Files likely touched:**
- `tests/` (несколько файлов)
- `scripts/smoke_check.py`

**Estimated scope:** M

---

#### Task 17b: Удалить `MUTABLE_LEGACY_NAMES` proxy

**Description:** Удалить `__getattr__` proxy в `core/config_and_data.py` и allowlist в `core/plate_runtime_state.py`. Оставить публичный API (`get_config`, `set_plate_lists_from_text`).

**Acceptance criteria:**
- [ ] `MUTABLE_LEGACY_NAMES` пуст или модуль без proxy
- [ ] Grep-gates зелёные

**Verification:**
- [ ] `pytest tests/test_config_and_data_proxy_boundary.py tests/test_config_and_data_module_semantics.py -q`
- [ ] `scripts/smoke_check.py` проходит

**Dependencies:** Task 17a

**Files likely touched:**
- `core/config_and_data.py`
- `core/plate_runtime_state.py`

**Estimated scope:** M

---

### Checkpoint 3: P6 Complete

- [ ] DoD 1: 30 дней ноль `/web/*` (подтверждено логами)
- [ ] DoD 2: `bot_archived/` удалён; нет `tests/archived/`
- [ ] DoD 3: defaults без `bot/data`
- [ ] DoD 4: активные планы с `kp_plate_id`
- [ ] DoD 5: PEP 562 proxy удалён
- [ ] DoD 6: `pytest tests/ -q` + `npm run build` зелёные
- [ ] Completion report: `ai_docs/develop/reports/p6-legacy-decommission-complete.md`

---

## PR Summary (maps to spec)

| PR | Tasks | Spec WP |
|----|-------|---------|
| #1 | 1–4 | WP6-A2 |
| #2 | 5–9 | WP6-A2 ops, A3, A4 |
| #3 | 10–11 | WP6-A1 |
| *(wait 30d)* | — | GATE |
| #4 | 12a–12b | WP6-B1 |
| #5 | 13–15 | WP6-B2, B3 |
| #6 | 16, 17a–17b | WP6-B4, B5 |

---

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Деплой A2 без копии `work_calendar.json` на VPS | High — пустой календарь | Task 5 обязателен до Task 6 |
| Backfill на prod без бэкапа БД | High | Ask first; snapshot `/data` |
| Удаление `/web/*` при живых закладках | Med | Task 10–11, GATE, nginx 301 option (Q3) |
| B4 до A4 | High — broken completion | Task 16 depends on Task 9 |
| PEP 562 removal ломает smoke_check | Med | Task 17a before 17b |
| `rg` gates регрессируют | Med | Checkpoint после каждого PR |

---

## Parallelization

| Parallel safe | Must be sequential |
|---------------|-------------------|
| Task 10 (runbook) ‖ Tasks 1–4 | Task 1 → 2 → 3 → 4 |
| Task 17a test rewrites ‖ Task 16 prep (after review) | GATE → 12 → 13 → 15 |
| Documentation reports ‖ code PRs | Task 9 → 16 |

---

## Open Questions (from spec — fill before Phase B)

| # | Вопрос | Ответ | Owner |
|---|--------|-------|-------|
| Q1 | Есть ли hits на `/web/*` за 30 дней? | | |
| Q2 | JSON-планы на VPS или только SQLite? | | Task 7 |
| Q3 | nginx: удалить `/web/` или 301 навсегда? | | Before Task 12b |
| Q4 | Дата старта grace period | | Task 11 |

---

## Progress Tracking

| Task | Status | Notes |
|------|--------|-------|
| 1–4 | pending | Start here |
| 5–9 | pending | After PR #1 merge |
| 10–11 | pending | Can start anytime |
| 12–17 | blocked | GATE: grace + Q1/Q3 |

**Next action:** Task 1 — перенос defaults в `core/config/settings.py`.
