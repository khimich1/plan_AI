# PLAN: Стабилизация P0 (аудит 2026-06-19)

> **Фаза SDD:** PLAN → **закрыт (implemented)** — 2026-06-19
> **Дата:** 2026-06-19
> **Спека:** [`../../specs/stabilizaciya-p0-audit-2026-06-19.md`](../../specs/stabilizaciya-p0-audit-2026-06-19.md)
> **Baseline:** [`../../specs/project-baseline.md`](../../specs/project-baseline.md)
> **Источник:** [`../audits/2026-06-19-full-project-audit.md`](../audits/2026-06-19-full-project-audit.md)

---

## Завершение спринта (2026-06-19)

| WP | Статус | Находка |
|----|--------|---------|
| WP0 | ✅ done | Structured errors (backend + frontend) |
| WP1 | ✅ done | Q1 — остатки, транзакции, 422/500 |
| WP2 | ✅ done | Q3 — unpriced_plates, без fallback |
| WP3 | ✅ done | A2 — DDL + PlanRepository |
| WP4 | ✅ done | A2 — миграция `bot/data/plans/` → SQLite |
| WP5 | ✅ done | A2 — web без file I/O, `version` в API |
| WP6 | ✅ done | A1 — pipeline validate→load→optimize в `core/` |
| WP7 | ✅ done | A1 — persist + тонкий web-адаптер |

**Верификация:** `pytest tests/ -q` — **726 passed**, 0 failed. Бот помечен deprecated (`bot/README.md`). `A3` и `S1` — следующий спринт.

---

## 0. Резюме плана

Два этапа, строго последовательно:

- **Этап A — данные:** структурированные ошибки → `Q1` (остатки) → `Q3` (цены) → `A2` (планы в SQLite с `version`, миграция, отключение file I/O).
- **Этап B — консолидация:** `A1` (pipeline планирования в `core/`, `persist` через `PlanRepository` из `A2`).

`A3` (runtime-globals) и удаление бота — **вне плана** (deferred / follow-up).

### Граф зависимостей

```
WP0 (error schema, backend+frontend)
        │
        ├──► WP1 (Q1: остатки)        [нужен WP0 для 422/500 контракта]
        ├──► WP2 (Q3: цены)           [нужен WP0 для 422 unpriced_plates]
        │
WP3 (A2: схема production_plans + PlanRepository on SQLite)
        │
        ├──► WP4 (A2: миграция bot/data/plans → SQLite)   [ДО WP5]
        │            │
        │            ▼
        └──► WP5 (A2: отключение file I/O в web + version в API)
                     │
                     ▼
WP6 (A1: pipeline validate→load→optimize)
                     │
                     ▼
WP7 (A1: persist через PlanRepository + web как тонкий адаптер)
```

**Параллельно** можно вести: `WP1`, `WP2` (после `WP0`) и `WP3` (независим от Q-фиксов).
**Строго последовательно:** `WP4 → WP5` (миграция до отключения файлов), `WP3 → WP4`, `WP5/WP3 → WP7`.

---

## Этап A — Стабилизация данных

### WP0 — Структурированные ошибки (фундамент для Q1/Q3/A2)

**Зачем:** сейчас `app/core/http_errors.py` отдаёт только `detail: string`; фронт (`httpClient.parseError`) читает только `detail`. Для `unpriced_plates` / `plan_version_conflict` нужен машиночитаемый `code` + payload.

**Backend:**
- Создать `app/schemas/errors.py`: Pydantic-модель `ApiErrorBody { code: str, message: str, details: dict | None }` (+ типовые `code`-литералы).
- Расширить `app/core/http_errors.py`: хелпер `raise_structured_error(*, status_code, code, message, details=None, where)` — логирует server-side, отдаёт structured body. Существующие `raise_*` оставить (обратная совместимость).
- Зарегистрировать/проверить, что `HTTPException(detail=<dict>)` отдаётся как JSON-объект (FastAPI это умеет; зафиксировать формат тела).

**Frontend:**
- Расширить `frontend/src/shared/lib/apiError.ts` (`ApiError`): добавить поля `code?: string`, `details?: unknown`.
- Обновить `httpClient.parseError`: читать `{ code, message, details }`, fallback на старый `detail: string`.

**Files (~4–5):** `app/schemas/errors.py` (new), `app/core/http_errors.py`, `frontend/src/shared/lib/apiError.ts`, `frontend/src/shared/api/httpClient.ts` (+ тест).
**Verify:** unit-тест backend на форму тела; `npm run build` + `npm run test`.
**Зависимости:** нет. Делать первым.

---

### WP1 — Q1: тихая потеря данных при создании остатков

**Где:** `app/services/production_completion_service.py` (~164–178).
- Убрать `except: pass` вокруг записи в `plate_rests`.
- Обернуть запись остатков в транзакцию (rollback при частичном сбое).
- Ошибка валидации (невалидный `kp_id`) → **422** через structured error (`code: "rest_validation_failed"`); ошибка БД → **500**.
- Лог `ERROR` с `plan_id` + контекст плиты.

**Files (~2–3):** сервис + тест `tests/test_production_completion_service.py` (расширить).
**Verify:** `pytest tests/test_production_completion_service.py -q` — mock-сбой → ошибка (422/500), нет частичного состояния.
**Зависимости:** `WP0`.

---

### WP2 — Q3: ошибки ценообразования маскируются fallback

**Где:** `core/commercial_offer.py` (~195–226), дубликат `core/commercial_offer_xlsx.py` (~77–108).
- Удалить тихий fallback `area * 4000`.
- `get_plate_price()` при отсутствии цены → собрать список непрорасценённых позиций; поток генерации поднимает доменную ошибку.
- В endpoint генерации/сохранения → **422** `{ code: "unpriced_plates", details: { positions: [...] } }`.
- `print()` → `logging` (≥ `WARNING`).
- Блокировать экспорт PDF/XLSX до устранения.
- (Минимизировать дубль примитива; полный DRY-вынос `Q8` — вне scope.)

**Files (~3–5):** `core/commercial_offer.py`, `core/commercial_offer_xlsx.py`, endpoint в `app/api/v1/endpoints/commercial.py`, тест `tests/test_commercial_pricing_errors.py` (new).
**Verify:** `pytest tests/test_commercial_pricing_errors.py -q` — позиция без цены → 422 `unpriced_plates` со списком; документ не генерируется.
**Зависимости:** `WP0`.

---

### WP3 — A2 (часть 1): схема `production_plans` + `PlanRepository` на SQLite

**Где:** `core/kp_db_schema.py` (DDL), `app/repositories/plan_repository.py` (сейчас passthrough к `plan_manager`).
- DDL: `production_plans(id TEXT PK, payload_json TEXT, version INTEGER NOT NULL DEFAULT 1, is_active INTEGER, created_at, updated_at)`. План — JSON в `payload_json` (нормализация треков — позже, `A7`). **Ask first перед DDL.**
- Реализовать новые методы `PlanRepository` поверх SQLite (`core/kp_db_common.py` connection, WAL+FK):
  - `get(plan_id) -> {payload, version}`; `list_metadata`; `get_active/set_active`.
  - `save(payload, expected_version) ` → `UPDATE … WHERE id=? AND version=?`, `version=version+1`; если 0 строк → `PlanVersionConflict`.
  - `create(payload)`; `delete`.
- Endpoint-слой при `PlanVersionConflict` → **409** `{ code: "plan_version_conflict" }`.

**Files (~3–5):** `core/kp_db_schema.py`, `app/repositories/plan_repository.py`, исключение (`PlanVersionConflict`), тест `tests/test_plan_repository.py` (new).
**Verify:** `tests/test_plan_repository.py` — CRUD + параллельная запись: устаревшая `version` → conflict; данные не перезаписаны.
**Зависимости:** `WP0` (для 409-контракта). Независим от `WP1/WP2` → может идти параллельно.

---

### WP4 — A2 (часть 2): миграция `bot/data/plans/` → SQLite (ДО отключения файлов)

**Порядок критичен:** миграция выполняется и проверяется **раньше** `WP5`.
- Скрипт `scripts/migrate_plans_to_sqlite.py`: читает `bot/data/plans/*.json` + `plans_metadata.json`, пишет в `production_plans` через `PlanRepository.create`, проставляет `is_active`, `version=1`.
- Idempotent + dry-run (`--dry-run`), бэкап исходных файлов, отчёт (кол-во/ошибки).
- Сверка: для каждого плана `load из БД == исходный JSON` (по ключевым полям).

**Files (~1–2):** `scripts/migrate_plans_to_sqlite.py` (new) + тест на копии фикстур.
**Verify:** прогон `--dry-run` без ошибок; реальный прогон на копии → счётчики совпадают, обратимость подтверждена бэкапом.
**Зависимости:** `WP3`.

---

### WP5 — A2 (часть 3): отключение file I/O в web + `version` в API

- Перевести web-пути (`app/services/production_planning_service.py`, `app/planning/plan_manager.py` callers, `archive_service`, `day_*_service`) на `PlanRepository` (SQLite). Прямой `open(...,'w')`/`json.dump` планов из web убрать (закрывает `S8`).
- `plan_manager` file I/O планов больше не вызывается из web (сам модуль не удаляем в этом WP — это `A7`).
- **Добавить `version` в API-ответы плана:** `GET /production/plans/{id}` (и в list, где уместно) — поле `version` для клиентского reload при 409. Обновить response-схему в `app/schemas/` и фронтовые типы `frontend/src/features/production/types/production.ts`.
- Записи планов из endpoints передают `expected_version`.

**Files (~5):** production service + затронутые services, `app/schemas/` (plan response), фронтовые типы, тест.
**Verify:** `pytest tests/test_production_planning_service.py -q`; ручная проверка `GET /production/plans/{id}` содержит `version`; конкурентный PATCH → 409.
**Зависимости:** `WP3`, `WP4` (миграция должна пройти раньше).

---

## Этап B — Консолидация логики (A1)

> Начинать только когда Этап A зелёный.

### WP6 — A1 (часть 1): pipeline `validate → load → optimize` в `core/`

**Где:** новый `core/production/planning.py`; источник правил — `app/services/production_planning_service.py` (~1030 строк), метод `build_plan` (~74–348).
- Вынести чистые функции по фазам: `validate(input)`, `load(...)`, `optimize(...)` с явными контрактами данных (dataclass/Pydantic-DTO в `core/`).
- `core/` **не импортирует** `app/` (проверка `tests/test_core_no_app_import.py`).
- Бот **не трогаем**.

**Files (~3–5):** `core/production/planning.py` (new) + DTO + тест `tests/test_core_production_planning.py` (new).
**Verify:** `pytest tests/test_core_production_planning.py -q`; `tests/test_core_no_app_import.py` зелёный.
**Зависимости:** Этап A завершён.

---

### WP7 — A1 (часть 2): `persist` через `PlanRepository` + web как тонкий адаптер

- Добавить в pipeline шаг `persist(plan, repo)` — пишет через `PlanRepository` (SQLite, `version`).
- `production_planning_service.build_plan` → тонкий адаптер: собирает вход, вызывает `core/production/planning` фазы, возвращает schema-ответ. Своей копии правил планирования не держит.
- Новые правила планирования — только в `core/`-pipeline.

**Files (~3–5):** `core/production/planning.py`, `app/services/production_planning_service.py`, тест.
**Verify:** `pytest tests/test_production_planning_service.py -q` + `tests/test_core_production_planning.py -q` зелёные.
**Зависимости:** `WP5` (PlanRepository persist), `WP6`.

---

## Риски и митигации (на уровне плана)

| Риск | Где | Митигация |
|------|-----|-----------|
| Структурированные ошибки ломают текущий фронт (ждёт `detail`) | WP0 | Fallback на `detail` в `parseError`; backend дублирует человекочитаемый `message` |
| Миграция планов теряет данные | WP4 | dry-run + бэкап + сверка load==source; миграция строго до WP5 |
| Конкурентный конфликт `version` ломает UX | WP3/WP5 | 409 `plan_version_conflict` + reload на клиенте; тест параллельной записи |
| `A1`-вынос задевает много web-кода | WP6/WP7 | Поэтапно (validate→load→optimize→persist); изолированные тесты pipeline до перевода сервиса |
| Случайное использование замороженного бота | — | Бот помечен deprecated; не поддерживаем, но и не удаляем в этом плане |
| Скрытая связь `plan_manager` ↔ другие сервисы | WP5 | Сначала перевести все web-вызовы на репозиторий, затем убрать file I/O |

---

## Контрольные точки верификации (gates)

| Gate | Условие перехода |
|------|------------------|
| G0 (после WP0) | structured error body отдаётся; фронт парсит `code`+fallback; `npm run build`/`test` зелёные |
| G1 (после WP1+WP2) | `Q1`/`Q3` тесты зелёные; нет `except: pass`/fallback без логирования |
| G2 (после WP3+WP4) | репозиторий на SQLite + миграция прошла на копии без потерь |
| G3 (после WP5) | web не пишет планы в файлы; `version` в API; конкурентный конфликт → 409 |
| **Gate Этапа A** | `pytest tests/ -q` зелёный → разрешён старт Этапа B |
| G4 (после WP6) | pipeline изолированно тестируется; `core ↛ app` зелёный |
| **Gate Этапа B** | `production_planning_service` — тонкий адаптер; все тесты зелёные |

---

## Дополнительно явно зафиксировано (по запросу)

1. **Error schema** в `app/schemas/errors.py` + **задача на фронт** для structured errors (сейчас только `detail: string`) → `WP0`.
2. **Поле `version`** в API-ответах плана (`GET /production/plans/{id}`) → `WP5`.
3. **Миграция `bot/data/plans/` → SQLite ДО отключения file I/O** в web → `WP4` строго перед `WP5`.
4. **Assumption #5 уточнён** в спеке: `production_plans(payload_json, version, is_active, …)` — JSON-колонка + `version`, нормализация треков позже (`A7`).

---

## Следующий шаг

~~После ревью этого PLAN → фаза **TASKS**…~~

**Спринт закрыт 2026-06-19.** Следующий фокус: `A3` (runtime-globals), `S1` (bot auth fail-closed), `S2`–`S5`, frontend reload при `plan_version_conflict` (опционально).

**Требует подтверждения перед стартом (Ask first):**
- DDL `production_plans` в `core/kp_db_schema.py` (`WP3`). — **выполнено**
- Изменение формата тела ошибок API (`WP0`) — **выполнено**
