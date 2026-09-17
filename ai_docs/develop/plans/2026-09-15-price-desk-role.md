# Implementation Plan: Стол прайсов

> **Идея:** [`ai_docs/ideas/price-desk-role.md`](../../ideas/price-desk-role.md)
> **ТЗ:** [`ai_docs/specs/price-desk-role.md`](../../specs/price-desk-role.md)
> **Дата:** 2026-09-15
> **Статус:** реализовано (PD-001 … PD-012, 2026-09-15)

## Overview

Экономист (и админ) заходит на `/prices`, бросает заводской Excel, видит дифф,
подтверждает — цены в `pb.db` обновляются. Новые КП сами берут их при расчёте.
Скрипты импорта остаются внутри, CLI больше не нужен для этой работы.

## Architecture Decisions

- Классификация **только по имени файла**, затем один парсер. Не max(rows).
- Превью и apply — два POST, файл оба раза, сверка `sha256`.
- Плиты — новый парсер колонки Никиты; остальные — существующие `import_*`.
- `INSERT OR REPLACE`, пропавшие марки не удаляем.
- **Админ видит стол** (как ГСМ). Экономист не видит КП / 1С / производство / ГСМ.
- `/prices` обязательно в `ROUTE_ACCESS`: сейчас неизвестный путь открыт всем
  (`canAccessRoute` → `true`).

## Dependencies

```
PD-001 ─┐
PD-002 ─┼─ PD-003 ─ PD-004 ─┐
        │                   ├─ PD-007 ─ PD-008 ─ PD-009 ─ PD-010
PD-005 ─┴─ PD-006 ──────────┘              │
                                           └─ PD-011
PD-012 (ручная приёмка, после PD-010)
```

## Task List

### Phase 1: Домен (без UI)

- [x] **PD-001: Классификатор имени файла**
  - **Description:** Якоря по порядку из спеки; `составн` раньше `цельн`;
    нет якоря → ошибка. Регистр и ё не важны.
  - **Acceptance:** тесты на все 7 якорей, отказ без якоря.
  - **Verify:** `pytest tests/test_price_desk_classify.py -q`
  - **Dependencies:** None
  - **Files:** `core/price_desk_classify.py`, `tests/test_price_desk_classify.py`
  - **Scope:** S

- [x] **PD-002: Парсер плит, колонка Никиты**
  - **Description:** Лист «Прайс»; 6/8/10/12.5 → 12; 16/21 skip; М400 игнор;
    пустой Никита skip. Старый `parse_plate_price_rows_from_xlsx` не вызываем.
  - **Acceptance:** синтетический лист покрывает правила спеки.
  - **Verify:** `pytest tests/test_plate_nikita_parser.py -q`
  - **Dependencies:** None
  - **Files:** `core/plate_nikita_parser.py`, `tests/test_plate_nikita_parser.py`
  - **Scope:** M

- [x] **PD-003: Даты в `prices` + запись из Никиты**
  - **Description:** Колонки `price_list_date`, `imported_at` (как у свай);
    `INSERT OR REPLACE`; дата из имени файла.
  - **Acceptance:** повторный импорт идемпотентен; дата пишется.
  - **Verify:** `pytest tests/test_plate_nikita_parser.py tests/test_plate_price_import.py -q`
  - **Dependencies:** PD-002
  - **Files:** `core/price_db.py`, тесты PD-002
  - **Scope:** S

- [x] **PD-004: Дифф текущая таблица ↔ файл**
  - **Description:** Счётчики changed/new/missing/unchanged + до 30 примеров;
    порог 0.005 ₽. Missing после replace не стирается (нет DELETE).
  - **Acceptance:** unit-кейсы на все четыре корзины.
  - **Verify:** `pytest tests/test_price_desk_diff.py -q`
  - **Dependencies:** PD-001, PD-002, PD-003
  - **Files:** `core/price_desk_diff.py`, `tests/test_price_desk_diff.py`
  - **Scope:** M

**Checkpoint 1:** доменные тесты зелёные, UI нет.

### Phase 2: Роль

- [x] **PD-005: Роль `economist` на бэкенде**
  - **Description:** `RegisterUserRequest` += `economist`;
    `REQUIRE_PRICES = require_roles("admin", "economist")`.
  - **Acceptance:** register economist; мусорная роль не проходит Literal.
  - **Verify:** `pytest tests/test_gsm_auth.py tests/test_price_desk_role.py -q`
  - **Dependencies:** None
  - **Files:** `app/schemas/auth.py`, `app/dependencies/auth.py`,
    `tests/test_price_desk_role.py`
  - **Scope:** S

- [x] **PD-006: Роуты и шапка**
  - **Description:** `/prices` в `ROUTE_ACCESS` (admin+economist);
    default route экономиста → `/prices`; пункт «Прайсы»;
    у экономиста нет 1С/КП/ГСМ.
  - **Acceptance:** vitest `roleRoutes` + header.
  - **Verify:** `cd frontend && npm run test -- --run src/shared/lib/roleRoutes.test.ts src/app/layout/AppHeader.test.tsx`
  - **Dependencies:** PD-005
  - **Files:** `frontend/src/shared/lib/roleRoutes.ts`,
    `frontend/src/app/router/AppRouter.tsx`,
    `frontend/src/app/layout/AppHeader.tsx`, тесты
  - **Scope:** M

**Checkpoint 2:** экономист не падает на `/new`; `/prices` может быть заглушкой.

### Phase 3: API

- [x] **PD-007: Сервис preview / apply / status**
  - **Description:** sha256; preview не пишет; apply при чужом sha256 не пишет;
    составные → «группа пока не ведётся»; имя «цельные» + марки ФБС → ошибка,
    не `pile_prices`; вызов существующих `import_*` + nikita.
  - **Acceptance:** unit на сервисе без HTTP.
  - **Verify:** `pytest tests/test_price_desk_service.py -q`
  - **Dependencies:** PD-001…PD-005
  - **Files:** `app/schemas/price_desk.py`, `app/services/price_desk_service.py`,
    `app/dependencies/services.py`, `tests/test_price_desk_service.py`
  - **Scope:** M

- [x] **PD-008: Endpoints**
  - **Description:** `GET /api/v1/prices/status`, `POST .../preview`,
    `POST .../apply`; `read_upload_file_capped`; manager 403; economist 200.
  - **Acceptance:** HTTP-критерии спеки.
  - **Verify:** `pytest tests/test_price_desk_api.py -q`
  - **Dependencies:** PD-007
  - **Files:** `app/api/v1/endpoints/price_desk.py`, `app/api/v1/router.py`,
    `tests/test_price_desk_api.py`
  - **Scope:** M

**Checkpoint 3:** превью без изменения БД, apply меняет, повтор того же файла — дифф пустой.

### Phase 4: UI

- [x] **PD-009: Клиент API**
  - **Description:** types, query hooks (status, preview, apply).
  - **Acceptance:** typecheck без ошибок.
  - **Verify:** `cd frontend && npm run typecheck`
  - **Dependencies:** PD-008
  - **Files:** `frontend/src/features/price-desk/api`, `types`, `hooks`
  - **Scope:** S

- [x] **PD-010: Страница стола**
  - **Description:** Даты по группам + drop-zone `.xls,.xlsx` + дифф +
    «Записать» только после превью, передаёт sha256.
    Паттерн `ImportScheduleDialog`, два шага.
  - **Acceptance:** vitest: дифф до записи; apply с sha256;
    экономист не видит «Выгрузка 1С».
  - **Verify:** `cd frontend && npm run test -- --run src/features/price-desk` + `npm run typecheck`
  - **Dependencies:** PD-006, PD-009
  - **Files:** `frontend/src/pages/prices/PricesPage.tsx`,
    `frontend/src/features/price-desk/`, `AppRouter.tsx` (подключение страницы)
  - **Scope:** M

**Checkpoint 4:** ручной прогон в браузере под economist и admin.

### Phase 5: Изоляция и приёмка

- [x] **PD-011: 403 экономиста на чужие API**
  - **Description:** КП, архив, ГСМ, `POST /nomenclature/import-1c`.
  - **Acceptance:** все перечисленные → 403 для economist.
  - **Verify:** `pytest tests/test_price_desk_api.py -q`
  - **Dependencies:** PD-008
  - **Files:** `tests/test_price_desk_api.py`
  - **Scope:** S

- [x] **PD-012: Ручная приёмка на живых файлах** (не CI)
  - **Description:** Счётчики 07.09 и Никита 277; составные отказ;
    ФБС не в сваях.
  - **Acceptance:** совпадает со success criteria спеки.
  - **Verify:** ручной прогон
  - **Dependencies:** PD-010
  - **Files:** нет
  - **Scope:** S

## Risks

| Риск | Что делать |
|------|------------|
| Имя файла без якоря | 422, не угадывать |
| Ложный парсер | только имя + проверка марок |
| `prices` без даты | PD-003 до экрана статуса |
| `/prices` открыт всем | PD-006 в `ROUTE_ACCESS` |

## Open Questions

Нет. Админ на столе — согласовано 15.09.2026.

## Not doing

Составные как продукт, GUID/1С в этом столе, живая таблица, вычитание марок,
округление до рубля, журнал файлов.
