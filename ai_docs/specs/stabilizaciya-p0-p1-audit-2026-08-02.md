# Spec: Стабилизация P0 + узкий P1 — аудит 2026-08-02

> **Тип:** remediation feature-spec (стабилизационный срез)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → IMPLEMENT ✅  
> **Дата:** 2026-08-02  
> **Статус:** IMPLEMENTED  
> **План:** [`../develop/plans/2026-08-02-stabilizaciya-p0-p1-audit.md`](../develop/plans/2026-08-02-stabilizaciya-p0-p1-audit.md)  
> **Scope выбран пользователем:** **B** — P0 + узкий P1  
> **Источник:** [`../develop/audits/2026-08-02-audit-comparison.md`](../develop/audits/2026-08-02-audit-comparison.md)  
> **Полный аудит:** [`../develop/audits/2026-08-02-full-project-audit.md`](../develop/audits/2026-08-02-full-project-audit.md)  
> **Baseline:** [`project-baseline.md`](./project-baseline.md)  
> **Связанные закрытые спеки:** [`core-domain-enums-a1.md`](./core-domain-enums-a1.md), [`move-to-production-atomicity-q1-q2.md`](./move-to-production-atomicity-q1-q2.md)

---

## Стратегия (одной фразой)

> Закрыть 3 P0-блокера (утечка финансов + 2 runtime NameError) и узкий P1 (утечка SQLite в API, freeze в read-path, тесты на PDF/Gantt) — без рефакторинга god-сервисов и без изменения схемы БД.

---

## ASSUMPTIONS I'M MAKING

1. Scope = **P0 (3 уникальных блокера) + узкий P1** из сравнения аудитов:
   - P0: financial leak logistics, missing `datetime`, missing `create_gantt_excel`
   - P1: SQLite errors → client, freeze в read-path (`archive` + `sgp` progress), тесты PDF/XLSX + Gantt
2. **A1 + S1** считаются **одной** проблемой (field-level ACL / slim DTO).
3. Financial leak закрываем **role-aware mapping на существующем** `GET /commercial/archive/search`, без нового logistics-роутера.
4. Для logistics в ответе поиска **нет** `discount_percent`, `subtotal`, `vat_amount`, `total_amount` (и прочих финансовых полей). UI (`CreateShipmentDialog`) использует только `kp_id` + `customer_name`.
5. Admin/manager на `/search` сохраняют полный `ArchiveOfferListItem` (как сейчас).
6. Freeze в read-path: **только чтение** `ordered_qty`; запись freeze остаётся в write-путях (`commit_move_to_production`, `ShipmentService.complete`, и т.п.).
7. Fallback `m=x` при `ordered_qty IS NULL` в `_shipped_progress` — **в scope** (связан с Q5): при отсутствии M возвращаем `shipped_progress: null` или `m: null`, не маскируем.
8. God-services, repository layer, rate limit Redis, CSP, OCR, DI refactor — **out of scope**.
9. Audit-файлы **не** помечаем FIXED в этом изменении (отдельный follow-up).
10. Эта спека — фаза SPECIFY. PLAN/TASKS — после вашего «ок».

→ Поправьте допущения или подтвердите — перейду к PLAN.

---

## 1. Objective

### Что строим и зачем

Стабилизационный срез после сравнения двух прогонов `/audit` (2026-08-02): устранить блокеры production и ближайшие security/data-integrity риски без большого архитектурного рефакторинга.

**Пользователи:**
- `logistics` — поиск КП для создания рейса (не должен видеть суммы/скидки)
- `manager` / `admin` — PDF/XLSX КП, Gantt актуального плана, архив без регрессий
- все роли logistics API — не должны получать текст SQLite-ошибок

### Problem statement

| ID (v2 / comparison) | Severity | Суть |
|----------------------|----------|------|
| A1/S1 | Critical | `/archive/search` для logistics отдаёт полный commercial aggregate с финансами |
| Q1 v2 | Critical | `offers_service.py` — `datetime.now()` без import → NameError на PDF/XLSX |
| Q2 v2 | Critical | `archive_service.py` — `create_gantt_excel` без import → 500 на Gantt |
| S3 / A22 | High | `ShipmentError(f"…: {exc}")` утекает в HTTP `message` |
| Q4 / A9 (+ Q5) | High/Med | `freeze_ordered_qty_if_needed` в read-path; `m=x` маскирует NULL |

### Reframe → success criteria

| Требование | Конкретный критерий |
|------------|---------------------|
| «Логист не видит деньги» | JSON ответа `/search` для `role=logistics` **не содержит** ключей `discount_percent`, `subtotal`, `vat_amount`, `total_amount` (или они отсутствуют в модели) |
| «PDF/XLSX работают» | `OffersService.generate_pdf/xlsx` с пустым `creation_date` не бросает `NameError`; тест зелёный |
| «Gantt работает» | `ArchiveService.build_current_plan_gantt` резолвит `create_gantt_excel`; тест/smoke зелёный |
| «Нет утечки SQLite» | При `sqlite3.Error` в pile catalog клиент видит стабильное сообщение без текста exception |
| «Read не пишет M» | `_shipped_progress` и `_sgp_progress_on_cursor` не вызывают UPDATE `ordered_qty` |

---

## 2. Tech Stack

Без изменений относительно baseline:

- Backend: Python 3, FastAPI, Pydantic v2, SQLite
- Frontend: React + Vite + TypeScript (минимальные правки types для slim search item)
- Тесты: `pytest tests/`

---

## 3. Commands

```bash
# Backend tests (целевой набор)
.venv/bin/pytest \
  tests/test_logistics_api.py \
  tests/test_archive_endpoints.py \
  tests/test_archive_authorization.py \
  tests/test_offers_service.py \
  tests/test_shipment_service.py \
  tests/test_sgp_service.py \
  -q

# Или после создания новых файлов:
.venv/bin/pytest tests/test_offers_service.py tests/test_archive_gantt.py -q

# Frontend (если трогаем types / CreateShipmentDialog)
cd frontend && npm test -- --run src/features/logistics

# Dev
uvicorn app.main:app --reload
cd frontend && npm run dev
```

---

## 4. Project Structure (затрагиваемые пути)

```
app/schemas/archive.py              → slim LogisticsKpSearchItem (+ response variant)
app/services/archive_service.py     → role-aware search mapping; gantt import; read-only progress
app/api/v1/endpoints/archive.py     → response_model / typing для search
app/services/offers_service.py      → import datetime
app/services/shipment_service.py    → generic ShipmentError message (pile catalog)
app/api/v1/endpoints/logistics.py   → при необходимости не пробрасывать str(exc) с DB detail
app/services/sgp_service.py         → read-only ordered_qty в _sgp_progress_on_cursor
core/kp_db_plates_completion.py     → опционально: read helper get_ordered_qty (если удобно)
core/gantt_excel.py                 → источник create_gantt_excel (без изменения логики)

frontend/src/features/commercial-archive/types/archive.ts  → тип slim item (если нужен)
frontend/src/features/logistics/components/CreateShipmentDialog.tsx → тип результатов поиска

tests/test_logistics_api.py         → assert: нет financial keys для logistics
tests/test_offers_service.py        → NEW (или расширить): PDF/XLSX без creation_date
tests/…                             → Gantt import / build_current_plan_gantt smoke
tests/test_sgp_service.py           → read-path не мутирует ordered_qty
```

Docs:

```
ai_docs/specs/stabilizaciya-p0-p1-audit-2026-08-02.md   → эта спека
ai_docs/develop/plans/…                                 → PLAN после approval
```

---

## 5. Code Style

Пример role-aware mapping (ориентир, не финальный код):

```python
# app/services/archive_service.py — search()
FINANCIAL_FIELDS = ("discount_percent", "subtotal", "vat_amount", "total_amount")

def search(self, *, user: dict, kp_id: int | None = None, customer: str | None = None):
    logistics_read = user.get("role") == "logistics"
    # ... load rows as today ...
    if logistics_read:
        items = [self._to_logistics_search_item(raw) for raw in rows]
    else:
        items = [self._to_list_item(raw) for raw in rows]
    return ArchiveSearchResponse(...)  # или LogisticsArchiveSearchResponse

def _to_logistics_search_item(self, raw: dict) -> LogisticsKpSearchItem:
    return LogisticsKpSearchItem(
        kp_id=int(raw.get("kp_id") or 0),
        customer_name=raw.get("customer_name"),
        status=raw.get("status") or None,
        product_type=str(raw.get("product_type") or "plates"),
        # без финансов, без sgp/shipped progress (не нужны диалогу)
    )
```

Read-only progress:

```python
# вместо freeze_ordered_qty_if_needed в read-path:
cur.execute("SELECT ordered_qty FROM kp_meta WHERE kp_id = ?", (kp_id,))
row = cur.fetchone()
ordered = int(row[0]) if row and row[0] is not None else None
if ordered is None:
    return None  # shipped_progress отсутствует — не маскировать m=x
```

Generic error:

```python
except sqlite3.Error as exc:
    logger.exception("Ошибка чтения каталога свай: %s", exc)
    raise ShipmentError(
        "Ошибка чтения каталога свай",
        code="pile_catalog_read_failed",
    ) from exc
```

Соглашения:
- минимальный diff; не рефакторить god-сервисы «заодно»;
- typed domain errors как уже принято (`ShipmentError` с `code`);
- financial strip — на сервере (не только UI).

---

## 6. Testing Strategy

| Уровень | Что покрыть | Где |
|---------|-------------|-----|
| API authz | logistics `/search` 200 и **без** financial keys; manager/admin полный item; list/details logistics 403 | `tests/test_logistics_api.py` |
| Unit | `generate_pdf` / `generate_xlsx` при `creation_date=None` → нет NameError | `tests/test_offers_service.py` (new) |
| Unit/integration | `build_current_plan_gantt` с mock `PlanDistributionService` + реальный import path **или** тест, что `create_gantt_excel` импортируется из `core.gantt_excel` в методе | новый/существующий archive test |
| Unit | pile catalog: при `sqlite3.Error` message не содержит sqlite detail | `tests/test_shipment_service.py` |
| Unit | `_shipped_progress` / `sgp_progress`: при `ordered_qty IS NULL` **нет** UPDATE; badge не врёт `x/x` | archive + sgp tests |

Coverage ожидание: каждый пункт Success Criteria имеет ≥1 автоматизированный тест.

---

## 7. Boundaries

### Always
- Серверный strip финансовых полей для `role=logistics` на `/archive/search`
- Логировать exception server-side; клиенту — стабильный текст
- Freeze M только в write-транзакциях
- Прогнать целевой pytest-набор перед сдачей среза
- Минимальный diff

### Ask first
- Новый endpoint `/logistics/kp-search` вместо role-aware mapping на `/archive/search`
- Изменение OpenAPI так, что admin/manager теряют поля в search
- Рефакторинг `ShipmentService` / repository layer
- Пометка FIXED в audit markdown
- Миграция схемы БД / смена fallback progress globally beyond archive+sgp read helpers

### Never
- Коммит секретов
- God-service split «заодно»
- Redis rate limiter / CSP enforce / OCR policy в этом срезе
- Удаление failing tests без approval
- Возврат финансовых полей logistics «для удобства UI»

### Out of scope (явно)

| Тема | Почему |
|------|--------|
| A2–A7 god-services / DI / Pydantic-in-service | большой рефакторинг |
| S2 rate limit Redis | деплой/инфра |
| S4 OCR external LLM | политика, не hotfix |
| Q3 duplicate order_data | отдельный quality ticket |
| Q6 unify move_to_production wrappers | уже есть atomic core; DRY — позже |
| Q9 free-only complete | отдельный business rule |
| Frontend god ShipmentItemsSection | UI debt |
| Bot / viz cleanup | scope expansion |

---

## 8. Scope breakdown (implementable slices)

### P0-1 — Runtime imports (самый маленький)

| | |
|--|--|
| **Files** | `app/services/offers_service.py`, `app/services/archive_service.py`, tests |
| **Do** | `from datetime import datetime`; `from core.gantt_excel import create_gantt_excel` |
| **Verify** | pytest offers + gantt/import smoke |

### P0-2 — Financial fields strip for logistics

| | |
|--|--|
| **Files** | `app/schemas/archive.py`, `archive_service.py`, `endpoints/archive.py`, frontend types/dialog (минимально), `tests/test_logistics_api.py` |
| **Do** | `LogisticsKpSearchItem` + mapping в `search()`; OpenAPI/response не обещает финансы logistics |
| **Verify** | logistics search без financial keys; manager path без регрессии; CreateShipmentDialog работает |

### P1-1 — SQLite error sanitization

| | |
|--|--|
| **Files** | `shipment_service.py` (+ audit других `ShipmentError(f"…{exc}")` в том же файле, если есть) |
| **Do** | generic message; detail только в log |
| **Verify** | unit/API test на `pile_catalog_read_failed` |

### P1-2 — Freeze out of read-path (+ Q5)

| | |
|--|--|
| **Files** | `archive_service.py` (`_shipped_progress`), `sgp_service.py` (`_sgp_progress_on_cursor`), optional read helper in `core/kp_db_plates_completion.py`, tests |
| **Do** | SELECT only; no UPDATE on GET/list/progress; NULL ordered → null progress, не `m=x` |
| **Verify** | тест: после read `ordered_qty` остаётся NULL; write-path freeze по-прежнему работает |

### P1-3 — Tests for PDF/XLSX + Gantt

| | |
|--|--|
| **Files** | new/extended pytest modules |
| **Do** | покрыть P0-1 acceptance (даже если imports уже починены) |
| **Verify** | CI-локальный pytest зелёный |

**Порядок:** P0-1 → P0-2 → P1-1 → P1-2 → P1-3 (P1-3 можно параллельно с P0-1).

---

## 9. Success Criteria (done = все true)

1. `role=logistics` на `GET /api/v1/commercial/archive/search` получает 200; в `items[]` **нет** финансовых ключей/полей.
2. `role=manager|admin` на том же endpoint сохраняют финансовые поля (где доступ разрешён owner/role-фильтром).
3. `CreateShipmentDialog` по-прежнему находит и выбирает КП (`kp_id`, `customer_name`).
4. `OffersService.generate_pdf` и `generate_xlsx` при отсутствии `creation_date` не падают с `NameError`.
5. `ArchiveService.build_current_plan_gantt` не падает с `NameError` на `create_gantt_excel`.
6. Ошибка чтения pile catalog не содержит sqlite exception text в HTTP `message`.
7. Read-path progress (`_shipped_progress`, `_sgp_progress_on_cursor`) не выполняет UPDATE `kp_meta.ordered_qty`.
8. При `ordered_qty IS NULL` shipped badge не показывает ложные 100% через `m=x`.
9. Целевой pytest-набор зелёный.
10. Diff не включает god-service split / Redis / CSP / OCR.

---

## 10. Open Questions

| # | Вопрос | Предложение по умолчанию |
|---|--------|--------------------------|
| Q1 | Отдельный endpoint `/logistics/kp-search` или role-aware `/archive/search`? | **Role-aware на существующем** (assumption 3) |
| Q2 | Какие поля оставить logistics? | `kp_id`, `customer_name`, `status`, `product_type` |
| Q3 | Нужны ли logistics `sgp_progress` / `shipped_progress` в search? | **Нет** — диалогу не нужны |
| Q4 | При NULL `ordered_qty` в SGP progress: null весь badge или compute ephemeral M без записи? | **Ephemeral compute без UPDATE** для SGP view (чтобы UI не пустел), но **без persist**; для `_shipped_progress` — `null` если M нет |
| Q5 | Санитизировать все `ShipmentError` с `{exc}` в файле или только pile catalog? | **Все вхождения `{exc}` в shipment_service**, если найдутся grep'ом |
| Q6 | Обновлять audit markdown FIXED? | **Нет** в этом срезе |

→ Ответьте на Q1–Q6 (или «defaults ok») и подтвердите спеку — затем Phase 2: PLAN.

---

## 11. Risks

| Риск | Mitigation |
|------|------------|
| OpenAPI `response_model=ArchiveSearchResponse` всё ещё документирует финансы | Отдельная slim response model или Union; тест на фактический JSON |
| Frontend type `ArchiveOfferListItem` ожидает обязательные number-поля | Slim type / optional fields для search results в logistics |
| SGP UI «пустеет» без ephemeral M | Q4 default: ephemeral read-only compute без UPDATE |
| Gantt test хрупкий (нужны планы) | Mock `get_all_plans_gantt_data` + assert callable import |

---

*SPECIFY complete. Waiting for human review before PLAN.*
