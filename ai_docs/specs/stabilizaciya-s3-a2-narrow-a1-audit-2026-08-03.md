# Spec: Стабилизация S3 + A2 + узкий A1 — аудит 2026-08-03

> **Тип:** remediation feature-spec (стабилизационный срез)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS ✅ → IMPLEMENT ✅  
> **Дата:** 2026-08-03  
> **Статус:** IMPLEMENTED  
> **Scope выбран пользователем:** **D** — S3 + A2 + узкий A1  
> **Источник:** [`../develop/audits/2026-08-03-full-project-audit.ru.md`](../develop/audits/2026-08-03-full-project-audit.ru.md)  
> **План:** [`../develop/plans/2026-08-03-stabilizaciya-s3-a2-narrow-a1-audit.md`](../develop/plans/2026-08-03-stabilizaciya-s3-a2-narrow-a1-audit.md)  
> **Baseline:** [`project-baseline.md`](./project-baseline.md)  
> **ADR:** [`../develop/architecture/deployment-single-instance.md`](../develop/architecture/deployment-single-instance.md)  
> **Связанные:** [`shipment-logistics.md`](./shipment-logistics.md), [`stabilizaciya-p0-p1-audit-2026-08-02.md`](./stabilizaciya-p0-p1-audit-2026-08-02.md)

---

## Стратегия (одной фразой)

> Закрыть списание чужого СГП через валидацию КП состава (S3), сделать single-worker hard-fail в production (A2) и вынести SQL + completion из god-`ShipmentService` без полного split propose/export (узкий A1).

---

## Decisions locked (подтверждено 2026-08-03)

| # | Решение |
|---|---------|
| D1 | Scope **D**: S3 + A2 + узкий A1 |
| D2 | A2 hard-fail **только** при `APP_ENV=production` + `single_instance` + `configured_workers > 1` |
| D3 | `configured_workers is None` → **не** hard-fail; усилить WARNING + `/health` (`undeclared` / note) |
| D4 | S3: проверка в `put_items` **и** `complete` (3a); patch-orders UX-гард (3b) — out of scope |
| D5 | Pile `kp_id=null` **разрешён**; заданный `kp_id` должен ∈ orders |
| D6 | `cancel` живёт в `ShipmentCompletionService` |
| D7 | Endpoints → только фасад `ShipmentService` (делегирование внутрь) |
| D8 | Audit FIXED / Redis / полный A1 Propose+Export / A4 — out of scope |

---

## 1. Objective

### Что строим и зачем

После аудита 2026-08-03 закрыть критичный риск целостности склада (S3), enforcement деплой-модели single-instance (A2) и снизить связность logistics-сервиса ровно настолько, чтобы completion + persistence тестировались отдельно (узкий A1).

**Пользователи / акторы:**
- `logistics` — формирует состав рейса и завершает выезд; не должен списывать плиты чужого КП
- ops / deploy — не должен случайно поднять multi-worker production на SQLite + in-process limits
- разработчики — completion и SQL можно трогать без чтения всего god-модуля

### Problem statement

| ID | Severity | Суть |
|----|----------|------|
| S3 | High (security/integrity) | `put_items`/`complete` не проверяют, что плита принадлежит КП заказов рейса |
| A2 | Critical (architecture) | single-instance только warning; multi-worker в prod ломает данные/rate limits |
| A1 | Critical (architecture) | `ShipmentService` ~1521 строк смешивает SQL, completion, propose, export |

### Reframe → success criteria

| Требование | Конкретный критерий |
|------------|---------------------|
| «Не списать чужой СГП» | `put_items` с `completed_plate_id` КП вне `shipment_orders` → `ShipmentError` с code `shipment_plate_kp_mismatch`; `complete` с уже сохранённым чужим item — тот же отказ до списания |
| «Pile с чужим kp_id» | pile с `kp_id` ∉ orders → `shipment_pile_kp_mismatch`; pile без `kp_id` проходит |
| «Prod не стартует на N workers» | при `APP_ENV=production` + `single_instance` + `UVICORN_WORKERS=2` lifespan/startup бросает исключение |
| «Dev не ломается» | при `APP_ENV=development` + workers>1 — только WARNING, приложение стартует |
| «Undeclared workers видны» | при `configured_workers is None` и in-process store — WARNING + поле в `/health` rate_limiting metadata |
| «Completion отдельно» | `put_items` / `complete` / `cancel` в `ShipmentCompletionService`; фасад делегирует |
| «SQL не в god-методах» | fetch/insert shipment/orders/items — через `ShipmentRepository` и/или `core/kp_db_shipments.py` |
| «Propose/export на месте» | `propose*` и `export_shipment_sheet_xlsx` остаются в `ShipmentService` |
| «Нет регрессии happy-path» | `tests/test_shipment_service.py` + logistics API зелёные |

---

## 2. Tech Stack

Без изменений относительно baseline:

- Backend: Python 3, FastAPI, Pydantic v2, SQLite
- Persistence: `core/kp_db_shipments.py`, новый `app/repositories/shipment_repository.py`
- Тесты: `pytest tests/`
- Frontend: **не трогаем**

---

## 3. Commands

```bash
.venv/bin/pytest \
  tests/test_shipment_service.py \
  tests/test_shipment_qty_balance.py \
  tests/test_shipment_kp_membership.py \
  tests/test_logistics_api.py \
  tests/test_rate_limit_deployment.py \
  -q

uvicorn app.main:app --reload --workers 1
```

---

## 4. Project Structure (затрагиваемые пути)

```
app/services/shipment_service.py              → фасад
app/services/shipment_completion_service.py    → NEW: put_items, complete, cancel + S3
app/repositories/shipment_repository.py       → NEW: CRUD SQL
core/kp_db_shipments.py                       → helper order kp_ids / reuse available_qty
app/security/login_rate_limit.py              → enforce_* + undeclared warning/health
app/main.py                                   → lifespan: enforce
app/dependencies/services.py                  → wiring repo/completion в фасад

ai_docs/develop/architecture/deployment-single-instance.md
ai_docs/develop/deploy-contract.md

tests/test_shipment_kp_membership.py          → NEW
tests/test_rate_limit_deployment.py           → hard-fail + undeclared cases
tests/test_shipment_service.py                → regression (адаптация фикстур при необходимости)
```

### Целевая граница ответственности

```
ShipmentService (фасад)
  create / reuse / get / list / patch / propose* / export / search_*
  → делегирует put_items / complete / cancel
        │
        ▼
ShipmentCompletionService
  put_items, complete, cancel + KP membership
        │
        ▼
ShipmentRepository + kp_db_shipments
```

---

## 5. Code Style

```python
def _assert_kp_in_shipment_orders(
    cur, shipment_id: int, kp_id: int, *, code: str, detail: str
) -> None:
    cur.execute(
        "SELECT 1 FROM shipment_orders WHERE shipment_id = ? AND kp_id = ?",
        (int(shipment_id), int(kp_id)),
    )
    if cur.fetchone() is None:
        raise ShipmentError(detail, code=code)
```

```python
def enforce_single_instance_workers(*, app_env: str, storage_layout: str) -> None:
    if app_env.lower() != "production" or storage_layout != "single_instance":
        return
    workers = configured_worker_count()
    if workers is not None and workers > 1:
        raise RuntimeError(
            f"Refusing to start: configured_workers={workers} with "
            "APP_ENV=production and APP_STORAGE_LAYOUT=single_instance. "
            "Set UVICORN_WORKERS=1."
        )
```

Соглашения: минимальный diff; typed `ShipmentError`; фасад сохраняет сигнатуры endpoints; repository без FastAPI.

---

## 6. Testing Strategy

| Уровень | Что | Где |
|---------|-----|-----|
| S3 put/complete/pile | mismatch codes + null-ok pile | `tests/test_shipment_kp_membership.py` |
| A2 enforce | prod raise / dev no-raise / undeclared health warning | `tests/test_rate_limit_deployment.py` |
| Regression | CRUD/propose/export/cancel/complete happy-path | `test_shipment_service.py`, `test_logistics_api.py` |

---

## 7. Boundaries

### Always
- KP membership в `put_items` и `complete`
- Hard-fail multi-worker только production + `single_instance`
- Endpoints через фасад
- Целевой pytest перед сдачей

### Ask first
- Полный Propose/Export split; A4; hard-fail non-prod; hard-fail при `workers is None`; schema; audit FIXED

### Never
- Секреты; Redis в этом срезе; рефакторинг CommercialWorkflow/Sgp «заодно»; удаление failing tests

### Out of scope
Полный A1, A3–A8, S1 Redis, S2 OCR, Q1–Q2 SGP DRY, frontend redesign, patch-orders UX-гард (3b), multi-replica shared_volume

---

## 8. Scope breakdown

1. **S3** KP membership + tests  
2. **A2** production hard-fail + undeclared visibility + docs  
3. **Узкий A1** repository + completion extract + facade delegate  

---

## 9. Success Criteria (чеклист done)

- [x] `put_items` → `shipment_plate_kp_mismatch` для чужого КП
- [x] `complete` отклоняет чужой состав до списания
- [x] pile чужой `kp_id` → `shipment_pile_kp_mismatch`; null-ok
- [x] production + workers>1 → отказ старта
- [x] development + workers>1 → warning, старт ок
- [x] undeclared workers → warning + health metadata
- [x] CompletionService: put_items/complete/cancel
- [x] ShipmentRepository владеет CRUD SQL рейса
- [x] propose/export в фасаде
- [x] целевой pytest зелёный
- [x] ADR/deploy-contract обновлены

---

## 10. Open Questions

Все закрыты (см. Decisions locked).

---

## Verification (фаза SPECIFY)

- [x] Spec покрывает Objective / Commands / Structure / Style / Testing / Boundaries
- [x] Human reviewed and approved
- [x] Success criteria конкретны и тестируемы
- [x] Boundaries заданы
- [x] Spec сохранён в репозитории
