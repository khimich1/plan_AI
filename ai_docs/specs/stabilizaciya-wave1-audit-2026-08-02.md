# Spec: Стабилизация волна 1 — аудит 2026-08-02

> **Тип:** remediation feature-spec (стабилизационный срез)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → IMPLEMENT ✅  
> **Дата:** 2026-08-02  
> **Статус:** IMPLEMENTED  
> **План:** [`../develop/plans/2026-08-02-stabilizaciya-wave1-audit.md`](../develop/plans/2026-08-02-stabilizaciya-wave1-audit.md)  
> **Scope:** **1A** — S4 + A2/S2 + S9 + Q2 + A8/S12  
> **ACL logistics (S12):** **B** — статусы `в работе`, `На СГП`  
> **Источник:** [`../develop/audits/2026-08-02-full-project-audit.md`](../develop/audits/2026-08-02-full-project-audit.md)  
> **Предшественник:** [`stabilizaciya-p0-p1-audit-2026-08-02.md`](./stabilizaciya-p0-p1-audit-2026-08-02.md) (IMPLEMENTED)  
> **Baseline:** [`project-baseline.md`](./project-baseline.md)  
> **Связанный ADR-кандидат:** [`../develop/architecture/rate-limiting.md`](../develop/architecture/rate-limiting.md)

---

## Стратегия (одной фразой)

> Закрыть быстрые security/deploy риски и дожать bounded context логистики (свой kp-search + ACL по статусам), плюс страховка тестами UI — без декомпозиции ShipmentService и без Redis.

---

## Почему ACL = B (рекомендация, зафиксировано)

| Вариант | Плюсы | Минусы |
|---------|-------|--------|
| **A** все КП | Простой UX | Logistics перебирает коммерческий архив (скидки уже скрыты, но номера/заказчики всех менеджеров — да) |
| **B** `в работе` + `На СГП` | Соответствует работе логиста; режет «в архиве»/черновики | Нельзя заранее набрать рейс до перевода в производство |
| **C** только с остатком на СГП | Максимальный least privilege | Сложнее SQL/UX; нельзя создать рейс «под ожидаемую» отгрузку |

**Выбор B:** логист работает с КП, уже запущенными в производство или лежащими на СГП. Это закрывает S12-enumeration архивных КП без тяжёлого join по `completed_plates`. Вариант C — возможный follow-up.

Статусы в scope поиска: `KpStatus.IN_WORK` (`в работе`), `KpStatus.ON_SGP` (`На СГП`).  
Статусы **вне** поиска: `в архиве`, `выполнено` (если понадобится — отдельное решение).

---

## ASSUMPTIONS I'M MAKING

1. Scope = полный список волны 1: **S4, A2+S2, S9, Q2, A8+S12**.
2. **Redis rate limit — out of scope**; A2/S2 закрываются ADR + startup/deploy guard под single-instance / один worker.
3. `frontend/package.json` уже имеет script `audit:ci` — нужно **вшить в CI/деплой-контракт** и обновить `react-router-dom` до patched.
4. Logistics **перестаёт** вызывать `GET /commercial/archive/search`; роль `logistics` **убирается** с archive search (только admin/manager).
5. Новый endpoint: `GET /api/v1/logistics/kp-search` — slim DTO без финансов + фильтр статусов B.
6. Q2 = тесты `draftItems.ts` + узкий component/smoke для критичного flow; **не** полный split `ShipmentItemsSection` (A11).
7. A1 god-ShipmentService, OCR (S3), CSP (S5), encryption (S7) — **out of scope**.
8. Audit markdown FIXED-маркеры — не трогаем в этом срезе.
9. Эта спека — фаза SPECIFY. PLAN — после «ок».

→ Поправьте или подтвердите (особенно: исключать ли `выполнено` из ACL B).

---

## 1. Objective

### Что и зачем

После P0/P1 (импорты, финансы, freeze) закрыть следующий слой аудита: dependency CVE, честная deployment-модель, DEBUG-guard, изоляция logistics search, тестовая страховка UI отгрузки.

**Пользователи:**
- Ops / deploy — понятные ограничения single-instance
- `logistics` — поиск КП только среди релевантных статусов, свой API
- Разработчики — `npm audit` в CI, тесты draft/items

### Reframe → success criteria

| Требование | Критерий |
|------------|----------|
| «Починить npm CVE» | `react-router-dom` ≥ patched; `npm run audit:ci` exit 0 |
| «Честный single-instance» | ADR + guard/warn при несовместимой prod-конфигурации; rate-limit doc ссылается на workers=1 |
| «DEBUG не в prod» | Старт с `APP_ENV=production` + `APP_DEBUG=true` → fail |
| «Свой kp-search» | `CreateShipmentDialog` → logistics API; archive/search без роли logistics |
| «ACL B» | KP со статусом `в архиве` / `выполнено` не возвращаются logistics search |
| «Тесты UI» | `draftItems` unit tests зелёные; есть smoke на ключевой logistics UI path |

---

## 2. Tech Stack

Без смены стека:

- Backend: FastAPI, Pydantic v2, SQLite, settings validators
- Frontend: React + Vite + TypeScript + Vitest
- Docs: `ai_docs/develop/architecture/` для ADR

---

## 3. Commands

```bash
# Frontend security + tests
cd frontend
npm install                    # после bump react-router-dom
npm run audit:ci
npm test -- --run src/features/logistics/lib/draftItems.test.ts
npm test -- --run src/features/logistics/components/CreateShipmentDialog.test.tsx
npm run typecheck

# Backend
.venv/bin/pytest \
  tests/test_logistics_api.py \
  tests/test_archive_authorization.py \
  tests/test_settings_guards.py \
  -q

# Dev
uvicorn app.main:app --reload --workers 1
cd frontend && npm run dev
```

---

## 4. Project Structure (затрагиваемое)

```
# S4
frontend/package.json
frontend/package-lock.json

# A2 / S2
ai_docs/develop/architecture/deployment-single-instance.md   → NEW ADR
ai_docs/develop/architecture/rate-limiting.md               → link to ADR / workers=1
core/config/settings.py                                     → optional guard / tighten warning
docker-compose.yml / deploy docs                            → workers=1 explicit if applicable

# S9
core/config/settings.py                                     → model_validator: production ∧ debug → error
tests/test_settings_guards.py                               → NEW or extend

# A8 / S12
app/schemas/logistics.py                                    → LogisticsKpSearchItem + response
app/api/v1/endpoints/logistics.py                           → GET /kp-search
app/services/…                                              → search helper (shipment or thin service)
app/api/v1/endpoints/archive.py                             → убрать logistics из require_roles на /search
frontend/.../logisticsApi.ts                                → searchKp()
frontend/.../CreateShipmentDialog.tsx                       → logisticsApi вместо archiveApi
tests/test_logistics_api.py                                 → ACL + slim fields

# Q2
frontend/.../lib/draftItems.ts
frontend/.../lib/draftItems.test.ts                         → NEW
frontend/.../components/ShipmentItemsSection.test.tsx       → NEW minimal smoke (optional if time)
```

---

## 5. Code Style

### DEBUG guard

```python
@model_validator(mode="after")
def reject_debug_in_production(self) -> Settings:
    if self.app_env.lower() == "production" and self.app_debug:
        raise ValueError(
            "APP_DEBUG must be false when APP_ENV=production"
        )
    return self
```

### Logistics kp-search (slim + ACL B)

```python
ALLOWED_LOGISTICS_KP_STATUSES = frozenset({
    KpStatus.IN_WORK.value,   # "в работе"
    KpStatus.ON_SGP.value,    # "На СГП"
})

class LogisticsKpSearchItem(BaseModel):
    kp_id: int
    customer_name: str | None = None
    status: str | None = None
    product_type: Literal["plates", "piles"] = "plates"
```

### Frontend switch

```typescript
// CreateShipmentDialog — вместо archiveApi.search
const response = await logisticsApi.searchKp(
  Number.isInteger(numeric) && numeric > 0
    ? { kpId: numeric }
    : { customer: raw },
);
```

Соглашения: минимальный diff; переиспользовать существующий slim mapping из archive_service где возможно (extract helper), не копипастить SQL.

---

## 6. Testing Strategy

| Срез | Что | Где |
|------|-----|-----|
| S4 | `npm run audit:ci` зелёный после bump | frontend CI / локально |
| S9 | Settings raise при production+debug | `tests/test_settings_guards.py` |
| A8/S12 | logistics kp-search 200 slim; архивный КП не находится; archive/search → 403 для logistics | `tests/test_logistics_api.py` |
| A8 | CreateShipmentDialog мокает `logisticsApi.searchKp` | frontend test |
| Q2 | draftItems: payload, manual weight, version | `draftItems.test.ts` |

---

## 7. Boundaries

### Always
- Slim DTO без финансов на logistics kp-search
- Фильтр статусов B на сервере (не только UI)
- Убрать `logistics` с `GET /commercial/archive/search`
- `workers=1` / single-instance задокументированы
- Целевые pytest + vitest + audit:ci зелёные

### Ask first
- Включить статус `выполнено` в ACL B
- Перейти на ACL C (остаток на СГП)
- Добавить Redis rate limiter
- Менять CI platform (если нет GitHub Actions — куда вешать `audit:ci`)
- Полный split ShipmentItemsSection (A11)

### Never
- Возвращать финансы на logistics search
- Оставлять logistics на archive/search «на всякий случай»
- Декомпозиция god-ShipmentService в этом срезе
- Коммит секретов; пометка FIXED в audit без отдельной просьбы

### Out of scope

| ID | Тема |
|----|------|
| A1 | Full ShipmentService split |
| A3–A7 | CommercialWorkflow / DI / repository wave |
| S3 | OCR external LLM policy |
| S5 | CSP enforcing |
| S7 | Encryption at rest |
| A11 | Full UI component split (только тесты Q2) |

---

## 8. Scope slices (для будущего PLAN)

1. **S4** — bump `react-router-dom` + verify `audit:ci`  
2. **S9** — DEBUG production guard + test  
3. **A2/S2** — ADR deployment-single-instance + sync rate-limiting.md + compose/docs workers=1  
4. **A8/S12** — backend kp-search + remove logistics from archive search + tests  
5. **A8 frontend** — logisticsApi + CreateShipmentDialog  
6. **Q2** — draftItems tests (+ optional ShipmentItemsSection smoke)

Порядок: S4 ∥ S9 ∥ A2 → A8 backend → A8 frontend → Q2.

---

## 9. Success Criteria

1. `cd frontend && npm run audit:ci` — exit 0; `react-router-dom` не 7.18.0 (patched).  
2. При `APP_ENV=production` и `APP_DEBUG=true` приложение не стартует (ValueError на settings).  
3. Существует ADR `deployment-single-instance.md`; rate-limiting doc указывает на single-worker constraint.  
4. `GET /api/v1/logistics/kp-search` доступен logistics; ответ без financial keys.  
5. КП со статусом `в архиве` не возвращается logistics search (по kp_id и по customer).  
6. `GET /api/v1/commercial/archive/search` для logistics → **403**.  
7. `CreateShipmentDialog` использует logistics API, не `archiveApi.search`.  
8. `draftItems.test.ts` покрывает ключевые преобразования payload.  
9. Diff не включает A1 split / Redis / OCR / CSP.

---

## 10. Open Questions

| # | Вопрос | Default |
|---|--------|---------|
| Q1 | Включать `выполнено` в ACL B? | **Нет** (только `в работе`, `На СГП`) |
| Q2 | Где гонять `npm run audit:ci`, если нет `.github/workflows`? | Документировать в ADR + добавить workflow **или** script в существующий deploy Makefile — уточнить у вас |
| Q3 | Startup guard: hard-fail при multi-worker или только warning + ADR? | **ADR + явный `workers=1` в compose/docs**; hard-fail multi-worker сложно детектить из settings — не блочить без сигнала |
| Q4 | Переносить ли существующий `LogisticsKpSearchItem` из `schemas/archive.py` в `schemas/logistics.py`? | **Да** — канон в logistics; archive slim для legacy role-aware можно удалить вместе с ролью logistics на archive search |

→ Ответьте на Q1–Q4 или **`defaults ok`** — затем Phase 2: PLAN.

---

## 11. Risks

| Риск | Mitigation |
|------|------------|
| Логист не находит нужный КП (ещё в архиве) | UX: сообщение «КП должно быть в работе / на СГП»; менеджер переводит в производство |
| npm bump ломает роутинг | Прогнать frontend tests + typecheck |
| Нет GitHub Actions | Q2 — зафиксировать место запуска audit:ci |
| Дублирование slim schema archive vs logistics | Q4 — один канон в logistics |

---

*SPECIFY complete. Waiting for human review before PLAN.*
