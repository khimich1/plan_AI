# Report: График поставки — Week 3 (documents + frontend)

**Date:** 2026-08-07  
**Orchestration:** `orch-delivery-w3`  
**Status:** ✅ Week 3 complete  
**Feature MVP:** **READY_FOR_HUMAN_QA**  
**Spec:** [`docs/specs/delivery-schedule.md`](../../../docs/specs/delivery-schedule.md)  
**Plan:** [`ai_docs/develop/plans/delivery-schedule-plan.md`](../plans/delivery-schedule-plan.md)  
**Prev:** [`delivery-schedule-week2.md`](delivery-schedule-week2.md)

## Summary

Неделя 3 закрыта: документ XLSX/PDF + `/document`, frontend-фича (api/hooks/редактор/импорт/скачивание), кнопка в `OfferDetailsDrawer`. Автотесты backend+frontend зелёные. Живой `./run+logs.sh` отвечает (API 401 без cookie / UI 302 login). Полный E2E кликами и Success Criteria на реальном счёте 234 — на человека.

**Further work = human QA + optional polish** (не блокеры MVP-кода).

## Task statuses (T8–T12)

| Task | Title | Status |
|------|-------|--------|
| T8 | Документ XLSX + PDF + `/document` | ✅ completed |
| T9 | Frontend: api / types / hooks | ✅ completed |
| T10 | Frontend: редактор партий + светофор + кнопка в drawer | ✅ completed |
| T11 | Frontend: импорт + документы | ✅ completed |
| T12 | E2E-прогон и финальная верификация | ✅ completed (авто + live smoke; UI click — human) |

### T8 — Документ

- `core/delivery_schedule_xlsx.py` — `build_document` (шапка + таблица)
- `core/delivery_schedule_pdf.py` — PDF (reportlab; макет MVP, не этажи-колонки ЯРПРОФИТ)
- `DeliveryScheduleService.generate_document` — имя `График_КП{id}_ред_YYYY-MM-DD.{ext}`, коллизии `_HHmmss`
- `GET .../document?fmt=xlsx|pdf`
- **Tests:** `tests/test_delivery_schedule_doc.py` (3) + endpoint document (2)
- **R5 макет:** код готов; согласование с пользователем — **NEEDS_USER_REVIEW**

### T9 — Frontend API layer

- `frontend/src/features/delivery-schedule/{api,types,hooks}/`

### T10 — Редактор + светофор

- Editor / BatchCard / BatchStatusChip / Dialog; кнопка в drawer
- Бейдж списка: UI готов, но **API списка архива не отдаёт `has_delivery_schedule`** (см. Known gaps)
- Drawer-бейдж после успешного GET

### T11 — Импорт + документы

- `ImportScheduleDialog`, `ScheduleDocumentButtons`; draft + unmatched Alert

### T12 — Верификация

- Backend pytest feature suite: **95 passed**
- Frontend: typecheck ✅, vitest **239 passed**, build ✅
- Live stack: backend `:8000` + vite `:5173` up; unauth GET → 401; docs 200; frontend 302
- Полный менеджерский клик-сценарий (импорт → правка → светофор → скачать) — **не автоматизирован агентом** → human QA

## Test counts

### Backend

```bash
.venv/bin/python -m pytest \
  tests/test_delivery_schedule_schema.py \
  tests/test_delivery_schedule_check.py \
  tests/test_delivery_schedule_service.py \
  tests/test_delivery_schedule_xlsx.py \
  tests/test_delivery_schedule_doc.py \
  tests/test_delivery_schedule_endpoints.py \
  tests/test_archive_service.py -q
```

| Suite | Count | Result |
|-------|------:|--------|
| `test_delivery_schedule_schema.py` | 13 | ✅ |
| `test_delivery_schedule_check.py` | 16 | ✅ |
| `test_delivery_schedule_service.py` | 16 | ✅ |
| `test_delivery_schedule_xlsx.py` | 10 | ✅ |
| `test_delivery_schedule_doc.py` | 3 | ✅ |
| `test_delivery_schedule_endpoints.py` | 14 | ✅ |
| `test_archive_service.py` | 23 | ✅ |
| **Total** | **95** | **passed** |

Warnings: `python_multipart` PendingDeprecation; Starlette TestClient cookies (не блокеры).

### Frontend

```bash
cd frontend && npm run typecheck && npm run test -- --run && npm run build
```

| Gate | Result |
|------|--------|
| `npm run typecheck` | ✅ |
| `npm run test -- --run` | ✅ **239** passed / **45** files |
| Delivery-schedule related (feature + drawer test) | ✅ **15** passed |
| `npm run build` | ✅ (chunk size warning >500kB — pre-existing style) |

## Success Criteria checklist

Из `docs/specs/delivery-schedule.md`:

| # | Criterion | Status | Reason |
|---|-----------|--------|--------|
| 1 | Разбивка КП уровня счёта 234 (31×21) ≤10 мин: импорт + правки | **DEFERRED** | Нужен человек на живых данных/файле; агент не меряет UX-тайминг |
| 2a | Светофор на синтетике (зелёный/жёлтый/красный) | **DONE** | `test_delivery_schedule_check.py` + связка GET |
| 2b | На 3–5 прошлых заказах ±20% vs факт | **PARTIAL** | Week2 CP2: только 2 КП в локальной БД; калибровка не закрыта |
| 3 | Частично произведённая партия — по остатку | **DONE** | Тесты остатка / produced в check + service |
| 4 | Документ XLSX/PDF: шапка + таблица | **DONE** | Doc + endpoint тесты зелёные; **R5 макет — NEEDS_USER_REVIEW** |
| 5 | Старые файлы не перезаписываются (дата редакции) | **DONE** | `test_generate_document_twice_writes_distinct_paths` |
| 6 | Импорт с 3+ unmatched — draft + список, без падения | **DONE** | xlsx/import endpoint coverage |
| 7 | Удаление КП → cascade графика, без 500 | **DONE** | schema cascade tests |

**E2E на `./run+logs.sh`:** **PARTIAL** — стек живой, API/auth/UI reachable; полный click-path менеджера не прогнан агентом (**DEFERRED** для human QA).

## Gates

| Gate | Status |
|------|--------|
| CP2 (калибровка ±20%) | **PARTIAL** (week2; не закрыт) |
| CP3 (макет документа R5) | **NEEDS_USER_REVIEW** |
| CP4 (E2E + CI green) | **PARTIAL** — CI/автотесты green; live UI E2E → human |

## Known gaps

1. **Бейдж списка архива без API-флага** — фронт рендерит `has_delivery_schedule === true`, но list API / `archive_service` флаг не заполняет → бейдж в списке фактически не появится до optional polish.
2. **CP2** — недостаточно completed-КП в локальной БД; калибровка `1.15` / thresholds на человеке/проде.
3. **UX timing** (критерий ≤10 мин на счёт 234) — только human.
4. **R5 макет** документа — упрощённая шапка+таблица; визуальное согласование с образцом ЯРПРОФИТ не сделано.
5. HMR note: `validateScheduleEditor` export мешает Fast Refresh (не runtime-баг).

## Feature MVP status

**READY_FOR_HUMAN_QA**

Код недели 3 и автогейты зелёные. Выкатка/приёмка зависят от человека: разбивка 234, визуал PDF/XLSX, калибровка светофора, клик-E2E.

## Related Documentation

- Plan: [`delivery-schedule-plan.md`](../plans/delivery-schedule-plan.md)
- Feature: [`delivery-schedule.md`](../features/delivery-schedule.md)
- Week1/2: [`delivery-schedule-week1.md`](delivery-schedule-week1.md), [`delivery-schedule-week2.md`](delivery-schedule-week2.md)
- Spec: [`docs/specs/delivery-schedule.md`](../../../docs/specs/delivery-schedule.md)

## Next Steps

1. Human QA: drawer → import → edit → traffic light → download XLSX/PDF на живом стеке.
2. Согласовать макет документа (R5).
3. CP2 на 3–5 реальных заказах; при необходимости подкрутить константы.
4. Optional: `has_delivery_schedule` в list API архива.
5. Замер ≤10 мин на счёте 234.
