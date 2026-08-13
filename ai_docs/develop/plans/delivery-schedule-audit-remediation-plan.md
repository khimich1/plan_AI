# Plan: Remediation аудита «График поставки» (P0 + «скоро»)

Дата: 2026-08-11.  
Спека: `docs/specs/delivery-schedule-audit-remediation.md` (объём подтверждён).  
Аудит: `ai_docs/develop/audits/2026-08-11-delivery-schedule-audit.md`.  
Статус: **реализовано** (T1–T6 done).

## Цель

Закрыть AuthZ IDOR и ближайшие риски корректности/UX из аудита минимальным diff,
без эпика god-service и без калибровки констант ёмкости.

## Компоненты

| ID | Компонент | Что меняет |
|----|-----------|------------|
| C1 | AuthZ read path | `get` / `import_draft` + endpoints: `user` + `assert_offer_read_access` |
| C2 | HTTP семантика 403/404 | порядок: offer exists → read ACL → schedule exists |
| C3 | Produced mapping (A5) | распределение `on_sgp` по `plate_id` без двойного вычета |
| C4 | Traffic light degrade (Q1) | узкий except + `traffic_light_degraded` в schema/view + Alert на FE |
| C5 | Import date conflict (Q4) | `parse_template`: conflicting dates → unmatched |
| C6 | Read-only UI (A3) | кнопка в drawer всегда enable для плит; `readOnly` уже есть |

## Решение по A5 (зафиксировано в плане)

**Не** отдавать полный `on_sgp` identity каждому `plate_id`.

Алгоритм (минимальный, без нового SQL):

1. Как сейчас: `list_positions` → `on_sgp` по identity-key.
2. Сгруппировать `kp_plates` текущего КП по тому же identity-key.
3. Для каждой группы с `S = on_sgp` и списком plate_id с qty `q_i`:
   - распределить целое `S` пропорционально `q_i` (largest remainder),  
     `Σ allocated_i = min(S, Σ q_i)`.
4. `produced[plate_id] = allocated_i`.

Итог: суммарный вычет по identity не превышает факт СГП; ложный green от «каждому полный on_sgp» исчезает.

Альтернатива «прямой SQL к completed_plates» — запасной путь, если readiness окажется непригоден; в MVP remediation не нужна.

## Порядок реализации

```
T1 AuthZ (C1+C2) ──→ T2 Produced (C3) ──→ T3 Degrade (C4) 
T4 Import dates (C5)  ‖ можно параллельно с T2/T3 (другой файл core/)
T5 Read-only UI (C6)  ‖ можно параллельно с T1–T4 (только frontend)
```

Рекомендуемая последовательность одним разработчиком:

1. **T1 AuthZ** — P0, блокирует безопасное использование  
2. **T4 Import dates** ‖ **T5 Read-only UI** — быстрые, независимые  
3. **T2 Produced** — корректность светофора  
4. **T3 Degrade** — после T2 (тот же `_enrich_with_traffic_light`)

## Риски

| Риск | Митигация |
|------|-----------|
| `assert_offer_*` бросает `HTTPException` из сервиса | Уже так в `generate_document`; FastAPI пробрасывает 403 — не оборачивать в 500 |
| Порядок 404/403: сначала «нет КП» vs «нет доступа» | Сначала `_fetch_offer`; нет offer → 404; есть → `assert_read`; потом schedule |
| Ломаются service-тесты: `get(kp_id)` без `user` | Все вызовы `get`/`import_draft` обновить на `user=`; в unit-тестах передавать admin-dict |
| Пропорция produced округляет не туда | Тест: 2 plates qty 10+10, on_sgp=10 → allocated 5+5 (или 5+5), сумма 10 |
| FE не показывает degraded | Минимум API-поле; Alert в Editor если `traffic_light_degraded` |

## Чекпоинты

- **CP-A (после T1):** pytest endpoints — чужой manager GET/import → 403; свой без графика → 404; PUT/document без регрессии.
- **CP-B (после T2–T5):** полный `pytest tests/test_delivery_schedule_*.py -q` + `npm run typecheck` + точечный FE test drawer.
- **CP-C:** обновить секцию Remediation в audit-отчёте.

## Tasks

- [x] **T1. AuthZ на GET и import (P0)**
  - Acceptance: `service.get(kp_id, user=...)` и `import_draft(..., user=...)` вызывают `_fetch_offer` + `assert_offer_read_access` до чтения данных; endpoints передают `user` (не `_user`); чужой manager → 403; свой КП без графика → 404; нет КП → 404; HTTPException 403 не маскируется.
  - Verify: `.venv/bin/python -m pytest tests/test_delivery_schedule_endpoints.py -q` (+ новые кейсы auth); существующие service-тесты поправлены под `user=`.
  - Files: `app/api/v1/endpoints/delivery_schedule.py`, `app/services/delivery_schedule_service.py`, `tests/test_delivery_schedule_endpoints.py`, `tests/test_delivery_schedule_service.py` (сигнатура)

- [x] **T2. Produced без ложного green (A5)**
  - Acceptance: `_load_produced_by_plate_id` распределяет `on_sgp` по plate_id группы пропорционально qty; сумма allocated ≤ on_sgp и ≤ Σ qty группы; тест с двумя строками одной марки.
  - Verify: `.venv/bin/python -m pytest tests/test_delivery_schedule_service.py -q -k produced`
  - Files: `app/services/delivery_schedule_service.py`, `tests/test_delivery_schedule_service.py`

- [x] **T3. Узкий except + traffic_light_degraded (Q1)**
  - Acceptance: в `_enrich_with_traffic_light` ловятся только ожидаемые сбои источников (readiness/calendar/workdays — явный tuple или обёртка); прочие исключения пробрасываются; при ожидаемом сбое — `traffic_light_degraded=True`, statuses null; поле в `DeliveryScheduleView` + TS type; Alert в Editor при true.
  - Verify: service-тест degrade; `npm run typecheck`
  - Files: `app/services/delivery_schedule_service.py`, `app/schemas/delivery_schedule.py`, `frontend/.../types/deliverySchedule.ts`, `frontend/.../DeliveryScheduleEditor.tsx`, tests

- [x] **T4. ‖ Конфликт дат при import (Q4)**
  - Acceptance: при merge партии с тем же `name`, если `deliver_from`/`deliver_to`/`produce_by` отличаются от уже накопленных — строка уходит в `unmatched_rows` с reason вроде `conflicting batch dates` (и/или русская константа); позиции не мержатся в испорченную партию.
  - Verify: `.venv/bin/python -m pytest tests/test_delivery_schedule_xlsx.py -q`
  - Files: `core/delivery_schedule_xlsx.py`, `tests/test_delivery_schedule_xlsx.py`

- [x] **T5. ‖ Read-only кнопка в drawer (A3)**
  - Acceptance: для не-simple КП кнопка «График поставки» не `disabled` из-за статуса; `readOnly={!canEditDeliverySchedule}` сохраняется; title подсказывает «просмотр» vs «редактирование»; Save/Import в readOnly скрыты (уже в Dialog).
  - Verify: `cd frontend && npm run test -- --run OfferDetailsDrawer` (или ручной чек) + typecheck
  - Files: `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`, при необходимости `.test.tsx`

- [x] **T6. Документация remediation**
  - Acceptance: в audit-отчёте секция Remediation: Fixed (список ID) + Remaining.
  - Verify: файл обновлён
  - Files: `ai_docs/develop/audits/2026-08-11-delivery-schedule-audit.md`

## Оценка

~0.5–1 день одним разработчиком (T1 критичен и короткий; T2–T5 — по 1–2 часа).

## Вне плана

A2 god-service, A4 constants, A6/A7 archive polish, S3–S5 upload hardening, Q5 produce_by warning — backlog аудита.
