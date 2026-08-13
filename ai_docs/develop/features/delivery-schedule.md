# График поставки (delivery schedule)

**Status**: ✅ MVP implemented (Week 1–3); **READY_FOR_HUMAN_QA**  
**Date**: 2026-08-07  
**Reports**: [week1](../reports/delivery-schedule-week1.md), [week2](../reports/delivery-schedule-week2.md), [week3](../reports/delivery-schedule-week3.md)  
**Spec**: [`docs/specs/delivery-schedule.md`](../../../docs/specs/delivery-schedule.md)

## Description

Разбиение КП на партии поставки с проверкой производственной ёмкости (светофор), импортом из XLSX-шаблона и генерацией документа «График поставки» (XLSX + PDF).

## How It Works

1. Схема: `delivery_schedule` / `delivery_batch` / `delivery_batch_item` (`core/kp_db_schema.py`).
2. Светофор: `core/delivery_schedule_check.check_batches` + константы `core/production_capacity.py`.
3. Сервис: `app/services/delivery_schedule_service.py` — PUT/GET, occupancy/readiness, `generate_document`.
4. XLSX шаблон/импорт: `core/delivery_schedule_xlsx.py`; документ XLSX там же; PDF — `core/delivery_schedule_pdf.py`.
5. UI: `frontend/src/features/delivery-schedule/` + кнопка в `OfferDetailsDrawer`.

## API Endpoints

- `GET/PUT /api/v1/commercial/archive/{kp_id}/delivery-schedule`
- `GET .../delivery-schedule/template`
- `POST .../delivery-schedule/import` (не сохраняет)
- `GET .../delivery-schedule/document?fmt=xlsx|pdf`

Roles: `admin`, `manager`.

## Components (frontend)

- `DeliveryScheduleDialog` / `DeliveryScheduleEditor` — редактор партий
- `BatchCard` / `BatchStatusChip` — карточка + светофор
- `ImportScheduleDialog` — dropzone XLSX
- `ScheduleDocumentButtons` — шаблон / XLSX / PDF

## Known Issues / Gaps

- Бейдж «есть график» в списке архива ждёт флаг API `has_delivery_schedule` (UI готов).
- CP2 калибровка светофора ±20% на 3–5 заказах — PARTIAL.
- Макет документа (R5) — NEEDS_USER_REVIEW.
- UX ≤10 мин на счёт 234 — human QA.

## Related Tasks

- Week 1: T1–T3 — done
- Week 2: T4–T7 — done
- Week 3: T8–T12 — done (report week3)
