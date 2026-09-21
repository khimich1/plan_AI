# Аудит: график поставки

**Дата**: 2026-09-21
**Скоуп**: модуль «график поставки» (delivery-schedule) — backend `app/api/v1/endpoints/delivery_schedule.py`, сервисы `delivery_schedule_service.py`, `capacity_gate_service.py`, schemas `app/schemas/delivery_schedule.py`, core `delivery_schedule_check.py`, `delivery_schedule_xlsx.py`, `delivery_schedule_pdf.py`, `production_capacity.py`; frontend `frontend/src/features/delivery-schedule/`; кнопка в `OfferDetailsDrawer`; тесты `test_delivery_schedule_*`. **Не входит**: bot, конструктор КП / OCR.
**Реестр**: [FINDINGS.md](./FINDINGS.md)
**Прогон 0**: [2026-08-11-delivery-schedule-audit.md](./2026-08-11-delivery-schedule-audit.md) (локальные ID, не FINDINGS). Последнее касание реестра в смежном скоупе: [2026-09-04-archive-production-audit.md](./2026-09-04-archive-production-audit.md).

## Дельта

| ID | Было | Стало | Суть |
|----|------|-------|------|
| A5 | open | open | Raw SQL в `delivery_schedule_service.py:494–500`, `:647–810`; репозитория нет |
| A7 | open | open | Неполный DI: `get_delivery_schedule_service` без constructor injection (platform, приложение) |
| A23 | — | open | **NEW P1** — God-module `DeliveryScheduleService` ~878 LOC, ~24 метода |
| S1 | by-design | by-design | Общий архив КП — не IDOR |
| A15 | by-design | by-design | `owner_user_id` вне policy доступа |
| S8 | open | open | `str(exc)` в HTTP: `delivery_schedule.py:53`, `:70`, `:73–76`, `:92`, `:122`, `:125–128`, `:150`, `:153–156` |
| S3 | open | open | In-process rate limit; DS import/document без rate limit (platform, приложение) |
| S9 | open | open | CSRF multipart до проверки токена; пересечение с DS upload (auth, приложение) |
| Q6 | open | open | Две реализации `get_global_calendar_info`; DS через `PlanDistributionService`, R2 fallback vs `plan_storage` MAX_TRACKS |
| Q9 | open | open | God-component `OfferDetailsDrawer`; wiring графика не корневая причина |
| Q18 | open | open | `DeliveryScheduleDialog.tsx:92` — `eslint-disable` exhaustive-deps |
| Q23 | open | open | Archive `round` vs production FE `ceil` (101 м → 2 vs 1); DS светофор использует `ceil` в `check_batches:200` |
| Q25 | — | open | **NEW** — FE бейдж `has_delivery_schedule`, list API поле не отдаёт |
| S17 | — | open | **NEW** — импорт XLSX без magic bytes / без лимита строк |
| S18 | — | open | **NEW** — formula injection в генерируемом XLSX |
| S19 | — | open | **NEW** — неограниченный PUT: batches / items / name |
| Q27 | — | open | **NEW** — предупреждение `produce_by ≤ deliver_from` не реализовано |
| Q28 | — | open | **NEW** — три дублирующих SELECT из `kp_plates` |
| Q29 | — | open | **NEW** — нет service-тестов import happy-path и GET red |
| Q30 | — | open | **NEW** — причины unmatched на английском в UI |

- Закрыто: 0
- Открыто (подтверждено): 9
- Новое: 9
- Не воспроизвелось: 0

**Открытые P0 / P1 в скоупе**: 0 / 2

## Действия (максимум 8)

1. **[A5] + [A23]** — `DeliveryScheduleRepository`; разрезать god-service: traffic-light / document / import отдельно от оркестрации.
2. **[Q25]** — заполнять `has_delivery_schedule` в list API архива (EXISTS / LEFT JOIN на `delivery_schedule`).
3. **[S8]** — generic HTTP-сообщения наружу; детали в лог (`delivery_schedule.py` endpoints).
4. **[S17]** — проверка magic bytes ZIP/XLSX и cap строк/листов при import.
5. **[S18]** — санитизация formula-префиксов (`=`, `+`, `-`, `@`) в экспортируемом XLSX.
6. **[S19]** — `Field(max_length=…)` / `max_items` на PUT-схеме (batches, items, name).
7. **[Q27]** — soft-warning `produce_by > deliver_from` в API (`warnings[]`) и/или Alert в редакторе.
8. **[Q29]** — service-тесты: import_draft happy-path; GET со светофором red через сервисный слой.

## Открытые P0 / P1

### [A5] Сервисы обходят repository, raw SQL
**Статус**: open
**Улика**: `app/services/delivery_schedule_service.py:494–500`, `:647–810`; grep `app/repositories` — нет `DeliveryScheduleRepository` / *delivery* repo
**Зачем**: Raw SQL в сервисе обходит слой репозиториев, дублирует запросы к `kp_plates` / `delivery_schedule_*`, усложняет миграцию схемы и изолированные unit-тесты persistence.

### [A23] God-module DeliveryScheduleService
**Статус**: open
**Улика**: `app/services/delivery_schedule_service.py` — 878 строк, ~24 метода (`replace`, `get`, `_enforce_capacity_gate`, `_enrich_with_traffic_light`, `import_draft`, `generate_document`, `_replace_batches`, …); PUT → `get()` повторно прогоняет I/O и светофор
**Зачем**: В одном классе смешаны CRUD, валидация, светофор, импорт XLSX, генерация документов, календарь и readiness-mapping; высокая связность — изменение одной ответственности рискует регрессией во всех остальных.

## Приложение

### Закрыто относительно прогона 0 (2026-08-11, локальные ID — не FINDINGS)

| Было (локально) | Статус |
|-----------------|--------|
| AuthZ GET/import (IDOR) | ✅ `assert_offer_read_access` + тесты 403 (`test_delivery_schedule_endpoints.py:321–354`) |
| GET `/template` AuthZ | ✅ `build_template_bytes` вызывает `assert_offer_read_access` (`:228`) |
| Produced false-green по identity | ✅ пропорциональное `_load_produced_by_plate_id`; тест `test_produced_splits_on_sgp_across_same_identity_plates` |
| Конфликт дат в одной партии при import | ✅ `REASON_CONFLICTING_BATCH_DATES` |
| Broad except / silent degrade светофора | ✅ `traffic_light_degraded` + узкий except `_TRAFFIC_LIGHT_SOURCE_ERRORS` |
| XSS в DS UI | Не воспроизведено (нет `dangerouslySetInnerHTML`) |

**Частично**: occupancy через `PlanDistributionService` (старый drift ёмкости). R2 fallback — два источника «5 дорожек» → **[Q6]**.

**Открыто без нового FINDINGS ID** (упоминание под A23): N+1 `_build_view` — SELECT items на каждую партию; `generate_document` не регистрирует файл в `kp_files`; `delivery_schedule.status` всегда `'draft'`.

### Platform / смежные скоупы (не в P1 модуля)

**[A7] Неполный DI** — `app/dependencies/services.py:41–156`: `get_delivery_schedule_service` — фабрика без constructor injection; DS использует inline `PlanDistributionService()`, `KpReadinessService()`.

**[S3] Rate limiting** — import/document DS без per-endpoint limit; platform in-process store (см. A2).

**[S9] CSRF multipart** — upload XLSX парсится до проверки CSRF-токена; kp-touchpoint: `csrf.py:42–47`, `httpClient.ts:130–132`.

**[S1] / [A15] by-design** — общий архив; `owner_user_id` вне policy (`offer_access.py:26–32`).

### Безопасность (Medium)

**[S8]** — `delivery_schedule.py` отдаёт `str(exc)` клиенту через `HTTPException` / `raise_unprocessable_client_error` на строках `:53`, `:70`, `:73–76`, `:92`, `:122`, `:125–128`, `:150`, `:153–156`.

**[S17]** — POST `/import`: нет проверки сигнатуры ZIP/XLSX (magic bytes); парсер без верхней границы строк/листов → DoS через zip bomb / огромный файл.

**[S18]** — экспорт XLSX: значения ячеек с префиксами `=`, `+`, `-`, `@` могут интерпретироваться Excel как формулы (formula injection).

**[S19]** — PUT `/delivery-schedule`: схема без `max_length` на name и без `max_items` на batches/items → переполнение БД / DoS большим payload.

### Качество / архитектура (Medium)

**[Q6]** — DS: `_enrich_with_traffic_light` вызывает `PlanDistributionService().get_global_calendar_info()` (`:603`); при сбое — R2 fallback из `production_capacity.py` (5 дорожек) vs `plan_storage.MAX_TRACKS_PER_DAY` — расхождение с layout-контуром.

**[Q9]** — `OfferDetailsDrawer` ~1091 строк; кнопка «График поставки» и `readOnly={!canEdit}` работают; размер drawer не вызван wiring DS.

**[Q18]** — `DeliveryScheduleDialog.tsx:92`: `eslint-disable-line react-hooks/exhaustive-deps` — намеренный reset mutation только при close.

**[Q23]** — archive BE `round` vs production FE `ceil`; DS `check_batches` (`:200`) использует `ceil` — согласован с production FE, не с archive.

**[Q25]** — `ArchiveOfferList.tsx:172` рендерит бейдж при `has_delivery_schedule === true`; list API / `archive_service` поле не заполняет → бейдж не появляется.

**[Q27]** — спека §215–216: soft-warning если `produce_by > deliver_from`; schema проверяет только `deliver_from ≤ deliver_to` (`delivery_schedule.py:62–64`).

**[Q28]** — три почти одинаковых SELECT из `kp_plates` в `delivery_schedule_service.py` (meta, import mapping, view enrichment).

**[Q29]** — pytest покрывает check/xlsx/schema/endpoints; нет service-level: import_draft happy-path; GET red traffic light через `_enrich_with_traffic_light`.

**[Q30]** — `unmatched_rows[].reason` (`unknown mark`, `bad date`, `conflicting batch dates`) — machine-readable English без локализации в UI.

### Что сделано хорошо

- Тонкий роутер → сервис → `core/*`; `delivery_schedule_check.py` без импортов `app.*`.
- AuthZ read/write закрыт: GET, import, template, document — `assert_offer_read_access`; PUT — `assert_offer_write_access`.
- Read-only просмотр архивных КП: кнопка не disabled, `readOnly={!canEdit}`.
- `traffic_light_degraded` + узкий except источников светофора; Alert в Editor.
- Produced mapping пропорционально qty (largest remainder).
- Параметризованный SQL; PDF escape через `xml.sax.saxutils.escape`.
