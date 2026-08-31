# Отчёт аудита проекта

**Дата**: 2026-08-15  
**Область**: Модуль «СГП и логистика» (backend + frontend)  
**Аудиторы**: senior-reviewer + security-auditor + reviewer  

**Охват (группы файлов):**
- Backend API: `app/api/v1/endpoints/production.py` (SGP-эндпоинты), `app/api/v1/endpoints/logistics.py`, связанные schemas (`app/schemas/logistics.py`, `app/schemas/production.py`)
- Backend services: `app/services/sgp_service.py`, `shipment_service.py`, `shipment_completion_service.py`, `carrier_service.py`, `kp_readiness_service.py`, `delivery_schedule_service.py`
- Persistence: `app/repositories/shipment_repository.py`, inline SQL в сервисах, `core/kp_db_shipments.py`
- Core helpers: `core/shipment_packing.py`, `core/delivery_schedule_*`
- Frontend: `frontend/src/features/production/` (SGP), `frontend/src/features/logistics/` (рейсы, перевозчики, график поставок)

---

## Краткое резюме

**Общая оценка здоровья**: **6.0/10**

**Расчёт Health Score:**
- Старт: 10
- Critical: 0 → −0
- High: 10 находок × 0.5 = 5.0, потолок −3 → **−3**
- Medium: 26 × 0.1 = 2.6, потолок −1 → **−1**
- Итог: 10 − 3 − 1 = **6.0/10**

| Серьёзность | Архитектура | Безопасность | Качество кода | Итого |
|-------------|-------------|--------------|---------------|-------|
| Critical    | 0           | 0            | 0             | **0** |
| High        | 5           | 2            | 3             | **10** |
| Medium      | 10          | 5            | 11            | **26** |
| Low         | 7           | 5            | 4             | **16** |

**Рекомендация**: устранить High security **S1–S2** и архитектурные **A1–A5** до расширения логистического функционала; god-сервисы и разрыв bounded context — через `/refactor` по инкрементам; ACL и пагинация — через `/implement` в первую очередь.

Модуль «СГП и логистика» функционален и опирается на сильные core-хелперы (`shipment_packing`, SSOT резервов, `ShipmentCompletionService`, DI), но накопил значительный архитектурный и качественный долг: два god-сервиса (~900 и ~1000 строк), разрыв bounded context между production и logistics, отсутствие пагинации и ACL на критичных путях. Безопасность: логист может привязать КП любого статуса к рейсу; полный дамп СГП создаёт риск DoS.

### Краткая архитектурная сводка (Phase 1)

**Слои:** API → Services → Persistence (`ShipmentRepository` | inline SQL) → core helpers (`kp_db_shipments`, `shipment_packing`, `delivery_schedule_*`).

**Границы контекстов:** СГП — под `production` (роль production); рейсы — под `logistics` (роль logistics); график поставок (`DeliverySchedule`) — в commercial/archive (роль manager).

**Сильные стороны:** `core/shipment_packing`, единый источник правды по резервам, выделенный `ShipmentCompletionService`, dependency injection на уровне FastAPI.

**Слабые стороны:** отсутствие `SgpRepository`, SQL inline в сервисах, ручная инвалидация кэша между FE-модулями production/logistics, два алгоритма propose без Strategy.

---

## Критические проблемы

Критических находок (**Critical = 0**) не выявлено.

---

## Проблемы высокого приоритета

### [A1] God-сервис SgpService (~893 строк)

**Категория**: Архитектура  
**Где**: `app/services/sgp_service.py` — CRUD плит, резервы, XLSX-экспорт, прогресс; SQL inline, нет `SgpRepository`  
**Влияние**: любое изменение логики склада/резервов требует правок монолита; сложно тестировать и ревьюить; нарушение SRP и слоя persistence.  
**Исправление**: выделить `SgpRepository`; разрезать сервис на use-case'ы (CRUD, резервы, экспорт); SQL — только в repository; `/refactor` инкрементально.

### [A2] God-сервис ShipmentService (~1001 строк)

**Категория**: Архитектура  
**Где**: `app/services/shipment_service.py` — CRUD рейсов, propose (legacy + v2), поиск КП, каталог свай, экспорт, события  
**Влияние**: высокий coupling, риск регрессий при любом изменении логистики; смешение orchestration и persistence.  
**Исправление**: разделить на `ShipmentProposeService`, `ShipmentCatalogService`, `ShipmentExportService`; pass-through методы — в repository; `/refactor`.

### [A3] Разрыв bounded context СГП → отгрузка

**Категория**: Архитектура  
**Где**: `production.py` + `/production/sgp/*` (роль production) vs `logistics.py` + `/logistics/*` (роль logistics); FE: `features/production` vs `features/logistics`; ручная инвалидация React Query кэша  
**Влияние**: рассинхрон UI после операций на складе/рейсе; дублирование ключей запросов; нет единого логистического контекста.  
**Исправление**: общий bounded context «Logistics» с подмодулями SGP/Shipments; shared query keys / event bus для инвалидации; единая роль или явный cross-context contract.

### [A4] DIP: repository импортирует ошибку из service

**Категория**: Архитектура  
**Где**: `app/repositories/shipment_repository.py:19` — `from app/services/shipment_errors import ShipmentError`  
**Влияние**: инверсия зависимостей — persistence зависит от application layer; циклические импорты, сложность тестирования repo изолированно.  
**Исправление**: перенести `ShipmentError` в `app/domain/` или `core/shipment/errors.py`; repository и service импортируют из domain.

### [A5] Нет пагинации list_plates / GET /production/sgp/plates

**Категория**: Архитектура (+ связано с S2)  
**Где**: `SgpService.list_plates`, эндпоинт `GET /production/sgp/plates`  
**Влияние**: при росте склада — полная выборка в память, медленный ответ, нагрузка на БД и клиент.  
**Исправление**: cursor/offset pagination, `limit`/`offset` в schema и SQL; `/implement`.

### [S1] ACL статуса КП только в kp-search

**Категория**: Безопасность  
**Где**: фильтр статуса КП есть в kp-search; отсутствует в `create/reuse-transport/patch` shipments — логист может привязать КП любого статуса  
**Влияние**: привязка черновика/архивного КП к рейсу; нарушение бизнес-правил и учёта готовности.  
**Исправление**: единый guard `_assert_kp_eligible_for_shipment(status)` на всех mutating-путях; `/implement` с тестами.

### [S2] Полный дамп СГП без пагинации — DoS / resource exhaustion

**Категория**: Безопасность  
**Где**: `GET /production/sgp/plates` без лимита (см. A5)  
**Влияние**: аутентифицированный пользователь может запросить весь склад → исчерпание памяти/CPU, блокировка для других.  
**Исправление**: обязательный `limit` (default + max), rate limit; `/implement` вместе с A5.

### [Q1] ShipmentItemsSection.tsx ~644 строк без тестов propose/confirm

**Категория**: Качество кода  
**Где**: `frontend/src/features/logistics/components/ShipmentItemsSection.tsx`  
**Влияние**: ключевой UI-сценарий (propose → confirm) без автотестов; регрессии при рефакторинге незаметны.  
**Исправление**: unit/integration тесты на propose flow, confirm, edge cases; `/implement`.

### [Q2] DRY: _fetch_propose_candidates vs available_by_kp

**Категория**: Качество кода  
**Где**: `app/services/shipment_service.py` — два почти идентичных SQL-запроса  
**Влияние**: расхождение логики при правке одного пути; дублирование ~50+ строк.  
**Исправление**: общий метод в `ShipmentRepository` или shared query builder; `/refactor`.

### [Q3] SQL-монолит list_positions в kp_readiness_service

**Категория**: Качество кода  
**Где**: `app/services/kp_readiness_service.py` — 4 одинаковых подзапроса в `list_positions`  
**Влияние**: сложность поддержки, риск ошибки при изменении одного подзапроса; плохая читаемость.  
**Исправление**: CTE или helper `_positions_subquery(kind)`; `/refactor`.

---

## Проблемы среднего приоритета

### Архитектура

#### [A6] Непоследовательный persistence

**Где**: `ShipmentRepository` есть; SGP и Carrier — SQL inline в сервисах  
**Влияние**: разный стиль доступа к данным; сложнее мокать и переиспользовать запросы.  
**Исправление**: `SgpRepository`, `CarrierRepository` по образцу `ShipmentRepository`.

#### [A7] Дубль _assert_in_work

**Где**: `ShipmentService` и `ShipmentCompletionService`  
**Влияние**: расхождение guard-логики при правке одного места.  
**Исправление**: общий helper в domain или mixin.

#### [A8] ~15 pass-through методов ShipmentService → repo

**Где**: `ShipmentService` делегирует в repository без добавленной логики  
**Влияние**: лишний слой, раздувание API сервиса.  
**Исправление**: вызывать repository из endpoint или тонкий facade только для оркестрации.

#### [A9] Два propose без Strategy

**Где**: t30plus legacy vs v2 packing в `ShipmentService`  
**Влияние**: условная логика разветвлена; сложно добавить v3 или A/B.  
**Исправление**: Strategy/Protocol `ProposeAlgorithm` с двумя реализациями.

#### [A10] FE logistics импортирует productionKeys/sgpKeys

**Где**: `frontend/src/features/logistics/` → ключи из `features/production`  
**Влияние**: coupling между bounded contexts на клиенте.  
**Исправление**: shared `features/logistics/queryKeys.ts` или cross-context invalidation module.

#### [A11] DeliverySchedule вне логистического контекста

**Где**: commercial/archive, роль manager; не в `logistics.py`  
**Влияние**: график поставок логически связан с рейсами, но живёт отдельно.  
**Исправление**: документировать границу или постепенно перенести под logistics API.

#### [A12] KpReadinessService создаёт SgpService inline

**Где**: `kp_readiness_service.py` — `SgpService()` без DI  
**Влияние**: скрытая зависимость, сложность тестирования с mock.  
**Исправление**: инжектировать `SgpService` через конструктор/FastAPI Depends.

#### [A13] SGP endpoints внутри god-router production.py

**Где**: `app/api/v1/endpoints/production.py`  
**Влияние**: смешение производства и склада в одном router-файле.  
**Исправление**: выделить `sgp.py` router, подключить в `router.py`.

#### [A14] Magic strings 'в производстве' в SgpService

**Где**: `app/services/sgp_service.py` вместо `PlateStatus`  
**Влияние**: опечатки, расхождение с enum; см. также Q4.  
**Исправление**: использовать `PlateStatus.IN_PRODUCTION` или константу из core.

#### [A15] Дубль SQL unlink/relink/reserve

**Где**: повторяющиеся блоки SQL в SgpService для unlink, relink, reserve  
**Влияние**: DRY-нарушение; риск inconsistent updates.  
**Исправление**: private helpers или repository methods.

### Безопасность

#### [S3] Нет rate limit на /logistics/* и /production/sgp/*

**Где**: FastAPI endpoints без throttling  
**Влияние**: злоупотребление API при компрометации учётки или баге клиента.  
**Исправление**: middleware rate limit (slowapi или nginx); per-role limits.

#### [S4] Нет max_length / bounds на строках и списках

**Где**: `app/schemas/logistics.py` — строки и списки без upper bound  
**Влияние**: oversized payload → память, медленный парсинг.  
**Исправление**: `Field(max_length=...)`, `max_items` на list schemas.

#### [S5] XLSX import delivery-schedule без лимита строк

**Где**: импорт графика поставок из XLSX  
**Влияние**: файл с миллионами строк → DoS при парсинге.  
**Исправление**: max rows per sheet (например 10 000), reject early.

#### [S6] weight_kg с клиента без сверки / верхней границы

**Где**: shipment schemas — `weight_kg` от клиента  
**Влияние**: некорректный вес в документах; возможны absurd values.  
**Исправление**: `ge=0, le=MAX_WEIGHT`; опционально пересчёт server-side.

#### [S7] _maybe_write_event пишет PII в exchange_export_dir

**Где**: экспорт событий в файловую систему без контроля прав/ротации  
**Влияние**: утечка PII на диск; неограниченный рост каталога.  
**Исправление**: ACL на директорию, retention policy, redact PII.

### Качество кода

#### [Q4] Magic string 'в производстве' (quality angle на A14)

**Где**: импорт/сравнение статуса без `PlateStatus`  
**Влияние**: хрупкость при переименовании статуса в БД.  
**Исправление**: enum/constant everywhere.

#### [Q5] DRY списание open demand в relink/reserve/reduce_selected

**Где**: повторяющаяся логика списания спроса в SgpService  
**Влияние**: copy-paste bugs.  
**Исправление**: `_reduce_open_demand(conn, ...)`.

#### [Q6] FE формат размеров плиты в 3 местах

**Где**: logistics + production UI helpers  
**Влияние**: inconsistent display (мм vs м, округление).  
**Исправление**: `shared/lib/plateDimensions.ts`.

#### [Q7] Дубль матрицы статусов _format_summary / _format_client_copy

**Где**: shipment formatting helpers  
**Влияние**: расхождение текстов при добавлении статуса.  
**Исправление**: single `STATUS_LABELS` map.

#### [Q8] Повторяющийся transaction boilerplate ≥12 методов

**Где**: SgpService, ShipmentService — `with conn`, commit/rollback  
**Влияние**: шум, риск забыть rollback.  
**Исправление**: `@transactional` decorator или context manager.

#### [Q9] Нет тестов free_plates, build_from_sgp_rows, export_plan_sgx_xlsx

**Где**: SgpService / export paths  
**Влияние**: регрессии экспорта и free_plates незаметны.  
**Исправление**: unit tests с fixture DB.

#### [Q10] Нет FE тестов sgpApi, DeliveryScheduleEditor/Dialog/Import

**Где**: `frontend/src/features/production/api/sgpApi.ts`, delivery schedule components  
**Влияние**: API contract drift без сигнала.  
**Исправление**: vitest mocks + component tests.

#### [Q11] Swallow Exception в plan day mapping без logger

**Где**: mapping дней плана — bare `except`  
**Влияние**: silent failures, сложная диагностика.  
**Исправление**: log warning + re-raise или structured error.

#### [Q12] type: ignore в kp_readiness и shipment_service

**Где**: `# type: ignore` комментарии  
**Влияние**: маскировка реальных type errors.  
**Исправление**: исправить типы или TypedDict/Protocol.

#### [Q13] Мёртвый параметр actor в CarrierService.merge

**Где**: `CarrierService.merge(..., actor=...)` не используется  
**Влияние**: misleading API; audit trail не пишется.  
**Исправление**: использовать для audit log или удалить параметр.

#### [Q14] Повторная валидация membership при complete

**Где**: `ShipmentCompletionService` — duplicate checks  
**Влияние**: лишний код; возможное расхождение сообщений об ошибках.  
**Исправление**: single validation entry point.

---

## Низкий приоритет / предложения

### Архитектура

- **[A16]** FE типы SGP в `sgpApi.ts`, а не в `types/` — Fix: `types/sgp.ts`, re-export из api.
- **[A17]** CarrierService без repository — Fix: `CarrierRepository` при росте логики.
- **[A18]** Lazy imports SgpService (coupling) — Fix: явный import + DI.
- **[A19]** Крупные UI: SgpWarehouseView 510, ShipmentDrawer 491, LogisticsRegistryView 394 — Fix: subcomponents по вкладкам/секциям; `/refactor`.
- **[A20]** Split SQL: kp_db_shipments vs ShipmentRepository vs inline propose — Fix: консолидация в repository layer.
- **[A21]** Readiness release — заглушка, не связана с complete() — Fix: wire release → complete flow или удалить stub.
- **[A22]** Нет FE интеграционных тестов cross-feature invalidation — Fix: e2e test production action → logistics cache refresh.

### Безопасность

- **[S8]** LIKE без ESCAPE `%`/`_` в CarrierService и search_pile_catalog — Fix: escape user input или parameterized ESCAPE clause.
- **[S9]** Детальные доменные ошибки раскрывают ID связей — Fix: generic message наружу, details в logs.
- **[S10]** Allowlist ролей только HTTP (ISS-001 defense-in-depth) — Fix: duplicate check at repository boundary.
- **[S11]** CSP Report-Only + unsafe-inline — Fix: tighten CSP when FE bundle allows.
- **[S12]** UI RequireRole — не барьер (напоминание) — Fix: документировать; сервер — единственный gate.

### Качество кода

- **[Q15]** EN/RU mixed error messages в SgpService — Fix: единый язык (RU) + error codes.
- **[Q16]** eslint-disable exhaustive-deps в logistics drawers — Fix: стабилизировать deps / useEffectEvent.
- **[Q17]** Слабая типизация maybe_write_event / _to_item Any — Fix: TypedDict/Pydantic models.
- **[Q18]** `params: list = []` без generics в CarrierService — Fix: `list[CarrierMergeParam]` или dataclass.

---

## Матрица приоритетов

| ID | Issue | Severity | Effort | Priority |
|----|-------|----------|--------|----------|
| S1 | ACL статуса КП не на всех shipment paths | High | Low | **P0** — немедленно |
| S2/A5 | Полный дамп СГП без пагинации (DoS) | High | Medium | **P0** — немедленно |
| A1 | God-сервис SgpService | High | High | **P1** — этот спринт |
| A2 | God-сервис ShipmentService | High | High | **P1** — этот спринт |
| A3 | Разрыв bounded context СГП↔отгрузка | High | High | **P1** — этот спринт |
| A4 | DIP: repo импортирует ShipmentError из service | High | Low | **P1** — этот спринт |
| Q1 | ShipmentItemsSection без тестов propose/confirm | High | Medium | **P1** — этот спринт |
| Q2 | DRY propose candidates SQL | High | Low | **P1** — этот спринт |
| Q3 | SQL-монолит list_positions | High | Medium | **P1** — этот спринт |
| S3 | Rate limit logistics/sgp | Medium | Medium | **P2** — следующий спринт |
| S4 | max_length/bounds в logistics schemas | Medium | Low | **P2** — следующий спринт |
| S5 | XLSX import row limit | Medium | Low | **P2** — следующий спринт |
| S6 | weight_kg bounds/validation | Medium | Low | **P2** — следующий спринт |
| S7 | PII в exchange_export_dir | Medium | Medium | **P2** — следующий спринт |
| A6–A15 | Persistence, DRY, Strategy, router split… | Medium | Medium–High | **P2** — следующий спринт |
| Q4–Q14 | Magic strings, DRY, tests, types, boilerplate | Medium | Low–Medium | **P2** — следующий спринт |
| A16–A22, S8–S12, Q15–Q18 | Low / hygiene / hardening | Low | Low | Backlog |

---

## Следующие шаги

1. **Немедленно (P0)**
   - Единый ACL статуса КП на create/reuse-transport/patch shipments (**S1**) — `/implement`.
   - Пагинация и max limit на `GET /production/sgp/plates` (**A5**, **S2**) — `/implement`.

2. **Этот спринт (P1)**
   - Разрезать god-сервисы: начать с `SgpRepository` + вынос SQL (**A1**); pass-through и propose Strategy в Shipment (**A2**, **A8**, **A9**) — `/refactor`.
   - Перенести `ShipmentError` в domain (**A4**).
   - DRY SQL: propose candidates (**Q2**), list_positions CTE (**Q3**).
   - FE тесты `ShipmentItemsSection` propose/confirm (**Q1**) — `/implement`.
   - План cross-context invalidation (**A3**, **A10**).

3. **Следующий спринт (P2)**
   - Rate limit (**S3**), schema bounds (**S4**), XLSX row cap (**S5**), weight validation (**S6**).
   - SGP router split (**A13**), PlateStatus вместо magic strings (**A14**, **Q4**).
   - Backend/FE тесты gaps (**Q9**, **Q10**); transaction helper (**Q8**).

4. **Backlog**
   - Low: **A16–A22**, **S8–S12**, **Q15–Q18** — по мере касания модулей.
   - UI split крупных компонентов (**A19**) — `/refactor` при следующем UX-изменении.

---

## Примечание по remediation

Автоматическое исправление **не запускалось**. Код в рамках Phase 4 **не менялся** — только консолидированный отчёт. Для remediation использовать `/refactor` (структурные изменения) и `/implement` (ACL, пагинация, тесты) после явного подтверждения приоритетов.
