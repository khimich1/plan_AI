# Отчёт аудита проекта

**Дата**: 2026-08-13  
**Область**: два последних блока кода — (1) Multi-nomenclature KP append loop (`6907428`); (2) Delivery schedule MVP + audit remediation (`e9457a4` + `300b522`)  
**Аудиторы**: senior-reviewer + security-auditor + reviewer

---

## Краткое резюме

**Оценка здоровья**: **2.0 / 10**

Расчёт: старт 10; Critical −2 каждый (макс. −6); High −0.5 каждый (макс. −3); Medium −0.1 каждый (макс. −1).

| Серьёзность | Архитектура | Безопасность | Качество кода | Итого |
|-------------|-------------|--------------|---------------|-------|
| Critical    | 1           | 1            | 0             | **2** |
| High        | 4           | 3            | 2             | **9** |
| Medium      | 6           | 7            | 9             | **22** |
| Low         | 4           | 3            | 6             | **13** |

**Рекомендация**: Закрыть 2 критические проблемы (общая корневая причина: AI edit paths стирают mixed drafts) до следующего релиза.

**Сильные стороны (кратко):** Блок 1 — `_compose_order_data_for_product_update`, стабильный `line_id`, thin append APIs, обширные тесты. Блок 2 — чистый домен в `core/delivery_schedule_check.py`, тонкие роутеры, graceful degradation через `traffic_light_degraded`.

**Положительные находки (безопасность):** SQL injection не обнаружен (параметризованные запросы + allowlist таблиц); нет захардкоженных секретов; path traversal при download коммерческих файлов mitigated; нет dangerouslySetInnerHTML в scoped frontend; draft ID traversal guarded.

---

## Критические проблемы (исправить немедленно)

### [A1] / [S1] AI edit paths bypass append compose и стирают mixed drafts

**Категория**: Архитектура + Безопасность (одна корневая причина)  
**Где**: `app/services/commercial_workflow_service.py` — `apply_ai_piles_instruction`, `apply_ai_marches_instruction`, `apply_ai_bridge_piles_instruction`, `apply_ai_fbs_instruction`, `apply_ai_steps_instruction`, `apply_ai_plates_instruction`  
**Проблема**: Методы присваивают `order_data=preview.order_data` напрямую (например, строки 888–893), тогда как text/OCR `update_draft_*` корректно проходят через `_compose_order_data_for_product_update`. При AI-редактировании текущего типа номенклатуры в multi-nomenclature draft все строки других типов молча удаляются.  
**Влияние**: Потеря данных в ключевом новом workflow; может быть сохранена через `save_offer`.  
**Исправление**: Прокинуть каждый `apply_ai_*` через тот же compose/seal pipeline, что `update_draft_piles`; добавить regression-тесты по образцу `tests/test_commercial_multi_append_flow.py`.

---

## Высокий приоритет (исправить в ближайшее время)

### [A2] God module `CommercialWorkflowService`

**Категория**: Архитектура  
**Где**: `app/services/commercial_workflow_service.py` (~2784 строк, ~63 методов)  
**Влияние**: Высокая связность, сложно менять одну ответственность без регрессий.  
**Исправление**: Выделить per-product handlers (Strategy) за тонким facade.

---

### [A3] `append_batches` — draft-only; export layout ломается после DB round-trip

**Категория**: Архитектура  
**Где**: `core/kp_persistence_service.py` — `append_batches` не персистится; archive regen читает `raw.get("append_batches")`, который из БД всегда None  
**Влияние**: Same-type multi-append теряет layout при экспорте после сохранения.  
**Исправление**: Персистить `append_batches` в `kp_meta` или выводить из durable `append_batch_id` на line rows.

---

### [A4] Delivery schedule без repository layer

**Категория**: Архитектура  
**Где**: `app/services/delivery_schedule_service.py` (~764 строк) — сервис содержит raw SQL  
**Исправление**: Ввести `DeliveryScheduleRepository`.

---

### [A5] Framework exceptions просачиваются в service layer

**Категория**: Архитектура  
**Где**: `assert_offer_read_access` / `assert_offer_write_access` поднимают FastAPI `HTTPException`  
**Исправление**: Domain errors → маппинг в HTTP только в endpoints.

---

### [S2] Text-only AI endpoints без rate limiting

**Категория**: Безопасность  
**Где**: `app/api/v1/endpoints/commercial.py` — `apply_ai_*_to_draft`; OCR имеет per-user limiting, text AI — нет  
**Исправление**: Общий per-user rate limiter.

---

### [S3] AI instruction input без верхней границы

**Категория**: Безопасность  
**Где**: `instruction: str = Form(...)`; workflow проверяет только `len >= 3`  
**Исправление**: `max_length` 2–4 KB на HTTP boundary.

---

### [S4] Manager role — глобальный write/delete на все КП

**Категория**: Безопасность  
**Где**: `app/security/offer_access.py` — `can_write_offer` true для любого manager независимо от `owner_user_id`  
**Влияние**: Insider IDOR.  
**Исправление**: Enforce `owner_user_id` unless product intent — shared-manager pool (документировать, если intentional).

---

### [Q1] Resumed legacy KPs без `line_id` ломают append sealing

**Категория**: Качество кода  
**Где**: `_seal_unbatched_lines` пропускает пустой `line_id`; hydrate не stamp identities  
**Исправление**: Stamp при hydrate или mint IDs перед sealing.

---

### [Q2] God component `CommercialOfferWizard`

**Категория**: Качество кода  
**Где**: `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx` (~1 200 строк)  
**Исправление**: Декомпозиция на подкомпоненты / feature slices.

---

## Средний приоритет (запланировать на следующий спринт)

### Архитектура

| ID | Проблема | Исправление (кратко) |
|----|----------|----------------------|
| **[A6]** | `KpPersistenceService` — растущий persistence monolith (~888 строк) | Разбить по доменным зонам или repository layer |
| **[A7]** | Нет dependency injection в `CommercialWorkflowService` — 15+ collaborators в `__init__` | DI container / factory |
| **[A8]** | Archive hydrate отбрасывает plate optimization context — `hydrate_draft_from_saved_kp` сохраняет `PlateOrder()` и пустой `OptimizationContext` | Восстанавливать контекст из persisted meta |
| **[A9]** | Множественные SQLite connections на delivery-schedule GET | Unit-of-work / один conn на request |
| **[A10]** | Traffic light пересчитывается на каждый GET без кэширования | TTL cache / memoization |
| **[A11]** | Full batch DELETE+INSERT на PUT — `_replace_batches` удаляет все `delivery_batch` rows и re-insert | Upsert / diff-based update |

### Безопасность

| ID | Проблема | Исправление (кратко) |
|----|----------|----------------------|
| **[S5]** | Versioned document generation заполняет `outputs_dir` без pruning (disk exhaustion) | Lifecycle policy / TTL cleanup |
| **[S6]** | Delivery schedule PUT — full DELETE+INSERT без concurrency control (last-write-wins) | Optimistic locking / ETag |
| **[S7]** | XLSX import парсит unbounded rows в рамках 50 MB size cap | Лимит строк/листов при parse |
| **[S8]** | Error responses раскрывают existence details — разные сообщения missing KP vs missing schedule | Generic messages наружу |
| **[S9]** | Unbounded string fields в delivery schedule API (`BatchIn.name`, invoice/contract numbers) | `Field(max_length=…)` |
| **[S10]** | LLM output treated as trusted business input (prompt injection / integrity) | Sanitize / validate LLM output |
| **[S11]** | Per-user OCR rate limiter in-process only (per worker) | Shared limiter (Redis / DB) |

### Качество кода

| ID | Проблема |
|----|----------|
| **[Q3]** | Шесть near-copy product input steps (~500 строк каждый) |
| **[Q4]** | `useCommercialOfferWizard` mutation factory sprawl (~200 строк copy-paste) |
| **[Q5]** | Дублированные preview-row builders across product types |
| **[Q6]** | `FileNotFoundError` используется для domain «not found» |
| **[Q7]** | `undo_last_append_batch` оставляет stale `metadata.product_type` |
| **[Q8]** | Delivery-schedule UI без component/integration tests (только `scheduleDraft.ts` и `BatchStatusChip.test.tsx`) |
| **[Q9]** | Дублированный KP plate-loading SQL в delivery schedule service (3× SELECT) |
| **[Q10]** | Client date validation coupled to UI и incomplete vs server |
| **[Q11]** | Дублированные `order_data` builders в `kp_order_data.py` |

---

## Низкий приоритет / рекомендации

### Архитектура

| ID | Проблема |
|----|----------|
| **[A12]** | Дублированные per-product update handlers в workflow service (6×) |
| **[A13]** | `generate_document` пишет в shared `outputs_dir` без lifecycle policy |
| **[A14]** | Remediation commit `300b522` включает unrelated GSM scripts (`scripts/build_gsm_routes_map.py`, `scripts/build_gsm_trip_pool.py`) |
| **[A15]** | `WizardProgress` step labels не адаптируются к skipped client step |

### Безопасность

| ID | Проблема |
|----|----------|
| **[S12]** | Admin bypasses draft ownership — acceptable for support; consider audit log |
| **[S13]** | `last_ai_instruction` persisted в draft metadata и returned в API |
| **[S14]** | Delivery schedule import lacks XLSX magic-byte validation |

### Качество кода

| ID | Проблема |
|----|----------|
| **[Q12]** | Outdated schema comment about T4/T5 traffic light |
| **[Q13]** | Delivery-schedule feature uses inline styles throughout |
| **[Q14]** | `CommercialCalculationService.is_*_draft` boilerplate |
| **[Q15]** | `BatchCard` rebuilds unused `plateById` lookup map |
| **[Q16]** | `eslint-disable` masks hook dependency gap in `DeliveryScheduleDialog.tsx` |
| **[Q17]** | No cross-field date rule for `produce_by` vs delivery window |

---

## Матрица приоритетов

| ID | Проблема | Серьёзность | Effort | Priority |
|----|----------|-------------|--------|----------|
| A1/S1 | AI edit paths стирают mixed drafts | Critical | M | **P0** |
| A3 | `append_batches` не персистится | High | M | **P1** |
| S2 | Text AI без rate limiting | High | S | **P1** |
| S3 | AI instruction unbounded | High | S | **P1** |
| S4 | Manager global write на все КП | High | M | **P1** |
| Q1 | Legacy KPs без `line_id` ломают sealing | High | M | **P1** |
| A2 | God module CommercialWorkflowService | High | L | **P2** |
| A4 | Delivery schedule без repository | High | M | **P2** |
| A5 | HTTPException в service layer | High | M | **P2** |
| Q2 | God component CommercialOfferWizard | High | L | **P2** |
| A6 | KpPersistenceService monolith | Medium | L | **P2** |
| A7 | Нет DI в CommercialWorkflowService | Medium | M | **P2** |
| A8 | Archive hydrate теряет optimization context | Medium | M | **P2** |
| A9 | Multiple SQLite conn на GET | Medium | S | **P3** |
| A10 | Traffic light без кэша | Medium | S | **P3** |
| A11 | Full DELETE+INSERT на PUT | Medium | M | **P2** |
| S5 | `outputs_dir` без pruning | Medium | S | **P3** |
| S6 | PUT без concurrency control | Medium | M | **P2** |
| S7 | XLSX unbounded rows | Medium | S | **P3** |
| S8 | Error responses leak existence | Medium | S | **P3** |
| S9 | Unbounded string fields | Medium | S | **P3** |
| S10 | LLM output as trusted input | Medium | M | **P2** |
| S11 | OCR rate limiter in-process | Medium | M | **P3** |
| Q3 | Six near-copy product steps | Medium | L | **P3** |
| Q4 | Mutation factory sprawl | Medium | M | **P3** |
| Q5 | Duplicated preview-row builders | Medium | M | **P3** |
| Q6 | FileNotFoundError for domain | Medium | S | **P3** |
| Q7 | Stale `metadata.product_type` after undo | Medium | S | **P3** |
| Q8 | Delivery-schedule UI test gap | Medium | M | **P3** |
| Q9 | Duplicated plate-loading SQL | Medium | S | **P3** |
| Q10 | Client date validation incomplete | Medium | S | **P3** |
| Q11 | Duplicated order_data builders | Medium | M | **P3** |
| A12–A15 | Low architecture items | Low | S–M | **P4** |
| S12–S14 | Low security items | Low | S | **P4** |
| Q12–Q17 | Low code quality items | Low | S | **P4** |

---

## Следующие шаги

1. **Немедленно (P0):** Прокинуть все `apply_ai_*` через `_compose_order_data_for_product_update` / seal pipeline; regression-тесты multi-nomenclature AI edit.
2. **Этот спринт (P1):** Персистить `append_batches`; rate limit + max_length на text AI; enforce owner на manager write; stamp `line_id` при hydrate legacy KPs.
3. **Следующий спринт (P2):** Repository для delivery schedule; domain errors вместо HTTPException в services; concurrency на PUT; LLM output validation; декомпозиция god modules.
4. **Backlog (P3–P4):** DI, кэш светофора, test coverage UI, DRY product steps, low-priority hygiene.

Для структурных проблем: `/refactor [file]`  
Для feature-level security fixes: `/implement [fix]`
