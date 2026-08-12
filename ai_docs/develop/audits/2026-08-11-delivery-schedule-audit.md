# Отчёт аудита: График поставки

**Дата**: 2026-08-11  
**Область**: модуль delivery-schedule (backend + frontend feature)  
**Аудиторы**: senior-reviewer + security-auditor + reviewer  
**Спека**: [`docs/specs/delivery-schedule.md`](../../../docs/specs/delivery-schedule.md)

---

## Краткое резюме

**Оценка здоровья**: **6.0 / 10**

Расчёт: старт 10; Critical 0 → 0; High 13 × 0.5 (потолок −3); Medium 29 × 0.1 (потолок −1) → **6.0**

| Серьёзность | Архитектура | Безопасность | Качество кода | Итого |
|-------------|-------------|--------------|---------------|-------|
| Critical    | 0           | 0            | 0             | **0** |
| High        | 5 (A1–A5)   | 2 (S1–S2)    | 6 (Q1–Q6)     | **13** |
| Medium      | 9 (A6–A14)  | 5 (S3–S7)    | 15 (Q7–Q21)   | **29** |
| Low         | 5 (A15–A19) | 5 (S8–S12)   | 7 (Q22–Q28)   | **17** |

**Рекомендация**: критических блокеров нет. До прод-нагрузки и human QA — закрыть **P0 AuthZ** (IDOR на GET/import), затем read-only UX для архивных КП, унификацию констант ёмкости и маппинг «произведено».

MVP «График поставки» (Week 1–3) реализован в духе проекта: тонкий FastAPI-роутер, оркестрация в сервисе, чистый светофор в `core/delivery_schedule_check.py`, изолированная frontend-feature. Главный риск — **broken access control**: read-операции не проверяют владельца КП, тогда как PUT и `/document` — проверяют. Вторичный технический долг — монолитный сервис (~708 строк), N+1 и множественные conn на GET, тихая деградация светофора, неполная интеграция с архивом (бейдж, kp_files).

**Метрики области**: `delivery_schedule_service.py` (~708 строк), `delivery_schedule_check.py` (~329 строк), 6 pytest-модулей (~60 тестов), 2 vitest-файла, frontend feature (~12 файлов).

---

## Критические проблемы

Нет.

---

## Высокий приоритет

### P0 AuthZ: IDOR на GET и import [A1] [S1] [S2]

**Категория**: Архитектура + Безопасность (одна корневая проблема)  
**Где**: `app/api/v1/endpoints/delivery_schedule.py` (GET, POST `/import`), `app/services/delivery_schedule_service.py` (`get`, `import_draft`)  
**Влияние**: роутер получает `user`, но в сервис для read-операций не передаёт; `assert_offer_read_access` вызывается только в `generate_document`, `replace` проверяет write-доступ. Manager с валидной сессией может по `kp_id` читать чужой график (счёт, договор, партии, `plate_id`, марки, qty, светофор) и импортировать XLSX — в ответе `plate_id`/`plate_name`, через `unmatched_rows[].raw` — содержимое ячеек, сопоставленное с позициями чужого КП. Разрыв с `archive_service` и PUT/`/document`.  
**Исправление**: прокинуть `user` в `get`/`import_draft`, загружать offer, вызывать `assert_offer_read_access`; regression-тест «чужой manager → 403» по образцу `test_archive_authorization.py`.

---

### [A2] / [Q6] God-service `DeliveryScheduleService` (~708 строк)

**Категория**: Архитектура + Качество  
**Где**: `app/services/delivery_schedule_service.py`  
**Влияние**: в одном классе — raw SQL CRUD, валидация, светофор, импорт XLSX, генерация документов, календарь, readiness-mapping. Высокая связность, сложно менять одну ответственность без регрессий.  
**Исправление**: выделить `DeliveryScheduleRepository` (SQL), `DeliveryScheduleTrafficLightEnricher` (occupancy + readiness + `check_batches`); в сервисе — оркестрация use-case.

---

### [A3] Read-only просмотр архивных КП не реализован (спека)

**Категория**: Архитектура  
**Где**: `frontend/.../OfferDetailsDrawer.tsx` (`canEditDeliverySchedule`, кнопка `disabled`); backend `replace` корректно отклоняет (`ALLOWED_EDIT_KP_STATUSES`), GET доступен  
**Влияние**: по спеке график у архивного/завершённого КП должен оставаться **read-only для просмотра** (протокол обещанных дат); сейчас кнопка полностью disabled — менеджер не может открыть сохранённый график и светофор.  
**Исправление**: кнопку enable для не-simple КП; `readOnly={!canEdit}` уже есть в `DeliveryScheduleDialog` — убрать `disabled` с кнопки.

---

### [A4] / [Q20] Drift констант ёмкости производства

**Категория**: Архитектура + Качество  
**Где**: `core/production_capacity.py` (`TRACKS_PER_DAY_DEFAULT`), `app/planning/plan_storage.py` (`MAX_TRACKS_PER_DAY`), `core/delivery_schedule_check.py` (fallback R2, `capacity_buffer=1.15`, `green_slack_workdays=5`) vs `plan_calendar.get_global_calendar_info()`  
**Влияние**: спека предполагала единый источник в `production_capacity.py`; при изменении только одной константы светофор и план расходятся; R2-fallback даёт 5 дорожек из `production_capacity`, `occupied` — из другого контура.  
**Исправление**: единый модуль констант; `plan_calendar` импортирует оттуда; в `check_batches` для дат внутри горizonta — `days_info[].max`.

---

### [A5] Маппинг «произведено» по identity → ложный green

**Категория**: Архитектура  
**Где**: `delivery_schedule_service._load_produced_by_plate_id` + агрегация `KpReadinessService.list_positions` по `(plate_name, length, width, load_class)`  
**Влияние**: при нескольких строках `kp_plates` с одной identity каждый `plate_id` получает полный `on_sgp` группы → завышенное вычитание → занижение остатка → **ложный green** (обратная сторона «безопасной» ошибки по складу из спеки).  
**Исправление**: распределять `on_sgp` пропорционально `qty` или считать produced на уровне `plate_id` через SQL к `completed_plates`.

---

### [Q1] Широкий `except Exception` в светофоре

**Категория**: Качество  
**Где**: `delivery_schedule_service._enrich_with_traffic_light` (~329–334)  
**Влияние**: ловит любое исключение, отдаёт view без `status`/`hint`; баги программирования выглядят как «светофор недоступен», без сигнала для мониторинга.  
**Исправление**: ловить только ожидаемые сбои (readiness, calendar); прочие — пробрасывать или логировать как error.

---

### [Q4] Импорт XLSX: конфликт дат в одной партии игнорируется

**Категория**: Качество  
**Где**: `core/delivery_schedule_xlsx.py` (~315–327)  
**Влияние**: строки с одинаковым именем партии группируются; при разных датах сохраняются даты **первой** строки, позиции мёржатся — тихая порча данных.  
**Исправление**: расхождение → `UnmatchedRow` с reason `conflicting batch dates` или warning в ответе import.

---

### [Q5] Не реализовано правило `produce_by ≤ deliver_from` (спека)

**Категория**: Качество  
**Где**: `app/schemas/delivery_schedule.py`, `frontend/.../scheduleDraft.ts`, спека §215–216  
**Влияние**: спека требует предупреждение (не блокирующую ошибку), если «произвести до» позже «поставка с»; проверяется только `deliver_from ≤ deliver_to`.  
**Исправление**: soft-warning в API (`warnings[]`) и/или жёлтый Alert в редакторе.

---

## Средний приоритет

### Архитектура

| ID | Проблема | Исправление (кратко) |
|----|----------|----------------------|
| **[A6]** | Документы не регистрируются в `kp_files` / архиве КП | После генерации — запись в `kp_files` или helper из `offers_write.py` |
| **[A7]** | Бейдж «есть график» в списке архива не работает (`has_delivery_schedule` не заполняется) | LEFT JOIN / EXISTS на `delivery_schedule` в list query |
| **[A8]** / **[Q2]** | N+1 при `_build_view` — отдельный SELECT items на каждую партию | Один JOIN + группировка в памяти |
| **[A9]** / **[Q3]** | Множественные conn на один GET (view, plates_meta, readiness) | Unit-of-work / один conn на request |
| **[A10]** / **[S9]** | Тихая деградация светофора без сигнала пользователю | `traffic_light_degraded: bool` / warning в ответе и UI |
| **[A11]** | `get_global_calendar_info()` на каждый GET — чтение планов с диска | Кэш TTL / `OccupancySource` |
| **[A12]** | Поле `delivery_schedule.status` не используется (всегда `draft`) | Wire lifecycle или убрать из контракта |
| **[A13]** | Импорт: дубликаты марок (`plates_by_name` — первая позиция) | `unmatched` с reason `duplicate mark` |
| **[A14]** | Нет предупреждения `produce_by > deliver_from` (см. Q5) | `warnings[]` или UI-hint |

### Безопасность

| ID | Проблема | Исправление (кратко) |
|----|----------|----------------------|
| **[S3]** | Upload XLSX без magic bytes / защиты от zip bomb | Проверка сигнатуры ZIP; лимит строк/листов при parse |
| **[S4]** | Нет верхних границ на PUT-payload (DoS / переполнение БД) | `Field(max_length=…)` на строки, партии, items |
| **[S5]** | Formula injection в генерируемых XLSX | Префикс `'` для `=`, `+`, `-`, `@` или явный string-тип |
| **[S6]** | Утечка деталей валидации в HTTP (`plate_id=123 не принадлежит…`) | Generic-сообщения наружу; детали в лог (особенно до закрытия AuthZ) |
| **[S7]** | Перечисление наличия графика по kp_id (404 vs 200 без ownership) | После AuthZ: чужой КП → 403, не 404 |

### Качество кода

| ID | Проблема |
|----|----------|
| **[Q7]** | DRY: три почти одинаковых SELECT из `kp_plates` |
| **[Q8]** | DRY: копипаст test-fixtures (`_fresh_db`, `_seed_kp`) |
| **[Q9]** | Dead code: неиспользуемый `logger` в pdf/xlsx |
| **[Q10]** | Вводящее имя `iter_document_table_rows` (возвращает list) |
| **[Q11]** | Устаревшие комментарии с номерами задач T4/T5 в schemas |
| **[Q12]** / **[A18]** | Machine-readable `reason` (`unknown mark`, `bad date`) без локализации |
| **[Q13]** | Test gap: `import_draft` без интеграционного теста сервиса |
| **[Q14]** | Test gap: деградация светофора при сбое calendar/readiness |
| **[Q15]** | Test gap: yellow/red через сервисный слой (только green в service-тестах) |
| **[Q16]** | Test gap: frontend hooks и диалоги (нет MSW/react-query tests) |
| **[Q17]** | Слабая типизация «сырых» dict/Any (occupancy, schedule view) |
| **[Q18]** | Мутабельный merge в dataclass `_merge_item` |
| **[Q19]** | Избыточный re-raise `except RuntimeError: raise` |
| **[Q21]** | Неверная зависимость тестов lib → component (`scheduleDraft.test.ts`) |

---

## Низкий приоритет

### Архитектура

- **[A15]** UI-масштабируемость: `BatchCard` рендерит все позиции КП в каждой партии (31×21 → тяжёлый DOM).
- **[A16]** Светофор только после сохранения — импорт даёт drafts с `status: null` до PUT.
- **[A17]** `KpReadinessService` создаётся inline — без DI.
- **[A18]** Причины unmatched на английском (см. Q12).
- **[A19]** PDF зависит от XLSX-модуля для layout — coupling документов.

### Безопасность

- **[S8]** XSS в UI — **не обнаружен** (React text nodes, PDF `escape()`).
- **[S9]** Silent degrade светофора — integrity/UX (см. A10).
- **[S10]** GET `/template` не проверяет доступ к `kp_id` (шаблон пустой — риск минимален).
- **[S11]** Frontend полагается на UI-gating, не на AuthZ API (defense-in-depth после P0).
- **[S12]** Нет rate limiting на import/document.

### Качество

- **[Q22]** Дублирующее обновление кэша после PUT (`setQueryData` + `invalidateQueries`).
- **[Q23]** Дублирование inline-стилей полей ввода (`Editor` / `BatchCard`).
- **[Q24]** `plateById` пересоздаётся на каждый render в `BatchCard`.
- **[Q25]** Избыточная проверка `sort_order is not None` (Pydantic default 0).
- **[Q26]** Test gap: merge qty одной марки при import.
- **[Q27]** TS `any` в feature — **не найдено** ✅.
- **[Q28]** PDF-тест проверяет только `%PDF`, не содержимое.

---

## Матрица приоритетов

| ID | Проблема | Серьёзность | Усилия | Приоритет |
|----|----------|-------------|--------|-----------|
| **A1, S1, S2** | AuthZ IDOR — GET/import без `assert_offer_read_access` | High | Низкие | **P0 — немедленно** |
| A3 | Read-only просмотр архивных КП (кнопка disabled) | High | Низкие | **P1 — этот спринт** |
| A4, Q20 | Drift констант ёмкости / buffer / slack | High | Средние | **P1 — этот спринт** |
| A5 | Produced mapping по identity → ложный green | High | Средние | **P1 — этот спринт** |
| Q1, A10 | Broad except / silent degrade светофора | High | Низкие | **P1 — этот спринт** |
| Q4, Q5, A13, A14 | Import conflicts, `produce_by` warning | High/Medium | Низкие–средние | **P1 — этот спринт** |
| A2, Q6 | God-service ~708 строк | High | Высокие | **P2 — следующий спринт (эпик)** |
| A8, A9, Q2, Q3 | N+1 и множественные conn | Medium/High | Средние | **P2** |
| A6, A7 | kp_files, бейдж в списке архива | Medium | Средние | **P2** |
| S3–S7 | Upload hardening, payload limits, formula injection, error leak | Medium | Низкие–средние | **P2** |
| Q7–Q21 | DRY, test gaps, TypedDict, naming | Medium | Разные | **P3 — backlog** |
| A15–A19, S8–S12, Q22–Q28 | Low | Low | Разные | **P4 — backlog** |

---

## Следующие шаги

1. **Немедленно**:
   - [A1], [S1], [S2] — `assert_offer_read_access` на GET и import; прокинуть `user`; regression-тест 403 для чужого КП;
   - [S7] — унифицировать 403 vs 404 после закрытия AuthZ.

2. **Этот спринт**:
   - [A3] — read-only просмотр графика для архивных КП (убрать `disabled` с кнопки);
   - [A4], [Q20] — единый источник констант ёмкости;
   - [A5] — produced mapping по `plate_id`, не identity;
   - [Q1], [A10], [S9] — сузить except, поле `traffic_light_degraded`;
   - [Q4], [Q5], [A13], [A14] — import conflicts и warning `produce_by ≤ deliver_from`;
   - [S3], [S4], [S5] — upload magic, payload limits, formula injection в XLSX.

3. **Следующий спринт**:
   - [A2], [Q6] — декомпозиция сервиса (repository + enricher);
   - [A8], [A9], [Q2], [Q3] — один JOIN, unit-of-work;
   - [A6], [A7] — kp_files и `has_delivery_schedule` в list API;
   - [Q13]–[Q16] — интеграционные и frontend-тесты.

4. **Backlog**: Medium [A11]–[A12], [S6]; Low [A15]–[A19], [S10]–[S12], [Q7]–[Q12], [Q17]–[Q28]; CP2 калибровка светофора ±20% (human QA).

---

## Что сделано хорошо

- **Слои и направление зависимостей**: тонкий роутер → сервис → `core/*`; `delivery_schedule_check.py` без импортов `app.*` — чистая доменная математика светофора.
- **Чистый `check_batches`**: детерминированный алгоритм, хорошо покрыт unit-тестами (green/yellow/red, workdays, occupancy fallback).
- **Тесты check/xlsx/schema**: ~60 pytest-тестов; парсинг шаблона, инварианты qty, генерация документов.
- **Frontend без `any`**: типизация синхронна с Pydantic; feature-структура `api/ components/ hooks/ types/`.
- **PDF escape**: все динамические поля через `xml.sax.saxutils.escape` — XSS/injection в PDF закрыт.
- **Write AuthZ**: `assert_offer_write_access` в PUT; read AuthZ в `/document`.
- **SQL**: параметризованные запросы; path traversal в имени документа исключён (server-side stem).
- **Спека Not Doing соблюдена**: нет версий, LLM, резервирования дорожек — прагматичный MVP.

---

## Связанные документы

- Спека: [`docs/specs/delivery-schedule.md`](../../../docs/specs/delivery-schedule.md)
- Feature doc: [`ai_docs/develop/features/delivery-schedule.md`](../features/delivery-schedule.md)
- Отчёты реализации: [week1](../reports/delivery-schedule-week1.md), [week2](../reports/delivery-schedule-week2.md), [week3](../reports/delivery-schedule-week3.md)
- План: [`ai_docs/develop/plans/delivery-schedule-plan.md`](../plans/delivery-schedule-plan.md)

---

## Remediation

**Дата remediation**: 2026-08-11  
**Спека**: [`docs/specs/delivery-schedule-audit-remediation.md`](../../../docs/specs/delivery-schedule-audit-remediation.md)  
**План**: [`ai_docs/develop/plans/delivery-schedule-audit-remediation-plan.md`](../plans/delivery-schedule-audit-remediation-plan.md)

### Fixed (P0 + «скоро»)

| ID | Что сделано |
|----|-------------|
| **A1, S1, S2, S7** | AuthZ: `assert_offer_read_access` на GET и POST `/import`; чужой manager → **403**, свой без графика → **404**, нет КП → **404**; HTTPException не маскируется в 500 |
| **A3** | Кнопка «График поставки» в drawer не disabled из-за статуса; `readOnly={!canEdit}`; title просмотр vs редактирование |
| **A5** | Produced: `on_sgp` по identity распределяется пропорционально qty (largest remainder); без ложного green от дублирования |
| **Q1, A10, S9** | Узкий except источников светофора; поле `traffic_light_degraded`; statuses null при degrade; Alert в Editor |
| **Q4** | Импорт: конфликт дат у одного имени партии → unmatched `conflicting batch dates` |

### Remaining (вне scope этой remediation)

- **A2 / Q6** — god-service refactor  
- **A4 / Q20** — drift констант ёмкости  
- **Q5 / A14** — soft-warning `produce_by ≤ deliver_from`  
- **A6 / A7** — kp_files, бейдж списка  
- **A8 / A9 / Q2 / Q3** — N+1, multi-conn  
- **S3–S6, S10–S12** — upload hardening, payload limits, formula injection, template AuthZ, rate limit  
- **A11–A19, Q7–Q28** — backlog качества / UX
