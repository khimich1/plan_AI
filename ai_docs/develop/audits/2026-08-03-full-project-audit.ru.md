# Отчёт по аудиту проекта

**Дата**: 2026-08-03  
**Область**: весь проект (`app/`, `core/`, `frontend/src/` — критичные ~20–30%)  
**Аудиторы**: senior-reviewer + security-auditor + reviewer

---

## Краткое резюме

**Общая оценка здоровья**: 2.0/10

| Серьёзность | Архитектура | Безопасность | Качество кода | Итого |
|-------------|-------------|--------------|---------------|-------|
| Critical    | 2           | 0            | 0             | **2** |
| High        | 6           | 3            | 7             | **16** |
| Medium      | 7           | 9            | 11            | **27** |
| Low         | 3           | 5            | 4             | **12** |

**Расчёт Health Score**:

- Старт: **10**
- Critical: 2 × −2 = **−4** (потолок −6)
- High: 16 × −0.5, потолок **−3**
- Medium: 27 × −0.1, потолок **−1**
- Low: не учитываются
- **Итого = 10 − 4 − 3 − 1 = 2.0/10**

**Рекомендация**: устранить 2 критические архитектурные проблемы до следующего релиза; также приоритезировать High-находки по безопасности S1–S3 (особенно S3 — целостность складских списаний).

---

## Критические проблемы (исправить немедленно)

### [A1] ShipmentService — god-модуль: смешение persistence, домена и API-маппинга

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `app/services/shipment_service.py` (~1521 строк, 40+ методов, ~96 прямых обращений к БД) |
| **Влияние** | Один модуль владеет CRUD, propose/confirm, packing, complete/cancel с побочными эффектами СГП, XLSX-экспортом, поиском КП/свай, маппингом схем. Нарушает SRP; подсистемы нельзя тестировать изолированно; любое изменение логистики рискует широкими регрессиями |
| **Исправление** | Разбить на `ShipmentRepository`, `ShipmentProposeService`, `ShipmentCompletionService`, `ShipmentExportService`; тонкий оркестратор; маппинг в Pydantic только в endpoints |

### [A2] Stateful-хранилище рассчитано на single-instance без жёсткого enforcement

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `core/kp_db_common.py` (SQLite+WAL), `app/services/draft_store.py` (черновики в ФС), `core/config/settings.py` (`APP_STORAGE_LAYOUT=single_instance`, только warning ~450–458), in-process rate limits (`app/security/login_rate_limit.py`) |
| **Влияние** | SQLite + файловые черновики + in-process счётчики работают только на одном инстансе. ADR есть, но при старте только предупреждение. Multi-worker / multi-replica → гонки данных, потеря черновиков, обход rate limits |
| **Исправление** | Enforce `UVICORN_WORKERS=1` / число реплик при старте в production; либо перенести черновики и счётчики в общее хранилище до multi-instance |

---

## Проблемы высокого приоритета (исправить скоро)

### Архитектура

#### [A3] CommercialWorkflowService — god-модуль + Service Locator

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `app/services/commercial_workflow_service.py` (~1039 строк); конструктор создаёт 10+ сервисов inline без DI |
| **Влияние** | Скрытая связность, сложно тестировать без полного графа зависимостей, нарушает DIP |
| **Исправление** | Декомпозировать по use-case; constructor injection через `app/dependencies/services.py` |

#### [A4] Application-сервисы возвращают Pydantic HTTP-схемы

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `shipment_service`, `sgp_service`, `carrier_service`, `archive_service` импортируют/возвращают `app/schemas/*` |
| **Влияние** | Транспортный слой (HTTP-схемы) смешан с application/domain; изменение контракта API тянет правки бизнес-логики |
| **Исправление** | Сервисы возвращают доменные модели; маппинг в Pydantic — только в endpoints |

#### [A5] SQL логистики/СГП живёт в сервисах, а не в repositories

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `ShipmentService`, `SgpService`, `CarrierService` используют `_connect` + raw SQL; частичные хелперы в `core/kp_db_shipments.py` |
| **Влияние** | SQL размазан по сервисам; нет единой границы persistence; сложнее тестировать и эволюционировать схему |
| **Исправление** | Вынести repository-слой; консолидировать SQL в модулях `core/kp_db_*` |

#### [A6] Непоследовательная dependency injection

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `app/dependencies/services.py`: только auth инжектит repo; logistics-фабрики собирают `KpRepository().db_path`; `CommercialWorkflow`/`Offers` самоконструируются |
| **Влияние** | Неединообразный wiring; часть сервисов тестируема, часть — нет; скрытые синглтоны |
| **Исправление** | Единый factory-паттерн; все зависимости через FastAPI `Depends` |

#### [A7] Lazy-импорты маскируют связность между сервисами

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `archive_service` runtime-импортирует `SgpService`, `KpReadinessService`, `kp_db_shipments`; аналогично в `plan_storage` |
| **Влияние** | Риск циклов зависимостей скрыт на этапе импорта; связность не видна статическому анализу |
| **Исправление** | Явный DI; разрывать циклы через интерфейсы или event-границы |

#### [A8] Легаси: неявное mutable global state для plate order / optimization

| Поле | Значение |
|------|----------|
| **Категория** | Архитектура |
| **Где** | `core/plate_runtime_state.py`, `core/config_and_data.py`, `core/domain/plate_order.py`; в HTTP смягчено middleware, но bot/scripts могут обойти |
| **Влияние** | Общее mutable-состояние между запросами в bot/scripts; гонки и устаревшие данные |
| **Исправление** | Передавать явный context; убрать module-level mutable state |

### Безопасность

#### [S1] In-process rate limits обходятся несколькими workers

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `login_rate_limit.py:49–86`, `commercial_upload_validation.py:23–45` |
| **Влияние** | Лимиты brute-force и злоупотребления загрузками неэффективны при нескольких workers/репликах |
| **Исправление** | Счётчики в Redis либо жёсткий single-worker деплой |

#### [S2] Commercial OCR отправляет документы во внешний LLM

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `core/ocr/recognition.py`, `providers/openai.py`, `commercial.py:69–81`; `OCR_EXTERNAL_ENABLED` |
| **Влияние** | Коммерческие документы клиентов уходят к стороннему LLM; риски резидентности и конфиденциальности данных |
| **Исправление** | По умолчанию выключено; audit trail; предпочтительно on-prem обработка |

#### [S3] Позиции отгрузки не валидируются относительно КП заказа отгрузки

| Поле | Значение |
|------|----------|
| **Категория** | Безопасность |
| **Где** | `shipment_service.py` `put_items`/`_prepare_item` (~589–713), `complete` (~727–804) |
| **Влияние** | Роль logistics может привязать любой `completed_plate_id` и завершить отгрузку — списание чужого склада/клиента |
| **Исправление** | Требовать `plate.kp_id ∈ shipment order KPs` в `_prepare_item` и `complete` |

### Качество кода

#### [Q1] Дублированный SQL сопоставления плит СГП в unlink/relink/reserve/free

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `sgp_service.py` |
| **Влияние** | Четыре почти одинаковых SQL-блока; правка в одном пути может не попасть в остальные |
| **Исправление** | Вынести общий хелпер `_match_plates` или метод repository |

#### [Q2] Дублированная проверка доступности `_prepare_item` vs `_preflight_availability`

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `shipment_service.py` |
| **Влияние** | Расхождение валидации между путями; тонкие баги при смене правил |
| **Исправление** | Единая функция проверки доступности |

#### [Q3] N+1 обращения к БД при сборке строк списка архива

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `archive_service._to_list_item` |
| **Влияние** | List-endpoint масштабируется как O(n) запросов; медленная страница архива под нагрузкой |
| **Исправление** | Batch-fetch связанных данных; join или prefetch в repository |

#### [Q4] Методы мутаций SgpService — почти дубликаты (~170 строк каждый)

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `sgp_service.py` — `unlink`/`relink`/`reserve_on_conn` |
| **Влияние** | ~170 строк дублирования на метод; высокая стоимость сопровождения и риск drift |
| **Исправление** | Вынести общий каркас транзакции и логику сопоставления плит |

#### [Q5] ShipmentItemsSection — stateful-компонент на 644 строки с хрупкой синхронизацией

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `frontend/src/features/logistics/components/ShipmentItemsSection.tsx` |
| **Влияние** | Сложная синхронизация local/server state; `eslint-disable exhaustive-deps` маскирует stale-closure баги |
| **Исправление** | Разбить на подкомпоненты; выводить state; починить dependency array |

#### [Q6] Тонкое unit-покрытие OffersService

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `tests/test_offers_service.py` — только 2 теста дат PDF/XLSX |
| **Влияние** | Ядро коммерческого workflow почти не покрыто на unit-уровне |
| **Исправление** | Добавить тесты move_to_production, валидации, error-путей |

#### [Q7] Пробелы во фронтенд-тестах: header save/complete/cancel, CarrierAutocomplete, invalidateRelated

| Поле | Значение |
|------|----------|
| **Категория** | Качество кода |
| **Где** | `ShipmentDrawer`, `CarrierAutocomplete`, `useLogisticsQueries.ts` |
| **Влияние** | Критичные пользовательские сценарии и инвалидация кэша не протестированы; вероятны регрессии |
| **Исправление** | Добавить component- и hook-тесты для перечисленных потоков |

---

## Проблемы среднего приоритета (план на следующий спринт)

### Архитектура

- **[A9]** `core/kp_db.py` — god-фасад, реэкспортирующий persistence — сузить surface или явно задокументировать границу.
- **[A10]** Параллельные слои оркестрации планирования — `app/planning/`, `production_planning_service.py`, `core/production/planning.py` — выбрать канонический слой; deprecate дубликаты.
- **[A11]** Frontend god-компонент — `ShipmentItemsSection.tsx` (~617–644 строк) — разбить на row, bulk actions, validation.
- **[A12]** Слишком широкая инвалидация React Query — `useLogisticsQueries.ts` `invalidateRelated` трогает logistics+archive+production+sgp — сузить до затронутых ключей.
- **[A13]** Тонкие re-export сервисы размывают границу app/core — `kp_persistence_service`, `plate_completion_service`, `rest_matching_service` — inline или переименовать как passthrough.
- **[A14]** Склад СГП вложен под production API — `production.py` vs `/logistics` — выровнять владение маршрутами с доменной границей.
- **[A15]** Обогащение списка архива пересекает домены логистики/СГП — `archive_service._to_list_item` — вынести enrichment в assembler или batch-query слой.

### Безопасность

- **[S4]** CSP Report-Only с `unsafe-inline` — `security_headers.py:11–37` — ужесточить CSP; перейти в enforce, когда готово.
- **[S5]** Нет rate limiting на аутентифицированных мутирующих бизнес-API — logistics/production/commercial/archive/offers — добавить per-user/per-endpoint лимиты.
- **[S6]** Чувствительные данные at rest без шифрования — SQLite + `drafts_dir` plaintext — шифровать том или ограничить доступ к ФС.
- **[S7]** Долгая сессия без idle timeout — `session.py` / settings по умолчанию 12ч — idle timeout и sliding expiration.
- **[S8]** CSRF-cookie читаема из JS — `csrf.py` httponly False (нужно для double-submit, но XSS усиливает риск) — минимизировать XSS; рассмотреть альтернативный CSRF-паттерн.
- **[S9]** Динамические имена колонок SQL без allowlist в сервисе — `shipment_service.py:282–286` (на API смягчает Pydantic) — явный allowlist в service-слое.
- **[S10]** Поиск КП в логистике показывает данные клиентов чужих менеджеров — намеренно для logistics? Подтвердить бизнес-правило и задокументировать или ограничить.
- **[S11]** Создание отгрузки не валидирует статус КП — только `_assert_kp_exists` — проверять, что КП в отгружаемом состоянии.
- **[S12]** Нет сканирования Python-зависимостей в CI — `frontend-audit.yml` только npm — добавить `pip-audit` или Dependabot для Python.

### Качество кода

- **[Q8]** Magic number 0.005 (допуск размеров) не централизован — вынести в общую константу domain/packing config.
- **[Q9]** Захардкожено `'в производстве'` вместо `PlateStatus.IN_PRODUCTION.value` — использовать enum.
- **[Q10]** Повторяющийся transaction boilerplate в ShipmentService — вынести `_with_transaction`.
- **[Q11]** `patch()` принимает `actor`, но не сохраняет — неполный audit trail — писать actor при мутации или убрать параметр.
- **[Q12]** Две реализации propose (legacy FIFO vs v2 packing) — deprecate legacy или явно feature-flag.
- **[Q13]** Дублированное progress enrichment в archive-мапперах — общий enrichment-хелпер.
- **[Q14]** Дублирование оркестрации `move_to_production` в OffersService/ArchiveService — один общий use-case сервис.
- **[Q15]** OffersService — нетипизированные строковые коды в `ValueError` — typed exception hierarchy или error enum.
- **[Q16]** XLSX-экспорт встроен в ShipmentService — перенести в `ShipmentExportService` (см. A1).
- **[Q17]** У `reserve_on_conn` нет прямых unit-тестов — изолированные тесты с mocked connection.
- **[Q18]** Дублирование `generate_pdf`/`generate_xlsx` в OffersService — общий хелпер генерации документов.

---

## Низкий приоритет / предложения

### Архитектура

- **[A16]** Двойная иерархия моделей PlateOrder — `core/domain` vs `app/domain` — консолидировать или задокументировать маппинг.
- **[A17]** `plan_manager.py` меняет `sys.path` при импорте — починить структуру импортов; убрать runtime path hack.
- **[A18]** Позитив: `core/shipment_packing/` хорошо ограничен — держать как шаблон для будущих доменных модулей.

### Безопасность

- **[S13]** CORS `allow_headers=["*"]` — `main.py:61–67` — ограничить нужными заголовками.
- **[S14]** `session_version` в auth-ответах — оценить необходимость; убрать, если клиент не использует.
- **[S15]** Нет заголовка Permissions-Policy — добавить ограничительную политику.
- **[S16]** Роль production может видеть список всех менеджеров — подтвердить бизнес-правило; ограничить при отсутствии нужды.
- **[S17]** Legacy web login блокируется без CSRF (мёртвый путь) — удалить dead code.

### Качество кода

- **[Q19]** Module-level mutable counter для React draft keys — `draftItems.ts` — `useRef` или UUID на инстанс.
- **[Q20]** Дублированное форматирование размеров: logistics vs SgpWarehouseView — общий formatter.
- **[Q21]** `type: ignore[arg-type]` на маппинге product_type — `shipment_service.py:1068` — починить типизацию у источника.
- **[Q22]** Пустой неотслеживаемый scratch-файл `_tmp_old.py` — удалить.

---

## Матрица приоритетов

| ID | Проблема | Серьёзность | Effort | Приоритет |
|----|----------|-------------|--------|-----------|
| A1 | ShipmentService god-модуль | Critical | High | P0 |
| A2 | Single-instance storage без enforcement | Critical | High | P0 |
| S3 | Позиции отгрузки не валидируются по КП заказа | High | Medium | P0 |
| S1 | In-process rate limits обходятся | High | Medium | P1 |
| A3 | CommercialWorkflowService god-модуль | High | High | P1 |
| A4 | Сервисы возвращают Pydantic HTTP-схемы | High | Medium | P1 |
| A5 | SQL в сервисах, не в repositories | High | High | P1 |
| Q1 | Дублированный SQL сопоставления плит СГП | High | Medium | P1 |
| Q2 | Дублированная проверка доступности | High | Low | P1 |
| Q3 | N+1 к БД в списке архива | High | Medium | P1 |
| S2 | OCR шлёт документы во внешний LLM | High | Low | P2 |
| A6 | Непоследовательный DI | High | Medium | P2 |
| A7 | Lazy-импорты маскируют связность | High | Medium | P2 |
| A8 | Легаси mutable global plate state | High | Medium | P2 |
| Q4 | Почти-дубликаты мутаций SgpService | High | Medium | P2 |
| Q5 | God-компонент ShipmentItemsSection | High | Medium | P2 |
| Q6 | Тонкие unit-тесты OffersService | High | Medium | P2 |
| Q7 | Пробелы фронтенд-тестов | High | Medium | P2 |
| S4 | CSP Report-Only с unsafe-inline | Medium | Medium | P3 |
| S5 | Нет rate limit на mutating API | Medium | Medium | P3 |
| S11 | Создание отгрузки без проверки статуса КП | Medium | Low | P3 |
| A11 | Frontend god-компонент ShipmentItemsSection | Medium | Medium | P3 |
| A12 | Слишком широкая инвалидация React Query | Medium | Low | P3 |
| Q14 | Дублирование move_to_production | Medium | Medium | P3 |
| A9 | God-фасад kp_db.py | Medium | Medium | P4 |
| A10 | Параллельные слои планирования | Medium | High | P4 |
| A13 | Тонкие re-export сервисы | Medium | Low | P4 |
| A14 | СГП под production API | Medium | Medium | P4 |
| A15 | Enrichment архива пересекает домены | Medium | Medium | P4 |
| S6 | Данные at rest без шифрования | Medium | High | P4 |
| S7 | Долгая сессия без idle timeout | Medium | Low | P4 |
| S8 | CSRF-cookie читаема из JS | Medium | Low | P4 |
| S9 | Динамические имена колонок SQL | Medium | Low | P4 |
| S10 | Поиск КП logistics cross-manager | Medium | Low | P4 |
| S12 | Нет скана Python-зависимостей в CI | Medium | Low | P4 |
| Q8–Q18 | Остальные medium по качеству | Medium | Low–Med | P4 |
| A16–A18 | Architecture low / позитив | Low | Low | P5 |
| S13–S17 | Security low | Low | Low | P5 |
| Q19–Q22 | Code quality low | Low | Low | P5 |

---

## Следующие шаги

1. **Немедленно** (до следующего коммита): критические фиксы — **A1** (начать декомпозицию), **A2** (startup guard для single-instance)
2. **Этот спринт**: high, особенно **S3** (целостность склада), **S1** (rate limits), **A3–A5** (границы service/repository)
3. **Следующий спринт**: medium — hardening безопасности (S4–S12), декомпозиция фронтенда (A11, Q5), покрытие тестами (Q6, Q7)
4. **Бэклог**: low — cleanup, заголовки, dead code, дедуп форматирования

Для структурных проблем: `/refactor [file]`.  
Для feature-level security-фиксов: `/implement [fix]`.

---

## Связанные документы

- Предыдущий аудит: [2026-08-02-full-project-audit.md](./2026-08-02-full-project-audit.md)
- Сравнение аудитов: [2026-08-02-audit-comparison.md](./2026-08-02-audit-comparison.md)
- ADR по деплою: [deployment-single-instance.md](../architecture/deployment-single-instance.md)
- Спеки стабилизации: [stabilizaciya-p0-p1-audit-2026-08-02.md](../../specs/stabilizaciya-p0-p1-audit-2026-08-02.md)
- Английская версия этого отчёта: [2026-08-03-full-project-audit.md](./2026-08-03-full-project-audit.md)
