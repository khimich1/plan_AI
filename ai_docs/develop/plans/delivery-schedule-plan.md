# Plan: График поставки (delivery schedule)

Дата: 2026-08-07. Спека: `docs/specs/delivery-schedule.md` (принята).
Статус: план принят, задачи разбиты (ниже), реализация не начата.

## Компоненты и зависимости

```
C1. Схема БД (3 таблицы в core/kp_db_schema.py)
C2. core/production_capacity.py — константы 101 м / 5 дор. (мини-рефакторинг
    archive_service: импорт оттуда; поведение не меняется)
C3. core/delivery_schedule_check.py — светофор (чистая логика)
        deps: C2
        uses: app/planning/plan_calendar.get_global_calendar_info (days_info),
              core/work_calendar.py, KpReadinessService.list_positions (остатки)
C4. core/delivery_schedule_xlsx.py — шаблон: build / parse / документ-XLSX
        deps: — (openpyxl уже в проекте)
C5. PDF документа (reportlab, по образцу core/commercial_offer.py)
        deps: модель данных из спеки (C1)
C6. app: schemas + delivery_schedule_service + router + регистрация
        deps: C1, C3, C4, C5
C7. Frontend feature delivery-schedule + кнопка в OfferDetailsDrawer + бейдж
        deps: контракт C6 (стабилен после спеки → можно параллельно)
C8. Тесты: идут внутри каждого компонента, не отдельной фазой
```

## Порядок (3 недели, 1 разработчик)

**Неделя 1 — ядро (C1, C2, C3).**
Схема + миграция; константы; движок светофора с детерминированными тестами
(зелёный/жёлтый/красный, остаток партии, выходные, пересечение партий).
Почему первым: единственный компонент с алгоритмическим риском.

**Неделя 2 — API и импорт (C6-бэк, C4).**
PUT/GET/import/template эндпоинты; валидация инварианта Σ ≤ qty;
round-trip тесты шаблона. C4 параллелен C3 — при двух руках идут рядом.

**Неделя 3 — документы и фронт (C5, C7, E2E).**
XLSX/PDF документ с шапкой договора; файлы в outputs с датой редакции;
редактор партий, светофор-чипы, dropzone импорта; E2E-сценарий менеджера.

## Параллельность

- C3 ‖ C4 (разные файлы, общий интерфейс — pydantic-схемы из спеки).
- C7 стартует с начала недели 2 по зафиксированным контрактам, бэк
  подменяется моками MSW-стилем (фронт уже так тестируется).
- Последовательны только: C1 → C6 → ручная проверка API → E2E.

## Риски и митигации

| Риск | Вероятность/влияние | Митигация |
|---|---|---|
| R1. Оценка ёмкости расходится с фактом (коэф. 1,15 усредняет переналадки) | Средняя / высокое | Валидация на 3–5 прошлых заказах (CP2) до выкатки; пороги и коэф. — константы модуля; ошибка в безопасную сторону (занижаем «зелёный») |
| R2. days_info не покрывает далёкие даты (график на полгода вперёд, планы — на месяц) | Средняя / среднее | Семантика по умолчанию: дата вне планов = occupied 0, max = default 5. Проверить фактическое поведение `get_global_calendar_info` в задаче C3, зафиксировать тестом |
| R3. Марки из файла клиента не матчатся с plate_name | Высокая / низкое | Импорт → черновик + список несматченных, ничего не сохраняем автоматом; точное совпадение — в MVP, нормализация имён — фаза 2 |
| R4. Позиции КП изменились после создания графика (Σ > qty) | Средняя / низкое | GET не падает: позиция помечается «изменилось количество», партия пересчитывается по min(qty, остаток) |
| R5. PDF-вёрстка таблицы с 21 партией | Низкая / низкое | Макет = упрощённая версия файла ЯРПРОФИТ (строки-партии, не этажи-колонки); образец согласовать на CP3 |

## Чекпоинты верификации

- **CP1 (конец нед. 1):** `pytest tests/test_delivery_schedule_*.py` зелёный;
  схема идемпотентна (двойной ensure_schema); cascade от KP_offers покрыт
  тестом. Gate: движок отвечает правильно на синтетике.
- **CP2 (середина нед. 2):** валидация на прошлых заказах — расчётная дата
  vs факт ≤ ±20%. Не сошлось → калибруем коэф./пороги, не трогаем API.
- **CP3 (конец нед. 2):** ручной прогон через uvicorn: PUT → GET (светофор)
  → import (файл ЯРПРОФИТ) → document. Согласовать макет документа.
- **CP4 (конец нед. 3):** E2E: кнопка в drawer → импорт → правка → светофор
  → скачан XLSX/PDF; `pytest tests/ -q`, `npm run test`, `npm run typecheck`,
  `npm run build` — все зелёные.

## Явно не в плане (из Not Doing спеки)

LLM-парсер чужих форматов, резервирование дорожек, история версий,
сводный экран всех графиков (фаза 1.5), учёт свободного СГП в расчёте.

## Tasks

Порядок = зависимости. ‖ = можно параллельно. Каждая задача ≤5 файлов.

### Неделя 1 — ядро

- [ ] **T1. Схема БД: 3 таблицы графика**
  - Acceptance: `delivery_schedule` / `delivery_batch` / `delivery_batch_item`
    создаются идемпотентно (двойной `ensure_schema` не падает); FK cascade
    от `KP_offers`; `UNIQUE kp_id`; индексы по schedule_id/batch_id/plate_id.
  - Verify: `pytest tests/test_delivery_schedule_schema.py -q`
  - Files: `core/kp_db_schema.py`, `tests/test_delivery_schedule_schema.py`

- [ ] **T2. Константы ёмкости → core/production_capacity.py**
  - Acceptance: `_MAX_TRACK_LENGTH_M`/`_DAYS_PER_TRACK_FACTOR` живут в новом
    модуле; `archive_service` импортирует оттуда; поведение estimate
    не изменилось.
  - Verify: `pytest tests/ -q -k archive` (все зелёные без правок тестов)
  - Files: `core/production_capacity.py` (новый), `app/services/archive_service.py`

- [ ] **T3. Движок светофора core/delivery_schedule_check.py**
  - Acceptance: чистая функция `check_batches(batches, occupancy, workdays,
    produced) -> [BatchCheck]`; детерминированные кейсы: зелёный/жёлтый/
    красный, остаток партии (вычет produced), выходные пропускаются,
    партии соревнуются за ёмкость в порядке produce_by, подсказка
    «+N дорожек до даты» для красных. R2: дата вне days_info = свободный
    день с дефолтным max — поведение зафиксировано тестом.
  - Verify: `pytest tests/test_delivery_schedule_check.py -q`
  - Files: `core/delivery_schedule_check.py`, `tests/test_delivery_schedule_check.py`

→ **CP1: pytest зелёный, движок корректен на синтетике.**

### Неделя 2 — API и импорт

- [ ] **T4. Pydantic-схемы + CRUD-сервис (без светофора)**
  - Acceptance: PUT принимает полный набор партий, валидирует (Σ по позиции
    ≤ qty в КП; deliver_from ≤ deliver_to; qty ≥ 1; 422 с понятным detail),
    идемпотентен (повторный PUT = тот же вид); GET отдаёт график или 404.
    КП не «в работе»/«На СГП» → 403/422 на PUT. Редактирование invoice_number.
  - Verify: `pytest tests/test_delivery_schedule_service.py -q`
  - Files: `app/schemas/delivery_schedule.py`,
    `app/services/delivery_schedule_service.py`,
    `tests/test_delivery_schedule_service.py`

- [ ] **T5. GET с живым светофором (связка сервиса и T3)**
  - Acceptance: GET собирает остатки из `KpReadinessService`, occupancy из
    `get_global_calendar_info`, рабочие дни из `work_calendar` и прогоняет
    `check_batches`; у каждой партии status/ready_date/hint; R4: позиция КП
    с уменьшенным qty помечается `changed: true`, падения нет.
  - Verify: `pytest tests/test_delivery_schedule_service.py -q` (расширенные)
  - Files: `app/services/delivery_schedule_service.py`,
    `tests/test_delivery_schedule_service.py`

- [ ] **T6. Роутер + регистрация**
  - Acceptance: `/commercial/archive/{kp_id}/delivery-schedule` GET/PUT;
    `require_roles("admin","manager")`; 403 для customer-роли; роутер
    зарегистрирован; OpenAPI-схема корректна.
  - Verify: `pytest tests/test_delivery_schedule_endpoints.py -q`
  - Files: `app/api/v1/endpoints/delivery_schedule.py`,
    `app/api/v1/endpoints/__init__.py`,
    `tests/test_delivery_schedule_endpoints.py`

- [ ] **T7. ‖ XLSX-шаблон: build + parse + эндпоинты /template, /import**
  - Acceptance: шаблон скачивается (колонки: Партия | Поставка с | Поставка
    по | Произвести до | Марка | Кол-во); parse возвращает черновик партий
    + unmatched_rows с причинами; /import НЕ сохраняет; файл ЯРПРОФИТ
    (счёт 234, лист 2) парсится после приведения к шаблону — round-trip
    тест template→parse.
  - Verify: `pytest tests/test_delivery_schedule_xlsx.py -q`
  - Files: `core/delivery_schedule_xlsx.py`,
    `app/api/v1/endpoints/delivery_schedule.py`,
    `tests/test_delivery_schedule_xlsx.py`

→ **CP2: валидация светофора на 3–5 прошлых заказах из БД (±20%), калибровка
констант. CP3: ручной прогон PUT→GET→import через uvicorn.**

### Неделя 3 — документы и фронт

- [ ] **T8. Документ XLSX + PDF + /document**
  - Acceptance: XLSX/PDF содержат шапку (договор, счёт или «КП №…», стороны)
    и таблицу партий с датами; файл сохраняется в outputs с именем
    `График_КП{kp_id}_ред_YYYY-MM-DD.{ext}`, старые не затираются; XLSX
    открывается в LibreOffice. Макет согласован с пользователем (R5).
  - Verify: `pytest tests/test_delivery_schedule_doc.py -q` + ручное
    открытие файла
  - Files: `core/delivery_schedule_xlsx.py`, `core/delivery_schedule_pdf.py`
    (новый), `app/services/delivery_schedule_service.py`,
    `tests/test_delivery_schedule_doc.py`

- [ ] **T9. Frontend: api/types/hooks фичи**
  - Acceptance: `deliveryScheduleApi` (get/put/import/template/document),
    TanStack-хуки с queryKey-фабриками по образцу useArchiveQueries,
    типы синхронны со схемами.
  - Verify: `cd frontend && npm run typecheck`
  - Files: `frontend/src/features/delivery-schedule/api/deliveryScheduleApi.ts`,
    `frontend/src/features/delivery-schedule/types/deliverySchedule.ts`,
    `frontend/src/features/delivery-schedule/hooks/useDeliveryScheduleQueries.ts`

- [ ] **T10. Frontend: редактор партий + светофор + кнопка в drawer**
  - Acceptance: кнопка «График поставки» в `OfferDetailsDrawer` (только
    «в работе»/«На СГП», иначе read-only/скрыта); редактор: позиции
    слева с остатками, партии справа, индикатор «разбито N из M»,
    валидация qty; светофор-чипы (зелёный/жёлтый/красный + hint);
    бейдж «есть график» в списке архива.
  - Verify: `cd frontend && npm run test && npm run typecheck`
  - Files: `frontend/src/features/delivery-schedule/components/DeliveryScheduleEditor.tsx`,
    `.../components/BatchCard.tsx`, `.../components/BatchStatusChip.tsx`,
    `commercial-archive/components/OfferDetailsDrawer.tsx`,
    `commercial-archive/components/ArchiveOfferList.tsx`

- [x] **T11. Frontend: импорт + документы**
  - Acceptance: dropzone XLSX → черновик в редактор, unmatched-строки
    списком сверху; кнопки «Скачать шаблон» / «XLSX» / «PDF»
    (через saveBlobAs).
  - Verify: `cd frontend && npm run test && npm run build`
  - Files: `frontend/src/features/delivery-schedule/components/ImportScheduleDialog.tsx`,
    `.../components/ScheduleDocumentButtons.tsx`

- [ ] **T12. E2E-прогон и финальная верификация**
  - Acceptance: сценарий менеджера из спеки (разбивка счёта 234 ≤10 мин)
    проходит на живом стеке `./run+logs.sh`; все Success Criteria спеки
    проверены; `pytest tests/ -q`, `npm run test`, `npm run typecheck`,
    `npm run build` — зелёные.
  - Verify: чек-лист Success Criteria из `docs/specs/delivery-schedule.md`
  - Files: — (только прогон; возможные фиксы — отдельными задачами)

→ **CP4: E2E зелёный, фича готова к выкатке.**
