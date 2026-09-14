# Implementation Plan: Календарь периодов в «В производство»

> **Спека:** [`ai_docs/specs/kalendar-periodov-v-proizvodstvo.md`](../../specs/kalendar-periodov-v-proizvodstvo.md)
> **Идея:** [`ai_docs/ideas/kalendar-periodov-v-proizvodstvo.md`](../../ideas/kalendar-periodov-v-proizvodstvo.md)
> **Дата:** 2026-09-04
> **Статус:** draft — НЕ выполнять до approve. Код не писать.

## Overview

Наглядность диалога «В производство»: жёлтая полоса окна и дата клиенту в
модалке; в drawer «Ёмкость» — месячная сетка (неделя = клик), список жильцов
(холд/обещание поимённо, план числом), бейдж холда с датой. Модель корзин не
меняется, кроме `week_count` 12→26. Красный `capacity-snapshot` гейт из этого
диалога уходит; график поставок не трогаем. `MonthCalendarGrid` /
`FactoryMiniCalendar` не копируем.

## Architecture Decisions

- **Occupants — отдельный GET**, не раздувать каждую котировку: клик недели
  дешёвый; quote остаётся лёгким. `week_start` только понедельник → 422.
- **Праздники на quote**, не `/production/work-calendar` (роли production).
- **Список жильцов включает текущее КП** (`is_current`); суммы quote по-прежнему
  `exclude_kp_id` — не съедать своё место в `free`.
- **Сетка не пишет срок.** Единственный колбек — `onSelectWeek(weekStart)`.
- **`PromiseWeekStrip` оставляем файл** (есть unit-тесты); из диалога и drawer
  убираем. Удаление файла — ask first.
- **`week_count=26` только default конструктора** `PromiseService`; тесты с
  явным `week_count=3` не ломаем. Формула allocate без изменений.
- **Красный snapshot-гейт** снимаем только в `MoveToProductionDialog`. Endpoint
  `capacity-snapshot` и график поставок без изменений.

```
PromiseRepository.list_week_allocs
        │
        ▼
PromiseService.list_week_occupants + holidays на quote + week_count=26
        │
        ▼
GET promise-quote (поля)  │  GET promise-weeks/{week_start}/occupants
        │                              │
        ▼                              ▼
PromiseWindowBand          PromisePeriodCalendar → PromiseWeekOccupants
        │                              │
        └──────── MoveToProductionDialog (drawer) ────────┘
```

## Task List

### Phase 1: Backend

- [ ] **Task 1: `week_count=26` + holidays/extra на quote**
  - **Description:** Default `PromiseService._week_count = 26`. В
    `PromiseQuoteResponse` добавить `holidays` и `extra_workdays` (даты в
    span `weeks[0].week_start` … последняя+6) через тот же `is_workday`, что
    корзины. Существующие тесты с явным `week_count=3` не трогать; один тест
    на default ≥ 26 недель и непустой/пустой список праздников.
  - **Acceptance:** без явного `week_count` в `weeks[]` ≥ 26; quote JSON содержит
    два новых поля; allocate/gate регресс зелёный.
  - **Verify:** `pytest tests/test_promise_service.py tests/test_promise_buckets.py tests/test_archive_endpoints.py -q`
  - **Dependencies:** None
  - **Files:** `app/services/promise_service.py`, `app/schemas/archive.py`,
    `tests/test_promise_service.py`, при необходимости `tests/test_archive_endpoints.py`
  - **Scope:** S

- [ ] **Task 2: GET occupants недели**
  - **Description:** `PromiseRepository.list_week_allocs(week_start, kinds)` —
    active hold+promise, expire холдов на чтении, **без** exclude kp.
    `PromiseService.list_week_occupants`: 422 если не понедельник;
    `customer_name` с карточки КП; `planned` по occupancy рабочих дней недели
    (неделя вне горизонта котировки — не 404); `is_current`. Роут
    `GET /{kp_id}/promise-weeks/{week_start}/occupants`, роли admin+manager.
    В payload нет `created_by`.
  - **Acceptance:** кейсы спеки (hold+promise, expire, is_current, 422, 404);
    manager 200; production 403.
  - **Verify:** `pytest tests/test_promise_service.py tests/test_archive_endpoints.py -q`
  - **Dependencies:** T1 (тот же сервис/схемы; holidays не обязательны для T2)
  - **Files:** `app/repositories/promise_repository.py`,
    `app/services/promise_service.py`, `app/schemas/archive.py`,
    `app/api/v1/endpoints/archive.py`, тесты
  - **Scope:** M

### Checkpoint: Phase 1

- [ ] `pytest tests/test_promise_service.py tests/test_archive_endpoints.py tests/test_move_to_production_atomicity.py -q`
- [ ] Ручной curl quote (26 weeks + holidays) и occupants на понедельник
- [ ] Review контракта перед UI

### Phase 2: Frontend slices

- [ ] **Task 3: Бейдж холда — дата клиенту**
  - **Description:** Видимый текст: `к {formatQuoteDayMonth(promised_date)} · до вечера`
    в `ArchiveOfferList` и `OfferDetailsDrawer`. Tooltip `created_by` можно
    оставить. `holdCreatedByTitle` не обязан совпадать с видимым текстом.
  - **Acceptance:** vitest «к 18.09 · до вечера»; нет голого «срок закреплён до сегодня»
    как единственной подписи.
  - **Verify:** `cd frontend && npm run test -- --run src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx src/features/commercial-archive/components/ArchiveOfferList.test.tsx`
  - **Dependencies:** None (∥ Phase 1)
  - **Files:** `ArchiveOfferList.tsx`, `OfferDetailsDrawer.tsx`, их тесты;
    при необходимости `promiseQuote.ts` (хелпер подписи)
  - **Scope:** S

- [ ] **Task 4: `PromiseWindowBand` в диалоге вместо strip**
  - **Description:** Новый компактный блок окна (`from_week`…`to_week`, маркер
    `promised_date`). В `MoveToProductionDialog` вместо верхнего
    `PromiseWeekStrip`. Drawer на этом шаге можно не трогать.
  - **Acceptance:** нет `data-testid="promise-week-strip"` в модалке; есть band;
    нет `window` → полоса не рендерится.
  - **Verify:** `cd frontend && npm run test -- --run src/features/factory-capacity/components/PromiseWindowBand.test.tsx src/features/commercial-archive/components/MoveToProductionDialog.test.tsx`
  - **Dependencies:** None (∥ T3; типы quote уже есть)
  - **Files:** `PromiseWindowBand.tsx` (NEW), `PromiseWindowBand.test.tsx` (NEW),
    `MoveToProductionDialog.tsx`, `MoveToProductionDialog.test.tsx`
  - **Scope:** S

- [ ] **Task 5: `PromisePeriodCalendar` (только сетка)**
  - **Description:** Месяц пн–вс, серые нерабочие из `holidays`/`extra_workdays`
    (+ сб/вс вне span). Клик любой клетки → `onSelectWeek(isoMonday)`. Жёлтые
    строки окна; маркер `promised_date`. Нет `onSelectDate`, нет кисти, не
    импортировать `MonthCalendarGrid`/`FactoryMiniCalendar`.
  - **Acceptance:** клик 10.09 и 18.09 одной недели → один `week_start`;
    навигация месяцев ограничивается пропсами `minMonth`/`maxMonth`.
  - **Verify:** `cd frontend && npm run test -- --run src/features/factory-capacity/components/PromisePeriodCalendar.test.tsx`
  - **Dependencies:** T1 желателен (поля holidays); без них пропсами можно
    кормить [] и weekend-эвристику
  - **Files:** `PromisePeriodCalendar.tsx` (NEW), `PromisePeriodCalendar.test.tsx` (NEW);
    мелкий хелпер ISO-понедельника в `factory-capacity/lib/` если не раздувать компонент
  - **Scope:** M

- [ ] **Task 6: `PromiseWeekOccupants` + API-клиент**
  - **Description:** Хук `usePromiseWeekOccupantsQuery(kpId, weekStart)` на новый
    GET. Список: КП, клиент, kind, дорожки, дата; `is_current` визуально; строка
    «Уже в плане: N»; пустой текст и подпись про холды — как в спеке. Без
    `created_by`.
  - **Acceptance:** vitest копирайта и строк; 422/ошибка — Alert.
  - **Verify:** `cd frontend && npm run test -- --run src/features/factory-capacity`
  - **Dependencies:** T2
  - **Files:** `promiseQuote.ts`, `PromiseWeekOccupants.tsx` (NEW),
    `PromiseWeekOccupants.test.tsx` (NEW)
  - **Scope:** S

- [ ] **Task 7: Собрать drawer + снять snapshot-гейт**
  - **Description:** Drawer «Ёмкость»: календарь → occupants выбранной недели →
    `PromiseKnobSettings`. Default неделя = ISO-неделя `promised_date`.
    `maxMonth` = max(конец quote.weeks, месяц распознанного срока). Убрать
    `PromiseWeekStrip` из drawer. Убрать `useCapacitySnapshotQuery` /
    `isCapacityRed` из диалога: submit не серый из-за red snapshot. Ширина
    drawer ≥ 380 (чуть шире ок, не две колонки). Обновить тесты диалога,
    которые ждут capacity gate.
  - **Acceptance:** AC спеки US-1…6; `npm run typecheck`; атомарность pytest.
  - **Verify:**
    ```
    pytest tests/test_move_to_production_atomicity.py tests/test_archive_endpoints.py -q
    cd frontend && npm run test -- --run src/features/factory-capacity src/features/commercial-archive && npm run typecheck
    ```
  - **Dependencies:** T4, T5, T6
  - **Files:** `MoveToProductionDialog.tsx`, `MoveToProductionDialog.test.tsx`;
    при необходимости ширина `Drawer`
  - **Scope:** M

### Checkpoint: Complete

- [ ] Команды из спеки зелёные
- [ ] Ручной прогон на КП №6 (или аналог): бейдж с датой; диалог — полоса и 18.09;
      Ёмкость — клик недели, свой холд в списке; ранняя дата — ошибка корзин, не
      snapshot; март в поле — месяц открывается
- [ ] `PromiseWeekStrip` не в диалоге/drawer; файл и его unit-тест живы
- [ ] Human review перед merge

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Тесты quote без `week_count` ждут 12 недель | Med | Grep `PromiseService(`; явный week_count в старых тестах оставить |
| Роут `promise-weeks` пересечётся с `/{kp_id}/...` | Low | Статический сегмент после kp_id, рядом с `promise-quote` |
| Occupants N+1 за customer_name | Low | v1: get_by_id в цикле (недель мало); batch — ask later |
| Снятие red-гейта регрессит график поставок | Med | Не трогать DeliverySchedule* и capacity-snapshot endpoint |
| Сетка читается как бронь дня | Med | Только `onSelectWeek`; vitest двух дат одной недели |

## Open Questions

Нет блокирующих. Ширина drawer — в T7 на глаз (≥ 380px, не > 480 без вопроса).

## Stop

План записан. **Реализацию не начинать** до явного approve («делай» / `/implement`).
