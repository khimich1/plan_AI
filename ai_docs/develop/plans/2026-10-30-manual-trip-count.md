# Implementation Plan: Ручное количество рейсов

**Created:** 2026-10-30  
**Status:** IMPLEMENT  
**Spec:** [`ai_docs/specs/manual-trip-count.md`](../../specs/manual-trip-count.md)  
**Idea:** [`ai_docs/ideas/manual-trip-count.md`](../../ideas/manual-trip-count.md)

## Overview

Менеджер вписывает целое «Кол. рейсов» у трёх готовых котлов: плиты, короткие сваи, ФБС/ЛС/ЛМ. Пока ключа котла нет, сумма и файлы остаются на нынешней формуле, и первое заполнение не пишет «было 0». Сохранённое число живёт в `trip_counts_json`. Смена массы своего котла больше чем на 0.01 кг заменяет число формулой и, если целое на карточке сменилось, оставляет напоминание, пока его не закроют. Тариф, скидка и целевая сумма число не трогают. Длинные сваи и неготовые котлы считаются как сейчас.

Подмена — в одном месте в `calculate_total_cost`, перед умножением на тариф. `cargo_delivery_trips_count` и `compute_pile_trips` остаются чистыми.

## Architecture Decisions

- **A1. Формулы не знают про ручной ввод.** `cargo_delivery_trips_count`, `_fbs_lm_trips_count` и `compute_pile_trips` не получают accepted. Новая чистая функция `resolve_trip_counts` в `core/trip_counts.py` сравнивает массу с `basis_kg` и возвращает принятое число, формулу, notice и обновлённый JSON.
- **A2. Один JSON на КП и на черновик.** `kp_meta.trip_counts_json` и `metadata.trip_counts`. Схема ключей как в спеке: `plate` / `pile` / `fbs_lm` = `{accepted, basis_kg, formula}`, плюс `notices`. Пустой объект и отсутствие ключа — «формула, записать без notice». Явный `accepted: 0` — ноль рейсов.
- **A3. Подмена до умножения.** `plate_trips`, `pile_trips`, `fbs_lm_trips` в ответе `calculate_total_cost` — принятое число (формула, если ключа ещё не было). Рядом `*_trips_formula` и `*_trips_notice`. Деньги: `round(тариф × принятое, 2)`. Котёл не готов — accepted игнорируется, сумма 0, ключ этого котла из JSON убирается, чтобы следующее «готов» было первым заполнением без notice.
- **A4. Масса котла.** Плиты — уже посчитанные кг плит. ФБС/ЛС/ЛМ — `cargo_kg` готового котла. Короткие сваи — сумма `weight_kg × qty` коротких строк (длиннее 13 м в эту массу не входят). Порог `abs(масса − basis_kg) > 0.01`. Тариф в сравнение не входит.
- **A5. Явный ввод не переживает сброс по массе.** OK шлёт целое ≥ 0 тем же каналом, что `logistics_cost` (`plate_trips` / `pile_trips` / `fbs_lm_trips`). Если в этом же пересчёте масса котла уже сдвинулась, в JSON пишется формула, а не число из запроса. Пустое, дробное и отрицательное отклоняются схемой и UI и не затирают ключ. Закрытие notice — `dismiss_trip_notices`, accepted не меняется.
- **A6. Файлы смотрят в тот же итог.** `kp_delivery_export_lines` не получает новый вид строки. PDF/XLSX прокидывают `trip_counts`. Сумма 0 по-прежнему прячет строку. Сноска вшитой доставки складывает `plate_trips` и `pile_trips` из итога. Длинные длины не пишутся в JSON.
- **A7. Грузоподъёмность, нормы, отгрузки, НДС.** 18 600 и 20 000 не меняются. `pile_catalog` не пишется. `pile_trip_overrides` остаётся вопросом «сколько машин». Процент скидки и инпут целевой суммы не переписываются из-за рейсов.

```
формула котла (чистая)
        │
        ▼
resolve_trip_counts(stored JSON, масса, ready)
        │ accepted / formula / notice / новый JSON
        ▼
calculate_total_cost  →  тариф × accepted
        │
        ├── черновик metadata.trip_counts
        ├── kp_meta.trip_counts_json
        └── kp_delivery_export_lines (trips, amount)
```

## Implementation order

| Phase | Focus | Depends |
|-------|-------|---------|
| 1 | Резолвер и деньги в `calculate_total_cost` | — |
| 2 | Колонка, черновик, архивный PATCH | 1 |
| 3 | PDF/XLSX и сноска | 1 |
| 4 | Поле на шаге итогов и в drawer | 2 |

## Task List

### Phase 1: Математика

## Task 1: RED — принятое число и деньги

**Description:** Тесты до кода. Резолвер: 215 220 кг → формула 12 без notice; явный 10 при той же массе; 0 при живом весе; шаг массы > 0.01 заменяет 10 на формулу и пишет notice, даже если формула всё ещё 12; та же масса оставляет 10; первое заполнение без notice; мусор в JSON как отсутствие ключа. `calculate_total_cost`: OK=10 и тариф 2 000 → доставка 20 000 и `plate_trips` 10; смена тарифа не сбрасывает 10 и не ставит notice; нет JSON — сумма как формула; pending сваи игнорируют accepted и обнуляют доставку; готовые короткие сваи умножаются на `pile_logistics_cost`; ФБС `ceil(кг/20000)` и общий тариф с плитами; два котла — две строки notice.

**Acceptance criteria:**
- [x] Кейсы выше красные: резолвера ещё нет или `calculate_total_cost` не принимает `trip_counts`
- [x] Старые тесты логистики не переписаны под ручное число

**Verification:**
- [x] `pytest tests/test_commercial_logistics_cost.py tests/test_commercial_fbs_lm_delivery.py tests/test_pile_trip_pricing.py -q` — новые падают, старые зелёные

**Dependencies:** None

**Files likely touched:**
- `tests/test_commercial_logistics_cost.py`
- `tests/test_commercial_fbs_lm_delivery.py`

**Estimated scope:** M

## Task 2: GREEN — `resolve_trip_counts` и подмена в итоге

**Description:** Чистый резолвер и вызов из `calculate_total_cost` перед умножением на тариф. Ответ дополняется `*_trips_formula`, `*_trips_notice`, `trip_counts`. Неготовый котёл не входит в сумму. Длинные сваи в JSON не пишутся.

**Acceptance criteria:**
- [x] Тесты Task 1 зелёные
- [x] `cargo_delivery_trips_count(215220)` по-прежнему 12, ёмкости 18600 и 20000 не изменены
- [x] `test_pile_trip_pricing.py` зелёный без правок

**Verification:**
- [x] `pytest tests/test_commercial_logistics_cost.py tests/test_commercial_fbs_lm_delivery.py tests/test_pile_trip_pricing.py -q`

**Dependencies:** Task 1

**Files likely touched:**
- `core/trip_counts.py`
- `core/commercial_pricing.py`
- `tests/test_commercial_logistics_cost.py`

**Estimated scope:** M

### Checkpoint: Core

- [x] Нет ключа → сумма равна формуле, notice пустой
- [x] Не идти в схему, пока тариф не оставляет ручные 10 рейсов

---

### Phase 2: JSON и API

## Task 3: Схема и разбор JSON

**Description:** `ALTER TABLE kp_meta ADD COLUMN trip_counts_json TEXT` рядом с `long_pile_delivery_json`. Повторный `ensure_schema` не падает. Чтение КП без колонки и с пустым текстом — пустой объект. Мусор не роняет расчёт.

**Acceptance criteria:**
- [x] Миграция идемпотентна
- [x] Старая строка без JSON не меняет сумму

**Verification:**
- [x] `pytest tests/test_commercial_logistics_cost.py tests/test_archive_endpoints.py -q`

**Dependencies:** Task 2

**Files likely touched:**
- `core/kp_db_schema.py`
- `core/trip_counts.py`
- `core/kp/offers_read.py`

**Estimated scope:** S

## Task 4: Черновик, сохранение, архивный PATCH

**Description:** `compute_totals_from_metadata` передаёт JSON. После пересчёта черновик записывает обновлённый `trip_counts`, если он изменился. Сохранение КП пишет колонку. Архивный PATCH принимает целые ≥ 0 по одному котлу и `dismiss_trip_notices`: итог пересчитывается сразу, закрытие строки не меняет accepted. Скидка и тариф пересчитывают деньги тем же JSON и не сбрасывают accepted. КП без JSON не меняет сумму.

**Acceptance criteria:**
- [x] PATCH `plate_trips=10` при тарифе 2 000 даёт доставку 20 000 и тот же JSON на чтении
- [x] Закрытие notice убирает строку и оставляет accepted
- [x] Отрицательное и дробное тело запроса не стирает сохранённое число
- [x] Смена скидки не затирает accepted

**Verification:**
- [x] `pytest tests/test_archive_endpoints.py tests/test_commercial_calculation_service.py tests/test_commercial_logistics_cost.py -q`

**Dependencies:** Task 3

**Files likely touched:**
- `app/schemas/commercial.py`
- `app/schemas/archive.py`
- `app/services/commercial_calculation_service.py`
- `app/services/commercial_draft_lifecycle.py`
- `app/services/archive_service.py`

**Estimated scope:** L — прокидку в `offers_write` / `kp_persistence_service` / эндпоинты резать следом, не меняя контракт Task 2

### Checkpoint: Persistence

- [x] Туда-обратно: записать 10, прочитать 10, закрыть notice, accepted на месте
- [x] Не идти в PDF, пока КП без JSON считается формулой

---

### Phase 3: Документы

## Task 5: PDF и XLSX

**Description:** Делегаты PDF/XLSX и генераторы черновика/архива прокидывают `trip_counts`. Строка доставки берёт `trips` и `amount` из итога. Accepted 0 или тариф 0 — строки нет. Сноска «включена доставка N рейс(ов)» складывает принятые числа вшитых котлов. Длинная свая остаётся строкой по длине.

**Acceptance criteria:**
- [x] Принято 10, тариф 2 000 — в строке количество 10 и сумма 20 000
- [x] Принято 0 — строки доставки нет
- [x] `test_embed_delivery_in_unit_price.py` зелёный; ожидание сноски меняется только если тест задаёт accepted

**Verification:**
- [x] `pytest tests/test_commercial_export_mixed.py tests/test_embed_delivery_in_unit_price.py tests/test_commercial_logistics_cost.py -q`

**Dependencies:** Task 2, Task 4

**Files likely touched:**
- `core/commercial_offer.py`
- `core/commercial_offer_xlsx.py`
- `app/services/commercial_export_service.py`
- `app/services/archive_service.py`
- `core/kp_persistence_service.py`

**Estimated scope:** M

### Checkpoint: Documents

- [x] Текстовый слой или строка экспорта содержит принятое число, не вторую колонку «по весу»
- [x] Длинные сваи без ручного ключа

---

### Phase 4: UI

## Task 6: RED — vitest шага итогов и drawer

**Description:** Тесты до разметки. Под тарифом плит есть «Кол. рейсов» со значением с сервера. OK целого вызывает тот же submit, что тариф, с этим числом. «по весу: N» только при `accepted ≠ formula`. Две строки «было / стало». Закрытие одной не снимает другую. Pending сваи: вопроса «сколько машин» нет рядом с новым полем, поля «Кол. рейсов» нет. Пустое, дробное и отрицательное OK не вызывают submit.

**Acceptance criteria:**
- [x] Новые кейсы красные, старые подписи «Рейсов свай» / «Рейсов ФБС/ЛС/ЛМ» и «сколько машин» не сломаны ожиданием

**Verification:**
- [x] `cd frontend && npm run test -- --run src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`

**Dependencies:** Task 4 (контракт полей totals)

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx`
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`

**Estimated scope:** S

## Task 7: Поле на шаге «Итоги и скидка»

**Description:** Инпут под тарифом плит, под тарифом коротких свай, если котёл готов, и под общим тарифом ФБС, если котёл готов. Напоминание над полями, кнопка закрытия не блокирует OK. Серая «по весу: N». Ответ сервера после сброса затирает несохранённый текст. Целевая сумма не подменяется из-за рейсов.

**Acceptance criteria:**
- [x] Vitest Task 6 зелёный
- [x] `npm run typecheck` без ошибок в затронутых типах

**Verification:**
- [x] `cd frontend && npm run test -- --run src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx`
- [x] `cd frontend && npm run typecheck`

**Dependencies:** Task 6

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.tsx`
- `frontend/src/features/commercial-offer/types/commercialOffer.ts`
- `frontend/src/features/commercial-offer/api/commercialOfferApi.ts`
- `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`
- `app/api/v1/endpoints/commercial.py`

**Estimated scope:** M

## Task 8: Drawer архива

**Description:** Те же поля и те же строки напоминания. OK вызывает PATCH логистики. Закрытие notice — тот же PATCH с `dismiss_trip_notices`. КП без числа показывает формулу из `plate_trips`. Неготовый котёл поле не рисует.

**Acceptance criteria:**
- [x] Смена числа в drawer уходит в `updateLogisticsCost` и не сбрасывает тариф соседнего котла
- [x] Закрытие одной строки не закрывает другую
- [x] Старые тесты «Рейсов ФБС/ЛС/ЛМ» остаются зелёными

**Verification:**
- [x] `cd frontend && npm run test -- --run src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
- [x] `cd frontend && npm run typecheck`

**Dependencies:** Task 4, Task 7

**Files likely touched:**
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`
- `frontend/src/features/commercial-archive/types/archive.ts`
- `frontend/src/features/commercial-archive/api/archiveApi.ts`
- `frontend/src/features/commercial-archive/hooks/useArchiveQueries.ts`
- `app/api/v1/endpoints/archive.py`

**Estimated scope:** M

### Checkpoint: Complete

- [x] Команды из спеки зелёные
- [ ] Ручной проход шага итогов: поле показывает формулу, OK подставляет тариф × вписанное целое. 2026-10-30: браузер открыл `/commercial-offer/new` и остановился на `/commercial-offer/login` (поля «Логин», «Пароль», кнопка «Войти»). Учётных данных нет, поле «Кол. рейсов» не нажималось.
- [x] 18 600, 20 000, нормы `pile_catalog` и вопрос «сколько машин» не переписаны

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| JSON не прокинут в один делегат PDF, и файл считает формулой | High | Task 5 сравнивает `trips` строки с `plate_trips` итога |
| Запись на чтении черновика затрёт ручное число старым ответом | High | Пишем только JSON, который вернул последний `calculate_total_cost`; явный OK применяется в этом же вызове |
| Пустое OK уйдёт как 0 | High | Схема `int ≥ 0` и отказ UI до запроса; `None` значит «не менять» |
| Скидка пересчитает доставку без JSON | High | `update_kp_discount` читает и пишет `trip_counts_json` |
| Масса коротких свай случайно включит длины > 13 м | Med | В массу короткого котла не входят длины с ключом длинной группы |
| Task 4 шире пяти файлов | Med | Контракт Task 2 не менять; прокидку резать на мета и PATCH |

## Open Questions

Нет. Спека утверждена.
