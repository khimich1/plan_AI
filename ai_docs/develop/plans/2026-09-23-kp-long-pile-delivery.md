# Implementation Plan: КП — доставка свай длиннее 13 м

**Created:** 2026-09-23  
**Status:** IMPLEMENT  
**Spec:** [`ai_docs/specs/kp-long-pile-delivery.md`](../../specs/kp-long-pile-delivery.md)  
**Idea:** [`ai_docs/ideas/kp-long-pile-delivery.md`](../../ideas/kp-long-pile-delivery.md)

## Overview

Сваи длиннее 13,0 м выходят из общего котла `pile_logistics_cost`. У каждой такой длины свой тариф рейса в этом КП. Рейсы считаются той же формулой, что сейчас у свай, но остаток 19 800 кг живёт только внутри длины. Марка без `pcs_per_20t` по-прежнему спрашивает число машин; пустой ответ обнуляет только свою длину. Старые КП не пересчитываются: флаг `long_pile_delivery_enabled` по образцу `fbs_lm_delivery_enabled`.

После Task 4 математика и строки документов проверяются без UI. После Task 7 новое КП со сваями сохраняется с флагом и теми же суммами в архиве. Task 8–9 только показывают тарифы в визарде и drawer.

## Architecture Decisions

- **A1. Старый расчёт не переписывается.** `compute_pile_trips` остаётся для флага выкл и для регрессии. Новая функция в том же `core/pile_trip_pricing.py` режет строки на короткий котёл и группы длин. `calculate_total_cost` вызывает её только при флаге.
- **A2. Ключ длины — `round(length_m × 10)`.** 13,0 м и короче — короткий котёл. 13,8 м → `138`, 14,0 м → `140`. Длина из `resolve_catalog_for_mark`. Нет каталога, но геометрия длиннее 13 м (`C18-40T8`, `C14-40T4`) — длинная группа без нормы. Нет ни каталога, ни геометрии — pending короткого котла, как сейчас.
- **A3. Один тариф на ключ.** `С140.35` и `С140.40` делят тариф и один `ceil(Σ кг / 19800)`. `-12` ключ не меняет. Остатки 14 м и 16 м не складываются.
- **A4. Число машин не дублируется.** N по-прежнему в `pile_trip_overrides`, ключ — марка строки (`С160.35-12`). Пустое поле не пишется. Явный 0 пишется. Рейсы длины = полные + остаток этой длины + сумма N. Есть pending внутри длины — `amount` длины 0, остальные длины и короткий котёл уже в сумме.
- **A5. Пустой тариф длины ≠ 0.** JSON только введённые цены: `{"140": {"trip_cost": 5000}}`. Нет ключа — длина не готова, сумма 0, строки в PDF нет. Явный 0 — готова, сумма 0, строки нет.
- **A6. Флаг `kp_meta.long_pile_delivery_enabled INTEGER DEFAULT 0` и `long_pile_delivery_json TEXT`.** Аддитивная миграция рядом с `fbs_lm_delivery_enabled`. Новый драфт ставит флаг `True`. Без флага все сваи в одном котле: один pending по `С160` обнуляет всю доставку свай.
- **A7. Новые ключи totals всегда есть** (нули и пустые списки при выкл): `long_pile_delivery_total`, `long_pile_lengths`, `long_pile_pending_marks`. При флаге `pile_delivery_total` — только короткие сваи. `delivery_total` прибавляет сумму готовых длин.
- **A8. Документы.** Готовая длина с суммой > 0 → «Доставка свай 14,0 м». Короткие сваи: единственная доставка в документе → «Услуга по доставке грузов», иначе «Доставка свай». Плиты и ФБС/ЛС/ЛМ не меняют формулу и подписи.
- **A9. `pile_catalog` не пишем.** Прочерк `pcs_per_20t` у `С160.30`, `С160.35`, `С160.40` остаётся. Embed-доставка в цену изделия не входит: `test_embed_delivery_in_unit_price.py` зелёный без смены ожиданий.

```
resolve_catalog_for_mark
        │ length_m > 13 ?
        ▼
compute_long_pile_groups ──► short: compute_pile_trips(только ≤13)
        │                     lengths: тариф × рейсы, pending локальный
        ▼
calculate_total_cost(flag) ──► pile_delivery_total + long_pile_delivery_total
        │
        ├── wizard / archive PATCH
        └── kp_delivery_export_lines ──► «Доставка свай 14,0 м»
```

## Implementation order

| Phase | Focus | Depends |
|-------|-------|---------|
| 1 | Группы длин и деньги в `calculate_total_cost` | — |
| 2 | Флаг и JSON: схема, драфт, сохранение, архив | 1 |
| 3 | Строки PDF/XLSX | 2 |
| 4 | Поля в визарде и drawer | 2 (контракт totals уже в спеке) |

## Task List

### Phase 1: Математика

## Task 1: RED — группы длин

**Description:** Тесты до кода в `tests/test_long_pile_delivery.py`. Контракт новой функции: короткий котёл и список длин. Кейсы из спеки без денег: `С130.35` в коротком котле; `С140.35` и `С150.35` нет; `С100.35` × 6 даёт 1 полный рейс, килограммы 14 м+ в этот остаток не входят; `С150.35` × 5 → 1 полная + остаток 4 650 кг = 2 рейса; `С140.35` и `С140.40` делят один остаток и не смешиваются с 15 м; `С140.35-12` в ключе 140; `С160.35-12` без N не обнуляет рейсы 15 м; явный N=0 закрывает марку; `C14-40T4` в ключе 140; `C18-40T8` — длинная группа без нормы; марка без геометрии остаётся pending короткого котла.

**Acceptance criteria:**
- [x] Кейсы выше красные: функции ещё нет
- [x] `tests/test_pile_trip_pricing.py` не меняется и остаётся про прежний один котёл

**Verification:**
- [x] `pytest tests/test_long_pile_delivery.py -q` — fail import, не skip

**Dependencies:** None

**Files likely touched:**
- `tests/test_long_pile_delivery.py`

**Estimated scope:** S

## Task 2: GREEN — нарезка в `pile_trip_pricing`

**Description:** Функция рядом с `compute_pile_trips`. Внутри длины та же формула: `floor(qty / pcs)` по марке с нормой, остаток только этих марок `ceil(кг / 19800)`, плюс N из `pile_trip_overrides`. Pending длины не затирает `total_trips` других групп. `compute_pile_trips` не меняет поведение.

**Acceptance criteria:**
- [x] Тесты Task 1 зелёные
- [x] `pytest tests/test_pile_trip_pricing.py -q` зелёный без правок

**Verification:**
- [x] `pytest tests/test_long_pile_delivery.py tests/test_pile_trip_pricing.py -q`

**Dependencies:** Task 1

**Files likely touched:**
- `core/pile_trip_pricing.py`
- `tests/test_long_pile_delivery.py`

**Estimated scope:** M

## Task 3: RED — деньги, флаг, строки документов

**Description:** Тесты `calculate_total_cost(..., long_pile_delivery_enabled=...)` и `kp_delivery_export_lines`. Короткий котёл: `С100.35` × 6 и тариф 1 000 ₽ → `pile_delivery_total` 1 000. Длина 15 м тариф 5 000 → 10 000. Две длины считаются раздельно. Нет тарифа 16 м → сумма 16 м = 0, короткий котёл и 15 м в итоге есть. `С160.35-12` × 4, N=2, тариф 7 000 → 14 000. Явный тариф 0 → длина готова, сумма 0, строки нет. Флаг выкл: все сваи в одном котле, пустой N по `С160` обнуляет `pile_delivery_total`, новые ключи нулевые. Подпись готовой длины «Доставка свай 14,0 м».

**Acceptance criteria:**
- [x] Кейсы красные: kwarg и ключи ещё не принимаются
- [x] Сигнатура без нового kwarg остаётся прежней (дефолт флага выкл)

**Verification:**
- [x] `pytest tests/test_long_pile_delivery.py -q` — fail на отсутствии ключей

**Dependencies:** Task 2

**Files likely touched:**
- `tests/test_long_pile_delivery.py`

**Estimated scope:** S

## Task 4: GREEN — `calculate_total_cost` и подписи

**Description:** Kwarg `long_pile_delivery_enabled: bool = False` и разбор JSON тарифов. При флаге `pile_delivery_total` только из короткого котла, `delivery_total` прибавляет суммы готовых длин. Ключи A7 всегда в ответе. `kp_delivery_export_lines` добавляет готовые длины с суммой > 0. Плиты 18 600 кг и ФБС/ЛС/ЛМ не трогаем.

**Acceptance criteria:**
- [x] Тесты Task 3 зелёные
- [x] `test_commercial_logistics_cost.py`, `test_pile_trip_pricing.py`, `test_commercial_export_mixed.py`, `test_commercial_fbs_lm_delivery.py`, `test_embed_delivery_in_unit_price.py` зелёные без правок ожиданий

**Verification:**
- [x] `pytest tests/test_long_pile_delivery.py tests/test_commercial_logistics_cost.py tests/test_pile_trip_pricing.py tests/test_commercial_export_mixed.py tests/test_commercial_fbs_lm_delivery.py tests/test_embed_delivery_in_unit_price.py -q`

**Dependencies:** Task 3

**Files likely touched:**
- `core/commercial_pricing.py`
- `tests/test_long_pile_delivery.py`

**Estimated scope:** M

### Checkpoint: Core

- [x] Группы и деньги зелёные, флаг выкл не меняет прежний котёл свай
- [x] Не идти в схему, пока `С160` без N не обнуляет соседние длины

---

### Phase 2: Флаг и JSON

## Task 5: Схема и разбор JSON

**Description:** Миграция `long_pile_delivery_enabled INTEGER DEFAULT 0` и `long_pile_delivery_json TEXT` в `kp_meta`, рядом с `fbs_lm_delivery_enabled`. Coerce: пустой тариф не хранится, явный 0 хранится, ключ — строка length_key. Повторный `ensure_schema` не падает.

**Acceptance criteria:**
- [x] Миграция идемпотентна
- [x] Старые строки без колонок читаются как флаг 0 и пустой словарь
- [x] Мусор в JSON не роняет расчёт: тариф этой длины считается не введённым

**Verification:**
- [x] `pytest tests/test_long_pile_delivery.py -q`

**Dependencies:** Task 4

**Files likely touched:**
- `core/kp_db_schema.py`
- `core/pile_trip_pricing.py`
- `tests/test_long_pile_delivery.py`

**Estimated scope:** S

## Task 6: Драфт, сохранение, архив

**Description:** Новый драфт ставит `long_pile_delivery_enabled: True` там же, где сейчас `fbs_lm_delivery_enabled`. `commercial_calculation_service` передаёт флаг и JSON в `calculate_total_cost`. Сохранение КП пишет обе колонки (`kp_persistence_service`, `kp_repository`, `offers_read` / `offers_write`). Архив читает их из raw и принимает PATCH JSON целиком: итоги пересчитываются сразу, PDF по кнопке, как у `pile_logistics_cost`.

**Acceptance criteria:**
- [x] Новый драфт содержит флаг; КП, сохранённое до миграции, считается одним котлом
- [x] Save → read raw: флаг и JSON на месте, `total_amount` включает готовые длины
- [x] PATCH тарифа 15 м меняет сумму и не трогает короткий котёл
- [x] N по `С160.35-12` по-прежнему в `pile_trip_overrides`

**Verification:**
- [x] `pytest tests/test_long_pile_delivery.py tests/test_archive_endpoints.py tests/test_archive_service.py -q`

**Dependencies:** Task 5

**Files likely touched:**
- `core/kp_persistence_service.py`
- `app/repositories/kp_repository.py`
- `core/kp/offers_read.py`
- `core/kp/offers_write.py`
- `app/services/commercial_draft_service.py`
- `app/services/commercial_calculation_service.py`
- `app/services/archive_service.py`
- `app/schemas/commercial.py`
- `app/schemas/archive.py`
- `tests/test_long_pile_delivery.py`

**Estimated scope:** L — если по ходу раздуется, резать на «сохранение» и «архивный PATCH», не меняя контракт Task 4

### Checkpoint: Persistence

- [x] КП с флагом и КП без флага дают разные `pile_delivery_total` на одном и том же заказе с 16 м
- [x] Не идти в PDF, пока PATCH не крутит JSON туда-обратно

---

### Phase 3: Документы

## Task 7: PDF и XLSX

**Description:** Делегаты `commercial_offer.py` и `commercial_offer_xlsx.py` прокидывают флаг и JSON (дефолт выкл). Генераторы и `commercial_export_service` берут их из metadata драфта или из архива. Строка «Доставка свай 14,0 м» только у готовой длины с суммой > 0.

**Acceptance criteria:**
- [x] Новый КП: 14 м и 16 м — две строки с разными тарифами
- [x] Неготовая длина строки не получает
- [x] Документ без флага не получает длинных строк
- [x] `test_commercial_export_mixed.py` и `test_embed_delivery_in_unit_price.py` зелёные без смены ожиданий

**Verification:**
- [x] `pytest tests/test_long_pile_delivery.py tests/test_commercial_export_mixed.py tests/test_embed_delivery_in_unit_price.py -q`

**Dependencies:** Task 6

**Files likely touched:**
- `core/commercial_offer.py`
- `core/commercial_offer_xlsx.py`
- `app/services/commercial_export_service.py`
- `app/services/product_draft_handler.py`
- `tests/test_long_pile_delivery.py`

**Estimated scope:** M

### Checkpoint: Documents

- [x] Текстовый слой PDF/XLSX содержит длину только у готовой суммы > 0
- [x] Подписи плит и ФБС/ЛС/ЛМ те же

---

### Phase 4: UI

## Task 8: Итоги визарда

**Description:** Если в заказе нет свай длиннее 13 м, шаг итогов как сейчас. Если есть — поле рейса коротких свай остаётся, ниже тариф на каждую длину и число рейсов. Вопрос «Для {марка} ({qty} шт.) нет нормы загрузки в справочнике. Сколько машин нужно?» только у марки из `long_pile_pending_marks` или у pending короткого котла. Пустой тариф не уходит на сервер как 0. Длина в подписи: ключ 140 → «14,0 м».

**Acceptance criteria:**
- [x] КП только до 13 м не показывает тарифы длин
- [x] 15 м показывает тариф и рейсы без вопроса про машины
- [x] `С160.35-12` показывает вопрос; введённое N пишется в `pile_trip_overrides`
- [x] Пустое поле тарифа не обнуляет соседнюю длину на экране после пересчёта

**Verification:**
- [x] `cd frontend && npm run test -- --run src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx`
- [x] `cd frontend && npm run typecheck`

**Dependencies:** Task 6

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.tsx`
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx`
- `frontend/src/features/commercial-offer/types/commercialOffer.ts`
- `frontend/src/features/commercial-offer/api/commercialOfferApi.ts`
- `frontend/src/features/commercial-offer/hooks/useCommercialOfferWizard.ts`

**Estimated scope:** M

## Task 9: Drawer архива

**Description:** В карточке КП с флагом те же тарифы длин и вопросы по маркам без нормы. Сохранение вызывает тот же PATCH, что Task 6. КП без флага показывает одно поле рейса свай, как сейчас.

**Acceptance criteria:**
- [x] Смена тарифа 14 м в drawer меняет сумму и не сбрасывает тариф 16 м
- [x] КП без флага не рисует поля длин
- [x] Явный 0 в тарифе сохраняется как 0, пустое поле тариф не затирает

**Verification:**
- [x] `cd frontend && npm run test -- --run src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
- [x] `cd frontend && npm run typecheck`

**Dependencies:** Task 6, Task 8 (общие типы)

**Files likely touched:**
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
- `frontend/src/features/commercial-archive/types/archive.ts`

**Estimated scope:** M

### Checkpoint: Complete

- [x] Команды из спеки зелёные
- [x] В `pile_catalog` у трёх `С160.*` пустой `pcs_per_20t`
- [ ] Ручной проход: КП со `С100` + `С150` + `С160` — вопрос только у 16 м, пустой ответ не гасит 10 м и 15 м
- [ ] Готово к ревью, не к продакшен-выкладке без этого прохода

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Флаг забудут прокинуть в один из делегатов, и PDF посчитает старым котлом | High | Task 7 сравнивает сумму строк документа с `long_pile_delivery_total` |
| Task 6 шире пяти файлов | Med | Резать на сохранение и PATCH, контракт Task 4 не менять |
| Пустое поле UI уйдёт как 0 и закроет длину | High | Тесты Task 8 и 9: пустое поле не в JSON |
| Остаток 14 м случайно попадёт в короткий котёл | High | Task 1 фиксирует килограммы до кода |

## Open Questions

Нет.
