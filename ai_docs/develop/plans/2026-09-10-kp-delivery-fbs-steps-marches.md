# Implementation Plan: КП — доставка для ФБС / ступеней / маршей по массе

**Created:** 2026-09-10  
**Status:** IMPLEMENT ✅ · Task 1–10 done (2026-09-10)  
**Spec:** [`ai_docs/specs/kp-delivery-fbs-steps-marches.md`](../../specs/kp-delivery-fbs-steps-marches.md)  
**Idea:** [`ai_docs/ideas/kp-delivery-fbs-steps-marches.md`](../../ideas/kp-delivery-fbs-steps-marches.md)

## Overview

Третий весовой котёл в `calculate_total_cost`: `ceil(Σкг{fbs,steps,marches} / 20 000) × logistics_cost`. Источник веса — новый справочник `product_weight_catalog` в plita.db, импорт из 1С-выгрузок с валидацией. Котёл включается только для новых КП флагом `fbs_lm_delivery_enabled` (D10) — архивные КП побайтово не меняются.

Это не горизонтальный «весь бэкенд потом UI»: после Task 4 математика проверяема изолированно; после Task 7 mono-ФБС КП сквозным путём визард → сохранение → архив считает доставку; Task 8–9 только открывают это в документах и UI.

## Architecture Decisions

- **A1. Один расчётный модуль — `core/commercial_pricing.py`.** Два других `calculate_total_cost` (`commercial_offer.py`, `commercial_offer_xlsx.py`) — тонкие делегаты (19 строк, alias `_calculate_total_cost`); им добавляется прокидка новых kwarg, логика не дублируется.
- **A2. Справочник в plita.db, не в pb.db** — конвенция `pile_catalog`: прайсы отдельно, каталоги отдельно. Путь резолвится тем же `_resolve_pile_catalog_db_path`-паттерном (новый параметр, не переиспользуем `pile_catalog_db_path` по имени).
- **A3. Resolver веса: точная марка → семейство ЛС-N → None (pending).** Нормализация как у `pile_catalog`/`fbs_price_db`: C↔С, T↔Т, без пробелов, запятая→точка. Семейный fallback только для ЛС (веса семейств однородны — проверено сверкой 2026-09-10); ФБС/ЛМ только точное совпадение. Fallback логируется.
- **A4. Тип строки — оба ключа:** `product_type ∈ {fbs,steps,marches}` (mixed-builder штампует) или `product_kind ∈ {fbs,step,march}` (mono/legacy-билдеры в `kp_order_data.py`). Плиты/сваи игнорируются симметрично тому, как их котлы игнорируют чужие типы.
- **A5. Флаг-гейт D10: `kp_meta.fbs_lm_delivery_enabled INTEGER DEFAULT 0`** (аддитивная миграция в `kp_db_schema.py` — паттерн `pile_trip_overrides_json`). Выставляется при создании драфта; сохраняется с КП; читается архивом из `raw`. Дефолт `False` в `calculate_total_cost` → все старые пути (бот-legacy, `offers_write`, архив) без флага не меняются.
- **A6. Новые ключи totals всегда присутствуют** (нули при выкл) — frontend-типы стабильны, снапшоты старых КП не ломаются: `fbs_lm_delivery_total`, `fbs_lm_cargo_kg`, `fbs_lm_trips`, `fbs_lm_delivery_ready`, `fbs_lm_pending_marks`. Старые ключи не меняются.
- **A7. Документы:** `kp_delivery_export_lines` + третья строка по конвенции: котлов несколько → «Доставка ФБС/ЛС/ЛМ»; котёл один в документе → «Услуга по доставке грузов». PDF/XLSX-генераторы получают флаг от архива; визард и `FileGenerationService` флаг берут из draft metadata.
- **A8. Импортер — офлайн-скрипт** `scripts/import_product_weights_from_xlsx.py` (паттерн `import_fbs_prices_from_xlsx.py`), `--dry-run` с отчётом обязателен перед реальной загрузкой. Валидация при импорте, не в рантайме КП.
- **A9. Сваи не трогаем.** Сверка 11 расхождений весов — отдельная data-задача, не из этого плана.
- **A10. Embed-XLSX (`xlsx_delivery_in_unit`)** — третий пул отдельным follow-up после мержа обеих фич; в этом плане embed-поведение не меняется (ФБС/ЛС/ЛМ надбавки не получают, пока их котёл вшивать нечему — регрессионный тест `test_embed_delivery_in_unit_price.py` должен остаться зелёным).

```
product_weight_catalog (plita.db)
        │ resolve(mark, type)
        ▼
compute_fbs_lm_delivery ──► calculate_total_cost(flag) ──► totals[fbs_lm_*]
        │                            │
        │        ┌───────────────────┼────────────────────────┐
        │        ▼                   ▼                        ▼
        │   wizard drafts      archive get_details      generators PDF/XLSX
        │   (calc service)     generate_document        (через делегаты)
        ▼
kp_delivery_export_lines ──► строка «Доставка ФБС/ЛС/ЛМ»
```

## Implementation order

| Phase | Focus | Depends |
|-------|-------|---------|
| 1 | Справочник весов: parse + validate + resolve + importer | — |
| 2 | Котёл: агрегация веса + `calculate_total_cost` с флагом | 1 (контракт resolver) |
| 3 | Флаг-гейт: схема, сохранение, делегаты, архив | 2 |
| 4 | Документы: третья строка доставки | 3 |
| 5 | UI: визард + drawer | 3 (контракт totals в спеке — тесты можно параллелить) |
| 6 | Реальный импорт + регрессия + docs | 1–5 |

## Task List

### Phase 1: Foundation — справочник весов

## Task 1: RED — контракт `product_weight_catalog`

**Description:** Тесты до кода. Парсинг xlsx (колонки Наименование/Вес/Объём/Код); валидация: плотность 1500–3500 кг/м³ → quarantine-список, код 1С в колонке объёма → reject строки, дубли марки с разным весом → reject файл/строки с отчётом; пропуск типов вне scope (ЛП, кросс-блоки). Resolve: точный, нормализация Т↔T/пробелы, ЛС-семейство (`ЛС14-Б` → вес `ЛС-14`), ФБС без семейного fallback, miss → None.

**Acceptance criteria:**
- [x] `parse_product_weights_from_xlsx(Блоки.xlsx, "fbs")` → 14 записей ФБС, кросс-блоки в skipped с причиной
- [x] ЛС.xlsx → 19 записей; «ЛС-15-1 закл» (плотность 251) в quarantine
- [x] ЛМ и ЛП.xlsx → ЛМ записи; 3 артефакта плотности в quarantine; ЛП в skipped
- [x] resolve: `ФБС 24.6.6-Т` ↔ `ФБС24.6.6-T`; `ЛС14-Б` → семейство 150 кг; `ЛС99` → None
- [x] upsert идемпотентен (повторный прогон = update, не дубль)
- [x] Тесты красные: модуля нет

**Verification:**
- [x] `pytest tests/test_product_weight_catalog.py -q` — fail import, не skip

**Dependencies:** None

**Files likely touched:**
- `tests/test_product_weight_catalog.py`

**Estimated scope:** S

## Task 2: GREEN — `core/product_weight_catalog.py` + importer CLI

**Description:** Модуль по спеке (schema `product_weight_catalog`: `mark_norm TEXT PK, product_type TEXT, weight_kg REAL NOT NULL, volume_m3 REAL, display_name TEXT, source_file TEXT, imported_at`), openpyxl read-only как в `pile_catalog`. CLI: `scripts/import_product_weights_from_xlsx.py <file> --type fbs|steps|marches [--db plita.db] [--dry-run]`. Dry-run печатает: импортировано / quarantine (с причиной) / skipped (с причиной).

**Acceptance criteria:**
- [x] Тесты Task 1 зелёные
- [x] `--dry-run` на трёх реальных файлах: ФБС 14 ok; ЛС 18 ok + 1 quarantine; ЛМ **9** ok + 3 quarantine; ЛП/КБ skipped (оценка «7 ЛМ» в черновике плана = покрытие прайса 7/7, не строк файла)
- [x] Без `--dry-run` пишет в plita.db; повторный запуск — 0 inserted, N updated
- [x] Нет I/O в resolve (только SELECT)

**Verification:**
- [x] `pytest tests/test_product_weight_catalog.py -q`
- [x] `python scripts/import_product_weights_from_xlsx.py "банк знаний/Новая папка/Блоки.xlsx" --type fbs --dry-run`

**Dependencies:** Task 1

**Files likely touched:**
- `core/product_weight_catalog.py`
- `scripts/import_product_weights_from_xlsx.py`
- `tests/test_product_weight_catalog.py`

**Estimated scope:** M

### Checkpoint: Foundation

- [x] Справочник зелёный изолированно; реальные файлы проходят dry-run с ожидаемым отчётом
- [x] Существующие тесты не запускались/не тронуты — модуль изолирован
- [x] Не переходить к котлу, пока ЛС-семейный fallback и quarantine не зелёные

---

### Phase 2: Core — котёл

## Task 3: RED — математика котла и флаг

**Description:** Тесты `tests/test_commercial_fbs_lm_delivery.py`. Контракт `compute_fbs_lm_delivery` + интеграция `calculate_total_cost(..., fbs_lm_delivery_enabled=...)`. Кейсы: 30 т × 50 000 → 100 000; граница 20 000 → 1 рейс, 20 001 → 2; qty=0 строки не участвуют; марка без веса → ready=False, total=0, марка в pending; три типа в одном котле; product_kind vs product_type; флаг off → все `fbs_lm_*` = 0/True/[], старые ключи побайтово; mixed плиты+ФБС → котлы независимы (18,6т и 20т параллельно).

**Acceptance criteria:**
- [x] Все кейсы выше красные (ключи отсутствуют)
- [x] Тест флага off фиксирует неизменность существующего вызова (сигнатура обратно-совместима)

**Verification:**
- [x] `pytest tests/test_commercial_fbs_lm_delivery.py -q` — fail

**Dependencies:** Task 2 (resolver-контракт зафиксирован)

**Files likely touched:**
- `tests/test_commercial_fbs_lm_delivery.py`

**Estimated scope:** S

## Task 4: GREEN — котёл в `calculate_total_cost`

**Description:** `core/cargo_delivery_pricing.py`: `FBS_LM_TRUCK_CAPACITY_KG = 20_000.0` + агрегация веса по типам через resolver каталога (НЕ через `resolve_kp_line_weight_kg` — он плитный, по length/width). `core/commercial_pricing.py`: kwarg `fbs_lm_delivery_enabled: bool = False`, `weight_catalog_db_path: str | None = None` (резолв как у pile-каталога); котёл в `delivery_total`; новые ключи totals всегда (A6).

**Acceptance criteria:**
- [x] Тесты Task 3 зелёные
- [x] `test_commercial_logistics_cost.py`, `test_pile_trip_pricing.py`, `test_commercial_export_mixed.py` зелёные **без правок**
- [x] Флаг off = прежний байт-в-байт результат (кроме добавленных нулевых ключей)

**Verification:**
- [x] `pytest tests/test_commercial_fbs_lm_delivery.py tests/test_commercial_logistics_cost.py tests/test_pile_trip_pricing.py -q`

**Dependencies:** Task 3

**Files likely touched:**
- `core/cargo_delivery_pricing.py`
- `core/commercial_pricing.py`
- `tests/test_commercial_fbs_lm_delivery.py`

**Estimated scope:** M

### Checkpoint: Core

- [x] Котёл зелёный, регрессия плит/свай зелёная без правок
- [x] Не идти в plumbing, пока pending-конвенция и граница 20 000 не зелёные

---

### Phase 3: Флаг-гейт (D10) сквозной путь

## Task 5: Схема + сохранение + драфт

**Description:** Миграция `kp_meta.fbs_lm_delivery_enabled INTEGER DEFAULT 0` (паттерн `pile_trip_overrides_json`). Создание драфта выставляет флаг в metadata. Сохранение КП (`kp_persistence_service` / `kp_repository`) персистит флаг; totals при сохранении считаются с флагом.

**Acceptance criteria:**
- [x] Миграция идемпотентна на существующей БД (повторный ensure_schema не падает)
- [x] Новый драфт → metadata содержит флаг; старые строки kp_meta → NULL/0
- [x] Сохранённый из визарда КП: флаг в kp_meta, `total_amount` учитывает котёл
- [x] Тест persistence: save → read raw → флаг на месте

**Verification:**
- [x] `pytest tests/test_kp_persistence_steps.py tests/test_kp_persistence_marches.py tests/test_commercial_fbs_lm_delivery.py -q`

**Dependencies:** Task 4

**Files likely touched:**
- `core/kp_db_schema.py`
- `core/kp_persistence_service.py`
- `app/repositories/kp_repository.py`
- `app/services/commercial_draft_service.py`
- `tests/test_commercial_fbs_lm_delivery.py`

**Estimated scope:** M (ровно 5 файлов)

## Task 6: Делегаты и генераторы + третья строка доставки

**Description:** Делегаты `commercial_offer.py` / `commercial_offer_xlsx.py` принимают и прокидывают `fbs_lm_delivery_enabled` + путь каталога; генераторы — новые kwargs (default off). `kp_delivery_export_lines`: третья строка по A7. PDF/XLSX smoke: mono-ФБС с флагом → строка есть; без флага → нет.

**Acceptance criteria:**
- [x] XLSX/PDF нового mono-ФБС КП содержат «Доставка ФБС/ЛС/ЛМ» с рейсами и суммой котла
- [x] Документ старого КП (без флага) побайтово как раньше (сравнение текстового слоя/ключевых строк)
- [x] Лейблы существующих строк («Доставка плит», «Доставка свай», «Услуга по доставке грузов») не изменились
- [x] `test_commercial_export_mixed.py`, `test_embed_delivery_in_unit_price.py` зелёные без правок

**Verification:**
- [x] `pytest tests/test_commercial_fbs_lm_delivery.py tests/test_commercial_export_mixed.py tests/test_embed_delivery_in_unit_price.py -q`

**Dependencies:** Task 5

**Files likely touched:**
- `core/commercial_offer.py`
- `core/commercial_offer_xlsx.py`
- `core/commercial_pricing.py` (только `kp_delivery_export_lines`)
- `tests/test_commercial_fbs_lm_delivery.py`

**Estimated scope:** M

## Task 7: Архив + визард-calc + offers_write

**Description:** `commercial_calculation_service.compute_totals(_from_metadata)` читает флаг из draft metadata. `archive_service.get_details`/`generate_document` — из `raw` и прокидывает в `calculate_total_cost` и генераторы. `kp/offers_write.py` (пересчёт при смене скидки/рейса) — читает флаг из сохранённого meta. D10-тесты: архивное mixed-КП без флага — totals побайтово прежние; с флагом — котёл считается; `update_logistics_cost` старого КП котёл не включает.

**Acceptance criteria:**
- [x] Карточка нового mono-ФБС КП: итог с котлом, `fbs_lm_*` в details
- [x] Карточка/документ архивного КП без флага: неизменны (D10-тест)
- [x] Смена скидки в архиве нового КП: котёл сохраняется и не дисконтируется
- [x] `test_archive_endpoints.py`, `test_archive_service.py` зелёные без правок ассертов

**Verification:**
- [x] `pytest tests/test_commercial_fbs_lm_delivery.py tests/test_archive_endpoints.py tests/test_archive_service.py tests/test_archive_pile.py -q`

**Dependencies:** Task 6

**Files likely touched:**
- `app/services/commercial_calculation_service.py`
- `app/services/archive_service.py`
- `core/kp/offers_write.py`
- `tests/test_commercial_fbs_lm_delivery.py` или `tests/test_archive_endpoints.py` (один из двух)

**Estimated scope:** M

### Checkpoint: Backend complete

- [x] Сквозной путь: новый mono-ФБС драфт → расчёт → сохранение → карточка архива → PDF/XLSX — везде котёл
- [x] Архивные КП без флага неизменны — D10-тест зелёный
- [x] Не приступать к UI, пока ключи totals и лейбл строки не совпадают со спекой

---

### Phase 4: UI

## Task 8: Визард — поле рейса и итоги

**Description:** `CalculationResultStep`: поле «Стоимость рейса» показывается при `hasPlateLines || hasFbsLmLines` (сейчас только `hasPlateLines`). Плашки: «Рейсов ФБС/ЛС/ЛМ» (как «Рейсов свай»), pending-марки котла (паттерн pile pending). Типы в `commercialOffer.ts`. Mixed-подписи полей не ломать (`Рейс плит`/`Рейс свай` при mixedDelivery как сейчас — общий тариф, подпись остаётся «Стоимость рейса», если котёл не плитно-свайный mixed).

**Acceptance criteria:**
- [x] Mono-ФБС драфт: поле рейса видно, применение обновляет итог (mock API)
- [x] Pending-марки отображаются; при ready=false котёл 0
- [x] Плиты/сваи UI без регрессии

**Verification:**
- [x] `cd frontend && npm run test -- --run src/features/commercial-offer`
- [x] `cd frontend && npm run typecheck`

**Dependencies:** Task 7 (для e2e); контракт totals из спеки — тесты можно писать параллельно с Task 5–7

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.tsx`
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx`
- `frontend/src/features/commercial-offer/types/commercialOffer.ts`

**Estimated scope:** M

## Task 9: Drawer — плашка и pending котла

**Description:** `OfferDetailsDrawer`: состав доставки показывает котёл ФБС/ЛС/ЛМ (сумма, рейсы) и pending-марки — по существующему паттерну pile. Типы в `archive.ts`. Кнопка «XLSX (доставка в цене)» не меняется (embed — follow-up A10).

**Acceptance criteria:**
- [x] Новое КП: доставка котла видна в карточке; pending-марки отображаются
- [x] Старое КП без флага: карточка не изменилась
- [x] disabled-логика кнопки embed не тронута

**Verification:**
- [x] `cd frontend && npm run test -- --run src/features/commercial-archive`
- [x] `cd frontend && npm run typecheck && npm run build`

**Dependencies:** Task 7; параллелимо с Task 8

**Files likely touched:**
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
- `frontend/src/features/commercial-archive/types/archive.ts`

**Estimated scope:** S

---

### Phase 5: Данные и sweep

## Task 10: Реальный импорт + регрессия + docs

**Description:** Импорт трёх файлов в plita.db (перед этим — бэкап `plita.db.bak-before-weight-catalog-<ts>`, конвенция бэкапов в репо). Сверка отчёта с ожиданиями Task 2. Полный focused pytest/vitest/typecheck/build. Обновить `ai_docs` (documenter): feature-заметка + статус плана. Ручной чеклист: mono-ФБС КП в визарде; mixed плиты+ФБС; архивное старое КП открыть — итог не изменился; новое КП в архиве — PDF/XLSX со строкой котла.

**Acceptance criteria:**
- [x] В справочнике: 14 ФБС + 18 ЛС + 9 ЛМ (после quarantine; «7 ЛМ» в черновике = прайс 7/7); отчёт приложен к ai_docs/develop/reports
- [x] Все Commands из спеки зелёные
- [ ] Ручной чеклист пройден в браузере (vitest покрыл UI; живой визард/PDF в этой сессии не открывались — нет browser tools)

**Verification:**
- [x] `pytest -q` (focused-набор из спеки) + `npm run test -- --run src/features/commercial-offer src/features/commercial-archive` + typecheck + build

**Dependencies:** Task 2, 8, 9

**Files likely touched:**
- `plita.db` (данные, не коммитить)
- `ai_docs/develop/reports/2026-09-10-product-weight-catalog-import.md`
- этот план (статусы чекбоксов)

**Estimated scope:** S

---

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Забытый caller `calculate_total_cost` без флага → карточка ≠ документ | High | A1: логика в одном модуле; grep-лист callers в Task 6–7 (commercial_offer, xlsx, kp_persistence ×2, offers_write ×2, archive ×2, calc service); D10-тест на каждый путь |
| Марка не смэтчилась → доставка молча 0 | High | A3 pending-конвенция + UI-плашка; семейный fallback только ЛС |
| Мусор из файлов в справочнике (6 артефактов) | Med | A8: валидация плотности при импорте; dry-run отчёт обязателен |
| Ретро-изменение архивных КП | High | A5 флаг-гейт + D10-тесты побайтово; `update_logistics_cost` флаг не включает |
| Конфликт с embed-фичей (обе правят totals/генератор) | Med | A10: этот план embed не трогает; третий пул — follow-up после мержа; регрессионный тест embed в каждой фазе |
| 20 000 кг / веса нетто не подтверждены логистикой | Med | Константа изолирована (A-список спеки); релиз после подтверждения, код не блокируется |
| Фура упрётся в объём раньше веса (габаритные ЛС) | Med | Проверка 1–2 реальных КП до релиза (Open Question спеки №2); валидатор уже хранит объём — задел под объёмную схему |
| Task 5/7 раздуваются >5 файлов | Med | Сплит зафиксирован (5 — persistence, 7 — читатели); schemas/commercial.py не трогаем — totals pass-through dict |

## Parallelization

| Можно параллельно | После |
|-------------------|--------|
| Task 8–9 UI-тесты (мок totals) ∥ Task 5–7 | Контракт totals зафиксирован в спеке (A6) |
| Task 1–2 (справочник) ∥ написание Task 3 RED | Resolver-контракт из спеки |

Один агент: строго 1→2→…→10. Параллельные агенты: UI (8–9) после Checkpoint backend complete.

## Open Questions

Нет блокирующих. Внешние проверки перед релизом (spec, Open Questions 1–3): 20 т и «нетто без палет» — логистика; объёмное ограничение ЛС — 1–2 реальных КП; вопросы бухгалтерии №1–7.

## DoD (не начинать IMPLEMENT без)

- [x] Каждая задача: acceptance + verify + ≤5 files + S/M
- [x] Чекпоинты после foundation / core / backend / UI
- [x] Граф зависимостей и риски
- [x] Человек подтвердил план (этот файл)

Первая задача реализации: **Task 1 RED** (`tests/test_product_weight_catalog.py`).
