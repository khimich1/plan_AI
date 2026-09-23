# Implementation Plan: Характеристики бетона на позиции КП

**Created:** 2026-09-23  
**Status:** PLAN ✅ · TASKS ✅ · IMPLEMENT ✅ (задачи 1–8 закрыты)  
**Spec:** [`ai_docs/specs/kp-concrete-spec.md`](../../specs/kp-concrete-spec.md)  
**Idea:** [`ai_docs/ideas/kp-concrete-spec.md`](../../ideas/kp-concrete-spec.md)

## Overview

На позиции с известным бетоном хранится снимок F, W и щебня. Новая строка получает гранитную пару заводской таблицы. Смена класса пересчитывает пару от того же щебня. Пара из «Другое» не меняется. Пустой снимок старого КП остаётся пустым. Цена, ступени, PDF и Excel не меняются.

Текст строки КП остаётся `марка класс количество`. Снимок живёт на элементе `order_data` и в колонках `kp_*`. После каждого пересбора строк из текста снимок прикрепляется заново по `line_id`.

## Architecture Decisions

- **A1. Lookup в `core/concrete_spec.py`.** Одна функция подсказки и одна функция прикрепления. Таблица — как в spec, включая М500 → F300 · W12. Фронт повторяет те же пары в `concreteSpec.ts`; vitest сверяет B25, B22.5, B30 и М500 с этой таблицей.
- **A2. В текст строки F/W не пишем.** Смена класса и массовое «применить класс» по-прежнему пересобирают текст и прогоняют прайс. Сразу после появления новых `order_data` вызывается `apply_line_specs(previous, new)`.
- **A3. Ключ сопоставления — `line_id`.** Его уже сохраняет `stamp_order_line_identity`. Нет пары по `line_id` — строка новая, ей гранитная подсказка. Пустой снимок у найденной старой строки не заполняется.
- **A4. Правило прикрепления.** `table` + щебень → пересчёт от нового класса. Рядового столбца нет (B30 и выше, плиты) → гранит, источник остаётся `table`. `manual` → F и W копируются как есть. Полей нет → поля не появляются.
- **A5. Правка пары не пересобирает текст.** Отдельный PATCH меняет четыре поля элемента `order_data` по `line_id`. Прайс не вызывается.
- **A6. SQL.** Nullable-колонки `frost_resistance`, `waterproofness`, `concrete_aggregate`, `concrete_spec_source` на `kp_piles`, `kp_marches`, `kp_fbs`, `kp_bridge_piles`, `kp_plates`. `kp_steps` без колонок. `SELECT *` в `offers_read` подхватит колонки сам. В INSERT/UPDATE писать только то, что пришло: отсутствие полей даёт NULL, не подсказку.
- **A7. Плиты.** Марку по-прежнему даёт `resolve_concrete_grade`. Подсказка считается от неё в момент сборки строки. В таблице цен колонки марки нет.
- **A8. Показ без редактора.** Итог расчёта и карточка архива читают уже сохранённый снимок. Пусто → «—».

```
новая строка / смена класса (текст как сейчас)
        │
        ▼
preview.order_data
        │
        ▼
apply_line_specs(previous by line_id)
        │
        ▼
order_data  ──PATCH пары──►  те же четыре поля
        │
        ▼
kp_piles | kp_marches | kp_fbs | kp_bridge_piles | kp_plates
        │
        ▼
order_data_from_kp_*  →  архив и повторное открытие
```

## Risks

| Риск | Что делаем |
|------|------------|
| Пересбор текста сотрёт рядовой щебень и «Другое» | `apply_line_specs` в том же месте, где уже есть `previous_order_data`: обычное обновление списка и `update_grades`. Проверка — задача 3, до UI |
| `line_id` на пересборе не совпал | Задача 3 падает, если после смены класса снимок сел не на ту строку. Индексный запасной ключ не добавляем молча |
| Повторное сохранение старого КП допишет дефолт | Задача 4: NULL на входе → NULL в SQL |
| Плита в черновике без марки | Подсказка только после `resolve_concrete_grade`. Нет марки и нет номера ПБ — пара пустая, не выдумываем М400 ради F/W |
| Схемы архива отрежут новые колонки | Задача 7 добавляет поля в pydantic и в типы архива. Пока их нет, UI архива не обещаем |

## Implementation order

| Phase | Focus | Depends |
|-------|-------|---------|
| 1 | Справочник и правило прикрепления | — |
| 2 | Черновик сваи: пересбор не теряет снимок, PATCH пары не трогает цену | 1 |
| 3 | Те же колонки и запись для маршей, ФБС, мостовых, плит | 2 |
| 4 | Колонка в мастере, показ в итоге и в архиве | 3 |

## Task List

### Phase 1: Справочник

## Task 1: RED+GREEN — подсказка по марке и классу

**Description:** Зафиксировать таблицу spec в `core/concrete_spec.py`. Коды `B7_5`, `B22_5`, `B30_granite`, `B30`, `M400` / `М400`, `M500` / `М500` приводятся к одной строке. Неизвестный код даёт `None`.

**Acceptance criteria:**
- [x] B25 → F200, W8, `granite`
- [x] B22.5 рядовой столбец доступен как W4 при том же F200; гранитная подсказка — W6
- [x] `B30` и `B30_granite` → F300, W10, рядового столбца нет
- [x] `М500` и `M500` → F300, W12
- [x] Пустая строка и неизвестный код → `None`

**Verification:**
- [x] `pytest tests/test_concrete_spec.py -q`

**Dependencies:** None

**Files likely touched:**
- `core/concrete_spec.py`
- `tests/test_concrete_spec.py`

**Estimated scope:** S

## Task 2: RED+GREEN — `apply_line_specs`

**Description:** Чистая функция без БД. Сопоставляет новые строки с предыдущими по `line_id` и проставляет снимок по A4.

**Acceptance criteria:**
- [x] Новая строка без прошлого `line_id` и класса B25 получает F200 · W8, источник `table`, щебень `granite`
- [x] Прошлая строка `table` + `ordinary` при новом классе B22.5 становится F200 · W4
- [x] Прошлая строка `ordinary` при новом классе B30 становится F300 · W10, источник `table`
- [x] Прошлая строка `manual` F150 · W4 при любом новом классе остаётся F150 · W4
- [x] Прошлая строка с теми же `line_id` и пустым снимком остаётся пустой
- [x] Цена и класс в функции не пересчитываются: на выходе тот `concrete_grade` и `unit_price`, что пришли в новой строке

**Verification:**
- [x] `pytest tests/test_concrete_spec.py -q`

**Dependencies:** Task 1

**Files likely touched:**
- `core/concrete_spec.py`
- `tests/test_concrete_spec.py`

**Estimated scope:** S

### Checkpoint: Справочник

- [x] Подсказка и прикрепление зелёные без FastAPI и без UI
- [x] Не переходить к черновику, пока пустой старый снимок в тесте остаётся пустым

---

### Phase 2: Черновик сваи

## Task 3: Вшить прикрепление в пересбор свай

**Description:** После сборки `preview.order_data` для свай вызывать `apply_line_specs` и с предыдущим циклом. Оба входа: обновление текста списка и `update_grades` / смена класса одной строки, потому что оба заново парсят текст.

**Acceptance criteria:**
- [x] Новый список `С110.35-12 B25 2` в черновике получает F200 · W8
- [x] Строка с `ordinary` после смены класса на B22.5 хранит F200 · W4 и тот же `line_id`
- [x] Строка `manual` после смены класса хранит прежние F и W
- [x] Строка без снимка после смены класса остаётся без снимка
- [x] `unit_price` до и после смены только W не используется: этот тест не ходит в PATCH пары, цена берётся из прайса по марке и классу как раньше

**Verification:**
- [x] `pytest tests/test_concrete_spec.py tests/test_commercial_pile_pricing.py -q`
- [x] Существующий тест смены класса свай остаётся зелёным; если его нет рядом с потоком черновика — добавить сценарий в ближайший тест `update_draft_pile` / grades, не заводить второй фреймворк

**Dependencies:** Task 2

**Files likely touched:**
- `app/services/product_draft_handler.py`
- тест потока черновика свай (существующий файл потока или `tests/test_commercial_pile_pricing.py`, если сценарий черновика уже там)

**Estimated scope:** M

## Task 4: PATCH пары без пересбора текста

**Description:** Эндпоинт меняет четыре поля одной незапечатанной строки черновика по `line_id`. Для источника `table` сервер сам ставит F и W из справочника. Для `manual` принимает F и W только из списков ГОСТ в spec. Прайс и текст списка не вызываются.

**Acceptance criteria:**
- [x] Выбор рядового на B25 пишет F200, W6, `ordinary`, `table`, `unit_price` тот же
- [x] «Другое» F150 W4 пишет `manual` и пустой щебень
- [x] F или W вне списка ГОСТ → 422, строка не меняется
- [x] Запечатанная строка и чужой `line_id` не меняются
- [x] Ступени этот эндпоинт не обслуживает

**Verification:**
- [x] `pytest` на новом тесте эндпоинта рядом с коммерческими draft-тестами

**Dependencies:** Task 3

**Files likely touched:**
- `app/api/v1/endpoints/commercial.py`
- `app/schemas/commercial.py`
- `app/services/product_draft_handler.py`
- `frontend/src/features/commercial-offer/api/commercialOfferApi.ts`
- тест эндпоинта

**Estimated scope:** M

### Checkpoint: Черновик сваи

- [x] Смена класса не стирает рядовой щебень и «Другое»
- [x] Смена пары не меняет цену
- [x] Не начинать остальные изделия, пока этот срез красный

---

### Phase 3: Остальные изделия и сохранённое КП

## Task 5: Колонки SQL и круговое чтение

**Description:** Nullable-колонки на пяти таблицах. Запись и обновление в `kp_persistence_service` копируют снимок с элемента. `order_data_from_kp_piles`, марши, ФБС, мостовые и плиты возвращают те же поля. Пустой вход остаётся NULL. `kp_steps` без колонок.

**Acceptance criteria:**
- [x] Свежая схема и миграция старой тестовой БД добавляют четыре колонки, повторный `ensure_schema` не падает
- [x] Сохранённая свая с F200 · W8 читается обратно в `order_data` с теми же полями
- [x] Сохранение элемента без полей оставляет NULL
- [x] В `kp_steps` этих колонок нет

**Verification:**
- [x] `pytest tests/test_kp_piles_schema.py tests/test_kp_marches_schema.py tests/test_kp_fbs_schema.py tests/test_kp_bridge_piles_schema.py tests/test_kp_steps_schema.py tests/test_kp_persistence_piles.py tests/test_kp_persistence_service.py -q`

**Dependencies:** Task 2

**Files likely touched:**
- `core/kp_db_schema.py`
- `core/kp_persistence_service.py`
- `core/kp_order_data.py`
- существующие тесты схемы и персистенции из команды выше

**Estimated scope:** M

## Task 6: Прикрепление для маршей, ФБС, мостовых и плит

**Description:** Тот же `apply_line_specs`, что у свай, на путях пересбора этих изделий. Для плиты перед подсказкой марка берётся из `resolve_concrete_grade`, явное значение не затирается. Новый список плит М500 получает F300 · W12. Старая плита без снимка его не получает.

**Acceptance criteria:**
- [x] Новый марш, ФБС и мостовая свая B25 получают F200 · W8
- [x] Мостовая свая B30 получает F300 · W10
- [x] Новая плита с маркой М500 получает F300 · W12, источник `table`
- [x] Новая плита с маркой М400 получает F300 · W10
- [x] Повторный пересбор плиты без снимка не заполняет поля

**Verification:**
- [x] `pytest tests/test_concrete_spec.py -q` плюс ближайшие тесты черновика маршей, ФБС, мостовых и плит, которые уже собирают `order_data`

**Dependencies:** Task 3, Task 5

**Files likely touched:**
- `app/services/product_draft_handler.py`
- место сборки плитных `order_data`, если оно не в этом handler
- тесты этих потоков

**Estimated scope:** M

### Checkpoint: Сохранённое КП

- [x] Круговое чтение свай и плит зелёное
- [x] Ступени и цена без diff поведения
- [x] Не переходить к UI, пока плита М500 не получает F300 · W12 в тесте

---

### Phase 4: Экран

## Task 7: Колонка «F / W» в мастере

**Description:** Строки превью свай, маршей, ФБС и мостовых несут снимок. В `KpGradedPreviewPanel` колонка после класса: готовые пары, пункт «Другое» двумя списками ГОСТ, пометка «не по таблице», пустой снимок как «—». Выбор вызывает PATCH из задачи 4. `KpPlatePreviewPanel` показывает ту же колонку без колонки марки. Запечатанная строка только текстом. «Применить класс ко всем» по-прежнему меняет класс; снимок дорисовывает сервер.

**Acceptance criteria:**
- [x] Свая B25 без клика показывает F200 · W8
- [x] Выбор рядового показывает F200 · W6 и уходит в PATCH
- [x] Смена класса в тесте панели не затирает локальную `manual`-пару до ответа сервера: панель показывает то, что пришло в черновике
- [x] У ступеней колонки нет
- [x] Плита М500 показывает F300 · W12, заголовка марки бетона нет

**Verification:**
- [x] `cd frontend && npm test -- src/features/commercial-offer/lib/concreteSpec.test.ts src/features/commercial-offer/components/KpGradedPreviewPanel.test.tsx src/features/commercial-offer/components/KpPlatePreviewPanel.test.tsx`

**Dependencies:** Task 4, Task 6

**Files likely touched:**
- `frontend/src/features/commercial-offer/lib/concreteSpec.ts`
- `frontend/src/features/commercial-offer/lib/concreteSpec.test.ts`
- `frontend/src/features/commercial-offer/lib/productTypePreview.ts`
- `frontend/src/features/commercial-offer/components/KpGradedPreviewPanel.tsx`
- `frontend/src/features/commercial-offer/components/KpPlatePreviewPanel.tsx`
- соответствующие `*.test.tsx` и билдеры строк превью, если без них колонка не видит поля

**Estimated scope:** M

Если билдеры строк не влезают в этот список вместе с панелями, сначала отдельным коммитом прокинуть четыре поля в `GradedPreviewRow` и `KpPreviewRow`, потом колонку. В один заход больше пяти содержательных файлов не мешать.

## Task 8: Показ в итоге и в архиве

**Description:** Итог расчёта и карточка архива показывают «F / W» текстом для свай, маршей, ФБС, мостовых и плит. Пусто — «—». Редактора там нет. Поля добавить в pydantic архива и в типы фронта, иначе `SELECT *` до экрана не доедет.

**Acceptance criteria:**
- [x] Итог свай показывает F200 · W8
- [x] Архивная свая и плита показывают сохранённую пару
- [x] Пустые колонки дают «—», не подсказку
- [x] У ступеней колонки нет
- [x] PDF и XLSX не получают новых колонок: файлы генерации документов не в diff

**Verification:**
- [x] `cd frontend && npm test -- src/features/commercial-offer/components/steps/CalculationResultStep.test.tsx src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`
- [x] `pytest tests/test_kp_persistence_piles.py -q`

**Dependencies:** Task 5, Task 7

**Files likely touched:**
- `frontend/src/features/commercial-offer/components/steps/CalculationResultStep.tsx`
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`
- `frontend/src/features/commercial-archive/types/archive.ts`
- `app/schemas/archive.py`
- тесты этих экранов

**Estimated scope:** M

### Checkpoint: Экран

- [x] Новая свая B25 показывает F200 · W8 в мастере, в итоге и в архиве
- [x] Старая строка без снимка показывает «—»
- [x] Цена, ступени, PDF и Excel не изменились

## Out of scope

Паспорт партии, производство, бот, печать КП, колонка марки на плитах, заполнение старых тестовых КП, свободный текст вместо списков ГОСТ.
