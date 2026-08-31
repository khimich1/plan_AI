# Implementation Plan: Подсветка непроизводимых плит и замена нагрузки

**Спека:** `ai_docs/specs/kp-unpriced-plates-replacement.md`  
**Идея:** `ai_docs/ideas/unpriced-plates-replacement.md`  
**Дата:** 18.08.2026 · **Статус:** на ревью (код не начат)

## Overview

Два вертикальных релиза. Релиз 1 закрывает тихий `unit_price = 0` и включает уже существующий гейт wizard/экспорта. Релиз 2 добавляет inline-замену нагрузки по образцу wide-plates: предложения с ценами, одно действие «Применить», перегенерация preview.

## Architecture Decisions

1. **Релиз 1 — только три точки правды.** Не трогаем `core/price_db.py::get_price` (общая точка сметы/производства). Меняем `lookup_plate_price` (контракт КП), `_build_order_data` (источник `unit_price` черновика) и UI панели плит.
2. **`unit_price = None`, не `0.0`.** Согласуется с сваями/ступенями/маршами и с `collect_unpriced_positions` (`is not None`).
3. **Релиз 2 — зеркало wide-plates, не новый wizard-шаг.** Metadata-пара `unpriced_plate_lines` + `unpriced_plates_resolved`; блокировка через `WizardNextRequiredAction.resolve_unpriced_plates`; resolve переписывает `input_text` / `normalized_lines` и вызывает `generate_preview` заново.
4. **Детекция в MVP — после preview, по `order_data` с `unit_price is None`.** Так же, как wide-plates не отменяют оптимизацию до разрешения. Ранний skip оптимизации — follow-up, не блокер.
5. **Предложения считает backend.** Меньшие `load_code` той же `length_dm` с `price > 0`, по убыванию класса. Фронт только рисует и шлёт выбранный `load_code`.
6. **Смета закупок (`build_price_rows`) не трогаем** (open question из спеки → зафиксировано для плана).
7. **Текст подсказки UI:** «Замена нагрузки требует согласования с заказчиком.»
8. **Telegram-бот вне скоупа** (ботом не пользуемся).

## Dependency Graph

```
R1.1 lookup_plate_price (price<=0 → error)
    └── R1.2 _build_order_data (unit_price=None)
            └── R1.3 KpPlatePreviewPanel («нет в прайсе»)
                    └── Checkpoint R1

R2.1 core helper: replacements for (length_dm, load_code)
    └── R2.2 metadata + draft_service serialize/detect
            └── R2.3 wizard ERR + next_action
                    └── R2.4 resolve endpoint + workflow
                            └── R2.5 frontend UnpricedPlatesInlineSection
                                    └── Checkpoint R2
```

R1 полностью независим от R2 и должен мержиться/проверяться отдельно.

---

## Task List

### Phase 1: Релиз 1 — предохранитель

#### Task 1: `lookup_plate_price` считает `price <= 0` отсутствующей

**Description:** В `core/commercial_pricing.py::lookup_plate_price` после `SELECT` трактовать `None` и `<= 0` одинаково — `PriceNotFoundError`. Существующие вызовы `ensure_order_priced` / export начнут ловить нули из БД.

**Acceptance criteria:**
- [ ] Строка `(75, 12, 0.0)` в tmp `pb.db` → `lookup_plate_price(7.5, 1.2, 1200)` бросает `PriceNotFoundError`
- [ ] Строка `(70, 12, 29210)` → возвращает `29210.0`
- [ ] Отсутствующая строка по-прежнему бросает `PriceNotFoundError`

**Verification:**
- [ ] `pytest tests/test_commercial_pricing_errors.py tests/ -q -k "lookup_plate_price or ensure_order_priced"` зелёный
- [ ] Новый тест в `tests/test_commercial_pricing_errors.py` (или рядом)

**Dependencies:** None  
**Files likely touched:**
- `core/commercial_pricing.py`
- `tests/test_commercial_pricing_errors.py`  
**Estimated scope:** S (1–2 files)

---

#### Task 2: `_build_order_data` пишет `unit_price = None` при нулевой/отсутствующей цене сметы

**Description:** В `app/services/commercial_service.py::_build_order_data` заменить дефолт `0.0` на `None`; после парсинга сметной ячейки, если значение `<= 0` или matching_row нет / парсинг упал — оставлять `None`. Тогда `collect_unpriced_positions` начнёт находить плиты в черновике.

**Acceptance criteria:**
- [ ] Сметная строка с `0,00` → в `order_data` у позиции `unit_price is None`
- [ ] Позиция без matching_row → `unit_price is None`
- [ ] Позиция с положительной ценой → float как раньше
- [ ] `CommercialCalculationService.unpriced_position_labels` возвращает имя такой плиты

**Verification:**
- [ ] Unit/service тест на `_build_order_data` или через generate_preview с stubbed price_rows
- [ ] `pytest tests/test_commercial_calculation_service.py tests/test_commercial_pricing_errors.py -q`

**Dependencies:** Task 1 (желательно; логически независимо, но вместе дают полный предохранитель)  
**Files likely touched:**
- `app/services/commercial_service.py`
- `tests/test_commercial_service_order_data.py` (новый) или расширение существующего  
**Estimated scope:** S–M (2–3 files)

---

#### Task 3: Панель плит показывает «нет в прайсе»

**Description:** В `KpPlatePreviewPanel` повторить паттерн `KpPilePreviewPanel`: `unitPrice === null` → текст «нет в прайсе» (не `"0"`), при необходимости предупреждение сверху списка. `formatOfferNumber` не менять глобально — только отображение в панели плит (чтобы не сломать другие экраны, где `0` осмысленен).

**Acceptance criteria:**
- [ ] Строка с `unit_price: null` рендерит «нет в прайсе»
- [ ] Строка с числом рендерит локализованное число как раньше
- [ ] Визуально согласовано со сваями (цвет/тон ошибки опционально, не блокер)

**Verification:**
- [ ] `npm run build` из `frontend/`
- [ ] Ручная проверка: пересоздать черновик с ПБ 75-12-12п — в превью «нет в прайсе», wizard не пускает на результат

**Dependencies:** Task 2  
**Files likely touched:**
- `frontend/src/features/commercial-offer/components/KpPlatePreviewPanel.tsx`
- при необходимости `buildKpPreviewRows.ts` (без смены контракта)  
**Estimated scope:** S (1–2 files)

---

### Checkpoint: Релиз 1

- [ ] `pytest tests/ -q` зелёный
- [ ] `npm run build` зелёный
- [ ] Ручной сценарий: OCR/текст с ПБ 75-12-12п → `unit_price: null` в draft JSON, UI «нет в прайсе», сообщение wizard «Нет цен для позиций: …», generate-files / save → ошибка, не КП с нулём
- [ ] **Ревью с человеком перед Релизом 2**

---

### Phase 2: Релиз 2 — флоу замены

#### Task 4: Core helper — предложения замен нагрузки

**Description:** Чистая функция (например `core/commercial_pricing.py` или `core/unpriced_plate_replacements.py`): на вход `length_m`/`length_dm` + текущий `load_code` + `db_path`; на выход список `{load_code, price}` для меньших кодов (10, 8, 6 — только ниже текущего) с `price > 0`, сортировка по убыванию `load_code`. Без UI/HTTP.

**Acceptance criteria:**
- [ ] Для 75 дм / 12 → `[10→31890, 8→29316, 6→27144]` (или актуальные цены из tmp DB фикстуры)
- [ ] Для длины без цен ни на одном классе → `[]`
- [ ] Не предлагает равный или больший класс

**Verification:**
- [ ] `pytest tests/test_unpriced_plate_replacements.py -q`

**Dependencies:** Checkpoint R1  
**Files likely touched:**
- `core/commercial_pricing.py` или новый `core/unpriced_plate_replacements.py`
- `tests/test_unpriced_plate_replacements.py`  
**Estimated scope:** S (1–2 files)

---

#### Task 5: Metadata — `unpriced_plate_lines` / `unpriced_plates_resolved`

**Description:** После `_build_order_data` собрать позиции с `unit_price is None`, для каждой вызвать helper из Task 4, сериализовать в структуру как у wide-plates (`id`, `line`/`name`, `qty`, `length_m`, `width_m`, `load_class`, `replacements: [{load_code, price}]`). Пробросить через `ParseResult` или `CommercialPreviewResult` → `build_preview_metadata`. Флаг `unpriced_plates_resolved = not bool(lines)`. Схемы Pydantic в `app/schemas/commercial.py`.

**Acceptance criteria:**
- [ ] Preview черновика с ПБ 75-12-12п содержит непустой `metadata.unpriced_plate_lines` с replacements
- [ ] После того как все позиции оценены — `unpriced_plate_lines == []`, `unpriced_plates_resolved == true`
- [ ] Схемы ответа draft details принимают новые поля

**Verification:**
- [ ] Service/API тест на metadata
- [ ] `pytest tests/ -q -k "unpriced"`

**Dependencies:** Task 4  
**Files likely touched:**
- `app/domain/models/parse_result.py` и/или `CommercialPreviewResult`
- `app/services/commercial_service.py`
- `app/services/commercial_draft_service.py`
- `app/schemas/commercial.py`
- тесты  
**Estimated scope:** M (3–5 files) — если раздуется, разрезать на «domain serialize» + «wire metadata»

---

#### Task 6: Wizard блокирует по `resolve_unpriced_plates`

**Description:** Добавить `WizardNextRequiredAction.resolve_unpriced_plates`, константу `ERR_UNPRICED_PLATES` (отдельное сообщение, не путать с API `ERROR_CODE_UNPRICED_PLATES`), зеркальные ветки в `commercial_calculation_service` / `commercial_wizard_step_service` рядом с wide-plates. Приоритет: wide-plates и unpriced оба блокируют продукт-шаг; порядок — сначала wide, потом unpriced (как в спеке «зеркало»).

**Acceptance criteria:**
- [ ] При непустых `unpriced_plate_lines` и `resolved=false` → `next_required_action == resolve_unpriced_plates`, шаг результата недоступен
- [ ] После `resolved=true` действие снимается (если нет других блокеров)
- [ ] Существующие тесты wizard зелёные + новые кейсы

**Verification:**
- [ ] `pytest tests/test_commercial_wizard_step_service.py tests/test_commercial_calculation_service.py -q`

**Dependencies:** Task 5  
**Files likely touched:**
- `app/schemas/commercial.py`
- `app/services/commercial_calculation_service.py`
- `app/services/commercial_wizard_step_service.py`
- `tests/test_commercial_wizard_step_service.py`
- `tests/test_commercial_calculation_service.py`  
**Estimated scope:** M (3–5 files)

---

#### Task 7: Endpoint + workflow `resolve_unpriced_plates`

**Description:** `POST /drafts/{id}/unpriced-plates/resolve`. Тело: решения `{line_id, action: replace_load|exclude, load_code?}`. Валидация: `load_code` только из `replacements` этой строки; все строки покрыты; после exclude/replace список не пуст (если пуст — ошибка, как у wide). Переписать нагрузку в тексте строк (`-12п` → `-10п` и `load_class`), обновить batches, `generate_preview`, сохранить decisions в metadata.

**Acceptance criteria:**
- [ ] `replace_load` 12→10 для ПБ 75-12 → в новом draft `ПБ 75-12-10п`, `unit_price > 0`, `unpriced_plates_resolved=true`
- [ ] `exclude` удаляет строку
- [ ] Чужой `load_code` / неполное покрытие → 400 с понятным сообщением
- [ ] Нет замен → только `exclude` принимается

**Verification:**
- [ ] Интеграционный тест по образцу wide-plates resolve в `tests/test_commercial_web_flow.py` (или новый файл)
- [ ] `pytest tests/ -q -k "unpriced"`

**Dependencies:** Task 6  
**Files likely touched:**
- `app/services/commercial_workflow_service.py`
- `app/api/v1/endpoints/commercial.py`
- `app/schemas/commercial.py` (request body)
- `tests/test_commercial_unpriced_plates_resolve.py`  
**Estimated scope:** M (4–5 files)

---

#### Task 8: Frontend — `UnpricedPlatesInlineSection` + wiring

**Description:** Клон UX `WidePlatesInlineSection` на `PlateInputStep`: карточка «Нет в прайсе / не производится», список позиций, селект замен с ценами (первый предвыбран), пункт «Исключить», подсказка про согласование с заказчиком, кнопка «Применить». Типы в `commercialOffer.ts`, вызов нового API из wizard (рядом с wide resolve). Пока не resolved — блокер как у wide.

**Acceptance criteria:**
- [ ] Секция видна только при непустых unresolved `unpriced_plate_lines`
- [ ] Для ПБ 75-12-12п видны 10п/8п/6п с ценами; 10п предвыбрана
- [ ] «Применить» обновляет draft; секция пропадает; можно идти дальше по wizard
- [ ] Нет замен → только «Исключить»
- [ ] `npm run build` зелёный

**Verification:**
- [ ] `cd frontend && npm run build`
- [ ] Ручной E2E: тот же заказ → секция → применить 10п → пересчёт → сохранение КП без нулей

**Dependencies:** Task 7  
**Files likely touched:**
- `frontend/.../UnpricedPlatesInlineSection.tsx` (новый)
- `frontend/.../steps/PlateInputStep.tsx`
- `frontend/.../CommercialOfferWizard.tsx`
- `frontend/.../types/commercialOffer.ts`
- API client hook/файл рядом с wide-plates  
**Estimated scope:** M (4–5 files) — при раздувании вынести API client в отдельный XS-task

---

### Checkpoint: Релиз 2 / Complete

- [ ] Все success criteria спеки (Р1 + Р2) выполнены
- [ ] `pytest tests/ -q` и `npm run build` зелёные
- [ ] Ручной прогон на заказе с ПБ 75-12-12п и ПБ 75-10,8-12п
- [ ] Готово к ревью / merge (коммит — только по явной просьбе)

---

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| `unit_price: None` ломает экспорт/сериализацию JSON или xlsx | High | Task 1–2 тесты + ручной generate-files на unpriced draft (ожидаем ошибку, не traceback) |
| Переписывание суффикса нагрузки в свободном тексте OCR хрупко (`12п` vs `12,5п`) | Med | Нормализовать через `make_plate_name` / парсинг имени, не тупой `str.replace`; тесты на `12,5п`→floor 12 |
| Wide + unpriced одновременно на одном заказе | Med | Явный приоритет: сначала wide, потом unpriced; оба блокера в validation_errors |
| Прайсовые нули в БД vs пустые ячейки Excel | Low (R1) | R1 не чинит импорт; детектит `<= 0`. Импорт/зачистка нулей — отдельная задача, спросить перед правкой `pb.db` |
| Оптимизация всё ещё крутится для «призрака» до resolve | Low | Зафиксировано как MVP-компромисс; follow-up: early skip |

## Open Questions (оставшиеся)

- Завод реально производит 7.3+ м на 10п/8п/6п? (вне кода; не блокирует реализацию — UI всё равно покажет то, что в прайсе)
- Финальная формулировка подсказки — зафиксирована в плане; поменять можно до Task 8

## Not in this plan

- Правка `build_price_rows` / производственной сметы
- Чистка нулей в `pb.db` / смена `PRICE_XLSX_PATH`
- Журнал спроса, Telegram-бот, skip оптимизации до resolve
- Ручной ввод договорной цены
