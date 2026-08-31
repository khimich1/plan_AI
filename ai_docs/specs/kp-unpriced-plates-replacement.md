# Spec: Подсветка непроизводимых плит и замена нагрузки в wizard КП

_Дата: 18.08.2026 · Статус: на ревью · One-pager: `ai_docs/ideas/unpriced-plates-replacement.md` · Инцидент: ПБ 75-12-12п с `unit_price = 0` в черновике `fb656f712bb3458fae787376d0153362`_

## Objective

**Проблема.** Плита, комбинации «длина × нагрузка» которой нет в прайсе (завод её не производит), сегодня молча получает `unit_price = 0.0` и проходит весь путь до сохранённого КП. Пользователь замечает это глазами — или не замечает.

**Цель.** Менеджер видит такие плиты сразу после распознавания заказа, wizard блокируется до разрешения, а замена на производимую нагрузку (та же длина/ширина, меньший класс) делается в 2 клика из предложенных вариантов.

**Пользователь:** менеджер, формирующий КП. **Успех:** нулевая цена невозможна ни в одном артефакте (черновик, КП, xlsx), а разрешение проблемы быстрее, чем ручная правка текста заказа.

**Два релиза:**
- **Релиз 1 — предохранитель:** нулевая/отсутствующая цена становится `None`, существующий гейт wizard начинает работать, UI показывает «нет в прайсе». Без нового флоу.
- **Релиз 2 — флоу замены:** детекция после парсинга, блокировка wizard, inline-замена нагрузки (зеркало wide-plates).

## Корневые причины (установлены 18.08.2026)

1. В `pb.db.prices` для нагрузки 12п цены есть только до 72 дм; строки 73–90 дм хранят `price = 0` (наследие старого прайса 19.08.24; новый прайс 14.07.26 держит эти ячейки пустыми, импорт нули не затирает). Аналогично: 10п — нули с 84 дм, 8п — с 91 дм.
2. `core/commercial_pricing.py::lookup_plate_price` бросает `PriceNotFoundError` только при отсутствии строки; строка с `price = 0` возвращается как валидная.
3. `app/services/commercial_service.py::_build_order_data` всегда пишет float в `unit_price` (дефолт `0.0`), никогда `None` → `collect_unpriced_positions` (`is not None`) для плит не срабатывает никогда.
4. `frontend/.../KpPlatePreviewPanel.tsx` рендерит `null` как `"0"` (`formatOfferNumber`); паттерн «нет в прайсе» есть у свай/ступеней/маршей/ФБС, у плит — нет.
5. Цена проверяется на этапе сметы, когда 2D-оптимизация уже посчитана для несуществующей плиты.

## Tech Stack

Backend: Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`, `pb.db`). Frontend: React + Vite + TypeScript. Тесты: pytest (`tests/`). venv: `.venv/`.

## Commands

```bash
# Backend тесты (из корня, venv активирован)
pytest tests/ -q

# Точечно по затронутым модулям
pytest tests/ -q -k "commercial_pricing or commercial_service or price"

# Backend dev
uvicorn app.main:app --reload

# Frontend (из frontend/)
npm run build
npm run dev
```

## Project Structure (затрагиваемые области)

```
core/commercial_pricing.py            → lookup_plate_price, collect_unpriced_positions
core/price_db.py                      → get_price (НЕ меняем: см. Boundaries)
app/services/commercial_service.py    → _build_order_data (источник unit_price черновика)
app/services/commercial_workflow_service.py → resolve_wide_plates (образец для Р2)
app/services/commercial_wizard_step_service.py → гейты wizard, ERR_* сообщения
app/services/commercial_calculation_service.py → unpriced_position_labels
app/schemas/errors.py                 → коды ошибок wizard
app/api/v1/endpoints/commercial.py    → resolve_draft_wide_plates (образец endpoint Р2)
core/plate_text_normalizer.py         → get_wide_plate_lines (образец детекции строк)
frontend/src/features/commercial-offer/components/KpPlatePreviewPanel.tsx
frontend/src/features/commercial-offer/components/WidePlatesInlineSection.tsx (образец Р2)
frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx
tests/                                → новые тесты обоих релизов
```

## Code Style

Слои: роутер → сервис → репозиторий; не смешивать ORM, схемы и бизнес-логику. Минимальный diff — не трогать несвязанный код. Pydantic v2 схемы в `app/schemas/`. Комментарии — только для неочевидных решений. Имена полей metadata в snake_case, флаги разрешения — парой `<name>_lines` + `<name>_resolved` (существующая конвенция wide-plates).

```python
# Образец стиля (существующий код):
if self.wide_lines_blocking(metadata):
    product_step = self.product_step(metadata)
    if stored in (WizardStepId.client, WizardStepId.result):
        return product_step
```

## Design

### Релиз 1 — предохранитель

1. `lookup_plate_price`: строка с `price <= 0` трактуется как отсутствующая → `PriceNotFoundError`.
2. `_build_order_data`: распарсенная из сметы цена `<= 0` (или нет matching_row) → `unit_price = None` (не `0.0`).
3. `KpPlatePreviewPanel`: `unit_price === null` → «нет в прайсе», сумма — «—» (копия паттерна `KpPilePreviewPanel`).
4. Эффект без дополнительного кода: `collect_unpriced_positions` начинает находить такие плиты → `unpriced_position_labels` → wizard показывает «Нет цен для позиций: …» и не пускает на шаг результата; `calculate_order_total(require_all_priced=True)` бросает `UnpricedPlatesError` → существующий обработчик в `app/core/http_errors.py`; экспорт xlsx (`_resolve_line_unit_price`) бросает `PriceNotFoundError` → generate-files для неразрешённого черновика невозможен.

Не входит: изменение отображения сметы закупок (`build_price_rows` продолжит показывать `0,00` — источник данных для черновика, не пользовательский артефакт; см. Open Questions).

### Релиз 2 — флоу замены (зеркало wide-plates)

**Детекция (после парсинга, до оптимизации).** При построении preview для каждой распарсенной строки плиты проверяется `(length_dm, load_code)` по `pb.db.prices` через существующий `get_price` (включая ceil-ключ и floor нагрузки 12.5→12). Цена `None` или `<= 0` → строка попадает в `metadata["unpriced_plate_lines"]` вместе с **предложениями замен**: меньшие классы нагрузки той же длины с ценой > 0, отсортированные по убыванию класса (10п → 8п → 6п), каждый с ценой из прайса. Порядок флагов аналогичен wide-plates: `unpriced_plates_resolved = not bool(unpriced_plate_lines)`.

**Wizard.** Новый код `ERR_UNPRICED_PLATES` (`app/schemas/errors.py`) и `WizardNextRequiredAction.resolve_unpriced_plates`; блокировка — по образцу `wide_lines_blocking`: пока флаг не resolved, wizard не уходит дальше шага продукта.

**API.** `POST /api/v1/commercial/drafts/{draft_id}/unpriced-plates/resolve` — зеркало `resolve_draft_wide_plates`. Тело: решения по каждой строке `{line_id, action}`:
- `replace_load` + `load_code` (только из предложенных — валидация на backend) → строка заказа переписывается: суффикс нагрузки в normalized-строке заменяется (`ПБ 75-12-12п` → `ПБ 75-12-10п`);
- `exclude` → строка удаляется (единственный вариант, когда предложений нет).
Если не по всем строкам есть решения → ошибка «Нужно выбрать действие для всех позиций без цены» (аналог wide-plates). Если список стал пустым → ошибка.
Применение = переписанный текст → `generate_preview` заново → переоптимизация штатная → `unpriced_plates_resolved=True`, решения сохраняются в `metadata["unpriced_plate_decisions"]`.

**Frontend.** `UnpricedPlatesInlineSection` на `PlateInputStep` (клон `WidePlatesInlineSection`): по каждой строке — имя плиты, радио/селект предложений с ценами («10п — 31 890 ₽»), первый (ближайший меньший) предвыбран, отдельный пункт «Исключить позицию». Кнопка «Применить» неактивна, пока не выбраны решения по всем строкам. Подсветка: «замена нагрузки требует согласования с заказчиком».

**Граничный случай «нет замен»** (длина вообще вне прайса, напр. > 96 дм): секция показывает «Производимых замен нет» и доступен только `exclude`.

## Testing Strategy

pytest, новые тесты в `tests/` рядом с существующими commercial/price тестами. Уровни: unit (lookup, детекция, генерация предложений, переписывание строки) + service (resolve endpoint через workflow service с tmp draft store).

Ключевые кейсы:
- Р1: `get_price`-стуб с `price=0` → `lookup_plate_price` бросает `PriceNotFoundError`; `_build_order_data` со сметной строкой `0,00` → `unit_price is None`; `collect_unpriced_positions` находит такую плиту.
- Р2: детекция `ПБ 75-12-12п` → предложения `[10п, 8п, 6п]` с ценами 31890/29316/27144; `replace_load` на 10п → в перегенерированном черновике `ПБ 75-12-10п` с `unit_price > 0`, флаг resolved; `exclude` удаляет строку; нерешённые строки → 400; валидация отклоняет `load_code` не из предложений; длина вне прайса → предложений нет, доступен только exclude.
- Регрессия: `pytest tests/ -q` зелёный; `npm run build` зелёный.

## Boundaries

- **Always:** минимальный diff; тесты перед завершением каждой задачи; схемы Pydantic для новых полей API; решения хранить в metadata черновика (аудит).
- **Ask first:** изменение `core/price_db.py::get_price` (общая точка — смета, производство, бот-архив); любые правки данных в `pb.db` (цены); изменение контракта `resolve_draft_wide_plates`.
- **Never:** коммитить без явной просьбы; трогать несвязанный код (ГСМ, shipments и пр.); молча менять поведение `build_price_rows` для производственной сметы.

## Success Criteria

**Релиз 1:**
- [ ] Черновик из заказа с ПБ 75-12-12п: у строки `unit_price = null`, панель плит показывает «нет в прайсе», сумма строки «—».
- [ ] Wizard: сообщение «Нет цен для позиций: Плиты ПБ 75-12-12п», шаг результата недоступен.
- [ ] Сохранение КП и generate-files для такого черновика завершаются понятной ошибкой, а не КП с нулём.
- [ ] `pytest tests/ -q` и `npm run build` зелёные.

**Релиз 2:**
- [ ] Тот же заказ: после распознавания wizard блокирован `resolve_unpriced_plates`; на шаге ввода видна секция с ПБ 75-12-12п и предложениями 10п/8п/6п с ценами; 10п предвыбрана.
- [ ] «Применить» → черновик пересчитан: ПБ 75-12-10п с ценой > 0, wizard разблокирован, оптимизация пересчитана для новой нагрузки.
- [ ] ПБ 75-10,8-12п (та же длина, ширина 1.08) детектируется той же детекцией — сегодняшний «3450 ₽ без базовой цены» невозможен.
- [ ] Плита вне прайса целиком: только «исключить»; после исключения всех проблемных строк флоу продолжается.
- [ ] `pytest tests/ -q` и `npm run build` зелёные.

## Open Questions

1. Смета закупок (`build_price_rows`) продолжит показывать `0,00` для таких плит — оставляем как есть в обоих релизах или в Р2 тоже подсвечиваем?
2. Текст подсветки «требует согласования с заказчиком» — финальная формулировка?
3. Завод реально производит 7.3+ м с нагрузкой 10п и ниже? (вне кода, звонок)
