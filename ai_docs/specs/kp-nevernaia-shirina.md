# Spec: Неверная ширина в составе КП

> **Источник идеи:** [`ai_docs/ideas/kp-nevernaia-shirina.md`](../ideas/kp-nevernaia-shirina.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS/IMPLEMENT ✅  
> **План:** [`ai_docs/develop/plans/2026-09-02-kp-nevernaia-shirina.md`](../develop/plans/2026-09-02-kp-nevernaia-shirina.md)  
> **Дата:** 2026-09-02  
> **Статус:** реализовано (2026-09-03), коммит не делался  
> **Связанные:** `UnpricedPlatesInlineSection`, `WidePlatesInlineSection`, `commercial_plate_resolve`, `parse_pb_width_to_m`, `kp-unpriced-plates-replacement.md`

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| 1 | Кто / где | Менеджер на шаге плит, экран «Состав КП (предпросмотр)» — до расчёта и сохранения |
| 2 | Что ловим | Любая разобранная ширина **≤ 12 дм**, мм которой **не** входят в таблицу резов |
| 3 | Таблица резов (мм) | 260–320, 460–530, 660–720, 860–920, 1020–1080, **1200** (точка) |
| 4 | Запись | Не regex «-8-». Парсер → `width_m` → `round(width_m * 1000)` мм |
| 5 | Жёсткость | Нельзя идти дальше. Только **заменить** на предложенную или **исключить**. Нет «оставить как есть» |
| 6 | Предложения | Одно правило: край диапазона **снизу** и **сверху**. 800 → 720 и 860 (марки 7,2 и 8,6). 1000 → 920 и 1020 (9,2 и 10,2). Нет особой клички 10→10,8 |
| 7 | Один сосед | Если снизу или сверху диапазона нет — одна кнопка (200 мм → только 2,6) |
| 8 | 0,3 м | 300 мм ∈ 260–320 — **не** ошибка (`-0.3-`, `-3-`, `-3,0-`) |
| 9 | 0,2 м | 200 мм < 260 — ошибка, замена **2,6** |
| 10 | Шире 12 дм | Только существующий гейт wide-plates (там можно confirm). Этот гейт их не видит |
| 11 | UI | Карточка под таблицей, зеркало unpriced. Строка в таблице — красный маркер «решение ниже» |
| 12 | Карандаш | Правка имени + «Список верен» пересобирает preview; если мм в диапазоне — гейт пуст |
| 13 | Цена в карточке | Показать, если `lookup_plate_price` нашёл цену > 0; нет цены — всё равно предлагаем марку (дальше может сработать unpriced) |

---

## Assumptions

1. `width_mm = int(round(float(width_m) * 1000))`. Попадание — включительные границы диапазонов; 1200 только точное равенство после округления.
2. Продукт только `plates`. Сваи / ступени / марши / ФБС не трогаем.
3. Детекция на тех же строках, что уже попали в preview/`order_data` (после парсинга). Неразобранные строки — по-прежнему unparsed, не этот гейт.
4. Порядок гейтов: **wide → invalid width → unpriced**. Сначала «15», потом «8», потом «нет в прайсе».
5. После «Применить» — тот же путь, что unpriced: переписать `normalized_text` → `generate_preview` → штатный пересчёт.
6. Константа диапазонов живёт в `core/` (один модуль). `plate_lists.add_items` / snap раскладки **не меняем** в этом релизе.
7. Пустой список плит после exclude всех проблемных + не осталось других строк → ошибка как у unpriced («список стал пустым»). Если остались валидные — идём дальше.

→ Если пункт неверен — сказать до PLAN.

---

## Objective

**Проблема.** Марка вроде `Плиты ПБ 29-8-8п` проходит в КП. Завод такую ширину не режет. Цена сейчас ещё и выше полной 1,2 м (на скрине 29-8 = 11 248 ₽, 29-12 = 9 914 ₽). В раскладке 800 мм притягивается к 720 мм — состав КП и производство расходятся.

**Цель.** Менеджер не может сохранить / посчитать КП, пока каждая ширина вне таблицы не заменена на соседний заводской рез или строка не снята.

**Пользователь:** менеджер, собирающий КП из OCR или текста.

**Успех:** в сохранённом КП нет плит с `width_mm` вне таблицы (кроме wide, прошедших свой гейт). На «восьмёрке» менеджер явно выбирает 7,2 или 8,6.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | менеджер на предпросмотре | видеть красным `ПБ 29-8-8п` | не отправить клиенту нерезабельную ширину |
| US-2 | менеджер | в карточке выбрать 7,2 или 8,6 (с ценой, если есть) | поставить заводскую марку в 2 клика |
| US-3 | менеджер | исключить строку | не тащить в КП то, чего не будет |
| US-4 | менеджер | поправить карандашом `29-8` → `29-8,6` и нажать «Список верен» | обойтись без карточки, если сам исправил |
| US-5 | менеджер с `-10-` | тот же гейт и соседи 9,2 / 10,2 | одно правило, без исключений |

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`, `pb.db`) |
| Domain | новый модуль в `core/` + детект в preview; resolve через `CommercialPlateResolve` |
| API | новый `POST /api/v1/commercial/drafts/{draft_id}/invalid-widths/resolve` |
| Frontend | React + TS, карточка на `PlateInputStep`, маркер в `KpPlatePreviewPanel` / `plateLineHighlights` |
| Тесты | pytest `tests/`; vitest рядом с unpriced/wide |

---

## Commands

```bash
# Backend (корень, venv)
pytest tests/ -q -k "invalid_width or factory_width or commercial_plate"

# Регрессия
pytest tests/ -q

# Frontend
cd frontend && npm run test -- --run src/features/commercial-offer
cd frontend && npm run typecheck
cd frontend && npm run build
```

---

## Project Structure

```
core/factory_width.py                 → НОВЫЙ: диапазоны, is_factory_width_mm, suggest, rewrite марки
core/config/constants.py              → parse_pb_width_to_m (без смены правил)
core/unpriced_plate_replacements.py   → образец rewrite + список замен
app/services/commercial_service.py    → детект после парсинга, metadata
app/services/commercial_plate_resolve.py → третий PlateResolveSpec kind="invalid_width"
app/services/commercial_wizard_step_service.py → ERR_ + next_required_action
app/services/commercial_calculation_service.py → blocking
app/schemas/commercial.py             → lines, action, WizardNextRequiredAction
app/schemas/errors.py                 → ERR_INVALID_WIDTHS
app/api/v1/endpoints/commercial.py    → resolve endpoint
frontend/.../UnpricedPlatesInlineSection.tsx  → образец карточки
frontend/.../InvalidWidthsInlineSection.tsx   → НОВЫЙ
frontend/.../plateLineHighlights.ts          → kind invalid_width
frontend/.../KpPlatePreviewPanel.tsx         → алерт + красные строки
frontend/.../wizardDraftStore.tsx            → invalidWidthActions
tests/test_factory_width.py                  → НОВЫЙ unit
tests/test_invalid_width_resolve.py          → НОВЫЙ service
```

Имена файлов при реализации можно сдвинуть на 1, если рядом уже есть лучшее место — логика та же.

---

## Code Style

Слои: роутер → сервис → `core/`. Metadata-пара: `invalid_width_lines` + `invalid_widths_resolved` (как wide / unpriced).

Действия API: `replace_width` \| `exclude`. Цель замены — `width_mm` **только из `replacements` этой строки**.

```python
FACTORY_WIDTH_RANGES_MM: tuple[tuple[int, int], ...] = (
    (260, 320),
    (460, 530),
    (660, 720),
    (860, 920),
    (1020, 1080),
    (1200, 1200),
)

def is_factory_width_mm(width_mm: int) -> bool:
    return any(lo <= width_mm <= hi for lo, hi in FACTORY_WIDTH_RANGES_MM)

def suggest_factory_width_mm(width_mm: int) -> list[int]:
    if is_factory_width_mm(width_mm):
        return []
    lower = [hi for lo, hi in FACTORY_WIDTH_RANGES_MM if hi < width_mm]
    upper = [lo for lo, hi in FACTORY_WIDTH_RANGES_MM if lo > width_mm]
    out: list[int] = []
    if lower:
        out.append(max(lower))
    if upper:
        out.append(min(upper))
    return out
```

Марка: переписать только часть W в `L-W-N`, формат ширины как в `format_plate_name` / `config_and_data` (720 → `7,2`, 860 → `8,6`, 1200 → `12`). Длина, нагрузка, qty, префикс «Плиты » — без изменений.

---

## Design

### Детекция

После парсинга preview (до или вместе с unpriced, но **после** классификации wide):

Для каждой плиты с `width_m`:

1. Если строка в `wide_plate_lines` и ещё не resolved как split/confirm — **пропуск** (wide гейт).
2. `width_mm = round(width_m * 1000)`.
3. Если `is_factory_width_mm` — ок.
4. Иначе строка в `invalid_width_lines`: `id`, `line`/`name`, `qty`, `width_mm`, `replacements: [{width_mm, width_label, price?}]`.

`invalid_widths_resolved = not bool(invalid_width_lines)`.

Примеры:

| Вход | мм | Гейт | Замены (label) |
|------|----|------|----------------|
| `ПБ 29-8-8п` | 800 | да | 7,2 (720), 8,6 (860) |
| `ПБ 29-8,0-8п` | 800 | да | то же |
| `ПБ 60-10-8п` | 1000 | да | 9,2 (920), 10,2 (1020) |
| `ПБ 60-12-8п` | 1200 | нет | — |
| `ПБ 78-0.3-8п` / `ПБ 78-3-8п` | 300 | нет | — |
| `ПБ 78-0.2-8п` / `ПБ 78-2-8п` | 200 | да | 2,6 (260) |
| `ПБ 60-4-8п` | 400 | да | 3,2 (320), 4,6 (460) |
| `ПБ 60-11-8п` | 1100 | да | 10,8 (1080), 12 (1200) |
| `ПБ 60-15-8п` | 1500 | нет (wide) | wide-flow |

### Wizard

- `ERR_INVALID_WIDTHS` в `app/schemas/errors.py`.
- `WizardNextRequiredAction.resolve_invalid_widths`.
- `invalid_width_lines_blocking` — как unpriced: есть lines и не resolved → нельзя на client/result, `calculate` падает понятной ошибкой.
- Сообщение: «Нестандартная ширина: замените на заводской рез или исключите позицию».

### API

`POST /api/v1/commercial/drafts/{id}/invalid-widths/resolve`

Тело: `{ decisions: [{ line_id, action, width_mm? }] }`

- `replace_width` + `width_mm` из `replacements` → перепись W в строке.
- `exclude` → удалить строку.
- Не все строки решены → 400.
- `width_mm` не из списка → 400.
- Пустой заказ после exclude → 400.

Решения в `metadata["invalid_width_decisions"]`. После успеха preview заново, флаг resolved по новой детекции (если карандаш потом снова введёт 8 — гейт вернётся).

### Frontend

`InvalidWidthsInlineSection` на `PlateInputStep` под таблицей / рядом с unpriced:

- Заголовок: «Нестандартная ширина».
- Подзаголовок: завод такую ширину не режет — выберите рез или исключите.
- На строку: исходная марка, радио замен («8,6 — 10 400 ₽» / «7,2 — …»), пункт «Исключить позицию». Предвыбор: **верхний** сосед (для 8 это 8,6 — ближе к заказанной «восьмёрке», чем 7,2).
- «Применить» неактивна, пока не по всем строкам есть действие.

Таблица предпросмотра: kind `invalid_width` в `plateLineHighlights` (фон как `wide`, `#fef3f2`). Алерт над таблицей: N позиций с шириной вне таблицы резов.

Карандаш/корзина без новых диалогов. После «Список верен» сервер пересобирает lines.

### Порядок на шаге плит

1. Wide card (если есть).
2. Invalid width card (если есть).
3. Unpriced card (если есть).

Можно не нажимать три кнопки «Применить» в одном кадре, если гейты последовательны (wide resolved → появляется invalid → затем unpriced). Допустимо показать invalid сразу, если wide пуст.

---

## Testing Strategy

**pytest**

- `is_factory_width_mm`: 300, 720, 860, 1080, 1200 — true; 200, 400, 600, 800, 1000, 1100, 1190 — false.
- `suggest_factory_width_mm(800) == [720, 860]`; `(1000) == [920, 1020]`; `(200) == [260]`; `(1100) == [1080, 1200]`; in-range → `[]`.
- Rewrite: `Плиты ПБ 29-8-8п` + 860 → `Плиты ПБ 29-8,6-8п`; qty и нагрузка на месте; `0.3` не рерайтится.
- Детект preview: заказ из скрина → 3 invalid (29-8, 32-8, 36-8), 12-е ширины не в списке; `78-0.3` нет в списке; `60-15` нет в этом списке (wide).
- Resolve: replace 29-8 → 8,6 → в новом draft нет invalid по этой строке, имя с `8,6`.
- exclude последней единственной плиты → ошибка пустого списка.
- `width_mm=800` в replace, когда replacements только 720/860 → 400.
- Wizard: пока не resolved — `next_required_action = resolve_invalid_widths`, calculate недоступен.
- Регрессия wide + unpriced: заказ только с 15 дм или только без цены — поведение как сейчас.

**vitest**

- Карточка: две замены + exclude; Apply disabled без выбора.
- Подсветка строки 29-8-8п kind `invalid_width`.
- 0,3 / 12 не подсвечиваются этим kind.

**Не в scope тестов:** браузерный прогон всего wizard (ручная проверка по Success Criteria).

---

## Boundaries

- **Always:** минимальный diff; Pydantic на новые поля; тесты на детекцию и resolve до объявления готовым; пары metadata `*_lines` / `*_resolved`.
- **Ask first:** менять snap/`add_items`/раскладку; расширять таблицу резов; клички `10→10,8`; трогать wide confirm; миграция БД; feature-flag.
- **Never:** «confirm as-is» для внедиапазонной ширины; regex только на `-8-`; коммит без просьбы; ГСМ / производство / отгрузки.

---

## Success Criteria

- [x] Заказ со скрина (смесь 12 и 8): три строки `-8-` в карточке с 7,2 и 8,6; двенадцатье — нет. Расчёт недоступен. — pytest `test_invalid_width_preview` + `test_invalid_width_wizard`; vitest карточки/подсветки.
- [x] Выбор 8,6 → «Применить» → в составе `ПБ 29-8,6-8п` (и аналоги), гейт пуст, цена не выше полной 1,2 м из‑за «ложной восьмёрки» как 800 мм с резом. — pytest `test_invalid_width_resolve`.
- [x] `ПБ 60-10-8п` → гейт, замены 9,2 и 10,2. — `suggest_factory_width_mm(1000)==[920,1020]`.
- [x] `ПБ 78-0.3-8п` и `ПБ 78-3-8п` — без гейта.
- [x] `ПБ 60-15-8п` — только wide, не эта карточка.
- [x] Карандаш `29-8` → `29-8,6` + «Список верен» — карточки нет. — тот же `generate_preview`; отдельного браузерного e2e нет (нет chrome-devtools MCP).
- [x] Exclude единственной неверной при наличии других валидных — заказ жив.
- [ ] Полный `pytest tests/ -q` — 8 падений **вне фичи** (admin reset, offer identity, capacity gate, plate audit, generate-files schema / MagicMock). Целевые `factory_width` / `invalid_width` / wizard / resolve — зелёные.
- [x] `cd frontend && npm run test -- --run src/features/commercial-offer` + `npm run typecheck` + `npm run build` — зелёные.
- [ ] Полный `npm run test -- --run` — 1 падение `GsmPage` (`useAuth` без `AuthProvider`), не этот diff.

---

## Open Questions

Нет блокирующих. После пилота: не душит ли `-10-` менеджеров (тогда отдельные клички, Ask first).

---

## Not in this spec

Задачи реализации — в плане. Код написан 2026-09-03 (IW-001…009). Коммит — только по просьбе.
