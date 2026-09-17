# Spec: КП — прайс слева и договорная цена на «нет в прайсе»

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅ (focused pytest + vitest + typecheck; браузер не прогонялся)  
**Дата**: 2026-09-15  
**One-pager**: [../ideas/kp-unpriced-price-drawer.md](../ideas/kp-unpriced-price-drawer.md)  
**Plan**: [../develop/plans/2026-09-15-kp-unpriced-price-drawer.md](../develop/plans/2026-09-15-kp-unpriced-price-drawer.md)  
**Related**: [kp-source-image-queue-drawer.md](./kp-source-image-queue-drawer.md) (`Drawer side="left"`), [kp-row-edit-delete-icons.md](./kp-row-edit-delete-icons.md) (PATCH строки; ручная цена там была Out), [kp-unpriced-plates-replacement.md](./kp-unpriced-plates-replacement.md) (плиты — другой флоу), [guid-price-sync-loop.md](../ideas/guid-price-sync-loop.md), [price-desk-role.md](../ideas/price-desk-role.md)

## Objective

**Проблема.** На предпросмотре состава красные «нет в прайсе» (C110.30-6, C110.40-8.1) блокируют переход к клиенту. Менеджер не видит, есть ли марка в прайсе вообще и сколько стоят соседи. Ввести цену в это КП нельзя: карандаш меняет только qty/марку, заводской прайс пишется Excel-очередью / столом экономиста.

**Цель.** (1) Кнопка «Прайс» открывает слева справочник текущей группы с поиском — чтобы увидеть дырку и соседние цены. (2) На красной строке — договорная ₽/шт **только в этот черновик**. Каталог `pb.db` не трогаем.

**Пользователь:** менеджер на шаге 1 после «Список верен».

**Успех:** по красным строкам понятно, что марки нет; соседи с ценами видны; договорная разблокирует wizard; следующее КП с той же маркой снова без цены.

---

## ASSUMPTIONS I'M MAKING

1. **Типы.** UI в `KpGradedPreviewPanel` → сваи, мостовые, ФБС, марши, ступени. **Плиты вне скоупа** (`KpPlatePreviewPanel` + замена нагрузки).
2. **Кнопка** только если есть непрорасценённые строки текущего (не sealed) цикла. Не на шаге клиента / Result / сверке с фото.
3. **Drawer** — существующий `Drawer side="left"` (как исходные фото). Справочник **только чтение**. Ввод цены — в таблице состава, не в панели.
4. **Поиск** по всему прайсу группы. При открытии: пин «нет в этом КП» из draft + префилл поиска общим префиксом непрорасценённых марок (для скрина — `C110`), иначе пустой поиск и полный список (кап ниже).
5. **Цена в каталоге** — своя у каждой марки (и класса, если группа с классами). Не среднее по семейству.
6. **Договорная** пишется в `order_data` строки: `unit_price` + `price_source: "oneoff"`. Не INSERT/UPDATE прайс-таблиц.
7. **Гейт wizard** уже пропускает `unit_price is not None` (`collect_unpriced_positions`, `calculate_total_cost`, PDF/XLSX `_resolve_line_unit_price`). Новый lookup-path не нужен, если флаг и число живут на строке.
8. **Смена класса / повторный ingest / карандаш «как в списке»** гоняют `generate_preview` → договорная **сгорает**. Qty-only PATCH сохраняет её (как сейчас сохраняет любой `unit_price`).
9. **PDF/XLSX клиенту** — обычная колонка цены, без слова «договорная». Во внутреннем предпросмотре — подпись «договорная».
10. **Скидка КП** на шаге результата действует на договорную так же, как на прайсовую (`discount_factor` в `calculate_total_cost`). Отдельного исключения нет.
11. **Счёт / GUID** не меняем: КП уходит, блок счёта без GUID остаётся.
12. **Без новых npm/pip.** Коммиты — по просьбе. `./run+logs.sh` не убивать.

→ Correct me now or these are locked for PLAN.

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| **D-job** | Зачем панель | Доказать «марки нет» + показать соседние цены |
| **D-write** | Куда цена | Только строка этого КП (`oneoff`) |
| **D-catalog** | Прайс завода | Не писать |
| **D-ux** | Панель | Left Drawer, поиск, read-only |
| **D-entry** | Ввод | Поле на красной строке таблицы |
| **D-pdf** | Документ клиенту | Без пометки «договорная» |
| **D-discount** | Скидка КП | Действует и на договорную, как на прайс |
| **D-grade** | Смена класса | Сброс договорной, снова lookup |
| **D-plates** | Плиты | Не в этой фиче |
| **D-signal** | Очередь экономисту | Нет |

---

## User Stories

- Как **менеджер**, видя «нет в прайсе», жму **«Прайс»** и слева вижу, что C110.30-6 в прайсе нет, а C110.30-9 стоит 27 585,43.
- Как **менеджер**, ищу любую марку текущей группы и вижу **её** цену (и класс, если есть).
- Как **менеджер**, в красной строке ввожу договорную цену за штуку, сумма строки считается, переход к клиенту открыт.
- Как **менеджер**, в следующем КП та же марка снова без цены — я не записал её на завод.

---

## Tech Stack

| Слой | Стек |
|------|------|
| Frontend | React 19, TS, Vitest, `Drawer`, `KpGradedPreviewPanel`, TanStack Query |
| Backend | FastAPI, `commercial_draft_lifecycle.patch_order_line`, прайс-таблицы `pb.db` **только SELECT** |
| API | REST `/api/v1/commercial/...` |

Новых пакетов нет.

## Commands

```
# Frontend
cd frontend && npm run test -- src/features/commercial-offer
cd frontend && npm run typecheck

# Backend
pytest tests/test_commercial_draft_append.py tests/test_pile_price_import.py -q
# плюс новые: catalog list + PATCH unit_price oneoff

# Dev: не убивать ./run+logs.sh
```

## Project Structure

```
core/price_catalog_query.py              → SELECT по product_type + q (новый)
app/api/v1/endpoints/commercial.py       → GET price-catalog; PATCH line + unit_price
app/schemas/commercial.py                → CommercialDraftLinePatchRequest.unit_price
app/services/commercial_draft_lifecycle.py → patch oneoff / clear
frontend/src/features/commercial-offer/
  components/KpGradedPreviewPanel.tsx    → кнопка, поле, подпись «договорная»
  components/PriceCatalogDrawer.tsx      → Drawer left + поиск + таблица
  api/commercialOfferApi.ts              → getPriceCatalog; patch unit_price
  lib/build*PreviewRows.ts               → прокинуть price_source
tests/                                   → catalog query + PATCH oneoff + гейт
```

## Code Style

Слои: роутер тонкий, SELECT каталога в `core/`, PATCH — в lifecycle рядом с qty. Русские `detail`. Минимальный diff.

Пример кнопки:

```tsx
{hasUnpricedRows ? (
  <Button type="button" variant="secondary" onClick={() => setCatalogOpen(true)}>
    Прайс
  </Button>
) : null}
```

Пример PATCH (договорная **или** сброс; не вместе с `source_text`):

```json
{ "unit_price": 27585.43 }
{ "unit_price": null }
{ "qty": 2, "unit_price": 27585.43 }
```

На строке после успеха: `"price_source": "oneoff"`. Сброс (`null`) снимает флаг и возвращает «нет в прайсе», если lookup снова пуст.

## Design

### Справочник

`GET /api/v1/commercial/price-catalog?product_type=piles&q=C110`

- Auth как у остальных commercial (менеджер, владение не требуется — это справочник).
- `product_type`: `piles` | `bridge_piles` | `fbs` | `marches` | `steps`.
- `q`: подстрока марки после нормализации C↔С / пробелы; пустой = все строки группы.
- Ответ: `{ items: [{ mark, concrete_grade | null, price }] }`. Ступени — `concrete_grade: null`.
- Кап: 2000 строк; если больше — 400 «уточните поиск». Сваи × 5 классов ≈ 1590 — проходит.
- **Не** отдаём цены ≤ 0. **Не** пишем в БД.

UI Drawer:

1. Заголовок «Прайс».
2. Поиск.
3. Блок «В этом КП нет в прайсе» — марки из draft с `unit_price === null` (даже если их нет в ответе каталога). Это доказательство дырки.
4. Таблица совпадений: марка, класс (если есть), цена. Подсветка строк, чей stem совпал с непрорасценёнными.

Панель не подставляет цену в строку и не меняет марку.

### Договорная цена

Расширить `PATCH /drafts/{id}/lines/{id}`:

- Новое опциональное поле `unit_price: float | null`.
- `> 0` и конечное; потолок 10_000_000 — иначе 400.
- Разрешено только если сейчас `unit_price is None` **или** уже `price_source == "oneoff"` (правка/сброс).
- Если у строки уже есть прайсовая цена (`unit_price is not None` и флаг не `oneoff`) → 400 «Цена из прайса. Для изменения используйте скидку.»
- `source_text` в том же запросе → 400 (замена марки и так пересоберёт preview).
- Sealed-строка → 400.
- Успех: `unit_price`, `line_total = unit_price * qty`, `price_source = "oneoff"`. Totals через существующий `persist_order_and_metadata` / `compute_totals`.
- `unit_price: null` на oneoff-строке: снять цену и флаг.

Предпросмотр: вместо «нет в прайсе» — число + серый текст «договорная». Красвый алерт панели **пропадает**, когда непрорасценённых не осталось.

Шаг 3: число как у остальных; подпись «договорная» если флаг есть. Кнопки «Прайс» нет.

### Что не ломаем

- `collect_unpriced_positions` / PDF / XLSX уже берут `unit_price` со строки — oneoff проходит без правок lookup, **если** число осталось на линии.
- Карандаш qty/марки и мусорка — как сейчас.
- Плиты, wide, invalid width — не трогаем.
- `generate_preview` после класса/ingest **не** восстанавливает oneoff (D-grade).

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit catalog | `q=C110` на seeded `pile_prices` → C110.30-9 есть, C110.30-6 нет; пустой q не падает на фикстуре |
| API PATCH | oneoff на null-строке → гейт пуст, `price_source=oneoff`; повторная правка; сброс `null`; отказ на прайсовой строке; отказ с `source_text`; qty+oneoff сохраняет оба |
| Wizard/service | после oneoff `unpriced_position_labels` пуст; calculate не бросает `UnpricedPlatesError` |
| RTL панель | нет кнопки без красных строк; кнопка открывает Drawer; пин отсутствующей марки; поле на красной строке; после oneoff — «договорная», алерт снят |
| RTL Drawer | поиск фильтрует; Esc/оверлей закрывают |
| Regress | PATCH qty без `unit_price` как сейчас; карандаш марки; плиты unpriced |

## Boundaries

- **Always:** SELECT-only в каталог; oneoff только на дырках; русские ошибки; тесты на гейт и отказ писать в прайс.
- **Ask first:** плиты; пометка на PDF; автоподстановка цены соседа; сигнал экономисту; писать в `pile_prices`.
- **Never:** новые зависимости; блок счёта без GUID снимать; commit без просьбы; трогать ГСМ / production layout.

## Success Criteria

| # | Критерий |
|---|----------|
| S1 | Есть красные строки → кнопка «Прайс» на предпросмотре простого типа |
| S2 | Drawer слева: поиск по группе; пин C110.30-6 как отсутствующей; C110.30-9 с ценой |
| S3 | Поиск находит марку не из семейства текущего КП |
| S4 | Договорная на красной строке → число на линии, wizard к клиенту, алерт «Нет цен…» снят |
| S5 | `pb.db` прайс-таблицы без INSERT/UPDATE после S4 |
| S6 | Новое КП с той же маркой снова «нет в прайсе» |
| S7 | Прайсовую строку PATCH unit_price не берёт (400) |
| S8 | PDF/XLSX печатают введённое число (уже из `unit_price` строки) |
| S9 | Плиты: кнопки «Прайс» нет, замена нагрузки as-is |
| S10 | Focused pytest + vitest панели/drawer + typecheck зелёные |

## Out of Scope

- Запись в заводской прайс / роль economist / Excel 💰
- Клик по соседу = смена марки
- Подставить цену соседа в другую марку
- Плиты
- «Договорная» на PDF клиенту
- Журнал спроса для завода
- Result-шаг: кнопка «Прайс»

## Open Questions

_Нет. Скидка на договорную — как на прайс (подтверждено 2026-09-15). IMPLEMENT: focused тесты зелёные; браузерный прогон C110 не делался._

## Risks

| Риск | Почему | Смягчение |
|------|--------|-----------|
| Договорная потеряется | `generate_preview` после класса/ingest | Зафиксировано: сброс; qty-only безопасен |
| Обход стола экономиста | Менеджер ставит любую сумму | Только дырки; следующее КП снова пустое; не пишем в БД прайса |
| Счёт без GUID | Марки нет в 1С | Не в скоупе; КП считать можно и сейчас |
| Каталог 1590 строк | Тяжёлый Drawer | Кап 2000; префилл поиска семейством |
