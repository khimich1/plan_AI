# Implementation Plan: Себестоимость цельных свай

> **Идея:** [`ai_docs/ideas/pile-cost-calculation.md`](../../ideas/pile-cost-calculation.md)
> **ТЗ:** [`ai_docs/specs/pile-cost-calculation.md`](../../specs/pile-cost-calculation.md)
> **Дата:** 2026-10-08
> **Статус:** план на ревью, код не начат

## Overview

Экономист на `/prices` меняет цены сырья, рецепт бетона и четыре статьи расходов.
Программа считает цену цельной сваи для покупателя. После диффа и подтверждения
числа пишутся в `pile_prices`. Новые КП берут их при расчёте. Файл «цельные сваи»
через drop-zone цены больше не записывает.

Норма (объём, сечение, килограммы металла, редкая добавка числом) — снимок
файла от 06.10.2026 в репозитории. `С130.35-13А**` в снимок не входит.

## Architecture Decisions

- Формула живёт в `core/pile_cost.py` и не знает ни Excel, ни SQLite.
- Щебень рецепта: `B30_granite` по цене 1,97, остальные классы по 1,80.
- Добавка = 0,6% цемента, кроме марок, где в снимке лежит своё число килограммов.
- Металл снимка — те ячейки, которые лист «Стоимость» умножает на цену,
  включая ссылку на соседнюю строку, если так посчитан прайс.
- Ручки хранятся в `pb.db`. Норма — JSON, API её не пишет.
- Превью не пишет `pile_prices`. Apply пишет только если `inputs_sha256`
  совпал с сохранёнными ручками. Пропавшие марки не удаляем.
- Порог «цена изменилась» — `abs(a - b) > 0.005`, как на столе прайсов.
- `parse_pile_price_rows_from_xlsx` не удаляем. Стол через него больше не пишет.

## Dependencies

```
PC-001 ─┬─ PC-002 ─┐
        └─ PC-003 ─┴─ PC-004 ─ PC-005 ─ PC-006 ─┬─ PC-007 ─ PC-008
                                                └─ PC-009
PC-010 (ручная приёмка, после PC-008)
```

PC-002 и PC-003 можно делать параллельно после PC-001.

## Task List

### Phase 1: Формула

- [ ] **PC-001: Чистая формула цены**
  - **Description:** `buyer_price` по спеке: кг из объёма и рецепта, добавка
    0,6% или число из нормы, щебень гранита по своей цене, металл из строк
    нормы, расходы `(сумма четырёх статей) × 1,03`, затем ×1,2 ×1,22 и ×1,02
    для сечений 30/35/40. Отрицательный объём и цена ≤ 0 — ошибка.
  - **Acceptance:** С30.20-3 / B15 = 3513.2818728 ± 0.01 и разложение
    1431.4179 / 968.3648 / 2399.7827. С30.20-3 / B30_granite = 3655.4166552 ± 0.01
    без 1,02. С40.30-1 / B15 = 9033.108758784 ± 0.01. Рецепт B15 не двигает B20.
    Цена цемента не двигает рубли металла. Цена Ø10 не двигает марку только с Ø12.
  - **Verify:** `pytest tests/test_pile_cost_formula.py -q`
  - **Dependencies:** None
  - **Files:** `core/pile_cost.py`, `tests/test_pile_cost_formula.py`
  - **Scope:** M

**Checkpoint 1:** золотые числа спеки зелёные, базы и экрана ещё нет.

### Phase 2: Снимок и ручки

- [ ] **PC-002: Снимок нормы**
  - **Description:** Один прогон локального xlsx от 06.10.2026 в
    `core/pile_cost_norm.json`. В строке: марка, класс, сечение, объём,
    строки `{article, kg}`, и килограммы добавки только там, где ячейка
    не формула `цемент × 0,006`. Марку `С130.35-13А**` не писать.
    Скрипт чтения xlsx не вызывается из API.
  - **Acceptance:** по снимку формула даёт цены листа «Прайс» для С30.20-3
    (пять классов) и С40.30-1 / B15. Строк `С130.35-13А**` в JSON нет.
    С30.20-3 / B15 в снимке: объём 0,13, спираль 3,1, поперечная 3,15, рабочая 8.
  - **Verify:** `pytest tests/test_pile_cost_norm.py -q`
  - **Dependencies:** PC-001
  - **Files:** `scripts/extract_pile_cost_norm.py`, `core/pile_cost_norm.json`,
    `tests/test_pile_cost_norm.py`
  - **Scope:** M

- [ ] **PC-003: Хранение ручек**
  - **Description:** В `pb.db` таблицы цен статей, рецепта по `GRADE_CODES`
    и четырёх статей расходов. Пустая база при первом чтении заполняется
    числами 06.10.2026 из спеки. Запись ручек не трогает `pile_prices`.
    Цена ≤ 0 и рецепт < 0 отклоняются.
  - **Acceptance:** повторное чтение отдаёт записанный цемент; `pile_prices`
    после записи ручек не меняется.
  - **Verify:** `pytest tests/test_pile_cost_inputs.py -q`
  - **Dependencies:** PC-001
  - **Files:** `core/pile_cost_inputs.py`, `tests/test_pile_cost_inputs.py`
  - **Scope:** M

**Checkpoint 2:** снимок и стартовые ручки дают те же две марки, что PC-001.
Экрана нет.

### Phase 3: Запись в прайс

- [ ] **PC-004: Превью и запись**
  - **Description:** Считать все марки снимка. Дифф к `pile_prices`:
    changed / new / missing / unchanged, до 30 примеров, порог 0.005.
    Preview не пишет. Apply принимает `inputs_sha256`; не совпал — цены
    не меняются. Совпал — `INSERT OR REPLACE` только ключей нормы.
    Missing не удалять. `price_list_date` — дата записи.
  - **Acceptance:** preview даёт 0 изменений `pile_prices`. Чужой sha256
    не пишет. Apply пяти цен С30.20-3 виден в `pile_prices`. Ключ без нормы
    остаётся со старой ценой.
  - **Verify:** `pytest tests/test_pile_cost_service.py -q`
  - **Dependencies:** PC-002, PC-003
  - **Files:** `app/schemas/pile_cost.py`, `app/services/pile_cost_service.py`,
    `tests/test_pile_cost_service.py`
  - **Scope:** M

- [ ] **PC-005: Файл цельных свай больше не источник**
  - **Description:** В `price_desk_service` kind `pile` на preview и apply
    отвечает ошибкой «цены цельных свай считаются на этом экране» и не
    вызывает `import_pile_prices_from_xlsx`. Остальные kind без изменений.
  - **Acceptance:** тестовый xlsx «цельные» не добавляет строк в `pile_prices`.
    Превью плит/ФБС по-прежнему разбирает файл.
  - **Verify:** `pytest tests/test_price_desk_service.py tests/test_pile_price_import.py -q`
  - **Dependencies:** PC-004
  - **Files:** `app/services/price_desk_service.py`,
    `tests/test_price_desk_service.py`
  - **Scope:** S

- [ ] **PC-006: HTTP**
  - **Description:** `GET` и `PUT /api/v1/prices/pile-cost`,
    `POST .../preview`, `POST .../apply`. `REQUIRE_PRICES`.
    Менеджер — 403. PUT не пишет `pile_prices`. Apply без sha256 — 409.
  - **Acceptance:** критерии API из спеки.
  - **Verify:** `pytest tests/test_pile_cost_api.py tests/test_price_desk_api.py -q`
  - **Dependencies:** PC-004, PC-005
  - **Files:** `app/api/v1/endpoints/price_desk.py`,
    `app/dependencies/services.py`, `tests/test_pile_cost_api.py`
  - **Scope:** M

- [ ] **PC-009: Новое КП берёт записанную цену**
  - **Description:** После apply новое КП на С30.20-3 / B15 получает новую
    цену. Строка уже сохранённого КП хранит старый `unit_price`.
  - **Acceptance:** оба утверждения в тесте.
  - **Verify:** `pytest tests/test_pile_cost_api.py tests/test_commercial_pile_pricing.py -q`
  - **Dependencies:** PC-006
  - **Files:** `tests/test_pile_cost_kp.py` или соседний тест КП
  - **Scope:** S

**Checkpoint 3:** превью не меняет прайс, apply меняет, файл «цельн» получает отказ,
сохранённое КП не пересчитано.

### Phase 4: Экран

- [ ] **PC-007: Клиент API**
  - **Description:** Типы синхронны с `app/schemas/pile_cost.py`. Хуки чтения,
    записи ручек, превью и apply.
  - **Acceptance:** `tsc` без ошибок в новых файлах.
  - **Verify:** `cd frontend && npm run typecheck`
  - **Dependencies:** PC-006
  - **Files:** `frontend/src/features/price-desk/types/pileCost.ts`,
    `frontend/src/features/price-desk/api/`, hooks
  - **Scope:** S

- [ ] **PC-008: Блок на `/prices`**
  - **Description:** Над drop-zone блок «Цельные сваи»: три группы полей,
    константы текстом, поиск марки, карточка пяти классов с разложением.
    Пока цены не записаны — плашка и «Показать, что изменится». «Записать
    в прайс» шлёт sha256 из превью. Drop-zone остальных групп остаётся.
  - **Acceptance:** vitest: правка цены цемента меняет карточку; «Записать»
    не вызывается без превью; apply уходит с sha256.
  - **Verify:** `cd frontend && npm run test -- --run src/features/price-desk && npm run typecheck`
  - **Dependencies:** PC-007
  - **Files:** `frontend/src/features/price-desk/components/`,
    `frontend/src/pages/prices/PricesPage.tsx`
  - **Scope:** M

**Checkpoint 4:** экономист меняет цемент, видит С30.20-3, подтверждает,
в базе новая цена. Менеджер блок не вызывает (403 на API).

### Phase 5: Приёмка

- [ ] **PC-010: Сверка с файлом 06.10.2026** (не CI)
  - **Description:** Локальный xlsx против снимка и формулы. В отчёте число
    совпавших цен и список расхождений. `С130.35-13А**` в списке отсутствует
    и в отчёт не входит как ошибка.
  - **Acceptance:** расхождений нет, кроме явно исключённой марки.
  - **Verify:** ручной прогон скрипта PC-002 в режиме сверки
  - **Dependencies:** PC-008
  - **Files:** нет продуктовых
  - **Scope:** S

## Risks

| Риск | Что делать |
|------|------------|
| Снимок повторит сдвиг строки Excel не там | PC-002 сверяет снимок с листом «Прайс» до коммита JSON |
| Звёздочка в имени | `С130.35-13А**` выкинуть строкой. `А*` и `С80.40-9 *` оставить |
| Две двери в `pile_prices` | PC-005 сразу после сервиса записи, не после экрана |
| Ручки потеряются при рестарте | PC-003 пишет их в `pb.db`, не в память процесса |
| Полный список 1 500 разложений на каждый GET | GET списка отдаёт марку и пять цен; разложение — по одной марке |

## Open Questions

Нет.

## Not doing

Правка нормы в UI, повторный импорт xlsx, движок Excel, коэффициенты 1,45–2,
плиты и остальные группы, удаление пропавших марок, округление до рубля,
пересчёт сохранённых КП, марка `С130.35-13А**`.
