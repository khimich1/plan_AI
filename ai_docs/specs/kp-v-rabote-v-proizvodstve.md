# Spec: КП в работе (вкладка в производстве)

> **Источник идеи:** [`ai_docs/ideas/kp-v-rabote-v-proizvodstve.md`](../ideas/kp-v-rabote-v-proizvodstve.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS/IMPLEMENT ✅  
> **План:** [`ai_docs/develop/plans/2026-09-02-kp-v-rabote-v-proizvodstve.md`](../develop/plans/2026-09-02-kp-v-rabote-v-proizvodstve.md)  
> **Дата:** 2026-09-02  
> **Статус:** ✅ implemented (MVP)  
> **Связанные:** `ProductionPage`, `ProductionTabs`, `GET /api/v1/production/kp-candidates`, `list_kps_in_production`, `SgpProgress`

---

## Decisions locked (ревью UX)

| # | Тема | Решение |
|---|------|---------|
| 1 | Порядок вкладок | Календарный план → Планы → **КП в работе** → СГП → Производственный календарь |
| 2 | Клик по КП | Строка **раскрывается на месте** (плиты под ней). Не drawer, не отдельный роут |
| 3 | Сортировка | По сроку: горящие и просроченные сверху |
| 4 | КП 100% в плане | В раскрытии — список этих плит с пометкой «в плане, ждёт отливки» |

## Assumptions (остальное)

1. Query-параметр: `scope=in_work` для вкладки. Без параметра — поведение мастера **как сейчас**. Неизвестный `scope` → 422.
2. id вкладки: `in-work`. URL: `/production?tab=in-work`. Подпись: **«КП в работе»**.
3. Новые поля **аддитивны** на обоих срезах (мастер их игнорирует): `remaining_qty`, `in_plan_qty`, `on_sgp_qty`. «Осталось шт» в UI = `remaining_qty + in_plan_qty`.
4. `on_sgp_qty` = `SUM(completed_plates.qty) WHERE kp_id = ?` (как `SgpProgress.n`). Свободные плиты СГП с `kp_id IS NULL` в это КП не входят.
5. Во вкладке в раскрытии `plates[]` — плиты **не на СГП**: статус `в производстве` **и** `в плане`. У каждой строки `bucket: "awaiting_plan" | "in_plan"`. В срезе мастера `plates[]` по-прежнему только `в производстве`.
6. Одновременно раскрыта **одна** строка (аккордеон). Повторный клик закрывает.
7. Сортировка на сервере: `execution_terms` как дата (`ДД.ММ.ГГГГ` или ISO); неразобранная строка — в конец; тай-брейк `kp_id` по возрастанию.
8. Роли: `admin` и `production`. Архив и `/api/v1/offers` не трогаем.
9. Нет нового эндпоинта, нет миграции БД, нет feature-flag, нет сумм/`total_amount` в ответе.
10. Пустой список вкладки: «Все плиты на СГП — смотрите склад готовой продукции».

---

## Objective

Дать цеху и админу **очередь работы по плитам** внутри `/production`: какие КП ещё не на СГП, сколько ждут дорожек, сколько уже в календаре, сколько на складе.

Это не копия архива «В производстве». Архив — коммерция (суммы, PDF, скидка). Здесь — остаток отливки.

Мастер «Начать планирование» **не меняет состав**: по-прежнему только то, что ещё можно класть на дорожки.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | роль `production` | вкладку «КП в работе» без архива | видеть очередь, не заходя в коммерцию |
| US-2 | админ на `/production` | ту же вкладку | не прыгать в архив за остатком плит |
| US-3 | мастер цеха | в строке КП, клиент, срок, осталось / в плане / на СГП | понять, что горит и что уже разложено |
| US-4 | мастер цеха | раскрыть строку и увидеть марки/размеры/нагрузку/шт не на СГП | понять, что лить, без цен и PDF |
| US-5 | планировщик в мастере | прежний список кандидатов | не предлагать разложить то, что уже на дорожках |

### Пример (инвариант срезов)

КП №47, заказ 10 шт. 4 «в производстве», 6 «в плане», 0 на СГП.

- Мастер (`GET …/kp-candidates`): КП в списке, `plates` на 4 шт.
- Вкладка (`?scope=in_work`): КП в списке, осталось 10, в плане 6, на СГП 0; в раскрытии 4 awaiting_plan + 6 in_plan.

Когда все 10 в плане и 0 на СГП: мастер **не** показывает КП; вкладка **показывает**.

Когда все 10 на СГП: ни мастер, ни вкладка.

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | FastAPI, Pydantic v2, SQLite (`plita.db`) |
| API | `app/api/v1/endpoints/production.py` — существующий `GET /kp-candidates` |
| Domain | `app/services/production_service.py`, `app/repositories/kp_repository.py`, qty СГП как в `SgpService._sgp_progress_on_cursor` |
| Frontend | React 19, Vite, TanStack Query, вкладка в `ProductionTabs` |
| Auth | `require_roles("admin", "production")` — без изменений состава ролей |

---

## Commands

```
# Backend
pytest tests/test_offers_production_authorization.py tests/test_production_api_integration.py tests/test_production_mixed_inclusion.py -q
pytest tests/ -q -k "kp_candidate or kp-candidates or in_work"

# Frontend
cd frontend && npm run test -- --run src/features/production
cd frontend && npm run typecheck

# Dev
./run+logs.sh
```

Новые тесты класть рядом с существующими (см. Testing Strategy), не плодить параллельный раннер.

---

## Project Structure

```
app/api/v1/endpoints/production.py          # Query scope на GET /kp-candidates
app/schemas/production.py                   # KpCandidateItem / plate bucket
app/services/production_service.py          # list_kp_candidates(scope=…)
app/repositories/kp_repository.py           # qty in_plan / remaining / on_sgp при необходимости

frontend/src/features/production/types/production.ts
frontend/src/features/production/api/productionApi.ts
frontend/src/features/production/hooks/useProductionQueries.ts
frontend/src/features/production/components/ProductionTabs.tsx
frontend/src/pages/production/ProductionPage.tsx
frontend/src/features/production/components/KpInWorkView.tsx          # список + раскрытие строки
frontend/src/features/production/components/KpInWorkView.test.tsx
frontend/src/features/production/components/ProductionTabs.test.tsx   # если нет — завести

tests/test_kp_candidates_scope.py           # срезы scope + qty (новый файл ок)
```

Документы: идея уже в `ai_docs/ideas/`; эта спека — источник правды до кода. План — только после апрува спеки.

Не импортировать `features/commercial-archive/*` во фронт производства.

---

## Code Style

Существующий контракт кандидатов расширяем, не ломаем. Пример целевой схемы:

```python
from typing import Literal

KpCandidatesScope = Literal["plan", "in_work"]
PlateBucket = Literal["awaiting_plan", "in_plan"]

class KpCandidatePlateItem(BaseModel):
    id: int
    plate_name: str
    length_m: float
    width_m: float
    load_class: int | None = None
    qty: int
    bucket: PlateBucket = "awaiting_plan"

class KpCandidateItem(BaseModel):
    kp_id: int
    customer_name: str
    creation_date: str
    execution_terms: str
    total_plates: int
    completed_plates: int
    completion_pct: float
    in_plan_pct: float
    total_length_m: float
    remaining_qty: int = 0
    in_plan_qty: int = 0
    on_sgp_qty: int = 0
    plates: list[KpCandidatePlateItem] = Field(default_factory=list)
```

Соглашения:

- Слой: роутер тонкий → сервис → репозиторий. Qty не считать в endpoint.
- Имена полей API — snake_case, как сейчас.
- UI вкладки: inline-стили и `Card`/`Alert` как соседние вкладки производства, не архивный список и не `Drawer`.
- Во фронте не рендерить `total_amount`, скидку, PDF, «перевести статус». Этих полей в payload быть не должно.
- Мастер продолжает звать `listKpCandidates()` **без** query. Вкладка — `listKpCandidates("in_work")`.

---

## Testing Strategy

| Уровень | Где | Что доказывает |
|---------|-----|----------------|
| API/сервис | `tests/test_kp_candidates_scope.py` + существующий `test_get_kp_candidates` | дефолт = только неразложенное; `scope=in_work` включает 100% в плане; полностью на СГП нет в обоих; `on_sgp_qty` / `in_plan_qty` / `remaining_qty`; неизвестный scope → 422 |
| Auth | `tests/test_offers_production_authorization.py` | `production` 200 на `kp-candidates` и на `?scope=in_work`; по-прежнему 403 на `/api/v1/offers` |
| Inclusion | `tests/test_production_mixed_inclusion.py` (и соседние exclusion) | во вкладке по-прежнему нет свай/ФБС/маршей/мостовых; mixed-with-plates есть |
| UI | vitest: `KpInWorkView.test.tsx`, вкладка в `ProductionTabs` | вкладка **после «Планы»**; строка без сумм; клик раскрывает плиты с bucket; 100% в плане — пометка «ждёт отливки»; пустое состояние |
| Регрессия мастера | `useCreatePlanWizardState.test.ts` | запрос кандидатов без `scope` (или без `in_work`) |

Покрытие: не гонимся за %; обязательны тесты на **два среза** и **отсутствие коммерческих полей** в UI.

---

## Boundaries

- **Always:** тесты на оба `scope` до merge; мастер без параметра не меняет состав; роли `admin`/`production`; без сумм в UI; плиты-only.
- **Ask first:** новый эндпоинт вместо query; миграция БД; пустить `manager` на вкладку; колонка метров в строке; кнопка «на дорожки» из карточки; счётчик на календаре.
- **Never:** доступ `production` в `/archive` или `/api/v1/offers`; импорт архивных компонентов; PDF/скидка/удаление КП из карточки; feature-flag «на всякий случай»; коммит секретов и живых `*.db`.

---

## Success Criteria

Конкретные, проверяемые условия done:

1. На `/production` у `admin` и `production` есть вкладка «КП в работе» **сразу после «Планы»**; `?tab=in-work` её открывает.
2. `GET /api/v1/production/kp-candidates` без query возвращает тот же состав, что сегодня (есть остаток «в производстве»).
3. `GET /api/v1/production/kp-candidates?scope=in_work` возвращает плиточные КП со статусом «в работе», у которых `remaining_qty + in_plan_qty > 0`. КП со 100% в плане и 0 на СГП **есть**. КП, полностью на СГП, **нет**.
4. В строке вкладки видны: номер КП, клиент, срок, осталось шт, уже в плане, на СГП. Нет суммы, скидки, PDF.
5. Клик по строке раскрывает плиты (марка, длина, ширина, нагрузка, шт). Позиции в плане подписаны «в плане, ждёт отливки». Нет drawer и действий архива.
6. Роль `production` по-прежнему не читает архив/offers (403).
7. Сваи/ФБС/марши/мостовые сваи не появляются в `scope=in_work`.
8. `pytest` по затронутым тестам и `npm run test -- --run src/features/production` + `typecheck` зелёные.

---

## API contract

`GET /api/v1/production/kp-candidates`

| Query | Default | Смысл |
|-------|---------|--------|
| `scope` | `plan` | `plan` — кандидаты мастера; `in_work` — очередь вкладки |

Оба ответа — `KpCandidatesResponse`. Поля qty аддитивны.

Фильтр `plan` (текущий код): `in_plan_pct < 100` **и** непустой `plates` (только «в производстве»).

Фильтр `in_work`: плиточный КП «в работе», `remaining_qty + in_plan_qty > 0`. `plates` = строки `kp_plates` со статусами «в производстве» и «в плане».

Инвариант qty (при корректном учёте):  
`remaining_qty + in_plan_qty + on_sgp_qty` согласован с заказным объёмом плит этого КП (как readiness: remaining + in_plan + on_sgp).

---

## UX

- Список: карточки-строки в стиле производства (`PlansList` / календарь), не `ArchiveOfferList`.
- Срок из `execution_terms`; пустой → «без срока».
- Клик по строке раскрывает плиты **под ней**. Снова клик — свернуть. Одновременно открыта одна строка.
- КП только в плане: в раскрытии те же марки с пометкой «в плане, ждёт отливки», не пустой блок и не один текст без списка.
- Загрузка: Spinner в `Card`; ошибка: `Alert`.
- Не добавлять кнопку внизу календаря («Начать планирование» без изменений).

---

## Out of scope (Not doing)

- Доступ производства в архив.
- Не-плитные номенклатуры в списке.
- Кнопка «как в архиве» на календаре.
- «На дорожки» из раскрытой строки и бейдж неразложенного на сетке.
- Drawer / коммерческая карточка КП, график поставки, смена статуса, удаление.
- Новый backend-роут, миграции, feature-flag.

---

## Open Questions

Продуктовые — закрыты (см. idea one-pager).

UX-вопросы закрыты ревью. Инженерные — допущениями 1–10. Если спека принята, открытых вопросов нет.
