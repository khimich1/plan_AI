# Spec: A — Менеджер видит завод (готовность КП)

> **Источник идеи:** [`ai_docs/ideas/kp-readiness-manager-view.md`](../ideas/kp-readiness-manager-view.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → IMPLEMENT ✅  
> **Статус:** ✅ implemented (MVP)  
> **Roadmap:** [`2026-07-28-development-roadmap-cifra-context.md`](../develop/plans/2026-07-28-development-roadmap-cifra-context.md) — направление **A**  
> **Дата:** 2026-07-28  
> **Зависит от:** СГП MVP ✅ ([`sgp-warehouse.md`](./sgp-warehouse.md))  
> **Не делает:** **J** (статус для клиента по токену), **B** (выдача/отгрузка)  
> **Связанные:** `ArchiveService`, `SgpService`, `OfferDetailsDrawer`, `ArchiveOfferList`, `kp_plates`, `completed_plates`, `kp_meta.ordered_qty`  
> **Addendum:** [`kp-readiness-expected-sgp-date.md`](./kp-readiness-expected-sgp-date.md) — ожидаемая дата на СГП (SPECIFY)

---

## Assumptions I'm Making

1. **Read-only feature** — никаких мутаций БД, только агрегация существующих `kp_plates` + `completed_plates`.
2. **Auth** — те же роли, что архив: `admin`, `manager` (не публичный endpoint).
3. **«Можно выдать»** = `SUM(qty)` на СГП с `kp_id` (linked), без резервирования — до модуля B.
4. **Колонка «На СГП»** в таблице включает весь складской баланс по позиции (дорожка + wizard close-from-SGP); отдельная колонка «с СГП» **не показывается**.
5. **Блок readiness** виден только при `status ∈ {в работе, На СГП}` — не в «в архиве» / «выполнено».
6. **Refetch** при открытии drawer достаточен; отдельная кнопка «Обновить» — out of MVP.
7. **SQLite**, один `plita.db` — без новых таблиц и миграций.

→ Если что-то неверно — поправьте до Phase 2 (Plan).

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| 1 | API summary | Расширить `GET /commercial/archive/{kp_id}` → поле `readiness` в `ArchiveOfferDetails` |
| 2 | API positions | Lazy: `GET /commercial/archive/{kp_id}/readiness/positions` |
| 3 | Единый read-model | `KpReadinessService` — один источник для list/details/positions (list уже через `SgpService.sgp_progress`) |
| 4 | Степпер | 5 шагов; «Выдача» и «Закрыто» — `disabled` до направления B |
| 5 | Таблица | По клику «Подробнее»; колонки: Заказ \| В плане \| На СГП \| Осталось |
| 6 | «Можно выдать X шт» | X = N из `sgp_progress` (linked qty на СГП) |
| 7 | N/M | Как в СГП spec: M = `kp_meta.ordered_qty`; N = `SUM(completed_plates.qty WHERE kp_id=?)` |
| 8 | J / B | Out of scope |
| 9 | Feature-flag | Нет |
| 10 | `sectionFromStatus` | Fix в том же MVP: «На СГП» → `in_production` |
| 11 | % в шапке drawer | **Убрать** — только в блоке readiness; в списке архива бейдж % остаётся |
| 12 | Positions wrong status | `200` + `items: []` для «в архиве» / «выполнено» |
| 13 | Сортировка таблицы | По `position_number` (как в PDF/КП) |
| 14 | Copy для клиента | Клиентский тон: «Здравствуйте! По вашему заказу…» |

---

## Objective

Дать **менеджеру продаж** один экран в карточке КП, чтобы **за <1 мин без звонка в цех** ответить клиенту: что на складе, что ещё производится, сколько можно забрать.

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | открыл КП «в работе» в архиве | видеть степпер, %, N/M и сводную фразу | сразу ответить клиенту по телефону |
| US-2 | клиент спрашивает по конкретной марке | раскрыть «Подробнее» и увидеть таблицу позиций | не звонить мастеру за детализацию |
| US-3 | нужно переслать статус | нажать «Скопировать для клиента» | отправить текст в мессенджер (J не нужен) |
| US-4 | ищу КП через поиск | видеть readiness в drawer | не зависеть от активной вкладки архива |
| US-5 | КП перешло на «На СГП» | видеть его во вкладке «В производстве» с бейджом | не искать в «Архиве» по ошибке |

### Acceptance criteria (MVP)

- [x] `GET /commercial/archive/{kp_id}` для `status ∈ {в работе, На СГП}` возвращает `readiness` с `steps`, `summary_text`, `client_copy_text`, `sgp_progress`, `completion_percentage`, `issuable_qty`, `in_production_qty`.
- [x] `GET …/readiness/positions` возвращает строки с колонками `ordered`, `in_plan`, `on_sgp`, `remaining`; сумма по позиции сходится: `ordered = in_plan + on_sgp + remaining`.
- [x] UI: блок «Статус производства» в `OfferDetailsDrawer` — степпер (5 шагов), метрики, фраза, «Подробнее», «Скопировать».
- [x] Таблица позиций грузится **только** после раскрытия «Подробнее» (TanStack Query lazy).
- [x] Блок **не показывается** для «в архиве» и «выполнено».
- [x] Шаги «Выдача» / «Закрыто» визуально disabled; подпись «Выдача с СГП — в следующем обновлении».
- [x] `sectionFromStatus`: `«На СГП»` → `in_production` (fix регрессии поиска/списка).
- [ ] Readiness **не меняет** данные БД; plate_loss regression остаётся PASS (не прогонялся в worker).

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | FastAPI, SQLite (`plita.db`), Pydantic v2 |
| Domain | `kp_plates`, `completed_plates`, `kp_meta`, `SgpService.sgp_progress` |
| API | `app/api/v1/endpoints/archive.py` |
| Frontend | React, Vite, TypeScript, TanStack Query |
| Tests | pytest (`tests/`), Vitest (`frontend/`) |

---

## Commands

```bash
# Backend
source .venv/bin/activate
uvicorn app.main:app --reload
pytest tests/test_kp_readiness_service.py -q
pytest tests/test_archive_endpoints.py -q
pytest tests/ -k "readiness or archive" -q

# Qty gate (readiness не должен ломать баланс)
./.venv/bin/python scripts/run_plate_loss_regression.py

# Frontend
cd frontend && npm run dev
cd frontend && npm test -- --run KpReadiness
cd frontend && npm run build
```

---

## Project Structure

```
app/services/kp_readiness_service.py     → NEW: summary + positions aggregation (read-only)
app/services/archive_service.py          → _to_details: attach readiness; get_readiness_positions
app/schemas/archive.py                   → KpReadinessSummary, KpReadinessStep, position DTOs
app/api/v1/endpoints/archive.py          → GET /{kp_id}/readiness/positions
app/dependencies/services.py             → get_kp_readiness_service (если нужен DI)

frontend/src/features/commercial-archive/
  components/KpReadinessBlock.tsx        → NEW: stepper, summary, expand, copy
  components/OfferDetailsDrawer.tsx      → встроить KpReadinessBlock
  api/archiveApi.ts                      → getReadinessPositions(kpId)
  hooks/useArchiveQueries.ts             → useKpReadinessPositionsQuery (enabled on expand)
  types/archive.ts                       → readiness types
  lib/kpReadinessCopy.ts                 → NEW: optional helper для clipboard text

frontend/src/pages/commercial-offer-archive/
  CommercialOfferArchivePage.tsx         → fix sectionFromStatus «На СГП»

tests/test_kp_readiness_service.py       → NEW: aggregation, formulas, edge cases
tests/test_archive_endpoints.py          → extend: readiness in details + positions route

ai_docs/ideas/kp-readiness-manager-view.md
ai_docs/specs/kp-readiness-manager-view.md   → эта спека
ai_docs/develop/plans/…                      → Phase 2 после approval
```

---

## Code Style

Read-only сервис — без side effects, переиспользование `SgpService`:

```python
# app/services/kp_readiness_service.py
class KpReadinessService:
    def __init__(self, *, db_path: str) -> None:
        self.db_path = db_path

    def build_summary(self, kp_id: int, *, status: str) -> KpReadinessSummary | None:
        if status not in _READINESS_STATUSES:
            return None
        progress = SgpService(db_path=self.db_path).sgp_progress(kp_id)
        completion = get_kp_completion_percentage(kp_id, self.db_path)
        in_production_qty = int(completion.get("in_production") or 0)
        issuable_qty = progress.n
        summary_text = self._format_summary(progress, in_production_qty)
        return KpReadinessSummary(
            completion_percentage=float(completion.get("percentage") or 0.0),
            sgp_progress=progress,
            issuable_qty=issuable_qty,
            in_production_qty=in_production_qty,
            summary_text=summary_text,
            client_copy_text=self._format_client_copy(kp_id, progress, summary_text),
            steps=self._build_steps(status, progress, in_production_qty),
            release_note="Выдача с СГП — в следующем обновлении",
        )
```

- Слои: router → `ArchiveService` → `KpReadinessService` → SQL / `core.kp.*`.
- Identity позиции: `plate_name` + `length_m` + `width_m` + `load_class` (ABS length/width `< 0.005` при join, как СГП).
- Тексты для UI на русском; поля API — snake_case.
- Frontend: функциональные компоненты, стили inline как в `OfferDetailsDrawer` / `ArchiveOfferList`.

---

## Domain model

### Два регистра (без изменений vs СГП spec)

| Регистр | Таблица | Readiness использует |
|---------|---------|----------------------|
| Потребность | `kp_plates` | `в плане`, `в производстве` |
| Физика | `completed_plates` | qty с `kp_id` = on_sgp |

### Агрегация по позиции (identity)

Для каждого `kp_id` собрать **объединение** identity из `kp_plates` и `completed_plates`:

| Поле | Формула |
|------|---------|
| **on_sgp** | `SUM(completed_plates.qty) WHERE kp_id=? AND identity` |
| **in_plan** | `SUM(kp_plates.qty) WHERE kp_id=? AND status='в плане' AND identity` |
| **remaining** | `SUM(kp_plates.qty) WHERE kp_id=? AND status='в производстве' AND identity` |
| **ordered** | `on_sgp + in_plan + remaining` |

**Инвариант строки:** `ordered = in_plan + on_sgp + remaining` (qty conservation; gate в тестах).

**Отображение:** статус wizard «с СГП» **не выводится** — его вклад уже в `on_sgp` (плиты учтены на складе / потребность закрыта через резерв).

**Пример (из ideation):**

ПБ 59-12-8: заказ 10 — 3 закрыты wizard «с СГП», 3 с дорожки, 4 в плане → **10 / 4 / 6 / 0**.

### Summary (уровень КП)

| Поле | Источник |
|------|----------|
| `completion_percentage` | `get_kp_completion_percentage` → `percentage` |
| `sgp_progress {n,m}` | `SgpService.sgp_progress` |
| `issuable_qty` | `n` |
| `in_production_qty` | `completion.in_production` (= `SUM(kp_plates.qty)`) |

### Шаблоны текстов

**`summary_text` (UI):**

| Условие | Текст |
|---------|-------|
| `n > 0`, `in_production_qty > 0` | «{n} из {m} шт на складе, {in_production_qty} в производстве. Можно выдать {n} шт.» |
| `n > 0`, `in_production_qty == 0` | «{n} из {m} шт на складе. Можно выдать {n} шт.» |
| `n == 0`, `in_production_qty > 0` | «Заказ в производстве ({in_production_qty} шт). На складе пока нет.» |
| иначе | «Данных о производстве пока нет.» |

**`client_copy_text` (clipboard, клиентский тон):**

| Условие | Текст |
|---------|-------|
| `n > 0`, `in_production_qty > 0` | «Здравствуйте! По вашему заказу №{kp_id}: {n} из {m} шт уже на складе, остальные ({in_production_qty} шт) в производстве. Можно забрать {n} шт.» |
| `n > 0`, `in_production_qty == 0` | «Здравствуйте! По вашему заказу №{kp_id}: {n} из {m} шт на складе, можно забрать.» |
| `n == 0`, `in_production_qty > 0` | «Здравствуйте! По вашему заказу №{kp_id}: заказ в производстве ({in_production_qty} шт), на складе пока нет.» |
| иначе | «Здравствуйте! По вашему заказу №{kp_id}: уточняем статус производства, скоро сообщим.» |

`summary_text` остаётся **внутренним** (для менеджера в UI); в clipboard — только `client_copy_text`.

Менеджер правит вручную при необходимости; PDF не прикладывается (J out).

### Lifecycle stepper

| Step | id | MVP state |
|------|-----|-----------|
| КП | `kp` | `done` |
| Производство | `production` | `active` если `in_plan + remaining > 0`; иначе `done` если `n > 0`; иначе `pending` |
| СГП | `sgp` | `done` если `n == m && m > 0`; `active` если `n > 0`; иначе `pending` |
| Выдача | `release` | `disabled` |
| Закрыто | `closed` | `disabled` |

Под степпером: `release_note`.

---

## API

| Method | Path | Response | Назначение |
|--------|------|----------|------------|
| `GET` | `/api/v1/commercial/archive/{kp_id}` | `ArchiveOfferDetails` | + поле `readiness: KpReadinessSummary \| null` |
| `GET` | `/api/v1/commercial/archive/{kp_id}/readiness/positions` | `KpReadinessPositionsResponse` | Lazy таблица позиций |

### Schemas (черновик)

```python
class KpReadinessStepState(str, Enum):
    DONE = "done"
    ACTIVE = "active"
    PENDING = "pending"
    DISABLED = "disabled"

class KpReadinessStep(BaseModel):
    id: Literal["kp", "production", "sgp", "release", "closed"]
    label: str
    state: KpReadinessStepState
    hint: str | None = None  # e.g. "72%" or "14/20"

class KpReadinessSummary(BaseModel):
    completion_percentage: float | None = None
    sgp_progress: SgpProgress | None = None
    issuable_qty: int = 0
    in_production_qty: int = 0
    summary_text: str = ""
    client_copy_text: str = ""
    steps: list[KpReadinessStep] = Field(default_factory=list)
    release_note: str | None = None

class KpReadinessPositionItem(BaseModel):
    position_number: int | None = None  # для сортировки; min по identity
    plate_name: str
    length_m: float | None = None
    width_m: float | None = None
    load_class: int | None = None
    label: str  # display, e.g. "ПБ 59-12-8"
    ordered: int
    in_plan: int
    on_sgp: int
    remaining: int

class KpReadinessPositionsResponse(BaseModel):
    items: list[KpReadinessPositionItem]
    count: int
```

`ArchiveOfferDetails.readiness: KpReadinessSummary | None = None`

### Errors

- `404` — КП не найдено (как сейчас).
- `422` — не применимо (нет body).
- Positions для КП без readiness («в архиве» / «выполнено») → **`200` + `items: []`** (не 404).

---

## UI

### `KpReadinessBlock` в `OfferDetailsDrawer`

Расположение: **после** шапки (клиент/статус), **до** блока «Итоги».

```
┌─ Статус производства ────────────────────────────────────────────┐
│  [stepper 5 steps]                                                │
│  Производство: 72%          СГП: 14/20                            │
│  «14 из 20 шт на складе…»                                         │
│  Выдача с СГП — в следующем обновлении                            │
│  [Подробнее ▼]  [Скопировать для клиента]                         │
│  ── table (if expanded) ──                                         │
└───────────────────────────────────────────────────────────────────┘
```

- `showReadiness = ["в работе", "На СГП"].includes(offer.status)`.
- «Скопировать» → `navigator.clipboard.writeText(readiness.client_copy_text)` + toast/alert.
- Блок «Готовность X%» в **шапке drawer убрать** — % только в readiness-блоке; бейдж % в **списке** архива без изменений.

### Lazy positions

```typescript
const [expanded, setExpanded] = useState(false);
const positionsQuery = useKpReadinessPositionsQuery(kpId, { enabled: expanded && showReadiness });
```

Refetch details при `open && kpId` (как сейчас `useArchiveOfferQuery`).

### Archive list

Без изменений в MVP (бейджи N/M и % уже есть). Опционально позже: tooltip со `summary_text`.

### Fix `sectionFromStatus`

```typescript
case "На СГП":
  return "in_production";
```

---

## Testing Strategy

| Уровень | Что | Где |
|---------|-----|-----|
| Unit | position aggregation: ordered = in_plan + on_sgp + remaining | `tests/test_kp_readiness_service.py` |
| Unit | summary templates: n=0, partial, full on SGP | `tests/test_kp_readiness_service.py` |
| Unit | stepper states for partial / full / empty SGP | `tests/test_kp_readiness_service.py` |
| Integration | `GET /archive/{id}` includes readiness for «в работе» | `tests/test_archive_endpoints.py` |
| Integration | readiness `null` for «в архиве» | `tests/test_archive_endpoints.py` |
| Integration | positions endpoint lazy contract | `tests/test_archive_endpoints.py` |
| Frontend | KpReadinessBlock: render, expand, copy | Vitest |
| Frontend | sectionFromStatus «На СГП» | Vitest |
| Regression | `run_plate_loss_regression.py` PASS | script |

Fixtures: переиспользовать паттерн `tests/test_sgp_service.py` (tmp db, kp_plates + completed_plates).

---

## Boundaries

### Always

- Read-only SQL; один `KpReadinessService` для summary и positions
- Переиспользовать `SgpService.sgp_progress` и `get_kp_completion_percentage` — не дублировать N/M
- Тесты aggregation + archive endpoints перед merge
- Видимость readiness строго по `status` КП
- Русские тексты summary согласованы с backend (не собирать фразу только на frontend)

### Ask first

- Публичный endpoint для J (токен без auth)
- Показ readiness в «Выполнено»
- Кнопка «Обновить» / polling / WebSocket
- ETA по дням производства / Gantt
- Изменение формулы N/M или `ordered_qty` freeze

### Never

- Мутировать `kp_plates` / `completed_plates` из readiness
- Писать «можно выдать» на основе резервирования до модуля B
- Отдельная колонка «с СГП» в UI
- Feature-flag для скрытия блока
- Коммитить секреты

---

## Success Criteria

1. **Drawer:** КП «в работе» с частичным СГП показывает степпер, 72%-подобный hint, N/M, фразу с «Можно выдать N шт».
2. **Positions:** для фикстуры «10 заказ, 4 в плане, 6 на СГП, 0 осталось» — одна строка `10/4/6/0`.
3. **Copy:** clipboard содержит «Здравствуйте! По вашему заказу №…» (`client_copy_text`), не дублирует `summary_text` дословно.
4. **Hidden:** «в архиве» / «выполнено» — `readiness: null`, блок не рендерится.
5. **Disabled steps:** «Выдача» и «Закрыто» не кликабельны, визуально серые.
6. **Search:** КП «На СГП» в результатах поиска → `sectionFromStatus` = `in_production`.
7. **API auth:** без login → 401; manager видит только свои КП (как архив).
8. **Regression:** plate_loss script PASS; readiness tests green; `npm run build` OK.

---

## Out of Scope

| Item | Why |
|------|-----|
| **J** — ссылка клиенту по токену | Ideation: skip; достаточно copy |
| **B** — выдача, списание, акт | Шаги 4–5 disabled до B |
| Push / email при изменении N | Нет политики оповещений |
| Readiness в «Выполнено» | Live-статус не нужен |
| Dashboard C | Отдельное направление |
| ETA / дни / дорожки | Усложнение; таблица по маркам достаточна |
| Кнопка «Обновить» | Refetch on open |

---

## Open Questions

_Нет блокирующих — Q11–Q14 закрыты 2026-07-28._

---

## Next (SDD)

1. ~~Human review spec~~ ✅
2. **Phase 2 Plan:** [`ai_docs/develop/plans/2026-07-28-kp-readiness-manager-view.md`](../develop/plans/2026-07-28-kp-readiness-manager-view.md) ✅
3. **Phase 3 Tasks** → **Phase 4 Implement** (RDY-100…)
