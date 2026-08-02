# Spec: Визуализация укладки рейса (layout visualization)

> **Источник идеи:** [`ai_docs/ideas/shipment-layout-visualization.md`](../ideas/shipment-layout-visualization.md) (ideation 2026-08-02)
> **Фаза SDD:** SPECIFY ✅ → PLAN ⏳ → TASKS ⏳ → IMPLEMENT ⏳
> **Связанные:** [`shipment-propose-v2.md`](./shipment-propose-v2.md), `core/shipment_packing/`, `app/services/shipment_service.py`, `frontend/src/features/logistics/`
> **Дата:** 2026-08-02
> **Scope:** expose layout из `core/shipment_packing` → propose API → блок «Укладка в кузов» в UI карточки рейса. **Без** 2-го листа XLSX (вынос из MVP по решению 2026-08-02), **без** пересчёта layout при ручных правках.

---

## ASSUMPTIONS (согласовано 2026-08-02)

1. **Layout metadata хранится только в `propose_snapshot`** (уже пишется в БД после каждого propose) — новых колонок/таблиц в `shipments` не создаём.
2. **Схема скрывается при ручных правках:** после propose блок виден; любая ручная правка состава — блок пропадает до следующего propose.
3. **XLSX в MVP не входит** — только UI-блок (решение владельца; 2-й лист «Схема укладки» — кандидат post-MVP).
4. **t30plus → layout отсутствует:** legacy weight-FIFO не считает геометрию, поле `layout = null`, UI показывает «Схема укладки недоступна для класса t30+».
5. **Frontend-стейт локальный:** `layout` приезжает в ответе `useProposeMutation` и хранится в state `ShipmentItemsSection` — `GET /shipments/{id}` в MVP не расширяем.
6. **Порядок погрузки** = обход штабелей по порядку, внутри штабеля снизу вверх; формируется на бэкенде.
7. **Ручной override длины кузова** («Добавить всё равно» из propose v2) в MVP не реализован — в схеме не учитываем.
8. **SVG/div'ы без новых зависимостей** — никаких chart-библиотек.

## Objective

Логист после «Предложить состав» видит, КАК плиты уложены в кузов (штабели, ярусы, порядок погрузки), чтобы проверить, что состав реально грузится, и снизить перекладки/ошибки на площадке.

**Пользователь:** логист (роль `logistics`); вторичный — склад (по экрану/с монитора).

**Проблема сейчас:** `propose` v2 считает штабели и ярусы внутри движка, но отдаёт только плоский список позиций. Логист не видит физическую укладку; склад получает только таблицу состава.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | логист | видеть схему кузова сразу после propose | проверить, что состав реально грузится |
| US-2 | логист | чтобы схема скрывалась после моих ручных правок | не показывать складу устаревшую укладку |
| US-3 | логист | видеть нумерованный порядок погрузки | продиктовать/показать складу последовательность |

### Acceptance criteria (MVP)

- [ ] `PackResult` расширен: `layout: LayoutMetadata | None` (stacks→tiers→units, `loading_steps`, `body_used_m`); новые dataclasses в `core/shipment_packing/models.py`
- [ ] `ShipmentProposeResponse.layout` — nullable; для `t20` (и отсутствия класса) заполняется, для `t30plus` = `null`
- [ ] `propose_snapshot` сохраняет layout (JSON); `scripts/shipment_propose_hitrate.py` не ломается (читает только `items`)
- [ ] UI: блок «Укладка в кузов» в `ShipmentItemsSection` после успешного propose:
  - полоска кузова 13,2 м — штабели пропорционально длине маркировки, подпись метража;
  - раскрытие штабеля по клику — ярусы с марками плит и ширинами;
  - список «Порядок погрузки» (нумерованный).
- [ ] UI: любая ручная правка состава после propose — блок скрывается полностью до следующего propose
- [ ] UI: `layout == null` (t30plus / пустой состав) — блок не рендерится или показывает «Схема укладки недоступна»
- [ ] Golden-тест layout metadata на эталонном рейсе (9 плит: штабели 8,9 + 4,3 м, 5 шагов погрузки)
- [ ] Существующие тесты shipment/logistics зелёные; `npm run build` зелёный

### Out of scope (Not Doing)

- 2-й лист «Схема укладки» в `sheet.xlsx` — **вынос из MVP** (решение 2026-08-02), кандидат post-MVP
- Пересчёт/валидация layout при ручных правках состава (live layout)
- Drag-and-drop редактор укладки
- 3D / фотореалистичный рендер
- ГОСТ-паспорт / схема загрузки для водителя
- Reuse `viz_modules` (другая предметная область — производственная раскладка на дорожках)
- QR → мобильная схема для крановщика
- Раскрытие `GET /shipments/{id}` layout'ом (восстановление схемы при переоткрытии карточки) — post-MVP

---

## Tech Stack

| Слой | Стек |
|------|------|
| Engine | `core/shipment_packing/` — новые `dataclass(frozen=True)`: `LayoutMetadata`, `LayoutStack`, `LayoutTier`, `LayoutUnit`, `LoadingStep` |
| Service | `app/services/shipment_service.py` — `_propose_v2_packing` маппит `result.layout` в ответ |
| API schema | `app/schemas/logistics.py` — `ShipmentLayout*` Pydantic-модели, поле `layout` в `ShipmentProposeResponse` |
| Frontend | `frontend/src/features/logistics/` — новый `LayoutBlock.tsx`, типы в `types/logistics.ts`, state в `ShipmentItemsSection.tsx` |
| Tests | `tests/test_shipment_packing.py` (layout), `tests/test_shipment_service.py` (snapshot), `tests/test_logistics_api.py` (API), `LayoutBlock.test.tsx` |

## Commands

```bash
# Backend
pytest tests/test_shipment_packing.py tests/test_shipment_service.py -q
pytest tests/ -k "shipment or logistics" -q

# Hit-rate метрика (не должна сломаться)
./.venv/bin/python scripts/shipment_propose_hitrate.py --verbose

# Frontend
cd frontend && npm run build
cd frontend && npm test -- --run src/features/logistics
```

## Project Structure

```
core/shipment_packing/
  models.py      → + LayoutMetadata, LayoutStack, LayoutTier, LayoutUnit, LoadingStep
  engine.py      → pack_shipment собирает layout из _Layout

app/schemas/logistics.py        → + ShipmentLayout* модели
app/services/shipment_service.py → _propose_v2_packing маппит layout

frontend/src/features/logistics/
  components/LayoutBlock.tsx       → NEW: полоска кузова + раскрытие штабеля + порядок
  components/LayoutBlock.test.tsx  → NEW
  components/ShipmentItemsSection.tsx → state layout + hide on manual edit
  types/logistics.ts               → + ShipmentLayout* типы
```

## Code Style

Backend — pure engine, frozen dataclasses для read-моделей:

```python
@dataclass(frozen=True)
class LayoutStack:
    index: int
    marking_length_m: float
    tiers: list[LayoutTier] = field(default_factory=list)
```

Frontend — named export, `React.CSSProperties`-объекты (как `ShipmentItemsSection`), русские тексты, рамки `#eaecf0`, белые карточки, без новых UI-зависимостей.

## Testing Strategy

| Уровень | Что проверяем | Где |
|---------|----------------|-----|
| Engine unit | структура layout: штабели, ярусы, шаги, body_used_m на эталонном рейсе | `tests/test_shipment_packing.py` |
| Service | layout в ответе propose и в `propose_snapshot` JSON | `tests/test_shipment_service.py` |
| API | `POST /propose` возвращает `layout` (TestClient) | `tests/test_logistics_api.py` |
| Frontend | рендер полоски, раскрытие штабеля, скрытие при правке, null-состояние | `LayoutBlock.test.tsx`, `ShipmentDrawer.test.tsx` |

Покрытие: новый код — тесты обязательны; регрессионный гейт — `pytest -k "shipment or logistics"`.

## Boundaries

- **Always:** pytest после backend-правок; `npm run build` после frontend; минимальный diff
- **Ask first:** новые зависимости; изменение формата `propose_snapshot`, ломающее hitrate-скрипт; новые колонки/таблицы БД
- **Never:** трогать `viz_modules`; live-валидация ручного состава; коммит без явной просьбы; править несвязанный код

## Success Criteria

1. На эталонном рейсе (3× ПБ 89-12-8п, 3× ПБ 80-12-8п, 1× ПБ 43-12-8п, 1× ПБ 42,6-5,3-10п, 1× ПБ 42-3,0-8п) UI показывает 2 штабеля (8,9 + 4,3 м) и 5 шагов погрузки.
2. После изменения qty любой строки блок схемы исчезает; после повторного propose появляется снова.
3. `propose_snapshot` всех новых propose по t20 содержит `layout` (100%).
4. Hitrate-скрипт работает без изменений.

## Open Questions

- XLSX со схемой (post-MVP): всегда 2-й лист или отдельная кнопка — решить при возврате к задаче
- Восстанавливать ли схему из snapshot при переоткрытии карточки рейса (post-MVP)
- Метрика misload (поле «перекладки» в рейсе) — post-MVP
