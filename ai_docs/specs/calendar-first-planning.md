# Spec: Calendar Brush Planning (кисть на сетке)

> **Источник:** [`ai_docs/ideas/calendar-first-planning.md`](../ideas/calendar-first-planning.md)  
> **Дата:** 2026-07-23  
> **Статус:** implemented (BRUSH-001…007)  
> **Plan:** [`ai_docs/develop/plans/2026-07-23-calendar-brush-planning.md`](../develop/plans/2026-07-23-calendar-brush-planning.md)  
> **Supersedes UX:** drawer-based add-to-basket из CAL-001…009 (backend/`fill_targets`/wizard без изменений)  
> **Связанные:** `FillBasket.tsx`, `MonthCalendarGrid.tsx`, `basketDayKind.ts`, `planNameFromDates.ts`

---

## Assumptions (validate before Plan)

```
ASSUMPTIONS I'M MAKING:
1. Sticky-бар с пресетом N всегда виден на «Календарном плане» (даже при пустой корзине).
2. Жесты MVP: клик = toggle день в корзине с текущим N; Shift+клик = диапазон от якоря до дня.
   Drag-select — OUT of MVP (можно follow-up).
3. Открыть DayDrawer: отдельный жест — кнопка/иконка «i» на ячейке или двойной клик.
   Одинарный клик НЕ открывает drawer.
4. Секцию «Положить дорожек / Добавить в план» из DayDrawer УДАЛЯЕМ сразу (без feature-flag).
5. Inline-правка N: на чипе в корзине (number input). На ячейке — показ N; клик по N на
   выделенной ячейке тоже может открыть быстрый edit (nice-to-have, не блокер MVP).
6. Повторный клик по дню уже в корзине = снять с корзины (toggle). Смена N пресета
   НЕ перекрашивает уже выбранные дни автоматически — только новые / повторная покраска.
7. Empty vs partial по-прежнему нельзя смешивать; кисть по «чужому» kind → warning, день не добавляется.
8. Backend и CreatePlanWizard (fill-only) не меняем.
→ Correct me now or I'll proceed with these.
```

---

## Objective

Планировщик задаёт **N дорожек заранее** (кисть), **выделяет диапазон дней на сетке** без drawer, при необходимости **подкручивает N** на отдельных днях (чипы), затем уходит в выбор плит через уже существующую корзину + `fill_targets`.

### User stories

| # | Как планировщик… | Я хочу… | Чтобы… |
|---|------------------|---------|--------|
| US-1 | начинаю план на несколько дней | выставить N и Shift-выделить диапазон | не открывать drawer на каждый день |
| US-2 | на одном дне нужно другое N | поменять N на чипе | не сбрасывать весь диапазон |
| US-3 | смотрю содержимое дня | открыть drawer отдельным жестом | не путать просмотр и выбор |
| US-4 | дозаполняю partial дни | той же кистью на partial-сетке | сохранить dual CTA |

### Acceptance criteria (MVP)

- [x] Sticky-бар всегда показывает пресет `Дорожек: N` (default = max свободных типично `max_per_day`, clamp 1..max)
- [x] Клик по selectable дню (empty или partial с freeSlots>0) → add/remove из корзины с текущим N (clamp к freeSlots дня)
- [x] Shift+клик → inclusive range от lastAnchor до клика; рабочие дни с подходящим kind добавляются
- [x] Выходные / full / completed — не выбираются кистью
- [x] Чип в корзине: editable N + ✕ remove
- [x] DayDrawer: **нет** UI добавления в корзину; открытие только secondary gesture
- [x] Dual CTA + kind validation + fill-only wizard — без регрессий
- [x] Backend unchanged; fill_targets smoke green

---

## Tech Stack

| Слой | Стек |
|------|------|
| Frontend | React 18, TypeScript, Vite, Vitest |
| Backend | без изменений (`fill_targets`) |
| State | `ProductionPage` basket + `brushTracks` (новый) |
| Grid | `MonthCalendarGrid` — расширить жестами |

---

## Commands

```bash
cd frontend && npm run dev
cd frontend && npm test -- --run
cd frontend && npm run build
pytest tests/test_production_planning_service_fill_targets_smoke.py -q
pytest tests/test_production_planning_service_fill_targets.py -q
```

---

## Project Structure

```
frontend/src/features/production/
  components/
    MonthCalendarGrid.tsx     → click / Shift+range; optional onOpenDay
    FillBasket.tsx            → пресет N + чипы с editable tracks
    GlobalCalendarView.tsx    → wiring brush; DayDrawer без fill section
    DayDrawer.tsx             → удалить onAddToFillBasket / fill UI
  lib/
    basketDayKind.ts          → без изменений контракта
    calendarRange.ts          → NEW: datesBetween(a,b), filterSelectable
    planNameFromDates.ts      → без изменений
  pages/production/
    ProductionPage.tsx        → brushTracks state; addRange helper
```

---

## Code Style

```typescript
/** Inclusive ISO date range, sorted ascending. */
export function datesBetweenInclusive(a: string, b: string): string[] {
  // walk day-by-day UTC/local consistently with MonthCalendarGrid formatISO
}

export function paintDays(args: {
  dates: string[];
  brushTracks: number;
  daysInfo: Record<string, DayInfo>;
  basketKind: BasketDayKind | null;
}): { added: FillTargetItem[]; error: string | null } {
  // for each date: getDayKind → canAdd → clamp tracks to freeSlots
}
```

- Пресет N и per-day N в корзине — разные поля: `brushTracks` vs `item.tracks`
- Смена `brushTracks` не мутирует существующие `basket` items
- Жесты: `onDayClick(iso, { shiftKey })`; якорь `selectionAnchor: string | null`

---

## As-Is → To-Be

| As-Is (CAL-001…009) | To-Be (brush) |
|---------------------|---------------|
| Клик → DayDrawer → N → «Добавить в план» | Клик → сразу в корзину с пресетом N |
| Нет диапазона | Shift+клик диапазон |
| Sticky только при непустой корзине | Sticky всегда: пресет N + чипы + CTA |
| Правка N через drawer «Заменить» | Правка N на чипе |
| Drawer = планирование + просмотр | Drawer = только просмотр/документы |

---

## Interaction model

```
brushTracks = 3 (sticky input)
lastAnchor = null

click day D (no Shift):
  if D in basket → remove D
  else → paint [D] with brushTracks; lastAnchor = D

Shift+click day D:
  if lastAnchor == null → treat as plain click
  else → paint datesBetween(lastAnchor, D); lastAnchor = D

double-click OR «i» on cell:
  open DayDrawer(D)  // no fill controls

chip N change:
  set item.tracks = clamp(value, 1, freeSlots(date))
```

**Selectable:** workday, not completed, freeSlots > 0, kind compatible with basket.

---

## Testing Strategy

| Уровень | Что | Где |
|---------|-----|-----|
| Unit | `datesBetweenInclusive`, `paintDays`, clamp/mix errors | `lib/calendarRange.test.ts` |
| Unit | FillBasket: пресет N, chip edit | optional RTL |
| Hook/page | add/remove/range не ломают kind validation | extend existing tests |
| Regression | wizard + fill_targets | existing vitest + pytest |
| Manual | Shift-диапазон empty; partial; mix reject; drawer без «Добавить» | browser |

---

## Boundaries

### Always
- Preserve kind validation messages
- Clamp N to day's free slots
- Run frontend tests + fill_targets smoke before done
- Keep programmatic `tab=create` + redirect

### Ask first
- Adding drag-select in MVP
- Changing dual-CTA labels
- New npm deps for gesture libs

### Never
- Backend / `fill_targets` semantics changes
- Mixing empty + partial in one basket
- Reintroducing wizard steps 1–2
- Feature-flag half-drawer-fill (удаляем fill UI целиком)

---

## Success Criteria

1. Пресет N=3 → Shift от 20.07 до 24.07 (рабочие) → чипы с 3 дор. → «Начать планирование» → plates → build
2. Изменить чип 22.07 на 5 → в `fill_targets` уезжает 5 только для этой даты
3. Partial-only корзина → «Дозаполнить»
4. Попытка кистью добавить partial в empty-корзину → warning, корзина не меняется
5. Двойной клик / «i» → drawer без кнопок добавления в план
6. `npm test -- --run` + fill_targets pytest green

---

## Out of Scope

- Drag-select
- Автораспределение бюджета дорожек
- Смешанная корзина
- Режим-toggle «Планирую» (жесты выше достаточны без toggle)
- Backend changes

---

## Open Questions (нужен ответ / подтверждение assumptions)

| # | Вопрос | Рекомендация в спеке |
|---|--------|----------------------|
| Q1 | Toggle «Планирую»? | **Нет** — клик=кисть, secondary=drawer |
| Q2 | Drag в MVP? | **Нет** — только Shift+диапазон |
| Q3 | Удалить fill из DayDrawer сразу? | **Да** |
| Q4 | Смена пресета N перекрашивает выбранные? | **Нет** |
| Q5 | Повторный клик = снять? | **Да** (toggle) |

---

## Next Steps

1. ~~Phase 2 — Plan~~ → [`2026-07-23-calendar-brush-planning.md`](../develop/plans/2026-07-23-calendar-brush-planning.md)
2. ~~Implement: BRUSH-001 → … → BRUSH-007~~
3. Manual browser smoke (Success Criteria 1–5) + review / commit по запросу

---

## Verification checklist (Phase 1)

- [x] Spec covers Objective, Commands, Structure, Code Style, Testing, Boundaries
- [ ] Human reviewed and approved spec (+ assumptions Q1–Q5)
- [x] Success criteria specific and testable
- [x] Spec saved: `ai_docs/specs/calendar-first-planning.md` (replaced; brush v2)
