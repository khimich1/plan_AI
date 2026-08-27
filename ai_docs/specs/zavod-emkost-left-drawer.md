# Spec: Ёмкость завода — left drawer по кнопке

> **Источник:** [`ai_docs/ideas/zavod-emkost-left-drawer.md`](../ideas/zavod-emkost-left-drawer.md)  
> **Родитель (гейт + алгоритм):** [`ai_docs/specs/zavod-emkost-vizual-gate.md`](zavod-emkost-vizual-gate.md)  
> **Plan:** [`ai_docs/develop/plans/2026-08-26-zavod-emkost-left-drawer.md`](../develop/plans/2026-08-26-zavod-emkost-left-drawer.md)  
> **Дата:** 2026-08-26  
> **Статус:** approved — UX delta; automated verify ✅; Human QA pending; коммит только по просьбе  
> **Связанные:** `Drawer` (`side="left"`), `MoveToProductionDialog`, `DeliveryScheduleDialog`,
> `FactoryCapacityPanel`

Эта спека **меняет только размещение UI**. Алгоритм светофора, backend gate,
`capacity-snapshot`, отсутствие override — без изменений из родительской спеки.

---

## Assumptions (locked)

```
ASSUMPTIONS:
1. Кнопка-триггер называется ровно «Ёмкость» (без бейджа статуса / Δ на кнопке).
2. Одинаковый паттерн в MoveToProduction и DeliverySchedule.
3. При red: короткий Alert/hint ОСТАЁТСЯ в модалке; drawer НЕ автооткрывается.
4. Drawer крепится к ЛЕВОМУ краю viewport (не к краю модалки).
5. Закрытие drawer: ✕, Esc, клик по backdrop.
6. Esc при открытом drawer закрывает СНАЧАЛА drawer (capture + stopImmediatePropagation),
   не родительскую модалку.
7. z-index drawer > modal (drawer ≥ 1100, modal 1000).
8. Inline-grid с календарём в форме — убран; модалка снова одноколоночная по ёмкости.
9. Backend / check_batches / гейт red→4xx — не трогаем в этой дельте.
10. ПК-only; мобилка out.
→ Correct me now if wrong.
```

---

## Objective

Менеджер на ПК видит узкую модалку (срок / партии) и открывает календарь
загрузки завода **только по кнопке «Ёмкость»** — выезд слева от края экрана —
в обоих критических шагах, без потери red-hint и без ослабления гейта.

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|----------------|---------|--------|
| US-1 | открыл «В производство» | не видеть календарь рядом со сроком | форма оставалась узкой |
| US-2 | нажал «Ёмкость» | получить drawer слева с панелью завода | оценить загрузку визуально |
| US-3 | срок red, drawer закрыт | видеть короткий hint в модалке и disabled submit | понять, что менять, без открытия календаря |
| US-4 | закрываю drawer кликом / Esc / ✕ | вернуться к модалке | не закрыть всю операцию случайно Esc |

### Acceptance criteria

- [x] В `MoveToProductionDialog` и `DeliveryScheduleDialog` нет inline `FactoryCapacityPanel` в grid формы
- [x] Кнопка «Ёмкость» открывает `Drawer side="left"` с `FactoryCapacityPanel`
- [x] При red hint виден в модалке при закрытом drawer; submit/save disabled
- [x] Закрытие: ✕, Esc (drawer first), клик по backdrop
- [x] Drawer у левого края viewport; z-index выше модалки
- [x] Родительский gate (backend + `isCapacityRed`) без регрессий
- [ ] Human QA: оба entry points на живых данных

---

## Tech Stack

| Слой | Стек |
|------|------|
| Frontend | React 19, TS, Vite, Vitest, TanStack Query |
| UI | существующий `shared/ui/Drawer` + `side="left"`; `FactoryCapacityPanel` без смены контракта |
| Backend | без изменений в этой дельте |

---

## Commands

```bash
cd frontend && npm run test -- --run \
  src/features/commercial-archive/components/MoveToProductionDialog.test.tsx \
  src/features/delivery-schedule/components/DeliveryScheduleDialog.test.tsx \
  src/features/factory-capacity
cd frontend && npm run typecheck
cd frontend && npm run build
# Full local (уже часто крутится)
./run+logs.sh
```

---

## Project Structure

```
frontend/src/shared/ui/Drawer.tsx          # side?: "left" | "right"; Esc capture
frontend/src/index.css                     # .app-drawer--left, z-index 1100
frontend/src/features/factory-capacity/
  components/FactoryCapacityPanel.tsx      # без смены API; живёт внутри drawer
frontend/src/features/commercial-archive/
  components/MoveToProductionDialog.tsx    # кнопка + drawer + hint в модалке
frontend/src/features/delivery-schedule/
  components/DeliveryScheduleDialog.tsx    # то же
ai_docs/ideas/zavod-emkost-left-drawer.md
ai_docs/specs/zavod-emkost-left-drawer.md  # этот файл
```

Не добавляем новую feature-folder: reuse `factory-capacity` + `Drawer`.

---

## Code Style

- Кнопка: текст `Ёмкость`, `variant="ghost"`, рядом с Отмена/Сохранить.
- Drawer: `title="Ёмкость завода"`, `side="left"`, ширина ~380px.
- Hint в модалке при red: тот же смысл, что в панели
  (`{hint}. Увеличьте срок…` / для графика — hint или текст про красную партию).
- Не дублировать вторую сетку календаря.

Пример wiring:

```tsx
<>
  <Modal open={open} onClose={handleModalClose} …>
    {/* form + red Alert when blocked */}
    <Button type="button" variant="ghost" onClick={() => setCapacityOpen(true)}>
      Ёмкость
    </Button>
  </Modal>
  <Drawer
    open={open && capacityOpen}
    onClose={() => setCapacityOpen(false)}
    title="Ёмкость завода"
    side="left"
    width={380}
  >
    <FactoryCapacityPanel snapshot={…} isLoading={…} errorMessage={…} />
  </Drawer>
</>
```

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Vitest dialog | red → hint в модалке; panel отсутствует до клика; после «Ёмкость» — panel + mini-calendar; submit disabled |
| Vitest panel | без регрессий шапки/календаря/isCapacityRed |
| Manual | оба диалога: open/close ✕/Esc/backdrop; Esc не закрывает модалку пока drawer открыт |
| Backend | не требуется для этой дельты (регресс gate — по желанию smoke) |

---

## Boundaries

**Always:**
- Один паттерн на оба entry point
- Hint в модалке при red при закрытом drawer
- Esc сначала закрывает drawer
- Не ослаблять backend gate

**Ask first:**
- Бейдж статуса на кнопке «Ёмкость»
- Авто-open drawer при red
- Смена ширины/анимации глобального Drawer, ломающая GSM/Production drawers

**Never:**
- Вернуть inline-календарь в grid формы
- Drawer к краю модалки вместо viewport
- Менять формулу ёмкости / СГП / override в этой дельте

---

## Success Criteria

1. Модалка «В производство» снова читается как одна колонка срока; календарь только после «Ёмкость» слева.
2. То же в графике поставок.
3. Red: hint виден без drawer; сохранить нельзя.
4. Esc / backdrop / ✕ закрывают drawer; повторный Esc — модалку.
5. Vitest MoveToProduction capacity-drawer сценарий зелёный.

---

## Relation to parent spec

Родительская [`zavod-emkost-vizual-gate.md`](zavod-emkost-vizual-gate.md) acceptance
«виджет виден в диалоге» уточняется:

| Было (v1 UI) | Стало (эта спека) |
|--------------|-------------------|
| Панель inline справа от формы | Панель в left drawer по кнопке |
| Календарь всегда на экране диалога | Календарь по запросу; hint при red всегда в модалке |

Алгоритм, API snapshot, enforce — без изменений.

---

## Not Doing

- Авто-open при red  
- Бейдж на кнопке  
- Прилипание к модалке  
- Mobile layout  
- Новые API / таблицы  

---

## Open Questions

_Нет блокирующих._

---

## Next step

Plan: [`../develop/plans/2026-08-26-zavod-emkost-left-drawer.md`](../develop/plans/2026-08-26-zavod-emkost-left-drawer.md)  
Orch: `orch-2026-08-26-12-19-zavod-emkost-left-drawer` — LDR-001 (stale DeliverySchedule test)
→ verify → human QA → docs. Коммит только по просьбе.
