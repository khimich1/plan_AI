# Spec: КП — очередь исходных фото на шаге 1 (Drawer)

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-09-02  
**One-pager**: [../ideas/kp-source-image-queue-drawer.md](../ideas/kp-source-image-queue-drawer.md)  
**Related**: multi-page OCR (`useMultiPageRecognize`, `multiPageSource`), `Drawer`, input steps (`PlateInputStep` / peers), archive-edit save reset

## Objective

**Проблема.** На сверке («Список верен») исходное фото видно; после подтверждения на шаге 1 остаётся только предпросмотр состава — картинки нет. При multi-page очередь тоже теряется слишком рано (revoke / `multiPage.reset`).

**Цель.** На **шаге 1** менеджер вызывает очередь исходных фото текущего захода (1..N) в **боковом Drawer**. Очередь живёт до нового круга набора или сохранения в архив. FE-only.

**Пользователь:** менеджер в конструкторе КП на шаге ввода номенклатуры.

**Успех:** после сверки кнопка «Исходные фото (N)» открывает Drawer с листанием; после нового круга / archive save очереди нет.

---

## ASSUMPTIONS I'M MAKING

1. **Только шаг 1** (plates/piles/steps/marches/bridge_piles/fbs input). Не client, не Result, не archive drawer.
2. **Очередь = страницы текущего захода** — и single-file, и multi-page (`PageSource[]` / эквивалент). Text-only без image → кнопки нет.
3. **UX:** кнопка → `Drawer` (side left или right — как удобнее рядом с составом; default **left**, как (i) у picker). Не постоянный split-view.
4. **Lifecycle clear** (revoke blob + empty queue):
   - `start-append-cycle` / начало нового круга (в т.ч. reset перед picker);
   - успешный archive save (`В архив` / `Сохранить изменения` с navigate/reset);
   - «Создать новое КП» / полный `reset` мастера.
5. **Не clear** на «Список верен» / переход к предпросмотру состава на том же шаге 1 / правку клиент↔input на первом проходе (если input ещё доступен). Сейчас `multiPage.reset()` после confirm — **изменить**: сохранить snapshot очереди для просмотра.
6. **Хранение:** только браузер (object URL / in-memory). Без API upload, без IndexedDB в MVP (F5 = очередь пропала — ok).
7. **Drawer содержимое:** текущее фото, имя файла (если есть), zoom/fit по возможности переиспользовать контролы сверки или упрощённый img + «Открыть в новой вкладке»; **листание** prev/next при N>1; счётчик `k / N`.
8. **Компонент:** `frontend/src/shared/ui/Drawer.tsx`.
9. **Без новых npm/pip.** Коммиты — по просьбе. Не убивать `./run+logs.sh`.
10. Single-page path (`useRecognizedImagePreview`) тоже попадает в ту же очередь (1 элемент), не только multiPage.

→ Correct me now or these are locked for PLAN.

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| **D-scope** | Где | Только шаг 1 |
| **D-ux** | Как | Кнопка + Drawer, не permanent block |
| **D-queue** | Что | 1..N страниц текущего захода |
| **D-life** | Clear | Новый круг / archive save / create new |
| **D-store** | Где данные | FE blob only |
| **D-confirm** | После «Список верен» | Очередь **сохраняется** (не revoke сразу) |

---

## User Stories

- Как **менеджер**, подтвердив список с фото, на шаге 1 жму **«Исходные фото»** и снова вижу картинку в боковом окне.
- Как **менеджер**, загрузив **несколько** страниц, листаю их в Drawer.
- Как **менеджер**, начав **новый круг** («Добавить другое наименование»), старая очередь исчезает.
- Как **менеджер**, сохранив КП в архив, очередь очищается.

---

## Tech Stack

| Слой | Стек |
|------|------|
| Frontend | React 19, TS, Vitest, existing `Drawer`, `useMultiPageRecognize`, input steps |
| Backend | без изменений |

## Commands

```
cd frontend && npm run test -- src/features/commercial-offer/hooks/useMultiPageRecognize
cd frontend && npm run test -- src/features/commercial-offer/lib/multiPageSource
cd frontend && npm run test -- src/features/commercial-offer/components/steps
cd frontend && npm run typecheck
```

## Project Structure

```
frontend/src/features/commercial-offer/
  hooks/useMultiPageRecognize.ts     → keep pages after confirm; expose sourceQueue; clear on reset
  hooks/useRecognizedImagePreview.ts → feed single-image into same queue abstraction (or wizard merges)
  lib/sourceImageQueue.ts            → optional pure helpers (index, clear policy) + tests
  components/SourceImageQueueDrawer.tsx  → Drawer UI + pager
  components/steps/*InputStep.tsx    → CTA «Исходные фото (N)» when queue.length > 0
  components/CommercialOfferWizard.tsx → wire queue; clear on append/save/create-new
ai_docs/ideas|specs|develop/plans/...
```

## Code Style

- Не дублировать zoom-lightbox сверки целиком, если достаточно img + open-in-tab + pager.
- Revoke object URLs **только** при clear очереди (не при confirm).
- `aria-label` на prev/next / открытие Drawer.

Пример CTA:

```tsx
{sourceQueue.length > 0 && (
  <Button type="button" variant="secondary" onClick={() => setSourceDrawerOpen(true)}>
    Исходные фото ({sourceQueue.length})
  </Button>
)}
```

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit queue / multiPage | после confirm страницы и URL живы; reset/clear → length 0 + revoke called |
| RTL Drawer | открытие по кнопке; N=2 → next/prev; N=1 → без обязательных стрелок или disabled |
| RTL input step | нет кнопки при пустой очереди; есть при ≥1 |
| Wizard wiring (по возможности) | start-append / create-new чистит очередь |
| typecheck | зелёный |

## Boundaries

- **Always:** clear на новый круг и archive save; не держать permanent preview block на предпросмотре; revoke при clear.
- **Ask first:** IndexedDB; server upload; показ на Result; persist после resume архива.
- **Never:** новые зависимости; хранить фото в SQLite/KP files в этом MVP.

## Success Criteria

| # | Критерий |
|---|----------|
| S1 | После «Список верен» на шаге 1 есть CTA при наличии фото |
| S2 | CTA открывает Drawer с текущим исходником |
| S3 | Multi-page: листание по всей очереди захода |
| S4 | Новый круг → очереди нет, CTA нет |
| S5 | Archive save / create new → очередь cleared |
| S6 | Text-only без image → CTA нет |
| S7 | Focused vitest + typecheck зелёные |

## Out of Scope

- IndexedDB / переживание F5
- Server-side OCR file storage
- Фото на шаге клиента / Result / в архиве
- Permanent split-view на предпросмотре

## Open Questions

_Нет блокирующих — defaults locked; PLAN ✅._

---

**Next:** implement per plan [`../develop/plans/2026-09-02-kp-source-image-queue-drawer.md`](../develop/plans/2026-09-02-kp-source-image-queue-drawer.md) (orch-2026-09-02-11-33-kp-source-image-queue).
