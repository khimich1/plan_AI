# Handoff: КП archive-edit + source photos + soft-lock fixes (2026-09-02)

> **Дата:** 2026-09-02  
> **Ветка:** текущая рабочая · **изменения в working tree, коммитов по этой сессии не было**  
> **Статус:** archive-edit ✅ · source-image queue ✅ · follow-up bugs (blob / confirm-batch) ✅ · Result→input block ✅  
> **Цель файла:** продолжить с чистого контекста — smoke / review / commit / split PR, **без** повторного ideation по уже locked решениям.  
> **Не коммитить** без явной просьбы пользователя.  
> **Не убивать** `./run+logs.sh`, если запущен.

---

## Как стартовать новую сессию (скопируй в первый промпт)

```
Прочитай handoff:
ai_docs/develop/handoffs/2026-09-02-kp-archive-edit-and-source-photos.md

Контекст уже IMPLEMENT ✅. Не /idea-refine заново по этим темам.

Нужно (выбери задачу):
A) Manual smoke archive-edit + «Исходные фото» + «Список верен» → «Готово, далее»
B) Code review + точечные тесты на confirm-batch-review (multi-page)
C) Подготовить commit(ы) / split PR — только по просьбе; не коммить секреты/plita.db/test_kp.pdf

Skills: plan-web-context. Не трогать run+logs.sh.
```

### Чеклист агента

1. Прочитать **этот** handoff.  
2. Спеки/планы в таблице ниже — статус IMPLEMENT ✅.  
3. Не переворачивать гейт статуса обратно на «в работе».  
4. При правках фото-очереди: promote из `File` (не шарить blob URL с preview).  
5. Multi-page «Список верен» **обязан** диспатчить `confirm-batch-review`.  
6. Не коммитить `.env` / `plita.db` / password changes.

---

## Артефакты

| Тема | Idea | Spec | Plan | Статус |
|------|------|------|------|--------|
| Правка архива через конструктор | [`ideas/kp-archive-edit-in-constructor.md`](../../ideas/kp-archive-edit-in-constructor.md) | [`specs/kp-archive-edit-in-constructor.md`](../../specs/kp-archive-edit-in-constructor.md) | [`plans/2026-09-02-kp-archive-edit-in-constructor.md`](../plans/2026-09-02-kp-archive-edit-in-constructor.md) | IMPLEMENT ✅ |
| Очередь исходных фото (Drawer) | [`ideas/kp-source-image-queue-drawer.md`](../../ideas/kp-source-image-queue-drawer.md) | [`specs/kp-source-image-queue-drawer.md`](../../specs/kp-source-image-queue-drawer.md) | [`plans/2026-09-02-kp-source-image-queue-drawer.md`](../plans/2026-09-02-kp-source-image-queue-drawer.md) | IMPLEMENT ✅ |
| Related (уже в tree с этой же даты) | append-preview, archive-only-save, breakdown-xlsx, multi-type-picker | см. `ai_docs/specs/kp-*.md` / `plans/2026-09-02-*` | | IMPLEMENT ✅ (отдельные срезы) |

**Orchestration ids:**  
- `orch-2026-09-02-09-57-kp-archive-edit`  
- `orch-2026-09-02-11-33-kp-source-image-queue`

---

## Что сделано в этой сессии

### A. Архив → конструктор (ARC-001…008)

**Поведение**
- Resume/update **только** статус «в архиве» (три гейта: hydrate, `save_offer` update, `KpPersistenceService.update_kp_from_order_data`).  
- «в работе» / производство — **без** CTA правок состава.  
- Drawer «в архиве»:  
  - **(+ Добавить)** → resume + `start-append-cycle` → picker  
  - **Редактировать** → resume → Result  
  - кнопки у шапки **Итоги**, variant primary  
- Итоги в архиве — **read-only** (без скидки/рейса/целевой суммы).  
- После save — тот же `kp_id`, статус остаётся «в архиве».  
- Nav: **«Конструктор КП»** (`AppHeader`).  
- Import navigate: `from "react-router"` (не `react-router-dom`).

**Follow-up UX (после первого IMPLEMENT)**
- `SaveOfferSection`: при `resume_kp_id` — активная **«Сохранить изменения»** (раньше блокировалось `saved_offer` → «Сохранено»).  
- После успешного resume-save → invalidate archive queries → `reset` wizard → `navigate("/archive")`.

**Ключевые файлы**
- `core/kp_persistence_service.py`  
- `app/services/commercial_draft_lifecycle.py`  
- `frontend/.../OfferDetailsDrawer.tsx`  
- `frontend/.../SaveOfferSection.tsx`  
- `frontend/.../CommercialOfferWizard.tsx` (`handleSave`)  
- `frontend/src/app/layout/AppHeader.tsx`

### B. Soft-lock: Result → шаг 1

- С Result **нельзя** открыть input step без нового круга (`isInputStepBlockedWithoutAppendCycle`).  
- «Назад» на Result: при `skipClient` скрыта; иначе только на client.  
- Новый круг — «Добавить другое наименование».

**Файлы:** `wizardStepOrder.ts`, `CommercialOfferWizard.tsx`, `CalculationResultStep.tsx`

### C. Очередь исходных фото (IMG-001…008)

**Поведение**
- После «Список верен» на шаге 1: CTA **«Исходные фото (N)»** → left `Drawer` + pager.  
- Promote-then-reset: snapshot в `useSourceImageQueue` **до** `multiPage.reset` / clear preview.  
- Clear очереди: новый круг / archive save / create-new / text-abandon.  
- FE-only (blob); F5 = пусто.

**Ключевые файлы**
- `lib/sourceImageQueue.ts`, `lib/promoteSourceImageQueue.ts`  
- `hooks/useSourceImageQueue.ts`, `hooks/useRecognizedImagePreview.ts`  
- `components/SourceImageQueueDrawer.tsx`, `SourceImageQueueControls.tsx`  
- wiring во все 6 `*InputStep` + wizard

### D. Hotfixes после smoke (обязательно знать)

| Симптом | Причина | Фикс |
|---------|---------|------|
| После «Список верен» «Готово, далее» серая / «лишнее окно» сверки без фото | Multi-page finalize **не** диспатчил `confirm-batch-review` → `pendingBatchReview` оставался true | В `handleConfirmBatch` multi-path после promote/reset → `dispatch({ type: "confirm-batch-review", ... })` |
| Drawer «Исходные фото» — битая картинка | В очередь попадал тот же blob URL, что потом revoke’ился | Preview хранит `file`; promote создаёт **новый** URL из File; `takePreview` не шарит display URL |
| Vite: failed to resolve `react-router-dom` | В проекте пакет `react-router` v8 | Import `useNavigate` from `"react-router"` |

### E. Прочее (не feature)

- Локальный пароль `admin` в `plita.db` обновлён через `AuthRepository` (как `scripts/create_admin.py`). **Не коммитить БД.** Значение пароля — только у пользователя.  
- Жёлтый баннер «Вторая проверка не подтвердила список» — **не блокер**: OCR verify не совпал с extract; список остаётся первым; нужно сверить и нажать «Список верен».

---

## Locked decisions (не переспрашивать)

1. Resume/update только **«в архиве»** (не «в архиве \| в работе»).  
2. Два landing’а: Добавить → picker; Редактировать → Result.  
3. Финансы в archive drawer — только сводка.  
4. Save из resume → тот же `kp_id`, статус «в архиве».  
5. Исходные фото — только шаг 1, Drawer по кнопке, очередь 1..N текущего захода.  
6. Clear фото-очереди на новый круг / archive save / create-new.  
7. Без server storage / IndexedDB для фото в этом MVP.  
8. Result → input только через append-цикл.

---

## Проверка

```bash
# Archive-edit / persistence
pytest tests/test_kp_persistence_mixed.py tests/test_archive_endpoints.py \
  tests/test_commercial_draft_append.py tests/test_commercial_multi_append_flow.py \
  -q -k "resume or archive or update_kp or hydrate or save_offer or sc6"

# FE focused
cd frontend && npm run test -- \
  src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx \
  src/features/commercial-offer/components/SaveOfferSection.test.tsx \
  src/app/layout/AppHeader.test.tsx \
  src/features/commercial-offer/lib/wizardStepOrder.test.ts \
  src/features/commercial-offer/hooks/useSourceImageQueue.test.ts \
  src/features/commercial-offer/hooks/useRecognizedImagePreview.test.ts \
  src/features/commercial-offer/components/SourceImageQueueDrawer.test.tsx \
  src/features/commercial-offer/components/SourceImageQueueControls.test.tsx

cd frontend && npm run typecheck
```

### Manual smoke

1. Архив → КП «в архиве» → **Редактировать** → правка строки → **Сохранить изменения** → возврат в `/archive`, тот же номер.  
2. Архив → **(+ Добавить)** → picker.  
3. Конструктор: фото (1 или N) → сверка → «Список верен» → CTA «Исходные фото» показывает картинку → **Готово, далее** активна.  
4. С Result сайдбар «1. Плиты» недоступен без «Добавить другое наименование».

---

## Риски / leftover

- Working tree **смешан** с другими фичами 2026-09-02 (append-preview, archive-only-save, breakdown-xlsx, type-picker) + возможно `test_kp.pdf` / `tsconfig.app.tsbuildinfo` — при commit **splitting** обязателен.  
- Нет отдельного RTL-теста на multi-page `confirm-batch-review` в wizard (стоит добавить при review).  
- Comment в `CommercialOfferHeaderBridge` может ещё говорить «Создать КП».  
- Discount/logistics PATCH API на архивном КП UI не зовёт (hardening out of scope).  
- F5 убивает фото-очередь — by design.

---

## Не делать в следующей сессии

- Не возвращать resume для «в работе».  
- Не хранить OCR-фото на сервере без новой спеки.  
- Не открывать permanent split-view фото на предпросмотре.  
- Не коммитить `plita.db` / пароли.  
- Не запускать полный `/idea-refine` по уже закрытым темам выше.
