# Spec: КП — правка архива через конструктор

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-09-02  
**One-pager**: [../ideas/kp-archive-edit-in-constructor.md](../ideas/kp-archive-edit-in-constructor.md)  
**Plan**: [../develop/plans/2026-09-02-kp-archive-edit-in-constructor.md](../develop/plans/2026-09-02-kp-archive-edit-in-constructor.md)  
**Related**: [kp-multi-nomenclature-append.md](./kp-multi-nomenclature-append.md) (R2 override), [kp-archive-only-save.md](./kp-archive-only-save.md), [kp-row-edit-delete-icons.md](./kp-row-edit-delete-icons.md), [kp-append-preview-and-fresh-breakdown.md](./kp-append-preview-and-fresh-breakdown.md)

## Objective

**Проблема.** КП в архиве можно смотреть и править скидку/рейс в drawer, но **нельзя** дописать/удалить/переименовать позиции через конструктор: `resume` / save-append разрешены только для «в работе». После archive-only-save «в работе» ≈ производство — как раз то, что править нельзя. Drawer «Итоги» дублирует урезанный шаг 3.

**Цель.** Из карточки КП **«в архиве»** менеджер заходит в конструктор двумя кнопками (допись / правка), меняет состав и финансы там, сохраняет в **тот же** `kp_id` со статусом **«в архиве»**. Drawer показывает только сводку. Nav: «Конструктор КП».

**Пользователь:** менеджер в Archive → drawer → конструктор.

**Успех:** архивное КП открывается в конструкторе; после «В архив» номер и статус «в архиве» сохранены; производство без CTA правок; финансы в drawer read-only.

---

## ASSUMPTIONS I'M MAKING

1. **Статус-модель.** «в архиве» = можно править. «в работе» / СГП / выполнено = производство и дальше — **не** правим состав через resume.
2. **Override R2** из [kp-multi-nomenclature-append](./kp-multi-nomenclature-append.md): было «resume только в работе» → теперь **resume/update только «в архиве»**.
3. **Два CTA, один API resume.** Оба вызывают `POST .../archive/{kp_id}/resume`. Разница только FE:
   - **(+ Добавить)** → `hydrate-draft` + `start-append-cycle` → picker (скр. 1)
   - **Редактировать** → `hydrate-draft` (metadata уже `current_step=result`) → шаг 3 (скр. 2)
4. **Итоги в drawer** для «в архиве»: вес, НДС, доставка, итого — **только текст**. Убрать инпуты/OK для целевой суммы, рейса, скидки. Мутации discount/logistics из drawer не вызываем.
5. **Save.** Resume-draft с `resume_kp_id` + mode `archive` → `update_offer_from_order_data` на том же `kp_id`, статус остаётся **«в архиве»** (не «в работе»).
6. **Убрать** кнопку «Добавить другое наименование» при `status === "в работе"` в drawer (больше не нужна).
7. **На Result** конструктора кнопка «Добавить другое наименование» / append loop **остаётся** (сессия в конструкторе).
8. **Nav:** во всех местах UI «Создать КП» → **«Конструктор КП»** (как минимум `AppHeader`; тесты/aria по тексту обновить).
9. **Без новых** npm/pip. Коммиты — только по просьбе. Не убивать `./run+logs.sh`.
10. **API discount/logistics PATCH** на архивном КП можно оставить (не hardening в MVP) — UI просто не зовёт.

→ Correct me now or these are locked for PLAN.

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| **D-status** | Кого правим | Только «в архиве» |
| **D-cta** | Две кнопки | (+ Добавить) → picker; Редактировать → Result |
| **D-drawer** | Итоги | Read-only сводка; без finance edits |
| **D-save** | После правки | Тот же `kp_id`, статус «в архиве» |
| **D-prod** | «в работе» | Без resume CTA |
| **D-nav** | Шапка | «Конструктор КП» |
| **D-api** | Discount PATCH | UI-only stop; API as-is |

---

## User Stories

- Как **менеджер**, в drawer КП «в архиве» жму **(+ Добавить)** и попадаю на выбор номенклатуры, чтобы дописать позиции.
- Как **менеджер**, жму **Редактировать** и попадаю на шаг 3 с полным составом, чтобы удалить/переименовать строки и править скидку/рейс.
- Как **менеджер**, после «В архив» вижу то же КП №N со статусом «в архиве» и обновлённым составом/суммой.
- Как **менеджер**, в drawer вижу итоги **без** полей скидки/рейса/целевой суммы.
- Как **менеджер**, у КП в производстве **нет** кнопок правки состава.
- Как **менеджер**, в шапке вижу пункт **«Конструктор КП»** вместо «Создать КП».

---

## Tech Stack

| Слой | Стек |
|------|------|
| Frontend | React 19, TS, Vite, Vitest + Testing Library, TanStack Query, React Router |
| Backend | FastAPI, `ArchiveService.resume_as_draft`, `CommercialDraftLifecycle.hydrate_draft_from_saved_kp` / `save_offer` |
| API | `POST /api/v1/commercial/archive/{kp_id}/resume`; save draft `mode: archive` |

Новых endpoint’ов нет — меняется статус-гейт и FE landing.

## Commands

```
# Frontend
cd frontend && npm run test -- src/features/commercial-archive/components/OfferDetailsDrawer
cd frontend && npm run test -- src/app/layout
cd frontend && npm run typecheck

# Backend
pytest tests/test_archive_endpoints.py -q -k resume
pytest tests/test_commercial_web_flow.py -q -k "resume or archive or save"
# (+ точечные тесты hydrate/save_offer status gate в draft lifecycle)
```

Dev: не убивать `./run+logs.sh`.

## Project Structure

```
frontend/src/app/layout/AppHeader.tsx
frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx
  → (+ Добавить) / Редактировать для «в архиве»
  → убрать resume CTA для «в работе»
  → Итоги read-only
frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx
app/services/commercial_draft_lifecycle.py
  → hydrate: status == «в архиве»
  → save_offer update path: allow «в архиве», persist «в архиве»
app/services/archive_service.py          → docstring/гейт согласовать
tests/test_archive_endpoints.py
ai_docs/ideas|specs|develop/plans/...
```

## Code Style

- Русские `detail` / `ValueError` сообщения в духе существующих («Дополнить КП можно только…»).
- Один `handleResume*` с параметром landing (`append` | `result`), без копипасты двух почти одинаковых handlers.
- Не трогать production move / readiness без нужды.

Пример (FE):

```ts
async function openInConstructor(landing: "append" | "result") {
  const draft = await archiveApi.resume(offer.kp_id);
  dispatch({ type: "hydrate-draft", payload: draft });
  if (landing === "append") {
    dispatch({ type: "start-append-cycle" });
  }
  navigate(`/new?draft=${encodeURIComponent(draft.draft_id)}`);
  onClose();
}
```

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Pytest hydrate | «в архиве» → 200 + `resume_kp_id`; «в работе» → 4xx/ошибка |
| Pytest save | resume draft + archive save → тот же `kp_id`, status «в архиве» |
| RTL Drawer | «в архиве»: видны (+ Добавить) и Редактировать; нет finance OK/inputs |
| RTL Drawer | «в работе»: нет этих CTA (и нет старой «Добавить другое наименование») |
| RTL/unit Header | текст «Конструктор КП» |
| typecheck | зелёный |

## Boundaries

- **Always:** тот же `kp_id`; статус после save = «в архиве»; производство без edit CTA; read-only Итоги в drawer для архива.
- **Ask first:** запрет discount/logistics API на архивных КП; правки для других статусов; redesign drawer layout.
- **Never:** новые зависимости; правка состава для производства в этом MVP; убивать `./run+logs.sh`; коммит без просьбы.

## Success Criteria

| # | Критерий |
|---|----------|
| S1 | Drawer «в архиве»: кнопки **(+ Добавить)** и **Редактировать** |
| S2 | (+ Добавить) → picker (append cycle) |
| S3 | Редактировать → шаг 3 Result с составом |
| S4 | Save «В архив» → тот же `kp_id`, status «в архиве», состав/суммы обновлены |
| S5 | Drawer «в архиве»: нет редактирования скидки/рейса/целевой суммы |
| S6 | Drawer «в работе»: нет resume/edit состава |
| S7 | Nav label = «Конструктор КП» |
| S8 | Backend гейт: resume/update только «в архиве» |
| S9 | Focused vitest + pytest + typecheck зелёные |

## Out of Scope

- Инлайн rename/delete в таблице «Состав заказа» drawer
- Версии PDF/XLSX
- Hardening PATCH discount/logistics
- Изменение MoveToProduction / readiness
- Редактирование КП после производства

## Open Questions

_Нет блокирующих — assumptions locked 2026-09-02. Plan готов._

---

**Done:** ARC-001…ARC-008 implemented; status gate only «в архиве»; dual CTA + read-only Итоги + nav rename.
