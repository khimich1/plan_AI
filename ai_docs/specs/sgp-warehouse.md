# Spec: Склад готовой продукции (СГП)

> **Источник идеи:** [`ai_docs/ideas/sgp-warehouse.md`](../ideas/sgp-warehouse.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS/IMPLEMENT  
> **План:** [`ai_docs/develop/plans/2026-07-27-sgp-warehouse.md`](../develop/plans/2026-07-27-sgp-warehouse.md)  
> **Дата:** 2026-07-27  
> **Статус:** ✅ implemented (MVP)  
> **Инварианты qty:** gate как в [`fix-orphan-kp-plates-commit.md`](./fix-orphan-kp-plates-commit.md); актуальный PASS — [`reports/plate_loss_regression_roman_20260503.md`](../../reports/plate_loss_regression_roman_20260503.md) (orphan Σ=0). Файл `…_120437.md` — pre-fix snapshot, не опираться.  
> **Связанные:** `ProductionCompletionService`, `completed_plates`, `kp_plates`, `plate_status_log`, `DayDrawer`, `ProductionTabs`, wizard планирования

---

## Decisions locked (Q1–Q10)

| # | Тема | Решение |
|---|------|---------|
| 1 | Таблица | Расширить `completed_plates` (`kp_id` NULLABLE), **без** rename |
| 2 | API дня | Оставить `POST …/days/{date}/complete`; UI/семантика = send to SGP |
| 3 | Бейдж N/M | **M** = исходный заказной qty КП; **N** = qty на СГП с `kp_id` |
| 4 | Orphan при send | **Fail весь день (422)**; страховка поверх уже починенного commit |
| 5 | Free plates → commit | `sgp_reservations[]` в одном `buildPlan`; UX: система **предлагает** «закрыть со склада», мастер **подтверждает** |
| 6 | Scope MVP | Ядро склада **и** wizard close-from-SGP — **в одном MVP** |
| 7 | Unlink в wizard | Потребность видна + **бейдж/подсказка** «есть N свободных — закрыть со склада?» |
| 7+ | Рез донора с СГП | **Фаза 2** (не MVP): предложение реза длина/ширина + донор в плане |
| 8 | Удаление плана | Плиты на СГП **остаются**; plan-связи снимаются |
| 9 | Документы дорожки | «с СГП» **не** в схеме/формовке; **+ кнопка плана** «Со склада (XLSX)» |
| 10 | Feature-flag | **Нет** — сразу новый UX |

Прочие (из ideation, без изменения):

- Роли: мастер делает send/unlink/relink.
- `check_and_update` → «На СГП», не «выполнено».
- Отгрузка / «выполнено» — OUT of MVP.

---

## Objective

Разделить **физический склад** (реально произведённые плиты, ещё не у клиента) и **потребность КП** (`kp_plates`), чтобы:

1. Мастер отправлял день на **СГП** вместо «отметить выполненным».
2. На вкладке **Склад готовой продукции** видел все плиты, фильтровал с КП / без КП.
3. Мог **отвязать** (qty) и **перепривязать** (strict match) без потери штук.
4. В wizard мог **учесть свободные плиты на СГП** и закрыть потребность без повторного производства.

### User stories

| # | Как мастер… | Я хочу… | Чтобы… |
|---|-------------|---------|--------|
| US-1 | закончил день на дорожках | нажать «Отправить на СГП» | плиты ушли на склад, а не «исчезли в выполнено» |
| US-2 | открыл производство | видеть вкладку СГП | знать, что лежит на складе |
| US-3 | ошибся с КП / поменялся заказ | отвязать qty от КП | потребность снова в пуле, физика осталась на СГП |
| US-4 | нашёл подходящий другой заказ | перепривязать плиту к КП-2 | закрыть потребность КП-2 без нового производства |
| US-5 | планирую новый план | выбрать свободные плиты с СГП | не класть их снова на дорожку |
| US-6 | смотрю КП в архиве | видеть «N/M на СГП» и статус «На СГП» | понимать прогресс без Excel |

### Acceptance criteria (MVP)

- [x] Кнопка дня: **«Отправить на СГП»** (вместо «Отметить выполненным»); при успехе день в `completed_days`, плиты в `completed_plates` со статусом/семантикой on_sgp
- [x] Брак: годные → СГП, брак → «в производстве» (как сейчас `return_rejected`)
- [x] Вкладка **«Склад готовой продукции»** + плоская таблица + фильтры: все / с КП / без КП
- [x] Отвязка с указанием qty → split строки; `kp_id=NULL` на СГП; +qty в `kp_plates` «в производстве»
- [x] Перепривязка plate-first, strict match (`plate_name`+length+width+load_class), только на КП с открытой позицией; закрывает потребность целевого КП
- [x] `kp_meta.status = «На СГП»` автоматически, когда все позиции КП на СГП **и** привязаны
- [x] Бейдж **«N/M на СГП»** в архиве, wizard, день производства, список СГП
- [x] Wizard: подсказка «есть N свободных» на позиции; кнопка **«Закрыть со склада»** (propose → confirm); `sgp_reservations[]` в buildPlan; строка **«с СГП»** без дорожки
- [x] Документы дорожки без «с СГП»; отдельная кнопка плана **«Со склада (XLSX)»**
- [x] Удаление плана: плиты на СГП остаются; plan-привязки снимаются
- [x] **Qty-инвариант** (plate_loss gate): никакая операция СГП не «теряет» qty из учёта
- [x] Orphan pre-flight при send → 422 на весь день

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | FastAPI, SQLite (`plita.db`), Pydantic v2 |
| Domain | `app/domain/enums.py`, `core/plate_completion_service.py`, `core/kp_db_plates_*` |
| API | `app/api/v1/endpoints/production.py` (+ новый sgp router или секция) |
| Frontend | React, Vite, TypeScript, TanStack Query |
| Tests | pytest (`tests/`), Vitest (`frontend/`) |
| Regression | `scripts/run_plate_loss_regression.py` |

---

## Commands

```bash
# Backend
source .venv/bin/activate
uvicorn app.main:app --reload
pytest tests/test_production_completion_service.py -q
pytest tests/test_plan_consistency.py -q
pytest tests/ -k "sgp or completion or kp_completion" -q

# Qty gate (как в orphan-спеке / отчёте)
./.venv/bin/python scripts/run_plate_loss_regression.py
# ожидание: баланс OK; для СГП-тестов — отдельные assert на SUM qty

# Frontend
cd frontend && npm run dev
cd frontend && npm test -- --run
cd frontend && npm run build
```

---

## Project Structure

```
app/domain/enums.py                    → KpStatus.ON_SGP; PlateStatus.ON_SGP;
                                         PlateTransitionReason.SGP_*
core/kp_db_schema.py                   → completed_plates.kp_id NULLABLE (+ миграция)
core/kp_db_plates_common.py            → _record_plate_completion → on_sgp semantics
core/kp_db_plates_completion.py        → check_and_update → «На СГП» вместо «выполнено»
app/services/production_completion_service.py  → send_to_sgp (alias complete_day)
app/services/sgp_service.py            → NEW: list / unlink / relink / free_plates
app/repositories/...                   → при необходимости тонкий слой SQL
app/api/v1/endpoints/production.py     → complete = send_to_sgp; + /sgp endpoints
app/schemas/sgp.py                     → NEW: request/response models
app/schemas/production.py              → write_off / labels если нужно

frontend/src/features/production/
  components/ProductionTabs.tsx        → вкладка «Склад готовой продукции»
  components/DayDrawer.tsx             → «Отправить на СГП»
  components/SgpWarehouseView.tsx      → NEW: таблица + фильтры + actions
  components/create-plan-wizard/...    → секция свободных плит
  api/productionApi.ts | sgpApi.ts     → list/unlink/relink/free
  types/production.ts                  → ProductionTab += "sgp"

frontend/src/features/commercial-archive/
  → бейдж «N/M на СГП», статус «На СГП» в «В производстве»

tests/
  test_sgp_service.py                  → NEW: unlink/relink/split/invariants
  test_production_completion_service.py → обновить ожидания статуса КП
  test_sgp_qty_balance.py              → NEW: баланс qty как plate_loss уровень C

ai_docs/ideas/sgp-warehouse.md         → идея (источник)
ai_docs/specs/sgp-warehouse.md         → эта спека
ai_docs/develop/plans/…                → появится после approval (Phase 2)
```

---

## Code Style

Пример доменного перехода (ориентир по существующему `_record_plate_completion`):

```python
# core — одна транзакция, audit рядом с UPDATE/INSERT
def record_sgp_send(cur, *, row_id, kp_id, plate_name, deduct, plan_id, day_number, actor):
    _deduct_kp_plate_qty(cur, row_id, deduct)
    _insert_completed_plate(cur, kp_id=kp_id, ..., qty=deduct)  # kp_id may be set
    audit_append(
        cur,
        plate_id=row_id,
        kp_id=kp_id,
        plate_name=plate_name,
        plan_id=plan_id,
        day_number=day_number,
        from_status=PlateStatus.IN_PLAN.value,
        to_status=PlateStatus.ON_SGP.value,  # "on_sgp"
        qty=deduct,
        reason=PlateTransitionReason.SGP_SEND.value,
        actor=actor,
    )
```

- Слои: router → service → repository/core SQL; не смешивать ORM-логику в endpoint.
- Enum-значения — единая точка правды в `app/domain/enums.py`.
- Частичные qty → **split** (как `insert_kp_plate_remainder_row`), не silent overwrite.
- Сообщения API на русском для UI; коды ошибок стабильные (`sgp_no_matching_demand`, `sgp_strict_mismatch`).

---

## Domain model

### Два регистра

| Регистр | Таблица | Смысл |
|---------|--------|-------|
| Потребность | `kp_plates` | что ещё нужно произвести / в плане |
| Физика | `completed_plates` (СГП) | что реально лежит на складе |

### Статусы

| Enum | Значение | Где |
|------|----------|-----|
| `PlateStatus.ON_SGP` | `on_sgp` | audit `to_status`; UI «на СГП» |
| `KpStatus.ON_SGP` | `На СГП` | `kp_meta.status` |
| Reasons | `sgp_send`, `sgp_unlink`, `sgp_relink`, `sgp_reserve` | `plate_status_log.reason` |

`PlateStatus.COMPLETED` / reason `completed` — deprecate для новых записей; читать старые как on_sgp при необходимости.

### Операции (инварианты)

| Операция | Физика (СГП) | Потребность (`kp_plates`) |
|----------|--------------|---------------------------|
| Send day | +qty, kp_id=X | −qty из «в плане» |
| Unlink qty | split; часть kp_id→NULL | +qty «в производстве» у исходного КП |
| Relink X→Y | kp_id: X→Y (strict) | +qty у X; −qty у Y (закрытие спроса) |
| Wizard reserve | kp_id: NULL→Y | −qty у Y; **без** дорожки |
| Delete plan | строки СГП остаются | plan_id/day связи плана снимаются |

**Инвариант баланса (plate_loss стиль):**

Для каждой номенклатуры / в рамках тестовой фикстуры:

```
Σ qty(kp_plates where status in план|производство)
+ Σ qty(completed_plates on_sgp)
+ Σ qty(отгружено)   # OUT MVP = 0
= исходный заказной qty  (± явный брак/списание, если задокументировано)
```

Каждый переход — строка в `plate_status_log` с `qty`.

### Strict match (relink)

Совпадение: `plate_name`, `length_m`, `width_m`, `load_class`.  
Целевой КП обязан иметь открытую строку «в производстве» с той же identity. Иначе 422 + подсказка.

### Статус КП

```
remaining_in_kp_plates = SUM(qty) WHERE kp_id=?
on_sgp_linked = SUM(qty) FROM completed_plates WHERE kp_id=? AND kp_id IS NOT NULL
on_sgp_unlinked не входит в «закрытие»

IF remaining_in_kp_plates == 0 AND все произведённые qty по заказу покрыты
   linked-строками на СГП (и нет «дыр» потребности):
    → kp_meta = «На СГП»
ELSE:
    → «в работе» (+ бейдж N/M)
```

Точная формула N/M (**Q3 = a**):

- **M** = исходный заказной qty КП (снимок заказа / сумма позиций КП).
- **N** = `SUM(qty)` на СГП с `kp_id = this` (только linked).

После unlink 2 из 10 → бейдж **8/10**, не 8/8. Частично: статус «в работе» + бейдж.

---

## API (черновик контракта)

| Method | Path | Назначение |
|--------|------|------------|
| `POST` | `/api/v1/production/days/{date}/complete` | Send to SGP (существующий путь; body как сейчас + rejected) |
| `GET` | `/api/v1/production/sgp/plates?filter=all\|linked\|unlinked` | Список склада |
| `POST` | `/api/v1/production/sgp/plates/{id}/unlink` | body: `{ qty }` |
| `POST` | `/api/v1/production/sgp/plates/{id}/relink` | body: `{ target_kp_id, qty }` |
| `GET` | `/api/v1/production/sgp/free-plates` | Для wizard (exact match hints) |
| `GET` | `/api/v1/production/plans/{plan_id}/sgp-export` (или `/days/.../sgp`) | XLSX «Со склада» |
| `GET` | `/api/v1/offers/...` или расширить archive DTO | `sgp_progress: { n, m }`, status |
| `POST` | buildPlan + `sgp_reservations[]` | Закрытие со склада в той же транзакции |

Ошибки:

- `422` pre-flight: нет kp_id / orphan mismatch / нет спроса на target
- `409` version conflict плана (как сейчас)

---

## UI

### Production tabs

`ProductionTab` += `"sgp"` → label «Склад готовой продукции».

### DayDrawer

- Текст кнопки: «Отправить на СГП» / «Уже на СГП».
- Success alert: «День отправлен на СГП. Списано: …, брак возвращён: …».

### SgpWarehouseView

Плоская таблица: плита, размер, КП, заказчик, срок, qty, дата, plan/day.  
Фильтры: все / с КП / без КП.  
Actions: Отвязать (qty dialog), Перепривязать (plate-first picker).

### Wizard

1. Потребность КП **всегда видна**.
2. Matching free plates → бейдж: *«на СГП есть N свободных — закрыть со склада?»*
3. Кнопка **«Закрыть со склада»**: система предлагает → мастер подтверждает.
4. В плане строка **«с СГП»** без track; в оптимизатор уходит только остаток потребности.

### Документы плана / дня

| Кнопка | Содержимое |
|--------|------------|
| Схема / Разбивка / Формовка | Только дорожки; **без** позиций «с СГП» |
| **Со склада (XLSX)** (новая) | Позиции плана, закрытые со СГП |

### Archive

КП со статусом «На СГП» остаётся во вкладке **«В производстве»** с бейджем.  
Не переносить в «Выполненные» до отгрузки (OUT).

---

## Testing Strategy

| Уровень | Что | Где |
|---------|-----|-----|
| Unit | split unlink/relink; strict match reject | `tests/test_sgp_service.py` |
| Unit | `check_and_update` → «На СГП», не «выполнено» | update completion tests |
| Integration | `complete_day` = send_to_sgp; брак; audit reasons | `test_production_completion_service.py` |
| Qty balance | SUM до/после unlink+relink = const | `test_sgp_qty_balance.py` |
| Plan consistency | day_view invariant после send | `test_plan_consistency.py` |
| E2E regression | `run_plate_loss_regression.py` не регрессирует (orphan/баланс) | script + CI local |
| Frontend | tab sgp; filters; button label | vitest / manual |
| Manual | wizard free plates → план без дорожки; archive badge | browser |

**Связь с plate_loss / orphan:**

Commit orphan **починен** (PASS на `roman_20260503`, orphan Σ=0). Для СГП всё равно:

1. Pre-flight: расхождение kp_plates vs day_view → **422 на весь день** (страховка).
2. Unlink/relink не создают «учёт без физики» / «физика без учёта».
3. Тест: после send — `Σ moved` = `Σ planned − rejected`; orphan не растёт.

---

## Boundaries

### Always

- Атомарные транзакции для send / unlink / relink / reserve (как P0 complete_day)
- Писать `plate_status_log` в той же транзакции
- Split при частичном qty
- Прогонять completion + sgp qty tests перед merge
- Сохранять инвариант: физика на СГП + потребность в kp_plates согласованы после каждой операции

### Ask first

- Rename таблицы `completed_plates` → `sgp_inventory`
- Смена URL `/complete` на `/send-to-sgp`
- Отдельные RBAC-роли для unlink/relink
- Включение перепила / отгрузки в этот же PR
- Изменение поведения удаления плана (сейчас: СГП остаётся)

### Never

- Ставить `kp_meta = «выполнено»` только потому что kp_plates пуст
- Хранить «на СГП» как status в `kp_plates` (смешение регистров)
- Перепривязка без strict match в MVP
- Тихий skip плит без kp_id при send (сейчас — ошибка; оставить)
- Удалять audit-историю при split
- Коммитить секреты / `.env`

---

## Success Criteria

Конкретные, проверяемые:

1. **UI дня:** кнопка «Отправить на СГП»; после клика день `completed`; позиции в списке СГП с тем же qty (минус брак).
2. **Unlink 2 of 5:** на СГП две строки (3 linked + 2 unlinked) или эквивалентный split; у КП +2 «в производстве»; `plate_status_log` содержит `sgp_unlink` qty=2.
3. **Relink:** unlinked плита → КП-2 с matching demand → kp_id=2; у КП-2 −qty потребности; у КП-1 без изменений сверх уже сделанного unlink.
4. **Strict reject:** relink на КП без позиции / другой длины → 422, БД без изменений.
5. **КП status:** все linked на СГП, kp_plates пуст → `На СГП`; после unlink → снова `в работе` + бейдж N/M.
6. **Wizard:** propose/confirm «Закрыть со склада» → «с СГП» без track; потребность закрыта; бейдж на позиции с free match.
7. **Delete plan:** СГП qty не уменьшается.
8. **Qty gate:** баланс-тесты + `run_plate_loss_regression.py` остаётся PASS (orphan Σ=0).
9. **Архив:** КП «На СГП» в «В производстве» с бейджем N/M (M = заказной qty).
10. **Кнопка «Со склада (XLSX)»** у плана; схема/формовка без этих позиций.
11. **Orphan при send:** 422, БД без изменений.

---

## Out of Scope

### Фаза 2 (сразу после MVP, отдельный кусок)

- Предложение **реза** свободной плиты (длиннее/шире) под меньшую потребность
- В плане: донор + способ реза; правила остатка после реза

### Позже / не сейчас

- Перепил с брака на дорожке (двухшаговый сценарий ideation)
- Замена из будущих дней плана / авто-перепланировка
- Отгрузка и кнопка «Выполнено»
- Поштучный QR / qty=1 всегда
- Откат дня после СГП
- Telegram UI для СГП
- True swap двух плит
- Flexible («похожая») перепривязка
- Feature-flag / параллельный старый complete-flow

---

## Open Questions (осталось на Plan)

| # | Вопрос | Статус |
|---|--------|--------|
| P1 | Где хранить снимок **M** (ordered_qty): колонка / пересчёт из истории КП? | Решить на Plan |
| P2 | Точный path экспорта «Со склада» (plan vs day) | Решить на Plan |
| P3 | Внутренний rename метода `complete_day` → `send_to_sgp` или только комментарии | Решить на Plan |

Q1–Q10 из review — **закрыты** (см. Decisions locked).

---

## As-Is → To-Be

| As-Is | To-Be |
|-------|-------|
| «Отметить выполненным» | «Отправить на СГП» (без feature-flag) |
| `completed_plates` = конечная «выполнено» | СГП-инвентарь, `kp_id` nullable |
| `check_and_update` → «выполнено» | → «На СГП» (all linked) |
| Нет UI склада | Вкладка СГП + фильтры + unlink/relink |
| Нет возврата потребности с склада | Unlink → потребность + free на СГП |
| Нет учёта свободных в wizard | Propose/confirm «Закрыть со склада» + строка «с СГП» |
| Документы = только дорожки | + кнопка «Со склада (XLSX)» |
| Orphan commit bug | Починен; send всё равно fail при mismatch |

---

## Verification (конец Phase SPECIFY)

- [x] Spec covers Objective, Commands, Structure, Style, Testing, Boundaries
- [x] Human reviewed Q1–Q10
- [x] Decisions locked in spec
- [ ] Success criteria — финальное «ок, можно PLAN»
- [ ] **Stop:** не IMPLEMENT до Plan + Tasks

---

## Next

1. ~~Phase PLAN~~ → [`ai_docs/develop/plans/2026-07-27-sgp-warehouse.md`](../develop/plans/2026-07-27-sgp-warehouse.md)
2. Tasks = чеклист SGP-000…604 в плане
3. Implement с SGP-000 (TDD, по checkpoint’ам)
