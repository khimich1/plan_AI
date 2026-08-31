# Spec addendum: A — Ожидаемая дата на СГП (100% запланировано)

> **Родительская spec:** [`kp-readiness-manager-view.md`](./kp-readiness-manager-view.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → IMPLEMENT ✅  
> **Дата:** 2026-07-28  
> **Триггер:** менеджер видит, **когда весь заказ окажется на СГП**, если всё уже расписано по плану.

---

## Assumptions I'm Making

1. **Read-only** — как MVP readiness; без мутаций БД.
2. **«100% запланировано»** = нет qty в `kp_plates` со статусом **«в производстве»** (всё либо **«в плане»**, либо уже на СГП). Эквивалент: `remaining_total == 0` при `n < m`.
3. **Дата** = календарный **последний производственный день** среди строк `kp_plates` со статусом **«в плане»**, с резолвом `plan_id` + `day_number` → дата через план.
4. **Формат даты:** `ДД.ММ.ГГГГ` (как в КП/PDF).
5. **Семантика дня:** плиты попадают на СГП **в день завершения дорожки** (как сейчас send-to-SGP); дата последнего дня = ориентир «всё на складе».
6. **Не показывать**, если `n == m` (уже полностью на СГП) или дату не удалось вычислить.
7. **Auth** — как архив (admin/manager).

→ Correct me now or proceed to Plan.

---

## Decisions locked (Q1–Q4)

| # | Тема | Решение |
|---|------|---------|
| 1 | Критерий «100% запланировано» | Нет «в производстве» без плана: `remaining_total == 0` и `n < m` |
| 2 | Источник даты | `MAX(date)` по `kp_plates` «в плане» с `plan_id` + `day_number` |
| 3 | UI | Отдельная строка под `summary_text`: **«Ожидаем на СГП к: ДД.ММ.ГГГГ»** |
| 4 | Copy | Дата **включается** в `client_copy_text`, когда показана в UI |

---

## Objective

Когда менеджер открывает readiness и заказ **полностью расписан** (ничего не висит «в производстве» без плана), но **ещё не целиком на СГП**, показать **ориентировочную дату**, к которой остаток окажется на складе — без звонка в цех.

### User story

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | вижу 307/553 на СГП и 0 «осталось» вне плана | строку «Ожидаем на СГП к: …» | сказать клиенту, когда ждать полный комплект |

### Acceptance criteria

- [x] При `remaining_total == 0`, `n < m` и успешном резолве дат → `expected_sgp_date` + `expected_sgp_date_label` в `KpReadinessSummary`.
- [x] UI: строка под summary (не в hint степпера).
- [x] `client_copy_text` дополняется фразой с датой (см. шаблоны).
- [x] Если `remaining_total > 0` → поля `null`, строка не показывается.
- [x] Если `n == m` → поля `null` (заказ уже на СГП).
- [x] Если у «в плане» нет `plan_id`/`day_number` или план не найден → `null`, без ошибки API.
- [x] Read-only; plate_loss regression PASS.

---

## Domain logic

### Условие показа (`fully_scheduled`)

```python
fully_scheduled = (
    remaining_total == 0
    and progress.n < progress.m
    and in_plan_total > 0  # ещё есть что производить/отправить
)
```

Где `remaining_total` / `in_plan_total` — суммы из `list_positions()` (статус «в производстве» / «в плане»).

**Смысл для менеджера:** каждая позиция либо уже на СГП, либо стоит в плане; «хвостов» вне плана нет.

### Расчёт даты

```sql
SELECT DISTINCT plan_id, day_number
FROM kp_plates
WHERE kp_id = ?
  AND status = 'в плане'
  AND plan_id IS NOT NULL
  AND day_number IS NOT NULL
```

Для каждого `plan_id` — `get_plan_day_to_date_mapping(plan_id)` из `app/planning/plan_manager.py`:

```python
date_by_day = get_plan_day_to_date_mapping(plan_id)
calendar_date = date_by_day.get(day_number)  # ISO YYYY-MM-DD
```

`expected_sgp_date_iso = max(resolved_dates)` — лексикографический max на ISO достаточен.

**Кеш:** в рамках одного `build_summary` — dict `plan_id → mapping`, чтобы не грузить план многократно.

### API fields (расширение `KpReadinessSummary`)

```python
expected_sgp_date: str | None = Field(
    default=None,
    description="ISO date YYYY-MM-DD; last planned production day",
)
expected_sgp_date_label: str | None = Field(
    default=None,
    description="Formatted DD.MM.YYYY for UI",
)
fully_scheduled: bool = False
```

### Тексты

**UI строка** (если `expected_sgp_date_label`):

> Ожидаем на СГП к: **15.08.2026**

**Дополнение к `client_copy_text`** (append через пробел или новое предложение):

> …Можно забрать 307 шт. **Ожидаем полный комплект на складе к 15.08.2026.**

Шаблоны по вариантам `summary` — в `KpReadinessService._append_expected_date_to_copy()`.

`summary_text` **не менять** (внутренний тон); только отдельная строка + copy.

---

## Tech Stack

Без изменений vs parent spec. Новая зависимость: `app.planning.plan_manager.get_plan_day_to_date_mapping`.

---

## Commands

```bash
pytest tests/test_kp_readiness_service.py -k expected_sgp -q
pytest tests/test_archive_endpoints.py -q
cd frontend && npm test -- --run KpReadiness && npm run build
```

---

## Project Structure

```
app/services/kp_readiness_service.py   → _resolve_expected_sgp_date(); build_summary()
app/schemas/archive.py                 → +3 поля в KpReadinessSummary
frontend/.../KpReadinessBlock.tsx      → строка expected_sgp_date_label
frontend/.../types/archive.ts          → типы
tests/test_kp_readiness_service.py     → fixtures: plan + kp_plates in_plan
```

---

## Code Style

```python
def _resolve_expected_sgp_date(
    self,
    cur: sqlite3.Cursor,
    kp_id: int,
    *,
    remaining_total: int,
    progress: SgpProgress,
    in_plan_total: int,
) -> tuple[str | None, str | None]:
    if remaining_total > 0 or progress.n >= progress.m or in_plan_total <= 0:
        return None, None
    # DISTINCT plan_id, day_number → max ISO date → label DD.MM.YYYY
```

- Не падать, если план удалён — пропустить пару `(plan_id, day)`.
- Формат label: `datetime.strptime(iso, "%Y-%m-%d").strftime("%d.%m.%Y")`.

---

## Testing Strategy

| Test | Fixture |
|------|---------|
| `remaining > 0` → null | 5 в производстве, 0 в плане |
| `n == m` → null | всё на СГП |
| Два дня в одном плане → max | day 1 = 2026-08-10, day 3 = 2026-08-14 → 14.08 |
| Два plan_id → max across plans | |
| Нет day_number → null | |
| Copy содержит дату | snapshot substring |
| Vitest: строка в UI | mock readiness с label |

Fixtures: in-memory DB + mock plan JSON в `plans` table (как `test_production_planning`).

---

## Boundaries

### Always

- Read-only SQL + read plan payloads
- Graceful degrade → `null` без 500
- Дата только при `fully_scheduled`

### Ask first

- Буфер +1 рабочий день
- ETA по позициям в таблице «Подробнее»
- Использовать `execution_terms` вместо плана

### Never

- Менять даты в плане / перепланировать
- Показывать дату при `remaining > 0`
- Обещать точное время (только дата)

---

## Success Criteria

1. КП 307/553, 0 remaining, последний день 14.08 → строка «Ожидаем на СГП к: 14.08.2026».
2. Copy клиента содержит «…к 14.08.2026».
3. КП с 246 «в производстве» → строки нет.
4. КП 553/553 на СГП → строки нет.
5. Tests green; build OK.

---

## Out of Scope

- Hint на степпере СГП (выбрана только summary line)
- Рабочий календарь / пропуск выходных для +1 дня
- Дата выдачи клиенту (модуль B)

---

## Open Questions

_Нет блокирующих._

---

## Next (SDD)

1. ~~Human review spec~~ ✅
2. ~~Plan: Phase 6~~ ✅
3. ~~Implement RDY-600…604~~ ✅
