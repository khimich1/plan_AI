# Spec: Атомарный «перевести в производство» (audit Q1 + Q2)

> **Источник:** [`ai_docs/develop/audits/2026-08-02-full-project-audit.md`](../develop/audits/2026-08-02-full-project-audit.md) — findings **Q1** (проглоченный freeze), **Q2** (нет атомарности)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → IMPLEMENT ✅  
> **Статус:** IMPLEMENTED  
> **Дата:** 2026-08-02  
> **План:** [`../develop/plans/2026-08-02-move-to-production-atomicity.md`](../develop/plans/2026-08-02-move-to-production-atomicity.md)  
> **Scope:** атомарность `execution_terms + status + freeze(ordered_qty)` в обоих путях `move_to_production`; убрать `except Exception: pass`. Без дедупа Q6, без фронтенда, без схемы БД.

---

## Objective

Сделать перевод КП «в архиве» → «в работе» **атомарным и наблюдаемым**: либо применяются все три шага, либо ни один; сбой freeze M больше не маскируется.

**Пользователи:** менеджер / админ (архив КП, offers API).  
**Проблема сейчас:**
- `update_execution_date`, `update_status`, `freeze_ordered_qty_if_needed` — три отдельных соединения с автокоммитом.
- После статуса `except Exception: pass` — API отвечает успехом при `ordered_qty IS NULL`.
- Прогресс N/M, % выполнения, отгрузка x/m и автостатусы опираются на M; silent corruption ломает их незаметно.

**Результат:** одна транзакция; при ошибке — ROLLBACK, лог, ошибка клиенту; КП остаётся «в архиве».

### Acceptance (user-facing)

1. Успешный перевод: статус «в работе», срок сохранён, `kp_meta.ordered_qty` зафиксирован (не NULL).
2. Любой сбой на шаге записи → КП не меняет статус/срок/M; клиент получает ошибку (не 200 с полусостоянием).
3. В логах при сбое freeze — traceback (`logger.exception`), не тишина.

### Out of scope (Not Doing)

- **Q6** — вынос общего domain-сервиса для двух `move_to_production` (можно позже; контракт атомарности одинаковый в обоих).
- Frontend / тексты UI.
- Миграция схемы SQLite.
- Lazy-freeze в read-путях (`_shipped_progress`, SGP и т.п.) — не менять в этом тикете.
- Пометка Q1/Q2 как FIXED в audit-файле.

---

## Tech Stack

- Backend: Python 3, FastAPI, SQLite (`plita.db`)
- Слои: endpoint → service → repository / `core.kp.offers_write` + `freeze_ordered_qty_if_needed`
- Паттерн транзакции: `_external_conn` (как в `kp_db_rests` / `PlateCompletionService`)

---

## Commands

```bash
# Затронутые тесты (после реализации)
pytest tests/test_archive_service.py \
  tests/test_archive_endpoints.py \
  tests/test_move_to_production_atomicity.py -q

# Регрессия offers auth (если трогаем OffersService)
pytest tests/test_offers_production_authorization.py \
  tests/test_offers_authorization.py -q

# Поиск проглоченных except вокруг freeze (гейт)
rg -n 'except Exception:\s*\n\s*pass' app/services/archive_service.py app/services/offers_service.py
```

Success: pytest зелёный; в двух сервисах нет `except Exception: pass` вокруг freeze.

---

## Project Structure

```
core/kp/offers_write.py              → update_kp_status / update_kp_execution_date:
                                       опциональный _external_conn (без commit/close, если передан)
core/kp_db_plates_completion.py      → freeze_ordered_qty_if_needed (без изменения контракта;
                                       вызывающая сторона трактует None как ошибку операции)

app/repositories/kp_archive_repository.py  → при необходимости: метод атомарного перевода
app/repositories/kp_repository.py          → то же для offers-пути (или прямой вызов offers_write)

app/services/archive_service.py      → move_to_production: одна транзакция; без bare except
app/services/offers_service.py       → move_to_production: та же семантика

app/api/v1/endpoints/archive.py      → при необходимости: маппинг ArchiveError → HTTP (не silent 200)
app/api/v1/endpoints/offers.py       → маппинг кода ошибки freeze/tx → 4xx/5xx (не 200)

tests/test_move_to_production_atomicity.py  → NEW: интеграция на tmp SQLite
tests/test_archive_service.py               → обновить моки под новый контракт репозитория (если API меняется)
```

---

## Code Style

- Следовать существующему `_external_conn`:
  - `own_conn = _external_conn is None`
  - при `own_conn`: `_connect` → commit/rollback/close
  - при внешнем conn: **не** commit/rollback/close
- Ошибки домена: `ArchiveError` / `ArchiveValidationError` в архиве; `ValueError("…")` коды в offers — согласовать с существующим endpoint-маппингом.
- Логирование: `logger.exception(...)` перед raise; без `except Exception: pass`.
- `freeze` вернул `None` (нет строки `kp_meta`) → считать провалом операции (rollback + ошибка), не успехом.

Пример целевого каркаса (иллюстрация, не финальный API):

```python
conn = _connect(db_path)
try:
    conn.execute("PRAGMA foreign_keys = ON")
    cur = conn.cursor()
    if not update_kp_execution_date(kp_id, terms, db_path, _external_conn=conn):
        raise ArchiveError(...)
    if not update_kp_status(kp_id, "в работе", db_path, _external_conn=conn):
        raise ArchiveError(...)
    ordered = freeze_ordered_qty_if_needed(cur, kp_id)
    if ordered is None:
        raise ArchiveError(f"Не удалось зафиксировать ordered_qty для КП №{kp_id}")
    conn.commit()
except Exception:
    conn.rollback()
    logger.exception("move_to_production failed for kp_id=%s", kp_id)
    raise
finally:
    conn.close()
```

---

## Testing Strategy

| Уровень | Что | Где |
|---------|-----|-----|
| Integration (tmp SQLite) | Happy path: статус + срок + `ordered_qty` после одного вызова | `tests/test_move_to_production_atomicity.py` |
| Integration | Сбой freeze (mock/`monkeypatch` raise или отсутствие `kp_meta`) → статус остаётся «в архиве», срок не меняется, `ordered_qty` NULL | тот же файл |
| Integration | Сбой mid-flight (mock update_status fail после execution_date на одном conn) → полный rollback | тот же файл |
| Unit (существующие) | Валидация «только из архива», нормализация сроков | `tests/test_archive_service.py` |
| Endpoint | Ошибка freeze/tx не даёт 200 | archive/offers endpoint tests при необходимости |
| Gate | Нет `except Exception: pass` в двух сервисах вокруг freeze | `rg` |

Покрытие: оба сервиса (`ArchiveService`, `OffersService`) — минимум один failure-path тест на каждый или общий helper + параметризация.

---

## Boundaries

- **Always:** атомарность трёх шагов; логировать сбои; pytest по затронутым тестам; не глотать Exception.
- **Ask first:** дедуп Q6 в общий domain service; смена HTTP-кодов/контракта ответа API; правка audit FIXED; расширение `_external_conn` на другие `offers_write.*`.
- **Never:** silent success при незафиксированном M; коммит секретов; «компенсирующий» UPDATE статуса отдельной транзакцией вместо единого ROLLBACK; трогать frontend без запроса.

---

## Success Criteria

1. В `archive_service.py` и `offers_service.py` нет `except Exception: pass` вокруг freeze.
2. `move_to_production` (оба сервиса) выполняет срок + статус + freeze в **одной** SQLite-транзакции; при ошибке любого шага — ROLLBACK всех трёх эффектов.
3. При успешном переводе `kp_meta.ordered_qty IS NOT NULL`.
4. При искусственном сбое freeze интеграционный тест доказывает: статус = «в архиве», срок прежний, `ordered_qty` NULL, исключение проброшено.
5. Сбой логируется через `logger.exception` (или эквивалент).
6. Существующие happy-path / validation тесты архива зелёные; новый atomicity-набор зелёный.
7. Audit-файл **не** помечен FIXED в этом изменении.

---

## Resolved decisions

| # | Вопрос | Решение |
|---|--------|---------|
| 1 | Поведение при ошибке freeze | **A:** ROLLBACK всей операции + доменная ошибка клиенту + `logger.exception` |
| 2 | Дедуп двух сервисов (Q6) | **Out of scope** — одинаковый контракт атомарности в обоих, общий helper позже |
| 3 | `freeze` → `None` | Считать ошибкой операции (rollback), не успехом |
| 4 | Имя спеки | `ai_docs/specs/move-to-production-atomicity-q1-q2.md` |

---

## Open Questions

Нет блокирующих. При review можно уточнить только HTTP-код для archive (`ArchiveError` сейчас часто уходит в 500) — допустимо оставить 500 для tx/freeze failure; offers уже мапит неизвестный `ValueError` в 400.

---

## Next phase

PLAN готов: [`../develop/plans/2026-08-02-move-to-production-atomicity.md`](../develop/plans/2026-08-02-move-to-production-atomicity.md) (задачи MTP-100…500).  
После approve плана → IMPLEMENT по TDD / `incremental-implementation`.
