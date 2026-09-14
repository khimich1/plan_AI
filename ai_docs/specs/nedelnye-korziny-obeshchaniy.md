# Spec: Недельные корзины обещаний (подбор срока КП)

> **Источник:** [`ai_docs/ideas/nedelnye-korziny-obeshchaniy.md`](../ideas/nedelnye-korziny-obeshchaniy.md) — decisions locked 2026-09-03
> **Дата:** 2026-09-03
> **Статус:** draft. Task 0 выполнен, buffer = 1.0 зафиксирован. Phase 1 в работе.
> **Связанные:** [`zavod-emkost-vizual-gate.md`](zavod-emkost-vizual-gate.md) (подневной гейт — остаётся в графике поставок; в диалоге «В производство» заменяется корзинами),
> [`planirovanie-po-srokam-podlozhki.md`](planirovanie-po-srokam-podlozhki.md) (подложки добивают остаток недели),
> [`move-to-production-atomicity-q1-q2.md`](move-to-production-atomicity-q1-q2.md) (транзакция коммита, в которую добавляется запись обещания),
> `core/delivery_schedule_check.py`, `core/production_capacity.py`, `core/work_calendar.py`

---

## Assumptions (locked)

```
ASSUMPTIONS:
1. Неделя = ISO пн–вс; week_start = понедельник. Ёмкость корзины =
   (рабочие дни недели по work_calendar: holidays + extra_workdays) × ручка
   promise_tracks_per_day (default 3, max TRACKS_PER_DAY_HARD_CAP=5).
   Текущая частичная неделя: рабочие дни от ЗАВТРА до конца недели.
2. tracks(KP) = ceil( Σ(length_m × qty) / MAX_TRACK_LENGTH_M × buffer ).
   buffer = 1.0 (безразмерный коэффициент к м/101, 0% надбавки) — подтверждено
   калибровкой Task 0 (отчёт ai_docs/develop/reports/2026-09-03-promise-buffers.md).
   Параметр в kp_setting, не зашитая константа. Соло-дни = ceil(tracks / ручка)
   рабочих дней от завтра.
3. Занятость корзины: free = max(0, capacity − planned − promised).
   planned = Σ days_info[день].occupied по дням недели (план, реальные мощности
   до 5/день). promised = Σ активных promise-аллокаций недели.
   ХОЛДЫ НЕ ВЫЧИТАЮТСЯ — показываются отдельным счётчиком.
4. Размещение КП: если tracks ≤ free какой-то одной недели → ЦЕЛИКОМ в первую
   такую неделю (разреза нет, фрагментация не проблема: вдали free=полная).
   Если tracks > ёмкости целой недели → ОКНО: жадно от первой недели с free>0,
   подряд, allocation = min(free, остаток) на неделю; promised_date =
   последний рабочий день последней недели окна. Только такие КП режутся
   по неделям («целиком или никак» для остальных).
5. Гейт: move-to-production с датой X разрешён только если X ≥ promised_date
   из ПЕРЕСЧЁТА корзин на момент перевода (не из устаревшей котировки).
   Иначе 4xx + ближайшая возможная дата в сообщении. Оба пути:
   POST /commercial/archive/{id}/move-to-production И
   PATCH /offers/{id}/move-to-production (сейчас без гейта — закрываем дыру).
6. Обещание неделимо: одно КП, одна promised_date, статус consumed только
   когда у КП не осталось незапланированных позиций. Коммит плана недели W
   гасит аллокации недели W вошедших КП; аллокация НЕ вошедшего КП → overdue
   (красный сигнал + уведомление), а не молчаливый сдвиг.
7. Холд: TTL до конца текущего дня (локальное время завода), ленивый expire
   при чтении. Конвертация холда в promise — в той же транзакции, что
   commit_move_to_production (срок из холда, повторный ввод не нужен; гейт
   п.5 всё равно выполняется).
8. Ручка promise_tracks_per_day: roles admin+manager, каждое изменение с
   updated_by/updated_at; применяется ТОЛЬКО к новым расчётам — активные
   обещания/холды задним числом не пересчитываются.
9. Статус КП в БД не меняется: «на рассмотрении» = производная от активного
   холда (бейдж в разделе «в архиве»), новых статусов и секций архива нет.
10. Миграция: чистый старт — существующие КП «в работе» НЕ регистрируются
    как обещания; они продолжают попадать в планы через блок срочных.
    Принимаем: первые недели корзины занижают реальную загрузку.
11. Роли: котировка/холд/перевод — admin, manager (как сейчас archive).
    Блок «Обещано» в wizard — admin, production (как build_plan).
12. Уведомления — только in-web (таблица + бейдж). Telegram-бот архивирован,
    не оживляем.
13. Подневной capacity-snapshot и check_batches НЕ меняются — остаются для
    графика поставок. Корзины — отдельный чистый модуль, не форк check_batches.
→ Correct me now if wrong.
```

---

## Objective

Менеджер на ПК, формируя КП и переводя его в производство, видит **честную
дату из недельных корзин** (план + уже обещанное), может закрепить её холдом
на время согласования с клиентом и перевести в производство одной кнопкой —
с гарантией, что планировщик увидит обещание при сборке недели и не снимет
его молча.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | менеджер открыл «В производство» | видеть первично «обещать к 25.09», вторично начало/соло | не угадывать «N дней» вслепую |
| US-2 | менеджер согласовывает с клиентом | «Закрепить срок» — холд до конца дня, бейдж в архиве | место не увели, пока я говорю с клиентом |
| US-3 | менеджер после согласования | «В производство» из холда без повторного ввода срока | одна кнопка, без трения |
| US-4 | менеджер ввёл слишком раннюю дату | 4xx с ближайшей возможной датой | не пообещать нереальное |
| US-5 | менеджер формирует КП | видеть «~N дорожек» на финальном шаге мастера | понимать масштаб заказа до архива |
| US-6 | планировщик собирает неделю | блок «Обещано на эту неделю» предвыбран | не потерять обещанное при смешивании |
| US-7 | планировщик снимает обещанное КП | обязательная причина + уведомление менеджеру | нет молчаливых срывов |
| US-8 | планировщик видит overdue | красный блок «обещано, но не в плане» | эскалация видна сразу |
| US-9 | менеджер с крупным КП (20 дор.) | окно 2 недели, обещание к концу второй | модель не ломается на больших заказах |

### Acceptance criteria (MVP)

- [ ] `GET promise-quote`: tracks, solo_days/solo_date, solo_week_end_date,
      earliest_start_week, window {from_week, to_week, promised_date},
      weeks[{week_start, workdays, capacity, planned, promised, held, free}], knob
- [ ] Корректность корзин: частичная текущая неделя, праздники/extra_workdays,
      крупное КП (окно), whole-only при фрагментации, free=max(0,…)
- [ ] Накопление: котировка второго КП учитывает promise первого
- [ ] Холд: создание по котировке, TTL конец дня, ленивый expire, бейдж
      «срок закреплён до сегодня» в архиве, счётчик «из них холды N» в чужих
      котировках, конвертация одной кнопкой
- [ ] Гейт: дата раньше promised_date → 4xx + earliest; оба пути перевода
- [ ] Атомарность: promise + execution_terms + status + freeze(ordered_qty)
      в одной SQLite-транзакции (расширение commit_move_to_production)
- [ ] Погашение: коммит плана недели W → consumed для вошедших; невошедшее →
      overdue, не исчезает молча
- [ ] Wizard: блок «Обещано на эту неделю» с предвыбором; снятие требует
      причины; причина пишется в журнал; менеджеру — in-web уведомление
- [ ] Удаление КП → release активных promise/hold; редактирование состава →
      пересчёт tracks и окна, при сдвиге promised_date — уведомление
- [ ] Ручка: GET/PUT, default 3, cap 5, аудит; старые расчёты не трогаем
- [ ] Мастер КП, финальный шаг: «~N дорожек»
- [x] Task 0: калибровочный скрипт отработал, отчёт в ai_docs/develop/reports,
      buffer = 1.0 зафиксирован в спеке (п.2 Assumptions)

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`) |
| Domain | новый чистый модуль `core/production/promise_buckets.py` (no I/O); reuse `core/work_calendar.py`, `core/production_capacity.py` |
| Frontend | React 19, TS, Vite, TanStack Query; feature-структура |
| Тесты | pytest (`tests/`), vitest (`frontend/src/`) |

Новых внешних зависимостей нет.

---

## Commands

```bash
source venv/bin/activate

# Task 0 — калибровка ДО UI (прецедент: validate_podlozhki_phase0.py)
python scripts/validate_promise_buffers.py --db plita.db \
  --report ai_docs/develop/reports/2026-09-XX-promise-buffers.md

# Backend
pytest tests/test_promise_buckets.py tests/test_promise_service.py -q
pytest tests/test_archive_endpoints.py tests/test_move_to_production_atomicity.py -q
uvicorn app.main:app --reload

# Frontend
cd frontend && npm run test -- --run
cd frontend && npm run typecheck
cd frontend && npm run build

./run+logs.sh
```

---

## Project Structure

```
core/
  production/promise_buckets.py        → NEW: чистая математика корзин
    (WeekBucket, allocate(tracks, weeks) → PromiseWindow; no I/O, no app.*)
  work_calendar.py                     → reuse: рабочие дни недели
  production_capacity.py               → reuse: MAX_TRACK_LENGTH_M, cap 5
  kp_db_schema.py                      → +таблицы kp_promise, kp_promise_alloc,
                                         kp_promise_exclusion, notifications,
                                         настройка promise_tracks_per_day
  kp/offers_write.py                   → commit_move_to_production + запись
                                         promise/конвертация холда (одна tx)
  plan_commit.py                       → погашение аллокаций + overdue

app/
  repositories/promise_repository.py   → NEW: CRUD журнала, ленивый expire холдов
  services/promise_service.py          → NEW: котировка, холды, гейт, погашение
  services/archive_service.py          → move_to_production через promise gate
  services/offers_service.py           → тот же гейт (второй путь)
  api/v1/endpoints/archive.py          → +GET promise-quote, +POST/DELETE promise-hold
  api/v1/endpoints/offers.py           → PATCH move-to-production → gate
  api/v1/endpoints/production.py       → promised block для wizard
  api/v1/endpoints/notifications.py    → NEW: GET list + POST read
  api/v1/endpoints/settings…           → GET/PUT promise_tracks_per_day (admin+manager)
  schemas/                             → +PromiseQuote, PromiseHold, …

frontend/src/
  features/factory-capacity/
    components/PromiseWeekStrip.tsx    → NEW: полоса недель (план/обещано/холды/свободно)
    components/PromiseQuoteBlock.tsx   → NEW: 4 числа котировки
    … FactoryMiniCalendar остаётся в графике поставок, не трогаем
  features/commercial-archive/
    components/MoveToProductionDialog.tsx → котировка, «Закрепить срок», week strip,
                                          кнопка настройки ручки (в drawer «Ёмкость»)
    components/OfferDetailsDrawer.tsx  → бейдж холда
  features/production/components/create-plan-wizard/
    PromisedWeekBlock.tsx              → NEW: предвыбор + причина при снятии + overdue
  features/notifications/              → NEW: бейдж в шапке + список
  features/commercial-offer/           → «~N дорожек» на финальном шаге мастера

scripts/
  validate_promise_buffers.py          → NEW: Task 0 калибровка A1
```

---

## Code Style

Чистая математика корзин — без I/O и без импортов `app.*` (как `core/production/capacity.py`):

```python
# core/production/promise_buckets.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass(frozen=True, slots=True)
class WeekBucket:
    week_start: date          # понедельник
    workdays: int             # рабочие дни (для текущей — от завтра)
    capacity: int             # workdays * promise_tracks_per_day
    planned: int              # Σ occupied из плана
    promised: int             # активные обещания
    held: int                 # активные холды (НЕ вычитаются)

    @property
    def free(self) -> int:
        return max(0, self.capacity - self.planned - self.promised)


@dataclass(frozen=True, slots=True)
class PromiseWindow:
    from_week: date
    to_week: date
    promised_date: date                       # последний рабочий день to_week
    allocations: tuple[tuple[date, int], ...]  # (week_start, tracks)


def allocate(tracks: int, weeks: list[WeekBucket]) -> PromiseWindow | None:
    """Целиком в первую неделю с free >= tracks; если tracks больше целой
    корзины — жадное окно от первой недели с free > 0."""
```

Слои: роутер → сервис → репозиторий. Все мутации обещаний — в транзакции
вызывающего (коммит перевода / коммит плана), без `except: pass`.

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| pytest pure | `test_promise_buckets.py`: частичная неделя, праздники, крупное КП-окно, whole-only при фрагментации, free=max(0,…), promised_date = последний рабочий день |
| pytest service | `test_promise_service.py`: холд TTL/expire, гейт 4xx + earliest, конвертация холда, погашение при коммите, overdue, release при удалении КП, пересчёт при редактировании |
| pytest api | quote/hold endpoints, оба пути move-to-production через гейт, ручка GET/PUT + аудит |
| vitest | диалог (4 числа, «Закрепить срок», disabled submit при ранней дате), PromiseWeekStrip, PromisedWeekBlock (предвыбор, причина, overdue), бейдж холда, бейдж уведомлений |
| Калибровка | `validate_promise_buffers.py` → отчёт; запуск вручную, результат фиксирует buffer в спеке |
| Регресс | `test_move_to_production_atomicity.py`, `test_capacity_gate.py` (график поставок не тронут) |

---

## Boundaries

**Always:**
- Запись promise/холда — в той же SQLite-транзакции, что перевод/коммит плана
- Холды не вычитаются из free — только отдельный счётчик
- Ручка ≤ TRACKS_PER_DAY_HARD_CAP (5); изменения с updated_by
- Один гейт на обоих путях move-to-production
- pytest + vitest зелёные перед отчётом

**Ask first:**
- Новые статусы КП в БД (сейчас «на рассмотрении» — производная, не статус)
- Изменения `check_batches` / подневного capacity-snapshot (график поставок)
- Telegram-уведомления, push, email
- Любые изменения оптимизатора / ILP

**Never:**
- Фантомная занятость в `days_info` (обещания — отдельный журнал)
- Ночной сброс обещаний (сгорают только холды)
- Жёсткий блок коммита плана (уровень 2: причина + уведомление)
- Лимиты холдов на менеджера (только при факте злоупотреблений — A3)
- Ослабление существующего гейта графика поставок

---

## Success Criteria

1. Менеджер называет клиенту дату из котировки, не угадывая «N дней».
2. Котировка следующего КП учитывает обещания предыдущих (накопление).
3. Холд сгорает в конце дня; конвертация в обещание — одной кнопкой.
4. Нет двойного счёта: после коммита плана недели W её аллокации consumed.
5. Исключение обещанного КП из плана всегда с причиной; менеджер уведомлён in-web.
6. Калибровка A1 выполнена, buffer зафиксирован; оба пути перевода под одним гейтом.

---

## Not Doing

- ILP с дедлайнами / изменения оптимизатора
- Автоплан из диалога, бронь дорожек в `days_info`
- СГП в формуле корзин
- Мобильная вёрстка
- Жёсткий блок коммита плана; структурный уровень (фикс-костяк) — по метрике A2
- Лимиты холдов — по факту A3
- Telegram/push/email-уведомления

---

## Open Questions

- Политика пересчёта окна при редактировании состава КП с активным promise:
  MVP — пересчёт + уведомление при сдвиге promised_date; детали в плане.

## Resolved при написании спеки (проверка кода 2026-09-03)

- **buffer = 1.0** — калибровка Task 0
  (`ai_docs/develop/reports/2026-09-03-promise-buffers.md`): оценка без надбавки
  покрывает факт (±1 дорожка); 1.15 систематически завышает. Параметр в
  `kp_setting`, дефолт при первом запуске = 1.0.
- **Хранение ручки** `promise_tracks_per_day`: в `plita.db` через
  `core/kp_db_schema.py` — там уже есть key-value паттерн (`gsm_setting`,
  schema:612) и аудит-паттерн (`day_capacity_override` с updated_by/updated_at,
  schema:641). Новая таблица `kp_setting` по образцу `gsm_setting`.
- **Notifications:** пользователи — `app_users` (`app/repositories/auth_repository.py:40`);
  таблица `notifications` рядом, `user_id` мягкой ссылкой (без FK через БД).
- **Точка погашения:** `commit_plan_plates` (`core/plan_commit.py:308`) —
  расширяем, новый параллельный коммит-путь не строим.
