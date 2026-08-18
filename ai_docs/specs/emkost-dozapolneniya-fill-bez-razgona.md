# Spec: Ёмкость дозаполнения — fill ≤ max, без тихого разгона

Дата: 2026-08-12. Статус: **approved** (решения уточнены 2026-08-12 вечер).
Идея: [`ai_docs/ideas/emkost-dozapolneniya-fill-bez-razgona.md`](../ideas/emkost-dozapolneniya-fill-bez-razgona.md).
Связанные: [`planirovanie-po-srokam-podlozhki.md`](./planirovanie-po-srokam-podlozhki.md),
фича [`../develop/features/planirovanie-po-srokam-podlozhki.md`](../develop/features/planirovanie-po-srokam-podlozhki.md).

## Decisions (locked)

| # | Решение |
|---|--------|
| Hard cap | На заводе **5 дорожек/день**. `max_tracks > 5` запрещён. |
| Кнопка дефицита | Не вызывает `PUT /day-capacity`. |
| UX дефицита | Система **предлагает варианты (список)**; **человек выбирает**, что принять. Без авто-merge в корзину. |
| Порядок поиска слотов | 1) добрать **в выбранных днях**, если там ещё `< 5` (есть free); 2) если в выбранных уже по 5 — **предыдущие дни ≥ today** со свободными дорожками; 3) если нет — **будущие дни**. |
| Закрытый день | День **не кандидат**, если переведён на СГП (`completed` после complete_day / списание на СГП). |
| Список | Показать **список** кандидатов; выбор за человеком (не одна «лучшая» кнопка без списка). |
| Ошибки API | Детальный `detail` для **`POST /plans/build`**. |
| `days_info.max` | Sync с day-capacity, всегда **≤5**. |
| Override >5 | Clamp при чтении; PUT >5 → **400**. |
| today | Календарная дата **Europe/Moscow**. |
| Строка списка | Показывать / предлагать **весь free** дня (`add_tracks = free`). |
| Длина списка | Max **10** options (порядок A→B→C, обрезать). |
| Ёмкость max | Допустимо **0…5** (0 = день вручную выключен, free=0, не в options). |

## ASSUMPTIONS

1. Hard cap = 5; floor = 0 (ручное выключение дня без СГП).
2. `free(d) = day_max(d) − occupied(d)`, `0 ≤ day_max ≤ 5`.
3. «В выбранных днях < 5» = у дня корзины есть headroom: `fill_tracks < free`.
4. Кандидаты вне корзины: `date ≥ today` (Moscow), не СГП, workday, `free > 0`, `day_max > 0`.
5. В option всегда `add_tracks = free` (весь свободный объём дня).
6. Empty/partial не смешиваются автоматически: человек выбирает из списка / уходит в календарь.
7. Кисть partial на календаре — ручной путь дозаполнения; список дефицита — помощь.

## Objective

При нехватке дорожек под срочные: показать **понятный список вариантов дозаполнения** по приоритету (свои дни → прошлые ≥ today → будущие), человек отмечает что взять; max дня никогда не разгоняется выше 5; build не падает из‑за silent override; ошибка build — с текстом.

## Tech Stack / Commands / Structure

Без изменений по стеку; см. предыдущую версию + FE: список опций вместо одной кнопки «+N».

```bash
.venv/bin/python -m pytest tests/test_production_capacity.py \
  tests/test_production_capacity_service.py \
  tests/test_production_planning_service_fill_targets.py \
  tests/test_core_production_planning.py \
  tests/test_production_api_integration.py \
  tests/test_day_capacity_repository.py -q
cd frontend && npm test -- --run src/features/production
```

## Behaviour: алгоритм вариантов

Вход: `fill_targets`, `tracks_missing` (или полный deficit), occupancy, completed, today, work calendar.

Построить **упорядоченный список** `options[]` (не один suggestion):

**Шаг A — выбранные дни**  
Для каждого `d` в `fill_targets`, если `headroom(d) = free(d) − fill_tracks(d) > 0`  
→ option `{ action: "bump_fill", date: d, add_tracks: headroom }`  
(`add_tracks` = весь доступный headroom = free − текущий fill; по смыслу «весь free» относительно уже занятого/заявленого).  
Порядок: по дате возрастанию.

**Шаг B — предыдущие дни** (если в выбранных нет headroom / не хватает)  
Кандидаты: `today ≤ date < min(fill_targets)`, не СГП, workday, `day_max > 0`, `free > 0`.  
→ `{ action: "propose_day", date, add_tracks: free, free }`  
Порядок: от более ранних к поздним.

**Шаг C — будущие дни**  
`date > max(fill_targets)`, не СГП, workday, `free > 0`, горизонт 30 календ. дней.  
→ те же propose_day, `add_tracks: free`.

Итоговый `options` = A+B+C, **max 10** элементов.

UI: список; человек выбирает; применение только по действию.  
Kind conflict → CTA «Открыть календарь на ДД.ММ».

## API delta

Было: одно `suggestion`.  
Стало (предпочтительно):

```json
"capacity_deficit": {
  "tracks_needed": 20,
  "tracks_available": 15,
  "tracks_missing": 5,
  "deficit_until": "2026-08-20",
  "options": [
    { "action": "bump_fill", "date": "2026-08-18", "add_tracks": 1, "free": 1 },
    { "action": "propose_day", "date": "2026-08-12", "add_tracks": 3, "free": 3 },
    { "action": "propose_day", "date": "2026-08-24", "add_tracks": 5, "free": 5 }
  ]
}
```

Обратная совместимость: поле `suggestion` можно оставить = `options[0]` на один релиз или убрать сразу (фронт наш — убираем, только `options`).

## Success Criteria

1. SC-1: нет PUT day-capacity из дефицита.  
2. SC-2: список options по порядку A→B→C; человек выбирает.  
3. SC-3: день после СГП не в options.  
4. SC-4: hard cap 5 в API/calendar/build.  
5. SC-5: build 422 с конкретным текстом.  
6. SC-6: календарь позволяет выбрать partial-дни и дозаполнить (регрессия brush).

## Boundaries

- **Always:** hard cap 5; max 0…5; список вариантов; человек решает; TDD; today=Moscow.  
- **Ask first:** смена N=10; смена empty/partial правил.  
- **Never:** max>5; auto PUT capacity; bot_archived; commit без просьбы.

## Open Questions

Нет — закрыты.

## Out of scope

Разгон >5; авто-выбор за человека; bot_archived; логика оптимизатора/подложек.

## Gate

SPECIFY ✓. Нужен **апрув плана** → IMPLEMENT.
