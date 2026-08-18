# Ёмкость дозаполнения: fill ≤ max, без тихого разгона

Дата: 2026-08-12. Статус: **done** (implemented).
Спека: [`../specs/emkost-dozapolneniya-fill-bez-razgona.md`](../specs/emkost-dozapolneniya-fill-bez-razgona.md).
План: [`../develop/plans/2026-08-12-emkost-dozapolneniya-fill-bez-razgona.md`](../develop/plans/2026-08-12-emkost-dozapolneniya-fill-bez-razgona.md).
Связь: [`planirovanie-po-srokam-podlozhki.md`](planirovanie-po-srokam-podlozhki.md).

## Problem Statement

Как планировщику, при дефиците дорожек под срочные плиты, расширять корзину
**в пределах нормальной ёмкости дня** (обычно 5) или **другим днём**, а не
получать тихий override на 6 и падение `build` с общим «операция не удалась»?

## Recommended Direction

Инцидент: алерт «не хватает дорожек» → кнопка «+N» писала
`day_capacity_override.max_tracks = 6` и поднимала `fill_targets` → analyze
принимал 6, `build`/`persist` считал потолок константой 5 →
«На 2026-08-20 свободно 5 дорожек, запрошено 6» → UI показывал только
`MSG_UNPROCESSABLE`.

Выбранная семантика: **на заводе ровно 5 дорожек — больше нельзя** (hard cap).
Поднимать fill только в пределах `5 − occupied` (или ниже, если день урезан);
иначе предложить **ближайший день** с free > 0 (вперёд/назад по горизонту).
Override >5 — запрещён (ни кнопка дефицита, ни «Ёмкость»).

1. **Кнопка дефицита** не вызывает `PUT /day-capacity`; только предлагает
   bump fill или другой день (подтверждение отдельно).
2. **Нет headroom** → искать дальше дни со свободными дорожками, не разгонять max.
3. **Единый free** в analyze, build, календаре: `free = day_max − occupied`, `day_max ≤ 5`.
4. **UI:** детальный текст ошибок build (scope API — в спеке).

## Key Assumptions to Validate

- [ ] A1. Дефолт 5 — «норма завода»; выше 5 — редкий ручной override в режиме «Ёмкость».
- [ ] A2. При дефиците почти всегда есть соседние дни с free (иначе UX упрётся в «некуда добавить»).
- [ ] A3. `fill_targets.tracks` = сколько дозаполнить (≤ free), не «новый max дня».
- [ ] A4. Детальный текст ошибок ок для ролей `admin` / `production`.

## MVP Scope

- Переписать suggestion в `calculate_capacity_deficit` (+ тесты): clamp к headroom;
  иначе ближайший день с free > 0.
- `addCapacityTracks`: только bump `fill_targets` / добавление дня; без
  `saveDayCapacity` из алерта.
- `persist` / `build`: валидация через capacity map + occupancy, не голый
  `MAX_TRACKS_PER_DAY = 5` без overrides.
- Проброс конкретного текста ошибки в 422 для production `plans/build`.
- Календарь: `days_info.max` из day-capacity (согласованность корзины и UI).

## Not Doing (and Why)

- Авто-override max > 5 из алерта дефицита — ломает модель «норма = 5».
- Полный рефактор всех `MAX_TRACKS_PER_DAY` в bot_archived — вне боли.
- Автовыбор дней оптимизатором без подтверждения менеджером — слишком много магии.
- Менять логику подложек / оптимизатора — проблема в ёмкости и UX, не в резах.

## Open Questions

- Если override уже задан вручную (например 7 в режиме «Ёмкость») — кнопка
  дефицита добивает fill до этого max или всегда ориентируется на default 5?
  (Черновик: уважать уже заданный override как потолок fill.)
- Критерий «ближайший день»: ближайший по календарю после `deficit_until`,
  только рабочие дни (work calendar), исключая уже полные — уточнить при спеке.
