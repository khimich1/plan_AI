# Handoff: Trim — соседние марки не воруют резы (шаг 1 ✅ → live / шаг 2)

> **Дата:** 2026-09-03  
> **Ветка:** текущая рабочая · **шаг 1 в working tree, не закоммичен**  
> **Статус:** Step 1 IMPLEMENT ✅ (2026-09-03) · **Step 2 NOT started**  
> **Цель файла:** открыть **новое окно в мультитаске** и продолжить (live-проверка разбивки / условный шаг 2), **без** повторного ideation.  
> **Не коммитить** без явной просьбы пользователя.  
> **Не убивать** `./run+logs.sh` (уже запущен).

---

## Как стартовать новую сессию (скопируй в первый промпт)

```
В мультитаске: продолжай trim-neighbor-sku-cuts. Не коммить. Не /idea-refine. Не трогай ILP и геометрию 20 мм (керф).

Контекст: прочитай целиком
ai_docs/develop/handoffs/2026-09-03-trim-neighbor-sku-cuts.md

Skills:
- .cursor/skills/plan-web-context/SKILL.md
- .cursor/skills/project-shishov/SKILL.md

Источник правды:
- ai_docs/specs/trim-neighbor-sku-cuts.md
- ai_docs/develop/plans/2026-09-03-trim-neighbor-sku-cuts.md

Шаг 1 уже в working tree (PLATE_WIDTH_MATCH_TOLERANCE_MM = 10, _width_matches_cut, тесты). Не переimplementируй.

Дальше (жди пользователя, если неясно):
1) опционально: pytest tests/test_procurement_trim_cuts.py tests/test_procurement_mixed_load_breakdown.py -q
2) live: пересчитать черновик КП/разбивку (df9b2510 / архив КП 3) — ПБ 56-7-8п: 1 поперечный, остаток 0,40 м, БЕЗ продольного 6,3 м; ПБ 56-7,2 свой 6,3 м; цена 56-7 ≈ 12 916 ₽
3) шаг 2 (target_order_key / consume-once) — ТОЛЬКО если live всё ещё ворует рез с Δ>10 мм или один secondary_instance_id биллится дважды

TDD, если будет код. Не коммить. Не убивать ./run+logs.sh.
```

### Чеклист агента в новом окне

1. Прочитать **этот** handoff целиком.  
2. `.cursor/skills/plan-web-context/SKILL.md` и `.cursor/skills/project-shishov/SKILL.md`.  
3. Спека + план (пути выше).  
4. **Не** запускать `/idea-refine` заново — решения locked.  
5. **Не** переimplementировать шаг 1.  
6. Опционально прогнать pytest (команда ниже).  
7. Live-проверка разбивки **только если пользователь хочет**; иначе ждать.  
8. Шаг 2 — **только** если live всё ещё крадёт рез с Δ>10 мм или один `secondary_instance_id` биллится дважды.  
9. Если будет код — TDD. Не трогать `MIN_BILLABLE_TRIM_MM`, `_longitudinal_cuts_for_rest_secondary`, `tolerance_width=20`, ILP.  
10. Не коммитить без просьбы. Не убивать `./run+logs.sh`.

**Режим:** multitask · **не** ideation · **не** полный orchestration с нуля · шаг 1 уже сделан.

---

## Артефакты

| Артефакт | Путь | Статус |
|----------|------|--------|
| **Idea** | [`ai_docs/ideas/trim-neighbor-sku-cuts.md`](../../ideas/trim-neighbor-sku-cuts.md) | locked |
| **Spec** | [`ai_docs/specs/trim-neighbor-sku-cuts.md`](../../specs/trim-neighbor-sku-cuts.md) | SPECIFY ✅ · IMPLEMENT ✅ шаг 1 |
| **Plan** | [`ai_docs/develop/plans/2026-09-03-trim-neighbor-sku-cuts.md`](../plans/2026-09-03-trim-neighbor-sku-cuts.md) | TRIM-001…004 ✅ · шаг 2 не открыт |
| Related (R1–R6, 20 мм как отход) | [`ai_docs/specs/plate-pricing-trim-bugs.md`](../../specs/plate-pricing-trim-bugs.md) | контекст, не этот срез |

---

## Что уже сделано (не переделывать) — Step 1 IMPLEMENT ✅ 2026-09-03

Баг: разбивка КП для **ПБ 56-7-8п** начисляла **2 поперечных** + продольный `460 × 6,3 × 1`, потому что `_width_matches_cut` сравнивал с допуском **20 мм** (керф / кромка). `|700−720|=20` склеивало **ПБ 56-7** и **ПБ 56-7,2**. Эталон: **ПБ 65-9-8п** (1 поперечный, без продольного на ребёнке).

После фикса 56-7 должна быть **~12 915,62 ₽**, не **18 380,45**. Сохранённое КП **само не пересчитывается** — нужен пересчёт черновика.

### Код в working tree (uncommitted)

- `core/config/constants.py` — `PLATE_WIDTH_MATCH_TOLERANCE_MM = 10`
- `core/config/__init__.py` — реэкспорт
- `viz_modules/procurement/trim.py` — `_width_matches_cut` использует 10 мм
- `tests/test_procurement_trim_cuts.py` — регрессия ниже

КП и производственная смета идут через один `trim.py`.

### Тесты (добавлены)

- `test_pb_56_7_does_not_steal_720_secondary`
- `test_pb_56_72_keeps_own_63m_cut`
- `test_pb_65_9_transverse_no_long_cut`
- `test_neighbor_delta20_does_not_share_cuts` (700/720, 300/320, 480/500, 880/900)
- `test_width_matches_cut_mark_tolerance_10mm`
- `test_waste_20_two_pieces_one_longitudinal_cut`

### Прогон, который уже зелёный

```bash
pytest tests/test_procurement_trim_cuts.py tests/test_procurement_mixed_load_breakdown.py -q
# → 50 passed
```

---

## Что делать в новом окне (по умолчанию)

1. **Не** переimplementировать шаг 1.  
2. Опционально: снова прогнать pytest выше.  
3. Live: перегенерировать КП/разбивку для текущего черновика **`df9b2510` / архив КП 3** и проверить:
   - **ПБ 56-7-8п** — 1 поперечный, остаток **только 0,40 м**, **нет** продольного 6,3 м;
   - **ПБ 56-7,2** — свой продольный 6,3 м;
   - цена 56-7 ≈ **12 916 ₽** (точно ~12 915,62).  
4. **Шаг 2** начинать **только если** live всё ещё ворует рез с **Δ>10 мм** или один `secondary_instance_id` биллится дважды.  
5. Если пользователь просит commit — git-helper / правила коммита пользователя; **не** коммитить секреты.

Шаг 2 (если понадобится): `target_order_key` / consume-once; матч secondary к строке сметы; каждый `secondary_instance_id` — не больше одной строки. Не возвращать допуск 20 мм в `_width_matches_cut`.

---

## Locked decisions (не переспрашивать)

| ID | Решение |
|----|---------|
| D-mark-10 | Допуск **марки** = **10 мм** (1 см на плиту, документация завода) → `PLATE_WIDTH_MATCH_TOLERANCE_MM = 10` |
| D-kerf-20 | Керф / второй продольный **не делали**: `MIN_BILLABLE_TRIM_MM = 20` **без изменений**; `_longitudinal_cuts_for_rest_secondary` **без изменений**; геометрия `tolerance_width=20` **без изменений** |
| D-700-720 | 700 vs 720 (Δ=20) — **не** делят резы |
| D-710-720 | 710 vs 720 (Δ=10) — **могут** матчиться; **не** триггер шага 2 |
| D-step2 | Consume-once / `target_order_key` — **только шаг 2**, не этот срез |
| D-one-trim | КП и производственная смета — один `trim.py` |
| D-no-commit | Не коммитить, пока пользователь не попросит |
| D-runlogs | Не убивать `./run+logs.sh` |
| D-no-ideate | Не перезапускать `/idea-refine` |

---

## Ключевые файлы кода (уже в WT)

```
core/config/constants.py              ← PLATE_WIDTH_MATCH_TOLERANCE_MM = 10
core/config/__init__.py               ← реэкспорт
viz_modules/procurement/trim.py       ← _width_matches_cut (10 мм)
tests/test_procurement_trim_cuts.py   ← регрессия шага 1
```

Шаг 2 (ещё не трогали): `viz_modules/procurement/plan_snapshot.py` — ключ заказа / `secondary_instance_id`; матч в `trim.py`. `price_rows.py` / `breakdown.py` не менять, если trim достаточен.

---

## Команды проверки

```bash
pytest tests/test_procurement_trim_cuts.py tests/test_procurement_mixed_load_breakdown.py -q
# Dev уже поднят: ./run+logs.sh — не убивать
```

Live smoke (если пользователь хочет):

1. Пересчитать **черновик** (сохранённое КП само не пересчитается). Draft `df9b2510` / архив КП 3.  
2. ПБ 56-7-8п: 1 поперечный, remainder 0,40 м, **без** `460 × 6,3`.  
3. ПБ 56-7,2: свой 6,3 м.  
4. Сумма 56-7 ≈ 12 915,62 ₽ (не 18 380,45).  
5. ПБ 65-9-8п без регрессии (1 поперечный, без продольного на ребёнке).

---

## Out of scope этого handoff

- Повторный implement шага 1  
- `/idea-refine`  
- ILP / `demand_tolerance_width` / геометрия `tolerance_width=20`  
- Менять `MIN_BILLABLE_TRIM_MM` или `_longitudinal_cuts_for_rest_secondary`  
- Возвращать `<= 20` в `_width_matches_cut`  
- Consume-once / `target_order_key`, пока live не покажет кражу Δ>10 мм или двойной биллинг  
- UI формулы остатка («0,40м» при сумме 0,40+0,70)  
- Пересчёт архива, `core/pricing/`  
- Factory strip 1020–1080 (R5)  
- Commit / PR без просьбы  
- Убивать `./run+logs.sh`

---

## Definition of done (новое окно)

- [ ] Подтвердить, что шаг 1 всё ещё зелёный (`pytest tests/test_procurement_trim_cuts.py tests/test_procurement_mixed_load_breakdown.py -q`)  
- [ ] Если пользователь хочет live: пересчёт КП, 56-7 ≈ 12 916 ₽, нет украденного 6,3 м  
- [ ] Шаг 2 — только если кража с Δ>10 мм остаётся  
- [ ] Краткий отчёт пользователю; **без commit**, пока не попросят
