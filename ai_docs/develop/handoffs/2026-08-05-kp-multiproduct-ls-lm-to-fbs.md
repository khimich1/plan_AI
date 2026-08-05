# Handoff: КП мультипродукт (ЛС → ЛМ → мостовые → ФБС)

> **Дата:** 2026-08-05  
> **Ветка:** `aleksey_web`  
> **Статус:** ЛС, ЛМ, мостовые сваи и **ФБС** реализованы (не закоммичено). **Следующий продукт: TBD.** Имеет смысл обсудить generic multi-product framework.  
> **Цель этого файла:** продолжить работу в новом окне/сессии без потери контекста.  
> **Актуально:** idea [`ai_docs/ideas/kp-bridge-piles-and-fbs.md`](../ideas/kp-bridge-piles-and-fbs.md); FBS report [`ai_docs/develop/reports/2026-08-05-kp-fbs-implementation.md`](../reports/2026-08-05-kp-fbs-implementation.md).

---

## Как стартовать новую сессию

1. Прочитать этот handoff.
2. Прочитать `.cursor/skills/plan-web-context/SKILL.md`.
3. Для следующего продукта: SPECIFY по новому прайсу (или обсудить framework).
4. Прайс ФБС (уже импортирован локально): `банк знаний/7_5 В прайс на ФБС  от 03.08.2026.xlsx`.

**Не коммитить** без явной просьбы пользователя.

---

## Что сделано

### 1–3. ЛС / ЛМ / мостовые сваи — DONE

См. предыдущие секции / reports:
- `ai_docs/develop/reports/2026-08-05-kp-stair-steps-implementation.md`
- `ai_docs/develop/reports/2026-08-05-kp-stair-marches-implementation.md`
- `ai_docs/develop/reports/2026-08-05-kp-bridge-piles-implementation.md`

### 4. КП на ФБС — DONE

| Артефакт | Путь |
|----------|------|
| Spec | `ai_docs/specs/kp-fbs.md` |
| Plan | `ai_docs/develop/plans/2026-08-05-kp-fbs.md` |
| Report | `ai_docs/develop/reports/2026-08-05-kp-fbs-implementation.md` |

- `product_type=fbs`, UI «ФБС»
- Grades: **B7_5 / B20 / B22_5 / B25**, default **B25**, плотная матрица
- Таблицы: `fbs_prices` (pb.db), `kp_fbs` (plita.db)
- Импорт: `python scripts/import_fbs_prices_from_xlsx.py "банк знаний/7_5 В прайс на ФБС  от 03.08.2026.xlsx" --sheet Прайс` → **56** строк (14×4)
- Без T/В alias groups
- OCR в MVP; preview hide until confirm; client-step errors не на вводе
- Production = plates only; shared `kp_id`

### 5. Сопутствующие UX/infra фиксы (все продукты)

См. предыдущий handoff: CSRF bootstrap, `/auth/me` invalidate, preview hide, client-step errors, сноска доборов удалена.

---

## Продуктовая линейка

```
Плиты → Сваи → Ступени ЛС → Марши ЛМ → Мостовые сваи → ФБС (DONE) → TBD
```

- Одно КП = один `product_type`.
- Production / СГП: whitelist **`plates`** only.
- Нумерация: общая серия `kp_id`.
- **Generic framework** — после ФБС имеет смысл обсудить рефакторинг дублирования (clone-паттерн разросся).

| Продукт | Модель цены | Шаблон |
|---------|-------------|--------|
| Плиты | dimensions + load | plates |
| Сваи | mark + grade | piles |
| Ступени ЛС | mark only | steps |
| Марши ЛМ | mark + grade | marches |
| Мостовые сваи | mark + grade (B25/B30) + aliases | bridge_piles |
| **ФБС** | mark + grade (B7_5/B20/B22_5/B25) | **fbs** (= piles + bridge wiring) |

---

## Ключевые файлы ФБС

### Domain / prices
- `core/fbs_price_db.py`, `fbs_line_parser.py`, `fbs_format_prompt.py`, `fbs_text_normalizer.py`
- `core/ocr/fbs_parser_gate.py`
- Shared: `commercial_pricing.py`, `commercial_offer*.py`, `kp_db_schema.py`, persistence / offers_read

### App
- `app/services/commercial_fbs_service.py`
- endpoints `.../fbs`, `.../ai`, `.../grades`

### Frontend
- `ProductTypePicker` «ФБС»
- `FbsInputStep`, `KpFbsPreviewPanel`, `buildFbsPreviewRows`, `fbsGrades`
- Archive badge/filter «ФБС»

### Scripts
- `scripts/import_fbs_prices_from_xlsx.py`

---

## Команды

```bash
./run+logs.sh

source venv/bin/activate
python scripts/import_fbs_prices_from_xlsx.py \
  "банк знаний/7_5 В прайс на ФБС  от 03.08.2026.xlsx" --sheet Прайс

pytest tests/ -k "fbs or bridge_pile or march or step or pile or wizard" -q
cd frontend && npm run typecheck && npm run test && npm run build
```

UI: http://localhost:5173/commercial-offer/new  

---

## Git / незакоммиченное

- Ветка: **`aleksey_web`**
- Много modified + untracked (steps, marches, bridge_piles, **fbs**, docs, CSRF/run fixes)
- **Коммита не было**
- `pb.db` локально с `fbs_prices` после CLI-импорта

---

## Следующий шаг: TBD / framework

1. Выбрать следующий продукт (или пауза на приёмку ФБС).
2. **Обсудить** generic multi-product framework vs продолжение точечных клонов — дублирование wiring уже заметное (5+ продуктов с grade/mark).
3. Manual browser smoke ЛС/ЛМ/мостовых/ФБС перед коммитом желателен.

---

## Известные gaps / не блокеры

- Нет dedicated OCR pipeline tests для новых продуктов
- `ShipmentDrawer` vitest timeout — flaky, не от ФБС
- `test_commercial_web_flow.py` старые падения (если всплывут вне focused `-k`) — не от ФБС
- Manual browser smoke желателен

---

## Definition of done для handoff-приемника

Новая сессия может:
1. Поднять `./run+logs.sh` и создать КП «Ступени» / «Марши» / «Мостовые сваи» / **«ФБС»** end-to-end.
2. Решить next product TBD или начать обсуждение framework.
3. Не ломать существующие продукты и не возвращать CSRF loop / manager error на шаге 1.
