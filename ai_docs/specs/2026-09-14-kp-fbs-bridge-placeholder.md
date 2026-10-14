# Spec: Пример ФБС и мостовых свай из прайса

> **Тип:** copy fix (wizard placeholders)  
> **Фаза SDD:** SPECIFY (этот документ) → PLAN ✅ → TASKS ✅ → IMPLEMENT ⏳  
> **Дата:** 2026-09-14  
> **Статус:** на реализацию после review спеки и плана  
> **Идея:** [`ai_docs/ideas/kp-fbs-bridge-placeholder.md`](../ideas/kp-fbs-bridge-placeholder.md)  
> **Plan:** [`ai_docs/develop/plans/2026-09-14-kp-fbs-bridge-placeholder.md`](../develop/plans/2026-09-14-kp-fbs-bridge-placeholder.md)  
> **Связанные:** [`kp-fbs.md`](./kp-fbs.md), [`kp-bridge-piles.md`](./kp-bridge-piles.md), план 2026-09-08 §5 ([`2026-09-08-kp-constructor-merge-input-steps.md`](../develop/plans/2026-09-08-kp-constructor-merge-input-steps.md))

---

## ASSUMPTIONS I'M MAKING

1. **Это «да» заказчика из плана 2026-09-08 §5.** Copy-paste leftover `С120.35-12 B25 5` / `С120.35-13и 3` у `fbs` и `bridge_piles` больше не сохраняем. Канон — две реальные марки из **текущего** прайса, не выдуманные SKU и не фикстуры тестов.
2. **Один мастер, один конфиг.** Менеджер и admin ходят на `/new` (`RequireRole allowedRoles={["admin", "manager"]}` → `CommercialOfferCreatePage` → тот же `PRODUCT_TYPE_CONFIG`). Роль-специфичных плейсхолдеров нет.
3. **Источник SKU — живой `pb.db` на момент реализации**, тот же файл, что читает рантайм: `core.project_paths.PRICE_DB_PATH` (`PRICE_DB_PATH` env, иначе `PB_DB_PATH`, иначе `PROJECT_ROOT/pb.db`). Таблицы: `fbs_prices`, `bridge_pile_prices`.
4. **Строки замораживаются в фронтовом конфиге.** Нет API «дай пример», нет запроса к БД при открытии шага. Если прайс позже сменится — это отдельный тикет.
5. **Пустой плейсхолдер — валидный исход.** Если в таблице меньше двух марок с `price > 0`, `labels.placeholder === ""`. HTML `placeholder=""` — пустое поле без примера. Не копировать сваи, не брать `ФБС 9.3.6-Т` / `C8-35T1` из pytest, если их нет в **этой** БД.
6. **Парсер / OCR / lookup / preview / calculate не меняются.** Меняется только копи в конфиге и пин-тесты, которые это копи замораживают.
7. **Другие типы КП вне scope.** Плейсхолдеры `plates` / `piles` / `steps` / `marches` не трогать.

→ Locked. Не переоткрывать.

---

## Objective

Сейчас шаг 1 ФБС и мостовых свай показывает пример обычных свай (`С120.35…`). Менеджер или admin вставляет его как есть → «Обработать текст» → строки **не в прайсе** этого типа.

Нужно: в пустом поле списка — две строки из текущего прайса (или пусто, если марок не хватает). Вставил как есть → получил цены.

**Пользователь:** менеджер **и** admin, шаг 1 конструктора КП (`/new`), типы `fbs` и `bridge_piles`.

**Успех:** paste placeholder as-is → «Обработать текст» → priced rows, без «не найдено в прайсе» для этих двух строк. Если в таблице меньше двух пригодных марок — поле без примера.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | **менеджер** на шаге 1 ФБС с пустым полем | видеть пример из прайса ФБС (не `С120.35…`) | вставить его как есть и получить цены |
| US-2 | **admin** на том же шаге 1 ФБС | тот же пример, тот же мастер | не учить второй UX |
| US-3 | **менеджер** на шаге 1 мостовых свай | видеть две марки из `bridge_pile_prices` (алиас как в колонке `mark`) | lookup нашёл цену, без «не в прайсе» |
| US-4 | **admin** на шаге 1 мостовых свай | тот же плейсхолдер, что у менеджера | один конфиг |
| US-5 | **менеджер или admin**, если в прайсе &lt; 2 марок этого типа | пустое поле, без чужого примера | не вставлять сваи / выдуманные SKU |

### Reframed success criteria

| Требование | Измеримый критерий |
|------------|-------------------|
| «Не сваи» | `PRODUCT_TYPE_CONFIG.fbs.labels.placeholder` и `…bridge_piles…` **не** равны `С120.35-12 B25 5\nС120.35-13и 3` |
| «Из этой БД» | Каждая непустая строка — `mark` + priced `concrete_grade` из той же `pb.db`, что открыли при реализации |
| «Обработать текст» | Paste as-is → preview с ценой по обеим строкам; нет unpriced-алерта на эти строки |
| «Пустой прайс» | &lt; 2 suitable marks в таблице → `placeholder === ""`; UI без примера |
| «Один мастер» | Нет ветвления по роли; `/new` для `admin` и `manager` читает один конфиг |
| «Только копи» | Diff продукта: `productTypeConfig.ts` (+ комментарий leftover) и `productTypeConfig.test.ts`. Парсер/OCR/pricing — 0 изменений |

---

## Problem (current vs desired)

### Current

В `frontend/src/features/commercial-offer/lib/productTypeConfig.ts`:

- `bridge_piles.labels.placeholder` = `"С120.35-12 B25 5\nС120.35-13и 3"` (коммент: leftover, план 2026-09-08 §5)
- `fbs.labels.placeholder` = то же

`SimpleProductInputStep` передаёт `placeholder={labels.placeholder}` в `SourceInputCard`. Пин-тесты в `productTypeConfig.test.ts` **замораживают leftover** и требуют `placeholder.length > 0` для **всех** типов, включая fbs/bridge.

План 2026-09-08 §5 явно откладывал фикс до «да» заказчика. Исторический канон из отчёта внедрения (`ФБС 9.3.6-Т 2`) **не** использовать, если марки нет в текущей БД.

### Desired

Два независимых плейсхолдера (ФБС отдельно, мостовые отдельно), выбранные по алгоритму ниже и вписанные литералами в конфиг.

---

## How to pick marks from the live price DB

**Когда:** только на IMPLEMENT, один раз. Результат — две строки (или `""`) в исходнике. Не в рантайме.

**Файл БД:** `str(core.project_paths.PRICE_DB_PATH)`. Не коммитить `pb.db`.

**Таблицы (подтверждено кодом):**

| Тип | Модуль | Таблица | PK | Lookup |
|-----|--------|---------|----|--------|
| ФБС | `core/fbs_price_db.py` | `fbs_prices` | `(mark, concrete_grade)` | `get_fbs_price` / `normalize_fbs_mark_for_lookup` (пробелы, регистр, `Т`↔`T`) |
| Мостовые сваи | `core/bridge_pile_price_db.py` | `bridge_pile_prices` | `(mark, concrete_grade)` | `get_bridge_pile_price` / `normalize_bridge_pile_mark_for_lookup` (`C`↔`С`, `B`↔`В`, `T`↔`Т`) |

Колонки, нужные для выбора: `mark TEXT`, `concrete_grade TEXT`, `price REAL`. Нулевые ячейки **не** импортируются (`price > 0` в импорте); всё равно фильтровать `price > 0`.

**Алгоритм (одинаковый для каждой таблицы, независимо):**

1. Suitable mark = DISTINCT `mark`, у которого есть хотя бы одна строка с `price > 0`.
2. Стабильный порядок: `ORDER BY mark` (SQLite default collation).
3. Если таких марок **&lt; 2** → placeholder этой номенклатуры = `""`. Стоп. Не брать соседнюю таблицу, не выдумывать SKU.
4. Иначе взять **первые две** марки.
5. Для каждой марки выбрать класс, у которого в БД есть цена, в том же порядке, что рантайм:
   - ФБС: `resolve_default_fbs_grade(mark)` — `B25`, если priced, иначе первый из `FBS_GRADE_CODES` (`B7_5`, `B20`, `B22_5`, `B25`) среди available.
   - Мостовые: `resolve_default_bridge_pile_grade(mark)` — единственный available; иначе `B25` если priced; иначе первый из `B25` / `B30`.
6. Токен класса в строке плейсхолдера — **display-форма**, которую парсер уже принимает:
   - ФБС: `formatFbsGradeLabel` → `B7.5` / `B20` / `B22.5` / `B25` (не сырой `B7_5` / `B22_5` из SQLite).
   - Мостовые: `B25` / `B30`.
7. Количество: первая строка **2**, вторая **3**.
8. Формат линии как у свай: `{mark} {grade} {qty}`, две линии через `\n`.
9. `mark` — **как в колонке таблицы** (после импорта алиасы уже разрезаны `split_alias_marks`: `C8-35T4; C8-35В4` → отдельные `mark`). Не склеивать через `;`. Lookup должен принять эту форму без правки парсера.

Пример формы (не канон, пока не подтверждён живой БД):

```
ФБС 12.4.6-Т B25 2
ФБС 24.6.6-Т B20 3
```

```
C8-35T1 B25 2
C8-35В4 B30 3
```

Записать выбранные литералы в конфиг и в пин-тест. Комментарий leftover / «до решения заказчика» — удалить или заменить на «заморожено из `pb.db` YYYY-MM-DD».

---

## Tech Stack

Как у конструктора КП: React 19, TypeScript, Vite, Vitest. Прайс — SQLite `pb.db` (`fbs_prices` / `bridge_pile_prices`). Backend Python не меняем.

## Commands

```bash
# Выбор SKU (IMPLEMENT, read-only). Путь = PRICE_DB_PATH || PB_DB_PATH || ./pb.db
python - <<'PY'
from core.project_paths import PRICE_DB_PATH
from core.fbs_price_db import resolve_default_fbs_grade
from core.bridge_pile_price_db import resolve_default_bridge_pile_grade
import sqlite3
print("db", PRICE_DB_PATH)
con = sqlite3.connect(str(PRICE_DB_PATH))
for table, resolver in (("fbs_prices", resolve_default_fbs_grade), ("bridge_pile_prices", resolve_default_bridge_pile_grade)):
    marks = [r[0] for r in con.execute(
        f"SELECT DISTINCT mark FROM {table} WHERE price > 0 ORDER BY mark"
    )]
    print(table, "count", len(marks), "first2", marks[:2])
    if len(marks) >= 2:
        for m in marks[:2]:
            print(" ", m, resolver(m, db_path=str(PRICE_DB_PATH)))
PY

cd frontend && npx vitest run src/features/commercial-offer/lib/productTypeConfig.test.ts
cd frontend && npm run typecheck
```

Ручной smoke (живой `./run+logs.sh`, `/new`):

1. Войти как **менеджер**. Новый КП → ФБС → пустое поле шага 1 показывает новый placeholder (или пусто). Вставить его as-is → «Обработать текст» → две строки с ценой, без unpriced.
2. Тот же пользователь: тип «Мостовые сваи» — то же.
3. Выйти, войти как **admin**. Повторить 1–2: тот же placeholder, тот же результат.
4. Если Task 1 записал `""` для типа — поле без примера, нет `С120.35`.

## Project Structure

```
frontend/src/features/commercial-offer/lib/productTypeConfig.ts
  → labels.placeholder для fbs и bridge_piles; JSDoc leftover
frontend/src/features/commercial-offer/lib/productTypeConfig.test.ts
  → пин точных строк; completeness: empty allowed только для fbs/bridge_piles
frontend/src/features/commercial-offer/components/steps/SimpleProductInputStep.tsx
  → не менять (уже placeholder={labels.placeholder})
core/fbs_price_db.py          → read-only: таблица fbs_prices, resolve_default_fbs_grade
core/bridge_pile_price_db.py  → read-only: таблица bridge_pile_prices, split aliases, lookup
```

## Code Style

Минимальный diff. Литералы как у свай — две линии, `\n`, без trailing newline. Не добавлять хелпер «собери пример из API». Не ослаблять пины других типов.

JSDoc `ProductTypeLabels.placeholder` сейчас: «Exact current placeholder, including known copy-paste leftovers» — после фикса убрать «leftovers»; empty string для fbs/bridge допустима.

Completeness-тест сегодня:

```ts
expect(labels.placeholder.length, type).toBeGreaterThan(0);
```

После фикса: для `plates` / `piles` / `steps` / `marches` по-прежнему `length > 0`; для `fbs` / `bridge_piles` — либо замороженная непустая строка, либо `""`.

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Vitest `productTypeConfig.test.ts` | Точные новые строки (или `""`) для `fbs` и `bridge_piles`. Пины plates/piles/steps/marches без изменений. Completeness не валит empty fbs/bridge. |
| Не добавлять | pytest парсера, OCR, preview, live SQL в CI (CI не обязан иметь тот же `pb.db`) |
| Manual | Сценарии 1–4 в Commands, менеджер **и** admin |

Пин-тесты **не** читают `pb.db`: они замораживают строки, выбранные в Task 1.

## Boundaries

- **Always:** плейсхолдеры fbs/bridge из текущей БД или `""`; пины совпадают с конфигом; plates/piles/steps/marches не трогать; один конфиг на обе роли.
- **Ask first:** живой endpoint примера; менять парсер, чтобы принять новую форму марки; править плейсхолдеры других типов.
- **Never:** фоллбек на `С120.35…`, на тестовые `ФБС 9.3.6-Т` / `C8-35T1`, если их нет в этой БД; рандом при рендере; кнопка «Вставить пример»; легенда формата; автокомплит; OCR-варнинги; коммитить `pb.db` / `.env`; менять pricing.

## Success Criteria

| # | Критерий |
|---|----------|
| S1 | Placeholder ФБС ≠ пример свай; либо 2 линии из `fbs_prices`, либо `""` |
| S2 | Placeholder мостовых ≠ пример свай; либо 2 линии из `bridge_pile_prices` (mark as stored), либо `""` |
| S3 | Менеджер: paste → «Обработать текст» → priced, без unpriced на эти строки (если placeholder непустой) |
| S4 | Admin: то же на том же `/new` |
| S5 | Empty-DB / &lt;2 марок: `""`, в UI нет `С120.35` |
| S6 | `productTypeConfig.test.ts` зелёный; `npm run typecheck` чистый |
| S7 | Нет diff в парсере, OCR, price_db, preview, calculate |

## Edge cases

| Кейс | Ожидание |
|------|----------|
| Ровно 1 suitable mark | `""` (нужно **две**) |
| 0 строк / нет таблицы | `""` |
| Марка только с одним классом (типичные мостовые) | Этот класс в строке, даже если B30 |
| Алиас-группа в Excel | В БД уже отдельные `mark`; в placeholder — значение колонки, не `a; b` |
| Поле уже с текстом черновика | Placeholder не виден (нативный HTML); не чинить в этом тикете |
| Grade `B7_5` в SQLite | В строке `B7.5` |

## Out of Scope (copied from idea)

- Живой пример с бэка при каждом открытии шага.
- Случайные марки каждый рендер.
- Эндпоинт «дай пример».
- Кнопка «Вставить пример».
- Легенда формата, автокомплит.
- OCR / правки варнингов OCR.
- Аудит плейсхолдеров plates / piles / steps / marches.
- Разные тексты для менеджера и админа.
- Фоллбек на сваи или pytest-фикстуры, если БД пустая.

## Open Questions

Нет.
