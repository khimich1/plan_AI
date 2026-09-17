# Spec: Стол прайсов

> **Источник идеи:** [`ai_docs/ideas/price-desk-role.md`](../ideas/price-desk-role.md)  
> **План:** [`ai_docs/develop/plans/2026-09-15-price-desk-role.md`](../develop/plans/2026-09-15-price-desk-role.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS → IMPLEMENT  
> **Статус:** план согласован, до реализации  
> **Дата:** 2026-09-15  
> **Связанные модули:** `app/schemas/auth.py`, `app/dependencies/auth.py`,
> `frontend/src/shared/lib/roleRoutes.ts`, `core/*_price_db.py`, `core/price_db.py`,
> `scripts/import_*_prices_from_xlsx.py`. Рядом, не это: `POST /api/v1/nomenclature/import-1c`.

---

## Assumptions I'm Making

1. **Админ видит стол прайсов**, как ГСМ. Экономист — нет КП/производства/ГСМ/логистики.
   Роль нужна для изоляции человека, не чтобы запретить админу поддержку.
2. **Пользователя-экономиста заводит админ** через существующий `POST /auth/register`
   (расширяем `Literal` ролей). Отдельной саморегистрации нет.
3. **Превью и запись — два запроса, файл оба раза.** `sha256` из превью должен
   совпасть при apply. Без серверного кэша Excel (нет гонки воркеров и хвостов в `/tmp`).
4. **У `prices` (плиты) нет `price_list_date` / `imported_at`.** Добавляем колонки
   по тому же шаблону, что у `pile_prices` — иначе экран «дата прайса» для плит пустой.
5. **Дробные цены не-плит** пишем как в файле (CLI уже так делает). Округление до рубля
   не делаем.
6. **Черновик КП**, открытый в момент заливки, при следующем превью может показать
   новые цены. В v1 не фиксируем цену в черновике.
7. **Журнала файлов нет.** В v1 хватает `imported_at` / `price_list_date` в таблицах.
8. **Фикстуры в тестах — синтетические листы**, не коммитим живые прайсы из
   `банк знаний/`. Реальные файлы 07.09 / 17.08 — ручной прогон при приёмке.
9. **Лимит загрузки** — тот же `COMMERCIAL_UPLOAD_MAX_BYTES` (по умолчанию 50 МБ).
   Цельные сваи ~1.4 МБ, расчёт ПБ ~1.1 МБ.
10. **Web-only.** Бот вне scope. Без новых pip/npm зависимостей (`xlrd` уже нужен
    pandas для `.xls`).

→ Если что-то из этого неверно — править spec до плана.

---

## Decisions locked (ideation 2026-09-15)

| # | Тема | Решение |
|---|------|---------|
| D1 | Кто | Роль `economist`. Стол: `admin` + `economist` |
| D2 | UI | `/prices`, один drop-zone, без живой таблицы |
| D3 | Подхват в КП | Lookup в `pb.db` при расчёте. Архивные КП не трогаем |
| D4 | Плиты | Лист «Прайс», колонка Никиты. Не М400/М500. Не 16/21. 12,5 → `load_code=12` |
| D5 | ФБС / ЛМ / ступени / мостовые / цельные | Лист «Прайс», текущие `import_*` |
| D6 | Составные | 422, текст «группа пока не ведётся» |
| D7 | 1С GUID | Другой стол (`Import1cDialog`). Сюда не класть |
| D8 | Маршрут файла | Якорь в **имени**, затем один парсер. Не «кто больше строк» |
| D9 | Запись | Только после подтверждения. `INSERT OR REPLACE`. Пропавшие марки остаются |
| D10 | Форматы | `.xls` и `.xlsx` |
| D11 | Apply | Повторная загрузка + `file_sha256` с превью |
| D12 | Даты плит | Колонки `price_list_date`, `imported_at` в `prices` |

---

## Objective

Экономист без доступа к КП сам кладёт заводской Excel. После подтверждения диффа
новые КП считаются по новым ценам.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | экономист | зайти и увидеть только «Прайсы» | не открывать чужие разделы |
| US-2 | экономист | бросить файл и увидеть группу + дифф | не записать вслепую |
| US-3 | экономист | подтвердить запись | следующие КП брали эти цены |
| US-4 | экономист | увидеть дату прайса по группам | знать, что сейчас в программе |
| US-5 | экономист | получить понятную ошибку на составные / чужой файл | не испортить сваи |
| US-6 | админ | завести пользователя с ролью экономист | не отдавать ему admin |
| US-7 | менеджер | считать новое КП как раньше | не ходить на стол прайсов |

### Reframed success criteria

| Требование | Критерий |
|------------|----------|
| Изоляция | `economist` → 200 на `/prices` и `GET /api/v1/prices/status`; 403 на КП, архив, производство, ГСМ, логистику, `POST /nomenclature/import-1c` |
| Админ | админ видит `/prices` и может apply |
| Превью без записи | `POST .../preview` не меняет `pb.db` |
| Apply = превью | `sha256` не совпал → 409, БД не меняется |
| Плиты | 277 цен Никиты на образце 17.08; 0 строк с нагрузкой 16/21; 12,5 как 12 |
| ФБС | 56 строк на образце 07.09 |
| ЛМ | 35 строк |
| Ступени | 42 строки |
| Мостовые | 114 строк |
| Цельные | 1590 строк |
| Составные | 422, в `pile_prices` / `bridge_pile_prices` 0 новых строк |
| Чужой парсер | файл ФБС не пишется в `pile_prices` |
| Дифф | в ответе: `changed` / `new` / `missing` / `unchanged` (счётчики + до 30 примеров) |
| Пропавшие | ключ есть в БД, нет в файле → в `missing`, после apply цена старая |
| КП | новое КП после apply берёт новую цену; сохранённое — старый `unit_price` |
| 1С | кнопка «Выгрузка 1С» у менеджера на месте, у экономиста её нет |

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2, SQLite `pb.db` |
| Domain | существующие `core/{pile,bridge_pile,fbs,march,step}_price_db.py`; новый парсер плит |
| Auth | cookie-сессия; `require_roles("admin", "economist")` |
| Frontend | React 19, TS, React Router 7, TanStack Query |
| Upload | `read_upload_file_capped` |
| Tests | pytest; vitest + Testing Library |

---

## Commands

```bash
source .venv/bin/activate   # или venv/bin/activate
pytest tests/test_price_desk_classify.py tests/test_plate_nikita_parser.py \
  tests/test_price_desk_diff.py tests/test_price_desk_api.py \
  tests/test_fbs_price_import.py tests/test_pile_price_import.py \
  tests/test_march_price_import.py tests/test_step_price_import.py \
  tests/test_bridge_pile_price_import.py tests/test_gsm_auth.py -q

cd frontend && npm run typecheck
cd frontend && npm run test -- --run src/features/price-desk src/shared/lib/roleRoutes.test.ts src/app/router
cd frontend && npm run build
```

Приёмка на живых файлах (не в CI, пути локальные):

```bash
# превью без записи, затем apply — сверка счётчиков со spec
```

---

## Project Structure

```
ai_docs/ideas/price-desk-role.md
ai_docs/specs/price-desk-role.md                          # этот spec

core/price_desk_classify.py                               # NEW: якорь имени → kind
core/plate_nikita_parser.py                               # NEW: лист Прайс, колонка Никиты
core/price_desk_diff.py                                   # NEW: diff текущей таблицы и файла
core/price_db.py                                          # колонки даты; import из nikita-парсера
core/{pile,bridge_pile,fbs,march,step}_price_db.py        # без смены формата листа; вызов как сейчас

app/schemas/auth.py                                       # + economist
app/schemas/price_desk.py                                 # NEW
app/dependencies/auth.py                                  # REQUIRE_PRICES
app/services/price_desk_service.py                        # NEW
app/api/v1/endpoints/price_desk.py                        # NEW
app/api/v1/router.py                                      # include

frontend/src/features/auth/types/user.ts
frontend/src/shared/lib/roleRoutes.ts
frontend/src/app/router/AppRouter.tsx
frontend/src/app/layout/AppHeader.tsx
frontend/src/features/price-desk/                         # NEW: страница, drop-zone, дифф
frontend/src/pages/prices/PricesPage.tsx                  # NEW

tests/test_price_desk_classify.py
tests/test_plate_nikita_parser.py
tests/test_price_desk_diff.py
tests/test_price_desk_api.py
```

---

## Code Style

### Классификация (имя файла, регистр/ё не важны)

Порядок **первый матч побеждает**:

| Порядок | Подстрока в имени | `product_kind` |
|--------|-------------------|----------------|
| 1 | `составн` | `composite` → не парсим, 422 |
| 2 | `расчет новых цен на пб` или (`расчет` и `пб`) | `plates` |
| 3 | `фбс` | `fbs` |
| 4 | `ступен` | `step` |
| 5 | `мостов` | `bridge_pile` |
| 6 | `цельн` | `pile` |
| 7 | `лм` | `march` |

Нет якоря → 422 «не понял группу по имени файла». После kind — проверка содержимого
(есть лист «Прайс», марки своей группы). Не сошлось → 422, не запись.

Запрещено: гонять все парсеры и брать max(len(rows)).

### Плиты: колонка Никиты

```python
NIKITA_HEADER = "цены по прайсу, который прислал никита"
ALLOWED_LOADS = {6, 8, 10, 12}  # 12.5 → 12; 16 и 21 отбрасываем

def parse_plate_nikita_rows(path: str) -> list[tuple[int, int, float]]:
    """Лист «Прайс». Имя ПБ {length}-12-{load}. Цена — колонка Никиты. price > 0."""
```

Текущий `parse_plate_price_rows_from_xlsx` **не** вызываем для этого файла (на 17.08
даёт 0 строк). Старый матричный формат («6 нагрузка») стол в v1 не принимает:
экономист приносит расчёт ПБ.

Имя `ПБ 17-12-6` → `length_dm=17`, `load_code=6`. `ПБ 18-12-12.5` → `load_code=12`.
Пустая ячейка Никиты — строка пропускается (не 0).

### API

```
GET  /api/v1/prices/status
POST /api/v1/prices/preview   multipart file
POST /api/v1/prices/apply     multipart file + file_sha256
```

Только `admin` | `economist`. Остальные → 403.

`preview` считает sha256 сырых байт, парсит, сравнивает с текущей таблицей группы,
**не пишет**.

`apply`: тот же файл, sha256 совпал, затем существующий `import_*` / nikita-import.
Не совпал → 409 «файл не тот, что в превью».

Цена ≤ 0 не импортируется (как ФБС).

Дифф, ключ группы:

| kind | ключ |
|------|------|
| plates | `(length_dm, load_code)` |
| fbs / pile / bridge_pile / march | `(mark, concrete_grade)` |
| step | `(mark,)` |

Сравнение цен: `abs(a - b) > 0.005` → `changed`.

Ответ превью (сжатый):

```python
class PriceDeskPreview:
    product_kind: Literal["plates", "fbs", "march", "step", "bridge_pile", "pile"]
    price_list_date: str | None  # из имени YYYY-MM-DD
    file_sha256: str
    parsed_rows: int
    changed: int
    new: int
    missing: int
    unchanged: int
    examples: PriceDeskDiffExamples  # до 30 штук на категорию
```

`status`: по каждой из шести групп `price_list_date`, `imported_at`, `row_count`
(`MAX(imported_at)` / `MAX(price_list_date)`). Пустая таблица — `null`.

Человекочитаемые 422: «группа пока не ведётся»; «это не прайс ПБ: нет колонки Никиты»;
«лист «Прайс» не найден»; «не понял группу по имени файла».

### Роли и UI

- `RegisterUserRequest.role` += `"economist"`.
- `REQUIRE_PRICES = require_roles("admin", "economist")`.
- `defaultRouteForRole("economist")` → `/prices`.
- Нав: пункт «Прайсы» для admin+economist. У экономиста нет Конструктора, Архива,
  «Выгрузка 1С», Производства, Логистики, ГСМ.
- Страница: таблица дат по группам + drop-zone `.xls,.xlsx` → превью → кнопки
  «Отмена» / «Записать в программу». Пока нет превью, «Записать» нет.

Паттерн диалога: `ImportScheduleDialog` (drag-and-drop), но **два шага**, не сразу import.

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Classify | якоря; `составн` раньше `цельн`; без якоря → ошибка; регистр/ё |
| Nikita parser | синтетический лист: 6/8/10/12.5 пишутся, 16/21 нет, М400 игнорируется, пустой Никита skip, 12.5→12 |
| Diff | changed/new/missing/unchanged; apply не трогает missing |
| Cross-route | файл с марками ФБС и именем «цельные» → 422, не 56 строк в `pile_prices` |
| API | economist preview 200 и 0 UPDATE; apply без sha256 422; чужой sha256 409; manager 403; составные 422 |
| Auth | register `economist`; default route; 403 на `/gsm` и `/new` |
| Vitest | drop-zone; дифф виден до записи; Записать зовёт apply с sha256; экономист не видит «Выгрузка 1С» |
| Регрессия | существующие `test_*_price_import.py` зелёные |

Живые файлы 07.09 / 17.08 — чеклист приёмки, не обязательный CI.

---

## Boundaries

**Always:**

- Сначала имя файла, потом один парсер
- Превью без записи
- `read_upload_file_capped`
- Тесты classify + nikita до UI
- ACL на API, не только скрытие пункта меню

**Ask first:**

- Вычитание марок, которых нет в файле
- Округление не-плит до рубля
- Журнал загрузок (кто/файл)
- Приём старого матричного прайса плит («6 нагрузка»)
- Составные сваи как продукт КП
- Менять бренд в шапке («Коммерческие предложения»)

**Never:**

- Писать расчёт М400/М500 плит в `prices`
- Класть выгрузку 1С в этот endpoint
- Автовыбор парсера по числу строк
- Автоapply без диффа
- Пересчёт `unit_price` в сохранённых КП
- Commit секретов / живых `pb.db` / файлов из `банк знаний/`
- Давать `economist` роль admin «чтобы было проще»

---

## Out of scope (из идеи)

- Составные сваи как номенклатура КП
- `POST /nomenclature/import-1c` и очередь GUID
- Живое редактирование ячеек
- Watch-папка
- Лист «Цена для покупателя»
- Нагрузки плит 16 и 21

---

## Open Questions

Нет. Закрыты в идее и в assumptions выше. Расхождения — править этот spec до плана.
