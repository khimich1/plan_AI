# Spec: График поставки — партии с датами внутри одного заказа

Дата: 2026-08-07. Статус: draft, на ревью.
Идея: `ai_docs/ideas/grafik-postavki-partii.md` (направление B, решения зафиксированы там).

## Objective

Менеджер после согласования счёта разбивает заказ (КП) на партии с датами
«поставка с/по» и «произвести до», сразу видит реалистичность сроков
(светофор по загрузке производства) и одной кнопкой получает документ
«График поставки» (XLSX + PDF) для клиента. По ходу производства светофор
пересчитывается по остатку партии — раннее предупреждение о срыве.

Пользователь: менеджер по продажам (роли admin/manager, как у архива КП).
Источники данных: импорт нашего XLSX-шаблона (высылается клиенту) или
ручной ввод «со слов». Пока нет номера счёта — график привязан к КП.

Принятые продуктовые решения (из сессии idea-refine):
- Направление B: партии как сущность + живой светофор + импорт шаблона + документ.
- Даты задаёт менеджер вручную (без автоподстановки буфера).
- Светофор «живой»: пересчёт по остатку при каждом открытии.
- СГП: вычитаем только произведённое, привязанное к этому КП; свободный
  склад не трогаем (ошибка только в безопасную сторону). Закрытие со склада =
  существующая операция привязки плиты к КП → автоматически учтётся.
- Истории версий нет: документ регенерируется целиком, файлы с датой редакции
  в имени не перезаписываются.
- График не резервирует дорожки: обязательство ≠ план.

## Tech Stack

- Backend: Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`).
- Frontend: React 19, TS, Vite, TanStack Query, feature-структура
  `api/ components/ hooks/ types/`.
- Документы: существующий стек генерации (openpyxl для XLSX, reportlab для
  PDF — как `core/commercial_offer.py`, `core/commercial_offer_xlsx.py`).
- Новых зависимостей не требуется.

## Data Model

Новые таблицы в `core/kp_db_schema.py` (стиль `CREATE TABLE IF NOT EXISTS`,
FK на `KP_offers` с `ON DELETE CASCADE`, как `kp_plates`):

```sql
delivery_schedule (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  kp_id INTEGER NOT NULL UNIQUE REFERENCES KP_offers(kp_id) ON DELETE CASCADE,
  invoice_number TEXT,                -- № счёта; NULL пока не выставлен
  contract_number TEXT,               -- № договора (шапка документа)
  status TEXT NOT NULL DEFAULT 'draft',  -- draft | active | completed
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
)

delivery_batch (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  schedule_id INTEGER NOT NULL REFERENCES delivery_schedule(id) ON DELETE CASCADE,
  name TEXT NOT NULL,                 -- «1 этаж, 2 подъезд» — свободный текст
  deliver_from TEXT NOT NULL,         -- ISO date
  deliver_to TEXT NOT NULL,           -- ISO date
  produce_by TEXT NOT NULL,           -- ISO date, вручную
  sort_order INTEGER NOT NULL DEFAULT 0
)

delivery_batch_item (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  batch_id INTEGER NOT NULL REFERENCES delivery_batch(id) ON DELETE CASCADE,
  plate_id INTEGER NOT NULL REFERENCES kp_plates(id) ON DELETE CASCADE,
  qty INTEGER NOT NULL CHECK (qty >= 1)
)
```

Инвариант (проверяется в сервисе, не в БД): Σ qty по всем batch_item
позиции ≤ `kp_plates.qty`. Частичная разбивка легальна (статус
«разбито N из M» — производное, не хранится).

Один график на КП (`UNIQUE kp_id`) — пересогласование = правка того же
графика + регенерация документа (истории версий нет, см. решения выше).

## API

Новый роутер `app/api/v1/endpoints/delivery_schedule.py`,
prefix `/commercial/archive/{kp_id}/delivery-schedule`, все хендлеры под
`require_roles("admin", "manager")`, схемы в `app/schemas/delivery_schedule.py`,
логика в `app/services/delivery_schedule_service.py`, чистая математика в
`core/delivery_schedule_check.py`.

| Метод | Путь | Назначение |
|---|---|---|
| GET | `/` | График + светофор по каждой партии (пересчёт при каждом вызове) |
| PUT | `/` | Полная замена партий (создание/редактирование одной операцией) |
| POST | `/import` | Импорт XLSX-шаблона → черновик партий (без сохранения) |
| GET | `/template` | Скачать пустой XLSX-шаблон для клиента |
| GET | `/document?fmt=xlsx\|pdf` | Сгенерировать документ; файл сохраняется в архив КП |

PUT принимает полный набор партий (идемпотентная замена) — проще транзакция,
проще валидация инварианта Σ ≤ qty, нет частичных состояний.
POST /import не сохраняет: возвращает распарсенный черновик + список
несматченных строк; менеджер проверяет и жмёт «Сохранить» (тот же PUT).

## Светофор: алгоритм (`core/delivery_schedule_check.py`)

Переиспользование существующего:
- Остатки по позициям: `KpReadinessService.list_positions` (агрегирует
  `kp_plates` + `completed_plates` по kp_id — это и есть «произведено,
  привязанное к КП»).
- Ёмкость: `days_info` глобального календаря (occupied/max на дату) +
  `core/work_calendar.py` (рабочие дни).
- Константы ёмкости: дорожка 101 м, 5 дорожек/день — сейчас дублируются в
  `archive_service.py` (`_MAX_TRACK_LENGTH_M`, `_DAYS_PER_TRACK_FACTOR`).
  Согласованный мини-рефакторинг: вынести в `core/production_capacity.py`,
  `archive_service` импортирует оттуда (diff 2 строки, поведение не меняется).

Шаги:
1. Остаток партии = Σ max(0, batch_item.qty − produced_linked) по позициям.
2. Потребность в дорожках = Σ (остаток × length_m) / 101 × коэф. запаса 1,15,
   округление вверх.
3. Симуляция от сегодня по рабочим дням: свободно = max − occupied;
   партии обрабатываются в порядке `produce_by`; каждая съедает ёмкость.
4. Расчётная дата готовности партии → статус:
   зелёный: готово ≤ produce_by − 5 раб. дней;
   жёлтый: готово ≤ produce_by;
   красный: позже produce_by + подсказка «нужно +N дорожек до <дата>»
   (дефицит ёмкости в окне сегодня → produce_by).

Пороги и коэф. 1,15 — константы модуля, калибруются после валидации
на 3–5 прошлых заказах (см. Success Criteria).

## XLSX-шаблон

`core/delivery_schedule_xlsx.py`: `build_template(path)` /
`parse_template(bytes, kp_plates) -> (batches_draft, unmatched_rows)`.

Колонки: `Партия | Поставка с | Поставка по | Произвести до | Марка | Кол-во`.
Матчинг марки → позиция КП: точное по `plate_name`, иначе строка уходит в
`unmatched_rows` с причиной. Даты: `ДД.ММ.ГГГГ` (парсер как
`core/execution_terms` по стилю, строгий).

## Frontend

Новая фича `frontend/src/features/delivery-schedule/`
(`api/ components/ hooks/ types/` по образцу `commercial-archive`).

- Точка входа: кнопка «График поставки» в `OfferDetailsDrawer`
  (график — атрибут конкретного КП; в шапке архива рядом с
  `CurrentPlanButton` кнопку НЕ ставим — там нет выбранного КП).
  В строке списка архива — бейдж «есть график» (если schedule существует).
- Экран/модал редактора: слева позиции КП с остатками к разбивке,
  справа карточки партий (имя, 3 даты, позиции с qty). Индикатор
  «разбито N из M» по каждой позиции. Светофор-чип на каждой партии.
- Импорт: dropzone XLSX → черновик подставляется в редактор,
  несматченные строки — списком сверху.
- Документ: кнопки «Скачать XLSX» / «Скачать PDF»; файлы также видны в
  существующем списке файлов КП.

## Commands

```bash
# Backend
source venv/bin/activate        # или .venv/
pytest tests/ -q
uvicorn app.main:app --reload

# Frontend
cd frontend && npm run dev
cd frontend && npm run test
cd frontend && npm run typecheck
cd frontend && npm run build

# Полный стек
./run+logs.sh
```

## Project Structure (новые/затрагиваемые файлы)

```
core/kp_db_schema.py                  # +3 таблицы (миграция в общем стиле)
core/delivery_schedule_check.py       # чистая логика светофора
core/delivery_schedule_xlsx.py        # шаблон + парсер + генерация XLSX
core/production_capacity.py           # константы 101 м / 5 дорожек (новое)
app/services/archive_service.py       # импорт констант из production_capacity
app/schemas/delivery_schedule.py      # pydantic-контракты
app/services/delivery_schedule_service.py
app/api/v1/endpoints/delivery_schedule.py
app/api/v1/endpoints/__init__.py      # регистрация роутера
frontend/src/features/delivery-schedule/**
frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx  # кнопка
tests/test_delivery_schedule_*.py     # схема, сервис, светофор, xlsx
frontend/src/features/delivery-schedule/**/*.test.tsx
```

## Code Style

Слои: роутер → сервис → core/репозиторий; ORM нет, чистый sqlite3 как в
`core/kp_db_*`. Хендлеры тонкие, роли через `Depends(require_roles(...))`:

```python
@router.get("", response_model=DeliveryScheduleView)
def get_schedule(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: DeliveryScheduleService = Depends(get_delivery_schedule_service),
) -> DeliveryScheduleView:
    return service.get_view(kp_id, user=user)
```

Frontend: TanStack Query хуки с `queryKey`-фабриками (по образцу
`useArchiveQueries.ts`), типы рядом в `types/`, сохранение файлов через
`saveBlobAs`.

## Testing Strategy

- pytest, новые тесты в `tests/` (зеркалят существующие по стилю,
  см. `tests/test_execution_terms.py`, `tests/test_archive_endpoints.py`):
  - схема: таблицы создаются идемпотентно, cascade от KP_offers;
  - сервис: PUT-валидация (Σ ≤ qty, даты from ≤ to, produce_by ≤ deliver_from
    — предупреждение, не ошибка), инвариант при повторном PUT;
  - светофор: детерминированные кейсы (зелёный/жёлтый/красный, частично
    произведённая партия, пересечение партий по ёмкости, выходные в
    work_calendar);
  - xlsx: round-trip template → parse; несматченные марки.
- Frontend: vitest — редактор партий (добавление/валидация остатков),
  рендер светофора по мок-ответу.
- Перед завершением: `pytest tests/ -q` + `npm run test` + `npm run typecheck`.

## Boundaries

- Always: тесты зелёные перед коммитом; минимальный diff, не трогать
  несвязанный код; схемы Pydantic — контракт, держать типы фронта в синхроне.
- Ask first: изменения схемы БД (эта спека — такое согласование);
  новые зависимости; правки `core/kp_db_*` вне новых таблиц; рефакторинг
  `archive_service` сверх импорта констант.
- Never: секреты/живые БД в коммит; обход `destructive_db_guard`; удаление
  или ослабление существующих тестов; резервирование дорожек плана под
  партии (осознанное продуктовое «нет»).

## Success Criteria

1. Менеджер разбивает КП уровня счёта 234 (31 марка × 21 партия) за
   ≤10 минут: импорт шаблона + правки.
2. Светофор на синтетическом кейсе с известным ответом (зелёный/жёлтый/
   красный) — совпадает; на 3–5 прошлых заказах расчётная дата готовности
   отличается от факта не более чем на ±20% (калибровка коэф.).
3. Частично произведённая партия: светофор считает по остатку
   (прогресс в completed_plates вычитается).
4. Документ XLSX открывается в Excel/LibreOffice, содержит шапку
   (договор/счёт/стороны) и таблицу партий; PDF генерируется тем же данными.
5. Старые файлы документа не перезаписываются (дата редакции в имени).
6. Импорт чужого файла с 3+ несматченными марками — черновик + понятный
   список проблем, ничего не падает.
7. Удаление КП удаляет график (cascade), 500-ок нет.

## Open Questions

Закрыты при ревью (2026-08-07):
- Валидация ёмкости (Success Criteria №2): заказы подбираются из БД
  автоматически (completed_date vs даты КП), без участия менеджера.
- Границы редактирования: создавать/править только у КП «в работе»/
  «На СГП». Обоснование: график — приложение к договору, которого у
  архивного КП нет; светофор для заказа вне производства считается против
  нерелевантной загрузки (ложный зелёный); состав КП до подписания
  нестабилен. Пресейл-прикидку закрывает существующий estimate_production.
  Если КП с графиком вернули в архив или он выполнен — график остаётся
  read-only (протокол обещанных дат).
- Сводный блок «красные партии» на странице производства — фаза 1.5,
  вне этой спеки.
