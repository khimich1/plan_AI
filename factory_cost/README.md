# Модуль расчёта заводской себестоимости плит ПБ

## Описание

Модуль `factory_cost` предназначен для расчёта **заводской себестоимости** производства плит ПБ на основании данных из Excel.

⚠️ **КРИТИЧЕСКИ ВАЖНО:**
- Модуль **ИЗОЛИРОВАН** от КП и таблицы `prices`
- **НЕ** считает цену продажи и маржу
- Использует **СУЩЕСТВУЮЩИЙ** парсер из `core.config_and_data.py`
- **НЕ** дублирует логику определения размеров и нагрузок

## Структура модуля

```
factory_cost/
├── __init__.py              # Экспорт API
├── db_schema.py             # Схема БД (factory_plate_costs, factory_plate_cost_components)
├── excel_reader.py          # Чтение данных из Excel
├── import_from_xlsx.py      # Импорт себестоимости в БД
├── cost_engine.py           # API для получения себестоимости
└── README.md               # Эта документация
```

## База данных

### Таблица `factory_plate_costs`

Хранит полную себестоимость плиты.

| Поле | Тип | Описание |
|------|-----|----------|
| `plate_name` | TEXT | Название плиты (например, "Плиты ПБ 71-12-10п") |
| `length_dm` | INTEGER | Длина в дециметрах (71 = 7.1м) |
| `width_dm` | INTEGER | Ширина в дециметрах (12 = 1.2м) |
| `load_code` | REAL | Нагрузка (8, 10, 12, 12.5) |
| `direct_cost` | REAL | Прямые затраты (руб) |
| `overhead_cost` | REAL | Накладные расходы (руб) |
| `full_cost` | REAL | Полная себестоимость без КЭФ (руб) |
| `kef` | REAL | КЭФ (коэффициент общих затрат) |
| `full_cost_with_kef` | REAL | Полная себестоимость с КЭФ (руб) |
| `volume_m3` | REAL | Объём бетона (м³) |
| `concrete_grade` | TEXT | Марка бетона (М400, М500, В30, В40) |
| `quality_flag` | TEXT | Флаг проблем ('components_mismatch' если сумма компонентов не сходится) |
| `source_file` | TEXT | Имя Excel-файла источника |
| `source_sheet` | TEXT | Название листа в Excel |
| `source_row` | INTEGER | Номер строки в Excel |
| `updated_at` | TEXT | Дата обновления |

**Primary Key:** `(length_dm, width_dm, load_code)`

### Таблица `factory_plate_cost_components`

Детализация себестоимости по компонентам.

| Поле | Тип | Описание |
|------|-----|----------|
| `plate_name` | TEXT | Название плиты |
| `component` | TEXT | Компонент ('reinforcement', 'concrete', 'loops', 'izoform') |
| `value` | REAL | Стоимость компонента (руб) |

**Primary Key:** `(plate_name, component)`

## Использование

### 1. Импорт данных из Excel

```bash
# Импорт из дефолтного файла
python scripts/import_factory_costs.py

# Импорт из конкретного файла
python scripts/import_factory_costs.py --file "path/to/file.xlsx"

# Добавление без очистки старых данных
python scripts/import_factory_costs.py --no-clear
```

### 2. Валидация данных

```bash
# Быстрая проверка
python scripts/validate_factory_costs.py

# Подробный отчёт
python scripts/validate_factory_costs.py --detailed
```

### 3. API в Python-коде

```python
from factory_cost import get_cost_by_plate_name, get_cost_by_params

# Поиск по названию плиты
cost = get_cost_by_plate_name("Плиты ПБ 71-12-10п")
if cost:
    print(f"Себестоимость: {cost['full_cost_with_kef']:.2f} руб")
    print(f"Прямые затраты: {cost['direct_cost']:.2f} руб")
    print(f"Компоненты: {cost['components']}")

# Поиск по параметрам (длина, ширина)
# Нагрузка определяется автоматически через get_load_code_for_plate()
cost = get_cost_by_params(length_m=7.1, width_m=1.2)
if cost:
    print(f"Плита: {cost['plate_name']}")
    print(f"Себестоимость: {cost['full_cost_with_kef']:.2f} руб")
else:
    print("Плита не найдена в базе себестоимости")
```

### 4. Детальная разбивка

```python
from factory_cost.cost_engine import get_cost_breakdown

breakdown = get_cost_breakdown("Плиты ПБ 71-12-10п")
if breakdown:
    print(f"Прямые затраты: {breakdown['direct_cost']:.2f} руб")
    print(f"КЭФ: {breakdown['kef']}")
    print(f"Полная себестоимость: {breakdown['full_cost_with_kef']:.2f} руб")
    print("\nСтруктура затрат:")
    for comp, data in breakdown['breakdown'].items():
        print(f"  {comp}: {data['value']:.2f} руб ({data['percentage']:.1f}%)")
```

## Формат Excel

### Лист "Стоимость" (обязательный)

Содержит прямые затраты на материалы.

**Необходимые колонки:**
- Наименование ЖБИ (например, "ПБ 71-12-10п")
- Длина, дм
- Бетон | марка
- Бетон | объем, м³
- Армирование | Стоимость | руб
- Бетон | Стоимость | руб
- Петли | Стоимость | руб
- Изоформ-Б "Экстра" | руб
- Итого стоимость | руб

### Лист "Себестоимость" (опциональный)

Содержит КЭФ - коэффициент общих затрат.

Модуль ищет на листе ячейку с текстом "КЭФ" и значение рядом с ней.

## Интеграция с существующим парсером

Модуль **ОБЯЗАТЕЛЬНО** использует функции из `core.config_and_data.py`:

```python
from core.config_and_data import (
    parse_load_code_from_name,    # Извлечение нагрузки из названия
    get_load_code_for_plate,      # Определение нагрузки по размерам
    make_plate_name               # Создание стандартного названия
)
```

### Логика определения нагрузки

1. **При импорте из Excel:** нагрузка извлекается из названия плиты через `parse_load_code_from_name()`
2. **При запросе по параметрам:** нагрузка определяется через `get_load_code_for_plate(length_m, width_m)`
   - Использует данные из заказа (если плита была в PLATE_LOAD_MAP)
   - Иначе: 6п для узких плит (<1.0м), 8п для широких

### Пример работы парсера

```python
# Парсинг названия
>>> parse_load_code_from_name("ПБ 71-12-10п")
10

>>> parse_load_code_from_name("ПБ 69-12-12,5п")
13  # 12.5 округляется до 13

# Определение нагрузки по размерам
>>> get_load_code_for_plate(7.1, 1.2)
8  # Широкая плита, дефолт 8п

>>> get_load_code_for_plate(6.3, 0.46)
6  # Узкая плита, дефолт 6п
```

## Валидации

Модуль автоматически проверяет:

1. **Сумма компонентов ≈ прямые затраты**
   - Допуск: max(50 руб, 2%)
   - Если не сходится → `quality_flag = 'components_mismatch'`

2. **Адекватность значений:**
   - Нет отрицательных значений
   - Себестоимость < 100 000 руб
   - Объём < 2 м³
   - Размеры в разумных пределах

3. **Дубликаты:** Не должно быть двух плит с одинаковыми (длина, ширина, нагрузка)

4. **Наличие компонентов:** У каждой плиты должны быть компоненты

## Тестирование

```bash
# Запуск всех тестов
python tests/test_factory_cost.py

# Тесты проверяют:
# - Инициализацию схемы БД
# - Чтение Excel
# - Импорт данных
# - API (поиск по названию и параметрам)
# - Интеграцию с существующим парсером
```

## Часто задаваемые вопросы

### Почему все плиты имеют `quality_flag = 'components_mismatch'`?

Это нормально. В Excel есть дополнительные статьи затрат (кроме армирования, бетона, петель и изоформа), которые не учитываются в 4 основных компонентах. Расхождение показывает, сколько приходится на прочие затраты.

### Как обновить себестоимость?

Просто запустите импорт заново:

```bash
python scripts/import_factory_costs.py
```

По умолчанию старые данные будут удалены и заменены новыми.

### Что делать, если плита не найдена?

Проверьте:
1. Импортирован ли файл с себестоимостью?
2. Есть ли плита с такой длиной и шириной в Excel?
3. Совпадает ли нагрузка? (можно посмотреть через `get_all_available_plates()`)

### Можно ли добавить свой парсер размеров?

**НЕТ!** Это явно запрещено в ТЗ. Модуль должен использовать существующий парсер из `core.config_and_data.py`.

### Как связать с КП?

Модуль **НЕ** предназначен для КП. Это две независимые системы:
- `factory_cost` → заводская себестоимость
- `prices` → цены для клиентов

Если нужно использовать себестоимость в КП, создайте отдельный модуль-адаптер.

## Примеры из практики

### Пример 1: Расчёт себестоимости для плиты из заказа

```python
from core.config_and_data import set_plate_lists_from_text
from factory_cost import get_cost_by_params

# Парсим заказ
order_text = """
Плиты ПБ 71-12-10п - 5 шт
Плиты ПБ 63-5,3-8п - 10 шт
"""

set_plate_lists_from_text(order_text)

# Получаем себестоимость
cost1 = get_cost_by_params(7.1, 1.2)  # Автоматически определит нагрузку 10п из заказа
cost2 = get_cost_by_params(6.3, 0.53)  # Нагрузка 8п из заказа

if cost1:
    print(f"ПБ 71-12: {cost1['full_cost_with_kef']:.2f} руб × 5 = {cost1['full_cost_with_kef']*5:.2f} руб")

if cost2:
    print(f"ПБ 63-5,3: {cost2['full_cost_with_kef']:.2f} руб × 10 = {cost2['full_cost_with_kef']*10:.2f} руб")
```

### Пример 2: Анализ структуры затрат

```python
from factory_cost.cost_engine import get_all_available_plates, get_cost_breakdown

# Получаем все плиты
plates = get_all_available_plates()

# Анализируем плиты с нагрузкой 12.5п
for plate in plates:
    if plate['load_code'] == 12.5:
        breakdown = get_cost_breakdown(plate['plate_name'])
        reinforcement_pct = breakdown['breakdown']['reinforcement']['percentage']
        
        print(f"{plate['plate_name']}: армирование {reinforcement_pct}% от себестоимости")
```

## Лицензия

Модуль является частью проекта расчёта ПБ ЖБК СТАРТ.

