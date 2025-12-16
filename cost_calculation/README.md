# Расчет себестоимости плит

## Описание

Модуль для расчета себестоимости плит ПБ на основе их названия. Все константы и нормы расхода хранятся в базе данных SQLite.

## Структура

### Модули

1. **`cost_calculation/db.py`** - работа с базой данных констант
   - Инициализация схемы БД
   - Хранение констант цен материалов
   - Хранение норм расхода материалов
   - Хранение норм расхода армирования
   - Хранение норм расхода изоформа

2. **`cost_calculation/calculation.py`** - расчет себестоимости
   - Парсинг названия плиты
   - Расчет объема плиты
   - Расчет компонентов себестоимости
   - Полный расчет себестоимости

3. **`cost_calculation/load_from_excel.py`** - загрузка данных из Excel
   - Загрузка объемов плит
   - Загрузка стоимости компонентов
   - Загрузка готовых расчетов

4. **`cost_calculation/tests/test_calculation.py`** - тесты
   - Тесты работы с БД
   - Тесты расчета компонентов
   - Интеграционные тесты

## Использование

### Базовое использование

```python
from cost_calculation import calculate_plate_cost, init_default_constants
import core.config_and_data as cfg

# Инициализация БД (один раз)
init_default_constants(cfg.PRICE_DB_PATH)

# Расчет себестоимости
result = calculate_plate_cost("ПБ 17-12-6", cfg.PRICE_DB_PATH)

if result:
    print(f"Себестоимость: {result['total_cost']:,.2f} руб")
    print(f"Компоненты: {result['components']}")
```

### Формат названия плиты

Формат: `ПБ {длина_дм}-{ширина_дм}-{нагрузка}п`

Примеры:
- `ПБ 17-12-6` - длина 1.7 м, ширина 1.2 м, нагрузка 6п
- `ПБ 20-12-8` - длина 2.0 м, ширина 1.2 м, нагрузка 8п
- `ПБ 30-12-10` - длина 3.0 м, ширина 1.2 м, нагрузка 10п

### Структура результата

```python
{
    'plate_name': 'ПБ 17-12-6',
    'parameters': {
        'length_dm': 17,
        'length_m': 1.7,
        'width_dm': 12,
        'width_m': 1.2,
        'load_code': 6,
        'concrete_grade': 'М400'
    },
    'volume_m3': 0.4488,
    'components': {
        'concrete': 59434.36,      # стоимость бетона
        'reinforcement': 0.0,       # стоимость армирования
        'loops': 286.0,            # стоимость петель
        'izoform': 14.25           # стоимость изоформа
    },
    'total_cost': 59734.61,        # общая себестоимость
    'breakdown': {                 # детальная разбивка
        'concrete': {...},
        'reinforcement': {...}
    }
}
```

## Компоненты себестоимости

### 1. Бетон
- Цемент ПЦ 500 Д0 (кг)
- Песок (м³)
- Щебень (м³)

Расчет: `Стоимость = (Цемент × Цена_цемента) + (Песок × Цена_песка) + (Щебень × Цена_щебня)`

### 2. Армирование
- Проволока Ø 5 Вр II (кг)
- Канат Ø 12 К7 (руб)

Расчет зависит от нагрузки (6п, 8п, 10п, 12п)

### 3. Петли
- Петли д 18 (для стандартных плит)
- Петли д 12, д 14 (для специальных случаев)

### 4. Изоформ
- Изоформ-Б "Экстра" (кг)

Расход зависит от объема плиты

## Обновление констант

### Обновление цен материалов

```python
import sqlite3
from core.config_and_data import PRICE_DB_PATH

conn = sqlite3.connect(PRICE_DB_PATH)
cur = conn.cursor()

# Обновить цену цемента
cur.execute("""
    UPDATE cost_constants 
    SET value = ?, updated_at = CURRENT_TIMESTAMP
    WHERE key = 'cement_price_per_kg'
""", (380.0,))

conn.commit()
conn.close()
```

### Обновление норм расхода

```python
# Обновить нормы бетона М400
cur.execute("""
    UPDATE concrete_norms 
    SET cement_kg_per_m3 = ?, 
        sand_m3_per_m3 = ?,
        gravel_m3_per_m3 = ?,
        updated_at = CURRENT_TIMESTAMP
    WHERE concrete_grade = 'М400'
""", (360.0, 0.62, 2.065))

# Обновить нормы армирования для нагрузки 8п
cur.execute("""
    UPDATE reinforcement_norms 
    SET wire_kg_per_m3 = ?,
        cable_cost_per_m3 = ?,
        updated_at = CURRENT_TIMESTAMP
    WHERE load_code = 8
""", (6.6, 0.0))
```

## Загрузка констант из Excel

Для загрузки реальных констант из файла "Расчет новых цен на ПБ 10.09.2025 (1).xls":

```python
from cost_calculation.load_from_excel import main
import core.config_and_data as cfg

# Загрузить все данные из Excel в БД
main()  # Использует cfg.PRICE_DB_PATH автоматически
```

Или используйте функцию напрямую:

```python
from cost_calculation.load_from_excel import load_volumes_from_excel, load_concrete_costs_from_excel
from cost_calculation.db import init_cost_schema
import core.config_and_data as cfg

db_path = cfg.PRICE_DB_PATH
init_cost_schema(db_path)

# Загрузить конкретные данные
load_volumes_from_excel(db_path)
load_concrete_costs_from_excel(db_path)
```

## Тестирование

Запуск всех тестов:

```bash
python -m unittest cost_calculation.tests.test_calculation -v
```

Запуск конкретного теста:

```bash
python -m unittest cost_calculation.tests.test_calculation.TestCostCalculation.test_calculate_plate_cost_full -v
```

## Примеры

См. файлы в папке `cost_calculation/examples/`:
- `example_basic.py` - базовые примеры использования
- `calculate_interactive.py` - интерактивный калькулятор себестоимости

Запуск примеров:

```bash
python -m cost_calculation.examples.example_basic
python -m cost_calculation.examples.calculate_interactive
```

## Примечания

1. **Константы по умолчанию** - это примерные значения. Для реальных расчетов нужно обновить константы из Excel файла.

2. **Высота плиты** - используется стандартная высота 0.22 м. Если нужно изменить, отредактируйте функцию `calculate_plate_volume()`.

3. **Марка бетона** - автоматически определяется по нагрузке:
   - Нагрузка < 12п → М400
   - Нагрузка ≥ 12п → М500

4. **Нормы изоформа** - зависят от объема плиты. Если объем не попадает в диапазоны, возвращается 0.

## Расширение функциональности

Для добавления новых компонентов себестоимости:

1. Добавьте константу в таблицу `cost_constants`
2. Создайте функцию расчета в `calculation.py`
3. Добавьте компонент в `calculate_plate_cost()`
4. Напишите тесты в `tests/test_calculation.py`

