# Таблица стоимости сырья и производственных расходов

## Описание

В базе данных `pb.db` создана таблица `raw_material_costs`, которая хранит данные о стоимости сырья и производственных расходов для каждой плиты.

## Структура таблицы

```sql
CREATE TABLE IF NOT EXISTS raw_material_costs (
    plate_name TEXT PRIMARY KEY,                    -- Название плиты (например, "ПБ 17-12-6")
    raw_material_and_production_cost REAL NOT NULL  -- Стоимость сырья + производственные расходы
);
```

## Импортированные данные

- **Источник**: `банк знаний\Расчет новых цен на ПБ 10.09.2025 (1).xlsx`
- **Лист**: `прайс`
- **Столбец**: `сырье+произ расходы`
- **Количество записей**: 277
- **Диапазон стоимости**: 3807.83 - 25795.91 руб.
- **Средняя стоимость**: 12613.88 руб.

## Использование

### Python API

Модуль `core/raw_material_db.py` предоставляет следующие функции:

```python
from core.raw_material_db import (
    get_raw_material_cost,
    get_all_costs,
    add_or_update_cost,
    get_statistics
)

# Получить стоимость для конкретной плиты
cost = get_raw_material_cost("ПБ 17-12-6")
print(f"Стоимость: {cost:.2f} руб.")

# Получить все данные
all_costs = get_all_costs()
for plate, cost in all_costs.items():
    print(f"{plate}: {cost:.2f} руб.")

# Добавить или обновить стоимость
add_or_update_cost("ПБ 17-12-6", 3807.83)

# Получить статистику
stats = get_statistics()
print(f"Всего записей: {stats['count']}")
print(f"Мин: {stats['min']:.2f}, Макс: {stats['max']:.2f}, Средняя: {stats['avg']:.2f}")
```

### SQL запросы

```sql
-- Получить стоимость для конкретной плиты
SELECT raw_material_and_production_cost 
FROM raw_material_costs 
WHERE plate_name = 'ПБ 17-12-6';

-- Получить все данные
SELECT * FROM raw_material_costs ORDER BY plate_name;

-- Получить статистику
SELECT 
    COUNT(*) as count,
    MIN(raw_material_and_production_cost) as min_cost,
    MAX(raw_material_and_production_cost) as max_cost,
    AVG(raw_material_and_production_cost) as avg_cost
FROM raw_material_costs;

-- Найти плиты с наименьшей стоимостью
SELECT * FROM raw_material_costs 
ORDER BY raw_material_and_production_cost 
LIMIT 10;

-- Найти плиты с наибольшей стоимостью
SELECT * FROM raw_material_costs 
ORDER BY raw_material_and_production_cost DESC 
LIMIT 10;
```

## Повторный импорт данных

Если нужно обновить данные из Excel файла:

```bash
python import_raw_material_costs.py
```

Скрипт автоматически:
1. Найдет лист с прайсом
2. Извлечет данные из столбца "сырье+произ расходы"
3. Обновит данные в базе (INSERT OR REPLACE)

## Примеры данных

| Плита | Стоимость (руб.) |
|-------|------------------|
| ПБ 17-12-6 | 3807.83 |
| ПБ 18-12-6 | 4021.07 |
| ПБ 19-12-6 | 4234.32 |
| ПБ 20-12-6 | 4447.56 |
| ПБ 21-12-6 | 4660.80 |

