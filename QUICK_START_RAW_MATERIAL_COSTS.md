# Быстрый старт: Таблица стоимости сырья и производственных расходов

## Что было сделано

✅ Создана таблица `raw_material_costs` в базе данных `pb.db`  
✅ Импортировано **277 записей** из Excel файла  
✅ Создан Python модуль для работы с данными  
✅ Добавлены утилиты для просмотра и импорта данных

## Структура БД

```sql
CREATE TABLE raw_material_costs (
    plate_name TEXT PRIMARY KEY,                    -- "ПБ 17-12-6"
    raw_material_and_production_cost REAL NOT NULL  -- 3807.83
);
```

## Быстрое использование

### 1. Просмотр всех данных

```bash
python view_raw_material_costs.py
```

Выведет статистику и первые 20 записей.

### 2. Использование в коде

```python
from core.raw_material_db import get_raw_material_cost

# Получить стоимость для плиты
cost = get_raw_material_cost("ПБ 56-12-8")
print(f"Стоимость: {cost:.2f} руб.")  # 12890.09 руб.
```

### 3. Обновление данных из Excel

```bash
python import_raw_material_costs.py
```

## Примеры данных

| Плита | Стоимость (руб.) |
|-------|------------------|
| ПБ 17-12-6 | 3807.83 |
| ПБ 56-12-8 | 12890.09 |
| ПБ 72-12-12,5 | 17122.22 |

**Всего**: 277 записей  
**Диапазон**: 3807.83 - 25795.91 руб.  
**Средняя**: 12613.88 руб.

## Файлы

- `core/raw_material_db.py` - Python API для работы с таблицей
- `import_raw_material_costs.py` - Импорт данных из Excel
- `view_raw_material_costs.py` - Просмотр данных
- `RAW_MATERIAL_COSTS_README.md` - Полная документация

## SQL запросы

```sql
-- Получить стоимость
SELECT raw_material_and_production_cost 
FROM raw_material_costs 
WHERE plate_name = 'ПБ 56-12-8';

-- Топ-10 самых дорогих
SELECT * FROM raw_material_costs 
ORDER BY raw_material_and_production_cost DESC 
LIMIT 10;
```

