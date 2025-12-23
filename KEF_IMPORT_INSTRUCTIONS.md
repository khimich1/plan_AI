# Инструкция по импорту значений КЭФ из Excel

## Что было сделано

1. ✅ Создана таблица `plate_kef_values` в БД для хранения КЭФ по плитам
2. ✅ Обновлена функция `get_kef()` - теперь она ищет КЭФ для конкретной плиты
3. ✅ Обновлена функция `calculate_plate_cost()` - использует КЭФ из БД для каждой плиты
4. ✅ Создана функция `load_kef_from_excel()` для чтения КЭФ из Excel
5. ✅ Создан скрипт `scripts/import_kef_from_excel.py` для импорта

## Как использовать

### Шаг 1: Импорт КЭФ из Excel

Запустите скрипт импорта:

```bash
python scripts/import_kef_from_excel.py
```

Или с указанием конкретного файла:

```bash
python scripts/import_kef_from_excel.py --file "банк знаний/Расчет новых цен на ПБ 10.09.2025 (1).xlsx"
```

Скрипт:
- Ищет колонку "КЭФ" в листах "Себестоимость" или "Прайс"
- Читает значения КЭФ для каждой плиты
- Сохраняет в таблицу `plate_kef_values` в БД

### Шаг 2: Использование в расчёте себестоимости

Теперь функция `calculate_plate_cost()` автоматически использует КЭФ из БД:

```python
from cost_calculation import calculate_plate_cost

result = calculate_plate_cost("ПБ 17-12-6")
print(f"КЭФ: {result['kef']}")  # Будет использован КЭФ из БД для этой плиты
print(f"Полная себестоимость: {result['full_cost_with_kef']:.2f} руб")
```

## Логика работы

1. **При расчёте себестоимости:**
   - Функция `get_kef()` ищет КЭФ для конкретной плиты (length_dm, width_dm, load_code)
   - Если найдено - использует это значение
   - Если не найдено - использует дефолтное значение 1.25 из `cost_constants`

2. **При импорте из Excel:**
   - Скрипт ищет колонку "КЭФ" в листах Excel
   - Для каждой плиты сохраняет значение КЭФ в БД
   - Значения могут быть разными для разных плит (1.22, 1.25, 1.30 и т.д.)

## Структура БД

Таблица `plate_kef_values`:
- `length_dm` - длина плиты в дециметрах
- `width_dm` - ширина плиты в дециметрах  
- `load_code` - код нагрузки
- `kef` - значение КЭФ
- `plate_name` - название плиты (для справки)
- `source_file` - имя файла Excel
- `source_row` - номер строки в Excel

## Примеры

### Проверка импорта

```python
import sqlite3
from cost_calculation.db import DEFAULT_DB

conn = sqlite3.connect(DEFAULT_DB)
cur = conn.cursor()

# Посмотреть все импортированные значения КЭФ
cur.execute("SELECT plate_name, kef FROM plate_kef_values LIMIT 10")
for row in cur.fetchall():
    print(f"{row[0]}: КЭФ = {row[1]}")

conn.close()
```

### Расчёт с использованием КЭФ из БД

```python
from cost_calculation import calculate_plate_cost

# Для плиты ПБ 17-12-6
result = calculate_plate_cost("ПБ 17-12-6")

print(f"Прямые затраты: {result['direct_cost']:.2f} руб")
print(f"КЭФ: {result['kef']}")
print(f"Накладные: {result['overhead_cost']:.2f} руб")
print(f"Полная себестоимость: {result['full_cost_with_kef']:.2f} руб")
```

## Важно

- Если КЭФ не был импортирован для конкретной плиты, используется дефолтное значение 1.25
- Значения КЭФ могут отличаться для разных плит (как в Excel)
- При обновлении Excel файла нужно запустить импорт заново

