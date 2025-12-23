# Быстрый старт: Модуль заводской себестоимости

## 1. Импорт данных (один раз)

```bash
python scripts/import_factory_costs.py
```

**Результат:**
- ✅ Импортировано 351 записей
- 📁 Данные в `pb.db` → таблицы `factory_plate_costs`, `factory_plate_cost_components`

## 2. Использование в коде

### Простой способ: по параметрам плиты

```python
from factory_cost import get_cost_by_params

# Нагрузка определяется автоматически!
cost = get_cost_by_params(length_m=7.1, width_m=1.2)

if cost:
    print(f"Плита: {cost['plate_name']}")
    print(f"Себестоимость: {cost['full_cost_with_kef']:.2f} руб")
else:
    print("Плита не найдена")
```

### Продвинутый способ: по названию

```python
from factory_cost import get_cost_by_plate_name

cost = get_cost_by_plate_name("Плиты ПБ 71-12-10п")

if cost:
    print(f"Себестоимость: {cost['full_cost_with_kef']:.2f} руб")
    print(f"Компоненты:")
    print(f"  Армирование: {cost['components']['reinforcement']:.2f} руб")
    print(f"  Бетон: {cost['components']['concrete']:.2f} руб")
    print(f"  Петли: {cost['components']['loops']:.2f} руб")
    print(f"  Изоформ: {cost['components']['izoform']:.2f} руб")
```

### Детальная разбивка затрат

```python
from factory_cost.cost_engine import get_cost_breakdown

breakdown = get_cost_breakdown("Плиты ПБ 71-12-10п")

if breakdown:
    print(f"КЭФ: {breakdown['kef']}")
    print(f"Прямые затраты: {breakdown['direct_cost']:.2f} руб")
    print(f"Накладные: {breakdown['overhead_cost']:.2f} руб")
    print(f"ИТОГО: {breakdown['full_cost_with_kef']:.2f} руб")
    
    print("\nСтруктура затрат:")
    for comp, data in breakdown['breakdown'].items():
        print(f"  {comp}: {data['percentage']:.1f}%")
```

## 3. Проверка данных

```bash
# Быстрая валидация
python scripts/validate_factory_costs.py

# Подробный отчёт
python scripts/validate_factory_costs.py --detailed
```

## 4. Обновление данных

Когда появился новый Excel с ценами:

```bash
python scripts/import_factory_costs.py --file "path/to/new_file.xlsx"
```

Готово! Старые данные автоматически заменятся новыми.

## Важные моменты

✅ **ДА:**
- Использовать для расчёта заводской себестоимости
- Использовать совместно с существующим парсером плит
- Обновлять при изменении цен на материалы

❌ **НЕТ:**
- **НЕ** использовать для КП (есть отдельная таблица `prices`)
- **НЕ** писать свой парсер размеров/нагрузок
- **НЕ** считать маржу и цену продажи

## Что дальше?

Читайте полную документацию: `factory_cost/README.md`

