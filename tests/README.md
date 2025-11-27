# Тесты и утилиты

Все тестовые файлы и утилиты находятся в этой папке.

## Основные тесты

### ✅ test_visualization.py
**Назначение:** Проверка визуализации раскладки плит  
**Что тестирует:** Функцию `build_layout_sequence()` и корректность отображения первичных/вторичных резов  
**Запуск:** `python tests/test_visualization.py`

### ✅ test_order.py  
**Назначение:** Тест оптимизации заказа  
**Что тестирует:** Функцию `optimize_with_cascading_longitudinal_cuts()` на реальных данных  
**Запуск:** `python tests/test_order.py`

### ✅ test_parse.py
**Назначение:** Тест парсинга текста заказа  
**Что тестирует:** Функцию `set_plate_lists_from_text()` - как бот понимает заказы  
**Запуск:** `python tests/test_parse.py`

### ✅ test_kp_generation.py
**Назначение:** Тест генерации коммерческого предложения в PDF  
**Что тестирует:** Функцию `generate_commercial_offer_pdf()` (включая поддержку кириллицы)  
**Запуск:** `python tests/test_kp_generation.py`

## Утилиты для обслуживания

### ⚙️ check_db.py
**Назначение:** Проверка базы данных цен  
**Что делает:** Показывает содержимое БД, ищет цены, проверяет корректность данных  
**Запуск:** `python tests/check_db.py`

### ⚙️ check_excel_and_reload.py
**Назначение:** Проверка Excel файла и перезагрузка БД  
**Что делает:** Читает Excel с ценами из папки "банк знаний" и обновляет базу данных  
**Запуск:** `python tests/check_excel_and_reload.py`

## Экспериментальные скрипты

### 🔬 experiments_compare.py
**Назначение:** Сравнение OLD vs NEW моделей оптимизации  
**Что делает:** Запускает тестовые кейсы с двумя конфигурациями оптимизации и сравнивает результаты  
**Запуск:** `python tests/experiments_compare.py`

## Запуск всех тестов

Для запуска всех основных тестов можно использовать:

```bash
python tests/test_visualization.py
python tests/test_order.py
python tests/test_parse.py
python tests/test_kp_generation.py
```

## Структура

Все тесты автоматически добавляют корень проекта в `sys.path`, поэтому они могут импортировать модули из корня проекта (`visualization`, `optimization`, `config_and_data`, и т.д.).

