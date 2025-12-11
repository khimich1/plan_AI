# 📋 Шпаргалка по командам

> **Для новичков**: Копируй и вставляй команды в терминал! 🚀

---

## 🔧 Установка и настройка

### 1. Установка зависимостей

```bash
# Установка ВСЕХ библиотек сразу
pip install -r requirements.txt

# Или по отдельности:
pip install aiogram==3.6.0        # Telegram Bot API
pip install easyocr>=1.7.0        # 🤖 Бесплатный OCR
pip install openai>=1.0.0         # 🧠 GPT-4o Vision
```

### 2. Проверка установки

```bash
# Проверь, что библиотеки установлены
pip list | findstr "aiogram easyocr openai"

# Должно показать:
# aiogram         3.6.0
# easyocr         1.7.0
# openai          1.x.x
```

---

## 🤖 Запуск бота

### Основные команды:

```bash
# Запуск бота (обычный режим)
python run_bot.py

# Запуск с отображением всех логов
python run_bot.py --debug

# Остановка бота
Ctrl+C
```

---

## 🧪 Тестирование OCR

### Проверка распознавания фото:

```bash
# Тест на конкретном фото
python test_ocr_comparison.py "банк знаний/1.jpeg"

# Тест на другом фото
python test_ocr_comparison.py "банк знаний/photo_2025-07-10_10-06-26.jpg"
```

**Что показывает тест:**
- ✅ Какие методы доступны (EasyOCR, GPT-4o)
- 📊 Результаты распознавания каждым методом
- ⚡ Сравнение точности и скорости
- 💰 Стоимость распознавания

---

## 🔑 Настройка API ключей

### Добавление OPENAI_API_KEY:

**Вариант 1: Через файл bot.env (рекомендуется)**

```bash
# Открой файл в блокноте
notepad bot\bot.env

# Добавь строку (убери #):
OPENAI_API_KEY=sk-твой_ключ_здесь

# Сохрани (Ctrl+S) и закрой
```

**Вариант 2: Через переменные среды Windows**

```bash
# Установка переменной (временно, до перезагрузки)
set OPENAI_API_KEY=sk-твой_ключ_здесь

# Проверка
echo %OPENAI_API_KEY%
```

**Вариант 3: Постоянная переменная среды**

```bash
# Открой настройки системы
sysdm.cpl

# Перейди: Дополнительно → Переменные среды
# Создай новую: OPENAI_API_KEY = sk-...
# Перезапусти терминал!
```

---

## 📊 Мониторинг и логи

### Просмотр логов бота:

```bash
# Запуск с подробными логами
python run_bot.py --verbose

# Сохранение логов в файл
python run_bot.py > logs.txt 2>&1

# Просмотр последних 20 строк лога
type logs.txt | more
```

### Проверка стоимости GPT:

```bash
# В браузере открой:
start https://platform.openai.com/settings/organization/billing/overview

# Или вручную зайди на сайт OpenAI → Settings → Billing
```

---

## 🗂️ Работа с файлами проекта

### Просмотр структуры:

```bash
# Список файлов в корне
dir

# Список файлов в папке core
dir core

# Список всех Python файлов
dir /s *.py
```

### Чтение файлов:

```bash
# Открыть в блокноте
notepad core\ocr_gpt.py

# Открыть в VSCode
code core\ocr_gpt.py

# Вывести содержимое в терминал
type core\ocr_gpt.py
```

---

## 🔄 Обновление проекта

### Обновление зависимостей:

```bash
# Обновление всех библиотек
pip install --upgrade -r requirements.txt

# Обновление конкретной библиотеки
pip install --upgrade openai

# Проверка версий
pip show openai
pip show easyocr
```

---

## 🧹 Очистка и обслуживание

### Очистка временных файлов:

```bash
# Удаление временных фото
del /q bot\tmp\*.jpg
del /q bot\tmp\*.png

# Удаление кэша Python
rd /s /q __pycache__
rd /s /q core\__pycache__
rd /s /q bot\__pycache__

# Удаление старых логов
del /q logs.txt
```

### Сброс базы данных (ОСТОРОЖНО!):

```bash
# Создание резервной копии
copy pb.db pb_backup.db

# Удаление БД (все данные удалятся!)
del pb.db

# Восстановление из резервной копии
copy pb_backup.db pb.db
```

---

## 🐛 Решение проблем

### Проблема: "ModuleNotFoundError: No module named 'openai'"

```bash
# Решение: Установи библиотеку
pip install openai
```

### Проблема: "OPENAI_API_KEY не найден"

```bash
# Решение 1: Проверь bot.env
notepad bot\bot.env
# Убедись, что строка OPENAI_API_KEY не закомментирована (#)

# Решение 2: Установи через переменные среды
set OPENAI_API_KEY=sk-твой_ключ
```

### Проблема: "Insufficient credits" (недостаточно денег)

```bash
# Решение: Пополни баланс OpenAI
start https://platform.openai.com/settings/organization/billing/overview
```

### Проблема: EasyOCR долго загружается первый раз

```bash
# Это нормально! Первый запуск скачивает модели (~100 МБ)
# Подожди 2-3 минуты
# Следующие запуски будут быстрыми
```

---

## 📦 Создание резервной копии

### Архивация проекта:

```bash
# Создание ZIP архива (через PowerShell)
powershell Compress-Archive -Path . -DestinationPath backup.zip

# Или вручную:
# 1. Выдели все файлы (Ctrl+A)
# 2. ПКМ → Отправить → Сжатая ZIP-папка
```

---

## 🚀 Полезные команды для разработки

### Быстрый запуск всех тестов:

```bash
# Запуск конкретного теста
python test_ocr_comparison.py "банк знаний/1.jpeg"

# Запуск всех тестов (если есть)
python -m pytest tests/

# Запуск с отладкой
python -m pdb test_ocr_comparison.py "банк знаний/1.jpeg"
```

### Проверка кода на ошибки:

```bash
# Проверка синтаксиса
python -m py_compile core\ocr_gpt.py

# Проверка стиля кода (если установлен flake8)
pip install flake8
flake8 core\ocr_gpt.py
```

---

## 📖 Дополнительные ресурсы

### Открыть документацию:

```bash
# Инструкция по настройке GPT
notepad SETUP_GPT_OCR.md

# Описание изменений
notepad "описание изменений\2025-11-27_GPT_OCR.md"

# Краткое резюме
type GPT_OCR_READY.txt
```

### Полезные ссылки:

```bash
# OpenAI API Keys
start https://platform.openai.com/api-keys

# OpenAI Billing
start https://platform.openai.com/settings/organization/billing/overview

# Документация aiogram
start https://docs.aiogram.dev/

# Документация EasyOCR
start https://github.com/JaidedAI/EasyOCR
```

---

## 💡 Советы для новичков

### Как скопировать команду:

1. **Выдели текст** мышкой
2. Нажми **Ctrl+C** (или ПКМ → Копировать)
3. Вставь в терминал **Ctrl+V** (или ПКМ → Вставить)
4. Нажми **Enter**

### Как открыть терминал в VSCode:

1. Нажми **Ctrl+`** (русская буква Ё)
2. Или: Меню → Вид → Терминал
3. Или: Ctrl+Shift+P → "Terminal: Create New Terminal"

### Как узнать текущую папку:

```bash
# Windows (CMD)
cd

# Windows (PowerShell)
pwd

# Результат должен быть:
# C:\Users\Роман\Desktop\Шишов
```

---

## ✅ Чеклист готовности

Проверь, что всё установлено:

```bash
# 1. Python установлен?
python --version
# Должно быть: Python 3.10+ или выше

# 2. Библиотеки установлены?
pip list | findstr "aiogram easyocr openai"
# Должны быть все три

# 3. API ключ добавлен?
type bot\bot.env | findstr "OPENAI_API_KEY"
# Должна быть строка без # в начале

# 4. Бот запускается?
python run_bot.py
# Должно появиться: "Бот запущен!"
```

---

**Готово! Теперь ты можешь работать с ботом как профи! 🎉**

*Если что-то непонятно — спрашивай! 😊*

