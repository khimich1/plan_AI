#!/bin/bash
# Скрипт подготовки окружения: создание venv и установка зависимостей (web path)

set -e
cd "$(dirname "$0")"

# Пауза перед закрытием при любом выходе (чтобы видеть результат)
trap 'read -p "Нажми Enter для закрытия..."' EXIT

echo "=== 1. Создание виртуального окружения venv ==="
python3 -m venv venv || {
    echo ""
    echo "Ошибка: python3-venv не установлен."
    echo "Выполни в терминале: sudo apt install python3.12-venv"
    exit 1
}

echo ""
echo "=== 2. Активация venv и обновление pip ==="
source venv/bin/activate
pip install --upgrade pip

echo ""
echo "=== 3. Установка зависимостей (web: requirements.txt) ==="
pip install -r requirements.txt

echo ""
echo "=== Готово! ==="
echo "Для запуска веб-приложения:"
echo "  source venv/bin/activate"
echo "  uvicorn app.main:app --reload"
echo ""
echo "Telegram-бот DEPRECATED (P5 WP1): см. bot/README.md"
echo "Опционально (архив): pip install -r requirements-bot.txt — только для справки"
