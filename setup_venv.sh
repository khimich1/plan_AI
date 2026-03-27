#!/bin/bash
# Скрипт подготовки бота к запуску: создание venv и установка зависимостей

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
echo "=== 3. Установка зависимостей из requirements.txt ==="
pip install -r requirements.txt

echo ""
echo "=== Готово! ==="
echo "Для запуска бота:"
echo "  source venv/bin/activate"
echo "  python run_bot.py"
echo ""
echo "Или используй: ./start_bot.sh"
