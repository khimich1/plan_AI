#!/bin/bash

# Скрипт для быстрого запуска Telegram-бота
# Автор: Автоматически создан для удобного запуска

# Переходим в папку проекта
cd "$(dirname "$0")"

# Цвета для красивого вывода (опционально)
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${GREEN}========================================${NC}"
echo -e "${GREEN}   Запуск Telegram-бота${NC}"
echo -e "${GREEN}========================================${NC}"
echo ""

# Создаём папку для логов, если её нет
LOGS_DIR="logs"
mkdir -p "$LOGS_DIR"
echo -e "${YELLOW}📁 Папка для логов: $LOGS_DIR${NC}"

# Создаём имя файла лога с датой и временем
LOG_FILE="$LOGS_DIR/bot_$(date +%Y%m%d_%H%M%S).log"
echo -e "${YELLOW}📝 Лог-файл: $LOG_FILE${NC}"
echo ""

# Проверяем наличие виртуального окружения
if [ ! -d "venv" ]; then
    echo -e "${RED}❌ Виртуальное окружение не найдено!${NC}"
    echo -e "${YELLOW}💡 Создаю виртуальное окружение...${NC}"
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo -e "${RED}❌ Ошибка создания виртуального окружения!${NC}"
        echo ""
        echo "Нажмите Enter для выхода..."
        read
        exit 1
    fi
    echo -e "${GREEN}✅ Виртуальное окружение создано${NC}"
    echo ""
    echo -e "${YELLOW}💡 Устанавливаю зависимости...${NC}"
    source venv/bin/activate
    pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo -e "${RED}❌ Ошибка установки зависимостей!${NC}"
        echo ""
        echo "Нажмите Enter для выхода..."
        read
        exit 1
    fi
    echo -e "${GREEN}✅ Зависимости установлены${NC}"
    echo ""
fi

# Активируем виртуальное окружение
echo -e "${YELLOW}🔧 Активирую виртуальное окружение...${NC}"
source venv/bin/activate

# Проверяем наличие файла bot_main.py
if [ ! -f "bot_main.py" ]; then
    echo -e "${RED}❌ Файл bot_main.py не найден!${NC}"
    echo ""
    echo "Нажмите Enter для выхода..."
    read
    exit 1
fi

# Проверяем наличие файла bot.env
if [ ! -f "bot.env" ]; then
    echo -e "${RED}❌ Файл bot.env не найден!${NC}"
    echo -e "${YELLOW}💡 Создайте файл bot.env с токеном бота${NC}"
    echo ""
    echo "Нажмите Enter для выхода..."
    read
    exit 1
fi

echo -e "${GREEN}🚀 Запускаю бота...${NC}"
echo -e "${YELLOW}💡 Все сообщения сохраняются в: $LOG_FILE${NC}"
echo -e "${YELLOW}💡 Для остановки нажмите Ctrl+C${NC}"
echo ""
echo -e "${GREEN}========================================${NC}"
echo ""

# Запускаем бота и перенаправляем вывод в лог-файл
# Также выводим в консоль для удобства
python bot_main.py 2>&1 | tee "$LOG_FILE"

# Сохраняем код выхода
EXIT_CODE=${PIPESTATUS[0]}

echo ""
echo -e "${GREEN}========================================${NC}"

# Проверяем код выхода
if [ $EXIT_CODE -eq 0 ]; then
    echo -e "${GREEN}✅ Бот завершил работу нормально${NC}"
elif [ $EXIT_CODE -eq 130 ]; then
    echo -e "${YELLOW}⚠️  Бот остановлен пользователем (Ctrl+C)${NC}"
else
    echo -e "${RED}❌ Бот завершился с ошибкой (код: $EXIT_CODE)${NC}"
    echo -e "${YELLOW}💡 Проверьте лог-файл: $LOG_FILE${NC}"
fi

echo -e "${GREEN}========================================${NC}"
echo ""
echo "Нажмите Enter для закрытия окна..."
read

