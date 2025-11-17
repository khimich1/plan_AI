#!/bin/bash

# Скрипт для остановки Telegram-бота

# Цвета для красивого вывода
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${GREEN}========================================${NC}"
echo -e "${GREEN}   Остановка Telegram-бота${NC}"
echo -e "${GREEN}========================================${NC}"
echo ""

# Ищем процесс бота
BOT_PID=$(ps aux | grep "[p]ython.*bot_main.py" | awk '{print $2}')

if [ -z "$BOT_PID" ]; then
    echo -e "${YELLOW}⚠️  Бот не запущен${NC}"
else
    echo -e "${YELLOW}🔍 Найден процесс бота (PID: $BOT_PID)${NC}"
    echo -e "${YELLOW}🛑 Останавливаю бота...${NC}"
    kill $BOT_PID
    
    # Ждём немного
    sleep 2
    
    # Проверяем, остановился ли процесс
    if ps -p $BOT_PID > /dev/null 2>&1; then
        echo -e "${RED}❌ Бот не остановился, принудительно завершаю...${NC}"
        kill -9 $BOT_PID
        sleep 1
    fi
    
    # Проверяем ещё раз
    if ps -p $BOT_PID > /dev/null 2>&1; then
        echo -e "${RED}❌ Не удалось остановить бота${NC}"
    else
        echo -e "${GREEN}✅ Бот успешно остановлен${NC}"
    fi
fi

echo ""
echo -e "${GREEN}========================================${NC}"
echo ""
echo "Нажмите Enter для закрытия окна..."
read

