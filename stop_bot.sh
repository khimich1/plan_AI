#!/bin/bash

set -u

GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m'

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo -e "${GREEN}========================================${NC}"
echo -e "${GREEN}   Остановка run_local.sh процессов${NC}"
echo -e "${GREEN}========================================${NC}"
echo ""

stopped_any=0

stop_pid_tree() {
    local root_pid="$1"
    local child

    for child in $(pgrep -P "$root_pid" 2>/dev/null); do
        stop_pid_tree "$child"
    done

    if kill -0 "$root_pid" 2>/dev/null; then
        kill "$root_pid" 2>/dev/null || true
        stopped_any=1
    fi
}

stop_by_pattern() {
    local pattern="$1"
    local description="$2"
    local pids

    pids="$(pgrep -f "$pattern" || true)"
    if [ -z "$pids" ]; then
        return
    fi

    echo -e "${YELLOW}🛑 Останавливаю: ${description}${NC}"
    for pid in $pids; do
        stop_pid_tree "$pid"
    done
}

stop_by_port() {
    local port="$1"
    local pids

    pids="$(lsof -ti :"$port" 2>/dev/null || true)"
    if [ -z "$pids" ]; then
        return
    fi

    echo -e "${YELLOW}🛑 Освобождаю порт ${port}${NC}"
    for pid in $pids; do
        stop_pid_tree "$pid"
    done
}

# 1) Останавливаем сам run_local.sh (если запущен)
stop_by_pattern "$SCRIPT_DIR/run_local.sh" "run_local.sh"

# 2) Останавливаем типичные процессы, которые он запускает
stop_by_pattern "uvicorn app.main:app --host 127.0.0.1 --port 8000 --reload" "backend uvicorn"
stop_by_pattern "npm run dev -- --host 127.0.0.1 --port 5173" "frontend vite dev server"

# 3) На случай, если команда запуска отличалась, чистим порты
stop_by_port 8000
stop_by_port 5173

sleep 1

# 4) Принудительно добиваем оставшиеся процессы на нужных портах
for port in 8000 5173; do
    pids="$(lsof -ti :"$port" 2>/dev/null || true)"
    if [ -n "$pids" ]; then
        echo -e "${RED}⚠️  Порт ${port} всё ещё занят. Принудительное завершение...${NC}"
        for pid in $pids; do
            kill -9 "$pid" 2>/dev/null || true
            stopped_any=1
        done
    fi
done

if [ "$stopped_any" -eq 1 ]; then
    echo ""
    echo -e "${GREEN}✅ Процессы, запущенные через run_local.sh, остановлены${NC}"
else
    echo -e "${YELLOW}⚠️  Активных процессов run_local.sh не найдено${NC}"
fi

echo ""
echo -e "${GREEN}========================================${NC}"

