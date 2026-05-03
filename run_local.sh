#!/bin/bash

set -u

cd "$(dirname "$0")"

GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m'

LOGS_DIR="logs"
mkdir -p "$LOGS_DIR"
TIMESTAMP="$(date +%Y%m%d_%H%M%S)"
BACKEND_LOG="$LOGS_DIR/backend_${TIMESTAMP}.log"
FRONTEND_LOG="$LOGS_DIR/frontend_${TIMESTAMP}.log"

BACKEND_PID=""
FRONTEND_PID=""
HAS_PAUSED_ON_EXIT=0
PYTHON_CMD=""
VENV_ACTIVATE_PATH=""

pause_before_exit() {
    # Для запуска из GUI/двойным кликом: окно не закроется мгновенно.
    # Если запуск не интерактивный (например, CI), паузу пропускаем.
    if [ "${HAS_PAUSED_ON_EXIT}" -eq 0 ] && [ -t 0 ] && [ -z "${CI:-}" ]; then
        echo ""
        echo "Нажмите Enter для закрытия окна..."
        read -r
        HAS_PAUSED_ON_EXIT=1
    fi
}

cleanup() {
    echo ""
    echo -e "${YELLOW}Останавливаю backend и frontend...${NC}"

    if [ -n "$BACKEND_PID" ] && kill -0 "$BACKEND_PID" 2>/dev/null; then
        kill "$BACKEND_PID" 2>/dev/null
    fi
    if [ -n "$FRONTEND_PID" ] && kill -0 "$FRONTEND_PID" 2>/dev/null; then
        kill "$FRONTEND_PID" 2>/dev/null
    fi

    wait "$BACKEND_PID" 2>/dev/null || true
    wait "$FRONTEND_PID" 2>/dev/null || true

    pause_before_exit
}

trap cleanup EXIT INT TERM

echo -e "${GREEN}========================================${NC}"
echo -e "${GREEN}   Запуск backend + frontend${NC}"
echo -e "${GREEN}========================================${NC}"
echo -e "${YELLOW}Backend лог:  ${BACKEND_LOG}${NC}"
echo -e "${YELLOW}Frontend лог: ${FRONTEND_LOG}${NC}"
echo ""

detect_python_command() {
    if command -v python >/dev/null 2>&1; then
        PYTHON_CMD="python"
        return 0
    fi
    if command -v python3 >/dev/null 2>&1; then
        PYTHON_CMD="python3"
        return 0
    fi
    return 1
}

detect_venv_activate_path() {
    if [ -f "venv/bin/activate" ]; then
        VENV_ACTIVATE_PATH="venv/bin/activate"
        return 0
    fi
    if [ -f "venv/Scripts/activate" ]; then
        VENV_ACTIVATE_PATH="venv/Scripts/activate"
        return 0
    fi
    return 1
}

if ! detect_python_command; then
    echo -e "${RED}❌ Команда python/python3 не найдена. Установи Python 3 и повтори запуск.${NC}"
    exit 1
fi

if [ ! -d "venv" ]; then
    echo -e "${YELLOW}venv не найден, создаю автоматически...${NC}"
    "$PYTHON_CMD" -m venv venv
    if [ $? -ne 0 ]; then
        echo -e "${RED}❌ Не удалось создать venv${NC}"
        exit 1
    fi
    echo -e "${GREEN}✅ venv создан${NC}"
fi

if [ ! -f "requirements.txt" ]; then
    echo -e "${RED}❌ Не найден файл requirements.txt${NC}"
    exit 1
fi

if [ ! -f "frontend/package.json" ]; then
    echo -e "${RED}❌ Не найден frontend/package.json${NC}"
    exit 1
fi

try_load_nvm() {
    export NVM_DIR="${NVM_DIR:-$HOME/.nvm}"
    if [ -s "$NVM_DIR/nvm.sh" ]; then
        # shellcheck disable=SC1090
        . "$NVM_DIR/nvm.sh"
    fi
}

ensure_modern_node() {
    NODE_VERSION_RAW="$(node -v 2>/dev/null || true)"
    NODE_MAJOR="$(echo "$NODE_VERSION_RAW" | sed -E 's/^v([0-9]+).*/\1/')"

    if [ -n "$NODE_MAJOR" ] && [ "$NODE_MAJOR" -ge 20 ]; then
        return 0
    fi

    try_load_nvm
    if command -v nvm >/dev/null 2>&1; then
        echo -e "${YELLOW}Обнаружена старая Node.js (${NODE_VERSION_RAW:-не найдена}), пробую nvm use 22...${NC}"
        if nvm use 22 >/dev/null 2>&1; then
            NODE_VERSION_RAW="$(node -v 2>/dev/null || true)"
            NODE_MAJOR="$(echo "$NODE_VERSION_RAW" | sed -E 's/^v([0-9]+).*/\1/')"
            echo -e "${GREEN}✅ Активирована Node.js ${NODE_VERSION_RAW} через nvm${NC}"
        fi
    fi

    if [ -z "$NODE_MAJOR" ] || [ "$NODE_MAJOR" -lt 20 ]; then
        echo -e "${RED}❌ Слишком старая версия Node.js: ${NODE_VERSION_RAW:-не найдена}${NC}"
        echo -e "${YELLOW}Нужно минимум Node.js 20 (лучше 22).${NC}"
        echo -e "${YELLOW}Самый простой способ через nvm:${NC}"
        echo "curl -fsSL https://raw.githubusercontent.com/nvm-sh/nvm/v0.40.3/install.sh | bash"
        echo "source ~/.bashrc"
        echo "nvm install 22"
        echo "nvm use 22"
        exit 1
    fi
}

if ! command -v npm >/dev/null 2>&1; then
    echo -e "${RED}❌ Команда npm не найдена.${NC}"
    echo -e "${YELLOW}Установи Node.js и npm, затем запусти скрипт снова.${NC}"
    echo -e "${YELLOW}Рекомендуемая версия Node.js: 20+ (лучше 22).${NC}"
    exit 1
fi

if ! command -v node >/dev/null 2>&1; then
    echo -e "${RED}❌ Команда node не найдена.${NC}"
    exit 1
fi

ensure_modern_node

echo -e "${YELLOW}Активирую venv и проверяю зависимости backend...${NC}"
if ! detect_venv_activate_path; then
    echo -e "${RED}❌ Не найден activate-скрипт виртуального окружения (venv/bin/activate или venv/Scripts/activate).${NC}"
    exit 1
fi
source "$VENV_ACTIVATE_PATH"
python -m pip install -r requirements.txt >/dev/null

echo -e "${YELLOW}Проверяю зависимости frontend...${NC}"
if [ ! -d "frontend/node_modules" ]; then
    (cd frontend && npm install --include=optional)
else
    # Иногда npm пропускает optional native-пакеты (rolldown binding),
    # поэтому дополнительно дотягиваем их перед запуском.
    (cd frontend && npm install --include=optional --no-audit --no-fund >/dev/null)
fi

echo -e "${YELLOW}Запускаю backend: http://127.0.0.1:8000${NC}"
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000 --reload >"$BACKEND_LOG" 2>&1 &
BACKEND_PID=$!
sleep 1
if ! kill -0 "$BACKEND_PID" 2>/dev/null; then
    echo -e "${RED}❌ Backend не запустился. Смотри лог: $BACKEND_LOG${NC}"
    exit 1
fi

echo -e "${YELLOW}Запускаю frontend: http://127.0.0.1:5173${NC}"
(cd frontend && npm run dev -- --host 127.0.0.1 --port 5173) >"$FRONTEND_LOG" 2>&1 &
FRONTEND_PID=$!
sleep 1
if ! kill -0 "$FRONTEND_PID" 2>/dev/null; then
    echo -e "${RED}❌ Frontend не запустился. Смотри лог: $FRONTEND_LOG${NC}"
    exit 1
fi

echo ""
echo -e "${GREEN}✅ Оба процесса запущены${NC}"
echo -e "${YELLOW}Открывай в браузере: http://127.0.0.1:5173${NC}"
echo -e "${YELLOW}Открывай в браузере (localhost): http://localhost:5173${NC}"
echo -e "${YELLOW}API backend: http://127.0.0.1:8000${NC}"
echo -e "${YELLOW}API backend (localhost): http://localhost:8000${NC}"
echo -e "${YELLOW}Вход в систему: http://127.0.0.1:8000/web/login${NC}"
echo -e "${YELLOW}Вход в систему (localhost): http://localhost:8000/web/login${NC}"
echo -e "${YELLOW}Остановка: Ctrl+C${NC}"
echo ""

while true; do
    if ! kill -0 "$BACKEND_PID" 2>/dev/null; then
        echo -e "${RED}❌ Backend остановился. Смотри лог: $BACKEND_LOG${NC}"
        exit 1
    fi
    if ! kill -0 "$FRONTEND_PID" 2>/dev/null; then
        echo -e "${RED}❌ Frontend остановился. Смотри лог: $FRONTEND_LOG${NC}"
        exit 1
    fi
    sleep 1
done

