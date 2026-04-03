#!/usr/bin/env bash
# Первичная подготовка окружения на Linux VPS в корне проекта «Шишов».
# Запуск: из корня репозитория после git clone:
#   chmod +x deploy/bootstrap_vps.sh && ./deploy/bootstrap_vps.sh
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

if ! command -v python3 >/dev/null 2>&1; then
  echo "Ошибка: нужен python3 в PATH"
  exit 1
fi

PYVER="$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')"
MAJOR="$(python3 -c 'import sys; print(sys.version_info.major)')"
MINOR="$(python3 -c 'import sys; print(sys.version_info.minor)')"
if [ "$MAJOR" -lt 3 ] || { [ "$MAJOR" -eq 3 ] && [ "$MINOR" -lt 10 ]; }; then
  echo "Предупреждение: рекомендуется Python 3.10+, сейчас $PYVER"
fi

if [ ! -f "$ROOT/requirements.txt" ]; then
  echo "Ошибка: requirements.txt не найден в $ROOT"
  exit 1
fi

if [ ! -d "$ROOT/.venv" ]; then
  python3 -m venv "$ROOT/.venv"
fi

# shellcheck source=/dev/null
source "$ROOT/.venv/bin/activate"
python -m pip install -U pip
pip install -r "$ROOT/requirements.txt"

echo "Готово: venv в $ROOT/.venv"
echo "Дальше: создайте bot/bot.env с BOT_TOKEN, скопируйте pb.db, plita.db, папку «банк знаний»."
echo "Проверка: source .venv/bin/activate && python scripts/smoke_check.py"
echo "Запуск:  python run_bot.py"
