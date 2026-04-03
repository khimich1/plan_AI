#!/usr/bin/env bash
# Резервная копия данных бота (SQLite и банк знаний).
# Запуск из корня проекта:
#   chmod +x scripts/backup_shishov_data.sh
#   ./scripts/backup_shishov_data.sh
# Переменная BACKUP_DIR — каталог для архивов (по умолчанию ./backups).
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
BACKUP_DIR="${BACKUP_DIR:-$ROOT/backups}"
STAMP="$(date +%Y%m%d_%H%M%S)"
NAME="shishov_backup_${STAMP}.tar.gz"
mkdir -p "$BACKUP_DIR"

ITEMS=()
[ -f "$ROOT/pb.db" ] && ITEMS+=("pb.db")
[ -f "$ROOT/plita.db" ] && ITEMS+=("plita.db")
[ -d "$ROOT/банк знаний" ] && ITEMS+=("банк знаний")

if [ "${#ITEMS[@]}" -eq 0 ]; then
  echo "Нечего архивировать: нет pb.db, plita.db и папки «банк знаний»."
  exit 1
fi

tar -czf "$BACKUP_DIR/$NAME" -C "$ROOT" "${ITEMS[@]}"
echo "Создан: $BACKUP_DIR/$NAME"
