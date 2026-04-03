#!/usr/bin/env bash
# Снять зависший apt после Ctrl+Z или обрыва сети (блокировки dpkg/apt).
# Запуск на VPS от root:  bash deploy/unstick_apt.sh
set -euo pipefail

if [[ "${EUID:-0}" -ne 0 ]]; then
  echo "Запустите от root: sudo bash $0"
  exit 1
fi

echo "Завершаю процессы apt/apt-get (если есть)..."
killall apt apt-get 2>/dev/null || true
sleep 2

echo "Удаляю lock-файлы (если остались после сбоя)..."
rm -f /var/lib/apt/lists/lock
rm -f /var/cache/apt/archives/lock
rm -f /var/lib/dpkg/lock
rm -f /var/lib/dpkg/lock-frontend

echo "dpkg --configure -a ..."
dpkg --configure -a || true

echo ""
echo "Готово. Дальше: почините DNS (./deploy/fix_dns_ubuntu.sh), затем apt update."
