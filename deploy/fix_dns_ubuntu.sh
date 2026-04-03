#!/usr/bin/env bash
# Явные DNS для Ubuntu с systemd-resolved (обход "Temporary failure resolving").
# Запуск на VPS от root:  bash deploy/fix_dns_ubuntu.sh
# Если репозитория ещё нет — скопируйте этот файл с ПК: scp deploy/fix_dns_ubuntu.sh root@IP:/root/
set -euo pipefail

if [[ "${EUID:-0}" -ne 0 ]]; then
  echo "Запустите от root: sudo bash $0"
  exit 1
fi

DROPIN_DIR=/etc/systemd/resolved.conf.d
DROPIN="${DROPIN_DIR}/99-shishov-dns.conf"

mkdir -p "$DROPIN_DIR"
[[ -f "$DROPIN" ]] && cp -a "$DROPIN" "${DROPIN}.bak.$(date +%Y%m%d%H%M%S)"

cat > "$DROPIN" << 'EOF'
# Добавлено для стабильного резолва (Google + Cloudflare).
# Удалите файл и выполните: systemctl restart systemd-resolved
[Resolve]
DNS=8.8.8.8 1.1.1.1
FallbackDNS=8.8.8.8 1.1.1.1
EOF

systemctl restart systemd-resolved
sleep 2

echo "=== Статус DNS (resolvectl) ==="
resolvectl status 2>/dev/null || true
echo ""
echo "=== Проверка имён ==="
if getent hosts github.com >/dev/null 2>&1; then
  echo "OK: github.com резолвится"
else
  echo "FAIL: github.com не резолвится — проверьте файрвол/панель хостинга или поддержку"
  exit 1
fi
if getent hosts archive.ubuntu.com >/dev/null 2>&1; then
  echo "OK: archive.ubuntu.com резолвится"
else
  echo "WARN: archive.ubuntu.com не резолвится (apt может не работать)"
fi

echo ""
echo "Дальше: apt update && apt install -y git python3 python3-venv python3-pip"
