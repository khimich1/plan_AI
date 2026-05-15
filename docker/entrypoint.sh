#!/bin/sh
set -e
# Том /data при первом монтировании часто root:root — даём appuser запись в постоянные файлы.
if [ "$(id -u)" = "0" ]; then
  mkdir -p /data/drafts /data/outputs /data/logs /data/plans
  # Снимки БД из образа: один раз копируем на пустой том (пути из PB_DB_PATH / PLITA_DB_PATH).
  SEED_DIR=/app/docker/seed
  if [ -n "${PB_DB_PATH:-}" ] && [ ! -f "$PB_DB_PATH" ]; then
    if [ -f "$SEED_DIR/pb.db" ]; then
      echo "[entrypoint] Инициализация $PB_DB_PATH из $SEED_DIR/pb.db"
      cp -a "$SEED_DIR/pb.db" "$PB_DB_PATH"
    else
      echo "[entrypoint] Пропуск pb.db: нет $SEED_DIR/pb.db в образе. Положите файлы в docker/seed/ на машине сборки и пересоберите образ, либо смонтируйте каталог с готовыми БД на /data (см. docker-compose.split.yml)."
    fi
  fi
  if [ -n "${PLITA_DB_PATH:-}" ] && [ ! -f "$PLITA_DB_PATH" ]; then
    if [ -f "$SEED_DIR/plita.db" ]; then
      echo "[entrypoint] Инициализация $PLITA_DB_PATH из $SEED_DIR/plita.db"
      cp -a "$SEED_DIR/plita.db" "$PLITA_DB_PATH"
    else
      echo "[entrypoint] Пропуск plita.db: нет $SEED_DIR/plita.db в образе. Иначе приложение создаст новую пустую БД при старте (init_schema)."
    fi
  fi
  chown -R appuser:appuser /data
  exec gosu appuser "$@"
fi
exec "$@"
