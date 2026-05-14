#!/bin/sh
set -e
# Том /data при первом монтировании часто root:root — даём appuser запись в постоянные файлы.
if [ "$(id -u)" = "0" ]; then
  mkdir -p /data/drafts /data/outputs /data/logs /data/plans
  chown -R appuser:appuser /data
  exec gosu appuser "$@"
fi
exec "$@"
