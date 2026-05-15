Каталог для стартовых SQLite, которые ВПИСЫВАЮТСЯ В ОБРАЗ при команде docker compose build
(контекст сборки копирует docker/seed/ в /app/docker/seed внутри образа).

ВАЖНО:
- Если на сервере только git pull + build БЕЗ ваших pb.db/plita.db в этом каталоге —
  в образе не будет сидов, entrypoint нечего копировать в /data.
- Файлы в /data при named volume (web_split_data) лежат в хранилище Docker, не в папке
  репозитория на хосте. Проверка: docker compose -f docker-compose.split.yml run --rm backend ls -la /data

Сценарий A — сиды в образе:
  1) cp …/pb.db docker/seed/pb.db && cp …/plita.db docker/seed/plita.db
  2) docker compose -f docker-compose.split.yml build
  3) первый запуск на пустом томе: скопирует в /data/pb.db и /data/plita.db
  Если контейнер уже создавал пустые БД — удалите том: docker compose … down -v и up снова.

Сценарий B — без пересборки, файлы на сервере:
  В compose для backend замените том на привязку каталога, положите туда db:
    volumes:
      - ./data:/data
  и скопируйте pb.db, plita.db в ./data/ на сервере.

Проверить, есть ли сиды в образе:
  docker run --rm plan-wed-backend:split ls -la /app/docker/seed
