Каталог для стартовых SQLite, которые ВПИСЫВАЮТСЯ В ОБРАЗ при команде docker compose build
(контекст сборки копирует docker/seed/ в /app/docker/seed внутри образа).

ВАЖНО:
- plita.db НЕ кладётся в seed и не коммитится в git (безопасность: нет дефолтного admin).
- При первом запуске backend создаёт пустую схему (init_schema в lifespan).
- Первого пользователя-web-admin создайте вручную:

    python scripts/create_admin.py --username admin

  (на сервере — в контейнере backend или с PLITA_DB_PATH на смонтированный /data)

- pb.db (прайсы) — опциональный seed; если файла нет в образе, entrypoint пропустит копирование.
- Файлы в /data при named volume (web_split_data) лежат в хранилище Docker, не в папке
  репозитория на хосте. Проверка: docker compose -f docker-compose.split.yml run --rm backend ls -la /data

Сценарий A — seed прайсов в образе:
  1) cp …/pb.db docker/seed/pb.db
  2) docker compose -f docker-compose.split.yml build
  3) первый запуск на пустом томе: скопирует pb.db в /data; plita.db — пустая схема при старте
  4) python scripts/create_admin.py --username admin
  Если контейнер уже создавал пустые БД — удалите том: docker compose … down -v и up снова.

Сценарий B — без пересборки, файлы на сервере:
  В compose для backend замените том на привязку каталога, положите туда db:
    volumes:
      - ./data:/data
  и скопируйте pb.db в ./data/ на сервере (plita.db создастся при старте или скопируйте свою).

Проверить, есть ли сиды в образе:
  docker run --rm plan-wed-backend:split ls -la /app/docker/seed
