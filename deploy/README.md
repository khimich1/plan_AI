# Деплой бота «Шишов» на Linux VPS

Краткая инструкция; полный контекст см. в плане проекта.

## 1. Подготовка

- Чеклист файлов вне Git: [ARTIFACTS_CHECKLIST.md](ARTIFACTS_CHECKLIST.md)
- Пример переменных окружения: [`../bot/bot.env.example`](../bot/bot.env.example)
- Локально: `python scripts/smoke_check.py` (из корня репозитория; код выхода **1**, если [FAIL] по токену или парсеру; [WARN] по БД не считаются ошибкой)
- Перед запуском на сервере остановите бота с тем же `BOT_TOKEN` на других машинах

## DNS: ошибка «Temporary failure resolving»

Если `apt update` / `git clone` падают с **Temporary failure resolving** — на сервере не работает DNS.

**Вариант A** — из уже клонированного репозитория (от root):

```bash
chmod +x deploy/fix_dns_ubuntu.sh && ./deploy/fix_dns_ubuntu.sh
```

**Вариант B** — репозитория ещё нет: с ПК по IP (SSH не требует работающего DNS на сервере для подключения по IPv4):

```bash
scp deploy/fix_dns_ubuntu.sh root@72.56.66.237:/root/
```

На сервере: `bash /root/fix_dns_ubuntu.sh`

После успешной проверки снова выполните `apt update` и шаги из раздела 2.

## 2. Код и данные на сервере

```bash
git clone <url> /opt/shishov
cd /opt/shishov
chmod +x deploy/bootstrap_vps.sh && ./deploy/bootstrap_vps.sh
```

Скопируйте на сервер в `/opt/shishov` (пути поправьте):

- `pb.db`, `plita.db`
- каталог `банк знаний/`

Пример с вашего ПК (OpenSSH):

```bash
scp pb.db plita.db root@YOUR_HOST:/opt/shishov/
scp -r "банк знаний" root@YOUR_HOST:/opt/shishov/
```

Создайте секреты (права только владельцу):

```bash
cp bot/bot.env.example bot/bot.env
nano bot/bot.env   # BOT_TOKEN=...
chmod 600 bot/bot.env
```

## 3. Проверка вручную

```bash
cd /opt/shishov
source .venv/bin/activate
python scripts/smoke_check.py
python run_bot.py
```

Убедитесь в логах, что бот стартовал; в Telegram проверьте ответ на сообщение. Остановите процесс (Ctrl+C) перед включением systemd.

## 4. systemd

```bash
sudo cp deploy/shishov-bot.service.example /etc/systemd/system/shishov-bot.service
sudo nano /etc/systemd/system/shishov-bot.service   # User, пути
sudo systemctl daemon-reload
sudo systemctl enable --now shishov-bot.service
journalctl -u shishov-bot.service -f
```

Создайте пользователя `shishov` и выдайте права на `/opt/shishov`, если не используете root в `User=`.

## 5. Резервное копирование

Скрипт из корня репозитория:

```bash
chmod +x scripts/backup_shishov_data.sh
./scripts/backup_shishov_data.sh
```

Архивы по умолчанию в `./backups/`. Другой каталог: `BACKUP_DIR=/var/backups/shishov ./scripts/backup_shishov_data.sh`.

Пример cron (ежедневно в 03:15):

```
15 3 * * * cd /opt/shishov && BACKUP_DIR=/var/backups/shishov ./scripts/backup_shishov_data.sh
```

## Зависимости и RAM

`requirements.txt` включает EasyOCR (тяжёлые зависимости). На слабом VPS может не хватить памяти при `pip install` или первом OCR — см. план деплоя (swap / тариф).
