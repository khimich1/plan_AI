# Отчёт: SFTP-вход для обмена с 1С

**Дата:** 2026-10-02  
**Статус:** сделано  
**Код приложения:** не менялся

## Задача

Отдельный вход SFTP, чтобы компьютер 1С забирал `new_*.json` и `delta_*.json` и позже клал `otvet_*.json`. WinSCP и компьютер 1С не настраивались. Существующий вход root не ломался.

## Что получилось

Пользователь `exchange1c` заходит только по SFTP в chroot `/srv/exchange-1c` и видит две папки:

| Удалённая папка | Что это | Куда смотрит приложение |
|---|---|---|
| `/to_1c` | Файлы, которые пишет backend | Тот же каталог, что `/data/exchange/to_1c` в контейнере и `/opt/plan_zavod/data/exchange/to_1c` на хосте |
| `/from_1c` | Сюда WinSCP кладёт `otvet_*.json` | В контейнер не смонтирована, приложение ответы 1С не читает |

Оболочки нет. TCP-перенаправление, X11, агент и туннели для этого пользователя выключены.

**Ограничение по IP компьютера 1С не сделано.** Файрвол не сужался и новые порты не открывались. Порт 22 слушает как раньше, зайти может любой адрес, с которого SSH уже был доступен.

## Параметры для WinSCP

Их прописывают на компьютере 1С. Сама WinSCP с сервера не настраивалась.

| | |
|---|---|
| Протокол | SFTP |
| Хост | `zavodstart.ru` (`5.129.200.30`) |
| Порт | `22` |
| Пользователь | `exchange1c` |
| Пароль | выдан один раз в переписке 2026-10-02, в репозиторий, `.env`, логи и этот отчёт не записывался |
| Забирать | `/to_1c` |
| Отдавать | `/from_1c` |

Два задания на компьютере 1С:

1. Из `/to_1c` забирать `new_*.json` и `delta_*.json` в `C:\Exchange` и удалять их на сервере после успешной передачи.
2. Из `C:\Exchange` увозить в `/from_1c` только `otvet_*.json`.

## Почему chroot не внутри проекта

OpenSSH принимает `ChrootDirectory`, только если сам каталог и все каталоги выше принадлежат root и не доступны на запись группе и остальным. `/opt/plan_zavod` и `/opt/plan_zavod/data` принадлежат uid 1000 и доступны группе на запись, поэтому корень jail там сделать нельзя.

Корень jail: `/srv/exchange-1c`, владелец `root:root`, режим `755`. `/` и `/srv` тоже `root:root` и `755`.

`/to_1c` внутри jail — bind того же каталога, куда пишет приложение:

```
/opt/plan_zavod/data/exchange/to_1c  /srv/exchange-1c/to_1c  none  bind  0  0
```

Строка добавлена в `/etc/fstab`, монтирование переживает перезагрузку сервера. Docker-монтирование `./data/exchange/to_1c:/data/exchange/to_1c` и `EXCHANGE_EXPORT_DIR=/data/exchange/to_1c` не менялись.

## Права

`docker/entrypoint.sh` при каждом старте backend делает `chown -R appuser:appuser /data`. Этот chown меняет владельца смонтированного `to_1c` на uid/gid 1000. Группу он тоже сбрасывает, биты режима и ACL — нет. Поэтому доступ SFTP повешен на ACL, а не на группу, и entrypoint менять не пришлось.

На каталоге `to_1c`:

- владелец `1000:1000`, приложению хватает прав владельца, чтобы создавать файлы;
- именованный пользователь `exchange1c`: `rwx` на каталог;
- default ACL: новые файлы получают для `exchange1c` чтение (на файле эффективно `r--`);
- остальные на каталог прав не имеют.

Запись в каталог нужна, чтобы WinSCP могла удалить файл после успешной передачи. Отдельного права «только удалить» в обычных правах Linux нет, поэтому `exchange1c` может и положить файл в `/to_1c`. Класть ответы 1С нужно в `/from_1c`.

`/from_1c`: владелец `exchange1c:exchange1c`, режим `700`.

Домашний каталог учётной записи — `/var/lib/exchange1c` (`700`, вне jail). Так PAM не создаёт служебные файлы вроде `.cache` внутри `/from_1c`. Оболочка — `/usr/sbin/nologin`. Uid и gid — `2001`.

На хосте установлен пакет `acl` (команды `setfacl` / `getfacl`).

## SSH

Файл `/etc/ssh/sshd_config.d/zz-exchange1c.conf` (по имени идёт после `50-cloud-init.conf`):

```
Match User exchange1c
    ChrootDirectory /srv/exchange-1c
    ForceCommand internal-sftp
    AllowTcpForwarding no
    X11Forwarding no
    AllowAgentForwarding no
    PermitTunnel no
    PermitTTY no
    PasswordAuthentication yes
    AuthorizedKeysFile /dev/null
```

`PasswordAuthentication yes` внутри `Match` относится только к `exchange1c`. Глобальный `PasswordAuthentication`, `PermitRootLogin` и `Port` не менялись. Перед reload выполнен `sshd -t`. Конфиг перечитан через `systemctl reload ssh`, не через restart. Действующая сессия root после reload осталась.

Для root `sshd -T` до и после файла совпал. Для `exchange1c` эффективные значения: chroot `/srv/exchange-1c`, `internal-sftp`, TCP/X11/туннель/TTY выключены.

## Проверка

С localhost:

- SFTP-вход есть, видны `/to_1c` и `/from_1c`;
- оболочка не выдаётся (`This service allows sftp connections only`);
- запрос X11 сервер отклоняет;
- локальный проброс TCP доходит до ssh-клиента, сервер отвечает `administratively prohibited`;
- `/opt/plan_zavod`, `/etc` и каталог тома с базами (`/var/lib/docker/volumes/plan_zavod_web_split_data/_data`) из сессии не видны;
- файл, созданный от uid 1000 с umask `077`, виден в `/to_1c`, читается и удаляется этим пользователем;
- загрузка в `/from_1c` проходит, файл принадлежит `exchange1c`.

Пробные файлы удалены. Имена `new_*.json` и `delta_*.json` не использовались.

Затем backend перезапущен так, чтобы entrypoint снова выполнил `chown`, без `down -v`:

```bash
docker stop plan_zavod-backend-1
docker compose -f docker-compose.split.yml up -d --no-deps backend
```

Обычный `up -d` на уже запущенном контейнере entrypoint не запускает, поэтому контейнер сначала остановлен. Frontend не перезапускался. После старта backend снова healthy, ACL и режим `to_1c` те же, повторный файл от uid 1000 снова виден и удаляется. `from_1c` в контейнере нет.

Пароль после проверки стёрт с диска. Поиск по репозиторию (кроме `.git`, `node_modules`, виртуальных окружений и `data`), по `.env`, `auth.log` и `syslog` совпадений не дал.

## Что не менялось

- код приложения, включая `docker/entrypoint.sh`;
- `docker-compose.split.yml` и монтирование `to_1c`;
- `Port`, `PermitRootLogin`, глобальный `PasswordAuthentication`;
- файрвол;
- WinSCP и компьютер 1С.
