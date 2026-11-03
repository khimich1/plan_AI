# Spec: Каталог обмена КП с 1С

> **Фаза SDD:** SPECIFY. План и код — после приёмки этой спеки.
> **Статус:** путь на сервере подтверждён 2026-11-02. Файл КП №183 проверен и не пересобирается.
> **Дата:** 2026-11-02
> **Связанные модули:** `core/config/settings.py` (`exchange_export_dir`), `app/services/archive_service.py` (`_write_invoice_file`), `app/main.py` (`lifespan`), `docker-compose.split.yml`, `docker/entrypoint.sh` (не менять).
> **Рядом, не это:** разбор `otvet_*.json` — отдельная спека [`exchange-import-otvet.md`](exchange-import-otvet.md). HTTP в 1С по-прежнему не входит.

---

## Assumptions I'm Making

1. Чинится только исходящая папка. Компьютер 1С уже забирает `new_*.json` и `delta_*.json` из `/opt/plan_zavod/data/exchange/to_1c` по SFTP и удаляет файл после успешной передачи.
2. КП №183 уже в статусе «на согласовании». Повторная кнопка «Отправить в 1С» не нажимается, статус и база не откатываются. В папку 1С кладётся уже записанный файл.
3. Compose на сервере — `/opt/plan_zavod/docker-compose.split.yml`. Относительный путь `./data/exchange/to_1c` считается от этого файла.
4. Локально, при `APP_ENV`, отличном от `production`, программа по-прежнему пишет в `data/exchange_export` внутри проекта.
5. `chown -R appuser:appuser /data` в `docker/entrypoint.sh` остаётся. Он меняет владельца на uid 1000 и не снимает ACL пользователя `exchange1c`.

→ Если что-то из этого неверно — править spec до кода и до команд на сервере.

---

## Objective

Менеджер нажимает «Отправить в 1С». Программа пишет UTF-8 JSON в каталог, который компьютер 1С уже читает. Уже готовый файл КП №183 попадает туда без второй сборки.

Сейчас программа пишет в `/app/data/exchange_export/` внутри контейнера. Это запасной путь из кода: переменная `EXCHANGE_EXPORT_DIR` в контейнере пустая, а единственный диск — том `web_split_data` на `/data`. Папка SFTP к контейнеру не подключена. Каталог `/data/exchange/to_1c` внутри тома существует и пуст, но это не `/opt/plan_zavod/data/exchange/to_1c`.

Успех:

- `new_183_20261102T075653342235.json` лежит в папке SFTP, 1С забирает его и удаляет;
- содержимое JSON, статус КП и номер счёта не менялись;
- backend после пересоздания healthy, том с базой на месте;
- при следующем выкладе контейнер с `APP_ENV=production` не стартует, если каталог обмена не смонтирован отдельно.

## Проверенный файл КП №183

Файл на момент проверки: `/app/data/exchange_export/new_183_20261102T075653342235.json`, 1815 байт, UTF-8, перевод строки в конце. Пересобирать его нельзя.

| Проверка | Факт |
|---|---|
| Вид | Первая отправка. Ключей `УИД_ЗК`, `Номер_ЗК`, `УИД_КП` нет |
| Шапка | «КоммерческоеПредложение», номер 183, дата `2026-11-02`, организация «ООО «Комбинат ЖБК»» |
| Клиент | ШТОЛЬФ ООО. Код, ИНН и КПП совпадают со справочником и карточкой КП. GUID контрагента пустой, как в уже принятом примере КП №22: 1С узнаёт клиента по коду |
| Договор | `ТБ-3/08/26Б`, ключ сразу после «Контрагент». В реестре это договор id 187 того же клиента |
| Склад | «Склад Готовой Продукции», самовывоз, доставка 0 |
| Состав | Одна строка «Плиты ПБ 66-12-8п», 4 шт, цена 23 734. Других позиций в КП нет. GUID плиты совпадает с воротами счёта по `pb.db` |
| Суммы | 94 936,00; НДС 17 119,61; скидка 0. Шапка равна строке и полям `KP_offers` |
| Поле «Статус» в JSON | Литерал «в архиве». Карточка при этом уже «на согласовании». Так собирается каждый исходящий файл |

## Tech Stack

Python 3, FastAPI, Pydantic Settings, Docker Compose, pytest. Новых библиотек нет.

## Commands

```bash
pytest tests/test_exchange_export_mount.py -q
docker compose -f /opt/plan_zavod/docker-compose.split.yml up -d --no-deps backend
```

`docker compose down -v` не выполнять: флаг `-v` удаляет том с базой.

## Project Structure

```
ai_docs/specs/exchange-export-dir.md   → эта спека
core/exchange_export_guard.py          → отказ старта, если каталог не смонтирован
app/main.py                            → вызов проверки в lifespan
docker-compose.split.yml               → переменная и монтирование
tests/test_exchange_export_mount.py    → четыре случая без живой базы и без JSON КП 183
```

`app/services/archive_service.py` и `core/invoice_export.py` не меняются: сборка документа уже верная, ошибка только в каталоге.

## Code Style

Проверка старта повторяет `fail_fast_if_destructive_flags_in_production`: в production либо молчит, либо бросает ошибку до приёма запросов. Разбор `/proc/self/mountinfo` вынесен в функцию, которую тест подменяет.

```python
def fail_fast_if_exchange_dir_not_mounted(
    app_env: str,
    directory: Path,
    *,
    environ: Mapping[str, str],
    mount_points: Iterable[Path],
) -> None:
    if app_env.strip().lower() != "production":
        return
    if "EXCHANGE_EXPORT_DIR" not in environ:
        raise ExchangeExportMisconfigured("EXCHANGE_EXPORT_DIR не задан")
    if directory.resolve() not in {path.resolve() for path in mount_points}:
        raise ExchangeExportMisconfigured("каталог обмена не смонтирован отдельно")
```

Вызов стоит в `lifespan` сразу после `fail_fast_if_destructive_flags_in_production`. Рабочий код читает `os.environ` и точки монтирования процесса. Тест передаёт свои.

В `docker-compose.split.yml` у сервиса `backend`:

```yaml
environment:
  EXCHANGE_EXPORT_DIR: /data/exchange/to_1c
volumes:
  - ./data/exchange/to_1c:/data/exchange/to_1c
```

Одной переменной недостаточно. Путь `/data/exchange/to_1c` без отдельного монтирования — пустой подкаталог тома базы, 1С его не видит.

## Testing Strategy

Один новый файл pytest, без сети, без `plita.db` и без файла КП №183.

| Случай | Ожидание |
|---|---|
| `development`, переменная не задана, каталог не смонтирован | исключения нет |
| `production`, ключа `EXCHANGE_EXPORT_DIR` нет в окружении | `ExchangeExportMisconfigured` |
| `production`, переменная задана, пути нет среди точек монтирования | `ExchangeExportMisconfigured` |
| `production`, переменная задана, путь есть среди точек монтирования | исключения нет |

Команда: `pytest tests/test_exchange_export_mount.py -q`.

Живая проверка на сервере — отдельный список ниже, не pytest.

## Boundaries

- Always: скопировать файл КП №183 до пересоздания контейнера; проверить чтение для `exchange1c`; не создавать пробные `new_*.json` и `delta_*.json`.
- Ask first: монтировать `/from_1c`, вызывать разбор ответов, менять `docker/entrypoint.sh`, откатывать статус КП, пересобирать образ backend.
- Never: `docker compose down -v`; править `plita.db`; менять байты JSON; удалять том `web_split_data`; нажимать «Отправить в 1С» или «исправление» для КП №183.

## Порядок на сервере

Пересоздание контейнера стирает `/app/data/exchange_export`. Пока файл не скопирован наружу, второй отправки уже нет.

1. Скопировать файл в папку 1С и в запасную копию вне неё:

```bash
docker cp plan_zavod-backend-1:/app/data/exchange_export/new_183_20261102T075653342235.json \
  /opt/plan_zavod/data/exchange/to_1c/
cp /opt/plan_zavod/data/exchange/to_1c/new_183_20261102T075653342235.json \
  /root/new_183_20261102T075653342235.json
```

2. У файла должно быть чтение для `exchange1c`. У каталога это право для новых файлов уже задано. Если `getfacl` его не показывает, добавить только на этот файл: `setfacl -m u:exchange1c:r`. Удаление после забора делает запись на каталог, она у `exchange1c` уже есть.
3. Вписать переменную и монтирование в серверный compose. Образ не пересобирать.
4. Пересоздать только backend: `docker compose -f /opt/plan_zavod/docker-compose.split.yml up -d --no-deps backend`.
5. Контейнер healthy. В окружении `EXCHANGE_EXPORT_DIR=/data/exchange/to_1c`. `findmnt /data/exchange/to_1c` показывает `/opt/plan_zavod/data/exchange/to_1c`, не том `web_split_data`. КП №183 по-прежнему «на согласовании», `order_number_1c` пустой.
6. Успешный забор: файл исчез из `/to_1c` после захода `exchange1c`. Это удаление компьютером 1С. Если файл остаётся дольше нескольких минут, смотреть ACL, а не отправлять КП заново.

Код отказа старта выкладывается после того, как этот файл уехал. Иначе проверка остановит сайт раньше, чем compose получит диск. Строки compose коммитятся вместе с проверкой: без них в git следующий выклад снова поднимет контейнер без папки 1С, а с проверкой такой контейнер не запустится молча.

## Success Criteria

- В папке SFTP побывал исходный `new_183_20261102T075653342235.json` длиной 1815 байт, затем 1С его забрала.
- SHA-256 файла в папке 1С совпадает с файлом, который лежал в контейнере до копирования.
- `kp_meta.status` КП №183 остаётся «на согласовании», `order_number_1c` остаётся пустым.
- Backend healthy, frontend не перезапускался, том `web_split_data` подключён к `/data`.
- `pytest tests/test_exchange_export_mount.py -q` проходит.
- При `APP_ENV=production` без отдельного монтирования `EXCHANGE_EXPORT_DIR` процесс не доходит до приёма запросов.

## Open Questions

Нет. Приём `otvet_*.json` и запись номера счёта в программу в эту спеку не входят.
