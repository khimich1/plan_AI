# Spec: Приём ответа 1С с номером счёта

> **Фаза SDD:** SPECIFY. План и код — после приёмки этой спеки.
> **Статус:** 2026-11-02 ответ уже лежит на диске, программа его не читает.
> **Дата:** 2026-11-02
> **Исходящая папка:** [`exchange-export-dir.md`](exchange-export-dir.md). Её не переоткрывать.
> **Связанные модули:** `app/services/invoice_reply_import.py`, `core/invoice_reply.py`, `core/kp_db_schema.py` (`record_kp_order_ids`), `app/main.py` (`lifespan`), `core/config/settings.py`, `docker-compose.split.yml`.
> **Рядом, не это:** HTTP в 1С, отметка оплаты, переход в производство, правка `docker/entrypoint.sh`, повторная отправка КП №183.

---

## Assumptions I'm Making

1. 1С уже забрала `new_183_…json` и положила ответ. Карточка ждёт только запись номера в базу.
2. Разбор файла не пишется заново. `parse_invoice_reply` и `import_otvet_file` остаются как есть: номер и два УИД в пустые колонки, файл удаляется после записи, чужой номер и битый JSON файл не удаляют.
3. Папка ответов на хосте — `/srv/exchange-1c/from_1c`. В контейнер она монтируется в `/exchange/from_1c`, не под `/data`.
4. Процесс приложения работает от uid 1000. У этого uid на каталоге уже есть `rwx`, на лежащем файле — чтение. Новые файлы получают то же чтение через default ACL. Права заново не назначаются, если проверка это подтверждает.
5. Постоянное чтение появляется только в новом образе. Уже лежащий ответ записывается текущим образом одним запуском `import_otvet_file`, без пересборки.
6. Появление номера не ставит оплату и не переводит КП в производство.

→ Если что-то из этого неверно — править spec до кода и до команд на сервере.

---

## Objective

Менеджер открывает КП, которое уже ушло в 1С, и видит номер счёта, как только компьютер 1С положил `otvet_*.json`. Для КП №183 это номер 782.

Сейчас файл `otvet_353400.json` лежит в `/srv/exchange-1c/from_1c` с 11:15. В нём `Номер_КП` 183, `Номер_ЗК` 782, оба УИД и текст, что коммерческое предложение и заказ клиента сформированы. Функция разбора есть в работающем образе с 14 октября, но её вызывают только тесты. Cron на сервере нет. Папка в контейнер не смонтирована. Поэтому карточка пишет «Счёт · номер ещё не пришёл из 1С».

Успех:

- у КП №183 в базе `order_number_1c` = `782`, оба УИД записаны, статус остаётся «на согласовании»;
- `otvet_353400.json` после записи удалён;
- следующий `otvet_*.json` в ту же папку сам попадает в базу не позже чем через минуту после появления;
- контейнер с `APP_ENV=production` не стартует, если каталог ответов не смонтирован отдельно.

## Проверенный ответ

`/srv/exchange-1c/from_1c/otvet_353400.json`, 428 байт, владелец `exchange1c`, время 2026-11-02 08:15 UTC.

| Поле | Значение |
|---|---|
| `Номер_КП` | 183 |
| `Номер_ЗК` | 782 |
| `УИД_ЗК` | есть, непустой |
| `УИД_КП` | есть, непустой |
| `Описание` | предложение и заказ клиента сформированы |

Эти поля уже принимает `parse_invoice_reply`. Файл не переписывать.

Права на 2026-11-02:

- каталог `exchange1c:exchange1c`, у uid 1000 ACL `rwx`;
- файл: у uid 1000 чтение;
- default ACL каталога даёт uid 1000 чтение на новые файлы.

Удаление делает запись на каталог, она у uid 1000 уже есть.

## Tech Stack

Python 3, FastAPI lifespan, Docker Compose, pytest. Новых библиотек нет. Разбор JSON не меняется.

## Commands

```bash
pytest tests/test_invoice_reply_import.py tests/test_exchange_import_mount.py -q
docker compose -f /opt/plan_zavod/docker-compose.split.yml up -d --no-deps --build backend
```

Образ пересобирается только после того, как лежащий ответ уже записан в базу. `docker compose down -v` не выполнять.

## Project Structure

```
ai_docs/specs/exchange-import-otvet.md     → эта спека
core/config/settings.py                    → EXCHANGE_IMPORT_DIR
core/exchange_import_guard.py              → отказ старта без отдельного монтирования
app/main.py                                → проверка и цикл чтения в lifespan
docker-compose.split.yml                   → переменная и монтирование вне /data
tests/test_exchange_import_mount.py        → четыре случая без живой базы и без файла КП 183
```

`app/services/invoice_reply_import.py` и `core/invoice_reply.py` не меняются.

## Code Style

Проверка старта — та же форма, что у исходящего каталога. Цикл только читает каталог из настройки и вызывает уже существующий `import_otvet_dir`.

```python
def fail_fast_if_exchange_import_not_mounted(
    app_env: str,
    directory: Path,
    *,
    environ: Mapping[str, str],
    mount_points: Iterable[Path],
) -> None:
    if app_env.strip().lower() != "production":
        return
    if "EXCHANGE_IMPORT_DIR" not in environ:
        raise ExchangeImportMisconfigured("EXCHANGE_IMPORT_DIR не задан")
    if directory.resolve() not in {path.resolve() for path in mount_points}:
        raise ExchangeImportMisconfigured("каталог ответов не смонтирован отдельно")
```

В `development` переменная может отсутствовать: цикл не запускается, запасной каталог не создаётся.

В compose у сервиса `backend`:

```yaml
environment:
  EXCHANGE_IMPORT_DIR: /exchange/from_1c
volumes:
  - /srv/exchange-1c/from_1c:/exchange/from_1c
```

Путь внутри контейнера не начинается с `/data`. `chown -R appuser:appuser /data` в `docker/entrypoint.sh` эту папку не видит и не меняется.

Цикл: один проход сразу при старте, дальше раз в 60 секунд, пока процесс жив. Ошибка одного файла логируется и не останавливает цикл. Повторный проход по уже записанному номеру удаляет файл и второе оповещение не пишет — так уже делает `import_otvet_file`.

## Testing Strategy

Существующий `tests/test_invoice_reply_import.py` остаётся зелёным и не переписывается.

Новый `tests/test_exchange_import_mount.py`, без сети, без `plita.db` и без `otvet_353400.json`.

| Случай | Ожидание |
|---|---|
| `development`, переменная не задана, каталог не смонтирован | исключения нет |
| `production`, ключа `EXCHANGE_IMPORT_DIR` нет | `ExchangeImportMisconfigured` |
| `production`, переменная задана, пути нет среди точек монтирования | `ExchangeImportMisconfigured` |
| `production`, переменная задана, путь есть среди точек монтирования | исключения нет |

Команда: `pytest tests/test_invoice_reply_import.py tests/test_exchange_import_mount.py -q`.

Живая проверка КП №183 — список ниже, не pytest.

## Boundaries

- Always: записать уже лежащий ответ до пересборки образа; проверить, что uid 1000 по-прежнему читает файл и может удалить его из каталога; не менять байты `otvet_353400.json`.
- Ask first: менять `docker/entrypoint.sh`, переписывать ACL, если проверка прав не сошлась, менять разбор JSON, ставить оплату вместе с номером.
- Never: монтировать ответы под `/data`; `docker compose down -v`; править `plita.db` руками; удалять `otvet_353400.json` до успешной записи номера; отправлять КП №183 в 1С ещё раз; записывать номер счёта в `supply_contract`.

## Порядок на сервере

Пересборка образа со отказом старта остановит сайт, если монтирования ещё нет. Поэтому номер 782 записывается текущим образом.

1. Проверить ACL. У uid 1000 на каталоге должно быть `rwx`, на `otvet_353400.json` — чтение. Если чтения нет, `setfacl` только на этот файл и default ACL каталога. Владельца `exchange1c` не менять.
2. Вписать в серверный compose переменную и монтирование `/srv/exchange-1c/from_1c:/exchange/from_1c`. Образ не пересобирать.
3. Пересоздать только backend: `docker compose -f /opt/plan_zavod/docker-compose.split.yml up -d --no-deps backend`.
4. Убедиться, что `findmnt /exchange/from_1c` показывает `/srv/exchange-1c/from_1c`, а `findmnt /data` по-прежнему том `web_split_data`. Файл внутри контейнера виден от uid 1000.
5. Одним запуском от uid 1000 вызвать `import_otvet_file` для этого пути и базы `/data/plita.db`. После этого файл на хосте отсутствует, у КП №183 номер `782`, оба УИД заполнены, статус «на согласовании».
6. Код цикла и отказа старта выкладывать отдельным образом только после шага 5. В том же выкладе на сервере уже должно быть монтирование из шага 2.

## Success Criteria

- Карточка КП №183 показывает «Счёт 782». `order_number_1c` равен `782`. Статус «на согласовании». `paid_at` пустой.
- `otvet_353400.json` удалён только после этой записи.
- Договор в `supply_contract` не изменился.
- Backend healthy, frontend не перезапускался, исходящее монтирование `/data/exchange/to_1c` на месте.
- Следующий файл `otvet_*.json`, положенный в `/srv/exchange-1c/from_1c`, исчезает после записи номера, и это делает цикл в процессе, без ручного запуска.
- `pytest tests/test_invoice_reply_import.py tests/test_exchange_import_mount.py -q` проходит.
- При `APP_ENV=production` без отдельного монтирования `EXCHANGE_IMPORT_DIR` процесс не доходит до приёма запросов.

## Open Questions

Нет.
