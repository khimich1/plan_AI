# Спецификация: безопасная выкладка договоров на zavodstart.ru

> **Фаза SDD:** SPECIFY. План: [`ai_docs/develop/plans/2026-11-03-vykladka-dogovorov-bez-obryva-obmena.md`](../develop/plans/2026-11-03-vykladka-dogovorov-bez-obryva-obmena.md).  
> **Статус:** выложена 2026-11-03. Backend и frontend пересобраны. В `contractBlank.ts` на диске сервера пустые доверенность, КПП и должность проверяются через `?.trim()`: без этого `tsc` образ frontend не собрал. Этот файл в git ещё не закоммичен.  
> **Код:** `origin/web` `968401c`.  
> **Рядом, не это:** подмена compose и `.env`, копирование `plita.db`, правка реестра, повторная отправка КП №183.

## Допущения

1. Выкладываем только то, что уже в `origin/web` на `968401c`. Незакоммиченные правки пробелов и локальный `plita.db` на сервер не едут.
2. Обмен с 1С уже работает и в этой выкладке не перенастраивается. Compose, `.env`, тома и `nginx.himych.conf` на сервере не подменяем.
3. Один образ backend принесёт и договоры, и загрузку прайса цельных свай. На диске сервера прайс уже новый, в процессе ещё старый. Отдельно его не отрезать.
4. Базу не копируем и не мигрируем руками. Старт не пересоздаёт `supply_contract`: дата и форма уже допускают пустое значение.
5. Смена правила счёта входит в выкладку. После неё договор больше не привязывается сам по уникальному ИНН.
6. КП №183 повторно не отправляем. Номер `782` и статус «на согласовании» должны остаться.

## Цель

Менеджер на сайте находит клиента 1С и либо привязывает уже известный номер договора, либо собирает новый бланк. Счёт по-прежнему пишется в папку, которую забирает 1С, а ответ с номером счёта по-прежнему читается из `/srv/exchange-1c/from_1c`.

Пользователь — менеджер на `zavodstart.ru`. Успех — новый экран есть, КП №183 не тронуто, обменные каталоги те же, живая база на месте.

Сейчас этого нет. Backend поднят 2 ноября, frontend собран 30 октября. В работающем коде нет `save_directory_blank` и `search_directory`. Отправка всё ещё вызывает `link_unique_inn_contracts`. Каталог сервера не git-репозиторий, поэтому `docker compose build` там соберёт старый сервис договора и старый `ContractCardList.tsx`.

## Стек

Тот же: FastAPI в `plan_zavod-backend-1`, React в `plan_zavod-frontend-1`, compose `/opt/plan_zavod/docker-compose.split.yml`, база `/data/plita.db` в томе `web_split_data`. Новых зависимостей нет.

## Команды

Локально, до копирования:

```bash
cd /home/roman/project/Шишов
git rev-parse --short HEAD
pytest tests/test_archive_directory_contracts.py tests/test_kp_invoice_existing_contract.py tests/test_exchange_export_mount.py tests/test_exchange_import_mount.py tests/test_price_desk_service.py -q
```

`HEAD` должен быть `968401c`.

На сервере после копирования только runtime-файлов, без тестов и `ai_docs`:

```bash
cd /opt/plan_zavod
docker compose -f docker-compose.split.yml build backend
docker compose -f docker-compose.split.yml up -d --no-deps backend
docker compose -f docker-compose.split.yml build frontend
docker compose -f docker-compose.split.yml up -d --no-deps frontend
```

`docker compose down` и `down -v` не выполнять.

## Структура проекта

На сервер копируются только эти пути из `968401c`, поверх тех же путей в `/opt/plan_zavod`:

```
app/api/v1/endpoints/archive.py
app/repositories/supply_contract_repository.py
app/schemas/archive.py
app/services/supply_contract_service.py
app/services/price_desk_service.py
core/supply_contract_card_parse.py
frontend/src/features/commercial-archive/api/archiveApi.ts
frontend/src/features/commercial-archive/components/ContractCardDrawer.tsx
frontend/src/features/commercial-archive/components/ContractCardList.tsx
frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx
frontend/src/features/commercial-archive/hooks/useArchiveQueries.ts
frontend/src/features/commercial-archive/lib/contractBlank.ts
frontend/src/features/commercial-archive/types/supplyContract.ts
```

Уже совпадает с работающим контейнером и вторично не копируется: `app/main.py`, `core/exchange_export_guard.py`, `core/exchange_import_guard.py`, `core/config/settings.py`, `docker-compose.split.yml`.

Не трогать: `.env`, `frontend/nginx.docker.conf`, `frontend/nginx.himych.conf`, `/data/plita.db`, `data/exchange/to_1c`, `/srv/exchange-1c/from_1c`.

## Стиль выкладки

Образ собирается из диска сервера. Пока новый сервис не лежит в `/opt/plan_zavod`, сборка не является выкладкой `968401c`.

Проверка после сборки backend, до того как считать выкладку успешной:

```bash
docker exec plan_zavod-backend-1 python -c "import pathlib; t=pathlib.Path('/app/app/services/supply_contract_service.py').read_text(); a=pathlib.Path('/app/app/services/archive_service.py').read_text(); print('blank', 'def save_directory_blank' in t); print('autolink', 'link_unique_inn_contracts' in a)"
```

Ожидание: `blank True`, `autolink False`.

## Стратегия тестирования

Локально pytest из блока команд, без живой базы и без SSH. На сервере автоматических тестов нет. Приёмка ручная:

- `https://zavodstart.ru/health` отвечает 200;
- `findmnt` в backend показывает `/opt/plan_zavod/data/exchange/to_1c` и `/srv/exchange-1c/from_1c`, а `/data` остаётся томом `web_split_data`;
- у КП №183 статус «на согласовании» и `order_number_1c` равен `782`;
- в образе есть сохранение бланка и нет автопривязки;
- на сайте у клиента без договора видна строка «в реестре договора нет», поиск «МД-СТРОЙ» по-прежнему находит карточку справочника;
- в nginx сайта остаётся `client_max_body_size 50m`.

## Границы

- Всегда: сначала файлы из списка, потом сборка. Backend и frontend по отдельности, с `--no-deps`. Снимок `/data/plita.db` до первой записи в реестр после выкладки. Health, монтирования и КП №183 проверить до того, как считать работу законченной.
- Сначала спроси: подмена compose или `.env`, копирование всей `plita.db`, правка реестра договоров, повторная отправка КП №183, выкладка чего-либо кроме `968401c`.
- Никогда: `docker compose down -v`. Копирование локального `plita.db` поверх серверной базы. Удаление файлов из папок обмена. Коммит секретов. Сборка frontend с чужим `nginx.docker.conf`.

## Критерии успеха

- В контейнере `save_directory_blank` есть, `link_unique_inn_contracts` в `archive_service.py` нет.
- На карточке архивного КП без договора есть «в реестре договора нет» и «Собрать договор».
- КП №183 не изменилось: «на согласовании», номер `782`, `paid_at` пустой.
- Число строк `KP_offers` и `counterparties` равно числу, снятому до сборки. На утро 3 ноября это было 181 и 4406.
- Оба каталога обмена смонтированы как раньше. Новых пробных `new_*.json` и `otvet_*.json` нет.
- Health 200, лимит тела 50 МБ.

## Открытые вопросы

Нет.
