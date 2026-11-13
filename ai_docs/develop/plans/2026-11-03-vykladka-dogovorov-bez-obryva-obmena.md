# План: выкладка договоров без обрыва обмена с 1С

**Дата:** 2026-11-03  
**Статус:** выложено 2026-11-03. Контрольные точки backend и frontend сошлись.  
**Спека:** [`ai_docs/specs/vykladka-dogovorov-bez-obryva-obmena.md`](../../specs/vykladka-dogovorov-bez-obryva-obmena.md)  
**Не это:** подмена compose и `.env`, копирование `plita.db`, правка реестра, повторная отправка КП №183, `docker compose down -v`.

## Подход

Код `968401c` уже на `origin/web`. Обмен на сервере уже живой: `app/main.py` в контейнере совпадает с этим коммитом, каталоги обмена смонтированы. Повторно их не копируем и не пересобираем ради них.

Новый экран договора на диск сервера ещё не положен. Сборка текущего `/opt/plan_zavod` его не даст. Поэтому сначала на диск кладутся только runtime-файлы из спеки, затем отдельно собирается backend и отдельно frontend. Образ backend заодно включит загрузку прайса цельных свай: файл на диске уже новый, процесс ещё старый.

Реестр и справочник не трогаем. Снимок базы нужен только как точка отката, не как источник данных.

## Компоненты и зависимости

```
локальный pytest и сверка HEAD = 968401c
        │
        └── снимок /data/plita.db и счётчики KP_offers / counterparties
                    │
                    └── 13 runtime-файлов на диск /opt/plan_zavod
                                │
                                ├── build + up backend ── blank True, autolink False, монтирования, КП 183, health
                                │
                                └── build + up frontend ── «Собрать договор», лимит 50 МБ, поиск МД-СТРОЙ
```

Копирование и сборка backend только последовательно. Frontend можно собирать после того, как файлы уже на диске, но сайт без нового backend эти кнопки не обслужит. Поэтому frontend поднимаем после контрольной точки backend.

Файлы, и только они:

- `app/api/v1/endpoints/archive.py`
- `app/repositories/supply_contract_repository.py`
- `app/schemas/archive.py`
- `app/services/supply_contract_service.py`
- `app/services/price_desk_service.py`
- `core/supply_contract_card_parse.py`
- `frontend/src/features/commercial-archive/api/archiveApi.ts`
- `frontend/src/features/commercial-archive/components/ContractCardDrawer.tsx`
- `frontend/src/features/commercial-archive/components/ContractCardList.tsx`
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`
- `frontend/src/features/commercial-archive/hooks/useArchiveQueries.ts`
- `frontend/src/features/commercial-archive/lib/contractBlank.ts`
- `frontend/src/features/commercial-archive/types/supplyContract.ts`

Тесты и `ai_docs` на сервер не копируются.

## Порядок

1. Локально убедиться, что `git rev-parse --short HEAD` равен `968401c`, и прогнать pytest из спеки.
2. На сервере снять `.backup` в `/data/plita.db.bak-before-contract-ui-20261103`. Записать число строк `KP_offers` и `counterparties`. Записать статус и `order_number_1c` КП №183.
3. Скопировать 13 файлов в те же пути `/opt/plan_zavod`. Compose, `.env`, nginx и папки обмена не копировать.
4. `docker compose -f docker-compose.split.yml build backend`, затем `up -d --no-deps backend`. Не `down`.
5. Контрольная точка backend. Если `blank` не `True`, `autolink` не `False`, монтирование обмена сбилось, health не 200 или КП №183 изменилось — frontend не собирать и реестр не трогать.
6. `build frontend`, затем `up -d --no-deps frontend`.
7. Контрольная точка сайта: «в реестре договора нет» и «Собрать договор» на КП без договора, поиск «МД-СТРОЙ», `client_max_body_size 50m`.

## Риски

| Риск | Как снижаем |
|---|---|
| Сборка старого диска выдаёт старый экран за новую выкладку | Файлы копируются до `build`. После backend в контейнере обязан быть `save_directory_blank`. |
| Пересоздание backend сотрёт обмен | Compose не подменяем. После старта `findmnt` обязан показать прежние два каталога. `down -v` не вызываем. |
| Автопривязка по ИНН исчезнет в момент старта нового backend | Это ожидаемое правило спеки. Реестр уже привязан. КП №183 не отправляем повторно. |
| Старт пересоздаст `supply_contract` | Дата и форма на сервере уже допускают пустое значение. Снимок базы всё равно снимается до сборки. |
| Сборка frontend подменит лимит 50 МБ | `nginx.docker.conf` не копируем. После сборки читаем лимит в сайте. |
| Прайс свай попадёт в тот же образ | Так и задумано: файл уже на диске, отдельный образ для него не собираем. |

## Контрольные точки

После шага 5 сайт ещё может показывать старый frontend, но backend уже без автопривязки и с сохранением бланка. КП №183 на месте, обменные диски на месте. Дальше шага 6 не идём, если эта точка не сошлась.

После шага 7 сходятся все критерии успеха спеки. Сама выкладка ждёт подтверждения этого плана.
