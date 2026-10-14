# Аудит: экономист

**Дата**: 2026-10-08
**Скоуп**: стол прайсов, роль economist, выгрузка 1С и задачи GUID со страницы «Прайсы и 1С»
**Реестр**: [FINDINGS.md](./FINDINGS.md)
**Прогон 0**: —

## Дельта

| ID | Было | Стало | Суть |
|----|------|-------|------|
| A29 | — | open | Составные пишутся в `composite_pile_prices`, UI показывает сырой ключ `composite_pile` |
| A30 | — | open | Сырой sqlite3 и f-string SQL в `PriceDeskService` |
| A31 | — | open | «Ввести цену»: два коммита, обрыв оставляет открытую задачу |
| Q40 | — | open | ЛМ, ступени и цельные сваи пишут цену ≤ 0 |
| Q41 | — | open | Имя «цельные» отсекает только марки с «ФБС», не ступени/мостовые |
| Q42 | — | open | Гонка двух превью показывает дифф не от того файла; apply ловит 409 |
| S8 | open | open | Тот же ID: `import-1c` отдаёт `detail=str(exc)`. Стол прайсов сам не течёт |
| S9 | open | open | Тот же ID: CSRF читает multipart до токена; бьёт preview/apply/import-1c |
| S17 | open | open | Тот же ID: Excel без magic bytes и без потолка строк |

- Закрыто: 0
- Открыто (подтверждено): 3
- Новое: 6
- Не воспроизвелось: 0

**Открытые P0 / P1 в скоупе**: 0 / 2

## Действия (максимум 8)

1. Q40 — не писать цену ≤ 0 для ЛМ, ступеней и цельных свай; тест на ноль.
2. Q41 — для имени «цельные» отклонять чужие группы, не только «ФБС».
3. Q42 — не применять ответ устаревшего превью.
4. A31 — одна транзакция «Ввести цену».
5. A30 — убрать сырой SQL из PriceDeskService.
6. A29 — подпись «Составные сваи» вместо ключа composite_pile.
7. S8 — не отдавать str(exc) с import-1c.
8. S17 — magic bytes и потолок строк на Excel стола и выгрузки 1С. (S9 — общий CSRF, не чинить только в этом модуле.)

## Открытые P0 / P1

### [Q40] ЛМ, ступени и цельные сваи пишут цену ≤ 0

**Статус**: open
**Улика**: нет фильтра `price <= 0` в `core/march_price_db.py:153-157`, `core/step_price_db.py:104-108` и `:142-145`, `core/pile_price_db.py:163-167`. Вызов: `app/services/price_desk_service.py:254-262`. Контраст: `core/fbs_price_db.py:169-170`, `core/plate_nikita_parser.py:124-125` отбрасывают цену ≤ 0. Отдача нуля: `core/commercial_pricing.py:163-168`, `:177-180`, `:183-190`. Теста на ноль для этих трёх групп нет.
**Зачем**: High P1. Нулевая и отрицательная цена ЛМ, ступеней и цельных свай попадает в прайс и затем в расчёт КП. ФБС и парсер Никиты такой ввод уже режут.

### [Q41] Имя «цельные» отсекает только марки с «ФБС»

**Статус**: open
**Улика**: `core/pile_price_db.py:83`, `app/services/price_desk_service.py:205-208`. Тест покрывает только ФБС: `test_pile_filename_with_fbs_marks_does_not_write_piles`. Ступени и мостовые с буквой «С» пишутся в `pile_prices`.
**Зачем**: High P1. Файл с именем «цельные» принимает чужие группы (ступени, мостовые) в таблицу цельных свай. Отсечение по подстроке «ФБС» этого не закрывает.

## Приложение

Medium. Подтверждённые S8, S9, S17 остаются Medium. Что проверено и не выносится — ниже. Без Health Score. Без «Start remediation?».

### [A29] Сырой ключ `composite_pile`

**Статус**: open Medium
**Улика**: `core/price_desk_classify.py:36-37`; `app/services/price_desk_service.py:62`, `:209-213`, `:263-265`; `app/schemas/price_desk.py:13-20`; `frontend/src/features/price-desk/types/priceDesk.ts:3-11`, `:59-63`; `frontend/src/features/price-desk/components/PricesView.tsx:129`; `frontend/src/features/nomenclature/types/nomenclature.ts:120-127`.
**Зачем**: составные пишутся в `composite_pile_prices`, UI показывает сырой ключ `composite_pile`. Подпись «Составные сваи» в карте видов стола.

### [A30] Сырой sqlite3 и f-string SQL

**Статус**: open Medium
**Улика**: `app/services/price_desk_service.py:270-301`. Имена таблиц из закрытого словаря.
**Зачем**: list/meta рядом с import_*, убрать sqlite3 из сервиса.

### [A31] «Ввести цену» — два коммита

**Статус**: open Medium
**Улика**: `app/services/nomenclature_queue_service.py:102-122`; `core/price_import_queue.py:231-232`; `frontend/src/features/price-desk/components/TasksSection.tsx:21-30`.
**Зачем**: обрыв между коммитами оставляет открытую задачу. Одна транзакция на цену, GUID и set_resolved.

### [Q42] Гонка двух превью

**Статус**: open Medium
**Улика**: `frontend/src/features/price-desk/components/PricesView.tsx:58-71`, `:90-96`, `:200`. Apply чужих байт даёт 409.
**Зачем**: дифф на экране может быть не от того файла. Не применять ответ устаревшего превью.

### [S8] `import-1c` отдаёт `detail=str(exc)`

**Статус**: open Medium (тот же ID, severity не повышалась)
**Улика**: `app/api/v1/endpoints/nomenclature.py:63-73`; `core/pricelist_1c_parser.py:60-65`, `:124-128`; `app/services/nomenclature_import_service.py:100-101`, `:120-121`, `:217-223`. Стол прайсов сам не течёт: `app/api/v1/endpoints/price_desk.py:34-42`, `:55-62` — PriceDeskError + raise_unexpected_server_error.
**Зачем**: утечка текста исключения с выгрузки 1С. Чинить `import-1c`, не стол прайсов.

### [S9] CSRF читает multipart до токена

**Статус**: open Medium (тот же ID, severity не повышалась)
**Улика**: `app/middleware/csrf.py:36-47`; бьёт `price_desk.py:25-31`, `:45-52` и `nomenclature.py:46-55`. FE шлёт заголовок `frontend/src/shared/api/httpClient.ts:130-132`.
**Зачем**: общий CSRF middleware, не чинить только в модуле экономиста.

### [S17] Excel без magic bytes и без потолка строк

**Статус**: open Medium (тот же ID, severity не повышалась)
**Улика**: суффикс только `app/services/price_desk_service.py:114-117`; `pd.read_excel` без nrows — `core/plate_nikita_parser.py:59` и group price DBs; `core/pricelist_1c_parser.py:79`; `core/universal_report_1c_parser.py:57`, `:77`. Потолок размера есть через `read_upload_file_capped`.
**Зачем**: magic bytes и потолок строк на Excel стола и выгрузки 1С.

### Проверено и не выносится

- Один парсер по имени.
- Нет якоря → 422.
- Preview не пишет.
- Чужой sha256 → 409.
- REQUIRE_PRICES на `/prices` и на nomenclature этой страницы.
- Economist не admin.
- Никита отбрасывает нагрузки 16/21 и цену ≤ 0.
- Apply `INSERT OR REPLACE` не удаляет missing и не трогает `unit_price` сохранённых КП.
- Выгрузка 1С — другой endpoint.
- Маршруты КП, производства, ГСМ и логистики экономисту закрыты.
- Своего in-process rate limiter нет (A2/S3 не переоткрывались).
