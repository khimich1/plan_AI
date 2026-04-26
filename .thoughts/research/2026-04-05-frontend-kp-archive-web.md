---
date: 2026-04-05
feature: frontend-kp-archive-web
---

# Frontend исследование: Web перенос «Создать КП + Архив»

## Legacy -> Web/API mapping

- `📝 Создать КП` (bot) -> блок "Создать КП" на `/web/managers` -> `POST /api/v1/commercial/generate-preview`
- `💾 Сохранить в БД` (bot) -> кнопка "Сохранить в БД (в работе)" -> `POST /api/v1/offers` (`save_mode=work`)
- `📦 В архив` (bot) -> кнопка "В архив" -> `POST /api/v1/offers` (`save_mode=archive`)
- `📁 Архив` разделы (bot) -> вкладки статусов в web -> `GET /api/v1/offers?status=...`
- `🔍 Найти по номеру КП` (bot) -> поле поиска в web -> `GET /api/v1/offers?kp_id=...`
- `Карточка КП` (bot) -> блок "Карточка КП" -> `GET /api/v1/offers/{kp_id}`
- `PDF/XLSX` (bot) -> кнопки в карточке -> `GET /api/v1/offers/{kp_id}/pdf|xlsx`
- `Изменить скидку` (bot) -> кнопка в карточке -> `PATCH /api/v1/offers/{kp_id}/discount`
- `В производство` (bot) -> кнопка в карточке -> `PATCH /api/v1/offers/{kp_id}/move-to-production`
- `Удалить КП` (bot) -> кнопка в карточке -> `DELETE /api/v1/offers/{kp_id}`

## Контракт данных

- `GET /api/v1/offers`
  - query: `status=archived|in_production|completed|all`, `limit`, `kp_id`
  - response: `{ items: OfferSummary[], count: number }`
- `GET /api/v1/offers/{kp_id}`
  - response: `OfferDetails` (`OfferSummary` + `plates[]`)
- `POST /api/v1/offers`
  - request:
    - `creation_date`, `customer_name`, `manager_name`
    - `discount_percent`, `delivery_conditions`, `payment_conditions`
    - `execution_terms_input`, `save_mode=work|archive`
    - `order_data[]`
  - response: `{ kp_id, status, execution_terms, used_default_execution_terms, offer }`
- `PATCH /api/v1/offers/{kp_id}/discount`
  - request: `{ discount_percent }`
  - response: обновленная карточка КП
- `PATCH /api/v1/offers/{kp_id}/move-to-production`
  - request: `{ execution_terms_input }`
  - response: `{ kp_id, execution_terms, used_default_execution_terms, offer }`
- `DELETE /api/v1/offers/{kp_id}`
  - response: `{ ok: true, kp_id }`
- `GET /api/v1/offers/{kp_id}/pdf|xlsx`
  - response: бинарный файл с `Content-Disposition: attachment`
- `POST /api/v1/commercial/recognize-screen`
  - request: `multipart/form-data` (`image`)
  - response: `{ recognized_text, normalized_text, lines, warnings, method, confidence }`
- `POST /api/v1/commercial/preview-check-xlsx`
  - request: `{ plates_text, recognized_text }`
  - response: бинарный `.xlsx` файл сверки (`как прислал -> распознано -> как в КП`)

## Access roles

- `GET /api/v1/offers`, `GET /api/v1/offers/{id}`, `GET pdf/xlsx`: `admin`, `manager`, `production`
- `POST /api/v1/offers`, `PATCH discount`, `PATCH move-to-production`, `DELETE`: `admin`, `manager`
- `/web/managers`: `admin`, `manager`, `production`
  - для `production` блок создания скрыт, доступен режим просмотра архива/карточек/файлов

## Backend gaps закрыты

- Добавлен API модуль `offers` с CRUD/архивными операциями.
- Добавлен сервис `OffersService`:
  - парсинг сроков (`дата / N дней / N недель / default 14 дней`)
  - генерация PDF/XLSX из данных БД
  - перенос в производство и изменение скидки
- Расширен `KpRepository` методами для операций архива.

## План UI-реализации (выполнено)

- Расширена страница `/web/managers`:
  - форма создания КП + превью
  - кнопка `Распознать СКРИН` (OCR скрина, вывод текста в отдельный блок)
  - кнопка `XLSX проверка` (bot-style preview сверка)
  - действия сохранения в `в работе` и `в архиве`
  - вкладки архивных статусов
  - поиск по номеру КП
  - карточка КП + действия (pdf/xlsx/скидка/перевод/удаление)
- Все действия выполняются через API (`fetch` с `credentials: "include"`).
- Добавлены базовые состояния интерфейса: loading/success/error/empty.

## Ручной test plan

1. Авторизация
   - Войти как `admin`.
   - Открыть `/web/managers`.
2. Создание КП -> в архив
   - Заполнить список плит + менеджер + клиент.
   - Нажать "Сгенерировать превью".
   - Нажать "В архив".
   - Проверить, что в вкладке `В архиве` появился новый КП.
3. Создание КП -> в работе
   - Создать превью.
   - Указать срок (`14 дней`).
   - Нажать "Сохранить в БД (в работе)".
   - Проверить, что КП в `В производстве` со статусом `в работе`.
4. Поиск и карточка
   - Ввести номер КП в поиск.
   - Открыть карточку, проверить позиции, сумму, статус.
5. Скачивание файлов
   - В карточке нажать `PDF`, затем `XLSX`.
   - Проверить, что оба файла скачиваются.
6. Скидка
   - Нажать "Изменить скидку", ввести `5`.
   - Проверить обновление суммы в карточке и таблице.
7. Архив -> производство
   - Для КП со статусом `в архиве` нажать "В производство".
   - Ввести срок, проверить переход в `В производстве`.
8. Удаление
   - Нажать "Удалить" и подтвердить.
   - Убедиться, что КП исчез из списка и не открывается по id.
9. Роли
   - Войти под `production`.
   - Проверить: блок создания скрыт; просмотр списка/карточки/скачивание работает; кнопки редактирования недоступны.
10. Ошибки
   - Попробовать создать КП без плит/клиента -> UI должен показать ошибку.
   - Ввести некорректную скидку (`200`) -> ошибка в UI/API.
11. OCR скрина
   - Нажать `Распознать СКРИН`, выбрать png/jpg.
   - Проверить, что распознанный текст появляется в отдельном блоке и не подставляется в textarea автоматически.
12. XLSX проверка
   - После OCR нажать `XLSX проверка`.
   - Проверить, что скачивается файл сверки с колонками `Как прислал пользователь / Распознано / Как в КП`.
