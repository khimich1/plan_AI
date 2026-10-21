# Spec: Контакты менеджера в PDF и XLSX из архива

> **Фаза SDD:** SPECIFY  
> **Статус:** спека до кода. Человек ещё не принимал этот файл как разрешение писать код.  
> **Дата:** 2026-10-21  
> **Идея:** [`../ideas/archive-pdf-manager-contacts.md`](../ideas/archive-pdf-manager-contacts.md)  
> **План:** [`../develop/plans/2026-10-21-archive-pdf-manager-contacts.md`](../develop/plans/2026-10-21-archive-pdf-manager-contacts.md)

---

## Assumptions I'm Making

1. Совпадение ФИО — точное, после `strip` по краям и `casefold`, как в `CommercialDraftLifecycle.resolve_manager_for_hydrate` (`app/services/commercial_draft_lifecycle.py`, строки 899–911). Внутренние пробелы не схлопываются. «Зубов Алексей Юрьевич» совпадает с той же строкой в `managers.fio`. Другое написание — это несовпадение.
2. При нескольких карточках с одним и тем же `casefold`-ФИО берётся первая в порядке `get_all_managers`: `ORDER BY fio` (`core/kp_db_managers.py`, строка 83). Новый SQL не добавляется.
3. Телефон в шапке форматирует уже существующий `format_phone` (`core/commercial_pricing.py`, строки 445–453). Для `79621872265` это `+7 (962) 187-22-65`. `format_phone` не меняется.
4. В `KP_offers` есть только `manager_name` (`core/kp_db_schema.py`, строки 32–44). Колонок телефона и почты у КП нет и не будет.
5. Пустое и пробельное имя — это «имени нет». Заводские контакты в этом случае остаются.
6. Имя есть, а в справочнике карточки нет — строка контактов пустая. Заводской номер не подставляется. То же самое, если карточка нашлась, но и телефон, и почта пустые.
7. Заполнен только телефон или только почта — печатается имеющееся поле. Второе поле не требуется.
8. Переименование менеджера после сохранения КП оставляет в КП старое ФИО. Поиск не найдёт карточку, шапка станет пустой, а не заводской. Это принятый риск, не повод хранить снимок контактов.
9. Мастер и `offers_service` в этой задаче не меняются. У мастера несовпавшее ФИО по-прежнему может дать заводскую шапку, потому что `resolve_manager_for_hydrate` кладёт пустые телефон и почту, а экспорт не передаёт флаг подавления.
10. Проверка не привязана к живому КП №11 и не открывает рабочую `plita.db`. Фикстура повторяет карточку Зубова в памяти.

→ Если что-то из этого неверно — править spec до кода.

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| D1 | Кто читает файл | И клиент, и менеджер. Один файл, одна шапка |
| D2 | Что считать успехом | В шапке есть и телефон, и почта найденного менеджера. Подпись остаётся его ФИО |
| D3 | Свежесть | Поиск по справочнику в момент скачивания. В КП телефон и почта не пишутся |
| D4 | Где | `ArchiveService.generate_document` для `pdf`, `xlsx` и `xlsx_delivery_in_unit` |
| D5 | Где не трогаем | Мастер (`commercial_export_service.py`), `offers_service.generate_pdf` / `generate_xlsx`, фронт, маршрут скачивания, схема `schema` |
| D6 | Имя есть, карточки нет | Пустая строка контактов. Не `8 (4852) 595 000` и не `info@zhbk.ru` |
| D7 | Имени нет | Заводские контакты, как сейчас |
| D8 | Частичная карточка | Печатать только заполненное поле. Завод не подмешивать |
| D9 | Как погасить завод | Новый аргумент `suppress_company_contacts: bool = False` в обоих генераторах. Архив ставит `True`, только если после `strip` имя менеджера непустое. Остальные вызовы аргумент не передают |
| D10 | Чего не переиспользовать | `resolve_manager_for_hydrate`. При непустом несовпавшем имени он возвращает `(1, "", "")` — id 1, не контакты найденной карточки (`commercial_draft_lifecycle.py`, строка 913) |

---

## Objective

При скачивании КП из архива шапка PDF и XLSX показывает текущие телефон и почту менеджера, чьё ФИО записано в КП. Клиент и менеджер видят в файле того, кто его подписал.

Сейчас `ArchiveService.generate_document` передаёт `manager_phone=None` и `manager_email=None` и в PDF (`app/services/archive_service.py`, строки 1006–1007), и в XLSX, включая `xlsx_delivery_in_unit` (строки 1031–1032). Оба генератора при пустой паре печатают завод: `Тел.: 8 (4852) 595 000, Email: info@zhbk.ru` (`core/commercial_offer.py`, строки 513–521; `core/commercial_offer_xlsx.py`, строки 251–259). Константы — `COMPANY_PHONE` / `COMPANY_EMAIL`, строки 47–48 в обоих файлах.

Для КП №11 в базе: заказчик ООО «АВТОВИД», `manager_name` «Зубов Алексей Юрьевич». В `managers` id=2: то же ФИО, `contact_number` `79621872265`, `email` `zubov@gbkstart.ru`. Ожидаемый фрагмент шапки: `Тел.: +7 (962) 187-22-65, email: zubov@gbkstart.ru`. Подпись остаётся «Зубов Алексей Юрьевич» (`commercial_offer.py`, строка 811; `commercial_offer_xlsx.py`, строка 650).

Успех: совпавшее ФИО даёт отформатированный телефон и почту в PDF и в XLSX; подпись не меняется; несовпавшее ФИО не показывает завод; пустое имя по-прежнему показывает завод; мастер не меняется; колонок в БД не добавляется; контакты читаются при скачивании.

### Пользователь

Менеджер или администратор скачивает файл из архива. Тот же файл уходит клиенту. Отдельной роли и отдельного клиента нет. «Скай» — формулировка разговора, в локальной базе такого заказчика нет. Ветки под одного клиента нет.

### Сценарии

1. **Совпавшее ФИО, оба контакта.** Скачивание PDF, XLSX и XLSX «доставка в цене» ставит в шапку `Тел.: +7 (962) 187-22-65, email: zubov@gbkstart.ru`. Подпись — ФИО из КП. Заводского телефона и `info@zhbk.ru` в файле нет.
2. **Только телефон или только почта.** В шапке только заполненное поле: `Тел.: …` либо `Email: …`. Второе поле и завод не добавляются.
3. **ФИО есть, карточки нет.** Строка контактов пустая. Завода нет. Подпись — ФИО из КП, как сейчас.
4. **ФИО есть, карточка есть, оба контакта пустые.** То же, что сценарий 3: пустая строка, без завода.
5. **Имени нет** (`NULL`, `""` или одни пробелы). Шапка заводская: `Тел.: 8 (4852) 595 000, Email: info@zhbk.ru`. Подпись по-прежнему запасное «Менеджер», если имени нет. Это уже делает генератор (`manager_name or "Менеджер"`), здесь не меняется.
6. **Мастер.** Экспорт по-прежнему берёт `metadata["manager_phone"]` и `metadata["manager_email"]` и не передаёт `suppress_company_contacts`. Пустая пара в мастере по-прежнему завод.
7. **Старый offers API.** `offers_service.generate_pdf` / `generate_xlsx` по-прежнему передают `None` / `None` (строки 195–196 и 219–220) и не получают флаг. Поведение старого API не чинится этой спекой.

---

## Tech Stack

Тот же контур, без новых зависимостей: Python 3, FastAPI, существующие ReportLab и openpyxl, pytest. Фронт не меняется.

Слои: `ArchiveService.generate_document` читает справочник и вызывает уже существующие `generate_commercial_offer_pdf` и `generate_commercial_offer_xlsx`. Маршрут `GET /api/v1/commercial/archive/{kp_id}/files/{kind}` (`app/api/v1/endpoints/archive.py`, `download_archive_document`, строки 510–524) остаётся. Фронт уже ходит туда: `buildDocumentUrl` в `frontend/src/features/commercial-archive/api/archiveApi.ts`, строки 151–152.

## Commands

Команды для будущей реализации. Сейчас их не гонять как приёмку: кода ещё нет.

```bash
.venv/bin/python -m pytest \
  tests/test_commercial_offer_header_contacts.py \
  tests/test_archive_manager_contact_lookup.py \
  tests/test_archive_document_manager_contacts.py \
  tests/test_archive_service.py::test_generate_document_pdf \
  tests/test_archive_pile.py \
  tests/test_archive_march.py -q
```

`test_archive_service.py::test_generate_document_pdf`, `test_archive_pile.py` и `test_archive_march.py` — регресс: прежние вызовы `generate_document` не должны падать из-за чтения справочника и не должны создавать sqlite-файл.

## Project Structure

```
app/services/manager_contact_lookup.py   # чистый поиск контактов и загрузка справочника, если файл БД есть
app/services/archive_service.py          # pdf / xlsx / xlsx_delivery_in_unit передают найденные контакты и флаг
core/commercial_offer.py                 # suppress_company_contacts в PDF
core/commercial_offer_xlsx.py            # тот же флаг в XLSX
tests/test_commercial_offer_header_contacts.py
tests/test_archive_manager_contact_lookup.py
tests/test_archive_document_manager_contacts.py
```

Не менять: `commercial_draft_lifecycle.py` (включая `resolve_manager_for_hydrate`), `commercial_export_service.py`, `file_generation_service.py`, `offers_service.py`, `app/api/v1/endpoints/archive.py`, фронт, схему `KP_offers`.

`FileGenerationService.generate_offer_pdf` / `generate_offer_xlsx` уже прокидывают `**kwargs` (строки 35–45). Новый аргумент с умолчанием `False` до мастера не доходит, пока мастер его не передаёт. Мастер не начинает его передавать.

## Механизм шапки

Сейчас пустая пара всегда означает завод. Этого недостаточно: архиву нужна пустая пара без завода, а мастеру и старому API — прежний завод. Поэтому генераторы получают явный флаг, а не «пустое значение как-нибудь».

Аргумент добавляется **в конец** обеих сигнатур, с умолчанием `False`, чтобы текущие позиционные и именные вызовы не сдвинулись:

```python
def generate_commercial_offer_pdf(
    order_data: List[Dict],
    offer_number: str,
    offer_date: str,
    customer_name: Optional[str] = None,
    manager_name: Optional[str] = None,
    manager_phone: Optional[str] = None,
    manager_email: Optional[str] = None,
    # ...прежние аргументы без изменения порядка...
    long_pile_delivery: Optional[Dict] = None,
    suppress_company_contacts: bool = False,
) -> io.BytesIO:
    phone = (manager_phone or "").strip()
    email = (manager_email or "").strip()
    if phone and email:
        contacts_text = f"Тел.: {format_phone(phone)}, email: {email}"
    elif phone:
        contacts_text = f"Тел.: {format_phone(phone)}"
    elif email:
        contacts_text = f"Email: {email}"
    elif suppress_company_contacts:
        contacts_text = ""
    else:
        contacts_text = f"Тел.: {COMPANY_PHONE}, Email: {COMPANY_EMAIL}"
```

В XLSX та же развилка для `contacts_line`. Регистр букв не выдумывать: у менеджера `email:` со строчной, у завода `Email:` с прописной — как в текущих строках 513–521 и 251–259.

`strip` нужен, чтобы строка из пробелов не считалась телефоном и не печатала `Тел.:` с пустым номером. Для обычных значений мастера поведение веток `if / elif` не меняется.

Пустая строка остаётся в том же слоте шапки: PDF по-прежнему добавляет абзац, XLSX по-прежнему пишет ячейку. Слот не удаляется, чтобы блок «адрес → контакты → номер КП» не съезжал. В слоте нет заводских констант.

Флаг не влияет на ветки, где телефон или почта непустые. Он меняет только ветку «оба пустые».

Подпись не собирается из телефона. Строки 809–811 и 649–650 не меняются: в подписи имя, контакты только в шапке.

### Поиск для архива

Отдельная чистая функция, не метод жизненного цикла черновика:

```python
def lookup_manager_contacts(
    manager_name: str | None,
    managers: list[dict],
) -> tuple[str, str]:
    name = (manager_name or "").strip()
    if not name:
        return "", ""
    needle = name.casefold()
    for manager in managers:
        fio = str(manager.get("fio") or "").strip()
        if fio.casefold() != needle:
            continue
        return (
            str(manager.get("contact_number") or "").strip(),
            str(manager.get("email") or "").strip(),
        )
    return "", ""


def header_contacts_for_archive(
    manager_name: str | None,
    managers: list[dict],
) -> tuple[str | None, str | None, bool]:
    name = (manager_name or "").strip()
    if not name:
        return None, None, False
    phone, email = lookup_manager_contacts(name, managers)
    return phone or None, email or None, True
```

Возврат — только телефон и почта. `id` не возвращается. Несовпадение не подставляет карточку id=1.

`load_managers_if_db_exists(db_path)` читает `ManagerRepository(db_path).list_managers()` только если `db_path` — `str` или `Path` и этот путь является файлом. Иначе возвращает `[]` и не вызывает `sqlite.connect`. Так тесты с `MagicMock` и с несуществующим `tmp_path / "plita.db"` не создают базу. В бою файл есть: строка КП только что прочитана из него. Отсутствующая таблица `managers` в существующем файле не глотается.

`generate_document` один раз, до ветки `kind`, считает тройку и передаёт её и в PDF, и в XLSX. `xlsx_delivery_in_unit` остаётся в той же ветке `kind in {"xlsx", "xlsx_delivery_in_unit"}` и получает те же контакты. `embed_delivery_in_unit_price` не меняется. Ветка `schema` контакты не получает.

Архив передаёт в генератор сырой `contact_number` (`79621872265`), не заранее отформатированную строку. Формат делает `format_phone` внутри генератора.

## Code Style

Имена и типы как в соседнем коде: `manager_phone`, `manager_email`, `Optional[str]`, ранний `strip`, без нового слоя репозитория. Поиск — функция модуля, сервис архива её вызывает. Комментарий у новой ветки короткий и по-русски, в духе уже стоящего «Fallback: контакты компании, если нет контактов менеджера».

Флаг не прячется в пустую строку-сентинель. Пустая строка и `None` по-прежнему значат «контакта нет». Смысл «не подставлять завод» несёт только `suppress_company_contacts`.

## Testing Strategy

pytest, временные данные в памяти, рабочую `plita.db` не открывать. Сначала тесты на развилку генератора и на поиск, затем подключение в `generate_document`. Уровень — модульные тесты. HTTP и фронт не нужны: маршрут уже вызывает `generate_document`.

Текст PDF читать так же, как `tests/test_long_pile_delivery.py` (`pdftotext -layout`). Текст XLSX собирать по всем ячейкам листа «КП», без жёсткого номера строки: логотип сдвигает шапку. Подпись в XLSX проверять по ячейке следующей строки после «С уважением,»: там ровно `manager_name`, без телефона и почты.

Обязательные случаи перечислены в Success Criteria. Имена тестов — в плане.

## Boundaries

- Always: считать контакты один раз на скачивание и передать их в PDF и в оба вида XLSX. Гасить завод только флагом и только когда имя менеджера непустое. Сохранить печать одного заполненного поля. Не создавать sqlite-файл, если пути к базе нет. Прогонять точечный pytest из Commands.
- Ask first: хранить телефон и почту на КП, менять мастер, чинить `offers_service`, печатать в шапке и завод, и менеджера, ослаблять точное совпадение ФИО до нечёткого поиска, менять `format_phone`.
- Never: вызывать `resolve_manager_for_hydrate` для шапки документа. Возвращать из поиска id=1 или контакты Александра при чужом ФИО. Менять подпись. Добавлять колонки. Делать ветку под одного заказчика. Коммитить в рамках этой спеки.

## Success Criteria

| Требование | Критерий |
|---|---|
| Совпавший менеджер, PDF | В тексте есть `Тел.: +7 (962) 187-22-65, email: zubov@gbkstart.ru`. Нет `8 (4852) 595 000` и `info@zhbk.ru`. ФИО подписи на месте, телефон в тексте один раз |
| Совпавший менеджер, XLSX | Та же строка шапки на листе «КП». Ячейка под «С уважением,» равна ФИО и не содержит телефон или почту |
| `xlsx_delivery_in_unit` | Тот же телефон, та же почта и `suppress_company_contacts=True`, что у обычного XLSX |
| Только телефон | Шапка `Тел.: +7 (962) 187-22-65`. Нет `email:`, нет `info@zhbk.ru`, нет заводского телефона |
| Только почта | Шапка `Email: zubov@gbkstart.ru`. Нет `Тел.:`, нет заводских контактов |
| ФИО есть, карточки нет | `suppress_company_contacts=True`, телефон и почта пустые. В файле нет заводских контактов. Подпись — это ФИО |
| Оба контакта карточки пустые | То же подавление завода |
| Имени нет | Флаг `False`, генератор печатает `Тел.: 8 (4852) 595 000, Email: info@zhbk.ru` |
| Мастер | `commercial_export_service.py` не передаёт флаг. Вызов генератора без флага и с пустой парой по-прежнему печатает завод |
| Старый offers API | `offers_service.py` не изменён |
| База | Новых колонок нет. Контакты не пишутся в `KP_offers` |
| Свежесть | Каждый вызов `generate_document` читает справочник заново |
| Поиск | Несовпавшее имя не возвращает контакты другой карточки, в том числе id=1 |
| Регресс тестов архива | `test_generate_document_pdf` и генерация в pile/march не создают файл базы и остаются зелёными |

## Open Questions

Нет. Пустая строка при несовпавшем ФИО зафиксирована. Частичная карточка печатает имеющееся поле.
