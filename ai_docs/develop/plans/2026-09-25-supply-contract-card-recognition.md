# План: Распознавание карточки договора

**Дата:** 2026-09-25  
**Orchestration:** orch-2026-09-25-20-15-card  
**Статус:** план, код по этой спеке не писать до приёмки  
**Идея:** `ai_docs/ideas/supply-contract-card-recognition.md`  
**Спека:** `ai_docs/specs/supply-contract-card-recognition.md`  
**Рядом:** `ai_docs/specs/supply-contract.md` (раздел «Распознавание карточки», D13). Спеку не переписывать.

## Overview

Менеджер на уже существующем экране договора загружает карточку покупателя кнопкой или перетаскиванием и видит поля, которым можно верить. Word, Excel (xlsx) и PDF с текстом разбираются локально, без GigaChat. Фото и скан без текста идут в один вызов `create_ocr_provider()`, затем в тот же контроль цифр; если контроль не прошёл, один повтор только по этим полям. Неверная цифра в `fields` не попадает. Номер договора эта работа не выдаёт.

## Architecture Decisions

- **Контроль цифр — одно место.** Функции `accept_*` в `core/supply_contract_card_parse.py`. Они не знают, пришёл текст из Word, xlsx или из модели. Пустой результат значит «в `fields` не класть», имя поля — в `doubtful`. `normalize_digits` остаётся в `core/supply_contract.py`.
- **Неверная цифра хуже пустого поля.** Длина или контрольная сумма не сошлись — значения в `fields` нет, имя есть в `doubtful`.
- **КПП — только 9 цифр.** Контрольной суммы нет.
- **Корсчёт — только 20 цифр.** Ключ вместе с БИК покупателя не считается.
- **Расчётный счёт.** Если БИК принят — контрольный ключ пары (Положение Банка России № 579-П): условный номер = три последние цифры БИК + 20 цифр счёта, коэффициенты `7, 1, 3` по кругу, сумма произведений делится на 10 без остатка. Без принятого БИК счёт верной длины остаётся в `fields.account` и попадает в `doubtful`. Неверная длина — поля нет, имя в `doubtful`.
- **БИК — 9 цифр.** Отдельного контрольного разряда у номера БИК положение ЦБ не задаёт, самодельные веса не добавлять. Иная длина — поля нет, имя в `doubtful`. Принятый БИК (ровно 9 цифр) идёт в ключ расчётного счёта.
- **ИНН.** 10 цифр: веса `2,4,10,3,5,9,4,6,8`, `(сумма % 11) % 10` равна последней цифре. 12 цифр: 11-я — веса `7,2,4,10,3,5,9,4,6,8` по первым десяти; 12-я — веса `3,7,2,4,10,3,5,9,4,6,8` по первым одиннадцати. Иначе поля нет.
- **ОГРН и ОГРНИП — одно поле формы `ogrn`.** 13 цифр: `int(первые 12) % 11 % 10` равно 13-й. 15 цифр: `int(первые 14) % 13 % 10` равно 15-й. Иная длина — поля нет.
- **Текст без зрения.** `SupplyContractService._parse_upload`: docx через `extract_docx_text`, xlsx через новый `extract_xlsx_text`, pdf с `has_text_layer` через `extract_pdf_text`. `complete_vision_json` на этих ветках не вызывается.
- **Excel — уже стоящий openpyxl.** `openpyxl.load_workbook` из `requirements.txt` (`openpyxl>=3.0.0`), тот же вызов, что в `parse_lawyer_sheet` (`app/services/supply_contract_import.py`): `BytesIO`, `read_only=True`, `data_only=True`. Новую зависимость не добавлять. Старый `.xls` не читать.
- **Зрение.** Фото (`_is_image`) и pdf без текстового слоя: один `create_ocr_provider()` из `core/ocr/pipeline.py`, затем `complete_vision_json`. Повтор — один и только по цифровым полям, которые в `fields` не попали. Прошедшее поле ответ повтора не затирает. Второго полного прохода нет: `get_contract_card_verify_prompt` по всей карточке и `apply_contract_verify`, который подменяет все поля, для этого пути не вызывать.
- **Несколько расчётных счетов.** Список `accounts`. Один прошедший счёт можно положить в `fields.account`. Два и больше — `fields.account` пустой, пока менеджер не выбрал. Первый счёт молча не подставлять.
- **Нет реквизитов.** После контроля нет ни ИНН, ни ОГРН, ни счёта, ни БИК — `SupplyContractValidationError` с текстом «это не карточка контрагента». Форма полей не заполняется. Номер не создаётся.
- **Файл только на экране договора.** Кнопка и перетаскивание внутри `SupplyContractDrawer`, не на всё окно программы. Несколько файлов в одном броске — только первый, как `event.dataTransfer.files?.[0]` в `PricesView`. Рамка зоны — как в `PricesView` и `TransactionsImportDialog` (`onDragEnter` / `onDragOver` / `onDragLeave` / `onDrop`, пунктир). `addFiles` из диалога ГСМ, который берёт все файлы, не копировать.
- **Живой API в тестах не вызывать.** Карточки из `банк знаний/варианты реквизитов` в git не класть. Фикстуры — короткие строки и минимальный docx/xlsx/png.
- **Новый сервис распознавания не заводить.** Разбор остаётся в `parse_card_text` и `recognize_contract_card`.

## Dependency graph

```
accept_* (длина и суммы)
    │
    ├── локальный разбор: слэш, банк, КФХ/ИП, основание, accounts
    │       │
    │       ├── xlsx через openpyxl в _parse_upload
    │       │       │
    │       │       └── отказ «это не карточка контрагента»
    │       │               │
    │       │               ├── точечный повтор зрения
    │       │               └── список счетов и зона файла на экране договора
    │       └── accounts в схеме ответа
    │
    └── тот же accept_* после Extract
```

## Task List

### Phase 1 — Контроль цифр

- [x] **CR-001: Контроль цифр**
  - **Type:** feat-be
  - **Description:** В `core/supply_contract_card_parse.py` функции приёма цифр: ИНН, КПП, ОГРН/ОГРНИП, БИК, расчётный счёт, корсчёт. Правила — в Architecture Decisions. `parse_card_text` кладёт в поле только принятое значение; отказ даёт `None` и имя в `doubtful`. Счёт без принятого БИК при длине 20 остаётся и подсвечивается.
  - **Acceptance:**
    - [x] ИНН с неверной суммой отсутствует в `fields` и есть в `doubtful`
    - [x] КПП не из 9 цифр и корсчёт не из 20 цифр в `fields` не попадают
    - [x] Строка `407028105771000149` (18 цифр) не становится `fields.account`
    - [x] Счёт из 20 цифр без БИК остаётся в `fields.account` и есть в `doubtful`
    - [x] Счёт из 20 цифр при принятом БИК с несошедшимся ключом в `fields` не попадает
  - **Verify:** `pytest tests/test_supply_contract_card_parse.py -q -k checksum`
  - **Dependencies:** —
  - **Files:** `core/supply_contract_card_parse.py`, `tests/test_supply_contract_card_parse.py`
  - **Scope:** M

### Checkpoint: Phase 1

- [x] Отказ цифры зелёный без HTTP и без зрения
- [x] Старые фикстуры с заведомо неверными ИНН/ОГРН/счётом больше не утверждают, что такая цифра лежит в `fields`

### Phase 2 — Локальный разбор

- [x] **CR-002: Макеты текста и список счетов**
  - **Type:** feat-be
  - **Description:** Дописать `parse_card_text`, не заменяя его промптом марок. Строка «ИНН / КПП» с двумя числами через слэш раскладывается на `inn` и `kpp`. ОГРН читается и из шапки «ИНН … КПП … ОГРН …», и из отдельной строки. Банк — строка с отделением или словом «банк», не только шаблон «Банк:» в `_bank_name`. Подписант: директор, генеральный директор, индивидуальный предприниматель, глава КФХ (`_signatory`). Если в файле и устав, и доверенность другого лица, `_authority` берёт основание того, чьё ФИО записано в `signatory_name`. Все расчётные счета — в `accounts` у `CardParseResult.as_dict`. Один принятый счёт можно положить в `fields.account`. Два и больше — `fields.account` пустой. `SupplyContractParseOut` в `app/schemas/supply_contract.py` отдаёт `accounts`, иначе `response_model` список срежет.
  - **Acceptance:**
    - [x] «ИНН/КПП» через слэш не кладёт ИНН в КПП
    - [x] Название отделения без слова «банк» попадает в `bank_name`
    - [x] Глава КФХ и ИП заполняют должность и ФИО
    - [x] Основание берётся у выбранного подписанта, а не у чужой доверенности
    - [x] Два счёта лежат в `accounts`, `fields.account` пустой
  - **Verify:** `pytest tests/test_supply_contract_card_parse.py -q -k layout`
  - **Dependencies:** CR-001
  - **Files:** `core/supply_contract_card_parse.py`, `app/schemas/supply_contract.py`, `tests/test_supply_contract_card_parse.py`
  - **Scope:** M

- [x] **CR-003: Excel без GigaChat**
  - **Type:** feat-be
  - **Description:** `extract_xlsx_text` читает xlsx через `openpyxl.load_workbook`, как `parse_lawyer_sheet`. Текст ячеек склеивается так же, как абзацы и ячейки в `extract_docx_text`, и уходит в `parse_card_text`. В `_parse_upload` ветка xlsx стоит рядом с docx и не вызывает `_recognize_image`. Отличить xlsx от docx: zip и `xl/workbook.xml`, не `word/document.xml` (`is_docx` уже требует `word/document.xml`). `.xls` по-прежнему даёт `MSG_UNSUPPORTED_CARD`.
  - **Acceptance:**
    - [x] xlsx с ИНН и счётом заполняет те же поля, что docx, и не вызывает зрение
    - [x] docx и pdf с текстом по-прежнему не вызывают `complete_vision_json`
    - [x] Файл `.xls` не принимается и новую библиотеку не тянет
  - **Verify:** `pytest tests/test_supply_contract_card_parse.py -q -k xlsx`
  - **Dependencies:** CR-002
  - **Files:** `core/supply_contract_card_parse.py`, `app/services/supply_contract_service.py`, `tests/test_supply_contract_card_parse.py`
  - **Scope:** S

- [x] **CR-004: Это не карточка контрагента**
  - **Type:** feat-be
  - **Description:** После локального разбора и после зрения, если в принятых полях нет ни `inn`, ни `ogrn`, ни `account`, ни `bik`, `_parse_upload` бросает `SupplyContractValidationError` «это не карточка контрагента». Существующий `parse_archive_supply_contract` уже отдаёт её как 400. Строка `supply_contract` не создаётся. Пустой docx, который сейчас возвращает пустые поля и `doubtful`, становится этой ошибкой.
  - **Acceptance:**
    - [x] Текст без ИНН, ОГРН, счёта и БИК — сообщение «это не карточка контрагента»
    - [x] Договор при этом не создаётся
    - [x] Карточка, где есть хотя бы один из четырёх реквизитов, ошибкой отказа не закрывается
  - **Verify:** `pytest tests/test_supply_contract_card_parse.py -q -k not_a_card`
  - **Dependencies:** CR-003
  - **Files:** `app/services/supply_contract_service.py`, `tests/test_supply_contract_card_parse.py`
  - **Scope:** S

### Checkpoint: Phase 2

- [x] Word, xlsx и текстовый PDF не вызывают зрение
- [x] Паспорт или КП без реквизитов не выглядит успешным разбором

### Phase 3 — Зрение и экран

- [x] **CR-005: Промпт и один точечный повтор**
  - **Type:** feat-be
  - **Description:** В `CONTRACT_CARD_EXTRACT_PROMPT` (`core/ocr/prompts.py`): шапка «ИНН КПП ОГРН» обязательна, даже если ОГРН нет в таблице; счёт и корсчёт копируются по цифрам, повторы `0` и `7` не сжимаются; ячейка «БИК … в <банк>» делится на БИК и название банка. `recognize_contract_card` делает один Extract через переданный провайдер. Дальше те же `accept_*`, что в CR-001. Если цифровое поле не принято, один повторный `complete_vision_json` только по этим именам. Ответ повтора пишет лишь пустые места. Поле, которое уже прошло, не меняется. Третьего вызова нет. `_default_card_provider` в `app/services/supply_contract_service.py` ходит в `create_ocr_provider()` и не гоняет полный Verify. Тесты зрения переносятся в `tests/test_supply_contract_card_ocr.py`. Провайдер и `GigaChatProvider` подменяются, как в текущем `test_png_with_ocr_provider_gigachat_uses_mock_not_openai`. Живой GigaChat и OpenAI не вызываются.
  - **Acceptance:**
    - [x] При `OCR_PROVIDER=gigachat` PNG идёт в mock, OpenAI не вызывается
    - [x] Поле, прошедшее контроль, повтор не меняет, даже если ответ повтора содержит другой ИНН
    - [x] Повтор не вызывается, если все цифровые поля уже приняты (один вызов, не два)
    - [x] Повтор вызывается один раз и только по полям, которых нет в `fields`
    - [x] Второго полного прохода по всей карточке нет
  - **Verify:** `pytest tests/test_supply_contract_card_ocr.py -q`
  - **Dependencies:** CR-004
  - **Files:** `core/supply_contract_card_ocr.py`, `core/ocr/prompts.py`, `app/services/supply_contract_service.py`, `tests/test_supply_contract_card_ocr.py`, `tests/test_supply_contract_card_parse.py`
  - **Scope:** M
  - **parallelSafe:** да, файлы не пересекаются с CR-006

- [x] **CR-006: Список счетов и файл на экране договора**
  - **Type:** ui
  - **Description:** В `SupplyContractDrawer` кнопка выбора и зона перетаскивания вызывают один `onFile`. Зона — модалка договора, не страница архива и не `window`. Подсветка, пока файл над зоной: пунктир и фон, как блок загрузки в `frontend/src/features/price-desk/components/PricesView.tsx` и `frontend/src/features/gsm/components/TransactionsImportDialog.tsx`. Из броска берётся только `files[0]`, как в `PricesView`. Несколько файлов диалога ГСМ не повторять. Ответ с `accounts` показывается списком; выбор записывает один счёт в поле формы. Пока выбора нет, поле счёта остаётся пустым. Ошибка «это не карточка контрагента» показывается текстом и не подставляет угаданные поля. Тип `accounts` — в `frontend/src/features/commercial-archive/types/supplyContract.ts`.
  - **Acceptance:**
    - [x] Список из двух счетов виден, до выбора `fields.account` в форме пустой, выбор записывает один счёт
    - [x] Drop файла на экран договора вызывает тот же `parseSupplyContract`, что кнопка
    - [x] Два файла в одном броске: запрос уходит с первым, второй не читается
    - [x] Ошибка разбора не заполняет поля формы
  - **Verify:** `cd frontend && npm run test -- src/features/commercial-archive/components/SupplyContractDrawer.test.tsx`
  - **Dependencies:** CR-004
  - **Files:** `frontend/src/features/commercial-archive/components/SupplyContractDrawer.tsx`, `frontend/src/features/commercial-archive/components/SupplyContractDrawer.test.tsx`, `frontend/src/features/commercial-archive/types/supplyContract.ts`
  - **Scope:** M
  - **parallelSafe:** да, файлы не пересекаются с CR-005

### Checkpoint: готово к реализации по срезам

- [x] CR-001…CR-006 реализованы
- [x] `cd frontend && npm run typecheck` — после CR-006
- [x] `pytest tests/test_supply_contract_card_parse.py tests/test_supply_contract_card_ocr.py -q` — после CR-005
- [x] Следующий шаг после приёмки плана — CR-001, не весь контур сразу

## Risks and Mitigations

| Risk | Impact | Mitigation |
|---|---|---|
| Текущие фикстуры (`7604010010`, `40702810000000000001`, `1027600000001`) не проходят настоящую сумму | Старые тесты начнут ждать цифру в `fields` | В CR-001 заменить ожидание: верная сумма остаётся, неверная уходит в `doubtful`. Алгоритм под фиктивные номера не подгонять |
| `test_png_with_ocr_provider_gigachat_uses_mock_not_openai` ждёт два вызова chat и полный Verify | После отмены второго полного прохода тест врёт | В CR-005 считать вызовы: один, если контроль прошёл; второй только по отказавшим полям. OpenAI по-прежнему не вызывается |
| xlsx тоже zip с `PK`, как docx | Файл Excel уедет в `extract_docx_text` или в зрение | Сначала `is_docx` по `word/document.xml`, xlsx — по `xl/workbook.xml` |
| Кривой текстовый слой PDF даст пустые поля | Менеджер не увидит цифры | Это приемлемо: в поле нет неверной цифры, номер не выдаётся |
| Повтор зрения вернёт другие цифры в уже принятое поле | Верный ИНН затрётся | Повтор пишет только поля, которых нет в `fields` |
| Два счёта, и UI подставит первый само | Менеджер подпишет не тот банк | `fields.account` пустой, пока нет явного выбора |

## Open Questions

Нет. Ограничения выше закрыты спекой и принятыми решениями. Включение выдачи номеров и лист юриста в эту работу не входят.

## Порядок выполнения

1. CR-001
2. CR-002
3. CR-003
4. CR-004
5. CR-005 и CR-006 параллельно (`parallelSafe`)
