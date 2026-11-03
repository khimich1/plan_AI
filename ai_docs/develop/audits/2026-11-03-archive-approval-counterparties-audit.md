# Аудит: архив, на согласовании, контрагенты

**Дата**: 2026-11-03  
**Скоуп**: архив КП, вкладка «На согласовании», контрагенты — backend `app/api/v1/endpoints/archive.py`, `archive_service.py`, `supply_contract_service.py`, `supply_contract_import.py`, `commercial_upload_validation.py`; фронт архива (`OfferDetailsDrawer`, `ArchiveOfferList`, `SupplyContractDrawer`, `SupplyContractFields`, `CounterpartyAutocomplete`, `MoveToProductionDialog`); спека `ai_docs/specs/archive-on-approval.md`; тесты `tests/test_archive_endpoints.py`, `tests/test_supply_contract.py`, `tests/test_specification_archive.py`, `SupplyContractDrawer.test.tsx`. Производственный план, GSM, layout, бот — вне скоупа.  
**Реестр**: [FINDINGS.md](./FINDINGS.md)  
**Прогон 0**: [2026-10-01-kp-audit.md](./2026-10-01-kp-audit.md) (конструктор + архив); [2026-09-04-archive-production-audit.md](./2026-09-04-archive-production-audit.md); см. также supply-contract [2026-09-27-supply-contract-audit.md](./2026-09-27-supply-contract-audit.md).

## Дельта

| ID | Было | Стало | Суть |
|----|------|-------|------|
| Q31 | open Medium | **resolved** | Поля `poa_number` / `poa_date` при «доверенность»: `SupplyContractFields.tsx:92-110`, wiring `SupplyContractDrawer.tsx:572-576`; тест `SupplyContractDrawer.test.tsx`. |
| Q36 | open Medium | **resolved** | Единый формат даты через `formatContractDate`: `SupplyContractDrawer.tsx:349`, `SupplyContractRegistry.tsx:81`. |
| S5 | open High P1 | open High P1 | OCR/vision карточки договора без `ensure_external_ocr_enabled` и rate limit; обновлены call sites в `archive.py` и сервисе. |
| A18 | open High P1 | open High P1 | Sync soffice 60s; архивные GET PDF договора/спецификации `archive.py:810-824`, `:1103-1113`; устаревшая улика `:742` снята. |
| A31 | — | open High P1 **NEW** | Толстый HTTP `archive.py` (~1451 LOC, 53 маршрута); inline supply-contract/parse vs тонкий `list_archive_offers`. |
| A11 | open Medium | open Medium | ArchiveService **~2003** строк, **~63** метода; split-proposal нет — без действия на неделю. |
| A27 | open Medium | open Medium | `update_contract_date` `:529-548` без `_assert_writer`; HTTP-роута нет. |
| A28 | open Medium | open Medium | Прямой `SupplyContractRepository`: `archive_service.py:538-551`, `:816-818`. |
| S8 | open Medium | open Medium | `detail=str(exc)` на маршрутах архива (напр. `:209`, `:873`, `:1443`). |
| S9 | open Medium | open Medium | CSRF читает multipart до токена; загрузки архива `:176-187`, `:441`, `:784`. |
| S17 | open Medium | open Medium | XLSX договоров: `load_workbook` + unbounded `iter_rows`; вход `archive.py:196-204`, `:462-468`. |
| Q9 | open Medium | open Medium | `OfferDetailsDrawer.tsx` **1773** строки (было 1658). |
| Q12 | open Medium | open Medium | Список глотает `archive_service.py:1239-1252`; `move_to_production` рейзит `:957-961`. |
| Q18 | open Low | open Low | Partial suppressions: `MoveToProductionDialog.tsx:87`, `CounterpartyAutocomplete.tsx:151`, `:163`; severity не повышать. |
| Q23 | open Medium | open Medium | Archive half `archive_service.py:997-998` `round`; FE производства не пересматривался. |
| Q25 | open Medium | open Medium | Badge `has_delivery_schedule` без поля в list API `schemas/archive.py:29-61`. |
| Q32 | open Medium | open Medium | 0 callers verify/dead helpers; `already_exists_message` `:1000`. |
| Q33 | open Medium | open Medium | Дубли `extract_contract_fields` / `reread_contract_fields` `:1345-1371`. |
| Q34 | open Medium | open Medium | `SupplyContractDrawer.tsx` **594** строки (было 741). |
| Q35 | open Medium | open Medium | FE `status: string` vs schema enum `:13-19`. |
| Q37 | open Medium | open Medium | Только ADMIN happy-path `tests/test_supply_contract.py:390-394`; пара A27. |
| Q38 | open Medium | open Medium | Enum статуса/скана ×3 (core, schema, FE types). |
| Q39 | open Medium | open Medium | Enum kind спецификации ×4 (core, schema, types, UI). |
| Q41 | — | open Medium **NEW** | `move_to_production` требует сохранённую спецификацию сверх спеки on-approval (`paid_at` + `order_number_1c`). |
| S1 | by-design | by-design | Общий архив — не IDOR. |
| A15 | by-design | by-design | `owner_user_id` вне policy — не переоткрывать. |
| S15 | open Low | open Low | В скоупе архива sessionStorage не найден (улика — wizard `draftStorage.ts`); статус open без смены last_seen. |

- Закрыто: 2  
- Открыто (подтверждено): 20  
- Новое: 2  
- Не воспроизвелось: 0  

**Открытые P0 / P1 в скоупе**: 0 / 3  

## Действия (максимум 8)

1. **[S5]** Прогнать разбор карточки договора из архива через `ensure_external_ocr_enabled` и rate limit до vision.
2. **[A18]** Убрать синхронный soffice из HTTP-потока PDF договора и спецификации в архиве.
3. **[A31]** Разнести договор, directory и promise из `archive.py`, оставив роутер карточки КП тонким.
4. **[Q41]** Согласовать ворота спецификации в `move_to_production` со спекой `archive-on-approval.md`: либо записать в спеку, либо убрать из кода и с карточки.
5. **[S8]** На маршрутах архива не отдавать клиенту `str(exc)`.
6. **[S17]** На импорте XLSX договоров: сигнатура файла и лимит строк.
7. **[A28]** Чтение договора из ArchiveService — через SupplyContractService, не напрямую в репозиторий.
8. **[Q25]** Либо поле `has_delivery_schedule` в list API, либо убрать мёртвый бейдж.

## Открытые P0 / P1

### [S5] OCR/vision карточки договора обходит внешний гейт и rate limit
**Статус**: open  
**Улика**: `archive.py:434-447`, `:777-790`, `:914-927` → `supply_contract_service.py:449-464`, `:495-522`, `:1337-1371`; grep `ensure_external_ocr` в сервисе договора = 0. Контраст: `commercial_upload_validation.py:60-66`, `:141-142`.  
**Зачем**: При выключенном внешнем OCR коммерческая загрузка блокируется, разбор карточки из архива — нет; обход лимита и политики vision.

### [A18] Синхронный LibreOffice/soffice в HTTP (архив + GSM)
**Статус**: open  
**Улика**: `supply_contract_service.py:106`, `:905-936` (timeout 60s); вызовы из архива `archive.py:810-824`, `:1103-1113`; тест `tests/test_archive_endpoints.py:2716`. Старая привязка только к `archive.py:742` неактуальна.  
**Зачем**: Блокировка worker на до 60s на запрос PDF — DoS и деградация latency; общий дефект для GSM и архива договора/спецификации.

### [A31] Толстый роутер `archive.py`
**Статус**: open  
**Улика**: ~**1451** LOC, **53** маршрута (`@router` / `@settings_router`); контраст тонкого `list_archive_offers` `:109-116` vs inline блоки supply-contract, directory, promise, parse `:352-945`.  
**Зачем**: Смешение карточки КП, договора, справочников и promise в одном HTTP-модуле усложняет сопровождение, тестирование и единообразие с тонкими роутерами.

## Приложение

### Закрыто в этом прогоне
**[Q31]** — форма доверенности с `poa_number` / `poa_date` на FE и в drawer.  
**[Q36]** — единое форматирование даты договора в drawer и реестре.

### Medium и контекст (без P1)
**[Q41]** — лишнее ворота спецификации: `archive_service.py:936-940`; FE `OfferDetailsDrawer.tsx:117-122`; спека `archive-on-approval.md:145` (только оплата и номер 1С); тест `tests/test_specification_archive.py:508-526`.  
**[A11]** — god-orchestrator ArchiveService; действие на неделю не планируется.  
**[A27]**, **[Q37]** — запись даты договора без writer-gate; тест только ADMIN.  
**[A28]** — обход фасада договора в ArchiveService.  
**[S8]**, **[S9]**, **[S17]** — утечки exc, CSRF/multipart, unbounded XLSX (см. действия 5–6).  
**[Q9]**, **[Q12]**, **[Q25]**, **[Q32]–[Q35]**, **[Q38]**, **[Q39]** — размер drawer, swallow списка, мёртвый badge, dead OCR path, дубли vision, слабые типы, enums.  
**[Q18]** — Low: подавления deps в архивном UI; глобально не закрыто.  
**[Q23]** — `round` в archive_service vs `ceil` на FE производства (production half не открывали).

### By-design
**[S1]** — общий архив КП (`offer_access.py:26-32`).  
**[A15]** — `owner_user_id` вне policy доступа.

### S15 (без смены статуса)
**[S15]** — в скоупе commercial-archive / CounterpartyAutocomplete / commercial-offer-archive sessionStorage не найден; реестровая улика остаётся wizard `draftStorage.ts` (вне скоупа). Статус **open**, last_seen не сдвигался.

### Домен держится (не находки)
- Общий архив менеджеров; логистика без финполей (`tests/test_logistics_api.py:528`).  
- Вкладка «На согласовании»: `core/kp/offers_read.py:421-422`.  
- Gate производства: `paid_at` + `order_number_1c` (`archive_service.py:928-935`) плюс отдельное требование спецификации (**Q41**).  
- Correction по diff (`archive_service.py:414-420`); `assert_invoice_guids` до записи файла (`:610-616`); путь JSON фиксирован (`:621-628`).  
- Bind контрагента только «в архиве» + `assert_offer_write_access` (`:305-308`).  
- Сканы договора — UUID-имя (`supply_contract_service.py:833-837`).

Глобальный счётчик P0/P1 из прогона --full **не пересчитывался**; в скоупе подтверждены **0 P0 / 3 P1** (S5, A18, A31).
