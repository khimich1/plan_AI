# Аудит: договор поставки (архив)

**Дата**: 2026-09-27
**Скоуп**: договор поставки (supply-contract) — backend `supply_contract_*` (`app/services/supply_contract_service.py`, `supply_contract_import.py`, `app/repositories/supply_contract_repository.py`, `app/schemas/supply_contract.py`, `core/supply_contract.py`, `core/supply_contract_card_ocr.py`, `core/supply_contract_card_parse.py`, `core/supply_contract_docx.py`); роуты в `app/api/v1/endpoints/archive.py`; OCR/промпты `core/ocr/prompts.py`; фронт `SupplyContractDrawer` / `SupplyContractRegistry`, кнопка и встраивание в `OfferDetailsDrawer`; тесты `tests/test_supply_contract.py`.
**Реестр**: [FINDINGS.md](./FINDINGS.md)
**Прогон 0**: — (ближайший по архиву: [2026-09-04-archive-production-audit.md](./2026-09-04-archive-production-audit.md))

## Дельта

| ID | Было | Стало | Суть |
|----|------|-------|------|
| A7 | open High P1 | open High P1 | Фабрика `get_supply_contract_service` без constructor injection. Evidence: `app/dependencies/services.py:120-121`; `app/services/supply_contract_service.py:78-81`, `:297-308`. |
| A11 | open Medium | open Medium | ArchiveService не проверялся; методы договора туда не добавлялись — не resolved. |
| A27 | — | open Medium **NEW** | `update_contract_date` пишет дату и user_id в обход `_assert_writer`. HTTP-роута нет. Evidence: `app/services/supply_contract_service.py:315-334`; `_assert_writer` `:336-337`; тест только happy-path `tests/test_supply_contract.py:340-369`. |
| S1 | by-design | by-design | Не переоткрывать. Общий архив. |
| A15 | by-design | by-design | Не переоткрывать. |
| S5 | open Medium | **open High P1** | Карточка договора уходит во внешний vision-LLM и **не** вызывает `ensure_external_ocr_enabled` / `check_commercial_ocr_rate_limit`. При `OCR_EXTERNAL_ENABLED=false` коммерческий OCR останавливается, `POST …/supply-contract/parse` — нет. Evidence: `app/api/v1/endpoints/archive.py:499-513`; `app/services/supply_contract_service.py:262-274`, `:283-295`, `:297-310`, `:469-503`; гейт `app/services/commercial_upload_validation.py:60-66`, `:141-142`; до двух vision-вызовов `core/supply_contract_card_ocr.py:93-105`. Старый смысл (документы уходят вовне при включённом OCR) сохранён в notes реестра; лидирует обход выключателя. |
| S8 | open Medium | open Medium | Supply-contract handlers отдают `detail=str(exc)`: `archive.py:190-194`, `:217-221`, `:258-268`, `:287-297`, `:435-468`, `:486-490`, `:520-524`. Неожиданный сбой parse — `raise_unexpected_server_error` `:526-529`. |
| S9 | open Medium | open Medium | CSRF читает multipart до отказа: `app/middleware/csrf.py:36-47`. Роуты `archive.py:159-174`, `:181-187`, `:499-506`. |
| S17 | open Medium | open Medium | Тот же класс: импорт листа юриста без лимита строк и без ZIP-сигнатуры. `app/services/supply_contract_import.py:35-60`; размер только `read_upload_file_capped` в `archive.py:187`. |
| Q9 | open Medium | open Medium | `OfferDetailsDrawer.tsx` 1369 строк (было 1276). Кнопка/встраивание `:1220-1225`, `:1273-1281`. Форма вынесена в `SupplyContractDrawer.tsx` (601) — проблема размера не снята. |
| Q31 | — | open Medium **NEW** | Форма не шлёт доверенность: select «доверенность» есть, полей `poa_number`/`poa_date` нет. `SupplyContractDrawer.tsx:494-503`; бэкенд `supply_contract_service.py:372-374` кидает «Укажите номер и дату доверенности». |
| Q32 | — | open Medium **NEW** | Мёртвый verify-путь OCR: `apply_contract_verify`, `already_exists_message`, `get_contract_card_verify_prompt`. `core/supply_contract_card_ocr.py:54-69, 258-259`; `supply_contract_service.py:409-410`; `core/ocr/prompts.py:122`. В production не вызывается. |
| Q33 | — | open Medium **NEW** | `extract_contract_fields` и `reread_contract_fields` — одинаковые тела. `supply_contract_service.py:477-503`. |
| Q34 | — | open Medium **NEW** | God-component `SupplyContractDrawer.tsx` 601 строка: кнопка, modal, OCR-превью, форма, вложенный реестр. |
| Q35 | — | open Medium **NEW** | FE `SupplyContract` урезан относительно `SupplyContractOut`. `frontend/src/features/commercial-archive/types/supplyContract.ts:15-29` vs `app/schemas/supply_contract.py:99-131`. `status: string` вместо union. |
| Q36 | — | open Medium **NEW** | Дата: drawer ISO `SupplyContractDrawer.tsx:279`, реестр `DD.MM.YYYY` `SupplyContractRegistry.tsx:21-27, 88`. |
| Q37 | — | open Medium **NEW** | Нет теста отказа non-writer на `update_contract_date` (пара к A27). `tests/test_supply_contract.py:340-369` только ADMIN happy path. |
| Q38 | — | open Medium **NEW** | Enum статусов/отметок скана в трёх местах: `core/supply_contract.py:11-18`, `app/schemas/supply_contract.py:13-20`, `frontend/.../types/supplyContract.ts:1-11`. |

- Закрыто: 0
- Открыто (подтверждено): 6
- Новое: 9
- Не проверялось: 1 (A11)

**Открытые P0 / P1 в скоупе**: 0 / 2

## Действия (максимум 8)

1. **[S5]** Прогнать распознавание карточки через `ensure_external_ocr_enabled` и `check_commercial_ocr_rate_limit` до vision-вызова.
2. **[A27]** Убрать `update_contract_date` или провести через `_assert_writer` / `update_contract`, плюс тест отказа **[Q37]**.
3. **[Q31]** Поля доверенности в форме при `authority_basis === "доверенность"`, либо убрать опцию из select.
4. **[S17]** Лимит строк и проверка ZIP-сигнатуры на импорте листа юриста.
5. **[S8]** На роутах договора не отдавать клиенту `str(exc)` доменных/SQL ошибок.
6. **[Q32]** Удалить неиспользуемый verify-путь карточки (функции и промпт).
7. **[Q34]** Разрезать `SupplyContractDrawer`: превью карточки и поля формы отдельно от оркестратора.
8. **[Q35]** Выровнять тип `SupplyContract` с `SupplyContractOut` (включая union статуса). Форматы даты **[Q36]** и тройное дублирование enum **[Q38]** — в том же проходе типизации/контракта.

## Открытые P0 / P1

### [A7] Неполный DI — фабрика supply-contract
**Статус**: open
**Улика**: `app/dependencies/services.py:120-121` — `get_supply_contract_service` возвращает `SupplyContractService()` без constructor injection; сервис сам создаёт зависимости `:78-81`, `:297-308`
**Зачем**: Нельзя подменить repository/OCR в тестах через FastAPI Depends; дублирует платформенный паттерн «толстых» фабрик и усложняет изоляцию договора поставки.

### [S5] OCR карточки договора обходит коммерческий OCR-гейт
**Статус**: open
**Улика**: `app/api/v1/endpoints/archive.py:499-513` → `supply_contract_service.py:262-274`, `:283-295`, `:297-310`, `:469-503`; vision `core/supply_contract_card_ocr.py:93-105` (до двух вызовов). Коммерческий гейт: `commercial_upload_validation.py:60-66`, `:141-142` — **не вызывается** на этом пути. При `OCR_EXTERNAL_ENABLED=false` коммерческий parse блокируется, supply-contract parse — нет.
**Зачем**: Политика «внешний OCR выключен» не единообразна; риск утечки сканов карточки во внешний LLM и обход rate limit, заданного для коммерческого OCR.

## Приложение

### Platform / by-design (без отдельного action в лимите 8)

**[A7]** — см. P1 выше; затрагивает и другие `get_*_service` в `services.py`.

**[A11] ArchiveService god-orchestrator** — open Medium; в этом прогоне длину не перемеряли; методы договора в `ArchiveService` не вливались (2026-09-27).

**[S1] / [A15] by-design** — общий архив КП; `owner_user_id` вне policy доступа.

**[S9] CSRF multipart до токена** — `app/middleware/csrf.py:36-47`; upload/parse договора `archive.py:159-174`, `:181-187`, `:499-506`.

### Безопасность (Medium)

**[S8]** — `str(exc)` на роутах supply-contract в `archive.py:190-194`, `:217-221`, `:258-268`, `:287-297`, `:435-468`, `:486-490`, `:520-524`; parse `:526-529` — `raise_unexpected_server_error`.

**[S17]** — импорт листа юриста: `supply_contract_import.py:35-60` без cap строк и magic bytes; только cap размера файла в `archive.py:187`. (DS XLSX — прежние улики в реестре.)

**[A27]** — `update_contract_date` `:315-334` пишет в БД без `_assert_writer` (`:336-337`); HTTP-эндпоинта нет, но API сервиса доступен из кода/тестов.

### Качество (Medium)

**[Q9]** — `OfferDetailsDrawer.tsx` 1369 строк; wiring договора `:1220-1225`, `:1273-1281`.

**[Q31]** — UI «доверенность» без `poa_number` / `poa_date` → 422 на бэкенде `:372-374`.

**[Q32]** — мёртвый verify OCR (`apply_contract_verify`, промпт `:122` в `prompts.py`).

**[Q33]** — дубли `extract_contract_fields` / `reread_contract_fields` `:477-503`.

**[Q34]** — `SupplyContractDrawer.tsx` 601 строка (modal + OCR + форма + реестр).

**[Q35]** — слабый FE-тип vs `SupplyContractOut`; **Q36** — ISO в drawer vs `DD.MM.YYYY` в реестре; **Q38** — три копии enum статусов/скана.

**[Q37]** — нет негативного теста non-writer для `update_contract_date` (пара к A27).
