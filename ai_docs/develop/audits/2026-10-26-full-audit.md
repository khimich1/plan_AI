# Аудит: --full / весь проект

**Дата**: 2026-10-26
**Скоуп**: живой код `app/`, `core/`, `frontend/src/`, `viz_modules/`; контуры GSM, auth, КП, производство, layout. Не живые и в прогон не входили: `bot/`, `bot_archived/`, `tests/archived/`. Доменные чеклисты применили ревьюеры. Этот документ собран из дельты координатора; код продукта на шаге отчёта не перечитывался.
**Реестр**: [FINDINGS.md](./FINDINGS.md)
**Прогон 0**: [2026-09-21-full-audit.md](./2026-09-21-full-audit.md)
**Рядом, другой скоуп**: [2026-10-26-production-audit.md](./2026-10-26-production-audit.md) — не изменялся.

## Дельта

| ID | Было | Стало | Суть |
|----|------|-------|------|
| S4 | resolved | open | npm audit в `frontend/`: exit 1; high=5, critical=0, moderate=4, total=9. Runtime `zod` 4.3.6 (`frontend/package.json:23`, `<4.6.6`); `vite` 8.0.16 (`package.json:34`, `8.0.0–8.1.5`). Транзитивные high: undici через jsdom, brace-expansion, source-map-js. High, до Critical не повышать |
| A1 | open | open | Partial HTTP+CPU на месте, fallback жив. `core/config_and_data.py:441-452`, `core/domain/plate_order.py:320-338`, `viz_modules/layout_sequence/from_plan.py:193-194`, `core/plate_runtime_state.py:138-153`, `app/middleware/plate_runtime_isolation.py:12-17`, `app/concurrency/cpu_bound.py:32-41` |
| A2 | open | open | Fail-fast только production+single_instance+явный workers>1. `app/security/login_rate_limit.py:49-50`, `:84-85`, `:179-191`; `app/main.py:47-52`; тот же паттерн `commercial_upload_validation.py:24` |
| A4 | open | open | `production.py` 705 строк, 26 маршрутов. `:18-24`, `:175-186` (журнал глотает сбой), `:441-453`, `:474-485`, `:505-513` |
| A5 | open | open | `sgp_service.py:5`, `:717-729`; `day_view_service.py` 615 строк `:231`, `:234-240`; `delivery_schedule_service.py:6`, `:91-93`; `shipment_service.py` 1001 строк, `import sqlite3`. delivery/readiness не сужать. ID не дробить |
| A6 | open | open | `core/production/planning.py:41-46`; `core/visualization/__init__.py:16-18`; `core/visualization/layout.py:20-24`, `:253`. Порт `core/ports/visualization.py` есть, обхода пакета нет |
| A7 | open | open | `app/dependencies/services.py:121-122`; `supply_contract_service.py:122-136`. Без Depends на репозиторий: `:102-103`, `:117-118` |
| A16 | open | open | `gsm_generation_service.py` 1177 строк (было 1030), 24 метода. `:56`, `:70`, `:328`, `:352`, `:455`, `:538`, `:727` |
| A17 | open | open | `core/gsm/generator.py` 1533 строки (было 1324), 59 функций, `generate` с `:134` |
| A18 | open | open | `gsm_export_service.py:54-70`; `gsm.py:662-672`. Тот же дефект: `supply_contract_service.py:778-806`, вызов `:840-841`, `archive.py:742-748`. Новый ID нет |
| A19 | open | open | `core/gsm/season.py:21-32`; `gsm_generation_service.py:850-856`; `gsm_registry_service.py:293-303`; `app/schemas/gsm.py:279-280`; `downstreamPreview.ts:33-44` |
| A20 | open | open | `gsm_transaction_service.py:65-114`; `gsm_repository.py:388`, `:404-424` |
| A22 | open | open | `production_completion_service.py` 702 строки; sqlite3 `:4`, `:123-125`; `complete_day` `:281` |
| A23 | open | open | `delivery_schedule_service.py` 878 строк, 24 метода, соединение `:75-93`, репозитория нет |
| A24 | open | open | Partial, не закрывать. `from_plan.py:693`, `:761`, `:965`; тест `test_golden_root_items_carry_cut_grade_not_resolver_default`; хеш `tests/test_layout_hash.py:1-4`, `:40-55`. Skip/rescue `plate_attribution` и legacy `day_view` не перечитывались |
| A26 | open | open | Partial, не закрывать. `coverage_verify.py:60` теперь `ok = not missing and not surplus`. ILP `ilp_model.py:261` по-прежнему `==`. `extract_cuts` и rescue в `plan_commit` не открывались |
| S3 | open | open | `commercial_upload_validation.py:23-24`, `:53-57`, `:141-142`; default 0 → выход `core/config/settings.py:150-154`; плюс строки A2 |
| S5 | open | open | `archive.py:367-375`, `:710-718`, `:847-859`; `supply_contract_service.py:421-439`, `:1031-1074`; grep `ensure_external_ocr_enabled` в сервисе договора — 0. Картинка и текст |
| S16 | open | open | `gsm.py:196-199`; 50 МБ на файл `settings.py:145-148`; `core/gsm/transactions.py:199-202` без лимита строк; `gsm_repository.py:404-424` |
| Q19 | open | open | `downstreamPreview.ts:23-26` Math.round; `core/gsm/balance.py:15-17`; `tests/test_gsm_balance.py:73-77`; `downstreamPreview.test.ts:54-55` не на границе `*.xx5` |
| Q20 | open | open | `gsm_registry_service.py:343-361` глотает; генерация `:850-858`; `gsm_export_service.py:208-215`; `gsm_report_service.py:294-301`; `tests/test_gsm_season.py:358-371`. `switch_season` перезаписывает из глотающего пути |
| Q40 | open | open | Не закрывать. Сборка пишет username иначе `str(id)`: `production.py:111-113`, `production_service.py:469-476`, `sgp_service.py:717-726`. Пустые: `production.py:297`, `:399`, `:580`, `:600`; `production_capacity_service.py:25-35`. Сессия: `auth_repository.py:84-85`, `app/dependencies/auth.py:21` |
| A25 | open | open | Не перепроверялся. Остаётся open P0. Код не открывали, last_seen не сдвигался |
| A3 | resolved | resolved | Фасад 873 строки; `calculate_draft` делегирует `commercial_workflow_service.py:691-692` |
| A29 | resolved | resolved | `reserve_on_conn` на том же conn, один commit `core/production/planning.py:341-385`, вызов `production_service.py:458-476`. Тест `test_sgp_reserve_fail_after_build_compensates_no_plan_plates_not_in_plan`. `set_active` `planning.py:392-398` в A29 не возвращать |
| A30 | resolved | resolved | Тот же commit; `commit_plan_deletion` `plan_storage.py:166-194`. Тест `test_plan_save_fail_after_plate_mark_rolls_back_no_plan_plates_stay_in_production` |
| S2 | resolved | resolved | `fastapi==0.141.1`, `starlette==1.6.0`; `pip-audit -r requirements.txt` — No known vulnerabilities found, exit 0 |
| S20 | resolved | resolved | `BuildPlanRequest.expected_version` `app/schemas/production.py:230`, `production.py:140`. Тест `test_append_stale_expected_version_returns_409_plan_and_plates_unchanged`. `fill_targets` без проверки версии S20 не переоткрывает |
| S21 | resolved | resolved | DELETE → `commit_plan_deletion` `production.py:261-267`, `production_service.py:156-163`, `plan_storage.py:166-194` |
| Q1 | resolved | resolved | `product_draft_config.py` — таблица типов |
| Q2 | resolved | resolved | Тонкие `_run_product_*` `commercial.py:117-156` |
| Q3 | resolved | resolved | Builder ~265 строк; хеш `tests/test_layout_hash.py:30-37` |
| Q5 | resolved | resolved | `CommercialPlateResolve.resolve_wide_plates` / `resolve_unpriced_plates`; workflow только делегирует |
| S1 | by-design | by-design | `offer_access.py:1-8`. Как IDOR не открывать |
| S11 | by-design | by-design | CSRF-cookie не HttpOnly ожидаемо `app/security/csrf.py:21-28` |
| A15 | by-design | by-design | Связан с S1. Строку реестра не меняли |
| A11, S8, S9, S10, Q4, Q6, Q8, Q10, Q12, Q23, Q24 | open | open | Medium со свежей уликой, last_seen 2026-10-26. В тело P0/P1 не входят |
| Q17, S13, S15 | open | open | Low, взгляд был, last_seen 2026-10-26. Новых Low нет |
| A8, A9, A12, A21, A27, A28 | open | open | Без изменения. Свежей улики файл:строка нет, last_seen не сдвигался. Это не «не воспроизвелось» и не закрытие |
| A10, A14, S6, S7, S12, S14, S17–S19, Q7, Q9, Q11, Q13–Q16, Q18, Q21, Q22, Q25, Q27–Q39 | open | open | Код не открывали. Статус и last_seen не менялись. A25 — строкой выше |

- Закрыто: 0 (ни один open не стал resolved)
- Переоткрыто: 1 (S4)
- Открыто (подтверждено): 35 (2 Critical: A1, A2; 19 High: A4, A5, A6, A7, A16–A20, A22, A23, A24, A26, S3, S5, S16, Q19, Q20, Q40; 11 Medium: A11, S8, S9, S10, Q4, Q6, Q8, Q10, Q12, Q23, Q24; 3 Low: Q17, S13, S15). S4 в это число не входит. A25 не входит
- Новое: 0
- Не воспроизвелось: 34 (код не открывали; статус open сохранён, last_seen не сдвигался), включая A25
- Без изменения: 6 (A8, A9, A12, A21, A27, A28)

**Открытые P0 / P1 в скоупе**: 3 / 20

Пересчёт по таблице после правки S4 (open + Critical = P0, open + High = P1):

| Scope | Open P0 | Open P1 |
|-------|---------|---------|
| layout | 2 (A1, A25) | 3 (A6, A24, A26) |
| platform | 1 (A2) | 3 (A7, S3, S4) |
| kp | 0 | 5 (A4, A5, A22, A23, S5) |
| gsm | 0 | 8 (A16, A17, A18, A19, A20, S16, Q19, Q20) |
| production | 0 | 1 (Q40) |
| auth | 0 | 0 |
| **итого** | **3** | **20** |

P0: A1, A2, A25. P1: A4, A5, A6, A7, A16, A17, A18, A19, A20, A22, A23, A24, A26, S3, S4, S5, S16, Q19, Q20, Q40. Расхождения с дельтой координатора нет. A29, A30, S20, S21 в таблице resolved и в 20 не входят.

Прогон 0 считал **2 P0 / 15 P1** и не включал A24–A26 (они уже лежали в таблице и в счётчик того прогона не входили), а также ещё не содержал S5 и Q40. Сейчас 3 / 20, потому что в реестре учтены A25 (P0), A24 и A26 (P1), S5, Q40, и S4 снова open. Это не «пять новых регрессий» поверх прогона 0: закрыто в этом прогоне 0, новых ID 0.

## Действия (максимум 8)

1. **[A1]** Довести явный PlateOrderContext: убрать чтение заказа через PEP 562 / TLS fallback.
2. **[A2]/[S3]** Общий store для rate limit или жёсткий single-worker во всех окружениях. OCR-limiter ещё и выключен при default 0.
3. **[S4]** Закрыть npm high, в первую очередь runtime `zod`.
4. **[S5]** `ensure_external_ocr_enabled` и rate limit на все пути parse карточки договора.
5. **[A18]** Вынести `soffice` из потока HTTP (ГСМ и PDF договора).
6. **[S16]** Лимит числа файлов импорта ГСМ.
7. **[Q19]** Одно правило округления литров на фронте и в `balance.py`.
8. **[Q40]** Писать актора из username/id на complete, снятии дорожки, unlink/relink и ёмкости дня.

A25 в эти восемь не входит: в прогоне не подтверждён. В реестре он остаётся открытым P0. Остальные открытые P1 в эту восьмёрку не входят.

## Открытые P0 / P1

### [A1] Неявное мутабельное глобальное состояние заказа плит

**Статус**: open (partial)
**Улика**: `core/config_and_data.py:441-452`, `core/domain/plate_order.py:320-338`, `viz_modules/layout_sequence/from_plan.py:193-194`, `core/plate_runtime_state.py:138-153`, `app/middleware/plate_runtime_isolation.py:12-17`, `app/concurrency/cpu_bound.py:32-41`
**Зачем**: Изоляция HTTP и CPU-пула на месте, чтение заказа через PEP 562 и TLS fallback живо.

### [A2] In-process state блокирует несколько воркеров

**Статус**: open
**Улика**: `app/security/login_rate_limit.py:49-50`, `:84-85`, `:179-191`; `app/main.py:47-52`. Тот же паттерн `commercial_upload_validation.py:24`
**Зачем**: Окно rate limit живёт в памяти процесса. Отказ от нескольких воркеров срабатывает только при production, single_instance и явном workers>1.

### [A25] Сироты и призраки двойной атрибуции

**Статус**: open
**Улика**: свежей нет. Код в этом прогоне не открывали. В реестре last_seen 2026-09-21, статус open, P0. Строки прошлого прогона здесь не повторяются как улика этого прогона.
**Зачем**: Открытый P0 остаётся открытым. Пропуск файлов в этом прогоне ID не закрывает.

### [A4] Толстый `production.py`

**Статус**: open
**Улика**: `production.py` **705** строк, **26** маршрутов. `:18-24`, `:175-186` (журнал исключений глотает сбой), `:441-453`, `:474-485`, `:505-513`
**Зачем**: Роутер производства держит оркестрацию. Журнал исключений после сборки глотает сбой.

### [A5] Сервисы обходят repository

**Статус**: open
**Улика**: `sgp_service.py:5`, `:717-729`; `day_view_service.py:231`, `:234-240` (**615** строк); `delivery_schedule_service.py:6`, `:91-93`; `shipment_service.py` **1001** строк, `import sqlite3`. delivery/readiness не сужать
**Зачем**: СГП, дневной вид, график поставки и отгрузка ходят в SQLite мимо репозитория. Это один ID, его не дробить.

### [A6] Planning импортирует visualization на загрузке

**Статус**: open
**Улика**: `core/production/planning.py:41-46`; `core/visualization/__init__.py:16-18` `matplotlib.use`; `core/visualization/layout.py:20-24`, `:253`. Порт `core/ports/visualization.py` есть, обхода пакета нет
**Зачем**: Импорт планирования поднимает matplotlib. Порт визуализации есть, пакет `core.visualization` при загрузке планирования импортируется.

### [A7] Неполный DI

**Статус**: open
**Улика**: `app/dependencies/services.py:121-122`; `supply_contract_service.py:122-136`. Рядом без Depends на репозиторий: `:102-103`, `:117-118`
**Зачем**: Фабрика возвращает сервис, который сам собирает репозитории.

### [A16] God-сервис жизненного цикла ПЛ

**Статус**: open
**Улика**: `gsm_generation_service.py` **1177** строк (было 1030), **24** метода. Класс `:56`; `:70` generate, `:328` list, `:352` patch, `:455` create, `:538` confirm, `:727` `_rechain_downstream`
**Зачем**: Генерация, список, правка, создание, подтверждение и пересчёт хвоста цепочки живут в одном сервисе. Файл вырос.

### [A17] God-модуль генератора ПЛ

**Статус**: open
**Улика**: `core/gsm/generator.py` **1533** строки (было 1324), **59** функций, `generate` с `:134`
**Зачем**: Генерация путевых листов собрана в одном модуле, он вырос.

### [A18] Блокирующий `soffice` в HTTP

**Статус**: open
**Улика**: `gsm_export_service.py:54-70`; `gsm.py:662-672`. Тот же дефект, новый ID не заводить: `supply_contract_service.py:778-806` timeout 60 с, вызов `:840-841`, синхронные GET `archive.py:742-748`
**Зачем**: Конвертация через `soffice` занимает поток HTTP и на экспорте ГСМ, и на PDF договора.

### [A19] Сезон на фронте и в API

**Статус**: open
**Улика**: `core/gsm/season.py:21-32`; генерация `gsm_generation_service.py:850-856`; API хвост журнала `gsm_registry_service.py:293-303`; схема `app/schemas/gsm.py:279-280`; фронт `downstreamPreview.ts:33-44`
**Зачем**: Сезонная логика продублирована в ядре, в превью и в хвосте журнала API.

### [A20] Импорт транзакций без unit of work

**Статус**: open
**Улика**: `gsm_transaction_service.py:65-114`; `gsm_repository.py:388` commit batch, `:404-424` commit строки
**Зачем**: Импорт коммитит пачку и строки по отдельности.

### [A22] God-сервис завершения производства

**Статус**: open
**Улика**: `production_completion_service.py` **702** строки; sqlite3 `:4`, `:123-125`; `complete_day` `:281`
**Зачем**: Завершение дня — монолит с прямым sqlite3.

### [A23] God-module графика поставки

**Статус**: open
**Улика**: `delivery_schedule_service.py` **878** строк, **24** метода, соединение `:75-93`, репозитория нет
**Зачем**: График поставки держит соединение и сценарии в одном классе.

### [A24] `concrete_grade` в сборке плана

**Статус**: open (partial, не закрывать)
**Улика**: три сборки копируют `concrete_grade` с реза `viz_modules/layout_sequence/from_plan.py:693`, `:761`, `:965`. Тест `test_golden_root_items_carry_cut_grade_not_resolver_default`. Хеш марку выкидывает `tests/test_layout_hash.py:1-4`, `:40-55`. Skip/rescue `plate_attribution` и legacy `day_view` не перечитывались
**Зачем**: Марка бетона на сборке берётся с реза. Хеш раскладки марку отбрасывает. Непрочитанные ветки не дают закрыть ID.

### [A26] Детекция surplus

**Статус**: open (partial, не закрывать)
**Улика**: `coverage_verify.py:60` теперь `ok = not missing and not surplus` (сдвиг относительно улики «только missing»). ILP `ilp_model.py:261` по-прежнему `==`. `extract_cuts` и rescue в `plan_commit` не открывались
**Зачем**: Проверка покрытия видит и surplus, равенство спроса в ILP на месте. Неоткрытые `extract_cuts` и rescue не дают закрыть ID.

### [S3] Rate limit OCR в процессе и выключен при 0

**Статус**: open
**Улика**: `commercial_upload_validation.py:23-24`, `:53-57`, `:141-142`; default `commercial_ocr_uploads_per_hour` = 0 → проверка сразу выходит `core/config/settings.py:150-154`; плюс строки A2
**Зачем**: OCR-limiter живёт в памяти процесса и при лимите по умолчанию 0 запросы не считает.

### [S4] npm audit high

**Статус**: open (переоткрыт; было resolved). Severity High, до Critical не повышать
**Улика**: `npm audit --json` в `frontend/`, exit 1; metadata high=5, critical=0, moderate=4, total=9. Прямые: `zod` 4.3.6 в dependencies (`frontend/package.json:23`, диапазон `<4.6.6`), `vite` 8.0.16 в devDependencies (`package.json:34`, диапазон `8.0.0–8.1.5`). Транзитивные high: undici через jsdom, brace-expansion, source-map-js
**Зачем**: В дереве фронта пять high, среди них runtime-пакет `zod`.

### [S5] Parse карточки договора без OCR-gate

**Статус**: open
**Улика**: `archive.py:367-375`, `:710-718`, `:847-859`; `supply_contract_service.py:421-439`, `:1031-1074`; grep `ensure_external_ocr_enabled` в сервисе договора — 0. Картинка и текст
**Зачем**: Разбор карточки договора не проходит флаг внешнего OCR и rate limit.

### [S16] Импорт ГСМ без лимита числа файлов

**Статус**: open
**Улика**: `gsm.py:196-199`; на файл 50 МБ `settings.py:145-148`; `core/gsm/transactions.py:199-202` без лимита строк; коммит строки `gsm_repository.py:404-424`
**Зачем**: Число файлов и число строк импорта не ограничены. Лимит размера одного файла есть.

### [Q19] Два правила округления литров

**Статус**: open
**Улика**: `downstreamPreview.ts:23-26` `Math.round`; `core/gsm/balance.py:15-17`; `tests/test_gsm_balance.py:73-77`; фронт-тест `downstreamPreview.test.ts:54-55` не на границе `*.xx5`
**Зачем**: Превью и `balance.py` округляют литры по-разному. Тест превью границу `*.xx5` не держит.

### [Q20] Битый `season_switches`

**Статус**: open
**Улика**: `gsm_registry_service.py:343-361` глотает; генерация `gsm_generation_service.py:850-858`; экспорт `gsm_export_service.py:208-215`; отчёт `gsm_report_service.py:294-301`; `tests/test_gsm_season.py:358-371`. `switch_season` читает глотающий путь и перезаписывает
**Зачем**: Реестр превращает битые переключатели в пустоту, генерация, экспорт и отчёт идут другим путём, а `switch_season` записывает результат глотающего чтения.

### [Q40] Пустой актор части мутаций производства

**Статус**: open (не закрывать)
**Улика**: сборка плана уже пишет username, иначе `str(id)`: `production.py:111-113`, `production_service.py:469-476`, `sgp_service.py:717-726`. Пустые вызовы: `production.py:297`, `:399`, `:580`, `:600`; `production_capacity_service.py:25-35`. Сессия: `auth_repository.py:84-85`, `app/dependencies/auth.py:21` — id и username, не email/login/user_id
**Зачем**: Complete, снятие дорожки, unlink/relink и ёмкость дня пишут пустого актора. В сессии есть username и id.

## Приложение

### Medium (в тело P0/P1 не входят)

| ID | Улика | Заметка |
|----|-------|---------|
| A11 | `archive_service.py` **2038** строк, методов 65 | Действия на эту неделю нет. Статус open |
| S8 | `detail=str(exc)` `production.py:148`, `:181`, `:210`, `:216`, `:703`; `archive.py` например `:206`, `:674`, `:805`. Часть маршрутов режется `raise_unexpected_server_error` `app/core/http_errors.py:67-71` | Severity не повышать |
| S9 | `app/middleware/csrf.py:36-47` читает multipart, если нет заголовка | |
| S10 | TTL 12 ч `core/config/settings.py:63-67`, без ротации | |
| Q4 | Копий `build_*_preview_metadata` уже **7** (было 5): `commercial_draft_service.py:197`, `:281`, `:354`, `:427`, `:500`, `:570`, `:643` | Свести в один сборщик до следующего типа изделия. Medium, не P1 |
| Q6 | Оба `get_global_calendar_info`: `plan_calendar.py:70`, `plan_distribution_service.py:135` | Не закрывать |
| Q8 | `useCreatePlanWizardState` **997** строк | В тело P1 не входит |
| Q10 | JSON-маршруты без `response_model` на месте | Свежих номеров строк нет |
| Q12 | Список глотает `archive_service.py:1275-1288`; `move_to_production` рейзит `:993-997` | |
| Q23 | FE производства `Math.ceil(length/101)` `productionEstimate.ts:24` | Archive `round` не перепроверялся. ID не закрывать |
| Q24 | Мёртвый raise после `raise_not_found` в разбивке и формовке `production.py:488-496`, `:515-523`. У схемы хвоста уже нет `:455-465` | Удалить два недостижимых блока. Medium, не P1 |

### Low

Взгляд был, статус open, last_seen 2026-10-26. Новых Low нет. В тело P0/P1 не входят.

| ID | Улика |
|----|-------|
| Q17 | Тексты `GsmGenerationError` на английском `gsm_generation_service.py:122`, `:856` |
| S13 | `/health` в production без метаданных `health.py:16-19` |
| S15 | `draftStorage.ts:8`, `:22` sessionStorage |

### Без изменения

Архитектура: действия на эту неделю нет, свежей улики файл:строка нет. Статус не менялся, last_seen не сдвигался. Это не «не воспроизвелось» и не закрытие.

A8, A9, A12, A21, A27, A28.

### Не перепроверялись

Статус не менялся, last_seen не сдвигался, resolved не ставился.

A10, A14, **A25** (Critical, P0 — открыт, код не открывали), S6, S7, S12, S14, S17, S18, S19, Q7, Q9, Q11, Q13, Q14, Q15, Q16, Q18, Q21, Q22, Q25, Q27, Q28, Q29, Q30, Q31, Q32, Q33, Q34, Q35, Q36, Q37, Q38, Q39.

A10 и S14 относятся к боту. Бот в этом прогоне не живой контур.

### Resolved держится

Статус не менялся. Свежая улика записана в реестр, last_seen 2026-10-26.

A3, A29, A30, S2, S20, S21, Q1, Q2, Q3, Q5 — см. строки дельты.

### By-design

S1 (`offer_access.py:1-8`) как IDOR не открывался. S11: CSRF-cookie не HttpOnly ожидаемо (`app/security/csrf.py:21-28`). A15 связан с S1, строка реестра не менялась.

### Инварианты

В находки не заносились: логистика без финансовых полей КП; черновики с проверкой пути и владельца; `destructive_db_guard` на admin reset и GSM reset-to-anchors; сессия HttpOnly.

### GSM, без нового Q

Шлюз закрытия месяца есть. Старт генерации — с последнего confirmed/exported. `_rechain_downstream` не трогает километраж месяцев confirmed/exported. Повторный экспорт спрашивает подтверждение в UI; флага на API нет.

### Снимок размеров из улик этого прогона

| Модуль | Строк |
|--------|-------|
| `archive_service.py` | 2038 (методов 65; в реестре до прогона было 1892) |
| `core/gsm/generator.py` | 1533 (было 1324) |
| `gsm_generation_service.py` | 1177 (было 1030), 24 метода |
| `shipment_service.py` | 1001 |
| `useCreatePlanWizardState` | 997 |
| `commercial_workflow_service.py` (фасад) | 873 |
| `delivery_schedule_service.py` | 878, 24 метода |
| `production.py` | 705, 26 маршрутов |
| `production_completion_service.py` | 702 |
| `day_view_service.py` | 615 |
| Копий `build_*_preview_metadata` | 7 (было 5) |

Новых ID нет. Свободные номера: **A31**, **S22**, **Q41**. A13 и Q26 не занимать.
