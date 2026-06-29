# Отчёт аудита: `core/kp_db.py`

**Дата:** 2026-06-03  
**Область:** `core/kp_db.py` (~4034 строки, ~54 публичные функции)  
**Проверено:** senior-reviewer + security-auditor + reviewer  

> **Remediation (2026-06-04):** P0/P1 закрыты; **phase 2 complete (stages 7–12)**. A1: модули `kp_db_*` + фасад. A2: `plate_completion_service`, `rest_matching_service`, `kp_persistence_service` в `core/`. A3: bot → `kp_persistence`; production services → slice modules (guard `test_production_services_kp_boundary`). A4: `ensure_schema` на старте. A5: `kp_db_audit.audit_append` + `PlateAuditRepository`. A7/S1: agent NDJSON убран из completion hot path. **Health Score (estimate): ~6.2 / 10** (было 2.0). Отчёты: [phase 2 complete](../reports/2026-06-04-kp-db-remediation-phase2-complete.md), stages [7](../reports/2026-06-04-kp-db-stage7-a2-rests.md)–[12](../reports/2026-06-04-kp-db-stage12-kp-persistence.md).

---

## Краткое резюме (Executive Summary)

**Health Score: 2.0 / 10** (pre-remediation; post phase 2 estimate: **~6.2 / 10** — see [phase 2 report](../reports/2026-06-04-kp-db-remediation-phase2-complete.md))

Формула: `10 − min(2×2, 6) − min(15×0.5, 3) − min(22×0.1, 1) = 10 − 4 − 3 − 1 = 2.0`

| Серьёзность | Архитектура | Безопасность | Качество кода | **Итого** |
|-------------|-------------|--------------|---------------|-----------|
| Critical    | 2           | 0            | 0             | **2**     |
| High        | 5           | 7            | 3             | **15**    |
| Medium      | 8           | 6            | 8             | **22**    |
| Low         | 5           | 4            | 6             | **15**    |
| **Всего**   | **20**      | **17**       | **17**        | **54**    |

**Вердикт:** Модуль `kp_db.py` — центральный persistence-слой коммерческих предложений, lifecycle плит, остатков и менеджеров. В текущем виде он сочетает **критическую архитектурную концентрацию** (god module + доменная логика в БД-слое) с **высокими рисками утечки PII/КП** (debug NDJSON, хардкод менеджеров) и **опасными операциями списания** (кросс-КП в `find_one_row`). SQL-инъекций и захардкоженных секретов в файле не обнаружено; параметризованные запросы используются последовательно.

**Рекомендация:** Перед дальнейшим расширением функционала закрыть **2 Critical** (декомпозиция + вынос доменной логики), затем пакет **P0/P1** из матрицы приоритетов (debug-логи, кросс-КП, `clear_all_*`, тесты на `move_plates_to_completed` / `find_one_row`). Remediation в рамках данного аудита **не выполнялся**.

> **Контекст:** Частичная декомпозиция уже начата (`core/kp_db_nomenclature.py`, re-export в `kp_db.py`). Настоящий отчёт фиксирует **остаточное состояние** монолита `kp_db.py` как отдельной области.

---

## Критические проблемы (исправить в первую очередь)

### [A1] God Module — монолит persistence + домена

**Категория:** Архитектура  
**Расположение:** `core/kp_db.py` (весь файл, ~4034 строки, ~54 функции)  
**Влияние:** В одном модуле смешаны CRUD КП, lifecycle плит (`move_plates_to_completed`, `assign_plates_to_plan`, `return_plates_to_production`), остатки (`plate_rests`, `find_matching_rests`), менеджеры, миграции DDL (`init_schema`), аудит статусов, отладочные записи. Любое изменение затрагивает несвязанные bounded context; code review и изолированное тестирование практически невозможны. `app/repositories/kp_repository.py` остаётся тонким прокси без реальной границы.  
**Исправление:** Поэтапная декомпозиция по bounded context: `repositories/kp_offers.py`, `repositories/plates.py`, `repositories/rests.py`, `repositories/managers.py`; миграции схемы — отдельный модуль; фасад `kp_db` — deprecated shim на переходный период.  
**Команда:** `/refactor core/kp_db.py`

---

### [A2] Доменная логика в persistence-слое

**Категория:** Архитектура  
**Расположение:**
- `move_plates_to_completed` — строки ~1274–1790 (~520 строк, вложенный `find_one_row` ~220 строк)
- `find_matching_rests` — ~2119+
- `save_kp_to_db` — ~412–560 (расчёт позиций, BLOB, статусы)

**Влияние:** Правила списания плит (шаги 0–7 в `find_one_row`, допуски длины ±0.02 м, кросс-КП для 61,1/61,2), подбор остатков и сохранение КП реализованы внутри SQL-транзакций без domain services. Бизнес-правила нельзя переиспользовать в `app/services` и боте единообразно; регрессии при правках SQL высоки.  
**Исправление:** Вынести оркестрацию в `PlateCompletionService`, `RestMatchingService`, `KpPersistenceService`; в repository — только CRUD и атомарные операции с явными DTO из `core/domain`.  
**Команда:** `/refactor core/kp_db.py` + `/implement` для domain services

---

## Высокий приоритет (High)

### Архитектура (5)

| ID | Проблема | Расположение | Краткое исправление |
|----|----------|--------------|-------------------|
| **[A3]** | Обход слоёв — прямые импорты `kp_db` из bot и `plan_commit` | `bot/handlers/commercial.py`, `production_create.py`, `production_execution.py`, `archive.py`, `admin.py`; `core/plan_commit.py:27`; `app/services/production_planning_service.py` (частично через сервис, но repository → kp_db) | Маршрутизировать через `app/services/*`; handlers — только Telegram + DI |
| **[A4]** | `init_schema()` на каждой операции | **48 вызовов** по файлу (проверено grep); определение ~121 | Однократная инициализация при старте приложения / Alembic; lazy flag `_schema_ready` |
| **[A5]** | Дублирование audit с `plate_audit_repository` | `_audit_append` ~75–118; комментарий дублирует `app.repositories.plate_audit_repository` | Единый `PlateAuditRepository`; core не пишет в `plate_status_log` напрямую |
| **[A6]** | `find_one_row` — O(n) эвристики внутри транзакции | Вложенная функция ~1344–1550 в `move_plates_to_completed`; шаги 2.55, 2.6 без индексов по эвристикам | Индексы + явные ключи поиска (`nomenclature_id`, `length_dm_raw`); убрать полный скан по допуску |
| **[A7]** | Debug NDJSON в production hot path | `_debug_session_write` ~36–52; множественные `open(_DEBUG_LOG*)` в `move_plates_to_completed` / `find_one_row` (~1347–1758) | Удалить или обернуть `if settings.DEBUG`; не писать в hot path |

### Безопасность (7)

| ID | Проблема | Расположение | Краткое исправление |
|----|----------|--------------|-------------------|
| **[S1]** | Debug NDJSON — утечка данных КП/PII в `debug_logs/` | `get_debug_log_path` ~23–33; запись plate_name, kp_id, qty в ~1347–1758 | Отключить в production; не логировать PII; ротация/очистка `debug_logs/` |
| **[S2]** | PII менеджеров в коде | `init_default_managers` ~3656–3662 (ФИО, телефоны, email) | Вынести в seed-скрипт / env; не коммитить реальные контакты |
| **[S3]** | Нет авторизации на слое БД | Все публичные функции — любой caller с импортом | Проверка прав только в `app/services` + bot middleware; не экспортировать destructive API в handlers |
| **[S4]** | Кросс-КП списание в `find_one_row` шаг 2.55 | ~1429–1453: поиск по длине **без** `kp_id` в WHERE | Требовать `allow_cross_kp=True` явно; по умолчанию только `prefer_kp_id` |
| **[S5]** | Path traversal на `xlsx_file_path` | `save_kp_to_db` ~539–552; `save_xlsx_file` ~893–933 | `os.path.realpath` + whitelist каталога; запрет `..` |
| **[S6]** | SQLite без шифрования, BLOB КП | `kp_files.xlsx_blob` ~539–546; схема `init_schema` | Шифрование тома / SQLCipher для production; лимит размера BLOB |
| **[S7]** | `clear_all_*` без предохранителей | `clear_all_plates_data` ~1027–1126; `clear_all_kp` ~1179+ | Подтверждение, `APP_ENV != production`, audit log actor |

### Качество кода (3)

| ID | Проблема | Расположение | Краткое исправление |
|----|----------|--------------|-------------------|
| **[Q1]** | Потеря `nomenclature_id` / `length_dm_raw` при split INSERT | Split INSERT ~2480–2486, ~2659–2667, ~2777–2785, ~3132–3140 — колонки отсутствуют vs `save_kp_to_db` ~524–528 | Копировать `nomenclature_id`, `length_dm_raw` во все INSERT/SELECT split |
| **[Q2]** | Мегафункции | `move_plates_to_completed` ~1274–1790; `find_one_row` ~1344–1550 | Разбить на стратегии поиска + unit-тесты на каждый шаг |
| **[Q3]** | Нет тестов для критических путей | Только `test_kp_db_search_by_customer.py`, `test_kp_db_update_logistics_cost.py` | Добавить tests для rests, `move_plates_to_completed`, `find_one_row` (golden cases) |

---

## Средний приоритет (Medium)

### Архитектура (A8–A15)

| ID | Проблема | Расположение |
|----|----------|--------------|
| **A8** | Несогласованность repository pattern — `kp_repository` проксирует монолит | `app/repositories/kp_repository.py` → `core.kp_db` |
| **A9** | Нет Unit of Work — каждая функция открывает своё соединение | `_connect` + `conn.close()` в каждой функции |
| **A10** | DDL и миграции в runtime `init_schema` | ~121–410: `CREATE TABLE`, `ALTER TABLE`, backfill `length_dm_raw` |
| **A11** | BLOB XLSX в SQLite — раздувание БД | `kp_files`, `save_kp_to_db` ~539–546 |
| **A12** | Dict-based API без TypedDict/Pydantic на границе | `List[Dict]`, `Dict[str, List[Dict]]` в публичных сигнатурах |
| **A13** | Дублирование логики split qty в 4 местах | ~2480, ~2659, ~2777, ~3132 |
| **A14** | Re-export номенклатуры из монолита | ~405: `from core.kp_db_nomenclature import ...` — смешение фасада |
| **A15** | Несогласованные контракты ошибок (bool / int / Dict / raise) | `delete_kp`, `clear_all_*`, `move_plates_to_completed` |

### Безопасность (M1–M6)

| ID | Проблема | Расположение |
|----|----------|--------------|
| **M1** | `print` + traceback в operational path | Множественные `print(f"[DB]...")` по файлу; исключения ~1785 |
| **M2** | Silent `except: pass` в debug-записях | ~1353–1354, ~1451–1452, `_debug_session_write` ~51–52 |
| **M3** | BLOB без лимита размера файла | `save_kp_to_db` / `save_xlsx_file` — чтение всего файла в память |
| **M4** | `db_path` параметр без валидации — подмена БД | Все функции с `db_path: str = DEFAULT_DB` |
| **M5** | Абсолютный `file_path` сохраняется в БД | `kp_files.file_path` ~546, ~927 |
| **M6** | `init_schema` на каждый вызов — лишняя нагрузка и race при миграциях | 48 вызовов (см. A4) |

### Качество кода (Q4–Q11)

| ID | Проблема | Расположение |
|----|----------|--------------|
| **Q4** | Мёртвые/избыточные debug-константы | `_DEBUG_*` ~27–33, не все используются |
| **Q5** | 4× дублирование split INSERT (DRY) | ~2480, ~2659, ~2777, ~3132 |
| **Q6** | Пробелы в audit при частичных операциях | `_audit_append` не везде при split/return |
| **Q7** | Несогласованные return-типы и сообщения об ошибках | См. A15 |
| **Q8** | `get_kp_completion_percentage` — потенциальная ошибка деления/логики | ~3809+ |
| **Q9** | `save_kp_to_db` — fallback НДС без явного контракта | ~412–560 |
| **Q10** | Magic constants цен резов | `LONG_CUT_PRICE_PER_M`, `TRANSVERSE_CUT_PRICE` ~2150–2151 |
| **Q11** | `#region agent log` — техдолг отладки в prod-коде | `move_plates_to_completed` ~1315–1758 |

---

## Низкий приоритет (Low)

### Архитектура (A16–A20)

| ID | Проблема | Расположение |
|----|----------|--------------|
| **A16** | Логирование через `print` вместо `logging` | По всему файлу (~50+ вхождений) |
| **A17** | `DEFAULT_DB` — жёсткий путь к `plita.db` в корне | ~26 |
| **A18** | Блок `if __name__ == '__main__'` с demo-данными | ~3980–4034 |
| **A19** | Lazy/runtime imports внутри функций | `import json` в hot path; `cfg.extract_length_dm_raw` в `init_schema` ~361 |
| **A20** | Magic cut prices без конфига | ~2150–2210 |

### Безопасность (L1–L4)

| ID | Проблема / позитив | Расположение |
|----|-------------------|--------------|
| **L1** | WAL + `foreign_keys=ON` — хорошая практика | `_connect` ~61–72 |
| **L2** | Параметризованные SQL — нет конкатенации user input в запросах | По файлу |
| **L3** | **`_escape_sql_like`** — корректное экранирование LIKE | ~3746–3769 |
| **L4** | Demo `__main__` создаёт тестовые КП при случайном запуске | ~4000–4034 |

### Качество кода (Q12–Q17)

| ID | Проблема | Расположение |
|----|----------|--------------|
| **Q12** | Неполные type hints на вложенных структурах | `List[Dict]` без TypedDict для plate items |
| **Q13** | Runtime imports в циклах/транзакциях | `move_plates_to_completed` ~1316 |
| **Q14** | Debug markers (`_is_target`, `_target_substrings`) в prod | ~1327–1342 |
| **Q15** | Устаревшие/избыточные комментарии «Простыми словами» | Docstrings по файлу |
| **Q16** | `recover_stuck_plan_plates` — метрики не в kp_db, но зависимость от API | Caller `scripts/recover_stuck_plan_plates.py` |
| **Q17** | Дублирование `PRAGMA foreign_keys` после `_connect` | ~1322, ~3697 и др. |

---

## Матрица приоритетов (топ 15)

| Приоритет | ID | Проблема | Серьёзность | Усилие |
|-----------|-----|----------|-------------|--------|
| **P0 — сейчас** | A1 | God Module — декомпозиция | Critical | High |
| **P0 — сейчас** | A2 | Доменная логика в persistence | Critical | High |
| **P0 — сейчас** | S1/A7 | Debug NDJSON / PII в `debug_logs/` | High | Low |
| **P0 — сейчас** | S4 | Кросс-КП списание шаг 2.55 | High | Medium |
| **P1 — спринт** | S7 | `clear_all_*` без guard | High | Low |
| **P1 — спринт** | Q1 | Потеря `nomenclature_id` при split | High | Medium |
| **P1 — спринт** | Q3 | Нет тестов rests / move / find_one_row | High | Medium |
| **P1 — спринт** | A3 | Bot/plan_commit → kp_db напрямую | High | High |
| **P1 — спринт** | S2 | PII менеджеров в коде | High | Low |
| **P1 — спринт** | S5 | Path traversal `xlsx_file_path` | High | Low |
| **P2 — следующий спринт** | A4/M6 | 48× `init_schema` на вызов | High/Medium | Medium |
| **P2** | A5 | Дублирование plate audit | High | Medium |
| **P2** | A6/Q2 | `find_one_row` сложность + O(n) | High | High |
| **P2** | S3 | Нет auth на слое БД | High | Medium (policy) |
| **P2** | S6 | SQLite без шифрования + BLOB | High | High |

---

## Следующие шаги

1. **Немедленно (до следующего релиза):** Удалить или зафлажить debug NDJSON ([S1], [A7]); ограничить кросс-КП списание явным флагом и аудитом ([S4]); добавить production-guard на `clear_all_plates_data` / `clear_all_kp` ([S7]).
2. **Текущий спринт:** Исправить split INSERT — сохранять `nomenclature_id` и `length_dm_raw` ([Q1]); написать regression-тесты на `move_plates_to_completed` и `find_one_row` ([Q3]); вынести PII менеджеров из `init_default_managers` ([S2]); валидировать `xlsx_file_path` ([S5]).
3. **Следующий спринт:** Начать декомпозицию [A1] — вынести plates/rests в отдельные модули; вынести `find_one_row` в domain service [A2]; заменить 48× `init_schema` на startup migration [A4]; унифицировать audit через `PlateAuditRepository` [A5].
4. **Бэклог:** Dict → TypedDict/Pydantic [A12]; единый UoW [A9]; шифрование БД [S6]; замена `print` на `logging` [A16]; удаление `__main__` demo [A18/L4].
5. **Повторный аудит:** После первого среза декомпозиции — targeted re-audit `core/kp_db.py` с пересчётом Health Score.

---

## Метрики аудита

| Метрика | Значение |
|---------|----------|
| Строк в файле | ~4034 |
| Публичных функций (`def`) | ~54 |
| Вызовов `init_schema` | 48 (проверено) |
| Прямых импортов `core.kp_db` / `kp_db` | 20+ модулей (app, bot, core, tests, scripts) |
| Тестовых файлов только для kp_db | 2 |
| Health Score | **2.0 / 10** |

---

## Связанная документация

- [Полный аудит проекта](2026-06-03-full-project-audit.md) — контекст [A2] как PARTIAL
- [Orchestration arch-triage](../reports/orchestration-arch-triage-2026-06-03.md) — первый срез `kp_db_nomenclature`
- План: `ai_docs/develop/plans/2026-06-03-architecture-triage-a1-a2-a3.md`

---

*Отчёт сформирован documenter по workflow audit-workflow. Remediation не применялся.*
