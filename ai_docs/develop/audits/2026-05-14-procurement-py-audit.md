# Project Audit Report

**Date:** 2026-05-14  
**Scope:** `viz_modules/procurement.py`  
**Audited by:** senior-reviewer, security-auditor, reviewer  

---

## Executive Summary (краткое резюме)

**Общий показатель здоровья (Health Score): 6.5 / 10**  
(формула: старт 10; −0.5 за каждый High с потолком −3 → при 5 High = −2.5; −0.1 за каждый Medium с потолком −1 → при 12 Medium = −1).

| Категория       | Critical | High | Medium | Low |
|-----------------|----------|------|--------|-----|
| Architecture    | 0        | 4 (A1–A4) | 4  | 3   |
| Security        | 0        | 1 (S1)    | 3  | 4   |
| Code Quality    | 0        | 0         | 5  | 5   |
| **Total**       | **0**    | **5**     | **12** | **12** |

Критических проблем по этому файлу не выявлено. **Рекомендация:** перед крупными изменениями функционала устранить пункты **High** (в первую очередь доверие к данным плана **S1** и структурные риски **A1–A4**), затем планомерно снижать накопленную среднюю серьёзность (**Medium**) в области дублирования, побочных эффектов и валидации.

---

## Critical Issues

**None identified.**

---

## High Priority

Объединённые по приоритету **High** находки из архитектуры и безопасности; для Code Quality в этой категории дополнительных пунктов нет (см. ревью).

### Architecture

- **A1** — Нарушение SRP / god module: один файл ~1888 строк, совмещает get_orders_from_opt_plan (35–93), build_procurement_items (281–542), build_price_rows (545–694), build_price_rows_production (697–1007), build_component_breakdown (1010–1329), build_component_breakdown_production (1332–1886), _calc_trim_components (154–278).

  *Предлагаемые меры кратко:* разбить модуль по зонам ответственности (заказы, прайс-строки, production-ветка, breakdown, trim); выделить сервисы/модули с явными публичными API.

- **A2** — Слабое разделение слоёв: прямые зависимости на core.config_and_data, core.optimization (ленивые импорты), core.price_db, raw_material_db, reinforcement_db — нет границы viz → application → ports.

  *Предлагаемые меры:* ввести порты/адаптеры для доступа к ценам и БД; держать viz-тонкий слой без прямого «протекания» core в расчётные функции.

- **A3** — Дублирование вместо единой стратегии: production ветки копируют поиск плана и разбор primary_cuts/secondary_cuts вместо переиспользования _calc_trim_components (build_price_rows_production ~778–926, build_component_breakdown_production ~1535–1684).

  *Предлагаемые меры:* единая функция/стратегия trim и разбора резов; убрать копипаст между production-ветками.

- **A4** — Зависимость от process-global OPT_* без инъекции плана — сложность параллелизма и тестов.

  *Предлагаемые меры:* передавать контекст плана/конфигурацию явно (параметры, DI), минимизировать глобальное состояние.

### Security

- **S1** — Trust boundary на данных плана: get_orders_from_opt_plan / build_procurement_items — минимальная валидация; load_code и размеры as-is → риск манипуляции котировками / resource pressure если глобальное состояние недостоверно.

  *Предлагаемые меры:* схема валидации входа (типы, диапазоны, допустимые коды), явная политика доверия к источнику плана, лимиты на размер/объём данных.

---

## Medium Priority

### Architecture

- **A5** — Высокая цикломатическая сложность и легаси-ветвления в build_procurement_items (346–542).

  *Кратко:* декомпозиция на шаги/подфункции, таблицы правил вместо глубоких if-цепочек.

- **A6** — Масштабируемость: O(n) get_price/get_raw_material_cost/get_reinforcement в циклах без батчинга.

  *Кратко:* батч-загрузка по ключам, кэш на время расчёта.

- **A7** — Побочные эффекты: print и NDJSON в build_price_rows (572–610) смешаны с расчётом.

  *Кратко:* вынести логирование/дамп в отдельный слой; чистые функции для расчёта.

- **A8** — Тестируемость/DIP: нет портов, всё на cfg и core get_*.

  *Кратко:* интерфейсы для цен/сырья/арматуры, моки в тестах.

### Security

- **S2** — print раскрывает коммерческие данные в stdout.

  *Кратко:* убрать или заменить структурированным логом с уровнем и маскированием.

- **S3** — NDJSON debug_logs с ценами на диск.

  *Кратко:* флаг окружения, редaction, ограничение путей и retention.

- **S4** — print `{e}` может раскрывать детали ошибок.

  *Кратко:* логировать ошибки безопасно, пользователю — обобщённое сообщение.

### Code Quality

- **Q1** — Дрейф production trim vs _calc_trim_components; mismatch base_price vs base_price_1_2m для width waste (~239–242 vs ~892–896).

  *Кратко:* единый источник правды для коэффициентов и ветвей trim.

- **Q2** — Повтор fallback order / regex / reconstruction между breakdown функциями.

  *Кратко:* общие утилиты/одна функция восстановления.

- **Q3** — Broad except Exception → silent data loss.

  *Кратко:* узкие исключения, логирование, явный fail или sentinel.

- **Q4** — Импорты json/os/time внутри hot loop в build_price_rows.

  *Кратко:* импорты на уровне модуля; минимизировать работу в цикле.

- **Q5** — Глубокая вложенность в _calc_trim_components и build_procurement_items.

  *Кратко:* ранние выходы, выделение подфункций.

---

## Low Priority

### Architecture

- **A9** — Магические константы (1020–1080, 0.170*80 и т.д.).

  *Кратко:* именованные константы/конфиг.

- **A10** — Ленивые импорты core.optimization внутри функций.

  *Кратко:* явная зависимость сверху модуля или ленивый сервис-объект с одним местом импорта.

- **A11** — Циклы core↔viz: обратная сторона не проверена из файла.

  *Кратко:* зафиксировать границы зависимостей в код-ревью смежных модулей.

### Security

- **S5** — Global state / multi-tenant footgun.

- **S6** — Swallowed I/O на debug логах.

- **S7** — Нет auth в модуле — контракт на вызывающий слой.

- **S8** — Пути файлов фиксированы через get_debug_log_path — path traversal не из этого файла.

### Code Quality

- **Q6** — Dead/redundant imports (unused OPT_CASCADING_PLAN и др.).

- **Q7** — Слабая типизация, mutable-default smell на price_rows.

- **Q8** — Комментарий про поперечный рез вводит в заблуждение.

- **Q9** — Дублирующиеся магические числа, двойной комментарий # Базовая цена.

- **Q10** — Пробелы в тестах для production/breakdown/get_orders.

---

## Priority Matrix (top items)

| ID  | Severity | Effort (estimate) | Priority |
|-----|----------|-------------------|----------|
| S1  | High     | Medium            | P0       |
| A1  | High     | High              | P0       |
| A2  | High     | High              | P1       |
| A3  | High     | Medium            | P1       |
| A4  | High     | Medium            | P1       |
| Q1  | Medium   | Medium            | P2       |
| Q2  | Medium   | Low–Medium        | P2       |
| Q3  | Medium   | Low               | P2       |

*Effort:* **Low** ≈ локальные правки; **Medium** — координация нескольких функций/тестов; **High** — выделение модулей и контрактов.

---

## Full Findings (verbatim source sections)

Ниже — полный текст находок с нормализованным форматированием; идентификаторы сохранены.

### Architecture Findings (from senior-reviewer)

#### Critical

- None — по одному только этому файлу нет признака обязательного «обрушения» слоёв; основные проблемы — сопротивляемость изменениям, тестируемость и дрейф логики.

#### High

- **A1** — Нарушение SRP / god module: один файл ~1888 строк, совмещает get_orders_from_opt_plan (35–93), build_procurement_items (281–542), build_price_rows (545–694), build_price_rows_production (697–1007), build_component_breakdown (1010–1329), build_component_breakdown_production (1332–1886), _calc_trim_components (154–278).
- **A2** — Слабое разделение слоёв: прямые зависимости на core.config_and_data, core.optimization (ленивые импорты), core.price_db, raw_material_db, reinforcement_db — нет границы viz → application → ports.
- **A3** — Дублирование вместо единой стратегии: production ветки копируют поиск плана и разбор primary_cuts/secondary_cuts вместо переиспользования _calc_trim_components (build_price_rows_production ~778–926, build_component_breakdown_production ~1535–1684).
- **A4** — Зависимость от process-global OPT_* без инъекции плана — сложность параллелизма и тестов.

#### Medium

- **A5** — Высокая цикломатическая сложность и легаси-ветвления в build_procurement_items (346–542).
- **A6** — Масштабируемость: O(n) get_price/get_raw_material_cost/get_reinforcement в циклах без батчинга.
- **A7** — Побочные эффекты: print и NDJSON в build_price_rows (572–610) смешаны с расчётом.
- **A8** — Тестируемость/DIP: нет портов, всё на cfg и core get_*.

#### Low

- **A9** — Магические константы (1020–1080, 0.170*80 и т.д.).
- **A10** — Ленивые импорты core.optimization внутри функций.
- **A11** — Циклы core↔viz: обратная сторона не проверена из файла.

### Security Findings (from security-auditor)

#### Critical

- None

#### High

- **S1** — Trust boundary на данных плана: get_orders_from_opt_plan / build_procurement_items — минимальная валидация; load_code и размеры as-is → риск манипуляции котировками / resource pressure если глобальное состояние недостоверно.

#### Medium

- **S2** — print раскрывает коммерческие данные в stdout.
- **S3** — NDJSON debug_logs с ценами на диск.
- **S4** — print `{e}` может раскрывать детали ошибок.

#### Low

- **S5** — Global state / multi-tenant footgun.
- **S6** — Swallowed I/O на debug логах.
- **S7** — Нет auth в модуле — контракт на вызывающий слой.
- **S8** — Пути файлов фиксированы через get_debug_log_path — path traversal не из этого файла.

### Code Quality Findings (from reviewer)

#### High

- None (дополнительно к архитектуре/безопасности)

#### Medium

- **Q1** — Дрейф production trim vs _calc_trim_components; mismatch base_price vs base_price_1_2m для width waste (~239–242 vs ~892–896).
- **Q2** — Повтор fallback order / regex / reconstruction между breakdown функциями.
- **Q3** — Broad except Exception → silent data loss.
- **Q4** — Импорты json/os/time внутри hot loop в build_price_rows.
- **Q5** — Глубокая вложенность в _calc_trim_components и build_procurement_items.

#### Low

- **Q6** — Dead/redundant imports (unused OPT_CASCADING_PLAN и др.).

- **Q7** — Слабая типизация, mutable-default smell на price_rows.
- **Q8** — Комментарий про поперечный рез вводит в заблуждение.
- **Q9** — Дублирующиеся магические числа, двойной комментарий # Базовая цена.
- **Q10** — Пробелы в тестах для production/breakdown/get_orders.

---

## Next Steps

1. Зафиксировать контракт доверия к данным плана (**S1**): валидация `load_code`, размеров, лимиты; документировать источник истины для плана.
2. Спланировать декомпозицию god-модуля (**A1**) и границы слоёв (**A2**) — использовать команду **`/refactor`** для структурных изменений.
3. Унифицировать trim/разбор резов (**A3**, **Q1**, **Q2**) через общий код-путь и тесты на расхождения production vs базовая ветка.
4. Убрать или параметризовать `OPT_*` (**A4**) для тестируемости и отсутствия гонок при параллели.
5. Вынести I/O логирования из hot path, заменить print на безопасное логирование (**A7**, **S2**, **S3**, **S4**); поведенческие правки безопасности — через **`/implement`**.
6. Добавить батчинг/кэш для запросов цен и сырья (**A6**); сузить `except` и устранить silent loss (**Q3**); поднять импорты из цикла (**Q4**).
7. Покрыть тестами критичные ветки после рефакторинга; закрыть пробелы по **Q10**.

---

## Note

- Для **структурных** изменений (разбиение модуля, порты, DIP) использовать **`/refactor`**.
- Для **поведенческих** исправлений безопасности (валидация плана, логирование, маскирование) использовать **`/implement`**.

---

*Конец отчёта.*
