# Аудит пакета core/optimization

**Дата**: 2026-05-05
**Область**: core/optimization/ (рефакторинг монолита)
**Модели**: senior-reviewer + security-auditor + reviewer (Composer 2)

---

## Краткое резюме

**Индекс здоровья**: 4.0/10 (формула skill: −2 за Critical max 6, −0.5 за High max 3, −0.1 за Medium max 1; Critical=1, High=6, Medium=10)

| Серьёзность | Архитектура | Безопасность | Качество кода | Всего |
|-------------|------------|--------------|-----------------|-------|
| Critical    | 1          | 0            | 0               | 1     |
| High        | 3          | 1            | 2               | 6     |
| Medium      | 3          | 2            | 5               | 10    |
| Low         | 3          | 3            | 4               | 10    |

**Рекомендация**: устранить критический контракт ошибок (пустой `{}`) и High-по приоритету (god-module, логи, сложность 2D) перед расширением API.

---

## Критические проблемы

- [A1] **`core/optimization/_implementation.py` + `app/services/optimization_service.py` — пустой `{}` как единственный сигнал ошибки.** В `_optimize_2d_with_lengths` при отсутствии PuLP, пустом `orders_2d`, статусе решателя `Infeasible`/`Undefined` и в ряде веток возвращается `{}`. В `optimize_cuts_pulp` пустой результат не пишет в `OPT_PLAN`, но вызывающий код всё равно получает неразличимые случаи. `OptimizationService.optimize` (`app/services/OptimizationService.py`) трактует ложный `result` как «нет плана» без кода причины (solver, validation, import). Для API и UI это риск молчаливого провала: нет стабильного контракта «успех / ошибка / частичный план».

**Предлагаемое исправление**: ввести явный контракт результата (например, структурированный объект или union «успех / ошибка / частичный план» с полем `reason`/`code`: `solver_infeasible`, `pulp_missing`, `validation_failed`, `empty_orders`, и т.д.); не возвращать голый `{}` без метаданных; на границе `OptimizationService.optimize` и API маппить коды в HTTP/ответ клиенту, чтобы UI не угадывал по отсутствию ключа.

---

## Высокий приоритет

### Архитектура

- [A2] **`core/optimization/_implementation.py` — остаточный god-module после выноса пакета.** Один файл объединяет: `verify_coverage`, legacy `apply_width_optimization` / `optimize_cuts_pulp`, конфиг `OptimizationConfig`, полный 2D пайплайн (геометрия → ILP → solve → разбор переменных → сортировка под цех → пост-коррекция / `no_sources_keys` / `PlateAudit`, слоты и `plate_assignments`), 1D ILP, десятки встроенных блоков записи в произвольные лог-файлы. Нарушение SRP: та же «взрывоопасность» изменений, что у монолита, несмотря на соседние модули `geometry.py`, `ilp_model.py`, `order_dispatch.py`.

- [A3] **`core/optimization/_implementation.py` ↔ `core/optimization/orchestrator.py` — хрупкий bootstrap импортов.** В конце `_implementation.py` выполняется `from core.optimization.orchestrator import optimize_with_cascading_longitudinal_cuts`, а `orchestrator.py` внутри публичной функции делает ленивый `import core.optimization as pkg` для вызова `pkg._optimize_*`. Цикл на уровне пакета обходится порядком загрузки и ленивым импортом; любой новый top-level импорт `core.optimization` или `_implementation` из `orchestrator`/`validation` легко восстанавливает частично инициализированный модуль (типичный риск «рефакторинга в пакет» без явного ядра).

- [A4] **`core/optimization/context.py` + `core/optimization/__init__.py` (`_OptimizationModule`) — TLS вместо глобалей, ограничение модели конкурентности.** Потоколокальные прокси `OPT_PLAN`, `OPT_WIDTH_PRIORITY` и т.д. устраняют гонки между **потоками**, но не изолируют две логически параллельные оптимизации на **одном OS-потоке** (например, сопрограммы asyncio без вынесения блокирующего вызова в executor). `OptimizationService.optimize` синхронный, но любой прямой вызов из `async def` без изоляции может смешать TLS-состояние для `OPT_*` между запросами.

### Безопасность

- [S1] Фиксированные пути логов + запись NDJSON с полями заказа вне единого gated/debug пути — риск утечки коммерческих данных и роста диска.

### Качество кода

- [Q1] Огромная `_optimize_2d_with_lengths` в `_implementation.py`.

- [Q2] `_canonical_length` → 0.0 при ошибке парсинга в `geometry.py`.

---

## Средний приоритет

### Архитектура

- [A5] **`core/optimization/__init__.py` и `__all__` в `_implementation.py` — размытая граница публичного API.** Реэкспорт `from ._implementation import *` плюс длинный `__all__`, включающий `_optimize_2d_with_lengths`, `_optimize_1d_widths_only`, `_peek_order_info`, `_build_proportional_slot_lists`, `_get_next_order_info`, `_append_actions`, `_group_plate_lengths`, `_residual_phys_key`, `_build_residual_balance_constraints` и др. Потребители в репозитории импортируют и пакет, и подмодули (`viz_modules`, `bot`, `tests`). Это фиксирует внутренности как де-факто контракт и усложняет дальнейшее разрезание без поломок.

- [A6] **Жёсткая связность с остальным `core` (не самодостаточный пакет).** `geometry.py` тянет `core.config_and_data`. `ilp_model.py` — `core.price_db`, `core.config_and_data`, плюс `debug_log`. `_implementation.py` — `cfg`, `canonical_plate_key`, `get_debug_log_path`, `core.plate_audit.PlateAudit`, локальные пути к `debug-*.log`. Для масштабирования граница «optimization package» не ясна: домен и инфраструктура смешаны по графу зависимостей.

- [A7] **Две политики отладочного I/O: `debug_log.py` vs прямые файлы в `_implementation.py`.** Часть путей уважает `OPT_DEBUG_LOG`, часть — нет.

### Безопасность

- [S2] Несогласованное redaction / `OPT_DEBUG_LOG`.

- [S3] TLS OPT_* и asyncio на одном потоке.

### Качество кода

- Дублирование `_peek_order_info` / `_get_next_order_info` в `order_dispatch.py`.
- DRY: повторяющиеся JSON debug блоки в нескольких файлах.
- Возможная несогласованность `normalize_load_code` vs сырой `load_code` в `order_dispatch.py`.
- Пробелы в тестах: `order_dispatch`, `optimize_tracks` / FFD.
- Слабая типизация на границах (`Any`, `list` без элементов).

---

## Низкий приоритет / пожелания

### Архитектура

- [A8] **`orchestrator.py` и большой объём `print` в `_implementation.py` — неидиоматично для сервисного слоя.**

- [A9] **`validation.py` — только `ValueError`.** Нет узких доменных исключений.

- [A10] **`ffd_packing.py` — удачный контрпример по границам** (stdlib-only).

### Безопасность

- [S4] ValueError раскрывает детали ввода.

- [S5] print → stdout.

- [S6] broad except на debug I/O.

### Качество кода

- Аннотации в `context.py` proxy methods.
- `ffd_packing` типы.
- Магические литералы без конфига.
- `{}` после ImportError PuLP без структурированного сигнала.

---

## Матрица приоритетов

| ID | Issue | Severity | Effort | Priority |
|----|-------|----------|--------|----------|
| A1 | Пустой `{}` как единственный сигнал ошибки (`_implementation.py`, `OptimizationService`) | Critical | High | P0 |
| A2 | God-module в `_implementation.py` | High | High | P0 |
| S1 | Логи/NDJSON вне gated пути, риск утечи данных и диска | High | Medium | P0 |
| Q1 | Огромная `_optimize_2d_with_lengths` | High | High | P0 |
| A3 | Хрупкий bootstrap импортов `_implementation` ↔ `orchestrator` | High | Medium | P1 |
| A4 | TLS OPT_* и параллелизм на одном потоке (asyncio) | High | Medium | P1 |
| Q2 | `_canonical_length` → 0.0 при ошибке парсинга (`geometry.py`) | High | Low | P1 |
| A5 | Размытая граница публичного API (`__init__`, `__all__`) | Medium | Medium | P2 |
| A6 | Связность пакета с остальным `core` | Medium | High | P2 |
| A7 | Две политики debug I/O (`debug_log` vs файлы в `_implementation`) | Medium | Medium | P2 |
| S2 | Несогласованное redaction / `OPT_DEBUG_LOG` | Medium | Medium | P2 |
| S3 | TLS OPT_* и asyncio (дублирует аспект A4) | Medium | Medium | P2 |
| — | Дубли `_peek_order_info` / `_get_next_order_info` (`order_dispatch.py`) | Medium | Low | P2 |
| — | Повторяющиеся JSON debug блоки | Medium | Low | P2 |
| — | `normalize_load_code` vs сырой `load_code` (`order_dispatch.py`) | Medium | Low | P2 |
| — | Пробелы в тестах (`order_dispatch`, FFD / `optimize_tracks`) | Medium | Medium | P2 |
| — | Слабая типизация на границах | Medium | Medium | P2 |
| A8 | `print` в `orchestrator` / `_implementation` | Low | Low | P3 |
| A9 | Только `ValueError` в `validation.py` | Low | Low | P3 |
| A10 | (информативно) `ffd_packing` как хорошая граница | Low | — | P3 |
| S4 | ValueError раскрывает детали ввода | Low | Low | P3 |
| S5 | print → stdout | Low | Low | P3 |
| S6 | broad except на debug I/O | Low | Low | P3 |
| — | Аннотации proxy methods в `context.py` | Low | Low | P3 |
| — | Типы в `ffd_packing` | Low | Low | P3 |
| — | Магические литералы | Low | Low | P3 |
| — | `{}` после ImportError PuLP без структурированного сигнала | Low | Low | P3 |

---

## Следующие шаги

1. Спроектировать и внедрить стабильный контракт результата оптимизации (успех / ошибка / частичный план с кодом причины), убрать неразличимый пустой `{}` на границе сервиса и API.
2. Вынести или разрезать `_implementation.py`: отдельные модули для legacy 1D/2D, записи логов, пост-обработки и публичного пайплайна; снизить связность с произвольными файловыми путями.
3. Устранить циклическую/ленивую схему импортов между `_implementation.py` и `orchestrator.py` (явное «ядро» пакета, односторонние зависимости).
4. Зафиксировать политику отладочного I/O: единый gated-путь, redaction, согласованность с `OPT_DEBUG_LOG`; убрать или централизовать фиксированные пути и NDJSON с чувствительными полями.
5. Разбить `_optimize_2d_with_lengths`, добавить тесты на `order_dispatch`, FFD / `optimize_tracks`; уточнить поведение `_canonical_length` при ошибке парсинга в `geometry.py`.
6. Документировать ограничение TLS для вызовов из asyncio и при необходимости изолировать вызовы через executor или контекст с явной передачей состояния.
7. По мере рефакторинга сузить публичный API пакета (`__all__`, реэкспорты) и постепенно мигрировать потребителей с приватных символов.
