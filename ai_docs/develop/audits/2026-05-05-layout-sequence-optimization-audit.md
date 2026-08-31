# Audit Report: Layout Sequence & Optimization Modules

**Date**: 2026-05-05  
**Scope**: `viz_modules/layout_sequence.py`, `core/optimization.py`  
**Audited by**: senior-reviewer + security-auditor + reviewer

---

## Executive Summary

Комплексный аудит двух критических модулей выявил **2 критические архитектурные проблемы** и **20 проблем среднего и низкого приоритета**. Модули содержат недопустимое глобальное изменяемое состояние, неустойчивые API контракты и уязвимости безопасности, которые делают код небезопасным для параллельной многопользовательской обработки. Основная рекомендация: исправить критические проблемы **перед** использованием этого кода в production API или многопоточном контексте.

**Overall Health Score**: **2.0/10**

| Severity | Architecture | Security | Code Quality | **Total** |
|----------|-------------|----------|--------------|----------|
| Critical | 2           | 0        | 0            | **2**    |
| High     | 4           | 2        | 1            | **7**    |
| Medium   | 5           | 2        | 3            | **10**   |
| Low      | 3           | 2        | 3            | **8**    |

**Executive Recommendation**: Обязательно устраните **2 критические проблемы** неявного глобального состояния и нестабильного контракта результата, а также **мертвый код** (вторая ветвь `if OPT_CASCADING_PLAN...`) перед использованием этого кода в системах с параллельной обработкой, повторными попытками или API, где вызывающие могут забыть о синхронизации глобального состояния.

---

## Critical Issues (fix immediately)

### [A1] Implicit Shared State — Unsafe for Concurrency

**Category**: Architecture  
**Location**: `core/optimization.py` — module-level globals `OPT_PLAN`, `OPT_CASCADING_PLAN`, `OPT_CASCADING_PLAN_BY_LOAD`  
**Impact**:  
- Function `build_layout_sequence()` читает глобальное состояние, которое не обновляется внутри `optimize_with_cascading_longitudinal_cuts()`  
- Вызывающие должны синхронизировать состояние вручную  
- **Критическая угроза**: в многопользовательских системах, асинхронных обработчиках или при повторных попытках глобальное состояние перезаписывается между запросами → утечка данных между пользователями, неправильные планы раскладки  
- API забвение вызывающего приводит к тихим ошибкам и состояниям гонки

**Fix**:  
1. Конвертировать глобалы в явный параметр `OptimizationState` DTO, передаваемый через `Depends`  
2. Убедиться, что `build_layout_sequence()` принимает `state: OptimizationState` аргумент вместо чтения глобалов  
3. Добавить type hints и валидацию, чтобы гарантировать, что состояние всегда явное и изолировано per-request

---

### [A2] Unstable Result Type — Runtime Type Inference

**Category**: Architecture  
**Location**: `viz_modules/layout_sequence.py` — функция `build_layout_sequence()`  
**Impact**:  
- Функция возвращает либо сгруппированный список словарей (`list[dict]` с полями группировки), либо плоский список plate-словарей  
- Вызывающие должны выводить форму результата во время выполнения (инспекция ключей, проверка наличия ключей)  
- Нет явного типа возврата → IDE не может помочь, нет документации о том, какая форма возвращается в каких условиях  
- **Критическая угроза**: консьюмеры кода падают или производят неправильный вывод, если случайно обработают неправильную форму

**Fix**:  
1. Создать явные типы возврата: `LayoutSequenceFlat = list[PlateDict]` и `LayoutSequenceGrouped = dict[str, list[PlateDict]]`  
2. Вернуть discriminated union: `LayoutSequence = Union[LayoutSequenceFlat, LayoutSequenceGrouped]` + маркер тип (enum) или отдельные функции  
3. Документировать условия: "возвращает grouped если есть separator, иначе flat"  
4. Обновить все вызовы, чтобы явно обрабатывать оба случая

---

## High Priority Issues (fix soon)

### [A3] Unreachable Dead Code — Cascading Branch Duplicate

**Category**: Architecture  
**Location**: `core/optimization.py` — вторая ветвь `if OPT_CASCADING_PLAN...` после ранней `return` в первом блоке  
**Impact**:  
- Сотни строк кода после `return` в первой ветке `if OPT_CASCADING_PLAN...` никогда не выполняются  
- Мертвый код усложняет тестирование, maintenance и увеличивает поверхность атаки при будущих изменениях  
- Разработчики теряют время, отлаживая или изменяя недостижимый код

**Fix**:  
1. Удалить полностью вторую ветвь `if OPT_CASCADING_PLAN...` (мертвый код)  
2. Убедиться, что первая ветка покрывает все необходимые cascading логики  
3. Добавить комментарий или issue, если вторая ветка была там для будущего функционала

---

### [A4] God Module: layout_sequence

**Category**: Architecture  
**Location**: `viz_modules/layout_sequence.py` — весь модуль (~700+ строк)  
**Impact**:  
- Один модуль содержит: DB запросы, конфиг, глобальные OPT_*, sequencing логику, separator выбор, layout_uid генерацию, debug-файлы  
- Нарушает Single Responsibility Principle → трудно тестировать, трудно переиспользовать, высокий coupling  
- Изменение одного аспекта требует понимания всей логики

**Fix**:  
1. Экстрактировать DB слой в `layout_repository.py`  
2. Экстрактировать sequencing логику в `sequencing_service.py`  
3. Экстрактировать separator выбор в `separator_selector.py`  
4. Экстрактировать debug/logging в `layout_logger.py`  
5. Оставить `layout_sequence.py` как орхестратор/фасад

---

### [A5] God Module: optimization

**Category**: Architecture  
**Location**: `core/optimization.py` — весь модуль (~2500+ строк)  
**Impact**:  
- Один модуль смешивает: PuLP 1D/2D модели, verify_coverage функции, глобальное состояние OPT_*, FFD-packing логику  
- Экспортирует 40+ функций без явной иерархии → трудно понять публичный API  
- Высокий риск регрессий при изменении

**Fix**:  
1. Создать `optimization_models.py` для PuLP model builders (1D, 2D)  
2. Создать `optimization_verifier.py` для coverage checks  
3. Создать `packing_ffd.py` для FFD-related логики  
4. Оставить `optimization.py` как орхестратор / entry point  
5. Явно определить public API с type hints

---

### [A6] Dependency Direction Violation

**Category**: Architecture  
**Location**: `viz_modules/layout_sequence.py` → `core/optimization.py` (module globals)  
**Impact**:  
- layout зависит от глобального состояния в optimizer, вместо того чтобы получать его через Dependency Injection  
- Нарушает Dependency Inversion Principle → трудно тестировать layout отдельно от optimizer

**Fix**:  
1. Создать `OptimizationStateDTO` (Pydantic model)  
2. Передавать `state: OptimizationStateDTO` в качестве параметра в `build_layout_sequence()`  
3. Убедиться, что optimizer заполняет эту DTO и передает в layout
4. Добавить Depends-инъекцию на уровне FastAPI endpoints

---

### [S1] Unguarded Agent Log Writes — Disk Growth & Information Disclosure

**Category**: Security  
**Location**: `core/optimization.py` — `open(..., "a")` writes для agent logs  
**Impact**:  
- Agentные логи пишутся без проверки debug gates → неограниченный рост диска  
- Логи содержат внутреннюю структуру раскладки, заказов, оптимизационные решения  
- **Угроза**:披露competitive intelligence, структуры заказов, pricing logic  
- В shared-hosting или multi-tenant сценариях один клиент может читать логи другого

**Fix**:  
1. Обернуть все `open(..., "a")` writes в единую, управляемую гейтом `DEBUG_LOGS` систему  
2. Убедиться, что логи пишутся только если явно включены в конфиге  
3. Переместить логи в `$TEMP` или `logs/` директорию с правильными permissions  
4. Ротировать логи при достижении размера лимита (логирование по ротации)  
5. Добавить authentication-check перед доступом к log файлам

---

### [S2] Process-Wide OPT_* Globals in Multi-User Workers

**Category**: Security  
**Location**: `core/optimization.py` — module-level globals в worker process  
**Impact**:  
- OPT_* переменные разделены между всеми потоками/greenlet'ами в процессе  
- В multi-user асинхронной системе одного пользователя может перезаписать optimization plan другого  
- **Критическая угроза**: data leakage, неправильные раскладки для других пользователей, мультитенант-изоляция нарушена

**Fix**: (см. [A1] выше — то же самое)

---

### [Q1] Bare `except:` in PuLP Solution Parsing

**Category**: Code Quality  
**Location**: `core/optimization.py:2301–2330` — PuLP solution value extraction  
**Impact**:  
- `except:` без типа ловит даже `KeyboardInterrupt`, `SystemExit`  
- Скрывает неоправданные исключения (баги, недостаток памяти) под тишину  
- Отладка практически невозможна

**Fix**:  
1. Заменить `except:` на `except (ValueError, KeyError) as e:`  
2. Залогировать исключение как WARNING перед возвратом fallback  
3. Добавить unit тесты для случаев, когда PuLP не находит решение

---

## Medium Priority Issues (plan for next sprint)

### [A7] In-Place Mutation of Plan/Cut Dicts

**Category**: Architecture  
**Location**: `core/optimization.py` — functions mutating plan/cut parameters  
**Impact**:  
- Функции модифицируют переданные дикты in-place без явного указания в docstring  
- Вызывающие могут не ожидать побочных эффектов → баги  
- Усложняет тестирование и debugging

**Fix**:  
1. Вернуть новые дикты вместо модификации in-place  
2. Документировать "immutable inputs" в docstrings  
3. Рассмотреть dataclasses или TypedDict для более явной структуры

---

### [A8] Duplicated/Divergent Paths vs `_build_sequence_from_plan`

**Category**: Architecture  
**Location**: `viz_modules/layout_sequence.py` — несколько путей sequencing  
**Impact**:  
- Существует несколько путей sequencing (cascading branch, `_build_sequence_from_plan`, основной путь)  
- Они diverge в обработке edge cases → несогласованное поведение  
- Высокий риск ошибок при изменении одного пути

**Fix**:  
1. Консолидировать в единый `_build_sequence()` путь  
2. Параметризовать различия (e.g., `strategy: SequencingStrategy enum`)  
3. Удалить ветки, которые были заменены

---

### [A9] Separator Choice / Misleading API — `_choose_best_separator`

**Category**: Architecture  
**Location**: `viz_modules/layout_sequence.py` — function `_choose_best_separator()`  
**Impact**:  
- Параметр `load_code` принимается, но не используется в lookup  
- Документация обещает выбор по `load_code`, но реально выбор по другой логике  
- Вызывающие передают `load_code`, ожидая эффекта, но его нет

**Fix**:  
1. Либо использовать `load_code` в логике выбора, либо удалить параметр  
2. Обновить docstring с явным объяснением алгоритма выбора  
3. Добавить тест, проверяющий, что `load_code` влияет на результат (если намеревается)

---

### [A10] Inconsistent Defaults — 800 vs 8

**Category**: Architecture  
**Location**: `core/optimization.py` и `viz_modules/layout_sequence.py`  
**Impact**:  
- Константы `800` и `8` используются в разных местах для normalization  
- Нет явных named constants → трудно понять, что означают магические числа  
- Несогласованность может привести к ошибкам при изменении логики

**Fix**:  
1. Создать `constants.py` с `LOAD_CODE_SCALE_FACTOR = 800`, `LOAD_CODE_DIVISOR = 8` и т.д.  
2. Использовать константы везде  
3. Документировать, откуда они берутся

---

### [A11] Observability — Inconsistent Logging Strategy

**Category**: Architecture  
**Location**: `core/optimization.py` и `viz_modules/layout_sequence.py`  
**Impact**:  
- Используются `print()`, env-gated debug writes, JSONL writes  
- Нет единого logging framework (не используется `logging` модуль)  
- Трудно включить/отключить логирование, трудно ротировать логи, нет уровней логирования

**Fix**:  
1. Создать `app/core/logging.py` с конфигурацией `logging` модуля  
2. Заменить все `print()` на `logger.info()`, `logger.debug()`, etc.  
3. Использовать `DEBUG` переменную конфига для управления уровнем логирования  
4. Ротировать логи с помощью `RotatingFileHandler`

---

### [S3] Layout Debug Logs — Unguarded Writes

**Category**: Security  
**Location**: `viz_modules/layout_sequence.py` — debug_logs writes  
**Impact**:  
- Аналогично [S1], но в layout_sequence  
- Пишет debug информацию без проверки gates

**Fix**: (см. [S1] выше — консолидировать logging)

---

### [S4] Verbose Stdout/Logs — PII / Commercial Data Leakage

**Category**: Security  
**Location**: `core/optimization.py` и `viz_modules/layout_sequence.py` — `[VISUAL]`, topology prints  
**Impact**:  
- Логируются: topology orders, layout details, customer orders structure, pricing data  
- В shared logging систем эта информация может утечь  
- **Угроза**:披露competitive intelligence, customer data, pricing

**Fix**:  
1. Убедиться, что sensitive data (customer names, order details, pricing) не выводятся в логи  
2. Создать separate "audit log" для sensitive операций (отдельно от debug логов)  
3. Шифровать audit logs или хранить только hashes

---

### [Q2] Duplicated Legacy Branch

**Category**: Code Quality  
**Location**: `viz_modules/layout_sequence.py` — cascading branch duplicate  
**Impact**:  
- (Тот же, что и [A3], но с точки зрения Code Quality)

**Fix**: (см. [A3] выше)

---

### [Q3] Broad `except Exception: pass`

**Category**: Code Quality  
**Location**: `core/optimization.py` — several places  
**Impact**:  
- Ловит все исключения, даже неожиданные, скрывает баги

**Fix**:  
1. Заменить на специфичные типы исключений  
2. Логировать исключения перед игнорированием  
3. Добавить тесты для случаев ошибок

---

### [Q4] Late `defaultdict` Import

**Category**: Code Quality  
**Location**: `core/optimization.py` — `defaultdict` imported late в модуле  
**Impact**:  
- Импорты должны быть в начале файла (PEP 8)  
- Создает confusion про то, где `defaultdict` определяется

**Fix**:  
1. Переместить import в секцию imports в начале файла

---

## Low Priority Issues

### [A12] `except: pass` on PuLP Value Reads

**Category**: Architecture  
**Location**: `core/optimization.py` — PuLP solution value extraction  
**Impact**: (см. [Q1] выше)  
**Fix**: (см. [Q1] выше)

---

### [A13] Organic Module Structure — Duplicate Imports, FFD Beside ILP

**Category**: Architecture  
**Location**: `core/optimization.py` — whole module layout  
**Impact**:  
- Импорты повторяются, функции перепутаны  
- FFD packing logic рядом с ILP solver logic без явной иерархии

**Fix**: (см. [A5] выше — модульный рефактор)

---

### [A14] Scalability — Globals + Per-Call DB/Config Walks

**Category**: Architecture  
**Location**: `core/optimization.py` + `viz_modules/layout_sequence.py`  
**Impact**:  
- Каждый вызов re-reads конфиг, re-connects к БД  
- Нет session isolation layer → потенциальные утечки соединений  
- Не масштабируется при большом количестве параллельных запросов

**Fix**:  
1. Создать `OptimizationContext` с re-используемой сессией БД, конфигом  
2. Передавать через Depends в все функции  
3. Убедиться, что сессия открывается/закрывается правильно

---

### [Q5] Misleading Docstring

**Category**: Code Quality  
**Location**: `viz_modules/layout_sequence.py` — `_choose_best_separator()` docstring  
**Impact**: (см. [A9] выше)  
**Fix**: (см. [A9] выше)

---

### [Q6] Wrong Type Hint

**Category**: Code Quality  
**Location**: `viz_modules/layout_sequence.py` — `load_code: int = None`  
**Impact**:  
- Type hint говорит `int`, но значение может быть `None`  
- Правильно: `load_code: Optional[int] = None` или `load_code: int | None = None`

**Fix**:  
1. Обновить type hint: `load_code: int | None = None`

---

### [Q7] Magic Numbers

**Category**: Code Quality  
**Location**: Throughout both modules  
**Impact**:  
- Числа `800`, `8`, `1000`, `100`, etc. используются без контекста  
- Трудно понять, что они означают

**Fix**: (см. [A10] выше)

---

## Priority Matrix

| ID | Issue | Category | Severity | Effort | Priority |
|----|-------|----------|----------|--------|----------|
| A1 | Implicit global state | Architecture | **Critical** | High | **P0** — до production |
| A2 | Unstable result type | Architecture | **Critical** | High | **P0** — до production |
| A3 | Dead cascading branch | Architecture | High | Low | **P1** — этот спринт |
| S1 | Unguarded log writes | Security | High | Medium | **P1** — этот спринт |
| S2 | OPT_* globals multi-user | Security | High | High | **P0** — до production (см. A1) |
| Q1 | Bare except: | Code Quality | High | Low | **P1** — этот спринт |
| A4 | God module layout_sequence | Architecture | High | Very High | **P2** — next sprint |
| A5 | God module optimization | Architecture | High | Very High | **P2** — next sprint |
| A6 | Dependency direction | Architecture | High | High | **P1** — this sprint |
| A7 | In-place mutation | Architecture | Medium | Medium | **P2** — next sprint |
| A8 | Duplicated sequencing paths | Architecture | Medium | High | **P2** — next sprint |
| A9 | Misleading _choose_best_separator | Architecture | Medium | Low | **P2** — next sprint |
| A10 | Inconsistent defaults | Architecture | Medium | Low | **P2** — next sprint |
| A11 | Inconsistent logging | Architecture | Medium | Medium | **P2** — next sprint |
| S3 | Layout debug logs | Security | Medium | Low | **P1** — this sprint (part of S1) |
| S4 | Verbose logs → data leakage | Security | Medium | Medium | **P1** — this sprint |
| Q2 | Duplicated legacy branch | Code Quality | Medium | Low | **P1** — (part of A3) |
| Q3 | Broad except Exception | Code Quality | Medium | Low | **P1** — this sprint |
| Q4 | Late defaultdict import | Code Quality | Medium | Low | **P3** — cleanup |
| Q5 | Misleading docstring | Code Quality | Low | Low | **P3** — cleanup (part of A9) |
| Q6 | Wrong type hint | Code Quality | Low | Low | **P3** — cleanup |
| Q7 | Magic numbers | Code Quality | Low | Low | **P3** — cleanup (part of A10) |

---

## Next Steps

### Immediate (before production use)
1. **[A1] + [A2] + [S2]**: Перейти на Dependency Injection для optimization state
   - Создать `OptimizationStateDTO` с type hints
   - Обновить `build_layout_sequence()` сигнатуру
   - Убедиться, что state всегда явное и изолировано per-request

### This Sprint
2. **[A3]**: Удалить мертвый cascading branch (~300-500 строк savings)
3. **[A6]**: Перейти на injected optimizer state (DI)
4. **[Q1] + [Q3]**: Заменить bare `except:` на специфичные типы + логирование
5. **[S1] + [S3] + [S4]**: Консолидировать логирование в unified `logging` module с gating

### Next Sprint
6. **[A4] + [A5]**: Рефактор God modules → extracted services (очень высокий effort)
7. **[A7] + [A8]**: Консолидировать sequencing paths
8. **[A9] + [A10]**: Исправить misleading API и constants

### Backlog
9. **[A11] + [A14]**: Scalability improvements (session isolation, context management)
10. **[Q5] + [Q6] + [Q7]**: Code cleanup (type hints, docstrings, constants)

---

## Recommendation for User

**Перед использованием в production:**
- ✅ Обязательно исправьте [A1], [A2], [S2] → Dependency Injection
- ✅ Удалите мертвый код [A3]
- ✅ Добавьте proper exception handling [Q1]
- ✅ Консолидируйте логирование [S1], [S3], [S4]

**Эти 4 задачи позволят коду быть безопасным для многопользовательского использования.**

Остальные issues могут быть адресованы в следующих спринтах (большой effort, но не блокирующие).

---

**Report Generated**: 2026-05-05 11:12 UTC+3  
**Scope**: `viz_modules/layout_sequence.py` (main sequencing logic), `core/optimization.py` (PuLP models, FFD packing)  
**Reviewers**: senior-reviewer (architecture), security-auditor (security), reviewer (code quality)
