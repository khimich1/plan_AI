# Отчёт: декомпозиция `_implementation.py` (тонкий фасад ref‑impl)

**Дата:** 2026-05-07  
**Оркестрация:** `orch-2026-05-07-15-00-ref-impl-decompose`  
**Статус:** Завершено  
**Оригинал (EN):** [`2026-05-07-ref-impl-decompose-optimization-implementation.md`](./2026-05-07-ref-impl-decompose-optimization-implementation.md)

## Кратко — цель

Заменить **монолит `core/optimization/_implementation.py`** (~1900+ строк) на **тонкий фасад импорта/реэкспорта** (~100 строк) и **сфокусированные модули** в `core/optimization/`. Сохранить **текущую публичную поверхность** для `core.optimization` (`from ._implementation import *` / стабильный `__all__`) и **контракт с оркестратором**, описанный в `core/optimization/orchestrator.py`.

**Эталонный план:** `ai_docs/develop/plans/2026-05-07-refactor-optimization-implementation.md`

## Выполненные задачи (OPT‑REF‑001 … OPT‑REF‑010)

| ID | Название | Зависимости | Результат |
|----|----------|-------------|-----------|
| **OPT‑REF‑001** | Вынести `OptimizationConfig` | — | `optimization_config.py`: `OptimizationConfig`, `DEFAULT_CONFIG`, `OLD_CONFIG`. |
| **OPT‑REF‑002** | Вынести coverage и хелперы PuLP для qty | — | `coverage_verify.py` (`verify_coverage`), `pulp_qty.py` (`_opt_1d_pulp_nonneg_qty`). |
| **OPT‑REF‑003** | Собрать отладочные хелперы реализации | — | `optimization_debug_impl.py` (вместе с использованием `debug_log.py` по плану); строки `location` в логах приведены к реальным местам вызова. |
| **OPT‑REF‑004** | Вынести legacy‑адаптеры ширины | — | `legacy_width_plan.py`: `_group_plate_lengths`, `_append_actions`, `apply_width_optimization`, `optimize_cuts_pulp`. |
| **OPT‑REF‑005** | Вынести размер батчей вторичных `z_sec` | — | `secondary_batches.py`: `_batch_sizes_for_secondary_z_sec`. |
| **OPT‑REF‑006** | Вынести оптимизатор ширин 1D | OPT‑REF‑002 | `optimize_1d_widths.py`: `_optimize_1d_widths_only`. |
| **OPT‑REF‑007** | Фаза 2D: подготовка + решение ILP | OPT‑REF‑001 | `optimize_2d/prep_solve.py`: `run_two_d_phase_a`. |
| **OPT‑REF‑008** | Фаза 2D: резы + порядок + родители | OPT‑REF‑007, OPT‑REF‑005 | `optimize_2d/extract_cuts.py`: `extract_two_d_phase_b` (связь с secondary batches). |
| **OPT‑REF‑009** | Фаза 2D: пост‑коррекция + аудит + атрибуция | OPT‑REF‑008, OPT‑REF‑002 | `optimize_2d/finalize.py`: `run_two_d_phase_finalize` (проверка покрытия, назначения, остатки). |
| **OPT‑REF‑010** | Уплотнить фасад + стабильный `__all__` | OPT‑REF‑009, 006, 004, 003 | `_implementation.py` агрегирует импорты/`__all__`; оркестрация только через делегирование (например `_optimize_2d_with_lengths` из `optimize_2d/with_lengths.py`). |

## Карта новых / основных модулей (`core/optimization/`)

| Путь | Зона ответственности |
|------|----------------------|
| `_implementation.py` | Тонкий фасад: реэкспорт TLS/context, символов geometry/ILP/FFD; вход в 2D через `optimize_2d/with_lengths`; **последней строкой** — подтягивание `optimize_with_cascading_longitudinal_cuts` из оркестратора (см. раздел про циклические импорты). |
| `optimization_config.py` | Dataclass конфигурации и значения по умолчанию. |
| `coverage_verify.py` | Проверка спроса/покрытия (`verify_coverage`). |
| `pulp_qty.py` | Хелпер неотрицательного количества PuLP для ветки 1D. |
| `optimization_debug_impl.py` | Отладочная инструментарий, вынесенный из бывших секций монолита. |
| `legacy_width_plan.py` | Legacy‑конвейер ширины / хелперы PuLP по ширине. |
| `secondary_batches.py` | Размер батчей вторичных `z_sec` (`_batch_sizes_for_secondary_z_sec`). |
| `optimize_1d_widths.py` | Современный оптимизатор только по ширине 1D (`_optimize_1d_widths_only`). |
| `optimize_2d/__init__.py` | Экспорт подпакета: точки входа фаз A/B + `TwoDPhaseAState`, `norm_demand_key`. |
| `optimize_2d/state.py` | Общее состояние / нормализация для фаз 2D. |
| `optimize_2d/prep_solve.py` | Фаза A: подготовка + сборка/решение ILP. |
| `optimize_2d/extract_cuts.py` | Фаза B: извлечение первичных/вторичных резов, родители и правила порядка. |
| `optimize_2d/finalize.py` | Фаза C: чекпоинты аудита, нормализация/пост‑коррекция, `verify_coverage`, назначение по слотам. |
| `optimize_2d/with_lengths.py` | Склейка фаз A→B→C: `_optimize_2d_with_lengths`. |

Соседние модули по роли не менялись: **`orchestrator.py`** (публичная точка входа + ленивый импорт пакета), **`context.py`** (TLS + `OPT_*`), **`ilp_model.py`**, **`ffd_packing.py`**, **`order_dispatch.py`**, **`geometry.py`**, **`validation.py`**, **`result_contract.py`**, **`debug_log.py`**.

## Регрессионные тесты (рекомендуемый набор)

Запуск из корня репозитория **с виртуальным окружением проекта / полным `requirements.txt`**:

```bash
pytest tests/test_optimization*.py tests/test_opt_1d_pulp_qty_extraction.py -v
```

**Файлы в явной области покрытия:**

- `tests/test_optimization_validation.py`
- `tests/test_optimization_baseline.py`
- `tests/test_optimization_config.py`
- `tests/test_optimization_result_contract.py`
- `tests/test_optimization_secondary_parent_assignment.py`
- `tests/test_optimization_semantics_and_tracks.py`
- `tests/test_optimization_thread_local_globals.py`
- `tests/test_optimization_verify_pulp_submodules_ref002.py`
- `tests/test_opt_1d_pulp_qty_extraction.py`

**Проверка на момент написания документа:** в этой среде автоматический прогон **не** полностью воспроизводим (нет `venv`; «голый» Python без полного дерева зависимостей — например `matplotlib` тянется через `core/__init__.py` → `visualization`). Перед релизом перезапустите те же команды в вашем venv на Windows.

## Риски и защита от циклических импортов

1. **`core.optimization.__init__.py`** сразу выполняет `from ._implementation import *`, поэтому при любом `import core.optimization` рано грузится `_implementation`.
2. **`_implementation.py`** импортирует **`optimize_with_cascading_longitudinal_cuts`** из **`orchestrator.py`** **в конце модуля** (`# Последним: оркестратор подтягивает пакет лениво внутри API`).
3. **`orchestrator.py`** не импортирует пакет на уровне модуля в тяжёлом пути; внутри `optimize_with_cascading_longitudinal_cuts` выполняется **`import core.optimization as pkg`** при делегировании в `pkg._optimize_2d_with_lengths` / `pkg._optimize_1d_widths_only`.

**Правило эксплуатации:** не добавлять **верхнеуровневый** `import core.optimization` (или цепочки, возвращающиеся к `_implementation`) в **`orchestrator.py`** или в другие модули, которые импортирует `_implementation` до завершения инициализации фасада — иначе получится цикл. Предпочитать **ленивые локальные импорты** в путях вызова (как уже сделано в оркестраторе).

**Подмодули:** импорты в `optimize_2d/*` должны оставаться **однонаправленными** к `ilp_model`, `geometry`, `secondary_batches`, `optimization_config` и т.д., без обратной подтяжки `_implementation`.

## Связанная документация

- План: [`ai_docs/develop/plans/2026-05-07-refactor-optimization-implementation.md`](../plans/2026-05-07-refactor-optimization-implementation.md)
- Воркспейс: `.cursor/workspace/active/orch-2026-05-07-15-00-ref-impl-decompose/` (`tasks.json`, `progress.json`, `links.json`)

## Следующие шаги (по желанию)

- После мерджей прогнать полный набор **`tests/test_optimization*.py`** + **`tests/test_opt_1d_pulp_qty_extraction.py`** в CI / venv проекта.
- Если автоматизация `links.json` расширится, привязать поле **`report`** к пути этого файла для удобного поиска.
