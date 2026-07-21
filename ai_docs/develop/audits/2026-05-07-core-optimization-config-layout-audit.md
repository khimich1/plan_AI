# Core Optimization, Config & Layout Sequence Audit Report

**Date**: 2026-05-07  
**Scope**: `core/optimization/` + `core/config_and_data.py` + `viz_modules/layout_sequence.py`  
**Audited by**: senior-reviewer + security-auditor + reviewer (orchestrated)

---

## Executive Summary

**Overall Health Score**: **6.0/10**

This review **extends** `ai_docs/develop/audits/2026-05-06-core-optimization-layout-sequence-audit.md` by bringing **`core/config_and_data.py`** into scope. The dominant cross-cutting theme is **dual global state**: **process-wide module globals** on `cfg` (`PLATES_*`, `PLATE_LOAD_DETAILS`, `PlateOrder.apply_to_globals()`, …) versus **thread-local** optimizer outputs (`OPT_PLAN`, `OPT_CASCADING_PLAN*`, …). **`viz_modules/layout_sequence.py` depends on both**, so any threading, worker-pool, or multi-job interleaving error can yield **wrong plans tied to the wrong order**—hard to reproduce and easy to misattribute.

There are **no Critical findings** in the aggregated counts. **High** severity is driven by **architecture** (monoliths, dead branches, coupling, ILP↔pricing boundary) and **security** (unconditional sensitive NDJSON to disk under repo paths), plus **quality** (unmaintainable function size, DRY gaps, print vs logging, broad silent `except`). **Recommendation**: before scaling concurrent or multi-tenant workloads, or running on shared/backed-up hosts, prioritize **gating/removing always-on debug file writes**, **unifying debug policy**, and **reducing cfg vs `OPT_*` implicit coupling** (explicit context objects or narrowed ports).

| Severity | Architecture | Security | Code Quality | Total |
|----------|-------------|----------|--------------|-------|
| Critical | 0           | 0        | 0            | **0** |
| High     | 7           | 2        | 4            | **13** |
| Medium   | 9           | 4        | 4            | **17** |
| Low      | 3           | 3        | 3            | **9** |

*Health score (per `.cursor/skills/audit-workflow/SKILL.md`): start 10; −2 per Critical (cap −6); −0.5 per High (cap −3); −0.1 per Medium (cap −1); floor 0; 1 decimal. Here: 10 − 3.0 (High cap) − 1.0 (Medium cap) = **6.0**.*

---

## Critical Issues (fix immediately)

*None in the aggregated severity counts for this scope.*

---

## High Priority Issues (fix soon)

### Architecture
- **[A1]** `_implementation.py` optimization monolith (~1.9k+ lines); SRP violated; regression isolation poor.
- **[A2]** `layout_sequence.py` merges viz, `pb.db`, process-wide `cfg`, and thread-local `OPT_*` without a single injected boundary.
- **[A3]** Unreachable control flow after early `return` in `build_layout_sequence`; large legacy body dead.
- **[A4]** `ilp_model.py` depends on `core.price_db.get_price` — commercial pricing tangled with ILP.
- **[A5]** Circular/fragile wiring: `orchestrator` lazy-imports package; `_implementation` imports orchestrator at EOF.
- **[A6]** Two global models: `cfg` globals vs TLS `OPT_*`; `layout_sequence` reads both — threading footgun.
- **[A7]** Optimizer depends concretely on `config_and_data` / implicit order contract — weak DIP.

### Security
- **[S1]** `get_price()` always appends NDJSON with commercial fields to disk — no debug gate (`price_db.py`).
- **[S2]** Unconditional agent NDJSON under repo paths in `layout_sequence` / `_implementation` — business-adjacent data beside source.

### Code Quality
- **[Q1]** `build_layout_sequence` / `_build_sequence_from_plan` — extreme size, nesting, mixed concerns.
- **[Q2]** DRY violation — duplicated solid/cut_groups/subgroup + `print` blocks in `layout_sequence.py`.
- **[Q3]** Widespread `print()` vs `logging` across optimization and layout paths.
- **[Q4]** Broad `except Exception` + `pass` around file/debug paths — silent failures.

---

## Medium Priority Issues (plan for next sprint)

### Architecture
- **[A8]** Non-idiomatic `sys.modules[__name__].__class__` for `OPT_*` interception.
- **[A9]** TLS proxies still implicit vs explicit parameters / `contextvars`.
- **[A10]** Duplicated narrative between `_build_sequence_from_plan` and dead legacy branch.
- **[A11]** Ad-hoc debug/trace I/O mixed with domain logic.
- **[A12]** Heavy sync PuLP work; no clear queue/worker contract in-module.
- **[A13]** `config_and_data.py` as god module (~1.2k+ lines).
- **[A14]** `PlateOrder` vs legacy globals — two parallel “current order” models.
- **[A15]** `apply_to_globals` / `set_plate_lists_from_text` → `kp_db` cache — hidden cross-layer effects.
- **[A16]** In-place mutation of plan dicts in `layout_sequence` — surprise for reuse.

### Security
- **[S3]** TLS `OPT_*` vs process-wide `cfg` — wrong pairing under concurrency (integrity/confidentiality).
- **[S4]** `LAST_PARSE_DIAGNOSTICS` retains `raw_input` — paste/PII leakage via exports/logs.
- **[S5]** Fixed local paths for DB/xlsx/docx — integrity risk if permissions loose.
- **[S6]** `get_debug_log_path()` — no `filename` normalization; future path traversal footgun.

### Code Quality
- **[Q5]** Weak typing / `Any` in layout, orchestrator, ILP.
- **[Q6]** Too-broad excepts where narrower types would do (`_canonical_load_code`, `format_reinforcement_from_load_code`, …).
- **[Q7]** `# type: ignore[index]` in `_canonical_target_order_key_tok`.
- **[Q8]** `config_and_data.py` overload without submodule split (maintainability).

---

## Low Priority / Suggestions

### Architecture
- **[A17]** `orchestrator` uses `print` for mode selection vs structured logging.
- **[A18]** `ffd_packing` isolated vs `geometry` still pulling `config_and_data` — uneven layering.
- **[A19]** Asymmetric `OPT_*` import style in `layout_sequence` (module vs function-local).

### Security
- **[S7]** `get_price(..., db_path=...)` API footgun if ever fed untrusted paths.
- **[S8]** Unbounded append-only debug logs — disk exhaustion / availability.
- **[S9]** Inconsistent debug policy vs `debug_log.py` `OPT_DEBUG_LOG` gate.

### Code Quality
- **[Q9]** `debug_log._dbg_open_append` no-op on error — hides permission/path failures.
- **[Q10]** Tests exist but branch coverage for layout/optimization combinations remains uncertain.
- **[Q11]** Dead second `if OPT_CASCADING_PLAN...` after `return` — readability noise.

---

## Priority Matrix

| ID   | Issue | Severity | Effort | Priority |
|------|-------|----------|--------|----------|
| S1   | Unconditional pricing NDJSON to disk | High | Low–Medium | P1 — before shared/backup/git-risk hosts |
| S2   | Unconditional layout/impl NDJSON under repo | High | Low–Medium | P1 — same |
| A3   | Dead branch / unreachable legacy in `build_layout_sequence` | High | Medium | P1 — removes review hazard |
| A6   | cfg globals vs TLS `OPT_*` threading model | High | High | P1 — explicit context / contract |
| A1   | `_implementation` monolith | High | High | P2 — phased extraction |
| A2   | `layout_sequence` layer merge | High | High | P2 — inject DTOs |
| A4   | ILP ↔ `price_db` boundary | High | Medium | P2 — pricing adapter |
| A5   | Orchestrator ↔ `_implementation` cycle | High | Medium | P2 — dependency injection |
| A7   | DIP / concrete cfg coupling | High | High | P2 |
| Q1–Q4| Size, DRY, print, silent except | High | Medium | P2 — alongside refactors |

---

## Next Steps

1. **Immediate (operational)**: Confirm whether `debug_logs/` and fixed NDJSON paths are `.gitignore`d and not shipped; rotate or delete accumulated logs on shared machines; align team on “no sensitive writes without explicit debug flag.”
2. **This sprint**: Gate **`price_db`** and **`layout_sequence` / `_implementation`** file appenders behind a single env/feature flag (extend `OPT_DEBUG_LOG` or equivalent); fix **dead branch** in `build_layout_sequence` for maintainability.
3. **Next sprint**: Prototype a **`RequestContext` / `OptimizationContext`** carrying order + plans explicitly into `layout_sequence`, reducing **cfg + `OPT_*`** implicit pairing; narrow **`get_price`** integration behind an interface for ILP.
4. **Backlog**: Split **`config_and_data`**, decompose **`_implementation`**, replace prints with logging, tighten exceptions and types; add targeted tests for plan/by_load combinations.

Use `/refactor` for structural/decoupling work; `/implement` for security gates and behavioral fixes.

---

## Architecture Findings

### Critical

- None. (No verified import-time cycles or hard failures in this slice; `core.optimization` ↔ `orchestrator` relies on lazy/bottom imports but loads in practice.)

### High

- **[A1]** `core/optimization/_implementation.py` is an optimization **monolith** (~1.9k+ lines): PuLP/ILP wiring, 1D/2D paths, coverage, dispatch, and ad-hoc cost heuristics (e.g. `plate_price = 12000`, bundled cut costs ~1838–1846) live in one module — **SRP** is heavily violated and regressions are hard to isolate.

- **[A2]** `viz_modules/layout_sequence.py` (~2k+ lines) **merges layers**: reinforcement lookups against `pb.db` (`Path(__file__).parent.parent / "pb.db"`, `get_reinforcement` inside `build_layout_sequence`), **process-wide** `import core.config_and_data as cfg` (e.g. `cfg.PLATE_LOAD_DETAILS`, `cfg.PLATES_*`, `cfg.normalize_load_code`), and **thread-local** optimizer state (`from core.optimization import OPT_PLAN`, `OPT_WIDTH_PRIORITY` at module scope; `OPT_CASCADING_PLAN*` imported inside `build_layout_sequence`). **Presentation/visualization** is tightly coupled to **persistence** and **implicit global/TLS state** instead of a single injected “order + plan DTO” boundary.

- **[A3]** **Unreachable control flow** in `build_layout_sequence`: after `if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('primary_cuts'):` the code calls `_build_sequence_from_plan`, then **`return`** — the following second `if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('primary_cuts'):` block is dead. Large **legacy** branches below (~450–~1016) are never executed, which is a major maintainability and review hazard.

- **[A4]** **`core/optimization/ilp_model.py`** imports **`get_price` from `core.price_db`** (`from core.price_db import get_price`, ~14). The **cutting ILP** layer depends on **commercial pricing infrastructure**, blurring the boundary between geometry/assignment math and price data.

- **[A5]** **Circular / fragile package wiring**: `core/optimization/orchestrator.py` uses a **lazy** `import core.optimization as pkg` to call implementation entrypoints (~39–44), while `_implementation.py` imports **`optimize_with_cascading_longitudinal_cuts` from `orchestrator` at EOF** (~1851–1852). Dependency direction is **bidirectional**; refactors or import reordering can reintroduce **real** cycles.

- **[A6]** **Two incompatible “global” models for one pipeline** — **`core.config_and_data`**: module-level **`PLATES_*`**, `PLATE_LOAD_DETAILS`, etc., plus **`PlateOrder.apply_to_globals()`** writing into that shared namespace (~413–453); versus **`OPT_PLAN` / `OPT_CASCADING_PLAN` / `OPT_CASCADING_PLAN_BY_LOAD`** backed by **`threading.local`** in `core/optimization/context.py` and **`_OptimizationModule.__setattr__`** in `core/optimization/__init__.py`. **`layout_sequence`** reads **both**: reinforcement map from **`cfg`** (~212–217) and plans from **`OPT_*`** (~278+, ~444+). **Coupling risk:** `OPT_*` is **per-thread**; **`cfg`** is **process-wide**. If optimization runs on **thread A** and visualization on **thread B**, or two jobs interleave on one thread without resetting TLS, **`OPT_*` can be empty/stale while `cfg` reflects another order** — asymmetric, hard-to-debug wrong output.

- **[A7]** **`core/optimization` depends concretely on `core.config_and_data`**: e.g. `_implementation.py` and `geometry.py` / `ilp_model.py` import **`cfg`**; **`PlateOrder`** / globals are the **implicit input contract** for much of the stack. **DIP** is weak: the optimizer is not driven by a narrow, injectable order/plan abstraction.

### Medium

- **[A8]** **`core/optimization/__init__.py`** replaces the module type via **`sys.modules[__name__].__class__ = _OptimizationModule`** to intercept **`OPT_*` assignments** — powerful but **non-idiomatic**; static analysis and “where does state live?” mental models suffer compared to explicit **`contextvars`** or a small **context object** passed through calls.

- **[A9]** **`core/optimization/context.py`**: TLS **proxies** improve isolation vs process-wide **`OPT_*`**, but state remains **implicit** (callers still think in terms of module attributes). **Async/interleaved work on one thread** without a dedicated worker or reset remains easy to get wrong vs **explicit parameters**.

- **[A10]** **`viz_modules/layout_sequence.py`**: **`_build_sequence_from_plan`** vs the (dead) **legacy** body under **[A3]** duplicates narrative (grouping, separators, secondary attach). **Divergence** if someone “fixes” the wrong branch.

- **[A11]** **Ad-hoc debug/trace I/O** in `layout_sequence` and parts of optimization (`_agent_seq_debug`, fixed log paths, `[VISUAL]` **`print`**) is **mixed with domain logic** — no single logging/feature-flag boundary.

- **[A12]** **Scalability**: heavy **synchronous** PuLP work in `_implementation.py` / `ilp_model.py` with **no** clear **queue/worker** contract in these modules — mostly **vertical** scaling unless callers always offload.

- **[A13]** **`core/config_and_data.py` as a god module** (~1.2k+ lines): **constants/paths**, **mutable global inventories**, **`PlateOrder` + JSON round-trip**, **text parsing** (`set_plate_lists_from_text` / `add_items`), **`make_plate_name` / load-code helpers**, **metadata buffers**, optional **NDJSON** in helpers — **many reasons to change**, weak **modular boundaries**.

- **[A14]** **`PlateOrder` vs legacy globals**: the dataclass is positioned as isolation (~211–216), but **`apply_to_globals` / `get_current_plate_order`** reintroduce a **single mutable `cfg` truth**. Callers can update **`PlateOrder`** without **`apply_to_globals`**, or mutate **globals** only — **two parallel models** of “current order.”

- **[A15]** **Hidden cross-layer effects**: **`apply_to_globals`** and **`set_plate_lists_from_text`** invoke **`core.kp_db.fill_plate_nomenclature_cache`** — **data/config loading** drags in **DB/cache population** without an explicit application/service seam.

- **[A16]** **`layout_sequence`** **mutates** plan structures in place (e.g. attaching **`reinforcement`** on cut dicts). Reusing the **same** plan dict elsewhere can **surprise** callers.

### Low

- **[A17]** **`core/optimization/orchestrator.py`** uses **`print`** for mode selection (~42–50) instead of structured logging — inconsistent with modules that use **`logging`**.

- **[A18]** **`core/optimization/ffd_packing.py`** is intentionally **stdlib-only** and isolated (~1–5), while **`geometry.py`** still pulls **`core.config_and_data`** globally (~7) — **uneven** layering discipline.

- **[A19]** **`layout_sequence` import style**: **`OPT_PLAN` / `OPT_WIDTH_PRIORITY`** at **module** import time vs **`OPT_CASCADING_*`** imported **inside** `build_layout_sequence` — minor inconsistency in when **`core.optimization`** is fully resolved.

---

**PlateOrder / globals / `OPT_*` / `layout_sequence` (summary):** `PlateOrder` can sync into **`cfg`** via **`apply_to_globals()`**; **`layout_sequence`** assumes **`cfg` (process-global)** and **`OPT_*` (thread-local)** describe **one** logical job. **`build_layout_sequence`** does not accept an explicit **`PlateOrder`** or plan dict; it **reads globals/TLS**, which maximizes **coupling** and **threading footguns** relative to a passed-in **OrderContext + OptimizationResult**.

---

## Security Findings

Scope reviewed: `core/optimization/` (incl. `ilp_model.py`, `_implementation.py`, `debug_log.py`, `context.py`), `core/config_and_data.py`, `viz_modules/layout_sequence.py`. Cross-read: `core/price_db.py` (used by `ilp_model`), `core/debug_paths.py`.

### Critical

- None identified in this scope for classic remotely exploitable issues (no auth layer here, no user-controlled SQL/path wiring found in these call paths; `get_price` uses bound parameters).

### High

- **[S1] Sensitive commercial data written to disk on every price lookup** — `get_price()` always appends NDJSON (length, load codes, resolved price) to `debug_logs/debug-db7a51.log`, with no `OPT_DEBUG_LOG`-style gate. When the ILP objective calls `get_price` (via `ilp_model.py`), this turns optimization runs into a durable pricing audit trail on the filesystem—high risk for **sensitive data exposure** (OWASP A02) if the host is shared, backed up, or indexed into git.

```169:176:c:\Users\Роман\Desktop\Шишов\core\price_db.py
        # #region agent log
        _log_path = _DEBUG_LOG_DB7A51
        try:
            with open(_log_path, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "db7a51", "hypothesisId": "get_price", "location": "price_db.py:get_price", "message": "get_price lookup", "data": {"length_m": length_m, "load_code": load_code, "length_dm": length_dm, "load_code_for_db": load_code_for_db, "result": result}, "timestamp": int(time.time() * 1000)}, ensure_ascii=False) + "\n")
        except Exception:
            pass
```

- **[S2] Unconditional “agent” NDJSON under repo paths (layout / optimization)** — Large parts of `layout_sequence.py` and `_implementation.py` append debug NDJSON **without** going through `_dbg_open_append()` / `OPT_DEBUG_LOG`. Examples: `_agent_seq_debug` always opens `debug-7e420e.log`; some branches write `PROJECT_ROOT / "debug-ef42ae.log"` and paths under `core/` (`debug-7e420e.log`). Content includes plan counts, sequence keys, reinforcement lookups—**order/layout-adjacent business data** persisted next to source and easy to commit or leak via backups.

```21:35:c:\Users\Роман\Desktop\Шишов\viz_modules\layout_sequence.py
def _agent_seq_debug(hypothesis_id: str, message: str, data: dict) -> None:
    # #region agent log
    try:
        payload = {
            "sessionId": "7e420e",
            "hypothesisId": hypothesis_id,
            "location": "layout_sequence._build_sequence_from_plan",
            "message": message,
            "data": data,
            "timestamp": int(time.time() * 1000),
        }
        with open(_DEBUG_LOG_7E420E, "a", encoding="utf-8") as f:
            f.write(json.dumps(payload, ensure_ascii=False, default=str) + "\n")
    except Exception:
        pass
```

`config_and_data.make_plate_name` similarly appends NDJSON to debug files on some branches (no env gate), mixing **user-derived dimensions** into logs.

### Medium

- **[S3] Broken isolation between thread-local optimizer state and process-wide config** — `core.optimization.context` keeps `OPT_*` in TLS, while `PlateOrder.apply_to_globals()` mutates module-level `cfg` globals (`PLATE_LOAD_DETAILS`, lists, caches). Under **concurrent requests** (async worker pool, threaded server), one request’s globals can pair with another thread’s `OPT_PLAN` → wrong layout/pricing outputs—integrity/confidentiality failure for multi-tenant or parallel workloads (OWASP A01 adjacent).

- **[S4] User-supplied text retained in diagnostics** — `LAST_PARSE_DIAGNOSTICS` stores `raw_input` per line from `set_plate_lists_from_text()`. Any consumer that logs or exports diagnostics can leak **PII or commercially sensitive paste** from chats/files.

- **[S5] Predictable local file attack surface** — `PRICE_DB_PATH`, `PRICE_XLSX_PATH`, `CUTS_DOCX_PATH`, `pb.db` are fixed under `BASE_DIR`. Not remote exploitation by itself, but a local attacker or swapped file can influence **pricing inputs** (integrity of quotes) if file permissions are loose.

- **[S6] `get_debug_log_path()` does not normalize `filename`** — Today only fixed literals call it; if a future caller passes a crafted `filename`, `DEBUG_LOGS_DIR / filename` can escape the intended directory (`..` segments). Defense-in-depth gap for **arbitrary file write** if the API is ever exposed.

### Low

- **[S7] API footgun: `get_price(..., db_path=...)`** — Parameter accepts any path; safe with current fixed `cfg.PRICE_DB_PATH` from `ilp_model`, but risky if reused with untrusted input later (**SQLite file read/write outside intended DB**).

- **[S8] Unbounded append-only debug logs** — No rotation/size caps on many writers → **disk exhaustion** /DoS on long-lived processes or CI (availability).

- **[S9] Inconsistent debug policy** — `core/optimization/debug_log.py` gates optimizer logs on `OPT_DEBUG_LOG`, but much of `layout_sequence`, `_implementation`, and `price_db` bypass it → uneven production posture and harder secret/data-handling review.

---

## Code Quality Findings

### Critical
- *(нет пунктов в заданном скоупе, которые по чисто качественным признакам блокируют корректность так же, как уровень «баг/потеря данных»; основные риски ниже смещены в высокую/среднюю серьёзность.)*

### High
- **[Q1]** Чрезмерный размер и вложенность: `viz_modules/layout_sequence.py` — `build_layout_sequence()` (~сотни строк, ~202–1078) и `_build_sequence_from_plan()` (~1305–конец файла) совмещают ветвление режимов, отладочные регионы и бизнес-логику; это серьёзно бьёт по читаемости, ревью и регрессионным тестам (сложность ≫ 30 строк / глубокая вложенность по чеклисту workflow).
- **[Q2]** Нарушение DRY: почти дублируются блоки формирования групп целых/резовых плит, разделителей и связанных `print` (сравните фрагменты порядка «solid / cut_groups / подгруппы» около ~560–690 и ~1535–1635 в `layout_sequence.py`) — выше риск расхождения поведения при правках.
- **[Q3]** `print()` вместо логирования: массовые `[OPT_*]`, `[DEBUG]`, `[VISUAL]` в `core/optimization/_implementation.py`, `orchestrator.py`, `ilp_model.py` и `viz_modules/layout_sequence.py` при том, что `core/config_and_data.py` уже использует `logging` — нет уровней, фильтрации и единого формата для серверного/CI окружения.
- **[Q4]** Широкий `except Exception` с `pass`/`no-op` вокруг отладочной записи в файлы (много вхождений в `core/optimization/_implementation.py`, `order_dispatch.py`, `viz_modules/layout_sequence.py`, плюс прямые `open(...); except Exception: pass` в `config_and_data.make_plate_name`) — типичный анти-паттерн: любые сбои I/O или сериализации исчезают без следа (в scope не найдено голого `except:`, зато массово `except Exception`).

### Medium
- **[Q5]** Слабая строгость типов: в `layout_sequence.py` у `_choose_best_separator`, `_build_sequence_from_plan`, вложенного `plate_label` и ряда хелперов параметры/возвраты описаны не полностью; в `core/optimization/orchestrator.py` / `ilp_model.py` — `Any` для конфига/модели PuLP; это ослабляет проверки и автодополнение при рефакторинге.
- **[Q6]** Избыточно широкие перехваты там, где достаточно `(TypeError, ValueError)`: например `_canonical_load_code` в `layout_sequence.py` (~82–85), `format_reinforcement_from_load_code` в `config_and_data.py` (~865–868) — смешивают «некорректный ввод» и любые программные ошибки.
- **[Q7]** Обход проверки типов: `L, w, lc = tok[0], tok[1], tok[2]  # type: ignore[index]` в `_canonical_target_order_key_tok` (~1089–1097) — лучше явный протокол/тип токена, чем `ignore`.
- **[Q8]** Перегруженный модуль `core/config_and_data.py` (≈1200+ строк: константы, глобальное состояние, парсинг, доменные типы, форматирование) — без разбиения на подмодули сложнее сопровождать и изолировать изменения (без повторения уже отмеченной темы TLS/globals).

### Low
- **[Q9]** `core/optimization/debug_log.py`: `_dbg_open_append` при любой ошибке открытия возвращает тихий no-op (~52–55) — удобно для устойчивости, но усложняет диагностику прав доступа/путей.
- **[Q10]** Покрытие тестами: для `viz_modules.layout_sequence` и `core.optimization` есть целенаправленные тесты (см. `tests/test_layout_*.py`, `tests/test_optimization_*.py`), но из-за размера и числа веток `build_layout_sequence` полнота сценариев (все комбинации plan / by_load / legacy) остаётся неочевидной — зона плотной ручной регрессии.
- **[Q11]** Недостижимый/мертвый участок после раннего `return`: в `build_layout_sequence` после `return sequence` для ветки с `primary_cuts` (~446–448) следует второй `if` с тем же условием и блоком с `print` (~450+) — чисто качественный запах «мертвый код» (архитектурный аспект уже отмечен отдельно; здесь — сопровождаемость и шум при чтении).
