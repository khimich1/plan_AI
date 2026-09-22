# Spec: Явное состояние процесса (A1 + A2/S3)

**Статус**: IDEATE ✅ · SPECIFY ✅ (на ревью) · PLAN ⬜ · IMPLEMENT ⬜
**Дата**: 2026-09-21
**One-pager**: [../ideas/explicit-process-state.md](../ideas/explicit-process-state.md)
**Источник**: [../develop/audits/FINDINGS.md](../develop/audits/FINDINGS.md) — A1 (Critical, P0), A2 (Critical, P0), S3 (High, P1); прогон 2 — [../develop/audits/2026-09-21-full-audit.md](../develop/audits/2026-09-21-full-audit.md)

## Objective

**Проблема.** Неявное состояние процесса не совпадает с топологией исполнения и деплоя:

- **A1** — заказ плит доступен через «магические» глобалы (`core/config_and_data.py` → PEP 562
  `__getattr__` → `get_plate_mutable_runtime()`). HTTP и CPU-пул изолированы, но
  viz/procurement/core-модули читают рантайм неявно. Цена ошибки: раскладка или план уезжают
  на завод **по плитам чужого КП**.
- **A2/S3** — лимиты логина, смены пароля и OCR считаются в памяти одного процесса
  (`_SlidingWindowRateLimiter`, `_CommercialOcrUploadLimiter`). Два воркера = лимит ×2.
  Fail-fast есть только для production + `single_instance` + явно объявленных workers.

**Цель.** Трек 2: неверная топология воркеров = отказ старта в любом окружении.
Трек 1: заказ плит передаётся явно; без привязанного контекста код видит пустой заказ
или падает — не чужой заказ никогда.

**Пользователь:** ops и разработчики (напрямую); менеджеры и завод (косвенно — корректность КП).

**Успех:** A1, A2, S3 → `resolved` в FINDINGS.md; следующий прогон аудита их не находит;
мисконфиг воркеров не стартует; неявный доступ к заказу логируется, затем падает.

---

## ASSUMPTIONS I'M MAKING

1. **Single-worker достаточно** на горизонте 6–12 мес (подтверждено 2026-09-21).
2. **`bot_archived` мёртв**; отрезаем вместе с proxy-именами (подтверждено).
3. **Прод env содержит `UVICORN_WORKERS=1`** — проверяется до включения undeclared-fail;
   если нет — сначала выставить, потом ужесточать.
4. **Hash-эталон раскладки воспроизводим** (прецедент Q3, hash sequence `77c686bd…`).
5. **Новых pip/npm зависимостей нет**; guard — stdlib logging + env.
6. **`shared_volume` перестаёт быть путём к репликам**, пока лимиты in-process —
   storage layout не смягчает fail-fast.
7. **Коммиты — по просьбе.** `./run+logs.sh` не убивать.
8. **Окно guard warn → raise — одна итерация**; критерий переключения «ноль срабатываний
   за N дней» фиксируется в плане Трека 1.

→ Correct me now or these are locked for PLAN.

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| **D-principle** | Формат | Два трека под одним принципом: «состояние процесса = топология деплоя; неявное — враг» |
| **D-single** | declared workers > 1 без shared store | Fail-fast во **всех** окружениях и при любом storage layout |
| **D-undeclared** | workers не объявлен | production → fail; non-production → warning одну итерацию → fail |
| **D-redis** | Redis / shared counters | Не трогаем; `RATE_LIMIT_SHARED_STORE=redis` остаётся NotImplementedError (P7) |
| **D-guard** | Точка guard | Одна: `get_plate_mutable_runtime()`; lazy-создание TLS-рантайма = неявный доступ. Proxy покрыт автоматически |
| **D-guard-mode** | Режимы guard | `PLATE_RUNTIME_STRICT=off\|warn\|raise`, default `warn` на первую итерацию |
| **D-order** | Порядок пачек миграции | procurement → layout_sequence → parsing → optimization → visualization/прочее → финал (proxy + bot_archived) |
| **D-hash** | Проверка пачки | Hash sequence эталонной раскладки неизменен + `pytest -q` зелёный |
| **D-bot** | `bot_archived/` | Удалить (после проверки systemd/cron), не мигрировать |
| **D-proxy** | PEP 562 | Удалить `__getattr__` и `MUTABLE_LEGACY_NAMES`; тесты proxy-семантики → явный runtime |

---

## User Stories

- Как **ops**, запускаю с `--workers 4` — процесс падает при старте с сообщением
  «Set UVICORN_WORKERS=1», а не работает с половинной защитой.
- Как **ops**, запускаю production без `UVICORN_WORKERS` — процесс падает при старте
  и требует объявить топологию явно.
- Как **разработчик**, пишу новый viz-модуль и зову рантайм без контекста — получаю
  warning в логе (после окна — ошибку) с моим файлом и строкой, а не чужой заказ.
- Как **менеджер**, не замечаю ничего: те же КП и раскладки; мой заказ не может
  попасть в чужой документ.
- Как **аудитор**, следующим прогоном не нахожу A1/A2/S3.

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python, FastAPI, uvicorn; только stdlib для новых проверок |
| Трек 2 | `app/security/login_rate_limit.py`, `app/main.py` (lifespan), health-метаданные |
| Трек 1 | `core/plate_runtime_state.py` (guard), `core/plate_order_context.py`, `core/config_and_data.py` (proxy), ~14 live-файлов viz/core/app |
| Тесты | pytest `tests/` |
| Доки | `ai_docs/develop/deploy-contract.md`, `docker-compose.yml`, `docker-compose.split.yml` |

Новых пакетов нет.

## Commands

```
# Трек 2 (focused)
pytest tests/test_rate_limit_deployment.py tests/test_auth_login_rate_limit.py \
  tests/test_web_login_rate_limit.py tests/test_password_change_rate_limit.py \
  tests/test_client_ip_resolution.py -q

# Трек 1 (focused)
pytest tests/test_plate_runtime_request_isolation.py \
  tests/test_plate_mutable_runtime_isolation.py tests/test_plate_order_context.py \
  tests/test_config_and_data_module_semantics.py \
  tests/test_config_and_data_proxy_boundary.py -q

# На каждую пачку миграции
pytest -q

# Dev: не убивать ./run+logs.sh
```

## Project Structure

```
# Трек 2
app/security/login_rate_limit.py     → enforce на все окружения + undeclared-fail в prod
app/main.py                          → lifespan: вызов уже есть (:46–50), сигнатура/логика
app/schemas/health.py                → метаданные rate_limiting (уже есть; сверить)
ai_docs/develop/deploy-contract.md   → контракт: обязательный UVICORN_WORKERS=1
docker-compose.yml                   → UVICORN_WORKERS=1 уже задан (:20); сверить split
tests/test_rate_limit_deployment.py  → матрица env × declared

# Трек 1
core/plate_runtime_state.py          → guard в get_plate_mutable_runtime() (warn/raise)
core/config_and_data.py              → финал: удалить __getattr__ (:423–434)
viz_modules/procurement/             → items, price_rows, breakdown, load_context (пачка 1)
viz_modules/layout_sequence/         → from_plan.py:194 default layout_cfg (пачка 2)
core/parsing/plate_lists.py          → пачка 3
core/optimization/                   → layout_runtime_snapshot, legacy_width_plan (пачка 4)
core/visualization/__init__.py       → пачка 5 (+ plates_preview_xlsx, kp_db_nomenclature)
core/commercial_offer_layout.py      → пачка 6 (+ app/services/product_draft_config.py)
bot_archived/                        → финал: удалить целиком
tests/                               → proxy-семантика → явный runtime
```

## Code Style

Явный параметр контекста вместо ambient-доступа. На hot paths допустим
`ctx.bound()` + getter — это механизм привязки, не публичный стиль нового кода.
Русские сообщения ошибок старта и 429. Минимальный diff, без новых зависимостей.

Guard — одна точка, режим по env:

```python
def get_plate_mutable_runtime() -> PlateMutableRuntime:
    ctx_rt = _plate_cv.get()
    if ctx_rt is not None:
        return ctx_rt
    if not hasattr(_tls, "runtime"):
        _report_implicit_plate_runtime_access()  # warn/raise по PLATE_RUNTIME_STRICT
        _tls.runtime = new_plate_mutable_runtime_empty()
    return _tls.runtime
```

## Design

### Трек 2: матрица fail-fast

```
                        declared = 1   declared > 1   undeclared
production              ok             fail           fail (новое)
non-production          ok             fail (новое)   warning → fail (через итерацию)
```

`enforce_single_instance_workers` теряет условия `app_env == production` и
`storage_layout == single_instance` в ветке declared>1: in-process лимиты сломаны
при любом layout. Отдельная ветка undeclared: production → `RuntimeError`;
non-production — текущий warning из `warn_if_multi_worker_without_shared_store`
на одну итерацию, затем fail. Сообщение об ошибке содержит инструкцию
(`Set UVICORN_WORKERS=1`). `validate_rate_limit_shared_store_config` не меняется.
Следствие: `shared_volume` больше не открывает несколько реплик — путь к репликам
только через Redis-триггеры (см. Not Doing).

### Трек 1: guard и миграция

- Guard в одной точке — `get_plate_mutable_runtime()`. Привязка через
  `plate_mutable_runtime_scope` / `PlateOrderContext.bound()` = явный доступ;
  lazy-создание TLS-рантайма = неявный. PEP 562 proxy идёт через тот же getter —
  покрыт автоматически.
- `warn`: `logger.warning` со stacklevel → файл:строка caller'а в логе; поведение
  не меняется. `raise`: `RuntimeError` с местом вызова.
- Окно warn — одна итерация; инвентаризация срабатываний (CLI, scripts/, редкие
  viz-пути) — вход для пачек. Критерий перехода на raise: ноль срабатываний за N дней.
- Порядок пачек — D-order. Каждая пачка: hash sequence эталонной раскладки до/после
  неизменен (приём Q3) + `pytest -q` зелёный. Отдельный PR на пачку, откат независим.
- Финал: удалить `bot_archived/`, `MUTABLE_LEGACY_NAMES`, `__getattr__`-proxy;
  тесты `test_config_and_data_module_semantics` / `test_config_and_data_proxy_boundary`
  переписать на явный runtime или удалить вместе с контрактом proxy.

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit Т2 | Матрица enforce: declared>1 fail во всех env; undeclared prod fail; undeclared dev warning |
| Startup Т2 | Lifespan падает с `RuntimeError` при declared>1; сообщение по-русски с инструкцией |
| Unit Т1 guard | `warn` логирует и пропускает; `raise` падает; bound scope не срабатывает; proxy покрыт |
| Isolation Т1 | `test_plate_runtime_request_isolation` / `test_plate_mutable_runtime_isolation` / `test_plate_order_context` зелёные без правок |
| Golden Т1 | Hash sequence эталонной раскладки неизменен на каждую пачку |
| Regress Т1 | Полный `pytest -q` после каждой пачки; после финала grep не находит `PLATES_` в live-коде |

## Boundaries

- **Always:** guard сначала в `warn`; hash-сверка sequence на каждую пачку; зелёный
  `pytest -q` перед merge пачки; русские сообщения; минимальный diff; отдельный PR на пачку.
- **Ask first:** удаление `bot_archived/` (после проверки systemd/cron); удаление
  `scripts/*`; новые env-переменные сверх `PLATE_RUNTIME_STRICT`; смена demo-order
  семантики в тестах; включение undeclared-fail без проверки прод-env.
- **Never:** новые зависимости (Redis и др.); миграции или изменения БД и файлов данных;
  смена HTTP-контрактов API; коммит без просьбы; убивать `./run+logs.sh`.

## Success Criteria

| # | Критерий |
|---|----------|
| S1 | declared workers > 1 без shared store → отказ старта во всех окружениях, сообщение с инструкцией |
| S2 | undeclared workers в production → отказ старта; в non-production — warning (одна итерация), затем fail |
| S3 | guard `warn`: неявный доступ логируется с местом вызова, поведение не меняется |
| S4 | guard `raise`: неявный доступ падает; HTTP/CPU пути зелёные (контекст привязан) |
| S5 | Каждая пачка миграции: hash sequence неизменен + `pytest -q` зелёный |
| S6 | Proxy и `MUTABLE_LEGACY_NAMES` удалены; `cfg.PLATES_*` → AttributeError; live-callers нет |
| S7 | `bot_archived/` удалён, suite зелёный |
| S8 | A1, A2, S3 → `resolved` в FINDINGS.md; прогон аудита их не находит |

## Out of Scope

- Redis / shared counters (триггеры пересмотра: второй хост; Redis по другой причине — A18)
- SQLite-счётчики для multi-worker на одном хосте
- Миграция `bot_archived` (отрезаем)
- Per-device revocation сессий (`session.py:20`, future work)
- Любые изменения данных, черновиков, планов, схем БД

## Open Questions

- Окно guard warn → raise: длина и критерий «ноль срабатываний за N дней» — зафиксировать в плане Трека 1
- Кто смотрит `/health` rate_limiting-метаданные — заметит ли кто-то warning?
- `scripts/_cleanup_stage5_a4.py` и подобные — одноразовые? Удалить вместе с миграцией?
