# Spec: A1 — явный контекст заказа плит

**Статус**: SPECIFY (на ревью) · PLAN ⬜ · TASKS ⬜ · IMPLEMENT ⬜
**Дата**: 2026-10-26
**Реестр**: [../develop/audits/FINDINGS.md](../develop/audits/FINDINGS.md) — A1, Critical, P0
**Прогон**: [../develop/audits/2026-10-26-full-audit.md](../develop/audits/2026-10-26-full-audit.md)
**Сужает**: трек 1 в [explicit-process-state.md](./explicit-process-state.md). Трек 2 (A2/S3, воркеры и rate limit) эта спецификация не исполняет.
**Идея-предшественник**: [../ideas/plate-runtime-guard-single-worker.md](../ideas/plate-runtime-guard-single-worker.md) — страж остаётся первым шагом. Удаление прокси входит в критерий «готово» этой спецификации, не в первую пачку.

## ASSUMPTIONS I'M MAKING

1. «Готово» для A1 — не страж в режиме warning, а отсутствие неявного заказа: lazy TLS в `get_plate_mutable_runtime()` не создаётся, PEP 562 имена `PLATES_*` удалены, живой код читает заказ только из привязанного `PlateOrderContext` или из явного аргумента.
2. A2 и S3 в эту работу не входят. Один воркер не считается исправлением A1.
3. `bot/` и `bot_archived/` не живой контур. Их не мигрируем и не удаляем в этих задачах.
4. Хеш эталонной раскладки не меняется. Меняется только способ достать заказ, не состав последовательности.
5. Новых pip/npm-зависимостей нет. Режим стража — переменная окружения и stdlib `logging`.
6. HTTP middleware и CPU-путь с переданным `plate_order_ctx` уже привязывают контекст. Их контракт не ломаем.
7. Коммиты только по просьбе. `./run+logs.sh` не останавливаем.
8. Окно `warn` → `raise` — одна поставка стража. В `warn` объект на потоке сохраняется, как сейчас, и пишется лог. `raise` по умолчанию включается сразу после зелёных пачек миграции в тестах, без календарной паузы на стенде. Запрет записывать runtime в TLS включается вместе с `raise`.

→ Поправь сейчас, иначе PLAN пойдёт с этими допущениями.

## Objective

Заказ плит не должен зависеть от того, что осталось в потоке.

Сейчас `get_plate_mutable_runtime()` берёт `ContextVar`, а если он пуст — лениво кладёт **пустой** `PlateMutableRuntime` в `threading.local` и дальше отдаёт тот же объект этому потоку (`core/plate_runtime_state.py:144-153`). Первый вызов без `bound()` видит пустой заказ. Следующий вызов на том же потоке видит то, что туда записал предыдущий неявный вызов. PEP 562 в `core/config_and_data.py:441-452` и `from_plan` при `layout_cfg is None` (`viz_modules/layout_sequence/from_plan.py:193-194`) ходят в тот же getter.

HTTP закрыт: middleware создаёт пустой `PlateOrderContext` и держит `ctx.bound()` на запрос (`app/middleware/plate_runtime_isolation.py:12-17`). CPU-пул изолирует только если вызывающий передал `plate_order_ctx` (`app/concurrency/cpu_bound.py:32-41`). Без аргумента задача уходит в `asyncio.to_thread` и при касании заказа попадает в TLS потока пула.

**Пользователь:** менеджер и завод получают раскладку своего КП. Разработчик, который забыл контекст, видит ошибку со своим файлом и строкой, а не чужие или пустые ширины.

**Успех:** два расчёта в одном процессе не читают списки плит друг друга; вызов без привязанного контекста не создаёт и не переиспользует заказ потока; A1 в реестре можно закрыть по уликам этой спецификации.

## Tech Stack

Python 3, FastAPI, `contextvars` + `threading.local` (уже есть). Тесты — pytest из корня репозитория. Новых пакетов нет.

Затрагиваемые места: `core/plate_runtime_state.py`, `core/plate_order_context.py`, `core/config_and_data.py`, `core/domain/plate_order.py` (`apply_to_globals`), `viz_modules/layout_sequence/from_plan.py`, живые вызовы `get_plate_mutable_runtime()` в `app/`, `core/`, `viz_modules/`.

## Commands

```bash
# Изоляция и контекст
pytest tests/test_plate_runtime_request_isolation.py \
  tests/test_plate_mutable_runtime_isolation.py \
  tests/test_plate_order_context.py \
  tests/test_config_and_data_module_semantics.py \
  tests/test_config_and_data_proxy_boundary.py -q

# Раскладка после каждой пачки миграции
pytest tests/test_layout_hash.py tests/test_plate_runtime_request_isolation.py -q

# Полный прогон перед закрытием A1
pytest tests/ -q
```

Dev-сервер этой спецификацией не перезапускается.

## Project Structure

```
core/plate_runtime_state.py       → страж в get_plate_mutable_runtime(); затем запрет lazy TLS
core/plate_order_context.py       → единственная публичная привязка заказа (bound / run_in_order_context)
core/config_and_data.py           → финал: нет __getattr__ на MUTABLE_LEGACY_NAMES
core/domain/plate_order.py        → финал: нет apply_to_globals
viz_modules/layout_sequence/      → from_plan не читает runtime при layout_cfg is None
viz_modules/procurement/          → items, price_rows, breakdown получают runtime аргументом или из bound()
app/middleware/plate_runtime_isolation.py → без смены контракта
app/concurrency/cpu_bound.py      → без ctx не подставляет чужой/пустой заказ
tests/                            → новые проверки стража и «поток не помнит прошлый заказ»
```

`bot/`, `bot_archived/`, `tests/archived/` не трогаем.

## Code Style

Новый код получает заказ аргументом или читает его внутри уже открытого `bound()`. Getter без привязки — не стиль, а устаревшая точка, которую закрываем.

```python
def get_plate_mutable_runtime() -> PlateMutableRuntime:
    ctx_rt = _plate_cv.get()
    if ctx_rt is not None:
        return ctx_rt
    _report_implicit_plate_runtime_access()  # warn или raise по PLATE_RUNTIME_STRICT
    raise RuntimeError("заказ плит не привязан")  # в режиме raise
```

В режиме `warn` getter ведёт себя как сейчас: пустой `ContextVar` создаёт или возвращает runtime потока и пишет warning с файлом и строкой вызывающего. Память потока в этом окне ещё жива. В режиме `raise` ветка TLS удаляется: пустой `ContextVar` даёт `RuntimeError`, атрибут `runtime` на потоке не создаётся.

Имена без аббревиатур сверх уже существующих `PlateOrderContext` и `PlateMutableRuntime`. Сообщение `RuntimeError` — по-русски, с файлом и строкой вызывающего.

## Testing Strategy

| Уровень | Что обязано стать зелёным |
|---------|---------------------------|
| Unit стража | Без `bound()` режим `raise` бросает `RuntimeError` и не пишет в TLS. Режим `warn` пишет warning с местом вызова и возвращает тот же объект потока, что и сегодня. С `bound()` страж молчит. |
| Утечка потока | Два последовательных вызова на одном потоке без `bound()` не видят списки друг друга. |
| HTTP | Существующие `test_plate_runtime_request_isolation` остаются зелёными без смены сценария. |
| CPU | `run_cpu_bound` с `plate_order_ctx` видит этот заказ. Без аргумента касание заказа в режиме `raise` падает, а не читает TLS пула. |
| Golden | `tests/test_layout_hash.py` — тот же хеш последовательности, что до пачки. |
| Финал | `pytest tests/ -q` зелёный. Поиск `MUTABLE_LEGACY_NAMES` / `__getattr__` proxy в `core/config_and_data.py` ничего не находит. Поиск `apply_to_globals` в `app/`, `core/`, `viz_modules/` ничего не находит. |

Покрытие процентом не задаём. Негативный тест на неявный доступ обязателен до удаления прокси.

## Boundaries

- **Always:** сначала страж, потом пачки вызывающих, прокси удаляем последним; хеш раскладки сверяем на каждую пачку; `pytest` из корня; минимальный diff; русское сообщение при отсутствии контекста.
- **Ask first:** удаление `bot_archived/`; правка или удаление одноразовых файлов в `scripts/`; смена демо-заказа в тестах; значение по умолчанию `PLATE_RUNTIME_STRICT=raise` на стенде до истечения окна `warn`; любая правка схемы БД.
- **Never:** новые зависимости; Redis; изменения `plita.db` / `pb.db`; смена HTTP-контрактов API; коммит без просьбы; остановка `./run+logs.sh`; правка A2/S3 «заодно»; закрытие A1, пока lazy TLS или PEP 562 proxy ещё отдают заказ.

## Design

Порядок фиксирован. Пачка N не начинается, пока пачка N−1 зелёная и хеш раскладки тот же.

1. **Страж.** Одна точка — `get_plate_mutable_runtime()`. `ContextVar` заполнен → этот runtime, без лога. Пуст → `_report_implicit_plate_runtime_access()` со stacklevel на вызывающего. `PLATE_RUNTIME_STRICT=warn|raise`, по умолчанию `warn`. В `warn` runtime потока сохраняется. В `raise` — `RuntimeError`, без записи в TLS. `plate_mutable_runtime_scope` / `PlateOrderContext.bound()` остаются явным доступом и страж не триггерят.
2. **Пачки живых вызывающих.** Каждая пачка либо принимает `PlateMutableRuntime` / срез конфига аргументом, либо вызывается только из `bound()`. Порядок: `viz_modules/procurement/` → `viz_modules/layout_sequence/from_plan.py` → остальные живые вызовы в `core/` и `app/`, которые сами делают `get_plate_mutable_runtime()` вне уже открытого `bound()`. В пачке `from_plan` аргумент `layout_cfg` становится обязательным. Тесты и хеш передают тот же пустой срез, который сегодня строит fallback из runtime; ветка `layout_cfg is None` удаляется. `run_cpu_bound` без `plate_order_ctx` не подставляет заказ сам: вызывающий расчёта передаёт контекст, иначе внутренний код падает в режиме `raise`. После зелёного `pytest` на этих пачках `PLATE_RUNTIME_STRICT` по умолчанию становится `raise`.
3. **Финал.** Удалить `__getattr__` и `MUTABLE_LEGACY_NAMES`. `apply_to_globals` удалить вместе с тестом, который его ещё вызывает (`tests/test_plate_order_context.py`). Тесты прокси-семантики переписать на явный runtime или удалить вместе с контрактом прокси. После этого getter без `ContextVar` только падает и не имеет ветки TLS.

`register_plate_metadata` / `consume_plate_metadata` / `get_exact_width` остаются функциями модуля, но читают runtime только через getter под стражем. Отдельный обход вокруг getter не добавляем.

## Success Criteria

- [ ] Два последовательных неявных чтения на одном потоке не возвращают один и тот же список плит.
- [ ] `PLATE_RUNTIME_STRICT=raise` и пустой `ContextVar` → `RuntimeError` по-русски, TLS-атрибут `runtime` не создаётся.
- [ ] `PLATE_RUNTIME_STRICT=warn` логирует файл и строку вызывающего и возвращает тот же runtime потока, что и сейчас.
- [ ] Параллельные HTTP-запросы по-прежнему изолированы существующим тестом request isolation.
- [ ] `run_cpu_bound(..., plate_order_ctx=ctx)` выполняет функцию в этом заказе.
- [ ] `_build_sequence_from_plan` требует `layout_cfg`. Вызов без среза — `TypeError`. Внутри функции нет чтения `get_plate_mutable_runtime()`.
- [ ] Хеш в `tests/test_layout_hash.py` совпадает с хешем до первой пачки.
- [ ] `cfg.PLATES_*` (имена из бывшего `MUTABLE_LEGACY_NAMES`) даёт `AttributeError`.
- [ ] В `app/`, `core/`, `viz_modules/` нет вызовов `apply_to_globals`.
- [ ] `pytest tests/ -q` зелёный.
- [ ] В FINDINGS.md строка A1 может стать `resolved` только когда все пункты выше выполнены. Сама смена статуса — отдельное действие, не часть коммита кода без просьбы.

## Out of Scope

- A2, S3: fail-fast воркеров, Redis, общий store лимитов.
- Удаление или перенос `bot/` и `bot_archived/`.
- Смена целевой функции ILP, хеша раскладки, ширины дорожки, схемы КП.
- Очередь LibreOffice (A18) и отдельный процесс на каждый расчёт.
- Данные в `plita.db` и уже сохранённые планы.

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| D-warn-tls | Поведение в `warn` | Объект заказа на потоке сохраняется, в лог пишется вызывающий. TLS отключается только в `raise` |
| D-raise-when | Когда `raise` по умолчанию | Сразу после зелёных пачек миграции в `pytest`. Семи дней на стенде нет |
| D-layout-cfg | `from_plan` без среза | В пачке миграции `layout_cfg` обязателен. Тесты передают тот же пустой срез, что сегодня даёт fallback. Хеш последовательности не меняется |
| D-scripts | `scripts/` | Отдельной миграции нет. `pool_overrun_experiment_2026.py` и `build_kz_pool_pair_report_2026.py` уже создают `PlateOrderContext`. Прямых вызовов `get_plate_mutable_runtime()` в `scripts/` нет. Если `raise` поймает скрипт, оборачиваем этот файл |

## Open Questions

Нет. Решения, которые блокировали план, зафиксированы выше.
