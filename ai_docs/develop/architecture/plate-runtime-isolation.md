# Plate runtime isolation (A3 / PLATE-CTX-001)

## Модель

Мутабельное состояние заказа плит (`PLATES_*`, `PLATE_LOAD_DETAILS`, OPT-снимок) живёт в:

- `PlateMutableRuntime` — списки и карты заказа
- `PlateOrderContext` — SSOT на один HTTP-запрос / апдейт бота (`plates` + `optimization`)
- `config_and_data` PEP 562 proxy — legacy-доступ `cfg.PLATES_*` → текущий рантайм (strangler, не удалён)

Изоляция между конкурентными запросами:

1. **FastAPI:** `PlateMutableRuntimeIsolationMiddleware` создаёт `PlateOrderContext.fresh_empty()` и вызывает `ctx.bound()` на время запроса.
2. **Telegram bot:** аналогичный middleware на каждый апдейт.
3. **Hot paths** (commercial preview, production `build_plan`, bot planning) передают request-scoped `plate_order_ctx` в pipeline и вызывают `ctx.bound()` / `run_in_order_context` перед legacy-кодом.

`bound()` вкладывает `contextvars.ContextVar` (asyncio-задачи) поверх `threading.local` (sync / worker threads).

## Deployment constraint

**Текущая гарантия изоляции — в рамках одного процесса Python.**

| Сценарий | Безопасно? | Примечание |
|----------|------------|------------|
| Один uvicorn worker, async concurrent requests | Да | ContextVar + middleware |
| Несколько asyncio-задач в одном worker | Да | Тесты `test_parallel_*` |
| Несколько gunicorn/uvicorn workers на одной машине | Да* | Каждый worker — отдельный процесс; глобали не шарятся между процессами |
| Фоновые потоки без `ctx.bound()` | Нет | Использовать `run_in_order_context` или явный scope |
| Скрипты/CLI без middleware | Нет | Обернуть в `PlateOrderContext.fresh_empty().bound()` |

\* Межпроцессной «TTL cache» для plate state **нет** — это осознанное ограничение WP3 (partial). Полный decommission `config_and_data` proxy отложен.

**Рекомендация для production:** один инстанс приложения **или** несколько workers с обязательным middleware на всех entry points (уже включено). Не полагаться на module-level globals вне `PlateOrderContext.bound()`.

## Что осталось (post-WP3)

- PEP 562 proxy в `core/config_and_data.py` — incremental migration
- Прямые вызовы `PlateOrderContext.fresh_empty()` внутри core без внешнего ctx — fallback для offline/tests
