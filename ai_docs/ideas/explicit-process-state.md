# Явное состояние процесса (A1 + A2/S3)

> Дата: 2026-09-21. Статус: к реализации закрыта.
> Спека: `ai_docs/specs/explicit-process-state.md`.
> Источник: реестр находок `ai_docs/develop/audits/FINDINGS.md` —
> A1 (Critical, P0), A2 (Critical, P0), S3 (High, P1); прогон 2 — `2026-09-21-full-audit.md`.

## Problem Statement

How might we сделать так, чтобы состояние процесса (заказ плит, счётчики лимитов)
всегда совпадало с реальной топологией исполнения и деплоя — by construction,
а не по договорённости?

## Recommended Direction

Два трека под одним принципом: **«состояние процесса должно совпадать с топологией
деплоя; неявное — враг»**.

### Трек 1 (A1, корректность): довести strangler-миграцию PlateOrderContext до конца

Инфраструктура уже есть: `PlateMutableRuntime` (ContextVar + TLS), `PlateOrderContext`
с `bound()` / `hydrate_from_order()`, изоляция HTTP-middleware и CPU-пула. Осталось:

- Шаг 0 — guard: неявный `get_plate_mutable_runtime()` без bound-контекста сначала
  логирует (warning-режим, инвентаризация callers), затем падает. Guard живёт в одной
  точке — getter'е рантайма; PEP 562 proxy идёт через него же и покрывается автоматически.
- Шаги 1..N — миграция live-файлов пачками: `viz_modules/procurement/*` →
  `layout_sequence` → `core/parsing` → `core/optimization` → `core/visualization` и пр.
  Каждая пачка: зелёный suite + hash-сверка sequence раскладки (приём Q3).
- Финал — отрезать `bot_archived`, удалить `MUTABLE_LEGACY_NAMES` и `__getattr__`-proxy,
  переписать тесты с proxy-семантики на явный runtime.

### Трек 2 (A2/S3, безопасность/деплой): жёсткий single-worker везде

- `enforce_single_instance_workers` — на все окружения и любой storage layout:
  declared workers > 1 без shared store → процесс не стартует.
- `UVICORN_WORKERS=1` обязателен и объявлен: undeclared → fail в production,
  одна итерация warning → затем fail везде.
- deploy-contract и docker-compose синхронизированы (compose уже `UVICORN_WORKERS=1`).
- Redis не трогаем: `RATE_LIMIT_SHARED_STORE=redis` остаётся NotImplementedError-заглушкой P7.

## Key Assumptions to Validate

- [ ] Golden-master/hash тесты раскладки покрывают мигрируемые пути — иначе сначала тесты
- [ ] Guard в warning-режиме не находит живых callers без bound-контекста (CLI, scripts/)
- [ ] `bot_archived` не запущен ни на одном сервере (проверить systemd/cron перед удалением)
- [ ] Прод-запуск реально имеет `UVICORN_WORKERS=1` в env (проверить .env / unit-файл
      до включения undeclared-fail)

## Locked decisions

| Тема | Решение |
|------|---------|
| Масштабирование | Одного воркера достаточно; Redis отложен до реального multi-host. Триггеры пересмотра: второй хост за балансировщиком; Redis, появившийся по другой причине (очередь OCR/LibreOffice, A18) |
| Глубина A1 | До конца: все live call sites на явный контекст, proxy удалить. Guard — первый шаг, не альтернатива |
| Формат | Два трека под одним принципом; Трек 2 первым (~1 день, закрывает A2+S3) |
| bot_archived | Отрезать вместе с proxy-именами (мёртвый контур, A10/S14) |

## MVP Scope

**Входит**

- Трек 2 целиком: fail-fast во всех окружениях, обязательный объявленный
  `UVICORN_WORKERS`, deploy-contract, расширенный `test_rate_limit_deployment`
- Трек 1, шаг 0: guard в warning-режиме + инвентаризация срабатываний

**Не входит** — см. Not Doing.

## Not Doing (and Why)

- **Redis / shared counters** — нет multi-host топологии; заглушка P7 уже стоит
- **SQLite-счётчики для multi-worker на одном хосте** — нет сценария, требующего >1 воркера
- **Миграция bot_archived** — мёртвый контур (A10/S14), отрезаем вместе с proxy-именами
- **Per-device revocation сессий** — отдельная future work (`session.py:20`)

## Open Questions

- Окно guard warn → raise: предложение — одна итерация, критерий «ноль срабатываний за N дней»
- Кто смотрит `/health` rate_limiting-метаданные — заметит ли кто-то warning?
- `scripts/_cleanup_stage5_a4.py` и подобные — одноразовые? Удалить вместе с миграцией?
