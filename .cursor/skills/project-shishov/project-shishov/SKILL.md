---
name: project-shishov
description: Контекст проекта «Шишов» — FastAPI backend, React frontend, Telegram-бот, оптимизация раскладки ЖБ плит, КП и производственное планирование. Использовать при любой задаче в этом репозитории до выбора остальных скиллов.
---

# Проект «Шишов»

## Обзор

Система автоматизации для завода ЖБ изделий: коммерческие предложения (КП), оптимизация раскладки плит, производственное планирование, закупки, Telegram-бот для менеджеров.

## Структура репозитория

| Путь | Назначение |
|------|------------|
| `app/` | FastAPI: `api/v1/endpoints/`, `services/`, `repositories/`, `schemas/`, `core/settings.py` |
| `core/` | Доменная логика: оптимизация (`core/optimization/`), парсинг, КП, OCR, конфиг |
| `frontend/` | React + Vite + TypeScript: страницы КП, производство, архив |
| `bot/` | Telegram-бот: handlers, middleware, security |
| `viz_modules/` | Визуализация раскладки, закупки |
| `tests/` | pytest — запускать из корня: `pytest tests/` |
| `docker/` | Docker, seed БД |
| `закупки/` | Прайсы поставщиков (не код) |

## Стек и соглашения

- **Backend:** Python 3, FastAPI, Pydantic v2, SQLite (`plita.db`, `pb.db`)
- **Frontend:** React, Vite, TypeScript
- **Тесты:** pytest в `tests/`
- **venv:** `.venv/` в корне — активировать перед запуском
- **Слои:** роутер → сервис → репозиторий; не смешивать ORM, схемы и бизнес-логику
- **Секреты:** `bot.env`, `*.env` — не коммитить
- **Документация агента:** `ai_docs/` (локально, в `.gitignore`)

## Типичные команды

```bash
# Backend
uvicorn app.main:app --reload

# Тесты
pytest tests/ -q

# Бот
python run_bot.py

# Frontend (из frontend/)
npm run dev
```

## Сопоставление задач и скиллов

| Задача | Скилл |
|--------|-------|
| Новый API endpoint | `api-and-interface-design` → `incremental-implementation` |
| React UI | `frontend-ui-engineering` |
| Баг / traceback | `debugging-and-error-recovery` |
| Рефакторинг оптимизации | `doubt-driven-development` + `test-driven-development` |
| Код-ревью перед merge | `code-review-and-quality` |
| Деплой Docker | `shipping-and-launch` |

## Ограничения проекта

- Минимальный diff — не трогать несвязанный код
- Не коммитить без явной просьбы пользователя
- `.cursor/` и `ai_docs/` в `.gitignore` — локальные настройки агента
- Windows: пути с кириллицей; в PowerShell использовать `Set-Location`, не `&&`

## Верификация

После изменений в backend: `pytest tests/` (релевантные тесты минимум).
После API: проверить схемы в `app/schemas/` и endpoint в `app/api/v1/`.
После frontend: `npm run build` из `frontend/` при существенных изменениях.
