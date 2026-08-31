# Telegram-бот — DEPRECATED

> **Статус:** заморожен с 2026-06-19 (спринт стабилизации P0).  
> **Не использовать в production.** Активная разработка и поддержка прекращены.

## Что это значит

- Новые фичи в бот **не добавляются**.
- Бот **не является** целью консолидации бизнес-логики — единый pipeline планирования живёт в `core/production/planning.py` и используется **веб-приложением** (FastAPI + React).
- Код бота **сохранён** в репозитории для совместимости и локальной отладки; полное удаление — отдельное согласование (follow-up).
- При запуске (`python run_bot.py`) в лог пишется предупреждение `WARNING`.

## Что использовать вместо бота

| Задача | Интерфейс |
|--------|-----------|
| Коммерческие предложения, КП | Веб UI (`frontend/`) + API `/api/v1/commercial/` |
| Производственное планирование | Веб UI + API `/api/v1/production/` |
| Планы производства | SQLite `plita.db`, таблица `production_plans` (через `PlanRepository`) |

## Запуск (только dev, на свой риск)

```powershell
Set-Location "c:\Users\Роман\Desktop\Шишов"
.\.venv\Scripts\Activate.ps1
python run_bot.py
```

Требуется `bot/bot.env` с валидным `BOT_TOKEN`. См. комментарии в `run_bot.py`.

## Связанные документы

- Спека P0: [`ai_docs/specs/stabilizaciya-p0-audit-2026-06-19.md`](../ai_docs/specs/stabilizaciya-p0-audit-2026-06-19.md)
- Аудит: [`ai_docs/develop/audits/2026-06-19-full-project-audit.md`](../ai_docs/develop/audits/2026-06-19-full-project-audit.md)
