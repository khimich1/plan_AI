# Отчёт о завершении оркестрации: Offers API и слой планирования

**Дата:** 2026-05-06  
**Тип:** completion report (оркестрация после аудита)  
**Контекст:** замечания A1/A2 из полного аудита проекта (расхождение графа API и циклические зависимости `app` ↔ `bot` через `plan_manager`).

---

## Цели

| ID | Проблема | Ожидаемый результат |
|----|-----------|---------------------|
| **A1** | Эндпоинты `/offers` описаны в `app/api/v1/endpoints/offers.py`, но не подключены в `app/api/v1/router.py`; веб-интерфейс менеджеров вызывает `/api/v1/offers/*`, которых в рантайме не было. | Роутер `offers` включён в API v1; маршруты доступны под префиксом приложения. |
| **A2** | Модули `app` импортировали `bot.handlers.plan_manager`, бот импортировал `app.services` / `app.domain`, что давало циклы и смешение слоёв. | Реализация хранения и операций над планами перенесена в `app/planning/`; бот использует shim-реэкспорт; в `app/` нет зависимостей от `bot.handlers.plan_manager`. |

---

## Выполненные изменения

### A1 — регистрация роутера offers

- Файл: `app/api/v1/router.py`
- Добавлены импорт модуля `offers` и вызов `router.include_router(offers.router)` (порядок согласован с остальными роутерами v1).

### A2 — перенос `plan_manager` под `app`

- Добавлен пакет `app/planning/`:
  - `app/planning/plan_manager.py` — полная прежняя реализация (пути к `bot/data/plans`, работа с метаданными, дорожки по дням и т.д.); корень репозитория и `BOT_DIR` пересчитаны относительно нового расположения файла.
  - `app/planning/__init__.py` — экспорт `plan_manager`.
- Файл `bot/handlers/plan_manager.py` заменён на thin shim: `from app.planning.plan_manager import *`, чтобы сохранить относительные импорты внутри `bot/handlers/*`.
- Все импорты в слое `app/` и затронутых тестах/скриптах переведены на `app.planning` (или точечный импорт из `app.planning.plan_manager`).
- Обновлена ссылка в докстринге: `core/plan_commit.py` (упоминание `save_plan`).

---

## Затронутые файлы (инвентарь)

| Путь | Роль изменения |
|------|----------------|
| `app/api/v1/router.py` | A1 |
| `app/planning/__init__.py` | новый |
| `app/planning/plan_manager.py` | новый (перенос логики) |
| `bot/handlers/plan_manager.py` | shim |
| `app/repositories/plan_repository.py` | импорты |
| `app/services/archive_service.py` | локальный импорт |
| `app/services/production_service.py` | импорты |
| `app/services/production_planning_service.py` | импорты |
| `app/services/day_view_service.py` | импорты |
| `app/services/day_documents_service.py` | импорты |
| `scripts/recover_stuck_plan_plates.py` | импорты |
| `core/plan_commit.py` | документация |
| `tests/test_*.py` (planning / plan_manager) | импорты и logger в caplog где нужно |

---

## Проверки

- **Импорты:** в каталоге `app/` отсутствуют обращения к `bot.handlers.plan_manager` и `from bot.handlers import plan_manager`.
- **Маршруты:** после сборки приложения под `/api/v1` доступны пути вида `/api/v1/offers`, `/api/v1/offers/{kp_id}`, PATCH discount / move-to-production, DELETE, PDF/XLSX (как описано в `endpoints/offers.py`).
- **Тесты:** подмножество тестов, завязанных на `plan_manager` и планирование (например `test_day_tracks_count`, `test_distribute_tracks_with_occupancy`, `test_plan_consistency`, `test_plate_audit`, серия `test_production_planning_service*`, `test_production_completion_service`) проходило успешно (43 теста в прогоне после изменений).

**Известное вне темы:** полный suite может содержать несвязанные падения (например исторический кейс «40 vs 40,0» в закупках); это не является регрессией данной оркестрации без отдельного доказательства.

---

## Риски и рекомендации

1. **Пути к данным:** при переносе модуля критична корректность `Path(__file__).resolve().parents[...]` и каталогов `bot/data/plans`. Рекомендуется смоук-тест: создать/прочитать план через API или бота и убедиться, что файлы по-прежнему в ожидаемой директории.
2. **Monkeypatch:** новые и будущие тесты должны патчить `app.planning.plan_manager`, а не устаревший путь через `bot.handlers`.
3. **Дальнейшее упрощение:** бот по-прежнему зависит от `app` (сервисы, домен) — это нормально; при желании следующим шагом можно вынести общие типы или тонкие порты без Telegram-специфики.

---

## Связанные материалы

- Аудиты: `ai_docs/develop/audits/` (полный аудит от 2026-05-06 и смежные документы).
- Конфиг путей документации: `.cursor/config.json` (`documentation.paths.reports` → данный файл в `ai_docs/develop/reports/`).
