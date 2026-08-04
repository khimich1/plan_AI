# Agent Skills в Cursor — руководство для проекта «Шишов»

Набор [agent-skills](https://github.com/addyosmani/agent-skills) адаптирован для этого репозитория. Скиллы лежат в `.cursor/skills/` — Cursor подхватывает их автоматически.

## Быстрый старт

1. Откройте проект в Cursor.
2. В чате опишите задачу на русском или английском.
3. Агент сам выберет скилл по описанию. Можно явно попросить: *«используй скилл test-driven-development»*.
4. Для любой задачи в этом репо сначала применяется скилл **`project-shishov`** (контекст проекта).

## Как это устроено

```
.cursor/
├── skills/                    # 25 скиллов (англ.) + project-shishov (рус.)
│   ├── using-agent-skills/    # Мета-скилл: какой скилл когда брать
│   ├── project-shishov/       # Контекст вашего проекта
│   └── …                      # spec, tdd, review, ship и др.
├── hooks.json                 # sessionStart — подсказка агенту при старте
├── hooks/session-start.ps1    # Windows-хук
└── СКИЛЛЫ-РУКОВОДСТВО.md       # этот файл

agent-skills-main/             # Исходники для переустановки (в репо)
├── skills/                    # копия скиллов
└── scripts/install-to-cursor.ps1
```

Папка `.cursor/` в `.gitignore` — настройки локальные. Чтобы переустановить скиллы на другой машине:

```powershell
Set-Location "путь\к\Шишов"
.\agent-skills-main\scripts\install-to-cursor.ps1
```

## Какой скилл когда использовать

| Ситуация | Скилл | Что делает |
|----------|-------|------------|
| Не знаю, что хочу | `interview-me` | Выясняет требования вопросами |
| Сырая идея | `idea-refine` | Прорабатывает идею, MVP |
| Новая фича | `spec-driven-development` | Спека до кода |
| Разбить на задачи | `planning-and-task-breakdown` | План с чеклистом |
| Писать код | `incremental-implementation` | Вертикальные срезы |
| API / FastAPI | `api-and-interface-design` | Контракты, endpoints |
| React UI | `frontend-ui-engineering` | Доступность, UX |
| Тесты | `test-driven-development` | Red → green → refactor |
| Ошибка / traceback | `debugging-and-error-recovery` | Воспроизвести → починить |
| Ревью кода | `code-review-and-quality` | 5 осей качества |
| Безопасность | `security-and-hardening` | OWASP, валидация |
| Упростить код | `code-simplification` | Меньше сложности |
| Git / коммиты | `git-workflow-and-versioning` | Атомарные коммиты |
| Документация | `documentation-and-adrs` | ADR, почему так |
| Деплой | `shipping-and-launch` | Чеклист релиза |
| **Любая задача здесь** | **`project-shishov`** | Структура репо, стек, команды |

Полная блок-схема выбора — в файле [skills/ОБЗОР-МЕТА-СКИЛЛА.ru.md](skills/ОБЗОР-МЕТА-СКИЛЛА.ru.md).

## Полезные фразы в чате

- «Проработай идею» → `idea-refine`
- «Напиши спеку на …» → `spec-driven-development`
- «Разбей на задачи» → `planning-and-task-breakdown`
- «Сделай по TDD» → `test-driven-development`
- «Проведи код-ревью» → `code-review-and-quality`
- «Почему падает …» → `debugging-and-error-recovery`

## Язык скиллов

- **Для вас (читать):** русские файлы `*.ru.md` в `.cursor/skills/` и это руководство.
- **Для агента:** `SKILL.md` на английском — так точнее следует инструкциям. Описание (`description`) у `project-shishov` на русском, чтобы агент чаще его подхватывал.

## Обновление скиллов

```powershell
# Скачать свежую версию с GitHub и установить
Set-Location "c:\Users\Роман\Desktop\Шишов"
git clone --depth 1 https://github.com/addyosmani/agent-skills.git agent-skills-update
Copy-Item agent-skills-update\skills\* .cursor\skills\ -Recurse -Force
Copy-Item skills\project-shishov .cursor\skills\ -Recurse -Force  # сохранить свой скилл
Remove-Item agent-skills-update -Recurse -Force
```

Или запустите `install-to-cursor.ps1` после обновления `agent-skills-main/skills/`.

## Лицензия

Скиллы Addy Osmani — [MIT](https://github.com/addyosmani/agent-skills/blob/main/LICENSE). Скилл `project-shishov` — локальный для этого проекта.
