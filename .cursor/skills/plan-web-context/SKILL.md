---
name: plan-web-context
description: Project facts for plan_web (commercial offers, plates, production planning). Use at the start of any non-trivial task, when switching domains (pricing, KP, layout, auth, frontend), or when another skill needs stack/paths/commands.
---

# Plan Web — Project Context

Read this before implementing. Prefer existing modules over inventing parallel ones.

## What This Product Is

Web app for **commercial offers (КП)**, **plate/pricing logic**, **production planning/layout**, and related admin. Domain language is mostly Russian (плиты, раскладка, прайс, КП, СГП, отходы).

Telegram bot is **deprecated** (`bot_archived/`, `requirements-bot.txt` optional only). Do not extend the live bot path.

## Stack

| Layer | Tech |
|-------|------|
| Backend | Python 3, FastAPI (`app.main:app`), Pydantic v2, uvicorn |
| Domain/shared | `core/` (pricing, KP DB, Excel/PDF, ILP via PuLP/CBC) |
| Frontend | React 19, TypeScript, Vite 8, React Router 7, TanStack Query, react-hook-form, Zod 4 |
| Data | SQLite (`plita.db`, `pb.db` — local/seed; do not commit secrets or live DBs) |
| Tests | Backend: pytest (`tests/`); Frontend: vitest + Testing Library |
| Run | `./run+logs.sh` / `./run_local.sh`; Docker via `docker-compose.yml` |

## Key Paths

```
app/                    # FastAPI app
  api/v1/endpoints/     # REST: auth, admin, commercial, offers, production, archive, managers, health
  domain/ services/ repositories/ schemas/ security/
  web/                  # SPA + legacy web routes
core/                   # Shared business logic (pricing, KP, layout, Excel)
frontend/src/           # SPA (app/, features/, pages/, shared/)
tests/                  # pytest (ignore tests/archived/)
ai_docs/                # AI plans/reports (local; see .cursor/config.json)
docs/                   # Human-facing docs / specs
.cursor/                # Rules, agents, commands, skills
```

## Commands

```bash
# Backend
source venv/bin/activate   # if needed
pytest                     # from repo root; archived bot tests skipped via pytest.ini
uvicorn app.main:app --reload

# Frontend
cd frontend && npm run dev
cd frontend && npm run test
cd frontend && npm run typecheck
cd frontend && npm run build

# Full local stack
./run+logs.sh              # or ./run_local.sh
```

## Architecture Habits

1. **API first in `app/api/v1/`** — routers thin; logic in `services/` / `core/`.
2. **Schemas in `app/schemas/`** — Pydantic models are the contract; keep frontend types aligned.
3. **Pricing/plates/KP** — check `core/` and existing endpoints (`commercial`, `offers`) before new modules.
4. **Destructive DB ops** — respect `core/destructive_db_guard.py` and admin guards; never bypass in production.
5. **Secrets** — `.env` / `bot.env` are gitignored; never commit tokens or live DBs.

## Documentation Map

Paths come from `.cursor/config.json`:

| Kind | Path |
|------|------|
| Plans | `ai_docs/develop/plans` |
| Reports | `ai_docs/develop/reports` |
| Issues | `ai_docs/develop/issues` |
| Architecture | `ai_docs/develop/architecture` |
| Features | `ai_docs/develop/features` |
| API notes | `ai_docs/develop/api` |
| Audits | `ai_docs/develop/audits` |
| Specs (human) | `docs/specs`, `ai_docs/specs` |

Also see: `ai_docs/develop/cursor-dot-folder-guide-RU.md`.

## Existing Cursor Workflows (prefer these for big work)

| Trigger | Skill / command |
|---------|-----------------|
| Complex feature | `/orchestrate` → `orchestration` |
| Small task | `/implement` → `simple-workflow` |
| Review | `/review` → `review-workflow` |
| Refactor | `/refactor` → `refactor-workflow` |
| Audit | `/audit` → `audit-workflow` |

Engineering-phase skills (`spec-driven-development`, `test-driven-development`, …) plug into these workflows; they do not replace them.

## Domain Hotspots (be careful)

- Plate pricing, longitudinal cuts, waste/cascade costing
- Commercial offer generation (PDF/XLSX/DOCX)
- Layout / reinforcement / production planning
- Auth + admin destructive actions
- Archive / KP DB boundaries (`app` vs `core.kp_db`)

When touching pricing or plates, add/adjust pytest coverage under `tests/` before claiming done.
