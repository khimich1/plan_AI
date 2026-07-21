# Telegram Bot — DEPRECATED

**Status:** soft-decommissioned (P5 WP1, 2026-06-21). Not used in production.

| Item | Location |
|------|----------|
| Archived source | [`../bot_archived/`](../bot_archived/) |
| Canonical product path | Web app (`app/`, `frontend/`) |
| Hard delete | Planned for **P6** — see `docs/specs/p6-legacy-decommission.md` |

## Do not use

- Do **not** run the bot in production or CI.
- Do **not** import from `bot_archived/` in `app/`, `core/`, or active `tests/`.
- `python run_bot.py` prints a deprecation message and exits with code 1.

## Rollback / reference

Git history and `bot_archived/` preserve handlers and services for audit only.
To inspect archived code, browse `bot_archived/` — it is not on the default Python path.
