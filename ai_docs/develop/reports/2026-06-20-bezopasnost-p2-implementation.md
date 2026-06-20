# Report: Безопасность P2 — rate limit, RBAC, frontend 409

**Дата:** 2026-06-20  
**Orchestration:** `orch-2026-06-20-bezopasnost-p2`  
**Статус:** ✅ Completed (G3 closed)

## Summary

Закрыт обязательный scope спринта P2: brute-force protection на login (S2), object-level RBAC для REST API offers/archive (S3), UX перезагрузки плана при optimistic-lock conflict (FE-409). WP3 (S4 npm CVE) отложен. Health Score: **~7/10**.

## What Was Built

| WP | Реализация | Верификация |
|----|------------|-------------|
| **WP0 S2** | `app/security/login_rate_limit.py`, wire в `auth.py` | `tests/test_auth_login_rate_limit.py` |
| **WP1 S3** | `app/security/offer_access.py`, `owner_user_id` в `kp_meta`, repository filters, offers + archive endpoints | `tests/test_offers_authorization.py`, `tests/test_archive_authorization.py` |
| **WP2 FE-409** | `frontend/src/shared/lib/planConflict.ts`, mutation handlers в production UI | `npm run test`, `npm run build` |
| **WP3 S4** | — | **deferred** |

## Metrics

- **pytest:** 756 passed, 12 skipped (`pytest tests/ -q`, 2026-06-20)
- **Frontend:** test + build green
- **Health Score:** ~6/10 → **~7/10** (S2, S3 RESOLVED)

## Known Gaps

- **Legacy web** (`app/web/router.py`) — без `offer_access`
- **Bot** — нет `owner_user_id` parity (bot deprecated)
- **S4** npm CVE — отдельный chore

## Related Documentation

- Spec: [`bezopasnost-p2-audit-2026-06-19.md`](../../specs/bezopasnost-p2-audit-2026-06-19.md) — closed
- Plan: [`2026-06-19-bezopasnost-p2.md`](../plans/2026-06-19-bezopasnost-p2.md) — closed
- Audit Post-P2: [`2026-06-19-full-project-audit.md#post-p2-remediation-status-2026-06-20`](../audits/2026-06-19-full-project-audit.md#post-p2-remediation-status-2026-06-20)

## Next Steps

1. Quality sprint: **Q5**, **Q6**
2. **S4** — npm audit + pinned versions + CI
3. Legacy web RBAC parity или deprecate (A10)
4. Backlog: S10, S6, strangler `cfg.PLATES_*`
