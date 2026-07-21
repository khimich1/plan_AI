# Post-P2 → P3 Direction

## Problem Statement

**Как зафиксировать реальное состояние после closure P2 и аудита 2026-06-21, чтобы следующий спринт не повторял сделанное и не спорил с закрытыми решениями?**

## Recommended Direction

P2 **не переоткрывать** — WP1–WP6 выполнены по заявленному scope (web/core, bot frozen). Аудит 2026-06-21 — **следующий слой** (full repo, новые ID). Следующий артеfact: **spec P3** + post-closure delta в P2 + cross-ref в audit.

Приоритет P3 (web-only):

1. DI gap: `offers.py`, `managers.py` (A3 audit-21)
2. Thin `PlanRepository` (A2) — P2 закрыл entry point, не thin repo
3. `core`↔`viz_modules` ports slice (A1)
4. Security quick wins: Swagger off prod (S2), role whitelist (S7), health metadata (S9)

## Key Assumptions to Validate

- [ ] Bot остаётся deprecated → bot findings = backlog, не P0
- [ ] Single-instance deploy → S1 Redis defer (decision D2 in P3 spec)
- [ ] `GET /managers` для production — accepted (decision D1 in P3 spec)

## MVP Scope (P3)

**In:** WP1–WP4 из [`stabilizaciya-p3-architecture-2026-06-21.md`](../specs/stabilizaciya-p3-architecture-2026-06-21.md)

**Out:** bot refactor, PostgreSQL, CreatePlanWizard, CSP enforce, OCR consent

## Not Doing

- Reopen P2 as `open`
- Fix bot god-modules in P3
- Chase full-repo Health Score 2→9 in one sprint

## Open Questions

- D1: `/managers` + production role — keep or restrict PII?
- D2: multi-worker soon → S1 in P0?
- D3: A2 thin repo — mandatory P3 or defer?

---

*2026-06-21 · idea-refine session output*
