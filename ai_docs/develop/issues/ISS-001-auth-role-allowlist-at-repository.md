# Issue: Role allowlist only at HTTP schema, not repository

**ID:** ISS-001  
**Discovered:** 2026-08-14 (during security audit of Task T2 accountant role)  
**Reported by:** security-auditor  
**Severity:** Low  
**Security Impact:** Low  
**Status:** Open  

## Description

`app_users.role` is stored as free `TEXT`. HTTP registration is correctly constrained by `RegisterUserRequest.role: Literal[...]`, but `AuthRepository.create_or_update_user` and `scripts/create_admin.py --role` accept any non-empty string. Authorization (`require_roles`) compares exact DB values and fails closed, so this is defense-in-depth rather than an active bypass.

## Impact

- Exploit difficulty: High (requires calling repository/CLI/direct DB, not the public register API as a non-admin)
- Data at risk: Privilege assignment only if a future caller skips schema validation
- Attack vector: Misconfigured internal script or new endpoint that passes unchecked `role` into `create_or_update_user`

## Why Not Fixed Now

- HTTP path for T2 is correctly allowlisted (`accountant` added to Literal; unknown roles rejected)
- `get_current_user` loads role from DB; cookie `role` claim is not used for AuthZ
- SQL uses parameterized binds — no injection via role text
- Current task (GSM accountant AuthZ) is not blocked

## Proposed Solution

- Reuse a single allowlist (constants / shared Literal) inside `create_or_update_user` (and CLI) so all write paths enforce the same roles
- Optionally map `REQUIRE_*` helpers to those constants to avoid string drift

## Priority

P3 (fix when touching auth registration next)
