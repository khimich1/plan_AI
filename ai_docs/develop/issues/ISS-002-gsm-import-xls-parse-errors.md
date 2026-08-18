# Issue: GSM import — unhandled xlrd errors / no content gate

**ID:** ISS-002  
**Discovered:** 2026-08-14 (during security audit of T3 POST /gsm/transactions/import)  
**Reported by:** security-auditor  
**Severity:** Medium  
**Security Impact:** Low (availability / noisy 500; AuthZ already required)  
**Status:** Open  

## Description

`POST /api/v1/gsm/transactions/import` reads uploads via `read_upload_file_capped` (size OK) and passes bytes to `xlrd.open_workbook(file_contents=...)` without:

1. Extension / magic / content-type allowlist for BIFF `.xls`
2. Catching `xlrd.XLRDError` (and similar) → map to HTTP 400

Non-xls bytes raise `XLRDError` and surface as an unhandled 500.

## Impact

- Exploit difficulty: Low (authenticated `admin` / `accountant` only)
- Data at risk: None (no RCE / no path write on this path)
- Attack vector: Upload garbage or non-BIFF payload → server error / possible detail leak if `APP_DEBUG`

## Why Not Fixed Now

- Scoped audit: AuthZ, per-file size cap, no FS write (path traversal N/A), parameterized SQL — all PASS
- Residual is input validation / error hygiene, not auth bypass or injection
- No rewrites requested in this audit pass

## Proposed Solution

- Reject non-xls early (filename suffix allowlist and/or BIFF magic / xlrd probe)
- Catch parse failures in endpoint or service → `400` with generic message
- Optionally basename-normalize `upload.filename` before storing in `gsm_import_batch` (defense in depth)

## Priority

P3 (fix when hardening GSM import UX / next GSM security pass)
