# Regression Log

## REG-V2-001 — SportyNet detection passed synthetic tests but failed real files

- **Discovered:** 2026-07-31
- **Affected build:** 2.0.0-dev
- **Evidence:** `SNETL - Week 31.xlsx`, `WEEK 31 V.4 1.xlsx`, and `Week 32.xlsx`
- **Failure:** valid SportyNet files were rejected or skipped.
- **Root cause:** the smoke test renamed one reference workbook for every filename case, so it did not exercise the real SNETL workbook or the `Mex` header variant. Structural checks also used `Range.Text`.
- **Correction:** `2.0.1-dev` uses exact fixtures, `Value2`, and dynamic BRT/CDMX source columns.
- **Regression test:** `AI/Tests/Run-Windows-SmokeTests.ps1`
- **Status:** fixed in code; pending Windows Excel confirmation.


## REG-002 — Valid SportyNet sheets fail during extraction

- **Status:** fixed in 2.0.2-dev; pending Windows validation
- **Last known result:** routing/layout validation succeeded for SNETL and Week 32.
- **Failure:** time or program extraction attempted to cast a COM `System.Object[]` directly to string. WEEK 31 V.4 1 also returned a generic strict-layout failure.
- **Correction:** scalar unwrapping, safe Value2 conversion, merged-cell top-left reads, stage diagnostics, and header probes.
- **Regression fixtures:** `SNETL - Week 31.xlsx`, `WEEK 31 V.4 1.xlsx`, `Week 32.xlsx`.
