# Issues Log

## OPEN-V2-001 — Windows Excel runtime regression testing

Static validation is complete, but V2 must be run on disposable copies of each supported grid type in Windows desktop Excel before replacing production.

## RESOLVED-V2-002 — Calculation property HRESULT 0x800A03EC

The V2 runtime does not set `Excel.Application.Calculation`. Formatting does not require that state change.

## OPEN-V2-003 — SportyNet golden parity

SportyNet Week 31 and the formatted reference are different weeks. Operator comparison remains required for final parity approval.

## RESOLVED-V2-004 — Broad SportyNet fallback scanning

The fallback now checks only row 2 of the first worksheet after all existing and known SportyNet filename routes fail. Full SportyNet scanning occurs only after that lightweight signature succeeds.

## RESOLVED-V2-005 — Valid SportyNet files rejected by strict layout

Observed Windows failures:

- `SNETL - Week 31.xlsx` was rejected even though its workbook structure was valid.
- `WEEK 31 V.4 1.xlsx` was rejected because `B2` contains `Mex` while BRT is in `J2`.
- `Week 32.xlsx` was skipped by the soft route.

Cause: structural labels were read with display-dependent `Range.Text`, and the strict layout did not support the Mexico-left source variant.

Correction in `2.0.1-dev`: use `Value2`, accept `Mex/Mexico/CDMX`, determine the authoritative BRT source column dynamically, and validate both time layouts.

Status: code and static fixture validation complete; Windows rerun remains required under `OPEN-V2-001`.


## ISS-005 — SportyNet COM value returned as System.Object[]

- **Status:** fixed in 2.0.2-dev; pending Windows rerun
- **Evidence:** `SNETL - Week 31.xlsx` and `Week 32.xlsx` reached the formatter but failed with `Unable to cast object of type System.Object[] to type System.String`.
- **Correction:** all SportyNet COM reads now pass through scalar-unwrapping helpers; direct string casts and `Range.Text` reads were removed from the affected path.
