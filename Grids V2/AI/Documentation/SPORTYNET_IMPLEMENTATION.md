# SportyNet Implementation

## Supported raw layouts

SportyNet is formatted in place on every selected matching worksheet. V2 accepts both observed source layouts:

1. **BRT-left layout**
   - `A2 = GMT`
   - `B2 = BRT` or `BRA`
   - seven dates in `C2:I2`
   - duplicate `BRT/BRA` and `GMT` in `J2:K2`

2. **Mexico-left layout**
   - `A2 = GMT`
   - `B2 = Mex`, `Mexico`, or `CDMX`
   - seven dates in `C2:I2`
   - authoritative `BRT/BRA` and `GMT` in `J2:K2`

Both layouts require 96 quarter-hour rows, a nonblank title in row 1, and the four-item legend. Structural text is read from `Value2`, not Excel's display-dependent `Text` property.

For the Mexico-left layout, V2 reads BRT from column J, validates the existing Mexico/CDMX times in column B, and writes the final standardized output as `GMT | BRT | CDMX | Monday ... Sunday`.

## Formatting behavior

The formatter normalizes source GMT/BRT values, including deterministic `6.45 -> 06:45`, validates the 15-minute sequences and GMT/BRT offset, validates an existing Mexico/CDMX column when present, inserts or standardizes CDMX, removes duplicate right-side time columns, preserves program merges/colors, applies the accepted text transformation and formatting, and saves the same workbook.

Already-formatted worksheets with `C2 = CDMX` are skipped.

## Routing guard

Existing grid filename routes always win. Known SportyNet filename families include `SNETL - Week 31` and versioned names such as `WEEK 31 V.4` or `WEEK 31 V.4 1`.

A plain unmatched name such as `Week 32.xlsx` receives only the lightweight first-worksheet row-2 signature check. The soft check accepts either BRT-left or Mexico-left SportyNet headers. Full tab scanning starts only after a known name or the soft check routes the workbook to SportyNet.

## Runtime defect fixed in 2.0.1-dev

The first Windows run exposed two detection defects:

- structural comparisons used `Range.Text`, which is display-dependent and caused a valid SNETL workbook to be rejected;
- the strict layout accepted only `BRT/BRA` in `B2`, while a real file used `Mex` in `B2` and authoritative BRT in `J2`.

V2 now uses `Value2` for structural strings and supports both source variants. Exact copies of `SNETL - Week 31.xlsx`, `WEEK 31 V.4 1.xlsx`, and `Week 32.xlsx` are retained as regression fixtures.


## Runtime defect fixed in 2.0.2-dev

The real Windows run exposed a second COM interop issue: a one-cell Excel property could arrive as `System.Object[]`, and direct `[string]` conversion failed during time/program extraction. The runtime now unwraps scalar COM values before conversion, uses `Value2` instead of `Range.Text`, safely reads merged-cell top-left values, and reports whether a remaining failure occurs in the time or program stage. Strict-layout failure messages now include the detected A2/B2/J2/K2 headers.
