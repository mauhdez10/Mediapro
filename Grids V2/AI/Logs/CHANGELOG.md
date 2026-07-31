# Changelog

## V2 development

- Created an isolated Grids V2 copy; production Grids remains unchanged.
- Added strict routing priority and SportyNet support.
- Added per-grid tab selection with All/First/Last behavior.
- Added bilingual English/Spanish console messages.
- Added a Settings UI that embeds configuration into the runtime script.
- Moved printer state into Settings with a no-Settings local fallback.
- Removed all dependency on setting `Excel.Application.Calculation`, addressing HRESULT `0x800A03EC`.
- Added Full, Manager, and Operator distribution builders.
- Tightened the unmatched-file SportyNet fallback to a first-worksheet row-2 header check. Plain names such as `Week 32.xlsx` use this soft check instead of direct filename classification.
- Consolidated Windows runtime validation into one smoke-test entry point to keep the AI test folder simple.

## 2.0.1-dev — SportyNet real-file detection fix

- Replaced display-dependent `Range.Text` structural checks with `Value2`.
- Added support for the observed `Mex`/`Mexico`/`CDMX` left-time layout with BRT in column J.
- Added validation of an existing Mexico/CDMX column against BRT minus three hours.
- Added exact real-file fixtures for `SNETL - Week 31.xlsx`, `WEEK 31 V.4 1.xlsx`, and `Week 32.xlsx`.
- Updated the Windows smoke test to use the actual files instead of renamed copies of one reference workbook.
- Kept all existing CATV, TVD, Pasiones, Todo Novelas, and REV TV routing unchanged.


## 2.0.2-dev — Excel COM scalar-unwrapping fix

- Added safe unwrapping for `System.Object[]` and multidimensional COM values.
- Removed `Range.Text` from SportyNet time extraction.
- Reads merged program/title cells through their scalar top-left value.
- Added time-stage and program-stage error prefixes.
- Added header probes to strict-layout failures.
- Updated exact-file regression tests and distribution packages.
