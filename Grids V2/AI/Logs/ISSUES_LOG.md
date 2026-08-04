# Issues Log

## PARTIALLY RESOLVED-V2-001 — Windows Excel runtime regression testing

Static validation is complete, but V2 must be run on disposable copies of each supported grid type in Windows desktop Excel before replacing production.

**2026-08-03 update:** Ran the live formatter (Windows desktop Excel, real COM automation, not a mock) against disposable copies:

- **SportyNet — genuinely regression-tested from fresh/unformatted input**, both layout variants (`BRT_LEFT` and `CDMX_LEFT_BRT_RIGHT`), across four filenames including one that matches no known naming pattern at all. All four formatted cleanly end to end (structural edits, time matrix, colors, program blocks, legend) with zero errors.
- **CATV, TVD, Pasiones LATAM, Pasiones US, Todo Novelas** — every real sample found on this machine (including older March/April 2026 copies pulled from Downloads) was already fully formatted, so only the "already done" skip-detection path was exercised, not a fresh transformation. Searched the whole machine for a genuinely unformatted sample of any of these types; none exists.
- **REV_TV** — no sample file of this type exists anywhere on this machine. Zero test coverage, not even the skip path.

None of today's code changes touch the CATV/TVD/Pasiones/Todo Novelas/REV_TV transformation logic at all (today's fixes were scoped to SportyNet's matrix-write COM bug, SportyNet's soft-signature routing, and the Settings UI), so there is no new regression risk from this session's work in those paths. But this item can't be fully closed until someone runs one genuinely fresh (never-formatted) copy of CATV/TVD/Pasiones/Todo Novelas/REV_TV through the formatter and confirms the output — that needs a real sample from an upcoming week, since none exists to test with today.

**2026-08-04 update:** Ran `AI/Tests/Run-Windows-SmokeTests.ps1` end to end on real Windows Excel COM for the first time (it now covers all supported SportyNet naming/layout variants plus an unrelated-file-untouched check and a Settings UI self-test) -- 7/7 pass. This is additional SportyNet-specific confirmation on top of the above; the CATV/TVD/Pasiones/Todo Novelas/REV_TV gap is unchanged since the smoke test's fixtures for those types are also pre-formatted, same limitation as before.

## RESOLVED-V2-002 — Calculation property HRESULT 0x800A03EC

The V2 runtime does not set `Excel.Application.Calculation`. Formatting does not require that state change.

## MOSTLY RESOLVED-V2-003 — SportyNet golden parity

SportyNet Week 31 and the formatted reference are different weeks. Operator comparison remains required for final parity approval.

**2026-08-03 update:** Since the golden reference (`AI/References/SportyNet/SNETL Transformat.xlsx`, Week 29) and the test data are genuinely different weeks, an exact content diff isn't meaningful -- schedule content should differ; only structure/styling should match. Ran a live COM comparison of every styling attribute the formatter controls (title merge/font/fill/alignment, header row, time-matrix number format/fonts/colors/borders, date-header row, a sampled program-block cell's font/wrap/alignment, legend row, column widths, row heights, shape count) between the golden reference and a freshly-formatted Week 31 output. 5 of ~30 checked attributes differed:

- Golden's title cell (A1) is **not** merged/centered, unlike the current code's `A1:C1` merge + center. Golden's time-matrix `NumberFormat` reads back blank (Excel returns that when a range has mixed formats internally, i.e. golden's own cells aren't uniformly `h:mm`). Both point to the golden reference itself predating or diverging from the current, intentional formatting rules -- not a defect in today's output, which is the more internally consistent of the two.
- Column widths differ by <1 unit (10 vs 10.71, 77 vs 77.71). This is Excel's own well-known `ColumnWidth` get/set asymmetry -- it rounds through pixel width using the workbook's default font metrics, so reading back a value you just set rarely matches exactly. Confirmed by checking the code: it sets `10.77734375` / `77.77734375` explicitly (`FormatGrids.ps1` line 750); the small drift on readback is Excel, not the formatter. Imperceptible visually.

Every other checked attribute matched exactly. This closes out the mechanical/structural side of parity; the subjective "does it look right" sign-off this issue originally asked for is still the operator's call, but there's no evidence of a formatting regression to sign off on.

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


## RESOLVED-ISS-005 — SportyNet COM value returned as System.Object[]

- **Status:** confirmed on real Windows Excel COM (2026-08-03/04), including the related matrix-write cast bug found in the same area (`Range.Value2` mis-resolving a bulk 2D array write on this install; fixed by using `Range.Value` instead).
- **Evidence:** `SNETL - Week 31.xlsx` and `Week 32.xlsx` reached the formatter but failed with `Unable to cast object of type System.Object[] to type System.String`.
- **Correction:** all SportyNet COM reads now pass through scalar-unwrapping helpers; direct string casts and `Range.Text` reads were removed from the affected path.

## RESOLVED-ISS-006 — WEEK 30 internal blank program line

- **Status:** confirmed on real Windows Excel COM (2026-08-04).
- **Evidence:** `WEEK 30 V2.xlsx`, source column D (pre-insert), row 73, span 8 -- contains `AMISTOSO INTERNACIONAL / 2026 / [blank] / Benfica x Villareal / 17/07 / vt`.
- **Previous behavior:** the formatter stopped at the program stage because the earlier plan treated every internal blank physical line as invalid.
- **Correction:** preserve physical positions while applying the two manual end-based joins; a blank join target consumes the next meaningful fragment without adding a leading separator. Remaining blank placeholders are removed afterward.
- **Confirmed output:** the block lands at post-format cell **E73** (source column D shifts by +1 for the CDMX column insert -- not D73), holding exactly `AMISTOSO INTERNACIONAL` / `2026` / `Benfica x Villareal` / `17/07 vt` on four lines.
- **Found and fixed in the same pass:** `AI/Tests/Run-Windows-SmokeTests.ps1` checked `Cells.Item(73,4)` (D73) instead of `Cells.Item(73,5)` (E73). D73 is a non-anchor cell of a *different* block (the column-C-sourced block spanning rows 68-75), so it legitimately reads empty regardless of correctness -- that's normal Excel merged-cell behavior, not a formatter bug. This is exactly the "pending Windows rerun" this issue and `OPEN-V2-001` called for; running it for the first time surfaced a real test bug, not a formatter bug. Fixed; full suite now passes 7/7.

## RESOLVED-ISS-007 — Settings layout controls were misleading

- **Status:** confirmed on real Windows (2026-08-04) via a live screenshot of the redesigned UI.
- **Evidence:** the Layout tab displayed the same generic values for every grid even though Pasiones/Todo and SportyNet use format-specific overrides.
- **Correction:** remove the Layout tab from the simplified V1 manager UI. Formatting dimensions remain implementation-owned and fixed per grid type.

## RESOLVED-ISS-008 — Settings launcher leaves a PowerShell console visible

- **Status:** confirmed on real Windows (2026-08-04) -- launched via `Open Settings.vbs` and checked running processes; only the Settings window itself has a visible window, no bare console.
- **Correction:** `Open Settings.bat` delegates to `Open Settings.vbs`, which launches Windows PowerShell with a hidden window. Startup errors are logged and displayed through a message box.
