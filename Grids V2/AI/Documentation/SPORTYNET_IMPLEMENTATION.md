# SportyNet Implementation — Grids V2 2.0.4-dev

SportyNet is integrated as an additional grid type after all existing filename routes. Known SportyNet names route directly; only unmatched workbooks receive the lightweight first-sheet signature check.

The formatter supports BRT-left and Mexico/CDMX-left source layouts and standardizes output to GMT, BRT, CDMX, and seven day columns.

## Program text

The manual text rule is applied using stable physical line indexes from the end:

1. Move the original final line into the previous line.
2. Move the original third-to-last line into its previous line.

Trailing empty lines are removed. A real client workbook (`WEEK 30 V2.xlsx`, `Sheet1!D73:D80`) uses an internal blank physical line as the destination of the second move. That pattern is valid: joining a blank target with a meaningful source returns the meaningful fragment without a leading separator. Any blank placeholders remaining after the two moves are removed.

The authoritative WEEK 30 result is:

```text
AMISTOSO INTERNACIONAL
2026
Benfica x Villareal
17/07 vt
```

Regression coverage is provided by `AI/Tests/Test-SportyNetProgramText.ps1` and the exact WEEK 30 fixture in the Windows smoke suite.

Desktop Excel COM testing is still required before V2 promotion.
