# GLM SportyNet Excel Parity Project

## 1. Project decision

Build **an Excel-first, deterministic parity tool** before attempting any redesign.

Recommended first deliverable:

- A separate macro-enabled controller workbook: `SportyNet_Formatter.xlsm`.
- The controller opens a client `.xlsx` file read-only, creates a new formatted `.xlsx` output, and never changes the source.
- VBA is the production implementation for Phase 1.
- Windows Excel COM/PowerShell is the source of truth for integration testing.
- Printing automation, a local app, or a non-VBA rewrite is deferred until the parity version is accepted.

This is safer than placing macros inside every weekly source workbook. An `.xlsx` output remains free of embedded macros, while the reusable automation remains in the controller workbook.

## 2. Required GLM configuration

Use the following for every critical pass until parity is accepted:

- **Model:** `GLM-5.2`
- **Thinking:** enabled
- **Effort:** `max`

In a Claude Code-style GLM session, use `/effort max`. Do not use `high` as a substitute: GLM maps `low`, `medium`, and `high` to its high tier, while `xhigh`, `max`, and `ultracode` map to the actual max tier.

Recommended model routing:

1. Workbook forensics and implementation plan: GLM-5.2, max.
2. Core VBA implementation: GLM-5.2, max.
3. Excel integration testing and repair: GLM-5.2, max.
4. Final independent review in a fresh session: GLM-5.2, max.
5. GLM-4.7 may be used only after acceptance for routine documentation, renaming, comments, and low-risk cleanup.

## 3. Files inspected and their roles

| File | Role |
|---|---|
| `Canada Aug 01 to Aug 07.xlsx` | Current beIN weekly source grid |
| `English - 08-01 to 08-07.xlsx` | Current beIN weekly source grid |
| `Spanish 08-01 to 08-07.xlsx` | Current beIN weekly source grid |
| `Template - Grid Format.xltm` | Current beIN macro/template implementation |
| `WEEK 31 V.4.xlsx` | SportyNet unformatted source example |
| `SNETL Transformat.xlsx` | SportyNet manually formatted visual/style reference |

### Important limitation

`WEEK 31 V.4.xlsx` is Week 31, while `SNETL Transformat.xlsx` is Week 29. They are not the same data. The formatted file is therefore a **style and structure reference**, not a complete golden-output pair.

Before final parity approval, obtain one of these:

- the original unformatted Week 29 file that produced `SNETL Transformat.xlsx`, or
- a manually formatted Week 31 output created from `WEEK 31 V.4.xlsx`.

That matching before/after pair becomes the authoritative golden fixture for cell-by-cell and print comparison.

## 4. What the current beIN macro actually does

The existing `Template - Grid Format.xltm` contains a worksheet macro named `format_fonts`. It processes three vertically stacked, single-day pasted grids using hard-coded row ranges.

Its actual operations are:

- unmerge the three grid headings at `A1`, `A118`, and `A228`;
- set all of column B to Excel `ColumnWidth = 100`;
- set column B to Calibri 22 pt;
- set columns A and C to 16 pt;
- unmerge all of column B;
- autofit rows;
- replace every line break in `B4:B400` with `" - "`;
- re-merge three fixed title blocks;
- set titles to Algerian 48 pt;
- delete entire rows when column B is blank within three fixed ranges.

This explains why it effectively works once: it destroys merge structures and deletes rows. It also relies on `Selection`, whole-column formatting, and hard-coded ranges.

**Do not extend this code directly for SportyNet.** Use it only as evidence of the current operator experience and as a warning about destructive operations.

## 5. SportyNet source structure found in `WEEK 31 V.4.xlsx`

The supplied source has this logical layout:

- Row 1: title stored in `A1`, merged across `A1:N1`.
- Row 2:
  - `A2 = GMT`
  - `B2 = BRT`
  - `C2:I2 = Monday through Sunday dates`
  - `J2 = BRT`
  - `K2 = GMT`
- Rows 3:98: 96 quarter-hour schedule rows.
- Columns C:I: seven program/day columns with vertical merges representing program duration.
- Rows 99:100: blank.
- Row 101: legend entries.
- The workbook contains formatting pollution through approximately `Z1000`, even though the real schedule is much smaller. Do not use Excel `UsedRange` as the layout detector.
- The provided file contains an empty drawing part but no embedded image. Future weekly files may still contain a logo or shape, so shape removal must remain supported.
- At least one supplied GMT value is malformed as a raw number (`6.45`). CDMX must therefore be calculated from the valid BRT column, not from GMT.

## 6. SportyNet manually formatted reference found in `SNETL Transformat.xlsx`

The supplied formatted reference has this logical target:

- `A = GMT`
- `B = BRT`
- `C = CDMX`
- `D:J = Monday through Sunday`
- the duplicate right-side BRT/GMT columns are gone;
- program merges and category fills are preserved;
- title styling occupies `A1:C1` visually;
- schedule rows are 18 pt high;
- program text is 14 pt and wrapped;
- time columns have complete thin borders;
- CDMX has a light-blue fill;
- weekday columns are much wider than the time columns;
- legend appears on row 100 in the reference.

Exact measurements in the reference workbook:

| Element | Reference value |
|---|---:|
| Time columns A:C | `ColumnWidth = 10.77734375`, approximately 80 px |
| Weekday columns D:J | `ColumnWidth = 77.77734375`, approximately 550 px |
| Row 1 | 22.8 pt |
| Row 2 | 24 pt |
| Rows 3:98 | 18 pt |
| Legend row | 13.2 pt |
| Program font | Arial 14 pt; existing bold/fill semantics preserved |
| Date header format | `mm/dd dddd` |
| CDMX fill | Theme Accent 1 with approximately 80% tint |
| Title fill | Theme Accent 2 with approximately 80% tint |

## 7. Conflicts that GLM must not silently guess about

Use the following precedence rule:

1. A matching same-week manually formatted golden output wins.
2. If no matching golden output exists, explicit written instructions win for behavior.
3. `SNETL Transformat.xlsx` wins for visual style when the written instruction is only descriptive.
4. Any remaining ambiguity must be documented in `docs/SPORTYNET_PARITY_DECISIONS.md` before implementation.

Known conflicts or omissions:

1. Written instruction says weekday columns should be 700 px; the formatted reference is approximately 550 px. The old beIN macro uses `ColumnWidth = 100`, which is approximately 700 px.
2. Written instruction says merge the first three title cells; the formatted reference has no actual row-1 merge.
3. Written instruction says `BRA`; both supplied files use `BRT`.
4. Written instruction mentions 40 px for the weekday row/hour area; the formatted reference has a 24 pt row-2 height and approximately 80 px time columns.
5. Written instructions omit deletion of the duplicate right-side BRT/GMT columns, but the formatted reference removes them.
6. Written instructions omit moving the legend from row 101 to row 100, but the formatted reference does so.
7. The reference makes GMT times bold while BRT and CDMX time values are regular; this is not stated in the written requirements.
8. The source title is in `A1`, not the second cell.
9. The supplied source has no visible logo object, although future files may.

### Recommended defaults until a matching golden pair exists

- Keep the label `BRT`.
- Remove duplicate right-side BRT/GMT columns.
- Use 80 px time columns.
- Use the formatted reference width of approximately 550 px for weekday columns, because the objective is current-manual parity. Treat 700 px as a possible future change.
- Merge `A1:C1`, because that is the explicit written instruction, and center the title.
- Place the legend at `scheduleLastRow + 2`.
- Preserve the existing bold state of all program cells and time cells unless the golden pair proves a forced change is required.
- Delete shapes/pictures if present.
- Warn about invalid GMT values but continue when BRT is valid.

## 8. Exact Phase-1 transformation contract

### 8.1 Safety and preflight

1. Prompt the operator to choose the source `.xlsx` file.
2. Open the source read-only.
3. Never save over the source.
4. Validate the layout before making an output:
   - locate the header row within the first 10 rows;
   - accept `BRT` or `BRA` as the left local-time header;
   - locate exactly seven contiguous date columns;
   - locate the duplicate right-side local-time and GMT headers if present;
   - locate 96 BRT quarter-hour values;
   - locate the legend by its labels rather than a hard-coded row;
   - reject the file with a clear message when the structure is not recognized.
5. Capture the source title before unmerging row 1.
6. Capture all day-column merged areas before any insert/delete operation.

### 8.2 Output construction

Create a fresh output workbook or fresh output worksheet so style pollution outside the real schedule is not carried forward.

The final schedule area must be:

- A: GMT
- B: BRT
- C: CDMX
- D:J: seven date/day columns

Do not retain the duplicate right-side BRT/GMT columns or trailing blank columns.

### 8.3 CDMX calculation

For every schedule row:

- calculate CDMX from BRT;
- `CDMX = BRT - 3 hours`, wrapped modulo 24 hours;
- write numeric Excel time values, not display strings;
- use a time number format equivalent to `h:mm`;
- do not derive CDMX from GMT because supplied GMT data can be malformed.

### 8.4 Program text line-break transformation

Normalize CRLF/CR to LF first. Trim joined line fragments and insert exactly one separating space.

For an unmerged day cell spanning one row:

- replace every line break with a single space;
- the result is one line.

For a vertically merged day cell spanning more than one row:

1. Work from the original line list.
2. If the original third-to-last line is not the first line, move it to the end of the preceding line with one space between them.
3. Move the original last line to the end of the preceding line with one space between them.
4. Preserve all other line breaks.

Examples:

```text
Input, six lines:
PROG00000001
CAMP. BRASILEIRO
2026
14ª RODADA
Team A x Team B
<Repeat>

Output:
PROG00000001
CAMP. BRASILEIRO
2026 14ª RODADA
Team A x Team B <Repeat>
```

```text
Input, three lines:
PROG00000002
ESPECIAL COPAS
<Repeat>

Output:
PROG00000002
ESPECIAL COPAS <Repeat>
```

```text
Input in a one-row unmerged cell:
PROG00000003
NO CENTRO DO JOGO
<Repeat>

Output:
PROG00000003 NO CENTRO DO JOGO <Repeat>
```

The function must be independently unit-tested before it is applied to workbook cells.

### 8.5 Formatting

Apply only the required changes and preserve schedule-category fills and existing program bold semantics.

- Time headers: Arial 14 pt, bold, center/middle.
- Time values: Arial 14 pt, center/middle, `h:mm`.
- CDMX values: light-blue fill matching the reference.
- All cells in the three time columns: thin border on all four sides.
- Date headers: Arial 14 pt, bold, center/middle, `mm/dd dddd`, preserve pale-yellow fill.
- Day/program columns: Arial 14 pt, wrap text, center/middle for populated program areas.
- Preserve all program merges and category colors.
- Weekday columns: use the accepted parity width.
- Schedule rows: use the accepted parity row height.
- Remove all pictures/shapes from the output schedule sheet when they exist.

### 8.6 Title row

1. Capture the original title from the first nonblank cell in row 1.
2. Unmerge all row-1 merged areas.
3. Clear row 1 across the output area.
4. Merge `A1:C1`, unless the accepted golden specification says not to merge.
5. Write the title.
6. Apply Arial 9 pt, bold, center/middle, light-red fill matching the reference.

### 8.7 Legend

- Locate the legend from content, not a fixed row.
- Preserve the four legend labels and their fills/borders.
- Place it at the accepted target row, recommended `scheduleLastRow + 2`.

### 8.8 Save behavior

- Save the output beside the source as `<original name>_FORMATTED.xlsx`.
- If the name already exists, use a timestamp or version suffix; never overwrite silently.
- Close the source without saving.
- Display a completion summary with the output path and validation results.

## 9. Phase-1 non-goals

Do not include these in the first parity release:

- automatic printing of all seven days;
- printer-specific settings;
- a local desktop application;
- a web UI;
- Python-only workbook rewriting;
- changes to the existing beIN macro/template;
- redesign of the manual line-break rules;
- automatic correction of source schedule data.

## 10. Repo structure

Recommended layout:

```text
docs/
  SPORTYNET_EXCEL_PARITY_PLAN.md
  SPORTYNET_REQUIREMENTS_MATRIX.md
  SPORTYNET_PARITY_DECISIONS.md
  CURRENT_BEIN_MACRO_AUDIT.md
src/
  vba/
    modEntryPoints.bas
    modLayoutDetection.bas
    modWorkbookTransform.bas
    modTextTransform.bas
    modFormatting.bas
    modValidation.bas
    modFileOutput.bas
  workbook/
    SportyNet_Formatter.xlsm
tests/
  vba/
    modTextTransformTests.bas
  powershell/
    Invoke-SportyNetParityTests.ps1
    Get-ExcelWorkbookSnapshot.ps1
  expected/
    README.md
fixtures/
  private/
    .gitkeep
scripts/
  Export-VbaModules.ps1
  Import-VbaModules.ps1
  Build-Formatter.ps1
dist/
  README.md
.gitignore
```

Rules:

- Put client grids in `fixtures/private/` and ignore them from Git unless the repository is confirmed private and explicit approval is given.
- Source-control exported VBA modules, not only the binary workbook.
- The binary controller workbook may be included as a release artifact after testing.
- Do not push until local testing and user acceptance are complete.

## 11. Implementation standards

- Use `Option Explicit` in every module.
- Fully qualify every `Workbook`, `Worksheet`, and `Range` reference.
- Do not use `Select`, `Selection`, `Activate`, or `ActiveSheet` in production code.
- Do not use whole-sheet `UsedRange` to locate the schedule.
- Do not hard-code week dates or sheet names.
- Minimize hard-coded row numbers; detect headers, dates, schedule rows, and legend by content.
- Preserve and restore `ScreenUpdating`, `EnableEvents`, `Calculation`, and `DisplayAlerts` in success and failure paths.
- Use structured error handling and actionable messages.
- Make the transformation idempotent: rerunning it must not add a second CDMX column or repeat text changes.
- Avoid destructive row deletion.
- Never save over the source file.

## 12. Automated acceptance tests

### 12.1 Source safety

- Source file hash is unchanged.
- Source workbook was opened read-only and closed without saving.
- Output file is a separate `.xlsx`.
- Output opens in desktop Excel without a repair warning.

### 12.2 Structure

- Exactly one title row, one header row, 96 quarter-hour schedule rows, and one legend area are detected.
- Final logical columns are GMT, BRT, CDMX, and seven date columns.
- No duplicate right-side BRT/GMT columns remain.
- No picture or shape remains on the output schedule sheet.
- Day-column merge start/end rows match the source after the one-column shift.

### 12.3 Time calculations

For all 96 rows:

- `CDMX = MOD(BRT - 3/24, 1)`.
- Values remain numeric Excel times.
- Midnight wrap is correct.
- Invalid GMT values do not affect CDMX.

### 12.4 Text rules

Include unit tests for:

- 1, 2, 3, 4, 5, and 6-line inputs;
- trailing blank lines;
- CRLF and LF inputs;
- a one-row unmerged cell;
- a two-row merged cell;
- an eight-row merged cell;
- accented Portuguese/Spanish characters;
- `<LIVE>`, `<Repeat>`, and `First Airing/Delayed` tags.

### 12.5 Formatting

Assert accepted values for:

- column widths;
- row heights;
- font names, sizes, and bold states;
- wrap text;
- horizontal and vertical alignment;
- number formats;
- time-column borders;
- CDMX fill;
- title fill and font;
- date-header format;
- preservation of program fills and merge borders.

### 12.6 Idempotence

Run the formatter against its own output and verify:

- no second CDMX column;
- no additional column deletion;
- no further text change;
- no merge change;
- no formatting drift.

### 12.7 Golden comparison

Once a matching before/after pair exists:

- compare values;
- compare merge maps;
- compare widths/heights;
- compare font/fill/border/alignment/number-format signatures;
- compare a PDF or print-preview capture for Monday and at least one later day;
- document every intentional difference.

## 13. Manual user acceptance

The first release is accepted only after the operator can:

1. run the controller against a new weekly SportyNet source;
2. open the formatted output;
3. print GMT+BRT+CDMX+Monday using the current manual selection method;
4. remove or hide Monday manually and repeat through Sunday;
5. confirm that the printed result matches the existing manual result.

Do not proceed to optimization merely because automated checks pass; the operator print comparison is mandatory.

## 14. Phase 2, only after parity is approved

Recommended next improvement:

- keep a non-destructive master week;
- add a day selector for Monday through Sunday;
- hide all nonselected day columns;
- set the print area through VBA;
- provide `Print Preview Current Day`;
- provide `Print/Export All Seven Days` with state restoration;
- never delete day columns.

After that works reliably, evaluate:

- an `.xlam` Excel add-in;
- a PowerShell/Python Excel-COM utility;
- a local desktop UI.

Do not choose a local app merely because it is newer. Move away from VBA only when the accepted workbook behavior is fully captured by tests and there is a clear operational benefit.

---

# Prompt A — GLM plan-only session

Copy this prompt into GLM-5.2 with effort set to max.

```text
You are working locally from my existing GitHub repository on my Windows desktop. The remote connection already works and you can see the repository.

Your task in this session is PLAN AND FORENSICS ONLY. Do not implement the formatter, do not modify production source code, do not modify the supplied Excel files, do not push, and do not create a release workbook.

Workflow:
1. Verify the repository remote and working-tree status.
2. Fetch remote changes and create or propose a feature branch named feature/sportynet-excel-parity. Do not push it.
3. Locate these six Excel files in the workspace or in the private fixture directory:
   - Canada Aug 01 to Aug 07.xlsx
   - English - 08-01 to 08-07.xlsx
   - Spanish 08-01 to 08-07.xlsx
   - Template - Grid Format.xltm
   - WEEK 31 V.4.xlsx
   - SNETL Transformat.xlsx
4. Inspect the workbooks using desktop Excel/COM and OOXML inspection. Extract and document the VBA from the xltm. Do not rely only on a Python spreadsheet library.
5. Confirm or correct every finding in GLM_SportyNet_Excel_Parity_Plan.md.
6. Create these documents only:
   - docs/SPORTYNET_EXCEL_PARITY_PLAN.md
   - docs/SPORTYNET_REQUIREMENTS_MATRIX.md
   - docs/SPORTYNET_PARITY_DECISIONS.md
   - docs/CURRENT_BEIN_MACRO_AUDIT.md
7. In the decisions document, explicitly list every conflict between the written instructions and SNETL Transformat.xlsx. Do not silently pick values.
8. Propose the exact VBA module design, Excel-COM test design, build approach for SportyNet_Formatter.xlsm, risks, rollback plan, and acceptance gates.
9. Confirm that the Week 31 source and Week 29 formatted file are not a matching golden pair and state exactly what matching fixture is still required.
10. End with a concise implementation checklist and the exact files you would create or change after approval.

Hard constraints:
- Excel/VBA first.
- Exact current-manual parity before optimization.
- Source files must remain unchanged.
- No Select, Selection, Activate, or ActiveSheet in production VBA.
- No UsedRange-based layout detection.
- No destructive blank-row deletion.
- No printing automation in Phase 1.
- No push until I approve the completed implementation.
- Use GLM-5.2 at max effort for this entire session.

Stop after producing and summarizing the plan. Wait for the exact message: APPROVED TO IMPLEMENT.
```

# Prompt B — implementation session after approval

Use only after reviewing the plan and sending `APPROVED TO IMPLEMENT`.

```text
APPROVED TO IMPLEMENT.

Implement the accepted SportyNet Excel parity plan on the feature/sportynet-excel-parity branch.

Use GLM-5.2 at max effort. Work locally and do not push until I explicitly approve the finished result.

Required deliverables:
- source-controlled VBA modules;
- a working SportyNet_Formatter.xlsm controller workbook;
- a separate formatted .xlsx output generated from WEEK 31 V.4.xlsx;
- PowerShell/Excel-COM integration tests;
- VBA unit tests for the line-break transformation;
- workbook snapshots and a test report;
- updated documentation and parity decisions;
- no changes to the existing beIN template.

Implementation requirements:
- open source read-only and never overwrite it;
- detect the layout dynamically;
- create a clean output without the source workbook's Z1000 style pollution;
- final logical columns: GMT, BRT, CDMX, Monday through Sunday;
- calculate CDMX from BRT minus three hours with midnight wrap;
- preserve day merges, fills, borders, and category semantics;
- apply the accepted line-break rules exactly;
- apply accepted widths, heights, fonts, alignment, borders, title styling, and legend placement;
- remove shapes when present;
- save as <source>_FORMATTED.xlsx;
- make the transform idempotent;
- fail safely with actionable messages;
- run every automated test in desktop Excel;
- compare against the matching golden pair when it is available;
- do not add Phase-2 printing automation.

Before declaring completion:
1. Run the unit tests.
2. Run the Excel-COM integration suite.
3. Open the generated workbook in Excel and verify there is no repair warning.
4. Produce a detailed parity report with any remaining visual or structural differences.
5. Review the full git diff and remove unrelated changes.
6. Stop and wait for my user-acceptance result. Do not push.
```

# Prompt C — fresh final audit

```text
Perform an independent final audit of the SportyNet Excel parity implementation. Use GLM-5.2 at max effort. Do not trust prior conclusions.

Read the accepted requirements, inspect the source-controlled VBA, run the unit tests and Excel-COM integration tests, regenerate the Week 31 output from a clean source, and compare it with the approved golden fixture.

Look specifically for:
- source-file modification;
- hard-coded sheet/week/date assumptions;
- UsedRange errors caused by style pollution;
- merge loss;
- incorrect midnight handling;
- line-break indexing mistakes;
- duplicate CDMX insertion on rerun;
- formatting drift;
- shape/logo remnants;
- title or legend discrepancies;
- Excel repair warnings;
- print-visible differences for Monday and a later day;
- unrelated git changes.

Return PASS only when every acceptance criterion is evidenced. Otherwise return FAIL with exact file, module, procedure, cell/range, and reproduction steps. Do not push.
```
