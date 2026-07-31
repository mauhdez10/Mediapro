# Grids V2 Project Overview

Grids V2 is a development copy of the current `Grids` formatter. Production `Grids` remains untouched until V2 passes Windows/Excel regression testing.

V2 preserves all existing filename routes and formatting logic, then adds SportyNet routing in this order:

1. Existing filename types: CATV, TVD, Pasiones, Todo Novelas, REV TV.
2. Known SportyNet names, including `SNETL - Week 31` and `WEEK 31 V.4 1`.
3. Only still-unmatched Excel files receive a lightweight SportyNet check on row 2 of the first worksheet. A positive match then triggers the full strict validation across eligible worksheets.
4. Other files are skipped as unknown.

The runtime updates workbooks in place and creates no PDF or day-output files.
