# Distribution Guide

Run `AI/Source/Build Distributions.bat` to recreate clean update packages in `AI/Packages`.

- **Full**: runtime + Settings + AI project files.
- **Manager**: runtime + Settings; no AI folder.
- **Operator**: runtime only; no Settings or AI folder.

The packages intentionally exclude the live Excel grid workbooks. Copy the update files into the coworkers' working grid folder. The active `FormatGrids.ps1` already contains the settings saved by the manager.

Before replacing production, run `AI/Tests/Run Windows Smoke Tests.bat` on a Windows computer with desktop Excel. The smoke test works only on disposable temporary copies.
