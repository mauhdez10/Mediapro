# Grids V2 Settings

Double-click `Open Settings.bat` to open the Settings window without leaving a PowerShell console visible.

The Settings UI changes only:

- PowerShell display language: English or Spanish
- UTC/GMT offset and label
- Automatic printer color setup and printer name
- How many matching worksheets each grid type formats: `X` for all, or a number using First/Last

Formatting dimensions are fixed inside each grid formatter and are intentionally not editable in this simplified V1 interface.

When you click **Save Settings**, the UI:

1. Creates a backup of `FormatGrids.ps1`.
2. Replaces only the protected managed-settings block.
3. validates the updated PowerShell script.
4. Activates it only if validation succeeds.

`FormatGrids.ps1` does not read `settings.json` during normal formatting. Removing the entire Settings folder does not prevent the formatter from running.
