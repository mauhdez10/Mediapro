# Settings Specification — Grids V2 2.0.4-dev

The manager Settings UI exposes only the settings requested for the V1 workflow:

- PowerShell display language
- UTC/GMT offset and label
- printer color automation and printer name
- worksheet selection per grid type (`X`, First N, Last N)

The UI starts in English and has an `Español`/`English` switch that changes only the Settings-window language. The separate PowerShell display-language selection controls the language embedded in `FormatGrids.ps1`.

Formatting dimensions are not user settings in the simplified V1 UI. Each grid formatter keeps its proven format-specific values. This avoids presenting generic values that do not describe Pasiones/Todo or SportyNet accurately.

Saving settings creates a backup, replaces only the managed block in `FormatGrids.ps1`, validates PowerShell syntax, and then activates the change. The runtime never reads `settings.json`.

`Open Settings.bat` launches `Open Settings.vbs`, which starts PowerShell hidden. Startup failures are written to `AI/Logs/settings_ui_errors.log.txt` when AI exists, otherwise to the Settings folder.
