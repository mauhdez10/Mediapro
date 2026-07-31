# Grid Formatter V2 Settings

1. Double-click **Open Settings.bat**.
2. Change language, UTC/GMT offset and label, printer color setup, or tab selection.
3. Use **X** to format all eligible tabs, or enter a number and choose First/Last.
4. Click **Save Settings**.

Saving updates only the protected settings block in `../FormatGrids.ps1`. The formatter never reads `settings.json` while processing grids. Deleting the entire Settings folder does not stop the active formatter.

`printer_color_set.txt` stores only the last day the color setup ran. Operator packages without Settings use a local Windows state file automatically.
