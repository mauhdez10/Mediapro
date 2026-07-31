# Settings Specification

The Settings UI edits only the managed block in `FormatGrids.ps1`, creates a backup, validates the resulting PowerShell syntax using the Windows PowerShell parser, and then replaces the active script. `settings.json` is UI state only.

Tab values accept `X` for all eligible sheets or an integer with First/Last. REV TV first filters sheets where `A1=Zone`. SportyNet first filters sheets using the strict SportyNet structure.

The global UTC setting applies to existing grid types that generate UTC columns. SportyNet uses the source GMT/BRT values and the fixed three-hour CDMX relationship.

When the optional AI folder exists, a successful Settings save also refreshes `AI/Source/FormatGrids_V2.ps1`. Failure to update the optional source snapshot never blocks the active runtime script.
