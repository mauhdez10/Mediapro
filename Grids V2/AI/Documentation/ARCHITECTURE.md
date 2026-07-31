# Architecture

## Runtime root

- `1.Run Grid Script.bat`
- `FormatGrids.ps1`
- Excel grids to process

## Optional Settings

The UI writes a managed block into `FormatGrids.ps1`. Runtime does not read JSON. The Settings folder may be deleted without breaking formatting.

## Optional AI

Documentation, references, sources, tests, logs, trash, and distribution packages. Runtime behavior does not require this folder.

## Distribution profiles

- Full: runtime + Settings + AI.
- Manager: runtime + Settings, no AI.
- Operator: runtime only, no Settings or AI.
