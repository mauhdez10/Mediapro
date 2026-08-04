$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scriptPath = Join-Path $projectRoot 'FormatGrids.ps1'
if (-not (Test-Path -LiteralPath $scriptPath)) { throw "Missing active script: $scriptPath" }

$tokens = $null; $parseErrors = $null
[System.Management.Automation.Language.Parser]::ParseFile($scriptPath,[ref]$tokens,[ref]$parseErrors) | Out-Null
if ($parseErrors.Count -gt 0) { throw ("PowerShell parse error: " + $parseErrors[0].Message) }
$content = Get-Content -LiteralPath $scriptPath -Raw
if ($content -match '\.Calculation\s*=') { throw 'Forbidden Excel.Application.Calculation assignment found.' }
if ($content -match '(?im)^\s*(Get-Content|Import-Clixml).*settings\.json') { throw 'Runtime dependency on settings.json found.' }

$tempRoot = Join-Path $env:TEMP ("GridsV2_Smoke_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot | Out-Null
$preserveTemp = $true
try {
    Copy-Item -LiteralPath $scriptPath -Destination (Join-Path $tempRoot 'FormatGrids.ps1')

    $fixtureMap = [ordered]@{
        'GRILLA CATV MASTER SMOKE.xlsx' = Join-Path $projectRoot 'GRILLA CATV MASTER SEMANA DEL 03 AL 09 DE AGOSTO 2026.xlsx'
        'GRILLA TVD MASTER SMOKE.xlsx' = Join-Path $projectRoot 'GRILLA TVD MASTER SEMANA DEL 03 AL 09 DE AGOSTO 2026.xlsx'
        'GRILLA MASTER PASIONES LATAM SMOKE.xlsx' = Join-Path $projectRoot 'GRILLA MASTER PASIONES LATAM 2026.xlsx'
        'TODO NOVELAS PROGRAMMING GRID SMOKE.xlsx' = Join-Path $projectRoot 'TODO NOVELAS  PROGRAMMING GRID 2026 7.13.xlsx'
        'SNETL - Week 31.xlsx' = Join-Path $projectRoot 'AI\References\SportyNet\SNETL - Week 31.xlsx'
        'WEEK 31 V.4 1.xlsx' = Join-Path $projectRoot 'AI\References\SportyNet\WEEK 31 V.4 1.xlsx'
        'Week 32.xlsx' = Join-Path $projectRoot 'AI\References\SportyNet\Week 32.xlsx'
        'WEEK 30 V2.xlsx' = Join-Path $projectRoot 'AI\References\SportyNet\WEEK 30 V2.xlsx'
        'Mystery Grid.xlsx' = Join-Path $projectRoot 'AI\References\SportyNet\Week 32.xlsx'
    }
    foreach ($entry in $fixtureMap.GetEnumerator()) {
        if (-not (Test-Path -LiteralPath $entry.Value)) { throw "Missing fixture: $($entry.Value)" }
        Copy-Item -LiteralPath $entry.Value -Destination (Join-Path $tempRoot $entry.Key)
    }

    # Create one unrelated workbook. It must be skipped and left byte-for-byte unchanged.
    $unknownPath = Join-Path $tempRoot 'Unrelated Workbook.xlsx'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    try {
        $unknownBook = $excel.Workbooks.Add()
        $unknownBook.Worksheets.Item(1).Cells.Item(1,1).Value2 = 'Not a supported grid'
        $unknownBook.SaveAs($unknownPath,51)
        $unknownBook.Close($false)
    } finally {
        try { $excel.Quit() } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    $unknownHashBefore = (Get-FileHash -LiteralPath $unknownPath -Algorithm SHA256).Hash

    $env:GRID_FORMATTER_NO_PAUSE = '1'
    $env:GRID_FORMATTER_SKIP_PRINTER = '1'
    $runner = Join-Path $tempRoot 'FormatGrids.ps1'
    $output = & "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $runner 2>&1 | Out-String
    $exitCode = $LASTEXITCODE
    Remove-Item Env:GRID_FORMATTER_NO_PAUSE -ErrorAction SilentlyContinue
    Remove-Item Env:GRID_FORMATTER_SKIP_PRINTER -ErrorAction SilentlyContinue

    Write-Host $output
    if ($exitCode -ne 0) { throw "Formatter exited with code $exitCode" }
    if ($output -match '(?im)^\s*Cannot process:' -or $output -match '(?im)^\s*No se puede procesar:') {
        throw 'Formatter reported a workbook processing failure.'
    }
    if ($output -match '(?im)^ERRORS\s*\(([1-9]\d*)\):' -or $output -match '(?im)^ERRORES\s*\(([1-9]\d*)\):') {
        throw 'Formatter final report contains errors.'
    }
    if ($output -match '0x800A03EC' -or $output -match 'Exception setting "Calculation"') {
        throw 'Calculation regression detected.'
    }

    $unknownHashAfter = (Get-FileHash -LiteralPath $unknownPath -Algorithm SHA256).Hash
    if ($unknownHashBefore -ne $unknownHashAfter) { throw 'Unrelated workbook was modified.' }

    # Verify all three known Sporty name styles and the unmatched soft-check path.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    try {
        foreach ($name in @('SNETL - Week 31.xlsx','WEEK 31 V.4 1.xlsx','Week 32.xlsx','WEEK 30 V2.xlsx','Mystery Grid.xlsx')) {
            $path = Join-Path $tempRoot $name
            $book = $excel.Workbooks.Open($path,$null,$true)
            try {
                $sheet = $book.Worksheets.Item(1)
                if (([string]$sheet.Cells.Item(2,3).Value2).Trim() -ne 'CDMX') { throw "$name was not formatted as SportyNet." }
                if ([Math]::Abs(([double]$sheet.Cells.Item(18,1).Value2) - (6.75/24.0)) -gt 0.0000001) { throw "$name A18 was not normalized to 06:45." }
                if ([Math]::Abs(([double]$sheet.Cells.Item(30,2).Value2) - (6.75/24.0)) -gt 0.0000001) { throw "$name B30 was not normalized to 06:45." }
                if ([Math]::Abs(([double]$sheet.Cells.Item(30,3).Value2) - (3.75/24.0)) -gt 0.0000001) { throw "$name C30 was not normalized to 03:45 CDMX." }
                if ($name -eq 'WEEK 30 V2.xlsx') {
                    $expectedProgram = "AMISTOSO INTERNACIONAL`n2026`nBenfica x Villareal`n17/07 vt"
                    $actualProgram = [string]$sheet.Cells.Item(73,4).Value2
                    if ($actualProgram -ne $expectedProgram) { throw "$name D73 internal-blank normalization failed. Actual: $actualProgram" }
                }
            } finally { $book.Close($false) }
        }
    } finally {
        try { $excel.Quit() } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    $settingsUi = Join-Path $projectRoot 'Settings\Settings UI.ps1'
    if (-not (Test-Path -LiteralPath $settingsUi)) { throw "Missing Settings UI: $settingsUi" }
    $env:GRID_SETTINGS_UI_SELFTEST='1'
    $uiOutput = & "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -STA -File $settingsUi 2>&1 | Out-String
    $uiExitCode=$LASTEXITCODE
    Remove-Item Env:GRID_SETTINGS_UI_SELFTEST -ErrorAction SilentlyContinue
    if ($uiExitCode -ne 0 -or $uiOutput -notmatch 'SETTINGS_UI_SELFTEST_PASS') { throw "Settings UI self-test failed: $uiOutput" }

    Write-Host 'PASS: PowerShell syntax' -ForegroundColor Green
    Write-Host 'PASS: No Calculation assignment' -ForegroundColor Green
    Write-Host 'PASS: Runtime works without Settings or AI folders' -ForegroundColor Green
    Write-Host 'PASS: Existing grid routes completed without reported errors' -ForegroundColor Green
    Write-Host 'PASS: Real SNETL, Mex-left, Week 32, Week 30 internal-blank, and soft-fallback SportyNet files formatted correctly' -ForegroundColor Green
    Write-Host 'PASS: Settings UI opens with English buttons and switches to Spanish' -ForegroundColor Green
    Write-Host 'PASS: Unrelated workbook was skipped and unchanged' -ForegroundColor Green
    $preserveTemp = $false
} finally {
    Remove-Item Env:GRID_FORMATTER_NO_PAUSE -ErrorAction SilentlyContinue
    Remove-Item Env:GRID_FORMATTER_SKIP_PRINTER -ErrorAction SilentlyContinue
    Remove-Item Env:GRID_SETTINGS_UI_SELFTEST -ErrorAction SilentlyContinue
    if ($preserveTemp) {
        Write-Host "Smoke-test files preserved for investigation: $tempRoot" -ForegroundColor Yellow
    } else {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
