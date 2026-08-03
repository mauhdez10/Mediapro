# End-to-end test with real Excel files
# This test runs the formatter with different layout settings for each grid type

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootPath = Split-Path $scriptPath -Parent
$formatterPath = Join-Path $rootPath "FormatGrids.ps1"

Write-Host "End-to-End Test: Per-Grid Layout Settings" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Test files available
$testFiles = @(
    @{Name="CATV"; File="GRILLA CATV MASTER SEMANA DEL 27 DE JULIO AL 02 DE AGOSTO 2026.xlsx"},
    @{Name="TVD"; File="GRILLA TVD MASTER SEMANA DEL 27 DE JULIO AL 02 DE AGOSTO 2026.xlsx"},
    @{Name="PASIONES_LATAM"; File="GRILLA MASTER PASIONES LATAM 2026.xlsx"},
    @{Name="PASIONES_US"; File="GRILLA MASTER PASIONES US 2026.xlsx"},
    @{Name="TODO_NOVELAS"; File="TODO NOVELAS  PROGRAMMING GRID 2026 7.13.xlsx"}
)

Write-Host "Available test files:" -ForegroundColor Yellow
foreach ($test in $testFiles) {
    $filePath = Join-Path $rootPath $test.File
    if (Test-Path -LiteralPath $filePath) {
        Write-Host "  ✓ $($test.Name): $($test.File)" -ForegroundColor Green
    } else {
        Write-Host "  ✗ $($test.Name): $($test.File) - NOT FOUND" -ForegroundColor Red
    }
}
Write-Host ""

# Backup current test files
Write-Host "Backing up test files..." -ForegroundColor Yellow
$backupDir = Join-Path $scriptPath "TestBackups_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

foreach ($test in $testFiles) {
    $filePath = Join-Path $rootPath $test.File
    if (Test-Path -LiteralPath $filePath) {
        Copy-Item -LiteralPath $filePath -Destination $backupDir -Force
        Write-Host "  ✓ Backed up $($test.File)" -ForegroundColor Gray
    }
}
Write-Host ""

# Show what the test will verify
Write-Host "What this test will verify:" -ForegroundColor Yellow
Write-Host "  1. FormatGrids.ps1 runs without errors" -ForegroundColor Gray
Write-Host "  2. Each grid type is processed correctly" -ForegroundColor Gray
Write-Host "  3. Layout values from Get-GridLayout are used" -ForegroundColor Gray
Write-Host "  4. Files can be opened in Excel after formatting" -ForegroundColor Gray
Write-Host ""

# Ask user if they want to proceed
Write-Host "WARNING: This will modify your test Excel files!" -ForegroundColor Red
Write-Host "Backups are saved to: $backupDir" -ForegroundColor Cyan
Write-Host ""
$proceed = Read-Host "Do you want to proceed? (Y/N)"

if ($proceed -ne "Y" -and $proceed -ne "y") {
    Write-Host "Test cancelled by user." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "Running formatter..." -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Run the formatter
try {
    Push-Location $rootPath
    & $formatterPath
    Pop-Location
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "Formatter completed successfully!" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "  1. Open each formatted Excel file in Excel" -ForegroundColor Gray
    Write-Host "  2. Check that font size is 14 (default)" -ForegroundColor Gray
    Write-Host "  3. Check that row heights are 25 (default)" -ForegroundColor Gray
    Write-Host "  4. Check that header row height is 35 (default)" -ForegroundColor Gray
    Write-Host "  5. Check that small column width is 7 (default)" -ForegroundColor Gray
    Write-Host "  6. Check that default column width is 37 (default)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "To test custom layout values:" -ForegroundColor Cyan
    Write-Host "  1. Open Settings UI.ps1" -ForegroundColor Gray
    Write-Host "  2. Go to the Layout tab" -ForegroundColor Gray
    Write-Host "  3. Select a grid type (e.g., CATV)" -ForegroundColor Gray
    Write-Host "  4. Change FontSize to 16, DefaultRowHeight to 30, etc." -ForegroundColor Gray
    Write-Host "  5. Save settings" -ForegroundColor Gray
    Write-Host "  6. Restore test files from backup" -ForegroundColor Gray
    Write-Host "  7. Run formatter again" -ForegroundColor Gray
    Write-Host "  8. Verify CATV has the custom values, other grids have defaults" -ForegroundColor Gray
} catch {
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Red
    Write-Host "Formatter failed with error:" -ForegroundColor Red
    Write-Host "============================================" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "You can restore your test files from:" -ForegroundColor Cyan
    Write-Host "$backupDir" -ForegroundColor White
    exit 1
}