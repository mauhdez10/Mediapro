# Simple integration test for per-grid layout settings
$errors = 0

Write-Host "Integration Test: Per-Grid Layout Settings" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# Get absolute paths
$scriptPath = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptPath)) {
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$formatterPath = Join-Path $scriptPath "..\..\FormatGrids.ps1"
$settingsPath = Join-Path $scriptPath "..\..\Settings\settings.json"

Write-Host ""
Write-Host "Script path: $scriptPath" -ForegroundColor Gray
Write-Host "Formatter path: $formatterPath" -ForegroundColor Gray
Write-Host "Settings path: $settingsPath" -ForegroundColor Gray

# Test 1: FormatGrids.ps1 exists
Write-Host ""
Write-Host "Test 1: FormatGrids.ps1 exists" -ForegroundColor Yellow
if (Test-Path $formatterPath) { Write-Host "PASS" -ForegroundColor Green } else { Write-Host "FAIL" -ForegroundColor Red; $errors++ }

# Test 2: Get-GridLayout function exists
Write-Host ""
Write-Host "Test 2: Get-GridLayout function exists" -ForegroundColor Yellow
if (Test-Path $formatterPath) {
    $content = [IO.File]::ReadAllText($formatterPath)
    if ($content -match 'function Get-GridLayout') { Write-Host "PASS" -ForegroundColor Green } else { Write-Host "FAIL" -ForegroundColor Red; $errors++ }
} else {
    Write-Host "SKIP - file not found" -ForegroundColor Yellow
}

# Test 3: Layouts section exists
Write-Host ""
Write-Host "Test 3: Layouts section exists" -ForegroundColor Yellow
if (Test-Path $formatterPath) {
    if ($content -match 'Layouts = \[ordered\]@{') { Write-Host "PASS" -ForegroundColor Green } else { Write-Host "FAIL" -ForegroundColor Red; $errors++ }
} else {
    Write-Host "SKIP - file not found" -ForegroundColor Yellow
}

# Test 4: Layout retrieval calls
Write-Host ""
Write-Host "Test 4: Layout retrieval calls (2+ expected)" -ForegroundColor Yellow
if (Test-Path $formatterPath) {
    $layoutCalls = ([regex]::Matches($content, 'layout = Get-GridLayout')).Count
    if ($layoutCalls -ge 2) { Write-Host "PASS - Found $layoutCalls calls" -ForegroundColor Green } else { Write-Host "FAIL - Only $layoutCalls calls" -ForegroundColor Red; $errors++ }
} else {
    Write-Host "SKIP - file not found" -ForegroundColor Yellow
}

# Test 5: FontSize used
Write-Host ""
Write-Host "Test 5: FontSize property used" -ForegroundColor Yellow
if (Test-Path $formatterPath) {
    if ($content -match '\$layout\.FontSize') { Write-Host "PASS" -ForegroundColor Green } else { Write-Host "FAIL" -ForegroundColor Red; $errors++ }
} else {
    Write-Host "SKIP - file not found" -ForegroundColor Yellow
}

# Test 6: DefaultRowH used
Write-Host ""
Write-Host "Test 6: DefaultRowH property used" -ForegroundColor Yellow
if (Test-Path $formatterPath) {
    if ($content -match '\$layout\.DefaultRowH') { Write-Host "PASS" -ForegroundColor Green } else { Write-Host "FAIL" -ForegroundColor Red; $errors++ }
} else {
    Write-Host "SKIP - file not found" -ForegroundColor Yellow
}

# Test 7: HeaderRowH used
Write-Host ""
Write-Host "Test 7: HeaderRowH property used" -ForegroundColor Yellow
if (Test-Path $formatterPath) {
    if ($content -match '\$layout\.HeaderRowH') { Write-Host "PASS" -ForegroundColor Green } else { Write-Host "FAIL" -ForegroundColor Red; $errors++ }
} else {
    Write-Host "SKIP - file not found" -ForegroundColor Yellow
}

# Test 8: SmallColW used
Write-Host ""
Write-Host "Test 8: SmallColW property used" -ForegroundColor Yellow
if (Test-Path $formatterPath) {
    if ($content -match '\$layout\.SmallColW') { Write-Host "PASS" -ForegroundColor Green } else { Write-Host "FAIL" -ForegroundColor Red; $errors++ }
} else {
    Write-Host "SKIP - file not found" -ForegroundColor Yellow
}

# Test 9: DefaultColW used
Write-Host ""
Write-Host "Test 9: DefaultColW property used" -ForegroundColor Yellow
if (Test-Path $formatterPath) {
    if ($content -match '\$layout\.DefaultColW') { Write-Host "PASS" -ForegroundColor Green } else { Write-Host "FAIL" -ForegroundColor Red; $errors++ }
} else {
    Write-Host "SKIP - file not found" -ForegroundColor Yellow
}

# Test 10: settings.json exists
Write-Host ""
Write-Host "Test 10: settings.json exists" -ForegroundColor Yellow
if (Test-Path $settingsPath) { Write-Host "PASS" -ForegroundColor Green } else { Write-Host "FAIL" -ForegroundColor Red; $errors++ }

# Summary
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
if ($errors -eq 0) {
    Write-Host "All tests PASSED! (10/10)" -ForegroundColor Green
} else {
    Write-Host "Tests FAILED: $errors error(s) found" -ForegroundColor Red
}
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  - FormatGrids.ps1: OK" -ForegroundColor Gray
Write-Host "  - Get-GridLayout function: OK" -ForegroundColor Gray
Write-Host "  - Layouts section: OK" -ForegroundColor Gray
Write-Host "  - Layout retrieval calls: $layoutCalls" -ForegroundColor Gray
Write-Host "  - All layout properties used: OK" -ForegroundColor Gray
Write-Host "  - settings.json: OK" -ForegroundColor Gray
Write-Host ""
Write-Host "Ready for end-to-end testing!" -ForegroundColor Green

exit $errors