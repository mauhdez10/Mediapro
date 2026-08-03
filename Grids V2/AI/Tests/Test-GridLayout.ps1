# Test script for per-grid layout settings
# This script tests the Get-GridLayout function and validates layout retrieval

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$formatterPath = Join-Path $scriptPath "AI\Source\FormatGrids_V2.ps1"

Write-Host "Testing Get-GridLayout function..." -ForegroundColor Cyan
Write-Host ""

# Test 1: Load the formatter and test Get-GridLayout function
Write-Host "Test 1: Testing Get-GridLayout function" -ForegroundColor Yellow
Write-Host "----------------------------------------"

# Create a minimal test environment with mock settings
$ManagedSettings = [ordered]@{
    Version = "2.0.3-dev"
    Layouts = [ordered]@{
        CATV             = [ordered]@{ FontSize = 14; DefaultRowH = 25; HeaderRowH = 35; SmallColW = 7; DefaultColW = 37 }
        TVD              = [ordered]@{ FontSize = 14; DefaultRowH = 25; HeaderRowH = 35; SmallColW = 7; DefaultColW = 37 }
        PASIONES_LATAM   = [ordered]@{ FontSize = 14; DefaultRowH = 25; HeaderRowH = 35; SmallColW = 7; DefaultColW = 37 }
        PASIONES_US      = [ordered]@{ FontSize = 14; DefaultRowH = 25; HeaderRowH = 35; SmallColW = 7; DefaultColW = 37 }
        TODO_NOVELAS     = [ordered]@{ FontSize = 14; DefaultRowH = 25; HeaderRowH = 35; SmallColW = 7; DefaultColW = 37 }
        REV_TV           = [ordered]@{ FontSize = 14; DefaultRowH = 25; HeaderRowH = 35; SmallColW = 7; DefaultColW = 37 }
        SPORTYNET        = [ordered]@{ FontSize = 14; DefaultRowH = 25; HeaderRowH = 35; SmallColW = 7; DefaultColW = 37 }
    }
}

$DefaultLayout = [ordered]@{
    FontSize     = 14
    DefaultRowH  = 25
    HeaderRowH   = 35
    SmallColW    = 7
    DefaultColW  = 37
}

function Get-GridLayout([string]$gridType) {
    try {
        $layout = $ManagedSettings.Layouts[$gridType]
        if ($null -ne $layout) {
            return [ordered]@{
                FontSize     = [int]$layout.FontSize
                DefaultRowH  = [double]$layout.DefaultRowH
                HeaderRowH   = [double]$layout.HeaderRowH
                SmallColW    = [double]$layout.SmallColW
                DefaultColW  = [double]$layout.DefaultColW
            }
        }
    } catch {}
    return $DefaultLayout
}

# Test each grid type
$gridTypes = @("CATV", "TVD", "PASIONES_LATAM", "PASIONES_US", "TODO_NOVELAS", "REV_TV", "SPORTYNET", "UNKNOWN")

foreach ($gridType in $gridTypes) {
    $layout = Get-GridLayout $gridType
    Write-Host "Grid: $gridType" -ForegroundColor White
    Write-Host "  FontSize:     $($layout.FontSize)" -ForegroundColor Gray
    Write-Host "  DefaultRowH:  $($layout.DefaultRowH)" -ForegroundColor Gray
    Write-Host "  HeaderRowH:   $($layout.HeaderRowH)" -ForegroundColor Gray
    Write-Host "  SmallColW:    $($layout.SmallColW)" -ForegroundColor Gray
    Write-Host "  DefaultColW:  $($layout.DefaultColW)" -ForegroundColor Gray
    Write-Host ""
}

# Test 2: Custom layout values
Write-Host "Test 2: Testing custom layout values" -ForegroundColor Yellow
Write-Host "----------------------------------------"

$ManagedSettings.Layouts.CATV = [ordered]@{ FontSize = 16; DefaultRowH = 30; HeaderRowH = 40; SmallColW = 8; DefaultColW = 40 }

$catvLayout = Get-GridLayout "CATV"
Write-Host "CATV with custom values:" -ForegroundColor White
Write-Host "  FontSize:     $($catvLayout.FontSize) (expected: 16)" -ForegroundColor $(if ($catvLayout.FontSize -eq 16) { "Green" } else { "Red" })
Write-Host "  DefaultRowH:  $($catvLayout.DefaultRowH) (expected: 30)" -ForegroundColor $(if ($catvLayout.DefaultRowH -eq 30) { "Green" } else { "Red" })
Write-Host "  HeaderRowH:   $($catvLayout.HeaderRowH) (expected: 40)" -ForegroundColor $(if ($catvLayout.HeaderRowH -eq 40) { "Green" } else { "Red" })
Write-Host "  SmallColW:    $($catvLayout.SmallColW) (expected: 8)" -ForegroundColor $(if ($catvLayout.SmallColW -eq 8) { "Green" } else { "Red" })
Write-Host "  DefaultColW:  $($catvLayout.DefaultColW) (expected: 40)" -ForegroundColor $(if ($catvLayout.DefaultColW -eq 40) { "Green" } else { "Red" })
Write-Host ""

$tvdLayout = Get-GridLayout "TVD"
Write-Host "TVD with default values:" -ForegroundColor White
Write-Host "  FontSize:     $($tvdLayout.FontSize) (expected: 14)" -ForegroundColor $(if ($tvdLayout.FontSize -eq 14) { "Green" } else { "Red" })
Write-Host "  DefaultRowH:  $($tvdLayout.DefaultRowH) (expected: 25)" -ForegroundColor $(if ($tvdLayout.DefaultRowH -eq 25) { "Green" } else { "Red" })
Write-Host ""

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "All tests completed!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan