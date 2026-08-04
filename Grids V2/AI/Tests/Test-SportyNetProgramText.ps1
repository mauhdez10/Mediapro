$ErrorActionPreference='Stop'
$projectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scriptPath = Join-Path $projectRoot 'FormatGrids.ps1'
$env:GRID_FORMATTER_FUNCTIONS_ONLY='1'
try { . $scriptPath }
finally { Remove-Item Env:GRID_FORMATTER_FUNCTIONS_ONLY -ErrorAction SilentlyContinue }

$tests = @(
    @{
        Name='WEEK 30 D73 internal blank target'
        Text="AMISTOSO INTERNACIONAL`n2026`n`nBenfica x Villareal`n17/07`nvt`n"
        Span=8
        Expected="AMISTOSO INTERNACIONAL`n2026`nBenfica x Villareal`n17/07 vt"
    },
    @{
        Name='Standard six-line merged program'
        Text="PROG00000001`nCAMP. BRASILEIRO`n2026`n14ª RODADA`nTeam A x Team B`n<Repeat>`n"
        Span=4
        Expected="PROG00000001`nCAMP. BRASILEIRO`n2026 14ª RODADA`nTeam A x Team B <Repeat>"
    },
    @{
        Name='Three-line merged program'
        Text="PROG00000002`nESPECIAL COPAS`n<Repeat>"
        Span=2
        Expected="PROG00000002`nESPECIAL COPAS <Repeat>"
    },
    @{
        Name='One-row program ignores blank placeholder'
        Text="PROGRAM`n`nMATCH`n<LIVE>"
        Span=1
        Expected='PROGRAM MATCH <LIVE>'
    }
)

$passed=0
foreach ($test in $tests) {
    $actual=Convert-SportyNetProgramText $test.Text $test.Span $test.Name
    if ($actual -ne $test.Expected) {
        throw "FAIL $($test.Name)`nExpected: [$($test.Expected)]`nActual:   [$actual]"
    }
    Write-Host "PASS: $($test.Name)" -ForegroundColor Green
    $passed++
}
Write-Host "PASS: $passed/$($tests.Count) SportyNet program-text tests" -ForegroundColor Green
