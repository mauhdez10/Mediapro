param([string]$Root = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent))
$ErrorActionPreference='Stop'
$packages=Join-Path $Root 'AI\Packages'; [void](New-Item -ItemType Directory -Path $packages -Force)
Add-Type -AssemblyName System.IO.Compression.FileSystem

function New-Package([string]$Name,[bool]$IncludeSettings,[bool]$IncludeAI) {
    $stage=Join-Path $env:TEMP ("GridsV2_"+[guid]::NewGuid().ToString('N'))
    $packageRoot=Join-Path $stage 'Grids V2'
    [void](New-Item -ItemType Directory -Path $packageRoot -Force)
    try {
        Copy-Item (Join-Path $Root '1.Run Grid Script.bat') $packageRoot
        Copy-Item (Join-Path $Root 'FormatGrids.ps1') $packageRoot
        if ($IncludeSettings) { Copy-Item (Join-Path $Root 'Settings') (Join-Path $packageRoot 'Settings') -Recurse }
        if ($IncludeAI) {
            Copy-Item (Join-Path $Root 'AI') (Join-Path $packageRoot 'AI') -Recurse
            Remove-Item (Join-Path $packageRoot 'AI\Packages\*') -Force -ErrorAction SilentlyContinue
            Remove-Item (Join-Path $packageRoot 'AI\Trash\*') -Force -Recurse -ErrorAction SilentlyContinue
        }
        $zip=Join-Path $packages $Name
        if (Test-Path $zip) { Remove-Item $zip -Force }
        [IO.Compression.ZipFile]::CreateFromDirectory($stage,$zip,[IO.Compression.CompressionLevel]::Optimal,$false)
        Write-Host "Created $zip"
    } finally { Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue }
}

New-Package 'Grids_V2_Full.zip' $true $true
New-Package 'Grids_V2_Manager.zip' $true $false
New-Package 'Grids_V2_Operator.zip' $false $false
