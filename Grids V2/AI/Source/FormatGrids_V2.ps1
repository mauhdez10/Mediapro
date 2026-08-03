# ===========================================================
#  GRID FORMATTER V2 - FormatGrids.ps1
# ===========================================================
#  Runtime is self-contained. The Settings UI only rewrites the
#  protected block below; FormatGrids.ps1 never reads settings.json.
# ===========================================================

# >>> BEGIN MANAGED SETTINGS >>>
$ManagedSettings = [ordered]@{
    Version             = "2.0.3-dev"
    Language            = "English"
    UtcOffset           = -4
    UtcLabel            = "UTC"
    PrinterColorEnabled = $true
    PrinterName         = "MediaproUS_Canon_Online"
    Tabs = [ordered]@{
        CATV             = [ordered]@{ Mode = "First"; Count = 1 }
        TVD              = [ordered]@{ Mode = "First"; Count = 1 }
        PASIONES_LATAM   = [ordered]@{ Mode = "Last";  Count = 5 }
        PASIONES_US      = [ordered]@{ Mode = "Last";  Count = 5 }
        TODO_NOVELAS     = [ordered]@{ Mode = "Last";  Count = 5 }
        REV_TV           = [ordered]@{ Mode = "All";   Count = 0 }
        SPORTYNET        = [ordered]@{ Mode = "All";   Count = 0 }
    }
    Layouts = [ordered]@{
        CATV             = [ordered]@{ FontSize = 14; DefaultRowHeight = 25; HeaderRowHeight = 35; SmallColumnWidth = 7; DefaultColumnWidth = 37 }
        TVD              = [ordered]@{ FontSize = 14; DefaultRowHeight = 25; HeaderRowHeight = 35; SmallColumnWidth = 7; DefaultColumnWidth = 37 }
        PASIONES_LATAM   = [ordered]@{ FontSize = 14; DefaultRowHeight = 25; HeaderRowHeight = 35; SmallColumnWidth = 7; DefaultColumnWidth = 37 }
        PASIONES_US      = [ordered]@{ FontSize = 14; DefaultRowHeight = 25; HeaderRowHeight = 35; SmallColumnWidth = 7; DefaultColumnWidth = 37 }
        TODO_NOVELAS     = [ordered]@{ FontSize = 14; DefaultRowHeight = 25; HeaderRowHeight = 35; SmallColumnWidth = 7; DefaultColumnWidth = 37 }
        REV_TV           = [ordered]@{ FontSize = 14; DefaultRowHeight = 25; HeaderRowHeight = 35; SmallColumnWidth = 7; DefaultColumnWidth = 37 }
        SPORTYNET        = [ordered]@{ FontSize = 14; DefaultRowHeight = 25; HeaderRowHeight = 35; SmallColumnWidth = 7; DefaultColumnWidth = 37 }
    }
}
# <<< END MANAGED SETTINGS <<<

$UTC_OFFSET  = [int]$ManagedSettings.UtcOffset
$UTC_LABEL   = [string]$ManagedSettings.UtcLabel
$UTC_START_H = 6 + [Math]::Abs($UTC_OFFSET)
$FolderPath  = $PSScriptRoot
$Language    = if ([string]$ManagedSettings.Language -eq "Spanish") { "Spanish" } else { "English" }
$NoPause = $env:GRID_FORMATTER_NO_PAUSE -eq "1"
$SkipPrinterForTest = $env:GRID_FORMATTER_SKIP_PRINTER -eq "1"

$FontSize    = 14
$DefaultRowH = 25
$HeaderRowH  = 35
$SmallColW   = 7
$DefaultColW = 37
$PF_TimeColW = 8
$PF_CenterW  = 30
$DataRowEnd  = 50

# Default layout values (used if not overridden in settings)
$DefaultLayout = [ordered]@{
    FontSize     = $FontSize
    DefaultRowH  = $DefaultRowH
    HeaderRowH   = $HeaderRowH
    SmallColW    = $SmallColW
    DefaultColW  = $DefaultColW
}

# SportyNet constants. These are intentionally independent of the
# global UTC setting because the source grid already supplies GMT/BRT.
$SportyScheduleRows = 96
$SportyQuarterHour  = 15.0 / 1440.0
$SportyThreeHours   = 3.0 / 24.0
$SportyTimeTolerance = 0.0000001
$SportyHhmmTolerance = 0.005
$XL_CENTER = -4108
$XL_CONTINUOUS = 1
$XL_THIN = 2
$XL_EDGE_LEFT = 7
$XL_EDGE_TOP = 8
$XL_EDGE_BOTTOM = 9
$XL_EDGE_RIGHT = 10
$XL_INSIDE_VERTICAL = 11
$XL_INSIDE_HORIZONTAL = 12
$ColorBlack = 0
$ColorYellow = 65535
$ColorCdmx = 16639961
$ColorDateHeader = 8052479
$ColorTitle = 14146299

function Show-Text {
    param([string]$English, [string]$Spanish, $Color = "White", [switch]$NoNewline)
    $message = if ($Language -eq "Spanish") { $Spanish } else { $English }
    if ($NoNewline) { Write-Host $message -ForegroundColor $Color -NoNewline }
    else { Write-Host $message -ForegroundColor $Color }
}

function Get-Text {
    param([string]$English, [string]$Spanish)
    if ($Language -eq "Spanish") { return $Spanish }
    return $English
}

function Throw-Text {
    param([string]$English, [string]$Spanish)
    throw (Get-Text $English $Spanish)
}

# -----------------------------------------------------------
function ColLetter([int]$n) {
    $s = ""
    while ($n -gt 0) {
        $r = ($n - 1) % 26
        $s = [char](65 + $r) + $s
        $n = [int](($n - 1) / 26)
    }
    return $s
}

function CellStr($cell) {
    try {
        $v = $cell.Value2
        if ($null -eq $v) { return "" }
        if ($v -is [double] -or $v -is [int]) { return "" }
        return "$v".Trim()
    } catch { return "" }
}

# Strips ALL spaces, dots, dashes then uppercases for comparison
function NormTZ([string]$val) {
    return ($val -replace '[\s\.\-]','').ToUpper()
}

function IsTZToDelete([string]$val) {
    $v = NormTZ $val
    if ($v -eq "")    { return $false }
    if ($v -eq "ET")  { return $false }
    if ($v -eq "UTC") { return $false }
    $deleteList = @("PT","PST","PDT","EST","EDT","RD","CA","CT","MT","AT","CST","CDT","MST","MDT")
    foreach ($d in $deleteList) { if ($v -eq $d) { return $true } }
    return $false
}

function Norm([string]$s) {
    return ([regex]::Replace($s.Trim(), '\s+', ' ')).ToUpper()
}
function HasAny([string]$name, [string[]]$list) {
    $n = Norm $name
    foreach ($p in $list) { if ($n.Contains((Norm $p))) { return $true } }
    return $false
}

# Fill UTC times - uses Range string refs to avoid COM casting
function Fill-UtcTimes($ws, [string[]]$dstCols, [int]$rowStart, [int]$rowEnd,
                       [int]$startMinutes, [int]$stepMinutes, [string]$utcLabel) {
    $hdrRow = $rowStart - 1
    foreach ($col in $dstCols) {
        try {
            $hdr = $ws.Range($col + "$hdrRow")
            if ($hdr.MergeCells) { $hdr.MergeArea.UnMerge() }
            $hdr.NumberFormat               = "@"
            $hdr.Value2                     = $utcLabel
            $hdr.Font.Color                 = 16777215   # white
            $hdr.Font.Bold                  = $true
            $hdr.HorizontalAlignment        = -4108
            $hdr.Interior.Color             = 0          # black fill
        } catch {}
    }
    foreach ($col in $dstCols) {
        try {
            $cIdx   = $ws.Range($col + "1").Column
            $startC = $ws.Cells.Item($rowStart, $cIdx)
            $endC   = $ws.Cells.Item($rowEnd,   $cIdx)
            $rng    = $ws.Range($startC, $endC)
            if ($rng.MergeCells) { $rng.MergeArea.UnMerge() }
            $rng.ClearContents()
        } catch {}
    }
    $minutes = $startMinutes
    for ($r = $rowStart; $r -le $rowEnd; $r++) {
        $frac = ($minutes % 1440) / 1440.0
        foreach ($col in $dstCols) {
            try {
                $cIdx = $ws.Range($col + "1").Column
                $cell = $ws.Cells.Item($r, $cIdx)
                $cell.Value2              = $frac
                $cell.NumberFormat        = "h:mm"
                $cell.Font.Color          = 0
                $cell.HorizontalAlignment = -4108
                $cell.Borders.LineStyle   = 1
                $cell.Borders.Weight      = 2
            } catch {}
        }
        $minutes += $stepMinutes
    }
}

function Apply-PrintSetup($ws, $xl) {
    try {
        $ps = $ws.PageSetup
        $ps.BlackAndWhite = $false ; $ps.Draft = $false
        $ps.Orientation   = 2      ; $ps.PaperSize = 1
        $m = $xl.InchesToPoints(0.1)
        $ps.LeftMargin = $m ; $ps.RightMargin  = $m
        $ps.TopMargin  = $m ; $ps.BottomMargin = $m
        $ps.HeaderMargin = $m ; $ps.FooterMargin = $m
        $ps.Zoom = $false ; $ps.FitToPagesWide = 1 ; $ps.FitToPagesTall = 0
    } catch {}
}

# -----------------------------------------------------------
# Routing and worksheet selection
# -----------------------------------------------------------
function Get-ExistingGridType([string]$baseName) {
    if (HasAny $baseName @("GRILLA CATV MASTER"))          { return "CATV" }
    if (HasAny $baseName @("GRILLA TVD MASTER"))           { return "TVD" }
    if (HasAny $baseName @("GRILLA MASTER PASIONES LATAM")){ return "PASIONES_LATAM" }
    if (HasAny $baseName @("GRILLA MASTER PASIONES US"))   { return "PASIONES_US" }
    if (HasAny $baseName @("TODO NOVELAS"))                { return "TODO_NOVELAS" }
    if (HasAny $baseName @("REV TV GRID", "REV_TV_GRID")) { return "REV_TV" }
    return $null
}

function Test-SportyNetFileName([string]$baseName) {
    if ([string]::IsNullOrWhiteSpace($baseName)) { return $false }

    # Known client naming families. Existing grid routes are evaluated before
    # this function, so CATV/TVD/Pasiones/Todo/REV TV can never be re-routed.
    if ($baseName -match '(?i)^\s*SNETL(?:\s*[-_]\s*|\s+).*WEEK\s*\d+(?:.*)?$') { return $true }
    if ($baseName -match '(?i)^\s*WEEK\s*\d+\s+V(?:ERSION)?\.?\s*\d+(?:[\s._-]+\d+)*\s*$') { return $true }
    return $false
}

function Get-TabRule([string]$gridType) {
    try {
        $rule = $ManagedSettings.Tabs[$gridType]
        if ($null -ne $rule) { return $rule }
    } catch {}
    return [ordered]@{ Mode = "All"; Count = 0 }
}

function Get-GridLayout([string]$gridType) {
    try {
        $gridLayout = $ManagedSettings.Layouts[$gridType]
        if ($null -ne $gridLayout) {
            return [ordered]@{
                FontSize     = [int]$gridLayout.FontSize
                DefaultRowH  = [double]$gridLayout.DefaultRowHeight
                HeaderRowH   = [double]$gridLayout.HeaderRowHeight
                SmallColW    = [double]$gridLayout.SmallColumnWidth
                DefaultColW  = [double]$gridLayout.DefaultColumnWidth
            }
        }
    } catch {}
    return $DefaultLayout
}

function Select-ItemsByRule($items, [string]$gridType) {
    $array = @($items)
    if ($array.Count -eq 0) { return @() }
    $rule = Get-TabRule $gridType
    $mode = ([string]$rule.Mode).Trim()
    $count = 0
    try { $count = [int]$rule.Count } catch { $count = 0 }
    if ($mode -eq "All" -or $count -le 0 -or $count -ge $array.Count) { return $array }
    if ($mode -eq "Last") {
        $start = $array.Count - $count
        return @($array[$start..($array.Count - 1)])
    }
    return @($array[0..($count - 1)])
}

# -----------------------------------------------------------
# SportyNet strict signature and formatting
# -----------------------------------------------------------
function Test-SportyDateHeader($value) {
    $value = Get-SportyScalarValue $value
    if ($null -eq $value) { return $false }
    if ($value -is [datetime]) { return $true }
    if ($value -is [byte] -or $value -is [sbyte] -or $value -is [int16] -or
        $value -is [uint16] -or $value -is [int32] -or $value -is [uint32] -or
        $value -is [int64] -or $value -is [uint64] -or $value -is [single] -or
        $value -is [double] -or $value -is [decimal]) {
        return ([double]$value -gt 1000.0)
    }
    return $false
}

# Structural and program-data reads must use Value2, not Range.Text. Excel COM
# can occasionally expose a one-cell property as System.Object[]/Object[,].
# These helpers unwrap the first scalar value before any string conversion so a
# valid SportyNet sheet cannot fail with "Unable to cast System.Object[] to String".
function Get-SportyScalarValue($value) {
    if ($null -eq $value) { return $null }
    if ($value -is [System.Array]) {
        foreach ($entry in $value) {
            $scalar = Get-SportyScalarValue $entry
            if ($null -ne $scalar) { return $scalar }
        }
        return $null
    }
    return $value
}

function Get-SportyCellRawValue($cell) {
    try { return (Get-SportyScalarValue $cell.Value2) }
    catch { return $null }
}

function ConvertTo-SportySafeString($value) {
    $scalar = Get-SportyScalarValue $value
    if ($null -eq $scalar) { return "" }
    try { return [System.Convert]::ToString($scalar, [System.Globalization.CultureInfo]::InvariantCulture) }
    catch {
        try { return $scalar.ToString() }
        catch { return "" }
    }
}

function Get-SportyCellAddress($cell) {
    try { return (ConvertTo-SportySafeString ($cell.Address($false,$false))) }
    catch { return "?" }
}

function Get-SportyCellHasFormula($cell) {
    try {
        $value = Get-SportyScalarValue $cell.HasFormula
        if ($null -eq $value) { return $false }
        return [bool]$value
    } catch { return $false }
}

function Get-SportyCellValueText($cell) {
    return (ConvertTo-SportySafeString (Get-SportyCellRawValue $cell)).Trim()
}

function Get-SportyHeaderToken($value) {
    return ((ConvertTo-SportySafeString $value).Trim() -replace '[\s\.\-_]', '').ToUpperInvariant()
}

function Test-SportyBrtHeader($value) {
    return (Get-SportyHeaderToken $value) -in @("BRT", "BRA")
}

function Test-SportyCdmxHeader($value) {
    return (Get-SportyHeaderToken $value) -in @("MEX", "MEXICO", "CDMX")
}

# Lightweight fallback used only for Excel files that match no known filename
# route. It inspects one header row on the first worksheet only. Full workbook
# validation is performed later after the file is routed as SportyNet.
function Test-SportyNetSoftSignature($wb) {
    try {
        if ($null -eq $wb -or [int]$wb.Worksheets.Count -lt 1) { return $false }
        $ws = $wb.Worksheets.Item(1)

        $gmt = Get-SportyCellValueText ($ws.Cells.Item(2,1))
        $leftLocal = Get-SportyCellValueText ($ws.Cells.Item(2,2))
        if ((Get-SportyHeaderToken $gmt) -ne "GMT") { return $false }
        if (-not (Test-SportyBrtHeader $leftLocal) -and -not (Test-SportyCdmxHeader $leftLocal)) { return $false }

        for ($c = 3; $c -le 9; $c++) {
            if (-not (Test-SportyDateHeader $ws.Cells.Item(2,$c).Value2)) { return $false }
        }

        $duplicateLocal = Get-SportyCellValueText ($ws.Cells.Item(2,10))
        $duplicateGmt = Get-SportyCellValueText ($ws.Cells.Item(2,11))
        return ((Test-SportyBrtHeader $duplicateLocal) -and (Get-SportyHeaderToken $duplicateGmt) -eq "GMT")
    } catch {
        return $false
    }
}

function Get-SportyNetWorksheetLayout($ws) {
    try {
        $gmt = Get-SportyCellValueText ($ws.Cells.Item(2,1))
        $leftLocal = Get-SportyCellValueText ($ws.Cells.Item(2,2))
        if ((Get-SportyHeaderToken $gmt) -ne "GMT") { return $null }

        $leftIsBrt = Test-SportyBrtHeader $leftLocal
        $leftIsCdmx = Test-SportyCdmxHeader $leftLocal
        if (-not $leftIsBrt -and -not $leftIsCdmx) { return $null }

        $third = Get-SportyCellValueText ($ws.Cells.Item(2,3))
        if ($leftIsBrt -and (Get-SportyHeaderToken $third) -eq "CDMX") {
            for ($c = 4; $c -le 10; $c++) {
                $rawValue = $ws.Cells.Item(2,$c).Value2
                $unwrappedValue = Get-SportyScalarValue $rawValue
                if (-not (Test-SportyDateHeader $unwrappedValue)) { return $null }
            }
            return [pscustomobject]@{
                Worksheet = $ws; SheetName = (ConvertTo-SportySafeString $ws.Name); IsFormatted = $true
                HeaderRow = 2; ScheduleStartRow = 3; ScheduleEndRow = 98
                GmtColumn = 1; LocalColumn = 2; DayStartColumn = 4; DayEndColumn = 10
                DuplicateLocalColumn = $null; DuplicateGmtColumn = $null
                GmtSourceColumn = 1; BrtSourceColumn = 2; CdmxSourceColumn = 3
                SourceVariant = "FORMATTED"
                LegendRow = $null; LegendStartColumn = $null; Title = ""
            }
        }

        for ($c = 3; $c -le 9; $c++) {
            $rawValue = $ws.Cells.Item(2,$c).Value2
            $unwrappedValue = Get-SportyScalarValue $rawValue
            if (-not (Test-SportyDateHeader $unwrappedValue)) { return $null }
        }

        $dupLocal = Get-SportyCellValueText ($ws.Cells.Item(2,10))
        $dupGmt = Get-SportyCellValueText ($ws.Cells.Item(2,11))
        if (-not (Test-SportyBrtHeader $dupLocal) -or (Get-SportyHeaderToken $dupGmt) -ne "GMT") { return $null }

        $brtSourceColumn = if ($leftIsBrt) { 2 } else { 10 }
        $cdmxSourceColumn = if ($leftIsCdmx) { 2 } else { $null }
        $requiredTimeColumns = @([int]1, [int]$brtSourceColumn)
        if ($null -ne $cdmxSourceColumn) { $requiredTimeColumns += [int]$cdmxSourceColumn }

        for ($r = 3; $r -le 98; $r++) {
            foreach ($c in $requiredTimeColumns) {
                if ($null -eq (Get-SportyCellRawValue ($ws.Cells.Item($r,$c)))) { return $null }
            }
        }

        $legendRow = $null; $legendCol = $null
        for ($r = 99; $r -le 112; $r++) {
            for ($c = 1; $c -le 18; $c++) {
                $v1 = Get-SportyCellValueText ($ws.Cells.Item($r,$c))
                $v2 = Get-SportyCellValueText ($ws.Cells.Item($r,$c+1))
                $v3 = Get-SportyCellValueText ($ws.Cells.Item($r,$c+2))
                $v4 = Get-SportyCellValueText ($ws.Cells.Item($r,$c+3))
                if ($v1 -ieq "Highlights" -and $v2 -ieq "First Airing/Delayed" -and
                    $v3 -ieq "LIVE" -and $v4 -ieq "<Repeat>") {
                    $legendRow = $r; $legendCol = $c; break
                }
            }
            if ($null -ne $legendRow) { break }
        }
        if ($null -eq $legendRow) { return $null }

        $title = ""
        for ($c = 1; $c -le 20; $c++) {
            try {
                $cell = $ws.Cells.Item(1,$c)
                $sourceCell = if ([bool](Get-SportyScalarValue $cell.MergeCells)) { $cell.MergeArea.Cells.Item(1,1) } else { $cell }
                $value = Get-SportyCellRawValue $sourceCell
                $valueText = ConvertTo-SportySafeString $value
                if (-not [string]::IsNullOrWhiteSpace($valueText)) {
                    $title = $valueText; break
                }
            } catch {}
        }
        if ([string]::IsNullOrWhiteSpace($title)) { return $null }

        $sourceVariant = if ($leftIsCdmx) { "CDMX_LEFT_BRT_RIGHT" } else { "BRT_LEFT" }

        return [pscustomobject]@{
            Worksheet = $ws; SheetName = (ConvertTo-SportySafeString $ws.Name); IsFormatted = $false
            HeaderRow = 2; ScheduleStartRow = 3; ScheduleEndRow = 98
            GmtColumn = 1; LocalColumn = 2; DayStartColumn = 3; DayEndColumn = 9
            DuplicateLocalColumn = 10; DuplicateGmtColumn = 11
            GmtSourceColumn = 1; BrtSourceColumn = [int]$brtSourceColumn
            CdmxSourceColumn = $cdmxSourceColumn
            SourceVariant = $sourceVariant
            LegendRow = $legendRow; LegendStartColumn = $legendCol; Title = $title
        }
    } catch {
        return $null
    }
}

function Get-SportyNetLayouts($wb) {
    $layouts = @()
    for ($i = 1; $i -le [int]$wb.Worksheets.Count; $i++) {
        $ws = $wb.Worksheets.Item($i)
        $layout = Get-SportyNetWorksheetLayout $ws
        if ($null -ne $layout) { $layouts += $layout }
    }
    return @($layouts)
}

function Get-ModuloOne([double]$value) { return $value - [Math]::Floor($value) }
function Get-CircularDistance([double]$first, [double]$second) {
    $difference = [Math]::Abs((Get-ModuloOne $first) - (Get-ModuloOne $second))
    if ($difference -gt 0.5) { $difference = 1.0 - $difference }
    return $difference
}

function ConvertTo-SportyNetTime($rawValue, [string]$displayedText, [string]$address, [bool]$hasFormula) {
    $classification = $null; $normalized = $null; $nonstandard = $false
    if ($null -eq $rawValue -or ($rawValue -is [string] -and [string]::IsNullOrWhiteSpace([string]$rawValue))) {
        Throw-Text "${address}: blank time value. Displayed='$displayedText'." "${address}: valor de hora vacio. Mostrado='$displayedText'."
    }
    if ($rawValue -is [bool]) {
        Throw-Text "${address}: Boolean is not a valid time." "${address}: un valor Booleano no es una hora valida."
    }
    $numeric = $rawValue -is [byte] -or $rawValue -is [sbyte] -or $rawValue -is [int16] -or
               $rawValue -is [uint16] -or $rawValue -is [int32] -or $rawValue -is [uint32] -or
               $rawValue -is [int64] -or $rawValue -is [uint64] -or $rawValue -is [single] -or
               $rawValue -is [double] -or $rawValue -is [decimal]
    if ($numeric) {
        $number = [double]$rawValue
        if ($number -ge -$SportyTimeTolerance -and $number -lt (1.0 - $SportyTimeTolerance)) {
            $normalized = Get-ModuloOne $number
            $classification = if ($hasFormula) { "FORMULA_TIME" } else { "VALID_TIME_FRACTION" }
        } elseif ($number -ge 0.0 -and $number -lt 24.0) {
            $hours = [int][Math]::Floor($number + $SportyTimeTolerance)
            $fraction = $number - $hours
            $minutes = [int][Math]::Round($fraction * 100.0, 0, [MidpointRounding]::AwayFromZero)
            $candidate = $hours + ($minutes / 100.0)
            if ($hours -ge 0 -and $hours -le 23 -and $minutes -in @(0,15,30,45) -and
                [Math]::Abs($number - $candidate) -le $SportyHhmmTolerance) {
                $normalized = (($hours * 60.0) + $minutes) / 1440.0
                $classification = if ($hasFormula) { "FORMULA_TIME" } else { "RECOGNIZED_HHMM_NUMERIC" }
                $nonstandard = $true
            } else {
                Throw-Text "${address}: invalid or ambiguous numeric time '$rawValue'." "${address}: hora numerica invalida o ambigua '$rawValue'."
            }
        } else {
            Throw-Text "${address}: numeric time is outside accepted ranges '$rawValue'." "${address}: la hora numerica esta fuera del rango aceptado '$rawValue'."
        }
    } elseif ($rawValue -is [string]) {
        $m = [regex]::Match([string]$rawValue, '^\s*(\d{1,2}):(\d{2})\s*$')
        if (-not $m.Success) {
            Throw-Text "${address}: invalid or ambiguous text time '$rawValue'." "${address}: hora de texto invalida o ambigua '$rawValue'."
        }
        $hours = [int]$m.Groups[1].Value; $minutes = [int]$m.Groups[2].Value
        if ($hours -lt 0 -or $hours -gt 23 -or $minutes -notin @(0,15,30,45)) {
            Throw-Text "${address}: text time is not a quarter-hour value '$rawValue'." "${address}: la hora de texto no corresponde a un cuarto de hora '$rawValue'."
        }
        $normalized = (($hours * 60.0) + $minutes) / 1440.0
        $classification = if ($hasFormula) { "FORMULA_TIME" } else { "RECOGNIZED_TEXT_TIME" }
        $nonstandard = $true
    } else {
        Throw-Text "${address}: unsupported time value type." "${address}: tipo de valor de hora no compatible."
    }
    if ($hasFormula) { $nonstandard = $true }
    return [pscustomobject]@{
        Address=$address; RawValue=$rawValue; DisplayedText=$displayedText
        Classification=$classification; NormalizedValue=[double](Get-ModuloOne $normalized)
        IsNonstandard=[bool]$nonstandard
    }
}

function Get-SportyNetTimeData($ws, $layout) {
    $gmt = @(); $brt = @(); $existingCdmx = @(); $anomalies = @()
    $hasExistingCdmx = $null -ne $layout.CdmxSourceColumn

    for ($r = 3; $r -le 98; $r++) {
        $definitions = @(
            @{Name="GMT"; Col=[int]$layout.GmtSourceColumn},
            @{Name="BRT"; Col=[int]$layout.BrtSourceColumn}
        )
        if ($hasExistingCdmx) {
            $definitions += @{Name="CDMX"; Col=[int]$layout.CdmxSourceColumn}
        }

        foreach ($def in $definitions) {
            $cell = $ws.Cells.Item($r,[int]$def.Col)
            $rawValue = Get-SportyCellRawValue $cell
            $displayedText = ConvertTo-SportySafeString $rawValue
            $addr = "$(ConvertTo-SportySafeString $ws.Name)!$(Get-SportyCellAddress $cell)"
            $conv = ConvertTo-SportyNetTime $rawValue $displayedText $addr (Get-SportyCellHasFormula $cell)
            switch ($def.Name) {
                "GMT"  { $gmt += [double]$conv.NormalizedValue }
                "BRT"  { $brt += [double]$conv.NormalizedValue }
                "CDMX" { $existingCdmx += [double]$conv.NormalizedValue }
            }
            if ($conv.IsNonstandard) { $anomalies += $conv }
        }
    }

    $sequences = @(
        @{Name="GMT"; Values=$gmt},
        @{Name="BRT"; Values=$brt}
    )
    if ($hasExistingCdmx) { $sequences += @{Name="CDMX"; Values=$existingCdmx} }

    foreach ($seq in $sequences) {
        for ($i = 1; $i -lt $seq.Values.Count; $i++) {
            $expected = Get-ModuloOne ($seq.Values[$i-1] + $SportyQuarterHour)
            if ((Get-CircularDistance $seq.Values[$i] $expected) -gt $SportyTimeTolerance) {
                Throw-Text "$($seq.Name) quarter-hour sequence failed at row $($i+3)." "La secuencia de 15 minutos de $($seq.Name) fallo en la fila $($i+3)."
            }
        }
    }

    for ($i = 0; $i -lt $SportyScheduleRows; $i++) {
        $expectedGmt = Get-ModuloOne ($brt[$i] + $SportyThreeHours)
        if ((Get-CircularDistance $gmt[$i] $expectedGmt) -gt $SportyTimeTolerance) {
            Throw-Text "GMT/BRT three-hour relationship failed at row $($i+3)." "La relacion de tres horas entre GMT y BRT fallo en la fila $($i+3)."
        }
    }

    $cdmx = @()
    if ($hasExistingCdmx) {
        for ($i = 0; $i -lt $SportyScheduleRows; $i++) {
            $expectedCdmx = Get-ModuloOne ($brt[$i] - $SportyThreeHours)
            if ((Get-CircularDistance $existingCdmx[$i] $expectedCdmx) -gt $SportyTimeTolerance) {
                Throw-Text "CDMX/Mex and BRT relationship failed at row $($i+3)." "La relacion entre CDMX/Mex y BRT fallo en la fila $($i+3)."
            }
            $cdmx += [double]$existingCdmx[$i]
        }
    } else {
        foreach ($value in $brt) { $cdmx += (Get-ModuloOne ($value - $SportyThreeHours)) }
    }

    return [pscustomobject]@{
        GMT=$gmt; BRT=$brt; CDMX=$cdmx; Anomalies=$anomalies
        SourceVariant=(ConvertTo-SportySafeString $layout.SourceVariant)
    }
}

function Convert-SportyNetProgramText([string]$text, [int]$span, [string]$address) {
    if ($null -eq $text) { $text = "" }
    $normalized = $text.Replace("`r`n","`n").Replace("`r","`n")
    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($part in $normalized.Split([char]10)) { [void]$parts.Add([string]$part) }
    while ($parts.Count -gt 0 -and $parts[$parts.Count-1] -eq "") { $parts.RemoveAt($parts.Count-1) }
    for ($i = 0; $i -lt $parts.Count; $i++) {
        if ([string]::IsNullOrWhiteSpace($parts[$i])) {
            Throw-Text "${address}: internal blank program line at line $($i+1)." "${address}: linea interna vacia en el programa, linea $($i+1)."
        }
    }
    $n = $parts.Count
    if ($n -eq 0) { return "" }
    if ($span -eq 1) { return [string]::Join(" ",$parts.ToArray()) }
    if ($n -eq 1) { return $parts[0] }
    if ($n -eq 2) { return $parts[0] + " " + $parts[1] }
    if ($n -eq 3) { return $parts[0] + "`n" + $parts[1] + " " + $parts[2] }
    $out = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -le ($n-5); $i++) { [void]$out.Add($parts[$i]) }
    [void]$out.Add($parts[$n-4] + " " + $parts[$n-3])
    [void]$out.Add($parts[$n-2] + " " + $parts[$n-1])
    return [string]::Join("`n",$out.ToArray())
}

function Get-SportyNetProgramData($ws, $layout) {
    $blocks = @()
    for ($c = 3; $c -le 9; $c++) {
        $r = 3
        while ($r -le 98) {
            $cell = $ws.Cells.Item($r,$c); $start=$r; $end=$r
            if ([bool]$cell.MergeCells) {
                $area = $cell.MergeArea
                $start = [int]$area.Row; $end = $start + [int]$area.Rows.Count - 1
                if ($start -ne $r -or [int]$area.Column -ne $c -or [int]$area.Columns.Count -ne 1 -or $end -gt 98) {
                    Throw-Text "Unsupported SportyNet merge at $($area.Address())." "Combinacion de celdas SportyNet no compatible en $($area.Address())."
                }
            }
            $sourceCell = if ([bool](Get-SportyScalarValue $cell.MergeCells)) { $cell.MergeArea.Cells.Item(1,1) } else { $cell }
            $raw = ConvertTo-SportySafeString (Get-SportyCellRawValue $sourceCell)
            $text = Convert-SportyNetProgramText $raw ($end-$start+1) "$(ConvertTo-SportySafeString $ws.Name)!$(Get-SportyCellAddress $sourceCell)"
            $blocks += [pscustomobject]@{
                SourceColumn=$c; StartRow=$start; EndRow=$end; Text=$text
                FillColor=[int]$cell.Interior.Color; FontBold=[bool]$cell.Font.Bold
            }
            $r = $end + 1
        }
    }
    $legend = @()
    for ($i=0; $i -lt 4; $i++) {
        $cell = $ws.Cells.Item($layout.LegendRow,$layout.LegendStartColumn+$i)
        $legend += [pscustomobject]@{ Text=(ConvertTo-SportySafeString (Get-SportyCellRawValue $cell)); FillColor=[int]$cell.Interior.Color }
    }
    return [pscustomobject]@{ Blocks=$blocks; Legend=$legend }
}

function Set-ThinBorders($range, [switch]$Inside) {
    $ids = @($XL_EDGE_LEFT,$XL_EDGE_TOP,$XL_EDGE_BOTTOM,$XL_EDGE_RIGHT)
    if ($Inside) { $ids += @($XL_INSIDE_VERTICAL,$XL_INSIDE_HORIZONTAL) }
    foreach ($id in $ids) {
        try {
            $border = $range.Borders.Item($id)
            $border.LineStyle=$XL_CONTINUOUS; $border.Weight=$XL_THIN; $border.Color=$ColorBlack
        } catch {}
    }
}

function Format-SportyNetWorksheet($ws, $layout, $xl) {
    if ($layout.IsFormatted) {
        Show-Text "  [$($ws.Name)] Already formatted - skipping" "  [$($ws.Name)] Ya esta formateada - se omite" DarkGray
        return [pscustomobject]@{ Changed=$false; Anomalies=@() }
    }
    try { $timeData = Get-SportyNetTimeData $ws $layout }
    catch { throw "SportyNet time stage failed: $($_.Exception.Message) | Stack: $($_.ScriptStackTrace)" }
    try { $programData = Get-SportyNetProgramData $ws $layout }
    catch { throw "SportyNet program stage failed: $($_.Exception.Message) | Stack: $($_.ScriptStackTrace)" }
    $title = ConvertTo-SportySafeString $layout.Title

    # Unmerge the title before structural column changes, then delete the
    # duplicate right-side GMT/BRT and insert CDMX at C.
    try { $ws.Rows(1).UnMerge() } catch {}
    $ws.Columns("K").Delete() | Out-Null
    $ws.Columns("J").Delete() | Out-Null
    $ws.Columns("C").Insert() | Out-Null

    try { $ws.Range("A1:N1").Clear() } catch { try { $ws.Range("A1:N1").ClearContents() } catch {} }
    $ws.Range("A1:C1").Merge() | Out-Null
    $ws.Range("A1:C1").Value2 = $title
    $ws.Range("A1:C1").Font.Name="Arial"; $ws.Range("A1:C1").Font.Size=9
    $ws.Range("A1:C1").Font.Bold=$true; $ws.Range("A1:C1").Font.Color=$ColorBlack
    $ws.Range("A1:C1").Interior.Color=$ColorTitle
    $ws.Range("A1:C1").HorizontalAlignment=$XL_CENTER; $ws.Range("A1:C1").VerticalAlignment=$XL_CENTER
    Set-ThinBorders ($ws.Range("A1:C1"))

    $ws.Cells.Item(2,1).Value2="GMT"; $ws.Cells.Item(2,2).Value2="BRT"; $ws.Cells.Item(2,3).Value2="CDMX"
    $ws.Range("A2:C2").Font.Name="Arial"; $ws.Range("A2:C2").Font.Size=14; $ws.Range("A2:C2").Font.Bold=$true
    $ws.Range("A2:C2").HorizontalAlignment=$XL_CENTER; $ws.Range("A2:C2").VerticalAlignment=$XL_CENTER
    Set-ThinBorders ($ws.Range("A2:C2")) -Inside

    $matrix = [object[,]]::new(96,3)
    try {
        for ($i=0; $i -lt 96; $i++) {
            $matrix[$i,0]=$timeData.GMT[$i]
            $matrix[$i,1]=$timeData.BRT[$i]
            $matrix[$i,2]=$timeData.CDMX[$i]
        }
    } catch {
        throw "Matrix construction failed at index ${i}: $($_.Exception.Message) | Type: $($_.Exception.GetType().FullName)"
    }
    try {
        # On this Excel/COM install, Range.Value2's property-set binder mis-resolves a
        # bulk 96x3 object[,] argument and throws "Unable to cast System.Object[,] to
        # System.String" -- reproducible via string-address Range, cell-pair Range, and
        # reflection InvokeMember alike. Range.Value does not hit this; use it for the
        # bulk numeric write (harmless here since these are plain doubles, not
        # currency/date-typed values).
        $ws.Range("A3:C98").Value=$matrix
    } catch {
        throw "Matrix assignment failed: $($_.Exception.Message) | Type: $($_.Exception.GetType().FullName)"
    }
    $ws.Range("A3:C98").NumberFormat="h:mm"; $ws.Range("A3:C98").Font.Name="Arial"; $ws.Range("A3:C98").Font.Size=14
    $ws.Range("A3:C98").Font.Color=$ColorBlack; $ws.Range("A3:A98").Font.Bold=$true; $ws.Range("B3:C98").Font.Bold=$false
    $ws.Range("A3:B98").Interior.Color=$ColorYellow; $ws.Range("C3:C98").Interior.Color=$ColorCdmx
    $ws.Range("A3:C98").HorizontalAlignment=$XL_CENTER; $ws.Range("A3:C98").VerticalAlignment=$XL_CENTER
    Set-ThinBorders ($ws.Range("A3:C98")) -Inside

    $ws.Range("D2:J2").NumberFormat="mm/dd dddd"; $ws.Range("D2:J2").Font.Name="Arial"; $ws.Range("D2:J2").Font.Size=14
    $ws.Range("D2:J2").Font.Bold=$true; $ws.Range("D2:J2").Font.Color=$ColorBlack; $ws.Range("D2:J2").Interior.Color=$ColorDateHeader
    $ws.Range("D2:J2").HorizontalAlignment=$XL_CENTER; $ws.Range("D2:J2").VerticalAlignment=$XL_CENTER

    foreach ($block in $programData.Blocks) {
        $outCol = [int]$block.SourceColumn + 1
        $top = $ws.Cells.Item($block.StartRow,$outCol); $bottom=$ws.Cells.Item($block.EndRow,$outCol)
        $range = $ws.Range($top,$bottom)
        if ($block.EndRow -gt $block.StartRow -and -not [bool]$top.MergeCells) { $range.Merge() | Out-Null }
        $blockText = ConvertTo-SportySafeString $block.Text
        if ([string]::IsNullOrEmpty($blockText)) { $top.ClearContents() } else { $top.Value2=$blockText }
        $range.Interior.Color=[int]$block.FillColor; $range.Font.Name="Arial"; $range.Font.Size=14
        $range.Font.Bold=[bool]$block.FontBold; $range.Font.Color=$ColorBlack; $range.WrapText=$true
        $range.HorizontalAlignment=$XL_CENTER; $range.VerticalAlignment=$XL_CENTER
        Set-ThinBorders $range
    }

    try { $ws.Range("A99:N112").Clear() } catch { try { $ws.Range("A99:N112").ClearContents() } catch {} }
    for ($i=0; $i -lt 4; $i++) {
        $cell=$ws.Cells.Item(100,6+$i); $cell.Value2=$programData.Legend[$i].Text
        $cell.Interior.Color=[int]$programData.Legend[$i].FillColor; $cell.Font.Name="Arial"; $cell.Font.Size=6
        $cell.Font.Bold=$true; $cell.Font.Color=$ColorBlack; $cell.WrapText=$true
        $cell.HorizontalAlignment=$XL_CENTER; $cell.VerticalAlignment=$XL_CENTER
        Set-ThinBorders $cell
    }

    $ws.Range("A1:J100").HorizontalAlignment=$XL_CENTER; $ws.Range("A1:J100").VerticalAlignment=$XL_CENTER
    $ws.Range("A:C").ColumnWidth=10.77734375; $ws.Range("D:J").ColumnWidth=77.77734375
    $ws.Rows.Item(1).RowHeight=23.25; $ws.Rows.Item(2).RowHeight=24
    $ws.Range("3:98").RowHeight=18; $ws.Rows.Item(100).RowHeight=12.75
    for ($i=[int]$ws.Shapes.Count; $i -ge 1; $i--) { try { $ws.Shapes.Item($i).Delete() } catch {} }

    # No SportyNet page breaks, PDF settings, or forced page scaling are added.
    # Operators keep control of the workbook's native print layout.

    foreach ($anomaly in $timeData.Anomalies) {
        Show-Text "  [$($ws.Name)] Normalized $($anomaly.Address): $($anomaly.RawValue) -> $([DateTime]::FromOADate($anomaly.NormalizedValue).ToString('HH:mm'))" `
                  "  [$($ws.Name)] Normalizado $($anomaly.Address): $($anomaly.RawValue) -> $([DateTime]::FromOADate($anomaly.NormalizedValue).ToString('HH:mm'))" DarkYellow
    }
    Show-Text "  [$($ws.Name)] SportyNet formatting completed" "  [$($ws.Name)] Formato SportyNet completado" Green
    return [pscustomobject]@{ Changed=$true; Anomalies=$timeData.Anomalies }
}


# -----------------------------------------------------------
# Optional project run log. The formatter works normally without AI.
# -----------------------------------------------------------
$TranscriptStarted=$false
$AiLogFolder=Join-Path $PSScriptRoot "AI\Logs"
if (Test-Path -LiteralPath $AiLogFolder -PathType Container) {
    try {
        $RunLog=Join-Path $AiLogFolder ("grid_run_{0}.log.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
        Start-Transcript -LiteralPath $RunLog -Force | Out-Null
        $TranscriptStarted=$true
    } catch {}
}

# -----------------------------------------------------------
# Color printer check. Settings is optional at runtime.
# -----------------------------------------------------------
if (-not $SkipPrinterForTest -and [bool]$ManagedSettings.PrinterColorEnabled -and -not [string]::IsNullOrWhiteSpace([string]$ManagedSettings.PrinterName)) {
    $settingsFolder = Join-Path $PSScriptRoot "Settings"
    if (Test-Path -LiteralPath $settingsFolder -PathType Container) {
        $colorFlagFile = Join-Path $settingsFolder "printer_color_set.txt"
    } else {
        $stateFolder = Join-Path $env:LOCALAPPDATA "Mediapro\GridFormatter"
        if (-not (Test-Path -LiteralPath $stateFolder)) { [void](New-Item -ItemType Directory -Path $stateFolder -Force) }
        $colorFlagFile = Join-Path $stateFolder "printer_color_set.txt"
    }
    $today=(Get-Date).ToString("yyyy-MM-dd"); $alreadyColor=$false
    if (Test-Path -LiteralPath $colorFlagFile) {
        try { if ((Get-Content -LiteralPath $colorFlagFile -Raw).Trim() -eq $today) { $alreadyColor=$true } } catch {}
    }
    if (-not $alreadyColor) {
        Show-Text "Setting printer to color..." "Configurando la impresora a color..." Cyan
        try {
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName UIAutomationClient
            Add-Type -AssemblyName UIAutomationTypes
            if (-not ("GridFormatterWinAPI" -as [type])) {
                Add-Type @"
using System;
using System.Runtime.InteropServices;
public class GridFormatterWinAPI {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int cmd);
}
"@
            }
            $printerName=[string]$ManagedSettings.PrinterName
            Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /e /n `"$printerName`"" | Out-Null
            Start-Sleep -Seconds 3
            $desktop=[System.Windows.Automation.AutomationElement]::RootElement; $prefWin=$null
            foreach ($w in $desktop.FindAll([System.Windows.Automation.TreeScope]::Children,[System.Windows.Automation.Condition]::TrueCondition)) {
                if ($w.Current.Name -like "*Printing Preferences*" -or $w.Current.Name -like "*Preferencias de impresion*") { $prefWin=$w; break }
            }
            if ($prefWin) {
                $hwnd=[IntPtr]$prefWin.Current.NativeWindowHandle
                [GridFormatterWinAPI]::ShowWindow($hwnd,9)|Out-Null; [GridFormatterWinAPI]::SetForegroundWindow($hwnd)|Out-Null
                Start-Sleep -Milliseconds 800
                1..4 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{TAB}"); Start-Sleep -Milliseconds 200 }
                [System.Windows.Forms.SendKeys]::SendWait("{HOME}"); Start-Sleep -Milliseconds 400
                [GridFormatterWinAPI]::SetForegroundWindow($hwnd)|Out-Null; Start-Sleep -Milliseconds 200
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                $today | Out-File -LiteralPath $colorFlagFile -Encoding utf8
                Show-Text "Printer set to color. It will not run again today." "Impresora configurada a color. No se repetira hoy." Green
            } else {
                Show-Text "Printer window not found - color setup skipped." "No se encontro la ventana de la impresora - se omitio la configuracion de color." DarkYellow
            }
        } catch {
            Show-Text "Printer color setup failed: $($_.Exception.Message)" "Fallo la configuracion de color: $($_.Exception.Message)" DarkYellow
        }
    } else {
        Show-Text "Printer already set to color today - skipping." "La impresora ya fue configurada a color hoy - se omite." DarkGray
    }
} else {
    Show-Text "Automatic printer color setup is disabled." "La configuracion automatica de color esta desactivada." DarkGray
}


Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Show-Text "  GRID FORMATTER V2  |  $UTC_LABEL offset $UTC_OFFSET  |  Start: ${UTC_START_H}:00 $UTC_LABEL" `
          "  FORMATEADOR DE GRILLAS V2  |  $UTC_LABEL diferencia $UTC_OFFSET  |  Inicio: ${UTC_START_H}:00 $UTC_LABEL" Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$files = Get-ChildItem -LiteralPath $FolderPath -File | Where-Object {
    $_.Name -notlike "~$*" -and $_.Extension -match '^\.(xlsx|xlsm|xls)$'
} | Sort-Object FullName -Unique

if ($files.Count -eq 0) {
    Show-Text "No Excel files found in: $FolderPath" "No se encontraron archivos de Excel en: $FolderPath" Red
    Show-Text "Press any key to exit..." "Presione cualquier tecla para salir..." White
    if (-not $NoPause) { $null=$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") }; exit
}
Show-Text "Excel files found ($($files.Count)):" "Archivos de Excel encontrados ($($files.Count)):" Green
$files | ForEach-Object { Write-Host "  $($_.Name)" }
Write-Host ""

$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.ScreenUpdating=$false
# Deliberately do not set Excel.Application.Calculation. Some environments reject
# that COM property with HRESULT 0x800A03EC; formatting does not require it.

$rptUpdated=@(); $rptSkipped=@(); $rptErrors=@(); $rptUnknown=@()

foreach ($f in $files) {
    Show-Text ">> $($f.Name)" ">> $($f.Name)" Yellow
    $wb=$null
    try {
        $gridType = Get-ExistingGridType $f.BaseName
        $sportyNameMatch = $false
        if ($null -eq $gridType) { $sportyNameMatch = Test-SportyNetFileName $f.BaseName }

        $wb=$xl.Workbooks.Open($f.FullName)
        if ($null -eq $gridType -and $sportyNameMatch) { $gridType="SPORTYNET" }
        if ($null -eq $gridType) {
            # Soft SportyNet fallback is used only after all existing routes
            # and known SportyNet names fail. It reads row 2 of the first tab
            # only; deeper validation happens only after a positive match.
            if (Test-SportyNetSoftSignature $wb) { $gridType="SPORTYNET" }
            else {
                Show-Text "  Unknown grid type - skipped" "  Tipo de grilla desconocido - se omite" DarkGray
                $rptUnknown += $f.Name
                try { $wb.Close($false) } catch {}; $wb=$null
                Write-Host ""; continue
            }
        }

        $workbookChanged=$false
        if ($gridType -eq "SPORTYNET") {
            $layouts=Get-SportyNetLayouts $wb
            if ($layouts.Count -eq 0) {
                $probeParts = @()
                for ($probeIndex = 1; $probeIndex -le [int]$wb.Worksheets.Count; $probeIndex++) {
                    $probeSheet = $wb.Worksheets.Item($probeIndex)
                    $probeParts += "$(ConvertTo-SportySafeString $probeSheet.Name): A2='$(Get-SportyCellValueText ($probeSheet.Cells.Item(2,1)))', B2='$(Get-SportyCellValueText ($probeSheet.Cells.Item(2,2)))', J2='$(Get-SportyCellValueText ($probeSheet.Cells.Item(2,10)))', K2='$(Get-SportyCellValueText ($probeSheet.Cells.Item(2,11)))'"
                }
                $probeText = [string]::Join("; ", [string[]]$probeParts)
                Throw-Text "SportyNet routing matched, but no worksheet has a supported layout. Detected: $probeText" `
                           "La deteccion SportyNet coincidio, pero ninguna hoja tiene una estructura compatible. Detectado: $probeText"
            }
            $selectedLayouts=Select-ItemsByRule $layouts "SPORTYNET"
            foreach ($layout in $selectedLayouts) {
                $lbl="$($f.Name) > $($layout.SheetName)"
                try {
                    $result=Format-SportyNetWorksheet $layout.Worksheet $layout $xl
                    if ($result.Changed) { $workbookChanged=$true; $rptUpdated += $lbl }
                    else { $rptSkipped += "$lbl ($(Get-Text 'already formatted' 'ya formateada'))" }
                } catch {
                    $errorMsg = "$($_.Exception.Message)"
                    $errorStack = "$($_.ScriptStackTrace)"
                    if ($_.Exception.InnerException) {
                        $errorMsg += " | Inner: $($_.Exception.InnerException.Message)"
                    }
                    Show-Text "  [$($layout.SheetName)] ERROR: $errorMsg" `
                              "  [$($layout.SheetName)] ERROR: $errorMsg" Red
                    Show-Text "  [$($layout.SheetName)] Stack: $errorStack" `
                              "  [$($layout.SheetName)] Stack: $errorStack" DarkYellow
                    $rptErrors += "$lbl : $errorMsg"
                }
            }
        } else {
            $isCATV = $gridType -eq "CATV"
            $isTVD = $gridType -eq "TVD"
            $isPasLatam = $gridType -eq "PASIONES_LATAM"
            $isPasUS = $gridType -eq "PASIONES_US"
            $isPas = $isPasLatam -or $isPasUS
            $isTodo = $gridType -eq "TODO_NOVELAS"
            $isRevTV = $gridType -eq "REV_TV"

            $candidates=@(); $total=[int]$wb.Worksheets.Count
            if ($isRevTV) {
                for ($i=1; $i -le $total; $i++) {
                    $testSheet=$wb.Worksheets.Item($i)
                    if ((CellStr $testSheet.Cells.Item(1,1)).Trim() -eq "Zone") { $candidates += $testSheet }
                }
            } else {
                for ($i=1; $i -le $total; $i++) { $candidates += $wb.Worksheets.Item($i) }
            }
            $sheets=Select-ItemsByRule $candidates $gridType

                    foreach ($ws in $sheets) {
                        $lbl = "$($f.Name) > $($ws.Name)"
                        try {

                            # Activate + unfreeze panes
                            $ws.Activate()
                            try { $xl.ActiveWindow.FreezePanes = $false } catch {}

                            # --------------------------------------------------
                            # REV TV: delete K+L, hide E, set widths, landscape 95%
                            # --------------------------------------------------
                            if ($isRevTV) {
                                # Already done check:
                                # If none of the columns we delete still exist, we already ran
                                $alreadyDone = $true
                                $checkNames = @("Program Synopsis","Special","Repeat","Schedule Item Log End Time")
                                for ($c = 1; $c -le 15; $c++) {
                                    try {
                                        $hdr = (CellStr $ws.Cells.Item(1,$c)).Trim()
                                        if ($checkNames -contains $hdr) { $alreadyDone = $false; break }
                                    } catch {}
                                }
                                # Also not done if True/False still visible in row 2
                                if ($alreadyDone) {
                                    for ($c = 1; $c -le 15; $c++) {
                                        try {
                                            $v2 = $ws.Cells.Item(2,$c).Value2
                                            $hidden = $ws.Columns((ColLetter $c)).Hidden
                                            if (-not $hidden -and ($v2 -is [bool] -or "$v2" -eq "True" -or "$v2" -eq "False")) {
                                                $alreadyDone = $false; break
                                            }
                                        } catch {}
                                    }
                                }
                                if ($alreadyDone) {
                                    Show-Text "  [$($ws.Name)] Already done - skipping" "  [$($ws.Name)] Ya estaba procesado - se omite" DarkGray
                                    $rptSkipped += "$lbl ($(Get-Text 'already done' 'ya procesada'))"; continue
                                }

                                # PASS 1: Read ALL header names and data row 2 values upfront
                                # This avoids calling UsedRange repeatedly after deletions
                                $revHeaders = @{}   # col index -> header name
                                $revHideCol = -1    # column index to hide (True/False)
                                $maxScan = 15
                                for ($c = 1; $c -le $maxScan; $c++) {
                                    $hdr = (CellStr $ws.Cells.Item(1, $c)).Trim()
                                    if ($hdr -ne "") { $revHeaders[$c] = $hdr }
                                    # Detect True/False column by header name
                                    if ($hdr -eq "Schedule Item Log End Time" -or $hdr -eq "Repeat") {
                                        $revHideCol = $c
                                    }
                                    # Also detect by checking if row 2 cell Value2 is boolean
                                    try {
                                        $v2 = $ws.Cells.Item(2, $c).Value2
                                        if ($v2 -is [bool] -or "$v2" -eq "True" -or "$v2" -eq "False") {
                                            $revHideCol = $c
                                        }
                                    } catch {}
                                }

                                # PASS 2: Delete unwanted columns RIGHT TO LEFT (so indices stay valid)
                                $deleteNames = @("Program Synopsis","Special","Repeat","Schedule Item Log End Time")
                                $deleteCols  = @()
                                foreach ($c in ($revHeaders.Keys | Sort-Object -Descending)) {
                                    if ($deleteNames -contains $revHeaders[$c]) {
                                        $deleteCols += $c
                                    }
                                }
                                foreach ($c in ($deleteCols | Sort-Object -Descending)) {
                                    try {
                                        $ws.Columns((ColLetter $c)).Delete() | Out-Null
                                        Show-Text "  [$($ws.Name)] Deleted column: $($revHeaders[$c])" "  [$($ws.Name)] Columna eliminada: $($revHeaders[$c])" Gray
                                        # Adjust hide col index if it was after a deleted col
                                        if ($revHideCol -gt $c) { $revHideCol-- }
                                    } catch {}
                                }

                                # PASS 3: Hide True/False column
                                if ($revHideCol -gt 0) {
                                    try {
                                        $ws.Columns((ColLetter $revHideCol)).Hidden = $true
                                        Show-Text "  [$($ws.Name)] Hidden column: $(ColLetter $revHideCol)" "  [$($ws.Name)] Columna oculta: $(ColLetter $revHideCol)" Gray
                                    } catch {}
                                }

                                # PASS 4: Set widths by header name (re-scan after deletions)
                                $widthMap = @{
                                    "Zone"                   = 5.86
                                    "Schedule Item Log Date" = 12.29
                                    "Schedule Item Log Time" = 10.43
                                    "Schedule Item Duration" = 7.57
                                    "Schedule Item End Time" = 9.57
                                    "Program Title"          = 25.86
                                    "File Name"              = 46.43
                                    "Premiere"               = 7.43
                                }
                                for ($c = 1; $c -le $maxScan; $c++) {
                                    try {
                                        $hdr = (CellStr $ws.Cells.Item(1, $c)).Trim()
                                        if ($widthMap.ContainsKey($hdr)) {
                                            $ws.Columns((ColLetter $c)).ColumnWidth = $widthMap[$hdr]
                                        }
                                    } catch {}
                                }

                                # Row height and print setup
                                try {
                                    $ur = $ws.UsedRange
                                    $ws.Rows("$([int]$ur.Row):$([int]($ur.Row+$ur.Rows.Count-1))").RowHeight = 15.75
                                } catch {}
                                try {
                                    $ps = $ws.PageSetup
                                    $ps.Orientation = 2 ; $ps.PaperSize = 1 ; $ps.Zoom = 95
                                    $m = $xl.InchesToPoints(0.1)
                                    $ps.LeftMargin=$m ; $ps.RightMargin=$m ; $ps.TopMargin=$m
                                    $ps.BottomMargin=$m ; $ps.HeaderMargin=$m ; $ps.FooterMargin=$m
                                } catch {}

                                # Give Excel a moment to settle after structural changes
                                Start-Sleep -Milliseconds 800

                                Show-Text "  [$($ws.Name)] Done OK" "  [$($ws.Name)] Completado correctamente" Green
                                $workbookChanged = $true
                                $rptUpdated += $lbl
                                continue
                            }

                            # --------------------------------------------------
                            # CATV / TVD: use old reliable hardcoded approach
                            # --------------------------------------------------
                            if ($isCATV -or $isTVD) {
                                # Skip sheets that don't have a real grid (no ET/RD/PT in rows 1-3)
                                $hasGrid = $false
                                for ($r = 1; $r -le 3; $r++) {
                                    for ($c = 1; $c -le 12; $c++) {
                                        $vn = NormTZ (CellStr $ws.Cells.Item($r,$c))
                                        if ($vn -eq "ET" -or $vn -eq "RD" -or $vn -eq "PT" -or $vn -eq "CA") {
                                            $hasGrid = $true; break
                                        }
                                    }
                                    if ($hasGrid) { break }
                                }
                                if (-not $hasGrid) {
                                    Show-Text "  [$($ws.Name)] No grid data - skipping" "  [$($ws.Name)] No contiene datos de grilla - se omite" DarkYellow
                                    $rptSkipped += "$lbl ($(Get-Text 'no grid data' 'sin datos de grilla'))"; continue
                                }

                                # Check already done
                                $alreadyDone = $false
                                for ($c = 1; $c -le 5; $c++) {
                                    for ($r = 1; $r -le 4; $r++) {
                                        if ((CellStr $ws.Cells.Item($r,$c)).ToUpper().StartsWith("UTC")) {
                                            $alreadyDone = $true; break
                                        }
                                    }
                                    if ($alreadyDone) { break }
                                }
                                if ($alreadyDone) {
                                    Show-Text "  [$($ws.Name)] Already done - skipping" "  [$($ws.Name)] Ya estaba procesado - se omite" DarkGray
                                    $rptSkipped += "$lbl ($(Get-Text 'already done' 'ya procesada'))"; continue
                                }

                                # CATV structure is always:
                                #   Row 1 = blank
                                #   Row 2 = title
                                #   Row 3 = header (RD, ET, dates)
                                #   Row 4+ = data
                                # TVD structure: Row 1=title, Row 2=header, Row 3+=data

                                # STEP 1: Save title
                                $catvTitle = $f.BaseName
                                if ($isCATV) {
                                    # Title is on row 2 for CATV
                                    for ($c = 1; $c -le 11; $c++) {
                                        try {
                                            $cell = $ws.Cells.Item(2, $c)
                                            $v = if ($cell.MergeCells) { CellStr $cell.MergeArea.Cells.Item(1,1) } else { CellStr $cell }
                                            if ($v -ne "") { $catvTitle = $v; break }
                                        } catch {}
                                    }
                                } else {
                                    # TVD: title on row 1
                                    for ($c = 1; $c -le 11; $c++) {
                                        try {
                                            $cell = $ws.Cells.Item(1, $c)
                                            $v = if ($cell.MergeCells) { CellStr $cell.MergeArea.Cells.Item(1,1) } else { CellStr $cell }
                                            if ($v -ne "") { $catvTitle = $v; break }
                                        } catch {}
                                    }
                                }
                                Show-Text "  [$($ws.Name)] Title: $catvTitle" "  [$($ws.Name)] Titulo: $catvTitle" Gray

                                # STEP 2: CATV - delete only row 1 (the blank row)
                                # Result: row1=title, row2=header(ET/RD), row3+=data
                                # Fill-UtcTimes(rowStart=3) writes UTC label to row2, data rows 3-50
                                # Row 1 title is never touched
                                if ($isCATV) {
                                    try { $ws.Rows(1).Delete() | Out-Null } catch {}
                                }

                                # STEP 3: Cleanup
                                try { $ws.Columns("L").Delete() | Out-Null } catch {}
                                try { $ws.Range("51:55").UnMerge() } catch {}
                                try { $ws.Rows("51:55").Delete() | Out-Null } catch {}

                                # STEP 4: Fill UTC
                                $startMins = $UTC_START_H * 60
                                # rowStart=3: UTC header -> row2, data -> rows 3-50 (same for both)
                                Fill-UtcTimes $ws @("A","K") 3 $DataRowEnd $startMins 30 $UTC_LABEL

                                # STEP 5: Formatting
                                $gridLayout = Get-GridLayout $gridType
                                $ur = $ws.UsedRange
                                try { $ur.WrapText = $true ; $ur.Font.Size = $gridLayout.FontSize } catch {}
                                if ($isCATV) { try { $ur.Font.Bold = $true } catch {} }
                                # TVD: bold the ET and UTC time columns
                                if ($isTVD) {
                                    try { $ws.Columns("A:B").Font.Bold = $true } catch {}
                                    try { $ws.Columns("J:K").Font.Bold = $true } catch {}
                                }
                                try {
                                    $ws.Columns("A:B").ColumnWidth = $gridLayout.SmallColW
                                    $ws.Columns("J:K").ColumnWidth = $gridLayout.SmallColW
                                    $ws.Columns("C:I").ColumnWidth = $gridLayout.DefaultColW
                                } catch {}
                                try {
                                    $fr = [int]$ur.Row ; $lr = [int]($ur.Row + $ur.Rows.Count - 1)
                                    $ws.Rows("${fr}:${lr}").RowHeight = $gridLayout.DefaultRowH
                                    $ws.Rows($fr).RowHeight = $gridLayout.HeaderRowH
                                    if ($lr -gt $fr) { $ws.Rows($fr + 1).RowHeight = $gridLayout.HeaderRowH }
                                } catch {}
                                Apply-PrintSetup $ws $xl


                                Show-Text "  [$($ws.Name)] Done OK" "  [$($ws.Name)] Completado correctamente" Green
                                $workbookChanged = $true
                                $rptUpdated += $lbl
                                continue
                            }

                            # --------------------------------------------------
                            # Pasiones / Todo Novelas: dynamic column detection
                            # KEY FIX: scan HEADER ROW directly for boundaries
                            # (not UsedRange which picks up stray far-right cells)
                            # --------------------------------------------------

                            # STEP 1 - Find header row
                            $hRow = -1
                            for ($r = 1; $r -le 6; $r++) {
                                for ($c = 1; $c -le 20; $c++) {
                                    $v = (CellStr $ws.Cells.Item($r,$c)).ToUpper()
                                    $vn = NormTZ $v
                                    if ($vn -eq "ET" -or $vn -eq "PT" -or $vn -eq "PST" -or
                                        $vn -eq "PDT" -or $vn -eq "RD" -or $vn -eq "CA"  -or
                                        $v.StartsWith("UTC")) {
                                        $hRow = $r; break
                                    }
                                }
                                if ($hRow -gt 0) { break }
                            }
                            if ($hRow -lt 0) {
                                Show-Text "  [$($ws.Name)] No time header - skipping" "  [$($ws.Name)] No se encontro encabezado de hora - se omite" DarkYellow
                                $rptSkipped += "$lbl ($(Get-Text 'no time header' 'sin encabezado de hora'))"; continue
                            }
                            Show-Text "  [$($ws.Name)] Header row: $hRow" "  [$($ws.Name)] Fila de encabezado: $hRow" Gray

                            # STEP 2 - Already done?
                            $alreadyDone = $false
                            for ($c = 1; $c -le 20; $c++) {
                                if ((CellStr $ws.Cells.Item($hRow,$c)).ToUpper().StartsWith("UTC")) {
                                    $alreadyDone = $true; break
                                }
                            }
                            if ($alreadyDone) {
                                Show-Text "  [$($ws.Name)] Already done - skipping" "  [$($ws.Name)] Ya estaba procesado - se omite" DarkGray
                                $rptSkipped += "$lbl ($(Get-Text 'already done' 'ya procesada'))"; continue
                            }

                            # STEP 3 - Find first and last populated column IN HEADER ROW
                            # This avoids stray cells far to the right from UsedRange
                            $hFc = -1 ; $hLc = -1
                            for ($c = 1; $c -le 30; $c++) {
                                $v = CellStr $ws.Cells.Item($hRow, $c)
                                if ($v -ne "") {
                                    if ($hFc -lt 0) { $hFc = $c }
                                    $hLc = $c
                                }
                            }
                            if ($hFc -lt 0) {
                                Show-Text "  [$($ws.Name)] Empty header row - skipping" "  [$($ws.Name)] Fila de encabezado vacia - se omite" DarkYellow
                                $rptSkipped += "$lbl ($(Get-Text 'empty header' 'encabezado vacio'))"; continue
                            }
                            Show-Text "  [$($ws.Name)] Header columns: $(ColLetter $hFc) to $(ColLetter $hLc)" "  [$($ws.Name)] Columnas de encabezado: $(ColLetter $hFc) a $(ColLetter $hLc)" Gray

                            # STEP 4 - Save title from row 1 (if header is not row 1)
                            # Always scan from col 1 - title merge anchor may be before $hFc
                            $titleTxt    = ""
                            $hasTitleRow = ($hRow -ge 2)
                            if ($hasTitleRow) {
                                for ($c = 1; $c -le 25; $c++) {
                                    try {
                                        $cell = $ws.Cells.Item(1, $c)
                                        $v    = if ($cell.MergeCells) {
                                                    CellStr $cell.MergeArea.Cells.Item(1,1)
                                                } else { CellStr $cell }
                                        if ($v -ne "") { $titleTxt = $v; break }
                                    } catch {}
                                }
                                if ($titleTxt -eq "") { $titleTxt = $f.BaseName }
                            } else {
                                $titleTxt = $f.BaseName
                            }
                            Show-Text "  [$($ws.Name)] Title: $titleTxt" "  [$($ws.Name)] Titulo: $titleTxt" Gray

                            # STEP 5 - Delete left TZ column if not ET/UTC
                            $leftVal = CellStr $ws.Cells.Item($hRow, $hFc)
                            if (IsTZToDelete $leftVal) {
                                $ws.Columns( (ColLetter $hFc) ).Delete() | Out-Null
                                Show-Text "  [$($ws.Name)] Removed left time zone: $leftVal" "  [$($ws.Name)] Zona horaria izquierda eliminada: $leftVal" Gray
                                # Recalculate - everything shifted left by 1
                                $hFc = $hFc        # stays same (next col moved here)
                                $hLc = $hLc - 1
                            }

                            # STEP 6 - Delete right TZ column if not ET/UTC
                            $rightVal = CellStr $ws.Cells.Item($hRow, $hLc)
                            if (IsTZToDelete $rightVal) {
                                $ws.Columns( (ColLetter $hLc) ).Delete() | Out-Null
                                Show-Text "  [$($ws.Name)] Removed right time zone: $rightVal" "  [$($ws.Name)] Zona horaria derecha eliminada: $rightVal" Gray
                                $hLc = $hLc - 1
                            }

                            # STEP 7 - Insert UTC left and right using column letters
                            $fcLetter = ColLetter $hFc
                            $ws.Columns($fcLetter).Insert() | Out-Null
                            # After left insert everything shifts +1
                            $utcLetter_L = ColLetter $hFc
                            $etLetter_L  = ColLetter ($hFc + 1)
                            $etLetter_R  = ColLetter ($hLc + 1)
                            $utcLetter_R = ColLetter ($hLc + 2)
                            $ws.Columns($utcLetter_R).Insert() | Out-Null

                            # STEP 8 - Fill UTC times
                            $dataStart = $hRow + 1
                            $startMins = $UTC_START_H * 60
                            Fill-UtcTimes $ws @($utcLetter_L, $utcLetter_R) $dataStart $DataRowEnd $startMins 30 $UTC_LABEL

                            # Get per-grid layout settings
                            $gridLayout = Get-GridLayout $gridType

                            # STEP 9 - Restore title row
                            # Column insertions in STEP 7 shift/break row 1 merge for all grids
                            # Always restore title after column ops
                            $newFc = [int]$ws.Range($utcLetter_L + "1").Column
                            $newLc = [int]$ws.Range($utcLetter_R + "1").Column
                            if ($hasTitleRow) {
                                try {
                                    try { $ws.Rows(1).UnMerge() } catch {}
                                    $ws.Cells.Item(1, $newFc).Value2 = $titleTxt
                                    $ws.Range($ws.Cells.Item(1,$newFc), $ws.Cells.Item(1,$newLc)).Merge() | Out-Null
                                    $ws.Cells.Item(1, $newFc).HorizontalAlignment = -4108
                                    $ws.Cells.Item(1, $newFc).Font.Bold           = $true
                                    $ws.Cells.Item(1, $newFc).Font.Size           = $gridLayout.FontSize
                                    $ws.Cells.Item(1, $newFc).Font.Color          = 0
                                    $ws.Rows(1).RowHeight = $gridLayout.HeaderRowH
                                } catch {}
                            } else {
                                try {
                                    $ws.Rows(1).Insert() | Out-Null
                                    $newFc2 = [int]$ws.Range($utcLetter_L + "1").Column
                                    $newLc2 = [int]$ws.Range($utcLetter_R + "1").Column
                                    $ws.Cells.Item(1, $newFc2).Value2 = $titleTxt
                                    $ws.Range($ws.Cells.Item(1,$newFc2), $ws.Cells.Item(1,$newLc2)).Merge() | Out-Null
                                    $ws.Cells.Item(1, $newFc2).HorizontalAlignment = -4108
                                    $ws.Cells.Item(1, $newFc2).Font.Bold           = $true
                                    $ws.Cells.Item(1, $newFc2).Font.Size           = $layout.FontSize
                                    $ws.Cells.Item(1, $newFc2).Font.Color          = 0
                                    $ws.Rows(1).RowHeight = $layout.HeaderRowH
                                } catch {}
                            }

                            # STEP 10 - Formatting
                            $ur = $ws.UsedRange
                            try { $ur.WrapText = $true ; $ur.Font.Size = $gridLayout.FontSize } catch {}
                            if ($isPas -or $isTodo) { try { $ur.Font.Bold = $true } catch {} }

                            $twStr = if ($isPas -or $isTodo) { $PF_TimeColW } else { $gridLayout.SmallColW }
                            try { $ws.Columns("${utcLetter_L}:${etLetter_L}").ColumnWidth = $twStr } catch {}
                            try { $ws.Columns("${etLetter_R}:${utcLetter_R}").ColumnWidth = $twStr } catch {}

                            $centerL = ColLetter ([int]$ws.Range($etLetter_L + "1").Column + 1)
                            $centerR = ColLetter ([int]$ws.Range($etLetter_R + "1").Column - 1)
                            $cw = if ($isPasLatam -or $isTodo) { $PF_CenterW } else { $gridLayout.DefaultColW }
                            try { $ws.Columns("${centerL}:${centerR}").ColumnWidth = $cw } catch {}

                            try {
                                $fr = [int]$ur.Row ; $lr = [int]($ur.Row + $ur.Rows.Count - 1)
                                $ws.Rows("${fr}:${lr}").RowHeight = $gridLayout.DefaultRowH
                                $ws.Rows($fr).RowHeight = $gridLayout.HeaderRowH
                                if ($lr -gt $fr) { $ws.Rows($fr + 1).RowHeight = $gridLayout.HeaderRowH }
                            } catch {}
                            if ($isPas -or $isTodo) {
                                try { $ws.Rows("3:50").RowHeight = 22 } catch {}
                            }

                            Apply-PrintSetup $ws $xl


                            Show-Text "  [$($ws.Name)] Done OK" "  [$($ws.Name)] Completado correctamente" Green
                                $workbookChanged = $true
                            $rptUpdated += $lbl

                        } catch {
                            Show-Text "  [$($ws.Name)] ERROR: $($_.Exception.Message)" "  [$($ws.Name)] ERROR: $($_.Exception.Message)" Red
                            $rptErrors += "$lbl : $($_.Exception.Message)"
                        }
                    }

        }

        if ($workbookChanged) {
            try { $wb.Save() }
            catch {
                Show-Text "  Save failed, retrying..." "  Fallo al guardar, intentando nuevamente..." DarkYellow
                Start-Sleep -Milliseconds 1500
                try { $wb.Save() } catch { Show-Text "  Save error: $($_.Exception.Message)" "  Error al guardar: $($_.Exception.Message)" Red }
            }
        }
        try { $wb.Close($false) } catch {}; $wb=$null
    } catch {
        Show-Text "  Cannot process: $($_.Exception.Message)" "  No se puede procesar: $($_.Exception.Message)" Red
        $rptErrors += "$($f.Name): $($_.Exception.Message)"
        if ($null -ne $wb) { try { $wb.Close($false) } catch {}; $wb=$null }
    }
    Write-Host ""
}

try { $xl.Quit() } catch {}
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Show-Text "  FINAL REPORT" "  INFORME FINAL" Cyan
Write-Host "============================================" -ForegroundColor Cyan
Show-Text "UPDATED ($($rptUpdated.Count)):" "ACTUALIZADAS ($($rptUpdated.Count)):" Green
$rptUpdated | ForEach-Object { Write-Host "  + $_" -ForegroundColor Green }
Write-Host ""
Show-Text "SKIPPED ($($rptSkipped.Count)):" "OMITIDAS ($($rptSkipped.Count)):" DarkGray
$rptSkipped | ForEach-Object { Write-Host "  - $_" -ForegroundColor DarkGray }
if ($rptUnknown.Count -gt 0) {
    Write-Host ""
    Show-Text "UNKNOWN ($($rptUnknown.Count)):" "DESCONOCIDAS ($($rptUnknown.Count)):" DarkGray
    $rptUnknown | ForEach-Object { Write-Host "  ? $_" -ForegroundColor DarkGray }
}
if ($rptErrors.Count -gt 0) {
    Write-Host ""
    Show-Text "ERRORS ($($rptErrors.Count)):" "ERRORES ($($rptErrors.Count)):" Red
    $rptErrors | ForEach-Object { Write-Host "  ! $_" -ForegroundColor Red }
}
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
if ($TranscriptStarted) { try { Stop-Transcript | Out-Null } catch {} }
if (-not $NoPause) {
    Show-Text "Press any key to exit..." "Presione cualquier tecla para salir..." White
    $null=$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
