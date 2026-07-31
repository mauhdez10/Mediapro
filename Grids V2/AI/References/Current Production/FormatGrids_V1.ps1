# ===========================================================
#  GRID FORMATTER - FormatGrids.ps1
# ===========================================================
#  UTC OFFSET - Update twice a year:
#    Mar-Nov : -4      Nov-Mar : -5
# ===========================================================
$UTC_OFFSET  = -4
$UTC_LABEL   = "UTC"
$UTC_START_H = 6 + [Math]::Abs($UTC_OFFSET)   # 10 for -4, 11 for -5

$FolderPath  = $PSScriptRoot

$GridPatterns = @(
    "GRILLA CATV MASTER*",
    "GRILLA MASTER PASIONES LATAM*",
    "GRILLA MASTER PASIONES US*",
    "GRILLA TVD MASTER*",
    "TODO NOVELAS*",
    "REV TV GRID*",
    "REV_TV_Grid*"
)

$LastFiveNames = @(
    "TODO NOVELAS",
    "GRILLA MASTER PASIONES US",
    "GRILLA MASTER PASIONES LATAM"
)

$FontSize    = 14
$DefaultRowH = 25
$HeaderRowH  = 35
$SmallColW   = 7
$DefaultColW = 37
$PF_TimeColW = 8
$PF_CenterW  = 30
$DataRowEnd  = 50

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
# ===========================================================
#  COLOR PRINTER CHECK - runs once per day automatically
# ===========================================================
$colorFlagFile = Join-Path $PSScriptRoot "printer_color_set.txt"
$today         = (Get-Date).ToString("yyyy-MM-dd")
$alreadyColor  = $false

if (Test-Path $colorFlagFile) {
    $saved = (Get-Content $colorFlagFile -Raw).Trim()
    if ($saved -eq $today) { $alreadyColor = $true }
}

if (-not $alreadyColor) {
    Write-Host "Setting printer to color..." -ForegroundColor Cyan

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes

    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinAPI {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int cmd);
}
"@

    $printerName = "MediaproUS_Canon_Online"
    Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /e /n `"$printerName`"" | Out-Null
    Start-Sleep -Seconds 3

    $desktop = [System.Windows.Automation.AutomationElement]::RootElement
    $prefWin = $null
    foreach ($w in $desktop.FindAll([System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.Condition]::TrueCondition)) {
        if ($w.Current.Name -like "*Printing Preferences*") { $prefWin = $w; break }
    }

    if ($prefWin) {
        $hwnd = [IntPtr]$prefWin.Current.NativeWindowHandle
        [WinAPI]::ShowWindow($hwnd, 9)       | Out-Null
        [WinAPI]::SetForegroundWindow($hwnd) | Out-Null
        Start-Sleep -Milliseconds 800

        [System.Windows.Forms.SendKeys]::SendWait("{TAB}") ; Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}") ; Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}") ; Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}") ; Start-Sleep -Milliseconds 400
        [System.Windows.Forms.SendKeys]::SendWait("{HOME}") ; Start-Sleep -Milliseconds 400

        [WinAPI]::SetForegroundWindow($hwnd) | Out-Null
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

        # Save today's date so it won't run again today
        $today | Out-File -FilePath $colorFlagFile -Encoding utf8
        Write-Host "Printer set to color. Won't run again today." -ForegroundColor Green
    } else {
        Write-Host "Printer window not found - skipping color fix." -ForegroundColor DarkYellow
    }

    Start-Sleep -Milliseconds 500
} else {
    Write-Host "Printer already set to color today - skipping." -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  GRID FORMATTER  |  UTC$UTC_OFFSET  |  Start: ${UTC_START_H}:00 UTC" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$files = @()
foreach ($pat in $GridPatterns) {
    $found = Get-ChildItem -Path $FolderPath -Filter $pat -File |
             Where-Object { $_.Extension -match "\.(xlsx|xlsm|xls)$" }
    $files += $found
}
$files = $files | Sort-Object FullName -Unique

if ($files.Count -eq 0) {
    Write-Host "No files found in: $FolderPath" -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}
Write-Host "Files found ($($files.Count)):" -ForegroundColor Green
$files | ForEach-Object { Write-Host "  $($_.Name)" }
Write-Host ""

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false ; $xl.DisplayAlerts = $false ; $xl.ScreenUpdating = $false

$rptUpdated = @() ; $rptSkipped = @() ; $rptErrors = @()

foreach ($f in $files) {
    Write-Host ">> $($f.Name)" -ForegroundColor Yellow
    try {
        $wb         = $xl.Workbooks.Open($f.FullName)
        $isLastFive = HasAny $f.BaseName $LastFiveNames
        $isCATV     = HasAny $f.BaseName @("GRILLA CATV")
        $isTVD      = HasAny $f.BaseName @("GRILLA TVD")
        $isPas      = HasAny $f.BaseName @("GRILLA MASTER PASIONES")
        $isTodo     = HasAny $f.BaseName @("TODO NOVELAS")
        $isPasLatam = HasAny $f.BaseName @("PASIONES LATAM")
        $isRevTV    = HasAny $f.BaseName @("REV TV GRID", "REV_TV_Grid")

        $total = [int]$wb.Worksheets.Count
        $sheets = @()
        if ($isCATV -or $isTVD) {
            # Only ever process Sheet1 - other sheets are unused templates
            $sheets += $wb.Worksheets.Item(1)
        } elseif ($isRevTV) {
            # Only process sheets where A1 = "Zone" (the data sheet)
            for ($i = 1; $i -le $total; $i++) {
                try {
                    $testSheet = $wb.Worksheets.Item($i)
                    $a1val = (CellStr $testSheet.Cells.Item(1,1)).Trim()
                    if ($null -ne $testSheet -and $a1val -eq "Zone") {
                        $sheets += $testSheet
                    }
                } catch {}
            }
        } elseif ($isLastFive -and $total -gt 5) {
            for ($i = ($total - 4); $i -le $total; $i++) { $sheets += $wb.Worksheets.Item($i) }
        } else {
            for ($i = 1; $i -le $total; $i++) { $sheets += $wb.Worksheets.Item($i) }
        }

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
                        Write-Host "  [$($ws.Name)] Already done - skipping" -ForegroundColor DarkGray
                        $rptSkipped += "$lbl (already done)"; continue
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
                            Write-Host "  [$($ws.Name)] Deleted col: $($revHeaders[$c])" -ForegroundColor Gray
                            # Adjust hide col index if it was after a deleted col
                            if ($revHideCol -gt $c) { $revHideCol-- }
                        } catch {}
                    }

                    # PASS 3: Hide True/False column
                    if ($revHideCol -gt 0) {
                        try {
                            $ws.Columns((ColLetter $revHideCol)).Hidden = $true
                            Write-Host "  [$($ws.Name)] Hidden col: $(ColLetter $revHideCol)" -ForegroundColor Gray
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

                    Write-Host "  [$($ws.Name)] Done OK" -ForegroundColor Green
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
                        Write-Host "  [$($ws.Name)] No grid data - skipping" -ForegroundColor DarkYellow
                        $rptSkipped += "$lbl (no grid data)"; continue
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
                        Write-Host "  [$($ws.Name)] Already done - skipping" -ForegroundColor DarkGray
                        $rptSkipped += "$lbl (already done)"; continue
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
                    Write-Host "  [$($ws.Name)] Title: $catvTitle" -ForegroundColor Gray

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
                    $ur = $ws.UsedRange
                    try { $ur.WrapText = $true ; $ur.Font.Size = $FontSize } catch {}
                    if ($isCATV) { try { $ur.Font.Bold = $true } catch {} }
                    # TVD: bold the ET and UTC time columns
                    if ($isTVD) {
                        try { $ws.Columns("A:B").Font.Bold = $true } catch {}
                        try { $ws.Columns("J:K").Font.Bold = $true } catch {}
                    }
                    try {
                        $ws.Columns("A:B").ColumnWidth = $SmallColW
                        $ws.Columns("J:K").ColumnWidth = $SmallColW
                        $ws.Columns("C:I").ColumnWidth = $DefaultColW
                    } catch {}
                    try {
                        $fr = [int]$ur.Row ; $lr = [int]($ur.Row + $ur.Rows.Count - 1)
                        $ws.Rows("${fr}:${lr}").RowHeight = $DefaultRowH
                        $ws.Rows($fr).RowHeight = $HeaderRowH
                        if ($lr -gt $fr) { $ws.Rows($fr + 1).RowHeight = $HeaderRowH }
                    } catch {}
                    Apply-PrintSetup $ws $xl


                    Write-Host "  [$($ws.Name)] Done OK" -ForegroundColor Green
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
                    Write-Host "  [$($ws.Name)] No time header - skipping" -ForegroundColor DarkYellow
                    $rptSkipped += "$lbl (no time header)"; continue
                }
                Write-Host "  [$($ws.Name)] Header row: $hRow" -ForegroundColor Gray

                # STEP 2 - Already done?
                $alreadyDone = $false
                for ($c = 1; $c -le 20; $c++) {
                    if ((CellStr $ws.Cells.Item($hRow,$c)).ToUpper().StartsWith("UTC")) {
                        $alreadyDone = $true; break
                    }
                }
                if ($alreadyDone) {
                    Write-Host "  [$($ws.Name)] Already done - skipping" -ForegroundColor DarkGray
                    $rptSkipped += "$lbl (already done)"; continue
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
                    Write-Host "  [$($ws.Name)] Empty header row - skipping" -ForegroundColor DarkYellow
                    $rptSkipped += "$lbl (empty header)"; continue
                }
                Write-Host "  [$($ws.Name)] Header cols: $(ColLetter $hFc) to $(ColLetter $hLc)" -ForegroundColor Gray

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
                Write-Host "  [$($ws.Name)] Title: $titleTxt" -ForegroundColor Gray

                # STEP 5 - Delete left TZ column if not ET/UTC
                $leftVal = CellStr $ws.Cells.Item($hRow, $hFc)
                if (IsTZToDelete $leftVal) {
                    $ws.Columns( (ColLetter $hFc) ).Delete() | Out-Null
                    Write-Host "  [$($ws.Name)] Removed left: $leftVal" -ForegroundColor Gray
                    # Recalculate - everything shifted left by 1
                    $hFc = $hFc        # stays same (next col moved here)
                    $hLc = $hLc - 1
                }

                # STEP 6 - Delete right TZ column if not ET/UTC
                $rightVal = CellStr $ws.Cells.Item($hRow, $hLc)
                if (IsTZToDelete $rightVal) {
                    $ws.Columns( (ColLetter $hLc) ).Delete() | Out-Null
                    Write-Host "  [$($ws.Name)] Removed right: $rightVal" -ForegroundColor Gray
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
                        $ws.Cells.Item(1, $newFc).Font.Size           = $FontSize
                        $ws.Cells.Item(1, $newFc).Font.Color          = 0
                        $ws.Rows(1).RowHeight = $HeaderRowH
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
                        $ws.Cells.Item(1, $newFc2).Font.Size           = $FontSize
                        $ws.Cells.Item(1, $newFc2).Font.Color          = 0
                        $ws.Rows(1).RowHeight = $HeaderRowH
                    } catch {}
                }

                # STEP 10 - Formatting
                $ur = $ws.UsedRange
                try { $ur.WrapText = $true ; $ur.Font.Size = $FontSize } catch {}
                if ($isPas -or $isTodo) { try { $ur.Font.Bold = $true } catch {} }

                $twStr = if ($isPas -or $isTodo) { $PF_TimeColW } else { $SmallColW }
                try { $ws.Columns("${utcLetter_L}:${etLetter_L}").ColumnWidth = $twStr } catch {}
                try { $ws.Columns("${etLetter_R}:${utcLetter_R}").ColumnWidth = $twStr } catch {}

                $centerL = ColLetter ([int]$ws.Range($etLetter_L + "1").Column + 1)
                $centerR = ColLetter ([int]$ws.Range($etLetter_R + "1").Column - 1)
                $cw = if ($isPasLatam -or $isTodo) { $PF_CenterW } else { $DefaultColW }
                try { $ws.Columns("${centerL}:${centerR}").ColumnWidth = $cw } catch {}

                try {
                    $fr = [int]$ur.Row ; $lr = [int]($ur.Row + $ur.Rows.Count - 1)
                    $ws.Rows("${fr}:${lr}").RowHeight = $DefaultRowH
                    $ws.Rows($fr).RowHeight = $HeaderRowH
                    if ($lr -gt $fr) { $ws.Rows($fr + 1).RowHeight = $HeaderRowH }
                } catch {}
                if ($isPas -or $isTodo) {
                    try { $ws.Rows("3:50").RowHeight = 22 } catch {}
                }

                Apply-PrintSetup $ws $xl


                Write-Host "  [$($ws.Name)] Done OK" -ForegroundColor Green
                $rptUpdated += $lbl

            } catch {
                Write-Host "  [$($ws.Name)] ERROR: $($_.Exception.Message)" -ForegroundColor Red
                $rptErrors += "$lbl : $($_.Exception.Message)"
            }
        }

        try {
            $wb.Save()
        } catch {
            Write-Host "  Save failed, retrying..." -ForegroundColor DarkYellow
            Start-Sleep -Milliseconds 1500
            try { $wb.Save() } catch { Write-Host "  Save error: $($_.Exception.Message)" -ForegroundColor Red }
        }
        try { $wb.Close($false) } catch {}

    } catch {
        Write-Host "  Cannot open: $($_.Exception.Message)" -ForegroundColor Red
        $rptErrors += "$($f.Name): $($_.Exception.Message)"
    }
    Write-Host ""
}

$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
[GC]::Collect()

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  FINAL REPORT" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "UPDATED ($($rptUpdated.Count)):" -ForegroundColor Green
$rptUpdated | ForEach-Object { Write-Host "  + $_" -ForegroundColor Green }
Write-Host ""
Write-Host "SKIPPED ($($rptSkipped.Count)):" -ForegroundColor DarkGray
$rptSkipped | ForEach-Object { Write-Host "  - $_" -ForegroundColor DarkGray }
if ($rptErrors.Count -gt 0) {
    Write-Host ""
    Write-Host "ERRORS ($($rptErrors.Count)):" -ForegroundColor Red
    $rptErrors | ForEach-Object { Write-Host "  ! $_" -ForegroundColor Red }
}
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")