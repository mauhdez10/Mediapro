# Updated Settings UI.ps1 with Layout Settings tab
# ================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$settingsFolder=$PSScriptRoot
$root=Split-Path $settingsFolder -Parent
$scriptPath=Join-Path $root "FormatGrids.ps1"
$jsonPath=Join-Path $settingsFolder "settings.json"
$beginMarker="# >>> BEGIN MANAGED SETTINGS >>>"
$endMarker="# <<< END MANAGED SETTINGS <<<"

$defaults=[ordered]@{
    language="English"; utcOffset=-4; utcLabel="UTC"; printerColorEnabled=$true; printerName="MediaproUS_Canon_Online"
    tabs=[ordered]@{
        CATV=[ordered]@{tabs="1";position="First"}; TVD=[ordered]@{tabs="1";position="First"}
        PASIONES_LATAM=[ordered]@{tabs="5";position="Last"}; PASIONES_US=[ordered]@{tabs="5";position="Last"}
        TODO_NOVELAS=[ordered]@{tabs="5";position="Last"}; REV_TV=[ordered]@{tabs="X";position="First"}
        SPORTYNET=[ordered]@{tabs="X";position="First"}
    }
    layout=[ordered]@{
        CATV=[ordered]@{fontSize="14";defaultRowHeight="25";headerRowHeight="35";smallColumnWidth="7";defaultColumnWidth="37"}
        TVD=[ordered]@{fontSize="14";defaultRowHeight="25";headerRowHeight="35";smallColumnWidth="7";defaultColumnWidth="37"}
        PASIONES_LATAM=[ordered]@{fontSize="14";defaultRowHeight="25";headerRowHeight="35";smallColumnWidth="7";defaultColumnWidth="37"}
        PASIONES_US=[ordered]@{fontSize="14";defaultRowHeight="25";headerRowHeight="35";smallColumnWidth="7";defaultColumnWidth="37"}
        TODO_NOVELAS=[ordered]@{fontSize="14";defaultRowHeight="25";headerRowHeight="35";smallColumnWidth="7";defaultColumnWidth="37"}
        REV_TV=[ordered]@{fontSize="14";defaultRowHeight="25";headerRowHeight="35";smallColumnWidth="7";defaultColumnWidth="37"}
        SPORTYNET=[ordered]@{fontSize="14";defaultRowHeight="25";headerRowHeight="35";smallColumnWidth="7";defaultColumnWidth="37"}
    }
}

function Clone-Defaults { return ($defaults | ConvertTo-Json -Depth 10 | ConvertFrom-Json) }
function Load-Settings {
    if (Test-Path -LiteralPath $jsonPath) {
        try { return (Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json) } catch {}
    }
    return Clone-Defaults
}
function Get-ManagedBlock($data) {
    $lines=[System.Collections.Generic.List[string]]::new()
    [void]$lines.Add($beginMarker)
    [void]$lines.Add('$ManagedSettings = [ordered]@{')
    [void]$lines.Add('    Version             = "2.0.3-dev"')
    [void]$lines.Add(('    Language            = "{0}"' -f $data.language))
    [void]$lines.Add(('    UtcOffset           = {0}' -f [int]$data.utcOffset))
    [void]$lines.Add(('    UtcLabel            = "{0}"' -f ([string]$data.utcLabel).Replace('"','`"')))
    [void]$lines.Add(('    PrinterColorEnabled = ${0}' -f ([string][bool]$data.printerColorEnabled).ToLower()))
    [void]$lines.Add(('    PrinterName         = "{0}"' -f ([string]$data.printerName).Replace('"','`"')))
    [void]$lines.Add('    Tabs = [ordered]@{')
    foreach ($name in @('CATV','TVD','PASIONES_LATAM','PASIONES_US','TODO_NOVELAS','REV_TV','SPORTYNET')) {
        $entry=$data.tabs.$name; $tabs=([string]$entry.tabs).Trim().ToUpper(); $position=[string]$entry.position
        if ($tabs -eq 'X') { $mode='All'; $count=0 } else { $mode=$position; $count=[int]$tabs }
        [void]$lines.Add(('        {0,-16} = [ordered]@{{ Mode = "{1}"; Count = {2} }}' -f $name,$mode,$count))
    }
    [void]$lines.Add('    }')
    [void]$lines.Add('    Layout = [ordered]@{')
    foreach ($name in @('CATV','TVD','PASIONES_LATAM','PASIONES_US','TODO_NOVELAS','REV_TV','SPORTYNET')) {
        $entry=$data.layout.$name
        [void]$lines.Add(('        {0,-16} = [ordered]@{{ FontSize = {1}; DefaultRowHeight = {2}; HeaderRowHeight = {3}; SmallColumnWidth = {4}; DefaultColumnWidth = {5} }}' -f $name,$entry.fontSize,$entry.defaultRowHeight,$entry.headerRowHeight,$entry.smallColumnWidth,$entry.defaultColumnWidth))
    }
    [void]$lines.Add('    }')
    [void]$lines.Add('}')
    [void]$lines.Add($endMarker)
    return [string]::Join("`r`n",$lines.ToArray())
}
function Save-Settings($data) {
    if (-not (Test-Path -LiteralPath $scriptPath)) { throw "FormatGrids.ps1 was not found: $scriptPath" }
    foreach ($name in @('CATV','TVD','PASIONES_LATAM','PASIONES_US','TODO_NOVELAS','REV_TV','SPORTYNET')) {
        $value=([string]$data.tabs.$name.tabs).Trim().ToUpper()
        if ($value -ne 'X') {
            $number=0
            if (-not [int]::TryParse($value,[ref]$number) -or $number -lt 1) { throw "Tabs for $name must be X or a number of 1 or more." }
        }
        # Validate layout settings
        $entry=$data.layout.$name
        if (-not [int]::TryParse($entry.fontSize,[ref]$number) -or $number -lt 8 -or $number -gt 72) { throw "FontSize for $name must be between 8 and 72." }
        if (-not [int]::TryParse($entry.defaultRowHeight,[ref]$number) -or $number -lt 10 -or $number -gt 100) { throw "DefaultRowHeight for $name must be between 10 and 100." }
        if (-not [int]::TryParse($entry.headerRowHeight,[ref]$number) -or $number -lt 10 -or $number -gt 100) { throw "HeaderRowHeight for $name must be between 10 and 100." }
        if (-not [int]::TryParse($entry.smallColumnWidth,[ref]$number) -or $number -lt 1 -or $number -gt 50) { throw "SmallColumnWidth for $name must be between 1 and 50." }
        if (-not [int]::TryParse($entry.defaultColumnWidth,[ref]$number) -or $number -lt 1 -or $number -gt 100) { throw "DefaultColumnWidth for $name must be between 1 and 100." }
    }
    if ([string]::IsNullOrWhiteSpace([string]$data.utcLabel)) { throw "UTC/GMT label cannot be blank." }
    $content=[IO.File]::ReadAllText($scriptPath)
    $start=$content.IndexOf($beginMarker); $end=$content.IndexOf($endMarker)
    if ($start -lt 0 -or $end -lt $start) { throw "Managed settings markers were not found in FormatGrids.ps1." }
    $end += $endMarker.Length
    $block=Get-ManagedBlock $data
    $updated=$content.Substring(0,$start)+$block+$content.Substring($end)
    $backupDir=Join-Path $settingsFolder "Backups"; [void](New-Item -ItemType Directory -Path $backupDir -Force)
    $backup=Join-Path $backupDir ("FormatGrids_{0}.ps1" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    Copy-Item -LiteralPath $scriptPath -Destination $backup -Force
    $temp="$scriptPath.tmp"
    [IO.File]::WriteAllText($temp,$updated,[Text.UTF8Encoding]::new($false))
    $tokens=$null; $errors=$null
    [System.Management.Automation.Language.Parser]::ParseFile($temp,[ref]$tokens,[ref]$errors) | Out-Null
    if ($errors.Count -gt 0) { Remove-Item -LiteralPath $temp -Force; throw ("PowerShell validation failed: " + $errors[0].Message) }
    Move-Item -LiteralPath $temp -Destination $scriptPath -Force

    $aiSource=Join-Path $root "AI\Source\FormatGrids_V2.ps1"
    if (Test-Path -LiteralPath (Split-Path $aiSource -Parent) -PathType Container) {
        try { Copy-Item -LiteralPath $scriptPath -Destination $aiSource -Force } catch {}
    }

    [IO.File]::WriteAllText($jsonPath,($data | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
    $aiLog=Join-Path $root "AI\Logs\settings_changes.log"
    if (Test-Path -LiteralPath (Split-Path $aiLog -Parent)) {
        "$(Get-Date -Format s) language=$($data.language) utc=$($data.utcOffset) label=$($data.utcLabel) printer=$($data.printerName)" | Add-Content -LiteralPath $aiLog -Encoding UTF8
    }
}

$data=Load-Settings
$form=New-Object Windows.Forms.Form
$form.Text="Grid Formatter V2 Settings / Configuracion"
$form.Size=New-Object Drawing.Size(850,700); $form.StartPosition='CenterScreen'; $form.MaximizeBox=$false

# Tab Control for organizing settings
$tabControl = New-Object Windows.Forms.TabControl
$tabControl.Location = '10,10'
$tabControl.Size = '820,600'
$tabControl.TabIndex = 0

# Tab 1: General & Tabs Settings
$tabGeneral = New-Object Windows.Forms.TabPage
$tabGeneral.Text = "General / General"
$tabControl.TabPages.Add($tabGeneral)

# Tab 2: Layout Settings
$tabLayout = New-Object Windows.Forms.TabPage
$tabLayout.Text = "Layout / Diseno"
$tabControl.TabPages.Add($tabLayout)

# ===== General Tab Controls =====
$labelLang=New-Object Windows.Forms.Label; $labelLang.Text='Console language / Idioma:'; $labelLang.Location='10,10'; $labelLang.AutoSize=$true
$comboLang=New-Object Windows.Forms.ComboBox; $comboLang.Location='200,6'; $comboLang.Width=150; $comboLang.DropDownStyle='DropDownList'; [void]$comboLang.Items.AddRange(@('English','Spanish')); $comboLang.SelectedItem=[string]$data.language
$labelOffset=New-Object Windows.Forms.Label; $labelOffset.Text='UTC/GMT offset:'; $labelOffset.Location='10,45'; $labelOffset.AutoSize=$true
$numOffset=New-Object Windows.Forms.NumericUpDown; $numOffset.Location='200,41'; $numOffset.Minimum=-12; $numOffset.Maximum=14; $numOffset.Value=[decimal]$data.utcOffset
$labelUtc=New-Object Windows.Forms.Label; $labelUtc.Text='UTC/GMT label:'; $labelUtc.Location='390,45'; $labelUtc.AutoSize=$true
$textUtc=New-Object Windows.Forms.TextBox; $textUtc.Location='520,41'; $textUtc.Width=180; $textUtc.Text=[string]$data.utcLabel
$checkPrinter=New-Object Windows.Forms.CheckBox; $checkPrinter.Text='Enable automatic printer color setup / Activar color'; $checkPrinter.Location='10,81'; $checkPrinter.Width=360; $checkPrinter.Checked=[bool]$data.printerColorEnabled
$labelPrinter=New-Object Windows.Forms.Label; $labelPrinter.Text='Printer name / Impresora:'; $labelPrinter.Location='390,85'; $labelPrinter.AutoSize=$true
$textPrinter=New-Object Windows.Forms.TextBox; $textPrinter.Location='520,81'; $textPrinter.Width=180; $textPrinter.Text=[string]$data.printerName

$info=New-Object Windows.Forms.Label; $info.Text='Tabs: enter X for all, or a number. Position controls First/Last. Rev TV and SportyNet always filter matching tabs first.'; $info.Location='10,120'; $info.Size='780,35'
$grid=New-Object Windows.Forms.DataGridView; $grid.Location='10,160'; $grid.Size='780,250'; $grid.AllowUserToAddRows=$false; $grid.RowHeadersVisible=$false; $grid.AutoSizeColumnsMode='Fill'
$c1=New-Object Windows.Forms.DataGridViewTextBoxColumn; $c1.Name='GridType'; $c1.HeaderText='Grid type / Tipo'; $c1.ReadOnly=$true
$c2=New-Object Windows.Forms.DataGridViewTextBoxColumn; $c2.Name='Tabs'; $c2.HeaderText='Tabs (X or number)'
$c3=New-Object Windows.Forms.DataGridViewComboBoxColumn; $c3.Name='Position'; $c3.HeaderText='Position / Posicion'; [void]$c3.Items.AddRange(@('First','Last'))
[void]$grid.Columns.AddRange(@($c1,$c2,$c3))
$display=[ordered]@{CATV='CATV';TVD='TVD';PASIONES_LATAM='Pasiones LATAM';PASIONES_US='Pasiones US';TODO_NOVELAS='Todo Novelas';REV_TV='REV TV';SPORTYNET='SportyNet'}
foreach ($name in $display.Keys) { $idx=$grid.Rows.Add($display[$name],[string]$data.tabs.$name.tabs,[string]$data.tabs.$name.position); $grid.Rows[$idx].Tag=$name }

$save=New-Object Windows.Forms.Button; $save.Text='Save Settings / Guardar'; $save.Location='10,440'; $save.Size='210,42'
$reset=New-Object Windows.Forms.Button; $reset.Text='Reset Defaults / Predeterminados'; $reset.Location='230,440'; $reset.Size='230,42'
$close=New-Object Windows.Forms.Button; $close.Text='Close / Cerrar'; $close.Location='470,440'; $close.Size='210,42'

$tabGeneral.Controls.AddRange(@($labelLang,$comboLang,$labelOffset,$numOffset,$labelUtc,$textUtc,$checkPrinter,$labelPrinter,$textPrinter,$info,$grid,$save,$reset,$close))

# ===== Layout Tab Controls =====
$infoLayout=New-Object Windows.Forms.Label; $infoLayout.Text='Set row heights, column widths, and font size per grid type. All values must be positive integers.'; $infoLayout.Location='10,10'; $infoLayout.Size='780,35'
$gridLayout=New-Object Windows.Forms.DataGridView; $gridLayout.Location='10,50'; $gridLayout.Size='780,360'; $gridLayout.AllowUserToAddRows=$false; $gridLayout.RowHeadersVisible=$false; $gridLayout.AutoSizeColumnsMode='Fill'

$l1=New-Object Windows.Forms.DataGridViewTextBoxColumn; $l1.Name='GridType'; $l1.HeaderText='Grid Type'; $l1.ReadOnly=$true; $l1.Width=120
$l2=New-Object Windows.Forms.DataGridViewTextBoxColumn; $l2.Name='FontSize'; $l2.HeaderText='Font'; $l2.Width=60
$l3=New-Object Windows.Forms.DataGridViewTextBoxColumn; $l3.Name='DefaultRowHeight'; $l3.HeaderText='Row H'; $l3.Width=60
$l4=New-Object Windows.Forms.DataGridViewTextBoxColumn; $l4.Name='HeaderRowHeight'; $l4.HeaderText='Header H'; $l4.Width=60
$l5=New-Object Windows.Forms.DataGridViewTextBoxColumn; $l5.Name='SmallColumnWidth'; $l5.HeaderText='Small Col'; $l5.Width=60
$l6=New-Object Windows.Forms.DataGridViewTextBoxColumn; $l6.Name='DefaultColumnWidth'; $l6.HeaderText='Col Width'; $l6.Width=60
[void]$gridLayout.Columns.AddRange(@($l1,$l2,$l3,$l4,$l5,$l6))
foreach ($name in $display.Keys) {
    $entry=$data.layout.$name
    $idx=$gridLayout.Rows.Add($display[$name],$entry.fontSize,$entry.defaultRowHeight,$entry.headerRowHeight,$entry.smallColumnWidth,$entry.defaultColumnWidth)
    $gridLayout.Rows[$idx].Tag=$name
}

$saveLayout=New-Object Windows.Forms.Button; $saveLayout.Text='Save Settings / Guardar'; $saveLayout.Location='10,440'; $saveLayout.Size='210,42'
$resetLayout=New-Object Windows.Forms.Button; $resetLayout.Text='Reset Defaults / Predeterminados'; $resetLayout.Location='230,440'; $resetLayout.Size='230,42
$closeLayout=New-Object Windows.Forms.Button; $closeLayout.Text='Close / Cerrar'; $closeLayout.Location='470,440'; $closeLayout.Size='210,42

$tabLayout.Controls.AddRange(@($infoLayout,$gridLayout,$saveLayout,$resetLayout,$closeLayout))

# ===== Form Controls =====
$form.Controls.Add($tabControl)

# ===== Event Handlers =====
$save.Add_Click({
    try {
        $new=[ordered]@{language=[string]$comboLang.SelectedItem;utcOffset=[int]$numOffset.Value;utcLabel=$textUtc.Text.Trim();printerColorEnabled=[bool]$checkPrinter.Checked;printerName=$textPrinter.Text.Trim();tabs=[ordered]@{};layout=[ordered]@{}}
        foreach ($row in $grid.Rows) { $new.tabs[$row.Tag]=[ordered]@{tabs=([string]$row.Cells['Tabs'].Value).Trim();position=[string]$row.Cells['Position'].Value} }
        foreach ($row in $gridLayout.Rows) {
            $new.layout[$row.Tag]=[ordered]@{
                fontSize=([string]$row.Cells['FontSize'].Value).Trim()
                defaultRowHeight=([string]$row.Cells['DefaultRowHeight'].Value).Trim()
                headerRowHeight=([string]$row.Cells['HeaderRowHeight'].Value).Trim()
                smallColumnWidth=([string]$row.Cells['SmallColumnWidth'].Value).Trim()
                defaultColumnWidth=([string]$row.Cells['DefaultColumnWidth'].Value).Trim()
            }
        }
        Save-Settings ([pscustomobject]$new)
        [Windows.Forms.MessageBox]::Show('Settings saved into FormatGrids.ps1. The formatter remains self-contained. / Configuracion guardada en FormatGrids.ps1.','Grid Formatter V2','OK','Information') | Out-Null
    } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Error','OK','Error') | Out-Null }
})
$reset.Add_Click({
    $d=Clone-Defaults; $comboLang.SelectedItem=$d.language; $numOffset.Value=[decimal]$d.utcOffset; $textUtc.Text=$d.utcLabel; $checkPrinter.Checked=$d.printerColorEnabled; $textPrinter.Text=$d.printerName
    foreach ($row in $grid.Rows) { $name=$row.Tag; $row.Cells['Tabs'].Value=[string]$d.tabs.$name.tabs; $row.Cells['Position'].Value=[string]$d.tabs.$name.position }
})
$close.Add_Click({$form.Close()})

$saveLayout.Add_Click({ & $save.Add_Click.Invoke() })
$resetLayout.Add_Click({
    $d=Clone-Defaults
    foreach ($row in $gridLayout.Rows) {
        $name=$row.Tag
        $row.Cells['FontSize'].Value=[string]$d.layout.$name.fontSize
        $row.Cells['DefaultRowHeight'].Value=[string]$d.layout.$name.defaultRowHeight
        $row.Cells['HeaderRowHeight'].Value=[string]$d.layout.$name.headerRowHeight
        $row.Cells['SmallColumnWidth'].Value=[string]$d.layout.$name.smallColumnWidth
        $row.Cells['DefaultColumnWidth'].Value=[string]$d.layout.$name.defaultColumnWidth
    }
})
$closeLayout.Add_Click({$form.Close()})

[void]$form.ShowDialog()