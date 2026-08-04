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
    [void]$lines.Add('    Version             = "2.0.4-dev"')
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
    # Millisecond precision plus a uniqueness loop: two saves within the same second
    # (or, for scripted callers, the same millisecond) must never collide and silently
    # overwrite an earlier backup -- that would quietly destroy the one thing recovery
    # depends on.
    $stamp=Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $backup=Join-Path $backupDir ("FormatGrids_{0}.ps1" -f $stamp)
    $suffix=1
    while (Test-Path -LiteralPath $backup) {
        $backup=Join-Path $backupDir ("FormatGrids_{0}_{1}.ps1" -f $stamp,$suffix)
        $suffix++
    }
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

function Get-AvailableBackups {
    $backupDir=Join-Path $settingsFolder "Backups"
    if (-not (Test-Path -LiteralPath $backupDir)) { return @() }
    return @(Get-ChildItem -LiteralPath $backupDir -Filter "FormatGrids_*.ps1" -File | Sort-Object LastWriteTime -Descending)
}

# Reconstructs a settings $data object by reading and evaluating the managed
# block out of a given FormatGrids.ps1 (or backup copy of one). Restoring a
# backup therefore reuses Save-Settings end to end -- same validation, same
# automatic backup-of-current-state-first, same AI/Source sync -- instead of
# duplicating that logic for the restore path.
function Read-ManagedBlockAsData([string]$ps1Path) {
    if (-not (Test-Path -LiteralPath $ps1Path)) { throw "File not found: $ps1Path" }
    $content=[IO.File]::ReadAllText($ps1Path)
    $start=$content.IndexOf($beginMarker); $end=$content.IndexOf($endMarker)
    if ($start -lt 0 -or $end -lt $start) { throw "Managed settings markers were not found in $(Split-Path $ps1Path -Leaf)." }
    $end += $endMarker.Length
    $blockLines = $content.Substring($start,$end-$start) -split "`r?`n" | Where-Object { $_ -ne $beginMarker -and $_ -ne $endMarker }
    $code = ($blockLines -join "`r`n") + "`r`n`$ManagedSettings"
    $tokens=$null; $errors=$null
    [System.Management.Automation.Language.Parser]::ParseInput($code,[ref]$tokens,[ref]$errors) | Out-Null
    if ($errors.Count -gt 0) { throw ("Backup failed validation: " + $errors[0].Message) }
    $ms = & ([scriptblock]::Create($code))

    $result=[ordered]@{
        language=[string]$ms.Language; utcOffset=[int]$ms.UtcOffset; utcLabel=[string]$ms.UtcLabel
        printerColorEnabled=[bool]$ms.PrinterColorEnabled; printerName=[string]$ms.PrinterName
        tabs=[ordered]@{}; layout=[ordered]@{}
    }
    foreach ($name in @('CATV','TVD','PASIONES_LATAM','PASIONES_US','TODO_NOVELAS','REV_TV','SPORTYNET')) {
        $t=$ms.Tabs.$name
        if ($null -eq $t) { throw "Backup is missing Tabs.$name -- likely from an older format, cannot restore safely." }
        if ([string]$t.Mode -eq 'All') { $tabsVal='X'; $posVal='First' } else { $tabsVal=[string]$t.Count; $posVal=[string]$t.Mode }
        $result.tabs[$name]=[ordered]@{tabs=$tabsVal;position=$posVal}
        $l=$ms.Layout.$name
        if ($null -eq $l) { throw "Backup is missing Layout.$name -- likely from an older format, cannot restore safely." }
        $result.layout[$name]=[ordered]@{
            fontSize=[string]$l.FontSize; defaultRowHeight=[string]$l.DefaultRowHeight
            headerRowHeight=[string]$l.HeaderRowHeight; smallColumnWidth=[string]$l.SmallColumnWidth
            defaultColumnWidth=[string]$l.DefaultColumnWidth
        }
    }
    return [pscustomobject]$result
}

$script:UiLanguage = 'English'
$script:data = $null
$script:Ui = @{}

$uiErrorFolder = Join-Path $root 'AI\Logs'
if (-not (Test-Path -LiteralPath $uiErrorFolder -PathType Container)) { $uiErrorFolder = $settingsFolder }
$uiErrorLog = Join-Path $uiErrorFolder 'settings_ui_errors.log.txt'

function Get-UiText([string]$key) {
    $english = @{
        FormTitle='Grid Formatter V2 Settings'; LanguageSwitch='Español'; GeneralGroup='General Settings'
        OutputLanguage='PowerShell display language'; Offset='UTC/GMT offset'; UtcLabel='UTC/GMT label'
        PrinterEnabled='Enable automatic printer color setup'; PrinterName='Printer name'
        TabsGroup='Worksheets to Format'; TabsInfo='Enter X to format all matching worksheets, or enter a number and choose First/Last.'
        FixedLayout='Formatting sizes are fixed per grid type and are not editable in this simplified V1 settings screen.'
        GridType='Grid type'; Tabs='Tabs (X or number)'; Position='Position (First/Last)'
        Save='Save Settings'; Reset='Reset Defaults'; Restore='Restore Backup'; Close='Close'
        Saved='Settings were saved into FormatGrids.ps1. The formatter remains self-contained.'
        ResetConfirm='Reset the visible settings to their defaults?'; NoBackups='No backups exist yet. A backup is created automatically whenever settings are saved.'
        Restored='Settings restored from'; PickerTitle='Restore Backup'; PickerLabel='Choose a backup to restore:'
        PickerRestore='Restore'; PickerCancel='Cancel'; ErrorTitle='Settings Error'
    }
    $spanish = @{
        FormTitle='Configuración del Formateador de Grillas V2'; LanguageSwitch='English'; GeneralGroup='Configuración General'
        OutputLanguage='Idioma mostrado por PowerShell'; Offset='Diferencia UTC/GMT'; UtcLabel='Etiqueta UTC/GMT'
        PrinterEnabled='Activar configuración automática de color'; PrinterName='Nombre de la impresora'
        TabsGroup='Hojas que se Formatearán'; TabsInfo='Escriba X para todas las hojas compatibles, o un número y seleccione First/Last.'
        FixedLayout='Los tamaños de formato son fijos para cada tipo de grilla y no se editan en esta pantalla simplificada V1.'
        GridType='Tipo de grilla'; Tabs='Hojas (X o número)'; Position='Posición (First/Last)'
        Save='Guardar Configuración'; Reset='Restablecer Valores'; Restore='Restaurar Respaldo'; Close='Cerrar'
        Saved='La configuración se guardó dentro de FormatGrids.ps1. El formateador continúa siendo independiente.'
        ResetConfirm='¿Restablecer la configuración visible a sus valores predeterminados?'; NoBackups='Todavía no existen respaldos. Se crea uno automáticamente cada vez que se guarda.'
        Restored='Configuración restaurada desde'; PickerTitle='Restaurar Respaldo'; PickerLabel='Seleccione un respaldo para restaurar:'
        PickerRestore='Restaurar'; PickerCancel='Cancelar'; ErrorTitle='Error de Configuración'
    }
    if ($script:UiLanguage -eq 'Spanish') { return [string]$spanish[$key] }
    return [string]$english[$key]
}

function Write-SettingsUiError($errorRecord) {
    try {
        $lines = @(
            ('=' * 72),
            (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
            "Message: $($errorRecord.Exception.Message)",
            "Type: $($errorRecord.Exception.GetType().FullName)",
            "Stack: $($errorRecord.ScriptStackTrace)",
            "Invocation: $($errorRecord.InvocationInfo.PositionMessage)"
        )
        $lines | Add-Content -LiteralPath $uiErrorLog -Encoding UTF8
    } catch {}
}

function Show-UiError($errorRecord) {
    Write-SettingsUiError $errorRecord
    [Windows.Forms.MessageBox]::Show($errorRecord.Exception.Message,(Get-UiText 'ErrorTitle'),'OK','Error') | Out-Null
}

function Show-BackupPicker($backups) {
    $picker=New-Object Windows.Forms.Form
    $picker.Text=Get-UiText 'PickerTitle'
    $picker.Size=New-Object Drawing.Size(460,390); $picker.StartPosition='CenterParent'
    $picker.FormBorderStyle='FixedDialog'; $picker.MaximizeBox=$false; $picker.MinimizeBox=$false
    $lbl=New-Object Windows.Forms.Label; $lbl.Text=Get-UiText 'PickerLabel'; $lbl.Location='15,15'; $lbl.Size='415,24'
    $list=New-Object Windows.Forms.ListBox; $list.Location='15,45'; $list.Size='415,250'
    foreach ($b in $backups) { [void]$list.Items.Add(('{0:yyyy-MM-dd HH:mm:ss}   {1}' -f $b.LastWriteTime,$b.Name)) }
    if ($list.Items.Count -gt 0) { $list.SelectedIndex=0 }
    $ok=New-Object Windows.Forms.Button; $ok.Text=Get-UiText 'PickerRestore'; $ok.Location='125,310'; $ok.Size='145,36'; $ok.DialogResult='OK'
    $cancel=New-Object Windows.Forms.Button; $cancel.Text=Get-UiText 'PickerCancel'; $cancel.Location='285,310'; $cancel.Size='145,36'; $cancel.DialogResult='Cancel'
    $picker.Controls.Add($lbl); $picker.Controls.Add($list); $picker.Controls.Add($ok); $picker.Controls.Add($cancel)
    $picker.AcceptButton=$ok; $picker.CancelButton=$cancel
    $picker.Add_Shown({ $list.Focus() })
    $result=$picker.ShowDialog()
    if ($result -eq 'OK' -and $list.SelectedIndex -ge 0) { return $backups[$list.SelectedIndex] }
    return $null
}

function Get-CurrentFormData {
    $new=[ordered]@{
        language=[string]$script:Ui.ComboLanguage.SelectedItem
        utcOffset=[int]$script:Ui.Offset.Value
        utcLabel=$script:Ui.UtcLabel.Text.Trim()
        printerColorEnabled=[bool]$script:Ui.PrinterEnabled.Checked
        printerName=$script:Ui.PrinterName.Text.Trim()
        tabs=[ordered]@{}
        # Layout values remain preserved for backward compatibility, but the
        # simplified V1 UI no longer exposes misleading generic dimensions.
        layout=$script:data.layout
    }
    foreach ($row in $script:Ui.Grid.Rows) {
        $name=[string]$row.Tag
        $new.tabs[$name]=[ordered]@{
            tabs=([string]$row.Cells['Tabs'].Value).Trim()
            position=[string]$row.Cells['Position'].Value
        }
    }
    return [pscustomobject]$new
}

function Refresh-FormFromData($d) {
    $script:Ui.ComboLanguage.SelectedItem=[string]$d.language
    $script:Ui.Offset.Value=[decimal]$d.utcOffset
    $script:Ui.UtcLabel.Text=[string]$d.utcLabel
    $script:Ui.PrinterEnabled.Checked=[bool]$d.printerColorEnabled
    $script:Ui.PrinterName.Text=[string]$d.printerName
    foreach ($row in $script:Ui.Grid.Rows) {
        $name=[string]$row.Tag
        $row.Cells['Tabs'].Value=[string]$d.tabs.$name.tabs
        $row.Cells['Position'].Value=[string]$d.tabs.$name.position
    }
}

function Apply-UiLanguage {
    $script:Ui.Form.Text=Get-UiText 'FormTitle'
    $script:Ui.LanguageSwitch.Text=Get-UiText 'LanguageSwitch'
    $script:Ui.GeneralGroup.Text=Get-UiText 'GeneralGroup'
    $script:Ui.LabelLanguage.Text=Get-UiText 'OutputLanguage'
    $script:Ui.LabelOffset.Text=Get-UiText 'Offset'
    $script:Ui.LabelUtc.Text=Get-UiText 'UtcLabel'
    $script:Ui.PrinterEnabled.Text=Get-UiText 'PrinterEnabled'
    $script:Ui.LabelPrinter.Text=Get-UiText 'PrinterName'
    $script:Ui.TabsGroup.Text=Get-UiText 'TabsGroup'
    $script:Ui.TabsInfo.Text=Get-UiText 'TabsInfo'
    $script:Ui.FixedLayout.Text=Get-UiText 'FixedLayout'
    $script:Ui.Grid.Columns['GridType'].HeaderText=Get-UiText 'GridType'
    $script:Ui.Grid.Columns['Tabs'].HeaderText=Get-UiText 'Tabs'
    $script:Ui.Grid.Columns['Position'].HeaderText=Get-UiText 'Position'
    $script:Ui.Save.Text=Get-UiText 'Save'
    $script:Ui.Reset.Text=Get-UiText 'Reset'
    $script:Ui.Restore.Text=Get-UiText 'Restore'
    $script:Ui.Close.Text=Get-UiText 'Close'
}

function Start-SettingsUi {
    $script:data=Load-Settings
    $form=New-Object Windows.Forms.Form
    $form.Size=New-Object Drawing.Size(880,650); $form.StartPosition='CenterScreen'
    $form.FormBorderStyle='FixedDialog'; $form.MaximizeBox=$false; $form.MinimizeBox=$true
    $form.Font=New-Object Drawing.Font('Segoe UI',9)

    $switch=New-Object Windows.Forms.Button; $switch.Location='735,15'; $switch.Size='105,34'

    $general=New-Object Windows.Forms.GroupBox; $general.Location='20,55'; $general.Size='820,145'
    $labelLang=New-Object Windows.Forms.Label; $labelLang.Location='20,33'; $labelLang.Size='210,24'
    $comboLang=New-Object Windows.Forms.ComboBox; $comboLang.Location='245,29'; $comboLang.Size='150,28'; $comboLang.DropDownStyle='DropDownList'
    [void]$comboLang.Items.Add('English'); [void]$comboLang.Items.Add('Spanish')
    $labelOffset=New-Object Windows.Forms.Label; $labelOffset.Location='430,33'; $labelOffset.Size='145,24'
    $numOffset=New-Object Windows.Forms.NumericUpDown; $numOffset.Location='590,29'; $numOffset.Size='80,28'; $numOffset.Minimum=-12; $numOffset.Maximum=14
    $labelUtc=New-Object Windows.Forms.Label; $labelUtc.Location='20,76'; $labelUtc.Size='210,24'
    $textUtc=New-Object Windows.Forms.TextBox; $textUtc.Location='245,72'; $textUtc.Size='150,27'
    $checkPrinter=New-Object Windows.Forms.CheckBox; $checkPrinter.Location='430,75'; $checkPrinter.Size='360,25'
    $labelPrinter=New-Object Windows.Forms.Label; $labelPrinter.Location='20,110'; $labelPrinter.Size='210,24'
    $textPrinter=New-Object Windows.Forms.TextBox; $textPrinter.Location='245,106'; $textPrinter.Size='545,27'
    foreach ($control in @($labelLang,$comboLang,$labelOffset,$numOffset,$labelUtc,$textUtc,$checkPrinter,$labelPrinter,$textPrinter)) { $general.Controls.Add($control) }

    $tabsGroup=New-Object Windows.Forms.GroupBox; $tabsGroup.Location='20,210'; $tabsGroup.Size='820,315'
    $info=New-Object Windows.Forms.Label; $info.Location='20,28'; $info.Size='775,38'
    $grid=New-Object Windows.Forms.DataGridView; $grid.Location='20,70'; $grid.Size='775,205'
    $grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.RowHeadersVisible=$false
    $grid.AutoSizeColumnsMode='None'; $grid.SelectionMode='CellSelect'; $grid.MultiSelect=$false
    $grid.RowTemplate.Height=28
    $c1=New-Object Windows.Forms.DataGridViewTextBoxColumn; $c1.Name='GridType'; $c1.ReadOnly=$true; $c1.Width=300
    $c2=New-Object Windows.Forms.DataGridViewTextBoxColumn; $c2.Name='Tabs'; $c2.Width=190
    $c3=New-Object Windows.Forms.DataGridViewComboBoxColumn; $c3.Name='Position'; $c3.Width=240
    [void]$c3.Items.Add('First'); [void]$c3.Items.Add('Last')
    $null=$grid.Columns.Add($c1); $null=$grid.Columns.Add($c2); $null=$grid.Columns.Add($c3)
    $display=[ordered]@{CATV='CATV';TVD='TVD';PASIONES_LATAM='Pasiones LATAM';PASIONES_US='Pasiones US';TODO_NOVELAS='Todo Novelas';REV_TV='REV TV';SPORTYNET='SportyNet'}
    foreach ($name in $display.Keys) {
        $idx=$grid.Rows.Add($display[$name],[string]$script:data.tabs.$name.tabs,[string]$script:data.tabs.$name.position)
        if ($idx -lt 0) { throw "Could not create settings row for $name." }
        $grid.Rows[$idx].Tag=$name
    }
    $fixedLayout=New-Object Windows.Forms.Label; $fixedLayout.Location='20,282'; $fixedLayout.Size='775,26'; $fixedLayout.ForeColor=[Drawing.Color]::DimGray
    $tabsGroup.Controls.Add($info); $tabsGroup.Controls.Add($grid); $tabsGroup.Controls.Add($fixedLayout)

    $save=New-Object Windows.Forms.Button; $save.Location='20,545'; $save.Size='185,42'
    $reset=New-Object Windows.Forms.Button; $reset.Location='220,545'; $reset.Size='185,42'
    $restore=New-Object Windows.Forms.Button; $restore.Location='420,545'; $restore.Size='185,42'
    $close=New-Object Windows.Forms.Button; $close.Location='620,545'; $close.Size='220,42'

    $form.Controls.Add($switch); $form.Controls.Add($general); $form.Controls.Add($tabsGroup)
    $form.Controls.Add($save); $form.Controls.Add($reset); $form.Controls.Add($restore); $form.Controls.Add($close)
    $form.AcceptButton=$save; $form.CancelButton=$close

    $script:Ui=@{
        Form=$form; LanguageSwitch=$switch; GeneralGroup=$general; LabelLanguage=$labelLang
        ComboLanguage=$comboLang; LabelOffset=$labelOffset; Offset=$numOffset; LabelUtc=$labelUtc
        UtcLabel=$textUtc; PrinterEnabled=$checkPrinter; LabelPrinter=$labelPrinter; PrinterName=$textPrinter
        TabsGroup=$tabsGroup; TabsInfo=$info; Grid=$grid; FixedLayout=$fixedLayout
        Save=$save; Reset=$reset; Restore=$restore; Close=$close
    }
    Refresh-FormFromData $script:data
    Apply-UiLanguage

    $switch.Add_Click({
        if ($script:UiLanguage -eq 'English') { $script:UiLanguage='Spanish' } else { $script:UiLanguage='English' }
        Apply-UiLanguage
    })
    $save.Add_Click({
        try {
            $new=Get-CurrentFormData
            Save-Settings $new
            $script:data=$new
            [Windows.Forms.MessageBox]::Show((Get-UiText 'Saved'),(Get-UiText 'FormTitle'),'OK','Information') | Out-Null
        } catch { Show-UiError $_ }
    })
    $reset.Add_Click({
        $result=[Windows.Forms.MessageBox]::Show((Get-UiText 'ResetConfirm'),(Get-UiText 'FormTitle'),'YesNo','Question')
        if ($result -eq 'Yes') { $script:data=Clone-Defaults; Refresh-FormFromData $script:data }
    })
    $restore.Add_Click({
        try {
            $backups=Get-AvailableBackups
            if ($backups.Count -eq 0) {
                [Windows.Forms.MessageBox]::Show((Get-UiText 'NoBackups'),(Get-UiText 'FormTitle'),'OK','Information') | Out-Null
                return
            }
            $chosen=Show-BackupPicker $backups
            if ($null -eq $chosen) { return }
            $restoredData=Read-ManagedBlockAsData $chosen.FullName
            Save-Settings $restoredData
            $script:data=$restoredData
            Refresh-FormFromData $script:data
            [Windows.Forms.MessageBox]::Show("$(Get-UiText 'Restored'): $($chosen.Name)",(Get-UiText 'FormTitle'),'OK','Information') | Out-Null
        } catch { Show-UiError $_ }
    })
    $close.Add_Click({ $form.Close() })

    if ($env:GRID_SETTINGS_UI_SELFTEST -eq '1') {
        if ($grid.Columns.Count -ne 3) { throw "Settings UI expected 3 columns, found $($grid.Columns.Count)." }
        if ($grid.Rows.Count -ne 7) { throw "Settings UI expected 7 grid rows, found $($grid.Rows.Count)." }
        if ($save.Text -ne 'Save Settings' -or $reset.Text -ne 'Reset Defaults' -or $restore.Text -ne 'Restore Backup' -or $close.Text -ne 'Close') {
            throw 'Settings UI did not start with English action buttons.'
        }
        $script:UiLanguage='Spanish'; Apply-UiLanguage
        if ($save.Text -ne 'Guardar Configuración' -or $close.Text -ne 'Cerrar') { throw 'Settings UI Spanish switch failed.' }
        $script:UiLanguage='English'; Apply-UiLanguage
        Write-Output 'SETTINGS_UI_SELFTEST_PASS'
        return
    }

    [void]$form.ShowDialog()
}

try { Start-SettingsUi }
catch {
    Write-SettingsUiError $_
    [Windows.Forms.MessageBox]::Show($_.Exception.Message,(Get-UiText 'ErrorTitle'),'OK','Error') | Out-Null
}
