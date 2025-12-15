#region Imports & Config
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $exe = Join-Path $PSHome 'powershell.exe'
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    Start-Process -FilePath $exe -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$scriptPath`""
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.ComponentModel
try {
    Add-Type -AssemblyName 'Microsoft.VisualBasic' -ErrorAction SilentlyContinue
} catch {}

$ScriptRootPath = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $PSScriptRoot) { $PSScriptRoot = $ScriptRootPath }
$modulesRoot = Join-Path $ScriptRootPath 'Modules'

. (Join-Path $modulesRoot 'Config.ps1') -ScriptRoot $ScriptRootPath
. (Join-Path $modulesRoot 'RuleBank.ps1')  # RuleBank root object (all assays)
. (Join-Path $modulesRoot 'RuleEngine.ps1')  # Pure helpers to read RuleBank

# --- Enkel loggning (ersätter saknad Logging-modul) ---
$script:LogControl = $null
function Set-LogOutputControl { param($Control) $script:LogControl = $Control }
function Gui-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Warn','Error')]$Level = 'Info',
        [switch]$Immediate
    )
    $ts = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts][$Level] $Message"
    if ($script:LogControl) { try { $script:LogControl.AppendText($line + [Environment]::NewLine) } catch {} }
    try { Add-Content -Path $global:LogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
    if ($Immediate) { try { [System.Windows.Forms.Application]::DoEvents() } catch {} }
}
function Add-AuditEntry {
    param(
        [string]$Lsp,
        [string]$Assay,
        [string]$BatchNumber,
        [int]$TestCount,
        [string]$Status,
        [string]$ReportPath
    )
    try {
        $line = ("{0};{1};{2};{3};{4};{5};{6}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $env:USERNAME, $Lsp, $Assay, $BatchNumber, $TestCount, $Status)
        Add-Content -Path $global:LogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}

. (Join-Path $modulesRoot 'Splash.ps1')
. (Join-Path $modulesRoot 'UiStyling.ps1')
. (Join-Path $modulesRoot 'DataHelpers.ps1')
. (Join-Path $modulesRoot 'SignatureHelpers.ps1')

$global:StartupReady = $true
$configStatus = $null

try {

    $configStatus = Test-Config
    if ($configStatus) {
        foreach ($err in $configStatus.Errors) { Gui-Log "❌ Konfig-fel: $err" 'Error' }
        foreach ($warn in $configStatus.Warnings) { Gui-Log "⚠️ Konfig-varning: $warn" 'Warn' }
        if (-not $configStatus.Ok) {
            $global:StartupReady = $false
            try { [System.Windows.Forms.MessageBox]::Show("Startkontroll misslyckades. Se logg för detaljer.","Startkontroll") | Out-Null } catch {}
        }
    }
} catch { Gui-Log "❌ Test-Config misslyckades: $($_.Exception.Message)" 'Error'; $global:StartupReady = $false }

$Host.UI.RawUI.WindowTitle = "Startar…"
Show-Splash "Laddar PnP.PowerShell…"
$global:SpConnected = $false
$global:SpError     = $null

try {
    $null = Get-PackageProvider -Name "NuGet" -ForceBootstrap -ErrorAction SilentlyContinue
} catch {}

try {
    Update-Splash "Laddar..."
    Import-Module PnP.PowerShell -ErrorAction Stop
} catch {
    try {
        Gui-Log "ℹ️ PowerShell-modulen saknas, installerar modulen..." 'Info'
        Install-Module PnP.PowerShell -MaximumVersion 1.12.0 -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Update-Splash "Laddar..."
        Import-Module PnP.PowerShell -ErrorAction Stop
    } catch {
        $global:SpError = "PnP-install/import misslyckades: $($_.Exception.Message)"
    }
}
$env:PNPPOWERSHELL_UPDATECHECK = "Off"
try { $null = Ensure-EPPlus -Version '4.5.3.3' } catch { Gui-Log "⚠️ EPPlus-förkontroll misslyckades: $($_.Exception.Message)" 'Warn' }

if (-not $global:SpError) {
    try {
        Update-Splash "Ansluter till SharePoint"
        Connect-PnPOnline -Url $global:SP_SiteUrl `
                          -Tenant $global:SP_Tenant `
                          -ClientId $global:SP_ClientId `
                          -CertificateBase64Encoded $global:SP_CertBase64 `
                          -ErrorAction Stop
        $global:SpConnected = $true
    } catch {
        $msg = "Connect-PnPOnline misslyckades: $($_.Exception.Message)"
        Update-Splash $msg
        $global:SpError = $msg
    }
}

#endregion Imports & Config

#region GUI Construction

Update-Splash "Startar gränssnitt…"
Close-Splash
$form = New-Object System.Windows.Forms.Form
$form.Text = "$ScriptVersion"
$form.AutoScaleMode = 'Dpi'
$form.Size = New-Object System.Drawing.Size(860,910)
$form.MinimumSize = New-Object System.Drawing.Size(860,910)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.AutoScroll  = $false
$form.MaximizeBox = $false
$form.Padding     = New-Object System.Windows.Forms.Padding(8)
$form.Font        = New-Object System.Drawing.Font('Segoe UI',10)
$form.KeyPreview = $true
$form.add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $form.Close() } })

# ---------- Menyrad ----------
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.Dock='Top'; $menuStrip.GripStyle='Hidden'
$menuStrip.ImageScalingSize = New-Object System.Drawing.Size(20,20)
$menuStrip.Padding = New-Object System.Windows.Forms.Padding(8,6,0,6)
$menuStrip.Font = New-Object System.Drawing.Font('Segoe UI',10)
$miArkiv = New-Object System.Windows.Forms.ToolStripMenuItem('🗂️ Arkiv')
$miScan  = New-Object System.Windows.Forms.ToolStripMenuItem('🔍 Sök filer')
$miBuild = New-Object System.Windows.Forms.ToolStripMenuItem('✅ Skapa rapport')
$miClear = New-Object System.Windows.Forms.ToolStripMenuItem('🧹 Rensa')
$miExit  = New-Object System.Windows.Forms.ToolStripMenuItem('❌ Avsluta')

# Arkiv menyn med endast arbetsflödesåtgärder (ingen extra tooling)
$miArkiv.DropDownItems.Clear()
$miArkiv.DropDownItems.AddRange(@(
    $miScan,
    $miBuild,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $miClear,
    $miExit
))

$menuStrip.Items.Clear()
$menuStrip.Items.Add($miArkiv) | Out-Null
$form.MainMenuStrip=$menuStrip

# ---------- Header ----------
$panelHeader = New-Object System.Windows.Forms.Panel
$panelHeader.Dock='Top'; $panelHeader.Height=64
$panelHeader.BackColor=[System.Drawing.Color]::SteelBlue
$panelHeader.Padding = New-Object System.Windows.Forms.Padding(10,8,10,8)

$picLogo = New-Object System.Windows.Forms.PictureBox
$picLogo.Dock='Left'; $picLogo.Width=50; $picLogo.BorderStyle='FixedSingle'
if(Test-Path $ikonSokvag){ $picLogo.Image=[System.Drawing.Image]::FromFile($ikonSokvag); $picLogo.SizeMode='Zoom' }

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text="$ScriptVersion - Skapa excelrapport för en lot."
$lblTitle.ForeColor=[System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold',13)
$lblTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblTitle.Padding = New-Object System.Windows.Forms.Padding(8,0,0,0)
$lblTitle.Dock='Fill'

$panelHeader.Controls.Add($lblTitle)
$panelHeader.Controls.Add($picLogo)

# ---------- Sök-rad ----------

$tlSearch = New-Object System.Windows.Forms.TableLayoutPanel
$tlSearch.Dock='Top'; $tlSearch.AutoSize=$true; $tlSearch.AutoSizeMode='GrowAndShrink'
$tlSearch.Padding = New-Object System.Windows.Forms.Padding(0,10,0,8)
$tlSearch.ColumnCount=3
[void]$tlSearch.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$tlSearch.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
[void]$tlSearch.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,130)))

$lblLSP = New-Object System.Windows.Forms.Label
$lblLSP.Text='LSP:'; $lblLSP.Anchor='Left'; $lblLSP.AutoSize=$true
$lblLSP.Margin = New-Object System.Windows.Forms.Padding(0,6,8,0)
$txtLSP = New-Object System.Windows.Forms.TextBox
$txtLSP.Dock='Fill'
$txtLSP.Margin = New-Object System.Windows.Forms.Padding(0,2,10,2)
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text='Sök filer'; $btnScan.Dock='Fill'; Set-AccentButton $btnScan -Primary
$btnScan.Margin= New-Object System.Windows.Forms.Padding(0,2,0,2)

$tlSearch.Controls.Add($lblLSP,0,0)
$tlSearch.Controls.Add($txtLSP,1,0)
$tlSearch.Controls.Add($btnScan,2,0)

$pLog = New-Object System.Windows.Forms.Panel
$pLog.Dock='Top'; $pLog.Height=220; $pLog.Padding=New-Object System.Windows.Forms.Padding(0,0,0,8)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline=$true; $outputBox.ScrollBars='Vertical'; $outputBox.ReadOnly=$true
$outputBox.BackColor='White'; $outputBox.Dock='Fill'
$outputBox.Font = New-Object System.Drawing.Font('Segoe UI',9)
$pLog.Controls.Add($outputBox)
try { Set-LogOutputControl -Control $outputBox } catch {}

$grpPick = New-Object System.Windows.Forms.GroupBox
$grpPick.Text='Välj filer för rapport'
$grpPick.Dock='Top'
$grpPick.Padding = New-Object System.Windows.Forms.Padding(10,12,10,14)
$grpPick.AutoSize=$false
$grpPick.Height = (78*3) + $grpPick.Padding.Top + $grpPick.Padding.Bottom +15

$tlPick = New-Object System.Windows.Forms.TableLayoutPanel
$tlPick.Dock='Fill'; $tlPick.ColumnCount=3; $tlPick.RowCount=3
$tlPick.GrowStyle=[System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize
[void]$tlPick.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$tlPick.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
[void]$tlPick.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,100)))
for($i=0;$i -lt 3;$i++){ [void]$tlPick.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,78))) }

function New-ListRow {
    param([string]$labelText,[ref]$lbl,[ref]$clb,[ref]$btn)
    $lbl.Value = New-Object System.Windows.Forms.Label
    $lbl.Value.Text=$labelText
    $lbl.Value.Anchor='Left'
    $lbl.Value.AutoSize=$true
    $lbl.Value.Margin=New-Object System.Windows.Forms.Padding(0,12,6,0)
    $clb.Value = New-Object System.Windows.Forms.CheckedListBox
    $clb.Value.Dock='Fill'
    $clb.Value.Margin=New-Object System.Windows.Forms.Padding(0,6,8,6)
    $clb.Value.Height=70
    $clb.Value.IntegralHeight=$false
    $clb.Value.CheckOnClick = $true
    $clb.Value.DisplayMember = 'Name'

    $btn.Value = New-Object System.Windows.Forms.Button
    $btn.Value.Text='Bläddra…'
    $btn.Value.Dock='Fill'
    $btn.Value.Margin=New-Object System.Windows.Forms.Padding(0,6,0,6)
    Set-AccentButton $btn.Value
}

$lblCsv=$null;$clbCsv=$null;$btnCsvBrowse=$null
New-ListRow -labelText 'CSV-fil:' -lbl ([ref]$lblCsv) -clb ([ref]$clbCsv) -btn ([ref]$btnCsvBrowse)
$lblNeg=$null;$clbNeg=$null;$btnNegBrowse=$null
New-ListRow -labelText 'Seal Test Neg:' -lbl ([ref]$lblNeg) -clb ([ref]$clbNeg) -btn ([ref]$btnNegBrowse)
$lblPos=$null;$clbPos=$null;$btnPosBrowse=$null
New-ListRow -labelText 'Seal Test Pos:' -lbl ([ref]$lblPos) -clb ([ref]$clbPos) -btn ([ref]$btnPosBrowse)

try {
    if ($tlPick.RowCount -lt 4) {
        $tlPick.RowCount = 4
        for ($i=$tlPick.RowStyles.Count; $i -lt 4; $i++) {
            $null = $tlPick.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 78)) )
        }
        $grpPick.Height = (78*4) + $grpPick.Padding.Top + $grpPick.Padding.Bottom + 15
    }
} catch {}

$lblLsp = $null; $clbLsp = $null; $btnLspBrowse = $null
New-ListRow -labelText 'Worksheet:' -lbl ([ref]$lblLsp) -clb ([ref]$clbLsp) -btn ([ref]$btnLspBrowse)

$tlPick.Controls.Add($lblLsp,  0, 3)
$tlPick.Controls.Add($clbLsp,  1, 3)
$tlPick.Controls.Add($btnLspBrowse, 2, 3)

$clbLsp.add_ItemCheck({
    param($s,$e)
    if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        for ($i=0; $i -lt $s.Items.Count; $i++) {
            if ($i -ne $e.Index) { $s.SetItemChecked($i, $false) }
        }
    }
})

$btnLspBrowse.Add_Click({
    try {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Excel|*.xlsx;*.xlsm|Alla filer|*.*"
        $dlg.Title  = "Välj LSP Worksheet"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $f = Get-Item -LiteralPath $dlg.FileName
            Add-CLBItems -clb $clbLsp -files @($f) -AutoCheckFirst
            if (Get-Command Update-StatusBar -ErrorAction SilentlyContinue) { Update-StatusBar }
        }
    } catch {
        Gui-Log ("⚠️ LSP-browse fel: " + $_.Exception.Message) 'Warn'
    }
})

# Lägg in i tabellen
$tlPick.Controls.Add($lblCsv,0,0); $tlPick.Controls.Add($clbCsv,1,0); $tlPick.Controls.Add($btnCsvBrowse,2,0)
$tlPick.Controls.Add($lblNeg,0,1); $tlPick.Controls.Add($clbNeg,1,1); $tlPick.Controls.Add($btnNegBrowse,2,1)
$tlPick.Controls.Add($lblPos,0,2); $tlPick.Controls.Add($clbPos,1,2); $tlPick.Controls.Add($btnPosBrowse,2,2)
$grpPick.Controls.Add($tlPick)

# ---------- Signatur ----------
$grpSign = New-Object System.Windows.Forms.GroupBox
$grpSign.Text = "Lägg till signatur i Seal Test-filerna"
$grpSign.Dock='Top'
$grpSign.Padding = New-Object System.Windows.Forms.Padding(10,8,10,10)
$grpSign.AutoSize = $false
$grpSign.Height = 88

$tlSign = New-Object System.Windows.Forms.TableLayoutPanel
$tlSign.Dock='Fill'; $tlSign.ColumnCount=2; $tlSign.RowCount=2
[void]$tlSign.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$tlSign.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
[void]$tlSign.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,28)))
[void]$tlSign.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,28)))

$lblSigner = New-Object System.Windows.Forms.Label
$lblSigner.Text = 'Fullständigt namn, signatur och datum:'
$lblSigner.Anchor='Left'; $lblSigner.AutoSize=$true

$txtSigner = New-Object System.Windows.Forms.TextBox
$txtSigner.Dock='Fill'; $txtSigner.Margin = New-Object System.Windows.Forms.Padding(6,2,0,2)
$chkWriteSign = New-Object System.Windows.Forms.CheckBox
$chkWriteSign.Text = 'Signera Seal Test-Filerna'
$chkWriteSign.Anchor='Left'
$chkWriteSign.AutoSize = $true

$chkOverwriteSign = New-Object System.Windows.Forms.CheckBox
$chkOverwriteSign.Text = 'Aktivera'

$chkOverwriteSign.Anchor='Left'
$chkOverwriteSign.AutoSize = $true
$chkOverwriteSign.Enabled = $false
$chkWriteSign.add_CheckedChanged({ $chkOverwriteSign.Enabled = $chkWriteSign.Checked })

$tlSign.Controls.Add($lblSigner,0,0); $tlSign.Controls.Add($txtSigner,1,0)
$tlSign.Controls.Add($chkWriteSign,0,1); $tlSign.Controls.Add($chkOverwriteSign,1,1)
$grpSign.Controls.Add($tlSign)

$grpSign.Visible = $true

# ---------- Primärknapp ----------
$btnBuild = New-Object System.Windows.Forms.Button
$btnBuild.Text='Skapa rapport'; $btnBuild.Dock='Top'; $btnBuild.Height=40
$btnBuild.Margin = New-Object System.Windows.Forms.Padding(0,16,0,8)
$btnBuild.Enabled=$false; Set-AccentButton $btnBuild -Primary

# ---------- Statusrad ----------
$status = New-Object System.Windows.Forms.StatusStrip
$status.SizingGrip=$false; $status.Dock='Bottom'; $status.Font=New-Object System.Drawing.Font('Segoe UI',9)
$status.ShowItemToolTips = $true
$slCount = New-Object System.Windows.Forms.ToolStripStatusLabel; $slCount.Text='0 filer valda'; $slCount.Spring=$false
$slSpacer= New-Object System.Windows.Forms.ToolStripStatusLabel; $slSpacer.Spring=$true

# --- Klickbar SharePoint-länk ---
$slBatchLink = New-Object System.Windows.Forms.ToolStripStatusLabel
$slBatchLink.IsLink   = $true
$slBatchLink.Text     = 'SharePoint: —'
$slBatchLink.Enabled  = $false
$slBatchLink.Tag      = $null
$slBatchLink.ToolTipText = 'Direktlänk aktiveras när Batch# hittas i filer.'
$slBatchLink.add_Click({
    if ($this.Enabled -and $this.Tag) {
        try { Start-Process $this.Tag } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna:`n$($this.Tag)`n$($_.Exception.Message)","Länk") | Out-Null
        }
    }
})

$status.Items.AddRange(@($slCount,$slSpacer,$slBatchLink))
$tsc = New-Object System.Windows.Forms.ToolStripContainer
$tsc.Dock = 'Fill'
$tsc.LeftToolStripPanelVisible  = $false
$tsc.RightToolStripPanelVisible = $false

$form.SuspendLayout()
$form.Controls.Clear()
$form.Controls.Add($tsc)

# Meny högst upp
$tsc.TopToolStripPanel.Controls.Add($menuStrip)
$form.MainMenuStrip = $menuStrip

# Status längst ner
$tsc.BottomToolStripPanel.Controls.Add($status)

# Content i mitten
$content = New-Object System.Windows.Forms.Panel
$content.Dock='Fill'
$content.BackColor = $form.BackColor
$tsc.ContentPanel.Controls.Add($content)

# Dock=Top: nedersta först
$content.SuspendLayout()
$content.Controls.Add($btnBuild)
$content.Controls.Add($grpSign)
$content.Controls.Add($grpPick)
$content.Controls.Add($pLog)
$content.Controls.Add($tlSearch)
$content.Controls.Add($panelHeader)
$content.ResumeLayout()
$form.ResumeLayout()
$form.PerformLayout()
$form.AcceptButton = $btnScan

#endregion GUI Construction

function Add-CLBItems {
    param([System.Windows.Forms.CheckedListBox]$clb,[System.IO.FileInfo[]]$files,[switch]$AutoCheckFirst)
    $clb.BeginUpdate()
    $clb.Items.Clear()
    foreach($f in $files){
        if ($f -isnot [System.IO.FileInfo]) { try { $f = Get-Item -LiteralPath $f } catch { continue } }
        [void]$clb.Items.Add($f, $false)
    }
    $clb.EndUpdate()
    if ($AutoCheckFirst -and $clb.Items.Count -gt 0) { $clb.SetItemChecked(0,$true) }
    Update-StatusBar
}

function Get-CheckedFilePath { param([System.Windows.Forms.CheckedListBox]$clb)
    for($i=0;$i -lt $clb.Items.Count;$i++){
        if ($clb.GetItemChecked($i)) {
            $fi = [System.IO.FileInfo]$clb.Items[$i]
            return $fi.FullName
        }
    }
    return $null
}

function Clear-GUI {
    $txtLSP.Text = ''
    $txtSigner.Text = ''
    $chkWriteSign.Checked = $false
    $chkOverwriteSign.Checked = $false
    Add-CLBItems -clb $clbCsv -files @()
    Add-CLBItems -clb $clbNeg -files @()
    Add-CLBItems -clb $clbPos -files @()
    Add-CLBItems -clb $clbLsp -files @() 
    $outputBox.Clear()
    Update-BuildEnabled
    Gui-Log "🧹 GUI rensat." 'Info'
    Update-BatchLink
}

$onExclusive = {
    $clb = $this
    if ($_.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        for ($i=0; $i -lt $clb.Items.Count; $i++) {
            if ($i -ne $_.Index -and $clb.GetItemChecked($i)) { $clb.SetItemChecked($i, $false) }
        }
    }
    $clb.BeginInvoke([Action]{ Update-BuildEnabled }) | Out-Null
}
$clbCsv.add_ItemCheck($onExclusive)
$clbNeg.add_ItemCheck($onExclusive)
$clbPos.add_ItemCheck($onExclusive)

function Get-SelectedFileCount {
    $n=0
    if (Get-CheckedFilePath $clbCsv) { $n++ }
    if (Get-CheckedFilePath $clbNeg) { $n++ }
    if (Get-CheckedFilePath $clbPos) { $n++ }
    if (Get-CheckedFilePath $clbLsp) { $n++ }
    return $n
}

function Update-StatusBar { $slCount.Text = "$(Get-SelectedFileCount) filer valda" }

function Update-BuildEnabled {
    $btnBuild.Enabled = ((Get-CheckedFilePath $clbNeg) -and (Get-CheckedFilePath $clbPos))
    Update-StatusBar
}

$script:LastScanResult = $null

function Update-BatchLink {
    try {
        $info = Get-BatchLinkInfo -SealPosPath (Get-CheckedFilePath $clbPos) -SealNegPath (Get-CheckedFilePath $clbNeg) -Lsp ($txtLSP.Text.Trim())
        if (-not $info -or -not $info.Batch) {
            $slBatchLink.Text        = 'SharePoint: —'
            $slBatchLink.Enabled     = $false
            $slBatchLink.Tag         = $null
            $slBatchLink.ToolTipText = 'Direktlänk aktiveras när Batch# hittas i sökt LSP.'
            return
        }
        $slBatchLink.Text        = "SharePoint: $($info.Batch)"
        $slBatchLink.Enabled     = $true
        $slBatchLink.Tag         = $info.Url
        $slBatchLink.ToolTipText = $info.Url
    } catch {
        Gui-Log "⚠️ Update-BatchLink: $($_.Exception.Message)" 'Warn'
    }
}

function Assert-StartupReady {
    if ($global:StartupReady) { return $true }
    Gui-Log "❌ Startkontroll misslyckades. Åtgärda konfigurationsfel innan du fortsätter." 'Error'
    return $false
}

#region Event Handlers
$miScan.add_Click({ $btnScan.PerformClick() })
$miBuild.add_Click({ if ($btnBuild.Enabled) { $btnBuild.PerformClick() } })
$miExit.add_Click({ $form.Close() })
$miClear.add_Click({ Clear-GUI })

function Set-Theme {
    $global:CurrentTheme = 'light'
    $form.BackColor        = [System.Drawing.Color]::WhiteSmoke
    $content.BackColor     = $form.BackColor
    $panelHeader.BackColor = [System.Drawing.Color]::SteelBlue
    $pLog.BackColor        = [System.Drawing.Color]::White
    $grpPick.BackColor     = $form.BackColor
    $grpSign.BackColor     = $form.BackColor
    $tlSearch.BackColor    = $form.BackColor
    $outputBox.BackColor   = [System.Drawing.Color]::White
    $outputBox.ForeColor   = [System.Drawing.Color]::Black
    $lblLSP.ForeColor      = [System.Drawing.Color]::Black
    $lblCsv.ForeColor      = [System.Drawing.Color]::Black
    $lblNeg.ForeColor      = [System.Drawing.Color]::Black
    $lblPos.ForeColor      = [System.Drawing.Color]::Black
    $grpPick.ForeColor     = [System.Drawing.Color]::Black
    $grpSign.ForeColor     = [System.Drawing.Color]::Black
    $pLog.ForeColor        = [System.Drawing.Color]::Black
    $tlSearch.ForeColor    = [System.Drawing.Color]::Black
}

$btnScan.Add_Click({
    if (-not (Assert-StartupReady)) { return }
    Gui-Log '🔎 Söker filer…' -Immediate
    try {
    $lsp = $txtLSP.Text.Trim()
    if (-not $lsp) { Gui-Log "⚠️ Ange ett LSP-nummer" 'Warn'; return }

    if ($script:LastScanResult -and $script:LastScanResult.Lsp -eq $lsp -and $script:LastScanResult.Folder -and (Test-Path -LiteralPath $script:LastScanResult.Folder)) {
        Gui-Log "♻️ Återanvänder senaste sökresultatet för $lsp." 'Info'
        Add-CLBItems -clb $clbCsv -files $script:LastScanResult.Csv -AutoCheckFirst
        Add-CLBItems -clb $clbNeg -files $script:LastScanResult.Neg -AutoCheckFirst
        Add-CLBItems -clb $clbPos -files $script:LastScanResult.Pos -AutoCheckFirst
        Add-CLBItems -clb $clbLsp -files $script:LastScanResult.LspFiles -AutoCheckFirst
        Update-BuildEnabled
        Update-BatchLink
        return
    }

    $folder = $null
    foreach ($path in $RootPaths) {
        $folder = Get-ChildItem $path -Directory -Recurse -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -match "#?$lsp" } |
                  Select-Object -First 1
        if ($folder) { break }
    }
    if (-not $folder) { Gui-Log "❌ Ingen LSP-mapp hittad för $lsp" 'Warn'; return }

    $files = Get-ChildItem $folder.FullName -File -ErrorAction SilentlyContinue
    $candCsv = $files | Where-Object { $_.Extension -ieq '.csv' -and ( $_.Name -match $lsp -or $_.Length -gt 100kb ) } | Sort-Object LastWriteTime -Descending
    $candNeg = $files | Where-Object { $_.Name -match '(?i)Neg.*\.xls[xm]$' -and $_.Name -match $lsp } | Sort-Object LastWriteTime -Descending
    $candPos = $files | Where-Object { $_.Name -match '(?i)Pos.*\.xls[xm]$' -and $_.Name -match $lsp } | Sort-Object LastWriteTime -Descending
    $candLsp = $files | Where-Object {
        ($_.Name -match '(?i)worksheet') -and ($_.Name -match [regex]::Escape($lsp)) -and ($_.Extension -match '^(\.xlsx|\.xlsm|\.xls)$')
    } | Sort-Object LastWriteTime -Descending
    Gui-Log "📂 Hittad mapp: $($folder.FullName)" 'Info'
    Add-CLBItems -clb $clbCsv -files $candCsv -AutoCheckFirst
    Add-CLBItems -clb $clbNeg -files $candNeg -AutoCheckFirst
    Add-CLBItems -clb $clbPos -files $candPos -AutoCheckFirst
    Add-CLBItems -clb $clbLsp -files $candLsp -AutoCheckFirst
    if ($candCsv.Count -eq 0) { Gui-Log "ℹ️ Ingen CSV hittad (endast .csv visas)." 'Info' }
    if ($candNeg.Count -eq 0) { Gui-Log "⚠️ Ingen Seal NEG hittad." 'Warn' }
    if ($candPos.Count -eq 0) { Gui-Log "⚠️ Ingen Seal POS hittad." 'Warn' }
    if ($candLsp.Count -eq 0) { Gui-Log "ℹ️ Ingen LSP Worksheet hittad." 'Info' }
    Update-BuildEnabled
    Update-BatchLink
    $script:LastScanResult = [pscustomobject]@{
        Lsp      = $lsp
        Folder   = $folder.FullName
        Csv      = $candCsv  | ForEach-Object { $_.FullName }
        Neg      = $candNeg  | ForEach-Object { $_.FullName }
        Pos      = $candPos  | ForEach-Object { $_.FullName }
        LspFiles = $candLsp  | ForEach-Object { $_.FullName }
    }
    } finally {
        Gui-Log '✅ Filer laddade'
    }
})

$btnCsvBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "CSV|*.csv|Alla filer|*.*"
    if ($dlg.ShowDialog() -eq 'OK') {
        $f = Get-Item -LiteralPath $dlg.FileName
        Add-CLBItems -clb $clbCsv -files @($f) -AutoCheckFirst
        Update-BuildEnabled
        Update-BatchLink
    }
})

$btnNegBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Excel|*.xlsx;*.xlsm|Alla filer|*.*"
    if ($dlg.ShowDialog() -eq 'OK') {
        $f = Get-Item -LiteralPath $dlg.FileName
        Add-CLBItems -clb $clbNeg -files @($f) -AutoCheckFirst
        Update-BuildEnabled
        Update-BatchLink
    }
})

$btnPosBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Excel|*.xlsx;*.xlsm|Alla filer|*.*"
    if ($dlg.ShowDialog() -eq 'OK') {
        $f = Get-Item -LiteralPath $dlg.FileName
        Add-CLBItems -clb $clbPos -files @($f) -AutoCheckFirst
        Update-BuildEnabled
        Update-BatchLink
    }
})

 

# ============================
# ===== RAPPORTLOGIK =========
# ============================

$btnBuild.Add_Click({
    if (-not (Assert-StartupReady)) { return }
    $request = @{
        CsvPath             = (Get-CheckedFilePath $clbCsv)
        SealNegPath         = Get-CheckedFilePath $clbNeg
        SealPosPath         = Get-CheckedFilePath $clbPos
        WorksheetPath       = Get-CheckedFilePath $clbLsp
        Lsp                 = ($txtLSP.Text + '').Trim()
        SignerText          = $txtSigner.Text
        SignSealTest        = $chkWriteSign.Checked
        OverwriteSignature  = $chkOverwriteSign.Checked
        IncludeSharePoint   = $true
        SaveInLsp           = $false
        SharePointLinkLabel = $slBatchLink
    }
    $resultPath = Invoke-AssayReportBuild @request
    if ($resultPath) { $global:LastReportPath = $resultPath }
})

#endregion Event Handlers

# === Tooltip-inställningar ===
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 8000
$toolTip.InitialDelay = 500
$toolTip.ReshowDelay  = 500
$toolTip.ShowAlways   = $true
$toolTip.SetToolTip($txtLSP, 'Ange LSP-numret utan ”#” och klicka på Sök filer.')
$toolTip.SetToolTip($btnScan, 'Sök efter LSP och lista tillgängliga filer.')
$toolTip.SetToolTip($clbCsv,  'Välj CSV-fil.')
$toolTip.SetToolTip($clbNeg,  'Välj Seal Test Neg-fil.')
$toolTip.SetToolTip($clbPos,  'Välj Seal Test Pos-fil.')
$toolTip.SetToolTip($btnCsvBrowse, 'Bläddra efter en CSV-fil manuellt.')
$toolTip.SetToolTip($btnNegBrowse, 'Bläddra efter Seal Test Neg-fil manuellt.')
$toolTip.SetToolTip($btnPosBrowse, 'Bläddra efter Seal Test Pos-fil manuellt.')
$toolTip.SetToolTip($txtSigner, 'Skriv fullständigt namn, signatur och datum (separerat med kommatecken).')
$toolTip.SetToolTip($chkWriteSign, 'Signatur appliceras på flikar.')
$toolTip.SetToolTip($chkOverwriteSign, 'Dubbelkontroll för att aktivera signering')
$toolTip.SetToolTip($btnBuild, 'Skapa och öppna rapporten baserat på de valda filerna.')
$txtLSP.add_TextChanged({ Update-BatchLink })

#region Main Run / Orchestration
# =============== SLUT ===============
function Enable-DoubleBuffer {
    $pi = [Windows.Forms.Control].GetProperty('DoubleBuffered',[Reflection.BindingFlags]'NonPublic,Instance')
    foreach($c in @($content,$pLog,$grpPick,$grpSign)) { if ($c) { $pi.SetValue($c,$true,$null) } }
}
try { Set-Theme 'light' } catch {}
Enable-DoubleBuffer
Update-BatchLink
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)

try{ Stop-Transcript | Out-Null }catch{}
#endregion Main Run / Orchestration
