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
. (Join-Path $modulesRoot 'Splash.ps1')
. (Join-Path $modulesRoot 'UiStyling.ps1')
. (Join-Path $modulesRoot 'Logging.ps1')
. (Join-Path $modulesRoot 'EpplusLoader.ps1')

# Ladda EPPlus tidigt så att typer + Excel-hjälpfunktioner alltid finns innan övriga moduler läses in.
try {
    Initialize-EPPlus -HintPath $Config.EpplusDllPath | Out-Null
} catch {
    Gui-Log ("⚠️ EPPlus kunde inte laddas vid start: {0}" -f $_.Exception.Message) 'Warn'
}

. (Join-Path $modulesRoot 'DataHelpers.ps1')
. (Join-Path $modulesRoot 'CsvBundle.ps1')
. (Join-Path $modulesRoot 'RuleBank.ps1')
. (Join-Path $modulesRoot 'RuleEngine.ps1')
. (Join-Path $modulesRoot 'SignatureHelpers.ps1')


# LÄGG TILL tidigt i Main.ps1
$perfModulePath = Join-Path $modulesRoot 'PerfHelper.psm1'
if (-not (Test-Path -LiteralPath $perfModulePath)) {
    throw "PerfHelper.psm1 hittades inte: $perfModulePath"
}

# Importera modulen (exporterade funktioner blir synliga i sessionen)
Import-Module -Name $perfModulePath -Force -ErrorAction Stop

# Valfritt: verifiera att kommandon finns
foreach ($fn in 'Start-PerfTimer','Stop-PerfTimer') {
    if (-not (Get-Command -Name $fn -ErrorAction SilentlyContinue)) {
        throw "Funktionen '$fn' saknas efter Import-Module. Kontrollera Export-ModuleMember i PerfHelper.psm1."
    }
}

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
try { Initialize-EPPlus -HintPath $Config.EpplusDllPath | Out-Null } catch { Gui-Log "⚠️ EPPlus-förkontroll misslyckades: $($_.Exception.Message)" 'Warn' }

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
$miArkiv   = New-Object System.Windows.Forms.ToolStripMenuItem('🗂️ Arkiv')
$miVerktyg = New-Object System.Windows.Forms.ToolStripMenuItem('🛠️ Verktyg')
$miSettings= New-Object System.Windows.Forms.ToolStripMenuItem('⚙️ Inställningar')
$miHelp    = New-Object System.Windows.Forms.ToolStripMenuItem('📖 Instruktioner')
$miAbout   = New-Object System.Windows.Forms.ToolStripMenuItem('ℹ️ Om')
$miScan  = New-Object System.Windows.Forms.ToolStripMenuItem('🔍 Sök filer')
$miBuild = New-Object System.Windows.Forms.ToolStripMenuItem('✅ Skapa rapport')
$miExit  = New-Object System.Windows.Forms.ToolStripMenuItem('❌ Avsluta')

# Rensa ev. gamla undermenyer
$miArkiv.DropDownItems.Clear()
$miVerktyg.DropDownItems.Clear()
$miSettings.DropDownItems.Clear()
$miHelp.DropDownItems.Clear()

# ----- Arkiv -----
$miNew         = New-Object System.Windows.Forms.ToolStripMenuItem('🆕 Nytt')
$miOpenRecent  = New-Object System.Windows.Forms.ToolStripMenuItem('📂 Öppna senaste rapport')
$miArkiv.DropDownItems.AddRange(@(
    $miNew,
    $miOpenRecent,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $miExit
))


# ----- Verktyg -----
$miScript1   = New-Object System.Windows.Forms.ToolStripMenuItem('📜 Kontrollprovsfilskript')
$miScript2   = New-Object System.Windows.Forms.ToolStripMenuItem('📅 Ändra datum-prefix för filnamn')
$miScript3   = New-Object System.Windows.Forms.ToolStripMenuItem('📅 TBD')
$miToggleSign = New-Object System.Windows.Forms.ToolStripMenuItem('✅ Aktivera Seal Test-signatur')
$miVerktyg.DropDownItems.AddRange(@(
    $miScript1,
    $miScript2,
    $miScript3,
    $miToggleSign
))

# ----- Inställningar -----
$miTheme = New-Object System.Windows.Forms.ToolStripMenuItem('🎨 Tema')
$miLightTheme = New-Object System.Windows.Forms.ToolStripMenuItem('☀️ Ljust (default)')
$miDarkTheme  = New-Object System.Windows.Forms.ToolStripMenuItem('🌙 Mörkt')
$miTheme.DropDownItems.AddRange(@($miLightTheme,$miDarkTheme))
$miSettings.DropDownItems.Add($miTheme)

# ----- Instruktioner -----
$miShowInstr   = New-Object System.Windows.Forms.ToolStripMenuItem('📖 Visa instruktioner')
$miFAQ         = New-Object System.Windows.Forms.ToolStripMenuItem('❓ Vanliga frågor (FAQ)')
$miHelpDlg     = New-Object System.Windows.Forms.ToolStripMenuItem('🆘 Hjälp')
$miHelp.DropDownItems.AddRange(@($miShowInstr,$miFAQ,$miHelpDlg))

$miGenvagar = New-Object System.Windows.Forms.ToolStripMenuItem('🔗 Genvägar')
$ShortcutGroups = @{
    '🗂️ IPT-mappar' = @(

        @{ Text='📂 IPT - PÅGÅENDE KÖRNINGAR';        Target='N:\QC\QC-1\IPT\2. IPT - PÅGÅENDE KÖRNINGAR' },
        @{ Text='📂 IPT - KLART FÖR SAMMANSTÄLLNING'; Target='N:\QC\QC-1\IPT\3. IPT - KLART FÖR SAMMANSTÄLLNING' },
        @{ Text='📂 IPT - KLART FÖR GRANSKNING';      Target='N:\QC\QC-1\IPT\4. IPT - KLART FÖR GRANSKNING' },
        @{ Text='📂 SPT Macro Assay';                 Target='N:\QC\QC-0\SPT\SPT macros\Assay' }

    )

    '📄 Dokument' = @(

        @{ Text='🧰 Utrustningslista';    Target=$UtrustningListPath },
        @{ Text='🧪 Kontrollprovsfil';    Target=$RawDataPath }

    )

    '🌐 Länkar' = @(

        @{ Text='⚡ IPT App';              Target='https://apps.powerapps.com/play/e/default-771c9c47-7f24-44dc-958e-34f8713a8394/a/fd340dbd-bbbf-470b-b043-d2af4cb62c83' },
        @{ Text='🌐 MES';                  Target='http://mes.cepheid.pri/camstarportal/?domain=CEPHEID.COM' },
        @{ Text='🌐 CSV Uploader';         Target='http://auw2wgxtpap01.cepaws.com/Welcome.aspx' },
        @{ Text='🌐 BMRAM';                Target='https://cepheid62468.coolbluecloud.com/' },
        @{ Text='🌐 Agile';                Target='https://agileprod.cepheid.com/Agile/default/login-cms.jsp' }

    )

}

foreach ($grp in $ShortcutGroups.GetEnumerator()) {

    $grpMenu = New-Object System.Windows.Forms.ToolStripMenuItem($grp.Key)
    foreach ($entry in $grp.Value) { Add-ShortcutItem -Parent $grpMenu -Text $entry.Text -Target $entry.Target }
    [void]$miGenvagar.DropDownItems.Add($grpMenu)

}

$miOm = New-Object System.Windows.Forms.ToolStripMenuItem('ℹ️ Om det här verktyget'); $miAbout.DropDownItems.Add($miOm)
$menuStrip.Items.AddRange(@($miArkiv,$miVerktyg,$miGenvagar,$miSettings,$miHelp,$miAbout))
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

$grpSign.Visible = $false
$baseHeight = $form.Height

# ---------- Utdatasparande ----------
$grpSave = New-Object System.Windows.Forms.GroupBox
$grpSave.Text = "Rapport-utdata"
$grpSave.Dock='Top'
$grpSave.Padding = New-Object System.Windows.Forms.Padding(10,8,10,10)
$grpSave.AutoSize = $false
$grpSave.Height = 62

$flSave = New-Object System.Windows.Forms.FlowLayoutPanel
$flSave.Dock='Fill'
$rbSaveInLsp = New-Object System.Windows.Forms.RadioButton
$rbSaveInLsp.Text = "Spara i LSP-mapp"
$rbSaveInLsp.Checked = $true
$rbSaveInLsp.AutoSize = $true
$rbTempOnly = New-Object System.Windows.Forms.RadioButton
$rbTempOnly.Text = "Öppna i temporärt läge"
$rbTempOnly.AutoSize = $true
$flSave.Controls.Add($rbSaveInLsp); $flSave.Controls.Add($rbTempOnly)
$chkSharePointInfo = New-Object System.Windows.Forms.CheckBox
$chkSharePointInfo.Text = "Inkludera SharePoint Info"
$chkSharePointInfo.AutoSize = $true
$chkSharePointInfo.Checked = $false
$flSave.Controls.Add($chkSharePointInfo)

$grpSave.Controls.Add($flSave)

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
$content.Controls.Add($grpSave)
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

function Get-BatchLinkInfo {
    param(
        [string]$SealPosPath,
        [string]$SealNegPath,
        [string]$Lsp
    )

    $batch = $null
    try { if ($SealPosPath) { $batch = Get-BatchNumberFromSealFile $SealPosPath } } catch {}
    if (-not $batch) {
        try { if ($SealNegPath) { $batch = Get-BatchNumberFromSealFile $SealNegPath } } catch {}
    }

    $batchEsc = if ($batch) { [uri]::EscapeDataString($batch) } else { '' }
    $lspEsc   = if ($Lsp)   { [uri]::EscapeDataString($Lsp) }   else { '' }

    $url = if ($SharePointBatchLinkTemplate) {
        ($SharePointBatchLinkTemplate -replace '\{BatchNumber\}', $batchEsc) -replace '\{LSP\}', $lspEsc
    } else {
        "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/AllItems.aspx?view=7&q=$batchEsc"
    }
    $linkText = if ($batch) { "Öppna $batch" } else { 'Ingen batch funnen' }

    return [pscustomobject]@{
        Batch    = $batch
        Url      = $url
        LinkText = $linkText
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
$miNew.add_Click({ Clear-GUI })

$miOpenRecent.add_Click({
    if ($global:LastReportPath -and (Test-Path -LiteralPath $global:LastReportPath)) {
        try { Start-Process -FilePath $global:LastReportPath } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna rapporten:\n$($_.Exception.Message)","Öppna senaste rapport") | Out-Null
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Ingen rapport har genererats i denna session.","Öppna senaste rapport") | Out-Null
    }
})

# Skript1..3
$miScript1.add_Click({
    $p = $Script1Path
    if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.Forms.MessageBox]::Show("Ange sökvägen till Skript1 i variabeln `$Script1Path.","Skript1") | Out-Null; return }
    if (-not (Test-Path -LiteralPath $p)) { [System.Windows.Forms.MessageBox]::Show("Filen hittades inte:\n$Script1Path","Skript1") | Out-Null; return }
    $ext=[System.IO.Path]::GetExtension($p).ToLowerInvariant()
    switch ($ext) {
        '.ps1' { Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$p`"" }
        '.bat' { Start-Process cmd.exe -ArgumentList "/c `"$p`"" }
        '.lnk' { Start-Process -FilePath $p }
        default { try { Start-Process -FilePath $p } catch { [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna filen:","Skript1") | Out-Null } }
    }
})

$miScript2.add_Click({
    $p = $Script2Path
    if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.Forms.MessageBox]::Show("Ange sökvägen till Skript2 i variabeln `$Script2Path.","Skript2") | Out-Null; return }
    if (-not (Test-Path -LiteralPath $p)) { [System.Windows.Forms.MessageBox]::Show("Filen hittades inte:\n$Script2Path","Skript2") | Out-Null; return }
    $ext=[System.IO.Path]::GetExtension($p).ToLowerInvariant()
    switch ($ext) {
        '.ps1' { Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$p`"" }
        '.bat' { Start-Process cmd.exe -ArgumentList "/c `"$p`"" }
        '.lnk' { Start-Process -FilePath $p }
        default { try { Start-Process -FilePath $p } catch { [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna filen:","Skript2") | Out-Null } }
    }
})

$miScript3.add_Click({
    $p = $Script3Path
    if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.Forms.MessageBox]::Show("...","Skript3") | Out-Null; return }
    if (-not (Test-Path -LiteralPath $p)) { [System.Windows.Forms.MessageBox]::Show("...","Skript3") | Out-Null; return }
    $ext=[System.IO.Path]::GetExtension($p).ToLowerInvariant()
    switch ($ext) {
        '.ps1' { Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$p`"" }
        '.bat' { Start-Process cmd.exe -ArgumentList "/c `"$p`"" }
        '.lnk' { Start-Process -FilePath $p }
        default { try { Start-Process -FilePath $p } catch { [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna filen:","Skript3") | Out-Null } }
    }
})

$miToggleSign.add_Click({
    $lsp = $txtLSP.Text.Trim()
    if (-not $lsp) {
        Gui-Log "⚠️ Ange och sök ett LSP först innan du aktiverar Seal Test-signatur." 'Warn'
        return
    }
    $selNeg = Get-CheckedFilePath $clbNeg
    $selPos = Get-CheckedFilePath $clbPos
    if (-not $selNeg -or -not $selPos) {
        Gui-Log "⚠️ Du måste först välja både Seal Test NEG och POS innan Seal Test-signatur kan aktiveras." 'Warn'
        return
    }
    $grpSign.Visible = -not $grpSign.Visible
    if ($grpSign.Visible) {
        $form.Height = $baseHeight + $grpSign.Height + 40
        $miToggleSign.Text  = '❌ Dölj Seal Test-signatur'
    }
    else {
        $form.Height = $baseHeight
        $miToggleSign.Text  = '✅ Aktivera Seal Test-signatur'
    }
})

function Set-Theme {
    param([string]$Theme)
    if ($Theme -eq 'dark') {
        $global:CurrentTheme = 'dark'
        $form.BackColor        = [System.Drawing.Color]::FromArgb(35,35,35)
        $content.BackColor     = $form.BackColor
        $panelHeader.BackColor = [System.Drawing.Color]::DarkSlateBlue
        $pLog.BackColor        = [System.Drawing.Color]::FromArgb(45,45,45)
        $grpPick.BackColor     = $form.BackColor   
        $grpSign.BackColor     = $form.BackColor
        $grpSave.BackColor     = $form.BackColor
        $tlSearch.BackColor    = $form.BackColor
        $outputBox.BackColor   = [System.Drawing.Color]::FromArgb(55,55,55)
        $outputBox.ForeColor   = [System.Drawing.Color]::White
        $lblLSP.ForeColor      = [System.Drawing.Color]::White
        $lblCsv.ForeColor      = [System.Drawing.Color]::White
        $lblNeg.ForeColor      = [System.Drawing.Color]::White
        $lblPos.ForeColor      = [System.Drawing.Color]::White
        $grpPick.ForeColor     = [System.Drawing.Color]::White
        $grpSign.ForeColor     = [System.Drawing.Color]::White
        $grpSave.ForeColor     = [System.Drawing.Color]::White
        $pLog.ForeColor        = [System.Drawing.Color]::White
        $tlSearch.ForeColor    = [System.Drawing.Color]::White
    } else {
        $global:CurrentTheme = 'light'
        $form.BackColor        = [System.Drawing.Color]::WhiteSmoke
        $content.BackColor     = $form.BackColor
        $panelHeader.BackColor = [System.Drawing.Color]::SteelBlue
        $pLog.BackColor        = [System.Drawing.Color]::White
        $grpPick.BackColor     = $form.BackColor
        $grpSign.BackColor     = $form.BackColor
        $grpSave.BackColor     = $form.BackColor
        $tlSearch.BackColor    = $form.BackColor
        $outputBox.BackColor   = [System.Drawing.Color]::White
        $outputBox.ForeColor   = [System.Drawing.Color]::Black
        $lblLSP.ForeColor      = [System.Drawing.Color]::Black
        $lblCsv.ForeColor      = [System.Drawing.Color]::Black
        $lblNeg.ForeColor      = [System.Drawing.Color]::Black
        $lblPos.ForeColor      = [System.Drawing.Color]::Black
        $grpPick.ForeColor     = [System.Drawing.Color]::Black
        $grpSign.ForeColor     = [System.Drawing.Color]::Black
        $grpSave.ForeColor     = [System.Drawing.Color]::Black
        $pLog.ForeColor        = [System.Drawing.Color]::Black
        $tlSearch.ForeColor    = [System.Drawing.Color]::Black
    }
}

$miLightTheme.add_Click({ Set-Theme 'light' })
$miDarkTheme.add_Click({ Set-Theme 'dark' })

# Instruktioner

$miShowInstr.add_Click({
    $msg = @"
Snabbguide

1. Skriv in ditt LSP och klicka "Sök Filer eller använd Bläddra..."

2. Välj fil:
   • 1x CSV
   • 1x Seal Test NEG
   • 1x Seal Test POS
   • 1x Worksheet

3. Välj Rapport-utdata:
   • Spara i LSP-mapp (default)
   • Öppna endast i temporärt läge
   • Inkludera flik "SharePoint Info"

4. Klicka på "Skapa rapport"

Excelrapport öppnas med följande flikar beroende på valda filer:
   • Information (generell)
   • Seal Test Info
   • STF Sum (och minusvärden)
   • Infinity/GX 
   • Kontrollmaterial
   • SharePoint Info

Felsökning:
   • Filen låst → Stäng Excelfiler.
"@
    [System.Windows.Forms.MessageBox]::Show($msg,"Instruktioner") | Out-Null
})

$miFAQ.add_Click({
    $faq = @"
Vad gör skriptet?

Det skapar en excel-rapport som jämför sökt LSP för Seal Test-Filer,
hämtar utrustningslista, rätt kontrollmaterial och SharePoint Info för sökt LSP.

1) Varför ser jag inte fliken “SharePoint Info”?
   • Kryssrutan “SharePoint Info” måste vara ibockad.
   • Inloggning kan saknas eller SharePoint-listan innehåller inte batchnumret.

2) UI fryser ibland – är det normalt?
   • Nej. PnP-kopplingen och läsningen görs i bakgrunden. Om det ändå känns segt:
     - Testa utan SharePoint först (avbocka) för att isolera.
     - Stäng tunga Excel-instans(er) i bakgrunden.

3) “Filen är låst/kan inte spara”
   • Stäng källfilen i Excel.
   • Kontrollera att OneDrive/SharePoint Sync inte håller filen exklusivt låst.
   • Spara till TEMP för att testa att genereringen fungerar.

4) Var sparas rapporten?
   • Välj “LSP-mapp” (samma mapp som ditt LSP) eller “TEMP” = sparas inte.
"@
    [System.Windows.Forms.MessageBox]::Show($faq,"Vanliga frågor") | Out-Null
})

$miHelpDlg.add_Click({
    $helpForm = New-Object System.Windows.Forms.Form
    $helpForm.Text = 'Skicka meddelande'
    $helpForm.Size = New-Object System.Drawing.Size(400,300)
    $helpForm.StartPosition = 'CenterParent'
    $helpForm.Font = $form.Font
    $helpBox = New-Object System.Windows.Forms.TextBox
    $helpBox.Multiline = $true
    $helpBox.ScrollBars = 'Vertical'
    $helpBox.Dock = 'Fill'
    $helpBox.Font = New-Object System.Drawing.Font('Segoe UI',9)
    $helpBox.Margin = New-Object System.Windows.Forms.Padding(10)
    $panelButtons = New-Object System.Windows.Forms.FlowLayoutPanel
    $panelButtons.Dock = 'Bottom'
    $panelButtons.FlowDirection = 'RightToLeft'
    $panelButtons.Padding = New-Object System.Windows.Forms.Padding(10)
    $btnSend = New-Object System.Windows.Forms.Button
    $btnSend.Text = 'Skicka'
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Avbryt'
    $panelButtons.Controls.Add($btnSend)
    $panelButtons.Controls.Add($btnCancel)
    $helpForm.Controls.Add($helpBox)
    $helpForm.Controls.Add($panelButtons)
    $btnSend.Add_Click({
        $msg = $helpBox.Text.Trim()
        if (-not $msg) { [System.Windows.Forms.MessageBox]::Show('Ange ett meddelande innan du skickar.','Hjälp') | Out-Null; return }
        try {
            $helpDir = Join-Path $PSScriptRoot 'help'
            if (-not (Test-Path $helpDir)) { New-Item -ItemType Directory -Path $helpDir -Force | Out-Null }
            $ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
            $file = Join-Path $helpDir "help_${ts}.txt"
            Set-Content -Path $file -Value $msg -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show('Meddelandet sparades. Tack!','Hjälp') | Out-Null
            $helpForm.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte spara meddelandet:\n$($_.Exception.Message)",'Hjälp') | Out-Null
        }
    })
    $btnCancel.Add_Click({ $helpForm.Close() })
    $helpForm.ShowDialog() | Out-Null
})

# Om
$miOm.add_Click({ [System.Windows.Forms.MessageBox]::Show("OBS! Detta verktyg är endast ett hjälpmedel och ersätter inte någon process hos PQC.`n $ScriptVersion`nav Jesper","Om") | Out-Null })

$btnScan.Add_Click({
    if (-not (Assert-StartupReady)) { return }
    Gui-Log '🔎 Söker filer…' -Immediate
    # Start performance timer for scan/search
    
    Import-Module 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less v06.9\Modules\PerfHelper.psm1' -Force
    $perfScan = Start-PerfTimer -Label 'Scan/Search'

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
        # Stop performance timer and log elapsed time
        try {
            if ($perfScan) { Stop-PerfTimer -Perf $perfScan }
        } catch {}
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

 

if (-not (Get-Command Write-SPSheet-Safe -ErrorAction SilentlyContinue)) {
    function Write-SPSheet-Safe {
        param(
            [OfficeOpenXml.ExcelPackage]$Pkg,
            [object]$Rows,                    
            [string[]]$DesiredOrder,            
            [string]$Batch
        )
        if (-not $Pkg) { return $false }
        $Rows = @($Rows)
        $name = "SharePoint Info"
        $wsOld = $Pkg.Workbook.Worksheets[$name]
        if ($wsOld) { $Pkg.Workbook.Worksheets.Delete($wsOld) }
        $ws = $Pkg.Workbook.Worksheets.Add($name)
        if ($Rows.Count -eq 0 -or $Rows[0] -eq $null) {
            $ws.Cells[1,1].Value = "No rows found (Batch=$Batch)"
            return $true
        }
        $isKV = ($Rows[0].psobject.Properties.Name -contains 'Rubrik') -and `
                ($Rows[0].psobject.Properties.Name -contains 'Värde')
        if ($isKV) {
            $ws.Cells[1,1].Value = "SharePoint Information"
            $ws.Cells[1,2].Value = ""
            $ws.Cells["A1:B1"].Merge = $true
            $ws.Cells["A1"].Style.Font.Bold = $true
            $ws.Cells["A1"].Style.Font.Size = 12
            $ws.Cells["A1"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
            $ws.Cells["A1"].Style.Fill.PatternType = "Solid"
            $ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkBlue)
            $ws.Cells["A1"].Style.HorizontalAlignment = "Center"
            $ws.Cells["A1"].Style.VerticalAlignment   = "Center"

            $r = 2
            foreach ($row in $Rows) {
                $ws.Cells[$r,1].Value = $row.Rubrik
                $ws.Cells[$r,2].Value = $row.'Värde'
                $r++
            }
            $lastRow = $r-1
            $ws.Cells["A2:A$lastRow"].Style.Font.Bold = $true
            $ws.Cells["A2:A$lastRow"].Style.Fill.PatternType = "Solid"
            $ws.Cells["A2:A$lastRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Gainsboro)
            $ws.Cells["B2:B$lastRow"].Style.Fill.PatternType = "Solid"
            $ws.Cells["B2:B$lastRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::WhiteSmoke)

            $rng = $ws.Cells["A1:B$lastRow"]
            $rng.Style.Font.Name = "Arial"
            $rng.Style.Font.Size = 10
            $rng.Style.HorizontalAlignment = "Left"
            $rng.Style.VerticalAlignment   = "Center"
            $rng.Style.Border.Top.Style    = "Thin"
            $rng.Style.Border.Bottom.Style = "Thin"
            $rng.Style.Border.Left.Style   = "Thin"
            $rng.Style.Border.Right.Style  = "Thin"
            $rng.Style.Border.BorderAround("Medium")
            try { $rng.AutoFitColumns() | Out-Null } catch {}
        }
        else {
            $cols = @()
            if ($DesiredOrder) { $cols += $DesiredOrder }
            foreach ($k in $Rows[0].psobject.Properties.Name) {
                if ($cols -notcontains $k) { $cols += $k }
            }

            for ($c=0; $c -lt $cols.Count; $c++) {
                $ws.Cells[1,$c+1].Value = $cols[$c]
                $ws.Cells[1,$c+1].Style.Font.Bold = $true
            }
            $r = 2
            foreach ($row in $Rows) {
                for ($c=0; $c -lt $cols.Count; $c++) {
                    $ws.Cells[$r,$c+1].Value = $row.$($cols[$c])
                }
                $r++
            }
            try {
                if ($ws.Dimension) {
                    $maxR = [Math]::Min($ws.Dimension.End.Row, 2000)
                    $ws.Cells[$ws.Dimension.Start.Row,$ws.Dimension.Start.Column,$maxR,$ws.Dimension.End.Column].AutoFitColumns() | Out-Null
                }
            } catch {}
        }
        return $true
    }
}

# ============================
# ===== RAPPORTLOGIK =========
# ============================

$btnBuild.Add_Click({
    if (-not (Assert-StartupReady)) { return }
    Gui-Log 'Skapar rapport…' -Immediate
    # Start performance timer for the build phase
    $perfBuild = Start-PerfTimer -Label 'Build'
    try {
    if (-not (Load-EPPlus)) { Gui-Log "❌ EPPlus kunde inte laddas – avbryter." 'Error'; return }
 
    $selCsv = Get-CheckedFilePath $clbCsv
    $selNeg = Get-CheckedFilePath $clbNeg
    $selPos = Get-CheckedFilePath $clbPos

    if (-not $selNeg -or -not $selPos) { Gui-Log "❌ Du måste välja en Seal NEG och en Seal POS." 'Error'; return }
    $lsp = ($txtLSP.Text + '').Trim()
    if (-not $lsp) { Gui-Log "⚠️ Ange ett LSP-nummer." 'Warn'; return }

    Gui-Log "📄 Neg-fil: $(Split-Path $selNeg -Leaf)" 'Info'
    Gui-Log "📄 Pos-fil: $(Split-Path $selPos -Leaf)" 'Info'
    if ($selCsv) { Gui-Log "📄 CSV: $(Split-Path $selCsv -Leaf)" 'Info' } else { Gui-Log "ℹ️ Ingen CSV vald." 'Info' }

#region REGION: CSV
    $csvBundle = $null
    $csvLines  = $null
    if ($selCsv -and (Test-Path -LiteralPath $selCsv)) {
        try {
            $csvBundle = Get-TestsSummaryBundle -Path $selCsv
            if ($csvBundle -and $csvBundle.Lines) { $csvLines = $csvBundle.Lines }
        } catch {
            Gui-Log ("⚠️ Kunde inte läsa CSV: " + $_.Exception.Message) 'Warn'
        }
        if (-not $csvLines) {
            try { $csvLines = Get-Content -LiteralPath $selCsv } catch {}
        }
    }

#endregion REGION: CSV
    $negWritable = $true; $posWritable = $true
    if ($chkWriteSign.Checked) {
        $negWritable = -not (Test-FileLocked $selNeg); if (-not $negWritable) { Gui-Log "🔒 NEG är låst (öppen i Excel?)." 'Warn' }
        $posWritable = -not (Test-FileLocked $selPos); if (-not $posWritable) { Gui-Log "🔒 POS är låst (öppen i Excel?)." 'Warn' }
    }
    $pkgNeg = $null; $pkgPos = $null; $pkgOut = $null
    try {
        try {
            $pkgNeg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selNeg))
            $pkgPos = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selPos))
        } catch {
            Gui-Log "❌ Kunde inte öppna NEG/POS: $($_.Exception.Message)" 'Error'
            return
        }
 
        $templatePathV5 = Join-Path $PSScriptRoot "output_template-v5.xlsx"
        $templatePathV4 = Join-Path $PSScriptRoot "output_template-v4.xlsx"
        $templatePath = if (Test-Path -LiteralPath $templatePathV5) { $templatePathV5 } elseif (Test-Path -LiteralPath $templatePathV4) { $templatePathV4 } else { $null }
if (-not $templatePath) { Gui-Log "❌ Mallfilen 'output_template-v5.xlsx' (eller v4) saknas!" 'Error'; return }
        try {
            $pkgOut = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($templatePath))
        } catch {
            Gui-Log "❌ Kunde inte läsa mall: $($_.Exception.Message)" 'Error'
            return
        }

        # ============================
        # === SIGNATUR I NEG/POS  ====
        # ============================

        $signToWrite = ($txtSigner.Text + '').Trim()
        if ($chkWriteSign.Checked) {
            if (-not $signToWrite) { Gui-Log "❌ Ingen signatur angiven (B47). Avbryter."; return }
            if (-not (Confirm-SignatureInput -Text $signToWrite)) { Gui-Log "🛑 Signatur ej bekräftad. Avbryter."; return }

            $negWritten = 0; $posWritten = 0; $negSkipped = 0; $posSkipped = 0
            foreach ($ws in $pkgNeg.Workbook.Worksheets) {
                if ($ws.Name -eq 'Worksheet Instructions') { continue }
                $h3 = ($ws.Cells['H3'].Text + '').Trim()
                if ($h3 -match '^[0-9]') {
                    $existing = ($ws.Cells['B47'].Text + '').Trim()
                    if ($existing -and -not $chkOverwriteSign.Checked) { $negSkipped++; continue }
                    $ws.Cells['B47'].Style.Numberformat.Format = '@'
                    $ws.Cells['B47'].Value = $signToWrite
                    $negWritten++

                } elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/\?A|NA|Tomt( innehåll)?)$') {
                    break
                }
            }
            foreach ($ws in $pkgPos.Workbook.Worksheets) {
                if ($ws.Name -eq 'Worksheet Instructions') { continue }
                $h3 = ($ws.Cells['H3'].Text + '').Trim()
                if ($h3 -match '^[0-9]') {
                    $existing = ($ws.Cells['B47'].Text + '').Trim()
                    if ($existing -and -not $chkOverwriteSign.Checked) { $posSkipped++; continue }
                    $ws.Cells['B47'].Style.Numberformat.Format = '@'
                    $ws.Cells['B47'].Value = $signToWrite
                    $posWritten++
                } elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/\?A|NA|Tomt( innehåll)?)$') {
                    break
                }
            }
            try {
                if ($negWritten -eq 0 -and $negSkipped -eq 0 -and $posWritten -eq 0 -and $posSkipped -eq 0) {
                    Gui-Log "ℹ️ Inga databladsflikar efter flik 1 att sätta signatur i (ingen åtgärd)."
                } else {
                    if ($negWritten -gt 0 -and $negWritable) { $pkgNeg.Save() } elseif ($negWritten -gt 0) { Gui-Log "🔒 Kunde inte spara NEG (låst)." 'Warn' }
                    if ($posWritten -gt 0 -and $posWritable) { $pkgPos.Save() } elseif ($posWritten -gt 0) { Gui-Log "🔒 Kunde inte spara POS (låst)." 'Warn' }
                    Gui-Log "🖊️ Signatur satt: NEG $negWritten blad (överhoppade $negSkipped), POS $posWritten blad (överhoppade $posSkipped)."
                }
            } catch {
                Gui-Log "⚠️ Kunde inte spara signatur i NEG/POS: $($_.Exception.Message)" 'Warn'
            }
        }

        # ============================
        # === CSV (Info/Control)  ====
        # ============================

        $runAssay = $null
        if ($csvBundle) {
            try { $runAssay = Get-AssayFromTestsSummaryBundle -Bundle $csvBundle } catch {}
        }
        if (-not $runAssay -and $selCsv) {
            try { $runAssay = Get-AssayFromCsv -Path $selCsv -StartRow 10 } catch {}
        }
        if ($runAssay) { Gui-Log "🔎 Assay från CSV: $runAssay" }
        
        $controlTab = $null
        if ($runAssay) { $controlTab = Get-ControlTabName -AssayName $runAssay }
        if ($controlTab) { Gui-Log "🧪 Control Material-flik: $controlTab" } else { Gui-Log "ℹ️ Ingen control-mappning (fortsätter utan)." }

        # ============================
        # === Läs avvikelser       ===
        # ============================

        $violationsNeg = @(); $violationsPos = @(); $failNegCount = 0; $failPosCount = 0
         foreach ($ws in $pkgNeg.Workbook.Worksheets) {
            if ($ws.Name -eq "Worksheet Instructions") { continue }
            if (-not $ws.Dimension) { continue }
            $obsC = Find-ObservationCol $ws
            for ($r = 3; $r -le 45; $r++) {
                $valK = $ws.Cells["K$r"].Value; $textL = $ws.Cells["L$r"].Text
                if ($valK -ne $null -and $valK -is [double]) {
                    if ($textL -eq "FAIL" -or $valK -le -3.0) {
                        $obsTxt = $ws.Cells[$r, $obsC].Text
                        $violationsNeg += [PSCustomObject]@{
                            Sheet      = $ws.Name
                            Cartridge  = $ws.Cells["H$r"].Text
                            InitialW   = $ws.Cells["I$r"].Value
                            FinalW     = $ws.Cells["J$r"].Value
                            WeightLoss = $valK
                            Status     = if ($textL -eq "FAIL") { "FAIL" } else { "Minusvärde" }
                            Obs        = $obsTxt
                        }
                        if ($textL -eq "FAIL") { $failNegCount++ }
                    }
                }
            }
        }
         foreach ($ws in $pkgPos.Workbook.Worksheets) {
            if ($ws.Name -eq "Worksheet Instructions") { continue }
            if (-not $ws.Dimension) { continue }
            $obsC = Find-ObservationCol $ws
            for ($r = 3; $r -le 45; $r++) {
                $valK = $ws.Cells["K$r"].Value; $textL = $ws.Cells["L$r"].Text
                if ($valK -ne $null -and $valK -is [double]) {
                    if ($textL -eq "FAIL" -or $valK -le -3.0) {
                        $obsTxt = $ws.Cells[$r, $obsC].Text
                        $violationsPos += [PSCustomObject]@{
                            Sheet      = $ws.Name
                            Cartridge  = $ws.Cells["H$r"].Text
                            InitialW   = $ws.Cells["I$r"].Value
                            FinalW     = $ws.Cells["J$r"].Value
                            WeightLoss = $valK
                            Status     = if ($textL -eq "FAIL") { "FAIL" } else { "Minusvärde" }
                            Obs        = $obsTxt
                        }
                        if ($textL -eq "FAIL") { $failPosCount++ }
                    }
                }
            }
        }
 
        # ============================
#region REGION: Formatting
        # === Seal Test Info (blad) ==
        # ============================

        $wsOut1 = $pkgOut.Workbook.Worksheets["Seal Test Info"]
        if (-not $wsOut1) { Gui-Log "❌ Fliken 'Seal Test Info' saknas i mallen"; return }

        for ($row = 3; $row -le 15; $row++) {
            $wsOut1.Cells["D$row"].Value = $null
            try { $wsOut1.Cells["D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::None } catch {}
        }

        $fields = @(
            @{ Label = "ROBAL";                         Cell = "F2"  }
            @{ Label = "Part Number";                   Cell = "B2"  }
            @{ Label = "Batch Number";                  Cell = "D2"  }
            @{ Label = "Cartridge Number (LSP)";        Cell = "B6"  }
            @{ Label = "PO Number";                     Cell = "B10" }
            @{ Label = "Assay Family";                  Cell = "D10" }
            @{ Label = "Weight Loss Spec";              Cell = "F10" }
            @{ Label = "Balance ID Number";             Cell = "B14" }
            @{ Label = "Balance Cal Due Date";          Cell = "D14" }
            @{ Label = "Vacuum Oven ID Number";         Cell = "B20" }
            @{ Label = "Vacuum Oven Cal Due Date";      Cell = "D20" }
            @{ Label = "Timer ID Number";               Cell = "B25" }
            @{ Label = "Timer Cal Due Date";            Cell = "D25" }
        )

        $forceText = @("ROBAL","Part Number","Batch Number","Cartridge Number (LSP)","PO Number","Assay Family","Balance ID Number","Vacuum Oven ID Number","Timer ID Number")
        $mismatchFields = $fields[0..6] | ForEach-Object { $_.Label }

        $row = 3
        foreach ($f in $fields) {
            $valNeg=''; $valPos=''
             foreach ($wsN in $pkgNeg.Workbook.Worksheets) {
                if ($wsN.Name -eq "Worksheet Instructions") { continue }
                $cell = $wsN.Cells[$f.Cell]
                if ($cell.Value -ne $null) { if ($cell.Value -is [datetime]) { $valNeg = $cell.Value.ToString('MMM-yy') } else { $valNeg = $cell.Text }; break }
            }
             foreach ($wsP in $pkgPos.Workbook.Worksheets) {
                if ($wsP.Name -eq "Worksheet Instructions") { continue }
                $cell = $wsP.Cells[$f.Cell]
                if ($cell.Value -ne $null) { if ($cell.Value -is [datetime]) { $valPos = $cell.Value.ToString('MMM-yy') } else { $valPos = $cell.Text }; break }
            }

            if ($forceText -contains $f.Label) {
                $wsOut1.Cells["B$row"].Style.Numberformat.Format = '@'
                $wsOut1.Cells["C$row"].Style.Numberformat.Format = '@'
            }

            $wsOut1.Cells["B$row"].Value = $valNeg
            $wsOut1.Cells["C$row"].Value = $valPos
            $wsOut1.Cells["B$row"].Style.Border.Right.Style = "Medium"
            $wsOut1.Cells["C$row"].Style.Border.Left.Style  = "Medium"

            if ($mismatchFields -contains $f.Label -and $valNeg -ne $valPos) {
                $wsOut1.Cells["D$row"].Value = "Mismatch"

                Style-Cell $wsOut1.Cells["D$row"] $true "FF0000" "Medium" "FFFFFF"
                Gui-Log "⚠️ Mismatch: $($f.Label) ($valNeg vs $valPos)"
            }
            $row++
        }

        # ============================
        # === Testare (B43)        ===
        # ============================

        $testersNeg = @(); $testersPos = @()
        foreach ($s in $pkgNeg.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) { $t=$s.Cells["B43"].Text; if ($t) { $testersNeg += ($t -split ",") } }
        foreach ($s in $pkgPos.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) { $t=$s.Cells["B43"].Text; if ($t) { $testersPos += ($t -split ",") } }
        $testersNeg = $testersNeg | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique
        $testersPos = $testersPos | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique

        $wsOut1.Cells["B16"].Value = "Name of Tester"
        $wsOut1.Cells["B16:C16"].Merge = $true
        $wsOut1.Cells["B16"].Style.HorizontalAlignment = "Center"

        $maxTesters = [Math]::Max($testersNeg.Count, $testersPos.Count)
        $initialRows = 11
        if ($maxTesters -lt $initialRows) { $wsOut1.DeleteRow(17 + $maxTesters, $initialRows - $maxTesters) }
        if ($maxTesters -gt $initialRows) {
            $rowsToAdd = $maxTesters - $initialRows
            $lastRow = 16 + $initialRows
            for ($i = 1; $i -le $rowsToAdd; $i++) { $wsOut1.InsertRow($lastRow + 1, 1, $lastRow) }
        }
        for ($i = 0; $i -lt $maxTesters; $i++) {
            $rowIndex = 17 + $i
            $wsOut1.Cells["A$rowIndex"].Value = $null
            $wsOut1.Cells["B$rowIndex"].Value = if ($i -lt $testersNeg.Count) { $testersNeg[$i] } else { "N/A" }
            $wsOut1.Cells["C$rowIndex"].Value = if ($i -lt $testersPos.Count) { $testersPos[$i] } else { "N/A" }

            $topStyle    = if ($i -eq 0) { "Medium" } else { "Thin" }
            $bottomStyle = if ($i -eq $maxTesters - 1) { "Medium" } else { "Thin" }
            foreach ($col in @("B","C")) {
                $cell = $wsOut1.Cells["$col$rowIndex"]
                $cell.Style.Border.Top.Style    = $topStyle
                $cell.Style.Border.Bottom.Style = $bottomStyle
                $cell.Style.Border.Left.Style   = "Medium"
                $cell.Style.Border.Right.Style  = "Medium"
                $cell.Style.Fill.PatternType = "Solid"
                $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
            }
        }

        # ============================
        # === Signatur-jämförelse  ===
        # ============================

        $negSigSet = Get-SignatureSetForDataSheets -Pkg $pkgNeg
        $posSigSet = Get-SignatureSetForDataSheets -Pkg $pkgPos
        $negSet = New-Object 'System.Collections.Generic.HashSet[string]'
        $posSet = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($n in $negSigSet.NormSet) { [void]$negSet.Add($n) }
        foreach ($p in $posSigSet.NormSet) { [void]$posSet.Add($p) }
        $hasNeg = ($negSet.Count -gt 0)
        $hasPos = ($posSet.Count -gt 0)
        $onlyNeg = @(); $onlyPos = @(); $sigMismatch = $false
        if ($hasNeg -and $hasPos) {
            foreach ($n in $negSet) { if (-not $posSet.Contains($n)) { $onlyNeg += $n } }
            foreach ($p in $posSet) { if (-not $negSet.Contains($p)) { $onlyPos += $p } }
            $sigMismatch = ($onlyNeg.Count -gt 0 -or $onlyPos.Count -gt 0)
        } else {
            $sigMismatch = $false
        }

        $mismatchSheets = @()
        if ($sigMismatch) {
            foreach ($k in $onlyNeg) {
                $raw = if ($negSigSet.RawByNorm.ContainsKey($k)) { $negSigSet.RawByNorm[$k] } else { $k }
                $where = if ($negSigSet.Occ.ContainsKey($k)) { ($negSigSet.Occ[$k] -join ', ') } else { '—' }
                $mismatchSheets += ("NEG: " + $raw + "  [Blad: " + $where + "]")
            }

            foreach ($k in $onlyPos) {
                $raw = if ($posSigSet.RawByNorm.ContainsKey($k)) { $posSigSet.RawByNorm[$k] } else { $k }
                $where = if ($posSigSet.Occ.ContainsKey($k)) { ($posSigSet.Occ[$k] -join ', ') } else { '—' }
                $mismatchSheets += ("POS: " + $raw + "  [Blad: " + $where + "]")
            }
            Gui-Log "⚠️ Mismatch: Print Full Name, Sign, and Date (NEG vs POS)"
        }

        function Set-MergedWrapAutoHeight {
            param([OfficeOpenXml.ExcelWorksheet]$Sheet,[int]$RowIndex,[int]$ColStart=2,[int]$ColEnd=3,[string]$Text)
            $rng = $Sheet.Cells[$RowIndex, $ColStart, $RowIndex, $ColEnd]
            $rng.Style.WrapText = $true
            $rng.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::None
            $Sheet.Row($RowIndex).CustomHeight = $false

            try {
                $wChars = [Math]::Floor(($Sheet.Column($ColStart).Width + $Sheet.Column($ColEnd).Width) - 2); if ($wChars -lt 1) { $wChars = 1 }
                $segments = $Text -split "(\r\n|\n|\r)"; $lineCount = 0
                foreach ($seg in $segments) { if (-not $seg) { $lineCount++ } else { $lineCount += [Math]::Ceiling($seg.Length / $wChars) } }
                if ($lineCount -lt 1) { $lineCount = 1 }
                $targetHeight = [Math]::Max(15, [Math]::Ceiling(15 * $lineCount * 2.15))
                if ($Sheet.Row($RowIndex).Height -lt $targetHeight) {
                    $Sheet.Row($RowIndex).Height = $targetHeight
                    $Sheet.Row($RowIndex).CustomHeight = $true
                }
            } catch { $Sheet.Row($RowIndex).CustomHeight = $false }
        }

        $signRow = 17 + $maxTesters + 3
        $displaySignNeg = $null; $displaySignPos = $null
        if ($signToWrite) { $displaySignNeg = $signToWrite; $displaySignPos = $signToWrite }
        else {
            $displaySignNeg = if ($negSigSet.RawFirst) { $negSigSet.RawFirst } else { '—' }
            $displaySignPos = if ($posSigSet.RawFirst) { $posSigSet.RawFirst } else { '—' }
        }

        $wsOut1.Cells["B$signRow"].Style.Numberformat.Format = '@'
        $wsOut1.Cells["C$signRow"].Style.Numberformat.Format = '@'
        $wsOut1.Cells["B$signRow"].Value = $displaySignNeg
        $wsOut1.Cells["C$signRow"].Value = $displaySignPos
        foreach ($col in @('B','C')) {

            $cell = $wsOut1.Cells["${col}$signRow"]
            Style-Cell $cell $false 'CCFFFF' 'Medium' $null
            $cell.Style.HorizontalAlignment = 'Center'
        }

        try { $wsOut1.Column(2).Width = 40; $wsOut1.Column(3).Width = 40 } catch {}
        try { $wsOut1.Column(4).Width = 10 } catch {}

        if ($sigMismatch) {
            $mismatchCell = $wsOut1.Cells["D$signRow"]
            $mismatchCell.Value = 'Mismatch'
            Style-Cell $mismatchCell $true 'FF0000' 'Medium' 'FFFFFF'

            if ($mismatchSheets.Count -gt 0) {
                for ($j = 0; $j -lt $mismatchSheets.Count; $j++) {
                    $rowIdx = $signRow + 1 + $j
                    try { $wsOut1.Cells["B$rowIdx:C$rowIdx"].Merge = $true } catch {}
                    $text = $mismatchSheets[$j]
                    $wsOut1.Cells["B$rowIdx"].Value = $text
                    foreach ($mc in $wsOut1.Cells["B$rowIdx:C$rowIdx"]) { Style-Cell $mc $false 'CCFFFF' 'Medium' $null }
                    $wsOut1.Cells["B$rowIdx:C$rowIdx"].Style.HorizontalAlignment = 'Center'

                    if ($text -like 'NEG:*' -or $text -like 'POS:*') {
                        Set-MergedWrapAutoHeight -Sheet $wsOut1 -RowIndex $rowIdx -ColStart 2 -ColEnd 3 -Text $text
                    }
                }
            }
        }
 
        # ============================
        # === STF Sum              ===
        # ============================

        $wsOut2 = $pkgOut.Workbook.Worksheets["STF Sum"]
        if (-not $wsOut2) { Gui-Log "❌ Fliken 'STF Sum' saknas i mallen!"; return }
        $totalRows = $violationsNeg.Count + $violationsPos.Count
        $currentRow = 2

        if ($totalRows -eq 0) {
            Gui-Log "✅ Seal Test hittades"
            $wsOut2.Cells["B1:H1"].Value = $null
            $wsOut2.Cells["A1"].Value = "Inga STF hittades!"
            Style-Cell $wsOut2.Cells["A1"] $true "D9EAD3" "Medium" "006100"
            $wsOut2.Cells["A1"].Style.HorizontalAlignment = "Left"
            if ($wsOut2.Dimension -and $wsOut2.Dimension.End.Row -gt 1) { $wsOut2.DeleteRow(2, $wsOut2.Dimension.End.Row - 1) }

        } else {
            Gui-Log "❗ $failNegCount avvikelser i NEG, $failPosCount i POS"
            $oldDataRows = 0
            if ($wsOut2.Dimension) { $oldDataRows = $wsOut2.Dimension.End.Row - 1; if ($oldDataRows -lt 0) { $oldDataRows = 0 } }
            if ($totalRows -lt $oldDataRows) { $wsOut2.DeleteRow(2 + $totalRows, $oldDataRows - $totalRows) }
            elseif ($totalRows -gt $oldDataRows) { $wsOut2.InsertRow(2 + $oldDataRows, $totalRows - $oldDataRows, 1 + $oldDataRows) }

            $currentRow = 2
            foreach ($v in $violationsNeg) {

                $wsOut2.Cells["A$currentRow"].Value = "NEG"
                $wsOut2.Cells["B$currentRow"].Value = $v.Sheet
                $wsOut2.Cells["C$currentRow"].Value = $v.Cartridge
                $wsOut2.Cells["D$currentRow"].Value = $v.InitialW
                $wsOut2.Cells["E$currentRow"].Value = $v.FinalW
                $wsOut2.Cells["F$currentRow"].Value = [Math]::Round($v.WeightLoss, 1)
                $wsOut2.Cells["G$currentRow"].Value = $v.Status
                $wsOut2.Cells["H$currentRow"].Value = if ([string]::IsNullOrWhiteSpace($v.Obs)) { 'NA' } else { $v.Obs }
                Style-Cell $wsOut2.Cells["A$currentRow"] $true "B5E6A2" "Medium" $null
                $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
                $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFF99"))
                $wsOut2.Cells["H$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["H$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#D9D9D9"))
 
                if ($v.Status -in @("FAIL","Minusvärde")) {
                    $wsOut2.Cells["F$currentRow"].Style.Font.Bold = $true
                    $wsOut2.Cells["F$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                    $wsOut2.Cells["G$currentRow"].Style.Font.Bold = $true
                    $wsOut2.Cells["G$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                }
                Set-RowBorder -ws $wsOut2 -row $currentRow -firstRow 2 -lastRow ($totalRows + 1)
                $currentRow++
            }
            foreach ($v in $violationsPos) {
                $wsOut2.Cells["A$currentRow"].Value = "POS"
                $wsOut2.Cells["B$currentRow"].Value = $v.Sheet
                $wsOut2.Cells["C$currentRow"].Value = $v.Cartridge
                $wsOut2.Cells["D$currentRow"].Value = $v.InitialW
                $wsOut2.Cells["E$currentRow"].Value = $v.FinalW
                $wsOut2.Cells["F$currentRow"].Value = [Math]::Round($v.WeightLoss, 1)
                $wsOut2.Cells["G$currentRow"].Value = $v.Status
                $wsOut2.Cells["H$currentRow"].Value = if ($v.Obs) { $v.Obs } else { 'NA' }
 
                Style-Cell $wsOut2.Cells["A$currentRow"] $true "FFB3B3" "Medium" $null
                $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
                $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFF99"))
                $wsOut2.Cells["H$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["H$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#D9D9D9"))

                if ($v.Status -in @("FAIL","Minusvärde")) {
                    $wsOut2.Cells["F$currentRow"].Style.Font.Bold = $true
                    $wsOut2.Cells["F$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                    $wsOut2.Cells["G$currentRow"].Style.Font.Bold = $true
                    $wsOut2.Cells["G$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                }
                Set-RowBorder -ws $wsOut2 -row $currentRow -firstRow 2 -lastRow ($totalRows + 1)
                $currentRow++
            }

            $wsOut2.Cells.Style.WrapText = $false
            $wsOut2.Cells["A1"].Style.HorizontalAlignment = "Left"
            try { $wsOut2.Cells[2,6,([Math]::Max($currentRow-1,2)),6].Style.Numberformat.Format = '0.0' } catch {}
            if ($wsOut2.Dimension) { $wsOut2.Cells[$wsOut2.Dimension.Address].AutoFitColumns() }
        }

# ============================
#endregion REGION: Formatting
# === Information-blad     ===
# ============================

try {
    if (-not (Get-Command Add-Hyperlink -ErrorAction SilentlyContinue)) {
        function Add-Hyperlink {
            param([OfficeOpenXml.ExcelRange]$Cell,[string]$Text,[string]$Url)
            try {
                $Cell.Value = $Text
                $Cell.Hyperlink = [Uri]$Url
                $Cell.Style.Font.UnderLine = $true
                $Cell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0,102,204))
            } catch {}
        }
    }
    if (-not (Get-Command Find-RegexCell -ErrorAction SilentlyContinue)) {
        function Find-RegexCell {
            param([OfficeOpenXml.ExcelWorksheet]$Ws,[regex]$Rx,[int]$MaxRows=200,[int]$MaxCols=40)
            if (-not $Ws -or -not $Ws.Dimension) { return $null }
            $rMax = [Math]::Min($Ws.Dimension.End.Row, $MaxRows)
            $cMax = [Math]::Min($Ws.Dimension.End.Column, $MaxCols)
            for ($r=1; $r -le $rMax; $r++) {
                for ($c=1; $c -le $cMax; $c++) {
            $t = Normalize-HeaderText ($Ws.Cells[$r,$c].Text + '')
            if ($t -and $Rx.IsMatch($t)) { return @{Row=$r;Col=$c;Text=$t} }
                }
            }
            return $null
        }
    }
 
    if (-not (Get-Command Get-SealHeaderDocInfo -ErrorAction SilentlyContinue)) {
        function Get-SealHeaderDocInfo {
            param([OfficeOpenXml.ExcelPackage]$Pkg)
            $result = [pscustomobject]@{ Raw=''; DocNo=''; Rev='' }
            if (-not $Pkg) { return $result }
            $ws = $Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
            if (-not $ws) { return $result }
            try {
                $lt = ($ws.HeaderFooter.OddHeader.LeftAlignedText + '').Trim()
                if (-not $lt) { $lt = ($ws.HeaderFooter.EvenHeader.LeftAlignedText + '').Trim() }
                $result.Raw = $lt
                $rx = [regex]'(?i)(?:document\s*(?:no|nr|#|number)\s*[:#]?\s*([A-Z0-9\-_\.\/]+))?.*?(?:rev(?:ision)?\.?\s*[:#]?\s*([A-Z0-9\-_\.]+))?'
                $m = $rx.Match($lt)
                if ($m.Success) {
                    if ($m.Groups[1].Value) { $result.DocNo = $m.Groups[1].Value.Trim() }
                    if ($m.Groups[2].Value) { $result.Rev   = $m.Groups[2].Value.Trim() }
                }
            } catch {}
            return $result
        }
    }

    $wsInfo = $pkgOut.Workbook.Worksheets['Information']
    if (-not $wsInfo) {
        $wsInfo = $pkgOut.Workbook.Worksheets.Add('Information')
    }
    try { $wsInfo.Cells.Style.Font.Name='Arial'; $wsInfo.Cells.Style.Font.Size=11 } catch {}
    try {
        $csvStats = $null
        if ($selCsv -and (Test-Path -LiteralPath $selCsv)) {
            try { $csvStats = Get-CsvStats -Path $selCsv -Lines $csvLines -Bundle $csvBundle } catch { Gui-Log ("⚠️ Get-CsvStats: " + $_.Exception.Message) 'Warn' }
        }
        if (-not $csvLines) { $csvLines = @() }
        if (-not $csvStats) {
            $csvStats = [pscustomobject]@{
                TestCount    = 0
                DupCount     = 0
                Duplicates   = @()
                LspValues    = @()
                LspOK        = $null
                InstrumentByType = [ordered]@{}
            }
        }

        $infSN = @()
        if ($script:GXINF_Map) {
            foreach ($k in $script:GXINF_Map.Keys) {
                if ($k -like 'Infinity-*') {
                    $infSN += ($script:GXINF_Map[$k].Split(',') | ForEach-Object { ($_ + '').Trim() } | Where-Object { $_ })
                }
            }
        }

        $infSN = $infSN | Select-Object -Unique
        $infSummary = '—'

        try {
            if ($selCsv -and (Test-Path -LiteralPath $selCsv) -and $infSN.Count -gt 0) {
                $infSummary = Get-InfinitySpFromCsvStrict -Path $selCsv -InfinitySerials $infSN -Lines $csvLines
            }
        } catch {
            Gui-Log ("Infinity SP fel: " + $_.Exception.Message) 'Warn'
        }

        $dupSampleCount = 0
        $dupSampleList  = @()
        if ($csvLines -and $csvLines.Count -gt 8) {
            try {
                $hdrIdx  = if ($csvBundle) { [int]$csvBundle.HeaderRowIndex } else { 7 }
                $dataIdx = if ($csvBundle) { [int]$csvBundle.DataStartRowIndex } else { 9 }
                $delim   = if ($csvBundle -and $csvBundle.Delimiter) { [string]$csvBundle.Delimiter } else { ',' }

                $headerFields = Parse-CsvLine -line $csvLines[$hdrIdx] -delim $delim

                $sampleIdx = -1
                if ($csvBundle -and $csvBundle.HeaderIndex -and $csvBundle.HeaderIndex.ContainsKey('sample id')) {
                    $sampleIdx = [int]$csvBundle.HeaderIndex['sample id']
                } else {
                    for ($i = 0; $i -lt $headerFields.Count; $i++) {
                        $hf = ($headerFields[$i] + '').Trim().ToLowerInvariant()
                        if ($hf -match 'sample') { $sampleIdx = $i; break }
                    }
                }

                if ($sampleIdx -ge 0) {
                    $counts = @{}
                    for ($r = $dataIdx; $r -lt $csvLines.Count; $r++) {
                        $line = $csvLines[$r]
                        if (-not $line -or -not $line.Trim()) { continue }
                        $fields = Parse-CsvLine -line $line -delim $delim
                        if ($fields.Count -gt $sampleIdx) {
                            $val = ($fields[$sampleIdx] + '').Trim().Trim('"')
                            if ($val) {
                                if (-not $counts.ContainsKey($val)) { $counts[$val] = 0 }
                                $counts[$val]++
                            }
                        }
                    }

                    $dupList = @()
                    foreach ($entry in $counts.GetEnumerator()) {
                        if ($entry.Value -gt 1) { $dupList += ("$($entry.Key) x$($entry.Value)") }
                    }
                    $dupSampleCount = $dupList.Count
                    $dupSampleList  = $dupList
                }
            } catch {
                Gui-Log ("⚠️ Fel vid analys av Sample ID: " + $_.Exception.Message) 'Warn'
            }
        }

        $dupSampleText = if ($dupSampleCount -gt 0) {
            $show = ($dupSampleList | Select-Object -First 8) -join ', '
            "$dupSampleCount ($show)"
        } else { 'N/A' }
        $dupCartText = if ($csvStats.DupCount -gt 0) {
            $show = ($csvStats.Duplicates | Select-Object -First 8) -join ', '
            "$($csvStats.DupCount) ($show)"
        } else { 'N/A' }
        $lspSummary = ''
        try {
            if ($csvLines -and $csvLines.Count -gt 8) {
                $dataIdx = if ($csvBundle) { [int]$csvBundle.DataStartRowIndex } else { 9 }
                $delim   = if ($csvBundle -and $csvBundle.Delimiter) { [string]$csvBundle.Delimiter } else { ',' }

                $idxLsp = 4
                if ($csvBundle -and $csvBundle.HeaderIndex -and $csvBundle.HeaderIndex.ContainsKey('reagent lot id')) {
                    $idxLsp = [int]$csvBundle.HeaderIndex['reagent lot id']
                }

                $counts = @{}
                for ($rr = $dataIdx; $rr -lt $csvLines.Count; $rr++) {
                    $ln = $csvLines[$rr]
                    if (-not $ln -or -not $ln.Trim()) { continue }
                    $fs = Parse-CsvLine -line $ln -delim $delim
                    if ($fs.Count -gt $idxLsp) {
                        $raw = ($fs[$idxLsp] + '').Trim().Trim('"')
                        if ($raw) {
                            $mLsp = [regex]::Match($raw,'(\d{5})')
                            $code = if ($mLsp.Success) { $mLsp.Groups[1].Value } else { $raw }
                            if (-not $counts.ContainsKey($code)) { $counts[$code] = 0 }
                            $counts[$code]++
                        }
                    }
                }

                if ($counts.Count -gt 0) {
                    $sorted = $counts.GetEnumerator() | Sort-Object Key
                    $lspSummaryParts = @()
                    foreach ($kvp in $sorted) {
                        $part = if ($kvp.Value -gt 1) { "$($kvp.Key) x$($kvp.Value)" } else { $kvp.Key }
                        $lspSummaryParts += $part
                    }
                    $total = $sorted.Count
                    if ($total -eq 1) {
                        $lspSummary = $sorted[0].Key
                    }
                    else {
                        $lspSummary = "$total (" + ($lspSummaryParts -join ', ') + ")"
                    }
                }
            }
        } catch {
            Gui-Log ("⚠️ Fel vid extraktion av LSP från CSV: " + $_.Exception.Message) 'Warn'
            $lspSummary = ''
        }

        $instText = if ($csvStats.InstrumentByType.Keys.Count -gt 0) {
            ($csvStats.InstrumentByType.GetEnumerator() | ForEach-Object { "$($_.Key)" } | Sort-Object) -join '; '
        } else { '' }

        function Find-InfoRow {
            param([OfficeOpenXml.ExcelWorksheet]$Ws, [string]$Label)
            if (-not $Ws -or -not $Ws.Dimension) { return $null }
            $maxRow = [Math]::Min($Ws.Dimension.End.Row, 300)
            for ($ri=1; $ri -le $maxRow; $ri++) {
                $txt = (($Ws.Cells[$ri,1].Text) + '').Trim()
                if (-not $txt) { continue }
                if ($txt.ToLowerInvariant() -eq $Label.ToLowerInvariant()) { return $ri }
            }
            return $null
        }

        $isNewLayout = $true

        try {
            $tmpRow = Find-InfoRow -Ws $wsInfo -Label 'CSV-Info'
            if ($tmpRow) { $isNewLayout = $true }

        } catch {}

        $rowCsvFile    = Find-InfoRow -Ws $wsInfo -Label 'CSV'
        $rowLsp        = Find-InfoRow -Ws $wsInfo -Label 'LSP'
        $rowAntal      = Find-InfoRow -Ws $wsInfo -Label 'Antal tester'
        $rowDupSample  = Find-InfoRow -Ws $wsInfo -Label 'Dubblett Sample ID'
        if (-not $rowDupSample) { $rowDupSample = Find-InfoRow -Ws $wsInfo -Label 'Dublett Sample ID' }
        $rowDupCart    = Find-InfoRow -Ws $wsInfo -Label 'Dubblett Cartridge S/N'
        if (-not $rowDupCart) { $rowDupCart = Find-InfoRow -Ws $wsInfo -Label 'Dublett Cartridge S/N' }
        $rowInst       = Find-InfoRow -Ws $wsInfo -Label 'Använda INF/GX'
        $rowBag = Find-InfoRow -Ws $wsInfo -Label 'Bag Numbers Tested Using Infinity'
if (-not $rowBag) { $rowBag = Find-InfoRow -Ws $wsInfo -Label 'Bag Numbers Tested Using Infinity:' }
if (-not $rowBag) { $rowBag = 14 } 

$wsInfo.Cells["B$rowBag"].Style.Numberformat.Format = '@'
$wsInfo.Cells["B$rowBag"].Value = $infSummary

        if ($isNewLayout) {
            $rowLsp = Find-InfoRow -Ws $wsInfo -Label 'LSP'

            if (-not $rowCsvFile)   { $rowCsvFile   = 8 }
            if (-not $rowLsp)       { $rowLsp       = 9 }
            if (-not $rowAntal)     { $rowAntal     = 10 }
            if (-not $rowDupSample) { $rowDupSample = 11 }
            if (-not $rowDupCart)   { $rowDupCart   = 12 }
            if (-not $rowInst)      { $rowInst      = 13 }

        }

        if ($selCsv) {
            $wsInfo.Cells["B$rowCsvFile"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowCsvFile"].Value = (Split-Path $selCsv -Leaf)
        } else {
            $wsInfo.Cells["B$rowCsvFile"].Value = ''
        }

        if ($lspSummary -and $lspSummary -ne '') {
            $wsInfo.Cells["B$rowLsp"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowLsp"].Value = $lspSummary
        } else {
            $wsInfo.Cells["B$rowLsp"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowLsp"].Value = $lsp
        }

        $wsInfo.Cells["B$rowAntal"].Value = $csvStats.TestCount
        $wsInfo.Cells["B$rowAntal"].Style.Numberformat.Format = '@'
        $wsInfo.Cells["B$rowAntal"].Value = "$($csvStats.TestCount)"

        if ($rowDupSample) {
            $wsInfo.Cells["B$rowDupSample"].Value = $dupSampleText

        }
        if ($rowDupCart) {
            $wsInfo.Cells["B$rowDupCart"].Value = $dupCartText
        }

        $wsInfo.Cells["B$rowInst"].Value = $instText
    } catch {
        Gui-Log ("⚠️ CSV data-fel: " + $_.Exception.Message) 'Warn'
    }

    $assayForMacro = ''
    if ($runAssay) {
        $assayForMacro = $runAssay
    } elseif ($wsOut1) {
        $assayForMacro = ($wsOut1.Cells['D10'].Text + '').Trim()
    }

    $miniVal = ''
    if (Get-Command Get-MinitabMacro -ErrorAction SilentlyContinue) {
        $miniVal = Get-MinitabMacro -AssayName $assayForMacro
    }
    if (-not $miniVal) { $miniVal = 'N/A' }

    $hdNeg = $null; $hdPos = $null
    try { $hdNeg = Get-SealHeaderDocInfo -Pkg $pkgNeg } catch {}
    try { $hdPos = Get-SealHeaderDocInfo -Pkg $pkgPos } catch {}
    if (-not $hdNeg) { $hdNeg = [pscustomobject]@{Raw='';DocNo='';Rev=''} }
    if (-not $hdPos) { $hdPos = [pscustomobject]@{Raw='';DocNo='';Rev=''} }

    $wsInfo.Cells['B2'].Value = $ScriptVersion
    $wsInfo.Cells['B3'].Value = $env:USERNAME
    $wsInfo.Cells['B4'].Value = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    $wsInfo.Cells['B5'].Value = if ($miniVal) { $miniVal } else { 'N/A' }
    $selLsp = $null

    try {
        if (Get-Variable -Name clbLsp -ErrorAction SilentlyContinue) {
            $selLsp = Get-CheckedFilePath $clbLsp
        }
    } catch {}
    $batchInfo = Get-BatchLinkInfo -SealPosPath $selPos -SealNegPath $selNeg -Lsp $lsp
    $batch = $batchInfo.Batch
    $wsInfo.Cells['A34'].Value = 'SharePoint Batch'
    $wsInfo.Cells['A34'].Style.Font.Bold = $true
    Add-Hyperlink -Cell $wsInfo.Cells['B34'] -Text $batchInfo.LinkText -Url $batchInfo.Url
    $linkMap = [ordered]@{

        'IPT App'      = 'https://apps.powerapps.com/play/e/default-771c9c47-7f24-44dc-958e-34f8713a8394/a/fd340dbd-bbbf-470b-b043-d2af4cb62c83'
        'MES Login'    = 'http://mes.cepheid.pri/camstarportal/?domain=CEPHEID.COM'
        'CSV Uploader' = 'http://auw2wgxtpap01.cepaws.com/Welcome.aspx'
        'BMRAM'        = 'https://cepheid62468.coolbluecloud.com/'
        'Agile'        = 'https://agileprod.cepheid.com/Agile/default/login-cms.jsp'
    }

    $rowLink = 35
    foreach ($key in $linkMap.Keys) {
        $wsInfo.Cells["A$rowLink"].Value = $key
        # Förkorta texten som visas i cellen till "LÄNK" enligt mallens stil
        Add-Hyperlink -Cell $wsInfo.Cells["B$rowLink"] -Text 'LÄNK' -Url $linkMap[$key]
        $rowLink++
    }

    # ----------------------------------------------------------------
#region REGION: Worksheet

    # WS (LSP Worksheet): hitta fil och skriv in i Information-bladet
    # ----------------------------------------------------------------
    try {
        if (-not $selLsp) {
            $probeDir = $null
            if ($selPos) { $probeDir = Split-Path -Parent $selPos }
            if (-not $probeDir -and $selNeg) { $probeDir = Split-Path -Parent $selNeg }
            if ($probeDir -and (Test-Path -LiteralPath $probeDir)) {
                $cand = Get-ChildItem -LiteralPath $probeDir -File -ErrorAction SilentlyContinue |
                        Where-Object {
                            ($_.Name -match '(?i)worksheet') -and ($_.Name -match [regex]::Escape($lsp)) -and ($_.Extension -match '^\.(xlsx|xlsm|xls)$')
                        } |
                        Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($cand) {
                    $selLsp = $cand.FullName
                }
            }
        }

        function Find-LabelValueRightward {
        $normLbl = Normalize-HeaderText $Label
        $pat = '^(?i)\s*' + [regex]::Escape($normLbl).Replace('\ ', '\s*') + '\s*[:\.]*\s*$'
        $rx  = [regex]::new($pat, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $hit = Find-RegexCell -Ws $Ws -Rx $rx -MaxRows $MaxRows -MaxCols $MaxCols
    if (-not $hit) { return $null }
    $cMax = [Math]::Min($Ws.Dimension.End.Column, $MaxCols)
    for ($c = $hit.Col + 1; $c -le $cMax; $c++) {
        $t = Normalize-HeaderText ($Ws.Cells[$hit.Row,$c].Text + '')
        if ($t) { return $t }
    }
    return $null
}

        if ($selLsp -and (Test-Path -LiteralPath $selLsp)) {
            Gui-Log ("🔎 WS hittad: " + (Split-Path $selLsp -Leaf)) 'Info'
        } else {
            Gui-Log "ℹ️ Ingen WS-fil vald/hittad (LSP Worksheet). Hoppar över WS-extraktion." 'Info'
        }
    } catch {
        Gui-Log ("⚠️ WS-block fel: " + $_.Exception.Message) 'Warn'
    }
        try {
            $headerWs  = $null
            $headerNeg = $null
            $headerPos = $null
                 if ($selLsp -and (Test-Path -LiteralPath $selLsp)) {
            try {
                    $tmpPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selLsp))                  
$eqInfo = $null
try {
    $eqInfo = Get-TestSummaryEquipment -Pkg $tmpPkg
    if ($eqInfo) {
        Gui-Log ("ℹ️ Utrustning hittad i WS '{0}': Pipetter={1}, Instrument={2}" -f $eqInfo.WorksheetName, ($eqInfo.Pipettes.Count), ($eqInfo.Instruments.Count)) 'Info'
    } else {
        Gui-Log "ℹ️ Utrustning gav tomt resultat." 'Info'
    }
} catch {
    Gui-Log ("⚠️ Kunde inte extrahera utrustning från Test Summary: " + $_.Exception.Message) 'Warn'
}

            $headerWs = Extract-WorksheetHeader -Pkg $tmpPkg
            $wsHeaderRows  = Get-WorksheetHeaderPerSheet -Pkg $tmpPkg
            $wsHeaderCheck = Compare-WorksheetHeaderSet   -Rows $wsHeaderRows
            try {
                if ($wsHeaderCheck.Issues -gt 0 -and $wsHeaderCheck.Summary) {
                    Gui-Log ("Worksheet header-avvikelser: {0} – se Information!" -f $wsHeaderCheck.Summary) 'Warn'
                } else {
                    Gui-Log "✅ Worksheet header korrekt" 'Info'
                }
            } catch {}
                $tmpPkg.Dispose()
                } catch {}
            }
#endregion REGION: Worksheet
#region REGION: SealTest
            try { $headerNeg = Extract-SealTestHeader -Pkg $pkgNeg } catch {}
            try { $headerPos = Extract-SealTestHeader -Pkg $pkgPos } catch {}
            try {
                if ($selLsp -and (Test-Path -LiteralPath $selLsp)) {
                    $tmpPkg2 = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selLsp))
                    $wsLsp   = $tmpPkg2.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                    if ($wsLsp) {
                        if (-not $headerWs -or -not $headerWs.PartNo) {
                            $val = $null
                            $labels = @(
                                'Part No.', 'Part No.:', 'Part No', 'Part Number', 'Part Number:', 'Part Number.', 'Part Number.:'
                            )
                            foreach ($lbl in $labels) {
                                $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl
                                if ($val) { break }
                            }
                            if ($val) { $headerWs.PartNo = $val }
                        }
                        if (-not $headerWs -or -not $headerWs.BatchNo) {
                            $val = $null
                            $labels = @(
                                'Batch No(s)', 'Batch No(s).', 'Batch No(s):', 'Batch No(s).:',
                                'Batch No', 'Batch No.', 'Batch No:', 'Batch No.:' ,
                                'Batch Number', 'Batch Number.', 'Batch Number:', 'Batch Number.:'
                            )
                            foreach ($lbl in $labels) {
                                $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl
                                if ($val) { break }
                            }
                            if ($val) { $headerWs.BatchNo = $val }
                        }
                        if (-not $headerWs -or -not $headerWs.CartridgeNo -or $headerWs.CartridgeNo -eq '.') {
                            $val = $null
                            $labels = @(
                                'Cartridge No. (LSP)', 'Cartridge No. (LSP):', 'Cartridge No. (LSP) :',
                                'Cartridge No (LSP)', 'Cartridge No (LSP):', 'Cartridge No (LSP) :',
                                'Cartridge Number (LSP)', 'Cartridge Number (LSP):', 'Cartridge Number (LSP) :',
                                'Cartridge No.', 'Cartridge No.:', 'Cartridge No. :', 'Cartridge No :',
                                'Cartridge Number', 'Cartridge Number:', 'Cartridge Number :',
                                'Cartridge No', 'Cartridge No:', 'Cartridge No :'
                            )
                            foreach ($lbl in $labels) {
                                $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl
                                if ($val) { break }
                            }
                            if (-not $val) {
                                $rxCart = [regex]::new('(?i)Cartridge.*\(LSP\)')
                                $maxCols = [Math]::Min($wsLsp.Dimension.End.Column, 100)
                                $hitCart = Find-RegexCell -Ws $wsLsp -Rx $rxCart -MaxRows 200 -MaxCols $maxCols
                                if ($hitCart) {
                                    for ($c = $hitCart.Col + 1; $c -le $wsLsp.Dimension.End.Column; $c++) {
                                        $cellVal = ($wsLsp.Cells[$hitCart.Row, $c].Text + '').Trim()
                                        if ($cellVal) { $val = $cellVal; break }
                                    }
                                }
                            }
                            if ($val) { $headerWs.CartridgeNo = $val }
                        }
                        if (-not $headerWs -or -not $headerWs.Effective) {
                            $val = Find-LabelValueRightward -Ws $wsLsp -Label 'Effective'
                            if (-not $val) { $val = Find-LabelValueRightward -Ws $wsLsp -Label 'Effective Date' }
                            if ($val) { $headerWs.Effective = $val }
                        }
            }
            try {

                if ($selLsp -and (-not $headerWs -or -not $headerWs.CartridgeNo -or $headerWs.CartridgeNo -eq '.' -or $headerWs.CartridgeNo -eq '')) {
                    $fn = Split-Path $selLsp -Leaf
                    $m = [regex]::Matches($fn, '(?<!\d)(\d{5,7})(?!\d)')
                    if ($m.Count -gt 0) {
                        $headerWs.CartridgeNo = $m[0].Groups[1].Value
                    }
                }
            } catch {}
                    $tmpPkg2.Dispose()
                }
            } catch {}
            try {
                if ($pkgPos -and -not $headerPos.Effective) {
                    $wsPos = $pkgPos.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                    if ($wsPos) {
                        $val = Find-LabelValueRightward -Ws $wsPos -Label 'Effective'
                        if (-not $val) { $val = Find-LabelValueRightward -Ws $wsPos -Label 'Effective Date' }
                        if ($val) { $headerPos.Effective = $val }
                    }
                }
            } catch {}
            try {
                if ($pkgNeg -and -not $headerNeg.Effective) {
                    $wsNeg = $pkgNeg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                    if ($wsNeg) {
                        $val = Find-LabelValueRightward -Ws $wsNeg -Label 'Effective'
                        if (-not $val) { $val = Find-LabelValueRightward -Ws $wsNeg -Label 'Effective Date' }
                        if ($val) { $headerNeg.Effective = $val }
                    }
                }
            } catch {}
            $wsBatch   = if ($headerWs -and $headerWs.BatchNo) { $headerWs.BatchNo } else { $null }
            $sealBatch = $batch
            if (-not $sealBatch) {
                try { if ($selPos) { $sealBatch = Get-BatchNumberFromSealFile $selPos } } catch {}
                if (-not $sealBatch) { try { if ($selNeg) { $sealBatch = Get-BatchNumberFromSealFile $selNeg } } catch {} }
            }
            $batchMatchFlag = $null
            if ($wsBatch -and $sealBatch) { $batchMatchFlag = ($wsBatch -eq $sealBatch) }
            $sealConsistentFlag = $null
            if ($headerNeg -and $headerPos) {
                if ($headerNeg.DocumentNumber -and $headerPos.DocumentNumber -and $headerNeg.Rev -and $headerPos.Rev -and $headerNeg.Effective -and $headerPos.Effective) {
                    $sealConsistentFlag = (($headerNeg.DocumentNumber -eq $headerPos.DocumentNumber) -and ($headerNeg.Rev -eq $headerPos.Rev) -and ($headerNeg.Effective -eq $headerPos.Effective))
                }
            }
            $noteStr = ''
            if ($headerNeg -and $headerNeg.DocumentNumber -and $headerNeg.DocumentNumber -ne 'D10552') { $noteStr += ("NEG DocNo (" + $headerNeg.DocumentNumber + ") != D10552; ") }
            if ($headerPos -and $headerPos.DocumentNumber -and $headerPos.DocumentNumber -ne 'D10552') { $noteStr += ("POS DocNo (" + $headerPos.DocumentNumber + ") != D10552; ") }
            $rowWsFile = Find-InfoRow -Ws $wsInfo -Label 'Worksheet'
            if (-not $rowWsFile) { $rowWsFile = 17 }
            $rowPart  = $rowWsFile + 1
            $rowBatch = $rowWsFile + 2
            $rowCart  = $rowWsFile + 3
            $rowDoc   = $rowWsFile + 4
            $rowRev   = $rowWsFile + 5
            $rowEff   = $rowWsFile + 6
            $rowPosFile = Find-InfoRow -Ws $wsInfo -Label 'Seal Test POS'
            if (-not $rowPosFile) {
                if ($rowWsFile) { $rowPosFile = $rowWsFile + 7 } else { $rowPosFile = 24 }
            }
            $rowPosDoc = $rowPosFile + 1
            $rowPosRev = $rowPosFile + 2
            $rowPosEff = $rowPosFile + 3
            $rowNegFile = Find-InfoRow -Ws $wsInfo -Label 'Seal Test NEG'
            if (-not $rowNegFile) {
                $rowNegFile = $rowPosFile + 4
            }
            $rowNegDoc = $rowNegFile + 1
            $rowNegRev = $rowNegFile + 2
            $rowNegEff = $rowNegFile + 3
            if ($selLsp) {
                $wsInfo.Cells["B$rowWsFile"].Style.Numberformat.Format = '@'
                $wsInfo.Cells["B$rowWsFile"].Value = (Split-Path $selLsp -Leaf)
            } else {
                $wsInfo.Cells["B$rowWsFile"].Value = ''
            }

            $consPart  = Get-ConsensusValue -Type 'Part'      -Ws $headerWs.PartNo      -Pos $headerPos.PartNumber   -Neg $headerNeg.PartNumber
            $consBatch = Get-ConsensusValue -Type 'Batch'     -Ws $headerWs.BatchNo     -Pos $headerPos.BatchNumber  -Neg $headerNeg.BatchNumber
            $consCart  = Get-ConsensusValue -Type 'Cartridge' -Ws $headerWs.CartridgeNo -Pos $headerPos.CartridgeNo  -Neg $headerNeg.CartridgeNo
 
            if (-not $consCart.Value -and $selLsp) {
                $fnCart = Split-Path $selLsp -Leaf
                $mCart  = [regex]::Match($fnCart,'(?<!\d)(\d{5,7})(?!\d)')
                if ($mCart.Success) {
                    $consCart = @{
                        Value  = $mCart.Groups[1].Value
                        Source = 'FILENAME'
                        Note   = 'Filename fallback'
                    }
                }
            }

            if ($consPart.Value)  { $wsInfo.Cells["B$rowPart"].Value = $consPart.Value }  else { $wsInfo.Cells["B$rowPart"].Value = '' }
            if ($consBatch.Value) { $wsInfo.Cells["B$rowBatch"].Value = $consBatch.Value } else { $wsInfo.Cells["B$rowBatch"].Value = '' }
            if ($consCart.Value)  { $wsInfo.Cells["B$rowCart"].Value = $consCart.Value }  else { $wsInfo.Cells["B$rowCart"].Value = '' }

            $batchMismatch = $false
            try {
                if ($headerNeg -and $headerPos -and $headerNeg.BatchNumber -and $headerPos.BatchNumber) {
                    $normNeg = Normalize-Id -Value $headerNeg.BatchNumber -Type 'Batch'
                    $normPos = Normalize-Id -Value $headerPos.BatchNumber -Type 'Batch'
                    if ($normNeg -and $normPos -and $normNeg -ne $normPos) { $batchMismatch = $true }
                }
            } catch {}
 
            if ($batchMismatch) {
                try { if ($consPart.Note)  { [void]$wsInfo.Cells["B$rowPart"].AddComment($consPart.Note,  'DocMerge') } } catch {}
                try { if ($consBatch.Note) { [void]$wsInfo.Cells["B$rowBatch"].AddComment($consBatch.Note, 'DocMerge') } } catch {}
                try { if ($consCart.Note)  { [void]$wsInfo.Cells["B$rowCart"].AddComment($consCart.Note,  'DocMerge') } } catch {}
            }

            try {
                if ($wsHeaderCheck -and $wsHeaderCheck.Details) {
                    $linesDev = ($wsHeaderCheck.Details -split "`r?`n")
                    $devPart  = $null
                    $devBatch = $null
                    $devCart  = $null
                    foreach ($ln in $linesDev) {
                        if ($ln -match '^-\s*PartNo[^:]*:\s*(.+)$') {
                            $devPart = $matches[1].Trim()
                        } elseif ($ln -match '^-\s*BatchNo[^:]*:\s*(.+)$') {
                            $devBatch = $matches[1].Trim()
                        } elseif ($ln -match '^-\s*CartridgeNo[^:]*:\s*(.+)$') {
                            $devCart = $matches[1].Trim()
                        }
                    }
                    if ($devPart) {
                        $wsInfo.Cells["C$rowPart"].Style.Numberformat.Format = '@'
                        $wsInfo.Cells["C$rowPart"].Value = 'Avvikande flik: ' + $devPart
                    }
                    if ($devBatch) {
                        $wsInfo.Cells["C$rowBatch"].Style.Numberformat.Format = '@'
                        $wsInfo.Cells["C$rowBatch"].Value = 'Avvikande flik: ' + $devBatch
                    }
                    if ($devCart) {
                        $wsInfo.Cells["C$rowCart"].Style.Numberformat.Format = '@'
                        $wsInfo.Cells["C$rowCart"].Value = 'Avvikande flik: ' + $devCart
                    }
                }
            } catch {}

            if ($headerWs) {
                $doc = $headerWs.DocumentNumber
                if ($doc) {
                    $doc = ($doc -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$', '').Trim()
                }
                if ($headerWs.Attachment -and ($doc -notmatch '(?i)\bAttachment\s+\w+\b')) {
                    $doc = "$doc Attachment $($headerWs.Attachment)"
                }
                $wsInfo.Cells["B$rowDoc"].Value = $doc
                $wsInfo.Cells["B$rowRev"].Value = $headerWs.Rev
                $wsInfo.Cells["B$rowEff"].Value = $headerWs.Effective
            } else {
                $wsInfo.Cells["B$rowDoc"].Value = ''
                $wsInfo.Cells["B$rowRev"].Value = ''
                $wsInfo.Cells["B$rowEff"].Value = ''
            }
 
            if ($selPos) {
                $wsInfo.Cells["B$rowPosFile"].Style.Numberformat.Format = '@'
                $wsInfo.Cells["B$rowPosFile"].Value = (Split-Path $selPos -Leaf)
            } else {
                $wsInfo.Cells["B$rowPosFile"].Value = ''
            }

            if ($headerPos) {
                $docPos = $headerPos.DocumentNumber
                if ($docPos) { $docPos = ($docPos -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$','').Trim() }
                $wsInfo.Cells["B$rowPosDoc"].Value = $docPos
                $wsInfo.Cells["B$rowPosRev"].Value = $headerPos.Rev
                $wsInfo.Cells["B$rowPosEff"].Value = $headerPos.Effective
            } else {
                $wsInfo.Cells["B$rowPosDoc"].Value = ''
                $wsInfo.Cells["B$rowPosRev"].Value = ''
                $wsInfo.Cells["B$rowPosEff"].Value = ''
            }
            if ($selNeg) {
                $wsInfo.Cells["B$rowNegFile"].Style.Numberformat.Format = '@'
                $wsInfo.Cells["B$rowNegFile"].Value = (Split-Path $selNeg -Leaf)
            } else {
                $wsInfo.Cells["B$rowNegFile"].Value = ''
            }
            # Seal Test NEG metadata
            if ($headerNeg) {
                # NEG: ta bort ev. "Rev/Effective" som följt med
                $docNeg = $headerNeg.DocumentNumber
                if ($docNeg) { $docNeg = ($docNeg -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$','').Trim() }
                $wsInfo.Cells["B$rowNegDoc"].Value = $docNeg
                $wsInfo.Cells["B$rowNegRev"].Value = $headerNeg.Rev
                $wsInfo.Cells["B$rowNegEff"].Value = $headerNeg.Effective
            } else {
                $wsInfo.Cells["B$rowNegDoc"].Value = ''
                $wsInfo.Cells["B$rowNegRev"].Value = ''
                $wsInfo.Cells["B$rowNegEff"].Value = ''
            }
#endregion REGION: SealTest
            # Töm eventuella överflödiga rader nedanför tabellen – ej nödvändig då layout definierad i mall
        } catch {
            Gui-Log ("⚠️ Header summary fel: " + $_.Exception.Message) 'Warn'
        }
} catch {
    Gui-Log "⚠️ Information-blad fel: $($_.Exception.Message)" 'Warn'
}

        # ============================
        # === Information2 (Rules)  ===
        # ============================
        $wsInfo2 = $null
        try {
            $wsInfo2 = $pkgOut.Workbook.Worksheets['Information2']
            if (-not $wsInfo2) { $wsInfo2 = $pkgOut.Workbook.Worksheets.Add('Information2') }
        } catch {
            Gui-Log "⚠️ Information2 (Rules) kunde inte skapa blad: $($_.Exception.Message)" 'Warn'
        }

        if ($wsInfo2) {
            $ctx = $null
            try {
                if (Get-Command New-AssayRuleContext -ErrorAction SilentlyContinue) {
                    $assayMap = $null
                    if ($Config -and $Config.AssayMap) { $assayMap = $Config.AssayMap }
                    elseif ($Config -and $Config.SlangAssay) { $assayMap = $Config.SlangAssay }

                    $ctx = New-AssayRuleContext -Bundle $csvBundle -AssayMap $assayMap -RuleBank $global:RuleBank
                    $eval = Invoke-AssayRuleEngine -Context $ctx -RuleBank $global:RuleBank
                    Write-GoldenStandardReport -Package $pkgOut -Context $ctx -Evaluation $eval -CsvPath $selCsv -ScriptVersion $ScriptVersion
                } else {
                    $fallbackEval = [pscustomobject]@{
                        OverallStatus    = 'FAIL'
                        OverallSeverity  = 'Error'
                        Findings         = @([pscustomobject]@{ Severity='Error'; RuleId='RULEENGINE_UNAVAILABLE'; Message='RuleEngine functions not loaded'; Count=0; Example=''; Evidence=''; Classification='Engine'; GeneratesRetest=$false })
                        AffectedTests    = @()
                        AffectedTestsTruncated = 0
                        ErrorSummary     = @()
                        ErrorCodesTable  = @()
                        PressureStats    = [ordered]@{ Max=$null; Avg=$null; OverWarn=0; OverFail=0; WarnThreshold=$null; FailThreshold=$null }
                        DuplicatesTable  = @()
                        Debug            = @()
                        UniqueErrorCodes = 0
                        IdentityFlags    = [ordered]@{}
                        SeverityCounts   = [ordered]@{ Error=1; Warn=0; Info=0 }
                    }
                    Write-GoldenStandardReport -Package $pkgOut -Context $ctx -Evaluation $fallbackEval -CsvPath $selCsv -ScriptVersion $ScriptVersion
                }
            } catch {
                    $msg = $_.Exception.Message
    $pos = $_.InvocationInfo.PositionMessage
    $stk = $_.ScriptStackTrace
    Gui-Log ("⚠️ ⚠️ Information2 (Rules) fel: {0}`n{1}`n{2}" -f $msg, $pos, $stk) 'Warn'
                $evidence = $null
                try { $evidence = $_.Exception.ToString() } catch {}
                $fallbackEval = [pscustomobject]@{
                    OverallStatus    = 'FAIL'
                    OverallSeverity  = 'Error'
                    Findings         = @([pscustomobject]@{ Severity='Error'; RuleId='RULEENGINE_EXCEPTION'; Message='RuleEngine failed'; Count=0; Example=''; Evidence=$evidence; Classification='Engine'; GeneratesRetest=$false })
                    AffectedTests    = @()
                    AffectedTestsTruncated = 0
                    ErrorSummary     = @()
                    ErrorCodesTable  = @()
                    PressureStats    = [ordered]@{ Max=$null; Avg=$null; OverWarn=0; OverFail=0; WarnThreshold=$null; FailThreshold=$null }
                    DuplicatesTable  = @()
                    Debug            = @()
                    UniqueErrorCodes = 0
                    IdentityFlags    = [ordered]@{}
                    SeverityCounts   = [ordered]@{ Error=1; Warn=0; Info=0 }
                }
                if (-not $ctx) {
                    $ctx = [pscustomobject]@{
                        TotalTests           = 0
                        StatusCounts         = @{}
                        AssayRaw             = ''
                        AssayCanonical       = ''
                        AssayVersion         = ''
                        ReagentLotIds        = @()
                        WorkCenters          = @()
                        MaxPressureMax       = $null
                        MaxPressureAvg       = $null
                        UniqueSampleIds      = 0
                        UniqueCartridgeSN    = 0
                        DuplicateCounts      = [ordered]@{ SampleId=0; CartridgeSN=0 }
                        AssayCanonicalSource = 'Error'
                    }
                }
                Write-GoldenStandardReport -Package $pkgOut -Context $ctx -Evaluation $fallbackEval -CsvPath $selCsv -ScriptVersion $ScriptVersion
            }
        }


        # ============================
        # === Equipment-blad       ===
        # ============================
        try {
            if (Test-Path -LiteralPath $UtrustningListPath) {
                $srcPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($UtrustningListPath))
                try {
                    $srcWs = $srcPkg.Workbook.Worksheets['Sheet1']
                    if (-not $srcWs) {
                        $srcWs = $srcPkg.Workbook.Worksheets[1]
                    }

                    if ($srcWs) {
                        $wsEq = $pkgOut.Workbook.Worksheets['Infinity/GX']
                        if ($wsEq) {
                            $pkgOut.Workbook.Worksheets.Delete($wsEq)
                        }
                        $wsEq = $pkgOut.Workbook.Worksheets.Add('Infinity/GX', $srcWs)

                        if ($wsEq.Dimension) {
                            foreach ($cell in $wsEq.Cells[$wsEq.Dimension.Address]) {
                                if ($cell.Formula -or $cell.FormulaR1C1) {
                                    $val = $cell.Value
                                    $cell.Formula     = $null
                                    $cell.FormulaR1C1 = $null
                                    $cell.Value       = $val
                                }
                            }
                            $colCount = $srcWs.Dimension.End.Column
                            for ($c = 1; $c -le $colCount; $c++) {
                                try {
                                    $wsEq.Column($c).Width = $srcWs.Column($c).Width
                                } catch {
                                }
                            }
                        }

                        if ($eqInfo) {
                            $wsName = $null
                            if ($selLsp) {
                                $wsName = Split-Path $selLsp -Leaf
                            } elseif ($eqInfo.PSObject.Properties['WorksheetName'] -and $eqInfo.WorksheetName) {
                                $wsName = $eqInfo.WorksheetName
                            } elseif ($headerWs -and $headerWs.PSObject.Properties['WorksheetName'] -and $headerWs.WorksheetName) {
                                $wsName = $headerWs.WorksheetName
                            } else {
                                $wsName = 'Test Summary'
                            }

                            $cellHeaderPip = $wsEq.Cells['A14']
                            $cellHeaderPip.Value = "PIPETTER hämtade från $wsName"
                            $cellHeaderPip.Style.Font.Bold = $true
                            $cellHeaderPip.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                            $cellHeaderPip.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

                            $cellHeaderInst = $wsEq.Cells['A18']
                            $cellHeaderInst.Value = "INSTRUMENT hämtade från $wsName"
                            $cellHeaderInst.Style.Font.Bold = $true
                            $cellHeaderInst.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                            $cellHeaderInst.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

                            function Convert-ToEqDate {
                                param(
                                    [Parameter(Mandatory = $false)]
                                    $Value
                                )

                                if (-not $Value -or $Value -eq 'N/A') {
                                    return $null
                                }

                                if ($Value -is [datetime]) {
                                    return $Value
                                }

                                if ($Value -is [double] -or $Value -is [int]) {
                                    try {
                                        $base = Get-Date '1899-12-30'
                                        return $base.AddDays([double]$Value)
                                    } catch {
                                        return $Value
                                    }
                                }

                                try {
                                    return (Get-Date -Date $Value -ErrorAction Stop)
                                } catch {
                                    return $Value
                                }
                            }

                            $pipetteIdCells  = @('B15','D15','F15','H15','J15','L15')
                            $pipetteDueCells = @('B16','D16','F16','H16','J16','L16')

                            $pipettes = @()
                            if ($eqInfo.PSObject.Properties['Pipettes'] -and $eqInfo.Pipettes) {
                                $pipettes = @($eqInfo.Pipettes)
                            }

                            for ($i = 0; $i -lt $pipetteIdCells.Count; $i++) {
                                $cellId  = $wsEq.Cells[$pipetteIdCells[$i]]
                                $cellDue = $wsEq.Cells[$pipetteDueCells[$i]]

                                if ($i -lt $pipettes.Count) {
                                    $p = $pipettes[$i]

                                    $id  = $null
                                    $due = $null

                                    if ($p -is [string]) {
                                        $id = $p
                                    } else {
                                        $idCandidates = @()
                                        foreach ($propName in 'Id','CepheidId','Name','PipetteId') {
                                            if ($p.PSObject.Properties[$propName]) {
                                                $idCandidates += $p.$propName
                                            }
                                        }
                                        $id = $idCandidates | Where-Object { $_ } | Select-Object -First 1

                                        $dueCandidates = @()
                                        foreach ($propName in 'CalibrationDueDate','DueDate','CalDue') {
                                            if ($p.PSObject.Properties[$propName]) {
                                                $dueCandidates += $p.$propName
                                            }
                                        }
                                        $due = $dueCandidates | Where-Object { $_ } | Select-Object -First 1
                                    }

                                    if ([string]::IsNullOrWhiteSpace($id) -or $id -eq 'N/A') {
                                        $cellId.Value  = 'N/A'
                                        $cellDue.Value = 'N/A'
                                    } else {
                                        $cellId.Value = $id

                                        $dt = Convert-ToEqDate -Value $due
                                        if ($dt -is [datetime]) {
                                            $cellDue.Value = $dt
                                            $cellDue.Style.Numberformat.Format = 'mmm-yy'
                                        } elseif ($dt) {
                                            $cellDue.Value = $dt
                                        } else {
                                            $cellDue.Value = 'N/A'
                                        }
                                    }
                                } else {
                                    $cellId.Value  = 'N/A'
                                    $cellDue.Value = 'N/A'
                                }

                                foreach ($c in @($cellId,$cellDue)) {
                                    $c.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                                    $c.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
                                }
                            }

                            $instIdCells  = @(
                                'B19','D19','F19','H19','J19','L19',
                                'B21','D21','F21','H21','J21','L21'
                            )
                            $instDueCells = @(
                                'B20','D20','F20','H20','J20','L20',
                                'B22','D22','F22','H22','J22','L22'
                            )

                            $instruments = @()
                            if ($eqInfo.PSObject.Properties['Instruments'] -and $eqInfo.Instruments) {
                                $instruments = @($eqInfo.Instruments)
                            }

                            for ($i = 0; $i -lt $instIdCells.Count; $i++) {
                                $cellId  = $wsEq.Cells[$instIdCells[$i]]
                                $cellDue = $wsEq.Cells[$instDueCells[$i]]

                                if ($i -lt $instruments.Count) {
                                    $inst = $instruments[$i]

                                    $id  = $null
                                    $due = $null

                                    if ($inst -is [string]) {
                                        $id = $inst
                                    } else {
                                        $idCandidates = @()
                                        foreach ($propName in 'Id','CepheidId','Name','InstrumentId') {
                                            if ($inst.PSObject.Properties[$propName]) {
                                                $idCandidates += $inst.$propName
                                            }
                                        }
                                        $id = $idCandidates | Where-Object { $_ } | Select-Object -First 1

                                        $dueCandidates = @()
                                        foreach ($propName in 'CalibrationDueDate','DueDate','CalDue') {
                                            if ($inst.PSObject.Properties[$propName]) {
                                                $dueCandidates += $inst.$propName
                                            }
                                        }
                                        $due = $dueCandidates | Where-Object { $_ } | Select-Object -First 1
                                    }

                                    if ([string]::IsNullOrWhiteSpace($id) -or $id -eq 'N/A') {
                                        $cellId.Value  = 'N/A'
                                        $cellDue.Value = 'N/A'
                                    } else {
                                        $cellId.Value = $id

                                        $dt = Convert-ToEqDate -Value $due
                                        if ($dt -is [datetime]) {
                                            $cellDue.Value = $dt
                                            $cellDue.Style.Numberformat.Format = 'mmm-yy'
                                        } elseif ($dt) {
                                            $cellDue.Value = $dt
                                        } else {
                                            $cellDue.Value = 'N/A'
                                        }
                                    }
                                } else {
                                    $cellId.Value  = 'N/A'
                                    $cellDue.Value = 'N/A'
                                }

                                foreach ($c in @($cellId,$cellDue)) {
                                    $c.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                                    $c.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
                                }
                            }

                            if ($pipettes.Count -gt $pipetteIdCells.Count -or
                                $instruments.Count -gt $instIdCells.Count) {
                                Gui-Log ("ℹ️ Infinity/GX: allt får inte plats i mallen (pipetter={0}, instrument={1})" -f $pipettes.Count, $instruments.Count) 'Info'
                            }

                        } else {
                            Gui-Log "ℹ️ Utrustning saknas – Infinity/GX lämnas som mall." 'Info'
                        }
                    }
                }
                finally {
                    if ($srcPkg) { $srcPkg.Dispose() }
                }
            } else {
                Gui-Log "ℹ️ Infinity/GX mall saknas: $($_.Exception.Message)" 'Info'
            }
        }
        catch {
            Gui-Log "⚠️ Kunde inte skapa 'Infinity/GX'-flik: $($_.Exception.Message)" 'Warn'
        }

        # ============================
        # === Control Material     ===
        # ============================
        try {
            if ($controlTab -and (Test-Path -LiteralPath $RawDataPath)) {
                $srcPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($RawDataPath))
                try { $srcPkg.Workbook.Calculate() } catch {}
                $candidates = if ($controlTab -match '\|') { $controlTab -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @($controlTab) }
                $srcWs = $null
                foreach ($cand in $candidates) {

                    $srcWs = $srcPkg.Workbook.Worksheets | Where-Object { $_.Name -eq $cand } | Select-Object -First 1
                    if ($srcWs) { break }
                    $srcWs = $srcPkg.Workbook.Worksheets | Where-Object { $_.Name -like "*$cand*" } | Select-Object -First 1
                    if ($srcWs) { break }
                }
                if ($srcWs) {
                    $safeName = if ($srcWs.Name.Length -gt 31) { $srcWs.Name.Substring(0,31) } else { $srcWs.Name }
                    $destName = $safeName; $n=1
                    while ($pkgOut.Workbook.Worksheets[$destName]) { $base = if ($safeName.Length -gt 27) { $safeName.Substring(0,27) } else { $safeName }; $destName = "$base($n)"; $n++ }
                    $wsCM = $pkgOut.Workbook.Worksheets.Add($destName, $srcWs)
                    if ($wsCM.Dimension) {
                        foreach ($cell in $wsCM.Cells[$wsCM.Dimension.Address]) {
                            if ($cell.Formula -or $cell.FormulaR1C1) { $v=$cell.Value; $cell.Formula=$null; $cell.FormulaR1C1=$null; $cell.Value=$v }
                        }
                        try { $wsCM.Cells[$wsCM.Dimension.Address].AutoFitColumns() | Out-Null } catch {}
                    }
                    Gui-Log "✅ Control Material kopierad: '$($srcWs.Name)' → '$destName'" 'Info'
                } else { Gui-Log "ℹ️ Hittade inget blad i kontrollfilen som matchar '$controlTab'." 'Info' }
                $srcPkg.Dispose()
            } else { Gui-Log "ℹ️ Ingen Control-flik skapad (saknar mappning eller kontrollfil)." 'Info' }
        } catch { Gui-Log "⚠️ Control Material-fel: $($_.Exception.Message)" 'Warn' }

        # ============================
        # === SharePoint Info      ===
        # ============================
        try {

            if ($chkSharePointInfo -and -not $chkSharePointInfo.Checked) {
                Gui-Log "ℹ️ SharePoint Info ej valt – hoppar över." 'Info'
                try { $old = $pkgOut.Workbook.Worksheets["SharePoint Info"]; if ($old) { $pkgOut.Workbook.Worksheets.Delete($old) } } catch {}
            } else {
                $spOk = $false
                if ($global:SpConnected) { $spOk = $true }
                elseif (Get-Command Get-PnPConnection -ErrorAction SilentlyContinue) {
                    try { $null = Get-PnPConnection; $spOk = $true } catch { $spOk = $false }
                }
 
                if (-not $spOk) {
                    $errMsg = if ($global:SpError) { $global:SpError } else { 'Okänt fel' }
                    Gui-Log ("⚠️ SharePoint ej tillgängligt: $errMsg") 'Warn'
                }

                $batchInfo = Get-BatchLinkInfo -SealPosPath $selPos -SealNegPath $selNeg -Lsp $lsp
                $batch = $batchInfo.Batch

                if (-not $batch) {
                    Gui-Log "ℹ️ Inget Batch # i POS/NEG – skriver tom SharePoint Info." 'Info'
                    if (Get-Command Write-SPSheet-Safe -ErrorAction SilentlyContinue) {
                        [void](Write-SPSheet-Safe -Pkg $pkgOut -Rows @() -DesiredOrder @() -Batch '—')
                    } else {
                        $wsSp = $pkgOut.Workbook.Worksheets["SharePoint Info"]; if ($wsSp) { $pkgOut.Workbook.Worksheets.Delete($wsSp) }
                        $wsSp = $pkgOut.Workbook.Worksheets.Add("SharePoint Info")
                        $wsSp.Cells[1,1].Value = "Rubrik"; $wsSp.Cells[1,2].Value = "Värde"
                        $wsSp.Cells[2,1].Value = "Batch";  $wsSp.Cells[2,2].Value = "—"
                        try { $wsSp.Cells[$wsSp.Dimension.Address].AutoFitColumns() | Out-Null } catch {}
                    }
                } else {
                    Gui-Log "🔎 Batch hittad: $batch" 'Info'

                    $fields = @(
                        'Work_x0020_Center','Title','Batch_x0023_','SAP_x0020_Batch_x0023__x0020_2',
                        'LSP','Material','BBD_x002f_SLED','Actual_x0020_startdate_x002f__x0',
                        'PAL_x0020__x002d__x0020_Sample_x','Sample_x0020_Reagent_x0020_P_x00',
                        'Order_x0020_quantity','Total_x0020_good','ITP_x0020_Test_x0020_results',
                        'IPT_x0020__x002d__x0020_Testing_0','MES_x0020__x002d__x0020_Order_x0'
                    )
                    $renameMap = @{
                        'Work Center'            = 'Work Center'
                        'Title'                  = 'Order#'
                        'Batch#'                 = 'SAP Batch#'
                        'SAP Batch# 2'           = 'SAP Batch# 2'
                        'LSP'                    = 'LSP'
                        'Material'               = 'Material'
                        'BBD/SLED'               = 'BBD/SLED'
                        'Actual startdate/_x0'   = 'ROBAL - Actual start date/time'
                        'PAL - Sample_x'         = 'Sample Reagent use'
                        'Sample Reagent P'       = 'Sample Reagent P/N'
                        'Order quantity'         = 'Order quantity'
                        'Total good'             = 'ROBAL - Till Packning'
                        'IPT Test results'       = 'IPT Test results'
                        'IPT - Testing_0'        = 'IPT - Testing Finalized'
                        'MES - Order_x0'         = 'MES Order'
                    }

                    $desiredOrder = @(
                        'Work Center','Order#','SAP Batch#','SAP Batch# 2','LSP','Material','BBD/SLED',
                        'ROBAL - Actual start date/time','Sample Reagent use','Sample Reagent P/N',
                        'Order quantity','ROBAL - Till Packning','IPT Test results',
                        'IPT - Testing Finalized','MES Order'
                    )

                    $dateFields      = @('BBD/SLED','ROBAL - Actual start date/time','IPT - Testing Finalized')
                    $shortDateFields = @('BBD/SLED')
                    $rows = @()
                    if ($spOk) {
                        try {
                            $items = Get-PnPListItem -List "Cepheid | Production orders" -Fields $fields -PageSize 2000 -ErrorAction Stop
                            $match = $items | Where-Object {
                                $v1 = $_['Batch_x0023_']; $v2 = $_['SAP_x0020_Batch_x0023__x0020_2']
                                $s1 = if ($null -ne $v1) { ([string]$v1).Trim() } else { '' }
                                $s2 = if ($null -ne $v2) { ([string]$v2).Trim() } else { '' }
                                $s1 -eq $batch -or $s2 -eq $batch
                            } | Select-Object -First 1
                            if ($match) {
                                foreach ($f in $fields) {
                                    $val = $match[$f]
                                    $label = $f -replace '_x0020_', ' ' `
                                                 -replace '_x002d_', '-' `
                                                 -replace '_x0023_', '#' `
                                                 -replace '_x002f_', '/' `
                                                 -replace '_x2013_', '–' `
                                                 -replace '_x00',''
                                    $label = $label.Trim()
                                    if ($renameMap.ContainsKey($label)) { $label = $renameMap[$label] }
                                    if ($null -ne $val -and $val -ne '') {
                                        if ($val -eq $true) { $val = 'JA' }
                                        elseif ($val -eq $false) { $val = 'NEJ' }

                                        $dt = $null
                                        if ($val -is [datetime]) { $dt = [datetime]$val }
                                        else { try { $dt = [datetime]::Parse($val) } catch { $dt = $null } }
                                        if ($dt -ne $null -and ($dateFields -contains $label)) {
                                            $fmt = if ($shortDateFields -contains $label) { 'yyyy-MM-dd' } else { 'yyyy-MM-dd HH:mm' }
                                            $val = $dt.ToString($fmt)
                                        }
                                        $rows += [pscustomobject]@{ Rubrik = $label; 'Värde' = $val }
                                    }
                                }
 
                                if ($rows.Count -gt 0) {
                                    $ordered = @()
                                    foreach ($label in $desiredOrder) {
                                        $hit = $rows | Where-Object { $_.Rubrik -eq $label } | Select-Object -First 1
                                        if ($hit) { $ordered += $hit }
                                    }
                                    if ($ordered.Count -gt 0) { $rows = $ordered }
                                }
                                Gui-Log "📄 SharePoint-post hittad – skriver blad." 'Info'
                            } else {
                                Gui-Log "ℹ️ Ingen post i SharePoint för Batch=$batch." 'Info'
                            }
                        } catch {
                            Gui-Log "⚠️ SP: Get-PnPListItem misslyckades: $($_.Exception.Message)" 'Warn'
                        }
                    }
                    if (Get-Command Write-SPSheet-Safe -ErrorAction SilentlyContinue) {
                        [void](Write-SPSheet-Safe -Pkg $pkgOut -Rows $rows -DesiredOrder $desiredOrder -Batch $batch)
                    } else {
                        $wsSp = $pkgOut.Workbook.Worksheets["SharePoint Info"]; if ($wsSp) { $pkgOut.Workbook.Worksheets.Delete($wsSp) }
                        $wsSp = $pkgOut.Workbook.Worksheets.Add("SharePoint Info")
                        $wsSp.Cells[1,1].Value = "Rubrik"; $wsSp.Cells[1,2].Value = "Värde"
                        if ($rows.Count -gt 0) {
                            $r=2; foreach($rowObj in $rows) { $wsSp.Cells[$r,1].Value = $rowObj.Rubrik; $wsSp.Cells[$r,2].Value = $rowObj.'Värde'; $r++ }
                        } else {
                            $wsSp.Cells[2,1].Value = "Batch";  $wsSp.Cells[2,2].Value = $batch
                            $wsSp.Cells[3,1].Value = "Info";   $wsSp.Cells[3,2].Value = "No matching SharePoint row"
                        }
                        try { $wsSp.Cells[$wsSp.Dimension.Address].AutoFitColumns() | Out-Null } catch {}
                    }
                    try {
                        if ($slBatchLink -and $batch) {
                            $slBatchLink.Text = "SharePoint: $batch"
                            $slBatchLink.Tag  = $batchInfo.Url
                            $slBatchLink.Enabled = $true
                        }
                    } catch {}
                    try {
                        $wsSP = $pkgOut.Workbook.Worksheets['SharePoint Info']
                        if ($wsSP -and $wsSP.Dimension) {
                            $labelCol = 1; $valueCol = 2
                            for ($r = 1; $r -le $wsSP.Dimension.End.Row; $r++) {
                                if (($wsSP.Cells[$r,$labelCol].Text).Trim() -eq 'Sample Reagent use') {
                                    $wsSP.Cells[$r,$valueCol].Style.WrapText = $true
                                    $wsSP.Cells[$r,$valueCol].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
                                    try { $wsSP.Column($valueCol).Width = 55 } catch {}
                                    $wsSP.Row($r).CustomHeight = $true
                                    break
                                }
                            }
                        }
                    } catch {
                        Gui-Log "⚠️ WrapText på 'Sample Reagent use' misslyckades: $($_.Exception.Message)" 'Warn'
                    }
                }
            }
        } catch {
            Gui-Log "⚠️ SP-blad: $($_.Exception.Message)" 'Warn'
        }

        # ============================
        # === Header watermark     ===
        # ============================

        try {
            foreach ($ws in $pkgOut.Workbook.Worksheets) {
                try {
                    $ws.HeaderFooter.OddHeader.CenteredText   = '&"Arial,Bold"&14 UNCONTROLLED'
                    $ws.HeaderFooter.EvenHeader.CenteredText  = '&"Arial,Bold"&14 UNCONTROLLED'
                    $ws.HeaderFooter.FirstHeader.CenteredText = '&"Arial,Bold"&14 UNCONTROLLED'
                } catch { Write-Warning "Kunde inte sätta header på blad: $($ws.Name)" }
            }
        } catch { Write-Warning "Fel vid vattenstämpling av rapporten." }

        # ============================
        # === Tab-färger (innan Save)
        # ============================

        try {
            $wsT = $pkgOut.Workbook.Worksheets['Information'];            if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 52, 152, 219) }
            $wsT = $pkgOut.Workbook.Worksheets['Infinity/GX'];            if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 33, 115, 70) }
            $wsT = $pkgOut.Workbook.Worksheets['SharePoint Info'];        if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 0, 120, 212) }
        } catch {
            Gui-Log "⚠️ Kunde inte sätta tab-färg: $($_.Exception.Message)" 'Warn'
        }

        # ============================
        # === Spara & Audit        ===
        # ============================

        $nowTs   = Get-Date -Format "yyyyMMdd_HHmmss"
        $baseName = "$($env:USERNAME)_output_${lsp}_$nowTs.xlsx"
        if ($rbSaveInLsp.Checked) {
            $saveDir = Split-Path -Parent $selNeg
            $SavePath = Join-Path $saveDir $baseName
            Gui-Log "💾 Sparläge: LSP-mapp → $saveDir"
        } else {
            $saveDir = $env:TEMP
            $SavePath = Join-Path $saveDir $baseName
            Gui-Log "💾 Sparläge: Temporärt → $SavePath"
        }
        try {
            $pkgOut.Workbook.View.ActiveTab = 0
            $wsInitial = $pkgOut.Workbook.Worksheets["Information"]
            if ($wsInitial) { $wsInitial.View.TabSelected = $true }
            $pkgOut.SaveAs($SavePath)
            Gui-Log "✅ Rapport sparad: $SavePath" 'Info'
            $global:LastReportPath = $SavePath

            try {
                $auditDir = Join-Path $PSScriptRoot 'audit'
                if (-not (Test-Path $auditDir)) { New-Item -ItemType Directory -Path $auditDir -Force | Out-Null }
                $auditObj = [pscustomobject]@{
                    DatumTid        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    Användare       = $env:USERNAME
                    LSP             = $lsp
                    ValdCSV         = if ($selCsv) { Split-Path $selCsv -Leaf } else { '' }
                    ValdSealNEG     = Split-Path $selNeg -Leaf
                    ValdSealPOS     = Split-Path $selPos -Leaf
                    SignaturSkriven = if ($chkWriteSign.Checked) { 'Ja' } else { 'Nej' }
                    OverwroteSign   = if ($chkOverwriteSign.Checked) { 'Ja' } else { 'Nej' }
                    SigMismatch     = if ($sigMismatch) { 'Ja' } else { 'Nej' }
                    MismatchSheets  = if ($mismatchSheets -and $mismatchSheets.Count -gt 0) { ($mismatchSheets -join ';') } else { '' }
                    ViolationsNEG   = $violationsNeg.Count
                    ViolationsPOS   = $violationsPos.Count
                    Violations      = ($violationsNeg.Count + $violationsPos.Count)
                    Sparläge        = if ($rbSaveInLsp.Checked) { 'LSP-mapp' } else { 'Temporärt' }
                    OutputFile      = $SavePath
                    Kommentar       = 'UNCONTROLLED rapport, ingen källfil ändrades automatiskt.'
                    ScriptVersion   = $ScriptVersion
                }

                $auditFile = Join-Path $auditDir ("$($env:USERNAME)_audit_${nowTs}.csv")
                $auditObj | Export-Csv -Path $auditFile -NoTypeInformation -Encoding UTF8
                try {
                    $statusText = 'OK'
                    if (($violationsNeg.Count + $violationsPos.Count) -gt 0 -or $sigMismatch -or ($mismatchSheets -and $mismatchSheets.Count -gt 0)) {
                        $statusText = 'Warnings'
                    }
                    $auditTests = $null
                    try { if ($csvStats) { $auditTests = $csvStats.TestCount } } catch {}
                    Add-AuditEntry -Lsp $lsp -Assay $runAssay -BatchNumber $batch -TestCount $auditTests -Status $statusText -ReportPath $SavePath
                } catch { Gui-Log "⚠️ Kunde inte skriva audit-CSV: $($_.Exception.Message)" 'Warn' }
            } catch { Gui-Log "⚠️ Kunde inte skriva revisionsfil: $($_.Exception.Message)" 'Warn' }

            try { Start-Process -FilePath "excel.exe" -ArgumentList "`"$SavePath`"" } catch {}
        }
        catch { Gui-Log "⚠️ Kunde inte spara/öppna: $($_.Exception.Message)" 'Warn' }
    } finally {
        try { if ($pkgNeg) { $pkgNeg.Dispose() } } catch {}
        try { if ($pkgPos) { $pkgPos.Dispose() } } catch {}
        try { if ($pkgOut) { $pkgOut.Dispose() } } catch {}
    }
    } finally {
        # Stop performance timer and log elapsed time
        try {
            if ($perfBuild) { Stop-PerfTimer -Perf $perfBuild }
        } catch {}
    }

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
$miToggleSign.ToolTipText = 'Visa eller dölj panelen för att lägga till signatur.'
$toolTip.SetToolTip($rbSaveInLsp, 'Spara rapporten i mappen för ditt LSP.')
$toolTip.SetToolTip($rbTempOnly, 'Skapa rapporten temporär utan att spara.')
$toolTip.SetToolTip($btnBuild, 'Skapa och öppna rapporten baserat på de valda filerna.')
$toolTip.SetToolTip($chkSharePointInfo, 'Exportera med SharePoint Info.')
$txtLSP.add_TextChanged({ Update-BatchLink })

#region Main Run / Orchestration
# =============== SLUT ===============
function Enable-DoubleBuffer {
    $pi = [Windows.Forms.Control].GetProperty('DoubleBuffered',[Reflection.BindingFlags]'NonPublic,Instance')
    foreach($c in @($content,$pLog,$grpPick,$grpSign,$grpSave)) { if ($c) { $pi.SetValue($c,$true,$null) } }
}
try { Set-Theme 'light' } catch {}
Enable-DoubleBuffer
Update-BatchLink
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)

try{ Stop-Transcript | Out-Null }catch{}
#endregion Main Run / Orchestration
