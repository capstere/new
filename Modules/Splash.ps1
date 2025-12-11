<#
    Module: Splash.ps1
    Purpose: Minimal splash screen helpers for the Validate-Assay WinForms UI.
    Platform: PowerShell 5.1, WinForms
    Notes:
      - Pure UI scaffolding; no data loading or business logic is triggered here.
#>

$global:Splash = $null

<#
    .SYNOPSIS
        Shows a borderless splash screen with a status message.

    .DESCRIPTION
        Creates a top-most WinForms splash window used during startup (module load,
        PnP connect). Keeps logic lightweight to avoid masking slow operations.

    .PARAMETER msg
        Optional message to display while the app initializes.

    .OUTPUTS
        None
    
    .NOTES
        UI-only; heavy work must remain outside to keep the form responsive.
#>
function Show-Splash([string]$msg = "Startar…") {

    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    $f = New-Object Windows.Forms.Form
    $f.FormBorderStyle = 'None'
    $f.StartPosition   = 'CenterScreen'
    $f.BackColor       = [Drawing.Color]::FromArgb(0,120,215)
    $f.ForeColor       = [Drawing.Color]::White
    $f.Size            = New-Object Drawing.Size(420,120)

    $lbl = New-Object Windows.Forms.Label
    $lbl.Dock = 'Fill'
    $lbl.TextAlign = 'MiddleCenter'
    $lbl.Font = New-Object Drawing.Font('Segoe UI Semibold',12)
    $f.Controls.Add($lbl)
    $f.TopMost = $true
    $f.Show()

    $global:Splash = @{ Form = $f; Label = $lbl }
    Update-Splash $msg
    [Windows.Forms.Application]::DoEvents()
}

<#
    .SYNOPSIS
        Updates the splash message while keeping the UI responsive.

    .PARAMETER msg
        Text to show on the splash label.

    .OUTPUTS
        None
    
    .NOTES
        Keeps DoEvents to flush the message pump during startup.
#>
function Update-Splash([string]$msg) {
    if ($global:Splash) {
        $global:Splash.Label.Text = $msg
        [Windows.Forms.Application]::DoEvents()
    }
}

<#
    .SYNOPSIS
        Closes and disposes the splash form.

    .OUTPUTS
        None
    
    .NOTES
        Safe to call multiple times; guard with the global Splash hashtable.
#>
function Close-Splash() {
    if ($global:Splash) {
        $global:Splash.Form.Close()
        $global:Splash.Form.Dispose()
        $global:Splash = $null
    }
}
