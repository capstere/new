#requires -Version 5.1
<#!
  RuleEngine.ps1 (facade)
  - Keep backward compatibility: Main.ps1 dot-sources this file.
  - This file only dot-sources the actual implementation files.
  - EPPlus 4.5.3.3 safe. PowerShell 5.1.
!#>

Set-StrictMode -Off

. (Join-Path $PSScriptRoot 'Rules\RuleEngine.Core.ps1')
. (Join-Path $PSScriptRoot 'Writers\Information2Writer.ps1')
