﻿#requires -Version 5.1
# Kör: Invoke-Pester -Path .\Tests
Describe "RuleEngine smoke tests" {
    It "Can dot-source RuleEngine without errors" {
        { . "$PSScriptRoot\..\Modules\RuleEngine.ps1" } | Should Not Throw
    }
}
