#Requires -Version 7.4
<#
.SYNOPSIS
    Runs the offline test suite (no tenant needed). Installs Pester 5+, PSScriptAnalyzer and
    PnP.PowerShell for the current user if they are missing.
#>
param([switch]$CI)
$ErrorActionPreference = 'Stop'
foreach ($m in @(@{ N = 'Pester'; V = '5.5.0' }, @{ N = 'PSScriptAnalyzer'; V = '1.21.0' }, @{ N = 'PnP.PowerShell'; V = '3.0.0' })) {
    if (-not (Get-Module -ListAvailable $m.N | Where-Object Version -ge ([version]$m.V))) {
        Write-Host "Installing $($m.N)..." -ForegroundColor Cyan
        Install-Module $m.N -MinimumVersion $m.V -Scope CurrentUser -Force -SkipPublisherCheck
    }
}
Import-Module Pester -MinimumVersion 5.5.0 -Force
$c = New-PesterConfiguration
$c.Run.Path = $PSScriptRoot
$c.Run.Exit = $true
$c.Output.Verbosity = 'Detailed'
if ($CI) { $c.TestResult.Enabled = $true; $c.TestResult.OutputPath = (Join-Path $PSScriptRoot '..' 'testResults.xml') }
Invoke-Pester -Configuration $c
