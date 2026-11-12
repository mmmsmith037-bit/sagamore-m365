#Requires -Version 7.4
<#
.SYNOPSIS
    One-time: creates the Entra ID app registration the installer signs in with, and writes its
    Client ID into config/sagamore.config.json.

.DESCRIPTION
    PnP.PowerShell no longer ships a shared multi-tenant app, so every tenant needs its own.
    Must be run by a Global Administrator (or Application Administrator + Privileged Role
    Administrator) of the Sagamore tenant, because it grants admin consent.

    Delegated permissions requested (the app can only ever do what the signed-in user can do):
      SharePoint: AllSites.FullControl
      Graph:      Group.ReadWrite.All, User.Read.All, Directory.Read.All,
                  Team.Create, TeamSettings.ReadWrite.All, TeamMember.ReadWrite.All,
                  Channel.Create, Channel.ReadBasic.All, TeamsTab.ReadWrite.All

.EXAMPLE
    ./scripts/Register-SagamoreApp.ps1
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..' 'config' 'sagamore.config.json'),
    [string]$ApplicationName = 'Sagamore M365 Installer',
    [switch]$DeviceLogin
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Common.psm1') -Force
Assert-SagamoreModules

$cfg = Read-SagamoreConfig -Path $ConfigPath
$tenant = "$($cfg.tenant.name).onmicrosoft.com"

$graph = @(
    'Group.ReadWrite.All', 'User.Read.All', 'Directory.Read.All',
    'Team.Create', 'TeamSettings.ReadWrite.All', 'TeamMember.ReadWrite.All',
    'Channel.Create', 'Channel.ReadBasic.All', 'TeamsTab.ReadWrite.All'
)

Write-Host "Registering '$ApplicationName' in $tenant ..." -ForegroundColor Cyan
Write-Host 'A browser window will open. Sign in as a Global Administrator and accept the consent prompt.' -ForegroundColor Cyan

if (-not $PSCmdlet.ShouldProcess($tenant, "Register Entra ID app '$ApplicationName'")) { return }

$params = @{
    ApplicationName               = $ApplicationName
    Tenant                        = $tenant
    GraphDelegatePermissions      = $graph
    SharePointDelegatePermissions = @('AllSites.FullControl')
}
if ($DeviceLogin) { $params.DeviceLogin = $true }

$result = Register-PnPEntraIDAppForInteractiveLogin @params

$clientId = $null
foreach ($name in 'AzureAppId/ClientId', 'AzureAppId', 'ClientId', 'AppId') {
    if ($result -and $result.PSObject.Properties[$name]) { $clientId = $result.$name; break }
    if ($result -is [System.Collections.IDictionary] -and $result.Contains($name)) { $clientId = $result[$name]; break }
}
if (-not $clientId) {
    Write-Host ($result | Out-String)
    throw 'App was registered but its Client ID could not be read from the output above. Copy it into tenant.clientId in the config manually.'
}

$cfg.tenant.clientId = [string]$clientId
Save-SagamoreConfig -Config $cfg -Path $ConfigPath
Write-Host "Client ID $clientId saved to $ConfigPath" -ForegroundColor Green
Write-Host 'Next: ./scripts/Install-Sagamore.ps1 -WhatIf' -ForegroundColor Green
