#Requires -Version 7.4
<#
.SYNOPSIS
    Removes what the installer created. Use for test tenants or a failed first install.

.DESCRIPTION
    Deletes the Microsoft 365 group, which removes the Team and the SharePoint site with ALL FILES.
    Microsoft keeps a deleted group (and its site) restorable for 30 days
    (Entra admin center > Groups > Deleted groups).
    Security groups are only removed with -RemoveSecurityGroups.

    You must type the site alias to confirm.

.EXAMPLE
    ./scripts/Uninstall-Sagamore.ps1 -ConfirmAlias SagamoreAlliance-Test -RemoveSecurityGroups
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][string]$ConfirmAlias,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..' 'config' 'sagamore.config.json'),
    [switch]$RemoveSecurityGroups
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Common.psm1') -Force
$cfg = Read-SagamoreConfig -Path $ConfigPath
if ($ConfirmAlias -cne $cfg.site.alias) { throw "ConfirmAlias '$ConfirmAlias' does not match site.alias '$($cfg.site.alias)' in config. Nothing was removed." }

Assert-SagamoreModules
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Tenant.psm1') -Force
Connect-SagamoreTenant -Config $cfg -Url "https://$($cfg.tenant.name).sharepoint.com"

$m365 = Find-SagamoreM365Group -Alias $cfg.site.alias
if ($m365) {
    $id = Get-SagamoreProp $m365 'id'
    if ($PSCmdlet.ShouldProcess("$($cfg.site.title) ($id)", 'DELETE Microsoft 365 group, Team and SharePoint site including all files')) {
        Remove-PnPMicrosoft365Group -Identity $id
        Write-Host "Deleted group $($cfg.site.alias). Restorable for 30 days from Entra ID > Deleted groups." -ForegroundColor Yellow
    }
} else { Write-Host "No Microsoft 365 group with alias $($cfg.site.alias)." }

if ($RemoveSecurityGroups) {
    foreach ($key in 'Owner', 'Member', 'Support') {
        $g = Find-SagamoreSecurityGroup -DisplayName $cfg.groups[$key]
        if ($g -and $PSCmdlet.ShouldProcess($cfg.groups[$key], 'Delete security group')) {
            Remove-PnPEntraIDGroup -Identity (Get-SagamoreProp $g 'id')
            Write-Host "Deleted security group $($cfg.groups[$key])" -ForegroundColor Yellow
        }
    }
}
