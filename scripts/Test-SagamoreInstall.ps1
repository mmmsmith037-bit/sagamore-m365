#Requires -Version 7.4
<#
.SYNOPSIS
    Read-only audit: compares what is actually in SharePoint/Teams with the config.

.DESCRIPTION
    Makes no changes. Checks groups, site, library permissions, every planned folder's permissions,
    any extra locked folders (Invoices) found under Clients/, and Teams channels (client channels must be
    private and contain exactly the invited people).
    Writes audit-<timestamp>.csv to the logs folder. Exit code 0 = everything matches, 1 = drift found.
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..' 'config' 'sagamore.config.json'),
    [string]$LogDirectory = (Join-Path $PSScriptRoot '..' 'logs')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Common.psm1') -Force
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
Initialize-SagamoreLog -Path (Join-Path $LogDirectory "audit-$stamp.log")

$cfg = Read-SagamoreConfig -Path $ConfigPath
$problems = @(Test-SagamoreConfig -Config $cfg)
if ($problems.Count) { $problems | ForEach-Object { Write-SagamoreLog $_ -Level ERROR }; exit 2 }

Assert-SagamoreModules
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Tenant.psm1') -Force

$siteUrl = Get-SagamoreSiteUrl -Config $cfg
$libUrl  = $cfg.site.library
$rows    = [System.Collections.Generic.List[object]]::new()
function Add-Row([string]$Area, [string]$Target, [string]$Status, [string]$Detail = '') {
    $rows.Add([pscustomobject]@{ Area = $Area; Target = $Target; Status = $Status; Detail = $Detail })
    $lvl = @{ PASS = 'SUCCESS'; FAIL = 'ERROR'; WARN = 'WARN'; INFO = 'INFO' }[$Status]
    Write-SagamoreLog ("{0,-5} {1,-10} {2} {3}" -f $Status, $Area, $Target, $Detail) -Level $lvl
}

try {
    Connect-SagamoreTenant -Config $cfg -Url $siteUrl

    Write-SagamoreSection 'Groups'
    $groupIds = @{}
    foreach ($key in 'Owner', 'Member', 'Support') {
        $g = Find-SagamoreSecurityGroup -DisplayName $cfg.groups[$key]
        if (-not $g) { Add-Row 'Group' $cfg.groups[$key] 'FAIL' 'missing'; continue }
        $groupIds[$key] = Get-SagamoreProp $g 'id'
        $members = Get-SagamoreGroupMemberUpns -GroupId $groupIds[$key]
        $wanted  = @($cfg.people | Where-Object { $_.role -eq $key } | ForEach-Object { $_.upn.ToLower() })
        $missing = @($wanted | Where-Object { $_ -notin $members })
        $extra   = @($members | Where-Object { $_ -notin $wanted })
        if ($missing.Count -or $extra.Count) {
            Add-Row 'Group' $cfg.groups[$key] 'FAIL' "missing: $($missing -join ',') extra: $($extra -join ',')"
        } else { Add-Row 'Group' $cfg.groups[$key] 'PASS' "$($members.Count) member(s)" }
    }
    if ($groupIds.Count -lt 3) { throw 'Security groups are missing; run the installer first.' }

    $list = Get-SagamoreLibrary -LibraryUrl $libUrl
    $ownerSp = (Get-PnPGroup -AssociatedOwnerGroup).LoginName.ToLower()

    $toDesired = {
        param($access)
        $d = @{}
        foreach ($k in $access.Keys) { $d[(ConvertTo-SagamoreLogin (Resolve-SagamorePrincipal -Key $k -GroupIds $groupIds))] = $access[$k] }
        $d[$ownerSp] = 'Full Control'
        $d
    }

    Write-SagamoreSection 'Library'
    if (-not (Get-PnPProperty -ClientObject $list -Property HasUniqueRoleAssignments)) {
        Add-Row 'Library' $libUrl 'FAIL' 'inherits site permissions -- every Team member can edit'
    } else {
        $diff = Compare-SagamoreAccess -Desired (& $toDesired $cfg.libraryAccess) -Actual (Get-SagamoreRoleAssignments -Securable $list)
        if ($diff.Count) { $diff | ForEach-Object { Add-Row 'Library' $libUrl 'FAIL' "$($_.Issue) $($_.Login) expected=$($_.Expected) actual=$($_.Actual)" } }
        else { Add-Row 'Library' $libUrl 'PASS' }
    }

    Write-SagamoreSection 'Folders'
    $siteClients = @()
    try { $siteClients = @(Get-PnPFolderItem -FolderSiteRelativeUrl "$libUrl/Clients" -ItemType Folder | ForEach-Object Name) } catch { }
    $configClients = @($cfg.clients | ForEach-Object { $_.name.ToLower() })
    $extra = @($siteClients | Where-Object { $_.ToLower() -notin $configClients })
    foreach ($x in $extra) { Add-Row 'Folder' "Clients/$x" 'WARN' 'on site but not in config (template applies)' }

    $plan = Get-SagamoreFolderPlan -Config $cfg -ExtraClients $extra
    foreach ($p in $plan) {
        try {
            $folder = Get-PnPFolder -Url "$libUrl/$($p.Path)" -ErrorAction Stop
        } catch { Add-Row 'Folder' $p.Path 'FAIL' 'missing'; continue }
        $item = Get-PnPProperty -ClientObject $folder -Property ListItemAllFields
        $unique = Test-SagamoreFolderUnique -Folder $folder
        if (-not $p.Access) {
            if ($unique) { Add-Row 'Folder' $p.Path 'FAIL' 'has unique permissions but config says inherit' }
            else { Add-Row 'Folder' $p.Path 'PASS' 'inherits' }
            continue
        }
        if (-not $unique) { Add-Row 'Folder' $p.Path 'FAIL' 'inherits but config sets unique permissions'; continue }
        $diff = Compare-SagamoreAccess -Desired (& $toDesired $p.Access) -Actual (Get-SagamoreRoleAssignments -Securable $item)
        if ($diff.Count) { $diff | ForEach-Object { Add-Row 'Folder' $p.Path 'FAIL' "$($_.Issue) $($_.Login) expected=$($_.Expected) actual=$($_.Actual)" } }
        else { Add-Row 'Folder' $p.Path 'PASS' }
    }

    Write-SagamoreSection 'Locked folders (any depth)'
    $planned = @($plan | ForEach-Object { $_.Path.ToLower() })
    $rootRel = (Get-PnPProperty -ClientObject $list -Property RootFolder).ServerRelativeUrl.TrimEnd('/')
    $all = @()
    try { $all = @(Get-PnPFolderItem -FolderSiteRelativeUrl "$libUrl/Clients" -ItemType Folder -Recursive) } catch { }
    foreach ($f in $all) {
        $locked = Get-SagamoreLockedAccess -Config $cfg -FolderName $f.Name
        if (-not $locked) { continue }
        $rel = $f.ServerRelativeUrl.Substring($rootRel.Length).TrimStart('/')
        if ($rel.ToLower() -in $planned) { continue }
        $folder = Get-PnPFolder -Url "$libUrl/$rel"
        if (-not (Test-SagamoreFolderUnique -Folder $folder)) { Add-Row 'Locked' $rel 'FAIL' 'NOT locked -- inherits parent access'; continue }
        $diff = Compare-SagamoreAccess -Desired (& $toDesired $locked) -Actual (Get-SagamoreRoleAssignments -Securable (Get-PnPProperty -ClientObject $folder -Property ListItemAllFields))
        if ($diff.Count) { Add-Row 'Locked' $rel 'FAIL' (($diff | ForEach-Object { "$($_.Issue) $($_.Login)" }) -join '; ') }
        else { Add-Row 'Locked' $rel 'PASS' }
    }

    $teamsCfg = Get-OrDefault $cfg 'teams' @{}
    if (Get-OrDefault $teamsCfg 'enabled' $true) {
        Write-SagamoreSection 'Teams'
        $m365 = Find-SagamoreM365Group -Alias $cfg.site.alias
        if (-not (@(Get-SagamoreProp $m365 'resourceProvisioningOptions') -contains 'Team')) {
            Add-Row 'Teams' $cfg.site.title 'FAIL' 'group is not a Team'
        } else {
            $gid = Get-SagamoreProp $m365 'id'
            $have = @{}
            foreach ($ch in @(Get-PnPTeamsChannel -Team $gid)) { $have[$ch.DisplayName.ToLower()] = $ch }
            $me = Get-SagamoreProp (Invoke-PnPGraphMethod -Url 'me?$select=userPrincipalName') 'userPrincipalName'
            $keepIds = @(@(Get-OrDefault $cfg.site 'adminUpns' @()) + @($me) | Where-Object { $_ } | ForEach-Object { Get-SagamoreUserId -Upn $_ } | Where-Object { $_ } | ForEach-Object { $_.ToLower() })
            foreach ($c in Get-SagamoreChannelPlan -Config $cfg -ExtraClients $extra) {
                $ch = $have[$c.Channel.ToLower()]
                if (-not $ch) { Add-Row 'Channel' $c.Channel 'FAIL' 'missing'; continue }
                if ($c.Type -ne 'Private') { Add-Row 'Channel' $c.Channel 'PASS' "$($ch.MembershipType)"; continue }
                if ("$($ch.MembershipType)" -ine 'private') { Add-Row 'Channel' $c.Channel 'FAIL' "SILO BREACH: $($ch.MembershipType) channel, visible to the whole Team"; continue }
                $desired = @{}; $upnById = @{}
                foreach ($u in @($c.Owners) + @($c.Members)) {
                    $id = Get-SagamoreUserId -Upn $u
                    if ($id) { $upnById[$id.ToLower()] = $u; $desired[$id.ToLower()] = if ($u -in $c.Owners) { 'owner' } else { 'member' } }
                    else { Add-Row 'Channel' $c.Channel 'FAIL' "$u is invited but does not exist in Entra ID" }
                }
                $actual = Get-SagamoreChannelMembers -TeamId $gid -Channel $ch.Id
                $diff = Compare-SagamoreChannelMembers -Desired $desired -Actual $actual -KeepIds $keepIds
                if ($diff.Count) {
                    $txt = ($diff | ForEach-Object { $who = if ($upnById[$_.UserId]) { $upnById[$_.UserId] } else { $actual[$_.UserId].Name }; "$($_.Issue) $who" }) -join '; '
                    Add-Row 'Channel' $c.Channel 'FAIL' $txt
                } else {
                    Add-Row 'Channel' $c.Channel 'PASS' "private, $($desired.Count) invited"
                }
            }
        }
    }
}
catch {
    Add-Row 'Audit' 'run' 'FAIL' $_.Exception.Message
}

$csv = Join-Path $LogDirectory "audit-$stamp.csv"
$rows | Export-Csv -Path $csv -NoTypeInformation -Encoding utf8
$fails = @($rows | Where-Object Status -eq 'FAIL').Count
Write-SagamoreSection 'Audit result'
Write-SagamoreLog "Report: $csv"
if ($fails) { Write-SagamoreLog "$fails check(s) failed. Re-run ./scripts/Install-Sagamore.ps1 -PermissionsOnly to fix permission drift." -Level ERROR; exit 1 }
Write-SagamoreLog 'Everything matches the config.' -Level SUCCESS
exit 0
