#Requires -Version 7.4
<#
.SYNOPSIS
    Installs (or re-applies) the Sagamore Alliance SharePoint + Teams environment from config.

.DESCRIPTION
    Idempotent: safe to run again at any time. Each run makes the tenant match the config:
      1. Entra ID security groups (Sagamore-Owner / -Member / -Support) and their members
      2. The Microsoft 365 group-connected team site
      3. Site settings: external sharing, versioning, "Add Only" permission level
      4. Library-level permissions (replaces the default "everyone in the Team can edit")
      5. Folder tree + per-folder permissions, including Invoices lock-down at any depth
      6. Microsoft Teams: team, owners/members, one channel per client + Files tab
      7. Permission report (CSV) written locally and to _Admin/Permissions-Log

.PARAMETER ConfigPath
    Path to sagamore.config.json (default: ../config/sagamore.config.json).

.PARAMETER PlanOnly
    Do not connect to Microsoft 365. Validate config and print the folder/permission/channel plan.

.PARAMETER PermissionsOnly
    Skip group/site/Teams creation. Re-stamp library + folder permissions and channels only.
    Use after importing clients or editing permissions in the dashboard.

.PARAMETER SkipTeams
    Do not create or update the Team, members, channels or tabs.

.EXAMPLE
    ./scripts/Install-Sagamore.ps1 -PlanOnly
.EXAMPLE
    ./scripts/Install-Sagamore.ps1 -WhatIf
.EXAMPLE
    ./scripts/Install-Sagamore.ps1
.EXAMPLE
    ./scripts/Install-Sagamore.ps1 -PermissionsOnly
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..' 'config' 'sagamore.config.json'),
    [switch]$PlanOnly,
    [switch]$PermissionsOnly,
    [switch]$SkipTeams,
    [string]$LogDirectory = (Join-Path $PSScriptRoot '..' 'logs')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Common.psm1') -Force
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
Initialize-SagamoreLog -Path (Join-Path $LogDirectory "install-$stamp.log")

$script:Config   = Read-SagamoreConfig -Path $ConfigPath
$script:Problems = @(Test-SagamoreConfig -Config $script:Config -AllowPlaceholderClientId:$PlanOnly)
if ($script:Problems.Count) {
    Write-SagamoreLog "Config has $($script:Problems.Count) problem(s):" -Level ERROR
    $script:Problems | ForEach-Object { Write-SagamoreLog "  - $_" -Level ERROR }
    exit 2
}

$cfg      = $script:Config
$siteUrl  = Get-SagamoreSiteUrl -Config $cfg
$rootUrl  = "https://$($cfg.tenant.name).sharepoint.com"
$libUrl   = $cfg.site.library
$teamsCfg = Get-OrDefault $cfg 'teams' @{}
$doTeams  = (Get-OrDefault $teamsCfg 'enabled' $true) -and -not $SkipTeams
$warnings = [System.Collections.Generic.List[string]]::new()
$adminUpns = @(Get-OrDefault $cfg.site 'adminUpns' @() | ForEach-Object { $_.ToLower() })

function Add-Warning([string]$Message) { $warnings.Add($Message); Write-SagamoreLog $Message -Level WARN }

function Show-Plan {
    param([object[]]$Plan, [object[]]$Channels)
    Write-SagamoreSection 'Plan: folders and permissions'
    Write-SagamoreLog "Library root ($libUrl): $((($cfg.libraryAccess.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '))" -Level PLAN
    foreach ($p in $Plan) {
        $acc = if ($p.Access) { ($p.Access.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; ' } else { '(inherits)' }
        Write-SagamoreLog ("{0,-55} {1}" -f $p.Path, $acc) -Level PLAN
    }
    if ($doTeams) {
        Write-SagamoreSection 'Plan: Teams channels'
        foreach ($c in $Channels) { Write-SagamoreLog ("{0,-40} {1}" -f $c.Channel, ($c.Folder ?? '')) -Level PLAN }
    }
}

# ------------------------------------------------------------------
# PLAN ONLY
# ------------------------------------------------------------------
if ($PlanOnly) {
    Show-Plan -Plan (Get-SagamoreFolderPlan -Config $cfg) -Channels (Get-SagamoreChannelPlan -Config $cfg)
    Write-SagamoreLog ''
    Write-SagamoreLog 'Plan only -- nothing was changed. Config is valid.' -Level SUCCESS
    exit 0
}

Assert-SagamoreModules
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Tenant.psm1') -Force
$wi = @{ WhatIf = $WhatIfPreference }

Write-SagamoreLog ('=' * 60)
Write-SagamoreLog "  Sagamore Alliance installer"
Write-SagamoreLog "  Mode:   $(if ($PermissionsOnly) { 'PERMISSIONS ONLY' } else { 'FULL' })$(if ($WhatIfPreference) { ' (WHATIF - no changes)' })"
Write-SagamoreLog "  Site:   $siteUrl"
Write-SagamoreLog "  Config: $((Resolve-Path $ConfigPath).Path)"
Write-SagamoreLog ('=' * 60)

try {
    # --------------------------------------------------------------
    # 1. SIGN IN
    # --------------------------------------------------------------
    Write-SagamoreSection 'Signing in'
    Connect-SagamoreTenant -Config $cfg -Url $rootUrl
    $me = (Get-SagamoreProp (Invoke-PnPGraphMethod -Url 'me?$select=userPrincipalName') 'userPrincipalName').ToLower()
    $allOwners = @($cfg.people | Where-Object { $_.role -eq 'Owner' } | ForEach-Object { $_.upn.ToLower() }) + $adminUpns
    Write-SagamoreLog "Signed in as: $me" -Level INFO
    if ($me -notin $allOwners) {
        throw "You are signed in as $me, who is not an Owner in config and not in site.adminUpns. Add $me to site.adminUpns (so the installer keeps access to every folder) and re-run."
    }

    # --------------------------------------------------------------
    # 2. SECURITY GROUPS (needed even in PermissionsOnly to resolve ids)
    # --------------------------------------------------------------
    Write-SagamoreSection 'Entra ID security groups'
    $groupIds = @{}
    foreach ($key in 'Owner', 'Member', 'Support') {
        $name  = $cfg.groups[$key]
        $group = Find-SagamoreSecurityGroup -DisplayName $name
        if ($group) {
            $groupIds[$key] = Get-SagamoreProp $group 'id'
            Write-SagamoreLog "Group exists: $name" -Level SKIP
        } elseif ($PermissionsOnly) {
            throw "Group '$name' does not exist. Run a full install first."
        } elseif ($PSCmdlet.ShouldProcess($name, 'Create Entra ID security group')) {
            $nick    = ($name -replace '[^A-Za-z0-9-]', '').ToLower()
            $created = New-PnPEntraIDGroup -DisplayName $name -Description "Sagamore Alliance $key -- managed by sagamore-m365" -MailNickname $nick -IsSecurityEnabled
            $groupIds[$key] = Get-SagamoreProp $created 'Id'
            if (-not $groupIds[$key]) { $groupIds[$key] = Get-SagamoreProp (Invoke-SagamoreRetry -What "find $name" -Action { Find-SagamoreSecurityGroup -DisplayName $name } ) 'id' }
            Write-SagamoreLog "Created group: $name" -Level SUCCESS
        } else {
            $groupIds[$key] = [guid]::NewGuid().ToString()   # placeholder so -WhatIf can continue
        }

        if (-not $PermissionsOnly) {
            $wanted  = @($cfg.people | Where-Object { $_.role -eq $key } | ForEach-Object { $_.upn.ToLower() })
            $current = if ($group) { Get-SagamoreGroupMemberUpns -GroupId $groupIds[$key] } else { @() }
            foreach ($upn in $wanted | Where-Object { $_ -notin $current }) {
                if (-not (Test-SagamoreUserExists -Upn $upn)) { Add-Warning "User not found in tenant, skipped: $upn"; continue }
                if ($PSCmdlet.ShouldProcess($upn, "Add to $name")) {
                    Add-PnPEntraIDGroupMember -Identity $groupIds[$key] -Users $upn
                    Write-SagamoreLog "Added $upn to $name" -Level SUCCESS
                }
            }
            foreach ($upn in $current | Where-Object { $_ -notin $wanted }) {
                Add-Warning "$upn is in $name but not in config with role $key. Remove them in Entra ID or add them to config."
            }
        }
    }

    # --------------------------------------------------------------
    # 3. SITE
    # --------------------------------------------------------------
    $m365 = Find-SagamoreM365Group -Alias $cfg.site.alias
    if (-not $PermissionsOnly) {
        Write-SagamoreSection 'Team site'
        if ($m365) {
            Write-SagamoreLog "Microsoft 365 group/site exists: $($cfg.site.alias)" -Level SKIP
        } elseif ($PSCmdlet.ShouldProcess($siteUrl, 'Create Microsoft 365 group-connected team site')) {
            $owners  = @($cfg.people | Where-Object { $_.role -eq 'Owner' } | ForEach-Object { $_.upn }) + $adminUpns
            $newUrl  = New-PnPSite -Type TeamSite -Title $cfg.site.title -Alias $cfg.site.alias -Owners $owners -Wait
            Write-SagamoreLog "Site created: $newUrl" -Level SUCCESS
            if ($newUrl -and $newUrl.TrimEnd('/') -ne $siteUrl) {
                throw "SharePoint created the site at $newUrl instead of $siteUrl (alias taken?). Set site.alias in config to match, then re-run."
            }
            $m365 = Invoke-SagamoreRetry -What 'find new Microsoft 365 group' -Action {
                $g = Find-SagamoreM365Group -Alias $cfg.site.alias; if (-not $g) { throw 'not visible yet' }; $g
            }
        } else {
            Write-SagamoreLog 'WhatIf: site does not exist yet, so folder/permission steps cannot be previewed against it.' -Level PLAN
            Show-Plan -Plan (Get-SagamoreFolderPlan -Config $cfg) -Channels (Get-SagamoreChannelPlan -Config $cfg)
            exit 0
        }
    } elseif (-not $m365) {
        throw "Site $siteUrl does not exist. Run a full install first."
    }
    $groupId = Get-SagamoreProp $m365 'id'

    Connect-SagamoreTenant -Config $cfg -Url $siteUrl
    $list = Get-SagamoreLibrary -LibraryUrl $libUrl
    # The site's own Owners group (= Microsoft 365 group owners: the Owner people + site.adminUpns)
    # keeps Full Control everywhere so the installer and Team owners are never locked out.
    $ownerSpGroup = Get-PnPGroup -AssociatedOwnerGroup

    # --------------------------------------------------------------
    # 4. SITE SETTINGS
    # --------------------------------------------------------------
    if (-not $PermissionsOnly) {
        Write-SagamoreSection 'Site settings'
        $sharing = Get-OrDefault $cfg.site 'externalSharing' 'ExistingExternalUserSharingOnly'
        if ($PSCmdlet.ShouldProcess($siteUrl, "Set external sharing to $sharing")) {
            try {
                Set-PnPSite -SharingCapability $sharing -DisableSharingForNonOwners
                Write-SagamoreLog "External sharing: $sharing (only owners can share)" -Level SUCCESS
            } catch { Add-Warning "Could not set sharing (needs SharePoint admin role): $($_.Exception.Message)" }
        }
        $versions = [int](Get-OrDefault $cfg.site 'versionLimit' 50)
        if ($PSCmdlet.ShouldProcess($libUrl, "Enable versioning ($versions major versions)")) {
            Set-PnPList -Identity $list -EnableVersioning $true -MajorVersions $versions | Out-Null
            Write-SagamoreLog "Versioning: $versions major versions" -Level SUCCESS
        }
    }

    $addOnly = Get-SagamoreAddOnlyRole
    $existingRole = $null
    try { $existingRole = Get-PnPRoleDefinition -Identity $addOnly.Name -ErrorAction Stop } catch { }
    if ($existingRole) {
        Write-SagamoreLog "Permission level exists: $($addOnly.Name)" -Level SKIP
    } elseif ($PSCmdlet.ShouldProcess($siteUrl, "Create permission level '$($addOnly.Name)'")) {
        Add-PnPRoleDefinition -RoleName $addOnly.Name -Clone $addOnly.Clone -Exclude $addOnly.Exclude -Description $addOnly.Description | Out-Null
        Write-SagamoreLog "Created permission level: $($addOnly.Name) (Contribute without delete)" -Level SUCCESS
    }

    # --------------------------------------------------------------
    # 5. LIBRARY PERMISSIONS
    # --------------------------------------------------------------
    Write-SagamoreSection 'Library permissions'
    $keep = [System.Collections.Generic.List[string]]::new()
    $keep.Add($ownerSpGroup.LoginName.ToLower())

    $hasUnique = Get-PnPProperty -ClientObject $list -Property HasUniqueRoleAssignments
    if (-not $hasUnique -and $PSCmdlet.ShouldProcess($libUrl, 'Break permission inheritance')) {
        # Copy first so nobody (incl. the installer) loses access mid-run; extras are trimmed below.
        # -ClearSubscopes is deliberately NOT used: folder-level permissions are preserved.
        Set-PnPList -Identity $list -BreakRoleInheritance -CopyRoleAssignments | Out-Null
        Write-SagamoreLog 'Library now has unique permissions' -Level SUCCESS
    }
    if ($PSCmdlet.ShouldProcess($libUrl, 'Grant site owners Full Control')) {
        Set-PnPListPermission -Identity $list -Group $ownerSpGroup -AddRole 'Full Control'
    }
    foreach ($k in $cfg.libraryAccess.Keys) {
        $principal = Resolve-SagamorePrincipal -Key $k -GroupIds $groupIds
        $keep.Add((ConvertTo-SagamoreLogin $principal))
        if ($PSCmdlet.ShouldProcess($libUrl, "Grant $k $($cfg.libraryAccess[$k])")) {
            Set-PnPListPermission -Identity $list -User $principal -AddRole $cfg.libraryAccess[$k]
            Write-SagamoreLog "Library: $k -> $($cfg.libraryAccess[$k])" -Level SUCCESS
        }
    }
    if (-not $WhatIfPreference) {
        Remove-SagamoreUnexpectedAssignments -Securable $list -KeepLogins $keep.ToArray() -Label "library $libUrl" @wi
    }

    # --------------------------------------------------------------
    # 6. FOLDERS + FOLDER PERMISSIONS
    # --------------------------------------------------------------
    Write-SagamoreSection 'Folders and permissions'
    $siteClients = @()
    try {
        $siteClients = @(Get-PnPFolderItem -FolderSiteRelativeUrl "$libUrl/Clients" -ItemType Folder -ErrorAction Stop |
                         Where-Object { $_.Name -ne 'Forms' } | ForEach-Object Name)
    } catch { }
    $configClients = @($cfg.clients | ForEach-Object { $_.name.ToLower() })
    $extra = @($siteClients | Where-Object { $_.ToLower() -notin $configClients })
    foreach ($x in $extra) { Add-Warning "Client folder '$x' exists on the site but not in config -- applying client template. Add it to config to customise." }

    $plan = Get-SagamoreFolderPlan -Config $cfg -ExtraClients $extra
    $report = [System.Collections.Generic.List[object]]::new()

    foreach ($p in $plan) {
        $sitePath = "$libUrl/$($p.Path)"
        $folder = $null
        if ($PSCmdlet.ShouldProcess($p.Path, 'Ensure folder exists')) {
            $folder = Resolve-PnPFolder -SiteRelativePath $sitePath
        }

        if ($p.Access) {
            if ($PSCmdlet.ShouldProcess($p.Path, 'Replace permissions (site owners keep Full Control)')) {
                Set-PnPFolderPermission -List $list -Identity $sitePath -Group $ownerSpGroup -AddRole 'Full Control' -ClearExisting
            }
            foreach ($k in $p.Access.Keys) {
                $principal = Resolve-SagamorePrincipal -Key $k -GroupIds $groupIds
                $role = $p.Access[$k]
                if ($PSCmdlet.ShouldProcess($p.Path, "Grant $k $role")) {
                    Set-PnPFolderPermission -List $list -Identity $sitePath -User $principal -AddRole $role
                }
                $report.Add([pscustomobject]@{ Folder = $p.Path; Principal = $k; Role = $role; Mode = 'Unique' })
            }
            Write-SagamoreLog "$($p.Path): $((($p.Access.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '))" -Level SUCCESS
        } else {
            $isUnique = $folder -and (Test-SagamoreFolderUnique -Folder $folder)
            if ($isUnique -and $PSCmdlet.ShouldProcess($p.Path, 'Reset to inherit permissions (config says inherit)')) {
                Set-PnPFolderPermission -List $list -Identity $sitePath -InheritPermissions
                Write-SagamoreLog "$($p.Path): reset to inherit" -Level SUCCESS
            } else {
                Write-SagamoreLog "$($p.Path): inherits" -Level SKIP
            }
            $report.Add([pscustomobject]@{ Folder = $p.Path; Principal = '(inherits)'; Role = ''; Mode = 'Inherit' })
        }
    }

    # Locked folders (Invoices) at any depth that the config doesn't mention, e.g. created by staff or by import
    Write-SagamoreSection 'Locked folder sweep'
    $planned = @($plan | ForEach-Object { $_.Path.ToLower() })
    $rootRel = (Get-PnPProperty -ClientObject $list -Property RootFolder).ServerRelativeUrl.TrimEnd('/')
    try {
        $all = @(Get-PnPFolderItem -FolderSiteRelativeUrl "$libUrl/Clients" -ItemType Folder -Recursive -ErrorAction Stop)
    } catch { $all = @() }
    $swept = 0
    foreach ($f in $all) {
        $locked = Get-SagamoreLockedAccess -Config $cfg -FolderName $f.Name
        if (-not $locked) { continue }
        $rel = $f.ServerRelativeUrl.Substring($rootRel.Length).TrimStart('/')
        if ($rel.ToLower() -in $planned) { continue }
        if ($PSCmdlet.ShouldProcess($rel, 'Lock: replace permissions (site owners keep Full Control)')) {
            Set-PnPFolderPermission -List $list -Identity "$libUrl/$rel" -Group $ownerSpGroup -AddRole 'Full Control' -ClearExisting
        }
        foreach ($k in $locked.Keys) {
            $principal = Resolve-SagamorePrincipal -Key $k -GroupIds $groupIds
            if ($PSCmdlet.ShouldProcess($rel, "Lock: grant $k $($locked[$k])")) {
                Set-PnPFolderPermission -List $list -Identity "$libUrl/$rel" -User $principal -AddRole $locked[$k]
            }
            $report.Add([pscustomobject]@{ Folder = $rel; Principal = $k; Role = $locked[$k]; Mode = 'Locked' })
        }
        $swept++
        Write-SagamoreLog "Locked: $rel" -Level SUCCESS
    }
    Write-SagamoreLog "Locked folder sweep complete ($swept extra folder(s))" -Level INFO

    # --------------------------------------------------------------
    # 7. TEAMS
    # --------------------------------------------------------------
    if ($doTeams) {
        Write-SagamoreSection 'Microsoft Teams'
        $isTeam = @(Get-SagamoreProp $m365 'resourceProvisioningOptions') -contains 'Team'
        if (-not $isTeam -and -not $PermissionsOnly) {
            if ($PSCmdlet.ShouldProcess($cfg.site.title, 'Enable Microsoft Teams on the group')) {
                Invoke-SagamoreRetry -What 'Teamify group' -Attempts 8 -DelaySeconds 30 -Action { New-PnPTeamsTeam -GroupId $groupId | Out-Null }
                Write-SagamoreLog 'Team created' -Level SUCCESS
                $isTeam = $true
            }
        } elseif ($isTeam) {
            Write-SagamoreLog 'Team exists' -Level SKIP
        }

        if (-not $PermissionsOnly) {
            $owners  = @($cfg.people | Where-Object { $_.role -eq 'Owner' } | ForEach-Object { $_.upn.ToLower() }) + $adminUpns
            $roles   = @('Member') + $(if (Get-OrDefault $teamsCfg 'supportAsTeamMembers' $true) { 'Support' } else { @() })
            $members = @($cfg.people | Where-Object { $_.role -in $roles } | ForEach-Object { $_.upn.ToLower() })
            $curOwners  = Get-SagamoreGroupOwnerUpns  -GroupId $groupId
            $curMembers = Get-SagamoreGroupMemberUpns -GroupId $groupId
            foreach ($u in $owners | Where-Object { $_ -notin $curOwners }) {
                if ($PSCmdlet.ShouldProcess($u, 'Add as Team owner')) { Add-PnPMicrosoft365GroupOwner -Identity $groupId -Users $u; Write-SagamoreLog "Team owner: $u" -Level SUCCESS }
            }
            foreach ($u in ($owners + $members) | Where-Object { $_ -notin $curMembers }) {
                if (-not (Test-SagamoreUserExists -Upn $u)) { continue }
                if ($PSCmdlet.ShouldProcess($u, 'Add as Team member')) { Add-PnPMicrosoft365GroupMember -Identity $groupId -Users $u; Write-SagamoreLog "Team member: $u" -Level SUCCESS }
            }
            foreach ($u in $curMembers | Where-Object { $_ -notin ($owners + $members) }) {
                Add-Warning "$u is a Team member but not in config. Remove them from the Team if they should not be there."
            }
        }

        if ($isTeam) {
            $channels = Get-SagamoreChannelPlan -Config $cfg -ExtraClients $extra
            $existing = @(Get-PnPTeamsChannel -Team $groupId | ForEach-Object { $_.DisplayName.ToLower() })
            foreach ($c in $channels) {
                if ($c.Channel.ToLower() -in $existing) {
                    Write-SagamoreLog "Channel exists: $($c.Channel)" -Level SKIP
                } elseif ($PSCmdlet.ShouldProcess($c.Channel, 'Create Teams channel')) {
                    try {
                        Add-PnPTeamsChannel -Team $groupId -DisplayName $c.Channel | Out-Null
                        Write-SagamoreLog "Created channel: $($c.Channel)" -Level SUCCESS
                    } catch { Add-Warning "Could not create channel $($c.Channel): $($_.Exception.Message)"; continue }
                }
                if ($c.Folder) {
                    $tabName = 'Client Files'
                    try {
                        $tabs = @(Get-PnPTeamsTab -Team $groupId -Channel $c.Channel | ForEach-Object DisplayName)
                        if ($tabName -in $tabs) { continue }
                        $url = "$siteUrl/" + (((@($libUrl) + ($c.Folder -split '/')) | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/')
                        if ($PSCmdlet.ShouldProcess($c.Channel, "Add '$tabName' tab -> $($c.Folder)")) {
                            Add-PnPTeamsTab -Team $groupId -Channel $c.Channel -DisplayName $tabName -Type DocumentLibrary -ContentUrl $url | Out-Null
                            Write-SagamoreLog "Tab added: $($c.Channel) -> $($c.Folder)" -Level SUCCESS
                        }
                    } catch { Add-Warning "Could not add Files tab to $($c.Channel): $($_.Exception.Message)" }
                }
            }
        }
    }

    # --------------------------------------------------------------
    # 8. REPORT
    # --------------------------------------------------------------
    Write-SagamoreSection 'Permission report'
    $csv = Join-Path $LogDirectory "permissions-$stamp.csv"
    $report | Export-Csv -Path $csv -NoTypeInformation -Encoding utf8
    Write-SagamoreLog "Report: $csv" -Level SUCCESS
    if (-not $WhatIfPreference) {
        try {
            Add-PnPFile -Path $csv -Folder "$libUrl/_Admin/Permissions-Log" | Out-Null
            Write-SagamoreLog 'Report uploaded to _Admin/Permissions-Log' -Level SUCCESS
        } catch { Add-Warning "Could not upload report: $($_.Exception.Message)" }
    }

    Write-SagamoreSection 'Done'
    if ($warnings.Count) {
        Write-SagamoreLog "$($warnings.Count) warning(s):" -Level WARN
        $warnings | ForEach-Object { Write-SagamoreLog "  - $_" -Level WARN }
    }
    Write-SagamoreLog 'Next: run ./scripts/Test-SagamoreInstall.ps1 to audit the result.' -Level INFO
    exit 0
}
catch {
    Write-SagamoreLog "FATAL: $($_.Exception.Message)" -Level ERROR
    Write-SagamoreLog $_.ScriptStackTrace -Level ERROR
    exit 1
}
