#Requires -Version 7.4
<#
    Sagamore.Tenant -- helpers that talk to Microsoft 365 through PnP.PowerShell.
    Every write goes through ShouldProcess on the calling script, so -WhatIf is honoured.
#>

Set-StrictMode -Version Latest

$script:LimitedRoles = @('Limited Access', 'Web-Only Limited Access')

function Invoke-SagamoreRetry {
    <# Retries transient Graph/SharePoint failures (new groups and Teams can take minutes to propagate). #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [string]$What = 'operation',
        [int]$Attempts = 6,
        [int]$DelaySeconds = 20
    )
    for ($i = 1; $i -le $Attempts; $i++) {
        try { return & $Action }
        catch {
            if ($i -eq $Attempts) { throw }
            Write-SagamoreLog "$What failed (attempt $i/$Attempts): $($_.Exception.Message) -- retrying in ${DelaySeconds}s" -Level WARN
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

function Connect-SagamoreTenant {
    <# Interactive sign-in using the Sagamore app registration. Returns nothing; sets the current PnP connection. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config, [Parameter(Mandatory)][string]$Url)
    Connect-PnPOnline -Url $Url -ClientId $Config.tenant.clientId -Tenant "$($Config.tenant.name).onmicrosoft.com" -Interactive -ErrorAction Stop
    Write-SagamoreLog "Connected: $Url" -Level SUCCESS
}

function Get-SagamoreGraphValue {
    <# Invoke-PnPGraphMethod returns either a hashtable or a PSObject depending on version; normalise to an array. #>
    param($Response)
    if ($null -eq $Response) { return @() }
    if ($Response -is [System.Collections.IDictionary]) { return @($Response['value']) }
    if ($Response.PSObject.Properties['value']) { return @($Response.value) }
    return @($Response)
}

function Get-SagamoreProp {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function Find-SagamoreSecurityGroup {
    param([Parameter(Mandatory)][string]$DisplayName)
    $safe = $DisplayName.Replace("'", "''")
    $res = Invoke-PnPGraphMethod -Url "groups?`$filter=displayName eq '$safe'&`$select=id,displayName,securityEnabled,mailEnabled,groupTypes"
    $groups = @(Get-SagamoreGraphValue $res | Where-Object { (Get-SagamoreProp $_ 'securityEnabled') })
    if ($groups.Count -gt 1) { throw "More than one security group is named '$DisplayName'. Rename the duplicates in Entra ID first." }
    if ($groups.Count -eq 1) { return $groups[0] }
    return $null
}

function Get-SagamoreGroupMemberUpns {
    param([Parameter(Mandatory)][string]$GroupId)
    $res = Invoke-PnPGraphMethod -Url "groups/$GroupId/members?`$select=id,userPrincipalName" -All
    @(Get-SagamoreGraphValue $res | ForEach-Object { Get-SagamoreProp $_ 'userPrincipalName' } | Where-Object { $_ } | ForEach-Object { $_.ToLower() })
}

function Get-SagamoreGroupOwnerUpns {
    param([Parameter(Mandatory)][string]$GroupId)
    $res = Invoke-PnPGraphMethod -Url "groups/$GroupId/owners?`$select=id,userPrincipalName" -All
    @(Get-SagamoreGraphValue $res | ForEach-Object { Get-SagamoreProp $_ 'userPrincipalName' } | Where-Object { $_ } | ForEach-Object { $_.ToLower() })
}

function Test-SagamoreUserExists {
    param([Parameter(Mandatory)][string]$Upn)
    try {
        $null = Invoke-PnPGraphMethod -Url "users/$([uri]::EscapeDataString($Upn))?`$select=id"
        return $true
    } catch { return $false }
}

$script:UserIdCache = @{}
function Get-SagamoreUserId {
    <# Entra object id for a UPN, or $null if the user does not exist. Cached per run. #>
    param([Parameter(Mandatory)][string]$Upn)
    $k = $Upn.ToLower()
    if (-not $script:UserIdCache.ContainsKey($k)) {
        try {
            $u = Invoke-PnPGraphMethod -Url "users/$([uri]::EscapeDataString($Upn))?`$select=id"
            $script:UserIdCache[$k] = Get-SagamoreProp $u 'id'
        } catch { $script:UserIdCache[$k] = $null }
    }
    return $script:UserIdCache[$k]
}

function Get-SagamoreChannelMembers {
    <# Members of a channel: @{ <userId lower> = @{ Name; Role = 'owner'|'member'|'guest'; Member = <TeamChannelMember> } } #>
    param([Parameter(Mandatory)][string]$TeamId, [Parameter(Mandatory)]$Channel)
    $out = @{}
    foreach ($m in @(Get-PnPTeamsChannelUser -Team $TeamId -Channel $Channel)) {
        $roles = @($m.Roles | ForEach-Object { "$_".ToLower() })
        $role = if ('owner' -in $roles) { 'owner' } elseif ('guest' -in $roles) { 'guest' } else { 'member' }
        $name = if ($m.Email) { $m.Email } else { $m.DisplayName }
        if ($m.UserId) { $out[$m.UserId.ToLower()] = @{ Name = $name; Role = $role; Member = $m } }
    }
    return $out
}

function Compare-SagamoreChannelMembers {
    <#
        Pure comparison. Desired: @{ userId = 'owner'|'member' }. Actual: output of Get-SagamoreChannelMembers.
        KeepIds (installer admins) are never reported as unexpected. Returns Missing / WrongRole / Unexpected rows.
    #>
    param([Parameter(Mandatory)][hashtable]$Desired, [Parameter(Mandatory)][hashtable]$Actual, [string[]]$KeepIds = @())
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($id in $Desired.Keys) {
        if (-not $Actual.ContainsKey($id)) { $rows.Add([pscustomobject]@{ Issue = 'Missing'; UserId = $id; Expected = $Desired[$id]; Actual = '' }) }
        elseif ($Actual[$id].Role -ne $Desired[$id]) { $rows.Add([pscustomobject]@{ Issue = 'WrongRole'; UserId = $id; Expected = $Desired[$id]; Actual = $Actual[$id].Role }) }
    }
    foreach ($id in $Actual.Keys) {
        if (-not $Desired.ContainsKey($id) -and $id -notin $KeepIds) {
            $rows.Add([pscustomobject]@{ Issue = 'Unexpected'; UserId = $id; Expected = ''; Actual = $Actual[$id].Role })
        }
    }
    return , $rows.ToArray()
}

function Find-SagamoreM365Group {
    param([Parameter(Mandatory)][string]$Alias)
    $safe = $Alias.Replace("'", "''")
    $res = Invoke-PnPGraphMethod -Url "groups?`$filter=mailNickname eq '$safe'&`$select=id,displayName,mailNickname,groupTypes,resourceProvisioningOptions"
    $g = @(Get-SagamoreGraphValue $res)
    if ($g.Count -gt 0) { return $g[0] }
    return $null
}

function Get-SagamoreLibrary {
    <# Finds the document library by its URL segment (e.g. "Shared Documents"), regardless of its display title. #>
    param([Parameter(Mandatory)][string]$LibraryUrl)
    $lists = Get-PnPList -Includes RootFolder
    $match = @($lists | Where-Object { $_.BaseTemplate -eq 101 -and $_.RootFolder.ServerRelativeUrl.TrimEnd('/') -like "*/$LibraryUrl" })
    if ($match.Count -ne 1) { throw "Could not find exactly one document library at URL '$LibraryUrl' on this site." }
    return $match[0]
}

function Test-SagamoreFolderUnique {
    <# True when the folder has its own (broken-inheritance) permissions. #>
    param([Parameter(Mandatory)]$Folder)
    $item = Get-PnPProperty -ClientObject $Folder -Property ListItemAllFields
    return [bool](Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments)
}

function ConvertTo-SagamoreLogin {
    <# Normalises a principal into the SharePoint login name used for comparisons. #>
    param([Parameter(Mandatory)][string]$Principal)
    if ($Principal -like 'c:0t.c|tenant|*' -or $Principal -like 'i:0#.f|membership|*') { return $Principal.ToLower() }
    if ($Principal -match '@') { return "i:0#.f|membership|$($Principal.ToLower())" }
    return $Principal.ToLower()
}

function Get-SagamoreRoleAssignments {
    <#
        Returns @{ <login lower> = @{ Title = ..; Roles = [string[]] } } for a securable object
        (a List or a folder's ListItemAllFields). Limited-access-only principals are dropped.
    #>
    param([Parameter(Mandatory)]$Securable)
    $ras = Get-PnPProperty -ClientObject $Securable -Property RoleAssignments
    $out = @{}
    foreach ($ra in $ras) {
        $null = Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings
        $roles = @($ra.RoleDefinitionBindings | ForEach-Object { $_.Name } | Where-Object { $_ -notin $script:LimitedRoles })
        if ($roles.Count -eq 0) { continue }
        $login = if ($ra.Member.PSObject.Properties['LoginName'] -and $ra.Member.LoginName) { $ra.Member.LoginName } else { $ra.Member.Title }
        $out[$login.ToLower()] = @{ Title = $ra.Member.Title; Roles = $roles; PrincipalId = $ra.PrincipalId }
    }
    return $out
}

function Remove-SagamoreUnexpectedAssignments {
    <# Deletes role assignments on $Securable whose login is not in $KeepLogins (lower-case). #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)]$Securable, [Parameter(Mandatory)][string[]]$KeepLogins, [string]$Label = 'object')
    $ras = Get-PnPProperty -ClientObject $Securable -Property RoleAssignments
    $toDelete = @()
    foreach ($ra in $ras) {
        $null = Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings
        $roles = @($ra.RoleDefinitionBindings | ForEach-Object { $_.Name } | Where-Object { $_ -notin $script:LimitedRoles })
        if ($roles.Count -eq 0) { continue }
        $login = if ($ra.Member.PSObject.Properties['LoginName'] -and $ra.Member.LoginName) { $ra.Member.LoginName } else { $ra.Member.Title }
        if ($login.ToLower() -notin $KeepLogins) { $toDelete += [pscustomobject]@{ Ra = $ra; Login = $login; Title = $ra.Member.Title } }
    }
    foreach ($d in $toDelete) {
        if ($PSCmdlet.ShouldProcess("$Label", "Remove access for $($d.Title)")) {
            $d.Ra.DeleteObject()
            Write-SagamoreLog "  removed unexpected access: $($d.Title) on $Label" -Level SUCCESS
        }
    }
    if ($toDelete.Count) { Invoke-PnPQuery }
}

function Compare-SagamoreAccess {
    <#
        Pure comparison used by the audit. Desired: @{ login = role }. Actual: output of Get-SagamoreRoleAssignments.
        Returns rows describing missing, wrong-role and unexpected principals.
    #>
    param([Parameter(Mandatory)][hashtable]$Desired, [Parameter(Mandatory)][hashtable]$Actual, [string[]]$IgnoreLogins = @())
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($login in $Desired.Keys) {
        if (-not $Actual.ContainsKey($login)) {
            $rows.Add([pscustomobject]@{ Issue = 'Missing'; Login = $login; Expected = $Desired[$login]; Actual = '' })
        } elseif ($Desired[$login] -notin $Actual[$login].Roles -or $Actual[$login].Roles.Count -gt 1) {
            $rows.Add([pscustomobject]@{ Issue = 'WrongRole'; Login = $login; Expected = $Desired[$login]; Actual = ($Actual[$login].Roles -join ', ') })
        }
    }
    foreach ($login in $Actual.Keys) {
        if (-not $Desired.ContainsKey($login) -and $login -notin $IgnoreLogins) {
            $rows.Add([pscustomobject]@{ Issue = 'Unexpected'; Login = $login; Expected = ''; Actual = ($Actual[$login].Roles -join ', ') })
        }
    }
    return , $rows.ToArray()
}

Export-ModuleMember -Function *
