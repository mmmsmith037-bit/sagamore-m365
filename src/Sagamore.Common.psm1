#Requires -Version 7.4
<#
    Sagamore.Common -- shared, mostly side-effect-free helpers used by every script.

    Everything that decides WHAT should exist (config validation, folder plan,
    channel plan, principal resolution) lives here and is covered by Pester tests.
    The scripts in /scripts only execute that plan against the tenant.
#>

Set-StrictMode -Version Latest

# ------------------------------------------------------------------
# Constants
# ------------------------------------------------------------------
$script:ValidRoles      = @('Full Control', 'Edit', 'Contribute', 'Add Only', 'Read')
$script:GroupKeys       = @('Owner', 'Member', 'Support')
$script:SharingValues   = @('Disabled', 'ExistingExternalUserSharingOnly', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing')
$script:InvalidPathChar = '["*:<>?\\|#%]'          # SharePoint-illegal or problematic in folder names
$script:InvalidChannel  = '[~#%&*{}+/\\:<>?|''"]'   # Teams-illegal channel name characters
$script:UpnPattern      = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
$script:LogFile         = $null

# The custom permission level used for "upload but cannot delete".
$script:AddOnlyRole = @{
    Name        = 'Add Only'
    Clone       = 'Contribute'
    Exclude     = @('DeleteListItems', 'DeleteVersions')
    Description = 'Can view, add and edit items but cannot delete. Created by Sagamore installer.'
}

function Get-SagamoreValidRoles { $script:ValidRoles }
function Get-SagamoreAddOnlyRole { $script:AddOnlyRole }

# ------------------------------------------------------------------
# Logging
# ------------------------------------------------------------------
function Initialize-SagamoreLog {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $script:LogFile = $Path
}

function Write-SagamoreLog {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)][AllowEmptyString()][string]$Message = '',
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'SKIP', 'PLAN')][string]$Level = 'INFO'
    )
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level.PadRight(7), $Message
    if ($script:LogFile) { Add-Content -Path $script:LogFile -Value $line }
    $color = @{ INFO = 'Cyan'; WARN = 'Yellow'; ERROR = 'Red'; SUCCESS = 'Green'; SKIP = 'DarkGray'; PLAN = 'Magenta' }[$Level]
    Write-Host $line -ForegroundColor $color
}

function Write-SagamoreSection {
    param([string]$Title)
    Write-SagamoreLog ''
    Write-SagamoreLog ('=' * 60)
    Write-SagamoreLog "  $Title"
    Write-SagamoreLog ('=' * 60)
}

# ------------------------------------------------------------------
# Config
# ------------------------------------------------------------------
function Read-SagamoreConfig {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) {
        throw "Config file not found: $Path. Copy config/sagamore.config.example.json to config/sagamore.config.json and edit it."
    }
    try {
        return Get-Content -Path $Path -Raw | ConvertFrom-Json -AsHashtable -Depth 20
    } catch {
        throw "Config file is not valid JSON ($Path): $($_.Exception.Message)"
    }
}

function Save-SagamoreConfig {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][hashtable]$Config, [Parameter(Mandatory)][string]$Path)
    if ($PSCmdlet.ShouldProcess($Path, 'Write config')) {
        $Config | ConvertTo-Json -Depth 20 | Set-Content -Path $Path -Encoding utf8
    }
}

function Get-OrDefault {
    param($Table, [string]$Key, $Default = $null)
    if ($Table -is [System.Collections.IDictionary] -and $Table.Contains($Key) -and $null -ne $Table[$Key]) { return $Table[$Key] }
    return $Default
}

function ConvertTo-SagamorePath {
    <# Normalises a folder path: forward slashes, no leading/trailing slashes, no empty segments. #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Path)
    (($Path -replace '\\', '/') -split '/' | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }) -join '/'
}

function ConvertTo-SagamoreFolderName {
    <# Turns an arbitrary local folder/client name into a safe SharePoint folder name. #>
    param([Parameter(Mandatory)][string]$Name)
    $n = $Name -replace $script:InvalidPathChar, '' -replace '\s+', '-' -replace '-{2,}', '-'
    $n = $n.Trim('-', '.', ' ')
    if ($n -eq '') { throw "Name '$Name' has no usable characters for a SharePoint folder." }
    if ($n.Length -gt 100) { $n = $n.Substring(0, 100).TrimEnd('-', '.') }
    return $n
}

function ConvertTo-SagamoreChannelName {
    param([Parameter(Mandatory)][string]$Name)
    $n = ($Name -replace $script:InvalidChannel, '').Trim().TrimEnd('.')
    if ($n.Length -gt 50) { $n = $n.Substring(0, 50).Trim() }
    if ($n -eq '') { throw "Name '$Name' has no usable characters for a Teams channel." }
    return $n
}

function Test-SagamoreConfig {
    <#
        Returns an array of human-readable problems. Empty array = valid.
        -AllowPlaceholderClientId lets the example config and -WhatIf runs pass before app registration.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param([Parameter(Mandatory)][hashtable]$Config, [switch]$AllowPlaceholderClientId)

    $errs = [System.Collections.Generic.List[string]]::new()

    foreach ($section in 'tenant', 'site', 'groups', 'people', 'folders', 'clientTemplate', 'clients', 'libraryAccess') {
        if (-not $Config.Contains($section)) { $errs.Add("Missing required section '$section'.") }
    }
    if ($errs.Count) { return $errs.ToArray() }

    # tenant
    $tenant = $Config.tenant
    if ((Get-OrDefault $tenant 'name' '') -notmatch '^[a-z0-9][a-z0-9-]{0,62}$') {
        $errs.Add("tenant.name must be the tenant prefix only (e.g. 'sagamorealliance' from sagamorealliance.sharepoint.com).")
    }
    $clientId = [string](Get-OrDefault $tenant 'clientId' '')
    $guid = [guid]::Empty
    if (-not [guid]::TryParse($clientId, [ref]$guid)) {
        $errs.Add('tenant.clientId must be a GUID. Run scripts/Register-SagamoreApp.ps1 to create one.')
    } elseif ($guid -eq [guid]::Empty -and -not $AllowPlaceholderClientId) {
        $errs.Add('tenant.clientId is still the placeholder. Run scripts/Register-SagamoreApp.ps1 first.')
    }

    # site
    $site = $Config.site
    if ((Get-OrDefault $site 'alias' '') -notmatch '^[A-Za-z0-9][A-Za-z0-9-]{0,62}$') { $errs.Add('site.alias may contain only letters, digits and hyphens.') }
    if (-not (Get-OrDefault $site 'title' '')) { $errs.Add('site.title is required.') }
    if (-not (Get-OrDefault $site 'library' '')) { $errs.Add('site.library is required (normally "Shared Documents").') }
    foreach ($a in @(Get-OrDefault $site 'adminUpns' @())) {
        if ([string]$a -notmatch $script:UpnPattern) { $errs.Add("site.adminUpns: '$a' is not a valid UPN.") }
    }
    $sharing = Get-OrDefault $site 'externalSharing' 'ExistingExternalUserSharingOnly'
    if ($sharing -notin $script:SharingValues) { $errs.Add("site.externalSharing must be one of: $($script:SharingValues -join ', ').") }

    # groups
    foreach ($k in $script:GroupKeys) {
        if (-not (Get-OrDefault $Config.groups $k '')) { $errs.Add("groups.$k must name an Entra ID security group.") }
    }

    # people
    $upns = @{}
    $ownerCount = 0
    foreach ($p in @($Config.people)) {
        $upn = [string](Get-OrDefault $p 'upn' '')
        if ($upn -notmatch $script:UpnPattern) { $errs.Add("people: '$upn' is not a valid UPN."); continue }
        if ($upns.ContainsKey($upn.ToLower())) { $errs.Add("people: '$upn' is listed twice.") }
        $upns[$upn.ToLower()] = $true
        $role = Get-OrDefault $p 'role' ''
        if ($role -notin $script:GroupKeys) { $errs.Add("people: '$upn' has role '$role'; use Owner, Member or Support.") }
        if ($role -eq 'Owner') { $ownerCount++ }
    }
    if ($ownerCount -lt 1) { $errs.Add('people: at least one person must have role Owner.') }

    # access maps
    $checkAccess = {
        param($access, [string]$where)
        if ($null -eq $access) { return }
        if ($access -isnot [System.Collections.IDictionary]) { $errs.Add("$where.access must be an object of principal -> role."); return }
        foreach ($principal in $access.Keys) {
            if ($principal -notin $script:GroupKeys) {
                if ($principal -notmatch $script:UpnPattern) { $errs.Add("$where.access: '$principal' is not Owner/Member/Support or a UPN.") }
                elseif (-not $upns.ContainsKey($principal.ToLower())) { $errs.Add("$where.access: '$principal' is not listed in people.") }
            }
            if ($access[$principal] -notin $script:ValidRoles) {
                $errs.Add("$where.access: role '$($access[$principal])' for '$principal' is not one of: $($script:ValidRoles -join ', ').")
            }
        }
    }

    & $checkAccess $Config.libraryAccess 'libraryAccess'
    if (-not (Get-OrDefault $Config.libraryAccess 'Owner')) { $errs.Add('libraryAccess must grant Owner access.') }

    $seen = @{}
    foreach ($f in @($Config.folders)) {
        $raw = [string](Get-OrDefault $f 'path' '')
        $path = ConvertTo-SagamorePath $raw
        if (-not $path) { $errs.Add('folders: an entry has an empty path.'); continue }
        if ($path -match $script:InvalidPathChar) { $errs.Add("folders: '$path' contains a character SharePoint does not allow.") }
        if ($seen.ContainsKey($path.ToLower())) { $errs.Add("folders: '$path' is listed twice.") }
        $seen[$path.ToLower()] = $true
        & $checkAccess (Get-OrDefault $f 'access') "folders['$path']"
    }

    $tpl = $Config.clientTemplate
    & $checkAccess (Get-OrDefault $tpl 'access') 'clientTemplate'
    foreach ($s in @(Get-OrDefault $tpl 'subfolders' @())) {
        & $checkAccess (Get-OrDefault $s 'access') "clientTemplate.subfolders['$($s.path)']"
    }

    $clientNames = @{}
    foreach ($c in @($Config.clients)) {
        $name = [string](Get-OrDefault $c 'name' '')
        if (-not $name) { $errs.Add('clients: an entry has no name.'); continue }
        if ($name -match $script:InvalidPathChar -or $name -match '/') { $errs.Add("clients: '$name' contains a character SharePoint does not allow.") }
        if ($clientNames.ContainsKey($name.ToLower())) { $errs.Add("clients: '$name' is listed twice.") }
        $clientNames[$name.ToLower()] = $true
        & $checkAccess (Get-OrDefault $c 'access') "clients['$name']"
        foreach ($cf in @(Get-OrDefault $c 'folders' @())) {
            $cp = ConvertTo-SagamorePath ([string](Get-OrDefault $cf 'path' ''))
            if (-not $cp) { $errs.Add("clients['$name'].folders: an entry has an empty path.") }
            elseif ($cp -match $script:InvalidPathChar) { $errs.Add("clients['$name'].folders: '$cp' contains a character SharePoint does not allow.") }
            & $checkAccess (Get-OrDefault $cf 'access') "clients['$name'].folders['$cp']"
        }
    }

    return $errs.ToArray()
}

# ------------------------------------------------------------------
# Plans
# ------------------------------------------------------------------
function Resolve-SagamoreAccess {
    <#
        Normalises an access map: copies it, and guarantees the Owner group keeps Full Control
        so nobody can lock Joe out of a folder from the config or dashboard.
    #>
    param($Access)
    if ($null -eq $Access) { return $null }
    $out = [ordered]@{ Owner = 'Full Control' }
    foreach ($k in $Access.Keys) { if ($k -ne 'Owner') { $out[$k] = $Access[$k] } }
    return $out
}

function Get-SagamoreFolderPlan {
    <#
        Expands the config into an ordered list of folders (parents before children):
          @{ Path = 'Clients/Meridian-Group/Invoices'; Access = [ordered]@{...} or $null; Source = '...' }
        Access $null  -> folder inherits from its parent.
        Access map    -> inheritance is broken and replaced by exactly this map.
        -ExtraClients adds client folder names discovered on the site (e.g. created by hand).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config, [string[]]$ExtraClients = @())

    $plan  = [ordered]@{}
    $add = {
        param([string]$path, $access, [string]$source)
        $path = ConvertTo-SagamorePath $path
        # make sure every ancestor exists in the plan (inheriting) before the child
        $parts = $path -split '/'
        for ($i = 1; $i -lt $parts.Count; $i++) {
            $anc = ($parts[0..($i - 1)]) -join '/'
            if (-not $plan.Contains($anc.ToLower())) {
                $plan[$anc.ToLower()] = [pscustomobject]@{ Path = $anc; Access = $null; Source = "implied by $path" }
            }
        }
        $key = $path.ToLower()
        if ($plan.Contains($key)) {
            if ($null -ne $access) { $plan[$key].Access = Resolve-SagamoreAccess $access; $plan[$key].Source = $source }
        } else {
            $plan[$key] = [pscustomobject]@{ Path = $path; Access = (Resolve-SagamoreAccess $access); Source = $source }
        }
    }

    foreach ($f in @($Config.folders)) { & $add $f.path (Get-OrDefault $f 'access') 'folders' }

    $tpl        = $Config.clientTemplate
    $tplAccess  = Get-OrDefault $tpl 'access'
    $tplSubs    = @(Get-OrDefault $tpl 'subfolders' @())
    $lockedNames = @(Get-OrDefault $tpl 'lockedFolderNames' @()) | ForEach-Object { $_.ToLower() }
    $lockedAccess = @{}
    foreach ($s in $tplSubs) {
        $leaf = (ConvertTo-SagamorePath $s.path).Split('/')[-1].ToLower()
        if ($leaf -in $lockedNames -and (Get-OrDefault $s 'access')) { $lockedAccess[$leaf] = $s.access }
    }

    $clients = [System.Collections.Generic.List[object]]::new()
    foreach ($c in @($Config.clients)) { $clients.Add($c) }
    $known = @($Config.clients | ForEach-Object { $_.name.ToLower() })
    foreach ($x in $ExtraClients) {
        if ($x -and $x.ToLower() -notin $known) { $clients.Add(@{ name = $x }); $known += $x.ToLower() }
    }

    foreach ($c in $clients) {
        $base = "Clients/$($c.name)"
        & $add $base (Get-OrDefault $c 'access' $tplAccess) "client $($c.name)"
        foreach ($s in $tplSubs) { & $add "$base/$($s.path)" (Get-OrDefault $s 'access') "clientTemplate" }
        foreach ($cf in @(Get-OrDefault $c 'folders' @())) {
            & $add "$base/$($cf.path)" (Get-OrDefault $cf 'access') "client $($c.name)"
        }
    }

    # Locked folder names (e.g. Invoices) are forced to the locked access at ANY depth under Clients/.
    foreach ($item in $plan.Values) {
        $leaf = $item.Path.Split('/')[-1].ToLower()
        if ($item.Path -like 'Clients/*/*' -and $lockedAccess.ContainsKey($leaf)) {
            $item.Access = Resolve-SagamoreAccess $lockedAccess[$leaf]
            $item.Source = "locked ($leaf)"
        }
    }

    # parents before children, stable within the same depth
    $i = 0
    return @($plan.Values |
        ForEach-Object { [pscustomobject]@{ Item = $_; Depth = ($_.Path.Split('/').Count); Order = $i++ } } |
        Sort-Object Depth, Order |
        ForEach-Object { $_.Item })
}

function Get-SagamoreLockedAccess {
    <# Access map that must be applied to a locked folder name (e.g. 'Invoices'), or $null. #>
    param([Parameter(Mandatory)][hashtable]$Config, [Parameter(Mandatory)][string]$FolderName)
    $tpl = $Config.clientTemplate
    $locked = @(Get-OrDefault $tpl 'lockedFolderNames' @()) | ForEach-Object { $_.ToLower() }
    if ($FolderName.ToLower() -notin $locked) { return $null }
    foreach ($s in @(Get-OrDefault $tpl 'subfolders' @())) {
        if ((ConvertTo-SagamorePath $s.path).Split('/')[-1] -ieq $FolderName -and (Get-OrDefault $s 'access')) {
            return Resolve-SagamoreAccess $s.access
        }
    }
    return Resolve-SagamoreAccess @{ Owner = 'Full Control' }
}

function Get-SagamoreChannelPlan {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config, [string[]]$ExtraClients = @())
    $teams = Get-OrDefault $Config 'teams' @{}
    if (-not (Get-OrDefault $teams 'enabled' $true)) { return @() }
    $names = [System.Collections.Generic.List[object]]::new()
    foreach ($x in @(Get-OrDefault $teams 'extraChannels' @())) {
        $names.Add([pscustomobject]@{ Channel = (ConvertTo-SagamoreChannelName $x); Folder = $null })
    }
    if (Get-OrDefault $teams 'channelPerClient' $true) {
        $all = @($Config.clients | ForEach-Object { $_.name }) + $ExtraClients | Where-Object { $_ } | Select-Object -Unique
        foreach ($c in $all) {
            $names.Add([pscustomobject]@{ Channel = (ConvertTo-SagamoreChannelName $c); Folder = "Clients/$c" })
        }
    }
    # Teams channel names are case-insensitively unique
    return @($names | Group-Object { $_.Channel.ToLower() } | ForEach-Object { $_.Group[0] })
}

function Resolve-SagamorePrincipal {
    <#
        Maps a config principal key to what SharePoint needs:
          Owner/Member/Support -> Entra group claim  c:0t.c|tenant|<objectId>
          user@domain          -> the UPN (SharePoint resolves it)
        $GroupIds is a hashtable of Owner/Member/Support -> Entra object id.
    #>
    param([Parameter(Mandatory)][string]$Key, [Parameter(Mandatory)][hashtable]$GroupIds)
    if ($Key -in $script:GroupKeys) {
        $id = $GroupIds[$Key]
        if (-not $id) { throw "No Entra group id known for '$Key'." }
        return "c:0t.c|tenant|$id"
    }
    return $Key
}

function Get-SagamoreSiteUrl {
    param([Parameter(Mandatory)][hashtable]$Config)
    "https://$($Config.tenant.name).sharepoint.com/sites/$($Config.site.alias)"
}

function Get-SagamoreRequiredModules {
    @(@{ Name = 'PnP.PowerShell'; MinimumVersion = '3.0.0' })
}

function Assert-SagamoreModules {
    foreach ($m in Get-SagamoreRequiredModules) {
        $found = Get-Module -ListAvailable -Name $m.Name | Where-Object { $_.Version -ge [version]$m.MinimumVersion }
        if (-not $found) {
            throw "Module $($m.Name) $($m.MinimumVersion)+ is required. Install it with: Install-Module $($m.Name) -Scope CurrentUser"
        }
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
}

Export-ModuleMember -Function *-Sagamore*, Get-OrDefault, ConvertTo-SagamorePath
