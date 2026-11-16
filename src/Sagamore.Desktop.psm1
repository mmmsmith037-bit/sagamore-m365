#Requires -Version 7.4
<#
    Sagamore.Desktop -- things that happen on the computer running the dashboard, not in the tenant:

      * OneDrive health: is the sync app installed, running, signed in to the Sagamore tenant, and is the
        Sagamore document library synced to this PC?
      * Office save settings: do Excel, Word and PowerPoint save into the synced Sagamore folder by default,
        with AutoSave on? (The dashboard can fix the user-level settings; policies are only reported.)
      * The desktop drop folder: files dragged into it are offered for upload to a Teams channel the next
        time the dashboard is opened (see scripts/Send-SagamoreDrop.ps1).

    Design: everything that reads the machine (registry, processes, the file system) is kept in small
    "Get-...Facts" / "Get-...Items" functions. Everything that DECIDES something (is OneDrive healthy? which
    files can be sent?) is a pure function over those facts, so it is covered by offline tests on any OS.

    Requires Sagamore.Common.psm1 to be imported first (uses Get-SagamoreSiteUrl, Get-SagamoreDropDestinations,
    Get-OrDefault, ConvertTo-SagamorePath).
#>

Set-StrictMode -Version Latest

# ------------------------------------------------------------------
# OneDrive
# ------------------------------------------------------------------

function Get-SagamoreOneDriveFacts {
    <#
        Reads raw OneDrive state for the CURRENT Windows user. Changes nothing.
        Returns a hashtable the pure evaluator (Test-SagamoreOneDriveState) understands:
          IsWindows   - $false on macOS/Linux (tests, CI); everything else is then empty
          ExePath     - path of OneDrive.exe, or $null when not installed
          Running     - OneDrive.exe is running in THIS logon session (another user's copy does not count)
          Accounts    - work/school accounts signed in to OneDrive: @{ Email; SpoResourceId; UserFolder }
          Libraries   - SharePoint libraries the sync app syncs: @{ Url; MountPoint }
        Sources (all HKCU, so no admin rights are needed):
          HKCU\Software\Microsoft\OneDrive\Accounts\Business<N>   -> UserEmail, SPOResourceId, UserFolder
          HKCU\Software\SyncEngines\Providers\OneDrive\<id>       -> UrlNamespace, MountPoint
        These keys are written by the OneDrive sync app itself; they are not a documented API, so every read
        is defensive and a missing value simply means "not known".
    #>
    [CmdletBinding()]
    param()

    $facts = @{ IsWindows = [bool]$IsWindows; ExePath = $null; Running = $false; Accounts = @(); Libraries = @() }
    if (-not $IsWindows) { return $facts }

    # --- installed? Per-user install is the default; per-machine installs live under Program Files. ---
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft OneDrive\OneDrive.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft OneDrive\OneDrive.exe')
    ) | Where-Object { $_ }
    $facts.ExePath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1

    # --- running in this session? (On a shared PC other users' OneDrive processes are visible too.) ---
    $mySession = [Diagnostics.Process]::GetCurrentProcess().SessionId
    $facts.Running = [bool](Get-Process -Name OneDrive -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $mySession })

    # --- signed-in work/school accounts ("Personal" is the consumer account and never the Sagamore one) ---
    $acctRoot = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    if (Test-Path $acctRoot) {
        $facts.Accounts = @(Get-ChildItem $acctRoot -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -like 'Business*' } | ForEach-Object {
            $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            $email = if ($p -and $p.PSObject.Properties['UserEmail']) { [string]$p.UserEmail } else { '' }
            if (-not $email) { return }   # an empty BusinessN key = an account slot that was signed out
            @{
                Email         = $email
                SpoResourceId = if ($p.PSObject.Properties['SPOResourceId']) { [string]$p.SPOResourceId } else { '' }
                UserFolder    = if ($p.PSObject.Properties['UserFolder']) { [string]$p.UserFolder } else { '' }
            }
        })
    }

    # --- libraries being synced (Teams/SharePoint libraries and "Add shortcut to My files" both appear here) ---
    $provRoot = 'HKCU:\Software\SyncEngines\Providers\OneDrive'
    if (Test-Path $provRoot) {
        $facts.Libraries = @(Get-ChildItem $provRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if (-not $p -or -not $p.PSObject.Properties['UrlNamespace']) { return }
            @{ Url = [string]$p.UrlNamespace; MountPoint = if ($p.PSObject.Properties['MountPoint']) { [string]$p.MountPoint } else { '' } }
        })
    }
    return $facts
}

function Test-SagamoreOneDriveState {
    <#
        PURE. Turns Get-SagamoreOneDriveFacts output into a status the dashboard can show:
          @{ Overall = 'ok'|'warn'|'fail'|'unsupported'; Checks = @( @{ Id; Label; Status; Detail; Fix } ); SyncedPath }
        Status per check: ok / warn / fail. Fix is a plain-English next step ('' when nothing to do).

        Signed in = a work account whose SharePoint host is <tenant>-my.sharepoint.com (the tenant in config).
        Falls back to the e-mail domain of the people in config when SPOResourceId is not recorded.
        Library synced = the sync app syncs a URL under the Sagamore site (https://<tenant>.sharepoint.com/sites/<alias>/...).
        A missing library sync is only a WARNING: uploads from the drop folder and the dashboard work without it.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Facts, [Parameter(Mandatory)][hashtable]$Config)

    if (-not $Facts.IsWindows) {
        return @{
            Overall = 'unsupported'; SyncedPath = $null
            Checks  = @(@{ Id = 'platform'; Label = 'OneDrive check'; Status = 'warn'; Detail = 'OneDrive can only be checked on Windows.'; Fix = '' })
        }
    }

    $tenant  = [string]$Config.tenant.name
    $siteUrl = (Get-SagamoreSiteUrl -Config $Config).TrimEnd('/').ToLower()
    $libUrl  = "$siteUrl/" + ([string]$Config.site.library).ToLower()
    $myHost  = "$tenant-my.sharepoint.com".ToLower()
    $domains = @(@($Config.people | ForEach-Object { ([string]$_.upn).Split('@')[-1].ToLower() }) +
                 @(@(Get-OrDefault $Config.site 'adminUpns' @()) | ForEach-Object { ([string]$_).Split('@')[-1].ToLower() }) | Select-Object -Unique)

    $checks = [System.Collections.Generic.List[hashtable]]::new()

    # 1. installed
    $installed = [bool]$Facts.ExePath
    $checks.Add(@{ Id = 'installed'; Label = 'OneDrive installed'; Status = $(if ($installed) { 'ok' } else { 'fail' })
        Detail = $(if ($installed) { $Facts.ExePath } else { 'OneDrive.exe was not found for this Windows user.' })
        Fix    = $(if ($installed) { '' } else { 'Install OneDrive from https://www.microsoft.com/microsoft-365/onedrive/download and sign in with your Sagamore account.' }) })

    # 2. running
    $running = [bool]$Facts.Running
    $checks.Add(@{ Id = 'running'; Label = 'OneDrive running'; Status = $(if ($running) { 'ok' } else { 'fail' })
        Detail = $(if ($running) { 'The sync app is running.' } else { 'The sync app is not running, so nothing on this computer is syncing.' })
        Fix    = $(if ($running -or -not $installed) { '' } else { 'Click "Start OneDrive".' }) })

    # 3. signed in to THIS tenant (a personal account or another company's account does not count)
    $accounts = @($Facts.Accounts)
    $match = @($accounts | Where-Object {
        $res = [string]$_.SpoResourceId
        if ($res) {
            try { ([uri]$res).Host.ToLower() -eq $myHost } catch { $false }
        } else {
            ([string]$_.Email).Split('@')[-1].ToLower() -in $domains
        }
    })
    $signedIn = $match.Count -gt 0
    $others = @($accounts | Where-Object { $_ -notin $match } | ForEach-Object { $_.Email })
    $checks.Add(@{ Id = 'signedIn'; Label = 'Signed in to Sagamore'; Status = $(if ($signedIn) { 'ok' } else { 'fail' })
        Detail = $(if ($signedIn) { "Signed in as $($match[0].Email)." }
                   elseif ($others.Count) { "Signed in only as $($others -join ', '), which is not a Sagamore account." }
                   else { 'No work or school account is signed in to OneDrive.' })
        Fix    = $(if ($signedIn) { '' } else { 'Open OneDrive (cloud icon by the clock) > Settings > Account > Add an account, and sign in with your Sagamore e-mail.' }) })

    # 4. Sagamore library synced to this PC (warning only)
    $lib = @($Facts.Libraries | Where-Object {
        $u = [uri]::UnescapeDataString([string]$_.Url).TrimEnd('/').ToLower()
        $u -eq $libUrl -or $u.StartsWith("$libUrl/") -or $u -eq $siteUrl -or $u.StartsWith("$siteUrl/")
    } | Select-Object -First 1)
    $synced = $lib.Count -gt 0
    $checks.Add(@{ Id = 'library'; Label = 'Sagamore files synced to this PC'; Status = $(if ($synced) { 'ok' } else { 'warn' })
        Detail = $(if ($synced) { "Synced to $($lib[0].MountPoint)" } else { 'The Sagamore document library is not synced to this computer.' })
        Fix    = $(if ($synced) { '' } else { 'Click "Open library", then "Sync" (or "Add shortcut to My files") in the toolbar. Only the folders you can open will sync.' }) })

    $statuses = @($checks | ForEach-Object { $_.Status })
    $overall = if ('fail' -in $statuses) { 'fail' } elseif ('warn' -in $statuses) { 'warn' } else { 'ok' }
    return @{ Overall = $overall; Checks = $checks.ToArray(); SyncedPath = $(if ($synced) { $lib[0].MountPoint } else { $null }) }
}

function Start-SagamoreOneDrive {
    <# Starts the OneDrive sync app for this user (it shows its own sign-in window if needed). Returns $true if started. #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][hashtable]$Facts)
    if (-not $Facts.IsWindows -or -not $Facts.ExePath -or $Facts.Running) { return $false }
    if ($PSCmdlet.ShouldProcess($Facts.ExePath, 'Start OneDrive')) {
        # /background starts it to the tray without opening a File Explorer window.
        Start-Process -FilePath $Facts.ExePath -ArgumentList '/background'
        return $true
    }
    return $false
}

# ------------------------------------------------------------------
# Office save settings
# ------------------------------------------------------------------
<#
    Goal: when someone in Excel, Word or PowerPoint presses Save, the file lands in the Sagamore SharePoint
    library (through the OneDrive-synced folder) and AutoSave is on. This uses Office's own features only,
    and the dashboard just makes sure they are set up.

    There is no single Office setting for "save to this SharePoint site by default". Office's cloud-first
    save (PreferCloudSaveLocations = 1) points at the user's PERSONAL OneDrive, not a SharePoint site. So the
    approach is:
      1. OneDrive syncs the Sagamore library to a local folder (checked by Test-SagamoreOneDriveState).
      2. PreferCloudSaveLocations = 0 ("Save to Computer by default") makes Office's default save location the
         per-app default folder...
      3. ...and each app's default folder is set to that synced folder, so a save there is uploaded to SharePoint
         by OneDrive and is the same file seen in Teams and on every other device.
      4. AutoSave works for files opened from a synced folder as long as nothing turns it off (a user
         "DontAutoSave" value or an admin policy) and Office/OneDrive integration is not blocked by policy.

    Everything the dashboard WRITES is under HKCU\Software\Microsoft\Office\16.0 (the user's own settings, no
    admin rights needed). Policy keys (HKCU\Software\Policies, HKLM) are only READ: if a policy gets in the way
    the check says so and names who has to change it, because a user-level change would be overridden anyway.

    Value names and locations (Microsoft 365 Apps / Office 16.0, Windows):
      Common\General\PreferCloudSaveLocations   DWORD  0 = "Save to Computer by default" ticked
      Excel\Options\DefaultPath                 EXPAND_SZ  Excel default local file location
      Word\Options\DOC-PATH                     EXPAND_SZ  Word default local file location
      PowerPoint\RecentFolderList\Default       EXPAND_SZ  PowerPoint default local file location
      Common\FileIO\EnableCloudOnlySaveAsMode   DWORD  1 = Save As shows cloud locations only (hides the synced folder)
      <App>\DontAutoSave                        DWORD  1 = legacy per-user "AutoSave off" switch
      Policies\...\<app>\AutoSaveByDefaultAdminChoice   set by the "Turn off AutoSave by default in <app>" policy
      HKLM\SOFTWARE\Policies\Microsoft\OneDrive\EnableAllOcsiClients   DWORD  0 = Office/OneDrive integration
                                                (AutoSave, co-authoring on synced files) blocked by an admin
#>

$script:OfficeRoot   = 'HKCU:\Software\Microsoft\Office\16.0'
$script:OfficePolicy = 'HKCU:\Software\Policies\Microsoft\Office\16.0'
# Where each app keeps its default local save folder. Order = order shown on the dashboard.
$script:OfficeApps = [ordered]@{
    Excel      = @{ Key = "$script:OfficeRoot\Excel\Options";                Name = 'DefaultPath' }
    Word       = @{ Key = "$script:OfficeRoot\Word\Options";                 Name = 'DOC-PATH' }
    PowerPoint = @{ Key = "$script:OfficeRoot\PowerPoint\RecentFolderList";  Name = 'Default' }
}

function Get-SagamoreRegValue {
    <# One registry value, or $null when the key or value does not exist. Never throws. #>
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Name)
    try {
        $p = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        return $p.$Name
    } catch { return $null }
}

function Get-SagamoreOfficeSaveFacts {
    <#
        Reads raw Office/OneDrive save settings for the CURRENT Windows user. Changes nothing.
        Returns a hashtable the pure functions below understand:
          IsWindows          - $false on macOS/Linux (tests, CI); everything else is then empty
          OfficeInstalled    - Microsoft 365 Apps / Office 16.0 found for this user or machine
          PreferCloud        - Common\General\PreferCloudSaveLocations (user), $null = never set (Office treats as 1)
          PreferCloudPolicy  - same value set by an admin policy, $null = no policy
          CloudOnly          - Common\FileIO\EnableCloudOnlySaveAsMode (user)
          CloudOnlyPolicy    - same, set by policy
          Paths              - @{ Excel = '...'; Word = '...'; PowerPoint = '...' } default local folders ($null = not set)
          DontAutoSave       - @{ Excel = 1; ... } legacy per-user AutoSave-off values ($null = not set)
          AutoSavePolicy     - @{ Excel = <value>; ... } "Turn off AutoSave by default" policy values ($null = no policy)
          OcsiPolicy         - HKLM OneDrive EnableAllOcsiClients policy ($null = no policy, the normal case)
    #>
    [CmdletBinding()]
    param()

    $facts = @{
        IsWindows = [bool]$IsWindows; OfficeInstalled = $false
        PreferCloud = $null; PreferCloudPolicy = $null; CloudOnly = $null; CloudOnlyPolicy = $null
        Paths = @{}; DontAutoSave = @{}; AutoSavePolicy = @{}; OcsiPolicy = $null
    }
    if (-not $IsWindows) { return $facts }

    # Click-to-Run (Microsoft 365 Apps) registers here; the HKCU 16.0 key exists once any Office app has run.
    $facts.OfficeInstalled = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration') -or (Test-Path $script:OfficeRoot)

    $facts.PreferCloud       = Get-SagamoreRegValue "$script:OfficeRoot\Common\General" 'PreferCloudSaveLocations'
    $facts.PreferCloudPolicy = Get-SagamoreRegValue "$script:OfficePolicy\Common\General" 'PreferCloudSaveLocations'
    $facts.CloudOnly         = Get-SagamoreRegValue "$script:OfficeRoot\Common\FileIO" 'EnableCloudOnlySaveAsMode'
    $facts.CloudOnlyPolicy   = Get-SagamoreRegValue "$script:OfficePolicy\Common\FileIO" 'EnableCloudOnlySaveAsMode'
    $facts.OcsiPolicy        = Get-SagamoreRegValue 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive' 'EnableAllOcsiClients'

    foreach ($app in $script:OfficeApps.Keys) {
        $a = $script:OfficeApps[$app]
        $facts.Paths[$app]          = Get-SagamoreRegValue $a.Key $a.Name
        $facts.DontAutoSave[$app]   = Get-SagamoreRegValue "$script:OfficeRoot\$app" 'DontAutoSave'
        $facts.AutoSavePolicy[$app] = Get-SagamoreRegValue "$script:OfficePolicy\$($app.ToLower())" 'AutoSaveByDefaultAdminChoice'
    }
    return $facts
}

function Test-SagamorePathUnder {
    <# PURE. True when $Path is $Root or a folder inside it (case-insensitive, ignores trailing slashes, expands %VARS%). #>
    param([AllowNull()][string]$Path, [Parameter(Mandatory)][string]$Root)
    if (-not $Path) { return $false }
    $p = [Environment]::ExpandEnvironmentVariables($Path).TrimEnd('\', '/').ToLower()
    $r = [Environment]::ExpandEnvironmentVariables($Root).TrimEnd('\', '/').ToLower()
    return ($p -eq $r -or $p.StartsWith("$r\") -or $p.StartsWith("$r/"))
}

function Get-SagamoreOfficeSaveChanges {
    <#
        PURE. The user-level registry changes needed so Office saves into the synced Sagamore folder with AutoSave
        on. Returns an array of:
          @{ Path; Name; Kind = 'DWord'|'ExpandString'; Value; Remove = $bool; Old; Why }
        Nothing is returned for things only an admin can change (policies); Test-SagamoreOfficeSaveState reports those.

        SyncedPath = local folder where OneDrive syncs the Sagamore library (Test-SagamoreOneDriveState.SyncedPath).
        Without it the save-location changes are skipped: pointing Office at a folder that does not exist yet
        would make every Save fail.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Facts, [AllowNull()][AllowEmptyString()][string]$SyncedPath)

    $changes = [System.Collections.Generic.List[hashtable]]::new()
    if (-not $Facts.IsWindows -or -not $Facts.OfficeInstalled) { return @() }

    if ($SyncedPath) {
        # Office treats "never set" as 1 (cloud first), so anything other than an explicit 0 needs setting.
        if ($Facts.PreferCloud -ne 0) {
            $changes.Add(@{ Path = "$script:OfficeRoot\Common\General"; Name = 'PreferCloudSaveLocations'; Kind = 'DWord'; Value = 0; Remove = $false
                Old = $Facts.PreferCloud; Why = 'Use the default folder below instead of personal OneDrive ("Save to Computer by default").' })
        }
        foreach ($app in $script:OfficeApps.Keys) {
            $cur = $Facts.Paths[$app]
            if (-not (Test-SagamorePathUnder -Path $cur -Root $SyncedPath)) {
                $a = $script:OfficeApps[$app]
                $changes.Add(@{ Path = $a.Key; Name = $a.Name; Kind = 'ExpandString'; Value = $SyncedPath; Remove = $false
                    Old = $cur; Why = "$app saves into the synced Sagamore folder by default." })
            }
        }
        # Cloud-only Save As hides PC folders, including the synced Sagamore folder set above.
        if ($Facts.CloudOnly -eq 1) {
            $changes.Add(@{ Path = "$script:OfficeRoot\Common\FileIO"; Name = 'EnableCloudOnlySaveAsMode'; Kind = 'DWord'; Value = 0; Remove = $false
                Old = 1; Why = 'Show the synced Sagamore folder in Save As (cloud-only Save As hides it).' })
        }
    }

    # Legacy per-user AutoSave-off switch: removing it restores Office's default (AutoSave on for cloud files).
    foreach ($app in $script:OfficeApps.Keys) {
        if ($Facts.DontAutoSave[$app] -eq 1) {
            $changes.Add(@{ Path = "$script:OfficeRoot\$app"; Name = 'DontAutoSave'; Kind = 'DWord'; Value = $null; Remove = $true
                Old = 1; Why = "Turn AutoSave back on by default in $app." })
        }
    }
    return $changes.ToArray()
}

function Test-SagamoreOfficeSaveState {
    <#
        PURE. Turns Get-SagamoreOfficeSaveFacts output (plus the synced library folder) into a dashboard status:
          @{ Overall = 'ok'|'warn'|'fail'|'unsupported'; Checks = @( @{ Id; Label; Status; Detail; Fix } ); Changes }
        Changes = Get-SagamoreOfficeSaveChanges output; when non-empty, "Fix save settings" can apply them.
        Problems the dashboard can fix itself are warnings; only a policy that blocks AutoSave/co-authoring
        outright is a failure, because no button here can fix it.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Facts, [AllowNull()][AllowEmptyString()][string]$SyncedPath)

    if (-not $Facts.IsWindows) {
        return @{ Overall = 'unsupported'; Changes = @()
            Checks = @(@{ Id = 'platform'; Label = 'Office save settings'; Status = 'warn'; Detail = 'Office settings can only be checked on Windows.'; Fix = '' }) }
    }
    if (-not $Facts.OfficeInstalled) {
        return @{ Overall = 'warn'; Changes = @()
            Checks = @(@{ Id = 'office'; Label = 'Office installed'; Status = 'warn'; Detail = 'Microsoft 365 Apps (Excel, Word, PowerPoint) were not found for this Windows user.'
                Fix = 'Install Microsoft 365 Apps from https://portal.office.com and sign in with your Sagamore account, then check again.' }) }
    }

    $changes = @(Get-SagamoreOfficeSaveChanges -Facts $Facts -SyncedPath $SyncedPath)
    $checks = [System.Collections.Generic.List[hashtable]]::new()

    # 1. Office <-> OneDrive integration. Without it AutoSave and co-authoring do not work on synced files.
    $blocked = $Facts.OcsiPolicy -eq 0
    $checks.Add(@{ Id = 'coauth'; Label = 'Office works with OneDrive'; Status = $(if ($blocked) { 'fail' } else { 'ok' })
        Detail = $(if ($blocked) { 'A policy on this PC stops Office and OneDrive working together, so AutoSave and co-authoring are off for synced files.' }
                   else { 'Office can AutoSave and co-author files in synced folders.' })
        Fix    = $(if ($blocked) { 'Ask the IT admin to remove the OneDrive policy "Coauthor and share in Office desktop apps" (EnableAllOcsiClients = 0) from this PC.' } else { '' }) })

    # 2. AutoSave on by default for cloud files.
    $userOff   = @($script:OfficeApps.Keys | Where-Object { $Facts.DontAutoSave[$_] -eq 1 })
    $policyOff = @($script:OfficeApps.Keys | Where-Object { $null -ne $Facts.AutoSavePolicy[$_] -and $Facts.AutoSavePolicy[$_] -ne 0 })
    $autoOk = -not $userOff.Count -and -not $policyOff.Count
    $checks.Add(@{ Id = 'autosave'; Label = 'AutoSave on by default'; Status = $(if ($autoOk) { 'ok' } else { 'warn' })
        Detail = $(if ($autoOk) { 'Files in SharePoint or OneDrive save every few seconds.' }
                   else { @(
                       $(if ($userOff.Count) { "Turned off on this PC for $($userOff -join ', ')." })
                       $(if ($policyOff.Count) { "An admin policy turns it off by default for $($policyOff -join ', ')." })
                   ) -join ' ' })
        Fix    = $(if ($autoOk) { '' }
                   elseif ($userOff.Count -and -not $policyOff.Count) { 'Click "Fix save settings".' }
                   elseif ($userOff.Count) { 'Click "Fix save settings", and ask the IT admin to remove the "Turn off AutoSave by default" policy.' }
                   else { 'Ask the IT admin to remove the "Turn off AutoSave by default" policy. Until then, switch AutoSave on with the toggle at the top left of each file.' }) })

    # 3. Default save location = the synced Sagamore folder.
    if (-not $SyncedPath) {
        $checks.Add(@{ Id = 'saveLocation'; Label = 'Saves go to Sagamore'; Status = 'warn'
            Detail = 'Office cannot default to the Sagamore library until it is synced to this PC.'
            Fix    = 'Sync the library first (see the OneDrive card), then click "Fix save settings".' })
    } else {
        $pending = @($changes | Where-Object { -not $_.Remove })
        $policyCloud = $Facts.PreferCloudPolicy -eq 1
        $policyCloudOnly = $Facts.CloudOnlyPolicy -eq 1
        $status = if ($pending.Count -or $policyCloud -or $policyCloudOnly) { 'warn' } else { 'ok' }
        $detail = if ($status -eq 'ok') { "Excel, Word and PowerPoint save to $SyncedPath, which OneDrive keeps in SharePoint and Teams." }
                  else { @(
                      $(if ($Facts.PreferCloud -ne 0) { 'Office saves to personal OneDrive by default, not the Sagamore library.' }
                        elseif (@($pending | Where-Object { $_.Kind -eq 'ExpandString' }).Count) { 'Some Office apps save to a folder outside Sagamore by default.' }
                        elseif ($pending.Count) { 'Save As is in cloud-only mode, which hides the synced Sagamore folder.' })
                      $(if ($policyCloud) { 'An admin policy forces saving to personal OneDrive first.' })
                      $(if ($policyCloudOnly) { 'An admin policy hides PC folders (including the synced Sagamore folder) in Save As.' })
                  ) -join ' ' }
        $fix = @(
            $(if ($pending.Count) { 'Click "Fix save settings", then close and reopen Excel, Word and PowerPoint.' })
            $(if ($policyCloud -or $policyCloudOnly) { 'Ask the IT admin to remove the Office policies "PreferCloudSaveLocations" / "EnableCloudOnlySaveAsMode" for Sagamore users.' })
        ) -join ' '
        $checks.Add(@{ Id = 'saveLocation'; Label = 'Saves go to Sagamore'; Status = $status; Detail = $detail; Fix = $fix })
    }

    $statuses = @($checks | ForEach-Object { $_.Status })
    $overall = if ('fail' -in $statuses) { 'fail' } elseif ('warn' -in $statuses) { 'warn' } else { 'ok' }
    return @{ Overall = $overall; Checks = $checks.ToArray(); Changes = $changes }
}

function Set-SagamoreOfficeSaveDefaults {
    <#
        Applies Get-SagamoreOfficeSaveChanges output. Returns the changes that were made (each with its Old value,
        so the caller can log them and they can be put back by hand).

        Defence in depth: refuses any change outside HKCU\Software\Microsoft\Office\16.0, so this can never be
        used to write policies or machine-wide keys whatever it is handed. Office reads these at start-up, so
        apps that are already open keep the old behaviour until they are restarted.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Changes)

    $done = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($c in $Changes) {
        if (-not ([string]$c.Path).StartsWith("$script:OfficeRoot\", [StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to change '$($c.Path)': only $script:OfficeRoot is managed."
        }
        $target = "$($c.Path)\$($c.Name)"
        if ($c.Remove) {
            if ($PSCmdlet.ShouldProcess($target, 'Remove value')) {
                Remove-ItemProperty -Path $c.Path -Name $c.Name -ErrorAction SilentlyContinue
                $done.Add($c)
            }
        } elseif ($PSCmdlet.ShouldProcess($target, "Set to '$($c.Value)'")) {
            if (-not (Test-Path $c.Path)) { New-Item -Path $c.Path -Force | Out-Null }
            New-ItemProperty -Path $c.Path -Name $c.Name -PropertyType $c.Kind -Value $c.Value -Force | Out-Null
            $done.Add($c)
        }
    }
    return $done.ToArray()
}

# ------------------------------------------------------------------
# Drop folder
# ------------------------------------------------------------------

# Never offered for upload: OS/Office litter and files still being written by a browser download.
$script:DropIgnore = @('desktop.ini', 'thumbs.db', '.ds_store')
$script:DropIgnoreLike = @('~$*', '*.tmp', '*.crdownload', '*.part', '*.partial', '*.download')
$script:DropSentName = '_Sent'

function Get-SagamoreDefaultDropFolder {
    <# "Sagamore Drop" on the user's desktop. Follows OneDrive "Desktop backup" if it has moved the desktop. #>
    $desktop = [Environment]::GetFolderPath('Desktop')
    if (-not $desktop) { $desktop = Join-Path $HOME 'Desktop' }
    Join-Path $desktop 'Sagamore Drop'
}

function Initialize-SagamoreDropFolder {
    <#
        Creates the drop folder (and its _Sent archive) if missing. Returns the full path.
        Puts a short "How this works" note in _Sent rather than the drop folder itself, so the note is never
        mistaken for a file to upload.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$Path)
    $full = [IO.Path]::GetFullPath($Path)
    $sent = Join-Path $full $script:DropSentName
    if (-not (Test-Path $sent) -and $PSCmdlet.ShouldProcess($full, 'Create drop folder')) {
        New-Item -ItemType Directory -Path $sent -Force | Out-Null
        $note = @(
            'Sagamore Drop'
            '============='
            'Drag files into the "Sagamore Drop" folder. The next time the Sagamore dashboard is opened it asks'
            'which Teams channel each batch belongs in, shows who will be able to open it, and uploads it.'
            'Uploaded files take the permissions of the channel folder they land in.'
            ''
            'After a successful upload the local copy is moved here, into _Sent\<date>. You can delete it once you'
            'have checked the file in Teams.'
            ''
            'Folders are not uploaded (a folder named "Invoices" would be open to everyone in the channel until'
            'the next permissions run). Drop the files themselves, or zip the folder first.'
        ) -join [Environment]::NewLine
        Set-Content -Path (Join-Path $sent 'README.txt') -Value $note -Encoding utf8
    }
    return $full
}

function Test-SagamoreDropIgnored {
    <# PURE. True for names the drop folder silently skips (our own _Sent archive, hidden/temp/Office lock files). #>
    param([Parameter(Mandatory)][string]$Name)
    $n = $Name.ToLower()
    if ($n.StartsWith('_') -or $n.StartsWith('.')) { return $true }
    if ($n -in $script:DropIgnore) { return $true }
    foreach ($p in $script:DropIgnoreLike) { if ($n -like $p) { return $true } }
    return $false
}

function Test-SagamoreFileBusy {
    <# True when another program still has the file open for writing (e.g. a download or a save in progress). #>
    param([Parameter(Mandatory)][string]$Path)
    try {
        $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $fs.Dispose()
        return $false
    } catch [IO.IOException] { return $true }
    catch [UnauthorizedAccessException] { return $true }
}

function Get-SagamoreDropItems {
    <#
        Lists what is waiting in the drop folder. Top level only.
          @{ Name; Kind = 'file'|'folder'|'busy'; Size; Modified; Sendable; Reason }
        Only Kind 'file' is Sendable. Folders are listed (so the user is told why they were not offered)
        but never uploaded; see Initialize-SagamoreDropFolder for the reason.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path -PathType Container)) { return @() }
    $items = foreach ($i in Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue) {
        if (Test-SagamoreDropIgnored $i.Name) { continue }
        if ($i.Attributes -band ([IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System)) { continue }
        if ($i.PSIsContainer) {
            @{ Name = $i.Name; Kind = 'folder'; Size = 0; Modified = $i.LastWriteTime.ToString('o'); Sendable = $false
               Reason = 'Folders are not uploaded. Drop the files inside it, or zip it first.' }
        } elseif (Test-SagamoreFileBusy $i.FullName) {
            @{ Name = $i.Name; Kind = 'busy'; Size = $i.Length; Modified = $i.LastWriteTime.ToString('o'); Sendable = $false
               Reason = 'Still open in another program (or still downloading). Close it and try again.' }
        } else {
            @{ Name = $i.Name; Kind = 'file'; Size = $i.Length; Modified = $i.LastWriteTime.ToString('o'); Sendable = $true; Reason = '' }
        }
    }
    return @($items | Sort-Object { $_.Name })
}

function Move-SagamoreDropItemToSent {
    <#
        After a successful upload: moves the local file to <drop>\_Sent\<yyyy-MM-dd>\, adding " (2)", " (3)"...
        if that name is already there. Returns the new path. Never deletes anything.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$DropFolder, [Parameter(Mandatory)][string]$Name, [datetime]$When = (Get-Date))
    $src = Join-Path $DropFolder $Name
    $dir = Join-Path (Join-Path $DropFolder $script:DropSentName) ($When.ToString('yyyy-MM-dd'))
    $base = [IO.Path]::GetFileNameWithoutExtension($Name); $ext = [IO.Path]::GetExtension($Name)
    $dest = Join-Path $dir $Name
    for ($n = 2; Test-Path -LiteralPath $dest; $n++) { $dest = Join-Path $dir "$base ($n)$ext" }
    if ($PSCmdlet.ShouldProcess($src, "Move to $dest")) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Move-Item -LiteralPath $src -Destination $dest
    }
    return $dest
}

function Test-SagamoreDropRequest {
    <#
        PURE (apart from looking at the drop folder listing it is given). Validates an upload request before
        anything signs in. Used by the dashboard server (to refuse bad requests at once) AND by
        Send-SagamoreDrop.ps1 (defence in depth: the script trusts nothing it is handed).

        Request: @{ destination = 'Clients/Meridian-Group/Invoices'; files = @('a.pdf'); overwrite = $false; acknowledged = $true }
        Items:   output of Get-SagamoreDropItems
        Returns: @{ Errors = [string[]]; Destination = <entry from Get-SagamoreDropDestinations> or $null; Files = [string[]] }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config, [Parameter(Mandatory)]$Request, [AllowEmptyCollection()][object[]]$Items = @())

    $errs = [System.Collections.Generic.List[string]]::new()
    $dest = ConvertTo-SagamorePath ([string](Get-OrDefault $Request 'destination' ''))
    $match = $null
    foreach ($d in Get-SagamoreDropDestinations -Config $Config) {
        if (@($d.Folders | Where-Object { $_ -ieq $dest }).Count) { $match = $d; break }
    }
    if (-not $dest) { $errs.Add('Pick a channel to send the files to.') }
    elseif (-not $match) { $errs.Add("'$dest' is not a Teams channel folder managed by the Sagamore config.") }

    # Must be the boolean true. PowerShell's -eq would also accept 'true', 'yes'-ish strings and 1.
    $ack = Get-OrDefault $Request 'acknowledged' $false
    if (-not ($ack -is [bool] -and $ack)) {
        $errs.Add('Confirm that the files will take the permissions of the channel folder.')
    }

    $byName = @{}
    foreach ($i in $Items) { $byName[$i.Name.ToLower()] = $i }
    $files = @(@(Get-OrDefault $Request 'files' @()) | ForEach-Object { [string]$_ } | Select-Object -Unique)
    if (-not $files.Count) { $errs.Add('Pick at least one file.') }
    foreach ($f in $files) {
        # Names only: anything with a path separator could point outside the drop folder.
        if ($f -match '[\\/]' -or $f -in @('.', '..')) { $errs.Add("'$f' is not a file name."); continue }
        $i = $byName[$f.ToLower()]
        if (-not $i) { $errs.Add("'$f' is not in the drop folder any more."); continue }
        if (-not $i.Sendable) { $errs.Add("'$f': $($i.Reason)") }
    }
    return @{ Errors = $errs.ToArray(); Destination = $match; Files = $files }
}

Export-ModuleMember -Function *-Sagamore*
