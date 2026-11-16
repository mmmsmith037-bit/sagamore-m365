#Requires -Version 7.4
<#
.SYNOPSIS
    Opens the Permission Manager with one-click Save, Preview, Apply and Audit.

.DESCRIPTION
    Starts a small web server on this computer only (http://localhost:<Port>), opens the dashboard in
    the default browser and lets it:
      - load and save config/sagamore.config.json directly (validated first; the previous file is backed up)
      - run the installer and audit, streaming their output into the page

    Only a fixed list of commands can be run (see $Actions). Nothing typed in the page reaches the command line.
    Requests must carry a random token that is handed to the browser in the URL fragment, so other web pages
    and other users on the machine cannot drive it. Apply is refused until a Preview of the exact same
    config has succeeded.

    Sign-in to Microsoft 365 happens in the installer's own browser window, exactly as when it is run by hand.
    Press Ctrl+C in this window to stop the server.

    Desktop features (see src/Sagamore.Desktop.psm1):
      - OneDrive: on start-up checks that the OneDrive sync app is installed, running, signed in to the
        Sagamore tenant and syncing the Sagamore library. If it is installed but not running it is started.
        While the dashboard is open the check repeats every minute and OneDrive is restarted if it quits
        (at most once every 5 minutes, so a crashing OneDrive is not restarted in a loop).
      - Office save settings: checks that Excel, Word and PowerPoint save into the synced Sagamore library by
        default with AutoSave on. "Fix save settings" in the page changes the user's own Office settings
        (HKCU only, previous values logged to logs/office-settings-<time>.json). Nothing is changed without
        that click; admin policies that get in the way are reported, never touched.
      - Drop folder: creates "Sagamore Drop" on the desktop. Files dragged into it are offered in the page
        for upload to a Teams channel (scripts/Send-SagamoreDrop.ps1), after the user confirms that they will
        take the channel folder's permissions.

.EXAMPLE
    ./scripts/Start-SagamoreDashboard.ps1
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..' 'config' 'sagamore.config.json'),
    [ValidateRange(1024, 65535)][int]$Port = 8765,
    [string]$LogDirectory = (Join-Path $PSScriptRoot '..' 'logs'),
    # Defaults to "Sagamore Drop" on the current user's desktop.
    [string]$DropFolder,
    # Do not start OneDrive automatically (the check is still shown).
    [switch]$NoOneDriveStart,
    [switch]$NoBrowser,
    # For tests only. Normally a fresh random token is generated each start.
    [string]$Token
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Common.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Desktop.psm1') -Force

$ConfigPath   = [IO.Path]::GetFullPath($ConfigPath)
$LogDirectory = [IO.Path]::GetFullPath($LogDirectory)
if (-not $DropFolder) { $DropFolder = Get-SagamoreDefaultDropFolder }
$DropFolder   = Initialize-SagamoreDropFolder -Path $DropFolder
$dashboardDir = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..' 'dashboard'))
$backupDir    = Join-Path $LogDirectory 'config-backups'
$pwshPath     = (Get-Process -Id $PID).Path
if (-not $Token) { $Token = [Convert]::ToHexString([Security.Cryptography.RandomNumberGenerator]::GetBytes(24)).ToLower() }
New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null

# The only commands the page can start. Kind groups a Preview with the Apply it unlocks.
$Actions = [ordered]@{
    plan        = @{ Script = 'Install-Sagamore.ps1';     Args = @('-PlanOnly');                    Label = 'Plan (offline)';             Kind = $null;   Role = 'check'   }
    preview     = @{ Script = 'Install-Sagamore.ps1';     Args = @('-PermissionsOnly', '-WhatIf'); Label = 'Preview permission changes'; Kind = 'perms'; Role = 'preview' }
    apply       = @{ Script = 'Install-Sagamore.ps1';     Args = @('-PermissionsOnly');            Label = 'Apply permission changes';   Kind = 'perms'; Role = 'apply'   }
    previewFull = @{ Script = 'Install-Sagamore.ps1';     Args = @('-WhatIf');                     Label = 'Preview full install';       Kind = 'full';  Role = 'preview' }
    installFull = @{ Script = 'Install-Sagamore.ps1';     Args = @();                              Label = 'Full install';               Kind = 'full';  Role = 'apply'   }
    audit       = @{ Script = 'Test-SagamoreInstall.ps1'; Args = @();                              Label = 'Audit';                      Kind = $null;   Role = 'check'   }
    # Only started through POST /api/drop/send (which validates the request and writes the request file);
    # POST /api/run refuses Role 'drop'.
    drop        = @{ Script = 'Send-SagamoreDrop.ps1';    Args = @();                              Label = 'Send files to Teams';        Kind = $null;   Role = 'drop'    }
}
$Static = @{
    '/dashboard/'           = @{ File = 'index.html'; Type = 'text/html; charset=utf-8' }
    '/dashboard/index.html' = @{ File = 'index.html'; Type = 'text/html; charset=utf-8' }
    '/dashboard/plan.js'    = @{ File = 'plan.js';    Type = 'text/javascript; charset=utf-8' }
}
$AllowedHosts   = @("localhost:$Port", "127.0.0.1:$Port")
$AllowedOrigins = @("http://localhost:$Port", "http://127.0.0.1:$Port")
$MaxBody        = 2MB
$Utf8           = [Text.UTF8Encoding]::new($false)

$script:Run      = $null    # the current or most recent run
$script:RunSeq   = 0
$script:Previewed = @{}     # Kind -> config hash that last passed a Preview
$script:OneDriveStartedAt = [datetime]::MinValue   # last time this server started OneDrive (restart throttle)
$OneDriveRestartEvery = [timespan]::FromMinutes(5)

# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
function Get-ConfigHash {
    if (-not (Test-Path $ConfigPath)) { return $null }
    (Get-FileHash -Path $ConfigPath -Algorithm SHA256).Hash
}

function Send-Bytes($Ctx, [int]$Status, [byte[]]$Bytes, [string]$ContentType) {
    $r = $Ctx.Response
    $r.StatusCode = $Status
    $r.ContentType = $ContentType
    $r.AddHeader('Cache-Control', 'no-store')
    $r.AddHeader('X-Content-Type-Options', 'nosniff')
    $r.AddHeader('Referrer-Policy', 'no-referrer')
    $r.ContentLength64 = $Bytes.Length
    $r.OutputStream.Write($Bytes, 0, $Bytes.Length)
    $r.Close()
}

function Send-Json($Ctx, [int]$Status, $Object) {
    Send-Bytes $Ctx $Status ($Utf8.GetBytes((ConvertTo-Json -InputObject $Object -Depth 10 -Compress))) 'application/json; charset=utf-8'
}

function Read-Body($Ctx) {
    $req = $Ctx.Request
    if ($req.ContentLength64 -gt $MaxBody) { throw [InvalidOperationException]::new('Request too large.') }
    $reader = [IO.StreamReader]::new($req.InputStream, $Utf8)
    try { $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Read-RunOutput([string]$File, [long]$From) {
    if (-not (Test-Path $File)) { return @{ Text = ''; Next = $From } }
    $fs = [IO.File]::Open($File, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        if ($From -gt $fs.Length) { $From = 0 }
        $null = $fs.Seek($From, [IO.SeekOrigin]::Begin)
        $buf = [byte[]]::new($fs.Length - $From)
        $n = $fs.Read($buf, 0, $buf.Length)
        return @{ Text = $Utf8.GetString($buf, 0, $n); Next = $From + $n }
    } finally { $fs.Dispose() }
}

function Update-RunState {
    $r = $script:Run
    if (-not $r -or $r.Done) { return }
    if (-not $r.Process.HasExited) { return }
    $r.Process.WaitForExit()
    $r.ExitCode = $r.Process.ExitCode
    $r.Done = $true
    $def = $Actions[$r.Action]
    if ($def.Role -eq 'preview' -and $r.ExitCode -eq 0 -and $r.ConfigHash -eq (Get-ConfigHash)) {
        $script:Previewed[$def.Kind] = $r.ConfigHash
    }
    if ($def.Role -eq 'apply') { $script:Previewed.Remove($def.Kind) }
}

function Get-RunStatus([long]$From = 0) {
    Update-RunState
    $hash = Get-ConfigHash
    $ready = @{}
    foreach ($k in 'perms', 'full') { $ready[$k] = [bool]($hash -and $script:Previewed[$k] -eq $hash) }
    $r = $script:Run
    if (-not $r) { return @{ id = $null; applyReady = $ready } }
    $out = Read-RunOutput $r.OutFile $From
    @{
        id         = $r.Id
        action     = $r.Action
        label      = $Actions[$r.Action].Label
        running    = -not $r.Done
        exitCode   = $r.ExitCode
        started    = $r.Started.ToString('o')
        output     = $out.Text
        next       = $out.Next
        logFile    = $r.OutFile
        applyReady = $ready
    }
}

function Start-Run([string]$Action, [System.Collections.IDictionary]$Named = @{}) {
    # $Named (parameter name -> value) comes only from this server (e.g. the request file path it just wrote),
    # never from the page. Names are checked and values single-quoted so a path with spaces or quotes cannot
    # break the command line.
    $def = $Actions[$Action]
    $hash = Get-ConfigHash
    if (-not $hash) { throw [IO.FileNotFoundException]::new("No config at $ConfigPath.") }
    if ($def.Role -eq 'apply' -and $script:Previewed[$def.Kind] -ne $hash) {
        return @{ error = 'Run the Preview for this exact config first. The config has changed since the last successful Preview, or no Preview has run.' }
    }
    $script:RunSeq++
    $outFile = Join-Path $LogDirectory ('dashboard-{0:yyyyMMdd-HHmmss}-{1}.txt' -f (Get-Date), $Action)
    $q = { param($s) "'" + ($s -replace "'", "''") + "'" }
    $scriptPath = Join-Path $PSScriptRoot $def.Script
    $extra = @($Named.Keys | ForEach-Object {
        if ($_ -notmatch '^[A-Za-z]+$') { throw "Bad parameter name '$_'." }
        "-$_ $(& $q ([string]$Named[$_]))"
    }) -join ' '
    $cmd = @(
        '$PSStyle.OutputRendering = ''PlainText'''
        '$global:LASTEXITCODE = 0'
        "try { & $(& $q $scriptPath) -ConfigPath $(& $q $ConfigPath) $($def.Args -join ' ') $extra *>&1 | Out-String -Stream -Width 300 }"
        'catch { Write-Output ($_ | Out-String); exit 1 }'
        'exit $LASTEXITCODE'
    ) -join "`n"
    $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($cmd))
    $p = Start-Process -FilePath $pwshPath -ArgumentList '-NoProfile', '-NonInteractive', '-EncodedCommand', $enc `
        -RedirectStandardOutput $outFile -NoNewWindow -PassThru
    $null = $p.Handle   # keeps ExitCode available after the process ends
    $script:Run = @{ Id = $script:RunSeq; Action = $Action; Process = $p; OutFile = $outFile; Started = Get-Date; Done = $false; ExitCode = $null; ConfigHash = $hash }
    Write-Host "  > $($def.Label) started (log: $outFile)" -ForegroundColor Cyan
    return $null
}

function Test-Running { Update-RunState; [bool]($script:Run -and -not $script:Run.Done) }

function Read-ConfigOrNull {
    <# The saved config if it exists and is valid, else $null plus a message. Desktop features need a valid one. #>
    if (-not (Test-Path $ConfigPath)) { return @{ Config = $null; Error = "No config at $ConfigPath yet." } }
    try { $c = Read-SagamoreConfig -Path $ConfigPath } catch { return @{ Config = $null; Error = $_.Exception.Message } }
    $errs = @(Test-SagamoreConfig -Config $c -AllowPlaceholderClientId)
    if ($errs.Count) { return @{ Config = $null; Error = 'The saved config has problems; fix and save it first.' } }
    return @{ Config = $c; Error = $null }
}

function Get-OneDriveStatus([switch]$AllowAutoStart) {
    <#
        Current OneDrive health for the page. With -AllowAutoStart (start-up and the page's periodic check),
        an installed-but-stopped OneDrive is started, at most once per $OneDriveRestartEvery.
    #>
    $r = Read-ConfigOrNull
    $facts = Get-SagamoreOneDriveFacts
    $autoStarted = $false
    if ($AllowAutoStart -and -not $NoOneDriveStart -and $facts.IsWindows -and $facts.ExePath -and -not $facts.Running -and
        ((Get-Date) - $script:OneDriveStartedAt) -gt $OneDriveRestartEvery) {
        $autoStarted = Start-SagamoreOneDrive -Facts $facts
        if ($autoStarted) {
            $script:OneDriveStartedAt = Get-Date
            Write-Host '  > OneDrive was not running; started it.' -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            $facts = Get-SagamoreOneDriveFacts
        }
    }
    if (-not $r.Config) {
        return @{ overall = 'unknown'; checks = @(); error = $r.Error; autoStarted = $autoStarted; libraryUrl = $null; syncedPath = $null }
    }
    $s = Test-SagamoreOneDriveState -Facts $facts -Config $r.Config
    $libUrl = (Get-SagamoreSiteUrl -Config $r.Config) + '/' + ([string]$r.Config.site.library -replace ' ', '%20')
    return @{
        overall     = $s.Overall
        checks      = @($s.Checks | ForEach-Object { @{ id = $_.Id; label = $_.Label; status = $_.Status; detail = $_.Detail; fix = $_.Fix } })
        syncedPath  = $s.SyncedPath
        libraryUrl  = $libUrl
        canStart    = [bool]($facts.IsWindows -and $facts.ExePath -and -not $facts.Running)
        autoStarted = $autoStarted
        error       = $null
    }
}

function Get-DropStatus {
    <# What is waiting in the drop folder, plus where it may go (from the SAVED config, which is what the uploader uses). #>
    $r = Read-ConfigOrNull
    $items = @(Get-SagamoreDropItems -Path $DropFolder)
    $dests = @()
    if ($r.Config) {
        $dests = @(Get-SagamoreDropDestinations -Config $r.Config | ForEach-Object {
            @{ channel = $_.Channel; type = $_.Type; root = $_.Root; folders = @($_.Folders) }
        })
    }
    @{
        folder       = $DropFolder
        items        = @($items | ForEach-Object { @{ name = $_.Name; kind = $_.Kind; size = $_.Size; modified = $_.Modified; sendable = $_.Sendable; reason = $_.Reason } })
        destinations = $dests
        error        = $r.Error
    }
}

function Get-OfficeSaveStatus {
    <#
        Office save-settings status for the page. Needs the OneDrive status for the synced library folder,
        because "save to Sagamore by default" means "save into that folder". Never changes anything.
    #>
    $od = Get-OneDriveStatus
    $facts = Get-SagamoreOfficeSaveFacts
    $s = Test-SagamoreOfficeSaveState -Facts $facts -SyncedPath $od.syncedPath
    return @{
        overall    = $s.Overall
        checks     = @($s.Checks | ForEach-Object { @{ id = $_.Id; label = $_.Label; status = $_.Status; detail = $_.Detail; fix = $_.Fix } })
        changes    = @($s.Changes | ForEach-Object { $_.Why })
        canFix     = [bool]@($s.Changes).Count
        syncedPath = $od.syncedPath
        error      = $od.error
    }
}

function Invoke-OfficeSaveFix {
    <#
        Applies the user-level Office changes (Set-SagamoreOfficeSaveDefaults) after logging the old values,
        so every change can be put back by hand. Recomputes the changes here rather than trusting the page.
    #>
    $od = Get-OneDriveStatus
    $facts = Get-SagamoreOfficeSaveFacts
    $changes = @(Get-SagamoreOfficeSaveChanges -Facts $facts -SyncedPath $od.syncedPath)
    if (-not $changes.Count) { return @{ applied = @(); log = $null } }
    $log = Join-Path $LogDirectory ('office-settings-{0:yyyyMMdd-HHmmss}.json' -f (Get-Date))
    $changes | ForEach-Object { @{ path = $_.Path; name = $_.Name; old = $_.Old; new = $(if ($_.Remove) { '(removed)' } else { $_.Value }); why = $_.Why } } |
        ConvertTo-Json -Depth 4 -AsArray | Set-Content -Path $log -Encoding utf8
    $done = @(Set-SagamoreOfficeSaveDefaults -Changes $changes)
    foreach ($c in $done) { Write-Host "  > Office: $($c.Why)" -ForegroundColor Yellow }
    return @{ applied = @($done | ForEach-Object { $_.Why }); log = $log }
}

# ------------------------------------------------------------------
# Request handling
# ------------------------------------------------------------------
function Invoke-Request($Ctx) {
    $req = $Ctx.Request
    $path = $req.Url.AbsolutePath
    $method = $req.HttpMethod

    if ($req.Headers['Host'] -notin $AllowedHosts) { return Send-Json $Ctx 403 @{ error = 'Bad host.' } }

    if ($method -eq 'GET' -and $path -eq '/') {
        $Ctx.Response.Redirect('/dashboard/'); $Ctx.Response.Close(); return
    }
    if ($method -eq 'GET' -and $Static.ContainsKey($path)) {
        $s = $Static[$path]
        return Send-Bytes $Ctx 200 ([IO.File]::ReadAllBytes((Join-Path $dashboardDir $s.File))) $s.Type
    }
    if (-not $path.StartsWith('/api/')) { return Send-Json $Ctx 404 @{ error = 'Not found.' } }

    # --- API: token + origin required ---
    $origin = $req.Headers['Origin']
    if ($origin -and $origin -notin $AllowedOrigins) { return Send-Json $Ctx 403 @{ error = 'Bad origin.' } }
    if ($req.Headers['X-Sagamore-Token'] -cne $Token) { return Send-Json $Ctx 403 @{ error = 'Missing or wrong token. Re-open the dashboard from the link printed by Start-SagamoreDashboard.ps1.' } }

    switch ("$method $path") {
        'GET /api/info' {
            return Send-Json $Ctx 200 @{
                configPath   = $ConfigPath
                configExists = (Test-Path $ConfigPath)
                actions      = @($Actions.Keys | ForEach-Object { @{ id = $_; label = $Actions[$_].Label; role = $Actions[$_].Role; kind = $Actions[$_].Kind } })
            }
        }
        'GET /api/config' {
            if (-not (Test-Path $ConfigPath)) {
                return Send-Json $Ctx 404 @{ error = "No config at $ConfigPath. Copy config/sagamore.config.example.json to config/sagamore.config.json first." }
            }
            return Send-Bytes $Ctx 200 ([IO.File]::ReadAllBytes($ConfigPath)) 'application/json; charset=utf-8'
        }
        'PUT /api/config' {
            if (Test-Running) { return Send-Json $Ctx 409 @{ error = 'A run is in progress. Save again when it finishes.' } }
            $body = Read-Body $Ctx
            try { $cfg = $body | ConvertFrom-Json -AsHashtable -Depth 20 } catch { return Send-Json $Ctx 400 @{ error = 'Not valid JSON.'; errors = @() } }
            if ($cfg -isnot [hashtable]) { return Send-Json $Ctx 400 @{ error = 'Not a Sagamore config.'; errors = @() } }
            $errs = @(Test-SagamoreConfig -Config $cfg -AllowPlaceholderClientId)
            if ($errs.Count) { return Send-Json $Ctx 400 @{ error = 'Config has problems; not saved.'; errors = $errs } }
            $backup = $null
            if (Test-Path $ConfigPath) {
                New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
                $backup = Join-Path $backupDir ('sagamore.config.{0:yyyyMMdd-HHmmss-fff}.json' -f (Get-Date))
                Copy-Item -Path $ConfigPath -Destination $backup
            }
            [IO.File]::WriteAllText($ConfigPath, $body, $Utf8)
            Write-Host "  > Config saved (backup: $backup)" -ForegroundColor Green
            return Send-Json $Ctx 200 @{ saved = $true; backup = $backup; status = (Get-RunStatus ([long]::MaxValue)) }
        }
        'GET /api/run' {
            $from = 0L
            [void][long]::TryParse([string]$req.QueryString['from'], [ref]$from)
            return Send-Json $Ctx 200 (Get-RunStatus $from)
        }
        'POST /api/run' {
            try { $action = [string]((Read-Body $Ctx | ConvertFrom-Json -AsHashtable).action) } catch { $action = '' }
            if (-not $action -or -not $Actions.Contains($action)) { return Send-Json $Ctx 400 @{ error = "Unknown action '$action'." } }
            if ($Actions[$action].Role -eq 'drop') { return Send-Json $Ctx 400 @{ error = 'Send files with the drop folder dialog.' } }
            if (Test-Running) { return Send-Json $Ctx 409 @{ error = 'Another run is still in progress.' } }
            $refused = Start-Run $action
            if ($refused) { return Send-Json $Ctx 409 $refused }
            return Send-Json $Ctx 202 (Get-RunStatus 0)
        }
        'POST /api/run/stop' {
            if (-not (Test-Running)) { return Send-Json $Ctx 409 @{ error = 'Nothing is running.' } }
            $script:Run.Process.Kill($true)
            Write-Host '  > Run stopped from the dashboard.' -ForegroundColor Yellow
            return Send-Json $Ctx 200 @{ stopped = $true }
        }

        # --- OneDrive ---
        'GET /api/onedrive' {
            # The page calls this on load and every minute; that is what keeps OneDrive running.
            return Send-Json $Ctx 200 (Get-OneDriveStatus -AllowAutoStart)
        }
        'POST /api/onedrive/start' {
            $facts = Get-SagamoreOneDriveFacts
            if (-not $facts.IsWindows) { return Send-Json $Ctx 409 @{ error = 'OneDrive can only be started on Windows.' } }
            if (-not $facts.ExePath) { return Send-Json $Ctx 409 @{ error = 'OneDrive is not installed for this Windows user.' } }
            if (Start-SagamoreOneDrive -Facts $facts) { $script:OneDriveStartedAt = Get-Date; Start-Sleep -Seconds 2 }
            return Send-Json $Ctx 200 (Get-OneDriveStatus)
        }

        # --- Drop folder ---
        # --- Office save settings ---
        'GET /api/officesave' { return Send-Json $Ctx 200 (Get-OfficeSaveStatus) }
        'POST /api/officesave/fix' {
            if (-not $IsWindows) { return Send-Json $Ctx 409 @{ error = 'Office settings can only be changed on Windows.' } }
            $r = Invoke-OfficeSaveFix
            $status = Get-OfficeSaveStatus
            $status.applied = $r.applied
            $status.log = $r.log
            return Send-Json $Ctx 200 $status
        }
        'GET /api/drop' { return Send-Json $Ctx 200 (Get-DropStatus) }
        'POST /api/drop/open' {
            # Opens the drop folder in File Explorer so the user can drag files in.
            if ($IsWindows) { Invoke-Item -LiteralPath $DropFolder }
            return Send-Json $Ctx 200 @{ opened = [bool]$IsWindows; folder = $DropFolder }
        }
        'POST /api/drop/send' {
            if (Test-Running) { return Send-Json $Ctx 409 @{ error = 'Another run is still in progress. Send again when it finishes.' } }
            try { $reqBody = Read-Body $Ctx | ConvertFrom-Json -AsHashtable } catch { return Send-Json $Ctx 400 @{ error = 'Not valid JSON.'; errors = @() } }
            if ($reqBody -isnot [hashtable]) { return Send-Json $Ctx 400 @{ error = 'Bad request.'; errors = @() } }
            $r = Read-ConfigOrNull
            if (-not $r.Config) { return Send-Json $Ctx 409 @{ error = $r.Error; errors = @() } }
            # Keep only the fields the uploader understands; anything else in the body is dropped.
            $clean = [ordered]@{
                destination  = [string](Get-OrDefault $reqBody 'destination' '')
                files        = @(@(Get-OrDefault $reqBody 'files' @()) | ForEach-Object { [string]$_ })
                # Booleans only: a string such as "true" or a number is NOT accepted as yes.
                overwrite    = $(($v = Get-OrDefault $reqBody 'overwrite' $false) -is [bool] -and $v)
                acknowledged = $(($v = Get-OrDefault $reqBody 'acknowledged' $false) -is [bool] -and $v)
                requestedAt  = (Get-Date).ToString('o')
            }
            $check = Test-SagamoreDropRequest -Config $r.Config -Request $clean -Items @(Get-SagamoreDropItems -Path $DropFolder)
            if ($check.Errors.Count) { return Send-Json $Ctx 400 @{ error = 'Not sent.'; errors = $check.Errors } }
            $reqFile = Join-Path $LogDirectory ('drop-request-{0:yyyyMMdd-HHmmss-fff}.json' -f (Get-Date))
            [IO.File]::WriteAllText($reqFile, ($clean | ConvertTo-Json -Depth 5), $Utf8)
            $refused = Start-Run 'drop' ([ordered]@{ RequestPath = $reqFile; DropFolder = $DropFolder })
            if ($refused) { return Send-Json $Ctx 409 $refused }
            return Send-Json $Ctx 202 (Get-RunStatus 0)
        }
        default { return Send-Json $Ctx 404 @{ error = 'Not found.' } }
    }
}

# ------------------------------------------------------------------
# Main loop
# ------------------------------------------------------------------
$listener = [Net.HttpListener]::new()
$listener.Prefixes.Add("http://localhost:$Port/")
try { $listener.Start() } catch {
    Write-Host "Could not listen on port $Port ($($_.Exception.Message)). Is the dashboard already running? Try -Port 8766." -ForegroundColor Red
    exit 1
}
$url = "http://localhost:$Port/dashboard/#token=$Token"
Write-Host ''
Write-Host 'Sagamore Permission Manager is running on this computer only.' -ForegroundColor Green
Write-Host "  Config: $ConfigPath"
Write-Host "  Drop:   $DropFolder  (drag files here to send them to a Teams channel)"
Write-Host "  Open:   $url"
Write-Host '  Press Ctrl+C to stop.'
Write-Host ''

# Start-up OneDrive check (also starts OneDrive if it is installed but not running).
try {
    $od = Get-OneDriveStatus -AllowAutoStart
    if ($od.error) { Write-Host "OneDrive check skipped: $($od.error)" -ForegroundColor Yellow }
    foreach ($c in $od.checks) {
        $color = @{ ok = 'Green'; warn = 'Yellow'; fail = 'Red' }[$c.status]
        Write-Host ('  OneDrive: [{0,-4}] {1} -- {2}' -f $c.status.ToUpper(), $c.label, $c.detail) -ForegroundColor $color
        if ($c.fix) { Write-Host "            $($c.fix)" -ForegroundColor $color }
    }
    Write-Host ''
} catch { Write-Host "OneDrive check failed: $($_.Exception.Message)" -ForegroundColor Yellow }
# Start-up Office save-settings check (report only; the fix is a button in the page).
try {
    $os = Get-OfficeSaveStatus
    foreach ($c in $os.checks) {
        $color = @{ ok = 'Green'; warn = 'Yellow'; fail = 'Red' }[$c.status]
        Write-Host ('  Office:   [{0,-4}] {1} -- {2}' -f $c.status.ToUpper(), $c.label, $c.detail) -ForegroundColor $color
        if ($c.fix) { Write-Host "            $($c.fix)" -ForegroundColor $color }
    }
    Write-Host ''
} catch { Write-Host "Office check failed: $($_.Exception.Message)" -ForegroundColor Yellow }
if (-not $NoBrowser) { Start-Process $url }

try {
    while ($listener.IsListening) {
        $task = $listener.GetContextAsync()
        while (-not $task.AsyncWaitHandle.WaitOne(250)) { }   # short waits so Ctrl+C is honoured
        $ctx = $task.GetAwaiter().GetResult()
        try { Invoke-Request $ctx }
        catch {
            Write-Host "  ! $($_.Exception.Message)" -ForegroundColor Red
            try { Send-Json $ctx 500 @{ error = $_.Exception.Message } } catch { Write-Verbose 'Client went away.' }
        }
    }
} finally {
    $listener.Stop(); $listener.Close()
    if ($script:Run -and -not $script:Run.Done -and -not $script:Run.Process.HasExited) {
        Write-Host "A run is still going in the background (log: $($script:Run.OutFile)). It will finish on its own." -ForegroundColor Yellow
    }
}
