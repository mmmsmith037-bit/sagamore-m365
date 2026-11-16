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

.EXAMPLE
    ./scripts/Start-SagamoreDashboard.ps1
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..' 'config' 'sagamore.config.json'),
    [ValidateRange(1024, 65535)][int]$Port = 8765,
    [string]$LogDirectory = (Join-Path $PSScriptRoot '..' 'logs'),
    [switch]$NoBrowser,
    # For tests only. Normally a fresh random token is generated each start.
    [string]$Token
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Common.psm1') -Force

$ConfigPath   = [IO.Path]::GetFullPath($ConfigPath)
$LogDirectory = [IO.Path]::GetFullPath($LogDirectory)
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

function Start-Run([string]$Action) {
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
    $cmd = @(
        '$PSStyle.OutputRendering = ''PlainText'''
        '$global:LASTEXITCODE = 0'
        "try { & $(& $q $scriptPath) -ConfigPath $(& $q $ConfigPath) $($def.Args -join ' ') *>&1 | Out-String -Stream -Width 300 }"
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
Write-Host "  Open:   $url"
Write-Host '  Press Ctrl+C to stop.'
Write-Host ''
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
