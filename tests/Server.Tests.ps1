# End-to-end tests for the local dashboard server (scripts/Start-SagamoreDashboard.ps1).
# Starts the real server on a random port against a temp copy of the example config. No tenant needed:
# the only command actually run is the offline -PlanOnly plan.
BeforeAll {
    $root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ('sagamore-srv-' + [guid]::NewGuid().ToString('n'))
    New-Item -ItemType Directory -Path $Tmp | Out-Null
    $script:Cfg = Join-Path $Tmp 'sagamore.config.json'
    Copy-Item (Join-Path $root 'config' 'sagamore.config.example.json') $Cfg
    $script:Logs = Join-Path $Tmp 'logs'
    $script:Port = Get-Random -Minimum 20000 -Maximum 40000
    $script:Tok = 'test-' + [guid]::NewGuid().ToString('n')
    $script:Base = "http://localhost:$Port"
    $script:Drop = Join-Path $Tmp 'Sagamore Drop'
    $pwsh = (Get-Process -Id $PID).Path
    $srvArgs = @{
        FilePath     = $pwsh
        PassThru     = $true
        ArgumentList = @('-NoProfile', '-File', (Join-Path $root 'scripts' 'Start-SagamoreDashboard.ps1'),
            '-ConfigPath', $Cfg, '-Port', $Port, '-LogDirectory', $Logs, '-DropFolder', "`"$Drop`"", '-NoOneDriveStart', '-NoBrowser', '-Token', $Tok)
    }
    # -WindowStyle only exists on Windows; elsewhere the server runs without a window anyway.
    if ($IsWindows) { $srvArgs.WindowStyle = 'Hidden' } else { $srvArgs.RedirectStandardOutput = (Join-Path $Tmp 'server.out') }
    $script:Srv = Start-Process @srvArgs
    $deadline = (Get-Date).AddSeconds(30)
    do {
        Start-Sleep -Milliseconds 300
        try { $null = Invoke-WebRequest "$Base/dashboard/" -TimeoutSec 2; $up = $true } catch { $up = $false }
    } until ($up -or (Get-Date) -gt $deadline -or $Srv.HasExited)
    if (-not $up) { throw 'Dashboard server did not start.' }

    function script:Api([string]$Method, [string]$Path, $Body, [hashtable]$Headers = @{ 'X-Sagamore-Token' = $Tok }) {
        $p = @{ Uri = "$Base$Path"; Method = $Method; Headers = $Headers; SkipHttpErrorCheck = $true; TimeoutSec = 10 }
        if ($null -ne $Body) { $p.Body = $Body; $p.ContentType = 'application/json' }
        $r = Invoke-WebRequest @p
        $data = $null; try { $data = $r.Content | ConvertFrom-Json -AsHashtable } catch { $data = $null }
        [pscustomobject]@{ Status = [int]$r.StatusCode; Data = $data; Text = $r.Content }
    }
    function script:Wait-Run {
        $deadline = (Get-Date).AddSeconds(90)
        do { Start-Sleep -Milliseconds 400; $s = (Api GET '/api/run').Data } while ($s.running -and (Get-Date) -lt $deadline)
        $s
    }
}

AfterAll {
    if ($Srv -and -not $Srv.HasExited) { $Srv.Kill($true) }
    Remove-Item $Tmp -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'Dashboard server' {
    It 'serves the dashboard and plan.js without a token' {
        (Invoke-WebRequest "$Base/dashboard/").Content | Should -Match 'Permission Manager'
        (Invoke-WebRequest "$Base/dashboard/plan.js").Content | Should -Match 'SagamorePlan'
    }

    It 'serves nothing outside the dashboard whitelist' {
        (Api GET '/config/sagamore.config.example.json' $null @{}).Status | Should -Be 404
        (Api GET '/dashboard/../config/sagamore.config.json' $null @{}).Status | Should -Not -Be 200
    }

    It 'refuses API calls without the token or with a wrong token' {
        (Api GET '/api/config' $null @{}).Status | Should -Be 403
        (Api GET '/api/config' $null @{ 'X-Sagamore-Token' = 'nope' }).Status | Should -Be 403
    }

    It 'refuses API calls from another origin even with the token' {
        (Api GET '/api/config' $null @{ 'X-Sagamore-Token' = $Tok; Origin = 'https://evil.example' }).Status | Should -Be 403
    }

    It 'returns the config file' {
        $r = Api GET '/api/config'
        $r.Status | Should -Be 200
        $r.Data.site.alias | Should -Be 'SagamoreAlliance'
    }

    It 'rejects an invalid config and leaves the file untouched' {
        $before = Get-FileHash $Cfg
        $bad = Get-Content $Cfg -Raw | ConvertFrom-Json -AsHashtable
        $bad.clients[0].access['Member'] = 'Read'          # whole group on a client silo
        $r = Api PUT '/api/config' ($bad | ConvertTo-Json -Depth 20)
        $r.Status | Should -Be 400
        ($r.Data.errors -join "`n") | Should -Match 'Meridian'
        (Get-FileHash $Cfg).Hash | Should -Be $before.Hash
        (Api PUT '/api/config' 'not json').Status | Should -Be 400
    }

    It 'saves a valid config and backs up the previous one' {
        $c = Get-Content $Cfg -Raw | ConvertFrom-Json -AsHashtable
        $c.clients[1].access['support1@sagamorealliance.com'] = 'Read'
        $text = ($c | ConvertTo-Json -Depth 20) + "`n"
        $r = Api PUT '/api/config' $text
        $r.Status | Should -Be 200
        $r.Data.backup | Should -Exist
        Get-Content $Cfg -Raw | Should -Be $text
    }

    It 'rejects unknown actions' {
        (Api POST '/api/run' '{"action":"rm -rf"}').Status | Should -Be 400
        (Api POST '/api/run' '{"action":"apply; whoami"}').Status | Should -Be 400
    }

    It 'refuses Apply until a Preview of the same config has succeeded' {
        $r = Api POST '/api/run' '{"action":"apply"}'
        $r.Status | Should -Be 409
        $r.Data.error | Should -Match 'Preview'
        (Api POST '/api/run' '{"action":"installFull"}').Status | Should -Be 409
        (Api GET '/api/run').Data.applyReady.perms | Should -BeFalse
    }

    It 'runs the offline plan and streams its output' {
        $r = Api POST '/api/run' '{"action":"plan"}'
        $r.Status | Should -Be 202
        $s = Wait-Run
        $s.running | Should -BeFalse
        $s.exitCode | Should -Be 0
        $s.output | Should -Match 'Plan only -- nothing was changed'
        $s.output | Should -Match 'Hartwell-Inc .*support1@sagamorealliance.com'
        $s.logFile | Should -Exist
    }

    It 'returns only new output when asked from an offset' {
        $all = (Api GET '/api/run').Data
        $tail = (Api GET "/api/run?from=$($all.next)").Data
        $tail.output | Should -BeNullOrEmpty
        $tail.next | Should -Be $all.next
    }
}

Describe 'Dashboard server: drop folder and OneDrive' {
    It 'creates the drop folder with its _Sent archive on start-up' {
        Join-Path $Drop '_Sent' | Should -Exist
        Join-Path $Drop '_Sent' 'README.txt' | Should -Exist
    }

    It 'lists waiting files, explains folders, and offers only channel destinations' {
        'pdf' | Set-Content (Join-Path $Drop 'Invoice 104.pdf')
        New-Item -ItemType Directory (Join-Path $Drop 'Invoices') -Force | Out-Null
        $d = (Api GET '/api/drop').Data
        $d.folder | Should -Be ([IO.Path]::GetFullPath($Drop))
        ($d.items | Where-Object name -eq 'Invoice 104.pdf').sendable | Should -BeTrue
        ($d.items | Where-Object name -eq 'Invoices').sendable | Should -BeFalse
        @($d.items | Where-Object name -eq '_Sent').Count | Should -Be 0
        @($d.destinations | ForEach-Object channel) | Should -Be @('Internal', 'Meridian-Group', 'Hartwell-Inc')
        ($d.destinations | Where-Object channel -eq 'Meridian-Group').folders | Should -Contain 'Clients/Meridian-Group/Invoices'
        @($d.destinations.folders) | Should -Not -Contain '_Admin'
    }

    It 'refuses a send that is not acknowledged, goes outside the channels, or names unknown files' {
        $r = Api POST '/api/drop/send' (@{ destination = 'Clients/Meridian-Group'; files = @('Invoice 104.pdf'); acknowledged = $false } | ConvertTo-Json)
        $r.Status | Should -Be 400
        ($r.Data.errors -join "`n") | Should -Match 'permissions'

        $r = Api POST '/api/drop/send' (@{ destination = 'Clients/Meridian-Group'; files = @('Invoice 104.pdf'); acknowledged = 'true' } | ConvertTo-Json)
        $r.Status | Should -Be 400

        $r = Api POST '/api/drop/send' (@{ destination = '_Admin'; files = @('Invoice 104.pdf'); acknowledged = $true } | ConvertTo-Json)
        $r.Status | Should -Be 400
        ($r.Data.errors -join "`n") | Should -Match 'not a Teams channel folder'

        $r = Api POST '/api/drop/send' (@{ destination = 'Clients/Meridian-Group'; files = @('../sagamore.config.json', 'Invoices', 'nope.pdf'); acknowledged = $true } | ConvertTo-Json)
        $r.Status | Should -Be 400
        $e = $r.Data.errors -join "`n"
        $e | Should -Match 'not a file name'
        $e | Should -Match 'Folders are not uploaded'
        $e | Should -Match 'not in the drop folder'
        @(Get-ChildItem $Logs -Filter 'drop-request-*' -ErrorAction SilentlyContinue).Count | Should -Be 0
    }

    It 'does not let the generic run endpoint start an upload' {
        (Api POST '/api/run' '{"action":"drop"}').Status | Should -Be 400
    }

    It 'starts the uploader for a valid, acknowledged request and records exactly what was asked' {
        $r = Api POST '/api/drop/send' (@{ destination = 'clients/meridian-group/INVOICES'; files = @('Invoice 104.pdf'); overwrite = $true; acknowledged = $true; extra = 'ignored' } | ConvertTo-Json)
        $r.Status | Should -Be 202
        $r.Data.action | Should -Be 'drop'
        $req = Get-ChildItem $Logs -Filter 'drop-request-*.json' | Select-Object -First 1
        $req | Should -Not -BeNullOrEmpty
        $saved = Get-Content $req.FullName -Raw | ConvertFrom-Json -AsHashtable
        $saved.files | Should -Be @('Invoice 104.pdf')
        $saved.overwrite | Should -BeTrue
        $saved.acknowledged | Should -BeTrue
        $saved.ContainsKey('extra') | Should -BeFalse
        # The example config has a placeholder app id, so the real uploader stops at validation (exit 2)
        # before any sign-in. That proves the script was started with the request and drop folder.
        $s = Wait-Run
        $s.exitCode | Should -Be 2
        $s.output | Should -Match 'clientId'
        Join-Path $Drop 'Invoice 104.pdf' | Should -Exist
    }

    It 'reports OneDrive status without failing on any platform' {
        $o = (Api GET '/api/onedrive').Data
        $o.overall | Should -BeIn @('ok', 'warn', 'fail', 'unsupported')
        $o.libraryUrl | Should -Be 'https://sagamorealliance.sharepoint.com/sites/SagamoreAlliance/Shared%20Documents'
        $o.autoStarted | Should -BeFalse     # -NoOneDriveStart
    }

    It 'reports Office save settings without changing anything, and only changes them on Windows' {
        $o = (Api GET '/api/officesave').Data
        $o.overall | Should -BeIn @('ok', 'warn', 'fail', 'unsupported')
        @($o.checks).Count | Should -BeGreaterThan 0
        if (-not $IsWindows) {
            (Api POST '/api/officesave/fix' '{}').Status | Should -Be 409
        }
        @(Get-ChildItem $Logs -Filter 'office-settings-*' -ErrorAction SilentlyContinue).Count | Should -Be 0
    }
}
