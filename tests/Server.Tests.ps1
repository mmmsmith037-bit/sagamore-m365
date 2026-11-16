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
    $pwsh = (Get-Process -Id $PID).Path
    $script:Srv = Start-Process $pwsh -PassThru -WindowStyle Hidden -ArgumentList @(
        '-NoProfile', '-File', (Join-Path $root 'scripts' 'Start-SagamoreDashboard.ps1'),
        '-ConfigPath', $Cfg, '-Port', $Port, '-LogDirectory', $Logs, '-NoBrowser', '-Token', $Tok)
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
