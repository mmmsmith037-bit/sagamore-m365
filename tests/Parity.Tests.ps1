# The dashboard (dashboard/plan.js) and the installer (src/Sagamore.Common.psm1) must agree on the
# folder plan and on validation, otherwise what you see in the dashboard is not what gets applied.
BeforeDiscovery { $script:HasNode = [bool](Get-Command node -ErrorAction SilentlyContinue) }

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Common.psm1') -Force
    $script:PlanJs  = (Resolve-Path (Join-Path $PSScriptRoot '..' 'dashboard' 'plan.js')).Path
    $script:Example = (Resolve-Path (Join-Path $PSScriptRoot '..' 'config' 'sagamore.config.example.json')).Path

    function Invoke-JsPlan([string]$ConfigFile) {
        $js = @"
const P = require(process.argv[1]);
const cfg = JSON.parse(require('fs').readFileSync(process.argv[2], 'utf8'));
const plan = P.folderPlan(cfg).map(p => ({ Path: p.Path, Access: p.Access }));
const errors = P.validate(cfg, { allowPlaceholderClientId: true });
process.stdout.write(JSON.stringify({ plan, errors }));
"@
        $out = node -e $js $script:PlanJs $ConfigFile
        if ($LASTEXITCODE -ne 0) { throw "node failed: $out" }
        $out | ConvertFrom-Json -AsHashtable
    }
    function Get-PsResult([string]$ConfigFile) {
        $cfg = Read-SagamoreConfig -Path $ConfigFile
        @{
            plan   = @(Get-SagamoreFolderPlan -Config $cfg | ForEach-Object { @{ Path = $_.Path; Access = $_.Access } })
            errors = @(Test-SagamoreConfig -Config $cfg -AllowPlaceholderClientId)
        }
    }
    function Format-Plan($plan) {
        $lines = foreach ($item in $plan) {
            $acc = if ($item.Access) { (@($item.Access.Keys) | Sort-Object | ForEach-Object { "$_=$($item.Access[$_])" }) -join ';' } else { '(inherit)' }
            "$($item.Path) | $acc"
        }
        $lines -join "`n"
    }
}

Describe 'Dashboard and installer agree' -Skip:(-not $HasNode) {
    It 'produce the same folder plan for the example config' {
        $js = Invoke-JsPlan $Example; $ps = Get-PsResult $Example
        Format-Plan $js.plan | Should -Be (Format-Plan $ps.plan)
    }
    It 'produce the same validation result for the example config' {
        $js = Invoke-JsPlan $Example; $ps = Get-PsResult $Example
        (@($js.errors) -join "`n") | Should -Be (@($ps.errors) -join "`n")
    }
    It 'agree on an edited config with errors, extra folders and nested locks' {
        $cfg = Read-SagamoreConfig -Path $Example
        $cfg.folders += @{ path = 'Deep/Er/Folder'; access = @{ Member = 'Read' } }
        $cfg.clients[1].folders += @{ path = 'Project-Q3Audit/Invoices'; access = @{ Member = 'Edit' } }
        $cfg.clients += @{ name = 'New-Client'; access = @{ Owner = 'Read'; Support = 'Edit' } }
        $cfg.folders[0].access['ghost@nowhere.com'] = 'Read'
        $cfg.libraryAccess.Member = 'Write'
        $tmp = Join-Path ([IO.Path]::GetTempPath()) "sag-parity-$([guid]::NewGuid()).json"
        try {
            $cfg | ConvertTo-Json -Depth 20 | Set-Content $tmp
            $js = Invoke-JsPlan $tmp; $ps = Get-PsResult $tmp
            Format-Plan $js.plan | Should -Be (Format-Plan $ps.plan)
            (@($js.errors) -join "`n") | Should -Be (@($ps.errors) -join "`n")
            @($ps.errors).Count | Should -BeGreaterThan 0
        } finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
    }
}
