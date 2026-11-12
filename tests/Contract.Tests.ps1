# Static checks that catch the class of bug that broke the original script (wrong cmdlet parameters)
# without needing a tenant: every PnP cmdlet the scripts call must exist in the installed module and
# every named parameter must be valid for it. Plus PSScriptAnalyzer errors and parse errors.
BeforeDiscovery {
    $root = Join-Path $PSScriptRoot '..'
    $script:Files = @(Get-ChildItem (Join-Path $root 'scripts'), (Join-Path $root 'src') -Include *.ps1, *.psm1 -Recurse |
        ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } })
}

BeforeAll {
    Import-Module PnP.PowerShell -ErrorAction Stop
    # Parameters PnP adds dynamically (only visible once -Type etc. is bound).
    $script:Dynamic = @{
        'New-PnPSite'     = @('Title', 'Alias', 'Owners', 'Description', 'IsPublic', 'Lcid', 'Url', 'Owner', 'SiteAlias', 'Classification')
        'Add-PnPTeamsTab' = @('ContentUrl', 'WebSiteUrl', 'EntityId')
    }
}

Describe '<Name>' -ForEach $Files {
    It 'parses without errors' {
        $errors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$errors)
        $errors | Should -BeNullOrEmpty
    }

    It 'only calls PnP cmdlets and parameters that exist in PnP.PowerShell' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
        $calls = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) |
                 Where-Object { $_.GetCommandName() -like '*-PnP*' }
        $bad = foreach ($c in $calls) {
            $name = $c.GetCommandName()
            $cmd = Get-Command $name -ErrorAction SilentlyContinue
            if (-not $cmd) { "line $($c.Extent.StartLineNumber): $name does not exist"; continue }
            $valid = @($cmd.Parameters.Keys) + @($Dynamic[$name]) + @($cmd.Parameters.Values | ForEach-Object Aliases)
            foreach ($e in $c.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] }) {
                $p = $e.ParameterName
                $match = @($valid | Where-Object { $_ -and $_ -like "$p*" })
                if (-not $match) { "line $($c.Extent.StartLineNumber): $name has no parameter -$p" }
            }
        }
        $bad | Should -BeNullOrEmpty
    }

    It 'has no PSScriptAnalyzer errors' {
        $r = Invoke-ScriptAnalyzer -Path $Path -Severity Error
        ($r | ForEach-Object { "line $($_.Line): $($_.RuleName) $($_.Message)" }) | Should -BeNullOrEmpty
    }
}
