BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Common.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Tenant.psm1') -Force
    $script:ExamplePath = Join-Path $PSScriptRoot '..' 'config' 'sagamore.config.example.json'
    function New-TestConfig { Read-SagamoreConfig -Path $script:ExamplePath }
}

Describe 'Config validation' {
    It 'accepts the example config when the placeholder client id is allowed' {
        Test-SagamoreConfig -Config (New-TestConfig) -AllowPlaceholderClientId | Should -BeNullOrEmpty
    }
    It 'rejects the placeholder client id for a real run' {
        Test-SagamoreConfig -Config (New-TestConfig) | Should -Contain 'tenant.clientId is still the placeholder. Run scripts/Register-SagamoreApp.ps1 first.'
    }
    It 'rejects a full domain as tenant name' {
        $c = New-TestConfig; $c.tenant.name = 'sagamorealliance.sharepoint.com'
        (Test-SagamoreConfig -Config $c -AllowPlaceholderClientId) -join "`n" | Should -Match 'tenant.name'
    }
    It 'rejects an unknown role name' {
        $c = New-TestConfig; $c.folders[0].access = @{ Owner = 'Full Control'; Member = 'Write' }
        (Test-SagamoreConfig -Config $c -AllowPlaceholderClientId) -join "`n" | Should -Match "role 'Write'"
    }
    It 'rejects a UPN in an access map that is not in people' {
        $c = New-TestConfig; $c.clients[0].access['stranger@example.com'] = 'Read'
        (Test-SagamoreConfig -Config $c -AllowPlaceholderClientId) -join "`n" | Should -Match 'not listed in people'
    }
    It 'requires at least one Owner' {
        $c = New-TestConfig; foreach ($p in $c.people) { if ($p.role -eq 'Owner') { $p.role = 'Member' } }
        (Test-SagamoreConfig -Config $c -AllowPlaceholderClientId) -join "`n" | Should -Match 'at least one person must have role Owner'
    }
    It 'rejects duplicate clients regardless of case' {
        $c = New-TestConfig; $c.clients += @{ name = 'meridian-group' }
        (Test-SagamoreConfig -Config $c -AllowPlaceholderClientId) -join "`n" | Should -Match 'listed twice'
    }
    It 'rejects illegal characters in folder paths' {
        $c = New-TestConfig; $c.folders += @{ path = 'Internal/Q3: plans' }
        (Test-SagamoreConfig -Config $c -AllowPlaceholderClientId) -join "`n" | Should -Match 'does not allow'
    }
    It 'rejects an invalid sharing value' {
        $c = New-TestConfig; $c.site.externalSharing = 'Anyone'
        (Test-SagamoreConfig -Config $c -AllowPlaceholderClientId) -join "`n" | Should -Match 'externalSharing'
    }
    It 'rejects role groups on a client folder (client silo)' {
        $c = New-TestConfig; $c.clients[0].access['Member'] = 'Read'
        (Test-SagamoreConfig -Config $c -AllowPlaceholderClientId) -join "`n" | Should -Match "clients\['Meridian-Group'\].access: 'Member' .* breaks the client silo"
    }
    It 'rejects role groups on client subfolders and the client template' {
        $c = New-TestConfig
        $c.clients[0].folders[1].access['Support'] = 'Read'
        $c.clientTemplate.access['Member'] = 'Edit'
        $errs = (Test-SagamoreConfig -Config $c -AllowPlaceholderClientId) -join "`n"
        $errs | Should -Match "Deliverables'\].access: 'Support'"
        $errs | Should -Match "clientTemplate.access: 'Member'"
    }
    It 'still allows role groups outside Clients' {
        $c = New-TestConfig; $c.folders[0].access['Member'] = 'Read'
        Test-SagamoreConfig -Config $c -AllowPlaceholderClientId | Should -BeNullOrEmpty
    }
    It 'reports missing sections instead of throwing' {
        Test-SagamoreConfig -Config @{ tenant = @{} } -AllowPlaceholderClientId | Should -Contain "Missing required section 'site'."
    }
}

Describe 'Folder plan' {
    BeforeAll { $script:Plan = Get-SagamoreFolderPlan -Config (New-TestConfig) }

    It 'lists every parent before its children' {
        $index = @{}; for ($i = 0; $i -lt $Plan.Count; $i++) { $index[$Plan[$i].Path] = $i }
        foreach ($p in $Plan) {
            if ($p.Path -match '/') {
                $parent = $p.Path.Substring(0, $p.Path.LastIndexOf('/'))
                $index[$parent] | Should -BeLessThan $index[$p.Path] -Because "$parent must be created before $($p.Path)"
            }
        }
    }
    It 'has no duplicate paths' {
        ($Plan.Path | ForEach-Object ToLower | Group-Object | Where-Object Count -gt 1) | Should -BeNullOrEmpty
    }
    It 'locks every client Invoices folder to Owners only' {
        $inv = @($Plan | Where-Object { $_.Path -like 'Clients/*/Invoices' })
        $inv.Count | Should -Be 2
        foreach ($i in $inv) { @($i.Access.Keys) | Should -Be @('Owner'); $i.Access.Owner | Should -Be 'Full Control' }
    }
    It 'always gives Owner Full Control on folders with unique access, even if config tries otherwise' {
        $c = New-TestConfig; $c.folders[0].access = @{ Owner = 'Read' }
        $p = Get-SagamoreFolderPlan -Config $c | Where-Object Path -eq '_Admin'
        $p.Access.Owner | Should -Be 'Full Control'
    }
    It 'uses per-client access when given, template otherwise' {
        ($Plan | Where-Object Path -eq 'Clients/Meridian-Group').Access.Keys | Should -Contain 'jacquelin@sagamorealliance.com'
        $c = New-TestConfig; $c.clients[1].Remove('access')
        $p = Get-SagamoreFolderPlan -Config $c | Where-Object Path -eq 'Clients/Hartwell-Inc'
        @($p.Access.Keys) | Should -Be @('Owner')
    }
    It 'leaves folders without access to inherit' {
        ($Plan | Where-Object Path -eq 'Clients/Meridian-Group/_ClientShared').Access | Should -BeNullOrEmpty
    }
    It 'adds implied parent folders' {
        $c = New-TestConfig; $c.folders = @(@{ path = 'A/B/C'; access = @{ Owner = 'Full Control' } })
        (Get-SagamoreFolderPlan -Config $c | Where-Object { $_.Path -in 'A', 'A/B' }).Count | Should -Be 2
    }
    It 'locks a nested Invoices folder deep inside a client project' {
        $c = New-TestConfig; $c.clients[1].folders += @{ path = 'Project-Q3Audit/Invoices'; access = @{ Owner = 'Full Control'; Member = 'Edit' } }
        $p = Get-SagamoreFolderPlan -Config $c | Where-Object Path -eq 'Clients/Hartwell-Inc/Project-Q3Audit/Invoices'
        @($p.Access.Keys) | Should -Be @('Owner')
    }
    It 'applies the client template to clients discovered on the site' {
        $p = Get-SagamoreFolderPlan -Config (New-TestConfig) -ExtraClients 'Blue-Ridge'
        ($p | Where-Object Path -eq 'Clients/Blue-Ridge/Invoices').Access.Keys | Should -Be @('Owner')
    }
    It 'normalises slashes and whitespace in paths' {
        ConvertTo-SagamorePath ' /Clients\\Foo// Bar / ' | Should -Be 'Clients/Foo/Bar'
    }
}

Describe 'Names' {
    It 'makes local folder names SharePoint-safe' {
        ConvertTo-SagamoreFolderName 'Meridian Group: "2026" #1' | Should -Be 'Meridian-Group-2026-1'
    }
    It 'throws on names with nothing usable' {
        { ConvertTo-SagamoreFolderName '***' } | Should -Throw
    }
    It 'truncates channel names to 50 characters and strips illegal characters' {
        $n = ConvertTo-SagamoreChannelName ('A&B ' + ('x' * 80))
        $n.Length | Should -BeLessOrEqual 50
        $n | Should -Not -Match '&'
    }
    It 'plans one channel per client plus extra channels' {
        $ch = Get-SagamoreChannelPlan -Config (New-TestConfig)
        $ch.Channel | Should -Contain 'Internal'
        $ch.Channel | Should -Contain 'Meridian-Group'
        ($ch | Where-Object Channel -eq 'Meridian-Group').Folder | Should -Be 'Clients/Meridian-Group'
    }
    It 'makes every client channel private and extra channels standard' {
        $ch = Get-SagamoreChannelPlan -Config (New-TestConfig)
        ($ch | Where-Object Channel -eq 'Internal').Type | Should -Be 'Standard'
        @($ch | Where-Object Folder | ForEach-Object Type | Select-Object -Unique) | Should -Be @('Private')
    }
    It 'invites exactly the people named on the client folder, with Owners as channel owners' {
        $m = Get-SagamoreChannelPlan -Config (New-TestConfig) | Where-Object Channel -eq 'Meridian-Group'
        $m.Owners  | Should -Be @('joe@sagamorealliance.com')
        $m.Members | Should -Be @('jacquelin@sagamorealliance.com', 'support1@sagamorealliance.com')
        # Laci only has Add Only on a Meridian subfolder: she is not invited to the channel
        $m.Members | Should -Not -Contain 'laci@sagamorealliance.com'
    }
    It 'invites nobody to a client with no access list or discovered on the site' {
        $c = New-TestConfig; $c.clients[1].Remove('access')
        $ch = Get-SagamoreChannelPlan -Config $c -ExtraClients 'Blue-Ridge'
        @(($ch | Where-Object Channel -eq 'Hartwell-Inc').Members).Count | Should -Be 0
        $b = $ch | Where-Object Channel -eq 'Blue-Ridge'
        $b.Type | Should -Be 'Private'; @($b.Members).Count | Should -Be 0
    }
    It 'plans no channels when Teams is disabled' {
        $c = New-TestConfig; $c.teams.enabled = $false
        Get-SagamoreChannelPlan -Config $c | Should -BeNullOrEmpty
    }
}

Describe 'Principals and access comparison' {
    It 'maps role keys to Entra group claims' {
        Resolve-SagamorePrincipal -Key 'Member' -GroupIds @{ Member = 'abc' } | Should -Be 'c:0t.c|tenant|abc'
    }
    It 'passes UPNs through' {
        Resolve-SagamorePrincipal -Key 'a@b.com' -GroupIds @{} | Should -Be 'a@b.com'
    }
    It 'normalises logins for comparison' {
        ConvertTo-SagamoreLogin 'Laci@Sagamore.com' | Should -Be 'i:0#.f|membership|laci@sagamore.com'
        ConvertTo-SagamoreLogin 'c:0t.c|tenant|ABC' | Should -Be 'c:0t.c|tenant|abc'
    }
    It 'finds missing, wrong-role and unexpected principals' {
        $desired = @{ 'a' = 'Edit'; 'b' = 'Read' }
        $actual  = @{ 'a' = @{ Roles = @('Read') }; 'c' = @{ Roles = @('Edit') } }
        $d = Compare-SagamoreAccess -Desired $desired -Actual $actual
        ($d | Where-Object Issue -eq 'WrongRole').Login  | Should -Be 'a'
        ($d | Where-Object Issue -eq 'Missing').Login    | Should -Be 'b'
        ($d | Where-Object Issue -eq 'Unexpected').Login | Should -Be 'c'
    }
    It 'finds channel members to invite, re-role and remove, but keeps installer admins' {
        $desired = @{ joe = 'owner'; jac = 'member'; laci = 'member' }
        $actual  = @{ joe = @{ Role = 'owner' }; jac = @{ Role = 'owner' }; sup2 = @{ Role = 'member' }; admin = @{ Role = 'owner' } }
        $d = Compare-SagamoreChannelMembers -Desired $desired -Actual $actual -KeepIds 'admin'
        ($d | Where-Object Issue -eq 'Missing').UserId    | Should -Be 'laci'
        ($d | Where-Object Issue -eq 'WrongRole').UserId  | Should -Be 'jac'
        ($d | Where-Object Issue -eq 'Unexpected').UserId | Should -Be 'sup2'
        @($d).Count | Should -Be 3
    }
    It 'reports nothing when access matches' {
        Compare-SagamoreAccess -Desired @{ a = 'Read' } -Actual @{ a = @{ Roles = @('Read') } } | Should -BeNullOrEmpty
    }
}
