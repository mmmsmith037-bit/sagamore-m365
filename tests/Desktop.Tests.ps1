# Offline tests for src/Sagamore.Desktop.psm1 (OneDrive evaluation, drop folder) and the drop destinations
# in src/Sagamore.Common.psm1. OneDrive is tested through its pure evaluator with hand-made "facts", so these
# run on any OS. Send-SagamoreDrop.ps1 is exercised in -PlanOnly mode (no sign-in).
BeforeAll {
    $script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $Root 'src' 'Sagamore.Common.psm1') -Force
    Import-Module (Join-Path $Root 'src' 'Sagamore.Desktop.psm1') -Force
    $script:ExamplePath = Join-Path $Root 'config' 'sagamore.config.example.json'
    function script:New-Cfg { Read-SagamoreConfig -Path $ExamplePath }

    # A healthy Windows machine; each test breaks one thing.
    function script:New-Facts {
        @{
            IsWindows = $true
            ExePath   = 'C:\Users\joe\AppData\Local\Microsoft\OneDrive\OneDrive.exe'
            Running   = $true
            Accounts  = @(@{ Email = 'joe@sagamorealliance.com'; SpoResourceId = 'https://sagamorealliance-my.sharepoint.com/'; UserFolder = 'C:\Users\joe\OneDrive - Sagamore' })
            Libraries = @(@{ Url = 'https://sagamorealliance.sharepoint.com/sites/SagamoreAlliance/Shared%20Documents'; MountPoint = 'C:\Users\joe\Sagamore Alliance' })
        }
    }
    function script:Status($State, $Id) { ($State.Checks | Where-Object { $_.Id -eq $Id }).Status }

    function script:New-Drop {
        $d = Join-Path ([IO.Path]::GetTempPath()) ('sagamore-drop-' + [guid]::NewGuid().ToString('n'))
        Initialize-SagamoreDropFolder -Path $d
    }
}

Describe 'OneDrive evaluation' {
    It 'is ok when installed, running, signed in to the tenant and syncing the library' {
        $s = Test-SagamoreOneDriveState -Facts (New-Facts) -Config (New-Cfg)
        $s.Overall | Should -Be 'ok'
        $s.SyncedPath | Should -Be 'C:\Users\joe\Sagamore Alliance'
    }

    It 'fails when OneDrive is not installed, and does not suggest starting it' {
        $f = New-Facts; $f.ExePath = $null; $f.Running = $false
        $s = Test-SagamoreOneDriveState -Facts $f -Config (New-Cfg)
        $s.Overall | Should -Be 'fail'
        Status $s 'installed' | Should -Be 'fail'
        ($s.Checks | Where-Object Id -eq 'running').Fix | Should -BeNullOrEmpty
    }

    It 'fails when installed but not running' {
        $f = New-Facts; $f.Running = $false
        $s = Test-SagamoreOneDriveState -Facts $f -Config (New-Cfg)
        Status $s 'running' | Should -Be 'fail'
        ($s.Checks | Where-Object Id -eq 'running').Fix | Should -Match 'Start OneDrive'
    }

    It 'fails when signed in only to another company''s tenant' {
        $f = New-Facts
        $f.Accounts = @(@{ Email = 'joe@othercorp.com'; SpoResourceId = 'https://othercorp-my.sharepoint.com/'; UserFolder = '' })
        $s = Test-SagamoreOneDriveState -Facts $f -Config (New-Cfg)
        Status $s 'signedIn' | Should -Be 'fail'
        ($s.Checks | Where-Object Id -eq 'signedIn').Detail | Should -Match 'othercorp'
    }

    It 'fails when no work account is signed in' {
        $f = New-Facts; $f.Accounts = @()
        Status (Test-SagamoreOneDriveState -Facts $f -Config (New-Cfg)) 'signedIn' | Should -Be 'fail'
    }

    It 'accepts an account by e-mail domain when OneDrive did not record the SharePoint host' {
        $f = New-Facts; $f.Accounts = @(@{ Email = 'laci@sagamorealliance.com'; SpoResourceId = ''; UserFolder = '' })
        Status (Test-SagamoreOneDriveState -Facts $f -Config (New-Cfg)) 'signedIn' | Should -Be 'ok'
    }

    It 'only warns when the Sagamore library is not synced (uploads still work)' {
        $f = New-Facts
        $f.Libraries = @(@{ Url = 'https://sagamorealliance.sharepoint.com/sites/SomethingElse/Shared%20Documents'; MountPoint = 'C:\x' })
        $s = Test-SagamoreOneDriveState -Facts $f -Config (New-Cfg)
        $s.Overall | Should -Be 'warn'
        Status $s 'library' | Should -Be 'warn'
        $s.SyncedPath | Should -BeNullOrEmpty
    }

    It 'counts a synced subfolder or shortcut of the Sagamore site as synced' {
        $f = New-Facts
        $f.Libraries = @(@{ Url = 'https://sagamorealliance.sharepoint.com/sites/SagamoreAlliance/Shared Documents/Clients/Meridian-Group'; MountPoint = 'C:\m' })
        Status (Test-SagamoreOneDriveState -Facts $f -Config (New-Cfg)) 'library' | Should -Be 'ok'
    }

    It 'reports unsupported (not a failure) off Windows' {
        (Test-SagamoreOneDriveState -Facts @{ IsWindows = $false } -Config (New-Cfg)).Overall | Should -Be 'unsupported'
    }

    It 'never starts OneDrive when it is already running, missing, or off Windows' {
        $f = New-Facts
        Start-SagamoreOneDrive -Facts $f | Should -BeFalse
        $f.Running = $false; $f.ExePath = $null
        Start-SagamoreOneDrive -Facts $f | Should -BeFalse
        Start-SagamoreOneDrive -Facts @{ IsWindows = $false; ExePath = 'x'; Running = $false } | Should -BeFalse
    }
}

Describe 'Office save settings' {
    BeforeAll {
        $script:Synced = 'C:\Users\joe\Sagamore Alliance'
        # Office set up exactly as wanted; each test breaks one thing.
        function script:New-OfficeFacts {
            @{
                IsWindows = $true; OfficeInstalled = $true
                PreferCloud = 0; PreferCloudPolicy = $null; CloudOnly = $null; CloudOnlyPolicy = $null
                Paths          = @{ Excel = $Synced; Word = "$Synced\Internal"; PowerPoint = "$Synced\" }
                DontAutoSave   = @{ Excel = $null; Word = $null; PowerPoint = $null }
                AutoSavePolicy = @{ Excel = $null; Word = $null; PowerPoint = $null }
                OcsiPolicy     = $null
            }
        }
    }

    It 'is ok, with nothing to change, when Office already saves into the synced library' {
        $s = Test-SagamoreOfficeSaveState -Facts (New-OfficeFacts) -SyncedPath $Synced
        $s.Overall | Should -Be 'ok'
        @($s.Changes).Count | Should -Be 0
    }

    It 'treats "never set" as cloud-first and points every app at the synced folder' {
        $f = New-OfficeFacts; $f.PreferCloud = $null; $f.Paths = @{ Excel = $null; Word = 'C:\Users\joe\Documents'; PowerPoint = $null }
        $s = Test-SagamoreOfficeSaveState -Facts $f -SyncedPath $Synced
        Status $s 'saveLocation' | Should -Be 'warn'
        ($s.Checks | Where-Object Id -eq 'saveLocation').Detail | Should -Match 'personal OneDrive'
        $c = @($s.Changes)
        ($c | Where-Object Name -eq 'PreferCloudSaveLocations').Value | Should -Be 0
        @($c | Where-Object Kind -eq 'ExpandString' | ForEach-Object Value) | Should -Be @($Synced, $Synced, $Synced)
        @($c | Where-Object Kind -eq 'ExpandString' | ForEach-Object Name) | Should -Be @('DefaultPath', 'DOC-PATH', 'Default')
    }

    It 'compares folders case-insensitively and accepts subfolders and %VARIABLES%' {
        Test-SagamorePathUnder -Path 'c:\users\JOE\sagamore alliance\Clients' -Root $Synced | Should -BeTrue
        Test-SagamorePathUnder -Path 'C:\Users\joe\Sagamore Alliance Old' -Root $Synced | Should -BeFalse
        Test-SagamorePathUnder -Path $null -Root $Synced | Should -BeFalse
    }

    It 'changes nothing about the save location until the library is synced' {
        $f = New-OfficeFacts; $f.PreferCloud = 1
        $s = Test-SagamoreOfficeSaveState -Facts $f -SyncedPath $null
        Status $s 'saveLocation' | Should -Be 'warn'
        ($s.Checks | Where-Object Id -eq 'saveLocation').Fix | Should -Match 'Sync the library first'
        @($s.Changes).Count | Should -Be 0
    }

    It 'turns off cloud-only Save As set by the user, but only reports it when a policy sets it' {
        $f = New-OfficeFacts; $f.CloudOnly = 1
        @(Get-SagamoreOfficeSaveChanges -Facts $f -SyncedPath $Synced).Name | Should -Be @('EnableCloudOnlySaveAsMode')
        $f = New-OfficeFacts; $f.CloudOnlyPolicy = 1
        $s = Test-SagamoreOfficeSaveState -Facts $f -SyncedPath $Synced
        Status $s 'saveLocation' | Should -Be 'warn'
        @($s.Changes).Count | Should -Be 0
        ($s.Checks | Where-Object Id -eq 'saveLocation').Fix | Should -Match 'IT admin'
    }

    It 'removes a per-user DontAutoSave switch even before the library is synced' {
        $f = New-OfficeFacts; $f.DontAutoSave.Excel = 1
        $s = Test-SagamoreOfficeSaveState -Facts $f -SyncedPath $null
        Status $s 'autosave' | Should -Be 'warn'
        $c = @($s.Changes)
        $c.Count | Should -Be 1
        $c[0].Remove | Should -BeTrue
        $c[0].Path | Should -Match 'Office\\16\.0\\Excel$'
    }

    It 'only reports an admin AutoSave policy, and fails when Office/OneDrive integration is blocked' {
        $f = New-OfficeFacts; $f.AutoSavePolicy.Word = 1
        $s = Test-SagamoreOfficeSaveState -Facts $f -SyncedPath $Synced
        Status $s 'autosave' | Should -Be 'warn'
        ($s.Checks | Where-Object Id -eq 'autosave').Fix | Should -Match 'IT admin'
        @($s.Changes).Count | Should -Be 0

        $f = New-OfficeFacts; $f.OcsiPolicy = 0
        $s = Test-SagamoreOfficeSaveState -Facts $f -SyncedPath $Synced
        $s.Overall | Should -Be 'fail'
        Status $s 'coauth' | Should -Be 'fail'
    }

    It 'warns when Office is missing and reports unsupported off Windows, with nothing to change' {
        $f = New-OfficeFacts; $f.OfficeInstalled = $false; $f.PreferCloud = 1
        $s = Test-SagamoreOfficeSaveState -Facts $f -SyncedPath $Synced
        $s.Overall | Should -Be 'warn'
        @($s.Changes).Count | Should -Be 0
        (Test-SagamoreOfficeSaveState -Facts @{ IsWindows = $false } -SyncedPath $Synced).Overall | Should -Be 'unsupported'
    }

    It 'refuses to write anywhere but the user''s own Office settings' {
        { Set-SagamoreOfficeSaveDefaults -Changes @(@{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'; Name = 'X'; Kind = 'DWord'; Value = 1; Remove = $false }) -WhatIf } |
            Should -Throw '*only*Office*16.0*'
        { Set-SagamoreOfficeSaveDefaults -Changes @(@{ Path = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Excel'; Name = 'X'; Kind = 'DWord'; Value = 1; Remove = $false }) -WhatIf } |
            Should -Throw
        @(Set-SagamoreOfficeSaveDefaults -Changes @() ).Count | Should -Be 0
    }
}

Describe 'Drop destinations' {
    It 'offers one entry per Teams channel, private client channels rooted at Clients/<client>' {
        $d = @(Get-SagamoreDropDestinations -Config (New-Cfg))
        @($d | ForEach-Object Channel) | Should -Be @('Internal', 'Meridian-Group', 'Hartwell-Inc')
        ($d | Where-Object Channel -eq 'Meridian-Group').Root | Should -Be 'Clients/Meridian-Group'
        ($d | Where-Object Channel -eq 'Meridian-Group').Type | Should -Be 'Private'
        ($d | Where-Object Channel -eq 'Internal').Root | Should -Be 'Internal'
    }

    It 'lists the root first and only folders managed by the plan under it' {
        $m = Get-SagamoreDropDestinations -Config (New-Cfg) | Where-Object Channel -eq 'Meridian-Group'
        $m.Folders[0] | Should -Be 'Clients/Meridian-Group'
        $m.Folders | Should -Contain 'Clients/Meridian-Group/Invoices'
        $m.Folders | Should -Contain 'Clients/Meridian-Group/Project-BrandRefresh/Deliverables'
        @($m.Folders | Where-Object { $_ -like 'Clients/Hartwell-Inc*' }).Count | Should -Be 0
    }

    It 'offers nothing when Teams is disabled' {
        $c = New-Cfg; $c.teams.enabled = $false
        @(Get-SagamoreDropDestinations -Config $c).Count | Should -Be 0
    }

    It 'offers a standard channel even when its folder is not in the plan' {
        $c = New-Cfg; $c.teams.extraChannels = @('Announcements')
        (Get-SagamoreDropDestinations -Config $c | Where-Object Channel -eq 'Announcements').Folders | Should -Be @('Announcements')
    }
}

Describe 'Drop folder' {
    BeforeEach { $script:D = New-Drop }
    AfterEach { Remove-Item $D -Recurse -Force -ErrorAction SilentlyContinue }

    It 'creates _Sent with a README, and never offers it' {
        Join-Path $D '_Sent' 'README.txt' | Should -Exist
        @(Get-SagamoreDropItems -Path $D).Count | Should -Be 0
    }

    It 'ignores Office lock files, temp files and unfinished downloads' {
        foreach ($n in '~$Report.docx', 'x.tmp', 'big.zip.crdownload', 'desktop.ini', '.hidden') { 'x' | Set-Content (Join-Path $D $n) }
        'x' | Set-Content (Join-Path $D 'Report.docx')
        @(Get-SagamoreDropItems -Path $D | ForEach-Object Name) | Should -Be @('Report.docx')
    }

    It 'lists folders but never makes them sendable' {
        New-Item -ItemType Directory (Join-Path $D 'Invoices') | Out-Null
        $i = Get-SagamoreDropItems -Path $D
        $i.Kind | Should -Be 'folder'
        $i.Sendable | Should -BeFalse
        $i.Reason | Should -Match 'zip'
    }

    It 'marks a file that another program is still writing as busy' -Skip:(-not $IsWindows) {
        # Exclusive locks are only enforced like this on Windows.
        $p = Join-Path $D 'open.xlsx'; 'x' | Set-Content $p
        $fs = [IO.File]::Open($p, 'Open', 'ReadWrite', 'None')
        try { (Get-SagamoreDropItems -Path $D).Kind | Should -Be 'busy' } finally { $fs.Dispose() }
    }

    It 'moves a sent file into _Sent/<date> without overwriting an earlier one' {
        'a' | Set-Content (Join-Path $D 'a.pdf')
        $first = Move-SagamoreDropItemToSent -DropFolder $D -Name 'a.pdf' -When ([datetime]'2026-11-16')
        'b' | Set-Content (Join-Path $D 'a.pdf')
        $second = Move-SagamoreDropItemToSent -DropFolder $D -Name 'a.pdf' -When ([datetime]'2026-11-16')
        $first | Should -Be (Join-Path $D '_Sent' '2026-11-16' 'a.pdf')
        $second | Should -Be (Join-Path $D '_Sent' '2026-11-16' 'a (2).pdf')
        Join-Path $D 'a.pdf' | Should -Not -Exist
        Get-Content $first | Should -Be 'a'
    }
}

Describe 'Drop request validation' {
    BeforeAll {
        $script:Items = @(
            @{ Name = 'a.pdf'; Kind = 'file'; Sendable = $true; Reason = '' }
            @{ Name = 'Invoices'; Kind = 'folder'; Sendable = $false; Reason = 'Folders are not uploaded.' }
        )
    }

    It 'accepts a channel subfolder regardless of case, and returns the channel' {
        $r = Test-SagamoreDropRequest -Config (New-Cfg) -Items $Items -Request @{ destination = 'clients/MERIDIAN-group/invoices'; files = @('a.pdf'); acknowledged = $true }
        $r.Errors | Should -BeNullOrEmpty
        $r.Destination.Channel | Should -Be 'Meridian-Group'
    }

    It 'requires the permission acknowledgement to be exactly true' {
        foreach ($ack in $false, 'true', 1, $null) {
            $r = Test-SagamoreDropRequest -Config (New-Cfg) -Items $Items -Request @{ destination = 'Internal'; files = @('a.pdf'); acknowledged = $ack }
            ($r.Errors -join ' ') | Should -Match 'permissions'
        }
    }

    It 'refuses folders the channels do not own, including Owner-only areas' {
        foreach ($d in '_Admin', 'Clients', '_Templates/SOW', 'Clients/Unknown-Client', '') {
            $r = Test-SagamoreDropRequest -Config (New-Cfg) -Items $Items -Request @{ destination = $d; files = @('a.pdf'); acknowledged = $true }
            $r.Errors.Count | Should -BeGreaterThan 0 -Because "'$d' must be refused"
        }
    }

    It 'refuses paths, folders, missing files and empty selections' {
        $r = Test-SagamoreDropRequest -Config (New-Cfg) -Items $Items -Request @{ destination = 'Internal'; files = @('..\x', 'sub/a.pdf', 'Invoices', 'gone.pdf'); acknowledged = $true }
        $e = $r.Errors -join "`n"
        $e | Should -Match "'\.\.\\x' is not a file name"
        $e | Should -Match "'sub/a.pdf' is not a file name"
        $e | Should -Match 'Folders are not uploaded'
        $e | Should -Match 'gone.pdf.*not in the drop folder'
        (Test-SagamoreDropRequest -Config (New-Cfg) -Items $Items -Request @{ destination = 'Internal'; files = @(); acknowledged = $true }).Errors | Should -Contain 'Pick at least one file.'
    }
}

Describe 'Send-SagamoreDrop.ps1 -PlanOnly' {
    BeforeAll {
        $script:D = New-Drop
        'x' | Set-Content (Join-Path $D 'Invoice 104.pdf')
        $script:Logs = Join-Path $D '_Sent' 'logs'   # inside _Sent so it is never listed as a drop item
        $script:Send = Join-Path $Root 'scripts' 'Send-SagamoreDrop.ps1'
        function script:Invoke-Send($Request) {
            $req = Join-Path $Logs ('req-' + [guid]::NewGuid().ToString('n') + '.json')
            New-Item -ItemType Directory $Logs -Force | Out-Null
            $Request | ConvertTo-Json | Set-Content $req
            $out = & pwsh -NoProfile -File $Send -RequestPath $req -ConfigPath $ExamplePath -DropFolder $D -LogDirectory $Logs -PlanOnly 2>&1 | Out-String
            [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out }
        }
    }
    AfterAll { Remove-Item $D -Recurse -Force -ErrorAction SilentlyContinue }

    It 'plans an upload and names the folder the permissions come from' {
        $r = Invoke-Send @{ destination = 'Clients/Meridian-Group/Project-BrandRefresh'; files = @('Invoice 104.pdf'); acknowledged = $true }
        $r.Exit | Should -Be 0
        $r.Out | Should -Match 'Channel:\s+Meridian-Group \(private\)'
        # Project-BrandRefresh has no custom access, so it inherits from the client folder
        $r.Out | Should -Match 'inherited from Clients/Meridian-Group ->.*jacquelin@sagamorealliance.com \(Edit\)'
        $r.Out | Should -Match 'nothing was uploaded or moved'
        Join-Path $D 'Invoice 104.pdf' | Should -Exist
    }

    It 'shows that a locked Invoices folder stays Owner-only' {
        $r = Invoke-Send @{ destination = 'Clients/Meridian-Group/Invoices'; files = @('Invoice 104.pdf'); acknowledged = $true }
        $r.Out | Should -Match 'inherited from Clients/Meridian-Group/Invoices -> Owner \(Full Control\)$|inherited from Clients/Meridian-Group/Invoices -> Owner \(Full Control\)\s'
    }

    It 'exits 2 without doing anything for an unacknowledged or out-of-channel request' {
        (Invoke-Send @{ destination = 'Internal'; files = @('Invoice 104.pdf'); acknowledged = $false }).Exit | Should -Be 2
        (Invoke-Send @{ destination = '_Admin'; files = @('Invoice 104.pdf'); acknowledged = $true }).Exit | Should -Be 2
        Join-Path $D 'Invoice 104.pdf' | Should -Exist
    }
}
