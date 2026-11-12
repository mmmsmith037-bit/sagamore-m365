#Requires -Version 7.4
<#
.SYNOPSIS
    Imports client folders (and optionally their files) from a local folder into Clients/.

.DESCRIPTION
    Each top-level subfolder of -SourcePath becomes a client, e.g.
        D:\Joe\Clients\Meridian Group\Project X\file.docx
     -> Shared Documents/Clients/Meridian-Group/Project-X/file.docx

    Safe order of operations (so nothing is ever visible to the wrong people):
      1. Preview the mapping (nothing changes with -WhatIf)
      2. Add new clients to config/sagamore.config.json
      3. Create the folder tree in SharePoint
      4. Run Install-Sagamore.ps1 -PermissionsOnly  (locks Invoices etc.)
      5. Only THEN upload files (-UploadFiles)

.PARAMETER SourcePath
    Local folder whose subfolders are the clients.

.PARAMETER Only
    Import only these client folder names (local names).

.PARAMETER UploadFiles
    Upload the files too. Without it, only the folder structure is created.

.EXAMPLE
    ./scripts/Import-SagamoreClients.ps1 -SourcePath 'D:\Joe\Clients' -WhatIf
.EXAMPLE
    ./scripts/Import-SagamoreClients.ps1 -SourcePath 'D:\Joe\Clients' -UploadFiles
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$SourcePath,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..' 'config' 'sagamore.config.json'),
    [string[]]$Only = @(),
    [switch]$UploadFiles,
    [string]$LogDirectory = (Join-Path $PSScriptRoot '..' 'logs')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Common.psm1') -Force
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
Initialize-SagamoreLog -Path (Join-Path $LogDirectory "import-$stamp.log")

if (-not (Test-Path $SourcePath -PathType Container)) { throw "SourcePath not found: $SourcePath" }
$source = (Resolve-Path $SourcePath).Path.TrimEnd('\', '/')

$cfg = Read-SagamoreConfig -Path $ConfigPath
$problems = @(Test-SagamoreConfig -Config $cfg -AllowPlaceholderClientId:$WhatIfPreference)
if ($problems.Count) { $problems | ForEach-Object { Write-SagamoreLog $_ -Level ERROR }; exit 2 }

# ------------------------------------------------------------------
# 1. Build the mapping
# ------------------------------------------------------------------
Write-SagamoreSection 'Import preview'
$clientDirs = @(Get-ChildItem -Path $source -Directory | Where-Object { -not $_.Name.StartsWith('.') })
if ($Only.Count) { $clientDirs = @($clientDirs | Where-Object { $_.Name -in $Only }) }
if (-not $clientDirs.Count) { Write-SagamoreLog 'No client folders found.' -Level WARN; exit 0 }

$map = foreach ($d in $clientDirs) {
    $clientName = ConvertTo-SagamoreFolderName $d.Name
    $sub = @(Get-ChildItem -Path $d.FullName -Directory -Recurse | ForEach-Object {
        $rel = $_.FullName.Substring($d.FullName.Length).TrimStart('\', '/')
        (($rel -split '[\\/]') | ForEach-Object { ConvertTo-SagamoreFolderName $_ }) -join '/'
    })
    $files = @(Get-ChildItem -Path $d.FullName -File -Recurse)
    [pscustomobject]@{ Local = $d.FullName; Name = $clientName; Folders = $sub; Files = $files
                       Bytes = ($files | Measure-Object Length -Sum).Sum }
}
$dupes = @($map | Group-Object { $_.Name.ToLower() } | Where-Object Count -gt 1)
if ($dupes.Count) { throw "These local folders map to the same SharePoint name: $(($dupes | ForEach-Object { ($_.Group.Local) -join ' + ' }) -join '; ')" }

$known = @($cfg.clients | ForEach-Object { $_.name.ToLower() })
foreach ($m in $map) {
    $state = if ($m.Name.ToLower() -in $known) { 'existing' } else { 'NEW' }
    Write-SagamoreLog ("{0,-8} {1,-40} {2,4} folders {3,6} files {4,10:N1} MB   <- {5}" -f $state, "Clients/$($m.Name)", $m.Folders.Count, $m.Files.Count, (($m.Bytes ?? 0) / 1MB), $m.Local) -Level PLAN
}
$long = @($map | ForEach-Object { $m = $_; $m.Files | Where-Object { ("$($cfg.site.library)/Clients/$($m.Name)/" + $_.FullName.Substring($m.Local.Length)).Length -gt 380 } })
if ($long.Count) { Write-SagamoreLog "$($long.Count) file path(s) are close to SharePoint's 400-character limit and may fail to upload." -Level WARN }

# ------------------------------------------------------------------
# 2. Update config
# ------------------------------------------------------------------
$new = @($map | Where-Object { $_.Name.ToLower() -notin $known })
if ($new.Count -and $PSCmdlet.ShouldProcess($ConfigPath, "Add $($new.Count) client(s) to config")) {
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($c in @($cfg.clients)) { $list.Add($c) }
    foreach ($n in $new) { $list.Add([ordered]@{ name = $n.Name; folders = @() }) }
    $cfg.clients = $list.ToArray()
    Save-SagamoreConfig -Config $cfg -Path $ConfigPath
    Write-SagamoreLog "Config updated with: $($new.Name -join ', ')" -Level SUCCESS
}

if ($WhatIfPreference) { Write-SagamoreLog 'WhatIf: stopping before any SharePoint changes.' -Level PLAN; exit 0 }

# ------------------------------------------------------------------
# 3. Create folders
# ------------------------------------------------------------------
Assert-SagamoreModules
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Tenant.psm1') -Force
$siteUrl = Get-SagamoreSiteUrl -Config $cfg
$lib = $cfg.site.library
Connect-SagamoreTenant -Config $cfg -Url $siteUrl

Write-SagamoreSection 'Creating folders'
foreach ($m in $map) {
    $null = Resolve-PnPFolder -SiteRelativePath "$lib/Clients/$($m.Name)"
    foreach ($f in $m.Folders) { $null = Resolve-PnPFolder -SiteRelativePath "$lib/Clients/$($m.Name)/$f" }
    Write-SagamoreLog "Clients/$($m.Name): $($m.Folders.Count) subfolder(s)" -Level SUCCESS
}

# ------------------------------------------------------------------
# 4. Permissions BEFORE files
# ------------------------------------------------------------------
Write-SagamoreSection 'Applying permissions before upload'
& (Join-Path $PSScriptRoot 'Install-Sagamore.ps1') -ConfigPath $ConfigPath -PermissionsOnly -LogDirectory $LogDirectory
if ($LASTEXITCODE -ne 0) { throw 'Permission stamping failed -- files were NOT uploaded. Fix the error above and re-run.' }
# Install-Sagamore leaves the connection on the site; reconnect defensively
Connect-SagamoreTenant -Config $cfg -Url $siteUrl

# ------------------------------------------------------------------
# 5. Upload
# ------------------------------------------------------------------
if (-not $UploadFiles) {
    Write-SagamoreLog 'Folder structure created and locked. Re-run with -UploadFiles to copy the files.' -Level INFO
    exit 0
}

Write-SagamoreSection 'Uploading files'
$ok = 0; $skipped = 0; $failed = [System.Collections.Generic.List[string]]::new()
foreach ($m in $map) {
    foreach ($file in $m.Files) {
        $relDir = Split-Path ($file.FullName.Substring($m.Local.Length).TrimStart('\', '/')) -Parent
        $spDir = if ($relDir) { (($relDir -split '[\\/]') | ForEach-Object { ConvertTo-SagamoreFolderName $_ }) -join '/' } else { '' }
        $target = "$lib/Clients/$($m.Name)" + $(if ($spDir) { "/$spDir" } else { '' })
        try {
            $existing = Get-PnPFile -Url "$target/$($file.Name)" -AsListItem -ErrorAction SilentlyContinue
            if ($existing -and [long]$existing['File_x0020_Size'] -eq $file.Length) { $skipped++; continue }
            Add-PnPFile -Path $file.FullName -Folder $target | Out-Null
            $ok++
            if ($ok % 25 -eq 0) { Write-SagamoreLog "  $ok uploaded..." -Level INFO }
        } catch {
            $failed.Add("$($file.FullName): $($_.Exception.Message)")
        }
    }
}
Write-SagamoreLog "Uploaded $ok, skipped $skipped unchanged, failed $($failed.Count)" -Level $(if ($failed.Count) { 'WARN' } else { 'SUCCESS' })
$failed | ForEach-Object { Write-SagamoreLog "  FAILED $_" -Level ERROR }

# New "Invoices" folders that only appeared in the files are covered by the locked sweep; run it once more.
& (Join-Path $PSScriptRoot 'Install-Sagamore.ps1') -ConfigPath $ConfigPath -PermissionsOnly -SkipTeams -LogDirectory $LogDirectory | Out-Null
exit ($failed.Count -gt 0 ? 1 : 0)
