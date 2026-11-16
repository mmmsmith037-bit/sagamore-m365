#Requires -Version 7.4
<#
.SYNOPSIS
    Uploads files from the desktop drop folder to a Teams channel folder, then makes sure every uploaded
    file takes (inherits) that folder's permissions.

.DESCRIPTION
    Normally started by the dashboard (Start-SagamoreDashboard.ps1) after the user has picked a channel and
    confirmed the permission warning. It can also be run by hand with a request file.

    Request file (JSON), written by the dashboard into logs/:
        { "destination": "Clients/Meridian-Group/Invoices", "files": ["Invoice-104.pdf"],
          "overwrite": false, "acknowledged": true }

    Order of operations, per run:
      1. Re-validate the request against the config and the drop folder (trusts nothing it is handed).
      2. Sign in (same app registration and browser sign-in as the installer).
      3. Pre-flight: the destination folder must exist, and the folder it inherits its permissions from must
         already carry the permissions from config. If the installer has never been applied there, NOTHING is
         uploaded: a file in an unstamped folder could be visible to the whole Team.
      4. For each file:
           - an existing file with the same name is skipped unless "overwrite" is set (a replace keeps the old
             copy in SharePoint version history)
           - upload
           - if the file has its own (unique) permissions -- e.g. a replaced file that had been shared with a
             link -- reset it to inherit from the folder, then read it back to verify
           - move the local copy to <drop>\_Sent\<date>\ (never deleted)
      5. Append every result to logs/drop-history.csv.

    Exit code: 0 = everything sent, 1 = some files failed or were skipped, 2 = bad request / pre-flight refused.

.PARAMETER PlanOnly
    Offline: validate the request and print what would happen. No sign-in, nothing uploaded or moved.

.EXAMPLE
    ./scripts/Send-SagamoreDrop.ps1 -RequestPath logs/drop-request-20261116-101500.json
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$RequestPath,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..' 'config' 'sagamore.config.json'),
    [string]$DropFolder,
    [string]$LogDirectory = (Join-Path $PSScriptRoot '..' 'logs'),
    [switch]$PlanOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Common.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Desktop.psm1') -Force

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
Initialize-SagamoreLog -Path (Join-Path $LogDirectory "drop-$stamp.log")
if (-not $DropFolder) { $DropFolder = Get-SagamoreDefaultDropFolder }
$DropFolder = [IO.Path]::GetFullPath($DropFolder)

# ------------------------------------------------------------------
# 1. Validate (offline)
# ------------------------------------------------------------------
$cfg = Read-SagamoreConfig -Path $ConfigPath
$problems = @(Test-SagamoreConfig -Config $cfg -AllowPlaceholderClientId:$PlanOnly)
if ($problems.Count) { $problems | ForEach-Object { Write-SagamoreLog $_ -Level ERROR }; exit 2 }

try { $request = Get-Content -LiteralPath $RequestPath -Raw | ConvertFrom-Json -AsHashtable }
catch { Write-SagamoreLog "Request file is missing or not valid JSON: $RequestPath" -Level ERROR; exit 2 }

$items = @(Get-SagamoreDropItems -Path $DropFolder)
$check = Test-SagamoreDropRequest -Config $cfg -Request $request -Items $items
if ($check.Errors.Count) { $check.Errors | ForEach-Object { Write-SagamoreLog $_ -Level ERROR }; exit 2 }

$dest      = ConvertTo-SagamorePath ([string]$request.destination)
$channel   = $check.Destination
$overwrite = $(($v = Get-OrDefault $request 'overwrite' $false) -is [bool] -and $v)
$lib       = $cfg.site.library
$target    = "$lib/$dest"

# The folder whose permissions the uploaded files will end up with: the destination itself if config gives it
# custom access, otherwise its nearest ancestor that does (normally Clients/<client>), otherwise the library.
$plan = @(Get-SagamoreFolderPlan -Config $cfg)
$source = $null
$cur = $dest
while ($cur) {
    $e = $plan | Where-Object { $_.Path -ieq $cur } | Select-Object -First 1
    if ($e -and $e.Access) { $source = $e; break }
    $cur = if ($cur.Contains('/')) { $cur.Substring(0, $cur.LastIndexOf('/')) } else { '' }
}
$sourceLabel = if ($source) { $source.Path } else { 'the library root' }
$sourceWho   = if ($source) { ($source.Access.Keys | ForEach-Object { "$_ ($($source.Access[$_]))" }) -join ', ' } else { 'library defaults' }

Write-SagamoreLog ('=' * 60)
Write-SagamoreLog '  Sagamore drop folder upload'
Write-SagamoreLog "  Channel:     $($channel.Channel) ($($channel.Type.ToLower()))"
Write-SagamoreLog "  Folder:      $dest"
Write-SagamoreLog "  Permissions: inherited from $sourceLabel -> $sourceWho"
Write-SagamoreLog "  Files:       $($check.Files.Count)$(if ($overwrite) { ' (replace existing)' })"
Write-SagamoreLog "  Drop folder: $DropFolder"
Write-SagamoreLog ('=' * 60)
$check.Files | ForEach-Object { Write-SagamoreLog "  $_" }

if ($PlanOnly) {
    Write-SagamoreLog ''
    Write-SagamoreLog 'Plan only -- nothing was uploaded or moved.' -Level SUCCESS
    exit 0
}

# ------------------------------------------------------------------
# 2. Sign in + 3. pre-flight
# ------------------------------------------------------------------
Assert-SagamoreModules
Import-Module (Join-Path $PSScriptRoot '..' 'src' 'Sagamore.Tenant.psm1') -Force
$siteUrl = Get-SagamoreSiteUrl -Config $cfg

$historyPath = Join-Path $LogDirectory 'drop-history.csv'
$who = [Environment]::UserName
$results = [System.Collections.Generic.List[object]]::new()
function Add-Result([string]$File, [long]$Size, [string]$Result, [string]$Detail) {
    $results.Add([pscustomobject]@{
        Time = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); WindowsUser = $who; File = $File; Bytes = $Size
        Channel = $channel.Channel; Folder = $dest; Result = $Result; Detail = $Detail })
}

try {
    Write-SagamoreSection 'Signing in'
    Connect-SagamoreTenant -Config $cfg -Url $siteUrl
    $list = Get-SagamoreLibrary -LibraryUrl $lib

    Write-SagamoreSection 'Pre-flight'
    try { $null = Get-PnPFolder -Url $target -ErrorAction Stop }
    catch {
        Write-SagamoreLog "Folder '$dest' does not exist in SharePoint yet. Run Preview + Apply (full install) in the dashboard first, then send again." -Level ERROR
        exit 2
    }
    if ($source) {
        $srcFolder = Get-PnPFolder -Url "$lib/$($source.Path)" -ErrorAction Stop
        if (-not (Test-SagamoreFolderUnique -Folder $srcFolder)) {
            Write-SagamoreLog "'$($source.Path)' does not have its permissions from config yet (it still inherits the whole site's access). Nothing was uploaded. Run Preview + Apply in the dashboard first." -Level ERROR
            exit 2
        }
    }
    Write-SagamoreLog "'$dest' exists and '$sourceLabel' carries its own permissions." -Level SUCCESS

    # ------------------------------------------------------------------
    # 4. Upload
    # ------------------------------------------------------------------
    Write-SagamoreSection 'Uploading'
    foreach ($name in $check.Files) {
        $local = Join-Path $DropFolder $name
        $size = (Get-Item -LiteralPath $local).Length
        try {
            $existing = Get-PnPFile -Url "$target/$name" -AsListItem -ErrorAction SilentlyContinue
            if ($existing -and -not $overwrite) {
                Write-SagamoreLog "SKIPPED $name -- a file with that name is already in $dest. Tick 'Replace files with the same name' to upload a new version." -Level WARN
                Add-Result $name $size 'Skipped' 'already exists'
                continue
            }

            Add-PnPFile -Path $local -Folder $target | Out-Null

            # Read the item back and make sure it inherits. A brand-new file always does; a replaced file keeps
            # whatever unique permissions / sharing links the old version had, so those are removed here.
            $item = Get-PnPFile -Url "$target/$name" -AsListItem -ErrorAction Stop
            $reset = $false
            if (Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments) {
                Set-PnPListItemPermission -List $list -Identity $item.Id -InheritPermissions
                $reset = $true
                $item = Get-PnPFile -Url "$target/$name" -AsListItem -ErrorAction Stop
                if (Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments) {
                    throw 'uploaded, but it still has its own permissions after the reset. Check it in SharePoint (Manage access) before sharing the channel.'
                }
            }

            $moved = Move-SagamoreDropItemToSent -DropFolder $DropFolder -Name $name
            $note = if ($reset) { 'removed its own permissions/sharing links; ' } else { '' }
            Write-SagamoreLog "SENT $name -> $dest (${note}inherits from $sourceLabel). Local copy moved to $moved" -Level SUCCESS
            Add-Result $name $size 'Sent' ($note + "inherits from $sourceLabel")
        } catch {
            Write-SagamoreLog "FAILED $name -- $($_.Exception.Message)" -Level ERROR
            Add-Result $name $size 'Failed' $_.Exception.Message
        }
    }
} finally {
    # ------------------------------------------------------------------
    # 5. History
    # ------------------------------------------------------------------
    if ($results.Count) {
        $results | Export-Csv -Path $historyPath -NoTypeInformation -Append -Encoding utf8
        Write-SagamoreLog "History: $historyPath" -Level INFO
    }
}

$sent = @($results | Where-Object Result -eq 'Sent').Count
$notSent = $results.Count - $sent
Write-SagamoreLog ''
Write-SagamoreLog "Sent $sent of $($check.Files.Count) file(s) to $($channel.Channel)." -Level $(if ($notSent) { 'WARN' } else { 'SUCCESS' })
exit ($notSent -gt 0 ? 1 : 0)
