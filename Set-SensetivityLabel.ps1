#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '3.0.0' }

<#
.SYNOPSIS
    Sets a Microsoft Purview sensitivity label on documents in SharePoint Online
    using the dedicated Add-PnPFileSensitivityLabel cmdlet with PnP batching.

.DESCRIPTION
    Scope is resolved automatically from the parameters supplied:
      -FilePath                  → single file only
      -LibraryName               → all files in that library
      (neither)                  → all non-system libraries on the site

    Uses Add-PnPFileSensitivityLabel (Graph API) with PnP batch execution for
    efficient bulk updates. Throttle responses are handled with exponential
    backoff and configurable retries.

    Authentication requires your own registered Entra ID app (the shared PnP
    Management Shell app was retired in September 2024).

    Library discovery is language-agnostic: BaseTemplate 101 identifies
    document libraries regardless of the site language setting.

.PARAMETER SiteUrl
    Full URL of the SharePoint Online site collection.
    Example: https://avarante.sharepoint.com/sites/HR

.PARAMETER ClientId
    Client ID (AppId) of your registered Entra ID app.
    Register one with: Register-PnPEntraIDAppForInteractiveLogin

.PARAMETER LabelId
    GUID of the sensitivity label from Microsoft Purview.
    Find in: Purview compliance portal → Information protection → Labels → GUID

.PARAMETER LabelName
    Display name of the label (logging only). Example: "Confidential"

.PARAMETER AssignmentMethod
    How the label is assigned. Valid values: Standard, Privileged, Auto.
    Default: Standard.
    Use Privileged when downgrading from a higher classification.

.PARAMETER JustificationText
    Optional justification text. Required by some label policies when
    downgrading a label (e.g. replacing Highly Confidential with General).

.PARAMETER LibraryName
    (Optional) Target a single document library by display name.
    Mutually exclusive with -FilePath.

.PARAMETER FilePath
    (Optional) Server-relative path to a single file.
    Example: /sites/HR/Documents/Contract.docx
    Mutually exclusive with -LibraryName.

.PARAMETER SkipAlreadyLabelled
    Skip files already carrying this exact target label. Default: $true

.PARAMETER OverwriteExistingLabel
    Allow overwriting files that already carry ANY sensitivity label other
    than the target. Default: $false — files with an existing different label
    are blocked with a warning to prevent silent clobbering of intentional
    labels. Set to $true to replace existing labels.

.PARAMETER BatchSize
    Number of label assignments queued per PnP batch execution. Default: 20.
    Lower this if you hit Graph batch size limits.

.PARAMETER MaxRetries
    Max retry attempts on throttle / transient errors. Default: 5.

.PARAMETER RetryBaseDelaySeconds
    Base delay in seconds for exponential backoff. Doubles each attempt.
    Default: 2  →  delays of ~2s, 4s, 8s, 16s, 32s.

.PARAMETER LogPath
    (Optional) Path to write a CSV log of all processed files.
    Example: C:\Logs\label-run.csv

.EXAMPLE
    # Register your Entra ID app (one-time per tenant)
    Register-PnPEntraIDAppForInteractiveLogin `
        -ApplicationName "PnP-SensitivityLabel" `
        -Tenant "avarante.onmicrosoft.com" `
        -Interactive

    # Single file
    .\Set-SensitivityLabel.ps1 `
        -SiteUrl   "https://avarante.sharepoint.com/sites/HR" `
        -ClientId  "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -LabelId   "7f142f37-5120-4b7b-a8de-9e48b50dc66e" `
        -LabelName "Confidential" `
        -FilePath  "/sites/HR/Documents/Contract.docx"

    # Single library — with justification text for label downgrade
    .\Set-SensitivityLabel.ps1 `
        -SiteUrl            "https://avarante.sharepoint.com/sites/HR" `
        -ClientId           "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -LabelId            "7f142f37-5120-4b7b-a8de-9e48b50dc66e" `
        -LabelName          "Confidential" `
        -LibraryName        "Documents" `
        -AssignmentMethod   "Privileged" `
        -JustificationText  "Bulk reclassification approved by compliance team"

    # Entire site collection — dry run with CSV log
    .\Set-SensitivityLabel.ps1 `
        -SiteUrl   "https://avarante.sharepoint.com/sites/HR" `
        -ClientId  "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -LabelId   "7f142f37-5120-4b7b-a8de-9e48b50dc66e" `
        -LabelName "Confidential" `
        -LogPath   "C:\Logs\label-run.csv" `
        -WhatIf

.NOTES
    Requirements:
      - PowerShell 7.4+          →  https://aka.ms/powershell
      - PnP.PowerShell 3.x+      →  Install-Module PnP.PowerShell
      - Site Collection Admin or Site Owner permissions
      - Sensitivity label published to the account running the script
      - Own Entra ID app registered (PnP Management Shell app retired Sept 2024)

    Batching notes:
      Add-PnPFileSensitivityLabel uses the Microsoft Graph API under the hood.
      PnP batching groups multiple Graph calls into a single HTTP request,
      significantly reducing round-trips vs. one call per file.
      BatchSize default of 20 stays well within Graph's 20-request batch limit.

    Version: 6.0  |  2026-03-17
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'SiteCollection')]
param (
    [Parameter(Mandatory)][string] $SiteUrl,
    [Parameter(Mandatory)][string] $ClientId,

    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string] $LabelId,

    [Parameter(Mandatory)][string] $LabelName,

    [ValidateSet('Standard', 'Privileged', 'Auto')]
    [string] $AssignmentMethod = 'Standard',

    [string] $JustificationText,

    [Parameter(ParameterSetName = 'Library')][string] $LibraryName,
    [Parameter(ParameterSetName = 'File')]   [string] $FilePath,

    [bool]   $SkipAlreadyLabelled    = $true,
    [bool]   $OverwriteExistingLabel = $false,
    [int]    $BatchSize              = 20,
    [int]    $MaxRetries             = 5,
    [int]    $RetryBaseDelaySeconds  = 2,
    [string] $LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
#  Counters and log
# ─────────────────────────────────────────────────────────────────────────────
$stats = @{ Ok = 0; Skipped = 0; Protected = 0; Failed = 0 }

$logEntries = [System.Collections.Generic.List[PSCustomObject]]::new()

function Add-LogEntry {
    param([string]$Library, [string]$File, [string]$Result, [string]$Detail = '')
    $logEntries.Add([PSCustomObject]@{
        Timestamp = (Get-Date -Format 'o')
        Library   = $Library
        FilePath  = $File
        Result    = $Result
        Detail    = $Detail
    })
}

# ─────────────────────────────────────────────────────────────────────────────
#  System library detection
# ─────────────────────────────────────────────────────────────────────────────
$SystemFolderNames = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
@(
    'Style Library','FormServerTemplates','Site Assets','Site Pages',
    'SiteAssets','SitePages','_catalogs','appdata','appfiles',
    'Lists','IWConvertedForms'
) | ForEach-Object { $SystemFolderNames.Add($_) | Out-Null }

function Test-IsSystemLibrary {
    param($List)
    return ($List.BaseTemplate -ne 101) -or
           $List.Hidden -or
           $SystemFolderNames.Contains($List.RootFolder.Name)
}

# ─────────────────────────────────────────────────────────────────────────────
#  Exponential backoff wrapper
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-WithRetry {
    param(
        [scriptblock] $Action,
        [int]         $MaxAttempts      = $MaxRetries,
        [int]         $BaseDelaySeconds = $RetryBaseDelaySeconds
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return (& $Action)
        }
        catch {
            $msg     = $_.Exception.Message
            $isRetry = $msg -match '429|throttl|Too Many Requests|Request throttled|503|ServiceUnavailable|timeout|timed out'
            $wait    = $BaseDelaySeconds * [math]::Pow(2, $attempt - 1)

            if ($msg -match 'Retry-After[: ]+(\d+)') { $wait = [int]$Matches[1] }

            $jitter = $wait * 0.1 * (Get-Random -Minimum -1.0 -Maximum 1.0)
            $wait   = [math]::Max(1, $wait + $jitter)

            if ($isRetry -and $attempt -lt $MaxAttempts) {
                Write-Warning "  Throttled / transient error (attempt $attempt / $MaxAttempts). Retrying in $([math]::Round($wait, 1))s..."
                Start-Sleep -Seconds $wait
            }
            else { throw }
        }
    }
}

function Invoke-LabelBatch {
    param(
        [object[]] $Items,
        [string]   $ListTitle,
        [int]      $LibIndex,
        [int]      $LibTotal
    )

    $count = $Items.Count
    $batch = New-PnPBatch
    $batchCount  = 0
    $fileIndex   = 0

    foreach ($item in $Items) {
        $fileIndex++
        $path    = $item['FileRef']
        $leaf    = [System.IO.Path]::GetFileName($path)
        $current = $item['_ComplianceTag']

        Write-Progress -Id 1 `
                       -Activity "[$LibIndex / $LibTotal]  $ListTitle" `
                       -Status   "Evaluating [$fileIndex / $count]  $leaf" `
                       -PercentComplete ([math]::Round($fileIndex / $count * 100))

        # ── Guard: already has the target label ───────────────────────────────
        if ($SkipAlreadyLabelled -and $current -eq $LabelId) {
            Write-Host "  │  SKIP    $leaf  (already has target label)" -ForegroundColor DarkGray
            Add-LogEntry -Library $ListTitle -File $path -Result 'Skipped' -Detail 'Already has target label'
            $stats.Skipped++
            continue
        }

        # ── Guard: has a DIFFERENT existing label — block unless opted in ─────
        if (-not $OverwriteExistingLabel -and $current -and $current -ne $LabelId) {
            Write-Host "  │  PROT    $leaf  (has existing label — use -OverwriteExistingLabel to replace)" -ForegroundColor Yellow
            Add-LogEntry -Library $ListTitle -File $path -Result 'Protected' -Detail "Existing label: $current"
            $stats.Protected++
            continue
        }

        # ── WhatIf path ───────────────────────────────────────────────────────
        if ($WhatIfPreference -eq 'Continue') {
            Write-Host "  │  DRYRUN  $path" -ForegroundColor DarkCyan
            Add-LogEntry -Library $ListTitle -File $path -Result 'WhatIf'
            continue
        }

        # ── Queue into batch ──────────────────────────────────────────────────
        $addParams = @{
            Identity         = $path
            SensitivityLabelId = $LabelId
            AssignmentMethod = $AssignmentMethod
            Batch            = $batch
        }
        if ($JustificationText) { $addParams['JustificationText'] = $JustificationText }

        Add-PnPFileSensitivityLabel @addParams
        $batchCount++

        # ── Flush batch when it reaches BatchSize ─────────────────────────────
        if ($batchCount -ge $BatchSize) {
            Write-Host "  │  Executing batch ($batchCount requests)..." -ForegroundColor DarkCyan
            try {
                Invoke-WithRetry -Action { Invoke-PnPBatch -Batch $batch }
                $stats.Ok += $batchCount
            }
            catch {
                Write-Host "  │  BATCH FAIL: $($_.Exception.Message)" -ForegroundColor Red
                Add-LogEntry -Library $ListTitle -File "(batch)" -Result 'Failed' -Detail $_.Exception.Message
                $stats.Failed += $batchCount
            }
            $batch      = New-PnPBatch
            $batchCount = 0
        }
    }

    # ── Flush remaining items ─────────────────────────────────────────────────
    if ($batchCount -gt 0) {
        Write-Host "  │  Executing final batch ($batchCount requests)..." -ForegroundColor DarkCyan
        try {
            Invoke-WithRetry -Action { Invoke-PnPBatch -Batch $batch }
            $stats.Ok += $batchCount
        }
        catch {
            Write-Host "  │  BATCH FAIL: $($_.Exception.Message)" -ForegroundColor Red
            Add-LogEntry -Library $ListTitle -File "(batch)" -Result 'Failed' -Detail $_.Exception.Message
            $stats.Failed += $batchCount
        }
    }

    Write-Progress -Id 1 -Completed
}

function Invoke-Library {
    param([string]$Title, [int]$Index, [int]$Total)

    Write-Progress -Id 0 `
                   -Activity 'Processing libraries' `
                   -Status   "[$Index / $Total]  $Title" `
                   -PercentComplete ([math]::Round(($Index - 1) / $Total * 100))

    Write-Host "  ┌─ Library: $Title" -ForegroundColor White

    $items = try {
        Invoke-WithRetry -Action {
            @(Get-PnPListItem -List $Title -PageSize 500 `
                -Fields 'ID','FileRef','FSObjType','_ComplianceTag' |
             Where-Object { $_['FSObjType'] -eq 0 })
        }
    }
    catch {
        Write-Host "  └─ FAIL  Could not retrieve items: $($_.Exception.Message)`n" -ForegroundColor Red
        $stats.Failed++
        return
    }

    $count = $items.Count
    Write-Host "  │  $count file(s) found" -ForegroundColor DarkGray

    if ($count -eq 0) {
        Write-Host "  └─ Empty, skipping.`n" -ForegroundColor DarkGray
        return
    }

    Invoke-LabelBatch -Items $items -ListTitle $Title -LibIndex $Index -LibTotal $Total

    Write-Host "  └─ Done`n" -ForegroundColor White
}

Write-Host "`n  Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
Write-Host "  Connected.`n" -ForegroundColor Green

switch ($PSCmdlet.ParameterSetName) {

    # ── Single file ───────────────────────────────────────────────────────────
    'File' {
        Write-Host "  Scope : Single file — $FilePath`n" -ForegroundColor Cyan

        $fileItem = try {
            Invoke-WithRetry -Action {
                Get-PnPFile -Url $FilePath -AsListItem -ErrorAction Stop
            }
        }
        catch {
            Write-Error "Could not retrieve file '$FilePath': $($_.Exception.Message)"
            Disconnect-PnPOnline; exit 1
        }

        $listTitle = (Get-PnPList -Identity $fileItem.ParentList.Id).Title
        Write-Host "  ┌─ Library: $listTitle" -ForegroundColor White

        Invoke-LabelBatch -Items @($fileItem) -ListTitle $listTitle -LibIndex 1 -LibTotal 1

        Write-Host "  └─ Done`n" -ForegroundColor White
    }

    # ── Single library ────────────────────────────────────────────────────────
    'Library' {
        Write-Host "  Scope : Single library — '$LibraryName'`n" -ForegroundColor Cyan

        $list = try {
            Get-PnPList -Identity $LibraryName `
                -Includes BaseTemplate,Hidden,RootFolder -ErrorAction Stop
        }
        catch {
            Write-Error "Library '$LibraryName' not found: $($_.Exception.Message)"
            Disconnect-PnPOnline; exit 1
        }

        if ($list.BaseTemplate -ne 101) {
            Write-Error "'$LibraryName' is not a document library (BaseTemplate $($list.BaseTemplate))."
            Disconnect-PnPOnline; exit 1
        }

        Invoke-Library -Title $LibraryName -Index 1 -Total 1
        Write-Progress -Id 0 -Completed
    }

    # ── Entire site collection ────────────────────────────────────────────────
    default {
        Write-Host "  Scope : Entire site collection" -ForegroundColor Cyan
        Write-Host "  Discovering document libraries ..." -ForegroundColor Cyan

        $libraries = @(
            Get-PnPList -Includes BaseTemplate,Hidden,RootFolder |
            Where-Object { -not (Test-IsSystemLibrary $_) }
        )

        if ($libraries.Count -eq 0) {
            Write-Host "  No eligible document libraries found. Exiting." -ForegroundColor Yellow
            Disconnect-PnPOnline; exit 0
        }

        $libCount = $libraries.Count
        Write-Host "  Found $libCount librar$(if ($libCount -eq 1) {'y'} else {'ies'}):`n" -ForegroundColor White
        $libraries | ForEach-Object { Write-Host "    • $($_.Title)" -ForegroundColor DarkCyan }
        Write-Host ''

        $i = 0
        foreach ($lib in $libraries) {
            $i++
            Invoke-Library -Title $lib.Title -Index $i -Total $libCount
        }

        Write-Progress -Id 0 -Completed
    }
}

if ($LogPath -and $logEntries.Count -gt 0) {
    $logEntries | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Log written → $LogPath`n" -ForegroundColor DarkCyan
}

Write-Host "══════════════════════════════════════════" -ForegroundColor White
Write-Host "  SUMMARY" -ForegroundColor White
Write-Host "  Labelled  : $($stats.Ok)"        -ForegroundColor Green
Write-Host "  Skipped   : $($stats.Skipped)"   -ForegroundColor DarkGray
Write-Host "  Protected : $($stats.Protected)" -ForegroundColor Yellow
Write-Host "  Failed    : $($stats.Failed)"    -ForegroundColor $(if ($stats.Failed) { 'Red' } else { 'DarkGray' })
Write-Host "══════════════════════════════════════════`n" -ForegroundColor White

Disconnect-PnPOnline