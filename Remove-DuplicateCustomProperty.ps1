<#
.SYNOPSIS
    Removes duplicate custom properties from docProps/custom.xml inside
    Word .docx files on SharePoint or locally.

.DESCRIPTION
    SharePoint can sometimes create duplicate entries in docProps/custom.xml
    when a property's type changes (e.g. filetime → lpwstr).  Word may
    ignore one of them or behave unpredictably.

    This script:
      1. Opens the .docx (in memory — no temp files).
      2. Parses docProps/custom.xml.
      3. For each property name that appears more than once, keeps only the
         FIRST occurrence (lowest pid — the original) and removes later duplicates.
      4. Re-numbers pids sequentially starting at 2 (Office requirement).
      5. Re-serializes with Indent=$false (byte-safe round-trip).
      6. Uploads back / saves locally.

    By default it targets "ACTQMSApprovedDate" only.  Use -PropertyName to
    specify a different (or multiple) property name(s), or use -All to
    de-duplicate every property that has duplicates.

.PARAMETER SiteUrl
    Full URL of the SharePoint site.  Required for SharePoint mode.

.PARAMETER ClientId
    Client ID of your Entra ID app registration.  Required for SharePoint mode.

.PARAMETER FileServerRelativeUrl
    Server-relative path of a single .docx file on SharePoint.

.PARAMETER LibraryName
    Name of a document library — processes all .docx files.

.PARAMETER LocalPath
    Path to a local .docx file or folder of .docx files.

.PARAMETER Overwrite
    When using -LocalPath, overwrite the original file instead of creating
    a .deduped.docx copy.

.PARAMETER PropertyName
    One or more property names to de-duplicate.
    Default: @("ACTQMSApprovedDate")

.PARAMETER All
    De-duplicate ALL properties that have more than one entry (ignores -PropertyName).

.PARAMETER FileExtensionFilter
    Extensions to include when using -LibraryName. Default: "docx".

.PARAMETER PageSize
    Items per page when enumerating a SharePoint library. Default: 500.

.EXAMPLE
    # Remove duplicate ACTQMSApprovedDate from a single SharePoint file
    .\Remove-DuplicateCustomProperty.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/docs" `
        -ClientId "your-client-id" `
        -FileServerRelativeUrl "/sites/docs/Shared Documents/Report.docx"

.EXAMPLE
    # Remove ALL duplicate custom properties from every .docx in a library
    .\Remove-DuplicateCustomProperty.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/docs" `
        -ClientId "your-client-id" `
        -LibraryName "Shared Documents" `
        -All

.EXAMPLE
    # Remove duplicate ACTQMSApprovedDate from a local file (in-place)
    .\Remove-DuplicateCustomProperty.ps1 -LocalPath "C:\Docs\Report.docx" -Overwrite
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = "SPSingleFile")]
param(
    [Parameter(Mandatory, ParameterSetName = "SPSingleFile")]
    [Parameter(Mandatory, ParameterSetName = "SPLibrary")]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory, ParameterSetName = "SPSingleFile")]
    [Parameter(Mandatory, ParameterSetName = "SPLibrary")]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [Parameter(Mandatory, ParameterSetName = "SPSingleFile")]
    [ValidateNotNullOrEmpty()]
    [string]$FileServerRelativeUrl,

    [Parameter(Mandatory, ParameterSetName = "SPLibrary")]
    [ValidateNotNullOrEmpty()]
    [string]$LibraryName,

    [Parameter(Mandatory, ParameterSetName = "Local")]
    [ValidateNotNullOrEmpty()]
    [string]$LocalPath,

    [Parameter(ParameterSetName = "Local")]
    [switch]$Overwrite,

    [string[]]$PropertyName = @("ACTQMSApprovedDate"),

    [switch]$All,

    [Parameter(ParameterSetName = "SPSingleFile")]
    [Parameter(ParameterSetName = "SPLibrary")]
    [switch]$SkipArchived,

    [Parameter(ParameterSetName = "SPLibrary")]
    [string]$FileExtensionFilter = "docx",

    [Parameter(ParameterSetName = "SPLibrary")]
    [ValidateRange(1, 5000)]
    [int]$PageSize = 500
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName "System.IO.Compression"
Add-Type -AssemblyName "System.IO.Compression.FileSystem"

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $col = switch ($Level) {
        "INFO"    { "Cyan"   }
        "WARN"    { "Yellow" }
        "ERROR"   { "Red"    }
        "SUCCESS" { "Green"  }
    }
    Write-Host "[$Level] $Message" -ForegroundColor $col
}

function Write-SectionHeader {
    param([string]$Title)
    Write-Host ""
    Write-Host "--- $Title ---" -ForegroundColor DarkGray
}

function ConvertTo-CleanXmlBytes {
    param([System.Xml.XmlDocument]$XmlDoc)

    $ms = New-Object System.IO.MemoryStream
    try {
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $settings.Indent = $false

        $writer = [System.Xml.XmlWriter]::Create($ms, $settings)
        try {
            $XmlDoc.Save($writer)
        }
        finally {
            $writer.Close()
        }

        return $ms.ToArray()
    }
    finally {
        $ms.Dispose()
    }
}

function Remove-DuplicatePropertiesFromBytes {
    <#
    .SYNOPSIS
        Takes raw .docx bytes, finds duplicate custom properties in
        docProps/custom.xml, keeps only the FIRST occurrence (lowest pid) of each,
        re-numbers pids, and returns modified bytes.

        Returns a hashtable:
          @{
              Bytes      = [byte[]]   (original bytes when unchanged)
              Changed    = [bool]
              Removed    = [string[]]   # list of "name (pid=N, type=T)"
          }
    #>
    param(
        [byte[]]$DocxBytes,
        [string[]]$TargetNames,
        [bool]$DedupeAll
    )

    $noChange = @{ Bytes = $DocxBytes; Changed = $false; Removed = @() }

    $scanMs = New-Object System.IO.MemoryStream(,$DocxBytes)
    try {
        $scanZip = New-Object System.IO.Compression.ZipArchive(
            $scanMs, [System.IO.Compression.ZipArchiveMode]::Read, $true)

        $customEntry = $scanZip.GetEntry("docProps/custom.xml")
        if (-not $customEntry) {
            Write-Host "    docProps/custom.xml not found — nothing to do"
            return $noChange
        }

        $stream = $customEntry.Open()
        $rawMs  = New-Object System.IO.MemoryStream
        try { $stream.CopyTo($rawMs) } finally { $stream.Close() }
        $rawBytes = $rawMs.ToArray()
        $rawMs.Dispose()
    }
    finally {
        $scanZip.Dispose()
        $scanMs.Dispose()
    }

    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.PreserveWhitespace = $true
    $rawMs2 = New-Object System.IO.MemoryStream(,$rawBytes)
    try {
        $reader = New-Object System.IO.StreamReader($rawMs2, [System.Text.Encoding]::UTF8, $true)
        try { $xmlDoc.Load($reader) } finally { $reader.Close() }
    }
    finally { $rawMs2.Dispose() }

    $opNs = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    $nsm  = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
    $nsm.AddNamespace("op", $opNs)

    $propertiesNode = $xmlDoc.SelectSingleNode("//op:Properties", $nsm)
    if (-not $propertiesNode) {
        Write-Host "    No <Properties> root element found"
        return $noChange
    }

    $allProps = @($propertiesNode.SelectNodes("op:property", $nsm))
    if ($allProps.Count -eq 0) {
        Write-Host "    No custom properties found"
        return $noChange
    }

    Write-Host "    Found $($allProps.Count) custom properties"

    $groups = @{}
    foreach ($prop in $allProps) {
        $name = $prop.GetAttribute("name")
        if (-not $groups.ContainsKey($name)) {
            $groups[$name] = [System.Collections.Generic.List[System.Xml.XmlElement]]::new()
        }
        $groups[$name].Add($prop)
    }

    $duplicateNames = @($groups.GetEnumerator() |
        Where-Object { $_.Value.Count -gt 1 } |
        ForEach-Object { $_.Key })

    if ($duplicateNames.Count -eq 0) {
        Write-Host "    No duplicate properties found"
        return $noChange
    }

    Write-Host "    Duplicate property names: $($duplicateNames -join ', ')" -ForegroundColor Yellow

    if (-not $DedupeAll) {
        $duplicateNames = @($duplicateNames | Where-Object { $_ -in $TargetNames })
        if ($duplicateNames.Count -eq 0) {
            Write-Host "    No targeted duplicates found (targets: $($TargetNames -join ', '))"
            return $noChange
        }
    }

    $removedList = [System.Collections.Generic.List[string]]::new()

    foreach ($name in $duplicateNames) {
        $entries = $groups[$name]
        $sorted = @($entries | Sort-Object { [int]$_.GetAttribute("pid") })
        $keeper = $sorted[0]
        $keeperPid  = $keeper.GetAttribute("pid")
        $keeperType = ""
        foreach ($child in $keeper.ChildNodes) {
            if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                $keeperType = $child.LocalName; break
            }
        }

        Write-Host "    Keeping: '$name' pid=$keeperPid ($keeperType)" -ForegroundColor Green

        for ($i = 1; $i -lt $sorted.Count; $i++) {
            $dup = $sorted[$i]
            $dupPid  = $dup.GetAttribute("pid")
            $dupType = ""
            foreach ($child in $dup.ChildNodes) {
                if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                    $dupType = $child.LocalName; break
                }
            }
            Write-Host "    Removing: '$name' pid=$dupPid ($dupType) value='$($dup.InnerText)'" -ForegroundColor Red

            $propertiesNode.RemoveChild($dup) | Out-Null
            $removedList.Add("$name (pid=$dupPid, type=$dupType)")
        }
    }

    if ($removedList.Count -eq 0) {
        return $noChange
    }

    $remainingProps = @($propertiesNode.SelectNodes("op:property", $nsm))
    $nextPid = 2
    foreach ($prop in $remainingProps) {
        [void]$prop.SetAttribute("pid", [string]$nextPid)
        $nextPid++
    }
    Write-Host "    Re-numbered $($remainingProps.Count) properties (pid 2..$($nextPid - 1))"

    $cleanBytes = ConvertTo-CleanXmlBytes -XmlDoc $xmlDoc

    $outputMs = New-Object System.IO.MemoryStream
    try {
        $outputMs.Write($DocxBytes, 0, $DocxBytes.Length)
        $outputMs.Position = 0

        $zip = New-Object System.IO.Compression.ZipArchive(
            $outputMs, [System.IO.Compression.ZipArchiveMode]::Update, $true)
        try {
            $zip.GetEntry("docProps/custom.xml").Delete()
            $newEntry = $zip.CreateEntry("docProps/custom.xml",
                [System.IO.Compression.CompressionLevel]::Optimal)
            $ws = $newEntry.Open()
            try { $ws.Write($cleanBytes, 0, $cleanBytes.Length) } finally { $ws.Close() }
        }
        finally { $zip.Dispose() }

        return @{
            Bytes   = $outputMs.ToArray()
            Changed = $true
            Removed = @($removedList)
        }
    }
    finally {
        $outputMs.Dispose()
    }
}

function Upload-DocxBytesBack {
    param(
        [byte[]]$Bytes,
        [string]$FileRef,
        [string]$WebUrl,
        [switch]$VersioningDisabled
    )

    $leafName   = [System.IO.Path]::GetFileName($FileRef)
    $folderFull = $FileRef.Substring(0, $FileRef.Length - $leafName.Length).TrimEnd("/")

    if ($WebUrl -ne "" -and $folderFull.StartsWith($WebUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
        $folderPath = $folderFull.Substring($WebUrl.Length).TrimStart("/")
    }
    else {
        $folderPath = $folderFull.TrimStart("/")
    }

    $spFile    = Get-PnPFile -Url $FileRef -AsListItem -ErrorAction Stop
    $uiVersion = [string]$spFile.FieldValues["_UIVersionString"]
    Write-Host "    Version     : $uiVersion"

    # Fail fast if already checked out by another user
    $checkedOutBy = $spFile.FieldValues["CheckoutUser"]
    if ($checkedOutBy) {
        $coUser = if ($checkedOutBy -is [Microsoft.SharePoint.Client.FieldUserValue]) { $checkedOutBy.LookupValue } else { [string]$checkedOutBy }
        $ctx = Get-PnPContext
        $currentLogin = $ctx.Web.CurrentUser.LoginName
        $coLogin = if ($checkedOutBy -is [Microsoft.SharePoint.Client.FieldUserValue]) { $checkedOutBy.Email } else { "" }
        if ($coUser -and $coLogin -ne $currentLogin) {
            throw "File is already checked out by '$coUser'. Cannot proceed."
        }
    }

    # -------------------------------------------------------------------
    # Two strategies depending on whether versioning is already disabled:
    #
    # A) VersioningDisabled = $true  (caller disabled versioning for batch)
    #    → Direct upload, no checkout/checkin needed. The file content is
    #      silently replaced because SharePoint won't track a new version.
    #
    # B) VersioningDisabled = $false  (minor versions — normal path)
    #    → Add-PnPFile -Checkout -CheckInType OverwriteCheckIn
    #      PnP handles: CheckOut → Upload → CheckIn(OverwriteCheckIn)
    # -------------------------------------------------------------------
    $checkinOk = $false

    if ($VersioningDisabled) {
        # --- Strategy A: versioning is off, just upload directly ----------
        $ms = New-Object System.IO.MemoryStream(,$Bytes)
        try {
            Write-Host "    Strategy    : direct upload (versioning disabled)"
            Add-PnPFile -FileName $leafName -Folder $folderPath -Stream $ms -ErrorAction Stop | Out-Null
            Write-Host "    Upload      : file replaced"
            Write-Host "    Result      : $uiVersion (unchanged)"
            $checkinOk = $true
        }
        catch {
            Write-Host "    ERROR       : Upload failed — $($_.Exception.Message)" -ForegroundColor Red
        }
        finally {
            $ms.Dispose()
        }
    }
    else {
        # --- Strategy B: checkout → upload → OverwriteCheckIn ------------
        $ms = New-Object System.IO.MemoryStream(,$Bytes)
        try {
            Write-Host "    Strategy    : Add-PnPFile -Checkout -CheckInType OverwriteCheckIn"
            Add-PnPFile -FileName $leafName -Folder $folderPath -Stream $ms `
                -Checkout `
                -CheckInComment "Removed duplicate custom properties" `
                -CheckInType OverwriteCheckIn `
                -ErrorAction Stop | Out-Null
            Write-Host "    Upload      : file replaced & checked in"
        }
        catch {
            Write-Host "    ERROR       : Upload/checkin failed — $($_.Exception.Message)" -ForegroundColor Red
            # If checkout happened but upload/checkin failed, undo checkout
            try {
                $postFile = Get-PnPFile -Url $FileRef -AsListItem -ErrorAction SilentlyContinue
                if ($postFile -and $postFile.FieldValues["CheckoutUser"]) {
                    Write-Host "    UndoCheckout: reverting due to failure" -ForegroundColor Yellow
                    Undo-PnPFileCheckout -Url $FileRef -ErrorAction Stop
                }
            } catch { Write-Host "    UndoCheckout: Warning — $($_.Exception.Message)" -ForegroundColor Red }
            return $false
        }
        finally {
            $ms.Dispose()
        }

        # Verify the file is no longer checked out and version is unchanged
        try {
            $postFile = Get-PnPFile -Url $FileRef -AsListItem -ErrorAction Stop
            $postVersion = [string]$postFile.FieldValues["_UIVersionString"]
            $postCheckout = $postFile.FieldValues["CheckoutUser"]

            if ($postCheckout) {
                Write-Host "    WARNING     : File is STILL checked out after OverwriteCheckIn!" -ForegroundColor Yellow
                try {
                    Write-Host "    Fallback    : explicit Set-PnPFileCheckedIn -CheckinType OverwriteCheckIn"
                    Set-PnPFileCheckedIn -Url $FileRef -Comment "Removed duplicate custom properties" -CheckinType OverwriteCheckIn -ErrorAction Stop
                    $postFile2 = Get-PnPFile -Url $FileRef -AsListItem -ErrorAction Stop
                    $postVersion = [string]$postFile2.FieldValues["_UIVersionString"]
                    $postCheckout2 = $postFile2.FieldValues["CheckoutUser"]
                    if ($postCheckout2) {
                        Write-Host "    ERROR       : File still checked out after fallback!" -ForegroundColor Red
                        return $false
                    }
                }
                catch {
                    Write-Host "    ERROR       : Fallback checkin failed — $($_.Exception.Message)" -ForegroundColor Red
                    return $false
                }
            }

            if ($postVersion -ne $uiVersion) {
                Write-Host "    WARNING     : Version changed from $uiVersion to $postVersion" -ForegroundColor Yellow
            } else {
                Write-Host "    Result      : $uiVersion (unchanged)"
            }
            $checkinOk = $true
        }
        catch {
            Write-Host "    Verify      : Could not verify post-upload state — $($_.Exception.Message)" -ForegroundColor Yellow
            $checkinOk = $true  # Upload succeeded; verification is best-effort
        }
    }

    return $checkinOk
}

function Process-SingleSPFile {
    param(
        [string]$FileRef,
        [string]$WebUrl,
        [string[]]$TargetNames,
        [bool]$DedupeAll,
        [switch]$VersioningDisabled
    )

    $FileRef = [System.Uri]::UnescapeDataString($FileRef)

    if ($FileRef -notmatch '(?i)\.docx$') {
        Write-Status "Not a .docx file — skipping." "WARN"
        return [pscustomobject]@{ FileRef = $FileRef; Status = "Skipped-NotDocx"; Removed = @() }
    }

    Write-Status "Processing: $FileRef" "INFO"

    $ms = Get-PnPFile -Url $FileRef -AsMemoryStream -ErrorAction Stop
    $bytes = $ms.ToArray()
    $ms.Dispose()
    Write-Host "    Downloaded  : $([math]::Round($bytes.Length / 1KB, 1)) KB"

    if ($bytes.Length -lt 4 -or $bytes[0] -ne 0x50 -or $bytes[1] -ne 0x4B) {
        Write-Status "  Not a valid ZIP — skipping." "WARN"
        return [pscustomobject]@{ FileRef = $FileRef; Status = "Skipped-NotZip"; Removed = @() }
    }

    $result = Remove-DuplicatePropertiesFromBytes `
        -DocxBytes $bytes `
        -TargetNames $TargetNames `
        -DedupeAll $DedupeAll

    if (-not $result.Changed) {
        Write-Host "    No duplicates — no changes needed."
        return [pscustomobject]@{ FileRef = $FileRef; Status = "Clean"; Removed = @() }
    }

    if (-not $PSCmdlet.ShouldProcess($FileRef, "Remove duplicate custom properties")) {
        return [pscustomobject]@{ FileRef = $FileRef; Status = "WhatIf"; Removed = $result.Removed }
    }

    $ok = Upload-DocxBytesBack -Bytes $result.Bytes -FileRef $FileRef -WebUrl $WebUrl -VersioningDisabled:$VersioningDisabled

    if ($ok) {
        Write-Status "    Done — $($result.Removed.Count) duplicate(s) removed." "SUCCESS"
        return [pscustomobject]@{ FileRef = $FileRef; Status = "Fixed"; Removed = $result.Removed }
    }
    else {
        Write-Status "    Upload succeeded but checkin failed — file is checked out!" "ERROR"
        return [pscustomobject]@{ FileRef = $FileRef; Status = "Error: CheckinFailed"; Removed = $result.Removed }
    }
}

function Process-LocalFile {
    param(
        [string]$FullPath,
        [string[]]$TargetNames,
        [bool]$DedupeAll,
        [switch]$Overwrite
    )

    if ($FullPath -notmatch '(?i)\.docx$') {
        return [pscustomobject]@{ FileRef = $FullPath; Status = "Skipped-NotDocx"; Removed = @() }
    }

    $bytes = [System.IO.File]::ReadAllBytes($FullPath)
    Write-Host "    Read: $([math]::Round($bytes.Length / 1KB, 1)) KB"

    if ($bytes.Length -lt 4 -or $bytes[0] -ne 0x50 -or $bytes[1] -ne 0x4B) {
        Write-Status "  Not a valid ZIP — skipping." "WARN"
        return [pscustomobject]@{ FileRef = $FullPath; Status = "Skipped-NotZip"; Removed = @() }
    }

    $result = Remove-DuplicatePropertiesFromBytes `
        -DocxBytes $bytes `
        -TargetNames $TargetNames `
        -DedupeAll $DedupeAll

    if (-not $result.Changed) {
        Write-Host "    No duplicates — no changes needed."
        return [pscustomobject]@{ FileRef = $FullPath; Status = "Clean"; Removed = @() }
    }

    if ($Overwrite) {
        $outPath = $FullPath
    }
    else {
        $dir  = [System.IO.Path]::GetDirectoryName($FullPath)
        $name = [System.IO.Path]::GetFileNameWithoutExtension($FullPath)
        $outPath = Join-Path $dir "$name.deduped.docx"
    }

    if (-not $PSCmdlet.ShouldProcess($outPath, "Remove duplicate custom properties")) {
        return [pscustomobject]@{ FileRef = $FullPath; Status = "WhatIf"; Removed = $result.Removed }
    }

    [System.IO.File]::WriteAllBytes($outPath, $result.Bytes)
    Write-Host "    Saved: $outPath" -ForegroundColor Green

    return [pscustomobject]@{ FileRef = $FullPath; Status = "Fixed"; Removed = $result.Removed }
}


try {
    $results = [System.Collections.Generic.List[pscustomobject]]::new()

    $targetList = $PropertyName
    $dedupeAll  = [bool]$All

    if ($dedupeAll) {
        Write-Status "Mode: de-duplicate ALL properties with duplicates" "INFO"
    }
    else {
        Write-Status "Target properties: $($targetList -join ', ')" "INFO"
    }

    if ($PSCmdlet.ParameterSetName -eq "Local") {
        Write-SectionHeader "Local File — Remove Duplicate Custom Properties"

        if (Test-Path -Path $LocalPath -PathType Container) {
            $files = @(Get-ChildItem -Path $LocalPath -Filter "*.docx" -File -Recurse |
                Where-Object { $_.Name -notlike '~`$*' })
            Write-Status "Found $($files.Count) .docx file(s) in '$LocalPath'" "INFO"

            $i = 0
            foreach ($f in $files) {
                $i++
                Write-Progress -Activity "De-duplicating" -Status $f.Name -PercentComplete (100 * $i / [Math]::Max($files.Count, 1))
                try {
                    $localResult = Process-LocalFile -FullPath $f.FullName -TargetNames $targetList -DedupeAll $dedupeAll -Overwrite:$Overwrite
                    $results.Add($localResult)
                }
                catch {
                    Write-Status "ERROR on $($f.FullName): $($_.Exception.Message)" "ERROR"
                    $results.Add([pscustomobject]@{
                        FileRef = $f.FullName
                        Status  = "Error: $($_.Exception.Message)"
                        Removed = @()
                    })
                }
            }

            Write-Progress -Activity "De-duplicating" -Completed
        }
        else {
            $resolvedPath = (Resolve-Path $LocalPath).Path
            try {
                $localSingleResult = Process-LocalFile -FullPath $resolvedPath -TargetNames $targetList -DedupeAll $dedupeAll -Overwrite:$Overwrite
                $results.Add($localSingleResult)
            }
            catch {
                Write-Status "ERROR on ${resolvedPath}: $($_.Exception.Message)" "ERROR"
                $results.Add([pscustomobject]@{
                    FileRef = $resolvedPath
                    Status  = "Error: $($_.Exception.Message)"
                    Removed = @()
                })
            }
        }
    }
    else {
        Import-Module PnP.PowerShell -ErrorAction Stop

        Write-Status "Connecting to $SiteUrl" "INFO"
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
        Write-Status "Connected." "SUCCESS"

        $ctx = Get-PnPContext
        $ctx.Load($ctx.Web.CurrentUser)
        $null = Invoke-PnPQuery
        Write-Status "Current user    : $($ctx.Web.CurrentUser.Title)" "INFO"

        $cachedWebUrl = (Get-PnPWeb).ServerRelativeUrl.TrimEnd("/")

        if ($PSCmdlet.ParameterSetName -eq "SPSingleFile") {
            Write-SectionHeader "Single File — Remove Duplicate Custom Properties"

            $decodedRef  = [System.Uri]::UnescapeDataString($FileServerRelativeUrl)
            $skipThis = $false
            $spItem = $null
            if ($SkipArchived) {
                $spItem = Get-PnPFile -Url $decodedRef -AsListItem -ErrorAction Stop
                $docStatus = [string]$spItem.FieldValues["ACTQMSDocumentStatus"]
                if ($docStatus -eq "Archived") {
                    Write-Status "Skipping (Archived): $decodedRef" "WARN"
                    $results.Add([pscustomobject]@{ FileRef = $decodedRef; Status = "Skipped-Archived"; Removed = @() })
                    $skipThis = $true
                }
            }

            if (-not $skipThis) {
                # Detect if file is at a published major version (e.g. "5.0")
                # Reuse $spItem from SkipArchived check if available
                if (-not $spItem) {
                    $spItem = Get-PnPFile -Url $decodedRef -AsListItem -ErrorAction Stop
                }
                $ver = [string]$spItem.FieldValues["_UIVersionString"]
                $isMajor = ($ver -match '^\d+\.0$' -and $ver -ne '0.0')

                if ($isMajor) {
                    # Infer library from first path segment after web URL
                    $relPath = $decodedRef
                    if ($cachedWebUrl -ne "" -and $relPath.StartsWith($cachedWebUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $relPath = $relPath.Substring($cachedWebUrl.Length)
                    }
                    $relPath = $relPath.TrimStart("/")
                    $libRelPath = [System.Uri]::UnescapeDataString($relPath.Split("/")[0])
                    if (-not $libRelPath) {
                        throw "Could not infer library name from path '$decodedRef' (web='$cachedWebUrl')."
                    }

                    $list = Get-PnPList -Identity $libRelPath -ErrorAction Stop
                    if (-not $list) {
                        throw "Library '$libRelPath' not found (inferred from '$decodedRef')."
                    }
                    $ctx2 = Get-PnPContext
                    $ctx2.Load($list)
                    $null = Invoke-PnPQuery
                    $savedVersioning       = $list.EnableVersioning
                    $savedMinorVersions    = $list.EnableMinorVersions
                    $savedForceCheckout    = $list.ForceCheckout
                    $savedDraftVisibility  = $list.DraftVersionVisibility

                    Write-Status "Major version detected — temporarily disabling versioning" "INFO"
                    try {
                        Set-PnPList -Identity $libRelPath -EnableVersioning $false -ErrorAction Stop
                        if ($savedForceCheckout) {
                            $list.ForceCheckout = $false; $list.Update(); $null = Invoke-PnPQuery
                        }

                        $singleResult = Process-SingleSPFile -FileRef $decodedRef `
                            -WebUrl $cachedWebUrl `
                            -TargetNames $targetList -DedupeAll $dedupeAll `
                            -VersioningDisabled
                        $results.Add($singleResult)
                    }
                    finally {
                        try {
                            Set-PnPList -Identity $libRelPath `
                                -EnableVersioning $savedVersioning `
                                -EnableMinorVersions $savedMinorVersions `
                                -ErrorAction Stop
                            $list3 = Get-PnPList -Identity $libRelPath -ErrorAction Stop
                            if ($savedForceCheckout) {
                                $list3.ForceCheckout = $true
                            }
                            $list3.DraftVersionVisibility = $savedDraftVisibility
                            $list3.Update()
                            $null = Invoke-PnPQuery
                            Write-Status "Restored versioning settings" "INFO"
                        }
                        catch {
                            Write-Status "WARNING: Failed to restore versioning settings: $($_.Exception.Message)" "ERROR"
                        }
                    }
                }
                else {
                    $singleResult = Process-SingleSPFile -FileRef $decodedRef `
                        -WebUrl $cachedWebUrl `
                        -TargetNames $targetList -DedupeAll $dedupeAll
                    $results.Add($singleResult)
                }
            }
        }
        else {
            Write-SectionHeader "Library Sweep — $LibraryName"

            $camlQuery = "<View Scope='RecursiveAll'><Query><Where>" +
                "<Eq><FieldRef Name='File_x0020_Type'/><Value Type='Text'>$FileExtensionFilter</Value></Eq>" +
                "</Where></Query><RowLimit>$PageSize</RowLimit></View>"

            $items = @(Get-PnPListItem -List $LibraryName -Query $camlQuery -PageSize $PageSize)
            Write-Status "Found $($items.Count) .$FileExtensionFilter file(s)" "INFO"

            # Separate files into major-version and minor-version batches
            $majorItems = [System.Collections.Generic.List[object]]::new()
            $minorItems = [System.Collections.Generic.List[object]]::new()

            foreach ($item in $items) {
                $fileRef = $item.FieldValues["FileRef"]

                # Filter: skip lock files
                $leafName = [System.IO.Path]::GetFileName($fileRef)
                if ($leafName -like '~`$*') { continue }

                if ($SkipArchived) {
                    $docStatus = [string]$item.FieldValues["ACTQMSDocumentStatus"]
                    if ($docStatus -eq "Archived") {
                        Write-Status "Skipping (Archived): $fileRef" "WARN"
                        $results.Add([pscustomobject]@{ FileRef = $fileRef; Status = "Skipped-Archived"; Removed = @() })
                        continue
                    }
                }

                $ver = [string]$item.FieldValues["_UIVersionString"]
                $isMajor = ($ver -match '^\d+\.0$' -and $ver -ne '0.0')
                if ($isMajor) { $majorItems.Add($item) } else { $minorItems.Add($item) }
            }

            Write-Status "Batches: $($majorItems.Count) major version(s), $($minorItems.Count) minor version(s)" "INFO"
            $totalFiles = $majorItems.Count + $minorItems.Count
            $processed = 0

            # ---------------------------------------------------------------
            # PASS 1 — Major versions: disable versioning once, process all,
            #          then re-enable. No checkout/checkin needed.
            # ---------------------------------------------------------------
            if ($majorItems.Count -gt 0) {
                Write-SectionHeader "Pass 1 — Major versions (versioning disabled)"

                $list = Get-PnPList -Identity $LibraryName -ErrorAction Stop
                $ctx2 = Get-PnPContext
                $ctx2.Load($list)
                $null = Invoke-PnPQuery
                $savedVersioning      = $list.EnableVersioning
                $savedMinorVersions   = $list.EnableMinorVersions
                $savedForceCheckout   = $list.ForceCheckout
                $savedDraftVisibility = $list.DraftVersionVisibility

                Write-Status "Disabling versioning on '$LibraryName'" "INFO"
                try {
                    Set-PnPList -Identity $LibraryName -EnableVersioning $false -ErrorAction Stop
                    if ($savedForceCheckout) {
                        $list.ForceCheckout = $false; $list.Update(); $null = Invoke-PnPQuery
                    }
                    Write-Status "Versioning disabled — processing $($majorItems.Count) major version file(s)" "INFO"

                    foreach ($item in $majorItems) {
                        $processed++
                        $fileRef = $item.FieldValues["FileRef"]
                        Write-Progress -Activity "De-duplicating (major)" -Status $fileRef -PercentComplete (100 * $processed / [Math]::Max($totalFiles, 1))

                        try {
                            $libResult = Process-SingleSPFile -FileRef $fileRef `
                                -WebUrl $cachedWebUrl `
                                -TargetNames $targetList -DedupeAll $dedupeAll `
                                -VersioningDisabled
                            $results.Add($libResult)
                        }
                        catch {
                            $errMsg = $_.Exception.Message
                            Write-Status "ERROR on ${fileRef}: $errMsg" "ERROR"
                            $results.Add([pscustomobject]@{
                                FileRef = $fileRef
                                Status  = "Error: $errMsg"
                                Removed = @()
                            })
                        }
                    }
                }
                finally {
                    # Always restore original settings
                    try {
                        Set-PnPList -Identity $LibraryName `
                            -EnableVersioning $savedVersioning `
                            -EnableMinorVersions $savedMinorVersions `
                            -ErrorAction Stop
                        $list2 = Get-PnPList -Identity $LibraryName -ErrorAction Stop
                        if ($savedForceCheckout) {
                            $list2.ForceCheckout = $true
                        }
                        $list2.DraftVersionVisibility = $savedDraftVisibility
                        $list2.Update()
                        $null = Invoke-PnPQuery
                        Write-Status "Restored versioning settings on '$LibraryName'" "SUCCESS"
                    }
                    catch {
                        Write-Status "WARNING: Failed to restore versioning settings: $($_.Exception.Message)" "ERROR"
                    }
                }
            }

            # ---------------------------------------------------------------
            # PASS 2 — Minor versions: OverwriteCheckIn (no setting changes)
            # ---------------------------------------------------------------
            if ($minorItems.Count -gt 0) {
                Write-SectionHeader "Pass 2 — Minor versions (OverwriteCheckIn)"

                foreach ($item in $minorItems) {
                    $processed++
                    $fileRef = $item.FieldValues["FileRef"]
                    Write-Progress -Activity "De-duplicating (minor)" -Status $fileRef -PercentComplete (100 * $processed / [Math]::Max($totalFiles, 1))

                    try {
                        $libResult = Process-SingleSPFile -FileRef $fileRef `
                            -WebUrl $cachedWebUrl `
                            -TargetNames $targetList -DedupeAll $dedupeAll
                        $results.Add($libResult)
                    }
                    catch {
                        $errMsg = $_.Exception.Message
                        Write-Status "ERROR on ${fileRef}: $errMsg" "ERROR"
                        $results.Add([pscustomobject]@{
                            FileRef = $fileRef
                            Status  = "Error: $errMsg"
                            Removed = @()
                        })
                    }
                }
            }

            Write-Progress -Activity "De-duplicating" -Completed
        }
    }

    Write-SectionHeader "Summary"

    $cntFixed    = @($results | Where-Object { $_.Status -eq "Fixed" }).Count
    $cntClean    = @($results | Where-Object { $_.Status -eq "Clean" }).Count
    $cntArchived = @($results | Where-Object { $_.Status -eq "Skipped-Archived" }).Count
    $cntSkipped  = @($results | Where-Object { $_.Status -like "Skipped*" -and $_.Status -ne "Skipped-Archived" }).Count
    $cntWhatIf   = @($results | Where-Object { $_.Status -eq "WhatIf" }).Count
    $cntErrors   = @($results | Where-Object { $_.Status -like "Error*" }).Count

    $totalRemoved = 0
    foreach ($r in $results) {
        if ($r.Removed) { $totalRemoved += $r.Removed.Count }
    }

    Write-Host "  Fixed         : $cntFixed  ($totalRemoved duplicate(s) removed)"
    Write-Host "  Clean         : $cntClean  (no duplicates)"
    Write-Host "  Archived      : $cntArchived  (skipped)"
    Write-Host "  Skipped       : $cntSkipped"
    Write-Host "  WhatIf        : $cntWhatIf"
    Write-Host "  Errors        : $cntErrors"

    if ($cntFixed -gt 0) {
        Write-Host ""
        Write-Status "Files with duplicates removed:" "SUCCESS"
        foreach ($r in $results) {
            if ($r.Status -eq "Fixed") {
                Write-Host "  $($r.FileRef)" -ForegroundColor Green
                foreach ($rem in $r.Removed) {
                    Write-Host "    - $rem" -ForegroundColor DarkGreen
                }
            }
        }
    }

    if ($cntErrors -gt 0) {
        Write-Host ""
        Write-Status "Files with errors:" "ERROR"
        $results | Where-Object { $_.Status -like "Error*" } |
            ForEach-Object { Write-Host "  $($_.FileRef)  —  $($_.Status)" -ForegroundColor Red }
        Write-Host ""
        Write-Status "Completed with errors." "WARN"
        exit 1
    }

    Write-Host ""
    Write-Status "Done." "SUCCESS"
}
catch {
    Write-Host ""
    Write-Status $_.Exception.Message "ERROR"
    Write-Host ""
    exit 1
}
