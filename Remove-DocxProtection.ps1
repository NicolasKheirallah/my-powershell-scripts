<#
.SYNOPSIS
    Removes document protection (forms / read-only / tracked-changes lock)
    from Word .docx files on SharePoint or locally.

.DESCRIPTION
    Word documents can have a <w:documentProtection> element in
    word/settings.xml that locks editing.  Typical modes include:

      • forms        — only form fields (SDTs / content controls) are editable
      • readOnly     — entire document is read-only
      • trackedChanges — forces Track Changes on
      • comments     — only commenting is allowed
      • sections     — editing restricted to marked sections

    This script removes the <w:documentProtection> element entirely, which
    is equivalent to going to Review → Restrict Editing → Stop Protection in
    Word (with or without knowing the password).

    It does NOT crack the password — the protection element is simply deleted.
    Word stores only a hash, not the plaintext password, so removal is the
    standard approach.

    After removal the document opens fully editable.  The original formatting
    and content are untouched.

    ── How it works ──
      1. Downloads the .docx from SharePoint (or reads a local file).
      2. Opens the ZIP in memory.
      3. Parses word/settings.xml and removes <w:documentProtection>.
      4. Re-serializes with Indent=$false (byte-safe round-trip).
      5. Uploads back to SharePoint (no version increment) or saves locally.

    No temp files are written to disk.

.PARAMETER SiteUrl
    Full URL of the SharePoint site.  Required for SharePoint mode.

.PARAMETER ClientId
    Client ID of your Entra ID app registration.  Required for SharePoint mode.

.PARAMETER FileServerRelativeUrl
    Server-relative path of a single .docx file on SharePoint.
    Mutually exclusive with -LibraryName and -LocalPath.

.PARAMETER LibraryName
    Name of a document library — processes all .docx files.
    Mutually exclusive with -FileServerRelativeUrl and -LocalPath.

.PARAMETER LocalPath
    Path to a local .docx file or folder of .docx files.
    Saves alongside with a .unprotected.docx suffix (or overwrites with -Overwrite).
    Mutually exclusive with SharePoint parameters.

.PARAMETER Overwrite
    When using -LocalPath, overwrite the original file instead of creating
    a .unprotected.docx copy.

.PARAMETER FileExtensionFilter
    Extensions to include when using -LibraryName. Default: "docx".

.PARAMETER CheckoutConflict
    What to do when a SharePoint file is checked out by another user.
      Skip (default) | Abort

.PARAMETER PageSize
    Items per page when enumerating a SharePoint library. Default: 500.

.EXAMPLE
    # Remove protection from a single SharePoint file
    .\Remove-DocxProtection.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/docs" `
        -ClientId "your-client-id" `
        -FileServerRelativeUrl "/sites/docs/Shared Documents/Report.docx"

.EXAMPLE
    # Remove protection from all .docx files in a SharePoint library
    .\Remove-DocxProtection.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/docs" `
        -ClientId "your-client-id" `
        -LibraryName "Shared Documents"

.EXAMPLE
    # Remove protection from a local file (creates Report.unprotected.docx)
    .\Remove-DocxProtection.ps1 -LocalPath "C:\Docs\Report.docx"

.EXAMPLE
    # Remove protection from a local file in-place
    .\Remove-DocxProtection.ps1 -LocalPath "C:\Docs\Report.docx" -Overwrite
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

    [Parameter(ParameterSetName = "SPLibrary")]
    [string]$FileExtensionFilter = "docx",

    [Parameter(ParameterSetName = "SPSingleFile")]
    [Parameter(ParameterSetName = "SPLibrary")]
    [ValidateSet("Skip", "Abort")]
    [string]$CheckoutConflict = "Skip",

    [Parameter(ParameterSetName = "SPLibrary")]
    [ValidateRange(1, 5000)]
    [int]$PageSize = 500
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName "System.IO.Compression"
Add-Type -AssemblyName "System.IO.Compression.FileSystem"

# ─────────────────────────────────────────────────────────────────────────────
# Console helpers
# ─────────────────────────────────────────────────────────────────────────────

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

# ─────────────────────────────────────────────────────────────────────────────
# XML round-trip helper — UTF-8 without BOM, NO indentation
# ─────────────────────────────────────────────────────────────────────────────

function ConvertTo-CleanXmlBytes {
    param([System.Xml.XmlDocument]$XmlDoc)

    $ms = New-Object System.IO.MemoryStream
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

# ─────────────────────────────────────────────────────────────────────────────
# Core: remove <w:documentProtection> from word/settings.xml
# ─────────────────────────────────────────────────────────────────────────────

function Remove-ProtectionFromBytes {
    <#
    .SYNOPSIS
        Takes raw .docx bytes, removes <w:documentProtection> from
        word/settings.xml, and returns the modified bytes.

        Returns a hashtable:
          @{ Bytes = [byte[]]; WasProtected = [bool]; ProtectionType = [string] }
    #>
    param([byte[]]$DocxBytes)

    $inputMs  = New-Object System.IO.MemoryStream(,$DocxBytes)
    $outputMs = New-Object System.IO.MemoryStream

    try { $inputMs.CopyTo($outputMs) }
    finally { $inputMs.Dispose() }

    $outputMs.Position = 0

    $zip = New-Object System.IO.Compression.ZipArchive(
        $outputMs,
        [System.IO.Compression.ZipArchiveMode]::Update,
        $true
    )

    $wasProtected = $false
    $protectionType = ""

    try {
        $settingsEntry = $zip.GetEntry("word/settings.xml")
        if (-not $settingsEntry) {
            Write-Host "    word/settings.xml not found — nothing to do"
            return @{
                Bytes          = $DocxBytes
                WasProtected   = $false
                ProtectionType = ""
            }
        }

        # Read the settings XML
        $stream = $settingsEntry.Open()
        $rawMs = New-Object System.IO.MemoryStream
        try { $stream.CopyTo($rawMs) } finally { $stream.Close() }
        $rawBytes = $rawMs.ToArray()
        $rawMs.Dispose()

        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.PreserveWhitespace = $true
        $xmlDoc.LoadXml([System.Text.Encoding]::UTF8.GetString($rawBytes))

        $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        $nsm = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
        $nsm.AddNamespace("w", $wNs)

        # Find and remove <w:documentProtection>
        $protNode = $xmlDoc.SelectSingleNode("//w:documentProtection", $nsm)

        if (-not $protNode) {
            Write-Host "    No <w:documentProtection> element — document is not protected"
            return @{
                Bytes          = $DocxBytes
                WasProtected   = $false
                ProtectionType = ""
            }
        }

        # Capture the protection details before removal
        $editAttr = $protNode.GetAttribute("edit", $wNs)
        if ([string]::IsNullOrEmpty($editAttr)) { $editAttr = $protNode.GetAttribute("w:edit") }
        $enforcementAttr = $protNode.GetAttribute("enforcement", $wNs)
        if ([string]::IsNullOrEmpty($enforcementAttr)) { $enforcementAttr = $protNode.GetAttribute("w:enforcement") }

        if ($editAttr) { $protectionType = $editAttr } else { $protectionType = "unknown" }
        $enforced = ($enforcementAttr -eq "1" -or $enforcementAttr -eq "true")

        Write-Host "    Protection found:" -ForegroundColor Yellow
        Write-Host "      Type        : $protectionType"
        Write-Host "      Enforced    : $enforced"

        # Show all attributes for transparency
        foreach ($attr in $protNode.Attributes) {
            if ($attr.LocalName -notin @("edit", "enforcement")) {
                Write-Host "      $($attr.LocalName) : $($attr.Value)" -ForegroundColor Gray
            }
        }

        # Remove the element
        $protNode.ParentNode.RemoveChild($protNode) | Out-Null
        $wasProtected = $true

        Write-Host "    Removed <w:documentProtection> element" -ForegroundColor Green

        # Write back
        $cleanBytes = ConvertTo-CleanXmlBytes -XmlDoc $xmlDoc
        $settingsEntry.Delete()
        $newEntry = $zip.CreateEntry("word/settings.xml", [System.IO.Compression.CompressionLevel]::Optimal)
        $ws = $newEntry.Open()
        try { $ws.Write($cleanBytes, 0, $cleanBytes.Length) } finally { $ws.Close() }
    }
    finally {
        $zip.Dispose()
    }

    return @{
        Bytes          = $outputMs.ToArray()
        WasProtected   = $wasProtected
        ProtectionType = $protectionType
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SharePoint upload helper (no version increment)
# ─────────────────────────────────────────────────────────────────────────────

function Upload-DocxBytesBack {
    param(
        [byte[]]$Bytes,
        [string]$FileRef,
        [string]$LibraryName
    )

    # Decode any URL-encoded characters
    $FileRef = [System.Uri]::UnescapeDataString($FileRef)

    # Build a site-relative folder path by stripping the web's server-relative URL
    $webUrl     = (Get-PnPWeb).ServerRelativeUrl.TrimEnd("/")
    $leafName   = [System.IO.Path]::GetFileName($FileRef)
    $folderFull = $FileRef.Substring(0, $FileRef.Length - $leafName.Length).TrimEnd("/")
    if ($webUrl -ne "" -and $folderFull.StartsWith($webUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
        $folderPath = $folderFull.Substring($webUrl.Length).TrimStart("/")
    }
    else {
        $folderPath = $folderFull.TrimStart("/")
    }

    # Resolve library name from the folder path if not provided
    if (-not $LibraryName) {
        $LibraryName = $folderPath.Split("/")[0]
    }

    # Get list settings
    $list = Get-PnPList -Identity $LibraryName -ErrorAction Stop
    $ctx  = Get-PnPContext
    $versioningWasOn         = [bool]$list.EnableVersioning
    $forceCheckoutWasOn      = [bool]$list.ForceCheckout

    $spFile = Get-PnPFile -Url $FileRef -AsListItem -ErrorAction Stop
    $uiVersion = [string]$spFile.FieldValues["_UIVersionString"]
    Write-Host "    Version     : $uiVersion"

    $isMajor = ($uiVersion -match '\.(\d+)$') -and ([int]$Matches[1] -eq 0)
    $checkedOut = $false

    try {
        if ($forceCheckoutWasOn) {
            Write-Host "    Checkout    : checking out file"
            Set-PnPFileCheckedOut -Url $FileRef -ErrorAction Stop
            $checkedOut = $true
        }

        if ($isMajor) {
            $listChanged = $false
            if ($versioningWasOn) {
                $list.EnableVersioning    = $false
                $list.EnableMinorVersions = $false
                $listChanged = $true
            }
            if ($forceCheckoutWasOn) {
                $list.ForceCheckout = $false
                $listChanged = $true
            }
            if ($listChanged) {
                Write-Host "    Versioning  : temporarily disabling (major version)"
                $list.Update()
                $ctx.ExecuteQuery()
            }
        }
        elseif ($forceCheckoutWasOn) {
            $list.ForceCheckout = $false
            Write-Host "    ForceCheckout: temporarily disabling (minor version)"
            $list.Update()
            $ctx.ExecuteQuery()
        }

        $ms = New-Object System.IO.MemoryStream(,$Bytes)
        try {
            Add-PnPFile -FileName $leafName -Folder $folderPath -Stream $ms -ErrorAction Stop | Out-Null
            if ($isMajor) {
                Write-Host "    Upload      : overwritten (no version increment)"
            } else {
                Write-Host "    Upload      : saved (minor version increment)"
            }
        }
        finally {
            $ms.Dispose()
        }
    }
    finally {
        if ($checkedOut) {
            if ($isMajor) {
                Write-Host "    Checkin     : checking in file (overwrite)"
                Set-PnPFileCheckedIn -Url $FileRef -Comment "Removed document protection" -CheckinType OverwriteCheckIn -ErrorAction Stop
            } else {
                Write-Host "    Checkin     : checking in file (minor)"
                Set-PnPFileCheckedIn -Url $FileRef -Comment "Removed document protection" -CheckinType MinorCheckIn -ErrorAction Stop
            }
        }

        $restoreNeeded = $false
        if ($isMajor -and $versioningWasOn) {
            $list.EnableVersioning    = $true
            $list.EnableMinorVersions = $true
            $list.DraftVersionVisibility = [Microsoft.SharePoint.Client.DraftVisibilityType]::Author
            $restoreNeeded = $true
        }
        if ($forceCheckoutWasOn) {
            $list.ForceCheckout = $true
            $restoreNeeded = $true
        }
        if ($restoreNeeded) {
            Write-Host "    Versioning  : restored"
            $list.Update()
            $ctx.ExecuteQuery()
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Process a single SharePoint file
# ─────────────────────────────────────────────────────────────────────────────

function Process-SingleSPFile {
    param(
        [string]$FileRef,
        [string]$LibraryName
    )

    # Decode any URL-encoded characters (e.g. %20 → space)
    $FileRef = [System.Uri]::UnescapeDataString($FileRef)

    if ($FileRef -notmatch '\.docx$') {
        Write-Status "Not a .docx file — skipping." "WARN"
        return [pscustomobject]@{ FileRef = $FileRef; Status = "Skipped-NotDocx"; ProtectionType = "" }
    }

    Write-Status "Processing: $FileRef" "INFO"

    # Check for checkout conflicts
    $file = Get-PnPFile -Url $FileRef -AsListItem -ErrorAction Stop
    $checkedOutTo = $file.FieldValues["CheckoutUser"]
    if ($checkedOutTo) {
        $coUser = $checkedOutTo.LookupValue
        Write-Status "  Checked out to '$coUser'" "WARN"
        if ($CheckoutConflict -eq "Abort") {
            throw "File is checked out to '$coUser'. Aborting."
        }
        return [pscustomobject]@{ FileRef = $FileRef; Status = "Skipped-Checkout"; ProtectionType = "" }
    }

    # Download
    $ms = Get-PnPFile -Url $FileRef -AsMemoryStream -ErrorAction Stop
    $bytes = $ms.ToArray()
    $ms.Dispose()
    Write-Host "    Downloaded  : $([math]::Round($bytes.Length / 1KB, 1)) KB"

    # Verify it's a ZIP
    if ($bytes.Length -lt 4 -or $bytes[0] -ne 0x50 -or $bytes[1] -ne 0x4B) {
        Write-Status "  Not a valid ZIP — skipping." "WARN"
        return [pscustomobject]@{ FileRef = $FileRef; Status = "Skipped-NotZip"; ProtectionType = "" }
    }

    # Remove protection
    $result = Remove-ProtectionFromBytes -DocxBytes $bytes

    if (-not $result.WasProtected) {
        Write-Host "    Not protected — no changes needed."
        return [pscustomobject]@{ FileRef = $FileRef; Status = "Clean"; ProtectionType = "" }
    }

    if (-not $PSCmdlet.ShouldProcess($FileRef, "Remove document protection ($($result.ProtectionType))")) {
        return [pscustomobject]@{ FileRef = $FileRef; Status = "WhatIf"; ProtectionType = $result.ProtectionType }
    }

    # Upload
    Upload-DocxBytesBack -Bytes $result.Bytes -FileRef $FileRef -LibraryName $LibraryName
    Write-Status "    Done — protection removed ($($result.ProtectionType))." "SUCCESS"

    return [pscustomobject]@{ FileRef = $FileRef; Status = "Unprotected"; ProtectionType = $result.ProtectionType }
}

# ─────────────────────────────────────────────────────────────────────────────
# Process a single local file
# ─────────────────────────────────────────────────────────────────────────────

function Process-LocalFile {
    param([string]$FullPath)

    if ($FullPath -notmatch '\.docx$') {
        return [pscustomobject]@{ FileRef = $FullPath; Status = "Skipped-NotDocx"; ProtectionType = "" }
    }

    $bytes = [System.IO.File]::ReadAllBytes($FullPath)
    Write-Host "    Read: $([math]::Round($bytes.Length / 1KB, 1)) KB"

    if ($bytes.Length -lt 4 -or $bytes[0] -ne 0x50 -or $bytes[1] -ne 0x4B) {
        Write-Status "  Not a valid ZIP — skipping." "WARN"
        return [pscustomobject]@{ FileRef = $FullPath; Status = "Skipped-NotZip"; ProtectionType = "" }
    }

    $result = Remove-ProtectionFromBytes -DocxBytes $bytes

    if (-not $result.WasProtected) {
        Write-Host "    Not protected — no changes needed."
        return [pscustomobject]@{ FileRef = $FullPath; Status = "Clean"; ProtectionType = "" }
    }

    if ($Overwrite) {
        $outPath = $FullPath
    }
    else {
        $dir  = [System.IO.Path]::GetDirectoryName($FullPath)
        $name = [System.IO.Path]::GetFileNameWithoutExtension($FullPath)
        $outPath = Join-Path $dir "$name.unprotected.docx"
    }

    if (-not $PSCmdlet.ShouldProcess($outPath, "Remove document protection ($($result.ProtectionType))")) {
        return [pscustomobject]@{ FileRef = $FullPath; Status = "WhatIf"; ProtectionType = $result.ProtectionType }
    }

    [System.IO.File]::WriteAllBytes($outPath, $result.Bytes)
    Write-Host "    Saved: $outPath" -ForegroundColor Green

    return [pscustomobject]@{ FileRef = $FullPath; Status = "Unprotected"; ProtectionType = $result.ProtectionType }
}

# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

try {
    $results = [System.Collections.Generic.List[pscustomobject]]::new()

    # ── Local mode ──
    if ($PSCmdlet.ParameterSetName -eq "Local") {
        Write-SectionHeader "Local File Protection Removal"

        if (Test-Path -Path $LocalPath -PathType Container) {
            $files = @(Get-ChildItem -Path $LocalPath -Filter "*.docx" -File -Recurse)
            Write-Status "Found $($files.Count) .docx file(s) in '$LocalPath'" "INFO"

            $i = 0
            foreach ($f in $files) {
                $i++
                Write-Progress -Activity "Removing protection" -Status $f.Name -PercentComplete (100 * $i / [Math]::Max($files.Count, 1))
                try {
                    $localResult = Process-LocalFile -FullPath $f.FullName
                    $results.Add($localResult)
                }
                catch {
                    Write-Status "ERROR on $($f.FullName): $($_.Exception.Message)" "ERROR"
                    $results.Add([pscustomobject]@{
                        FileRef        = $f.FullName
                        Status         = "Error: $($_.Exception.Message)"
                        ProtectionType = ""
                    })
                }
            }

            Write-Progress -Activity "Removing protection" -Completed
        }
        else {
            $localSingleResult = Process-LocalFile -FullPath (Resolve-Path $LocalPath).Path
            $results.Add($localSingleResult)
        }
    }
    else {
        # ── SharePoint mode — connect ──
        Import-Module PnP.PowerShell -ErrorAction Stop

        Write-Status "Connecting to $SiteUrl" "INFO"
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
        Write-Status "Connected." "SUCCESS"

        $ctx = Get-PnPContext
        $ctx.Load($ctx.Web.CurrentUser)
        Invoke-PnPQuery
        Write-Status "Current user    : $($ctx.Web.CurrentUser.Title)" "INFO"

        if ($PSCmdlet.ParameterSetName -eq "SPSingleFile") {
            # ── Single file ──
            Write-SectionHeader "Single File Protection Removal"
            $singleResult = Process-SingleSPFile -FileRef $FileServerRelativeUrl
            $results.Add($singleResult)
        }
        else {
            # ── Library sweep ──
            Write-SectionHeader "Library Protection Removal — $LibraryName"

            $camlQuery = "<View Scope='RecursiveAll'><Query><Where>" +
                "<Eq><FieldRef Name='File_x0020_Type'/><Value Type='Text'>$FileExtensionFilter</Value></Eq>" +
                "</Where></Query><RowLimit>$PageSize</RowLimit></View>"

            $items = @(Get-PnPListItem -List $LibraryName -Query $camlQuery -PageSize $PageSize)
            Write-Status "Found $($items.Count) .$FileExtensionFilter file(s)" "INFO"

            $i = 0
            foreach ($item in $items) {
                $i++
                $fileRef = $item.FieldValues["FileRef"]
                Write-Progress -Activity "Removing protection" -Status $fileRef -PercentComplete (100 * $i / [Math]::Max($items.Count, 1))

                try {
                    $libResult = Process-SingleSPFile -FileRef $fileRef -LibraryName $LibraryName
                    $results.Add($libResult)
                }
                catch {
                    $errMsg = $_.Exception.Message
                    Write-Status "ERROR on ${fileRef}: $errMsg" "ERROR"
                    $results.Add([pscustomobject]@{
                        FileRef        = $fileRef
                        Status         = "Error: $errMsg"
                        ProtectionType = ""
                    })
                }
            }

            Write-Progress -Activity "Removing protection" -Completed
        }
    }

    # ── Summary ──
    Write-SectionHeader "Summary"

    $cntUnprotected = @($results | Where-Object { $_.Status -eq "Unprotected" }).Count
    $cntClean       = @($results | Where-Object { $_.Status -eq "Clean" }).Count
    $cntSkipped     = @($results | Where-Object { $_.Status -like "Skipped*" }).Count
    $cntWhatIf      = @($results | Where-Object { $_.Status -eq "WhatIf" }).Count
    $cntErrors      = @($results | Where-Object { $_.Status -like "Error*" }).Count

    Write-Host "  Unprotected : $cntUnprotected"
    Write-Host "  Clean       : $cntClean  (no protection found)"
    Write-Host "  Skipped     : $cntSkipped"
    Write-Host "  WhatIf      : $cntWhatIf"
    Write-Host "  Errors      : $cntErrors"

    if ($cntUnprotected -gt 0) {
        Write-Host ""
        Write-Status "Files with protection removed:" "SUCCESS"
        $results | Where-Object { $_.Status -eq "Unprotected" } |
            ForEach-Object {
                Write-Host "  $($_.FileRef)  —  was: $($_.ProtectionType)" -ForegroundColor Green
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
