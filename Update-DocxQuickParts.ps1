<#
.SYNOPSIS
    Syncs Word Quick Parts and SharePoint-bound content controls with the latest
    SharePoint metadata and makes them refresh dynamically when the document opens.

.DESCRIPTION
    This script processes Word .docx files either from SharePoint Online or locally.

    For each document it:
      1. Opens the .docx in memory.
        2. Detects both classic DOCPROPERTY field references and SharePoint-bound
            content controls used in the document, including main document, headers,
            footers, footnotes, and endnotes.
      3. In SharePoint mode, reads the current list item metadata and maps matching
         SharePoint columns to custom document properties.
      4. Removes duplicate custom properties for the same name, keeping the lowest pid.
      5. Updates matching custom properties in docProps/custom.xml.
        6. Updates bound customXml metadata stores for SharePoint content controls.
        7. Marks Quick Part fields as dirty and enables update-on-open in word/settings.xml,
            so Word refreshes them dynamically.
        8. Uploads back to SharePoint or saves locally.

.PARAMETER SiteUrl
    Full URL of the SharePoint site. Required for direct SharePoint modes.
    When used without -FileServerRelativeUrl, -LibraryName, or
    -AllSitesAndLibraries, the script processes all visible document libraries on
    the specified site.
    When combined with -IncludeSubsites in site-sweep mode, the script also
    traverses discovered subsites and their visible document libraries.
    When using -AllSitesAndLibraries, provide any site URL in the target tenant so
    the script can derive the tenant admin endpoint.

.PARAMETER IncludeSubsites
    When using -SiteUrl without -FileServerRelativeUrl or -LibraryName, also
    process visible document libraries in all discovered subsites beneath the
    specified site.

.PARAMETER SiteLibraryCsvPath
    Path to a CSV file containing SharePoint sites and optional document library
    names to process.
    The CSV must contain a SiteURL column and may contain a DocLib column.
    When DocLib is blank for a site row, the script processes all visible document
    libraries on that site only.

.PARAMETER ClientId
    Client ID of your Entra ID app registration. Required for SharePoint mode.

.PARAMETER FileServerRelativeUrl
    Server-relative path of a single .docx file on SharePoint.

.PARAMETER LibraryName
    Name of a document library. Processes all .docx files in the library.

.PARAMETER AllSitesAndLibraries
    Traverse all non-personal site collections in the tenant, all discovered
    subsites, and each visible document library within them.

.PARAMETER LocalPath
    Path to a local .docx file or folder of .docx files.
    Local mode can mark fields for refresh and optionally update property values via
    -PropertyValue, but it cannot pull live SharePoint metadata.

.PARAMETER Overwrite
    When using -LocalPath, overwrite the original file instead of creating a
    .quickparts.docx copy.

.PARAMETER PropertyValue
    Optional hashtable of custom property values to apply in Local mode.
    Example: -PropertyValue @{ ACTQMSApprovedDate = '2026-03-17'; Owner = 'Nicolas' }

.PARAMETER FieldName
    SharePoint/internal field names to treat as authoritative metadata fields.
    Use this for an ad hoc list on a single run. When specified, it overrides
    both the built-in default field list and any profile loaded via
    -FieldConfigPath.

.PARAMETER FieldConfigPath
    Optional path to a reusable field configuration file.
    Supported formats:
      - .json: ["Title","ACTQMSApprovedDate"] or
               { "default": [...], "QMSCore": { "Fields": [...] } }
      - .psd1: @{ default = @('Title'); QMSCore = @{ Fields = @('Title') } }
      - .txt : one field name per line

.PARAMETER FieldProfile
    Named field profile to load from -FieldConfigPath when the config file
    contains multiple profiles. Default: default.

.PARAMETER SkipArchived
    Skip documents where ACTQMSDocumentStatus equals Archived.

.PARAMETER SkipCreatedAfter
    Skip documents whose Created date is on or after this date.
    Accepts a [datetime] value. Files created from this date onward are skipped.
    Example: -SkipCreatedAfter '2025-01-01'

.PARAMETER FileExtensionFilter
    Extensions to include when processing a site, a specific library, or the
    entire tenant. Default: docx.

.PARAMETER PageSize
    Items per page when enumerating a SharePoint library. Default: 500.
    Applies to both single-library and tenant sweep modes.

.PARAMETER EnableUpdateOnOpen
    Controls whether Word fields are configured to refresh when the document is
    opened. Default: $true.
    When set to $false, the script removes the update-on-open setting and clears
    dirty flags from tracked classic fields to reduce Word's field-update prompt.

.EXAMPLE
    # Single SharePoint file
    .\Update-DocxQuickParts.ps1 `
        -SiteUrl "https://avarante.sharepoint.com/sites/docs" `
        -ClientId "your-client-id" `
        -FileServerRelativeUrl "/sites/docs/Shared Documents/Report.docx"

.EXAMPLE
    # Entire library
    .\Update-DocxQuickParts.ps1 `
        -SiteUrl "https://avarante.sharepoint.com/sites/docs" `
        -ClientId "your-client-id" `
        -LibraryName "Shared Documents"

.EXAMPLE
    # Entire site: all visible document libraries
    .\Update-DocxQuickParts.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/docs" `
        -ClientId "your-client-id"

.EXAMPLE
    # Entire site and all subsites: all visible document libraries
    .\Update-DocxQuickParts.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/docs" `
        -ClientId "your-client-id" `
        -IncludeSubsites

.EXAMPLE
    # Entire tenant: all site collections, subsites, and visible document libraries
    .\Update-DocxQuickParts.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/docs" `
        -ClientId "your-client-id" `
        -AllSitesAndLibraries

.EXAMPLE
    # CSV-driven sweep: listed sites and optional libraries
    .\Update-DocxQuickParts.ps1 `
        -SiteLibraryCsvPath ".\sharepoint-sites.csv" `
        -ClientId "your-client-id"

.EXAMPLE
    # Local file, mark Quick Parts for refresh on open and update selected values
    .\Update-DocxQuickParts.ps1 `
        -LocalPath "C:\Docs\Report.docx" `
        -PropertyValue @{ ACTQMSApprovedDate = '2026-03-17' } `
        -Overwrite

.EXAMPLE
    # SharePoint file, but do not force Word to update fields on open
    .\Update-DocxQuickParts.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/docs" `
        -ClientId "your-client-id" `
        -FileServerRelativeUrl "/sites/docs/Shared Documents/Report.docx" `
        -EnableUpdateOnOpen:$false

.EXAMPLE
    # Use a reusable field profile from a config file
    .\Update-DocxQuickParts.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/docs" `
        -ClientId "your-client-id" `
        -LibraryName "Shared Documents" `
        -FieldConfigPath ".\Update-DocxQuickParts.fields.json" `
        -FieldProfile "QMSCore"
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'SPSiteLibraries')]
param(
    [Parameter(Mandatory, ParameterSetName = 'SPSiteLibraries')]
    [Parameter(Mandatory, ParameterSetName = 'SPSingleFile')]
    [Parameter(Mandatory, ParameterSetName = 'SPLibrary')]
    [Parameter(Mandatory, ParameterSetName = 'SPTenantSweep')]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory, ParameterSetName = 'SPSiteLibraries')]
    [Parameter(Mandatory, ParameterSetName = 'SPSingleFile')]
    [Parameter(Mandatory, ParameterSetName = 'SPLibrary')]
    [Parameter(Mandatory, ParameterSetName = 'SPTenantSweep')]
    [Parameter(Mandatory, ParameterSetName = 'SPCsvSweep')]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [Parameter(Mandatory, ParameterSetName = 'SPCsvSweep')]
    [ValidateNotNullOrEmpty()]
    [string]$SiteLibraryCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'SPSingleFile')]
    [ValidateNotNullOrEmpty()]
    [string]$FileServerRelativeUrl,

    [Parameter(Mandatory, ParameterSetName = 'SPLibrary')]
    [ValidateNotNullOrEmpty()]
    [string]$LibraryName,

    [Parameter(Mandatory, ParameterSetName = 'SPTenantSweep')]
    [switch]$AllSitesAndLibraries,

    [Parameter(Mandatory, ParameterSetName = 'Local')]
    [ValidateNotNullOrEmpty()]
    [string]$LocalPath,

    [Parameter(ParameterSetName = 'Local')]
    [switch]$Overwrite,

    [Parameter(ParameterSetName = 'Local')]
    [hashtable]$PropertyValue,

    [string[]]$FieldName,

    [string]$FieldConfigPath,

    [ValidateNotNullOrEmpty()]
    [string]$FieldProfile = 'default',

    [Parameter(ParameterSetName = 'SPSiteLibraries')]
    [Parameter(ParameterSetName = 'SPSingleFile')]
    [Parameter(ParameterSetName = 'SPLibrary')]
    [Parameter(ParameterSetName = 'SPTenantSweep')]
    [Parameter(ParameterSetName = 'SPCsvSweep')]
    [switch]$SkipArchived,

    [Parameter(ParameterSetName = 'SPSiteLibraries')]
    [Parameter(ParameterSetName = 'SPSingleFile')]
    [Parameter(ParameterSetName = 'SPLibrary')]
    [Parameter(ParameterSetName = 'SPTenantSweep')]
    [Parameter(ParameterSetName = 'SPCsvSweep')]
    [Nullable[datetime]]$SkipCreatedAfter = $null,

    [Parameter(ParameterSetName = 'SPSiteLibraries')]
    [Parameter(ParameterSetName = 'SPLibrary')]
    [Parameter(ParameterSetName = 'SPTenantSweep')]
    [Parameter(ParameterSetName = 'SPCsvSweep')]
    [string]$FileExtensionFilter = 'docx',

    [Parameter(ParameterSetName = 'SPSiteLibraries')]
    [Parameter(ParameterSetName = 'SPLibrary')]
    [Parameter(ParameterSetName = 'SPTenantSweep')]
    [Parameter(ParameterSetName = 'SPCsvSweep')]
    [ValidateRange(1, 5000)]
    [int]$PageSize = 500,

    [Parameter(ParameterSetName = 'SPSiteLibraries')]
    [Parameter(ParameterSetName = 'SPLibrary')]
    [Parameter(ParameterSetName = 'SPTenantSweep')]
    [Parameter(ParameterSetName = 'SPCsvSweep')]
    [switch]$RetryFailed,

    [Parameter(ParameterSetName = 'SPSiteLibraries')]
    [switch]$IncludeSubsites,

    [bool]$EnableUpdateOnOpen = $true,

    [string]$LogFilePath = ".\Update-DocxQuickParts_WarningsErrors.log"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName 'System.IO.Compression'
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'

# Cache for user resolution checks so we don't call Ensure-PnPUser repeatedly
# for the same identity across multiple files. Key = "id:<LookupId>", Value = $true/$false.
$script:resolvedUserCache = @{}
$script:sharePointAccessTokenCache = @{}
$script:CurrentLogContext = $null

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $color = switch ($Level) {
        'INFO' { 'Cyan' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        'SUCCESS' { 'Green' }
    }

    Write-Host "[$Level] $Message" -ForegroundColor $color

    if (($Level -eq 'WARN' -or $Level -eq 'ERROR') -and -not [string]::IsNullOrWhiteSpace($script:LogFilePath)) {
        try {
            $logMessage = $Message
            if (-not [string]::IsNullOrWhiteSpace($script:CurrentLogContext)) {
                $logMessage = "[$($script:CurrentLogContext)] $Message"
            }
            $logLine = "$([datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')) [$Level] $logMessage"
            Add-Content -Path $script:LogFilePath -Value $logLine -ErrorAction SilentlyContinue
        }
        catch {}
    }
}

function Get-SharePointUserIdentityLabel {
    param([object]$UserValue)

    if ($null -eq $UserValue) {
        return ''
    }

    foreach ($propertyName in @('Email', 'LookupValue', 'Title', 'LoginName', 'LookupId')) {
        $candidate = [string](Get-SafeObjectPropertyValue -InputObject $UserValue -PropertyName $propertyName)
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            return $candidate
        }
    }

    return [string]$UserValue
}

function Write-SectionHeader {
    param([string]$Title)
    Write-Host ''
    Write-Host "--- $Title ---" -ForegroundColor DarkGray
}

function Test-ResultRecord {
    param([object]$InputObject)

    if ($null -eq $InputObject) {
        return $false
    }

    foreach ($propertyName in @('FileRef', 'Status', 'QuickPartNames', 'Updated', 'Missing')) {
        if ($null -eq $InputObject.PSObject.Properties[$propertyName]) {
            return $false
        }
    }

    return $true
}

function Add-ResultRecords {
    param(
        [System.Collections.Generic.List[pscustomobject]]$Target,
        [AllowNull()][object]$InputObject,
        [string]$SourceDescription = 'operation'
    )

    if ($null -eq $Target -or $null -eq $InputObject) {
        return
    }

    $items = if ((Test-ResultRecord -InputObject $InputObject) -or $InputObject -is [string]) {
        @($InputObject)
    }
    elseif ($InputObject -is [System.Collections.IEnumerable]) {
        @($InputObject)
    }
    else {
        @($InputObject)
    }

    foreach ($item in $items) {
        if ($null -eq $item) {
            continue
        }

        if (-not (Test-ResultRecord -InputObject $item)) {
            Write-Status "Ignoring unexpected output from $SourceDescription ($($item.GetType().FullName))." 'WARN'
            continue
        }

        $Target.Add([pscustomobject]$item)
    }
}

function New-NameLookup {
    param([string[]]$Names)

    $lookup = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $Names) {
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $lookup.Add($name) | Out-Null
        }
    }

    return (, $lookup)
}

function Get-PropertyAliasMap {
    $aliases = @{}
    $aliases['_UIVersionString'] = @('DLCPolicyLabelValue')
    $aliases['Title'] = @('title')
    return $aliases
}

function Get-DefaultFieldNameList {
    return @(
        'ACTDocumentLanguage',
        'ACTLocation',
        'ACTOrganisation',
        'ACTQMSEndToEndProcess',
        'ACTQMSSteeringDocumentType',
        'ACTTypeOfProjectDocument',
        'ACTQMSApprovedBy',
        'ACTQMSFunction',
        'ACTQMSManagementSystem',
        'ACTQMSPeriodOfValidity',
        'ACTQMSDocumentStatus',
        'ACTQMSApprovedDate',
        'ACTQMSUnitCodeLocal',
        'ACTQMSArchivedBy',
        'ACTQMSExpiryDate',
        'ACTQMSDocumentOwner',
        'MigratedChapter',
        '_dlc_DocId',
        '_UIVersionString',
        'Title'
    )
}

function Get-ObjectMemberValue {
    param(
        [object]$InputObject,
        [string]$MemberName
    )

    if ($null -eq $InputObject -or [string]::IsNullOrWhiteSpace($MemberName)) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) {
            if ([string]::Equals([string]$key, $MemberName, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $InputObject[$key]
            }
        }

        return $null
    }

    foreach ($property in $InputObject.PSObject.Properties) {
        if ([string]::Equals($property.Name, $MemberName, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $property.Value
        }
    }

    return $null
}

function Get-ObjectMemberNames {
    param([object]$InputObject)

    if ($null -eq $InputObject) {
        return @()
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        return @($InputObject.Keys | ForEach-Object { [string]$_ } | Select-Object -Unique)
    }

    return @($InputObject.PSObject.Properties | ForEach-Object { $_.Name } | Select-Object -Unique)
}

function ConvertTo-FieldNameList {
    param(
        [object]$InputObject,
        [string]$SourceDescription = 'field list'
    )

    if ($null -eq $InputObject) {
        return @()
    }

    $fieldsValue = Get-ObjectMemberValue -InputObject $InputObject -MemberName 'Fields'
    if ($null -ne $fieldsValue) {
        $InputObject = $fieldsValue
    }

    $fieldNames = [System.Collections.Generic.List[string]]::new()

    if ($InputObject -is [string]) {
        if (-not [string]::IsNullOrWhiteSpace($InputObject)) {
            $fieldNames.Add($InputObject.Trim())
        }
    }
    elseif ($InputObject -is [System.Collections.IDictionary]) {
        throw "$SourceDescription must be a string array or expose a 'Fields' property containing a string array."
    }
    elseif ($InputObject -is [System.Collections.IEnumerable]) {
        foreach ($entry in $InputObject) {
            if ($null -eq $entry) {
                continue
            }

            if ($entry -isnot [string]) {
                throw "$SourceDescription must contain only field name strings."
            }

            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                $fieldNames.Add($entry.Trim())
            }
        }
    }
    else {
        throw "$SourceDescription must be a string array or expose a 'Fields' property containing a string array."
    }

    return @($fieldNames | Select-Object -Unique)
}

function Import-FieldConfiguration {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'FieldConfigPath cannot be empty.'
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Field config file was not found: $Path"
    }

    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    $extension = [System.IO.Path]::GetExtension($resolvedPath).ToLowerInvariant()

    switch ($extension) {
        '.json' {
            $rawContent = Get-Content -LiteralPath $resolvedPath -Raw -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($rawContent)) {
                throw "Field config file is empty: $resolvedPath"
            }

            $data = ConvertFrom-Json -InputObject $rawContent -ErrorAction Stop
        }
        '.psd1' {
            $data = Import-PowerShellDataFile -Path $resolvedPath
        }
        '.txt' {
            $data = @(
                Get-Content -LiteralPath $resolvedPath -ErrorAction Stop |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                ForEach-Object { $_.Trim() }
            )
        }
        default {
            throw "Unsupported field config format '$extension'. Use .json, .psd1, or .txt."
        }
    }

    return [pscustomobject]@{
        Path = $resolvedPath
        Data = $data
    }
}

function Get-SelectedFieldConfiguration {
    param(
        [string[]]$FieldName,
        [bool]$FieldNameWasProvided,
        [string]$FieldConfigPath,
        [string]$FieldProfile = 'default'
    )

    if ($FieldNameWasProvided) {
        $selectedFieldNames = ConvertTo-FieldNameList -InputObject $FieldName -SourceDescription 'FieldName'
        if ($selectedFieldNames.Count -eq 0) {
            throw 'FieldName did not contain any usable field names.'
        }

        return [pscustomobject]@{
            FieldNames        = $selectedFieldNames
            Source            = 'FieldName'
            SourceDescription = 'inline -FieldName parameter'
            Profile           = $null
            Path              = $null
        }
    }

    if ([string]::IsNullOrWhiteSpace($FieldConfigPath)) {
        $selectedFieldNames = Get-DefaultFieldNameList
        return [pscustomobject]@{
            FieldNames        = $selectedFieldNames
            Source            = 'Default'
            SourceDescription = 'built-in default field list'
            Profile           = 'default'
            Path              = $null
        }
    }

    $fieldConfig = Import-FieldConfiguration -Path $FieldConfigPath
    $configData = $fieldConfig.Data
    $profilesNode = Get-ObjectMemberValue -InputObject $configData -MemberName 'Profiles'
    if ($null -ne $profilesNode) {
        $configData = $profilesNode
    }

    $directFields = if ($null -eq $profilesNode) { Get-ObjectMemberValue -InputObject $configData -MemberName 'Fields' } else { $null }
    $selectedValue = $null
    $availableProfiles = @()

    if ($null -ne $directFields) {
        if (-not [string]::Equals($FieldProfile, 'default', [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Field profile '$FieldProfile' was requested, but field config '$($fieldConfig.Path)' only defines a single 'Fields' list."
        }

        $selectedValue = $directFields
    }
    elseif ($configData -is [string] -or ($configData -is [System.Collections.IEnumerable] -and $configData -isnot [System.Collections.IDictionary] -and $configData -isnot [pscustomobject])) {
        if (-not [string]::Equals($FieldProfile, 'default', [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Field profile '$FieldProfile' was requested, but field config '$($fieldConfig.Path)' contains only a single field list."
        }

        $selectedValue = $configData
    }
    else {
        $availableProfiles = Get-ObjectMemberNames -InputObject $configData
        $selectedValue = Get-ObjectMemberValue -InputObject $configData -MemberName $FieldProfile
        if ($null -eq $selectedValue) {
            $availableProfilesText = if ($availableProfiles.Count -gt 0) { $availableProfiles -join ', ' } else { 'none' }
            throw "Field profile '$FieldProfile' was not found in '$($fieldConfig.Path)'. Available profiles: $availableProfilesText"
        }
    }

    $selectedFieldNames = ConvertTo-FieldNameList -InputObject $selectedValue -SourceDescription "field config '$($fieldConfig.Path)' profile '$FieldProfile'"
    if ($selectedFieldNames.Count -eq 0) {
        throw "Field config '$($fieldConfig.Path)' profile '$FieldProfile' did not contain any usable field names."
    }

    return [pscustomobject]@{
        FieldNames        = $selectedFieldNames
        Source            = 'FieldConfig'
        SourceDescription = "field config '$($fieldConfig.Path)' profile '$FieldProfile'"
        Profile           = $FieldProfile
        Path              = $fieldConfig.Path
    }
}

function Expand-AllowedPropertyNames {
    param(
        [string[]]$Names,
        [hashtable]$AliasMap
    )

    $expandedNames = [System.Collections.Generic.List[string]]::new()
    foreach ($name in $Names) {
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $expandedNames.Add($name)
        if ($AliasMap.ContainsKey($name)) {
            foreach ($alias in $AliasMap[$name]) {
                if (-not [string]::IsNullOrWhiteSpace($alias)) {
                    $expandedNames.Add($alias)
                }
            }
        }
    }

    return @($expandedNames | Select-Object -Unique)
}

function ConvertTo-CleanXmlBytes {
    param([System.Xml.XmlDocument]$XmlDoc)

    $memoryStream = New-Object System.IO.MemoryStream
    try {
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $settings.Indent = $false

        $writer = [System.Xml.XmlWriter]::Create($memoryStream, $settings)
        try {
            $XmlDoc.Save($writer)
        }
        finally {
            $writer.Close()
        }

        return $memoryStream.ToArray()
    }
    finally {
        $memoryStream.Dispose()
    }
}

function Read-ZipEntryBytes {
    param([System.IO.Compression.ZipArchiveEntry]$Entry)

    $stream = $Entry.Open()
    $memoryStream = New-Object System.IO.MemoryStream
    try {
        $stream.CopyTo($memoryStream)
        return $memoryStream.ToArray()
    }
    finally {
        $stream.Close()
        $memoryStream.Dispose()
    }
}

function Set-ZipEntryBytes {
    param(
        [System.IO.Compression.ZipArchive]$ZipArchive,
        [string]$EntryName,
        [byte[]]$Bytes
    )

    $existingEntry = $ZipArchive.GetEntry($EntryName)
    if ($existingEntry) {
        $existingEntry.Delete()
    }

    $entry = $ZipArchive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $stream.Write($Bytes, 0, $Bytes.Length)
    }
    finally {
        $stream.Close()
    }
}

function Get-XmlDocumentFromBytes {
    param([byte[]]$Bytes)

    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.PreserveWhitespace = $true

    $memoryStream = New-Object System.IO.MemoryStream(, $Bytes)
    try {
        $reader = New-Object System.IO.StreamReader($memoryStream, [System.Text.Encoding]::UTF8, $true)
        try {
            $xmlDoc.Load($reader)
        }
        finally {
            $reader.Close()
        }
    }
    finally {
        $memoryStream.Dispose()
    }

    return $xmlDoc
}

function Get-QuickPartPropertyNamesFromInstruction {
    param([string]$InstructionText)

    if ([string]::IsNullOrWhiteSpace($InstructionText)) {
        return @()
    }

    $regexMatches = [System.Text.RegularExpressions.Regex]::Matches(
        $InstructionText,
        'DOCPROPERTY\s+"([^"]+)"|DOCPROPERTY\s+([^\s\\]+)',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($match in $regexMatches) {
        $name = if ($match.Groups[1].Success) { $match.Groups[1].Value } else { $match.Groups[2].Value }
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $names.Add($name.Trim())
        }
    }

    return @($names | Select-Object -Unique)
}

function Set-FieldDirtyOnParent {
    param(
        [System.Xml.XmlElement]$FieldNode,
        [string]$WordNamespace,
        [bool]$EnableUpdateOnOpen = $true
    )

    if ($FieldNode.LocalName -eq 'fldSimple') {
        if ($EnableUpdateOnOpen) {
            if ($FieldNode.GetAttribute('dirty', $WordNamespace) -eq 'true') {
                return $false
            }
            $null = $FieldNode.SetAttribute('dirty', $WordNamespace, 'true')
        }
        else {
            if (-not $FieldNode.HasAttribute('dirty', $WordNamespace)) {
                return $false
            }
            $FieldNode.RemoveAttribute('dirty', $WordNamespace) | Out-Null
        }
        return $true
    }

    if ($FieldNode.LocalName -eq 'fldChar') {
        if ($FieldNode.GetAttribute('fldCharType', $WordNamespace) -eq 'begin') {
            if ($EnableUpdateOnOpen) {
                if ($FieldNode.GetAttribute('dirty', $WordNamespace) -eq 'true') {
                    return $false
                }
                $null = $FieldNode.SetAttribute('dirty', $WordNamespace, 'true')
            }
            else {
                if (-not $FieldNode.HasAttribute('dirty', $WordNamespace)) {
                    return $false
                }
                $FieldNode.RemoveAttribute('dirty', $WordNamespace) | Out-Null
            }
            return $true
        }
    }

    return $false
}

function Get-NextXmlNodeInDocumentOrder {
    param([System.Xml.XmlNode]$Node)

    if (-not $Node) {
        return $null
    }

    if ($Node.FirstChild) {
        return $Node.FirstChild
    }

    $cursor = $Node
    while ($cursor) {
        if ($cursor.NextSibling) {
            return $cursor.NextSibling
        }
        $cursor = $cursor.ParentNode
    }

    return $null
}

function Get-ComplexFieldInstructionText {
    param(
        [System.Xml.XmlElement]$StartFieldChar,
        [string]$WordNamespace
    )

    $parts = [System.Collections.Generic.List[string]]::new()
    $nestedDepth = 0
    $cursor = Get-NextXmlNodeInDocumentOrder -Node $StartFieldChar
    $done = $false

    while ($cursor -and -not $done) {
        if ($cursor -is [System.Xml.XmlElement] -and $cursor.NamespaceURI -eq $WordNamespace) {
            if ($cursor.LocalName -eq 'fldChar') {
                $fieldCharType = $cursor.GetAttribute('fldCharType', $WordNamespace)
                switch ($fieldCharType) {
                    'begin' {
                        $nestedDepth++
                    }
                    'separate' {
                        if ($nestedDepth -eq 0) {
                            $done = $true
                        }
                    }
                    'end' {
                        if ($nestedDepth -eq 0) {
                            $done = $true
                        }
                        else {
                            $nestedDepth--
                        }
                    }
                }
            }
            elseif ($cursor.LocalName -eq 'instrText' -and $nestedDepth -eq 0) {
                $parts.Add($cursor.InnerText)
            }
        }

        if (-not $done) {
            $cursor = Get-NextXmlNodeInDocumentOrder -Node $cursor
        }
    }

    return ($parts -join '')
}

function Get-FieldDisplayValueFromInstruction {
    param(
        [string]$InstructionText,
        [object]$Value
    )

    $dateFormatMatch = [System.Text.RegularExpressions.Regex]::Match(
        $InstructionText,
        '\\@\s+"([^"]+)"',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    if ($dateFormatMatch.Success) {
        $dateValue = Get-NormalizedDateValue -Value $Value
        if ($dateValue) {
            return $dateValue.ToString($dateFormatMatch.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture)
        }
    }

    return [string](ConvertTo-PropertyValue -Value $Value)
}

function Set-TextNodesValue {
    param(
        [System.Xml.XmlNode[]]$TextNodes,
        [string]$Value
    )

    if (-not $TextNodes -or $TextNodes.Count -eq 0) {
        return $false
    }

    $alreadyMatches = (($TextNodes[0].InnerText) -eq $Value)
    for ($index = 1; $index -lt $TextNodes.Count; $index++) {
        if ($TextNodes[$index].InnerText -ne '') {
            $alreadyMatches = $false
            break
        }
    }

    if ($alreadyMatches) {
        return $false
    }

    $TextNodes[0].InnerText = $Value
    for ($index = 1; $index -lt $TextNodes.Count; $index++) {
        $TextNodes[$index].InnerText = ''
    }

    return $true
}

function Get-ComplexFieldResultTextNodes {
    param(
        [System.Xml.XmlElement]$StartFieldChar,
        [string]$WordNamespace
    )

    $textNodes = [System.Collections.Generic.List[System.Xml.XmlNode]]::new()
    $nestedDepth = 0
    $inResult = $false
    $cursor = Get-NextXmlNodeInDocumentOrder -Node $StartFieldChar

    while ($cursor) {
        if ($cursor -is [System.Xml.XmlElement] -and $cursor.NamespaceURI -eq $WordNamespace) {
            if ($cursor.LocalName -eq 'fldChar') {
                $fieldCharType = $cursor.GetAttribute('fldCharType', $WordNamespace)
                switch ($fieldCharType) {
                    'begin' {
                        if ($inResult) {
                            $nestedDepth++
                        }
                    }
                    'separate' {
                        if ($nestedDepth -eq 0) {
                            $inResult = $true
                        }
                    }
                    'end' {
                        if ($nestedDepth -eq 0) {
                            break
                        }
                        $nestedDepth--
                    }
                }
            }
            elseif ($inResult -and $nestedDepth -eq 0 -and $cursor.LocalName -eq 't') {
                $textNodes.Add($cursor)
            }
        }

        $cursor = Get-NextXmlNodeInDocumentOrder -Node $cursor
    }

    return @($textNodes)
}

function Update-WordSettingsDocument {
    param(
        [System.Xml.XmlDocument]$XmlDoc,
        [bool]$EnableUpdateOnOpen = $true
    )

    $wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $nsManager = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
    $nsManager.AddNamespace('w', $wordNs)

    $settingsNode = $XmlDoc.SelectSingleNode('/w:settings', $nsManager)
    if (-not $settingsNode) {
        return $false
    }

    $updateFieldsNode = $XmlDoc.SelectSingleNode('/w:settings/w:updateFields', $nsManager)
    if ($EnableUpdateOnOpen) {
        if ($updateFieldsNode) {
            if ($updateFieldsNode.GetAttribute('val', $wordNs) -eq 'true') {
                return $false
            }
        }
        else {
            $updateFieldsNode = $XmlDoc.CreateElement('w', 'updateFields', $wordNs)
            $null = $settingsNode.AppendChild($updateFieldsNode)
        }

        $null = $updateFieldsNode.SetAttribute('val', $wordNs, 'true')
        return $true
    }

    if (-not $updateFieldsNode) {
        return $false
    }

    $null = $settingsNode.RemoveChild($updateFieldsNode)
    return $true
}

function Update-FieldDocument {
    param(
        [System.Xml.XmlDocument]$XmlDoc,
        [hashtable]$PropertyValues,
        [System.Collections.Generic.HashSet[string]]$AllowedNames,
        [bool]$EnableUpdateOnOpen = $true
    )

    $wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $nsManager = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
    $nsManager.AddNamespace('w', $wordNs)

    $foundNames = [System.Collections.Generic.List[string]]::new()
    $dirtyCount = 0
    $displayCount = 0

    $simpleFields = @($XmlDoc.SelectNodes('//w:fldSimple[contains(translate(@w:instr, ''abcdefghijklmnopqrstuvwxyz'', ''ABCDEFGHIJKLMNOPQRSTUVWXYZ''), ''DOCPROPERTY'')]', $nsManager))
    foreach ($field in $simpleFields) {
        $instruction = $field.GetAttribute('instr', $wordNs)
        $matchedNames = [System.Collections.Generic.List[string]]::new()
        foreach ($name in (Get-QuickPartPropertyNamesFromInstruction -InstructionText $instruction)) {
            if ($AllowedNames.Contains($name)) {
                $foundNames.Add($name)
                $matchedNames.Add($name)
            }
        }
        if ($matchedNames.Count -eq 0) {
            continue
        }

        foreach ($name in $matchedNames) {
            if (-not $PropertyValues.ContainsKey($name)) {
                continue
            }

            $displayValue = Get-FieldDisplayValueFromInstruction -InstructionText $instruction -Value $PropertyValues[$name]
            $textNodes = @($field.SelectNodes('.//w:t', $nsManager))
            if (Set-TextNodesValue -TextNodes $textNodes -Value $displayValue) {
                $displayCount++
                break
            }
        }

        if (Set-FieldDirtyOnParent -FieldNode $field -WordNamespace $wordNs -EnableUpdateOnOpen:$EnableUpdateOnOpen) {
            $dirtyCount++
        }
    }

    $complexFieldBegins = @($XmlDoc.SelectNodes('//w:fldChar[@w:fldCharType=''begin'']', $nsManager))
    foreach ($fieldCharNode in $complexFieldBegins) {
        $instructionText = Get-ComplexFieldInstructionText -StartFieldChar $fieldCharNode -WordNamespace $wordNs
        $names = @(Get-QuickPartPropertyNamesFromInstruction -InstructionText $instructionText)
        if ($names.Count -eq 0) {
            continue
        }

        $matchedNames = [System.Collections.Generic.List[string]]::new()
        foreach ($name in $names) {
            if ($AllowedNames.Contains($name)) {
                $foundNames.Add($name)
                $matchedNames.Add($name)
            }
        }

        if ($matchedNames.Count -eq 0) {
            continue
        }

        foreach ($name in $matchedNames) {
            if (-not $PropertyValues.ContainsKey($name)) {
                continue
            }

            $displayValue = Get-FieldDisplayValueFromInstruction -InstructionText $instructionText -Value $PropertyValues[$name]
            $resultTextNodes = Get-ComplexFieldResultTextNodes -StartFieldChar $fieldCharNode -WordNamespace $wordNs
            if (Set-TextNodesValue -TextNodes $resultTextNodes -Value $displayValue) {
                $displayCount++
                break
            }
        }

        if (Set-FieldDirtyOnParent -FieldNode $fieldCharNode -WordNamespace $wordNs -EnableUpdateOnOpen:$EnableUpdateOnOpen) {
            $dirtyCount++
        }
    }

    return [pscustomobject]@{
        PropertyNames = @($foundNames | Select-Object -Unique)
        DirtyCount    = $dirtyCount
        DisplayCount  = $displayCount
    }
}

function Add-NamespaceMappingsFromString {
    param(
        [System.Xml.XmlNamespaceManager]$NamespaceManager,
        [string]$PrefixMappings
    )

    if ([string]::IsNullOrWhiteSpace($PrefixMappings)) {
        return
    }

    $regexMatches = [System.Text.RegularExpressions.Regex]::Matches(
        $PrefixMappings,
        "xmlns:([^=\s]+)='([^']+)'"
    )

    foreach ($match in $regexMatches) {
        $prefix = $match.Groups[1].Value
        $namespaceUri = $match.Groups[2].Value
        if (-not [string]::IsNullOrWhiteSpace($prefix) -and -not [string]::IsNullOrWhiteSpace($namespaceUri)) {
            $NamespaceManager.AddNamespace($prefix, $namespaceUri)
        }
    }
}

function ConvertTo-NamespaceAgnosticXPath {
    param([string]$XPath)

    if ([string]::IsNullOrWhiteSpace($XPath) -or $XPath -match 'local-name\(\)') {
        return $XPath
    }

    $rebuiltParts = [System.Collections.Generic.List[string]]::new()
    foreach ($part in [System.Text.RegularExpressions.Regex]::Split($XPath, '(/+)')) {
        if ([string]::IsNullOrEmpty($part)) {
            continue
        }

        if ($part -match '^/+$') {
            $rebuiltParts.Add($part)
            continue
        }

        if ($part -match "^\*\[local-name\(\)=") {
            $rebuiltParts.Add($part)
            continue
        }

        if ($part -match '^(?:[^:\[]+:)?([^\[]+)(\[\d+\])?$') {
            $rebuiltParts.Add("*[local-name()='$($Matches[1])']$($Matches[2])")
            continue
        }

        $rebuiltParts.Add($part)
    }

    return ($rebuiltParts -join '')
}

function Get-BindingSourceKey {
    param(
        [System.Xml.XmlElement]$SdtNode,
        [System.Xml.XmlNamespaceManager]$NamespaceManager,
        [string]$WordNamespace
    )

    $tagNode = $SdtNode.SelectSingleNode('w:sdtPr/w:tag', $NamespaceManager)
    if ($tagNode) {
        $tagValue = $tagNode.GetAttribute('val', $WordNamespace)
        if (-not [string]::IsNullOrWhiteSpace($tagValue)) {
            return $tagValue
        }
    }

    $bindingNode = $SdtNode.SelectSingleNode('w:sdtPr/w:dataBinding', $NamespaceManager)
    if (-not $bindingNode) {
        return $null
    }

    $xpath = $bindingNode.GetAttribute('xpath', $WordNamespace)
    if ([string]::IsNullOrWhiteSpace($xpath)) {
        return $null
    }

    $parts = @($xpath -split '/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $parts = @($parts | Where-Object { $_ -notmatch '^ns0:properties(\[\d+\])?$' -and $_ -notmatch '^documentManagement(\[\d+\])?$' })
    [array]::Reverse($parts)

    foreach ($part in $parts) {
        if ($part -match 'local-name\(\)=''([^'']+)''') {
            return $Matches[1]
        }

        $cleanPart = ($part -replace '\[\d+\]', '')
        if ($cleanPart -match '^[^:]+:(.+)$') {
            $cleanPart = $Matches[1]
        }

        if ($cleanPart -eq 'title' -and $xpath -match 'coreProperties') {
            return 'Title'
        }

        if ($cleanPart -notin @('DisplayName', 'UserInfo', 'Url', 'Description')) {
            return $cleanPart
        }
    }

    return $null
}

function ConvertTo-BoundXmlValue {
    param([object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [datetime]) {
        return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssK')
    }

    return [string](ConvertTo-PropertyValue -Value $Value)
}

function Get-NormalizedDateValue {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [datetime]) {
        return $Value.Date
    }

    if ($Value -is [datetimeoffset]) {
        return $Value.Date.Date
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $isoDateMatch = [System.Text.RegularExpressions.Regex]::Match(
        $text,
        '^\s*(\d{4}-\d{2}-\d{2})(?:[T\s].*)?$'
    )
    if ($isoDateMatch.Success) {
        return [datetime]::ParseExact(
            $isoDateMatch.Groups[1].Value,
            'yyyy-MM-dd',
            [System.Globalization.CultureInfo]::InvariantCulture
        )
    }

    $dateTimeOffsetValue = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse(
            $text,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind,
            [ref]$dateTimeOffsetValue
        )) {
        return $dateTimeOffsetValue.Date.Date
    }

    $dateTimeValue = [datetime]::MinValue
    if ([datetime]::TryParse(
            $text,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind,
            [ref]$dateTimeValue
        )) {
        return $dateTimeValue.Date
    }

    return $null
}

function Format-BoundDisplayValue {
    param(
        [System.Xml.XmlElement]$SdtNode,
        [System.Xml.XmlNamespaceManager]$NamespaceManager,
        [string]$WordNamespace,
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    $dateNode = $SdtNode.SelectSingleNode('w:sdtPr/w:date', $NamespaceManager)
    if ($dateNode) {
        $dateFormatNode = $dateNode.SelectSingleNode('w:dateFormat', $NamespaceManager)
        $formatString = if ($dateFormatNode) { $dateFormatNode.GetAttribute('val', $WordNamespace) } else { '' }
        $dateValue = Get-NormalizedDateValue -Value $Value
        if ($dateValue) {
            if (-not [string]::IsNullOrWhiteSpace($formatString)) {
                return $dateValue.ToString($formatString, [System.Globalization.CultureInfo]::InvariantCulture)
            }
            return $dateValue.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
        }
    }

    return [string](ConvertTo-PropertyValue -Value $Value)
}

function Set-ContentControlDisplayValue {
    param(
        [System.Xml.XmlElement]$SdtNode,
        [System.Xml.XmlNamespaceManager]$NamespaceManager,
        [string]$DisplayValue
    )

    $textNodes = @($SdtNode.SelectNodes('w:sdtContent//w:t', $NamespaceManager))
    if ($textNodes.Count -eq 0) {
        return $false
    }

    $alreadyMatches = ($textNodes[0].InnerText -eq $DisplayValue)
    for ($index = 1; $index -lt $textNodes.Count; $index++) {
        if ($textNodes[$index].InnerText -ne '') {
            $alreadyMatches = $false
            break
        }
    }

    if ($alreadyMatches) {
        return $false
    }

    $textNodes[0].InnerText = $DisplayValue
    for ($index = 1; $index -lt $textNodes.Count; $index++) {
        $textNodes[$index].InnerText = ''
    }

    return $true
}

function Test-WordXmlEntryMayContainTrackedContent {
    param(
        [string]$EntryName,
        [byte[]]$EntryBytes
    )

    if ($EntryName -eq 'word/settings.xml') {
        return $false
    }

    $entryText = [System.Text.Encoding]::UTF8.GetString($EntryBytes)
    return (
        $entryText.IndexOf('DOCPROPERTY', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
        $entryText.IndexOf('dataBinding', [System.StringComparison]::OrdinalIgnoreCase) -ge 0
    )
}

function Get-CustomXmlStoreMap {
    param([System.IO.Compression.ZipArchive]$ZipArchive)

    $map = @{}
    $itemPropsEntries = @($ZipArchive.Entries | Where-Object { $_.FullName -match '^customXml/itemProps\d+\.xml$' })

    foreach ($itemPropsEntry in $itemPropsEntries) {
        $propsDoc = Get-XmlDocumentFromBytes -Bytes (Read-ZipEntryBytes -Entry $itemPropsEntry)
        $nsManager = New-Object System.Xml.XmlNamespaceManager($propsDoc.NameTable)
        $nsManager.AddNamespace('ds', 'http://schemas.openxmlformats.org/officeDocument/2006/customXml')

        $dataStoreNode = $propsDoc.SelectSingleNode('/ds:datastoreItem', $nsManager)
        if (-not $dataStoreNode) {
            continue
        }

        $itemId = $dataStoreNode.GetAttribute('itemID', 'http://schemas.openxmlformats.org/officeDocument/2006/customXml')
        if ([string]::IsNullOrWhiteSpace($itemId)) {
            continue
        }

        $itemNumber = [System.Text.RegularExpressions.Regex]::Match($itemPropsEntry.FullName, 'itemProps(\d+)\.xml$').Groups[1].Value
        if ([string]::IsNullOrWhiteSpace($itemNumber)) {
            continue
        }

        $itemEntryName = "customXml/item$itemNumber.xml"
        if ($ZipArchive.GetEntry($itemEntryName)) {
            $map[$itemId] = $itemEntryName
        }
    }

    return $map
}

function Get-BoundStoreMap {
    param([System.IO.Compression.ZipArchive]$ZipArchive)

    $map = Get-CustomXmlStoreMap -ZipArchive $ZipArchive

    $builtInStoreMap = @{
        '{6C3C8BC8-F283-45AE-878A-BAB7291924A1}' = 'docProps/core.xml'
    }

    foreach ($storeItemId in $builtInStoreMap.Keys) {
        $entryName = $builtInStoreMap[$storeItemId]
        if ($ZipArchive.GetEntry($entryName)) {
            $map[$storeItemId] = $entryName
        }
    }

    return $map
}

function Update-BoundContentControlsDocument {
    param(
        [System.Xml.XmlDocument]$XmlDoc,
        [hashtable]$PropertyValues,
        [hashtable]$StoreMap,
        [hashtable]$StoreDocuments,
        [System.Collections.Generic.HashSet[string]]$AllowedNames
    )

    $wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $nsManager = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
    $nsManager.AddNamespace('w', $wordNs)

    $updatedNames = [System.Collections.Generic.List[string]]::new()
    $authoritativeNames = [System.Collections.Generic.List[string]]::new()
    $fallbackNames = [System.Collections.Generic.List[string]]::new()
    $missingNames = [System.Collections.Generic.List[string]]::new()
    $updatedStores = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $customPropertyNames = [System.Collections.Generic.List[string]]::new()
    $displayCount = 0

    $sdtNodes = @($XmlDoc.SelectNodes('//w:sdt[w:sdtPr/w:dataBinding]', $nsManager))
    foreach ($sdtNode in $sdtNodes) {
        $bindingNode = $sdtNode.SelectSingleNode('w:sdtPr/w:dataBinding', $nsManager)
        if (-not $bindingNode) {
            continue
        }

        $storeItemId = $bindingNode.GetAttribute('storeItemID', $wordNs)
        $xpath = $bindingNode.GetAttribute('xpath', $wordNs)
        $prefixMappings = $bindingNode.GetAttribute('prefixMappings', $wordNs)
        if ([string]::IsNullOrWhiteSpace($storeItemId) -or [string]::IsNullOrWhiteSpace($xpath)) {
            continue
        }

        $sourceKey = Get-BindingSourceKey -SdtNode $sdtNode -NamespaceManager $nsManager -WordNamespace $wordNs
        if ([string]::IsNullOrWhiteSpace($sourceKey)) {
            continue
        }

        if (-not $AllowedNames.Contains($sourceKey)) {
            continue
        }

        if (-not $StoreMap.ContainsKey($storeItemId)) {
            $missingNames.Add($sourceKey)
            continue
        }

        $storeEntryName = $StoreMap[$storeItemId]
        if (-not $StoreDocuments.ContainsKey($storeEntryName)) {
            continue
        }

        $storeDoc = $StoreDocuments[$storeEntryName]
        $storeNsManager = New-Object System.Xml.XmlNamespaceManager($storeDoc.NameTable)
        Add-NamespaceMappingsFromString -NamespaceManager $storeNsManager -PrefixMappings $prefixMappings

        $storeChanged = $false
        $documentChanged = $false
        $targetNode = $storeDoc.SelectSingleNode($xpath, $storeNsManager)
        if (-not $targetNode) {
            $fallbackXPath = ConvertTo-NamespaceAgnosticXPath -XPath $xpath
            if ($fallbackXPath -ne $xpath) {
                $targetNode = $storeDoc.SelectSingleNode($fallbackXPath, $storeNsManager)
                if ($targetNode) {
                    $null = $bindingNode.SetAttribute('xpath', $wordNs, $fallbackXPath)
                    $documentChanged = $true
                }
            }
        }
        if (-not $targetNode) {
            $missingNames.Add($sourceKey)
            continue
        }

        $hasAuthoritativeValue = $PropertyValues.ContainsKey($sourceKey)
        if ($hasAuthoritativeValue) {
            $rawValue = $PropertyValues[$sourceKey]
            $boundValue = ConvertTo-BoundXmlValue -Value $rawValue
            if ($targetNode.InnerText -ne $boundValue) {
                $targetNode.InnerText = $boundValue
                $storeChanged = $true
            }
            $authoritativeNames.Add($sourceKey)
        }
        else {
            $rawValue = $targetNode.InnerText
            $PropertyValues[$sourceKey] = $rawValue
            $fallbackNames.Add($sourceKey)
        }

        if ($targetNode -is [System.Xml.XmlElement]) {
            if ($targetNode.HasAttribute('nil', 'http://www.w3.org/2001/XMLSchema-instance')) {
                $targetNode.RemoveAttribute('nil', 'http://www.w3.org/2001/XMLSchema-instance') | Out-Null
                $storeChanged = $true
            }
        }

        $displayValue = Format-BoundDisplayValue -SdtNode $sdtNode -NamespaceManager $nsManager -WordNamespace $wordNs -Value $rawValue
        if (Set-ContentControlDisplayValue -SdtNode $sdtNode -NamespaceManager $nsManager -DisplayValue $displayValue) {
            $displayCount++
            $documentChanged = $true
        }

        $dateNode = $sdtNode.SelectSingleNode('w:sdtPr/w:date', $nsManager)
        $normalizedDateValue = if ($dateNode) { Get-NormalizedDateValue -Value $rawValue } else { $null }
        if ($dateNode -and $normalizedDateValue) {
            $fullDateValue = $normalizedDateValue.ToString('yyyy-MM-ddT00:00:00Z', [System.Globalization.CultureInfo]::InvariantCulture)
            if ($dateNode.GetAttribute('fullDate', $wordNs) -ne $fullDateValue) {
                $null = $dateNode.SetAttribute('fullDate', $wordNs, $fullDateValue)
                $documentChanged = $true
            }
        }

        if ($storeChanged) {
            $updatedStores.Add($storeEntryName) | Out-Null
        }
        if ($storeEntryName -match '^customXml/') {
            $customPropertyNames.Add($sourceKey)
        }
        if ($storeChanged -or $documentChanged) {
            $updatedNames.Add($sourceKey)
        }
    }

    return [pscustomobject]@{
        UpdatedNames        = @($updatedNames | Select-Object -Unique)
        AuthoritativeNames  = @($authoritativeNames | Select-Object -Unique)
        FallbackNames       = @($fallbackNames | Select-Object -Unique)
        MissingNames        = @($missingNames | Select-Object -Unique)
        UpdatedStores       = @($updatedStores)
        CustomPropertyNames = @($customPropertyNames | Select-Object -Unique)
        DisplayCount        = $displayCount
    }
}

function Get-SafeObjectPropertyValue {
    param(
        [object]$InputObject,
        [string]$PropertyName
    )

    if ($null -eq $InputObject -or [string]::IsNullOrWhiteSpace($PropertyName)) {
        return $null
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($property) {
        return $property.Value
    }

    return $null
}

function Test-ObjectTypeName {
    param(
        [object]$InputObject,
        [string[]]$TypeNames
    )

    if ($null -eq $InputObject -or -not $TypeNames) {
        return $false
    }

    $currentType = $InputObject.GetType()
    while ($currentType) {
        foreach ($typeName in $TypeNames) {
            if (-not [string]::IsNullOrWhiteSpace($typeName) -and ($currentType.FullName -eq $typeName -or $currentType.Name -eq $typeName)) {
                return $true
            }
        }
        $currentType = $currentType.BaseType
    }

    return $false
}

function Test-SharePointUserResolvable {
    param(
        [string]$Email
    )
    if ([string]::IsNullOrWhiteSpace($Email)) {
        return $false
    }

    $cacheKey = "email:$Email"
    if ($script:resolvedUserCache.ContainsKey($cacheKey)) {
        return $script:resolvedUserCache[$cacheKey]
    }

    try {
        $ctx = Get-PnPContext
        $user = $ctx.Web.EnsureUser($Email)
        $ctx.Load($user)
        $null = Invoke-PnPQuery
        
        $script:resolvedUserCache[$cacheKey] = $true
        $script:resolvedUserCache["login:$Email"] = $user.LoginName
        return $true
    }
    catch {
        Write-Status "User '$Email' could not be resolved via native verify — likely deleted from directory" -Level WARN
        $script:resolvedUserCache[$cacheKey] = $false
        return $false
    }
}

function Get-RequiredUploadFieldValues {
    param(
        $ListItem,
        [object[]]$Fields
    )

    $values = @{}
    if (-not $ListItem -or -not $Fields) {
        return $values
    }

    $excludedNames = New-NameLookup -Names @(
        'FileLeafRef',
        'FileRef',
        'File_x0020_Type',
        'CheckoutUser',
        'Modified',
        'Created',
        'Author',
        'Editor',
        '_UIVersionString',
        'ContentTypeId'
    )

    foreach ($field in $Fields) {
        if (-not $field) {
            continue
        }

        $internalName = [string](Get-SafeObjectPropertyValue -InputObject $field -PropertyName 'InternalName')
        if ([string]::IsNullOrWhiteSpace($internalName) -or $excludedNames.Contains($internalName)) {
            continue
        }

        $isRequired = [bool](Get-SafeObjectPropertyValue -InputObject $field -PropertyName 'Required')
        $isHidden = [bool](Get-SafeObjectPropertyValue -InputObject $field -PropertyName 'Hidden')
        $isReadOnly = [bool](Get-SafeObjectPropertyValue -InputObject $field -PropertyName 'ReadOnlyField')
        $isSealed = [bool](Get-SafeObjectPropertyValue -InputObject $field -PropertyName 'Sealed')

        if (-not $isRequired -or $isHidden -or $isReadOnly -or $isSealed) {
            continue
        }

        if (-not $ListItem.FieldValues.ContainsKey($internalName)) {
            continue
        }

        $value = $ListItem.FieldValues[$internalName]
        if ($null -eq $value) {
            continue
        }

        $fieldType = [string](Get-SafeObjectPropertyValue -InputObject $field -PropertyName 'TypeAsString')
        if ($fieldType -in @('User', 'UserMulti')) {
            if (Test-ObjectTypeName -InputObject $value -TypeNames @('Microsoft.SharePoint.Client.FieldUserValue')) {
                $userEmail = [string](Get-SafeObjectPropertyValue -InputObject $value -PropertyName 'Email')
                $userLabel = Get-SharePointUserIdentityLabel -UserValue $value
                if (-not (Test-SharePointUserResolvable -Email $userEmail)) {
                    Write-Status "    SKIP field  : '$internalName' — user '$userLabel' could not be resolved" -Level WARN
                    continue
                }
                $value = $script:resolvedUserCache["login:$userEmail"]
            }
            elseif ($value -is [System.Collections.IEnumerable]) {
                $ids = [System.Collections.Generic.List[object]]::new()
                foreach ($entry in $value) {
                    $entryLookupId = Get-SafeObjectPropertyValue -InputObject $entry -PropertyName 'LookupId'
                    if ($null -eq $entryLookupId) { continue }
                    $entryLookupId = [int]$entryLookupId
                    $entryEmail = [string](Get-SafeObjectPropertyValue -InputObject $entry -PropertyName 'Email')
                    $entryLabel = Get-SharePointUserIdentityLabel -UserValue $entry
                    if (-not (Test-SharePointUserResolvable -Email $entryEmail)) {
                        Write-Status "    SKIP user   : '$entryLabel' in field '$internalName' could not be resolved" -Level WARN
                        continue
                    }
                    $ids.Add($script:resolvedUserCache["login:$entryEmail"])
                }
                if ($ids.Count -eq 0) { continue }
                $value = $ids.ToArray()
            }
        }
        elseif ($fieldType -in @('Lookup', 'LookupMulti')) {
            if (Test-ObjectTypeName -InputObject $value -TypeNames @('Microsoft.SharePoint.Client.FieldLookupValue')) {
                $value = $value.LookupId
            }
            elseif ($value -is [System.Collections.IEnumerable]) {
                $ids = [System.Collections.Generic.List[int]]::new()
                foreach ($entry in $value) {
                    $lookupId = Get-SafeObjectPropertyValue -InputObject $entry -PropertyName 'LookupId'
                    if ($null -ne $lookupId) {
                        $ids.Add([int]$lookupId)
                    }
                }
                if ($ids.Count -eq 0) { continue }
                $value = $ids.ToArray()
            }
        }
        elseif ($fieldType -in @('TaxonomyFieldType', 'TaxonomyFieldTypeMulti')) {
            # Managed Metadata: convert TaxonomyFieldValue objects to "Label|TermGuid"
            # strings that Add-PnPFile -Values expects.
            $termLabel = Get-SafeObjectPropertyValue -InputObject $value -PropertyName 'Label'
            $termGuid = Get-SafeObjectPropertyValue -InputObject $value -PropertyName 'TermGuid'
            if (-not [string]::IsNullOrWhiteSpace($termGuid)) {
                $value = "$termLabel|$termGuid"
            }
            elseif ($value -is [System.Collections.IEnumerable]) {
                $termParts = [System.Collections.Generic.List[string]]::new()
                foreach ($entry in $value) {
                    $entryLabel = Get-SafeObjectPropertyValue -InputObject $entry -PropertyName 'Label'
                    $entryGuid = Get-SafeObjectPropertyValue -InputObject $entry -PropertyName 'TermGuid'
                    if (-not [string]::IsNullOrWhiteSpace($entryGuid)) {
                        $termParts.Add("$entryLabel|$entryGuid")
                    }
                }
                if ($termParts.Count -eq 0) { continue }
                $value = $termParts -join ';#'
            }
        }

        $values[$internalName] = $value
    }

    return $values
}

function Undo-PnPFileCheckoutCompat {
    param([string]$FileRef)

    if (Get-Command -Name 'Undo-PnPFileCheckedOut' -ErrorAction SilentlyContinue) {
        Invoke-WithRetry -OperationName "Undo checkout for '$FileRef'" -Action {
            Undo-PnPFileCheckedOut -Url $FileRef -ErrorAction Stop
        }
        return
    }

    if (Get-Command -Name 'Undo-PnPFileCheckout' -ErrorAction SilentlyContinue) {
        Invoke-WithRetry -OperationName "Undo checkout for '$FileRef'" -Action {
            Undo-PnPFileCheckout -Url $FileRef -ErrorAction Stop
        }
        return
    }

    throw "No supported PnP undo-checkout cmdlet was found for '$FileRef'."
}

function Convert-ValueForExistingProperty {
    param(
        [object]$Value,
        [string]$ElementName
    )

    if ($null -eq $Value) {
        return ''
    }

    switch ($ElementName) {
        'filetime' {
            try {
                return ([datetime]$Value).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
            }
            catch {
                return [string]$Value
            }
        }
        'bool' {
            return ([System.Convert]::ToBoolean($Value)).ToString().ToLowerInvariant()
        }
        'i4' {
            return [string][int]$Value
        }
        'i8' {
            return [string][long]$Value
        }
        'r8' {
            return ([double]$Value).ToString([System.Globalization.CultureInfo]::InvariantCulture)
        }
        default {
            if ($Value -is [datetime]) {
                return $Value.ToString('yyyy-MM-ddTHH:mm:ss')
            }
            return [string]$Value
        }
    }
}

function New-VtElementForValue {
    param(
        [System.Xml.XmlDocument]$XmlDoc,
        [object]$Value
    )

    $vtNs = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'

    if ($Value -is [datetime]) {
        $element = $XmlDoc.CreateElement('vt', 'filetime', $vtNs)
        $element.InnerText = $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        return $element
    }

    if ($Value -is [bool]) {
        $element = $XmlDoc.CreateElement('vt', 'bool', $vtNs)
        $element.InnerText = $Value.ToString().ToLowerInvariant()
        return $element
    }

    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32]) {
        $element = $XmlDoc.CreateElement('vt', 'i4', $vtNs)
        $element.InnerText = [string]$Value
        return $element
    }

    if ($Value -is [int64]) {
        $element = $XmlDoc.CreateElement('vt', 'i8', $vtNs)
        $element.InnerText = [string]$Value
        return $element
    }

    if ($Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
        $element = $XmlDoc.CreateElement('vt', 'r8', $vtNs)
        $element.InnerText = ([double]$Value).ToString([System.Globalization.CultureInfo]::InvariantCulture)
        return $element
    }

    $textElement = $XmlDoc.CreateElement('vt', 'lpwstr', $vtNs)
    $textElement.InnerText = [string]$Value
    return $textElement
}

function ConvertTo-PropertyValue {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    if (Test-ObjectTypeName -InputObject $Value -TypeNames @('Microsoft.SharePoint.Client.FieldUserValue')) {
        return $Value.LookupValue
    }

    if (Test-ObjectTypeName -InputObject $Value -TypeNames @('Microsoft.SharePoint.Client.FieldLookupValue')) {
        return $Value.LookupValue
    }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $parts = [System.Collections.Generic.List[string]]::new()
        foreach ($entry in $Value) {
            if ($null -eq $entry) {
                continue
            }
            if (Test-ObjectTypeName -InputObject $entry -TypeNames @('Microsoft.SharePoint.Client.FieldUserValue')) {
                $parts.Add($entry.LookupValue)
                continue
            }
            if (Test-ObjectTypeName -InputObject $entry -TypeNames @('Microsoft.SharePoint.Client.FieldLookupValue')) {
                $parts.Add($entry.LookupValue)
                continue
            }
            $parts.Add([string]$entry)
        }
        return ($parts -join '; ')
    }

    if ($Value -is [datetime]) {
        return $Value
    }

    if ($Value -is [bool] -or
        $Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int32] -or
        $Value -is [int64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]) {
        return $Value
    }

    return [string]$Value
}

function New-PropertyResolver {
    param(
        [object]$ListItem,
        [object[]]$Fields,
        [System.Collections.Generic.HashSet[string]]$AllowedNames,
        [hashtable]$AliasMap
    )

    $resolver = @{}

    foreach ($key in $ListItem.FieldValues.Keys) {
        if ([string]::IsNullOrWhiteSpace([string]$key)) {
            continue
        }
        if ($AllowedNames.Contains([string]$key)) {
            $resolver[$key] = $ListItem.FieldValues[$key]
        }
        if ($AliasMap.ContainsKey([string]$key)) {
            foreach ($alias in $AliasMap[[string]$key]) {
                if ($AllowedNames.Contains([string]$alias)) {
                    $resolver[$alias] = $ListItem.FieldValues[$key]
                }
            }
        }
    }

    foreach ($field in $Fields) {
        if (-not $field) {
            continue
        }

        $candidateKeys = @($field.InternalName, $field.StaticName, $field.Title, $field.EntityPropertyName) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique

        $matchedValue = $null
        $matched = $false
        foreach ($candidateKey in $candidateKeys) {
            if (-not $AllowedNames.Contains([string]$candidateKey)) {
                continue
            }
            if ($ListItem.FieldValues.ContainsKey($candidateKey)) {
                $matchedValue = $ListItem.FieldValues[$candidateKey]
                $matched = $true
                break
            }
        }

        if (-not $matched) {
            continue
        }

        foreach ($candidateKey in $candidateKeys) {
            if ($AllowedNames.Contains([string]$candidateKey)) {
                $resolver[$candidateKey] = $matchedValue
            }
            if ($AliasMap.ContainsKey([string]$candidateKey)) {
                foreach ($alias in $AliasMap[[string]$candidateKey]) {
                    if ($AllowedNames.Contains([string]$alias)) {
                        $resolver[$alias] = $matchedValue
                    }
                }
            }
        }
    }

    return $resolver
}

function Update-CustomPropertiesDocument {
    param(
        [System.Xml.XmlDocument]$XmlDoc,
        [hashtable]$PropertyValues,
        [string[]]$TargetNames,
        [System.Collections.Generic.HashSet[string]]$AllowedNames
    )

    $opNs = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
    $vtNs = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
    $fmtid = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'

    $nsManager = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
    $nsManager.AddNamespace('op', $opNs)
    $nsManager.AddNamespace('vt', $vtNs)

    $propertiesNode = $XmlDoc.SelectSingleNode('/op:Properties', $nsManager)
    if (-not $propertiesNode) {
        $propertiesNode = $XmlDoc.CreateElement('op', 'Properties', $opNs)
        $null = $propertiesNode.SetAttribute('xmlns:vt', $vtNs)
        if (-not $XmlDoc.DocumentElement) {
            $null = $XmlDoc.AppendChild($propertiesNode)
        }
        else {
            $null = $XmlDoc.ReplaceChild($propertiesNode, $XmlDoc.DocumentElement)
        }
    }

    $allProps = @($propertiesNode.SelectNodes('op:property', $nsManager))
    $groups = @{}
    foreach ($prop in $allProps) {
        $name = $prop.GetAttribute('name')
        if (-not $groups.ContainsKey($name)) {
            $groups[$name] = [System.Collections.Generic.List[System.Xml.XmlElement]]::new()
        }
        $groups[$name].Add($prop)
    }

    $removedDuplicates = [System.Collections.Generic.List[string]]::new()
    foreach ($name in @($groups.Keys)) {
        if (-not $AllowedNames.Contains($name)) {
            continue
        }

        $entries = $groups[$name]
        if ($entries.Count -le 1) {
            continue
        }

        $sorted = @($entries | Sort-Object { [int]$_.GetAttribute('pid') })
        for ($index = 1; $index -lt $sorted.Count; $index++) {
            $duplicate = $sorted[$index]
            $null = $propertiesNode.RemoveChild($duplicate)
            $removedDuplicates.Add($name)
        }

        $groups[$name] = [System.Collections.Generic.List[System.Xml.XmlElement]]::new()
        $groups[$name].Add($sorted[0])
    }

    $updatedNames = [System.Collections.Generic.List[string]]::new()
    $addedNames = [System.Collections.Generic.List[string]]::new()
    $missingNames = [System.Collections.Generic.List[string]]::new()

    foreach ($name in $TargetNames) {
        if (-not $PropertyValues.ContainsKey($name)) {
            $missingNames.Add($name)
            continue
        }

        $value = ConvertTo-PropertyValue -Value $PropertyValues[$name]
        $propertyNode = $null
        if ($groups.ContainsKey($name) -and $groups[$name].Count -gt 0) {
            $propertyNode = $groups[$name][0]
        }

        if ($propertyNode) {
            $valueElement = $null
            foreach ($child in $propertyNode.ChildNodes) {
                if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                    $valueElement = [System.Xml.XmlElement]$child
                    break
                }
            }

            if ($valueElement) {
                $convertedValue = Convert-ValueForExistingProperty -Value $value -ElementName $valueElement.LocalName
                if ($valueElement.InnerText -ne $convertedValue) {
                    $valueElement.InnerText = $convertedValue
                    $updatedNames.Add($name)
                }
            }
            else {
                $null = $propertyNode.AppendChild((New-VtElementForValue -XmlDoc $XmlDoc -Value $value))
                $updatedNames.Add($name)
            }
        }
        else {
            $propertyNode = $XmlDoc.CreateElement('op', 'property', $opNs)
            $null = $propertyNode.SetAttribute('fmtid', $fmtid)
            $null = $propertyNode.SetAttribute('name', $name)
            $null = $propertyNode.AppendChild((New-VtElementForValue -XmlDoc $XmlDoc -Value $value))
            $null = $propertiesNode.AppendChild($propertyNode)
            $addedNames.Add($name)
        }
    }

    $remainingProps = @($propertiesNode.SelectNodes('op:property', $nsManager))
    $nextPid = 2
    foreach ($prop in $remainingProps) {
        $null = $prop.SetAttribute('pid', [string]$nextPid)
        $nextPid++
    }

    return [pscustomobject]@{
        UpdatedNames      = @($updatedNames)
        AddedNames        = @($addedNames)
        MissingNames      = @($missingNames | Select-Object -Unique)
        RemovedDuplicates = @($removedDuplicates)
    }
}

function Ensure-CustomPropertiesRelationshipDocument {
    param([System.Xml.XmlDocument]$XmlDoc)

    $relNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
    $targetPath = 'docProps/custom.xml'

    $nsManager = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
    $nsManager.AddNamespace('rel', $relNs)

    $relationshipsNode = $XmlDoc.SelectSingleNode('/rel:Relationships', $nsManager)
    if (-not $relationshipsNode) {
        $relationshipsNode = $XmlDoc.CreateElement('Relationships', $relNs)
        if ($XmlDoc.DocumentElement) {
            $null = $XmlDoc.ReplaceChild($relationshipsNode, $XmlDoc.DocumentElement)
        }
        else {
            $null = $XmlDoc.AppendChild($relationshipsNode)
        }
    }

    $relationshipNode = $XmlDoc.SelectSingleNode("/rel:Relationships/rel:Relationship[@Type='$relType']", $nsManager)
    if ($relationshipNode) {
        if ($relationshipNode.GetAttribute('Target') -ne $targetPath) {
            $null = $relationshipNode.SetAttribute('Target', $targetPath)
            return $true
        }
        return $false
    }

    $maxId = 0
    foreach ($existingNode in @($relationshipsNode.SelectNodes('rel:Relationship', $nsManager))) {
        $idValue = $existingNode.GetAttribute('Id')
        if ($idValue -match '^rId(\d+)$') {
            $maxId = [Math]::Max($maxId, [int]$Matches[1])
        }
    }

    $newRelationship = $XmlDoc.CreateElement('Relationship', $relNs)
    $null = $newRelationship.SetAttribute('Id', "rId$($maxId + 1)")
    $null = $newRelationship.SetAttribute('Type', $relType)
    $null = $newRelationship.SetAttribute('Target', $targetPath)
    $null = $relationshipsNode.AppendChild($newRelationship)
    return $true
}

function Ensure-CustomPropertiesContentTypeDocument {
    param([System.Xml.XmlDocument]$XmlDoc)

    $contentTypeNs = 'http://schemas.openxmlformats.org/package/2006/content-types'
    $partName = '/docProps/custom.xml'
    $contentType = 'application/vnd.openxmlformats-officedocument.custom-properties+xml'

    $nsManager = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
    $nsManager.AddNamespace('ct', $contentTypeNs)

    $typesNode = $XmlDoc.SelectSingleNode('/ct:Types', $nsManager)
    if (-not $typesNode) {
        $typesNode = $XmlDoc.CreateElement('Types', $contentTypeNs)
        if ($XmlDoc.DocumentElement) {
            $null = $XmlDoc.ReplaceChild($typesNode, $XmlDoc.DocumentElement)
        }
        else {
            $null = $XmlDoc.AppendChild($typesNode)
        }
    }

    $overrideNode = $XmlDoc.SelectSingleNode("/ct:Types/ct:Override[@PartName='$partName']", $nsManager)
    if ($overrideNode) {
        if ($overrideNode.GetAttribute('ContentType') -ne $contentType) {
            $null = $overrideNode.SetAttribute('ContentType', $contentType)
            return $true
        }
        return $false
    }

    $newOverride = $XmlDoc.CreateElement('Override', $contentTypeNs)
    $null = $newOverride.SetAttribute('PartName', $partName)
    $null = $newOverride.SetAttribute('ContentType', $contentType)
    $null = $typesNode.AppendChild($newOverride)
    return $true
}

function Update-DocxQuickPartsFromBytes {
    param(
        [byte[]]$DocxBytes,
        [hashtable]$PropertyValues,
        [System.Collections.Generic.HashSet[string]]$AllowedNames,
        [bool]$EnableUpdateOnOpen = $true
    )

    $memoryStream = New-Object System.IO.MemoryStream
    try {
        $memoryStream.Write($DocxBytes, 0, $DocxBytes.Length)
        $memoryStream.Position = 0
        $resultData = $null

        $zip = New-Object System.IO.Compression.ZipArchive(
            $memoryStream,
            [System.IO.Compression.ZipArchiveMode]::Update,
            $true
        )

        try {
            $storeMap = Get-BoundStoreMap -ZipArchive $zip
            $storeDocuments = @{}
            foreach ($storeEntryName in $storeMap.Values | Select-Object -Unique) {
                $storeEntry = $zip.GetEntry($storeEntryName)
                if ($storeEntry) {
                    $storeDocuments[$storeEntryName] = Get-XmlDocumentFromBytes -Bytes (Read-ZipEntryBytes -Entry $storeEntry)
                }
            }

            $wordXmlEntries = @($zip.Entries | Where-Object {
                    $_.FullName -match '^word/.+\.xml$' -and $_.Name -notmatch '^_rels'
                })

            $fieldNames = [System.Collections.Generic.List[string]]::new()
            $dirtyCount = 0
            $updatedParts = [System.Collections.Generic.List[string]]::new()
            $boundNames = [System.Collections.Generic.List[string]]::new()
            $syncedBoundNames = [System.Collections.Generic.List[string]]::new()
            $fallbackBoundNames = [System.Collections.Generic.List[string]]::new()
            $boundMissingNames = [System.Collections.Generic.List[string]]::new()
            $boundCustomPropertyNames = [System.Collections.Generic.List[string]]::new()
            $updatedStoreNames = [System.Collections.Generic.List[string]]::new()
            $boundDisplayCount = 0
            $packageMetadataUpdated = $false

            foreach ($entry in $wordXmlEntries) {
                $entryBytes = Read-ZipEntryBytes -Entry $entry
                if (-not (Test-WordXmlEntryMayContainTrackedContent -EntryName $entry.FullName -EntryBytes $entryBytes)) {
                    continue
                }

                $xmlDoc = Get-XmlDocumentFromBytes -Bytes $entryBytes
                $fieldResult = Update-FieldDocument -XmlDoc $xmlDoc -PropertyValues $PropertyValues -AllowedNames $AllowedNames -EnableUpdateOnOpen:$EnableUpdateOnOpen
                $boundResult = Update-BoundContentControlsDocument `
                    -XmlDoc $xmlDoc `
                    -PropertyValues $PropertyValues `
                    -StoreMap $storeMap `
                    -StoreDocuments $storeDocuments `
                    -AllowedNames $AllowedNames

                foreach ($name in $fieldResult.PropertyNames) {
                    $fieldNames.Add($name)
                }
                foreach ($name in $boundResult.UpdatedNames) {
                    $boundNames.Add($name)
                }
                foreach ($name in $boundResult.AuthoritativeNames) {
                    $syncedBoundNames.Add($name)
                }
                foreach ($name in $boundResult.FallbackNames) {
                    $fallbackBoundNames.Add($name)
                }
                foreach ($name in $boundResult.MissingNames) {
                    $boundMissingNames.Add($name)
                }
                foreach ($name in $boundResult.CustomPropertyNames) {
                    $boundCustomPropertyNames.Add($name)
                }
                foreach ($storeName in $boundResult.UpdatedStores) {
                    $updatedStoreNames.Add($storeName)
                }
                $dirtyCount += $fieldResult.DirtyCount
                $boundDisplayCount += $fieldResult.DisplayCount
                $boundDisplayCount += $boundResult.DisplayCount

                if ($fieldResult.DirtyCount -eq 0 -and $fieldResult.DisplayCount -eq 0 -and $boundResult.UpdatedNames.Count -eq 0 -and $entry.FullName -ne 'word/settings.xml') {
                    continue
                }

                Set-ZipEntryBytes -ZipArchive $zip -EntryName $entry.FullName -Bytes (ConvertTo-CleanXmlBytes -XmlDoc $xmlDoc)
                $updatedParts.Add($entry.FullName)
            }

            $fieldNameList = @($fieldNames | Select-Object -Unique)

            $settingsUpdated = $false
            $settingsEntry = $zip.GetEntry('word/settings.xml')
            if ($settingsEntry) {
                $settingsDoc = Get-XmlDocumentFromBytes -Bytes (Read-ZipEntryBytes -Entry $settingsEntry)
                $settingsUpdated = Update-WordSettingsDocument -XmlDoc $settingsDoc -EnableUpdateOnOpen:$EnableUpdateOnOpen
                Set-ZipEntryBytes -ZipArchive $zip -EntryName 'word/settings.xml' -Bytes (ConvertTo-CleanXmlBytes -XmlDoc $settingsDoc)
            }

            $propertyTargetNames = @($fieldNameList + $boundCustomPropertyNames | Select-Object -Unique)
            $propertyNamesWithValues = @($propertyTargetNames | Where-Object { $PropertyValues.ContainsKey($_) } | Select-Object -Unique)

            $propertyChangeResult = [pscustomobject]@{
                UpdatedNames      = @()
                AddedNames        = @()
                MissingNames      = @()
                RemovedDuplicates = @()
            }

            if ($propertyTargetNames.Count -gt 0) {
                $customEntry = $zip.GetEntry('docProps/custom.xml')
                if (-not $customEntry -and $propertyNamesWithValues.Count -eq 0) {
                    $customDoc = $null
                    $propertyChangeResult = [pscustomobject]@{
                        UpdatedNames      = @()
                        AddedNames        = @()
                        MissingNames      = @($propertyTargetNames)
                        RemovedDuplicates = @()
                    }
                }
                elseif ($customEntry) {
                    $customDoc = Get-XmlDocumentFromBytes -Bytes (Read-ZipEntryBytes -Entry $customEntry)
                }
                else {
                    $customDoc = New-Object System.Xml.XmlDocument
                    $null = $customDoc.AppendChild($customDoc.CreateXmlDeclaration('1.0', 'UTF-8', $null))
                    $propertiesElement = $customDoc.CreateElement('op', 'Properties', 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties')
                    $null = $propertiesElement.SetAttribute('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes')
                    $null = $customDoc.AppendChild($propertiesElement)
                }

                if ($customDoc) {
                    $propertyChangeResult = Update-CustomPropertiesDocument `
                        -XmlDoc $customDoc `
                        -PropertyValues $PropertyValues `
                        -TargetNames $propertyTargetNames `
                        -AllowedNames $AllowedNames

                    Set-ZipEntryBytes -ZipArchive $zip -EntryName 'docProps/custom.xml' -Bytes (ConvertTo-CleanXmlBytes -XmlDoc $customDoc)

                    $packageRelsEntry = $zip.GetEntry('_rels/.rels')
                    if ($packageRelsEntry) {
                        $packageRelsDoc = Get-XmlDocumentFromBytes -Bytes (Read-ZipEntryBytes -Entry $packageRelsEntry)
                    }
                    else {
                        $packageRelsDoc = New-Object System.Xml.XmlDocument
                        $null = $packageRelsDoc.AppendChild($packageRelsDoc.CreateXmlDeclaration('1.0', 'UTF-8', $null))
                        $relationshipsElement = $packageRelsDoc.CreateElement('Relationships', 'http://schemas.openxmlformats.org/package/2006/relationships')
                        $null = $packageRelsDoc.AppendChild($relationshipsElement)
                    }

                    if (Ensure-CustomPropertiesRelationshipDocument -XmlDoc $packageRelsDoc) {
                        $packageMetadataUpdated = $true
                    }
                    Set-ZipEntryBytes -ZipArchive $zip -EntryName '_rels/.rels' -Bytes (ConvertTo-CleanXmlBytes -XmlDoc $packageRelsDoc)

                    $contentTypesEntry = $zip.GetEntry('[Content_Types].xml')
                    if ($contentTypesEntry) {
                        $contentTypesDoc = Get-XmlDocumentFromBytes -Bytes (Read-ZipEntryBytes -Entry $contentTypesEntry)
                        if (Ensure-CustomPropertiesContentTypeDocument -XmlDoc $contentTypesDoc) {
                            $packageMetadataUpdated = $true
                        }
                        Set-ZipEntryBytes -ZipArchive $zip -EntryName '[Content_Types].xml' -Bytes (ConvertTo-CleanXmlBytes -XmlDoc $contentTypesDoc)
                    }
                }
            }

            foreach ($storeEntryName in @($updatedStoreNames | Select-Object -Unique)) {
                Set-ZipEntryBytes -ZipArchive $zip -EntryName $storeEntryName -Bytes (ConvertTo-CleanXmlBytes -XmlDoc $storeDocuments[$storeEntryName])
            }

            $changed = $dirtyCount -gt 0 -or
            $boundDisplayCount -gt 0 -or
            $settingsUpdated -or
            $propertyChangeResult.UpdatedNames.Count -gt 0 -or
            $propertyChangeResult.AddedNames.Count -gt 0 -or
            $propertyChangeResult.RemovedDuplicates.Count -gt 0 -or
            $updatedStoreNames.Count -gt 0 -or
            $packageMetadataUpdated

            $resultData = [ordered]@{
                Changed            = $changed
                QuickPartNames     = $fieldNameList
                BoundNames         = @($boundNames | Select-Object -Unique)
                SyncedBoundNames   = @($syncedBoundNames | Select-Object -Unique)
                FallbackBoundNames = @($fallbackBoundNames | Select-Object -Unique)
                DirtyCount         = $dirtyCount
                BoundDisplayCount  = $boundDisplayCount
                SettingsUpdated    = $settingsUpdated
                UpdatedProperties  = $propertyChangeResult.UpdatedNames
                AddedProperties    = $propertyChangeResult.AddedNames
                MissingProperties  = @($propertyChangeResult.MissingNames + $boundMissingNames | Select-Object -Unique)
                RemovedDuplicates  = $propertyChangeResult.RemovedDuplicates
                UpdatedParts       = @($updatedParts | Select-Object -Unique)
                UpdatedStores      = @($updatedStoreNames | Select-Object -Unique)
            }
        }
        finally {
            $zip.Dispose()
        }

        if (-not $resultData) {
            throw 'Could not build DOCX update result.'
        }

        return [pscustomobject]@{
            Bytes              = $memoryStream.ToArray()
            Changed            = $resultData.Changed
            QuickPartNames     = $resultData.QuickPartNames
            BoundNames         = $resultData.BoundNames
            SyncedBoundNames   = $resultData.SyncedBoundNames
            FallbackBoundNames = $resultData.FallbackBoundNames
            DirtyCount         = $resultData.DirtyCount
            BoundDisplayCount  = $resultData.BoundDisplayCount
            SettingsUpdated    = $resultData.SettingsUpdated
            UpdatedProperties  = $resultData.UpdatedProperties
            AddedProperties    = $resultData.AddedProperties
            MissingProperties  = $resultData.MissingProperties
            RemovedDuplicates  = $resultData.RemovedDuplicates
            UpdatedParts       = $resultData.UpdatedParts
            UpdatedStores      = $resultData.UpdatedStores
        }
    }
    finally {
        $memoryStream.Dispose()
    }
}

function Upload-DocxBytesBack {
    param(
        [byte[]]$Bytes,
        [string]$FileRef,
        [string]$WebUrl,
        $ExistingListItem,
        [object[]]$Fields,
        [switch]$VersioningDisabled
    )

    $leafName = [System.IO.Path]::GetFileName($FileRef)
    $folderFull = $FileRef.Substring(0, $FileRef.Length - $leafName.Length).TrimEnd('/')

    if ($WebUrl -ne '' -and $folderFull.StartsWith($WebUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
        $folderPath = $folderFull.Substring($WebUrl.Length).TrimStart('/')
    }
    else {
        $folderPath = $folderFull.TrimStart('/')
    }

    $spFile = if ($ExistingListItem) {
        $ExistingListItem
    }
    else {
        Invoke-WithRetry -OperationName "Load list item for '$FileRef'" -Action {
            Get-PnPFile -Url $FileRef -AsListItem -ErrorAction Stop
        }
    }
    $uiVersion = [string]$spFile.FieldValues['_UIVersionString']
    $requiredUploadValues = Get-RequiredUploadFieldValues -ListItem $spFile -Fields $Fields
    Write-Host "    Version     : $uiVersion"
    if ($requiredUploadValues.Count -gt 0) {
        Write-Host "    Metadata    : preserving $($requiredUploadValues.Count) required field(s)"
    }

    $isCurrentUserCheckout = $false
    $checkedOutBy = $spFile.FieldValues['CheckoutUser']
    if ($checkedOutBy) {
        $checkoutLookupId = Get-SafeObjectPropertyValue -InputObject $checkedOutBy -PropertyName 'LookupId'
        $checkoutLookupValue = Get-SafeObjectPropertyValue -InputObject $checkedOutBy -PropertyName 'LookupValue'
        $checkoutEmail = Get-SafeObjectPropertyValue -InputObject $checkedOutBy -PropertyName 'Email'
        $coUser = if (Test-ObjectTypeName -InputObject $checkedOutBy -TypeNames @('Microsoft.SharePoint.Client.FieldUserValue')) { [string]$checkoutLookupValue } else { [string]$checkedOutBy }
        $ctx = Get-PnPContext
        $ctx.Load($ctx.Web.CurrentUser)
        $null = Invoke-WithRetry -OperationName 'Load current SharePoint user' -Action {
            Invoke-PnPQuery
        }

        $currentUser = $ctx.Web.CurrentUser
        $currentCandidates = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($candidate in @(
                [string]$currentUser.LoginName,
                [string]$currentUser.Email,
                [string]$currentUser.Title,
                [string]$currentUser.Id
            )) {
            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                $currentCandidates.Add($candidate) | Out-Null
                if ($candidate.Contains('|')) {
                    $currentCandidates.Add($candidate.Substring($candidate.LastIndexOf('|') + 1)) | Out-Null
                }
            }
        }

        $checkoutCandidates = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($candidate in @(
                [string]$checkoutLookupId,
                [string]$checkoutLookupValue,
                [string]$checkoutEmail
            )) {
            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                $checkoutCandidates.Add($candidate) | Out-Null
            }
        }

        foreach ($candidate in $checkoutCandidates) {
            if ($currentCandidates.Contains($candidate)) {
                $isCurrentUserCheckout = $true
                break
            }
        }

        if ($coUser -and -not $isCurrentUserCheckout) {
            throw "File is already checked out by '$coUser'. Cannot proceed."
        }
    }

    $checkinOk = $false

    $refreshWopiCache = {
        try {
            $ctx = Get-PnPContext
            $spFileObj = $ctx.Web.GetFileByServerRelativeUrl($FileRef)
            $ctx.Load($spFileObj)
            $ctx.Load($spFileObj.ListItemAllFields)
            Invoke-PnPQuery
            $spFileObj.ListItemAllFields.UpdateOverwriteVersion()
            Invoke-PnPQuery
            Write-Host '    WopiRefresh : ETag updated'
        }
        catch {
            Write-Status "    WopiRefresh : Could not force ETag update - $($_.Exception.Message)" -Level WARN
        }
    }

    if ($VersioningDisabled) {
        $memoryStream = New-Object System.IO.MemoryStream(, $Bytes)
        try {
            Write-Host '    Strategy    : direct upload (versioning disabled)'
            $addFileParameters = @{
                FileName    = $leafName
                Folder      = $folderPath
                Stream      = $memoryStream
                ErrorAction = 'Stop'
            }
            if ($requiredUploadValues.Count -gt 0) {
                $addFileParameters['Values'] = $requiredUploadValues
            }
            try {
                Invoke-WithRetry -OperationName "Upload '$FileRef'" -Action {
                    if ($memoryStream.CanSeek) {
                        $memoryStream.Position = 0
                    }
                    Add-PnPFile @addFileParameters | Out-Null
                }
            }
            catch {
                if ($requiredUploadValues.Count -gt 0) {
                    Write-Status "    Upload with metadata failed ($($_.Exception.Message)) — retrying without metadata" -Level WARN
                    $memoryStream = New-Object System.IO.MemoryStream(, $Bytes)
                    Invoke-WithRetry -OperationName "Upload '$FileRef' without metadata" -Action {
                        if ($memoryStream.CanSeek) {
                            $memoryStream.Position = 0
                        }
                        Add-PnPFile -FileName $leafName -Folder $folderPath -Stream $memoryStream -ErrorAction Stop | Out-Null
                    }
                }
                else {
                    throw
                }
            }
            Write-Host '    Upload      : file replaced'
            Write-Host "    Result      : $uiVersion (unchanged)"
            & $refreshWopiCache
            $checkinOk = $true
        }
        catch {
            Write-Status "    Upload failed - $($_.Exception.Message)" -Level ERROR
        }
        finally {
            $memoryStream.Dispose()
        }
    }
    else {
        $memoryStream = New-Object System.IO.MemoryStream(, $Bytes)
        $performedCheckout = $false
        try {
            $addFileParameters = @{
                FileName    = $leafName
                Folder      = $folderPath
                Stream      = $memoryStream
                ErrorAction = 'Stop'
            }

            if ($null -ne $isCurrentUserCheckout -and $isCurrentUserCheckout) {
                Write-Host '    Strategy    : Add-PnPFile (Already checked out - updating draft)'
            }
            else {
                Write-Host '    Strategy    : explicit checkout -> upload -> Set-PnPFileCheckedIn OverwriteCheckIn'
                Invoke-WithRetry -OperationName "Check out '$FileRef'" -Action {
                    Set-PnPFileCheckedOut -Url $FileRef -ErrorAction Stop
                }
                $performedCheckout = $true
            }

            if ($requiredUploadValues.Count -gt 0) {
                $addFileParameters['Values'] = $requiredUploadValues
            }
            try {
                Invoke-WithRetry -OperationName "Upload '$FileRef'" -Action {
                    if ($memoryStream.CanSeek) {
                        $memoryStream.Position = 0
                    }
                    Add-PnPFile @addFileParameters | Out-Null
                }
            }
            catch {
                if ($requiredUploadValues.Count -gt 0) {
                    Write-Status "    Upload with metadata failed ($($_.Exception.Message)) — retrying without metadata" -Level WARN
                    $memoryStream = New-Object System.IO.MemoryStream(, $Bytes)
                    $addFileParameters.Remove('Values')
                    $addFileParameters['Stream'] = $memoryStream
                    Invoke-WithRetry -OperationName "Upload '$FileRef' without metadata" -Action {
                        if ($memoryStream.CanSeek) {
                            $memoryStream.Position = 0
                        }
                        Add-PnPFile @addFileParameters | Out-Null
                    }
                }
                else {
                    throw
                }
            }

            if ($performedCheckout) {
                Invoke-WithRetry -OperationName "Check in '$FileRef'" -Action {
                    Set-PnPFileCheckedIn -Url $FileRef -Comment 'Synced Quick Parts from SharePoint metadata' -CheckinType OverwriteCheckIn -ErrorAction Stop
                }
            }
            Write-Host '    Upload      : file replaced & checked in'
        }
        catch {
            Write-Status "    Upload/checkin failed - $($_.Exception.Message)" -Level ERROR
            try {
                $postFile = Invoke-WithRetry -OperationName "Verify checkout state for '$FileRef'" -Action {
                    Get-PnPFile -Url $FileRef -AsListItem -ErrorAction SilentlyContinue
                }
                if ($postFile -and $postFile.FieldValues['CheckoutUser'] -and $performedCheckout) {
                    Write-Status '    UndoCheckout: reverting due to failure' -Level WARN
                    Undo-PnPFileCheckoutCompat -FileRef $FileRef
                }
            }
            catch {
                Write-Status "    UndoCheckout: Warning - $($_.Exception.Message)" -Level ERROR
            }
            return $false
        }
        finally {
            $memoryStream.Dispose()
        }

        try {
            $postFile = Invoke-WithRetry -OperationName "Verify upload result for '$FileRef'" -Action {
                Get-PnPFile -Url $FileRef -AsListItem -ErrorAction Stop
            }
            $postVersion = [string]$postFile.FieldValues['_UIVersionString']
            $postCheckout = $postFile.FieldValues['CheckoutUser']

            if ($postCheckout) {
                if ($isCurrentUserCheckout) {
                    Write-Host '    Strategy    : Maintaining existing draft checkout state'
                }
                else {
                    Write-Status '    File is STILL checked out after OverwriteCheckIn!' -Level WARN
                    try {
                        Write-Host '    Fallback    : explicit Set-PnPFileCheckedIn -CheckinType OverwriteCheckIn'
                        Invoke-WithRetry -OperationName "Fallback check in '$FileRef'" -Action {
                            Set-PnPFileCheckedIn -Url $FileRef -Comment 'Synced Quick Parts from SharePoint metadata' -CheckinType OverwriteCheckIn -ErrorAction Stop
                        }
                        $postFile2 = Invoke-WithRetry -OperationName "Verify fallback check-in for '$FileRef'" -Action {
                            Get-PnPFile -Url $FileRef -AsListItem -ErrorAction Stop
                        }
                        $postVersion = [string]$postFile2.FieldValues['_UIVersionString']
                        $postCheckout2 = $postFile2.FieldValues['CheckoutUser']
                        if ($postCheckout2) {
                            Write-Status '    File still checked out after fallback!' -Level ERROR
                            return $false
                        }
                    }
                    catch {
                        Write-Status "    Fallback checkin failed - $($_.Exception.Message)" -Level ERROR
                        return $false
                    }
                }
            }

            if ($postVersion -ne $uiVersion) {
                Write-Status "    Version changed from $uiVersion to $postVersion" -Level WARN
            }
            else {
                Write-Host "    Result      : $uiVersion (unchanged)"
            }
            & $refreshWopiCache
            $checkinOk = $true
        }
        catch {
            Write-Status "    Verify   : Could not verify post-upload state - $($_.Exception.Message)" -Level WARN
            $checkinOk = $true
        }
    }

    return $checkinOk
}

function Invoke-WithRetry {
    param(
        [scriptblock]$Action,
        [int]$MaxAttempts = 5,
        [int]$BaseDelayMs = 2000,
        [int]$MaxDelayMs = 30000,
        [string]$OperationName = 'SharePoint operation'
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return (& $Action)
        }
        catch {
            $isThrottled = Test-SharePointThrottleException -Exception $_.Exception
            $isTransient = $isThrottled -or (Test-SharePointTransientException -Exception $_.Exception)

            if (-not $isTransient) {
                throw
            }

            if ($attempt -eq $MaxAttempts) {
                throw
            }

            $delayMs = Get-SharePointRetryDelayMilliseconds -Exception $_.Exception -Attempt $attempt -BaseDelayMs $BaseDelayMs -MaxDelayMs $MaxDelayMs
            $failureLabel = if ($isThrottled) { 'SharePoint throttling' } else { 'transient SharePoint failure' }
            Write-Status "$OperationName hit $failureLabel on attempt $attempt of $MaxAttempts. Retrying in ${delayMs}ms. $($_.Exception.Message)" 'WARN'
            Start-Sleep -Milliseconds $delayMs
        }
    }
}

function Test-SharePointThrottleException {
    param([System.Exception]$Exception)

    if ($null -eq $Exception) {
        return $false
    }

    $messageParts = [System.Collections.Generic.List[string]]::new()
    $cursor = $Exception
    while ($cursor) {
        if (-not [string]::IsNullOrWhiteSpace($cursor.Message)) {
            $messageParts.Add($cursor.Message)
        }

        $statusCode = Get-SafeObjectPropertyValue -InputObject $cursor -PropertyName 'StatusCode'
        if ($null -ne $statusCode) {
            $messageParts.Add([string]$statusCode)
        }

        $serverErrorTypeName = Get-SafeObjectPropertyValue -InputObject $cursor -PropertyName 'ServerErrorTypeName'
        if (-not [string]::IsNullOrWhiteSpace([string]$serverErrorTypeName)) {
            $messageParts.Add([string]$serverErrorTypeName)
        }

        $cursor = $cursor.InnerException
    }

    $combinedMessage = $messageParts -join ' | '
    return $combinedMessage -match '(?i)\b429\b' -or
        $combinedMessage -match '(?i)too many requests' -or
        $combinedMessage -match '(?i)throttl' -or
        $combinedMessage -match '(?i)rate limit' -or
        $combinedMessage -match '(?i)request limit' -or
        $combinedMessage -match '(?i)server too busy' -or
        $combinedMessage -match '(?i)retry after'
}

function Test-SharePointTransientException {
    param([System.Exception]$Exception)

    if ($null -eq $Exception) {
        return $false
    }

    $messageParts = [System.Collections.Generic.List[string]]::new()
    $cursor = $Exception
    while ($cursor) {
        if (-not [string]::IsNullOrWhiteSpace($cursor.Message)) {
            $messageParts.Add($cursor.Message)
        }

        $statusCode = Get-SafeObjectPropertyValue -InputObject $cursor -PropertyName 'StatusCode'
        if ($null -ne $statusCode) {
            $messageParts.Add([string]$statusCode)
        }

        $cursor = $cursor.InnerException
    }

    $combinedMessage = $messageParts -join ' | '
    return $combinedMessage -match '(?i)\b502\b' -or
        $combinedMessage -match '(?i)\b503\b' -or
        $combinedMessage -match '(?i)\b504\b' -or
        $combinedMessage -match '(?i)timeout' -or
        $combinedMessage -match '(?i)temporarily unavailable' -or
        $combinedMessage -match '(?i)connection (was )?closed' -or
        $combinedMessage -match '(?i)server too busy'
}

function Get-SharePointRetryDelayMilliseconds {
    param(
        [System.Exception]$Exception,
        [int]$Attempt,
        [int]$BaseDelayMs = 2000,
        [int]$MaxDelayMs = 30000
    )

    $cursor = $Exception
    while ($cursor) {
        foreach ($propertyName in @('RetryAfter', 'RetryAfterSeconds', 'RetryAfterInSeconds', 'ServerErrorRetryAfterSeconds')) {
            $retryValue = Get-SafeObjectPropertyValue -InputObject $cursor -PropertyName $propertyName
            if ($retryValue -is [TimeSpan]) {
                return [int][math]::Min($MaxDelayMs, [math]::Max(1000, [int]$retryValue.TotalMilliseconds))
            }

            $retrySeconds = 0
            if ($null -ne $retryValue -and [int]::TryParse([string]$retryValue, [ref]$retrySeconds) -and $retrySeconds -gt 0) {
                return [int][math]::Min($MaxDelayMs, [math]::Max(1000, $retrySeconds * 1000))
            }
        }

        foreach ($headersPropertyName in @('ResponseHeaders', 'Headers')) {
            $headers = Get-SafeObjectPropertyValue -InputObject $cursor -PropertyName $headersPropertyName
            if ($null -ne $headers) {
                foreach ($headerName in @('Retry-After', 'x-ms-retry-after-ms')) {
                    $headerValue = Get-SafeObjectPropertyValue -InputObject $headers -PropertyName $headerName
                    $headerSeconds = 0
                    if ($null -ne $headerValue -and [int]::TryParse([string]$headerValue, [ref]$headerSeconds) -and $headerSeconds -gt 0) {
                        $multiplier = if ($headerName -eq 'x-ms-retry-after-ms') { 1 } else { 1000 }
                        return [int][math]::Min($MaxDelayMs, [math]::Max(1000, $headerSeconds * $multiplier))
                    }
                }
            }
        }

        $cursor = $cursor.InnerException
    }

    $exponentialDelayMs = [int]($BaseDelayMs * [math]::Pow(2, $Attempt - 1))
    $jitterMs = Get-Random -Minimum 250 -Maximum 1000
    return [int][math]::Min($MaxDelayMs, $exponentialDelayMs + $jitterMs)
}

function Invoke-WithVersioningDisabled {
    param(
        [string]$LibraryIdentity,
        [scriptblock]$Action
    )

    $list = Invoke-WithRetry -OperationName "Load library settings for '$LibraryIdentity'" -Action {
        Get-PnPList -Identity $LibraryIdentity -ErrorAction Stop
    }
    $ctx = Get-PnPContext
    $ctx.Load($list)
    $null = Invoke-WithRetry -OperationName "Load client context for '$LibraryIdentity'" -Action {
        Invoke-PnPQuery
    }
    $savedVersioning = $list.EnableVersioning
    $savedMinorVersions = $list.EnableMinorVersions
    $savedForceCheckout = $list.ForceCheckout
    $savedDraftVisibility = $list.DraftVersionVisibility

    Write-Status "Library settings before toggle: versioning=$savedVersioning, minorVersions=$savedMinorVersions, forceCheckout=$savedForceCheckout" 'INFO'

    if (-not $savedVersioning) {
        Write-Status "Versioning is already disabled on '$LibraryIdentity'; skipping temporary toggle." 'INFO'
        & $Action
        return
    }

    Write-Status "Disabling versioning on '$LibraryIdentity'" 'INFO'
    try {
        Invoke-WithRetry -OperationName "Disable versioning on '$LibraryIdentity'" -Action {
            Set-PnPList -Identity $LibraryIdentity -EnableVersioning $false -ErrorAction Stop | Out-Null
        }
        if ($savedForceCheckout) {
            $list.ForceCheckout = $false
            $list.Update()
            $null = Invoke-WithRetry -OperationName "Disable forced checkout on '$LibraryIdentity'" -Action {
                Invoke-PnPQuery
            }
        }

        & $Action
    }
    finally {
        try {
            Invoke-WithRetry -OperationName "Restore versioning on '$LibraryIdentity'" -Action {
                Set-PnPList -Identity $LibraryIdentity -EnableVersioning $savedVersioning -EnableMinorVersions $savedMinorVersions -ErrorAction Stop | Out-Null
            }
            $listRestore = Invoke-WithRetry -OperationName "Reload library settings for '$LibraryIdentity'" -Action {
                Get-PnPList -Identity $LibraryIdentity -ErrorAction Stop
            }
            $listRestore.ForceCheckout = $savedForceCheckout
            $listRestore.DraftVersionVisibility = $savedDraftVisibility
            $listRestore.Update()
            $null = Invoke-WithRetry -OperationName "Restore checkout settings on '$LibraryIdentity'" -Action {
                Invoke-PnPQuery
            }
            Write-Status "Restored versioning and checkout settings on '$LibraryIdentity'" 'SUCCESS'
        }
        catch {
            Write-Status "WARNING: Failed to restore versioning settings: $($_.Exception.Message)" 'ERROR'
            throw "Failed to restore versioning and checkout settings on '$LibraryIdentity': $($_.Exception.Message)"
        }
    }
}

function Get-TenantAdminUrl {
    param([string]$SiteUrl)

    if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
        throw 'SiteUrl is required to derive the tenant admin URL.'
    }

    $siteUri = [System.Uri]$SiteUrl
    if ($siteUri.Host -match '^[^.]+-admin\.') {
        return "$($siteUri.Scheme)://$($siteUri.Host)"
    }

    $hostParts = $siteUri.Host.Split('.')
    if ($hostParts.Count -lt 3 -or -not [string]::Equals($hostParts[1], 'sharepoint', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Could not derive the tenant admin URL from '$SiteUrl'."
    }

    $adminHost = "$($hostParts[0])-admin." + ($hostParts[1..($hostParts.Count - 1)] -join '.')
    return "$($siteUri.Scheme)://$adminHost"
}

function Get-SharePointHostKey {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        throw 'A SharePoint URL is required.'
    }

    return ([System.Uri]$Url).Host.ToLowerInvariant()
}

function Test-SharePointUnauthorizedException {
    param([System.Exception]$Exception)

    if ($null -eq $Exception) {
        return $false
    }

    $messageParts = [System.Collections.Generic.List[string]]::new()
    $cursor = $Exception
    while ($cursor) {
        if (-not [string]::IsNullOrWhiteSpace($cursor.Message)) {
            $messageParts.Add($cursor.Message)
        }
        $cursor = $cursor.InnerException
    }

    $combinedMessage = $messageParts -join ' | '
    return $combinedMessage -match '(?i)\bunauthorized\b' -or
        $combinedMessage -match '(?i)\b401\b' -or
        $combinedMessage -match '(?i)\bforbidden\b' -or
        $combinedMessage -match '(?i)\b403\b' -or
        $combinedMessage -match '(?i)access denied'
}

function Test-SharePointRetryableTokenRejectionException {
    param([System.Exception]$Exception)

    if ($null -eq $Exception) {
        return $false
    }

    $messageParts = [System.Collections.Generic.List[string]]::new()
    $cursor = $Exception
    while ($cursor) {
        if (-not [string]::IsNullOrWhiteSpace($cursor.Message)) {
            $messageParts.Add($cursor.Message)
        }
        $cursor = $cursor.InnerException
    }

    $combinedMessage = $messageParts -join ' | '
    $isAccessDenied =
        $combinedMessage -match '(?i)\bforbidden\b' -or
        $combinedMessage -match '(?i)\b403\b' -or
        $combinedMessage -match '(?i)access denied' -or
        $combinedMessage -match '(?i)attempted to perform an unauthorized operation'

    if ($isAccessDenied) {
        return $false
    }

    return $combinedMessage -match '(?i)\bunauthorized\b' -or
        $combinedMessage -match '(?i)\b401\b' -or
        $combinedMessage -match '(?i)token' -or
        $combinedMessage -match '(?i)expired' -or
        $combinedMessage -match '(?i)invalid[_ -]?jwt' -or
        $combinedMessage -match '(?i)invalid[_ -]?token'
}

function Get-SharePointAccessTokenExpirationUtc {
    param([string]$AccessToken)

    if ([string]::IsNullOrWhiteSpace($AccessToken)) {
        return $null
    }

    $tokenParts = $AccessToken.Split('.')
    if ($tokenParts.Count -lt 2) {
        return $null
    }

    $payload = $tokenParts[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
        1 { return $null }
    }

    try {
        $payloadJson = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload))
        $payloadObject = ConvertFrom-Json -InputObject $payloadJson -ErrorAction Stop
        $expValue = Get-SafeObjectPropertyValue -InputObject $payloadObject -PropertyName 'exp'
        if ($null -eq $expValue) {
            return $null
        }

        $unixSeconds = 0L
        if (-not [long]::TryParse([string]$expValue, [ref]$unixSeconds)) {
            return $null
        }

        return [System.DateTimeOffset]::FromUnixTimeSeconds($unixSeconds).UtcDateTime
    }
    catch {
        return $null
    }
}

function New-SharePointAccessTokenCacheEntry {
    param(
        [string]$AccessToken,
        [AllowNull()][string]$LastRejectedAccessToken = $null
    )

    return [pscustomobject]@{
        AccessToken             = $AccessToken
        ExpiresOnUtc            = Get-SharePointAccessTokenExpirationUtc -AccessToken $AccessToken
        LastRejectedAccessToken = $LastRejectedAccessToken
    }
}

function Test-SharePointAccessTokenNeedsRefresh {
    param(
        [AllowNull()][object]$CacheEntry,
        [int]$RefreshWindowMinutes = 5
    )

    if ($null -eq $CacheEntry) {
        return $true
    }

    $accessToken = [string](Get-SafeObjectPropertyValue -InputObject $CacheEntry -PropertyName 'AccessToken')
    if ([string]::IsNullOrWhiteSpace($accessToken)) {
        return $true
    }

    $expiresOnUtc = Get-SafeObjectPropertyValue -InputObject $CacheEntry -PropertyName 'ExpiresOnUtc'
    if ($expiresOnUtc -isnot [datetime]) {
        $expiresOnUtc = Get-SharePointAccessTokenExpirationUtc -AccessToken $accessToken
    }

    if ($expiresOnUtc -isnot [datetime]) {
        return $false
    }

    return $expiresOnUtc -le [datetime]::UtcNow.AddMinutes($RefreshWindowMinutes)
}

function Confirm-SharePointConnectionAccess {
    param([string]$Url)

    return (Invoke-WithRetry -OperationName "Confirm access to '$Url'" -Action {
            Get-PnPWeb -Includes Id, Title, Url, ServerRelativeUrl -ErrorAction Stop
        })
}

function Connect-SharePointOnlineCached {
    param(
        [string]$Url,
        [string]$ClientId,
        [switch]$ForceInteractive,
        [switch]$PassThru
    )

    $interactiveConnectParameters = @{
        Url         = $Url
        Interactive = $true
        ErrorAction = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
        $interactiveConnectParameters.ClientId = $ClientId
    }

    $hostKey = Get-SharePointHostKey -Url $Url
    $cacheEntry = if ($script:sharePointAccessTokenCache.ContainsKey($hostKey)) {
        $script:sharePointAccessTokenCache[$hostKey]
    }
    else {
        $null
    }

    if ($ForceInteractive -or (Test-SharePointAccessTokenNeedsRefresh -CacheEntry $cacheEntry)) {
        Connect-PnPOnline @interactiveConnectParameters
        $confirmedWeb = Confirm-SharePointConnectionAccess -Url $Url
        $script:sharePointAccessTokenCache[$hostKey] = New-SharePointAccessTokenCacheEntry -AccessToken (Get-PnPAccessToken)
        if ($PassThru) {
            return $confirmedWeb
        }
        return
    }

    try {
        Connect-PnPOnline -Url $Url -AccessToken $cacheEntry.AccessToken -ErrorAction Stop
        $confirmedWeb = Confirm-SharePointConnectionAccess -Url $Url
        if ($PassThru) {
            return $confirmedWeb
        }
        return
    }
    catch {
        if ((Test-SharePointRetryableTokenRejectionException -Exception $_.Exception) -and $cacheEntry.LastRejectedAccessToken -ne $cacheEntry.AccessToken) {
            Write-Status "Cached SharePoint token for host '$hostKey' was rejected. Refreshing interactive sign-in for '$Url'." 'WARN'
            $rejectedAccessToken = $cacheEntry.AccessToken
            try {
                Connect-PnPOnline @interactiveConnectParameters
                $confirmedWeb = Confirm-SharePointConnectionAccess -Url $Url
                $script:sharePointAccessTokenCache[$hostKey] = New-SharePointAccessTokenCacheEntry -AccessToken (Get-PnPAccessToken) -LastRejectedAccessToken $rejectedAccessToken
                if ($PassThru) {
                    return $confirmedWeb
                }
                return
            }
            catch {
                $script:sharePointAccessTokenCache[$hostKey] = New-SharePointAccessTokenCacheEntry -AccessToken $rejectedAccessToken -LastRejectedAccessToken $rejectedAccessToken
                throw
            }
        }

        throw
    }
}

function Get-TenantSiteCollectionUrls {
    $siteUrls = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $tenantSites = @(Invoke-WithRetry { Get-PnPTenantSite -Detailed -ErrorAction Stop })

    foreach ($tenantSite in $tenantSites) {
        $url = [string]$tenantSite.Url
        if ([string]::IsNullOrWhiteSpace($url)) {
            continue
        }

        try {
            $uri = [System.Uri]$url
        }
        catch {
            continue
        }

        if ($uri.Host -like '*-my.sharepoint.*') {
            continue
        }

        $template = [string](Get-SafeObjectPropertyValue -InputObject $tenantSite -PropertyName 'Template')
        if ($template -like 'SPSPERS*') {
            continue
        }

        $siteUrls.Add($url.TrimEnd('/')) | Out-Null
    }

    return @($siteUrls | Sort-Object)
}

function Import-SiteLibraryCsvTargets {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'SiteLibraryCsvPath cannot be empty.'
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "CSV file was not found: $Path"
    }

    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    $headerLine = @(
        Get-Content -LiteralPath $resolvedPath -ErrorAction Stop |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -First 1
    )[0]

    if ([string]::IsNullOrWhiteSpace($headerLine)) {
        throw "CSV file '$resolvedPath' is empty."
    }

    $delimiter = ','
    if ($headerLine.Contains(';') -and -not $headerLine.Contains(',')) {
        $delimiter = ';'
    }
    elseif ($headerLine.Contains("`t") -and -not $headerLine.Contains(',') -and -not $headerLine.Contains(';')) {
        $delimiter = "`t"
    }

    $rows = @(Import-Csv -LiteralPath $resolvedPath -Delimiter $delimiter -ErrorAction Stop)
    if ($rows.Count -eq 0) {
        throw "CSV file '$resolvedPath' does not contain any data rows."
    }

    $columnNames = @(Get-ObjectMemberNames -InputObject $rows[0])
    if ($columnNames -notcontains 'SiteURL') {
        $availableColumns = if ($columnNames.Count -gt 0) { $columnNames -join ', ' } else { 'none' }
        throw "CSV file '$resolvedPath' must contain a SiteURL column. Detected delimiter '$delimiter'. Available columns: $availableColumns"
    }

    $targets = [System.Collections.Generic.List[pscustomobject]]::new()
    $seenTargets = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $rowNumber = 1

    foreach ($row in $rows) {
        $rowNumber++

        $siteUrl = [string](Get-ObjectMemberValue -InputObject $row -MemberName 'SiteURL')
        if ([string]::IsNullOrWhiteSpace($siteUrl)) {
            throw "CSV row $rowNumber is missing a SiteURL value."
        }

        $siteUrl = $siteUrl.Trim()
        $siteUri = $null
        if (-not [System.Uri]::TryCreate($siteUrl, [System.UriKind]::Absolute, [ref]$siteUri)) {
            throw "CSV row $rowNumber has an invalid SiteURL value: $siteUrl"
        }

        if ([string]::IsNullOrWhiteSpace($siteUri.Host)) {
            throw "CSV row $rowNumber has an invalid SiteURL value: $siteUrl"
        }

        $normalizedSiteUrl = $siteUri.AbsoluteUri.TrimEnd('/')
        $docLib = [string](Get-ObjectMemberValue -InputObject $row -MemberName 'DocLib')
        $normalizedDocLib = if ([string]::IsNullOrWhiteSpace($docLib)) { $null } else { $docLib.Trim() }
        $dedupeKey = if ($null -eq $normalizedDocLib) { "$normalizedSiteUrl|*" } else { "$normalizedSiteUrl|$normalizedDocLib" }

        if (-not $seenTargets.Add($dedupeKey)) {
            continue
        }

        $targets.Add([pscustomobject]@{
                SiteUrl = $normalizedSiteUrl
                DocLib  = $normalizedDocLib
            })
    }

    if ($targets.Count -eq 0) {
        throw "CSV file '$resolvedPath' did not yield any usable site targets."
    }

    return @($targets)
}

function Get-SharePointWebInventory {
    param([string]$ClientId)

    $webs = [System.Collections.Generic.List[pscustomobject]]::new()
    $seenUrls = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $pendingWebUrls = [System.Collections.Generic.Queue[string]]::new()

    $rootWeb = Invoke-WithRetry -OperationName 'Load root web' -Action {
        Get-PnPWeb -Includes Title, Url, ServerRelativeUrl -ErrorAction Stop
    }
    $currentConnectedWebUrl = ([string]$rootWeb.Url).TrimEnd('/')

    if ($rootWeb -and $seenUrls.Add($currentConnectedWebUrl)) {
        $webs.Add([pscustomobject]@{
                Title             = [string]$rootWeb.Title
                Url               = $currentConnectedWebUrl
                ServerRelativeUrl = [string]$rootWeb.ServerRelativeUrl
            })
        $pendingWebUrls.Enqueue($currentConnectedWebUrl)
    }

    while ($pendingWebUrls.Count -gt 0) {
        $parentWebUrl = $pendingWebUrls.Dequeue()

        try {
            if (-not [string]::Equals($parentWebUrl, $currentConnectedWebUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
                $connectedWeb = Connect-SharePointOnlineCached -Url $parentWebUrl -ClientId $ClientId -PassThru
                $currentConnectedWebUrl = ([string]$connectedWeb.Url).TrimEnd('/')
            }

            $childWebs = @(Invoke-WithRetry -OperationName "Enumerate child webs under '$parentWebUrl'" -Action {
                    Get-PnPSubWeb -Includes Title, Url, ServerRelativeUrl -ErrorAction Stop
                })
        }
        catch {
            if (Test-SharePointUnauthorizedException -Exception $_.Exception) {
                Write-Status "Skipping child web discovery under '$parentWebUrl' because the current connection does not have access: $($_.Exception.Message)" 'WARN'
                continue
            }

            throw
        }

        foreach ($subWeb in $childWebs) {
            $subWebUrl = ([string]$subWeb.Url).TrimEnd('/')
            if ([string]::IsNullOrWhiteSpace($subWebUrl) -or -not $seenUrls.Add($subWebUrl)) {
                continue
            }

            $webs.Add([pscustomobject]@{
                    Title             = [string]$subWeb.Title
                    Url               = $subWebUrl
                    ServerRelativeUrl = [string]$subWeb.ServerRelativeUrl
                })
            $pendingWebUrls.Enqueue($subWebUrl)
        }
    }

    return @($webs)
}

function Get-VisibleDocumentLibraries {
    $lists = @(Invoke-WithRetry -OperationName 'Enumerate visible document libraries' -Action {
            Get-PnPList -Includes Title, Hidden, BaseTemplate -ErrorAction Stop
        })
    return @(
        $lists |
        Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden } |
        Sort-Object Title
    )
}

function Get-SharePointListFields {
    param([string]$ListIdentity)

    return @(Invoke-WithRetry -OperationName "Load fields for '$ListIdentity'" -Action {
            Get-PnPField -List $ListIdentity -ErrorAction Stop
        })
}

function Process-SharePointLibrary {
    [CmdletBinding()]
    param(
        [string]$LibraryName,
        [string]$WebUrl,
        [object[]]$Fields,
        [System.Collections.Generic.HashSet[string]]$AllowedNames,
        [hashtable]$AliasMap,
        [bool]$EnableUpdateOnOpen = $true,
        [switch]$SkipArchived,
        [Nullable[datetime]]$SkipCreatedAfter = $null,
        [string]$FileExtensionFilter = 'docx',
        [int]$PageSize = 500,
        [switch]$RetryFailed
    )

    $libraryResults = [System.Collections.Generic.List[pscustomobject]]::new()
    $librarySettings = Invoke-WithRetry -OperationName "Load settings for library '$LibraryName'" -Action {
        Get-PnPList -Identity $LibraryName -ErrorAction Stop
    }
    $versioningEnabled = [bool]$librarySettings.EnableVersioning

    Write-SectionHeader "Library Sweep - $LibraryName"

    $escapedFilter = [System.Security.SecurityElement]::Escape($FileExtensionFilter)
    $camlQuery = "<View Scope='RecursiveAll'><Query><Where>" +
    "<Eq><FieldRef Name='File_x0020_Type'/><Value Type='Text'>$escapedFilter</Value></Eq>" +
    "</Where></Query><RowLimit>$PageSize</RowLimit></View>"

    $items = @(Invoke-WithRetry { Get-PnPListItem -List $LibraryName -Query $camlQuery -PageSize $PageSize })
    Write-Status "Found $($items.Count) .$FileExtensionFilter file(s)" 'INFO'

    if ($RetryFailed) {
        if (-not (Test-Path $LogFilePath)) {
            throw "Log file not found at $LogFilePath."
        }

        $failedRefs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        Get-Content $LogFilePath | ForEach-Object {
            if ($_ -match '\[ERROR\]\s+(.*?\.docx)\s+- Error') {
                $failedRefs.Add($matches[1]) | Out-Null
            }
        }

        if ($failedRefs.Count -gt 0) {
            $items = @($items | Where-Object { $failedRefs.Contains([string]$_.FieldValues['FileRef']) })
            Write-Status "Retrying $($items.Count) failed file(s) found in log." 'INFO'
        }
        else {
            Write-Status 'No failed files found in log.' 'WARN'
        }
    }

    $majorItems = [System.Collections.Generic.List[object]]::new()
    $minorItems = [System.Collections.Generic.List[object]]::new()

    foreach ($item in $items) {
        $fileRef = [string]$item.FieldValues['FileRef']
        $leafName = [System.IO.Path]::GetFileName($fileRef)
        if ($leafName -like '~`$*') {
            continue
        }

        if ($SkipArchived) {
            $docStatus = [string]$item.FieldValues['ACTQMSDocumentStatus']
            if ($docStatus -eq 'Archived') {
                Write-Status "Skipping (Archived): $fileRef" 'WARN'
                $libraryResults.Add([pscustomobject]@{ FileRef = $fileRef; Status = 'Skipped-Archived'; QuickPartNames = @(); Updated = @(); Missing = @() })
                continue
            }
        }

        if ($null -ne $SkipCreatedAfter) {
            $createdDate = $item.FieldValues['Created']
            if ($createdDate -is [datetime] -and $createdDate -ge $SkipCreatedAfter) {
                Write-Status "Skipping (Created $($createdDate.ToString('yyyy-MM-dd'))): $fileRef" 'WARN'
                $libraryResults.Add([pscustomobject]@{ FileRef = $fileRef; Status = 'Skipped-CreatedAfter'; QuickPartNames = @(); Updated = @(); Missing = @() })
                continue
            }
        }

        $versionString = [string]$item.FieldValues['_UIVersionString']
        $isMajorVersion = ($versionString -match '^\d+\.0$' -and $versionString -ne '0.0')
        if ($versioningEnabled -and $isMajorVersion) {
            $majorItems.Add($item)
        }
        else {
            $minorItems.Add($item)
        }
    }

    Write-Status "Batches: $($majorItems.Count) major version(s), $($minorItems.Count) minor version(s)" 'INFO'
    $totalFiles = $majorItems.Count + $minorItems.Count
    $processed = 0
    $progressActivity = "Syncing Quick Parts - $LibraryName"

    if ($majorItems.Count -gt 0) {
        if ($WhatIfPreference) {
            Write-SectionHeader 'Pass 1 - Major versions (WhatIf)'

            foreach ($item in $majorItems) {
                $processed++
                $fileRef = [string]$item.FieldValues['FileRef']
                Write-Progress -Activity $progressActivity -Status $fileRef -PercentComplete (100 * $processed / [Math]::Max($totalFiles, 1))
                try {
                    Add-ResultRecords -Target $libraryResults -InputObject (Process-SharePointFile -FileRef $fileRef -WebUrl $WebUrl -Fields $Fields -AllowedNames $AllowedNames -AliasMap $AliasMap -EnableUpdateOnOpen:$EnableUpdateOnOpen -SkipArchived:$SkipArchived -SkipCreatedAfter $SkipCreatedAfter -ListItem $item) -SourceDescription "SharePoint file '$fileRef'"
                }
                catch {
                    $libraryResults.Add([pscustomobject]@{
                            FileRef        = $fileRef
                            Status         = "Error: $($_.Exception.Message)"
                            QuickPartNames = @()
                            Updated        = @()
                            Missing        = @()
                        })
                }
            }
        }
        else {
            Write-SectionHeader 'Pass 1 - Major versions (versioning disabled)'

            Invoke-WithVersioningDisabled -LibraryIdentity $LibraryName -Action {
                foreach ($item in $majorItems) {
                    $processed++
                    $fileRef = [string]$item.FieldValues['FileRef']
                    Write-Progress -Activity $progressActivity -Status $fileRef -PercentComplete (100 * $processed / [Math]::Max($totalFiles, 1))
                    try {
                        Add-ResultRecords -Target $libraryResults -InputObject (Process-SharePointFile -FileRef $fileRef -WebUrl $WebUrl -Fields $Fields -AllowedNames $AllowedNames -AliasMap $AliasMap -EnableUpdateOnOpen:$EnableUpdateOnOpen -VersioningDisabled -SkipArchived:$SkipArchived -SkipCreatedAfter $SkipCreatedAfter -ListItem $item) -SourceDescription "SharePoint file '$fileRef'"
                    }
                    catch {
                        $libraryResults.Add([pscustomobject]@{
                                FileRef        = $fileRef
                                Status         = "Error: $($_.Exception.Message)"
                                QuickPartNames = @()
                                Updated        = @()
                                Missing        = @()
                            })
                    }
                }
            }
        }
    }

    if ($minorItems.Count -gt 0) {
        Write-SectionHeader 'Pass 2 - Minor versions (OverwriteCheckIn)'

        foreach ($item in $minorItems) {
            $processed++
            $fileRef = [string]$item.FieldValues['FileRef']
            Write-Progress -Activity $progressActivity -Status $fileRef -PercentComplete (100 * $processed / [Math]::Max($totalFiles, 1))
            try {
                Add-ResultRecords -Target $libraryResults -InputObject (Process-SharePointFile -FileRef $fileRef -WebUrl $WebUrl -Fields $Fields -AllowedNames $AllowedNames -AliasMap $AliasMap -EnableUpdateOnOpen:$EnableUpdateOnOpen -SkipArchived:$SkipArchived -SkipCreatedAfter $SkipCreatedAfter -ListItem $item) -SourceDescription "SharePoint file '$fileRef'"
            }
            catch {
                $libraryResults.Add([pscustomobject]@{
                        FileRef        = $fileRef
                        Status         = "Error: $($_.Exception.Message)"
                        QuickPartNames = @()
                        Updated        = @()
                        Missing        = @()
                    })
            }
        }
    }

    Write-Progress -Activity $progressActivity -Completed
    return @($libraryResults)
}

function Process-SharePointFile {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$FileRef,
        [string]$WebUrl,
        [object[]]$Fields,
        [System.Collections.Generic.HashSet[string]]$AllowedNames,
        [hashtable]$AliasMap,
        [bool]$EnableUpdateOnOpen = $true,
        [switch]$VersioningDisabled,
        [switch]$SkipArchived,
        [Nullable[datetime]]$SkipCreatedAfter = $null,
        $ListItem
    )

    $FileRef = [System.Uri]::UnescapeDataString($FileRef)
    $previousLogContext = $script:CurrentLogContext
    $script:CurrentLogContext = $FileRef

    try {
        if ($FileRef -notmatch '(?i)\.docx$') {
            return [pscustomobject]@{ FileRef = $FileRef; Status = 'Skipped-NotDocx'; QuickPartNames = @(); Updated = @(); Missing = @() }
        }

        if (-not $ListItem) {
            $ListItem = Invoke-WithRetry -OperationName "Load list item for '$FileRef'" -Action {
                Get-PnPFile -Url $FileRef -AsListItem -ErrorAction Stop
            }
        }
        if ($SkipArchived) {
            $docStatus = [string]$ListItem.FieldValues['ACTQMSDocumentStatus']
            if ($docStatus -eq 'Archived') {
                Write-Status "Skipping (Archived): $FileRef" 'WARN'
                return [pscustomobject]@{ FileRef = $FileRef; Status = 'Skipped-Archived'; QuickPartNames = @(); Updated = @(); Missing = @() }
            }
        }
        if ($null -ne $SkipCreatedAfter) {
            $createdDate = $ListItem.FieldValues['Created']
            if ($createdDate -is [datetime] -and $createdDate -ge $SkipCreatedAfter) {
                Write-Status "Skipping (Created $($createdDate.ToString('yyyy-MM-dd'))): $FileRef" 'WARN'
                return [pscustomobject]@{ FileRef = $FileRef; Status = 'Skipped-CreatedAfter'; QuickPartNames = @(); Updated = @(); Missing = @() }
            }
        }

        Write-Status "Processing: $FileRef" 'INFO'

        $resolver = New-PropertyResolver -ListItem $ListItem -Fields $Fields -AllowedNames $AllowedNames -AliasMap $AliasMap
        $propertyValues = @{}
        foreach ($key in $resolver.Keys) {
            $propertyValues[$key] = $resolver[$key]
        }

        $memoryStream = Invoke-WithRetry { Get-PnPFile -Url $FileRef -AsMemoryStream -ErrorAction Stop }
        try {
            $docxBytes = $memoryStream.ToArray()
        }
        finally {
            $memoryStream.Dispose()
        }
        Write-Host "    Downloaded  : $([math]::Round($docxBytes.Length / 1KB, 1)) KB"

        if ($docxBytes.Length -lt 4 -or $docxBytes[0] -ne 0x50 -or $docxBytes[1] -ne 0x4B) {
            Write-Status '  Not a valid ZIP - skipping.' 'WARN'
            return [pscustomobject]@{ FileRef = $FileRef; Status = 'Skipped-NotZip'; QuickPartNames = @(); Updated = @(); Missing = @() }
        }

        $result = Update-DocxQuickPartsFromBytes -DocxBytes $docxBytes -PropertyValues $propertyValues -AllowedNames $AllowedNames -EnableUpdateOnOpen:$EnableUpdateOnOpen

        if (-not $result.Changed) {
            return [pscustomobject]@{
                FileRef        = $FileRef
                Status         = 'Clean'
                QuickPartNames = @($result.QuickPartNames + $result.BoundNames | Select-Object -Unique)
                Updated        = @()
                Missing        = @($result.MissingProperties + $result.FallbackBoundNames | Select-Object -Unique)
            }
        }

        if (-not $PSCmdlet.ShouldProcess($FileRef, 'Sync Quick Parts from SharePoint metadata')) {
            return [pscustomobject]@{
                FileRef        = $FileRef
                Status         = 'WhatIf'
                QuickPartNames = @($result.QuickPartNames + $result.BoundNames | Select-Object -Unique)
                Updated        = @($result.UpdatedProperties + $result.AddedProperties + $result.SyncedBoundNames | Select-Object -Unique)
                Missing        = @($result.MissingProperties + $result.FallbackBoundNames | Select-Object -Unique)
            }
        }

        $ok = Upload-DocxBytesBack -Bytes $result.Bytes -FileRef $FileRef -WebUrl $WebUrl -ExistingListItem $ListItem -Fields $Fields -VersioningDisabled:$VersioningDisabled

        if (-not $ok) {
            return [pscustomobject]@{
                FileRef        = $FileRef
                Status         = 'Error: CheckinFailed'
                QuickPartNames = @($result.QuickPartNames + $result.BoundNames | Select-Object -Unique)
                Updated        = @($result.UpdatedProperties + $result.AddedProperties + $result.SyncedBoundNames | Select-Object -Unique)
                Missing        = @($result.MissingProperties + $result.FallbackBoundNames | Select-Object -Unique)
            }
        }

        return [pscustomobject]@{
            FileRef        = $FileRef
            Status         = 'Fixed'
            QuickPartNames = @($result.QuickPartNames + $result.BoundNames | Select-Object -Unique)
            Updated        = @($result.UpdatedProperties + $result.AddedProperties + $result.SyncedBoundNames | Select-Object -Unique)
            Missing        = @($result.MissingProperties + $result.FallbackBoundNames | Select-Object -Unique)
        }
    }
    finally {
        $script:CurrentLogContext = $previousLogContext
    }
}

function Process-LocalFile {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$FullPath,
        [System.Collections.Generic.HashSet[string]]$AllowedNames,
        [hashtable]$PropertyValue,
        [bool]$EnableUpdateOnOpen = $true,
        [switch]$Overwrite
    )

    $previousLogContext = $script:CurrentLogContext
    $script:CurrentLogContext = $FullPath

    try {
        if ($FullPath -notmatch '(?i)\.docx$') {
            return [pscustomobject]@{ FileRef = $FullPath; Status = 'Skipped-NotDocx'; QuickPartNames = @(); Updated = @(); Missing = @() }
        }

        $docxBytes = [System.IO.File]::ReadAllBytes($FullPath)
        Write-Host "    Read: $([math]::Round($docxBytes.Length / 1KB, 1)) KB"

        if ($docxBytes.Length -lt 4 -or $docxBytes[0] -ne 0x50 -or $docxBytes[1] -ne 0x4B) {
            Write-Status '  Not a valid ZIP - skipping.' 'WARN'
            return [pscustomobject]@{ FileRef = $FullPath; Status = 'Skipped-NotZip'; QuickPartNames = @(); Updated = @(); Missing = @() }
        }

        $propertyValues = @{}
        if ($PropertyValue) {
            foreach ($key in $PropertyValue.Keys) {
                if ($AllowedNames.Contains([string]$key)) {
                    $propertyValues[[string]$key] = $PropertyValue[$key]
                }
            }
        }

        $result = Update-DocxQuickPartsFromBytes -DocxBytes $docxBytes -PropertyValues $propertyValues -AllowedNames $AllowedNames -EnableUpdateOnOpen:$EnableUpdateOnOpen

        if (-not $result.Changed) {
            return [pscustomobject]@{
                FileRef        = $FullPath
                Status         = 'Clean'
                QuickPartNames = @($result.QuickPartNames + $result.BoundNames | Select-Object -Unique)
                Updated        = @()
                Missing        = $result.MissingProperties
            }
        }

        if ($Overwrite) {
            $outPath = $FullPath
        }
        else {
            $directory = [System.IO.Path]::GetDirectoryName($FullPath)
            $name = [System.IO.Path]::GetFileNameWithoutExtension($FullPath)
            $outPath = Join-Path $directory "$name.quickparts.docx"
        }

        return [pscustomobject]@{
            FileRef        = $FullPath
            Status         = if (-not $PSCmdlet.ShouldProcess($outPath, 'Sync Quick Parts and enable refresh on open')) { 'WhatIf' } else {
                [System.IO.File]::WriteAllBytes($outPath, $result.Bytes)
                'Fixed'
            }
            QuickPartNames = @($result.QuickPartNames + $result.BoundNames | Select-Object -Unique)
            Updated        = @($result.UpdatedProperties + $result.AddedProperties + $result.BoundNames | Select-Object -Unique)
            Missing        = $result.MissingProperties
        }
    }
    finally {
        $script:CurrentLogContext = $previousLogContext
    }
}

try {
    $results = [System.Collections.Generic.List[pscustomobject]]::new()
    $propertyAliasMap = Get-PropertyAliasMap
    $fieldSelection = Get-SelectedFieldConfiguration `
        -FieldName $FieldName `
        -FieldNameWasProvided:$PSBoundParameters.ContainsKey('FieldName') `
        -FieldConfigPath $FieldConfigPath `
        -FieldProfile $FieldProfile
    $allowedNameLookup = New-NameLookup -Names (Expand-AllowedPropertyNames -Names $fieldSelection.FieldNames -AliasMap $propertyAliasMap)

    Write-Status "Field source: $($fieldSelection.SourceDescription)" 'INFO'
    Write-Status "Tracking $($fieldSelection.FieldNames.Count) field(s): $($fieldSelection.FieldNames -join ', ')" 'INFO'

    if ($PSCmdlet.ParameterSetName -eq 'Local') {
        Write-SectionHeader 'Local File - Sync Quick Parts'

        if (Test-Path -Path $LocalPath -PathType Container) {
            $files = @(Get-ChildItem -Path $LocalPath -Filter '*.docx' -File -Recurse |
                Where-Object { $_.Name -notlike '~`$*' })

            Write-Status "Found $($files.Count) .docx file(s) in '$LocalPath'" 'INFO'

            $index = 0
            foreach ($file in $files) {
                $index++
                Write-Progress -Activity 'Syncing Quick Parts' -Status $file.Name -PercentComplete (100 * $index / [Math]::Max($files.Count, 1))
                try {
                    Add-ResultRecords -Target $results -InputObject (Process-LocalFile -FullPath $file.FullName -AllowedNames $allowedNameLookup -PropertyValue $PropertyValue -EnableUpdateOnOpen:$EnableUpdateOnOpen -Overwrite:$Overwrite) -SourceDescription "local file '$($file.FullName)'"
                }
                catch {
                    $results.Add([pscustomobject]@{
                            FileRef        = $file.FullName
                            Status         = "Error: $($_.Exception.Message)"
                            QuickPartNames = @()
                            Updated        = @()
                            Missing        = @()
                        })
                }
            }

            Write-Progress -Activity 'Syncing Quick Parts' -Completed
        }
        else {
            $resolvedPath = (Resolve-Path $LocalPath).Path
            Add-ResultRecords -Target $results -InputObject (Process-LocalFile -FullPath $resolvedPath -AllowedNames $allowedNameLookup -PropertyValue $PropertyValue -EnableUpdateOnOpen:$EnableUpdateOnOpen -Overwrite:$Overwrite) -SourceDescription "local file '$resolvedPath'"
        }
    }
    else {
        Import-Module PnP.PowerShell -ErrorAction Stop

        if ($PSCmdlet.ParameterSetName -eq 'SPTenantSweep') {
            $tenantAdminUrl = Get-TenantAdminUrl -SiteUrl $SiteUrl
            Write-SectionHeader 'Tenant Sweep - All Sites and Libraries'
            Write-Status "Connecting to tenant admin $tenantAdminUrl" 'INFO'
            Connect-SharePointOnlineCached -Url $tenantAdminUrl -ClientId $ClientId -ForceInteractive
            Write-Status 'Connected to tenant admin.' 'SUCCESS'

            $siteCollectionUrls = @(Get-TenantSiteCollectionUrls)
            Write-Status "Found $($siteCollectionUrls.Count) site collection(s) in tenant scope" 'INFO'

            foreach ($siteCollectionUrl in $siteCollectionUrls) {
                Write-SectionHeader "Site Collection - $siteCollectionUrl"

                try {
                    Connect-SharePointOnlineCached -Url $siteCollectionUrl -ClientId $ClientId
                    $webInfos = @(Get-SharePointWebInventory -ClientId $ClientId)
                    Write-Status "Found $($webInfos.Count) web(s)" 'INFO'
                }
                catch {
                    $status = if (Test-SharePointUnauthorizedException -Exception $_.Exception) { 'Skipped-Unauthorized' } else { "Error: $($_.Exception.Message)" }
                    $results.Add([pscustomobject]@{
                            FileRef        = $siteCollectionUrl
                            Status         = $status
                            QuickPartNames = @()
                            Updated        = @()
                            Missing        = @()
                        })
                    if ($status -eq 'Skipped-Unauthorized') {
                        Write-Status "Skipping site collection '$siteCollectionUrl' because the current connection does not have access: $($_.Exception.Message)" 'WARN'
                    }
                    else {
                        Write-Status "Failed to inventory site collection '$siteCollectionUrl': $($_.Exception.Message)" 'ERROR'
                    }
                    continue
                }

                foreach ($webInfo in $webInfos) {
                    Write-SectionHeader "Web - $($webInfo.Url)"

                    try {
                        $currentWeb = Connect-SharePointOnlineCached -Url $webInfo.Url -ClientId $ClientId -PassThru
                        $currentWebUrl = $currentWeb.ServerRelativeUrl.TrimEnd('/')
                        $libraries = @(Get-VisibleDocumentLibraries)
                    }
                    catch {
                        $status = if (Test-SharePointUnauthorizedException -Exception $_.Exception) { 'Skipped-Unauthorized' } else { "Error: $($_.Exception.Message)" }
                        $results.Add([pscustomobject]@{
                                FileRef        = $webInfo.Url
                                Status         = $status
                                QuickPartNames = @()
                                Updated        = @()
                                Missing        = @()
                            })
                        if ($status -eq 'Skipped-Unauthorized') {
                            Write-Status "Skipping web '$($webInfo.Url)' because the current connection does not have access: $($_.Exception.Message)" 'WARN'
                        }
                        else {
                            Write-Status "Failed to enumerate libraries for '$($webInfo.Url)': $($_.Exception.Message)" 'ERROR'
                        }
                        continue
                    }

                    if ($libraries.Count -eq 0) {
                        Write-Status 'No visible document libraries found.' 'WARN'
                        continue
                    }

                    Write-Status "Found $($libraries.Count) visible document librar$(if ($libraries.Count -eq 1) { 'y' } else { 'ies' })" 'INFO'

                    foreach ($library in $libraries) {
                        $fields = @()
                        try {
                            $fields = @(Get-SharePointListFields -ListIdentity $library.Title)
                        }
                        catch {
                            Write-Status "Could not load library fields for '$($library.Title)' in '$($webInfo.Url)': $($_.Exception.Message)" 'WARN'
                        }

                        try {
                            Add-ResultRecords -Target $results -InputObject (Process-SharePointLibrary -LibraryName $library.Title -WebUrl $currentWebUrl -Fields $fields -AllowedNames $allowedNameLookup -AliasMap $propertyAliasMap -EnableUpdateOnOpen:$EnableUpdateOnOpen -SkipArchived:$SkipArchived -SkipCreatedAfter $SkipCreatedAfter -FileExtensionFilter $FileExtensionFilter -PageSize $PageSize -RetryFailed:$RetryFailed) -SourceDescription "library '$($library.Title)' in '$($webInfo.Url)'"
                        }
                        catch {
                            $results.Add([pscustomobject]@{
                                    FileRef        = "$($webInfo.Url) [$($library.Title)]"
                                    Status         = "Error: $($_.Exception.Message)"
                                    QuickPartNames = @()
                                    Updated        = @()
                                    Missing        = @()
                                })
                            Write-Status "Library sweep failed for '$($library.Title)' in '$($webInfo.Url)': $($_.Exception.Message)" 'ERROR'
                        }
                    }
                }
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'SPCsvSweep') {
            $csvTargets = @(Import-SiteLibraryCsvTargets -Path $SiteLibraryCsvPath)
            $siteGroups = @($csvTargets | Group-Object SiteUrl)
            $seedSiteUrl = $csvTargets[0].SiteUrl

            Write-SectionHeader 'CSV Sweep - Listed Sites and Libraries'
            Write-Status "Loaded $($csvTargets.Count) CSV target row(s) across $($siteGroups.Count) site(s)" 'INFO'
            Write-Status "Connecting to seed site $seedSiteUrl" 'INFO'
            Connect-SharePointOnlineCached -Url $seedSiteUrl -ClientId $ClientId -ForceInteractive
            Write-Status 'Connected.' 'SUCCESS'

            foreach ($siteGroup in $siteGroups) {
                $targetSiteUrl = [string]$siteGroup.Name

                Write-SectionHeader "Site - $targetSiteUrl"

                try {
                    $web = Connect-SharePointOnlineCached -Url $targetSiteUrl -ClientId $ClientId -PassThru
                    $webUrl = $web.ServerRelativeUrl.TrimEnd('/')
                }
                catch {
                    $status = if (Test-SharePointUnauthorizedException -Exception $_.Exception) { 'Skipped-Unauthorized' } else { "Error: $($_.Exception.Message)" }
                    $results.Add([pscustomobject]@{
                            FileRef        = $targetSiteUrl
                            Status         = $status
                            QuickPartNames = @()
                            Updated        = @()
                            Missing        = @()
                        })
                    if ($status -eq 'Skipped-Unauthorized') {
                        Write-Status "Skipping site '$targetSiteUrl' because the current connection does not have access: $($_.Exception.Message)" 'WARN'
                    }
                    else {
                        Write-Status "Failed to connect to site '$targetSiteUrl': $($_.Exception.Message)" 'ERROR'
                    }
                    continue
                }

                $groupRows = @($siteGroup.Group)
                $processAllLibraries = @($groupRows | Where-Object { [string]::IsNullOrWhiteSpace($_.DocLib) }).Count -gt 0

                if ($processAllLibraries) {
                    $libraries = @(Get-VisibleDocumentLibraries)
                    if ($libraries.Count -eq 0) {
                        Write-Status 'No visible document libraries found on this site.' 'WARN'
                        continue
                    }

                    Write-Status "Processing all visible document librar$(if ($libraries.Count -eq 1) { 'y' } else { 'ies' }) on '$targetSiteUrl' from CSV site row(s)" 'INFO'
                    foreach ($library in $libraries) {
                        $libraryFields = @()
                        try {
                            $libraryFields = @(Get-SharePointListFields -ListIdentity $library.Title)
                        }
                        catch {
                            Write-Status "Could not load library fields for '$($library.Title)' on '$targetSiteUrl': $($_.Exception.Message)" 'WARN'
                        }

                        Add-ResultRecords -Target $results -InputObject (Process-SharePointLibrary -LibraryName $library.Title -WebUrl $webUrl -Fields $libraryFields -AllowedNames $allowedNameLookup -AliasMap $propertyAliasMap -EnableUpdateOnOpen:$EnableUpdateOnOpen -SkipArchived:$SkipArchived -SkipCreatedAfter $SkipCreatedAfter -FileExtensionFilter $FileExtensionFilter -PageSize $PageSize -RetryFailed:$RetryFailed) -SourceDescription "library '$($library.Title)' from CSV site '$targetSiteUrl'"
                    }

                    continue
                }

                $requestedLibraries = @($groupRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.DocLib) } | Select-Object -ExpandProperty DocLib -Unique)
                foreach ($requestedLibrary in $requestedLibraries) {
                    $libraryFields = @()
                    try {
                        $library = Invoke-WithRetry -OperationName "Load library '$requestedLibrary' on '$targetSiteUrl'" -Action {
                            Get-PnPList -Identity $requestedLibrary -Includes Title, Hidden, BaseTemplate -ErrorAction Stop
                        }
                    }
                    catch {
                        $results.Add([pscustomobject]@{
                                FileRef        = "$targetSiteUrl [$requestedLibrary]"
                                Status         = "Error: $($_.Exception.Message)"
                                QuickPartNames = @()
                                Updated        = @()
                                Missing        = @()
                            })
                        Write-Status "Failed to find library '$requestedLibrary' on '$targetSiteUrl': $($_.Exception.Message)" 'ERROR'
                        continue
                    }

                    if ($library.BaseTemplate -ne 101) {
                        $results.Add([pscustomobject]@{
                                FileRef        = "$targetSiteUrl [$requestedLibrary]"
                                Status         = 'Skipped-NotDocumentLibrary'
                                QuickPartNames = @()
                                Updated        = @()
                                Missing        = @()
                            })
                        Write-Status "Skipping '$requestedLibrary' on '$targetSiteUrl' because it is not a document library." 'WARN'
                        continue
                    }

                    try {
                        $libraryFields = @(Get-SharePointListFields -ListIdentity $library.Title)
                    }
                    catch {
                        Write-Status "Could not load library fields for '$($library.Title)': $($_.Exception.Message)" 'WARN'
                    }

                    Add-ResultRecords -Target $results -InputObject (Process-SharePointLibrary -LibraryName $library.Title -WebUrl $webUrl -Fields $libraryFields -AllowedNames $allowedNameLookup -AliasMap $propertyAliasMap -EnableUpdateOnOpen:$EnableUpdateOnOpen -SkipArchived:$SkipArchived -SkipCreatedAfter $SkipCreatedAfter -FileExtensionFilter $FileExtensionFilter -PageSize $PageSize -RetryFailed:$RetryFailed) -SourceDescription "library '$($library.Title)' from CSV site '$targetSiteUrl'"
                }
            }
        }
        else {
            Write-Status "Connecting to $SiteUrl" 'INFO'
            $web = Connect-SharePointOnlineCached -Url $SiteUrl -ClientId $ClientId -ForceInteractive -PassThru
            Write-Status 'Connected.' 'SUCCESS'

            $webUrl = $web.ServerRelativeUrl.TrimEnd('/')

            $fieldLibraryName = $LibraryName
            if ($PSCmdlet.ParameterSetName -eq 'SPSingleFile') {
                $decodedRef = [System.Uri]::UnescapeDataString($FileServerRelativeUrl)
                $relativePath = $decodedRef
                if ($webUrl -ne '' -and $relativePath.StartsWith($webUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $relativePath = $relativePath.Substring($webUrl.Length)
                }
                $relativePath = $relativePath.TrimStart('/')
                $fieldLibraryName = [System.Uri]::UnescapeDataString($relativePath.Split('/')[0])
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'SPSiteLibraries') {
                $fieldLibraryName = $null
            }

            $fields = @()
            if (-not [string]::IsNullOrWhiteSpace($fieldLibraryName)) {
                try {
                    $fields = @(Get-SharePointListFields -ListIdentity $fieldLibraryName)
                }
                catch {
                    Write-Status "Could not load library fields for '$fieldLibraryName': $($_.Exception.Message)" 'WARN'
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'SPSingleFile') {
                Write-SectionHeader 'Single File - Sync Quick Parts'

                $decodedRef = [System.Uri]::UnescapeDataString($FileServerRelativeUrl)
                $spItem = Invoke-WithRetry -OperationName "Load single file metadata for '$decodedRef'" -Action {
                    Get-PnPFile -Url $decodedRef -AsListItem -ErrorAction Stop
                }

                # Check skip conditions before potentially disabling versioning
                $skipSingleFile = $false
                if ($SkipArchived) {
                    $docStatus = [string]$spItem.FieldValues['ACTQMSDocumentStatus']
                    if ($docStatus -eq 'Archived') {
                        Write-Status "Skipping (Archived): $decodedRef" 'WARN'
                        $results.Add([pscustomobject]@{ FileRef = $decodedRef; Status = 'Skipped-Archived'; QuickPartNames = @(); Updated = @(); Missing = @() })
                        $skipSingleFile = $true
                    }
                }
                if (-not $skipSingleFile -and $null -ne $SkipCreatedAfter) {
                    $createdDate = $spItem.FieldValues['Created']
                    if ($createdDate -is [datetime] -and $createdDate -ge $SkipCreatedAfter) {
                        Write-Status "Skipping (Created $($createdDate.ToString('yyyy-MM-dd'))): $decodedRef" 'WARN'
                        $results.Add([pscustomobject]@{ FileRef = $decodedRef; Status = 'Skipped-CreatedAfter'; QuickPartNames = @(); Updated = @(); Missing = @() })
                        $skipSingleFile = $true
                    }
                }

                if (-not $skipSingleFile) {
                    $relativePath = $decodedRef
                    if ($webUrl -ne '' -and $relativePath.StartsWith($webUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $relativePath = $relativePath.Substring($webUrl.Length)
                    }
                    $relativePath = $relativePath.TrimStart('/')
                    $libraryIdentity = [System.Uri]::UnescapeDataString($relativePath.Split('/')[0])
                    if (-not $libraryIdentity) {
                        throw "Could not infer library name from path '$decodedRef' (web='$webUrl')."
                    }

                    $singleFileLibrary = Invoke-WithRetry -OperationName "Load library '$libraryIdentity'" -Action {
                        Get-PnPList -Identity $libraryIdentity -ErrorAction Stop
                    }
                    $versioningEnabled = [bool]$singleFileLibrary.EnableVersioning
                    $versionString = [string]$spItem.FieldValues['_UIVersionString']
                    $isMajorVersion = ($versionString -match '^\d+\.0$' -and $versionString -ne '0.0')

                    if ($versioningEnabled -and $isMajorVersion -and -not $WhatIfPreference) {

                        Write-Status 'Major version detected - temporarily disabling versioning' 'INFO'
                        Invoke-WithVersioningDisabled -LibraryIdentity $libraryIdentity -Action {
                            Add-ResultRecords -Target $results -InputObject (Process-SharePointFile -FileRef $decodedRef -WebUrl $webUrl -Fields $fields -AllowedNames $allowedNameLookup -AliasMap $propertyAliasMap -EnableUpdateOnOpen:$EnableUpdateOnOpen -VersioningDisabled -SkipArchived:$false -SkipCreatedAfter $null -ListItem $spItem) -SourceDescription "SharePoint file '$decodedRef'"
                        }
                    }
                    else {
                        Add-ResultRecords -Target $results -InputObject (Process-SharePointFile -FileRef $decodedRef -WebUrl $webUrl -Fields $fields -AllowedNames $allowedNameLookup -AliasMap $propertyAliasMap -EnableUpdateOnOpen:$EnableUpdateOnOpen -SkipArchived:$false -SkipCreatedAfter $null -ListItem $spItem) -SourceDescription "SharePoint file '$decodedRef'"
                    }
                }
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'SPLibrary') {
                Add-ResultRecords -Target $results -InputObject (Process-SharePointLibrary -LibraryName $LibraryName -WebUrl $webUrl -Fields $fields -AllowedNames $allowedNameLookup -AliasMap $propertyAliasMap -EnableUpdateOnOpen:$EnableUpdateOnOpen -SkipArchived:$SkipArchived -SkipCreatedAfter $SkipCreatedAfter -FileExtensionFilter $FileExtensionFilter -PageSize $PageSize -RetryFailed:$RetryFailed) -SourceDescription "library '$LibraryName'"
            }
            else {
                if ($IncludeSubsites) {
                    Write-SectionHeader 'Site Sweep - Site and Subsites'

                    $webInfos = @(Get-SharePointWebInventory -ClientId $ClientId)
                    Write-Status "Found $($webInfos.Count) web(s)" 'INFO'

                    foreach ($webInfo in $webInfos) {
                        Write-SectionHeader "Web - $($webInfo.Url)"

                        try {
                            $currentWeb = Connect-SharePointOnlineCached -Url $webInfo.Url -ClientId $ClientId -PassThru
                            $currentWebUrl = $currentWeb.ServerRelativeUrl.TrimEnd('/')
                            $libraries = @(Get-VisibleDocumentLibraries)
                        }
                        catch {
                            $status = if (Test-SharePointUnauthorizedException -Exception $_.Exception) { 'Skipped-Unauthorized' } else { "Error: $($_.Exception.Message)" }
                            $results.Add([pscustomobject]@{
                                    FileRef        = $webInfo.Url
                                    Status         = $status
                                    QuickPartNames = @()
                                    Updated        = @()
                                    Missing        = @()
                                })
                            if ($status -eq 'Skipped-Unauthorized') {
                                Write-Status "Skipping web '$($webInfo.Url)' because the current connection does not have access: $($_.Exception.Message)" 'WARN'
                            }
                            else {
                                Write-Status "Failed to enumerate libraries for '$($webInfo.Url)': $($_.Exception.Message)" 'ERROR'
                            }
                            continue
                        }

                        if ($libraries.Count -eq 0) {
                            Write-Status 'No visible document libraries found.' 'WARN'
                            continue
                        }

                        Write-Status "Found $($libraries.Count) visible document librar$(if ($libraries.Count -eq 1) { 'y' } else { 'ies' })" 'INFO'
                        foreach ($library in $libraries) {
                            $libraryFields = @()
                            try {
                                $libraryFields = @(Get-SharePointListFields -ListIdentity $library.Title)
                            }
                            catch {
                                Write-Status "Could not load library fields for '$($library.Title)' in '$($webInfo.Url)': $($_.Exception.Message)" 'WARN'
                            }

                            Add-ResultRecords -Target $results -InputObject (Process-SharePointLibrary -LibraryName $library.Title -WebUrl $currentWebUrl -Fields $libraryFields -AllowedNames $allowedNameLookup -AliasMap $propertyAliasMap -EnableUpdateOnOpen:$EnableUpdateOnOpen -SkipArchived:$SkipArchived -SkipCreatedAfter $SkipCreatedAfter -FileExtensionFilter $FileExtensionFilter -PageSize $PageSize -RetryFailed:$RetryFailed) -SourceDescription "library '$($library.Title)' in '$($webInfo.Url)'"
                        }
                    }
                }
                else {
                    Write-SectionHeader 'Site Sweep - All Document Libraries'

                    $libraries = @(Get-VisibleDocumentLibraries)
                    if ($libraries.Count -eq 0) {
                        Write-Status 'No visible document libraries found on this site.' 'WARN'
                    }
                    else {
                        Write-Status "Found $($libraries.Count) visible document librar$(if ($libraries.Count -eq 1) { 'y' } else { 'ies' })" 'INFO'
                        foreach ($library in $libraries) {
                            $libraryFields = @()
                            try {
                                $libraryFields = @(Get-SharePointListFields -ListIdentity $library.Title)
                            }
                            catch {
                                Write-Status "Could not load library fields for '$($library.Title)': $($_.Exception.Message)" 'WARN'
                            }

                            Add-ResultRecords -Target $results -InputObject (Process-SharePointLibrary -LibraryName $library.Title -WebUrl $webUrl -Fields $libraryFields -AllowedNames $allowedNameLookup -AliasMap $propertyAliasMap -EnableUpdateOnOpen:$EnableUpdateOnOpen -SkipArchived:$SkipArchived -SkipCreatedAfter $SkipCreatedAfter -FileExtensionFilter $FileExtensionFilter -PageSize $PageSize -RetryFailed:$RetryFailed) -SourceDescription "library '$($library.Title)'"
                        }
                    }
                }
            }
        }
    }

    Write-SectionHeader 'Summary'

    $fixedCount = @($results | Where-Object { $_.Status -eq 'Fixed' }).Count
    $cleanCount = @($results | Where-Object { $_.Status -eq 'Clean' }).Count
    $archivedCount = @($results | Where-Object { $_.Status -eq 'Skipped-Archived' }).Count
    $createdAfterCount = @($results | Where-Object { $_.Status -eq 'Skipped-CreatedAfter' }).Count
    $unauthorizedCount = @($results | Where-Object { $_.Status -eq 'Skipped-Unauthorized' }).Count
    $skippedCount = @($results | Where-Object { $_.Status -like 'Skipped*' -and $_.Status -ne 'Skipped-Archived' -and $_.Status -ne 'Skipped-CreatedAfter' -and $_.Status -ne 'Skipped-Unauthorized' }).Count
    $whatIfCount = @($results | Where-Object { $_.Status -eq 'WhatIf' }).Count
    $errorCount = @($results | Where-Object { $_.Status -like 'Error*' }).Count

    Write-Host "  Fixed         : $fixedCount"
    Write-Host "  Clean         : $cleanCount"
    Write-Host "  Archived      : $archivedCount"
    Write-Host "  CreatedAfter  : $createdAfterCount"
    Write-Host "  Unauthorized  : $unauthorizedCount"
    Write-Host "  Skipped       : $skippedCount"
    Write-Host "  WhatIf        : $whatIfCount"
    Write-Host "  Errors        : $errorCount"

    if ($unauthorizedCount -gt 0) {
        Write-Host ''
        Write-Status 'Sites or webs skipped due to missing access:' 'WARN'
        foreach ($result in $results | Where-Object { $_.Status -eq 'Skipped-Unauthorized' }) {
            Write-Host "  $($result.FileRef)" -ForegroundColor Yellow
        }
    }

    if ($fixedCount -gt 0) {
        Write-Host ''
        Write-Status 'Files updated:' 'SUCCESS'
        foreach ($result in $results | Where-Object { $_.Status -eq 'Fixed' }) {
            Write-Host "  $($result.FileRef)" -ForegroundColor Green
            if ($result.QuickPartNames.Count -gt 0) {
                Write-Host "    Quick Parts : $($result.QuickPartNames -join ', ')" -ForegroundColor DarkGreen
            }
            if ($result.Updated.Count -gt 0) {
                Write-Host "    Synced      : $($result.Updated -join ', ')" -ForegroundColor DarkGreen
            }
            if ($result.Missing.Count -gt 0) {
                Write-Host "    Missing     : $($result.Missing -join ', ')" -ForegroundColor Yellow
            }
        }
    }

    if ($errorCount -gt 0) {
        Write-Host ''
        Write-Status 'Files with errors:' 'ERROR'
        foreach ($result in $results | Where-Object { $_.Status -like 'Error*' }) {
            Write-Status "$($result.FileRef) - $($result.Status)" -Level ERROR
        }
        Write-Host ''
        Write-Status 'Completed with errors.' 'WARN'
        exit 1
    }

    Write-Host ''
    Write-Status 'Done.' 'SUCCESS'
}
catch {
    Write-Host ''
    Write-Status $_.Exception.Message 'ERROR'
    Write-Host ''
    exit 1
}
finally {
    if ($PSCmdlet.ParameterSetName -ne 'Local') {
        try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
    }
}
