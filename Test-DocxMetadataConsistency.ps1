<#
.SYNOPSIS
    Inspects a Word .docx and reports mismatches between Quick Part fields,
    bound content controls, custom document properties, and SharePoint metadata.

.DESCRIPTION
    This is a read-only diagnostics script intended to explain cases where Word
    Online edit mode and view mode appear to show different values.

    The script opens the .docx package, reads the following sources, and groups
    them by property name:
      1. Classic DOCPROPERTY field display text.
      2. SharePoint-bound content control display text.
      3. The underlying bound XML store values.
      4. Custom document property values from docProps/custom.xml.
      5. In SharePoint mode, the current list item field value that the updater
         script would resolve for the same property name.

    By default the script prints a summary table. Use -PassThru to return raw
    objects for further filtering or export.

.PARAMETER LocalPath
    Path to a local .docx file.

.PARAMETER SiteUrl
    Full URL of the SharePoint site.

.PARAMETER FileServerRelativeUrl
    Server-relative URL of the SharePoint .docx file to inspect.

.PARAMETER ClientId
    Optional Entra app/client ID used with Connect-PnPOnline -Interactive.

.PARAMETER PropertyName
    Optional list of property names to limit the report to.

.PARAMETER OnlyMismatches
    Only output rows where the document contains conflicting values, duplicate
    custom properties, or the SharePoint value does not match any document-side
    value.

.PARAMETER PassThru
    Return raw objects instead of formatting a table.

.EXAMPLE
    .\Test-DocxMetadataConsistency.ps1 -LocalPath 'C:\Docs\Report.docx'

.EXAMPLE
    .\Test-DocxMetadataConsistency.ps1 `
        -SiteUrl 'https://contoso.sharepoint.com/sites/qms' `
        -ClientId 'your-client-id' `
        -FileServerRelativeUrl '/sites/qms/Shared Documents/Procedure.docx' `
        -OnlyMismatches

.EXAMPLE
    .\Test-DocxMetadataConsistency.ps1 `
        -SiteUrl 'https://contoso.sharepoint.com/sites/qms' `
        -ClientId 'your-client-id' `
        -FileServerRelativeUrl '/sites/qms/Shared Documents/Procedure.docx' `
        -PropertyName '_UIVersionString','DLCPolicyLabelValue' `
        -PassThru | ConvertTo-Json -Depth 6
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, ParameterSetName = 'LocalFile')]
    [string]$LocalPath,

    [Parameter(Mandatory, ParameterSetName = 'SharePointFile')]
    [string]$SiteUrl,

    [Parameter(Mandatory, ParameterSetName = 'SharePointFile')]
    [string]$FileServerRelativeUrl,

    [Parameter(ParameterSetName = 'SharePointFile')]
    [string]$ClientId,

    [string[]]$PropertyName,

    [switch]$OnlyMismatches,

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName 'System.IO.Compression'
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'

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
            if (Test-ObjectTypeName -InputObject $entry -TypeNames @('Microsoft.SharePoint.Client.FieldUserValue', 'Microsoft.SharePoint.Client.FieldLookupValue')) {
                $parts.Add([string]$entry.LookupValue)
                continue
            }
            $parts.Add([string]$entry)
        }
        return ($parts -join '; ')
    }

    return $Value
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
                    'begin' { $nestedDepth++ }
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

        if ($part -match '^\*\[local-name\(\)=') {
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
        if ($part -match "local-name\(\)='([^']+)'") {
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

function Get-WordSettingsUpdateFieldsFlag {
    param([System.IO.Compression.ZipArchive]$ZipArchive)

    $entry = $ZipArchive.GetEntry('word/settings.xml')
    if (-not $entry) {
        return $null
    }

    $xmlDoc = Get-XmlDocumentFromBytes -Bytes (Read-ZipEntryBytes -Entry $entry)
    $wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $nsManager = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
    $nsManager.AddNamespace('w', $wordNs)

    $updateFieldsNode = $xmlDoc.SelectSingleNode('/w:settings/w:updateFields', $nsManager)
    if (-not $updateFieldsNode) {
        return $false
    }

    $value = $updateFieldsNode.GetAttribute('val', $wordNs)
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $true
    }

    return ($value -eq 'true' -or $value -eq '1' -or $value -eq 'on')
}

function Get-CustomPropertyMap {
    param([System.IO.Compression.ZipArchive]$ZipArchive)

    $map = @{}
    $entry = $ZipArchive.GetEntry('docProps/custom.xml')
    if (-not $entry) {
        return $map
    }

    $xmlDoc = Get-XmlDocumentFromBytes -Bytes (Read-ZipEntryBytes -Entry $entry)
    $opNs = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
    $nsManager = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
    $nsManager.AddNamespace('op', $opNs)

    $propertyNodes = @($xmlDoc.SelectNodes('/op:Properties/op:property', $nsManager))
    foreach ($propertyNode in $propertyNodes) {
        $name = $propertyNode.GetAttribute('name')
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $valueNode = $null
        foreach ($child in $propertyNode.ChildNodes) {
            if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                $valueNode = $child
                break
            }
        }

        $value = if ($valueNode) { [string]$valueNode.InnerText } else { '' }
        if (-not $map.ContainsKey($name)) {
            $map[$name] = [System.Collections.Generic.List[string]]::new()
        }
        $map[$name].Add($value)
    }

    return $map
}

function Join-NodeInnerText {
    param([System.Xml.XmlNode[]]$Nodes)

    if (-not $Nodes -or $Nodes.Count -eq 0) {
        return ''
    }

    return (($Nodes | ForEach-Object { $_.InnerText }) -join '')
}

function Get-FieldOccurrencesFromXml {
    param(
        [System.Xml.XmlDocument]$XmlDoc,
        [string]$EntryName
    )

    $wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $nsManager = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
    $nsManager.AddNamespace('w', $wordNs)

    $records = [System.Collections.Generic.List[pscustomobject]]::new()

    $simpleFields = @($XmlDoc.SelectNodes('//w:fldSimple[contains(translate(@w:instr, ''abcdefghijklmnopqrstuvwxyz'', ''ABCDEFGHIJKLMNOPQRSTUVWXYZ''), ''DOCPROPERTY'')]', $nsManager))
    foreach ($field in $simpleFields) {
        $instruction = $field.GetAttribute('instr', $wordNs)
        $displayText = Join-NodeInnerText -Nodes @($field.SelectNodes('.//w:t', $nsManager))
        $isDirty = ($field.GetAttribute('dirty', $wordNs) -eq 'true')

        foreach ($name in (Get-QuickPartPropertyNamesFromInstruction -InstructionText $instruction)) {
            $records.Add([pscustomobject]@{
                    PropertyName = $name
                    SourceType   = 'FieldSimple'
                    Part         = $EntryName
                    DisplayText  = $displayText
                    StoreValue   = $null
                    StoreEntry   = $null
                    XPath        = $null
                    IsDirty      = $isDirty
                })
        }
    }

    $complexFieldBegins = @($XmlDoc.SelectNodes('//w:fldChar[@w:fldCharType=''begin'']', $nsManager))
    foreach ($fieldCharNode in $complexFieldBegins) {
        $instructionText = Get-ComplexFieldInstructionText -StartFieldChar $fieldCharNode -WordNamespace $wordNs
        $names = @(Get-QuickPartPropertyNamesFromInstruction -InstructionText $instructionText)
        if ($names.Count -eq 0) {
            continue
        }

        $displayText = Join-NodeInnerText -Nodes (Get-ComplexFieldResultTextNodes -StartFieldChar $fieldCharNode -WordNamespace $wordNs)
        $isDirty = ($fieldCharNode.GetAttribute('dirty', $wordNs) -eq 'true')

        foreach ($name in $names) {
            $records.Add([pscustomobject]@{
                    PropertyName = $name
                    SourceType   = 'FieldComplex'
                    Part         = $EntryName
                    DisplayText  = $displayText
                    StoreValue   = $null
                    StoreEntry   = $null
                    XPath        = $null
                    IsDirty      = $isDirty
                })
        }
    }

    return @($records)
}

function Get-ContentControlOccurrencesFromXml {
    param(
        [System.Xml.XmlDocument]$XmlDoc,
        [string]$EntryName,
        [hashtable]$StoreMap,
        [hashtable]$StoreDocuments
    )

    $wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $nsManager = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
    $nsManager.AddNamespace('w', $wordNs)

    $records = [System.Collections.Generic.List[pscustomobject]]::new()
    $sdtNodes = @($XmlDoc.SelectNodes('//w:sdt[w:sdtPr/w:dataBinding]', $nsManager))
    foreach ($sdtNode in $sdtNodes) {
        $bindingNode = $sdtNode.SelectSingleNode('w:sdtPr/w:dataBinding', $nsManager)
        if (-not $bindingNode) {
            continue
        }

        $sourceKey = Get-BindingSourceKey -SdtNode $sdtNode -NamespaceManager $nsManager -WordNamespace $wordNs
        if ([string]::IsNullOrWhiteSpace($sourceKey)) {
            continue
        }

        $storeItemId = $bindingNode.GetAttribute('storeItemID', $wordNs)
        $xpath = $bindingNode.GetAttribute('xpath', $wordNs)
        $prefixMappings = $bindingNode.GetAttribute('prefixMappings', $wordNs)
        $storeEntryName = if ($StoreMap.ContainsKey($storeItemId)) { [string]$StoreMap[$storeItemId] } else { $null }

        $storeValue = $null
        if (-not [string]::IsNullOrWhiteSpace($storeEntryName) -and $StoreDocuments.ContainsKey($storeEntryName) -and -not [string]::IsNullOrWhiteSpace($xpath)) {
            $storeDoc = $StoreDocuments[$storeEntryName]
            $storeNsManager = New-Object System.Xml.XmlNamespaceManager($storeDoc.NameTable)
            Add-NamespaceMappingsFromString -NamespaceManager $storeNsManager -PrefixMappings $prefixMappings

            $targetNode = $storeDoc.SelectSingleNode($xpath, $storeNsManager)
            if (-not $targetNode) {
                $fallbackXPath = ConvertTo-NamespaceAgnosticXPath -XPath $xpath
                if ($fallbackXPath -ne $xpath) {
                    $targetNode = $storeDoc.SelectSingleNode($fallbackXPath, $storeNsManager)
                }
            }

            if ($targetNode) {
                $storeValue = [string]$targetNode.InnerText
            }
        }

        $displayText = Join-NodeInnerText -Nodes @($sdtNode.SelectNodes('w:sdtContent//w:t', $nsManager))
        $records.Add([pscustomobject]@{
                PropertyName = $sourceKey
                SourceType   = 'ContentControl'
                Part         = $EntryName
                DisplayText  = $displayText
                StoreValue   = $storeValue
                StoreEntry   = $storeEntryName
                XPath        = $xpath
                IsDirty      = $null
            })
    }

    return @($records)
}

function Normalize-DiagnosticValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ($null -eq $text) {
        return $null
    }

    $text = $text.Trim()
    if ($text -eq '') {
        return $null
    }

    return $text
}

function Get-UniqueNormalizedValues {
    param([object[]]$Values)

    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($value in $Values) {
        $normalized = Normalize-DiagnosticValue -Value $value
        if ($null -ne $normalized) {
            $set.Add($normalized) | Out-Null
        }
    }

    return @($set)
}

function Join-DisplayValues {
    param([object[]]$Values)

    $uniqueValues = Get-UniqueNormalizedValues -Values $Values
    if ($uniqueValues.Count -eq 0) {
        return ''
    }

    return ($uniqueValues -join ' | ')
}

function Get-DocxConsistencyAnalysis {
    param([byte[]]$DocxBytes)

    $memoryStream = New-Object System.IO.MemoryStream(, $DocxBytes)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($memoryStream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
        try {
            $wordXmlEntries = @($zip.Entries | Where-Object { $_.FullName -match '^word/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$' })
            $storeMap = Get-BoundStoreMap -ZipArchive $zip
            $storeDocuments = @{}
            foreach ($entryName in @($storeMap.Values | Select-Object -Unique)) {
                $entry = $zip.GetEntry($entryName)
                if ($entry) {
                    $storeDocuments[$entryName] = Get-XmlDocumentFromBytes -Bytes (Read-ZipEntryBytes -Entry $entry)
                }
            }

            $fieldOccurrences = [System.Collections.Generic.List[pscustomobject]]::new()
            $contentControlOccurrences = [System.Collections.Generic.List[pscustomobject]]::new()

            foreach ($entry in $wordXmlEntries) {
                $xmlDoc = Get-XmlDocumentFromBytes -Bytes (Read-ZipEntryBytes -Entry $entry)
                foreach ($record in (Get-FieldOccurrencesFromXml -XmlDoc $xmlDoc -EntryName $entry.FullName)) {
                    $fieldOccurrences.Add($record)
                }
                foreach ($record in (Get-ContentControlOccurrencesFromXml -XmlDoc $xmlDoc -EntryName $entry.FullName -StoreMap $storeMap -StoreDocuments $storeDocuments)) {
                    $contentControlOccurrences.Add($record)
                }
            }

            return [pscustomobject]@{
                UpdateFieldsOnOpen = Get-WordSettingsUpdateFieldsFlag -ZipArchive $zip
                CustomProperties   = Get-CustomPropertyMap -ZipArchive $zip
                FieldOccurrences   = @($fieldOccurrences)
                ContentControls    = @($contentControlOccurrences)
            }
        }
        finally {
            $zip.Dispose()
        }
    }
    finally {
        $memoryStream.Dispose()
    }
}

function Resolve-SharePointFileContext {
    param(
        [string]$SiteUrl,
        [string]$FileServerRelativeUrl,
        [string]$ClientId
    )

    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        throw 'PnP.PowerShell is required for SharePoint mode. Install-Module PnP.PowerShell -Scope CurrentUser'
    }

    Import-Module PnP.PowerShell -ErrorAction Stop | Out-Null

    $connectParameters = @{
        Url         = $SiteUrl
        Interactive = $true
        ErrorAction = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
        $connectParameters.ClientId = $ClientId
    }

    Connect-PnPOnline @connectParameters

    $decodedRef = [System.Uri]::UnescapeDataString($FileServerRelativeUrl)
    $listItem = Get-PnPFile -Url $decodedRef -AsListItem -ErrorAction Stop
    $memoryStream = Get-PnPFile -Url $decodedRef -AsMemoryStream -ErrorAction Stop
    try {
        $docxBytes = $memoryStream.ToArray()
    }
    finally {
        $memoryStream.Dispose()
    }

    $web = Get-PnPWeb -Includes ServerRelativeUrl -ErrorAction Stop
    $webUrl = ([string]$web.ServerRelativeUrl).TrimEnd('/')
    $relativePath = $decodedRef
    if ($webUrl -ne '' -and $relativePath.StartsWith($webUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
        $relativePath = $relativePath.Substring($webUrl.Length)
    }
    $relativePath = $relativePath.TrimStart('/')
    $libraryIdentity = [System.Uri]::UnescapeDataString($relativePath.Split('/')[0])
    if ([string]::IsNullOrWhiteSpace($libraryIdentity)) {
        throw "Could not infer library name from path '$decodedRef'."
    }

    $fields = @(Get-PnPField -List $libraryIdentity -ErrorAction Stop)

    return [pscustomobject]@{
        TargetLabel = "$SiteUrl $decodedRef"
        DocxBytes   = $docxBytes
        ListItem    = $listItem
        Fields      = $fields
    }
}

function Resolve-LocalFileContext {
    param([string]$LocalPath)

    if (-not (Test-Path -LiteralPath $LocalPath -PathType Leaf)) {
        throw "File not found: $LocalPath"
    }

    $resolvedPath = (Resolve-Path -LiteralPath $LocalPath).Path
    return [pscustomobject]@{
        TargetLabel = $resolvedPath
        DocxBytes   = [System.IO.File]::ReadAllBytes($resolvedPath)
        ListItem    = $null
        Fields      = @()
    }
}

function Get-ConsistencySummaryRows {
    param(
        [pscustomobject]$Analysis,
        [AllowNull()][object]$ListItem,
        [object[]]$Fields,
        [string[]]$PropertyNameFilter
    )

    $discoveredNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $Analysis.CustomProperties.Keys) {
        $discoveredNames.Add([string]$name) | Out-Null
    }
    foreach ($entry in $Analysis.FieldOccurrences) {
        $discoveredNames.Add([string]$entry.PropertyName) | Out-Null
    }
    foreach ($entry in $Analysis.ContentControls) {
        $discoveredNames.Add([string]$entry.PropertyName) | Out-Null
    }

    if ($PropertyNameFilter -and $PropertyNameFilter.Count -gt 0) {
        $filterSet = New-NameLookup -Names $PropertyNameFilter
        $names = @($discoveredNames | Where-Object { $filterSet.Contains($_) } | Sort-Object)
    }
    else {
        $names = @($discoveredNames | Sort-Object)
    }

    $resolver = @{}
    if ($null -ne $ListItem) {
        $aliasMap = Get-PropertyAliasMap
        $allowedNames = New-NameLookup -Names (Expand-AllowedPropertyNames -Names $names -AliasMap $aliasMap)
        $resolver = New-PropertyResolver -ListItem $ListItem -Fields $Fields -AllowedNames $allowedNames -AliasMap $aliasMap
    }

    $rows = [System.Collections.Generic.List[pscustomobject]]::new()
    foreach ($name in $names) {
        $fieldEntries = @($Analysis.FieldOccurrences | Where-Object { $_.PropertyName -eq $name })
        $contentEntries = @($Analysis.ContentControls | Where-Object { $_.PropertyName -eq $name })
        $customValues = if ($Analysis.CustomProperties.ContainsKey($name)) { @($Analysis.CustomProperties[$name]) } else { @() }

        $fieldValues = @($fieldEntries | ForEach-Object { $_.DisplayText })
        $contentValues = @($contentEntries | ForEach-Object { $_.DisplayText })
        $boundValues = @($contentEntries | ForEach-Object { $_.StoreValue })
        $documentValues = @($fieldValues + $contentValues + $boundValues + $customValues)
        $distinctDocumentValues = Get-UniqueNormalizedValues -Values $documentValues

        $sharePointValue = $null
        $sharePointMatches = $null
        if ($resolver.ContainsKey($name)) {
            $sharePointValue = Normalize-DiagnosticValue -Value (ConvertTo-PropertyValue -Value $resolver[$name])
            if ($null -ne $sharePointValue) {
                $sharePointMatches = ($distinctDocumentValues -contains $sharePointValue)
            }
        }

        $notes = [System.Collections.Generic.List[string]]::new()
        $duplicateCustomProperty = ($customValues.Count -gt 1)
        $documentMismatch = ($distinctDocumentValues.Count -gt 1)

        if ($duplicateCustomProperty) {
            $notes.Add('Duplicate custom properties')
        }
        if ($documentMismatch) {
            $notes.Add('Multiple document-side values')
        }
        if ($null -ne $sharePointMatches -and -not $sharePointMatches) {
            $notes.Add('SharePoint value differs from all document-side values')
        }

        $rows.Add([pscustomobject]@{
                PropertyName             = $name
                SharePointValue          = if ($null -ne $sharePointValue) { $sharePointValue } else { '' }
                FieldDisplays            = Join-DisplayValues -Values $fieldValues
                ContentControlDisplays   = Join-DisplayValues -Values $contentValues
                BoundStoreValues         = Join-DisplayValues -Values $boundValues
                CustomPropertyValues     = Join-DisplayValues -Values $customValues
                DocumentValueMismatch    = $documentMismatch
                SharePointMatchesAny     = $sharePointMatches
                DuplicateCustomProperty  = $duplicateCustomProperty
                FieldOccurrences         = $fieldEntries.Count
                ContentControlOccurrences = $contentEntries.Count
                Notes                    = ($notes -join '; ')
            })
    }

    return @($rows)
}

if ($PSCmdlet.ParameterSetName -eq 'SharePointFile') {
    Write-Status "Connecting to SharePoint and downloading '$FileServerRelativeUrl'" 'INFO'
    $context = Resolve-SharePointFileContext -SiteUrl $SiteUrl -FileServerRelativeUrl $FileServerRelativeUrl -ClientId $ClientId
}
else {
    Write-Status "Loading local file '$LocalPath'" 'INFO'
    $context = Resolve-LocalFileContext -LocalPath $LocalPath
}

if ($context.DocxBytes.Length -lt 4 -or $context.DocxBytes[0] -ne 0x50 -or $context.DocxBytes[1] -ne 0x4B) {
    throw 'Target file is not a valid .docx/ZIP package.'
}

$analysis = Get-DocxConsistencyAnalysis -DocxBytes $context.DocxBytes
$rows = @(Get-ConsistencySummaryRows -Analysis $analysis -ListItem $context.ListItem -Fields $context.Fields -PropertyNameFilter $PropertyName)

if ($OnlyMismatches) {
    $rows = @($rows | Where-Object {
            $_.DocumentValueMismatch -or
            $_.DuplicateCustomProperty -or
            ($null -ne $_.SharePointMatchesAny -and -not $_.SharePointMatchesAny)
        })
}

$mismatchCount = @($rows | Where-Object {
        $_.DocumentValueMismatch -or
        $_.DuplicateCustomProperty -or
        ($null -ne $_.SharePointMatchesAny -and -not $_.SharePointMatchesAny)
    }).Count

Write-Status "Target: $($context.TargetLabel)" 'INFO'
Write-Status "Word updateFields on open: $($analysis.UpdateFieldsOnOpen)" 'INFO'
Write-Status "Properties reported: $($rows.Count); suspicious rows: $mismatchCount" 'INFO'

if ($PassThru) {
    $rows | Sort-Object PropertyName
}
elseif ($rows.Count -eq 0) {
    Write-Status 'No rows matched the current filter.' 'WARN'
}
else {
    $rows |
        Sort-Object PropertyName |
        Format-Table PropertyName, SharePointValue, FieldDisplays, ContentControlDisplays, BoundStoreValues, CustomPropertyValues, DocumentValueMismatch, SharePointMatchesAny, DuplicateCustomProperty, Notes -Wrap -AutoSize
}