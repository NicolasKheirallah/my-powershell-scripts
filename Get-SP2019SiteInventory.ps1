#Requires -Version 5.1
<#
.SYNOPSIS
    Exports all site collections and their subsites from a SharePoint 2019 on-premises farm to CSV.

.DESCRIPTION
    Enumerates every Web Application, Site Collection, and subsite (recursively) in the farm.
    Outputs a flat CSV with key metadata for each web object found.

    Must be run on a SharePoint server with the SharePoint snap-in available,
    by an account with Farm Administrator privileges.

.PARAMETER OutputPath
    Full path to the output CSV file.
    Defaults to "SP2019_SiteInventory_<timestamp>.csv" in the current directory.

.PARAMETER ExcludeAdminSites
    Switch. When specified, Central Administration site collections are excluded.

.PARAMETER WebApplicationUrl
    Optional. Limit the scan to a single Web Application URL.
    This should ideally be the exact URL returned by Get-SPWebApplication.
    If a site collection URL is supplied instead, the script will attempt to resolve
    it back to its owning Web Application.
    If omitted, all Web Applications in the farm are processed.

.PARAMETER Credential
    Optional. Runs the script in a new Windows PowerShell process using the supplied
    credential. Use this when the inventory must be executed as a different farm account.

.PARAMETER PromptForCredential
    Switch. Prompts for alternate credentials, then relaunches the script under that account.

.EXAMPLE
    .\Get-SP2019SiteInventory.ps1

.EXAMPLE
    .\Get-SP2019SiteInventory.ps1 -OutputPath "C:\Reports\sites.csv" -ExcludeAdminSites

.EXAMPLE
    .\Get-SP2019SiteInventory.ps1 -WebApplicationUrl "https://intranet.contoso.com"

.EXAMPLE
    .\Get-SP2019SiteInventory.ps1 -PromptForCredential

.EXAMPLE
    .\Get-SP2019SiteInventory.ps1 -WebApplicationUrl "https://intranet.contoso.com" -Credential (Get-Credential)
#>

[CmdletBinding()]
param (
    [string]$OutputPath = ".\SP2019_SiteInventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [switch]$ExcludeAdminSites,
    [string]$WebApplicationUrl,
    [System.Management.Automation.PSCredential]$Credential,
    [switch]$PromptForCredential,
    [switch]$RelaunchedWithCredential
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Add-SharePointSnapin {
    $sharePointSnapin = Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

    if (($sharePointSnapin | Measure-Object).Count -lt 1) {
        Add-PSSnapin Microsoft.SharePoint.PowerShell
    }
}

function ConvertTo-SingleQuotedPowerShellString {
    param (
        [AllowNull()]
        [string]$Value
    )

    if ($null -eq $Value) {
        return "''"
    }

    return "'" + ($Value -replace "'", "''") + "'"
}

function Start-ScriptAsCredential {
    param (
        [System.Management.Automation.PSCredential]$Credential,
        [string]$OutputPath,
        [switch]$ExcludeAdminSites,
        [string]$WebApplicationUrl
    )

    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        throw "Unable to relaunch the script because PSCommandPath is not available. Run the script from a .ps1 file."
    }

    $scriptPathLiteral = ConvertTo-SingleQuotedPowerShellString -Value $PSCommandPath
    $outputPathLiteral = ConvertTo-SingleQuotedPowerShellString -Value $OutputPath
    $commandParts = [System.Collections.Generic.List[string]]::new()

    $commandParts.Add("& $scriptPathLiteral")
    $commandParts.Add("-OutputPath $outputPathLiteral")

    if ($ExcludeAdminSites.IsPresent) {
        $commandParts.Add("-ExcludeAdminSites")
    }

    if (-not [string]::IsNullOrWhiteSpace($WebApplicationUrl)) {
        $webApplicationUrlLiteral = ConvertTo-SingleQuotedPowerShellString -Value $WebApplicationUrl
        $commandParts.Add("-WebApplicationUrl $webApplicationUrlLiteral")
    }

    $commandParts.Add("-RelaunchedWithCredential")

    $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes(($commandParts -join ' ')))
    $powershellArguments = "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCommand"

    Write-Host "Launching a new PowerShell session as $($Credential.UserName)..." -ForegroundColor Yellow

    $process = Start-Process -FilePath "powershell.exe" -Credential $Credential -ArgumentList $powershellArguments -Wait -PassThru
    exit $process.ExitCode
}

function Get-ObjectPropertyValue {
    param (
        [object]$Object,
        [string]$PropertyName
    )

    if ($null -eq $Object) {
        return $null
    }

    $property = $Object.PSObject.Properties[$PropertyName]
    if ($null -ne $property) {
        return $property.Value
    }

    return $null
}

if ($PromptForCredential -and $null -ne $Credential) {
    throw "Specify either -Credential or -PromptForCredential, not both."
}

$effectiveCredential = $Credential

if ($PromptForCredential) {
    $effectiveCredential = Get-Credential -Message "Enter the SharePoint account to run the inventory under"
}

if ($null -ne $effectiveCredential -and -not $RelaunchedWithCredential) {
    Start-ScriptAsCredential -Credential $effectiveCredential -OutputPath $OutputPath -ExcludeAdminSites:$ExcludeAdminSites -WebApplicationUrl $WebApplicationUrl
}

function Get-SafeValue {
    param (
        [scriptblock]$ScriptBlock,
        [string]$FieldName,
        [System.Collections.Generic.List[string]]$Warnings,
        [object]$DefaultValue = $null
    )

    try {
        return (& $ScriptBlock)
    }
    catch {
        if ($null -ne $Warnings) {
            $message = $_.Exception.Message -replace "`r`n|`r|`n", " "
            $Warnings.Add("${FieldName}: $message")
        }

        return $DefaultValue
    }
}

function Get-PrincipalLoginName {
    param (
        [object]$Principal
    )

    return Get-ObjectPropertyValue -Object $Principal -PropertyName "LoginName"
}

function Get-PrincipalEmail {
    param (
        [object]$Principal
    )

    return Get-ObjectPropertyValue -Object $Principal -PropertyName "Email"
}

function Get-GroupName {
    param (
        [object]$Group
    )

    $title = Get-ObjectPropertyValue -Object $Group -PropertyName "Title"
    if (-not [string]::IsNullOrWhiteSpace($title)) {
        return $title
    }

    return Get-ObjectPropertyValue -Object $Group -PropertyName "Name"
}

function ConvertTo-MB {
    param (
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    return [math]::Round(([double]$Value / 1MB), 2)
}

function Format-DateValue {
    param (
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    return ([datetime]$Value).ToString("yyyy-MM-dd HH:mm:ss")
}

function Initialize-OutputPath {
    param (
        [string]$Path
    )

    $resolvedPath = [System.IO.Path]::GetFullPath($Path)
    $directoryPath = Split-Path -Path $resolvedPath -Parent

    if (-not [string]::IsNullOrWhiteSpace($directoryPath) -and -not (Test-Path -LiteralPath $directoryPath)) {
        New-Item -Path $directoryPath -ItemType Directory -Force | Out-Null
    }

    return $resolvedPath
}

function Normalize-UrlString {
    param (
        [string]$Url
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $null
    }

    return $Url.Trim().TrimEnd('/').ToLowerInvariant()
}

function Get-TargetWebApplications {
    param (
        [string]$WebApplicationUrl,
        [bool]$IncludeCentralAdministration
    )

    $webApplications = @(Get-SPWebApplication -IncludeCentralAdministration:$IncludeCentralAdministration)

    if ([string]::IsNullOrWhiteSpace($WebApplicationUrl)) {
        return $webApplications
    }

    $normalizedRequestedUrl = Normalize-UrlString -Url $WebApplicationUrl
    $matchedWebApplications = @(
        $webApplications | Where-Object {
            (Normalize-UrlString -Url $_.Url) -eq $normalizedRequestedUrl
        }
    )

    if ($matchedWebApplications.Count -gt 0) {
        return $matchedWebApplications
    }

    $site = $null

    try {
        $site = Get-SPSite -Identity $WebApplicationUrl -ErrorAction Stop
        $resolvedWebApplication = Get-ObjectPropertyValue -Object $site -PropertyName "WebApplication"

        if ($null -ne $resolvedWebApplication) {
            return @($resolvedWebApplication)
        }
    }
    catch {
    }
    finally {
        if ($null -ne $site) {
            $site.Dispose()
        }
    }

    $availableUrls = @(
        $webApplications |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.Url) } |
            ForEach-Object { $_.Url } |
            Sort-Object -Unique
    )

    $availableUrlMessage = if ($availableUrls.Count -gt 0) {
        "Available web application URLs: $($availableUrls -join ', ')"
    }
    else {
        "No web applications were returned from Get-SPWebApplication."
    }

    throw "No SharePoint web application matched '$WebApplicationUrl'. $availableUrlMessage"
}

function Write-InventoryRow {
    param (
        [PSCustomObject]$Row,
        [string]$OutputPath
    )

    if (-not $script:CsvInitialized) {
        $Row | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force
        $script:CsvInitialized = $true
    }
    else {
        $Row | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Append
    }

    $script:RowCount++

    if ($Row.RowStatus -ne "OK") {
        $script:PartialRowCount++
    }
}

function New-FallbackWebRow {
    param (
        [Microsoft.SharePoint.SPWeb]$Web,
        [string]$WebApplicationUrl,
        [string]$SiteCollectionUrl,
        [int]$Depth,
        [string]$ParentWebUrl,
        [string]$Message
    )

    $webUrl = $null
    $webTitle = $null

    try {
        $webUrl = $Web.Url
    }
    catch {
    }

    try {
        $webTitle = $Web.Title
    }
    catch {
    }

    [PSCustomObject]@{
        WebApplicationUrl              = $WebApplicationUrl
        WebApplicationId               = $null
        WebApplicationName             = $null
        SiteCollectionUrl              = $SiteCollectionUrl
        SiteId                         = $null
        ContentDatabase                = $null
        ContentDatabaseServer          = $null
        OwnerLogin                     = $null
        OwnerEmail                     = $null
        SecondaryOwnerLogin            = $null
        SecondaryOwnerEmail            = $null
        LockState                      = $null
        ReadOnly                       = $null
        CompatibilityLevel             = $null
        QuotaTemplate                  = $null
        QuotaWarningMB                 = $null
        QuotaMaximumMB                 = $null
        WebId                          = $null
        WebUrl                         = $webUrl
        ServerRelativeUrl              = $null
        ParentWebUrl                   = $ParentWebUrl
        WebTitle                       = $webTitle
        WebTemplate                    = $null
        WebTemplateName                = $null
        Depth                          = $Depth
        IsRootWeb                      = ($Depth -eq 0)
        Language                       = $null
        Locale                         = $null
        Created                        = $null
        LastModified                   = $null
        HasUniquePermissions           = $null
        NoCrawl                        = $null
        RequestAccessEmail             = $null
        AssociatedOwnerGroup           = $null
        AssociatedMemberGroup          = $null
        AssociatedVisitorGroup         = $null
        Author                         = $null
        Description                    = $null
        ListCount                      = $null
        SiteCollectionStorageUsedMB    = $null
        SiteCollectionHits             = $null
        SiteCollectionVisits           = $null
        SiteCollectionBandwidthMB      = $null
        RowStatus                      = "Error"
        RowWarnings                    = $Message
    }
}

function Get-WebRow {
    param (
        [Microsoft.SharePoint.SPWeb]$Web,
        [string]$WebApplicationUrl,
        [string]$SiteCollectionUrl,
        [int]$Depth,
        [string]$ParentWebUrl
    )

    $warnings = [System.Collections.Generic.List[string]]::new()

    $site = Get-SafeValue -FieldName "Site" -Warnings $warnings -ScriptBlock { $Web.Site }
    $webApplication = Get-SafeValue -FieldName "WebApplication" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $site -PropertyName "WebApplication" }
    $usage = Get-SafeValue -FieldName "Usage" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $site -PropertyName "Usage" }
    $quota = Get-SafeValue -FieldName "Quota" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $site -PropertyName "Quota" }
    $contentDatabase = Get-SafeValue -FieldName "ContentDatabase" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $site -PropertyName "ContentDatabase" }
    $owner = Get-SafeValue -FieldName "Owner" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $site -PropertyName "Owner" }
    $secondaryOwner = Get-SafeValue -FieldName "SecondaryContact" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $site -PropertyName "SecondaryContact" }
    $author = Get-SafeValue -FieldName "Author" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $Web -PropertyName "Author" }
    $locale = Get-SafeValue -FieldName "Locale" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $Web -PropertyName "Locale" }
    $webApplicationName = Get-SafeValue -FieldName "WebApplicationName" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $webApplication -PropertyName "DisplayName" }
    $contentDatabaseServer = Get-SafeValue -FieldName "ContentDatabaseServer" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $contentDatabase -PropertyName "Server" }

    if ($contentDatabaseServer -isnot [string] -and $null -ne $contentDatabaseServer) {
        $contentDatabaseServer = Get-SafeValue -FieldName "ContentDatabaseServerName" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $contentDatabaseServer -PropertyName "Name" }
    }

    if ([string]::IsNullOrWhiteSpace($webApplicationName)) {
        $webApplicationName = Get-SafeValue -FieldName "WebApplicationNameFallback" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $webApplication -PropertyName "Name" }
    }

    $isRootWeb = Get-SafeValue -FieldName "IsRootWeb" -Warnings $warnings -ScriptBlock { $Web.IsRootWeb } -DefaultValue ($Depth -eq 0)
    $siteCollectionStorageUsedMB = $null
    $siteCollectionHits = $null
    $siteCollectionVisits = $null
    $siteCollectionBandwidthMB = $null

    if ($isRootWeb) {
        $siteCollectionStorageUsedMB = Get-SafeValue -FieldName "SiteCollectionStorageUsedMB" -Warnings $warnings -ScriptBlock { ConvertTo-MB -Value (Get-ObjectPropertyValue -Object $usage -PropertyName "Storage") }
        $siteCollectionHits = Get-SafeValue -FieldName "SiteCollectionHits" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $usage -PropertyName "Hits" }
        $siteCollectionVisits = Get-SafeValue -FieldName "SiteCollectionVisits" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $usage -PropertyName "Visits" }
        $siteCollectionBandwidthMB = Get-SafeValue -FieldName "SiteCollectionBandwidthMB" -Warnings $warnings -ScriptBlock { ConvertTo-MB -Value (Get-ObjectPropertyValue -Object $usage -PropertyName "Bandwidth") }
    }

    $rowStatus = if ($warnings.Count -gt 0) { "Partial" } else { "OK" }
    $rowWarnings = if ($warnings.Count -gt 0) { $warnings -join "; " } else { $null }

    [PSCustomObject]@{
        WebApplicationUrl              = $WebApplicationUrl
        WebApplicationId               = Get-SafeValue -FieldName "WebApplicationId" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $webApplication -PropertyName "Id" }
        WebApplicationName             = $webApplicationName
        SiteCollectionUrl              = $SiteCollectionUrl
        SiteId                         = Get-SafeValue -FieldName "SiteId" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $site -PropertyName "ID" }
        ContentDatabase                = Get-SafeValue -FieldName "ContentDatabaseName" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $contentDatabase -PropertyName "Name" }
        ContentDatabaseServer          = $contentDatabaseServer
        OwnerLogin                     = Get-SafeValue -FieldName "OwnerLogin" -Warnings $warnings -ScriptBlock { Get-PrincipalLoginName -Principal $owner }
        OwnerEmail                     = Get-SafeValue -FieldName "OwnerEmail" -Warnings $warnings -ScriptBlock { Get-PrincipalEmail -Principal $owner }
        SecondaryOwnerLogin            = Get-SafeValue -FieldName "SecondaryOwnerLogin" -Warnings $warnings -ScriptBlock { Get-PrincipalLoginName -Principal $secondaryOwner }
        SecondaryOwnerEmail            = Get-SafeValue -FieldName "SecondaryOwnerEmail" -Warnings $warnings -ScriptBlock { Get-PrincipalEmail -Principal $secondaryOwner }
        LockState                      = Get-SafeValue -FieldName "LockState" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $site -PropertyName "LockState" }
        ReadOnly                       = Get-SafeValue -FieldName "ReadOnly" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $site -PropertyName "ReadOnly" }
        CompatibilityLevel             = Get-SafeValue -FieldName "CompatibilityLevel" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $site -PropertyName "CompatibilityLevel" }
        QuotaTemplate                  = Get-SafeValue -FieldName "QuotaTemplate" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $site -PropertyName "QuotaTemplate" }
        QuotaWarningMB                 = Get-SafeValue -FieldName "QuotaWarningMB" -Warnings $warnings -ScriptBlock { ConvertTo-MB -Value (Get-ObjectPropertyValue -Object $quota -PropertyName "WarningStorageMaximumLevel") }
        QuotaMaximumMB                 = Get-SafeValue -FieldName "QuotaMaximumMB" -Warnings $warnings -ScriptBlock { ConvertTo-MB -Value (Get-ObjectPropertyValue -Object $quota -PropertyName "StorageMaximumLevel") }
        WebId                          = Get-SafeValue -FieldName "WebId" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $Web -PropertyName "ID" }
        WebUrl                         = Get-SafeValue -FieldName "WebUrl" -Warnings $warnings -ScriptBlock { $Web.Url }
        ServerRelativeUrl              = Get-SafeValue -FieldName "ServerRelativeUrl" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $Web -PropertyName "ServerRelativeUrl" }
        ParentWebUrl                   = $ParentWebUrl
        WebTitle                       = Get-SafeValue -FieldName "WebTitle" -Warnings $warnings -ScriptBlock { $Web.Title }
        WebTemplate                    = Get-SafeValue -FieldName "WebTemplate" -Warnings $warnings -ScriptBlock { "$($Web.WebTemplate)#$($Web.Configuration)" }
        WebTemplateName                = Get-SafeValue -FieldName "WebTemplateName" -Warnings $warnings -ScriptBlock { $Web.WebTemplate }
        Depth                          = $Depth
        IsRootWeb                      = $isRootWeb
        Language                       = Get-SafeValue -FieldName "Language" -Warnings $warnings -ScriptBlock { $Web.Language }
        Locale                         = Get-SafeValue -FieldName "LocaleName" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $locale -PropertyName "Name" }
        Created                        = Get-SafeValue -FieldName "Created" -Warnings $warnings -ScriptBlock { Format-DateValue -Value $Web.Created }
        LastModified                   = Get-SafeValue -FieldName "LastModified" -Warnings $warnings -ScriptBlock { Format-DateValue -Value $Web.LastItemModifiedDate }
        HasUniquePermissions           = Get-SafeValue -FieldName "HasUniquePermissions" -Warnings $warnings -ScriptBlock { $Web.HasUniqueRoleAssignments }
        NoCrawl                        = Get-SafeValue -FieldName "NoCrawl" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $Web -PropertyName "NoCrawl" }
        RequestAccessEmail             = Get-SafeValue -FieldName "RequestAccessEmail" -Warnings $warnings -ScriptBlock { Get-ObjectPropertyValue -Object $Web -PropertyName "RequestAccessEmail" }
        AssociatedOwnerGroup           = Get-SafeValue -FieldName "AssociatedOwnerGroup" -Warnings $warnings -ScriptBlock { Get-GroupName -Group (Get-ObjectPropertyValue -Object $Web -PropertyName "AssociatedOwnerGroup") }
        AssociatedMemberGroup          = Get-SafeValue -FieldName "AssociatedMemberGroup" -Warnings $warnings -ScriptBlock { Get-GroupName -Group (Get-ObjectPropertyValue -Object $Web -PropertyName "AssociatedMemberGroup") }
        AssociatedVisitorGroup         = Get-SafeValue -FieldName "AssociatedVisitorGroup" -Warnings $warnings -ScriptBlock { Get-GroupName -Group (Get-ObjectPropertyValue -Object $Web -PropertyName "AssociatedVisitorGroup") }
        Author                         = Get-SafeValue -FieldName "AuthorLogin" -Warnings $warnings -ScriptBlock { Get-PrincipalLoginName -Principal $author }
        Description                    = Get-SafeValue -FieldName "Description" -Warnings $warnings -ScriptBlock { $Web.Description -replace "`r`n|`r|`n", " " }
        ListCount                      = Get-SafeValue -FieldName "ListCount" -Warnings $warnings -ScriptBlock { $Web.Lists.Count }
        SiteCollectionStorageUsedMB    = $siteCollectionStorageUsedMB
        SiteCollectionHits             = $siteCollectionHits
        SiteCollectionVisits           = $siteCollectionVisits
        SiteCollectionBandwidthMB      = $siteCollectionBandwidthMB
        RowStatus                      = $rowStatus
        RowWarnings                    = $rowWarnings
    }
}

function Get-SubwebsRecursive {
    param (
        [Microsoft.SharePoint.SPWeb]$Web,
        [string]$WebApplicationUrl,
        [string]$SiteCollectionUrl,
        [int]$Depth,
        [string]$ParentWebUrl,
        [string]$OutputPath
    )

    try {
        $row = Get-WebRow -Web $Web -WebApplicationUrl $WebApplicationUrl -SiteCollectionUrl $SiteCollectionUrl -Depth $Depth -ParentWebUrl $ParentWebUrl
    }
    catch {
        $fallbackMessage = "RowBuild: $($_.Exception.Message -replace "`r`n|`r|`n", " ")"
        $row = New-FallbackWebRow -Web $Web -WebApplicationUrl $WebApplicationUrl -SiteCollectionUrl $SiteCollectionUrl -Depth $Depth -ParentWebUrl $ParentWebUrl -Message $fallbackMessage
    }

    Write-InventoryRow -Row $row -OutputPath $OutputPath

    foreach ($subweb in $Web.Webs) {
        try {
            Get-SubwebsRecursive -Web $subweb -WebApplicationUrl $WebApplicationUrl -SiteCollectionUrl $SiteCollectionUrl -Depth ($Depth + 1) -ParentWebUrl $row.WebUrl -OutputPath $OutputPath
        }
        catch {
            $script:SubwebFailureCount++
            Write-Warning "Failed to process subsite '$($subweb.Url)': $_"
        }
        finally {
            if ($null -ne $subweb) { $subweb.Dispose() }
        }
    }
}

Add-SharePointSnapin

if ($ExcludeAdminSites -and $WebApplicationUrl) {
    Write-Warning "-ExcludeAdminSites has no effect when -WebApplicationUrl is specified. The target web application will always be included."
}

$OutputPath = Initialize-OutputPath -Path $OutputPath
$script:CsvInitialized = $false
$script:RowCount = 0
$script:PartialRowCount = 0
$script:SubwebFailureCount = 0
$script:SiteCollectionFailureCount = 0

try {
    $webApplications = @(Get-TargetWebApplications -WebApplicationUrl $WebApplicationUrl -IncludeCentralAdministration (-not $ExcludeAdminSites.IsPresent))

    $waCount = $webApplications.Count
    $waIndex = 0

    foreach ($wa in $webApplications) {
        $waIndex++
        Write-Progress -Id 1 -Activity "Scanning Web Applications" -Status "$($wa.Url)" -PercentComplete (($waIndex / $waCount) * 100)

        $scCount = $wa.Sites.Count
        $scIndex = 0

        foreach ($site in $wa.Sites) {
            try {
                $scIndex++
                Write-Progress -Id 2 -ParentId 1 -Activity "Scanning Site Collections" -Status "$($site.Url)" -PercentComplete (($scIndex / $scCount) * 100)

                $rootWeb = $null
                $rootWeb = $site.RootWeb
                try {
                    Get-SubwebsRecursive -Web $rootWeb -WebApplicationUrl $wa.Url -SiteCollectionUrl $site.Url -Depth 0 -ParentWebUrl $null -OutputPath $OutputPath
                }
                finally {
                    if ($null -ne $rootWeb) { $rootWeb.Dispose() }
                }
            }
            catch {
                $script:SiteCollectionFailureCount++
                Write-Warning "Failed to process site collection '$($site.Url)': $_"
            }
            finally {
                $site.Dispose()
            }
        }
    }
}
catch {
    Write-Progress -Id 1 -Activity "Scanning Web Applications" -Completed
    Write-Progress -Id 2 -Activity "Scanning Site Collections" -Completed
    Write-Error "Script aborted: $_"
    exit 1
}

Write-Progress -Id 1 -Activity "Scanning Web Applications" -Completed
Write-Progress -Id 2 -Activity "Scanning Site Collections" -Completed

if ($script:RowCount -eq 0) {
    Write-Warning "No webs were collected. The CSV will not be written."
    exit 0
}

Write-Host ""
Write-Host "Export complete." -ForegroundColor Green
Write-Host "  Total webs             : $($script:RowCount)"
Write-Host "  Partial rows           : $($script:PartialRowCount)"
Write-Host "  Subsite failures       : $($script:SubwebFailureCount)"
Write-Host "  Site collection failures: $($script:SiteCollectionFailureCount)"
Write-Host "  Output file            : $((Resolve-Path $OutputPath).Path)"
