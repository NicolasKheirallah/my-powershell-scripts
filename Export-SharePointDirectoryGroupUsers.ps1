#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '3.1.0' }

<#
.SYNOPSIS
    Exports users from Entra ID / AD groups that were added to a SharePoint Online site.

.DESCRIPTION
    This script uses the latest stable PnP PowerShell module line (3.1.0 or newer)
    to inspect a SharePoint Online site and export ALL users who have access,
    including:

      1. Individual users granted direct site permissions
      2. Individual users who are members of SharePoint groups
      3. Users expanded from Entra ID / AD directory groups (DLs, security groups)
         granted through SharePoint groups, direct site permissions, unique
         list/library, folder, or item/file permissions
      4. Members of the connected Microsoft 365 group (group-connected sites)

    Supported directory group types through Get-PnPEntraIDGroup /
    Get-PnPEntraIDGroupMember include:
      - Distribution groups that are resolvable through Microsoft Graph
      - Mail-enabled security groups
      - Security groups
      - Microsoft 365 groups

    Limitations:
      - Dynamic distribution groups are not supported through Microsoft Graph
        and cannot be expanded by this script.

    Notes:
      - Connect-PnPOnline now requires your own Entra ID app registration when
        using interactive auth. Pass -ClientId or configure ENTRAID_APP_ID.
      - Your app/user needs enough SharePoint permissions to read site groups and
        enough Microsoft Graph permissions to read group membership, typically
        GroupMember.Read.All and User.Read.All or broader equivalents.
      - -AllSites requires access to the SharePoint tenant admin site and uses
        Get-PnPTenantSite to enumerate site collections.
      - If -GrantExpandedUsersDirectly is used, the script also writes SharePoint
        permissions or SharePoint group membership. Use -WhatIf first.

.PARAMETER SiteUrl
    Full URL of the SharePoint Online site.
    Optional when -SitesCsvPath or -AllSites is used.

.PARAMETER SitesCsvPath
    Optional path to a CSV file containing site URLs to process.
    The CSV must contain a column named "New Site URL" unless you override
    it with -SitesCsvColumnName.

.PARAMETER SitesCsvColumnName
    Column name in -SitesCsvPath that contains the SharePoint site URLs.
    Default: New Site URL

.PARAMETER AllSites
    Enumerates all SharePoint Online site collections in the tenant and
    processes each site collection. Combine with -IncludeSubsites to also
    process every subsite under each site collection.

.PARAMETER TenantAdminUrl
    Optional SharePoint Online tenant admin URL used with -AllSites.
    Example: https://contoso-admin.sharepoint.com
    If omitted, the script will try to derive it from -SiteUrl or the first URL
    found in -SitesCsvPath.

.PARAMETER ClientId
    Entra ID app registration client ID used with Connect-PnPOnline -Interactive.
    Optional if ENTRAID_APP_ID or a managed app id is already configured.

.PARAMETER OutputCsvPath
    Local path for the generated CSV file.
    Default: .\SharePoint-GroupUsers-<timestamp>.csv

.PARAMETER IncludeTransitiveMembers
    When $true, nested group membership is expanded transitively.
    Default: $true

.PARAMETER UploadFolderServerRelativeUrl
    Optional SharePoint folder where the generated CSV should also be uploaded.
    Example: /sites/HR/Shared Documents/Reports

.PARAMETER IncludeSubsites
    When set, enumerates all subsites (sub-webs) under the given SiteUrl
    recursively and processes each one.

.PARAMETER SiteLevelOnly
    When set, only inspects site-level (web) and SharePoint group permissions.
    Skips scanning lists, libraries, folders, and items for unique permissions.
    Much faster for large sites when you only need the top-level access picture.

.PARAMETER PageSize
    Page size used when enumerating list items looking for unique permissions.
    Default: 500

.PARAMETER GrantExpandedUsersDirectly
    When set, resolved users from directory groups are granted direct
    SharePoint access:
      - If the directory group is inside a SharePoint group, the users are added
        directly to that SharePoint group.
      - If the directory group has direct permissions on the web, list, folder,
        file, or list item, the same permission levels are granted directly to
        the resolved users.

.PARAMETER RemoveExpandedDirectoryGroupsAfterGrant
    Optional follow-up to -GrantExpandedUsersDirectly. After at least one user
    is granted direct access for a source directory group assignment, the script
    removes the original directory group access from SharePoint.
    Use with caution and prefer -WhatIf first.

.EXAMPLE
    .\Export-SharePointDirectoryGroupUsers.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/Finance" `
        -ClientId "00000000-0000-0000-0000-000000000000"

.EXAMPLE
    .\Export-SharePointDirectoryGroupUsers.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/Finance" `
        -ClientId "00000000-0000-0000-0000-000000000000" `
        -OutputCsvPath "C:\Temp\finance-site-users.csv" `
        -UploadFolderServerRelativeUrl "/sites/Finance/Shared Documents/Reports"

.EXAMPLE
    .\Export-SharePointDirectoryGroupUsers.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
        -ClientId "00000000-0000-0000-0000-000000000000" `
        -IncludeSubsites

.EXAMPLE
    .\Export-SharePointDirectoryGroupUsers.ps1 `
        -SitesCsvPath "C:\Temp\sites.csv" `
        -ClientId "00000000-0000-0000-0000-000000000000"

    The CSV contains a column named "New Site URL".

.EXAMPLE
    .\Export-SharePointDirectoryGroupUsers.ps1 `
        -AllSites `
        -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
        -ClientId "00000000-0000-0000-0000-000000000000" `
        -IncludeSubsites

.EXAMPLE
    .\Export-SharePointDirectoryGroupUsers.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/Finance" `
        -ClientId "00000000-0000-0000-0000-000000000000" `
        -GrantExpandedUsersDirectly `
        -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$SiteUrl,

    [Parameter()]
    [string]$SitesCsvPath,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$SitesCsvColumnName = 'New Site URL',

    [Parameter()]
    [switch]$AllSites,

    [Parameter()]
    [string]$TenantAdminUrl,

    [Parameter()]
    [string]$ClientId,

    [Parameter()]
    [string]$OutputCsvPath = (Join-Path -Path (Get-Location) -ChildPath ("SharePoint-GroupUsers-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))),

    [Parameter()]
    [bool]$IncludeTransitiveMembers = $true,

    [Parameter()]
    [string]$UploadFolderServerRelativeUrl,

    [Parameter()]
    [switch]$IncludeSubsites,

    [Parameter()]
    [switch]$SiteLevelOnly,

    [Parameter()]
    [ValidateRange(1, 5000)]
    [int]$PageSize = 500,

    [Parameter()]
    [switch]$GrantExpandedUsersDirectly,

    [Parameter()]
    [switch]$RemoveExpandedDirectoryGroupsAfterGrant
)

$ErrorActionPreference = 'Stop'

$results = [System.Collections.Generic.List[object]]::new()
$warnings = [System.Collections.Generic.List[string]]::new()
$entraGroupIdentityCache = @{}
$entraGroupDisplayNameCache = @{}
$sharePointGroupMemberCache = @{}
$sharePointGroupCache = @{}
$allEntraGroups = $null
$userDetailCache = @{}
$exportRowKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$userReadPermissionWarningShown = $false
$userReadSparseWarningShown = $false
$userReadCmdletWarningShown = $false
$userReadGraphSuccessShown = $false
$activeDirectoryInitialized = $false
$activeDirectoryAvailable = $false
$activeDirectoryImportAttempted = $false
$activeDirectoryFallbackWarningShown = $false
$adGroupCache = @{}
$directGrantActionKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$sourceReplacementActionKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$directGrantSuccessCount = 0
$sourceGroupRemovalCount = 0
$scriptRootCmdlet = $PSCmdlet

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO ] $Message" -ForegroundColor Cyan
}

function Write-WarnMessage {
    param([string]$Message)
    $warnings.Add($Message) | Out-Null
    Write-Warning $Message
}

function Invoke-ShouldProcess {
    param(
        [Parameter(Mandatory)][string]$Target,
        [Parameter(Mandatory)][string]$Action
    )

    if ($null -eq $script:scriptRootCmdlet) {
        return $true
    }

    return $script:scriptRootCmdlet.ShouldProcess($Target, $Action)
}

function Test-IsBenignGrantError {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return $false
    }

    return $Message -match 'already exists|already a member|already has|duplicate|same key|has already been added|role assignment exists'
}

function Test-IsBenignRemovalError {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return $false
    }

    return $Message -match 'does not exist|not found|is not a member|cannot find|no role assignment'
}

function Get-RoleNamesFromPermissionLevels {
    param([string]$PermissionLevels)

    if ([string]::IsNullOrWhiteSpace($PermissionLevels)) {
        return @()
    }

    return @(
        $PermissionLevels -split ';' |
            ForEach-Object { $_.Trim() } |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_) -and
                $_ -ne 'Limited Access'
            } |
            Select-Object -Unique
    )
}

function Get-GrantableUserLoginName {
    param([Parameter(Mandatory)]$MemberDetails)

    if ($MemberDetails.MemberObjectType -ne 'User') {
        return $null
    }

    if (-not [string]::IsNullOrWhiteSpace($MemberDetails.UserPrincipalName)) {
        return $MemberDetails.UserPrincipalName
    }

    if (-not [string]::IsNullOrWhiteSpace($MemberDetails.Mail)) {
        return $MemberDetails.Mail
    }

    return $null
}

function Get-VerificationIdentityTokens {
    param(
        [Parameter(Mandatory)][string]$UserLoginName,
        [Parameter(Mandatory)]$MemberDetails
    )

    $tokens = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($value in @(
        $UserLoginName
        (Get-UserIdentityFromLoginName -LoginName $UserLoginName)
        $MemberDetails.UserPrincipalName
        $MemberDetails.Mail
    )) {
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $tokens.Add($value.Trim().ToLowerInvariant()) | Out-Null
        }
    }

    return $tokens
}

function Test-AssignmentMatchesResolvedUser {
    param(
        [Parameter(Mandatory)]$Assignment,
        [Parameter(Mandatory)][System.Collections.Generic.HashSet[string]]$IdentityTokens
    )

    foreach ($candidate in @(
        $Assignment.PrincipalEmail
        $Assignment.PrincipalLoginName
        (Get-UserIdentityFromLoginName -LoginName $Assignment.PrincipalLoginName)
    )) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and $IdentityTokens.Contains($candidate.Trim().ToLowerInvariant())) {
            return $true
        }
    }

    return $false
}

function Get-VerificationAssignmentsForSourcePrincipal {
    param([Parameter(Mandatory)]$SourcePrincipal)

    switch ($SourcePrincipal.AccessPath) {
        'DirectWebPermission' {
            $web = Get-PnPWeb -Includes RoleAssignments
            return @(Get-RoleAssignmentsForSecurableObject -ClientObject $web -AccessPath $SourcePrincipal.AccessPath -AccessContainer $SourcePrincipal.AccessContainer)
        }
        'ListPermission' {
            $list = Get-PnPList -Identity $SourcePrincipal.ListId -Includes RoleAssignments, Title
            return @(Get-RoleAssignmentsForSecurableObject -ClientObject $list -AccessPath $SourcePrincipal.AccessPath -AccessContainer $SourcePrincipal.AccessContainer -List $list)
        }
        'FolderPermission' {
            $list = Get-PnPList -Identity $SourcePrincipal.ListId -Includes Title
            $folderItem = Get-PnPFolder -Url $SourcePrincipal.AccessContainer -AsListItem
            return @(Get-RoleAssignmentsForSecurableObject -ClientObject $folderItem -AccessPath $SourcePrincipal.AccessPath -AccessContainer $SourcePrincipal.AccessContainer -List $list -ListItem $folderItem)
        }
        'FilePermission' {
            $list = Get-PnPList -Identity $SourcePrincipal.ListId -Includes Title
            $listItem = Get-PnPListItem -List $SourcePrincipal.ListId -Id $SourcePrincipal.ItemId
            return @(Get-RoleAssignmentsForSecurableObject -ClientObject $listItem -AccessPath $SourcePrincipal.AccessPath -AccessContainer $SourcePrincipal.AccessContainer -List $list -ListItem $listItem)
        }
        'ListItemPermission' {
            $list = Get-PnPList -Identity $SourcePrincipal.ListId -Includes Title
            $listItem = Get-PnPListItem -List $SourcePrincipal.ListId -Id $SourcePrincipal.ItemId
            return @(Get-RoleAssignmentsForSecurableObject -ClientObject $listItem -AccessPath $SourcePrincipal.AccessPath -AccessContainer $SourcePrincipal.AccessContainer -List $list -ListItem $listItem)
        }
        default {
            return @()
        }
    }
}

function Verify-ResolvedUserDirectAccess {
    param(
        [Parameter(Mandatory)]$SourcePrincipal,
        [Parameter(Mandatory)]$MemberDetails,
        [Parameter(Mandatory)][string]$UserLoginName
    )

    $identityTokens = Get-VerificationIdentityTokens -UserLoginName $UserLoginName -MemberDetails $MemberDetails

    if (Test-HasFlagValue -Value $SourcePrincipal.GrantedToType -FlagName 'SharePointGroup') {
        $groupIdentity = if ($null -ne $SourcePrincipal.GrantedToId -and "$($SourcePrincipal.GrantedToId)" -ne '') {
            $SourcePrincipal.GrantedToId
        }
        else {
            $SourcePrincipal.GrantedToTitle
        }

        $groupMembers = @(Get-PnPGroupMember -Group $groupIdentity)
        foreach ($groupMember in $groupMembers) {
            $memberAssignment = [PSCustomObject]@{
                PrincipalLoginName = $groupMember.LoginName
                PrincipalEmail     = (Get-PrincipalEmail -Principal $groupMember)
            }
            if (Test-AssignmentMatchesResolvedUser -Assignment $memberAssignment -IdentityTokens $identityTokens) {
                return $true
            }
        }

        return $false
    }

    $expectedRoles = @(Get-RoleNamesFromPermissionLevels -PermissionLevels $SourcePrincipal.PermissionLevels)
    if ($expectedRoles.Count -eq 0) {
        return $false
    }

    $assignments = @(Get-VerificationAssignmentsForSourcePrincipal -SourcePrincipal $SourcePrincipal)
    foreach ($assignment in $assignments) {
        if (-not (Test-AssignmentMatchesResolvedUser -Assignment $assignment -IdentityTokens $identityTokens)) {
            continue
        }

        $actualRoles = @(Get-RoleNamesFromPermissionLevels -PermissionLevels $assignment.PermissionLevels)
        $missingRoles = @($expectedRoles | Where-Object { $_ -notin $actualRoles })
        if ($missingRoles.Count -eq 0) {
            return $true
        }
    }

    return $false
}

function Grant-ResolvedUserDirectAccess {
    param(
        [Parameter(Mandatory)]$SourcePrincipal,
        [Parameter(Mandatory)]$MemberDetails
    )

    $userLoginName = Get-GrantableUserLoginName -MemberDetails $MemberDetails
    if ([string]::IsNullOrWhiteSpace($userLoginName)) {
        Write-WarnMessage "Skipping direct grant for '$($MemberDetails.DisplayName)' because no usable SharePoint login identity was resolved."
        return $false
    }

    $grantKey = @(
        $SourcePrincipal.AccessPath
        $SourcePrincipal.AccessContainer
        $SourcePrincipal.PermissionLevels
        $SourcePrincipal.GrantedToLoginName
        $SourcePrincipal.PrincipalLoginName
        $userLoginName
    ) -join '|'
    if (-not $script:directGrantActionKeys.Add($grantKey)) {
        return $true
    }

    $targetLabel = if (-not [string]::IsNullOrWhiteSpace($SourcePrincipal.GrantedToTitle)) {
        $SourcePrincipal.GrantedToTitle
    }
    else {
        $SourcePrincipal.AccessContainer
    }

    if (Test-HasFlagValue -Value $SourcePrincipal.GrantedToType -FlagName 'SharePointGroup') {
        $groupIdentity = if ($null -ne $SourcePrincipal.GrantedToId -and "$($SourcePrincipal.GrantedToId)" -ne '') {
            $SourcePrincipal.GrantedToId
        }
        else {
            $SourcePrincipal.GrantedToTitle
        }

        if (-not (Invoke-ShouldProcess -Target "SharePoint group '$targetLabel'" -Action "Add member '$userLoginName'")) {
            return $false
        }

        try {
            Add-PnPGroupMember -Group $groupIdentity -LoginName $userLoginName
        }
        catch {
            if (Test-IsBenignGrantError -Message $_.Exception.Message) {
                Write-Verbose "User '$userLoginName' is already a direct member of SharePoint group '$targetLabel'"
            }
            else {
                throw
            }
        }

        if (Verify-ResolvedUserDirectAccess -SourcePrincipal $SourcePrincipal -MemberDetails $MemberDetails -UserLoginName $userLoginName) {
            $script:directGrantSuccessCount++
            Write-Info "  Verified '$userLoginName' is directly present in SharePoint group '$targetLabel'"
            return $true
        }

        Write-WarnMessage "Direct add verification failed for '$userLoginName' in SharePoint group '$targetLabel'."
        return $false
    }

    $roles = @(Get-RoleNamesFromPermissionLevels -PermissionLevels $SourcePrincipal.PermissionLevels)
    if ($roles.Count -eq 0) {
        Write-WarnMessage "No assignable permission levels were found for '$($SourcePrincipal.AccessPath)' on '$($SourcePrincipal.AccessContainer)'."
        return $false
    }

    $grantedAnyRole = $false
    foreach ($role in $roles) {
        $action = "Grant role '$role' to '$userLoginName' on '$($SourcePrincipal.AccessContainer)'"
        if (-not (Invoke-ShouldProcess -Target $SourcePrincipal.AccessContainer -Action $action)) {
            continue
        }

        try {
            switch ($SourcePrincipal.AccessPath) {
                'DirectWebPermission' {
                    Set-PnPWebPermission -User $userLoginName -AddRole $role
                    break
                }
                'ListPermission' {
                    Set-PnPListPermission -Identity $SourcePrincipal.ListId -User $userLoginName -AddRole $role
                    break
                }
                'FolderPermission' {
                    Set-PnPFolderPermission -List $SourcePrincipal.ListId -Identity $SourcePrincipal.AccessContainer -User $userLoginName -AddRole $role -SystemUpdate
                    break
                }
                'FilePermission' {
                    Set-PnPListItemPermission -List $SourcePrincipal.ListId -Identity $SourcePrincipal.ItemId -User $userLoginName -AddRole $role -SystemUpdate
                    break
                }
                'ListItemPermission' {
                    Set-PnPListItemPermission -List $SourcePrincipal.ListId -Identity $SourcePrincipal.ItemId -User $userLoginName -AddRole $role -SystemUpdate
                    break
                }
                default {
                    Write-WarnMessage "Direct grant is not supported for access path '$($SourcePrincipal.AccessPath)'."
                    return $false
                }
            }
            $grantedAnyRole = $true
        }
        catch {
            if (Test-IsBenignGrantError -Message $_.Exception.Message) {
                Write-Verbose "Role '$role' already granted directly to '$userLoginName' on '$($SourcePrincipal.AccessContainer)'"
                $grantedAnyRole = $true
                continue
            }
            throw
        }
    }

    if ($grantedAnyRole) {
        if (Verify-ResolvedUserDirectAccess -SourcePrincipal $SourcePrincipal -MemberDetails $MemberDetails -UserLoginName $userLoginName) {
            $script:directGrantSuccessCount++
            Write-Info "  Verified direct access for '$userLoginName' on '$($SourcePrincipal.AccessContainer)'"
            return $true
        }

        Write-WarnMessage "Direct permission verification failed for '$userLoginName' on '$($SourcePrincipal.AccessContainer)'."
    }

    return $false
}

function Remove-SourceDirectoryGroupAccess {
    param([Parameter(Mandatory)]$SourcePrincipal)

    $removalKey = @(
        $SourcePrincipal.AccessPath
        $SourcePrincipal.AccessContainer
        $SourcePrincipal.PermissionLevels
        $SourcePrincipal.GrantedToLoginName
        $SourcePrincipal.PrincipalLoginName
    ) -join '|'
    if (-not $script:sourceReplacementActionKeys.Add($removalKey)) {
        return
    }

    if (Test-HasFlagValue -Value $SourcePrincipal.GrantedToType -FlagName 'SharePointGroup') {
        $groupIdentity = if ($null -ne $SourcePrincipal.GrantedToId -and "$($SourcePrincipal.GrantedToId)" -ne '') {
            $SourcePrincipal.GrantedToId
        }
        else {
            $SourcePrincipal.GrantedToTitle
        }

        if (-not (Invoke-ShouldProcess -Target "SharePoint group '$($SourcePrincipal.GrantedToTitle)'" -Action "Remove directory group '$($SourcePrincipal.PrincipalLoginName)'")) {
            return
        }

        try {
            Remove-PnPGroupMember -Group $groupIdentity -LoginName $SourcePrincipal.PrincipalLoginName
            $script:sourceGroupRemovalCount++
            Write-Info "  Removed directory group '$($SourcePrincipal.PrincipalTitle)' from SharePoint group '$($SourcePrincipal.GrantedToTitle)'"
            return
        }
        catch {
            if (Test-IsBenignRemovalError -Message $_.Exception.Message) {
                return
            }
            throw
        }
    }

    $roles = @(Get-RoleNamesFromPermissionLevels -PermissionLevels $SourcePrincipal.PermissionLevels)
    foreach ($role in $roles) {
        if (-not (Invoke-ShouldProcess -Target $SourcePrincipal.AccessContainer -Action "Remove role '$role' from '$($SourcePrincipal.PrincipalLoginName)'")) {
            continue
        }

        try {
            switch ($SourcePrincipal.AccessPath) {
                'DirectWebPermission' {
                    Set-PnPWebPermission -User $SourcePrincipal.PrincipalLoginName -RemoveRole $role
                    break
                }
                'ListPermission' {
                    Set-PnPListPermission -Identity $SourcePrincipal.ListId -User $SourcePrincipal.PrincipalLoginName -RemoveRole $role
                    break
                }
                'FolderPermission' {
                    Set-PnPFolderPermission -List $SourcePrincipal.ListId -Identity $SourcePrincipal.AccessContainer -User $SourcePrincipal.PrincipalLoginName -RemoveRole $role -SystemUpdate
                    break
                }
                'FilePermission' {
                    Set-PnPListItemPermission -List $SourcePrincipal.ListId -Identity $SourcePrincipal.ItemId -User $SourcePrincipal.PrincipalLoginName -RemoveRole $role -SystemUpdate
                    break
                }
                'ListItemPermission' {
                    Set-PnPListItemPermission -List $SourcePrincipal.ListId -Identity $SourcePrincipal.ItemId -User $SourcePrincipal.PrincipalLoginName -RemoveRole $role -SystemUpdate
                    break
                }
                default {
                    Write-WarnMessage "Removal of source directory-group access is not supported for access path '$($SourcePrincipal.AccessPath)'."
                    return
                }
            }
        }
        catch {
            if (Test-IsBenignRemovalError -Message $_.Exception.Message) {
                continue
            }
            throw
        }
    }

    $script:sourceGroupRemovalCount++
    Write-Info "  Removed direct directory-group access for '$($SourcePrincipal.PrincipalTitle)' from '$($SourcePrincipal.AccessContainer)'"
}

function Grant-ExpandedUsersDirectAccessFromMembers {
    param(
        [Parameter(Mandatory)]$SourcePrincipal,
        [Parameter(Mandatory)][object[]]$Members
    )

    if (-not $GrantExpandedUsersDirectly) {
        return
    }

    $successfulGrantCount = 0
    foreach ($member in $Members) {
        $details = Resolve-MemberUserDetails -Member $member
        if ($details.MemberObjectType -ne 'User') {
            continue
        }

        $userLoginName = Get-GrantableUserLoginName -MemberDetails $details
        if ([string]::IsNullOrWhiteSpace($userLoginName)) {
            Write-WarnMessage "Skipping direct grant for '$($details.DisplayName)' because no user principal name or mail value was resolved."
            continue
        }

        try {
            if (Grant-ResolvedUserDirectAccess -SourcePrincipal $SourcePrincipal -MemberDetails $details) {
                $successfulGrantCount++
            }
        }
        catch {
            Write-WarnMessage "Failed to grant '$userLoginName' direct access for '$($SourcePrincipal.AccessContainer)': $($_.Exception.Message)"
        }
    }

    if ($RemoveExpandedDirectoryGroupsAfterGrant -and $successfulGrantCount -gt 0) {
        try {
            Remove-SourceDirectoryGroupAccess -SourcePrincipal $SourcePrincipal
        }
        catch {
            Write-WarnMessage "Failed to remove source directory-group access for '$($SourcePrincipal.PrincipalTitle)': $($_.Exception.Message)"
        }
    }
}

function Reset-SiteScopedCaches {
    $script:sharePointGroupMemberCache = @{}
    $script:sharePointGroupCache = @{}
}

function New-ConnectParameters {
    param([Parameter(Mandatory)][string]$Url)

    $parameters = @{
        Url         = $Url
        Interactive = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
        $parameters.ClientId = $ClientId
    }

    return $parameters
}

function Add-UniqueSiteUrl {
    param(
        [ValidateNotNull()][System.Collections.Generic.List[string]]$SiteUrls,
        [ValidateNotNull()][System.Collections.Generic.HashSet[string]]$Seen,
        [string]$Url
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return
    }

    $normalizedUrl = $Url.Trim().TrimEnd('/')
    if ($Seen.Add($normalizedUrl)) {
        $SiteUrls.Add($normalizedUrl) | Out-Null
    }
}

function Get-SharePointAdminUrlFromSiteUrl {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $null
    }

    try {
        $uri = [Uri]$Url
    }
    catch {
        return $null
    }

    if ($uri.Host -match '^(?<tenant>[^.]+)-admin\.sharepoint\.(?<suffix>.+)$') {
        return "{0}://{1}" -f $uri.Scheme, $uri.Host
    }

    if ($uri.Host -match '^(?<tenant>[^.]+)\.sharepoint\.(?<suffix>.+)$') {
        return "{0}://{1}-admin.sharepoint.{2}" -f $uri.Scheme, $Matches['tenant'], $Matches['suffix']
    }

    return $null
}

function Get-ResolvedTenantAdminUrl {
    if (-not [string]::IsNullOrWhiteSpace($TenantAdminUrl)) {
        return $TenantAdminUrl.Trim().TrimEnd('/')
    }

    $candidateUrls = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($SiteUrl)) {
        $candidateUrls.Add($SiteUrl) | Out-Null
    }

    if (-not [string]::IsNullOrWhiteSpace($SitesCsvPath) -and (Test-Path -LiteralPath $SitesCsvPath)) {
        try {
            $firstCsvRow = Import-Csv -LiteralPath $SitesCsvPath | Select-Object -First 1
            if ($firstCsvRow -and $firstCsvRow.PSObject.Properties[$SitesCsvColumnName]) {
                $csvUrl = "$($firstCsvRow.$SitesCsvColumnName)".Trim()
                if (-not [string]::IsNullOrWhiteSpace($csvUrl)) {
                    $candidateUrls.Add($csvUrl) | Out-Null
                }
            }
        }
        catch {
            Write-Verbose "Could not inspect '$SitesCsvPath' while deriving the tenant admin URL: $($_.Exception.Message)"
        }
    }

    foreach ($candidateUrl in $candidateUrls) {
        $adminUrl = Get-SharePointAdminUrlFromSiteUrl -Url $candidateUrl
        if (-not [string]::IsNullOrWhiteSpace($adminUrl)) {
            return $adminUrl
        }
    }

    return $null
}

function Get-AllTenantSiteUrls {
    param([Parameter(Mandatory)][string]$AdminUrl)

    Write-Info "Connecting to tenant admin site: $AdminUrl"
    $adminConnectParameters = New-ConnectParameters -Url $AdminUrl
    $null = Connect-PnPOnline @adminConnectParameters

    try {
        Write-Info "Enumerating all site collections with Get-PnPTenantSite"
        $tenantSites = @(Get-PnPTenantSite -Detailed)
        $urls = [System.Collections.Generic.List[string]]::new()
        $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($tenantSite in $tenantSites) {
            $template = "$($tenantSite.Template)"
            $siteUrl = "$($tenantSite.Url)"
            if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                continue
            }

            if ($template -like 'SPSPERS*') {
                continue
            }

            Add-UniqueSiteUrl -SiteUrls $urls -Seen $seen -Url $siteUrl
        }

        Write-Info "Discovered $($urls.Count) site collections from the tenant admin center"
        return @($urls)
    }
    finally {
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            Write-Verbose "Disconnect-PnPOnline after tenant enumeration: $($_.Exception.Message)"
        }
    }
}

function Get-TargetSiteUrls {
    $siteUrls = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    if ($AllSites) {
        $resolvedTenantAdminUrl = Get-ResolvedTenantAdminUrl
        if ([string]::IsNullOrWhiteSpace($resolvedTenantAdminUrl)) {
            throw "TenantAdminUrl is required for -AllSites unless it can be derived from -SiteUrl or the first URL in -SitesCsvPath."
        }

        foreach ($tenantSiteUrl in @(Get-AllTenantSiteUrls -AdminUrl $resolvedTenantAdminUrl)) {
            Add-UniqueSiteUrl -SiteUrls $siteUrls -Seen $seen -Url $tenantSiteUrl
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($SiteUrl)) {
        Add-UniqueSiteUrl -SiteUrls $siteUrls -Seen $seen -Url $SiteUrl
    }

    if (-not [string]::IsNullOrWhiteSpace($SitesCsvPath)) {
        if (-not (Test-Path -LiteralPath $SitesCsvPath)) {
            throw "SitesCsvPath does not exist: $SitesCsvPath"
        }

        $csvRows = @(Import-Csv -LiteralPath $SitesCsvPath)
        if ($csvRows.Count -eq 0) {
            throw "SitesCsvPath is empty: $SitesCsvPath"
        }

        $firstRow = $csvRows[0]
        if ($null -eq $firstRow.PSObject.Properties[$SitesCsvColumnName]) {
            throw "CSV '$SitesCsvPath' does not contain the required column '$SitesCsvColumnName'."
        }

        foreach ($row in $csvRows) {
            $siteUrlFromCsv = "$($row.$SitesCsvColumnName)"
            Add-UniqueSiteUrl -SiteUrls $siteUrls -Seen $seen -Url $siteUrlFromCsv
        }
    }

    if ($siteUrls.Count -eq 0) {
        throw "Provide -SiteUrl, -SitesCsvPath, or -AllSites with at least one SharePoint site URL."
    }

    return @($siteUrls)
}

function Invoke-SiteCollectionProcessing {
    param([Parameter(Mandatory)][string]$RootSiteUrl)

    Reset-SiteScopedCaches

    $connectParameters = New-ConnectParameters -Url $RootSiteUrl

    Write-Info "Connecting to $RootSiteUrl"
    $null = Connect-PnPOnline @connectParameters

    $webUrlsToProcess = [System.Collections.Generic.List[string]]::new()
    $rootWeb = Get-PnPWeb
    $webUrlsToProcess.Add($rootWeb.Url) | Out-Null

    if ($IncludeSubsites) {
        Write-Info "Enumerating subsites recursively"
        try {
            $subWebs = @(Get-PnPSubWeb -Recurse)
            foreach ($subWeb in $subWebs) {
                $subUrl = Get-PnPProperty -ClientObject $subWeb -Property Url
                $webUrlsToProcess.Add($subUrl) | Out-Null
            }
            Write-Info "Found $($subWebs.Count) subsites ($($webUrlsToProcess.Count) webs total)"
        }
        catch {
            Write-WarnMessage "Failed to enumerate subsites for '$RootSiteUrl': $($_.Exception.Message)"
        }
    }

    Process-SingleWeb -WebUrl $rootWeb.Url

    Write-Info "Checking for connected Microsoft 365 group"
    $connectedM365Group = Get-ConnectedM365GroupMembers
    if ($connectedM365Group) {
        try {
            $m365Members = Get-ExpandedGroupUsers -EntraGroup $connectedM365Group
        }
        catch {
            Write-WarnMessage "Failed to expand connected M365 group members for '$RootSiteUrl': $($_.Exception.Message)"
            $m365Members = @()
        }

        if ($m365Members -and $m365Members.Count -gt 0) {
            $m365SourcePrincipal = [PSCustomObject]@{
                AccessPath          = 'ConnectedM365Group'
                AccessContainer     = $rootWeb.Url
                PermissionLevels    = 'M365 Group Member'
                GrantedToTitle      = $connectedM365Group.DisplayName
                GrantedToLoginName  = $connectedM365Group.Mail
                GrantedToType       = 'Microsoft365Group'
                PrincipalTitle      = $connectedM365Group.DisplayName
                PrincipalLoginName  = $connectedM365Group.Mail
                PrincipalType       = 'Microsoft365Group'
            }
            Add-ExportRows -WebUrl $rootWeb.Url -SourcePrincipal $m365SourcePrincipal -EntraGroup $connectedM365Group -Members $m365Members
            Write-Info "Added $($m365Members.Count) members from connected M365 group '$($connectedM365Group.DisplayName)'"
        }
    }

    for ($i = 1; $i -lt $webUrlsToProcess.Count; $i++) {
        $subWebUrl = $webUrlsToProcess[$i]
        Write-Info "Reconnecting to subsite: $subWebUrl"
        $subConnectParams = New-ConnectParameters -Url $subWebUrl
        try {
            $null = Connect-PnPOnline @subConnectParams
            Process-SingleWeb -WebUrl $subWebUrl
        }
        catch {
            Write-WarnMessage "Failed to process subsite '$subWebUrl': $($_.Exception.Message)"
        }
    }

    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    catch {
        Write-Verbose "Disconnect-PnPOnline: $($_.Exception.Message)"
    }

    return $webUrlsToProcess.Count
}

function Test-IsLookupNotFoundError {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return $false
    }

    return $Message -match 'Request_ResourceNotFound|does not exist|not found|cannot find|no group found|resource .* was not found'
}

function Test-HasFlagValue {
    param(
        [Parameter(Mandatory)]$Value,
        [Parameter(Mandatory)][string]$FlagName
    )

    try {
        $flag = [Microsoft.SharePoint.Client.Utilities.PrincipalType]::$FlagName
        return (([int]$Value -band [int]$flag) -ne 0)
    }
    catch {
        return ("$Value" -match [regex]::Escape($FlagName))
    }
}

function Test-IsDirectoryGroupPrincipal {
    param([Parameter(Mandatory)]$Principal)

    # Check login name claim format first — SharePoint often reports Entra ID
    # groups with PrincipalType 'User' when they use c:0t.c|tenant| claims
    $loginName = $null
    $loginProp = $Principal.PSObject.Properties['PrincipalLoginName']
    if ($loginProp) { $loginName = $loginProp.Value }
    if (-not $loginName) {
        $loginProp = $Principal.PSObject.Properties['LoginName']
        if ($loginProp) { $loginName = $loginProp.Value }
    }
    if (-not [string]::IsNullOrWhiteSpace($loginName) -and (Test-IsEntraGroupClaim -LoginName $loginName)) {
        return $true
    }

    if ($null -eq $Principal.PSObject.Properties['PrincipalType']) {
        return $false
    }

    return (Test-HasFlagValue -Value $Principal.PrincipalType -FlagName 'SecurityGroup') -or
           (Test-HasFlagValue -Value $Principal.PrincipalType -FlagName 'DistributionList')
}

function Test-IsSharePointGroupPrincipal {
    param([Parameter(Mandatory)]$Principal)

    if ($null -eq $Principal.PSObject.Properties['PrincipalType']) {
        return $false
    }

    return Test-HasFlagValue -Value $Principal.PrincipalType -FlagName 'SharePointGroup'
}

function Test-IsUserPrincipal {
    param([Parameter(Mandatory)]$Principal)

    if ($null -eq $Principal.PSObject.Properties['PrincipalType']) {
        return $false
    }

    return Test-HasFlagValue -Value $Principal.PrincipalType -FlagName 'User'
}

function Test-IsSpecialClaimPrincipal {
    param([Parameter(Mandatory)][string]$LoginName)

    return $LoginName -match '^c:0-.f\|rolemanager\|' -or
           $LoginName -match '^SharingLinks\.' -or
           $LoginName -match '^app@sharepoint$'
}

function Test-IsEntraGroupClaim {
    param([Parameter(Mandatory)][string]$LoginName)

    # c:0t.c|tenant|<guid>  — Entra ID security/M365 group via object ID
    # c:0o.c|federateddirectoryclaimprovider|<guid>  — security group via object ID
    # c:0o.c|federateddirectoryclaimprovider|email@domain.com — mail-enabled security group or DL via email
    return $LoginName -match '^c:0t\.c\|tenant\|[0-9a-f]{8}-' -or
           $LoginName -match '^c:0o\.c\|federateddirectoryclaimprovider\|.+'
}

function Test-IsProbablyGroupPrincipal {
    param([Parameter(Mandatory)]$Principal)

    $loginName = $null
    $loginProp = $Principal.PSObject.Properties['PrincipalLoginName']
    if ($loginProp) { $loginName = $loginProp.Value }
    if (-not $loginName) {
        $loginProp = $Principal.PSObject.Properties['LoginName']
        if ($loginProp) { $loginName = $loginProp.Value }
    }

    # Known Entra group claim patterns
    if (-not [string]::IsNullOrWhiteSpace($loginName) -and (Test-IsEntraGroupClaim -LoginName $loginName)) {
        return $true
    }

    # CSOM PrincipalType explicitly says group
    $principalType = $null
    $ptProp = $Principal.PSObject.Properties['PrincipalType']
    if ($ptProp) { $principalType = $ptProp.Value }

    if ($null -ne $principalType) {
        if ((Test-HasFlagValue -Value $principalType -FlagName 'SharePointGroup') -or
            (Test-HasFlagValue -Value $principalType -FlagName 'SecurityGroup') -or
            (Test-HasFlagValue -Value $principalType -FlagName 'DistributionList')) {
            return $true
        }
    }

    # Catch-all: any c:0 claim with a GUID or email-like value after the last pipe
    # is likely a group reference (covers edge cases not matched above)
    if (-not [string]::IsNullOrWhiteSpace($loginName) -and $loginName -match '^c:0') {
        # Has an object ID → likely a group
        if (Get-DirectoryObjectIdFromLoginName -LoginName $loginName) {
            return $true
        }
        # Has c:0o or c:0t prefix with any value → likely a group claim
        if ($loginName -match '^c:0[to]\.') {
            return $true
        }
    }

    return $false
}

function Get-UserIdentityFromLoginName {
    param([string]$LoginName)

    if ([string]::IsNullOrWhiteSpace($LoginName)) {
        return $LoginName
    }

    $email = Get-EmailAddressFromText -Text $LoginName
    if (-not [string]::IsNullOrWhiteSpace($email)) {
        return $email
    }

    if ($LoginName -match '\|([^|]+)$') {
        return $Matches[1]
    }

    return $LoginName
}

function Get-DirectoryObjectIdFromLoginName {
    param([string]$LoginName)

    if ([string]::IsNullOrWhiteSpace($LoginName)) {
        return $null
    }

    $match = [regex]::Match($LoginName, '(?i)[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}')
    if ($match.Success) {
        return $match.Value
    }

    return $null
}

function Get-EmailAddressFromText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $match = [regex]::Match($Text, '(?i)([a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,})')
    if ($match.Success) {
        return $match.Groups[1].Value
    }

    return $null
}

function Get-PropertyValue {
    param(
        [Parameter(Mandatory)][object]$InputObject,
        [Parameter(Mandatory)][string[]]$PropertyNames
    )

    # If the input is a dictionary/hashtable (e.g. from Invoke-PnPGraphMethod), read keys directly
    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($propertyName in $PropertyNames) {
            if ($InputObject.Contains($propertyName)) {
                $value = $InputObject[$propertyName]
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace("$value")) {
                    return $value
                }
            }
        }
        return $null
    }

    foreach ($propertyName in $PropertyNames) {
        $property = $InputObject.PSObject.Properties[$propertyName]
        if ($property -and $null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace("$($property.Value)")) {
            return $property.Value
        }
    }

    # Check AdditionalData (Graph SDK v5+) and AdditionalProperties (Graph SDK v4)
    foreach ($dictName in @('AdditionalData', 'AdditionalProperties')) {
        $additional = $InputObject.PSObject.Properties[$dictName]
        if ($additional -and $additional.Value -and $additional.Value -is [System.Collections.IDictionary]) {
            foreach ($propertyName in $PropertyNames) {
                if ($additional.Value.Contains($propertyName)) {
                    $value = $additional.Value[$propertyName]
                    if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace("$value")) {
                        return $value
                    }
                }
            }
        }
    }

    return $null
}

function Resolve-EntraUserByIdentity {
    param([string]$Identity)

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        return $null
    }

    if ($script:userDetailCache.ContainsKey($Identity)) {
        return $script:userDetailCache[$Identity]
    }

    try {
        $user = Get-PnPEntraIDUser -Identity $Identity -Select 'Id','DisplayName','UserPrincipalName','Mail','UserType'
        # Validate the returned object actually has useful properties — PnP may return sparse SDK objects
        $testDisplayName = Get-PropertyValue -InputObject $user -PropertyNames @('DisplayName', 'displayName')
        $testUpn = Get-PropertyValue -InputObject $user -PropertyNames @('UserPrincipalName', 'userPrincipalName')
        if (-not [string]::IsNullOrWhiteSpace($testDisplayName) -or -not [string]::IsNullOrWhiteSpace($testUpn)) {
            $script:userDetailCache[$Identity] = $user
            return $user
        }
        if (-not $script:userReadSparseWarningShown) {
            $script:userReadSparseWarningShown = $true
            $objType = if ($user) { $user.GetType().FullName } else { '(null)' }
            $allProps = if ($user) { @($user.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ', ' } else { '(none)' }
            Write-WarnMessage "Get-PnPEntraIDUser returned sparse object for '$Identity' (type=$objType). Properties: $allProps — falling through to Graph REST API."
        }
        else {
            Write-Verbose "  Get-PnPEntraIDUser returned sparse object for '$Identity', falling through to Graph REST API"
        }
    }
    catch {
        if (-not $script:userReadCmdletWarningShown) {
            $script:userReadCmdletWarningShown = $true
            Write-WarnMessage "Get-PnPEntraIDUser failed for '$Identity': $($_.Exception.Message)"
        }
        else {
            Write-Verbose "  Get-PnPEntraIDUser failed for '$Identity' : $($_.Exception.Message)"
        }
        $message = $_.Exception.Message
        if (Test-IsLookupNotFoundError -Message $message) {
            $script:userDetailCache[$Identity] = $null
            return $null
        }
    }

    try {
        $encodedIdentity = [uri]::EscapeDataString($Identity)
        $graphResult = Invoke-PnPGraphMethod -Url "v1.0/users/$encodedIdentity?`$select=id,displayName,userPrincipalName,mail,userType" -Method Get
        if ($graphResult) {
            $grDisplayName = Get-PropertyValue -InputObject $graphResult -PropertyNames @('displayName', 'DisplayName')
            $grUpn = Get-PropertyValue -InputObject $graphResult -PropertyNames @('userPrincipalName', 'UserPrincipalName')
            if (-not $script:userReadGraphSuccessShown) {
                $script:userReadGraphSuccessShown = $true
                Write-Info "  Graph REST API resolved user '$Identity' → DisplayName='$grDisplayName', UPN='$grUpn'"
            }
        }
        $script:userDetailCache[$Identity] = $graphResult
        return $graphResult
    }
    catch {
        if (-not $script:userReadPermissionWarningShown) {
            $script:userReadPermissionWarningShown = $true
            Write-WarnMessage "Graph REST API failed for user '$Identity': $($_.Exception.Message)"
            if ($_.Exception.Message -match 'Authorization_RequestDenied|Insufficient privileges|Forbidden|access is denied|does not have permission') {
                Write-WarnMessage "The PnP app can read group membership but cannot read user profile details. Add Microsoft Graph User.Read.All or Directory.Read.All and grant admin consent, then reconnect and rerun the script."
            }
        }
        else {
            Write-Verbose "  Graph REST API failed for '$Identity' : $($_.Exception.Message)"
        }
        if (Test-IsLookupNotFoundError -Message $_.Exception.Message) {
            $script:userDetailCache[$Identity] = $null
        }
    }

    return $null
}

function Initialize-ActiveDirectoryFallback {
    if ($script:activeDirectoryInitialized) {
        return $script:activeDirectoryAvailable
    }

    $script:activeDirectoryInitialized = $true

    try {
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            return $false
        }

        if (-not $script:activeDirectoryImportAttempted) {
            $script:activeDirectoryImportAttempted = $true
            Import-Module ActiveDirectory -ErrorAction Stop | Out-Null
        }

        $script:activeDirectoryAvailable = $true
        Write-Info "ActiveDirectory module detected. Local AD fallback is available."
        return $true
    }
    catch {
        Write-WarnMessage "ActiveDirectory module was found but could not be loaded. Local AD fallback will not be used. $($_.Exception.Message)"
        return $false
    }
}

function Escape-AdFilterValue {
    param([string]$Value)

    if ($null -eq $Value) {
        return ''
    }

    return ($Value -replace "'", "''")
}

function Get-AdSafeSamAccountName {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $candidate = $Value.Trim()
    if ($candidate -match '\|') {
        $candidate = ($candidate -split '\|')[-1]
    }

    if ($candidate -match '@') {
        return $null
    }

    if ($candidate -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') {
        return $null
    }

    if ($candidate -notmatch '^[A-Za-z0-9._-]{1,256}$') {
        return $null
    }

    return $candidate
}

function Resolve-AdGroupByCandidates {
    param([object[]]$Candidates)

    if (-not (Initialize-ActiveDirectoryFallback)) {
        return $null
    }

    $candidateTypePriority = @{
        Mail           = 1
        SamAccountName = 2
        DisplayName    = 3
        Name           = 4
    }

    $normalizedCandidates = @(
        foreach ($candidate in $Candidates) {
            if ($null -eq $candidate) {
                continue
            }

            if ($candidate -is [string]) {
                if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                    [PSCustomObject]@{
                        Type  = if ($candidate -match '@') { 'Mail' } else { 'SamAccountName' }
                        Value = $candidate.Trim()
                    }
                }
                continue
            }

            $candidateType = Get-PropertyValue -InputObject $candidate -PropertyNames @('Type', 'LookupType')
            $candidateValue = Get-PropertyValue -InputObject $candidate -PropertyNames @('Value', 'LookupValue')
            if (-not [string]::IsNullOrWhiteSpace("$candidateType") -and -not [string]::IsNullOrWhiteSpace("$candidateValue")) {
                [PSCustomObject]@{
                    Type  = "$candidateType"
                    Value = "$candidateValue".Trim()
                }
            }
        }
    ) | Group-Object Type, Value | ForEach-Object { $_.Group[0] } |
        Sort-Object @{ Expression = { $candidateTypePriority[$_.Type] } }, @{ Expression = { $_.Value } }

    foreach ($candidate in $normalizedCandidates) {
        $cacheKey = "{0}:{1}" -f $candidate.Type, $candidate.Value
        if ($script:adGroupCache.ContainsKey($cacheKey)) {
            $cached = $script:adGroupCache[$cacheKey]
            if ($cached) {
                return $cached
            }
            continue
        }

        $escapedCandidate = Escape-AdFilterValue -Value $candidate.Value
        $filters = @()
        switch ($candidate.Type) {
            'Mail' {
                $filters += "mail -eq '$escapedCandidate'"
                $filters += "proxyAddresses -eq 'SMTP:$escapedCandidate'"
                $filters += "proxyAddresses -eq 'smtp:$escapedCandidate'"
                break
            }
            'SamAccountName' {
                $filters += "samAccountName -eq '$escapedCandidate'"
                break
            }
            'DisplayName' {
                $filters += "displayName -eq '$escapedCandidate'"
                break
            }
            'Name' {
                $filters += "name -eq '$escapedCandidate'"
                break
            }
            default {
                Write-Verbose "Unsupported AD group lookup candidate type '$($candidate.Type)' for value '$($candidate.Value)'"
                $script:adGroupCache[$cacheKey] = $null
                continue
            }
        }

        foreach ($filter in $filters) {
            try {
                $matches = @(Get-ADGroup -Filter $filter -Properties mail,displayName,groupScope -ErrorAction Stop)
                if ($matches.Count -eq 1) {
                    $script:adGroupCache[$cacheKey] = $matches[0]
                    if ($candidate.Type -in @('DisplayName', 'Name')) {
                        Write-Info "  Matched on-prem AD group '$($matches[0].Name)' using exact $($candidate.Type) '$($candidate.Value)'"
                    }
                    return $matches[0]
                }

                if ($matches.Count -gt 1) {
                    Write-WarnMessage "Local AD fallback found multiple groups for candidate '$($candidate.Value)' using $($candidate.Type) and filter [$filter]. Skipping this candidate."
                    break
                }
            }
            catch {
                Write-Verbose "AD group lookup failed for candidate '$($candidate.Value)' with filter [$filter]: $($_.Exception.Message)"
            }
        }

        $script:adGroupCache[$cacheKey] = $null
    }

    return $null
}

function Convert-AdMemberToExportObject {
    param([Parameter(Mandatory)]$AdObject)

    $objectClass = @($AdObject.ObjectClass)[-1]
    $mail = $AdObject.Mail
    $upn = $AdObject.UserPrincipalName
    $displayName = $AdObject.DisplayName
    $id = $null
    if ($AdObject.ObjectGUID) {
        $id = "$($AdObject.ObjectGUID)"
    }

    $memberObjectType = switch -Regex ($objectClass) {
        'user' { 'User'; break }
        'contact' { 'OrganizationalContact'; break }
        'group' { 'Group'; break }
        'computer' { 'Device'; break }
        default { 'Unknown' }
    }

    return [PSCustomObject]@{
        Id                = $id
        DisplayName       = $displayName
        UserPrincipalName = $upn
        Mail              = $mail
        UserType          = if ($memberObjectType -eq 'User') { 'Member' } else { '' }
        MemberObjectType  = $memberObjectType
    }
}

function Get-LocalAdGroupMembers {
    param(
        [Parameter(Mandatory)]$SourcePrincipal,
        [Parameter(Mandatory)]$EntraGroup
    )

    if (-not (Initialize-ActiveDirectoryFallback)) {
        return @()
    }

    $candidates = [System.Collections.Generic.List[object]]::new()
    foreach ($mailCandidate in @(
        $EntraGroup.Mail
        $SourcePrincipal.PrincipalEmail
        (Get-EmailAddressFromText -Text $SourcePrincipal.PrincipalLoginName)
    )) {
        if (-not [string]::IsNullOrWhiteSpace($mailCandidate)) {
            $candidates.Add([PSCustomObject]@{
                Type  = 'Mail'
                Value = $mailCandidate
            }) | Out-Null
        }
    }

    $samAccountNameCandidate = Get-AdSafeSamAccountName -Value $SourcePrincipal.PrincipalLoginName
    if (-not [string]::IsNullOrWhiteSpace($samAccountNameCandidate)) {
        $candidates.Add([PSCustomObject]@{
            Type  = 'SamAccountName'
            Value = $samAccountNameCandidate
        }) | Out-Null
    }

    foreach ($nameCandidate in @(
        $EntraGroup.DisplayName
        $SourcePrincipal.PrincipalTitle
    )) {
        if (-not [string]::IsNullOrWhiteSpace($nameCandidate)) {
            $trimmedNameCandidate = $nameCandidate.Trim()
            $candidates.Add([PSCustomObject]@{
                Type  = 'DisplayName'
                Value = $trimmedNameCandidate
            }) | Out-Null
            $candidates.Add([PSCustomObject]@{
                Type  = 'Name'
                Value = $trimmedNameCandidate
            }) | Out-Null
        }
    }

    $adGroup = Resolve-AdGroupByCandidates -Candidates $candidates
    if (-not $adGroup) {
        if (-not $script:activeDirectoryFallbackWarningShown) {
            $script:activeDirectoryFallbackWarningShown = $true
            Write-WarnMessage "Local AD fallback is enabled but the SharePoint/Entra group could not be matched to an on-prem AD group. The script tried mail, proxyAddresses, samAccountName, and exact unique name/displayName matches."
        }
        return @()
    }

    Write-Info "  Using local AD fallback for group '$($adGroup.Name)'"
    $resolvedMembers = [System.Collections.Generic.List[object]]::new()

    try {
        $adMembers = @(Get-ADGroupMember -Identity $adGroup.DistinguishedName -Recursive -ErrorAction Stop)
        foreach ($adMember in $adMembers) {
            try {
                switch -Regex ($adMember.ObjectClass) {
                    '^user$' {
                        $fullUser = Get-ADUser -Identity $adMember.DistinguishedName -Properties displayName,mail,userPrincipalName,objectGUID -ErrorAction Stop
                        $resolvedMembers.Add((Convert-AdMemberToExportObject -AdObject $fullUser)) | Out-Null
                        break
                    }
                    '^contact$' {
                        $fullContact = Get-ADObject -Identity $adMember.DistinguishedName -Properties displayName,mail,objectGUID,objectClass -ErrorAction Stop
                        $resolvedMembers.Add((Convert-AdMemberToExportObject -AdObject $fullContact)) | Out-Null
                        break
                    }
                    default {
                        $fullObject = Get-ADObject -Identity $adMember.DistinguishedName -Properties displayName,mail,objectGUID,objectClass,userPrincipalName -ErrorAction Stop
                        $resolvedMembers.Add((Convert-AdMemberToExportObject -AdObject $fullObject)) | Out-Null
                    }
                }
            }
            catch {
                Write-Verbose "  AD member lookup failed for '$($adMember.DistinguishedName)': $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-WarnMessage "Failed to enumerate local AD members for group '$($adGroup.Name)': $($_.Exception.Message)"
        return @()
    }

    return @($resolvedMembers)
}

function Test-ShouldUseLocalAdFallback {
    param([object[]]$Members)

    if (-not $Members -or $Members.Count -eq 0) {
        return $false
    }

    foreach ($member in $Members) {
        $memberObjectType = Get-DirectoryMemberObjectType -Member $member
        if ($memberObjectType -in @('Group', 'OrganizationalContact', 'ServicePrincipal', 'Device')) {
            continue
        }

        $id = Get-PropertyValue -InputObject $member -PropertyNames @('Id', 'id')
        $displayName = Get-PropertyValue -InputObject $member -PropertyNames @('DisplayName', 'displayName')
        $upn = Get-PropertyValue -InputObject $member -PropertyNames @('UserPrincipalName', 'userPrincipalName')
        $mail = Get-PropertyValue -InputObject $member -PropertyNames @('Mail', 'mail')

        if (-not [string]::IsNullOrWhiteSpace($id) -and
            (
                [string]::IsNullOrWhiteSpace($displayName) -or
                [string]::IsNullOrWhiteSpace($upn) -or
                [string]::IsNullOrWhiteSpace($mail)
            )) {
            return $true
        }
    }

    return $false
}

function Get-PrincipalEmail {
    param([Parameter(Mandatory)]$Principal)

    $email = Get-PropertyValue -InputObject $Principal -PropertyNames @('Email', 'email', 'Mail', 'mail')
    if (-not [string]::IsNullOrWhiteSpace("$email")) {
        return $email
    }

    try {
        $loadedEmail = Get-PnPProperty -ClientObject $Principal -Property Email
        if (-not [string]::IsNullOrWhiteSpace("$loadedEmail")) {
            return $loadedEmail
        }
    }
    catch {
        Write-Verbose "Could not load Email property for principal: $($_.Exception.Message)"
    }

    return $null
}

function Test-IsUserDirectoryObject {
    param([Parameter(Mandatory)][object]$Member)

    # 1. Check @odata.type if available
    $odataType = Get-PropertyValue -InputObject $Member -PropertyNames @('@odata.type', 'odata.type', 'ODataType')
    if ($odataType -match 'group$') {
        return $false
    }
    if ($odataType -match 'user$') {
        return $true
    }

    # 2. Check user-specific properties
    $upn = Get-PropertyValue -InputObject $Member -PropertyNames @('UserPrincipalName', 'userPrincipalName')
    $userType = Get-PropertyValue -InputObject $Member -PropertyNames @('UserType', 'userType')
    if (-not [string]::IsNullOrWhiteSpace("$upn") -or -not [string]::IsNullOrWhiteSpace("$userType")) {
        return $true
    }

    # 3. Check group-specific properties — if present, it's a group, not a user
    $securityEnabled = Get-PropertyValue -InputObject $Member -PropertyNames @('SecurityEnabled', 'securityEnabled')
    $groupTypes = Get-PropertyValue -InputObject $Member -PropertyNames @('GroupTypes', 'groupTypes')
    $mailEnabled = Get-PropertyValue -InputObject $Member -PropertyNames @('MailEnabled', 'mailEnabled')
    if ($null -ne $securityEnabled -or $null -ne $groupTypes -or $null -ne $mailEnabled) {
        return $false
    }

    return $false
}

function Get-DirectoryMemberObjectType {
    param([Parameter(Mandatory)][object]$Member)

    $odataType = Get-PropertyValue -InputObject $Member -PropertyNames @('@odata.type', 'odata.type', 'ODataType')
    if ($odataType -match 'organizationalcontact$|orgcontact$|contact$') {
        return 'OrganizationalContact'
    }
    if ($odataType -match 'serviceprincipal$') {
        return 'ServicePrincipal'
    }
    if ($odataType -match 'device$') {
        return 'Device'
    }
    if ($odataType -match 'group$') {
        return 'Group'
    }
    if ($odataType -match 'user$') {
        return 'User'
    }

    $securityEnabled = Get-PropertyValue -InputObject $Member -PropertyNames @('SecurityEnabled', 'securityEnabled')
    $groupTypes = Get-PropertyValue -InputObject $Member -PropertyNames @('GroupTypes', 'groupTypes')
    $mailEnabled = Get-PropertyValue -InputObject $Member -PropertyNames @('MailEnabled', 'mailEnabled')
    if ($null -ne $securityEnabled -or $null -ne $groupTypes -or $null -ne $mailEnabled) {
        return 'Group'
    }

    if (Test-IsUserDirectoryObject -Member $Member) {
        return 'User'
    }

    $appId = Get-PropertyValue -InputObject $Member -PropertyNames @('AppId', 'appId')
    if (-not [string]::IsNullOrWhiteSpace("$appId")) {
        return 'ServicePrincipal'
    }

    return 'Unknown'
}

function Get-GroupKind {
    param([Parameter(Mandatory)]$Group)

    $groupTypes = @($Group.GroupTypes)

    if ($groupTypes -contains 'Unified') {
        return 'Microsoft365Group'
    }

    if ($Group.MailEnabled -and $Group.SecurityEnabled) {
        return 'MailEnabledSecurityGroup'
    }

    if ($Group.MailEnabled) {
        return 'DistributionGroup'
    }

    if ($Group.SecurityEnabled) {
        return 'SecurityGroup'
    }

    return 'Other'
}

function Resolve-EntraGroupByIdentity {
    param([Parameter(Mandatory)][string]$Identity)

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        return $null
    }

    if ($entraGroupIdentityCache.ContainsKey($Identity)) {
        return $entraGroupIdentityCache[$Identity]
    }

    try {
        $resolved = Get-PnPEntraIDGroup -Identity $Identity -ErrorAction Stop
        $entraGroupIdentityCache[$Identity] = $resolved
        Write-Verbose "Resolved Entra group by identity '$Identity': $($resolved.DisplayName) ($($resolved.Id))"
        return $resolved
    }
    catch {
        $message = $_.Exception.Message
        Write-Verbose "Failed to resolve Entra group by identity '$Identity': $message"
        if (Test-IsLookupNotFoundError -Message $message) {
            $entraGroupIdentityCache[$Identity] = $null
        }
        return $null
    }
}

function Resolve-EntraGroupByDisplayName {
    param([Parameter(Mandatory)][string]$DisplayName)

    if ([string]::IsNullOrWhiteSpace($DisplayName)) {
        return $null
    }

    if ($entraGroupDisplayNameCache.ContainsKey($DisplayName)) {
        return $entraGroupDisplayNameCache[$DisplayName]
    }

    if ($null -eq $script:allEntraGroups) {
        Write-Info "Loading Entra ID groups for exact display-name validation"
        $script:allEntraGroups = @(Get-PnPEntraIDGroup)
    }

    $matchedGroups = @($script:allEntraGroups | Where-Object { $_.DisplayName -eq $DisplayName })

    if ($matchedGroups.Count -eq 1) {
        $entraGroupDisplayNameCache[$DisplayName] = $matchedGroups[0]
        return $matchedGroups[0]
    }

    if ($matchedGroups.Count -gt 1) {
        Write-WarnMessage "Display name '$DisplayName' matched multiple Entra ID groups. Skipping to avoid exporting the wrong users."
    }

    $entraGroupDisplayNameCache[$DisplayName] = $null
    return $null
}

function Resolve-EntraGroupFromPrincipal {
    param([Parameter(Mandatory)]$Principal)

    $candidateIdentities = [System.Collections.Generic.List[string]]::new()

    $objectId = Get-DirectoryObjectIdFromLoginName -LoginName $Principal.LoginName
    if ($objectId) {
        $candidateIdentities.Add($objectId) | Out-Null
    }

    if (-not [string]::IsNullOrWhiteSpace($Principal.Email)) {
        $candidateIdentities.Add($Principal.Email) | Out-Null
    }

    $mailFromLoginName = Get-EmailAddressFromText -Text $Principal.LoginName
    if (-not [string]::IsNullOrWhiteSpace($mailFromLoginName)) {
        $candidateIdentities.Add($mailFromLoginName) | Out-Null
    }

    Write-Info "  Attempting resolution with candidates: $($candidateIdentities -join ', ')"
    foreach ($identity in ($candidateIdentities | Select-Object -Unique)) {
        $group = Resolve-EntraGroupByIdentity -Identity $identity
        if ($group) {
            Write-Info "  Resolved via identity '$identity' → '$($group.DisplayName)'"
            return $group
        }
    }

    Write-Info "  Falling back to display name lookup: '$($Principal.Title)'"
    return Resolve-EntraGroupByDisplayName -DisplayName $Principal.Title
}

function Get-RoleAssignmentsForSecurableObject {
    param(
        [Parameter(Mandatory)]$ClientObject,
        [Parameter(Mandatory)][string]$AccessPath,
        [Parameter(Mandatory)][string]$AccessContainer,
        $List,
        $ListItem
    )

    $assignments = [System.Collections.Generic.List[object]]::new()
    $roleAssignments = Get-PnPProperty -ClientObject $ClientObject -Property RoleAssignments

    $listId = $null
    $listTitle = $null
    $itemId = $null
    if ($null -ne $List) {
        $listId = $List.Id
        $listTitle = $List.Title
    }
    if ($null -ne $ListItem) {
        $itemId = $ListItem.Id
    }

    foreach ($roleAssignment in $roleAssignments) {
        $member = Get-PnPProperty -ClientObject $roleAssignment -Property Member
        $roleDefinitionBindings = Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings
        $permissionLevels = @($roleDefinitionBindings | ForEach-Object { $_.Name }) -join '; '
        $principalEmail = Get-PrincipalEmail -Principal $member

        $assignments.Add([PSCustomObject]@{
            AccessPath          = $AccessPath
            AccessContainer     = $AccessContainer
            PermissionLevels    = $permissionLevels
            PrincipalId         = $member.Id
            PrincipalTitle      = $member.Title
            PrincipalLoginName  = $member.LoginName
            PrincipalEmail      = $principalEmail
            PrincipalType       = $member.PrincipalType
            ListId              = $listId
            ListTitle           = $listTitle
            ItemId              = $itemId
        }) | Out-Null
    }

    return $assignments
}

function Get-DirectMembersFromSharePointGroup {
    param([Parameter(Mandatory)]$SharePointGroup)

    $directMembers = [System.Collections.Generic.List[object]]::new()
    try {
        $rawMembers = @(Get-PnPGroupMember -Group $SharePointGroup)
        Write-Info "  SP group '$($SharePointGroup.Title)' has $($rawMembers.Count) raw members"
        foreach ($member in $rawMembers) {
            Write-Verbose "    Member: '$($member.Title)' Login='$($member.LoginName)' Type='$($member.PrincipalType)'"
            if (Test-IsSpecialClaimPrincipal -LoginName $member.LoginName) {
                Write-Verbose "    → Skipped (special claim)"
                continue
            }

            $memberObj = [PSCustomObject]@{
                PrincipalId         = $member.Id
                PrincipalTitle      = $member.Title
                PrincipalLoginName  = $member.LoginName
                PrincipalEmail      = (Get-PrincipalEmail -Principal $member)
                PrincipalType       = "$($member.PrincipalType)"
            }
            $isGroup = Test-IsProbablyGroupPrincipal -Principal $memberObj
            Write-Info "    Member: '$($member.Title)' Login='$($member.LoginName)' Type='$($member.PrincipalType)' IsGroup=$isGroup"
            $directMembers.Add($memberObj) | Out-Null
        }
    }
    catch {
        Write-WarnMessage "Failed to enumerate members of SharePoint group '$($SharePointGroup.Title)': $($_.Exception.Message)"
    }

    return @($directMembers)
}

function Get-MembersFromSharePointGroup {
    param(
        [Parameter(Mandatory)]$SharePointGroup,
        [System.Collections.Generic.HashSet[string]]$VisitedGroupIds
    )

    $cacheKey = "$($SharePointGroup.Id)"
    $isTopLevel = ($PSBoundParameters.ContainsKey('VisitedGroupIds') -eq $false)

    if ($isTopLevel -and $sharePointGroupMemberCache.ContainsKey($cacheKey)) {
        return $sharePointGroupMemberCache[$cacheKey]
    }

    if ($null -eq $VisitedGroupIds) {
        $VisitedGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    if (-not $VisitedGroupIds.Add($cacheKey)) {
        Write-WarnMessage "Detected nested SharePoint group cycle at '$($SharePointGroup.Title)'. Skipping repeated traversal."
        return @()
    }

    $flattenedMembers = [System.Collections.Generic.List[object]]::new()
    foreach ($memberObj in @(Get-DirectMembersFromSharePointGroup -SharePointGroup $SharePointGroup)) {
        if (Test-IsSharePointGroupPrincipal -Principal $memberObj) {
            try {
                $nestedGroup = Get-SharePointGroupById -GroupId $memberObj.PrincipalId
                foreach ($nestedMember in @(Get-MembersFromSharePointGroup -SharePointGroup $nestedGroup -VisitedGroupIds $VisitedGroupIds)) {
                    $flattenedMembers.Add($nestedMember) | Out-Null
                }
            }
            catch {
                Write-WarnMessage "Failed to expand nested SharePoint group '$($memberObj.PrincipalTitle)': $($_.Exception.Message)"
            }

            continue
        }

        $flattenedMembers.Add($memberObj) | Out-Null
    }

    if ($isTopLevel) {
        $sharePointGroupMemberCache[$cacheKey] = @($flattenedMembers)
        return $sharePointGroupMemberCache[$cacheKey]
    }

    return @($flattenedMembers)
}

function Get-SharePointGroupById {
    param([Parameter(Mandatory)]$GroupId)

    $cacheKey = "$GroupId"
    if ($sharePointGroupCache.ContainsKey($cacheKey)) {
        return $sharePointGroupCache[$cacheKey]
    }

    $sharePointGroup = Get-PnPGroup -Identity $GroupId
    $sharePointGroupCache[$cacheKey] = $sharePointGroup
    return $sharePointGroup
}

function Add-Principal {
    param(
        [ValidateNotNull()][System.Collections.Generic.List[object]]$Principals,
        [ValidateNotNull()][System.Collections.Generic.HashSet[string]]$Seen,
        [Parameter(Mandatory)]$Assignment,
        [Parameter(Mandatory)]$Principal,
        [Parameter(Mandatory)]$GrantedThroughPrincipal
    )

    if (Test-IsSpecialClaimPrincipal -LoginName $Principal.PrincipalLoginName) {
        return
    }

    $key = "{0}|{1}|{2}|{3}|{4}" -f $Assignment.AccessPath, $Assignment.AccessContainer, $Assignment.PermissionLevels, $GrantedThroughPrincipal.PrincipalLoginName, $Principal.PrincipalLoginName
    if (-not $Seen.Add($key)) {
        return
    }

    $Principals.Add([PSCustomObject]@{
        AccessPath          = $Assignment.AccessPath
        AccessContainer     = $Assignment.AccessContainer
        PermissionLevels    = $Assignment.PermissionLevels
        ListId              = $Assignment.ListId
        ListTitle           = $Assignment.ListTitle
        ItemId              = $Assignment.ItemId
        GrantedToId         = $GrantedThroughPrincipal.PrincipalId
        GrantedToTitle      = $GrantedThroughPrincipal.PrincipalTitle
        GrantedToLoginName  = $GrantedThroughPrincipal.PrincipalLoginName
        GrantedToType       = "$($GrantedThroughPrincipal.PrincipalType)"
        PrincipalId         = $Principal.PrincipalId
        PrincipalTitle      = $Principal.PrincipalTitle
        PrincipalLoginName  = $Principal.PrincipalLoginName
        PrincipalEmail      = $Principal.PrincipalEmail
        PrincipalType       = "$($Principal.PrincipalType)"
        IsDirectUser        = -not (Test-IsProbablyGroupPrincipal -Principal $Principal)
    }) | Out-Null
}

function Expand-AssignmentToPrincipals {
    param(
        [ValidateNotNull()][System.Collections.Generic.List[object]]$Principals,
        [ValidateNotNull()][System.Collections.Generic.HashSet[string]]$Seen,
        [Parameter(Mandatory)]$Assignment
    )

    if (Test-IsDirectoryGroupPrincipal -Principal $Assignment) {
        Add-Principal -Principals $Principals -Seen $Seen -Assignment $Assignment -Principal $Assignment -GrantedThroughPrincipal $Assignment
        return
    }

    if (Test-IsUserPrincipal -Principal $Assignment) {
        if (-not (Test-IsSpecialClaimPrincipal -LoginName $Assignment.PrincipalLoginName)) {
            Add-Principal -Principals $Principals -Seen $Seen -Assignment $Assignment -Principal $Assignment -GrantedThroughPrincipal $Assignment
        }
        return
    }

    if (-not (Test-IsSharePointGroupPrincipal -Principal $Assignment)) {
        return
    }

    $sharePointGroup = Get-SharePointGroupById -GroupId $Assignment.PrincipalId
    foreach ($spGroupMember in @(Get-MembersFromSharePointGroup -SharePointGroup $sharePointGroup)) {
        Add-Principal -Principals $Principals -Seen $Seen -Assignment $Assignment -Principal $spGroupMember -GrantedThroughPrincipal ([PSCustomObject]@{
            PrincipalId         = $Assignment.PrincipalId
            PrincipalTitle      = $Assignment.PrincipalTitle
            PrincipalLoginName  = $Assignment.PrincipalLoginName
            PrincipalType       = $Assignment.PrincipalType
        })
    }
}

function Get-ListItemAccessPath {
    param(
        [Parameter(Mandatory)]$List,
        [Parameter(Mandatory)]$ListItem
    )

    $fileSystemObjectType = "$($ListItem.FileSystemObjectType)"
    if ([string]::IsNullOrWhiteSpace($fileSystemObjectType)) {
        switch ("$($ListItem.FieldValues['FSObjType'])") {
            '1' { $fileSystemObjectType = 'Folder' }
            '0' {
                if ($List.BaseType -eq 'DocumentLibrary') {
                    $fileSystemObjectType = 'File'
                }
                else {
                    $fileSystemObjectType = 'ListItem'
                }
            }
        }
    }

    if ($fileSystemObjectType -eq 'Folder') {
        return 'FolderPermission'
    }

    if ($fileSystemObjectType -eq 'File') {
        return 'FilePermission'
    }

    return 'ListItemPermission'
}

function Get-ListItemAccessContainer {
    param(
        [Parameter(Mandatory)]$List,
        [Parameter(Mandatory)]$ListItem
    )

    $fileRef = $ListItem.FieldValues['FileRef']
    if (-not [string]::IsNullOrWhiteSpace("$fileRef")) {
        return $fileRef
    }

    return "{0} [Item ID {1}]" -f $List.Title, $ListItem.Id
}

function Get-SitePrincipals {
    $principals = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $web = Get-PnPWeb -Includes RoleAssignments, Title, Url

    Write-Info "Inspecting direct site permissions"
    foreach ($assignment in @(Get-RoleAssignmentsForSecurableObject -ClientObject $web -AccessPath 'DirectWebPermission' -AccessContainer $web.Url)) {
        Expand-AssignmentToPrincipals -Principals $principals -Seen $seen -Assignment $assignment
    }

    if ($SiteLevelOnly) {
        Write-Info "SiteLevelOnly: skipping list, library, folder, and item permission scanning"
        return $principals
    }

    Write-Info "Inspecting list and library permissions"
    foreach ($list in Get-PnPList -Includes HasUniqueRoleAssignments, Hidden, Title, RootFolder, BaseType, ItemCount) {
        if ($list.Hidden) {
            continue
        }

        $rootFolder = Get-PnPProperty -ClientObject $list -Property RootFolder
        $listPath = Get-PnPProperty -ClientObject $rootFolder -Property ServerRelativeUrl

        if ($list.HasUniqueRoleAssignments) {
            foreach ($assignment in (Get-RoleAssignmentsForSecurableObject -ClientObject $list -AccessPath 'ListPermission' -AccessContainer $listPath -List $list)) {
                Expand-AssignmentToPrincipals -Principals $principals -Seen $seen -Assignment $assignment
            }
        }

        if ($list.ItemCount -lt 1) {
            continue
        }

        Write-Info "  Scanning '$($list.Title)' ($($list.ItemCount) items) for unique permissions"
        $itemIndex = 0
        foreach ($listItem in @(Get-PnPListItem -List $list -PageSize $PageSize -Fields 'ID', 'FileRef', 'FileLeafRef', 'FSObjType')) {
            $itemIndex++
            if ($itemIndex % 250 -eq 0) {
                Write-Info "    Processed $itemIndex / $($list.ItemCount) items in '$($list.Title)'"
            }
            $hasUniqueRoleAssignments = Get-PnPProperty -ClientObject $listItem -Property HasUniqueRoleAssignments
            if (-not $hasUniqueRoleAssignments) {
                continue
            }

            $itemPath = Get-ListItemAccessContainer -List $list -ListItem $listItem
            $itemAccessPath = Get-ListItemAccessPath -List $list -ListItem $listItem

            foreach ($assignment in (Get-RoleAssignmentsForSecurableObject -ClientObject $listItem -AccessPath $itemAccessPath -AccessContainer $itemPath -List $list -ListItem $listItem)) {
                Expand-AssignmentToPrincipals -Principals $principals -Seen $seen -Assignment $assignment
            }
        }
    }

    return $principals
}

function Get-ExpandedGroupUsers {
    param([Parameter(Mandatory)]$EntraGroup)

    if ($IncludeTransitiveMembers) {
        try {
            return @(Get-PnPEntraIDGroupMember -Identity $EntraGroup.Id -Transitive)
        }
        catch {
            $message = $_.Exception.Message
            if ($message -match 'transitive|unsupported|Request_UnsupportedQuery|not supported') {
                Write-WarnMessage "Transitive expansion is not supported for '$($EntraGroup.DisplayName)'. Retrying with direct members only. $message"
            }
            else {
                throw
            }
        }
    }

    return @(Get-PnPEntraIDGroupMember -Identity $EntraGroup.Id)
}

function Resolve-MemberUserDetails {
    param([Parameter(Mandatory)][object]$Member)

    $id = Get-PropertyValue -InputObject $Member -PropertyNames @('Id', 'id')
    $displayName = Get-PropertyValue -InputObject $Member -PropertyNames @('DisplayName', 'displayName')
    $upn = Get-PropertyValue -InputObject $Member -PropertyNames @('UserPrincipalName', 'userPrincipalName')
    $mail = Get-PropertyValue -InputObject $Member -PropertyNames @('Mail', 'mail')
    $userType = Get-PropertyValue -InputObject $Member -PropertyNames @('UserType', 'userType')
    $sourceIdentity = $null
    $memberObjectType = Get-DirectoryMemberObjectType -Member $Member
    $resolutionStatus = 'Resolved'

    # Group member expansion often returns sparse user objects. Hydrate from Graph
    # whenever any key identity property is missing and we have a stable identifier.
    $canAttemptUserHydration =
        ($memberObjectType -eq 'User') -or
        (
            $memberObjectType -eq 'Unknown' -and
            (
                -not [string]::IsNullOrWhiteSpace($id) -or
                -not [string]::IsNullOrWhiteSpace($upn) -or
                -not [string]::IsNullOrWhiteSpace($mail)
            )
        )

    $needsHydration =
        $canAttemptUserHydration -and
        (
            -not [string]::IsNullOrWhiteSpace($id) -or
            -not [string]::IsNullOrWhiteSpace($upn) -or
            -not [string]::IsNullOrWhiteSpace($mail)
        ) -and (
            [string]::IsNullOrWhiteSpace($displayName) -or
            [string]::IsNullOrWhiteSpace($upn) -or
            [string]::IsNullOrWhiteSpace($mail) -or
            [string]::IsNullOrWhiteSpace($userType)
        )

    if ($needsHydration) {
        $lookupIdentities = [System.Collections.Generic.List[string]]::new()
        foreach ($identity in @($id, $upn, $mail)) {
            if (-not [string]::IsNullOrWhiteSpace($identity)) {
                $lookupIdentities.Add($identity) | Out-Null
            }
        }
        $fullUser = $null
        foreach ($identity in ($lookupIdentities | Select-Object -Unique)) {
            Write-Verbose "  Fetching user details for identity '$identity'"
            $fullUser = Resolve-EntraUserByIdentity -Identity $identity
            if ($fullUser) {
                $sourceIdentity = $identity
                break
            }
        }

        if ($fullUser) {
            $fullId = Get-PropertyValue -InputObject $fullUser -PropertyNames @('Id', 'id')
            $fullDisplayName = Get-PropertyValue -InputObject $fullUser -PropertyNames @('DisplayName', 'displayName')
            $fullUpn = Get-PropertyValue -InputObject $fullUser -PropertyNames @('UserPrincipalName', 'userPrincipalName')
            $fullMail = Get-PropertyValue -InputObject $fullUser -PropertyNames @('Mail', 'mail')
            $fullUserType = Get-PropertyValue -InputObject $fullUser -PropertyNames @('UserType', 'userType')

            if (-not [string]::IsNullOrWhiteSpace($fullId)) { $id = $fullId }
            if (-not [string]::IsNullOrWhiteSpace($fullDisplayName)) { $displayName = $fullDisplayName }
            if (-not [string]::IsNullOrWhiteSpace($fullUpn)) { $upn = $fullUpn }
            if (-not [string]::IsNullOrWhiteSpace($fullMail)) { $mail = $fullMail }
            if (-not [string]::IsNullOrWhiteSpace($fullUserType)) { $userType = $fullUserType }
            $memberObjectType = 'User'
        }
    }

    if ([string]::IsNullOrWhiteSpace($mail)) {
        $mail = $upn
    }

    if ([string]::IsNullOrWhiteSpace($displayName)) {
        $displayName = $upn
    }

    if ([string]::IsNullOrWhiteSpace($displayName) -and
        [string]::IsNullOrWhiteSpace($upn) -and
        [string]::IsNullOrWhiteSpace($mail)) {
        $warningIdentity = $sourceIdentity
        if ([string]::IsNullOrWhiteSpace($warningIdentity)) {
            $warningIdentity = $id
        }
        if ([string]::IsNullOrWhiteSpace($warningIdentity)) {
            $warningIdentity = '(unknown identity)'
        }
        if ($memberObjectType -eq 'User' -or $canAttemptUserHydration) {
            Write-WarnMessage "Could not hydrate user details for member '$warningIdentity'. Verify Microsoft Graph user-read permissions such as User.Read.All or Directory.Read.All for the PnP app."
            $displayName = "(unreadable user)"
            $resolutionStatus = 'UnreadableUserDetails'
            $memberObjectType = 'User'
        }
        else {
            $displayName = "(non-user member)"
            $resolutionStatus = "NonUserMember:$memberObjectType"
        }
    }
    elseif ($memberObjectType -eq 'Unknown') {
        $resolutionStatus = 'UnknownDirectoryMember'
    }
    elseif ($memberObjectType -ne 'User') {
        $resolutionStatus = "NonUserMember:$memberObjectType"
    }

    return [PSCustomObject]@{
        Id                = $id
        DisplayName       = $displayName
        UserPrincipalName = $upn
        Mail              = $mail
        UserType          = $userType
        MemberObjectType  = $memberObjectType
        ResolutionStatus  = $resolutionStatus
    }
}

function Add-ExportRows {
    param(
        [Parameter(Mandatory)][string]$WebUrl,
        [Parameter(Mandatory)]$SourcePrincipal,
        [Parameter(Mandatory)]$EntraGroup,
        [Parameter(Mandatory)][object[]]$Members
    )

    $groupKind = Get-GroupKind -Group $EntraGroup

    foreach ($member in $Members) {
        $details = Resolve-MemberUserDetails -Member $member
        $memberId = $details.Id
        $memberDisplayName = $details.DisplayName
        $memberUpn = $details.UserPrincipalName
        $memberMail = $details.Mail
        $memberUserType = $details.UserType
        $memberObjectType = $details.MemberObjectType
        $resolutionStatus = $details.ResolutionStatus
        $row = [PSCustomObject]@{
            SiteUrl                = $WebUrl
            AccessPath             = $SourcePrincipal.AccessPath
            AccessContainer        = $SourcePrincipal.AccessContainer
            PermissionLevels       = $SourcePrincipal.PermissionLevels
            GrantedToPrincipal     = $SourcePrincipal.GrantedToTitle
            GrantedToLoginName     = $SourcePrincipal.GrantedToLoginName
            GrantedToPrincipalType = $SourcePrincipal.GrantedToType
            SharePointPrincipal    = $SourcePrincipal.PrincipalTitle
            SharePointLoginName    = $SourcePrincipal.PrincipalLoginName
            SharePointPrincipalType= $SourcePrincipal.PrincipalType
            EntraGroupId           = $EntraGroup.Id
            EntraGroupName         = $EntraGroup.DisplayName
            EntraGroupMail         = $EntraGroup.Mail
            EntraGroupType         = $groupKind
            UserId                 = $memberId
            UserDisplayName        = $memberDisplayName
            UserPrincipalName      = $memberUpn
            UserMail               = $memberMail
            UserType               = $memberUserType
            MemberObjectType       = $memberObjectType
            ResolutionStatus       = $resolutionStatus
        }

        $rowKey = @(
            $row.SiteUrl
            $row.AccessPath
            $row.AccessContainer
            $row.PermissionLevels
            $row.GrantedToLoginName
            $row.SharePointLoginName
            $row.EntraGroupId
            $row.UserId
            $row.UserPrincipalName
            $row.UserMail
        ) -join '|'

        if ($exportRowKeys.Add($rowKey)) {
            $results.Add($row) | Out-Null
        }
    }
}

function Add-DirectUserExportRow {
    param(
        [Parameter(Mandatory)][string]$WebUrl,
        [Parameter(Mandatory)]$UserPrincipal
    )

    $userIdentity = Get-UserIdentityFromLoginName -LoginName $UserPrincipal.PrincipalLoginName
    $userEmail = $UserPrincipal.PrincipalEmail
    if ([string]::IsNullOrWhiteSpace($userEmail)) {
        $userEmail = Get-EmailAddressFromText -Text $UserPrincipal.PrincipalLoginName
    }

    $row = [PSCustomObject]@{
        SiteUrl                = $WebUrl
        AccessPath             = $UserPrincipal.AccessPath
        AccessContainer        = $UserPrincipal.AccessContainer
        PermissionLevels       = $UserPrincipal.PermissionLevels
        GrantedToPrincipal     = $UserPrincipal.GrantedToTitle
        GrantedToLoginName     = $UserPrincipal.GrantedToLoginName
        GrantedToPrincipalType = $UserPrincipal.GrantedToType
        SharePointPrincipal    = $UserPrincipal.PrincipalTitle
        SharePointLoginName    = $UserPrincipal.PrincipalLoginName
        SharePointPrincipalType= $UserPrincipal.PrincipalType
        EntraGroupId           = ''
        EntraGroupName         = ''
        EntraGroupMail         = ''
        EntraGroupType         = 'DirectUser'
        UserId                 = ''
        UserDisplayName        = $UserPrincipal.PrincipalTitle
        UserPrincipalName      = $userIdentity
        UserMail               = $userEmail
        UserType               = ''
        MemberObjectType       = 'User'
        ResolutionStatus       = 'Resolved'
    }

    $rowKey = @(
        $row.SiteUrl
        $row.AccessPath
        $row.AccessContainer
        $row.PermissionLevels
        $row.GrantedToLoginName
        $row.SharePointLoginName
        ''
        ''
        $row.UserPrincipalName
        $row.UserMail
    ) -join '|'

    if ($exportRowKeys.Add($rowKey)) {
        $results.Add($row) | Out-Null
    }
}

function Add-UnresolvedGroupExportRow {
    param(
        [Parameter(Mandatory)][string]$WebUrl,
        [Parameter(Mandatory)]$GroupPrincipal
    )

    $row = [PSCustomObject]@{
        SiteUrl                = $WebUrl
        AccessPath             = $GroupPrincipal.AccessPath
        AccessContainer        = $GroupPrincipal.AccessContainer
        PermissionLevels       = $GroupPrincipal.PermissionLevels
        GrantedToPrincipal     = $GroupPrincipal.GrantedToTitle
        GrantedToLoginName     = $GroupPrincipal.GrantedToLoginName
        GrantedToPrincipalType = $GroupPrincipal.GrantedToType
        SharePointPrincipal    = $GroupPrincipal.PrincipalTitle
        SharePointLoginName    = $GroupPrincipal.PrincipalLoginName
        SharePointPrincipalType= $GroupPrincipal.PrincipalType
        EntraGroupId           = (Get-DirectoryObjectIdFromLoginName -LoginName $GroupPrincipal.PrincipalLoginName)
        EntraGroupName         = $GroupPrincipal.PrincipalTitle
        EntraGroupMail         = $GroupPrincipal.PrincipalEmail
        EntraGroupType         = 'Unresolved'
        UserId                 = ''
        UserDisplayName        = ''
        UserPrincipalName      = ''
        UserMail               = ''
        UserType               = ''
        MemberObjectType       = 'Group'
        ResolutionStatus       = 'UnresolvedGroup'
    }

    $rowKey = @(
        $row.SiteUrl
        $row.AccessPath
        $row.AccessContainer
        $row.PermissionLevels
        $row.GrantedToLoginName
        $row.SharePointLoginName
        $row.EntraGroupId
        ''
        ''
        ''
    ) -join '|'

    if ($exportRowKeys.Add($rowKey)) {
        $results.Add($row) | Out-Null
    }
}

function Add-EmptyGroupExportRow {
    param(
        [Parameter(Mandatory)][string]$WebUrl,
        [Parameter(Mandatory)]$SourcePrincipal,
        [Parameter(Mandatory)]$EntraGroup
    )

    $groupKind = Get-GroupKind -Group $EntraGroup
    $row = [PSCustomObject]@{
        SiteUrl                = $WebUrl
        AccessPath             = $SourcePrincipal.AccessPath
        AccessContainer        = $SourcePrincipal.AccessContainer
        PermissionLevels       = $SourcePrincipal.PermissionLevels
        GrantedToPrincipal     = $SourcePrincipal.GrantedToTitle
        GrantedToLoginName     = $SourcePrincipal.GrantedToLoginName
        GrantedToPrincipalType = $SourcePrincipal.GrantedToType
        SharePointPrincipal    = $SourcePrincipal.PrincipalTitle
        SharePointLoginName    = $SourcePrincipal.PrincipalLoginName
        SharePointPrincipalType= $SourcePrincipal.PrincipalType
        EntraGroupId           = $EntraGroup.Id
        EntraGroupName         = $EntraGroup.DisplayName
        EntraGroupMail         = $EntraGroup.Mail
        EntraGroupType         = $groupKind
        UserId                 = ''
        UserDisplayName        = ''
        UserPrincipalName      = ''
        UserMail               = ''
        UserType               = '(empty group)'
        MemberObjectType       = 'Group'
        ResolutionStatus       = 'EmptyGroup'
    }

    $rowKey = @(
        $row.SiteUrl
        $row.AccessPath
        $row.AccessContainer
        $row.PermissionLevels
        $row.GrantedToLoginName
        $row.SharePointLoginName
        $row.EntraGroupId
        ''
        ''
        ''
    ) -join '|'

    if ($exportRowKeys.Add($rowKey)) {
        $results.Add($row) | Out-Null
    }
}

function Get-ConnectedM365GroupMembers {
    try {
        $site = Get-PnPSite -Includes GroupId
        $groupId = Get-PnPProperty -ClientObject $site -Property GroupId

        if ($null -eq $groupId -or $groupId -eq [Guid]::Empty) {
            Write-Info "Site is not connected to a Microsoft 365 group"
            return $null
        }

        $groupIdString = $groupId.ToString()
        Write-Info "Site is connected to M365 group: $groupIdString"

        $entraGroup = Resolve-EntraGroupByIdentity -Identity $groupIdString
        if (-not $entraGroup) {
            Write-WarnMessage "Could not resolve connected M365 group ($groupIdString) in Entra ID."
            return $null
        }

        return $entraGroup
    }
    catch {
        Write-WarnMessage "Failed to detect connected M365 group: $($_.Exception.Message)"
        return $null
    }
}

function Process-SingleWeb {
    param([Parameter(Mandatory)][string]$WebUrl)

    Write-Info "========================================="
    Write-Info "Processing web: $WebUrl"
    Write-Info "========================================="

    Write-Info "Finding all principals and uniquely secured content"
    $sitePrincipals = Get-SitePrincipals

    $directUserPrincipals = @($sitePrincipals | Where-Object { $_.IsDirectUser -eq $true })
    $groupPrincipals = @($sitePrincipals | Where-Object { $_.IsDirectUser -ne $true })

    Write-Info "Found $($directUserPrincipals.Count) direct user entries and $($groupPrincipals.Count) directory group entries"
    foreach ($gp in $groupPrincipals) {
        Write-Verbose "  Group principal: '$($gp.PrincipalTitle)' Login='$($gp.PrincipalLoginName)' Type='$($gp.PrincipalType)' GrantedThrough='$($gp.GrantedToTitle)'"
    }

    foreach ($userPrincipal in $directUserPrincipals) {
        Add-DirectUserExportRow -WebUrl $WebUrl -UserPrincipal $userPrincipal
    }

    if ($groupPrincipals.Count -gt 0) {
        foreach ($principal in $groupPrincipals) {
            Write-Info "Resolving group '$($principal.PrincipalTitle)'"

            $entraGroup = Resolve-EntraGroupFromPrincipal -Principal ([PSCustomObject]@{
                LoginName = $principal.PrincipalLoginName
                Email     = $principal.PrincipalEmail
                Title     = $principal.PrincipalTitle
            })

            if (-not $entraGroup) {
                Write-WarnMessage "Could not resolve '$($principal.PrincipalTitle)' ($($principal.PrincipalLoginName)) to an Entra ID group. Adding unresolved row."
                Add-UnresolvedGroupExportRow -WebUrl $WebUrl -GroupPrincipal $principal
                continue
            }

            try {
                $members = Get-ExpandedGroupUsers -EntraGroup $entraGroup
            }
            catch {
                Write-WarnMessage "Failed to expand members for '$($entraGroup.DisplayName)': $($_.Exception.Message)"
                $members = @()
            }

            $adFallbackMembers = @()
            if ((-not $members -or $members.Count -eq 0) -or (Test-ShouldUseLocalAdFallback -Members $members)) {
                $adFallbackMembers = @(Get-LocalAdGroupMembers -SourcePrincipal $principal -EntraGroup $entraGroup)
                if ($adFallbackMembers.Count -gt 0) {
                    Write-Info "  Local AD fallback returned $($adFallbackMembers.Count) members"
                    Add-ExportRows -WebUrl $WebUrl -SourcePrincipal $principal -EntraGroup $entraGroup -Members $adFallbackMembers
                    Grant-ExpandedUsersDirectAccessFromMembers -SourcePrincipal $principal -Members $adFallbackMembers
                    continue
                }
            }

            if (-not $members -or $members.Count -eq 0) {
                Write-WarnMessage "No members returned for '$($entraGroup.DisplayName)'. The group may be empty or transitive expansion may not be supported."
                Add-EmptyGroupExportRow -WebUrl $WebUrl -SourcePrincipal $principal -EntraGroup $entraGroup
                continue
            }

            $userMembers = @(
                $members | Where-Object {
                    (Get-DirectoryMemberObjectType -Member $_) -in @('User', 'Unknown')
                }
            )
            Write-Info "  Expanded '$($entraGroup.DisplayName)': $($members.Count) total members, $($userMembers.Count) user-like members"

            # Dump first member's properties for diagnostics (only when verbose)
            if ($members.Count -gt 0) {
                $firstMember = $members[0]
                $directProps = @($firstMember.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ', '
                Write-Verbose "  First member direct properties: $directProps"
                foreach ($apDictName in @('AdditionalData', 'AdditionalProperties')) {
                    $additionalProps = $firstMember.PSObject.Properties[$apDictName]
                    if ($additionalProps -and $additionalProps.Value -and $additionalProps.Value -is [System.Collections.IDictionary]) {
                        $apStr = @($additionalProps.Value.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '
                        Write-Verbose "  First member $apDictName`: $apStr"
                        break
                    }
                }
            }

            foreach ($m in $userMembers) {
                $mDetails = Resolve-MemberUserDetails -Member $m
                Write-Info "    - $($mDetails.DisplayName) ($($mDetails.UserPrincipalName))"
            }

            Add-ExportRows -WebUrl $WebUrl -SourcePrincipal $principal -EntraGroup $entraGroup -Members $members
            Grant-ExpandedUsersDirectAccessFromMembers -SourcePrincipal $principal -Members $members
        }
    }
    else {
        Write-Info "No directory-backed groups found on this web."
    }
}

# --- Main execution ---

if ([string]::IsNullOrWhiteSpace($ClientId) -and [string]::IsNullOrWhiteSpace($env:ENTRAID_APP_ID)) {
    throw "ClientId is required. Pass -ClientId or configure ENTRAID_APP_ID for PnP PowerShell."
}

if ($RemoveExpandedDirectoryGroupsAfterGrant -and -not $GrantExpandedUsersDirectly) {
    throw "RemoveExpandedDirectoryGroupsAfterGrant requires -GrantExpandedUsersDirectly."
}

$targetSiteUrls = @(Get-TargetSiteUrls)
$totalWebsProcessed = 0

Write-Info "Target root sites to process: $($targetSiteUrls.Count)"
foreach ($targetSiteUrl in $targetSiteUrls) {
    try {
        $processedWebCount = Invoke-SiteCollectionProcessing -RootSiteUrl $targetSiteUrl
        $totalWebsProcessed += $processedWebCount
    }
    catch {
        Write-WarnMessage "Failed to process root site '$targetSiteUrl': $($_.Exception.Message)"
    }
}

# Export results
$outputDirectory = Split-Path -Path $OutputCsvPath -Parent
if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$finalRows = $results |
    Sort-Object SiteUrl, AccessPath, AccessContainer, GrantedToLoginName, SharePointLoginName, EntraGroupName, UserPrincipalName, UserMail

$finalRows | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding utf8

Write-Host ""
Write-Host "CSV exported to: $OutputCsvPath" -ForegroundColor Green
Write-Host "Root sites    : $($targetSiteUrls.Count)" -ForegroundColor Green
Write-Host "Webs processed: $totalWebsProcessed" -ForegroundColor Green
Write-Host "Rows exported : $($finalRows.Count)" -ForegroundColor Green
if ($GrantExpandedUsersDirectly) {
    Write-Host "Direct grants : $directGrantSuccessCount" -ForegroundColor Green
}
if ($RemoveExpandedDirectoryGroupsAfterGrant) {
    Write-Host "Groups removed: $sourceGroupRemovalCount" -ForegroundColor Green
}

if (-not [string]::IsNullOrWhiteSpace($UploadFolderServerRelativeUrl)) {
    $uploadTargetSite = $targetSiteUrls[0]
    $uploadConnectParameters = New-ConnectParameters -Url $uploadTargetSite
    Write-Info "Reconnecting to '$uploadTargetSite' for upload"
    $null = Connect-PnPOnline @uploadConnectParameters
    Write-Info "Uploading CSV back to SharePoint"
    Add-PnPFile -Path $OutputCsvPath -Folder $UploadFolderServerRelativeUrl | Out-Null
    Write-Host "Uploaded to   : $UploadFolderServerRelativeUrl" -ForegroundColor Green
    Write-Host "Upload site   : $uploadTargetSite" -ForegroundColor Green
}

if ($warnings.Count -gt 0) {
    Write-Host ""
    Write-Host "Warnings:" -ForegroundColor Yellow
    $warnings | Sort-Object -Unique | ForEach-Object { Write-Host " - $_" -ForegroundColor Yellow }
}

try {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
catch {
    Write-Verbose "Disconnect-PnPOnline: $($_.Exception.Message)"
}
