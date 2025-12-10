#requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter()]
    [string]$GroupDisplayName = 'DEV-Common-Localization',

    [Parameter()]
    [string]$NewOwnerUserPrincipalName = 'GroupUserKristaps@wearedots.com',

    [Parameter()]
    [switch]$ForceReconnect
)

function Ensure-Module {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not installed. Install with: Install-Module $Name"
    }

    Import-Module $Name -ErrorAction Stop | Out-Null
}

function Get-PrincipalLabel {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphDirectoryObject]$Principal
    )

    $props = @{}
    if ($Principal.PSObject.Properties['AdditionalProperties']) {
        $props = $Principal.AdditionalProperties
    }

    $displayName = $props['displayName']
    if (-not $displayName -and $Principal.PSObject.Properties['DisplayName']) {
        $displayName = $Principal.DisplayName
    }

    $upn = $props['userPrincipalName']
    if (-not $upn -and $props['mail']) {
        $upn = $props['mail']
    }

    if ($displayName -and $upn) {
        return "$displayName <$upn>"
    }

    if ($displayName) {
        return $displayName
    }

    if ($upn) {
        return $upn
    }

    return $Principal.Id
}

function Get-DirectoryObjectResourceSegment {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphDirectoryObject]$DirectoryObject
    )

    $odataType = $null
    if ($DirectoryObject.PSObject.Properties['AdditionalProperties']) {
        $odataType = $DirectoryObject.AdditionalProperties['@odata.type']
    }

    if (-not $odataType -and $DirectoryObject.PSObject.TypeNames) {
        foreach ($typeName in $DirectoryObject.PSObject.TypeNames) {
            if ($typeName -match 'MicrosoftGraph([A-Za-z]+)$') {
                $odataType = "#microsoft.graph.{0}" -f $matches[1].ToLower()
                break
            }
        }
    }

    switch -Regex ($odataType) {
        'serviceprincipal' { return 'servicePrincipals' }
        'group' { return 'groups' }
        'user' { return 'users' }
        default { return 'directoryObjects' }
    }
}

function Add-GroupOwnerReference {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphDirectoryObject]$OwnerObject
    )

    $ownerId = $OwnerObject.Id
    if ([string]::IsNullOrWhiteSpace($ownerId)) {
        throw 'OwnerObject.Id was not populated.'
    }

    $resourceSegment = Get-DirectoryObjectResourceSegment -DirectoryObject $OwnerObject

    $addCmd = Get-Command -Name Add-MgGroupOwnerByRef -ErrorAction SilentlyContinue
    $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/$resourceSegment/$ownerId" }

    if ($addCmd) {
        Add-MgGroupOwnerByRef -GroupId $GroupId -BodyParameter $body -ErrorAction Stop
        return
    }

    $jsonBody = $body | ConvertTo-Json -Depth 3 -Compress
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/owners/\$ref" -Body $jsonBody -ContentType 'application/json' -ErrorAction Stop
}

foreach ($module in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users')) {
    Ensure-Module -Name $module
}

if ($ForceReconnect -or -not (Get-MgContext)) {
    Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
    Connect-MgGraph -Scopes @('Group.ReadWrite.All', 'GroupMember.Read.All', 'User.Read.All') -NoWelcome
}

if (-not (Get-MgContext)) {
    throw 'Unable to acquire a Microsoft Graph context. Verify your credentials and consent.'
}

if ([string]::IsNullOrWhiteSpace($GroupDisplayName)) {
    throw 'GroupDisplayName cannot be empty.'
}

$escapedName = $GroupDisplayName.Replace("'", "''")
Write-Host ("Looking up group '{0}'" -f $GroupDisplayName) -ForegroundColor Cyan

try {
    $matchingGroups = Get-MgGroup -Filter "displayName eq '$escapedName'" -ConsistencyLevel eventual -CountVariable ignored
} catch {
    throw "Failed to query Microsoft Graph for group '$GroupDisplayName'. $_"
}

if (-not $matchingGroups) {
    throw "No group named '$GroupDisplayName' was found. Verify the display name."
}

if ($matchingGroups.Count -gt 1) {
    $ids = $matchingGroups | Select-Object -ExpandProperty Id
    throw "Multiple groups named '$GroupDisplayName' were found. Specify the exact object Id instead. Matches: $($ids -join ', ')"
}

$targetGroup = $matchingGroups | Select-Object -First 1
Write-Host ("Target group Id: {0}" -f $targetGroup.Id)

$placeholderOwner = $null
if (-not [string]::IsNullOrWhiteSpace($NewOwnerUserPrincipalName)) {
    try {
        $placeholderOwner = Get-MgUser -UserId $NewOwnerUserPrincipalName -ErrorAction Stop
        Write-Host ("Resolved placeholder owner '{0}' (ObjectId: {1})" -f $NewOwnerUserPrincipalName, $placeholderOwner.Id)
    } catch {
        throw "Unable to resolve placeholder owner '$NewOwnerUserPrincipalName'. $_"
    }
}

try {
    $members = @(Get-MgGroupMember -GroupId $targetGroup.Id -All -ErrorAction Stop)
} catch {
    throw "Failed to retrieve members for group '$GroupDisplayName'. $_"
}

try {
    $owners = @(Get-MgGroupOwner -GroupId $targetGroup.Id -All -ErrorAction Stop)
} catch {
    throw "Failed to retrieve owners for group '$GroupDisplayName'. $_"
}

if ($members.Count -eq 0) {
    Write-Host "Group '$GroupDisplayName' has no members." -ForegroundColor Yellow
} else {
    Write-Host ("Removing {0} members from '{1}'" -f $members.Count, $GroupDisplayName) -ForegroundColor Red
}

if ($owners.Count -eq 0) {
    Write-Host "Group '$GroupDisplayName' has no owners." -ForegroundColor Yellow
} else {
    Write-Host ("Removing {0} owners from '{1}'" -f $owners.Count, $GroupDisplayName) -ForegroundColor Red
}

if ($placeholderOwner) {
    $isPlaceholderAlreadyOwner = $owners | Where-Object { $_.Id -eq $placeholderOwner.Id }
    if (-not $isPlaceholderAlreadyOwner) {
        Write-Host ("Adding placeholder owner '{0}' before removals..." -f $NewOwnerUserPrincipalName) -ForegroundColor Cyan
        $shouldAdd = $PSCmdlet.ShouldProcess($NewOwnerUserPrincipalName, "Add placeholder owner to $GroupDisplayName")
        try {
            if ($shouldAdd) {
                Add-GroupOwnerReference -GroupId $targetGroup.Id -OwnerObject $placeholderOwner
                # Refresh owner list to include newly added owner
                $owners = @( $owners + $placeholderOwner )
            } else {
                Write-Host '  Skipped adding placeholder owner because ShouldProcess was denied.' -ForegroundColor Yellow
            }
        } catch {
            throw "Failed to add placeholder owner '$NewOwnerUserPrincipalName'. $_"
        }
    } else {
        Write-Host ("Placeholder owner '{0}' is already assigned to the group." -f $NewOwnerUserPrincipalName) -ForegroundColor Cyan
    }
}

$memberRemoved = 0
$memberFailed = 0
$ownerRemoved = 0
$ownerFailed = 0

foreach ($member in $members) {
    $label = Get-PrincipalLabel -Principal $member
    if (-not $PSCmdlet.ShouldProcess($label, "Remove member from $GroupDisplayName")) {
        continue
    }

    try {
        Remove-MgGroupMemberByRef -GroupId $targetGroup.Id -DirectoryObjectId $member.Id -ErrorAction Stop
        $memberRemoved++
        Write-Host ("  Member removed: {0}" -f $label)
    } catch {
        $memberFailed++
        Write-Warning ("Failed to remove member {0}. {1}" -f $label, $_.Exception.Message)
    }
}

$ownersToRemove = $owners
if ($placeholderOwner) {
    $ownersToRemove = $owners | Where-Object { $_.Id -ne $placeholderOwner.Id }
}

foreach ($owner in $ownersToRemove) {
    $label = Get-PrincipalLabel -Principal $owner
    if (-not $PSCmdlet.ShouldProcess($label, "Remove owner from $GroupDisplayName")) {
        continue
    }

    try {
        Remove-MgGroupOwnerByRef -GroupId $targetGroup.Id -DirectoryObjectId $owner.Id -ErrorAction Stop
        $ownerRemoved++
        Write-Host ("  Owner removed : {0}" -f $label)
    } catch {
        $ownerFailed++
        Write-Warning ("Failed to remove owner {0}. {1}" -f $label, $_.Exception.Message)
    }
}

Write-Host ''
Write-Host 'Summary' -ForegroundColor Cyan
Write-Host ("  Members removed : {0}" -f $memberRemoved)
Write-Host ("  Member failures : {0}" -f $memberFailed)
Write-Host ("  Owners removed  : {0}" -f $ownerRemoved)
Write-Host ("  Owner failures  : {0}" -f $ownerFailed)

if ($memberFailed -gt 0 -or $ownerFailed -gt 0) {
    throw "Finished with removal failures (Members: $memberFailed, Owners: $ownerFailed). Review warnings above."
}

Write-Host ("Completed member and owner removal for '{0}'." -f $GroupDisplayName) -ForegroundColor Green
