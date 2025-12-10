#requires -Version 5.1
#requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users, ImportExcel

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter()]
    [string]$WorksheetName,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [string]$PlaceholderOwnerUserPrincipalName = 'GroupUserKristaps@wearedots.com',

    [Parameter()]
    [switch]$ForceReconnect
)

# Hard-coded Excel source for validation runs.
$ExcelPath = 'C:\Users\KristapsD\OneDrive - WeAreDots\Desktop\DeleteUsers.xlsx'

function Ensure-Module {
    param([Parameter(Mandatory = $true)][string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not installed. Install with: Install-Module $Name"
    }

    Import-Module $Name -ErrorAction Stop | Out-Null
}

function ConvertTo-CleanGroupRow {
    param([Parameter(Mandatory = $true)][psobject]$Row)

    $normalized = [ordered]@{}

    foreach ($prop in $Row.PSObject.Properties) {
        $rawName = $prop.Name
        if (-not $rawName) { continue }

        $trimmedName = $rawName.Trim()
        if (-not $trimmedName) { continue }

        $standardName = switch -Regex ($trimmedName) {
            '^displayname$' { 'DisplayName'; break }
            '^id$' { 'Id'; break }
            '^source$' { 'Source'; break }
            '^grouptype$' { 'GroupType'; break }
            default { $trimmedName }
        }

        if ($normalized.Contains($standardName)) { continue }

        $value = $prop.Value
        if ($value -is [string]) { $value = $value.Trim() }
        $normalized[$standardName] = $value
    }

    return [pscustomobject]$normalized
}

function Get-PrincipalLabel {
    param([Parameter(Mandatory = $true)][Microsoft.Graph.PowerShell.Models.IMicrosoftGraphDirectoryObject]$Principal)

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

    if ($displayName -and $upn) { return "$displayName <$upn>" }
    if ($displayName) { return $displayName }
    if ($upn) { return $upn }
    return $Principal.Id
}

function New-PrincipalRecord {
    param(
        [Parameter(Mandatory = $true)][string]$GroupName,
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][ValidateSet('Owner', 'Member')][string]$Role,
        [Parameter(Mandatory = $true)][Microsoft.Graph.PowerShell.Models.IMicrosoftGraphDirectoryObject]$Principal
    )

    $props = @{}
    if ($Principal.PSObject.Properties['AdditionalProperties']) {
        $props = $Principal.AdditionalProperties
    }

    $displayName = $props['displayName']
    if (-not $displayName -and $Principal.PSObject.Properties['DisplayName']) {
        $displayName = $Principal.DisplayName
    }

    $upnOrMail = $props['userPrincipalName']
    if (-not $upnOrMail -and $props['mail']) {
        $upnOrMail = $props['mail']
    }

    $type = $props['@odata.type']
    if ($type) {
        $type = $type -replace '#microsoft.graph.', ''
    }

    return [pscustomobject]@{
        GroupDisplayName           = $GroupName
        GroupId                    = $GroupId
        Role                       = $Role
        PrincipalDisplayName       = $displayName
        PrincipalUserPrincipalName = $upnOrMail
        PrincipalObjectId          = $Principal.Id
        PrincipalType              = $type
    }
}

function Get-DirectoryObjectResourceSegment {
    param([Parameter(Mandatory = $true)][Microsoft.Graph.PowerShell.Models.IMicrosoftGraphDirectoryObject]$DirectoryObject)

    $userTypes = @(
        [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser],
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
    )
    if ($userTypes | Where-Object { $DirectoryObject -is $_ }) { return 'users' }

    $spTypes = @(
        [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphServicePrincipal],
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphServicePrincipal]
    )
    if ($spTypes | Where-Object { $DirectoryObject -is $_ }) { return 'servicePrincipals' }

    $groupTypes = @(
        [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphGroup],
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphGroup]
    )
    if ($groupTypes | Where-Object { $DirectoryObject -is $_ }) { return 'groups' }

    $odataType = $null
    if ($DirectoryObject.PSObject.Properties['AdditionalProperties']) {
        $odataType = $DirectoryObject.AdditionalProperties['@odata.type']
    }

    switch -Regex ($odataType) {
        'user' { return 'users' }
        'serviceprincipal' { return 'servicePrincipals' }
        'group' { return 'groups' }
        default { return 'directoryObjects' }
    }
}

function Add-GroupOwnerReference {
    param(
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][Microsoft.Graph.PowerShell.Models.IMicrosoftGraphDirectoryObject]$OwnerObject
    )

    $ownerId = $OwnerObject.Id
    if ([string]::IsNullOrWhiteSpace($ownerId)) {
        throw 'OwnerObject.Id was not populated.'
    }

    $resourceSegment = Get-DirectoryObjectResourceSegment -DirectoryObject $OwnerObject
    $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/$resourceSegment/$ownerId" }

    $addCmd = Get-Command -Name Add-MgGroupOwnerByRef -ErrorAction SilentlyContinue
    if ($addCmd) {
        Add-MgGroupOwnerByRef -GroupId $GroupId -DirectoryObjectId $ownerId -ErrorAction Stop
        return
    }

    $jsonBody = $body | ConvertTo-Json -Depth 3 -Compress
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/owners/`$ref"
    Invoke-MgGraphRequest -Method POST -Uri $uri -Body $jsonBody -ContentType 'application/json' -ErrorAction Stop
}

foreach ($module in @('ImportExcel', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users')) {
    Ensure-Module -Name $module
}

if (-not (Test-Path $ExcelPath -PathType Leaf)) {
    throw "Excel file was not found at '$ExcelPath'. Update the script if the location changes."
}

if (-not $OutputPath) {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path -Path (Split-Path -Path $ExcelPath -Parent) -ChildPath "RemovedPrincipals_$timestamp.xlsx"
}

if ($ForceReconnect -or -not (Get-MgContext)) {
    Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
    Connect-MgGraph -Scopes @('Group.ReadWrite.All', 'GroupMember.Read.All', 'User.Read.All') -NoWelcome
}

if (-not (Get-MgContext)) {
    throw 'Unable to acquire a Microsoft Graph context. Verify your credentials and consent.'
}

Write-Host "Reading groups from '$ExcelPath'" -ForegroundColor Cyan

$importParams = @{ Path = $ExcelPath }
if ($WorksheetName) {
    $importParams['WorksheetName'] = $WorksheetName
}

try {
    $groupRows = Import-Excel @importParams
} catch {
    throw "Failed to import Excel file '$ExcelPath'. $_"
}

if (-not $groupRows) {
    Write-Warning 'The Excel file contained no rows.'
    return
}

$groupRows = @(
    foreach ($row in $groupRows) {
        $cleanRow = ConvertTo-CleanGroupRow -Row $row
        if (-not $cleanRow.PSObject.Properties['Id']) { continue }
        $idValue = ([string]$cleanRow.Id).Trim()
        if ([string]::IsNullOrWhiteSpace($idValue)) { continue }
        $cleanRow | Add-Member -NotePropertyName 'Id' -NotePropertyValue $idValue -Force
        $cleanRow
    }
)

if (-not $groupRows) {
    Write-Warning "The Excel file contained no rows with a usable 'Id' column."
    return
}

$removalRecords = @()
$totalAttempted = 0
$totalGroupsProcessed = 0
$totalFailures = 0

foreach ($groupRecord in $groupRows) {
    $totalAttempted++

    $groupId = ([string]$groupRecord.Id).Trim()
    $groupName = if ([string]::IsNullOrWhiteSpace([string]$groupRecord.DisplayName)) { '<unnamed>' } else { [string]$groupRecord.DisplayName }

    Write-Host "`nProcessing $groupName ($groupId)" -ForegroundColor Cyan

    try {
        $null = Get-MgGroup -GroupId $groupId -ErrorAction Stop
        $totalGroupsProcessed++
    } catch {
        $totalFailures++
        Write-Warning ("  Unable to locate group with Id {0}. {1}" -f $groupId, $_.Exception.Message)
        continue
    }

    $members = @()
    $owners = @()

    try { $members = @(Get-MgGroupMember -GroupId $groupId -All -ErrorAction Stop) } catch { Write-Warning ("  Failed to retrieve members. {0}" -f $_.Exception.Message) }
    try { $owners = @(Get-MgGroupOwner -GroupId $groupId -All -ErrorAction Stop) } catch { Write-Warning ("  Failed to retrieve owners. {0}" -f $_.Exception.Message) }

    if ($owners.Count -eq 0) {
        Write-Host "  Owners       : none" -ForegroundColor Yellow
    } else {
        Write-Host ("  Owners       : {0}" -f $owners.Count)
    }

    if ($members.Count -eq 0) {
        Write-Host "  Members      : none" -ForegroundColor Yellow
    } else {
        Write-Host ("  Members      : {0}" -f $members.Count)
    }

    $placeholderOwner = $null
    if ($PlaceholderOwnerUserPrincipalName -and $owners.Count -gt 0) {
        try {
            $placeholderOwner = Get-MgUser -UserId $PlaceholderOwnerUserPrincipalName -ErrorAction Stop
            $placeholderIsOwner = $owners | Where-Object { $_.Id -eq $placeholderOwner.Id }
            if (-not $placeholderIsOwner) {
                Write-Host ("  Adding placeholder owner '{0}'" -f $PlaceholderOwnerUserPrincipalName) -ForegroundColor Cyan
                if ($PSCmdlet.ShouldProcess($PlaceholderOwnerUserPrincipalName, "Add placeholder owner to $groupName")) {
                    try {
                        Add-GroupOwnerReference -GroupId $groupId -OwnerObject $placeholderOwner
                        $owners = @( $owners + $placeholderOwner )
                    } catch {
                        Write-Warning ("  Failed to add placeholder owner. {0}" -f $_.Exception.Message)
                        $placeholderOwner = $null
                    }
                } else {
                    Write-Host '  Skipped adding placeholder owner because ShouldProcess was denied.' -ForegroundColor Yellow
                    $placeholderOwner = $null
                }
            }
        } catch {
            Write-Warning ("  Unable to resolve placeholder owner '{0}'. {1}" -f $PlaceholderOwnerUserPrincipalName, $_.Exception.Message)
            $placeholderOwner = $null
        }
    }

    foreach ($member in $members) {
        $label = Get-PrincipalLabel -Principal $member
        if (-not $PSCmdlet.ShouldProcess($label, "Remove member from $groupName")) { continue }
        try {
            Remove-MgGroupMemberByRef -GroupId $groupId -DirectoryObjectId $member.Id -ErrorAction Stop
            $removalRecords += New-PrincipalRecord -GroupName $groupName -GroupId $groupId -Role 'Member' -Principal $member
            Write-Host ("    Removed member: {0}" -f $label)
        } catch {
            $totalFailures++
            Write-Warning ("    Failed to remove member {0}. {1}" -f $label, $_.Exception.Message)
        }
    }

    $ownersToRemove = $owners
    if ($placeholderOwner) {
        $ownersToRemove = $owners | Where-Object { $_.Id -ne $placeholderOwner.Id }
    }

    foreach ($owner in $ownersToRemove) {
        $label = Get-PrincipalLabel -Principal $owner
        if (-not $PSCmdlet.ShouldProcess($label, "Remove owner from $groupName")) { continue }
        try {
            Remove-MgGroupOwnerByRef -GroupId $groupId -DirectoryObjectId $owner.Id -ErrorAction Stop
            $removalRecords += New-PrincipalRecord -GroupName $groupName -GroupId $groupId -Role 'Owner' -Principal $owner
            Write-Host ("    Removed owner : {0}" -f $label)
        } catch {
            $totalFailures++
            Write-Warning ("    Failed to remove owner {0}. {1}" -f $label, $_.Exception.Message)
        }
    }
}

Write-Host "`nSummary" -ForegroundColor Cyan
Write-Host ("  Rows evaluated  : {0}" -f $totalAttempted)
Write-Host ("  Groups processed : {0}" -f $totalGroupsProcessed)
Write-Host ("  Failures         : {0}" -f $totalFailures)
Write-Host ("  Principals removed: {0}" -f $removalRecords.Count)

if (-not $removalRecords -or $removalRecords.Count -eq 0) {
    Write-Warning 'No owners or members were removed.'
    return
}

if ($OutputPath) {
    Write-Host ("Exporting removal log to '{0}'" -f $OutputPath) -ForegroundColor Cyan
    $removalRecords | Export-Excel -Path $OutputPath -WorksheetName 'RemovedPrincipals' -AutoSize -FreezeTopRow -TableName 'RemovedPrincipals' -BoldTopRow
}

Write-Host ("Completed removal run for {0} groups." -f $totalGroupsProcessed) -ForegroundColor Green
