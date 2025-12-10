#requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, ImportExcel

[CmdletBinding()]
param(
    [Parameter()]
    [string]$WorksheetName,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [switch]$ForceReconnect
)

# Hard-coded Excel source for validation runs.
$ExcelPath = 'C:\Users\KristapsD\OneDrive - WeAreDots\Desktop\DeleteUsers.xlsx'

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

function New-PrincipalRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName,

        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Owner', 'Member')]
        [string]$Role,

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

function ConvertTo-CleanGroupRow {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Row
    )

    $normalized = [ordered]@{}

    foreach ($prop in $Row.PSObject.Properties) {
        $rawName = $prop.Name
        if (-not $rawName) {
            continue
        }

        $trimmedName = $rawName.Trim()
        if (-not $trimmedName) {
            continue
        }

        $standardName = switch -Regex ($trimmedName) {
            '^displayname$' { 'DisplayName'; break }
            '^id$' { 'Id'; break }
            '^source$' { 'Source'; break }
            '^grouptype$' { 'GroupType'; break }
            default { $trimmedName }
        }

        if ($normalized.Contains($standardName)) {
            continue
        }

        $value = $prop.Value
        if ($value -is [string]) {
            $value = $value.Trim()
        }

        $normalized[$standardName] = $value
    }

    return [pscustomobject]$normalized
}

foreach ($module in @('ImportExcel', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')) {
    Ensure-Module -Name $module
}

if (-not (Test-Path $ExcelPath -PathType Leaf)) {
    throw "Excel file was not found at '$ExcelPath'. Update the script if the location changes."
}

if (-not $OutputPath) {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path -Path (Split-Path -Path $ExcelPath -Parent) -ChildPath "GroupPrincipals_$timestamp.xlsx"
}

if ($ForceReconnect -or -not (Get-MgContext)) {
    Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
    Connect-MgGraph -Scopes @('Group.Read.All', 'GroupMember.Read.All') -NoWelcome
}

if (-not (Get-MgContext)) {
    throw 'Unable to acquire a Microsoft Graph context. Verify your credentials and consent.'
}

Write-Host "Reading groups from '$ExcelPath'"

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

        if (-not $cleanRow.PSObject.Properties['Id']) {
            continue
        }

        $idValue = ([string]$cleanRow.Id).Trim()
        if ([string]::IsNullOrWhiteSpace($idValue)) {
            continue
        }

        $cleanRow | Add-Member -NotePropertyName 'Id' -NotePropertyValue $idValue -Force
        $cleanRow
    }
)

if (-not $groupRows) {
    Write-Warning "The Excel file contained no rows with a usable 'Id' column."
    return
}

$principalRows = @()
$attemptedGroups = 0
$resolvedGroups = 0
$missingGroups = 0

foreach ($groupRecord in $groupRows) {
    $attemptedGroups++

    $groupId = ([string]$groupRecord.Id).Trim()
    $groupName = if ([string]::IsNullOrWhiteSpace([string]$groupRecord.DisplayName)) { '<unnamed>' } else { [string]$groupRecord.DisplayName }

    Write-Host "`nProcessing $groupName ($groupId)" -ForegroundColor Cyan

    try {
        $null = Get-MgGroup -GroupId $groupId -ErrorAction Stop
        $resolvedGroups++
    } catch {
        $missingGroups++
        Write-Warning ("  Unable to locate group with Id {0}. {1}" -f $groupId, $_.Exception.Message)
        continue
    }

    try {
        $owners = @(Get-MgGroupOwner -GroupId $groupId -All -ErrorAction Stop)
    } catch {
        Write-Warning ("  Failed to read owners. {0}" -f $_.Exception.Message)
        $owners = @()
    }

    if ($owners.Count -eq 0) {
        Write-Host '  Owners     : none' -ForegroundColor Yellow
    } else {
        Write-Host ("  Owners     : {0}" -f $owners.Count)
        foreach ($owner in $owners) {
            $principalRows += New-PrincipalRecord -GroupName $groupName -GroupId $groupId -Role 'Owner' -Principal $owner
        }
    }

    try {
        $members = @(Get-MgGroupMember -GroupId $groupId -All -ErrorAction Stop)
    } catch {
        Write-Warning ("  Failed to read members. {0}" -f $_.Exception.Message)
        $members = @()
    }

    if ($members.Count -eq 0) {
        Write-Host '  Members    : none' -ForegroundColor Yellow
    } else {
        Write-Host ("  Members    : {0}" -f $members.Count)
        foreach ($member in $members) {
            $principalRows += New-PrincipalRecord -GroupName $groupName -GroupId $groupId -Role 'Member' -Principal $member
        }
    }
}

Write-Host "`nSummary" -ForegroundColor Cyan
Write-Host ("  Rows evaluated : {0}" -f $attemptedGroups)
Write-Host ("  Found in Graph : {0}" -f $resolvedGroups)
Write-Host ("  Missing        : {0}" -f $missingGroups)

if (-not $principalRows -or $principalRows.Count -eq 0) {
    Write-Warning 'No owners or members were returned by Microsoft Graph.'
    if ($resolvedGroups -eq 0) {
        Write-Warning 'Verify that the Id column contains valid Entra ID group object IDs and that you have consent for Group.Read.All/GroupMember.Read.All.'
    }
    return
}

Write-Host "`nExporting results to '$OutputPath'" -ForegroundColor Cyan
$principalRows |
    Export-Excel -Path $OutputPath -WorksheetName 'GroupPrincipals' -AutoSize -FreezeTopRow -TableName 'GroupPrincipals' -BoldTopRow

Write-Host ('Completed export for {0} groups. {1} rows written.' -f $resolvedGroups, $principalRows.Count) -ForegroundColor Green
