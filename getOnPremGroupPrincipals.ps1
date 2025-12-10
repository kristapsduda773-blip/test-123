#requires -Version 5.1
#requires -Modules ActiveDirectory, ImportExcel

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ExcelPath = 'C:\Users\KristapsD\OneDrive - WeAreDots\Desktop\DeleteUsers.xlsx',

    [Parameter()]
    [string]$WorksheetName,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [switch]$IncludeNestedMembers,

    [Parameter()]
    [switch]$ForceModuleReload
)

function Ensure-Module {
    param([Parameter(Mandatory = $true)][string]$Name)

    if ($ForceModuleReload) {
        Remove-Module -Name $Name -ErrorAction SilentlyContinue
    }

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not installed. Install with: Install-WindowsFeature RSAT-AD-PowerShell"
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

function New-PrincipalRecord {
    param(
        [Parameter(Mandatory = $true)][string]$GroupName,
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][ValidateSet('Owner', 'Member')][string]$Role,
        [Parameter(Mandatory = $true)][psobject]$Principal
    )

    return [pscustomobject]@{
        GroupDisplayName           = $GroupName
        GroupId                    = $GroupId
        Role                       = $Role
        PrincipalDisplayName       = $Principal.DisplayName
        PrincipalUserPrincipalName = $Principal.UserPrincipalName
        PrincipalObjectId          = $Principal.ObjectId
        PrincipalType              = $Principal.ObjectClass
    }
}

function Resolve-PrincipalDetails {
    param(
        [Parameter(Mandatory = $true)][Microsoft.ActiveDirectory.Management.ADObject]$DirectoryObject
    )

    $mail = $DirectoryObject.mail
    $upn = $DirectoryObject.UserPrincipalName
    if (-not $upn -and $DirectoryObject.SamAccountName) {
        $upn = $DirectoryObject.SamAccountName
    }

    [pscustomobject]@{
        DisplayName       = $DirectoryObject.DisplayName
        UserPrincipalName = $upn
        ObjectId          = $DirectoryObject.ObjectGUID.Guid
        ObjectClass       = $DirectoryObject.ObjectClass
    }
}

Ensure-Module -Name ActiveDirectory
Ensure-Module -Name ImportExcel

if (-not (Test-Path $ExcelPath -PathType Leaf)) {
    throw "Excel file was not found at '$ExcelPath'. Update -ExcelPath parameter if needed."
}

if (-not $OutputPath) {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path -Path (Split-Path -Path $ExcelPath -Parent) -ChildPath "OnPremGroupPrincipals_$timestamp.xlsx"
}

$importParams = @{ Path = $ExcelPath }
if ($WorksheetName) { $importParams['WorksheetName'] = $WorksheetName }

try {
    $groupRows = Import-Excel @importParams
} catch {
    throw "Failed to import Excel file '$ExcelPath'. $_"
}

if (-not $groupRows) {
    Write-Warning 'The Excel file contained no rows to process.'
    return
}

$groupRows = @(
    foreach ($row in $groupRows) {
        $cleanRow = ConvertTo-CleanGroupRow -Row $row
        if (-not $cleanRow.Id -and -not $cleanRow.DisplayName) { continue }
        $cleanRow
    }
)

if (-not $groupRows) {
    Write-Warning 'No usable rows (DisplayName or Id) were present in the spreadsheet.'
    return
}

$principalRecords = @()
$processedGroups = 0
$missingGroups = 0

foreach ($groupRecord in $groupRows) {
    $groupId = $null
    $groupIdentity = $null

    if ($groupRecord.Id) {
        $guidRef = [ref]([guid]::Empty)
        if ([guid]::TryParse([string]$groupRecord.Id, $guidRef)) {
            $groupIdentity = $guidRef.Value
        }
    }

    if (-not $groupIdentity -and $groupRecord.DisplayName) {
        $groupIdentity = $groupRecord.DisplayName
    }

    if (-not $groupIdentity) {
        Write-Warning "Skipping row because no Id or DisplayName is available."
        continue
    }

    Write-Host "`n------------------------------------------------------------"
    Write-Host ("Group lookup : {0}" -f $groupIdentity) -ForegroundColor Cyan

    try {
        $group = Get-ADGroup -Identity $groupIdentity -Properties ManagedBy, mail, SamAccountName, DisplayName, ObjectGuid -ErrorAction Stop
        $processedGroups++
        $groupName = if ($group.DisplayName) { $group.DisplayName } else { $group.SamAccountName }
        $groupId = $group.ObjectGUID.Guid
        Write-Host ("DisplayName : {0}" -f $groupName)
        Write-Host ("ObjectGuid  : {0}" -f $groupId)
        Write-Host ("ManagedBy   : {0}" -f ($group.ManagedBy ?? '<not set>'))
    } catch {
        $missingGroups++
        Write-Warning ("Unable to find group '{0}'. {1}" -f $groupIdentity, $_.Exception.Message)
        continue
    }

    $owners = @()
    if ($group.ManagedBy) {
        try {
            $ownerObject = Get-ADObject -Identity $group.ManagedBy -Properties DisplayName, SamAccountName, mail, UserPrincipalName, ObjectClass
            $owners = @(Resolve-PrincipalDetails -DirectoryObject $ownerObject)
            foreach ($owner in $owners) {
                $principalRecords += New-PrincipalRecord -GroupName $group.DisplayName -GroupId $groupId -Role 'Owner' -Principal $owner
            }
            Write-Host ("Owners     : {0}" -f ($owners.Count))
        } catch {
            Write-Warning ("Failed to resolve ManagedBy reference '{0}'. {1}" -f $group.ManagedBy, $_.Exception.Message)
        }
    } else {
        Write-Host 'Owners     : <none>' -ForegroundColor Yellow
    }

    try {
        $memberParams = @{ Identity = $group.DistinguishedName; ErrorAction = 'Stop' }
        if ($IncludeNestedMembers) { $memberParams['Recursive'] = $true }
        $members = @(Get-ADGroupMember @memberParams)
    } catch {
        Write-Warning ("Failed to retrieve members for '{0}'. {1}" -f $group.DisplayName, $_.Exception.Message)
        continue
    }

    if ($members.Count -eq 0) {
        Write-Host 'Members    : <none>' -ForegroundColor Yellow
    } else {
        Write-Host ("Members    : {0}" -f $members.Count)
    }

    foreach ($member in $members) {
        try {
            $memberObject = Get-ADObject -Identity $member.DistinguishedName -Properties DisplayName, SamAccountName, mail, UserPrincipalName, ObjectClass
            $memberDetails = Resolve-PrincipalDetails -DirectoryObject $memberObject
            $principalRecords += New-PrincipalRecord -GroupName $group.DisplayName -GroupId $groupId -Role 'Member' -Principal $memberDetails
        } catch {
            Write-Warning ("Unable to resolve member {0}. {1}" -f $member.DistinguishedName, $_.Exception.Message)
        }
    }
}

Write-Host "`nSummary" -ForegroundColor Cyan
Write-Host ("  Groups processed : {0}" -f $processedGroups)
Write-Host ("  Groups missing   : {0}" -f $missingGroups)
Write-Host ("  Principals found : {0}" -f $principalRecords.Count)

if (-not $principalRecords -or $principalRecords.Count -eq 0) {
    Write-Warning 'No owners or members were enumerated from Active Directory.'
    return
}

Write-Host ("Exporting results to '{0}'" -f $OutputPath) -ForegroundColor Cyan
$principalRecords | Export-Excel -Path $OutputPath -WorksheetName 'OnPremGroupPrincipals' -AutoSize -FreezeTopRow -TableName 'OnPremGroupPrincipals' -BoldTopRow

Write-Host 'Completed on-prem group principal enumeration.' -ForegroundColor Green
