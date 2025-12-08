#requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, ImportExcel

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ExcelPath,

    [Parameter()]
    [string]$WorksheetName,

    [Parameter()]
    [switch]$ForceReconnect
)

$DryRun = $true

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

function Format-DirectoryObject {
    param(
        [Parameter(Mandatory = $true)]
        [object]$DirectoryObject
    )

    $displayName = $DirectoryObject.DisplayName
    $principal = $DirectoryObject.UserPrincipalName

    if ($DirectoryObject.AdditionalProperties) {
        if (-not $displayName -and $DirectoryObject.AdditionalProperties.ContainsKey('displayName')) {
            $displayName = $DirectoryObject.AdditionalProperties['displayName']
        }

        if (-not $principal) {
            foreach ($key in @('userPrincipalName', 'mail', 'appId', 'servicePrincipalType')) {
                if ($DirectoryObject.AdditionalProperties.ContainsKey($key)) {
                    $principal = $DirectoryObject.AdditionalProperties[$key]
                    break
                }
            }
        }
    }

    if (-not $displayName) {
        $displayName = 'Unknown name'
    }

    if (-not $principal) {
        $principal = $DirectoryObject.Id
    }

    return '{0} ({1})' -f $displayName, $principal
}

foreach ($module in @('ImportExcel', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')) {
    Ensure-Module -Name $module
}

if ($ForceReconnect -or -not (Get-MgContext)) {
    Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
    Connect-MgGraph -Scopes @('Group.Read.All', 'GroupMember.Read.All') -NoWelcome
}

if (-not (Get-MgContext)) {
    throw 'Unable to acquire a Microsoft Graph context. Verify your credentials and consent.'
}

Write-Host ("Dry run mode is {0}" -f $(if ($DryRun) { 'ON' } else { 'OFF' })) -ForegroundColor Yellow

$importParams = @{
    Path = $ExcelPath
}

if ($WorksheetName) {
    $importParams['WorksheetName'] = $WorksheetName
}

try {
    $groupRows = Import-Excel @importParams
} catch {
    throw "Failed to import Excel file '$ExcelPath'. $_"
}

if (-not $groupRows) {
    Write-Warning 'The Excel file contained no rows to evaluate.'
    return
}

$cloudGroups = $groupRows | Where-Object { $_.Source -eq 'CloudOnly' }

if (-not $cloudGroups) {
    Write-Host "No groups with Source = 'CloudOnly' were found in the provided file."
    return
}

foreach ($groupRecord in $cloudGroups) {
    if (-not $groupRecord.Id) {
        Write-Warning ("Skipping row with DisplayName '{0}' because Id is missing." -f $groupRecord.DisplayName)
        continue
    }

    Write-Host "`n------------------------------------------------------------"

    try {
        $groupObject = Get-MgGroup -GroupId $groupRecord.Id -Property Id,DisplayName,GroupTypes -ErrorAction Stop
    } catch {
        Write-Warning ("Unable to retrieve Entra ID record for group Id {0}. {1}" -f $groupRecord.Id, $_.Exception.Message)
        $groupObject = $null
    }

    if ($groupObject) {
        $groupName = if ([string]::IsNullOrWhiteSpace($groupObject.DisplayName)) { '<unknown>' } else { [string]$groupObject.DisplayName }
    } else {
        $groupName = if ([string]::IsNullOrWhiteSpace([string]$groupRecord.DisplayName)) { '<unknown>' } else { [string]$groupRecord.DisplayName }
    }

    Write-Host ("DisplayName : {0}" -f $groupName) -ForegroundColor Cyan
    Write-Host ("Group Id    : {0}" -f $groupRecord.Id)
    Write-Host ("Source      : {0}" -f ($groupRecord.Source -as [string]))

    if ($groupObject -and $groupObject.GroupTypes) {
        $groupTypeValue = $groupObject.GroupTypes -join ', '
    } elseif ($groupRecord.GroupType) {
        $groupTypeValue = [string]$groupRecord.GroupType
    } else {
        $groupTypeValue = '<not specified>'
    }

    Write-Host ("GroupType   : {0}" -f $groupTypeValue)

    try {
        $owners = Get-MgGroupOwner -GroupId $groupRecord.Id -All -ErrorAction Stop
    } catch {
        Write-Warning ("Failed to retrieve owners for group {0}. {1}" -f $groupRecord.Id, $_.Exception.Message)
        $owners = @()
    }

    try {
        $members = Get-MgGroupMember -GroupId $groupRecord.Id -All -ErrorAction Stop
    } catch {
        Write-Warning ("Failed to retrieve members for group {0}. {1}" -f $groupRecord.Id, $_.Exception.Message)
        $members = @()
    }

    $ownerSummaries = @()
    foreach ($owner in $owners) {
        $ownerSummaries += (Format-DirectoryObject -DirectoryObject $owner)
    }

    $memberSummaries = @()
    foreach ($member in $members) {
        $memberSummaries += (Format-DirectoryObject -DirectoryObject $member)
    }

    Write-Host ("Owners   ({0})" -f $ownerSummaries.Count)
    if ($ownerSummaries.Count -gt 0) {
        $ownerSummaries | ForEach-Object { Write-Host ("  - {0}" -f $_) }
    } else {
        Write-Host '  - None found'
    }

    Write-Host ("Members  ({0})" -f $memberSummaries.Count)
    if ($memberSummaries.Count -gt 0) {
        $memberSummaries | ForEach-Object { Write-Host ("  - {0}" -f $_) }
    } else {
        Write-Host '  - None found'
    }

    if ($DryRun) {
        Write-Host '  [DryRun] Validation only. No owner/member changes were made.' -ForegroundColor Yellow
    }
}

Write-Host "`nCompleted cloud-only group evaluation." -ForegroundColor Green
