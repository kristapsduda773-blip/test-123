#requires -Version 5.1
#requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, ImportExcel

[CmdletBinding()]
param(
    [Parameter()]
    [string[]]$GroupNames,

    [Parameter()]
    [string]$GraphScopes = 'Group.Read.All,GroupMember.Read.All',

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [switch]$ForceReconnect
)

$DefaultGroupNames = @(
    'DEV-ACTODLU','DEV-ADL','DEV-ADR','DEV-AIHEN','DEV-ALFAPV','DEV-ALFAPV-red',
    'DEV-ALFA-WSO2','DEV-ALTI','DEV-ALTUM','DEV-Apaksstacijas','DEV-APR','DEV-APUS',
    'DEV-Autostrade','DEV-AVIS','DEV-BDAS','DEV-Blokweb','DEV-CAKEHR','DEV-CSP-KLASIS',
    'DEV-DELTAPV','DEV-DVS','DEV-EIS','DEV-ELIS','DEV-EPS','DEV-ESKORT-EMCS',
    'DEV-ESTAPIKS','DEV-ESTAPIKS2','DEV-GAMMAPV','DEV-GPIS','DEV-IDMRII','DEV-INDTRA',
    'DEV-INTRANET','DEV-ITKC','DEV-LAMBDA','DEV-LAMBDAPV','DEV-LGAP','DEV-LNBACTO',
    'DEV-LP','DEV-LRPV','DEV-LSR','DEV-LVC','DEV-MANAGEMENT','DEV-MobitouchID',
    'DEV-NILDA','DEV-NILDA2','DEV-OMEGAPV','DEV-PERUZA-SIKPAKAS','DEV-PRESERVICA',
    'DEV-REID','DEV-RigasAcs','DEV-RSIIS','DEV-SAMS','DEV-SIGMAPV','DEV-SKUS','DEV-SNIP',
    'DEV-VirtualaisBirojs','DEV-VISTA2','DEV-VKL','DEV-VPiepirkums','DEV-VPPP',
    'DEV-VR-ABC-varti','DEV-WSO2','DEV-ZETAPV','PV-Zeta'
)

function Ensure-Module {
    param([Parameter(Mandatory = $true)][string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not installed. Install with: Install-Module $Name"
    }

    Import-Module $Name -ErrorAction Stop | Out-Null
}

foreach ($module in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'ImportExcel')) {
    Ensure-Module -Name $module
}

if (-not $GroupNames -or $GroupNames.Count -eq 0) {
    $GroupNames = $DefaultGroupNames
}

$GroupNames = $GroupNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique

if (-not $GroupNames) {
    throw 'No group names were provided.'
}

if (-not $OutputPath) {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path -Path (Get-Location) -ChildPath "IntuneGroupSnapshot_$timestamp.xlsx"
}

if ($ForceReconnect -or -not (Get-MgContext)) {
    Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
    $scopes = $GraphScopes.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    Connect-MgGraph -Scopes $scopes -NoWelcome
}

if (-not (Get-MgContext)) {
    throw 'Unable to acquire a Microsoft Graph context. Verify credentials and consent.'
}

$groupRecords = @()
$memberRecords = @()
$ownerRecords = @()

foreach ($groupName in $GroupNames) {
    Write-Host "`n------------------------------------------------------------"
    Write-Host ("Looking up '{0}'" -f $groupName) -ForegroundColor Cyan

    try {
        $group = Get-MgGroup -Filter "displayName eq '$($groupName.Replace("'","''"))'" -ConsistencyLevel eventual -CountVariable ignored
    } catch {
        Write-Warning ("  Graph query failed for {0}. {1}" -f $groupName, $_.Exception.Message)
        continue
    }

    if (-not $group) {
        Write-Warning ("  No Entra ID group named '{0}' was found." -f $groupName)
        continue
    }

    if ($group.Count -gt 1) {
        Write-Warning ("  Multiple groups named '{0}' detected. Using the first result." -f $groupName)
    }

    $groupObject = $group | Select-Object -First 1
    $groupRecords += [pscustomobject]@{
        DisplayName = $groupObject.DisplayName
        GroupId     = $groupObject.Id
        Mail        = $groupObject.Mail
        MailNickname= $groupObject.MailNickname
        Description = $groupObject.Description
        GroupTypes  = ($groupObject.GroupTypes -join ', ')
        SecurityEnabled = $groupObject.SecurityEnabled
        MailEnabled = $groupObject.MailEnabled
        OnPremisesSyncEnabled = $groupObject.OnPremisesSyncEnabled
    }

    Write-Host ("  GroupId : {0}" -f $groupObject.Id)
    Write-Host ("  Sync    : {0}" -f $groupObject.OnPremisesSyncEnabled)

    try {
        $owners = @(Get-MgGroupOwner -GroupId $groupObject.Id -All -ErrorAction Stop)
    } catch {
        Write-Warning ("  Failed to fetch owners. {0}" -f $_.Exception.Message)
        $owners = @()
    }

    if ($owners.Count -eq 0) {
        Write-Host '  Owners : none' -ForegroundColor Yellow
    } else {
        Write-Host ("  Owners : {0}" -f $owners.Count)
        foreach ($owner in $owners) {
            $ownerRecords += [pscustomobject]@{
                GroupDisplayName = $groupObject.DisplayName
                GroupId          = $groupObject.Id
                OwnerDisplayName = $owner.AdditionalProperties['displayName']
                OwnerUPN         = $owner.AdditionalProperties['userPrincipalName']
                OwnerObjectId    = $owner.Id
                OwnerType        = ($owner.AdditionalProperties['@odata.type'] -replace '#microsoft.graph.', '')
            }
        }
    }

    try {
        $members = @(Get-MgGroupMember -GroupId $groupObject.Id -All -ErrorAction Stop)
    } catch {
        Write-Warning ("  Failed to fetch members. {0}" -f $_.Exception.Message)
        $members = @()
    }

    if ($members.Count -eq 0) {
        Write-Host '  Members: none' -ForegroundColor Yellow
    } else {
        Write-Host ("  Members: {0}" -f $members.Count)
        foreach ($member in $members) {
            $memberRecords += [pscustomobject]@{
                GroupDisplayName = $groupObject.DisplayName
                GroupId          = $groupObject.Id
                MemberDisplayName= $member.AdditionalProperties['displayName']
                MemberUPN        = $member.AdditionalProperties['userPrincipalName']
                MemberObjectId   = $member.Id
                MemberType       = ($member.AdditionalProperties['@odata.type'] -replace '#microsoft.graph.', '')
            }
        }
    }
}

if (-not $groupRecords) {
    Write-Warning 'No matching groups were found in Entra ID.'
    return
}

Write-Host ("Exporting snapshot to '{0}'" -f $OutputPath) -ForegroundColor Cyan
$exportData = @{
    'Groups'  = $groupRecords
    'Members' = $memberRecords
    'Owners'  = $ownerRecords
}

foreach ($sheetName in $exportData.Keys) {
    $data = $exportData[$sheetName]
    if ($data -and $data.Count -gt 0) {
        $data | Export-Excel -Path $OutputPath -WorksheetName $sheetName -AutoSize -FreezeTopRow -TableName $sheetName -BoldTopRow -AutoFilter
    } else {
        # create an empty sheet to document absence
        $null = @([pscustomobject]@{ Note = 'No data returned' }) | Export-Excel -Path $OutputPath -WorksheetName $sheetName -AutoSize -FreezeTopRow -TableName $sheetName -BoldTopRow -AutoFilter
    }
}

Write-Host 'Completed Intune/Entra group snapshot export.' -ForegroundColor Green
