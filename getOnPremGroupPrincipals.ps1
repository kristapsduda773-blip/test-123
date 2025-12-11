#requires -Version 5.1
#requires -Modules ActiveDirectory

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter()]
    [string[]]$GroupNames,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [switch]$IncludeNestedMembers,

    [Parameter()]
    [switch]$ForceModuleReload
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

    if ($ForceModuleReload) {
        Remove-Module -Name $Name -ErrorAction SilentlyContinue
    }

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not installed. Install RSAT tools to proceed."
    }

    Import-Module $Name -ErrorAction Stop | Out-Null
}

function Resolve-PrincipalDetails {
    param([Parameter(Mandatory = $true)][Microsoft.ActiveDirectory.Management.ADObject]$DirectoryObject)

    $displayName = if ($DirectoryObject.DisplayName) { $DirectoryObject.DisplayName } else { $DirectoryObject.Name }
    $upnOrMail = $DirectoryObject.UserPrincipalName
    if (-not $upnOrMail -and $DirectoryObject.mail) { $upnOrMail = $DirectoryObject.mail }
    if (-not $upnOrMail -and $DirectoryObject.SamAccountName) { $upnOrMail = $DirectoryObject.SamAccountName }

    [pscustomobject]@{
        DisplayName       = $displayName
        UserPrincipalName = $upnOrMail
        ObjectId          = $DirectoryObject.ObjectGUID.Guid
        ObjectClass       = $DirectoryObject.ObjectClass
    }
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

function Get-ADGroupByIdentity {
    param([Parameter(Mandatory = $true)][string]$Identity)

    $properties = @('ManagedBy','mail','SamAccountName','DisplayName','ObjectGuid','DistinguishedName')

    try {
        return Get-ADGroup -Identity $Identity -Properties $properties -ErrorAction Stop
    } catch {
        # Fall through to filter search
    }

    $escaped = $Identity.Replace("'", "''")
    $result = Get-ADGroup -Filter "DisplayName -eq '$escaped'" -Properties $properties
    if (-not $result) {
        return $null
    }

    if ($result.Count -gt 1) {
        Write-Warning ("Multiple groups with DisplayName '{0}' detected. Using the first result." -f $Identity)
        return $result | Select-Object -First 1
    }

    return $result
}

Ensure-Module -Name ActiveDirectory

if (-not $GroupNames -or $GroupNames.Count -eq 0) {
    $GroupNames = $DefaultGroupNames
}

$GroupNames = $GroupNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique

if (-not $GroupNames) {
    throw 'No group names were supplied after filtering empty values.'
}

if (-not $OutputPath) {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path -Path (Get-Location) -ChildPath "OnPremGroupPrincipals_$timestamp.csv"
}

Write-Host "Processing {0} group names..." -f $GroupNames.Count -ForegroundColor Cyan

$principalRecords = @()
$processedGroups = 0
$missingGroups = 0
$membersRemovedTotal = 0
$memberRemovalFailures = 0

foreach ($groupName in $GroupNames) {
    Write-Host "`n------------------------------------------------------------"
    Write-Host ("Group lookup : {0}" -f $groupName) -ForegroundColor Cyan

    $group = Get-ADGroupByIdentity -Identity $groupName
    if (-not $group) {
        $missingGroups++
        Write-Warning ("Unable to locate group '{0}' in Active Directory." -f $groupName)
        continue
    }

    $processedGroups++
    $resolvedName = if ($group.DisplayName) { $group.DisplayName } else { $group.SamAccountName }
    $groupId = $group.ObjectGUID.Guid

    Write-Host ("DisplayName : {0}" -f $resolvedName)
    Write-Host ("ObjectGuid  : {0}" -f $groupId)
    $managedByValue = if ([string]::IsNullOrWhiteSpace($group.ManagedBy)) { '<not set>' } else { $group.ManagedBy }
    Write-Host ("ManagedBy   : {0}" -f $managedByValue)

    $owners = @()
    if ($group.ManagedBy) {
        try {
            $ownerObject = Get-ADObject -Identity $group.ManagedBy -Properties DisplayName, SamAccountName, mail, UserPrincipalName, ObjectClass, ObjectGuid
            $ownerDetails = Resolve-PrincipalDetails -DirectoryObject $ownerObject
            $owners = @($ownerDetails)
        } catch {
            Write-Warning ("Failed to resolve ManagedBy reference '{0}'. {1}" -f $group.ManagedBy, $_.Exception.Message)
        }
    }

    Write-Host ("Owners     : {0}" -f ($owners.Count))

    try {
        $memberParams = @{ Identity = $group.DistinguishedName; ErrorAction = 'Stop' }
        if ($IncludeNestedMembers) { $memberParams['Recursive'] = $true }
        $members = @(Get-ADGroupMember @memberParams)
    } catch {
        Write-Warning ("Failed to retrieve members for '{0}'. {1}" -f $resolvedName, $_.Exception.Message)
        continue
    }

    Write-Host ("Members    : {0}" -f $members.Count)

    foreach ($member in $members) {
        try {
            $memberObject = Get-ADObject -Identity $member.DistinguishedName -Properties DisplayName, SamAccountName, mail, UserPrincipalName, ObjectClass, ObjectGuid
            $memberDetails = Resolve-PrincipalDetails -DirectoryObject $memberObject

            $label = if ($memberDetails.UserPrincipalName) { $memberDetails.UserPrincipalName } else { $member.DistinguishedName }
            if (-not $PSCmdlet.ShouldProcess($label, "Remove from $resolvedName")) {
                continue
            }

            try {
                Remove-ADGroupMember -Identity $group.DistinguishedName -Members $member -Confirm:$false -ErrorAction Stop
                $membersRemovedTotal++
                Write-Host ("    Removed member: {0}" -f $label)
                $principalRecords += New-PrincipalRecord -GroupName $resolvedName -GroupId $groupId -Role 'Member' -Principal $memberDetails
            } catch {
                $memberRemovalFailures++
                Write-Warning ("    Failed to remove member {0}. {1}" -f $label, $_.Exception.Message)
            }
        } catch {
            Write-Warning ("Unable to resolve member '{0}'. {1}" -f $member.DistinguishedName, $_.Exception.Message)
        }
    }
}

Write-Host "`nSummary" -ForegroundColor Cyan
Write-Host ("  Groups processed : {0}" -f $processedGroups)
Write-Host ("  Groups missing   : {0}" -f $missingGroups)
Write-Host ("  Members removed  : {0}" -f $membersRemovedTotal)
Write-Host ("  Removal failures : {0}" -f $memberRemovalFailures)

if ($memberRemovalFailures -gt 0) {
    Write-Warning 'One or more member removals failed. Review warnings above.'
}

if (-not $principalRecords -or $principalRecords.Count -eq 0) {
    Write-Warning 'No members were removed from Active Directory.'
    return
}

Write-Host ("Exporting results to '{0}'" -f $OutputPath) -ForegroundColor Cyan
$principalRecords | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host 'Completed on-prem group principal enumeration.' -ForegroundColor Green
