[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId = $env:AZURE_TENANT_ID,

    [Parameter(Mandatory = $false)]
    [string]$ClientId = $env:AZURE_CLIENT_ID,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,

    [Parameter(Mandatory = $false)]
    [string]$NamesFile,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "group-source-report.csv",

    [Parameter(Mandatory = $false)]
    [int]$TimeoutSec = 15,

    [switch]$VerboseLogging
)

$ErrorActionPreference = "Stop"
if ($VerboseLogging.IsPresent) {
    $VerbosePreference = "Continue"
}

function Resolve-RequiredValue {
    param(
        [string]$Value,
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "Required parameter or environment variable '$Name' was not provided."
    }

    return $Value
}

$TenantId = Resolve-RequiredValue -Value $TenantId -Name "TenantId / AZURE_TENANT_ID"
$ClientId = Resolve-RequiredValue -Value $ClientId -Name "ClientId / AZURE_CLIENT_ID"
$ClientSecret = Resolve-RequiredValue -Value $ClientSecret -Name "ClientSecret / AZURE_CLIENT_SECRET"

$DefaultGroupNames = @(
    "DEV-ATD"
    "DEV-AVIS-cloud"
    "DEV-BDAS-cloud"
    "DEV-CAKEHR-cloud"
    "DEV-DELTAPV-cloud"
    "DEV-ESTAPIKS2-cloud"
    "DEV-Evo-Roads"
    "DEV-FITS-cloud"
    "DEV-INTRANET-cloud"
    "DEV-LAMBDAPV-cloud"
    "DEV-MANAGEMENT-cloud"
    "DEV-NILDA2-cloud"
    "DEV-OPVS–CargoRail"
    "DEV-PRESERVICA-cloud"
    "DEV-VADDVS"
    "Dots-Sales"
    "Product Group"
    "BW-DEV-ATD"
    "BW-DEV-Common"
    "BW-DEV-EvoRoads"
    "BW-DEV-Kappa"
    "BW-DEV-LDz-OPVS"
    "BW-DEV-SAGE-MAGS"
    "DEV-AIHEN-cloud"
    "DEV-AIROS"
    "DEV-Digitalizacija"
    "DEV-EXT-AKKA-LAA"
    "DEV-External-ATD-SMARTIN"
    "DEV-External-SMARTIN-DESIGN"
    "DEV-EXT-Estapiks2"
    "DEV-EXT-Fits"
    "DEV-EXT-KAMIS"
    "DEV-EXT-Peruza"
    "DEV-EXT-Preservica"
    "DEV-FITS"
    "DEV-IC-FITS"
    "SQL-00-PBI-Sync"
)

function Get-GroupNames {
    param(
        [string]$NamesFilePath
    )

    if ([string]::IsNullOrWhiteSpace($NamesFilePath)) {
        return $DefaultGroupNames
    }

    if (-not (Test-Path -Path $NamesFilePath -PathType Leaf)) {
        throw "Names file '$NamesFilePath' was not found."
    }

    $lines = Get-Content -Path $NamesFilePath -Encoding UTF8 |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith("#") }

    if (-not $lines) {
        throw "Names file '$NamesFilePath' did not contain any valid group names."
    }

    return $lines
}

function Get-GraphAccessToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }

    Write-Verbose "Requesting Microsoft Graph token..."
    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ErrorAction Stop
    if (-not $tokenResponse.access_token) {
        throw "Failed to obtain Microsoft Graph token."
    }

    return $tokenResponse.access_token
}

function Get-GraphGroup {
    param(
        [string]$AccessToken,
        [string]$DisplayName,
        [int]$TimeoutSec = 15
    )

    $escapedName = $DisplayName.Replace("'", "''")
    $filter = "displayName eq '$escapedName'"
    $encodedFilter = [Uri]::EscapeDataString($filter)
    $selectFields = "id,displayName,onPremisesSyncEnabled,onPremisesDomainName,onPremisesSecurityIdentifier"
    $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=$encodedFilter&`$select=$selectFields&`$count=true"

    $headers = @{
        "Authorization"    = "Bearer $AccessToken"
        "ConsistencyLevel" = "eventual"
    }

    Write-Verbose "Querying Microsoft Graph for '$DisplayName'"
    $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -TimeoutSec $TimeoutSec
    $matches = @($response.value)

    $exact = $matches | Where-Object { $_.displayName -and ($_.displayName.ToLower() -eq $DisplayName.ToLower()) } | Select-Object -First 1
    if ($exact) {
        return $exact
    }

    return $matches | Select-Object -First 1
}

function Get-SourceClassification {
    param(
        [object]$Group
    )

    if (-not $Group) {
        return @{
            Source = "not-found"
            Reason = "Group was not returned by Microsoft Graph"
        }
    }

    if ($Group.onPremisesSyncEnabled -eq $true) {
        return @{ Source = "on-prem"; Reason = "onPremisesSyncEnabled = true" }
    }

    if (-not [string]::IsNullOrWhiteSpace($Group.onPremisesDomainName)) {
        return @{ Source = "on-prem"; Reason = "onPremisesDomainName populated" }
    }

    if (-not [string]::IsNullOrWhiteSpace($Group.onPremisesSecurityIdentifier)) {
        return @{ Source = "on-prem"; Reason = "onPremisesSecurityIdentifier populated" }
    }

    if ($Group.onPremisesSyncEnabled -eq $false) {
        return @{ Source = "cloud"; Reason = "onPremisesSyncEnabled = false" }
    }

    return @{ Source = "cloud"; Reason = "No on-premises attributes detected" }
}

function Invoke-GroupClassification {
    param(
        [string[]]$Names,
        [string]$AccessToken,
        [int]$TimeoutSec
    )

    $results = foreach ($name in $Names) {
        try {
            $group = Get-GraphGroup -AccessToken $AccessToken -DisplayName $name -TimeoutSec $TimeoutSec
            $classification = Get-SourceClassification -Group $group
            [PSCustomObject]@{
                RequestedName = $name
                MatchedName   = $group.displayName
                Source        = $classification.Source
                Reason        = $classification.Reason
                GroupId       = $group.id
                Domain        = $group.onPremisesDomainName
            }
        }
        catch {
            [PSCustomObject]@{
                RequestedName = $name
                MatchedName   = $null
                Source        = "error"
                Reason        = $_.Exception.Message
                GroupId       = $null
                Domain        = $null
            }
        }
    }

    return $results
}

try {
    $groupNames = Get-GroupNames -NamesFilePath $NamesFile
    $token = Get-GraphAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    $records = Invoke-GroupClassification -Names $groupNames -AccessToken $token -TimeoutSec $TimeoutSec

    $sorted = $records | Sort-Object RequestedName
    $table = $sorted | Format-Table RequestedName, MatchedName, Source, Domain, Reason, GroupId -AutoSize | Out-String
    Write-Host $table

    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        $outputDirectory = Split-Path -Path $OutputPath -Parent
        if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory | Out-Null
        }
        $sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Saved CSV results to '$OutputPath'"
    }

    return $sorted
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
