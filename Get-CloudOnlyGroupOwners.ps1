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

# Top-level safeguard. Keep true while validating.
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

foreach ($module in @('ImportExcel', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')) {
    Ensure-Module -Name $module
}

if ($ForceReconnect -or -not (Get-MgContext)) {
    Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
    Connect-MgGraph -Scopes @('Group.Read.All') -NoWelcome
}

if (-not (Get-MgContext)) {
    throw 'Unable to acquire a Microsoft Graph context. Verify your credentials and consent.'
}

Write-Host ("Dry run mode is {0}" -f $(if ($DryRun) { 'ON' } else { 'OFF' })) -ForegroundColor Yellow

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
    Write-Warning 'The Excel file contained no rows to evaluate.'
    return
}

$cloudGroups = $groupRows | Where-Object { $_.Source -eq 'CloudOnly' }

if (-not $cloudGroups) {
    Write-Host "No groups with Source = 'CloudOnly' were found in the provided file."
    return
}

$foundCount = 0
$missingCount = 0

foreach ($groupRecord in $cloudGroups) {
    if (-not $groupRecord.Id) {
        Write-Warning ("Skipping row with DisplayName '{0}' because Id is missing." -f $groupRecord.DisplayName)
        continue
    }

    $groupName = if ([string]::IsNullOrWhiteSpace([string]$groupRecord.DisplayName)) { '<unknown>' } else { [string]$groupRecord.DisplayName }

    Write-Host "`n------------------------------------------------------------"
    Write-Host ("DisplayName : {0}" -f $groupName) -ForegroundColor Cyan
    Write-Host ("Group Id    : {0}" -f $groupRecord.Id)
    Write-Host ("Source      : {0}" -f ($groupRecord.Source -as [string]))
    Write-Host ("GroupType   : {0}" -f ([string]$groupRecord.GroupType ?? '<not specified>'))

    try {
        $groupObject = Get-MgGroup -GroupId $groupRecord.Id -ErrorAction Stop
        $foundCount++
        Write-Host 'Status      : Found in Entra ID' -ForegroundColor Green
        if ($DryRun) {
            Write-Host '  [DryRun] Validation only. No changes were made.' -ForegroundColor Yellow
        }
    } catch {
        $missingCount++
        Write-Warning ("Status      : NOT found in Entra ID. {0}" -f $_.Exception.Message)
    }
}

Write-Host "`nSummary" -ForegroundColor Cyan
Write-Host ("  Found groups   : {0}" -f $foundCount)
Write-Host ("  Missing groups : {0}" -f $missingCount)
Write-Host "Completed cloud-only group lookup." -ForegroundColor Green
