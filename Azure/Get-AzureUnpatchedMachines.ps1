<#
.SYNOPSIS
    Reports Azure machines that are not up to date according to Azure Update Manager.

.DESCRIPTION
    Queries Azure Resource Graph for Update Manager patch assessment and installation results
    (the same data backing the Update Manager "Compliance" view in the portal) and flags
    Windows Server machines that:
      - have never been assessed or patched,
      - have a pending critical/security patch count greater than zero, or
      - haven't had a successful patch installation within -DaysThreshold days.

    Requires the Az.Accounts and Az.ResourceGraph modules, and read access to the target
    subscription(s).

.PARAMETER SubscriptionId
    One or more subscription IDs to scan. Defaults to every subscription accessible in the
    current tenant/context.

.PARAMETER DaysThreshold
    Number of days since the last successful patch installation before a machine is considered
    out of date. Default is 30.

.PARAMETER IncludeArcMachines
    Also include Arc-enabled (hybrid/on-prem) servers in addition to native Azure VMs.

.PARAMETER OutputPath
    Optional path to export the results as CSV.

.EXAMPLE
    .\Get-AzureUnpatchedMachines.ps1

    Scans all accessible subscriptions using the default 30-day threshold.

.EXAMPLE
    .\Get-AzureUnpatchedMachines.ps1 -SubscriptionId "sub-guid-1","sub-guid-2" -DaysThreshold 14 -IncludeArcMachines -OutputPath .\report.csv

    Scans two subscriptions, flags anything not patched in the last 14 days, includes Arc
    machines, and exports the results to a CSV file.

.NOTES
    Azure Update Manager's Resource Graph schema (tables `patchassessmentresources` and
    `patchinstallationresources`) can evolve. If results look wrong, run:
        Search-AzGraph -Query "patchassessmentresources | limit 5" | ConvertTo-Json -Depth 5
    to confirm the current property names against your tenant's data.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string[]]$SubscriptionId,

    [Parameter()]
    [int]$DaysThreshold = 30,

    [Parameter()]
    [switch]$IncludeArcMachines,

    [Parameter()]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

foreach ($mod in 'Az.Accounts', 'Az.ResourceGraph') {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        throw "The $mod PowerShell module isn't installed. Run: Install-Module Az -Scope CurrentUser"
    }
    Import-Module $mod -ErrorAction Stop
}

if (-not (Get-AzContext)) {
    throw "No active Azure session. Run Connect-AzAccount (or .\Connect-AzureTenant.ps1) first."
}

$machineTypes = @('microsoft.compute/virtualmachines')
if ($IncludeArcMachines) {
    $machineTypes += 'microsoft.hybridcompute/machines'
}
$typeFilter = ($machineTypes | ForEach-Object { "'$_'" }) -join ', '

# Every VM/Arc machine in scope, so machines with no assessment/install record at all still show up.
# Filtered to Windows only (Arc machines report osType via a different property than native Azure VMs).
$machinesQuery = @"
resources
| where type in ($typeFilter)
| extend osType = case(type =~ 'microsoft.compute/virtualmachines', tostring(properties.storageProfile.osDisk.osType), type =~ 'microsoft.hybridcompute/machines', tostring(properties.osType), '')
| where osType =~ 'windows'
| project machineId = tolower(id), name, resourceGroup, subscriptionId, osType, type
"@

# Latest assessment per machine.
$assessmentQuery = @'
patchassessmentresources
| where type =~ "microsoft.compute/virtualmachines/patchassessmentresults" or type =~ "microsoft.hybridcompute/machines/patchassessmentresults"
| extend machineId = tolower(tostring(split(id, "/patchAssessmentResults")[0]))
| summarize arg_max(todatetime(properties.lastModifiedDateTime), properties) by machineId
| project machineId, assessmentTime = todatetime(properties.lastModifiedDateTime), assessmentStatus = tostring(properties.status), criticalPending = toint(properties.availablePatchCountByClassification.critical), securityPending = toint(properties.availablePatchCountByClassification.security), rebootPending = tobool(properties.rebootPending)
'@

# Latest install per machine.
$installQuery = @'
patchinstallationresources
| where type =~ "microsoft.compute/virtualmachines/patchinstallationresults" or type =~ "microsoft.hybridcompute/machines/patchinstallationresults"
| extend machineId = tolower(tostring(split(id, "/patchInstallationResults")[0]))
| summarize arg_max(todatetime(properties.lastModifiedDateTime), properties) by machineId
| project machineId, installTime = todatetime(properties.lastModifiedDateTime), installStatus = tostring(properties.status), installedCount = toint(properties.installedPatchCount), pendingCount = toint(properties.pendingPatchCount), failedCount = toint(properties.failedPatchCount)
'@

function Invoke-PagedGraphQuery {
    param([Parameter(Mandatory)][string]$Query)

    $results = [System.Collections.Generic.List[object]]::new()
    $pageSize = 1000
    $skip = 0
    do {
        $params = @{ Query = $Query; First = $pageSize }
        if ($skip -gt 0) { $params['Skip'] = $skip }
        if ($SubscriptionId) { $params['Subscription'] = $SubscriptionId }
        $page = Search-AzGraph @params
        if ($page) { $results.AddRange([object[]]$page) }
        $skip += $pageSize
    } while ($page -and $page.Count -eq $pageSize)

    return $results
}

try {
    Write-Host "Querying Azure Resource Graph for machines..." -ForegroundColor Cyan
    $machines = Invoke-PagedGraphQuery -Query $machinesQuery

    Write-Host "Querying patch assessment results..." -ForegroundColor Cyan
    $assessments = Invoke-PagedGraphQuery -Query $assessmentQuery

    Write-Host "Querying patch installation results..." -ForegroundColor Cyan
    $installs = Invoke-PagedGraphQuery -Query $installQuery
}
catch {
    throw "Resource Graph query failed (schema may have changed): $($_.Exception.Message)"
}

$assessmentMap = @{}
foreach ($a in $assessments) { $assessmentMap[$a.machineId] = $a }

$installMap = @{}
foreach ($i in $installs) { $installMap[$i.machineId] = $i }

if (-not $machines) {
    Write-Warning "No machines found in scope (check subscription access and -IncludeArcMachines)."
    return
}

$cutoff = (Get-Date).ToUniversalTime().AddDays(-$DaysThreshold)

$report = foreach ($m in $machines) {
    $a = $assessmentMap[$m.machineId]
    $i = $installMap[$m.machineId]

    $status =
        if (-not $a -and -not $i) { 'NeverAssessed' }
        elseif ($a -and ($a.criticalPending -gt 0 -or $a.securityPending -gt 0)) { 'MissingCriticalOrSecurityPatches' }
        elseif (-not $i -or $i.installTime -lt $cutoff -or $i.installStatus -ne 'Succeeded') { 'NotRecentlyPatched' }
        else { 'Compliant' }

    if ($status -ne 'Compliant') {
        [pscustomobject]@{
            Name                = $m.name
            ResourceGroup       = $m.resourceGroup
            SubscriptionId      = $m.subscriptionId
            OSType              = $m.osType
            MachineType         = $m.type
            Status              = $status
            LastAssessmentTime  = $a.assessmentTime
            CriticalPending     = $a.criticalPending
            SecurityPending     = $a.securityPending
            RebootPending       = $a.rebootPending
            LastInstallTime     = $i.installTime
            LastInstallStatus   = $i.installStatus
            InstallPendingCount = $i.pendingCount
            InstallFailedCount  = $i.failedCount
        }
    }
}

$report = $report | Sort-Object Status, Name

if (-not $report) {
    Write-Host "`nAll scanned machines are compliant within the last $DaysThreshold day(s)." -ForegroundColor Green
    return
}

Write-Host "`nFound $($report.Count) machine(s) not fully patched (threshold: $DaysThreshold day(s)):" -ForegroundColor Yellow
$report | Format-Table Name, ResourceGroup, Status, LastInstallTime, CriticalPending, SecurityPending, RebootPending -AutoSize

if ($OutputPath) {
    $report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nExported to $OutputPath" -ForegroundColor Green
}
