<#
.SYNOPSIS
    Reports the VM size (SKU) for a list of Azure virtual machines.

.DESCRIPTION
    Looks up the compute SKU (VM size, e.g. Standard_D2s_v3) for up to 10 VMs entered
    manually, or for any number of VMs supplied via a CSV file.

    Uses Azure Resource Graph, so VMs are found regardless of which subscription they live
    in (any subscription your account can access, or a specific set via -SubscriptionId) -
    no need to know or switch to the right subscription beforehand.

    CSV format (header required):
        Name,ResourceGroupName
        vm-web-01,rg-prod
        vm-web-02,rg-prod
        vm-sql-01,

    Only "Name" is mandatory. "ResourceGroupName" is optional and only used to disambiguate
    when the same VM name exists in more than one resource group/subscription.

    Requires the Az.Accounts and Az.ResourceGraph modules, and an active Connect-AzAccount
    session (see .\Connect-AzureTenant.ps1).

.PARAMETER VMName
    Up to 10 VM names to look up manually.

.PARAMETER ResourceGroupName
    Optional. Resource group(s) matching -VMName by position, used to disambiguate when a VM
    name exists in more than one resource group. Supply one value to apply it to every VM, or
    one value per VM.

.PARAMETER CsvPath
    Path to a CSV file with Name and (optional) ResourceGroupName columns. Use this instead of
    -VMName for larger lists.

.PARAMETER SubscriptionId
    Optional. Limit the search to these subscription IDs instead of every subscription
    accessible in the current tenant/context.

.PARAMETER OutputPath
    Optional path to export the results as CSV.

.EXAMPLE
    .\skureview.ps1 -VMName vm-web-01, vm-web-02

    Looks up the SKU for two VMs, searching every subscription you can access.

.EXAMPLE
    .\skureview.ps1 -VMName vm-web-01, vm-sql-01 -ResourceGroupName rg-prod, rg-data

    Looks up each VM, using the resource group to disambiguate if the name isn't unique.

.EXAMPLE
    .\skureview.ps1 -CsvPath .\vms.csv -OutputPath .\sku-report.csv

    Looks up every VM listed in vms.csv across all accessible subscriptions and exports the
    results to sku-report.csv.
#>

[CmdletBinding(DefaultParameterSetName = 'Manual')]
param(
    [Parameter(ParameterSetName = 'Manual', Position = 0)]
    [ValidateCount(1, 10)]
    [string[]]$VMName,

    [Parameter(ParameterSetName = 'Manual')]
    [string[]]$ResourceGroupName,

    [Parameter(ParameterSetName = 'Csv', Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter()]
    [string[]]$SubscriptionId,

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

# Build the list of VMs to resolve, regardless of which parameter set was used.
$targets = [System.Collections.Generic.List[pscustomobject]]::new()

if ($PSCmdlet.ParameterSetName -eq 'Csv') {
    $rows = Import-Csv -Path $CsvPath
    if (-not $rows) {
        throw "No rows found in '$CsvPath'."
    }
    foreach ($row in $rows) {
        $name = if ($row.Name) { $row.Name } else { $row.VMName }
        if (-not $name) {
            Write-Warning "Skipping a CSV row with no Name/VMName value."
            continue
        }
        $targets.Add([pscustomobject]@{
            Name              = $name.Trim()
            ResourceGroupName = if ($row.ResourceGroupName) { $row.ResourceGroupName.Trim() } else { $null }
        })
    }
}
else {
    if (-not $VMName) {
        throw "Provide -VMName (up to 10 VMs) or -CsvPath."
    }
    if ($ResourceGroupName -and $ResourceGroupName.Count -ne 1 -and $ResourceGroupName.Count -ne $VMName.Count) {
        throw "-ResourceGroupName must have either 1 value (applied to all VMs) or one value per -VMName."
    }
    for ($i = 0; $i -lt $VMName.Count; $i++) {
        $rg = $null
        if ($ResourceGroupName) {
            $rg = if ($ResourceGroupName.Count -eq 1) { $ResourceGroupName[0] } else { $ResourceGroupName[$i] }
        }
        $targets.Add([pscustomobject]@{
            Name              = $VMName[$i]
            ResourceGroupName = $rg
        })
    }
}

if (-not $targets) {
    throw "No VMs to look up."
}

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

$namesForQuery = ($targets.Name | Select-Object -Unique | ForEach-Object { "'$($_.Replace("'", "''"))'" }) -join ', '

# Covers both native Azure VMs and Arc-enabled (on-prem/hybrid) servers, since both show up as
# "machines" in this tenant and a name alone doesn't tell you which one you're dealing with.
$vmQuery = @"
resources
| where type in ('microsoft.compute/virtualmachines', 'microsoft.hybridcompute/machines')
| where name in~ ($namesForQuery)
| extend vmSize = case(type =~ 'microsoft.compute/virtualmachines', tostring(properties.hardwareProfile.vmSize), tostring(properties.detectedProperties.vmSize))
| extend osType = case(type =~ 'microsoft.compute/virtualmachines', tostring(properties.storageProfile.osDisk.osType), tostring(properties.osType))
| project name, resourceGroup, subscriptionId, location, machineType = type, vmSize, osType
"@

Write-Host "Querying Azure Resource Graph across accessible subscriptions..." -ForegroundColor Cyan
try {
    $found = Invoke-PagedGraphQuery -Query $vmQuery
}
catch {
    throw "Resource Graph query failed: $($_.Exception.Message)"
}

# If nothing at all came back, fall back to a fuzzy search so the user can see near-matches
# (typos, domain suffixes, trailing whitespace) instead of a bare NotFound.
if (-not $found) {
    $containsClauses = ($targets.Name | Select-Object -Unique | ForEach-Object { "name contains '$($_.Replace("'", "''"))'" }) -join ' or '
    $fuzzyQuery = @"
resources
| where type in ('microsoft.compute/virtualmachines', 'microsoft.hybridcompute/machines')
| where $containsClauses
| project name, resourceGroup, subscriptionId, type
"@
    try {
        $fuzzyMatches = Invoke-PagedGraphQuery -Query $fuzzyQuery
    }
    catch { $fuzzyMatches = @() }

    if ($fuzzyMatches) {
        Write-Warning "No exact name match, but found similarly-named resource(s):"
        $fuzzyMatches | Format-Table name, resourceGroup, subscriptionId, type -AutoSize | Out-Host
    }
    else {
        Write-Warning "No matching resource (VM or Arc machine) found anywhere in the subscriptions Resource Graph can see for your account. Confirm you have at least Reader access to the subscription(s) hosting these machines."
    }
}

$report = [System.Collections.Generic.List[pscustomobject]]::new()

foreach ($t in $targets) {
    $matches = @($found | Where-Object { $_.name -eq $t.Name })
    if ($t.ResourceGroupName -and $matches.Count -gt 1) {
        $narrowed = @($matches | Where-Object { $_.resourceGroup -eq $t.ResourceGroupName })
        if ($narrowed) { $matches = $narrowed }
    }

    if (-not $matches) {
        $report.Add([pscustomobject]@{
            Name              = $t.Name
            ResourceGroupName = $t.ResourceGroupName
            SubscriptionId    = $null
            MachineType       = $null
            VMSize            = $null
            Location          = $null
            OSType            = $null
            Status            = 'NotFound'
        })
        continue
    }

    if ($matches.Count -gt 1) {
        Write-Warning "'$($t.Name)' matches $($matches.Count) VMs across subscriptions/resource groups; specify -ResourceGroupName to disambiguate. Reporting all matches."
    }

    foreach ($vm in $matches) {
        $report.Add([pscustomobject]@{
            Name              = $vm.name
            ResourceGroupName = $vm.resourceGroup
            SubscriptionId    = $vm.subscriptionId
            MachineType       = $vm.machineType
            VMSize            = $vm.vmSize
            Location          = $vm.location
            OSType            = $vm.osType
            Status            = 'Found'
        })
    }
}

$report = $report | Sort-Object Status, Name

Write-Host "`nSKU lookup results ($($report.Count) row(s)):" -ForegroundColor Cyan
$report | Format-Table Name, ResourceGroupName, SubscriptionId, MachineType, VMSize, Location, OSType, Status -AutoSize

if ($OutputPath) {
    $report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nExported to $OutputPath" -ForegroundColor Green
}
