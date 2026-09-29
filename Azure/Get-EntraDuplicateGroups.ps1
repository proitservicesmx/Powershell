<#
.SYNOPSIS
    Finds duplicate groups in Microsoft Entra ID and lists their Object IDs.

.DESCRIPTION
    Always prompts for the Tenant ID / domain so you connect to the right customer, signs in
    with Microsoft Graph scoped to that tenant, and asks you to confirm the tenant name before
    reading anything. Any existing Graph session is disconnected first so you don't end up
    mixing contexts between tenants.

    Groups are considered duplicates when they share the same value for -MatchOn
    (DisplayName by default). Comparison is case-insensitive and ignores leading/trailing
    whitespace.

    Results are shown on screen and exported to CSV (one row per group, with a DuplicateSet
    number so the groups that belong together are easy to filter).

    Requires the Microsoft.Graph.Authentication and Microsoft.Graph.Groups modules, and the
    Group.Read.All delegated permission (admin consent may be needed the first time).

.PARAMETER MatchOn
    Property used to detect duplicates: DisplayName (default), MailNickname or Mail.

.PARAMETER OutputPath
    Folder for the CSV export. Defaults to the current directory.

.PARAMETER NoExport
    Only show results on screen; don't write a CSV.

.EXAMPLE
    .\Get-EntraDuplicateGroups.ps1

    Prompts for the tenant, confirms it, then lists groups with duplicate display names.

.EXAMPLE
    .\Get-EntraDuplicateGroups.ps1 -MatchOn MailNickname -OutputPath C:\Reports

    Finds groups sharing the same mail nickname and exports the CSV to C:\Reports.
#>

[CmdletBinding()]
param(
    [ValidateSet('DisplayName', 'MailNickname', 'Mail')]
    [string]$MatchOn = 'DisplayName',

    [string]$OutputPath = (Get-Location).Path,

    [switch]$NoExport
)

$ErrorActionPreference = 'Stop'

foreach ($module in 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups') {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        throw "The $module module isn't installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
    }
    Import-Module $module -ErrorAction Stop
}

# Always ask for the tenant - this script is used across multiple customer environments.
$TenantId = Read-Host "Tenant ID or domain (e.g. contoso.onmicrosoft.com)"
if (-not $TenantId) {
    throw "A tenant ID or domain is required."
}

# Clear any existing session so contexts from a previous tenant don't linger.
if (Get-MgContext) {
    Disconnect-MgGraph | Out-Null
}

Write-Host "Signing in to tenant '$TenantId'..." -ForegroundColor Cyan
Connect-MgGraph -TenantId $TenantId -Scopes 'Group.Read.All' -NoWelcome

$context = Get-MgContext
$tenantName = $TenantId
try {
    $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName'
    if ($org.value) { $tenantName = $org.value[0].displayName }
}
catch {
    Write-Verbose "Couldn't read organization name: $_"
}

Write-Host "`nConnected as : $($context.Account)" -ForegroundColor Green
Write-Host "Tenant       : $tenantName ($($context.TenantId))" -ForegroundColor Green
$confirm = Read-Host "`nIs this the correct tenant? (Y/N)"
if ($confirm -notmatch '^(y|yes|s|si|sí)$') {
    Disconnect-MgGraph | Out-Null
    Write-Warning "Cancelled - disconnected from '$tenantName'."
    return
}

Write-Host "`nRetrieving all groups (this can take a while in large tenants)..." -ForegroundColor Cyan
$properties = 'Id', 'DisplayName', 'MailNickname', 'Mail', 'GroupTypes', 'SecurityEnabled',
              'MailEnabled', 'OnPremisesSyncEnabled', 'CreatedDateTime', 'Description'
$groups = Get-MgGroup -All -Property $properties | Select-Object $properties
Write-Host "Groups retrieved: $(@($groups).Count)"

function Get-GroupKind {
    param($Group)
    $kind = if ($Group.GroupTypes -contains 'Unified') { 'Microsoft 365' }
            elseif ($Group.SecurityEnabled -and $Group.MailEnabled) { 'Mail-enabled security' }
            elseif ($Group.SecurityEnabled) { 'Security' }
            else { 'Distribution' }
    if ($Group.GroupTypes -contains 'DynamicMembership') { $kind += ' (Dynamic)' }
    $kind
}

$duplicateSets = $groups |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_.$MatchOn) } |
    Group-Object -Property { $_.$MatchOn.Trim().ToLowerInvariant() } |
    Where-Object Count -gt 1 |
    Sort-Object Name

if (-not $duplicateSets) {
    Write-Host "`nNo duplicate groups found by $MatchOn in '$tenantName'." -ForegroundColor Green
    return
}

$setNumber = 0
$report = foreach ($set in $duplicateSets) {
    $setNumber++
    foreach ($g in ($set.Group | Sort-Object CreatedDateTime)) {
        [PSCustomObject]@{
            DuplicateSet   = $setNumber
            MatchValue     = $g.$MatchOn
            DisplayName    = $g.DisplayName
            ObjectId       = $g.Id
            GroupType      = Get-GroupKind $g
            MailNickname   = $g.MailNickname
            Mail           = $g.Mail
            SyncedFromOnPrem = [bool]$g.OnPremisesSyncEnabled
            CreatedDateTime  = $g.CreatedDateTime
            Description    = $g.Description
        }
    }
}

Write-Host ("`nFound {0} duplicate set(s) by {1}, {2} group(s) in total." -f $setNumber, $MatchOn, @($report).Count) -ForegroundColor Yellow
$report | Format-Table DuplicateSet, DisplayName, ObjectId, GroupType, SyncedFromOnPrem, CreatedDateTime -AutoSize

if (-not $NoExport) {
    $safeTenant = ($tenantName -replace '[\\/:*?"<>|\s]', '_')
    $file = Join-Path $OutputPath ("DuplicateGroups_{0}_{1}_{2}.csv" -f $safeTenant, $MatchOn, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $report | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported to: $file" -ForegroundColor Green
}
