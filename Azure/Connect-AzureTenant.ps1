<#
.SYNOPSIS
    Connects your interactive Azure account to a specific tenant using the Az PowerShell module.

.DESCRIPTION
    Prompts for (or accepts) a Tenant ID / domain, signs in with Connect-AzAccount scoped to
    that tenant, then lets you pick which subscription in that tenant to make active.
    Any existing Az session is disconnected first so you don't end up mixing contexts
    between tenants.

.PARAMETER TenantId
    The tenant's GUID or verified domain (e.g. contoso.onmicrosoft.com). If omitted, you'll be prompted.

.PARAMETER SubscriptionId
    Optional. Subscription (Id or Name) to select automatically after sign-in, skipping the picker.

.EXAMPLE
    .\Connect-AzureTenant.ps1

    Prompts for a Tenant ID, signs in, then shows a menu of subscriptions in that tenant.

.EXAMPLE
    .\Connect-AzureTenant.ps1 -TenantId contoso.onmicrosoft.com -SubscriptionId "Contoso Prod"

    Signs in directly to the given tenant and subscription, no prompts.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$TenantId,

    [Parameter(Position = 1)]
    [string]$SubscriptionId
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
    throw "The Az PowerShell module isn't installed. Run: Install-Module Az -Scope CurrentUser"
}
Import-Module Az.Accounts -ErrorAction Stop

if (-not $TenantId) {
    $TenantId = Read-Host "Tenant ID or domain (e.g. contoso.onmicrosoft.com)"
}
if (-not $TenantId) {
    throw "A tenant ID or domain is required."
}

# Clear any existing session so contexts from a previous tenant don't linger.
if (Get-AzContext) {
    Disconnect-AzAccount | Out-Null
}

Write-Host "Signing in to tenant '$TenantId'..." -ForegroundColor Cyan
$account = Connect-AzAccount -Tenant $TenantId

$subscriptions = Get-AzSubscription -TenantId $TenantId
if (-not $subscriptions) {
    Write-Warning "Signed in, but no subscriptions were found in this tenant for your account."
    return
}

$selected = $null
if ($SubscriptionId) {
    $selected = $subscriptions | Where-Object { $_.Id -eq $SubscriptionId -or $_.Name -eq $SubscriptionId }
    if (-not $selected) {
        Write-Warning "Subscription '$SubscriptionId' not found in this tenant; falling back to picker."
    }
}

if (-not $selected) {
    if (@($subscriptions).Count -eq 1) {
        $selected = $subscriptions[0]
    }
    else {
        Write-Host "`nSubscriptions available in this tenant:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $subscriptions.Count; $i++) {
            "[{0}] {1} ({2})" -f $i, $subscriptions[$i].Name, $subscriptions[$i].Id | Write-Host
        }
        $choice = Read-Host "`nSelect a subscription by number"
        if ($choice -notmatch '^\d+$' -or [int]$choice -ge $subscriptions.Count) {
            throw "Invalid selection: $choice"
        }
        $selected = $subscriptions[[int]$choice]
    }
}

Set-AzContext -Tenant $TenantId -Subscription $selected.Id | Out-Null

Write-Host "`nConnected." -ForegroundColor Green
Get-AzContext | Format-List Account, Tenant, Subscription, Environment
