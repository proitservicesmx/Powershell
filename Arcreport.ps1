<#
.SYNOPSIS
  Servidores Azure Arc con fallas de parcheo atribuibles al agente.
  Ejecuta 2 consultas de Azure Resource Graph y llena ArcPatch_FallasAgente.xlsx (y genera CSV).

.EXAMPLE
  .\Invoke-ArcPatchReport.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
  .\Invoke-ArcPatchReport.ps1 -TenantId "<tenant>" -SubscriptionId "<sub1>","<sub2>"
  .\Invoke-ArcPatchReport.ps1 -TenantId "<tenant>" -ForzarLogin
#>
param(
    [Parameter(Mandatory)] [string]   $TenantId,
    [string[]] $SubscriptionId,                       # opcional: si no se indica, consulta todo el tenant
    [string]   $Libro = ".\ArcPatch_FallasAgente.xlsx",
    [switch]   $SoloCsv,
    [switch]   $ForzarLogin                           # limpia la sesion guardada y pide login con device code
)
$ErrorActionPreference = "Stop"

# --- 1. Modulos ---
foreach ($m in @("Az.Accounts", "Az.ResourceGraph") + $(if (-not $SoloCsv) { "ImportExcel" })) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "Instalando $m..." -ForegroundColor Yellow
        Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m
}

# --- 2. Login ---
if ($ForzarLogin) {
    Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
    Clear-AzContext -Force -ErrorAction SilentlyContinue | Out-Null
}
if (-not (Get-AzContext) -or (Get-AzContext).Tenant.Id -ne $TenantId) {
    # Device code: no depende de WAM ni del navegador predeterminado
    Connect-AzAccount -TenantId $TenantId -UseDeviceAuthentication | Out-Null
}
$ctx = Get-AzContext
Write-Host "Conectado como $($ctx.Account.Id) en tenant $($ctx.Tenant.Id)" -ForegroundColor Green

# --- 3. Consultas ---
$kqlParcheo = @"
patchinstallationresources
| where type =~ "microsoft.hybridcompute/machines/patchinstallationresults"
| extend machineId = tolower(tostring(split(id, "/patchInstallationResults/")[0]))
| extend lastModified = todatetime(properties.lastModifiedDateTime)
| summarize arg_max(lastModified, *) by machineId          // solo la última corrida por servidor
| where properties.status in~ ("Failed", "CompletedWithWarnings")
| extend errorCode = tostring(properties.errorDetails.code),
         errorMsg  = tostring(properties.errorDetails.message),
         errorFull = tostring(properties.errorDetails)
| where errorFull has_any ("agent", "himds", "GuestAgent", "PatchExtension", "extension")
| join kind=leftouter (
    resources
    | where type =~ "microsoft.hybridcompute/machines"
    | project machineId = tolower(id), serverName = name, resourceGroup, subscriptionId,
              arcStatus = tostring(properties.status),
              agentVersion = tostring(properties.agentVersion),
              osType = tostring(properties.osType),
              lastStatusChange = todatetime(properties.lastStatusChange)
  ) on machineId
| project serverName, subscriptionId, resourceGroup, osType, arcStatus, agentVersion,
          lastStatusChange, patchStatus = tostring(properties.status),
          startedBy = tostring(properties.startedBy), lastModified, errorCode, errorMsg
| order by lastModified desc
"@

$kqlAgente = @"
resources
| where type =~ "microsoft.hybridcompute/machines"
| extend machineId = tolower(id), arcStatus = tostring(properties.status),
         agentVersion = tostring(properties.agentVersion), osType = tostring(properties.osType)
| join kind=leftouter (
    resources
    | where type =~ "microsoft.hybridcompute/machines/extensions"
    | where name in~ ("WindowsPatchExtension", "LinuxPatchExtension")
    | project machineId = tolower(tostring(split(id, "/extensions/")[0])),
              patchExt = name,
              extState = tostring(properties.provisioningState),
              extMsg = tostring(properties.instanceView.status.message)
  ) on machineId
| where arcStatus !~ "Connected" or extState =~ "Failed"
| project name, subscriptionId, resourceGroup, osType, arcStatus, agentVersion,
          lastStatusChange = todatetime(properties.lastStatusChange), patchExt, extState, extMsg
| order by arcStatus asc
"@

function Invoke-ArgQuery([string]$Query) {
    $all = @(); $skip = $null
    do {
        $p = @{ Query = $Query; First = 1000 }
        if ($SubscriptionId) { $p.Subscription = $SubscriptionId } else { $p.UseTenantScope = $true }
        if ($skip) { $p.SkipToken = $skip }
        $page = Search-AzGraph @p
        $all += $page; $skip = $page.SkipToken
    } while ($skip)
    return $all
}

Write-Host "Ejecutando consulta 1 (parcheo fallido)..." -ForegroundColor Cyan
$parcheo = @(Invoke-ArgQuery $kqlParcheo | Select-Object serverName, subscriptionId, resourceGroup, osType, arcStatus,
    agentVersion, lastStatusChange, patchStatus, startedBy, lastModified, errorCode, errorMsg)
Write-Host "Ejecutando consulta 2 (agente / extension)..." -ForegroundColor Cyan
$agente = @(Invoke-ArgQuery $kqlAgente | Select-Object name, subscriptionId, resourceGroup, osType, arcStatus,
    agentVersion, lastStatusChange, patchExt, extState, extMsg)

Write-Host ("Consulta 1: {0} servidores | Consulta 2: {1} servidores" -f $parcheo.Count, $agente.Count) -ForegroundColor Green

# --- 4. CSV (siempre) ---
$stamp = Get-Date -Format "yyyyMMdd_HHmm"
$parcheo | Export-Csv ".\arc_parcheo_fallido_agente_$stamp.csv" -NoTypeInformation -Encoding UTF8
$agente  | Export-Csv ".\arc_agente_extension_$stamp.csv"      -NoTypeInformation -Encoding UTF8

# --- 5. Excel ---
if (-not $SoloCsv) {
    if (-not (Test-Path $Libro)) { throw "No se encontro $Libro. Colocalo en la misma carpeta que el script o usa -Libro <ruta>." }
    try { [IO.File]::Open((Resolve-Path $Libro), 'Open', 'ReadWrite', 'None').Close() }
    catch { throw "El libro esta abierto en Excel. Cierralo y vuelve a ejecutar." }

    $respaldo = $Libro -replace '\.xlsx$', "_respaldo_$stamp.xlsx"
    Copy-Item $Libro $respaldo

    if ($parcheo.Count) { $parcheo | Export-Excel -Path $Libro -WorksheetName "Resultados Parcheo" -StartRow 5 -NoHeader }
    if ($agente.Count)  { $agente  | Export-Excel -Path $Libro -WorksheetName "Resultados Agente"  -StartRow 5 -NoHeader }
    Write-Host "Libro actualizado: $Libro (respaldo: $respaldo)" -ForegroundColor Green
}