<#
.SYNOPSIS
    Lista las maquinas cuyo agente de Azure Arc esta roto (desconectado o con errores),
    causa raiz de fallos en Azure Update Manager (AUM).

.DESCRIPTION
    Recorre, via Azure Resource Graph, todas las suscripciones accesibles en el tenant actual
    (o las indicadas en -SubscriptionId) y combina dos señales:
      1. Maquinas Arc (Microsoft.HybridCompute/machines) cuyo estado de conexion no es
         'Connected', o que reportan errorDetails del propio agente.
      2. Resultados de evaluacion de parches de Update Manager (patchassessmentresources)
         con errorDetails no vacio, que normalmente son consecuencia directa de un agente
         Arc roto.

    Requiere los modulos Az.Accounts y Az.ResourceGraph, y acceso de lectura a las
    suscripciones objetivo.

.PARAMETER SubscriptionId
    Una o mas suscripciones a analizar. Por defecto, todas las accesibles en el tenant actual.

.PARAMETER OutputPath
    Ruta opcional para exportar el resultado combinado a CSV.

.EXAMPLE
    .\reviewarcaget.ps1

    Analiza todas las suscripciones accesibles y muestra en pantalla las maquinas con el
    agente Arc roto.

.EXAMPLE
    .\reviewarcaget.ps1 -SubscriptionId "sub-guid-1","sub-guid-2" -OutputPath .\arc_broken.csv

    Analiza solo esas dos suscripciones y exporta el resultado a CSV.

.NOTES
    El esquema de Resource Graph para Update Manager y Arc puede cambiar con el tiempo.
    Si algun campo aparece vacio de forma inesperada, valida el esquema actual con:
        Search-AzGraph -Query "resources | where type =~ 'microsoft.hybridcompute/machines' | limit 5" | ConvertTo-Json -Depth 5
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string[]]$SubscriptionId,

    [Parameter()]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

foreach ($mod in 'Az.Accounts', 'Az.ResourceGraph') {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        throw "El modulo $mod no esta instalado. Ejecuta: Install-Module Az -Scope CurrentUser"
    }
    Import-Module $mod -ErrorAction Stop
}

if (-not (Get-AzContext)) {
    throw "No hay sesion activa de Azure. Ejecuta Connect-AzAccount (o .\Connect-AzureTenant.ps1) primero."
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

# Señal 1: maquinas Arc realmente desconectadas o en error (el campo errorDetails viene
# poblado con un array vacio incluso en maquinas sanas, asi que no sirve como filtro).
$arcHealthQuery = @'
resources
| where type =~ 'microsoft.hybridcompute/machines'
| extend status = tostring(properties.status), lastStatusChange = todatetime(properties.lastStatusChange), agentVersion = tostring(properties.agentVersion)
| where status in~ ('Disconnected', 'Error', 'Expired')
| project name, resourceGroup, subscriptionId, source = 'ArcAgentStatus', status, lastStatusChange, agentVersion, errorMessage = strcat('Agente Arc con estado: ', status)
'@

# Señal 2: evaluaciones de Update Manager que fallaron de verdad (status 'Failed'; el mismo
# campo errorDetails tambien viene poblado en evaluaciones exitosas, por eso se filtra por status).
$aumErrorsQuery = @'
patchassessmentresources
| where type =~ 'microsoft.hybridcompute/machines/patchassessmentresults'
| extend prop = properties
| extend lastModified = todatetime(prop.lastModifiedDateTime), status = tostring(prop.status), errDetails = prop.errorDetails
| where status =~ 'Failed'
| project name = extract('machines/([^/]+)/', 1, tolower(id)), resourceGroup, subscriptionId, source = 'AUMAssessmentError', status, lastStatusChange = lastModified, agentVersion = '', errorMessage = tostring(errDetails.message)
'@

try {
    Write-Host "Consultando estado del agente Arc en todas las suscripciones..." -ForegroundColor Cyan
    $arcIssues = Invoke-PagedGraphQuery -Query $arcHealthQuery

    Write-Host "Consultando errores de evaluacion de Update Manager..." -ForegroundColor Cyan
    $aumIssues = Invoke-PagedGraphQuery -Query $aumErrorsQuery
}
catch {
    throw "La consulta a Resource Graph fallo (puede haber cambiado el esquema): $($_.Exception.Message)"
}

$combined = @($arcIssues) + @($aumIssues)

if (-not $combined) {
    Write-Host "`nNo se encontraron maquinas con el agente Arc roto ni errores de Update Manager." -ForegroundColor Green
    return
}

# Nombre de suscripcion legible en el reporte.
$subNames = @{}
foreach ($s in (Get-AzSubscription)) { $subNames[$s.Id] = $s.Name }

$report = $combined | ForEach-Object {
    [pscustomobject]@{
        VMName            = $_.name
        ResourceGroup     = $_.resourceGroup
        SubscriptionId    = $_.subscriptionId
        SubscriptionName  = $subNames[$_.subscriptionId]
        Origen            = $_.source
        Estado            = $_.status
        UltimoCambio      = $_.lastStatusChange
        Error             = $_.errorMessage
    }
} | Sort-Object SubscriptionName, VMName

Write-Host "`nSe encontraron $($report.Count) registro(s) de maquinas con el agente Arc roto / errores de AUM:" -ForegroundColor Yellow
$report | Format-Table VMName, SubscriptionName, Origen, Estado, Error -AutoSize -Wrap

if ($OutputPath) {
    $report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nExportado a $OutputPath" -ForegroundColor Green
}
