<#
.SYNOPSIS
    Diagnoses network-level connectivity issues between this app server and a SQL Server instance.

.DESCRIPTION
    Runs a sequence of non-invasive network checks (no SQL login required) commonly used to
    isolate "can't connect to SQL Server" problems:

        1. DNS resolution of the target hostname
        2. ICMP ping (informational only - many SQL Servers block ICMP by policy)
        3. TCP port reachability test (default 1433, or a custom port)
        4. SQL Server Browser service lookup (UDP 1434) when a named instance is given,
           to discover the dynamic TCP port a named instance is actually listening on
        5. Optional traceroute, only run automatically when the TCP test fails, to show
           where along the path the connection is being dropped

    Each check prints a PASS / WARN / FAIL line plus a short remediation hint, and a final
    summary is shown. Optionally the full transcript is written to a log file.

.PARAMETER SqlServer
    Hostname, FQDN, or IP address of the target SQL Server. Required.

.PARAMETER Port
    TCP port to test. Defaults to 1433 (the standard SQL Server port). Ignored if -InstanceName
    is supplied and the Browser lookup succeeds (the discovered dynamic port is used instead).

.PARAMETER InstanceName
    Named instance (e.g. "SQLEXPRESS"). When supplied, the script queries the SQL Server Browser
    service (UDP 1434) on $SqlServer to resolve the instance's actual listening TCP port before
    running the port test.

.PARAMETER TimeoutSeconds
    Timeout, in seconds, used for the TCP and UDP probes. Defaults to 5.

.PARAMETER LogPath
    Optional path to a text file. When supplied, all output is also written there (via Start-Transcript).

.EXAMPLE
    .\Test-SqlConnectivity.ps1 -SqlServer sqlprod01.contoso.local

.EXAMPLE
    .\Test-SqlConnectivity.ps1 -SqlServer 10.20.30.40 -Port 1433 -LogPath C:\Temp\sql-conn-test.log

.EXAMPLE
    .\Test-SqlConnectivity.ps1 -SqlServer sqlprod01 -InstanceName SQLEXPRESS

.NOTES
    Run from the APPLICATION server (the machine having trouble reaching SQL Server), not from
    the SQL Server itself. Requires PowerShell 5.1+ (Test-NetConnection, Test-Connection are built in).
    No SQL credentials are used or required - this script only tests network reachability.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SqlServer,

    [int]$Port = 1433,

    [string]$InstanceName,

    [int]$TimeoutSeconds = 5,

    [string]$LogPath
)

$ErrorActionPreference = 'Continue'

if ($LogPath) {
    try {
        Start-Transcript -Path $LogPath -Append | Out-Null
        Write-Host "Logging to $LogPath" -ForegroundColor DarkGray
    } catch {
        Write-Warning "Could not start transcript at '$LogPath': $($_.Exception.Message)"
    }
}

$results = New-Object System.Collections.Generic.List[Object]

function Write-Result {
    param(
        [string]$Check,
        [ValidateSet('PASS','WARN','FAIL')][string]$Status,
        [string]$Detail
    )
    $color = switch ($Status) {
        'PASS' { 'Green' }
        'WARN' { 'Yellow' }
        'FAIL' { 'Red' }
    }
    Write-Host ("[{0,-4}] {1,-28} {2}" -f $Status, $Check, $Detail) -ForegroundColor $color
    $results.Add([PSCustomObject]@{
        Check  = $Check
        Status = $Status
        Detail = $Detail
    })
}

Write-Host ""
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host " SQL Server connectivity test" -ForegroundColor Cyan
Write-Host " Target      : $SqlServer" -ForegroundColor Cyan
Write-Host " Port        : $Port$(if ($InstanceName) { " (may be overridden by Browser lookup)" })" -ForegroundColor Cyan
if ($InstanceName) { Write-Host " Instance    : $InstanceName" -ForegroundColor Cyan }
Write-Host " Ran at      : $(Get-Date)" -ForegroundColor Cyan
Write-Host " From host   : $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host ""

# ----------------------------------------------------------------------------
# 1. DNS resolution
# ----------------------------------------------------------------------------
Write-Host "-- 1. DNS resolution --" -ForegroundColor White
$resolvedIPs = @()
try {
    $dnsResult = [System.Net.Dns]::GetHostAddresses($SqlServer)
    $resolvedIPs = $dnsResult | ForEach-Object { $_.IPAddressToString }
    if ($resolvedIPs.Count -gt 0) {
        Write-Result -Check 'DNS resolution' -Status 'PASS' -Detail ("Resolved to: " + ($resolvedIPs -join ', '))
    } else {
        Write-Result -Check 'DNS resolution' -Status 'FAIL' -Detail 'No addresses returned.'
    }
} catch {
    Write-Result -Check 'DNS resolution' -Status 'FAIL' -Detail "Could not resolve '$SqlServer': $($_.Exception.Message). Check DNS server config, /etc/hosts equivalent, or use the IP directly."
}
Write-Host ""

# ----------------------------------------------------------------------------
# 2. ICMP ping (informational - many SQL Servers/firewalls block ICMP)
# ----------------------------------------------------------------------------
Write-Host "-- 2. ICMP ping (informational only) --" -ForegroundColor White
try {
    $ping = Test-Connection -ComputerName $SqlServer -Count 2 -ErrorAction Stop
    $avg = ($ping | Measure-Object -Property ResponseTime -Average).Average
    Write-Result -Check 'ICMP ping' -Status 'PASS' -Detail "Host responded, avg ~$([math]::Round($avg,1)) ms."
} catch {
    Write-Result -Check 'ICMP ping' -Status 'WARN' -Detail "No ICMP reply (often normal - ICMP is commonly blocked by firewalls even when SQL traffic is allowed). Not a reliable indicator by itself."
}
Write-Host ""

# ----------------------------------------------------------------------------
# 3. SQL Server Browser lookup (only if a named instance was given)
# ----------------------------------------------------------------------------
if ($InstanceName) {
    Write-Host "-- 3. SQL Server Browser lookup (UDP 1434) for instance '$InstanceName' --" -ForegroundColor White
    $discoveredPort = $null
    try {
        $udpClient = New-Object System.Net.Sockets.UdpClient
        $udpClient.Client.ReceiveTimeout = $TimeoutSeconds * 1000
        $udpClient.Connect($SqlServer, 1434)
        $requestBytes = [byte[]]@(0x02)
        $udpClient.Send($requestBytes, $requestBytes.Length) | Out-Null

        $remoteEndpoint = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
        $responseBytes = $udpClient.Receive([ref]$remoteEndpoint)
        $response = [System.Text.Encoding]::ASCII.GetString($responseBytes)
        $udpClient.Close()

        # Response looks like: ServerName;<name>;InstanceName;<inst>;IsClustered;No;Version;<v>;tcp;<port>;;
        $entries = $response -split ';;' | Where-Object { $_ -match "InstanceName;$InstanceName;" -or $_ -match "(?i)InstanceName;$([regex]::Escape($InstanceName));" }
        if (-not $entries) { $entries = @($response) }

        foreach ($entry in $entries) {
            if ($entry -match "InstanceName;$([regex]::Escape($InstanceName));.*?tcp;(\d+)") {
                $discoveredPort = [int]$Matches[1]
                break
            }
        }
        # Fallback: single generic tcp;PORT match anywhere in the response
        if (-not $discoveredPort -and $response -match 'tcp;(\d+)') {
            $discoveredPort = [int]$Matches[1]
        }

        if ($discoveredPort) {
            Write-Result -Check 'SQL Browser lookup' -Status 'PASS' -Detail "Instance '$InstanceName' is listening on TCP port $discoveredPort. This port will be used for the port test."
            $Port = $discoveredPort
        } else {
            Write-Result -Check 'SQL Browser lookup' -Status 'WARN' -Detail "Browser service responded but the port for instance '$InstanceName' could not be parsed. Falling back to configured port $Port. Raw: $response"
        }
    } catch {
        Write-Result -Check 'SQL Browser lookup' -Status 'WARN' -Detail "Could not reach SQL Browser service on UDP 1434: $($_.Exception.Message). This is common if UDP 1434 is firewalled - falling back to configured port $Port. If the instance uses a static port, this is not a problem."
    }
    Write-Host ""
}

# ----------------------------------------------------------------------------
# 4. TCP port reachability (the real test)
# ----------------------------------------------------------------------------
$stepNum = if ($InstanceName) { 4 } else { 3 }
Write-Host "-- $stepNum. TCP port test ($SqlServer`:$Port) --" -ForegroundColor White
$tcpOk = $false
try {
    $tnc = Test-NetConnection -ComputerName $SqlServer -Port $Port -WarningAction SilentlyContinue -InformationLevel Detailed
    if ($tnc.TcpTestSucceeded) {
        $tcpOk = $true
        Write-Result -Check 'TCP port test' -Status 'PASS' -Detail "Connected to $SqlServer`:$Port successfully (via local address $($tnc.SourceAddress))."
    } else {
        Write-Result -Check 'TCP port test' -Status 'FAIL' -Detail "Could not open a TCP connection to $SqlServer`:$Port. Likely causes: SQL Server not listening on this port/protocol (check SQL Server Configuration Manager > TCP/IP enabled), a firewall on the app server, network path, or the SQL Server host blocking the app server's IP."
    }
} catch {
    Write-Result -Check 'TCP port test' -Status 'FAIL' -Detail "Test-NetConnection error: $($_.Exception.Message)"
}
Write-Host ""

# ----------------------------------------------------------------------------
# 5. Traceroute - only run if the TCP test failed, to help locate the break
# ----------------------------------------------------------------------------
if (-not $tcpOk) {
    $stepNum++
    Write-Host "-- $stepNum. Traceroute (TCP test failed, tracing path for diagnosis) --" -ForegroundColor White
    try {
        $trace = Test-NetConnection -ComputerName $SqlServer -TraceRoute -WarningAction SilentlyContinue
        if ($trace.TraceRoute) {
            Write-Host "Path to $SqlServer :" -ForegroundColor DarkGray
            $hopNum = 1
            foreach ($hop in $trace.TraceRoute) {
                Write-Host ("   {0,2}. {1}" -f $hopNum, $hop) -ForegroundColor DarkGray
                $hopNum++
            }
            Write-Result -Check 'Traceroute' -Status 'WARN' -Detail "See hop list above. The last responding hop before the path stops or times out is usually where the block is (firewall, ACL, routing)."
        } else {
            Write-Result -Check 'Traceroute' -Status 'WARN' -Detail "No trace data returned."
        }
    } catch {
        Write-Result -Check 'Traceroute' -Status 'WARN' -Detail "Traceroute failed: $($_.Exception.Message)"
    }
    Write-Host ""
}

# ----------------------------------------------------------------------------
# Summary
# ----------------------------------------------------------------------------
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host " Summary" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
$results | Format-Table -AutoSize | Out-String | Write-Host

$fails = $results | Where-Object { $_.Status -eq 'FAIL' }
$warns = $results | Where-Object { $_.Status -eq 'WARN' }

if ($fails.Count -eq 0) {
    Write-Host "Overall: no hard failures detected. If the app is still failing to connect, the issue is likely at the SQL/auth layer (login, permissions, TLS/encryption settings, or connection string) rather than the network." -ForegroundColor Green
} else {
    Write-Host "Overall: $($fails.Count) check(s) FAILED. Start with the TCP port test result above - that's almost always the actionable one." -ForegroundColor Red
    Write-Host ""
    Write-Host "Common next steps when the TCP port test fails:" -ForegroundColor Yellow
    Write-Host "  - On the SQL Server: confirm the SQL Server service is running and TCP/IP is enabled (SQL Server Configuration Manager > SQL Server Network Configuration > Protocols)."
    Write-Host "  - On the SQL Server: confirm Windows Firewall (or any host firewall) has an inbound rule allowing TCP $Port from the app server's IP."
    Write-Host "  - Between the servers: check any network firewall / security group / NSG / ACL along the path for a rule allowing $SqlServer`:$Port from this host ($env:COMPUTERNAME, $((Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike '169.*' -and $_.IPAddress -ne '127.0.0.1' } | Select-Object -First 1 -ExpandProperty IPAddress))) ."
    Write-Host "  - Confirm the SQL Server isn't restricted to specific IPs (e.g. via a firewall allow-list) that doesn't yet include this app server."
    Write-Host "  - If using a named instance, confirm UDP 1434 (SQL Browser) is reachable, or hardcode the correct static port in the app's connection string."
}

Write-Host ""

if ($LogPath) {
    try { Stop-Transcript | Out-Null } catch {}
}