<#
.SYNOPSIS
    Validates forward and reverse DNS resolution for a list of servers.

.DESCRIPTION
    Reads hostnames and/or IP addresses from a text file (one per line; blank lines and lines
    starting with # are ignored) and checks that forward (A/AAAA) and reverse (PTR) DNS records
    round-trip back to each other:
      - For a hostname entry: resolves it to its IP address(es), then looks up the PTR record
        for each IP and checks that at least one PTR name matches the original hostname.
      - For an IP address entry: looks up its PTR record, then resolves that name back to an
        IP address and checks it includes the original IP.

    A mismatch (both lookups succeed but don't point back to each other) usually means a stale
    or incorrect PTR record. A failure on either side is reported separately so you can tell a
    missing PTR record apart from a missing A/AAAA record.

    If you have multiple reverse lookup zones split across different DNS servers (common with
    several sites/subnets), a lookup can fail simply because the server you asked doesn't hold
    that particular zone - not because the record is actually missing. Pass -DnsServer with
    more than one server and each lookup tries them in order until one answers; the report
    shows which server actually resolved each entry (AnsweredBy) so you can tell "genuinely
    missing" apart from "exists, just not on the server I queried".

    Querying an explicit -DnsServer bypasses the client-side DNS suffix search list that tools
    like ping.exe apply automatically to unqualified short names. So a short name (no dots) is
    also tried with your machine's DNS suffixes appended, the same way ping would resolve it -
    otherwise a short name that works fine in ping could wrongly show up as ForwardFailed here.

    On top of DNS, each entry is also pinged (ICMP) by its name and by its IP address, so you
    can tell "DNS resolves fine but the host doesn't actually answer" apart from a DNS problem.

.PARAMETER InputPath
    Path to the server list text file. Defaults to .\serverlist.txt.

.PARAMETER DnsServer
    Optional. One or more DNS servers to query instead of the system's configured resolver(s).
    When more than one is given, each lookup tries them in order and stops at the first one
    that answers - useful when different reverse zones live on different servers.

.PARAMETER SkipConnectivityTest
    Skip the ping-by-name / ping-by-IP checks and only validate DNS. Useful when ICMP is
    blocked on the network you're running from.

.PARAMETER PingTimeoutMs
    Timeout in milliseconds for each ping. Default is 1000.

.PARAMETER OutputPath
    Optional path to export the results as CSV.

.EXAMPLE
    .\Test-DnsForwardReverse.ps1

    Validates every entry in .\serverlist.txt using the system's default DNS resolver, and
    pings each one by name and by IP.

.EXAMPLE
    .\Test-DnsForwardReverse.ps1 -InputPath .\servers.txt -DnsServer 10.0.0.10 -OutputPath .\dns-report.csv

    Validates entries in servers.txt against a specific DNS server and exports the results.

.EXAMPLE
    .\Test-DnsForwardReverse.ps1 -DnsServer dc1.contoso.com, dc2.contoso.com, dc-siteb.contoso.com

    Tries each DNS server in order per lookup - handy when different sites/subnets have their
    own reverse lookup zones that aren't all hosted on (or forwarded by) the same server.

.EXAMPLE
    .\Test-DnsForwardReverse.ps1 -SkipConnectivityTest

    Only validates DNS, without pinging the hosts (e.g. when ICMP is blocked).
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$InputPath = '.\serverlist.txt',

    [Parameter()]
    [string[]]$DnsServer,

    [Parameter()]
    [switch]$SkipConnectivityTest,

    [Parameter()]
    [int]$PingTimeoutMs = 1000,

    [Parameter()]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) {
    throw "Resolve-DnsName isn't available. This script requires the DnsClient module (built into Windows)."
}

$entries = Get-Content -Path $InputPath |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -and -not $_.StartsWith('#') } |
    Select-Object -Unique

if (-not $entries) {
    throw "No entries found in '$InputPath'."
}

# Servers to try per lookup, in order. An empty entry means "system default resolver".
$serversToTry = if ($DnsServer) { $DnsServer } else { @($null) }

# Resolve-DnsName -Server queries that server directly for the literal name given - it skips
# the client-side DNS suffix search list / devolution that ping.exe (and Resolve-DnsName with
# no -Server) applies automatically. So an unqualified short name that ping resolves fine can
# come back empty here unless we append the candidate suffixes ourselves.
$dnsSuffixes = [System.Collections.Generic.List[string]]::new()
try {
    $primarySuffix = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName
    if ($primarySuffix) { $dnsSuffixes.Add($primarySuffix) }
}
catch { }
try {
    foreach ($s in (Get-DnsClientGlobalSetting -ErrorAction Stop).SuffixSearchList) {
        if ($s -and $dnsSuffixes -notcontains $s) { $dnsSuffixes.Add($s) }
    }
}
catch { }

function Get-NameCandidates {
    param([Parameter(Mandatory)][string]$Name)

    $candidates = [System.Collections.Generic.List[string]]::new()
    $candidates.Add($Name)
    if ($Name -notmatch '\.') {
        foreach ($suffix in $dnsSuffixes) { $candidates.Add("$Name.$suffix") }
    }
    return $candidates
}

function Resolve-Forward {
    param([Parameter(Mandatory)][string]$Name)

    foreach ($candidate in (Get-NameCandidates -Name $Name)) {
        foreach ($srv in $serversToTry) {
            $params = @{ Name = $candidate; Type = 'A_AAAA'; ErrorAction = 'Stop' }
            if ($srv) { $params['Server'] = $srv }
            try {
                $records = @(Resolve-DnsName @params | Where-Object { $_.Type -in 'A', 'AAAA' })
                if ($records) {
                    # Use the answer's own owner name, not the candidate we queried with - the
                    # DNS client can still apply devolution/suffixing itself even against an
                    # explicit -Server, so the record's Name is the true FQDN, unlike $candidate
                    # which may just be the bare short name we happened to query.
                    return [pscustomobject]@{
                        Values       = @($records.IPAddress)
                        Server       = $(if ($srv) { $srv } else { 'Default' })
                        ResolvedName = $records[0].Name.TrimEnd('.')
                    }
                }
            }
            catch { }
        }
    }
    return [pscustomobject]@{ Values = @(); Server = $null; ResolvedName = $null }
}

function Resolve-Reverse {
    param([Parameter(Mandatory)][string]$IPAddress)

    foreach ($srv in $serversToTry) {
        $params = @{ Name = $IPAddress; Type = 'PTR'; ErrorAction = 'Stop' }
        if ($srv) { $params['Server'] = $srv }
        try {
            $records = @(Resolve-DnsName @params | Where-Object { $_.Type -eq 'PTR' })
            # A resolver/firewall that blocks or fails a PTR query can return a bogus "answer" that
            # just echoes the reverse-zone query name back (e.g. 8.8.8.8.in-addr.arpa) instead of a
            # real hostname or a proper NXDOMAIN. Treat that as no PTR record and keep trying the
            # next server rather than a match.
            $names = @($records.NameHost.TrimEnd('.') | Where-Object { $_ -notmatch '\.(in-addr|ip6)\.arpa$' })
            if ($names) {
                return [pscustomobject]@{ Values = $names; Server = $(if ($srv) { $srv } else { 'Default' }) }
            }
        }
        catch { }
    }
    return [pscustomobject]@{ Values = @(); Server = $null }
}

function Get-DnsDomain {
    param([string]$Fqdn)

    if ($Fqdn -match '^[^.]+\.(.+)$') { return $Matches[1] }
    return $null
}

function Test-NameRoundTrip {
    param(
        [Parameter(Mandatory)][string]$Entry,
        [string[]]$Candidates
    )

    if (-not $Candidates) { return $false }
    if ($Entry -match '\.') {
        # Entry was already given as an FQDN, so the domain has to match too.
        return @($Candidates) -contains $Entry
    }
    # Entry was a short name - a PTR/forward result normally comes back fully qualified, so
    # only compare the host label; requiring an exact match against the FQDN would flag every
    # correctly-configured short-name entry as a false "Mismatch".
    $labels = $Candidates | ForEach-Object { ($_ -split '\.')[0] }
    return @($labels) -contains $Entry
}

function Test-HostResponse {
    param(
        [string]$Target,
        [int]$TimeoutMs
    )

    if (-not $Target) { return 'N/A' }
    try {
        $ping = [System.Net.NetworkInformation.Ping]::new()
        $reply = $ping.Send($Target, $TimeoutMs)
        if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) { 'OK' } else { 'Failed' }
    }
    catch {
        'Failed'
    }
}

$report = [System.Collections.Generic.List[pscustomobject]]::new()
$total = $entries.Count
$i = 0

foreach ($entry in $entries) {
    $i++
    Write-Host "[$i/$total] Testing '$entry'..." -ForegroundColor Cyan

    $isIp = [System.Net.IPAddress]::TryParse($entry, [ref]$null)

    if ($isIp) {
        $rev = Resolve-Reverse -IPAddress $entry
        $reverseNames = $rev.Values
        $reverseServer = $rev.Server

        $forwardIPs = @()
        $forwardServer = $null
        $forwardResolvedName = $null
        foreach ($n in $reverseNames) {
            $fwd = Resolve-Forward -Name $n
            $forwardIPs += $fwd.Values
            if (-not $forwardServer) { $forwardServer = $fwd.Server }
            if (-not $forwardResolvedName) { $forwardResolvedName = $fwd.ResolvedName }
        }
        $forwardIPs = $forwardIPs | Select-Object -Unique

        $status =
            if (-not $reverseNames) { 'ReverseFailed' }
            elseif (-not $forwardIPs) { 'ForwardFailed' }
            elseif ($forwardIPs -contains $entry) { 'Match' }
            else { 'Mismatch' }
    }
    else {
        $fwd = Resolve-Forward -Name $entry
        $forwardIPs = $fwd.Values
        $forwardServer = $fwd.Server
        $forwardResolvedName = $fwd.ResolvedName

        $reverseNames = @()
        $reverseServer = $null
        foreach ($addr in $forwardIPs) {
            $rev = Resolve-Reverse -IPAddress $addr
            $reverseNames += $rev.Values
            if (-not $reverseServer) { $reverseServer = $rev.Server }
        }
        $reverseNames = $reverseNames | Select-Object -Unique

        $status =
            if (-not $forwardIPs) { 'ForwardFailed' }
            elseif (-not $reverseNames) { 'ReverseFailed' }
            elseif (Test-NameRoundTrip -Entry $entry -Candidates $reverseNames) { 'Match' }
            else { 'Mismatch' }
    }

    # Pick one name and one IP to ping - prefer the fully-qualified name we actually resolved
    # (needed when the entry was a short name and only a non-default -DnsServer could find it),
    # falling back to the entry itself.
    $testName = if ($isIp) { $reverseNames | Select-Object -First 1 } else { if ($forwardResolvedName) { $forwardResolvedName } else { $entry } }
    $testIP = if ($isIp) { $entry } else { $forwardIPs | Select-Object -First 1 }

    # Prefer the FQDN our own forward lookup resolved to (handles short names where we had to
    # append a DNS suffix ourselves); fall back to the PTR name, then to the entry if it was
    # already given as an FQDN. Each source only has a domain to give if it's actually an FQDN,
    # so keep trying the next one rather than stopping at the first non-null source.
    $domain = $null
    if ($forwardResolvedName) { $domain = Get-DnsDomain -Fqdn $forwardResolvedName }
    if (-not $domain -and $reverseNames) { $domain = Get-DnsDomain -Fqdn ($reverseNames | Select-Object -First 1) }
    if (-not $domain -and -not $isIp -and $entry -match '\.') { $domain = Get-DnsDomain -Fqdn $entry }

    if ($SkipConnectivityTest) {
        $pingByName = 'Skipped'
        $pingByIP = 'Skipped'
    }
    else {
        $pingByName = Test-HostResponse -Target $testName -TimeoutMs $PingTimeoutMs
        $pingByIP = Test-HostResponse -Target $testIP -TimeoutMs $PingTimeoutMs
    }

    $report.Add([pscustomobject]@{
        Entry        = $entry
        Type         = if ($isIp) { 'IPAddress' } else { 'Hostname' }
        Domain       = if ($domain) { $domain } else { 'Unknown' }
        ForwardIPs   = ($forwardIPs -join ', ')
        ReverseNames = ($reverseNames -join ', ')
        Status       = $status
        AnsweredBy   = "Fwd:$(if ($forwardServer) { $forwardServer } else { '-' }) / Rev:$(if ($reverseServer) { $reverseServer } else { '-' })"
        PingByName   = $pingByName
        PingByIP     = $pingByIP
    })
}

$report = $report | Sort-Object Status, Entry

Write-Host "`nDNS and connectivity validation results ($($report.Count) entr$(if ($report.Count -eq 1) {'y'} else {'ies'})):" -ForegroundColor Cyan
$report | Format-Table Entry, Type, Domain, ForwardIPs, ReverseNames, Status, AnsweredBy, PingByName, PingByIP -AutoSize

$dnsIssues = @($report | Where-Object { $_.Status -ne 'Match' })
if ($dnsIssues) {
    Write-Warning "$($dnsIssues.Count) entr$(if ($dnsIssues.Count -eq 1) {'y'} else {'ies'}) failed forward/reverse DNS validation."
}
else {
    Write-Host "`nAll entries have matching forward and reverse DNS records." -ForegroundColor Green
}

if (-not $SkipConnectivityTest) {
    $pingIssues = @($report | Where-Object { $_.PingByName -eq 'Failed' -or $_.PingByIP -eq 'Failed' })
    if ($pingIssues) {
        Write-Warning "$($pingIssues.Count) entr$(if ($pingIssues.Count -eq 1) {'y'} else {'ies'}) didn't respond to ping by name and/or by IP."
    }
    else {
        Write-Host "All entries responded to ping by name and by IP." -ForegroundColor Green
    }
}

if ($OutputPath) {
    $report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nExported to $OutputPath" -ForegroundColor Green
}
