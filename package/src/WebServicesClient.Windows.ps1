<#
.SYNOPSIS
    Web Services Client for Exchange - Windows authentication: Negotiate, NTLM, Kerberos (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    The handshake is done by the tool itself with System.Net.Security.NegotiateAuthentication (.NET 7
    and later, SSPI on Windows), not hidden inside the HTTP client: every leg is a request of the trace.
      NTLM       request + type 1 (negotiate) -> 401 + type 2 (challenge) -> request + type 3 (authenticate) -> 200
      Kerberos   request + AP-REQ (service ticket for HTTP/<host>) -> 200 (+ mutual authentication token)
      Negotiate  Kerberos when a KDC of the realm answers and knows the SPN, NTLM otherwise (like Windows)
    The legs of NTLM must use the same TCP connection: the HTTP client of a Windows run keeps one
    connection to each server (MaxConnectionsPerServer = 1).

    The tokens are decoded for the trace: NTLM messages (flags, names; the type 2 names the server,
    its NetBIOS and DNS domain and its Windows version), SPNEGO (mechanism chosen, state). The
    NTLM response to the challenge and the Kerberos ticket are never written, only their length.

    Credentials: the current Windows account (Kerberos needs a domain account and a domain controller
    in reach) or -Credential (user and password, kept in memory). A computer out of the domain can use
    NTLM with -Credential; Kerberos also needs a KDC of the realm reachable (DNS SRV _kerberos._tcp,
    port 88) and the SPN of the EWS host name (HTTP/<host>, the alternate service account of Exchange
    for a load-balanced name).

    Extended Protection: every handshake carries the channel binding token (CBT) of the TLS connection,
    like Windows and Outlook: 'tls-server-end-point' (RFC 5929), the hash of the certificate the server
    presents (SHA-256, or the hash of its signature when stronger). Without it Exchange refuses the NTLM
    token with STATUS_BAD_BINDINGS (0xC000035B, Security event 4625), even with tokenChecking Allow - seen
    behind an IIS ARR reverse proxy that re-encrypts to Exchange with the same certificate. The type 3
    message decoded in the trace shows the CBT and the SPN sent.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:NtlmSignature = [byte[]](0x4E, 0x54, 0x4C, 0x4D, 0x53, 0x53, 0x50, 0x00)

# Channel binding of a TLS connection for NegotiateAuthentication: a SEC_CHANNEL_BINDINGS structure (32 bytes)
# followed by the application data 'tls-server-end-point:' + hash of the server certificate.
if (-not ('WscChannelBinding' -as [type])) {
    Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices; using System.Security.Authentication.ExtendedProtection;
public sealed class WscChannelBinding : ChannelBinding {
    private readonly int _size;
    public WscChannelBinding(byte[] applicationData) : base(true) {
        _size = 32 + applicationData.Length;
        IntPtr p = Marshal.AllocHGlobal(_size);
        for (int i = 0; i < 32; i++) Marshal.WriteByte(p, i, 0);
        Marshal.WriteInt32(p, 24, applicationData.Length);
        Marshal.WriteInt32(p, 28, 32);
        Marshal.Copy(applicationData, 0, p + 32, applicationData.Length);
        SetHandle(p);
    }
    public override int Size { get { return _size; } }
    protected override bool ReleaseHandle() { Marshal.FreeHGlobal(handle); return true; }
}
'@
}

function Get-WscTlsServerCertificate {
    <# The certificate a server presents (direct TLS connection, default validation), or $null. #>
    param([Parameter(Mandatory = $true)][string]$HostName, [int]$Port = 443, [int]$TimeoutSeconds = 10)

    $tcp = [Net.Sockets.TcpClient]::new()
    try {
        if (-not $tcp.ConnectAsync($HostName, $Port).Wait([TimeSpan]::FromSeconds($TimeoutSeconds))) { return $null }
        $ssl = [Net.Security.SslStream]::new($tcp.GetStream(), $false)
        try { $ssl.AuthenticateAsClient($HostName); return [Security.Cryptography.X509Certificates.X509Certificate2]::new($ssl.RemoteCertificate) }
        finally { $ssl.Dispose() }
    }
    catch { return $null }
    finally { $tcp.Dispose() }
}

function Get-WscChannelBindingData {
    <#
        Application data of the tls-server-end-point channel binding (RFC 5929) for a server: the hash of its
        certificate - SHA-256, or SHA-384 / SHA-512 when the certificate is signed with it. Cached per host.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][Uri]$Uri)

    $key = $Uri.Authority
    if ($Context.ChannelBindings.ContainsKey($key)) { return $Context.ChannelBindings[$key] }
    $cert = Get-WscTlsServerCertificate -HostName $Uri.Host -Port $Uri.Port -TimeoutSeconds ([Math]::Min(15, [int]$Context.Config.HttpTimeoutSeconds))
    $value = $null
    if ($cert) {
        $name = [string]$cert.SignatureAlgorithm.FriendlyName
        $hash = if ($name -match 'sha384') { [Security.Cryptography.SHA384]::HashData($cert.RawData); $algo = 'SHA-384' } elseif ($name -match 'sha512') { [Security.Cryptography.SHA512]::HashData($cert.RawData); $algo = 'SHA-512' } else { [Security.Cryptography.SHA256]::HashData($cert.RawData); $algo = 'SHA-256' }
        $value = [pscustomobject]@{
            Data = [byte[]]([Text.Encoding]::ASCII.GetBytes('tls-server-end-point:') + $hash)
            Text = "tls-server-end-point, $algo of $($cert.Subject) (thumbprint $($cert.Thumbprint))"
        }
    }
    $Context.ChannelBindings[$key] = $value
    return $value
}
# OIDs (DER, with tag and length) of the mechanisms in SPNEGO.
$script:SpnegoMechs = [ordered]@{
    'Kerberos'           = [byte[]](0x06, 0x09, 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x12, 0x01, 0x02, 0x02)
    'Kerberos (MS)'      = [byte[]](0x06, 0x09, 0x2A, 0x86, 0x48, 0x82, 0xF7, 0x12, 0x01, 0x02, 0x02)
    'NTLM'               = [byte[]](0x06, 0x0A, 0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x02, 0x0A)
    'NegoEx'             = [byte[]](0x06, 0x0A, 0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x02, 0x1E)
}

function Find-WscBytes {
    <# Position of a byte sequence in an array, or -1. #>
    param([Parameter(Mandatory = $true)][byte[]]$Data, [Parameter(Mandatory = $true)][byte[]]$Pattern, [int]$Start = 0)

    for ($i = $Start; $i -le $Data.Length - $Pattern.Length; $i++) {
        $hit = $true
        for ($j = 0; $j -lt $Pattern.Length; $j++) { if ($Data[$i + $j] -ne $Pattern[$j]) { $hit = $false; break } }
        if ($hit) { return $i }
    }
    return -1
}

function ConvertFrom-WscBase64Safe {
    param([AllowEmptyString()][string]$Text)
    try { return [Convert]::FromBase64String($Text.Trim()) } catch { return $null }
}

function Read-WscNtlmField {
    <# A security buffer of an NTLM message (length, max length, offset) as text (UTF-16 or OEM) or bytes. #>
    param([Parameter(Mandatory = $true)][byte[]]$Data, [Parameter(Mandatory = $true)][int]$Offset, [int]$Base = 0, [switch]$Unicode, [switch]$Bytes)

    if ($Base + $Offset + 8 -gt $Data.Length) { return $null }
    $length = [BitConverter]::ToUInt16($Data, $Base + $Offset)
    $start = [BitConverter]::ToUInt32($Data, $Base + $Offset + 4)
    if ($length -eq 0 -or $Base + $start + $length -gt $Data.Length) { return $(if ($Bytes) { [byte[]]@() } else { '' }) }
    if ($Bytes) { return [byte[]]$Data[($Base + $start)..($Base + $start + $length - 1)] }
    $encoding = if ($Unicode) { [Text.Encoding]::Unicode } else { [Text.Encoding]::ASCII }
    return $encoding.GetString($Data, $Base + $start, $length)
}

function ConvertFrom-WscNtlmMessage {
    <#
        Fields of an NTLM message (MS-NLMP) starting at Base: type, flags, names, server information.
        Never returns the challenge responses, only their length.
    #>
    param([Parameter(Mandatory = $true)][byte[]]$Data, [int]$Base = 0)

    if ($Data.Length -lt $Base + 12) { return $null }
    $type = [BitConverter]::ToUInt32($Data, $Base + 8)
    $info = [ordered]@{ Type = [int]$type; Name = switch ($type) { 1 { 'negotiate' } 2 { 'challenge' } 3 { 'authenticate' } default { 'unknown' } } }
    $version = {
        param([int]$At)
        if ($Data.Length -lt $Base + $At + 8) { return $null }
        $b = $Data[$Base + $At]; $m = $Data[$Base + $At + 1]; $build = [BitConverter]::ToUInt16($Data, $Base + $At + 2)
        if ($b -eq 0) { return $null }
        "Windows $b.$m build $build"
    }
    switch ($type) {
        1 {
            $flags = [BitConverter]::ToUInt32($Data, $Base + 12)
            $info.Flags = '0x{0:X8}' -f $flags
            $info.Domain = Read-WscNtlmField -Data $Data -Offset 16 -Base $Base
            $info.Workstation = Read-WscNtlmField -Data $Data -Offset 24 -Base $Base
            if ($flags -band 0x02000000) { $info.ClientVersion = & $version 32 }
        }
        2 {
            $flags = [BitConverter]::ToUInt32($Data, $Base + 20)
            $unicode = [bool]($flags -band 1)
            $info.Flags = '0x{0:X8}' -f $flags
            $info.Target = Read-WscNtlmField -Data $Data -Offset 12 -Base $Base -Unicode:$unicode
            $avs = Read-WscNtlmField -Data $Data -Offset 40 -Base $Base -Bytes
            $names = @{ 1 = 'NetBiosComputer'; 2 = 'NetBiosDomain'; 3 = 'DnsComputer'; 4 = 'DnsDomain'; 5 = 'DnsForest'; 9 = 'TargetName' }
            for ($i = 0; $avs -and $i + 4 -le $avs.Length) {
                $id = [BitConverter]::ToUInt16($avs, $i); $len = [BitConverter]::ToUInt16($avs, $i + 2)
                if ($id -eq 0 -or $i + 4 + $len -gt $avs.Length) { break }
                if ($names.ContainsKey([int]$id)) { $info[$names[[int]$id]] = [Text.Encoding]::Unicode.GetString($avs, $i + 4, $len) }
                $i += 4 + $len
            }
            if ($flags -band 0x02000000) { $info.ServerVersion = & $version 48 }
        }
        3 {
            $nt = Read-WscNtlmField -Data $Data -Offset 20 -Base $Base -Bytes
            $flags = [BitConverter]::ToUInt32($Data, $Base + 60)
            $unicode = [bool]($flags -band 1)
            $info.Flags = '0x{0:X8}' -f $flags
            $info.Domain = Read-WscNtlmField -Data $Data -Offset 28 -Base $Base -Unicode:$unicode
            $info.User = Read-WscNtlmField -Data $Data -Offset 36 -Base $Base -Unicode:$unicode
            $info.Workstation = Read-WscNtlmField -Data $Data -Offset 44 -Base $Base -Unicode:$unicode
            $info.Response = if ($nt.Length -gt 24) { "NTLMv2, $($nt.Length) bytes, never written" } elseif ($nt.Length) { "NTLMv1, $($nt.Length) bytes, never written" } else { 'anonymous' }
            if ($nt.Length -gt 48) {
                # AV pairs of the NTLMv2 client challenge (after NTProofStr and 28 bytes): SPN and channel binding sent.
                for ($i = 44; $i + 4 -le $nt.Length) {
                    $id = [BitConverter]::ToUInt16($nt, $i); $len = [BitConverter]::ToUInt16($nt, $i + 2)
                    if ($id -eq 0 -or $i + 4 + $len -gt $nt.Length) { break }
                    if ($id -eq 9) { $info.Spn = if ($len) { [Text.Encoding]::Unicode.GetString($nt, $i + 4, $len) } else { 'none' } }
                    if ($id -eq 10) { $info.ChannelBinding = if (@($nt[($i + 4)..($i + 3 + $len)] | Where-Object { $_ -ne 0 }).Count) { 'present (hash of the TLS binding)' } else { 'none (zeros)' } }
                    $i += 4 + $len
                }
            }
        }
    }
    return [pscustomobject]$info
}

function Get-WscNegotiateInfo {
    <#
        What a Negotiate, NTLM or Kerberos token contains: Mechanism (NTLM, Kerberos), Form (raw NTLM,
        SPNEGO init or response), SPNEGO state, and the NTLM message decoded.
    #>
    param([AllowEmptyString()][string]$Base64)

    $data = ConvertFrom-WscBase64Safe $Base64
    $info = [ordered]@{ Bytes = 0; Form = 'unknown'; Mechanism = $null; State = $null; Ntlm = $null }
    if (-not $data -or -not $data.Length) { return [pscustomobject]$info }
    $info.Bytes = $data.Length
    $ntlmAt = Find-WscBytes -Data $data -Pattern $script:NtlmSignature
    if ($ntlmAt -eq 0) { $info.Form = 'NTLM'; $info.Mechanism = 'NTLM' }
    elseif ($data[0] -eq 0x60) { $info.Form = 'SPNEGO init' }
    elseif ($data[0] -eq 0xA1) { $info.Form = 'SPNEGO response' }
    elseif ($data[0] -eq 0x6E -or $data[0] -eq 0x6F) { $info.Form = 'Kerberos'; $info.Mechanism = 'Kerberos' }
    if ($info.Form -like 'SPNEGO*') {
        $first = $null; $firstAt = [int]::MaxValue
        foreach ($name in $script:SpnegoMechs.Keys) {
            $at = Find-WscBytes -Data $data -Pattern $script:SpnegoMechs[$name]
            if ($at -ge 0 -and $at -lt $firstAt) { $first = $name; $firstAt = $at }
        }
        $info.Mechanism = if ($ntlmAt -gt 0) { 'NTLM' } elseif ($first) { ($first -replace ' \(MS\)$', '') } else { $null }
        if ($info.Form -eq 'SPNEGO response') {
            # negState: A0 03 0A 01 <value>
            $at = Find-WscBytes -Data $data -Pattern ([byte[]](0xA0, 0x03, 0x0A, 0x01))
            if ($at -ge 0 -and $at + 4 -lt $data.Length) {
                $info.State = switch ($data[$at + 4]) { 0 { 'accept-completed' } 1 { 'accept-incomplete' } 2 { 'reject' } 3 { 'request-mic' } default { 'unknown' } }
            }
        }
    }
    if ($ntlmAt -ge 0) { $info.Ntlm = ConvertFrom-WscNtlmMessage -Data $data -Base $ntlmAt }
    return [pscustomobject]$info
}

function Get-WscNegotiateKind {
    <# Short label of a token for the list of requests: NTLM type 1, Kerberos ticket... #>
    param([AllowEmptyString()][string]$Base64)

    $i = Get-WscNegotiateInfo -Base64 $Base64
    if ($i.Ntlm) { return "NTLM type $($i.Ntlm.Type) ($($i.Ntlm.Name))" }
    if ($i.Mechanism -eq 'Kerberos' -and $i.Form -ne 'SPNEGO response') { return 'Kerberos ticket' }
    if ($i.Form -eq 'SPNEGO response') { return "SPNEGO $($i.State)" }
    return "$($i.Form) token"
}

function Format-WscNegotiateSummary {
    <# A token decoded on one line for the trace; secrets (responses, tickets) only by their length. #>
    param([AllowEmptyString()][string]$Base64)

    $i = Get-WscNegotiateInfo -Base64 $Base64
    if (-not $i.Bytes) { return "token not decodable, $($Base64.Length) characters" }
    $parts = [Collections.Generic.List[string]]::new()
    $wrap = if ($i.Form -like 'SPNEGO*') { "$($i.Form)$(if ($i.State) { " $($i.State)" }), " } else { '' }
    if ($i.Ntlm) {
        $n = $i.Ntlm
        $parts.Add("$($wrap)NTLM type $($n.Type) ($($n.Name)), $($i.Bytes) bytes")
        foreach ($k in 'Target', 'NetBiosDomain', 'NetBiosComputer', 'DnsComputer', 'DnsDomain', 'DnsForest', 'ServerVersion', 'Domain', 'User', 'Workstation', 'ClientVersion', 'Response', 'Spn', 'ChannelBinding', 'Flags') {
            $p = $n.PSObject.Properties[$k]
            if ($p -and $p.Value) { $parts.Add("$k $($p.Value)") }
        }
    }
    elseif ($i.Mechanism -eq 'Kerberos') {
        $what = if ($i.Form -eq 'SPNEGO response') { 'mutual authentication token' } else { 'service ticket (AP-REQ), never written' }
        $parts.Add("$($wrap)Kerberos $what, $($i.Bytes) bytes")
    }
    else {
        $parts.Add("$($wrap)$($i.Bytes) bytes$(if ($i.Mechanism) { ", $($i.Mechanism)" })")
    }
    return $parts -join '; '
}

function Get-WscWindowsErrorText {
    <# What a status of NegotiateAuthentication means for the test. #>
    param([Parameter(Mandatory = $true)][string]$Status, [Parameter(Mandatory = $true)][string]$Package, [string]$Spn)

    switch ($Status) {
        'TargetUnknown' { return "the KDC does not know the SPN $($Spn): register it on the account of Exchange (the alternate service account, ASA, for a load-balanced name), or use NTLM." }
        { $_ -in 'UnknownCredentials', 'InvalidCredentials' } {
            if ($Package -eq 'Kerberos') { return 'no Kerberos ticket: no domain controller of the realm answers from this computer (outside the domain, VPN, port 88), or the user name or password is wrong.' }
            return 'the user name or the password is wrong.'
        }
        'CredentialsExpired' { return 'the password has expired.' }
        'Unsupported' { return "the $Package package is not available on this computer." }
        default {
            if ($Package -eq 'Kerberos') { return "SSPI answered $($Status): Kerberos needs a domain controller of the realm in reach (DNS _kerberos._tcp, port 88) and the SPN $($Spn)." }
            return "SSPI answered $Status."
        }
    }
}

function Invoke-WscWindowsRequest {
    <#
        One request authenticated with Windows. When the connection is already authenticated (an
        earlier handshake on it), the request goes without a token first, like a browser; a 401 starts
        a new handshake. NewRequest builds a fresh copy of the request for each leg (same body).
        Returns Response, Package (Kerberos or NTLM as negotiated), Legs, ServerInfo (NTLM type 2).
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Context,
        [Parameter(Mandatory = $true)][scriptblock]$NewRequest,
        [string]$Label,
        [switch]$ForceHandshake
    )

    $cfg = $Context.Config
    $package = [string]$cfg.WindowsPackage
    $scheme = if ($package -eq 'NTLM') { 'NTLM' } else { 'Negotiate' }
    if ($Context.WindowsSession -and -not $ForceHandshake) {
        $request = & $NewRequest
        try { $response = Invoke-WscHttp -HttpClient $Context.HttpClient -Request $request -Label $Label } finally { $request.Dispose() }
        if ($response.StatusCode -ne 401) {
            return [pscustomobject]@{ Response = $response; Package = $Context.WindowsPackageUsed; Legs = 0; ServerInfo = $null; Reused = $true }
        }
        $Context.WindowsSession = $false
        $script:WscWindowsSession = $false
    }
    $hostName = ([Uri]$(& { $r = & $NewRequest; try { $r.RequestUri.AbsoluteUri } finally { $r.Dispose() } })).Host
    $spn = "HTTP/$hostName"
    $options = [Net.Security.NegotiateAuthenticationClientOptions]::new()
    $options.Package = $package
    $options.TargetName = $spn
    $options.RequiredProtectionLevel = [Net.Security.ProtectionLevel]::None
    $options.Credential = if ($Context.Credential) { $Context.Credential.GetNetworkCredential() } else { [Net.CredentialCache]::DefaultNetworkCredentials }
    # Extended Protection: the channel binding of the TLS connection, like Windows and Outlook.
    $cbt = $null
    $probeRequest = & $NewRequest
    try { $target = $probeRequest.RequestUri } finally { $probeRequest.Dispose() }
    if ($target.Scheme -eq 'https') { $cbt = Get-WscChannelBindingData -Context $Context -Uri $target }
    if ($cbt) { $options.Binding = [WscChannelBinding]::new($cbt.Data) }
    $client = [Net.Security.NegotiateAuthentication]::new($options)
    $incoming = [NullString]::Value
    $serverInfo = $null
    $legs = 0
    try {
        for ($leg = 1; $leg -le 4; $leg++) {
            $status = [Net.Security.NegotiateAuthenticationStatusCode]::GenericFailure
            $blob = $client.GetOutgoingBlob($incoming, [ref]$status)
            if ($status -notin 'Completed', 'ContinueNeeded' -or [string]::IsNullOrEmpty($blob)) {
                throw "Windows could not build the $package token for $($spn): $(Get-WscWindowsErrorText -Status ([string]$status) -Package $package -Spn $spn)"
            }
            $request = & $NewRequest
            $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new($scheme, $blob)
            $legs++
            try { $response = Invoke-WscHttp -HttpClient $Context.HttpClient -Request $request -Label "$Label leg $leg".Trim() }
            finally { $request.Dispose() }
            $token = $null
            foreach ($c in @(Get-WscField $response 'Challenges')) {
                $m = [regex]::Match([string]$c, "^\s*$scheme\s+(\S+)", 'IgnoreCase')
                if ($m.Success) { $token = $m.Groups[1].Value; break }
            }
            if ($token) {
                $decoded = Get-WscNegotiateInfo -Base64 $token
                if ($decoded.Ntlm -and $decoded.Ntlm.Type -eq 2) { $serverInfo = $decoded.Ntlm }
            }
            if ($response.StatusCode -ne 401) {
                # Kerberos: the 200 can carry the mutual authentication token of the server.
                if ($token) { $s2 = $status; [void]$client.GetOutgoingBlob($token, [ref]$s2) }
                $used = try { [string]$client.Package } catch { $package }
                if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 300) {
                    $Context.WindowsSession = $true
                    $Context.WindowsPackageUsed = $used
                    $script:WscWindowsSession = $true
                }
                return [pscustomobject]@{ Response = $response; Package = $used; Legs = $legs; ServerInfo = $serverInfo; Reused = $false; Spn = $spn; ChannelBinding = $(if ($cbt) { $cbt.Text } else { 'none' }) }
            }
            if (-not $token) {
                # 401 without a token to continue: the credentials were refused.
                $used = try { [string]$client.Package } catch { $package }
                return [pscustomobject]@{ Response = $response; Package = $used; Legs = $legs; ServerInfo = $serverInfo; Reused = $false; Spn = $spn; ChannelBinding = $(if ($cbt) { $cbt.Text } else { 'none' }) }
            }
            $incoming = $token
        }
        throw "The $package handshake did not end after 4 legs."
    }
    finally {
        $client.Dispose()
        if ($options.Binding) { $options.Binding.Dispose() }
    }
}

function Test-WscKerberosRealm {
    <#
        Kerberos from this computer for a DNS domain: the SRV record _kerberos._tcp.<domain> and a TCP
        connection to port 88 of the first domain controller found. No ticket is requested.
    #>
    param([Parameter(Mandatory = $true)][string]$Domain, [int]$TimeoutSeconds = 3)

    $result = [ordered]@{ Domain = $Domain; Kdc = $null; Port88 = $false; Error = $null; DomainJoined = $false; ComputerDomain = $null }
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $result.DomainJoined = [bool]$cs.PartOfDomain
        $result.ComputerDomain = [string]$cs.Domain
    }
    catch { }
    try {
        $records = @(Resolve-DnsName -Name "_kerberos._tcp.$Domain" -Type SRV -DnsOnly -QuickTimeout -ErrorAction Stop | Where-Object { $_.QueryType -eq 'SRV' } | Sort-Object Priority, Weight)
        if (-not $records.Count) { $result.Error = "no SRV record _kerberos._tcp.$Domain"; return [pscustomobject]$result }
        $result.Kdc = [string]$records[0].NameTarget
        $tcp = [Net.Sockets.TcpClient]::new()
        try {
            $result.Port88 = $tcp.ConnectAsync($result.Kdc, 88).Wait([TimeSpan]::FromSeconds($TimeoutSeconds))
            if (-not $result.Port88) { $result.Error = "no TCP connection to $($result.Kdc):88 within $TimeoutSeconds s" }
        }
        catch { $inner = $_.Exception; while ($inner.InnerException) { $inner = $inner.InnerException }; $result.Error = $inner.Message }
        finally { $tcp.Dispose() }
    }
    catch {
        $result.Error = "DNS: $($_.Exception.Message)"
    }
    return [pscustomobject]$result
}
