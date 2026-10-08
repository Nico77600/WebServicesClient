<#
.SYNOPSIS
    Web Services Client for Exchange - Autodiscover and the prerequisites without sign-in (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    Autodiscover, like Outlook:
      1. Autodiscover v2 (JSON, anonymous): https://autodiscover.<domain>, https://<domain>, and
         https://outlook.office365.com (every Exchange Online mailbox), /autodiscover/autodiscover.json/v1.0/<address>?Protocol=EWS.
      2. Classic Autodiscover (POX, authenticated, Basic or Windows only): the same hosts, the host of the
         HTTP redirect method (http://autodiscover.<domain>) and of the DNS SRV record _autodiscover._tcp.<domain>.
      3. Target.EwsUrl, when Autodiscover gives no answer (the server typed by hand).
    With OAuth only the anonymous Autodiscover v2 is possible before the sign-in: the audience of the
    token is the EWS server, so the EWS URL must be known first (this is what Outlook does too).

    Prerequisites (no sign-in, nothing changed): authorization server (AD FS metadata, Entra ID tenant
    and user realm), TLS certificates, the authentication schemes EWS offers to an anonymous request,
    OAuth offered to the mailbox (authorization_uri), an NTLM challenge (names of the server and of the
    domain, no credentials), the Kerberos realm, and a forged token or a wrong password refused.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

#region Autodiscover -------------------------------------------------------------------------------

function Get-WscAutodiscoverV2 {
    <# Autodiscover v2 on one host: the EWS URL, following up to 4 HTTPS redirects. Returns Url, Request, Error. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$BaseUrl)

    $mailbox = [string]$Context.Config.Mailbox
    $uri = "$BaseUrl/autodiscover/autodiscover.json/v1.0/$([Uri]::EscapeDataString($mailbox))?Protocol=EWS"
    try {
        $response = $null
        for ($hop = 0; $hop -lt 4; $hop++) {
            Assert-WscNotCancelled
            $response = Invoke-WscWebRequest -HttpClient $Context.HttpClient -Uri $uri -UserAgent ([string]$Context.Config.UserAgent) -Headers @{ Accept = 'application/json' }
            if ($response.StatusCode -in 301, 302, 307, 308 -and [string]$response.Location -like 'https://*') { $uri = $response.Location; continue }
            break
        }
        if ($response.StatusCode -eq 200) {
            $json = $response.Content | ConvertFrom-Json -ErrorAction Stop
            $url = [string](Get-WscField $json 'Url')
            if ($url) { return [pscustomobject]@{ Url = $url; Request = $uri; Error = $null } }
            return [pscustomobject]@{ Url = $null; Request = $uri; Error = "HTTP 200 without URL ($([string](Get-WscField $json 'ErrorCode')) $([string](Get-WscField $json 'ErrorMessage')))".Trim() }
        }
        return [pscustomobject]@{ Url = $null; Request = $uri; Error = "HTTP $($response.StatusCode)" }
    }
    catch {
        return [pscustomobject]@{ Url = $null; Request = $uri; Error = $_.Exception.Message }
    }
}

function Get-WscAutodiscoverPox {
    <#
        Classic Autodiscover (POX) on one host, with the credentials of the run (Basic or Windows).
        Returns the external EWS URL (EXPR), else the internal one (EXCH), or a redirect address.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$BaseUrl)

    $cfg = $Context.Config
    $uri = "$BaseUrl/autodiscover/autodiscover.xml"
    $body = '<?xml version="1.0" encoding="utf-8"?><Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006"><Request>' +
        "<EMailAddress>$(ConvertTo-WscXmlText ([string]$cfg.Mailbox))</EMailAddress><AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema></Request></Autodiscover>"
    $wscPoxBuild = @{ Uri = $uri; Body = $body; Config = $cfg; Credential = $Context.Credential }
    $build = {
        $b = $wscPoxBuild
        $r = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Post, $b.Uri)
        $r.Content = [Net.Http.ByteArrayContent]::new([Text.Encoding]::UTF8.GetBytes($b.Body))
        $r.Content.Headers.ContentType = [Net.Http.Headers.MediaTypeHeaderValue]::Parse('text/xml; charset=utf-8')
        $null = $r.Headers.TryAddWithoutValidation('User-Agent', [string]$b.Config.UserAgent)
        $null = $r.Headers.TryAddWithoutValidation('X-AnchorMailbox', [string]$b.Config.Mailbox)
        if ([string]$b.Config.Authentication -eq 'Basic' -and $b.Credential) {
            $pair = '{0}:{1}' -f $b.Credential.UserName, $b.Credential.GetNetworkCredential().Password
            $r.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Basic', [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair)))
        }
        $r
    }
    try {
        if ([string]$cfg.Authentication -eq 'Windows') { $response = (Invoke-WscWindowsRequest -Context $Context -NewRequest $build -Label 'Autodiscover' -ForceHandshake).Response }
        else { $request = & $build; try { $response = Invoke-WscHttp -HttpClient $Context.HttpClient -Request $request -Label 'Autodiscover' } finally { $request.Dispose() } }
        if ($response.StatusCode -ne 200) { return [pscustomobject]@{ Url = $null; Request = $uri; Error = "HTTP $($response.StatusCode)"; Redirect = $null } }
        $xml = [xml][Text.Encoding]::UTF8.GetString([byte[]]$response.Body)
        $protocols = @($xml.SelectNodes('//*[local-name()="Protocol"]'))
        $pick = { param([string]$Type) $p = $protocols | Where-Object { $_.SelectSingleNode('*[local-name()="Type"]').InnerText -eq $Type } | Select-Object -First 1; if ($p) { $n = $p.SelectSingleNode('*[local-name()="EwsUrl"]'); if ($n) { $n.InnerText } } }
        $url = & $pick 'EXPR'
        if (-not $url) { $url = & $pick 'EXCH' }
        $redirect = $xml.SelectSingleNode('//*[local-name()="RedirectAddr"]')
        if ($url) { return [pscustomobject]@{ Url = $url; Request = $uri; Error = $null; Redirect = $null } }
        if ($redirect) { return [pscustomobject]@{ Url = $null; Request = $uri; Error = "redirect to $($redirect.InnerText)"; Redirect = $redirect.InnerText } }
        $error = $xml.SelectSingleNode('//*[local-name()="Error"]/*[local-name()="Message"]')
        return [pscustomobject]@{ Url = $null; Request = $uri; Error = "no EWS URL$(if ($error) { ": $($error.InnerText)" })"; Redirect = $null }
    }
    catch {
        return [pscustomobject]@{ Url = $null; Request = $uri; Error = $_.Exception.Message; Redirect = $null }
    }
}

function Get-WscAutodiscoverRedirectHost {
    <# HTTP redirect method: GET http://autodiscover.<domain>/autodiscover/autodiscover.xml -> 302 to an HTTPS URL. Returns its base, or $null. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Domain)
    try {
        $r = Invoke-WscWebRequest -HttpClient $Context.HttpClient -Uri "http://autodiscover.$Domain/autodiscover/autodiscover.xml" -UserAgent ([string]$Context.Config.UserAgent)
        if ($r.StatusCode -in 301, 302, 307, 308 -and [string]$r.Location -match '^(https://[^/]+)/autodiscover/autodiscover\.xml') { return $Matches[1] }
    }
    catch { }
    return $null
}

function Get-WscAutodiscoverSrvHost {
    <# DNS SRV method: _autodiscover._tcp.<domain>. Returns https://<target>, or $null. #>
    param([Parameter(Mandatory = $true)][string]$Domain)
    try {
        $srv = Resolve-DnsName -Name "_autodiscover._tcp.$Domain" -Type SRV -DnsOnly -QuickTimeout -ErrorAction Stop | Where-Object { $_.QueryType -eq 'SRV' } | Sort-Object Priority, Weight | Select-Object -First 1
        if ($srv) { return "https://$($srv.NameTarget.TrimEnd('.'))$(if ($srv.Port -and $srv.Port -ne 443) { ":$($srv.Port)" })" }
    }
    catch { }
    return $null
}

function Set-WscEwsUrl {
    <# Applies the EWS URL found to the run and recomputes the endpoints (resource and scope of the token). #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Url, [Parameter(Mandatory = $true)][string]$Source)
    $Context.Config.EwsUrl = $Url
    $Context.Endpoints = Resolve-WscEndpoints -Configuration $Context.Config
    $Context.EwsUrlSource = $Source
}

function Invoke-WscStageAutodiscover {
    <# Finds the EWS URL of the mailbox like Outlook. Stops the scenario when no URL is found. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Autodiscover'
    $cfg = $Context.Config
    $mailbox = [string]$cfg.Mailbox
    $domain = $mailbox.Split('@')[-1]
    $typed = [string]$cfg.EwsUrl
    $attempts = [Collections.Generic.List[string]]::new()
    $bases = [Collections.Generic.List[string]]::new()
    foreach ($b in "https://autodiscover.$domain", "https://$domain") { $bases.Add($b) }
    # Exchange Online answers Autodiscover v2 for every mailbox of every tenant (onmicrosoft.com domains included), like Outlook asks it.
    $bases.Add('https://outlook.office365.com')

    $found = $null
    foreach ($base in $bases) {
        $v2 = Get-WscAutodiscoverV2 -Context $Context -BaseUrl $base
        if ($v2.Url) { $found = [pscustomobject]@{ Url = $v2.Url; Method = 'Autodiscover v2'; Request = $v2.Request }; break }
        [void]$attempts.Add("v2 $($base): $($v2.Error)")
    }
    if (-not $found -and ([string]$cfg.Authentication -eq 'Windows' -or ([string]$cfg.Authentication -eq 'Basic' -and $Context.Credential))) {
        # The classic Autodiscover needs credentials: Basic or Windows can send them before knowing EWS.
        $pox = [Collections.Generic.List[string]]::new()
        foreach ($b in $bases) { if ($b -notlike '*office365*') { $pox.Add($b) } }
        $redirectHost = Get-WscAutodiscoverRedirectHost -Context $Context -Domain $domain
        if ($redirectHost -and -not $pox.Contains($redirectHost)) { $pox.Add($redirectHost) }
        $srvHost = Get-WscAutodiscoverSrvHost -Domain $domain
        if ($srvHost -and -not $pox.Contains($srvHost)) { $pox.Add($srvHost) }
        foreach ($base in $pox) {
            $classic = Get-WscAutodiscoverPox -Context $Context -BaseUrl $base
            if ($classic.Url) { $found = [pscustomobject]@{ Url = $classic.Url; Method = 'classic Autodiscover (POX)'; Request = $classic.Request }; break }
            [void]$attempts.Add("POX $($base): $($classic.Error)")
            if ($classic.Error -like 'HTTP 401*') { break }
        }
    }
    $details = [ordered]@{ Mailbox = $mailbox; Attempts = $attempts -join ' | ' }
    if ($found) {
        $details.Method = $found.Method
        $details.Request = $found.Request
        $details.EwsUrl = $found.Url
        if ($found.Url -notmatch '/EWS/Exchange\.asmx$') {
            Add-WscStep $Context $stage 'EWS URL' Failed "Autodiscover answered $($found.Url), which is not an EWS URL (/EWS/Exchange.asmx): check the ExternalUrl of the EWS virtual directory." $details
            $Context.Stop = $true
            return
        }
        Set-WscEwsUrl -Context $Context -Url $found.Url -Source "Autodiscover ($($found.Method))"
        $online = if ($Context.Endpoints.ExchangeOnline) { ' The mailbox is in Exchange Online: sign in with Entra ID.' } else { '' }
        if ($typed -and $typed.TrimEnd('/') -ine $found.Url.TrimEnd('/')) {
            Add-WscStep $Context $stage 'EWS URL' Warning "Autodiscover ($($found.Method)) gives $($found.Url) for $mailbox, not $($typed): the test uses the Autodiscover URL, like Outlook.$online" $details
        }
        else {
            Add-WscStep $Context $stage 'EWS URL' Passed "Autodiscover ($($found.Method)) gives the EWS URL of $mailbox ($($found.Url)).$online" $details
        }
        if ($Context.Endpoints.ExchangeOnline -and [string]$cfg.Authentication -eq 'OAuth' -and [string]$cfg.Authority -eq 'ADFS') {
            Add-WscStep $Context $stage 'Authorization server' Failed 'The mailbox is in Exchange Online, which accepts only Entra ID tokens: run again with -Authority EntraID (or Auto).' ([ordered]@{ EwsUrl = $found.Url })
            $Context.Stop = $true
        }
        return
    }
    if ($typed) {
        Set-WscEwsUrl -Context $Context -Url $typed -Source 'Configuration (Autodiscover gave no answer)'
        $details.Typed = $typed
        Add-WscStep $Context $stage 'EWS URL' Warning ("No usable Autodiscover answer for $mailbox ($($attempts -join '; ')): Outlook and the EWS applications that use Autodiscover would not find EWS. " +
            "The test goes on with $typed (Target.EwsUrl). Publish autodiscover.$domain (or the SRV record), or use -Discovery Manual.") $details
        return
    }
    Add-WscStep $Context $stage 'EWS URL' Failed "No usable Autodiscover answer for $mailbox ($($attempts -join '; ')): publish Autodiscover (autodiscover.$domain), or give the URL with -EwsUrl and -Discovery Manual." $details
    $Context.Stop = $true
}

#endregion

#region Prerequisites ------------------------------------------------------------------------------

function Add-WscTlsCheck {
    <# Certificate a server presents (direct TLS connection), added to the trace and checked: trusted, expiry. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage, [Parameter(Mandatory = $true)][string]$HostName, [Parameter(Mandatory = $true)][int]$Port)

    $cfg = $Context.Config
    $cert = Get-WscTlsCertificate -HostName $HostName -Port $Port -TimeoutSeconds ([Math]::Min(15, [int]$cfg.HttpTimeoutSeconds))
    $name = "TLS certificate ($HostName)"
    $details = [ordered]@{ Host = "$($HostName):$Port"; Subject = $cert.Subject; Issuer = $cert.Issuer; NotAfterUtc = $cert.NotAfterUtc; DaysLeft = $cert.DaysLeft; Protocol = $cert.Protocol; Error = $cert.Error }
    $received = if (-not $cert.Reachable) { "No TLS connection: $($cert.Error)" }
    elseif ($cert.Interrupted) { "No certificate: the connection was closed during the TLS handshake ($($cert.Error))" }
    else { "Certificate`nSubject: $($cert.Subject)`nIssuer: $($cert.Issuer)`nValid until: $($cert.NotAfterUtc) ($($cert.DaysLeft) day(s))`nProtocol: $($cert.Protocol)`nTrusted by this computer: $(if ($cert.Valid) { 'yes' } else { "no - $($cert.Error)" })" }
    Add-WscTraceEntry -Method 'TLS' -Url "tls://$($HostName):$Port" -Request "TLS handshake (ClientHello)`nServer name (SNI): $HostName`nPort: $Port" -Response $received -Note 'Direct TLS connection (no HTTP request): the certificate the server presents.'
    if (-not $cert.Reachable) { Add-WscStep $Context $Stage $name Warning "No direct TLS connection ($($cert.Error)). Expected behind a proxy (the HTTPS checks use the system proxy); otherwise check DNS and the firewall." $details }
    elseif ($cert.Interrupted) {
        Add-WscStep $Context $Stage $name Failed ("TLS handshake interrupted before the server sent its certificate ($($cert.Error)): the certificate is not in question. " +
            'Something on the path closes the connection: a firewall or NSG that filters the source address, a reverse proxy, or a VPN or Global Secure Access client that tunnels this address.') $details
    }
    elseif (-not $cert.Valid) { Add-WscStep $Context $Stage $name Failed "Certificate not trusted by this computer: $($cert.Error)" $details }
    elseif ($cert.DaysLeft -lt [int]$cfg.CertificateWarningDays) { Add-WscStep $Context $Stage $name Warning "Certificate expires in $($cert.DaysLeft) day(s) ($($cert.NotAfterUtc))." $details }
    else { Add-WscStep $Context $Stage $name Passed "Certificate trusted, valid $($cert.DaysLeft) more day(s), $($cert.Protocol)." $details }
}

function Set-WscAuthority {
    <# Applies the authorization server found (AD FS root or Entra ID tenant) to the run and recomputes the endpoints. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][ValidateSet('ADFS', 'EntraID')][string]$Kind, [string]$AdfsRoot, [string]$TenantId, [string]$Source)

    $cfg = $Context.Config
    $cfg.Authority = $Kind
    if ($Kind -eq 'ADFS') { $cfg.AdfsUrl = $AdfsRoot }
    elseif ($TenantId) { $cfg.TenantId = $TenantId }
    $Context.Endpoints = Resolve-WscEndpoints -Configuration $cfg
    if ($Source) { $Context.AuthoritySource = $Source }
}

function Add-WscEntraChecks {
    <# Entra ID before any sign-in: the tenant (OpenID configuration) and the user realm. Returns $false when the tenant is not found. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage, [string]$Hint, [string]$Source)

    $cfg = $Context.Config
    $user = if ([string]$cfg.SignInUser -match '@') { [string]$cfg.SignInUser } else { [string]$cfg.Mailbox }
    $domain = $user.Split('@')[-1]
    if ([string]$cfg.TenantId) { $name = [string]$cfg.TenantId; $from = 'Target.TenantId' }
    elseif ($Hint -and $Hint -notin 'common', 'organizations', 'consumers') { $name = $Hint; $from = 'the authorization URL of Exchange' }
    else { $name = $domain; $from = "the domain of $user" }
    $tenant = try { Resolve-WscEntraTenant -Context $Context -Name $name } catch { [pscustomobject]@{ TenantId = $null; Issuer = $null; Url = $null; Error = $_.Exception.Message } }
    $details = [ordered]@{ Tenant = $name; From = $from; TenantId = $tenant.TenantId; Issuer = $tenant.Issuer; Url = $tenant.Url; Error = $tenant.Error }
    if (-not $tenant.TenantId) {
        Add-WscStep $Context $Stage 'Entra ID tenant' Failed "Entra ID does not know the tenant $name ($($tenant.Error)): check the domain (a verified domain of the tenant) or set Target.TenantId." $details
        return $false
    }
    $Context.TenantId = $tenant.TenantId
    Set-WscAuthority -Context $Context -Kind EntraID -TenantId $tenant.TenantId -Source $Source
    Add-WscStep $Context $Stage 'Entra ID tenant' Passed "Tenant $($tenant.TenantId), found from $($from): the token is requested from $($Context.Endpoints.EntraRoot)." $details
    if ([string]$cfg.Context -eq 'Application') { return $true }
    $realm = try { Get-WscUserRealm -Context $Context -User $user } catch { [pscustomobject]@{ NameSpaceType = $null; DomainName = $null; FederationBrandName = $null; AuthUrl = $null } }
    $realmDetails = [ordered]@{ User = $user; NameSpaceType = $realm.NameSpaceType; DomainName = $realm.DomainName; AuthUrl = $realm.AuthUrl }
    switch ($realm.NameSpaceType) {
        'Managed' { Add-WscStep $Context $Stage 'User realm' Passed "Entra ID signs in $user itself (managed domain $($realm.DomainName))." $realmDetails }
        'Federated' { Add-WscStep $Context $Stage 'User realm' Passed "Entra ID sends $user to the federation server $($realm.AuthUrl) (federated domain $($realm.DomainName)): the password is typed there." $realmDetails }
        default { Add-WscStep $Context $Stage 'User realm' Warning "Entra ID does not know the domain of $user (NameSpaceType $(if ($realm.NameSpaceType) { $realm.NameSpaceType } else { 'not returned' })): sign in with the UPN of the user (Target.SignInUser)." $realmDetails }
    }
    return $true
}

function Add-WscAdfsMetadataCheck {
    <# OpenID configuration published by AD FS, then the certificate of AD FS. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage)

    $ep = $Context.Endpoints
    try {
        $meta = Invoke-WscHttpGet -HttpClient $Context.HttpClient -Uri $ep.MetadataEndpoint
        $details = [ordered]@{ Issuer = [string](Get-WscField $meta 'issuer'); TokenEndpoint = [string](Get-WscField $meta 'token_endpoint'); DeviceAuthorizationEndpoint = [string](Get-WscField $meta 'device_authorization_endpoint') }
        if ($details.TokenEndpoint) { Add-WscStep $Context $Stage 'AD FS metadata' Passed "OpenID configuration published by AD FS (issuer $($details.Issuer))." $details }
        else { Add-WscStep $Context $Stage 'AD FS metadata' Warning 'The OpenID configuration of AD FS has no token_endpoint.' $details }
    }
    catch {
        Add-WscStep $Context $Stage 'AD FS metadata' Warning "OpenID configuration not readable ($($_.Exception.Message)). Sign-in can still work if this endpoint is disabled in AD FS." ([ordered]@{ Url = $ep.MetadataEndpoint })
    }
    if ($ep.AdfsHost -and ($ep.AdfsHost -ine $ep.EwsHost -or $ep.AdfsPort -ne $ep.EwsPort)) { Add-WscTlsCheck -Context $Context -Stage $Stage -HostName $ep.AdfsHost -Port $ep.AdfsPort }
}

function Test-WscMailboxChallenge {
    <#
        What Exchange tells a client about OAuth for the mailbox: EWS request with an empty bearer header
        and X-AnchorMailbox, as Outlook sends it. The authorization URL is returned only when the
        authentication policy of the user allows modern authentication for EWS.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $mailbox = if ([string]$Context.Config.SignInUser -match '@') { [string]$Context.Config.SignInUser } else { [string]$Context.Config.Mailbox }
    $answer = Invoke-WscEws -Context $Context -Operation 'GetFolder' -Body (New-WscGetFolderBody -Context $Context) -Credentials EmptyBearer -Anchor $mailbox -ExtraHeaders @{ 'X-User-Identity' = $mailbox } -NoImpersonation
    $info = Get-WscChallengeInfo -Challenges @(Get-WscField $answer.Response 'Challenges')
    $details = [ordered]@{ Request = "POST with an empty bearer header, X-AnchorMailbox and X-User-Identity $mailbox"; HttpStatus = $answer.HttpStatus; Schemes = $info.Schemes -join ', '; AuthorizationUri = $info.AuthorizationUri; IssuerKind = $info.IssuerKind; TrustedIssuers = $info.TrustedIssuers -join ', '; Diagnostics = $answer.Diagnostics }
    $outcome = { param([string]$Status, [string]$Message) [pscustomobject]@{ Status = $Status; Message = $Message; Details = $details; AuthorizationUri = $info.AuthorizationUri; TrustedIssuers = @($info.TrustedIssuers) } }
    $advertised = Get-WscAuthorityInfo -Uri ([string]$info.AuthorizationUri)
    if ($answer.HttpStatus -ne 401) { return & $outcome 'Failed' "HTTP $($answer.HttpStatus) instead of 401 to a request without a token: check the URL and the publishing (reverse proxy, load balancer)." }
    $mismatch = if ($info.Bearer -and $advertised.Kind -ne 'None') { Test-WscExpectedAuthority -Context $Context -Advertised $advertised } else { $null }
    if ($mismatch) { return & $outcome 'Warning' "Exchange offers OAuth to $mailbox, but: $mismatch" }
    if ($info.Bearer -and $advertised.Kind -ne 'None') {
        $how = if ($advertised.Kind -eq 'EntraID') { if ($Context.Endpoints.ExchangeOnline) { 'Entra ID (Exchange Online)' } else { 'Entra ID (hybrid modern authentication)' } } else { "AD FS ($($advertised.Host))" }
        return & $outcome 'Passed' "Exchange offers OAuth for EWS to $mailbox and names the authorization server: $how ($($info.AuthorizationUri))."
    }
    if ($answer.Diagnostics -match 'oauth_not_available') {
        return & $outcome 'Failed' ("Exchange does not offer OAuth to $mailbox ($($answer.Diagnostics)): the authentication policy of the user blocks modern authentication for EWS (BlockModernAuthWebServices), or the domain is not an accepted domain. " +
            'Check Get-User <user> | Format-List AuthenticationPolicy and Get-AuthenticationPolicy | Format-List Name, BlockModernAuthWebServices.')
    }
    if ($info.Bearer) { return & $outcome 'Warning' "Exchange accepts OAuth for $mailbox but gives no authorization URL: clients cannot find the authorization server. Check Get-AuthServer (IsDefaultAuthorizationEndpoint)." }
    return & $outcome 'Failed' "No OAuth challenge for $mailbox (schemes: $($details.Schemes)): check OAuth on the EWS virtual directory (Get-WebServicesVirtualDirectory | Format-List OAuthAuthentication), the authorization server (Get-AuthServer) and the reverse proxy."
}

function Add-WscSchemesCheck {
    <# The schemes EWS offers to an anonymous request, compared with the authentication of the run. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage)

    $answer = Invoke-WscEws -Context $Context -Operation 'GetFolder' -Body (New-WscGetFolderBody -Context $Context) -Credentials None -NoImpersonation
    $info = Get-WscChallengeInfo -Challenges @(Get-WscField $answer.Response 'Challenges')
    $schemes = @($info.Schemes)
    $Context.OfferedSchemes = $schemes
    $details = [ordered]@{ HttpStatus = $answer.HttpStatus; Schemes = $schemes -join ', '; AuthorizationUri = $info.AuthorizationUri; TrustedIssuers = $info.TrustedIssuers -join ', '; Servers = $answer.Server }
    $auth = [string]$Context.Config.Authentication
    $list = if ($schemes.Count) { $schemes -join ', ' } else { 'none' }
    if ($answer.HttpStatus -eq 200) { Add-WscStep $Context $Stage 'Authentication offered' Warning 'EWS answered an anonymous request with HTTP 200: anonymous access is not expected.' $details; return }
    if ($answer.HttpStatus -ne 401) {
        Add-WscStep $Context $Stage 'Authentication offered' Warning "The anonymous request got HTTP $($answer.HttpStatus) instead of 401: this probe cannot tell which authentication EWS offers; the sign-in decides." $details
        return
    }
    $wanted = switch ($auth) { 'Basic' { @('Basic') } 'Windows' { if ([string]$Context.Config.WindowsPackage -eq 'NTLM') { @('NTLM', 'Negotiate') } else { @('Negotiate') } } default { @('Bearer') } }
    $offered = @($wanted | Where-Object { $schemes -contains $_ }).Count
    $name = switch ($auth) { 'Basic' { 'Basic' } 'Windows' { "Windows ($($wanted -join ' or '))" } default { 'OAuth (Bearer)' } }
    if ($offered) { Add-WscStep $Context $Stage 'Authentication offered' Passed "EWS offers $name (schemes: $list)." $details; return }
    if ($auth -eq 'OAuth') {
        # Exchange Server 2019 CU13+ and SE give their bearer challenge only to an empty bearer header: the next check decides.
        Add-WscStep $Context $Stage 'Authentication offered' Passed "Schemes offered to an anonymous request: $list. Exchange gives its OAuth challenge to an empty bearer header (next check)." $details
        return
    }
    $fix = switch ($auth) {
        'Basic' { if ($Context.Endpoints.ExchangeOnline) { 'Exchange Online no longer accepts Basic for EWS: use OAuth with Entra ID.' } else { 'Basic is disabled on the EWS virtual directory (Get-WebServicesVirtualDirectory | Format-List Server, BasicAuthentication) or removed by the reverse proxy.' } }
        default { if ($Context.Endpoints.ExchangeOnline) { 'Exchange Online does not offer Windows authentication: use OAuth with Entra ID.' } else { 'Windows authentication is disabled on the EWS virtual directory (WindowsAuthentication) or not let through by the reverse proxy (pre-authentication, NTLM not relayed).' } }
    }
    Add-WscStep $Context $Stage 'Authentication offered' Failed "EWS does not offer $name (schemes: $list): $fix" $details
}

function Add-WscNtlmChallengeCheck {
    <#
        An NTLM negotiate message sent without credentials: the challenge of Exchange names the server, its
        NetBIOS and DNS domain (the Kerberos realm) and its Windows version. Nothing is authenticated.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage)

    $options = [Net.Security.NegotiateAuthenticationClientOptions]::new()
    $options.Package = 'NTLM'
    $options.Credential = [Net.NetworkCredential]::new('wsc-probe', [guid]::NewGuid().ToString(), 'WSC')
    $options.TargetName = "HTTP/$($Context.Endpoints.EwsHost)"
    $client = [Net.Security.NegotiateAuthentication]::new($options)
    try {
        $status = [Net.Security.NegotiateAuthenticationStatusCode]::GenericFailure
        $type1 = $client.GetOutgoingBlob([NullString]::Value, [ref]$status)
    }
    finally { $client.Dispose() }
    $scheme = if (@($Context.OfferedSchemes) -contains 'NTLM') { 'NTLM' } else { 'Negotiate' }
    $answer = Invoke-WscEws -Context $Context -Operation 'GetFolder' -Body (New-WscGetFolderBody -Context $Context) -Credentials None -NoImpersonation -ExtraHeaders @{ Authorization = "$scheme $type1" }
    $token = $null
    foreach ($c in @(Get-WscField $answer.Response 'Challenges')) { $m = [regex]::Match([string]$c, "^\s*$scheme\s+(\S+)", 'IgnoreCase'); if ($m.Success) { $token = $m.Groups[1].Value } }
    $ntlm = if ($token) { (Get-WscNegotiateInfo -Base64 $token).Ntlm } else { $null }
    if (-not $ntlm -or $ntlm.Type -ne 2) {
        Add-WscStep $Context $Stage 'NTLM challenge' Warning "EWS did not answer the NTLM negotiate message with a challenge (HTTP $($answer.HttpStatus)): NTLM is not offered, or a reverse proxy does not relay it." ([ordered]@{ HttpStatus = $answer.HttpStatus; Scheme = $scheme })
        return $null
    }
    $details = [ordered]@{ Scheme = $scheme; Server = $ntlm.PSObject.Properties['DnsComputer'].Value; NetBiosDomain = $ntlm.PSObject.Properties['NetBiosDomain'].Value; DnsDomain = $ntlm.PSObject.Properties['DnsDomain'].Value; DnsForest = $ntlm.PSObject.Properties['DnsForest'].Value; ServerVersion = $ntlm.PSObject.Properties['ServerVersion'].Value }
    $Context.NtlmServer = $details
    Add-WscStep $Context $Stage 'NTLM challenge' Passed "Exchange answers the NTLM handshake: server $($details.Server), domain $($details.NetBiosDomain) ($($details.DnsDomain)), $($details.ServerVersion). No credentials were sent." $details
    return $details
}

function Add-WscKerberosCheck {
    <# Kerberos from this computer: domain membership, KDC of the realm (SRV record, port 88), SPN of the EWS host. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage, [string]$Realm)

    $cfg = $Context.Config
    $user = if ($Context.Credential) { $Context.Credential.UserName } else { "$([Environment]::UserDomainName)\$([Environment]::UserName)" }
    if (-not $Realm) { $Realm = if ($user -match '@(.+)$') { $Matches[1] } else { ([string]$cfg.Mailbox).Split('@')[-1] } }
    $k = Test-WscKerberosRealm -Domain $Realm
    $spn = "HTTP/$($Context.Endpoints.EwsHost)"
    $details = [ordered]@{ Realm = $Realm; Kdc = $k.Kdc; Port88 = $k.Port88; Spn = $spn; DomainJoined = $k.DomainJoined; ComputerDomain = $k.ComputerDomain; User = $user; Error = $k.Error }
    $package = [string]$cfg.WindowsPackage
    if ($k.Port88) {
        $member = if ($k.DomainJoined) { "this computer is in the domain $($k.ComputerDomain)" } else { 'this computer is not in the domain, Kerberos works with -Credential' }
        Add-WscStep $Context $Stage 'Kerberos realm' Passed "A KDC of $Realm answers ($($k.Kdc):88) and $($member): Kerberos is possible if the SPN $spn is registered (alternate service account for a load-balanced name)." $details
        return
    }
    $why = "No KDC of $Realm in reach from this computer ($($k.Error))"
    if ($package -eq 'Kerberos') {
        Add-WscStep $Context $Stage 'Kerberos realm' Failed "$($why): Kerberos cannot work here. Run from a computer in the domain or with a VPN to the domain controllers, or use -WindowsPackage NTLM." $details
    }
    elseif ($package -eq 'Negotiate') {
        Add-WscStep $Context $Stage 'Kerberos realm' Warning "$($why): Negotiate will fall back to NTLM, like Windows. Use a computer that reaches a domain controller to test Kerberos." $details
    }
    else {
        Add-WscStep $Context $Stage 'Kerberos realm' Passed "$($why); not needed: the run uses NTLM." $details
    }
}

function Invoke-WscStageDiscovery {
    <#
        Checks that need no sign-in and change nothing. A failed check does not stop the scenario.
          OAuth    authorization server, certificates, schemes offered, OAuth for the mailbox, trusted tenant, forged token
          Basic    certificate, Basic offered, wrong password refused (user that does not exist)
          Windows  certificate, Negotiate or NTLM offered, NTLM challenge (server, domain), Kerberos realm
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Discovery'
    $cfg = $Context.Config
    $auth = [string]$cfg.Authentication
    if ($auth -eq 'OAuth') {
        if ([string]$cfg.Authority -eq 'ADFS') { Add-WscAdfsMetadataCheck -Context $Context -Stage $stage }
        elseif ([string]$cfg.Authority -eq 'EntraID') { [void](Add-WscEntraChecks -Context $Context -Stage $stage -Source 'Configuration') }
    }
    if ($Context.Protocol -eq 'Graph') { Invoke-WscStageDiscoveryGraph -Context $Context; return }
    $ep = $Context.Endpoints
    Add-WscTlsCheck -Context $Context -Stage $stage -HostName $ep.EwsHost -Port $ep.EwsPort
    try {
        Add-WscSchemesCheck -Context $Context -Stage $stage
        switch ($auth) {
            'OAuth' {
                $challenge = Test-WscMailboxChallenge -Context $Context
                Add-WscStep $Context $stage 'OAuth for the mailbox' $challenge.Status $challenge.Message $challenge.Details
                if ([string]$cfg.Authority -eq 'Auto') {
                    $server = Get-WscAuthorityInfo -Uri ([string]$challenge.AuthorizationUri)
                    if ($server.Kind -eq 'ADFS') { Set-WscAuthority -Context $Context -Kind ADFS -AdfsRoot $server.AdfsRoot -Source 'Exchange challenge (authorization_uri)'; Add-WscAdfsMetadataCheck -Context $Context -Stage $stage }
                    elseif ($server.Kind -eq 'EntraID') { [void](Add-WscEntraChecks -Context $Context -Stage $stage -Hint $server.Tenant -Source 'Exchange challenge (authorization_uri)') }
                }
                if ($Context.TenantId -and @($challenge.TrustedIssuers).Count) {
                    $trusted = @($challenge.TrustedIssuers | Where-Object { $_ -ilike "*@$($Context.TenantId)" -or $_ -like '*@`*' }).Count
                    $d = [ordered]@{ TrustedIssuers = $challenge.TrustedIssuers -join ', '; TenantId = $Context.TenantId }
                    if ($trusted) { Add-WscStep $Context $stage 'Tenant trusted by Exchange' Passed "Exchange trusts the Entra ID tokens of tenant $($Context.TenantId) (trusted_issuers of its challenge)." $d }
                    else { Add-WscStep $Context $stage 'Tenant trusted by Exchange' Warning "Exchange trusts the tokens of $($challenge.TrustedIssuers -join ', '), not of tenant $($Context.TenantId): it will refuse the token. Check Get-AuthServer (EvoSts) and run the Hybrid Configuration Wizard again." $d }
                }
                $forged = Invoke-WscEws -Context $Context -Operation 'GetFolder' -Body (New-WscGetFolderBody -Context $Context) -Credentials Token -Token $script:InvalidToken -NoImpersonation
                $d = [ordered]@{ HttpStatus = $forged.HttpStatus; Diagnostics = $forged.Diagnostics }
                if ($forged.HttpStatus -eq 401) { Add-WscStep $Context $stage 'Forged token' Passed 'A forged bearer token is refused (HTTP 401).' $d }
                elseif ($forged.HttpStatus -ge 200 -and $forged.HttpStatus -lt 300) { Add-WscStep $Context $stage 'Forged token' Failed "Exchange accepted a forged bearer token (HTTP $($forged.HttpStatus)): investigate the publishing chain immediately." $d }
                else { Add-WscStep $Context $stage 'Forged token' Warning "A forged bearer token returned HTTP $($forged.HttpStatus) (401 expected)." $d }
            }
            'Basic' {
                $mailbox = [string]$cfg.Mailbox
                $user = '{0}{1}@{2}' -f $script:InvalidBasicUserPrefix, [guid]::NewGuid().ToString('N').Substring(0, 12), $mailbox.Split('@')[-1]
                $wrong = [pscredential]::new($user, (ConvertTo-SecureString ([guid]::NewGuid().ToString()) -AsPlainText -Force))
                $answer = Invoke-WscEws -Context $Context -Operation 'GetFolder' -Body (New-WscGetFolderBody -Context $Context) -Credentials Basic -BasicCredential $wrong -NoImpersonation
                $d = [ordered]@{ User = $user; HttpStatus = $answer.HttpStatus; Diagnostics = $answer.Diagnostics }
                if ($answer.HttpStatus -eq 401) { Add-WscStep $Context $stage 'Wrong password' Passed "A wrong user name and password are refused (HTTP 401). The test uses a user that does not exist ($user): no account can be locked." $d }
                elseif ($answer.HttpStatus -ge 200 -and $answer.HttpStatus -lt 300) { Add-WscStep $Context $stage 'Wrong password' Failed "Exchange accepted a user that does not exist (HTTP $($answer.HttpStatus)): investigate the publishing chain immediately." $d }
                else { Add-WscStep $Context $stage 'Wrong password' Warning "A wrong user name and password returned HTTP $($answer.HttpStatus) (401 expected)." $d }
            }
            'Windows' {
                $ntlm = Add-WscNtlmChallengeCheck -Context $Context -Stage $stage
                if ([string]$cfg.WindowsPackage -ne 'NTLM' -or -not $ntlm) {
                    $realm = if ($ntlm -and $ntlm.DnsDomain) { [string]$ntlm.DnsDomain } else { $null }
                    Add-WscKerberosCheck -Context $Context -Stage $stage -Realm $realm
                }
            }
        }
    }
    catch {
        Add-WscStep $Context $stage 'EWS reachable' Failed "EWS not reachable: $($_.Exception.Message)" ([ordered]@{ Url = $ep.EwsUrl })
    }
}

#endregion
