<#
.SYNOPSIS
    Web Services Client for Exchange - helpers and OAuth building blocks (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    Shared with EAS OAuth Mailbox (same author, same behaviour): object helpers, waits that keep the
    window responsive, traced GET and form POST, TLS certificate of a server, WWW-Authenticate
    challenges, authorization server named by Exchange, device-code sign-in, token claims, Entra ID
    tenant and user realm, and the answers of the AD FS and Entra ID sign-in pages.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

function Get-WscField {
    <# Property of an object or key of a dictionary; $null when absent (safe under strict mode). #>
    param([AllowNull()][object]$Object, [Parameter(Mandatory = $true)][string]$Name)

    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function ConvertTo-WscBase64Url {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)
    return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function ConvertFrom-WscBase64Url {
    param([Parameter(Mandatory = $true)][string]$Text)
    $value = $Text.Replace('-', '+').Replace('_', '/')
    switch ($value.Length % 4) { 2 { $value += '==' } 3 { $value += '=' } }
    return [Convert]::FromBase64String($value)
}

function Invoke-WscUiPump {
    <# Keeps the GUI responsive during long waits (no effect on the command line). #>
    if ($script:Ui -and $script:Ui.Pump) { & $script:Ui.Pump }
}

function Assert-WscNotCancelled {
    if ($script:Ui -and $script:Ui.Cancel) { throw 'Cancelled by the operator.' }
}

function Wait-WscSeconds {
    param([Parameter(Mandatory = $true)][double]$Seconds)
    $until = [DateTimeOffset]::UtcNow.AddSeconds($Seconds)
    while ([DateTimeOffset]::UtcNow -lt $until) {
        Invoke-WscUiPump
        Assert-WscNotCancelled
        Start-Sleep -Milliseconds 200
    }
}

function Invoke-WscHttpGet {
    <# JSON GET without credentials (AD FS metadata), traced. Throws when the answer is not HTTP 200 with JSON. #>
    param([Parameter(Mandatory = $true)][Net.Http.HttpClient]$HttpClient, [Parameter(Mandatory = $true)][string]$Uri)

    $response = Invoke-WscWebRequest -HttpClient $HttpClient -Uri $Uri -UserAgent $script:WscUserAgent -Headers @{ Accept = 'application/json' }
    if ($response.StatusCode -ne 200) { throw "HTTP $($response.StatusCode)" }
    return ($response.Content | ConvertFrom-Json -ErrorAction Stop)
}

function Invoke-WscWebRequest {
    <#
        GET without credentials and without following redirects (Autodiscover, AD FS pages), traced.
        Returns StatusCode, Location and Content (text).
    #>
    param(
        [Parameter(Mandatory = $true)][Net.Http.HttpClient]$HttpClient,
        [Parameter(Mandatory = $true)][string]$Uri,
        [string]$UserAgent,
        [hashtable]$Headers
    )

    $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Get, $Uri)
    try {
        if ($UserAgent) { $null = $request.Headers.TryAddWithoutValidation('User-Agent', $UserAgent) }
        if ($Headers) { foreach ($name in $Headers.Keys) { $null = $request.Headers.TryAddWithoutValidation([string]$name, [string]$Headers[$name]) } }
        $response = Invoke-WscHttp -HttpClient $HttpClient -Request $request
        [pscustomobject]@{
            StatusCode = $response.StatusCode
            Location   = if ($response.Headers -and $response.Headers.ContainsKey('Location')) { [string]$response.Headers['Location'] } else { $null }
            Content    = if ($response.Body) { [Text.Encoding]::UTF8.GetString([byte[]]$response.Body) } else { '' }
        }
    }
    finally {
        $request.Dispose()
    }
}

function Invoke-WscFormPost {
    <# POST of form fields to AD FS (device code, token), traced. Returns StatusCode and the JSON answer (or $null). #>
    param(
        [Parameter(Mandatory = $true)][Net.Http.HttpClient]$HttpClient,
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Fields,
        [string]$UserAgent,
        [string]$Collapse
    )

    $pairs = [Collections.Generic.List[Collections.Generic.KeyValuePair[string, string]]]::new()
    foreach ($key in $Fields.Keys) { $pairs.Add([Collections.Generic.KeyValuePair[string, string]]::new([string]$key, [string]$Fields[$key])) }
    $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Post, $Uri)
    try {
        $request.Content = [Net.Http.FormUrlEncodedContent]::new($pairs)
        $null = $request.Headers.TryAddWithoutValidation('User-Agent', $(if ($UserAgent) { $UserAgent } else { $script:WscUserAgent }))
        $null = $request.Headers.TryAddWithoutValidation('Accept', 'application/json')
        $response = Invoke-WscHttp -HttpClient $HttpClient -Request $request -Collapse $Collapse
        $json = $null
        if ($response.Body -and $response.Body.Length) {
            try { $json = [Text.Encoding]::UTF8.GetString([byte[]]$response.Body) | ConvertFrom-Json -ErrorAction Stop } catch { $json = $null }
        }
        [pscustomobject]@{ StatusCode = $response.StatusCode; Json = $json }
    }
    finally {
        $request.Dispose()
    }
}

function Get-WscTlsCertificate {
    <#
        Server certificate of HostName:Port with the default Windows validation. Reachable = $false
        when no direct TCP connection is possible (proxy, firewall); Valid = $false when the TLS
        handshake fails: Interrupted = $true when the connection was closed before the server sent a
        certificate (firewall, NSG, proxy or VPN client on the path), $false when the certificate was
        received and rejected.
    #>
    param([Parameter(Mandatory = $true)][string]$HostName, [int]$Port = 443, [int]$TimeoutSeconds = 10)

    $result = [ordered]@{ HostName = $HostName; Port = $Port; Reachable = $false; Valid = $false; Interrupted = $false; Subject = $null; Issuer = $null; NotAfterUtc = $null; DaysLeft = $null; Protocol = $null; Error = $null }
    $tcp = [Net.Sockets.TcpClient]::new()
    try {
        try {
            if (-not $tcp.ConnectAsync($HostName, $Port).Wait([TimeSpan]::FromSeconds($TimeoutSeconds))) {
                $result.Error = "no TCP connection within $TimeoutSeconds s"
                return [pscustomobject]$result
            }
        }
        catch {
            $inner = $_.Exception; while ($inner.InnerException) { $inner = $inner.InnerException }
            $result.Error = $inner.Message
            return [pscustomobject]$result
        }
        $result.Reachable = $true
        $tcp.ReceiveTimeout = $TimeoutSeconds * 1000
        $tcp.SendTimeout = $TimeoutSeconds * 1000
        # Same decision as Windows (no policy error), and a trace of whether a certificate arrived at all.
        $received = @{ Certificate = $false; Errors = $null }
        $validate = [Net.Security.RemoteCertificateValidationCallback]{
            param($sender, $certificate, $chain, $errors)
            if ($certificate) { $received.Certificate = $true }
            if ($errors -ne [Net.Security.SslPolicyErrors]::None) {
                $status = @($chain.ChainStatus | ForEach-Object { [string]$_.Status } | Where-Object { $_ -ne 'NoError' } | Select-Object -Unique)
                $received.Errors = "The remote certificate is invalid: $errors$(if ($status) { " ($($status -join ', '))" })"
            }
            return $errors -eq [Net.Security.SslPolicyErrors]::None
        }.GetNewClosure()
        $ssl = [Net.Security.SslStream]::new($tcp.GetStream(), $false, $validate)
        try {
            try {
                $ssl.AuthenticateAsClient($HostName)
            }
            catch {
                $inner = $_.Exception; while ($inner.InnerException) { $inner = $inner.InnerException }
                $result.Error = if ($received.Errors) { $received.Errors } else { $inner.Message }
                $result.Interrupted = -not $received.Certificate
                return [pscustomobject]$result
            }
            $certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($ssl.RemoteCertificate)
            $notAfter = $certificate.NotAfter.ToUniversalTime()
            $result.Valid = $true
            $result.Subject = $certificate.Subject
            $result.Issuer = $certificate.Issuer
            $result.NotAfterUtc = $notAfter.ToString('yyyy-MM-ddTHH:mm:ssZ')
            $result.DaysLeft = [int][Math]::Floor(($notAfter - [DateTime]::UtcNow).TotalDays)
            $result.Protocol = [string]$ssl.SslProtocol
            return [pscustomobject]$result
        }
        finally {
            $ssl.Dispose()
        }
    }
    finally {
        $tcp.Dispose()
    }
}

function Get-WscChallengeInfo {
    <# Schemes of the WWW-Authenticate challenges, and the parameters of the Bearer challenge. #>
    param([AllowEmptyCollection()][string[]]$Challenges)

    $schemes = [Collections.Generic.List[string]]::new()
    $parameters = @{}
    $bearer = $false
    foreach ($challenge in @($Challenges)) {
        if ([string]::IsNullOrWhiteSpace($challenge)) { continue }
        $scheme = ($challenge.Trim() -split '\s+', 2)[0].TrimEnd(',')
        if (-not $schemes.Contains($scheme)) { [void]$schemes.Add($scheme) }
        if ($scheme -ieq 'Bearer') {
            $bearer = $true
            foreach ($m in [regex]::Matches($challenge, '([A-Za-z_]+)\s*=\s*"([^"]*)"')) { $parameters[$m.Groups[1].Value.ToLowerInvariant()] = $m.Groups[2].Value }
        }
    }
    [pscustomobject]@{
        Schemes          = @($schemes)
        Bearer           = $bearer
        AuthorizationUri = $parameters['authorization_uri']
        IssuerKind       = $parameters['issuer_kind']
        # Entra ID (hybrid modern authentication): <token service ID>@<tenant ID>, comma separated.
        TrustedIssuers   = @(([string]$parameters['trusted_issuers']).Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        Error            = $parameters['error']
    }
}

function Get-WscAuthorityInfo {
    <#
        What an authorization URL given by Exchange points to: AD FS (https://<host>/adfs/oauth2/authorize,
        AdfsRoot) or Entra ID (login.microsoftonline.com, login.windows.net..., Tenant: the path segment:
        common, organizations, a tenant ID or a domain). Kind None when there is no URL, Other otherwise.
    #>
    param([AllowEmptyString()][string]$Uri)

    $info = [ordered]@{ Kind = 'None'; Uri = $Uri; Host = $null; AdfsRoot = $null; Tenant = $null; Name = 'no authorization server' }
    $parsed = $null
    if (-not $Uri -or -not [Uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$parsed)) { return [pscustomobject]$info }
    $info.Host = $parsed.Host
    $adfs = [regex]::Match($Uri, '^(https://[^/?#]+/adfs)/oauth2/authorize', 'IgnoreCase')
    if ($adfs.Success) {
        $info.Kind = 'ADFS'; $info.AdfsRoot = $adfs.Groups[1].Value; $info.Name = "AD FS ($($parsed.Host))"
    }
    elseif ($script:Entra.Hosts -contains $parsed.Host.ToLowerInvariant()) {
        $info.Kind = 'EntraID'
        $info.Tenant = ($parsed.AbsolutePath.Trim('/') -split '/')[0]
        $info.Name = "Entra ID ($($parsed.Host))"
    }
    else {
        $info.Kind = 'Other'; $info.Name = $parsed.Host
    }
    [pscustomobject]$info
}

function Test-WscExpectedAuthority {
    <#
        Compares the authorization server named by Exchange with the one the test expects (Target.Authority).
        Returns $null when they match (or when the test takes the one of Exchange: Auto), otherwise
        the text of the warning.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][pscustomobject]$Advertised)

    $ep = $Context.Endpoints
    if ($Advertised.Kind -eq 'None') { return $null }
    switch ([string]$Context.Config.Authority) {
        'ADFS' {
            if ($Advertised.Kind -eq 'EntraID') {
                return "Exchange sends clients to Entra ID ($($Advertised.Uri)): hybrid modern authentication is enabled (Get-AuthServer: EvoSts is the default authorization endpoint), not AD FS $($ep.AdfsHost). Test it with -Authority EntraID."
            }
            if ($ep.AdfsHost -and $Advertised.Host -ine $ep.AdfsHost) {
                return "Exchange names the authorization server $($Advertised.Host), not $($ep.AdfsHost): clients will sign in there. Check Get-AuthServer (IsDefaultAuthorizationEndpoint)."
            }
        }
        'EntraID' {
            if ($Advertised.Kind -ne 'EntraID') {
                $evo = 'Set-AuthServer ''EvoSts - <ID>'' -IsDefaultAuthorizationEndpoint $true and Set-OrganizationConfig -OAuth2ClientProfileEnabled $true'
                return "Exchange sends clients to $($Advertised.Name), not to Entra ID: hybrid modern authentication is not enabled. Run the Hybrid Configuration Wizard, then $evo."
            }
            if ($Context.TenantId -and $Advertised.Tenant -match '^[0-9a-fA-F-]{36}$' -and $Advertised.Tenant -ine $Context.TenantId) {
                return "Exchange sends clients to the Entra ID tenant $($Advertised.Tenant), not to $($Context.TenantId): check Get-AuthServer (EvoSts, IsDefaultAuthorizationEndpoint) and Target.TenantId."
            }
        }
    }
    return $null
}

function Get-WscOverallStatus {
    <# Failed > Blocked > Warning > Passed. Skipped steps never decide the status alone. #>
    param([AllowEmptyCollection()][object[]]$Steps)

    $statuses = @($Steps | ForEach-Object { $_.Status })
    if ($statuses -contains 'Failed') { return 'Failed' }
    if ($statuses -contains 'Blocked') { return 'Blocked' }
    if ($statuses -contains 'Warning') { return 'Warning' }
    if (@($statuses | Where-Object { $_ -eq 'Passed' }).Count) { return 'Passed' }
    return 'Failed'
}

function Get-WscEntraErrorHint {
    <# What to check for the Entra ID errors (AADSTS codes) a sign-in meets most often. #>
    param([AllowEmptyString()][string]$Description, [string]$Resource, [switch]$ExchangeOnline)

    $code = [regex]::Match($Description, 'AADSTS(\d+)').Groups[1].Value
    switch ($code) {
        '500011' {
            if ($ExchangeOnline) { return "Exchange Online ($Resource) is not found in this tenant: check the tenant (Target.TenantId, the domain of the mailbox) and that it has Exchange Online." }
            return "the URL $Resource is not a service principal name of Office 365 Exchange Online ($($script:Entra.ExchangeApp)) in this tenant: run the Hybrid Configuration Wizard, or add the external and internal EWS URLs to its servicePrincipalNames (Microsoft Graph, Update-MgServicePrincipal)."
        }
        '65001' { return 'the user or an administrator has not consented to this client for Exchange: grant the consent in Entra ID (Enterprise applications).' }
        '90094' { return 'an administrator must consent to this client (users cannot consent in this tenant): grant admin consent to the application in Entra ID (Enterprise applications).' }
        '53003' { return 'a Conditional Access policy blocked the sign-in: read the sign-in log of the user in Entra ID (Conditional Access tab).' }
        '50105' { return 'the user is not assigned to the application (assignment required).' }
        '700016' { return 'this client ID does not exist in the tenant (Identity.ClientId or Identity.AppClientId).' }
        '7000218' { return 'the client is not a public client: the device-code flow needs a public client (Allow public client flows).' }
        { $_ -in '50020', '50034', '90072' } { return 'the account used in the browser is not a user of this tenant.' }
        '50076' { return 'multi-factor authentication is required: complete it in the browser.' }
        '7000215' { return 'the client secret is wrong or expired (Certificates & secrets of the application).' }
        '700027' { return 'the certificate is not one of the application (Certificates & secrets), or the assertion is invalid.' }
        '1002012' { return 'the scope of an application must end with /.default.' }
        '90002' { return 'the tenant does not exist (Target.TenantId or the domain of the mailbox).' }
        default { return $null }
    }
}

function Invoke-WscDeviceCodeAuthentication {
    <#
        Device-code flow (RFC 8628) with AD FS or Entra ID, every request traced. The verification
        page is opened; the code is shown in the console and the GUI. UserAgent: the one of the
        client played.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Configuration,
        [Parameter(Mandatory = $true)][pscustomobject]$Endpoints,
        [Parameter(Mandatory = $true)][Net.Http.HttpClient]$HttpClient,
        [string]$UserAgent
    )

    $server = if ($Endpoints.Authority -eq 'EntraID') { 'Entra ID' } else { 'AD FS' }
    $explain = {
        param([string]$Description)
        $first = ([string]$Description -split "`r?`n")[0]
        $hint = if ($Endpoints.Authority -eq 'EntraID') { Get-WscEntraErrorHint -Description $first -Resource $Endpoints.Resource -ExchangeOnline:$Endpoints.ExchangeOnline } else { $null }
        if ($hint) { "$first Cause: $hint" } else { $first }
    }
    $requested = Invoke-WscFormPost -HttpClient $HttpClient -Uri $Endpoints.DeviceCodeEndpoint -UserAgent $UserAgent `
        -Fields ([ordered]@{ client_id = [string]$Configuration.ClientId; scope = $Endpoints.Scope })
    $deviceCode = $requested.Json
    if ($requested.StatusCode -ne 200) {
        throw ('{0} refused the device-code request (HTTP {1}): {2} - {3}' -f $server, $requested.StatusCode, [string](Get-WscField $deviceCode 'error'), (& $explain ([string](Get-WscField $deviceCode 'error_description'))))
    }
    $deviceCodeValue = [string](Get-WscField $deviceCode 'device_code')
    if ([string]::IsNullOrWhiteSpace($deviceCodeValue)) { throw "$server did not return a device code." }
    $expiresIn = 0
    if (-not [int]::TryParse([string](Get-WscField $deviceCode 'expires_in'), [ref]$expiresIn) -or $expiresIn -le 0) {
        throw "$server did not return a valid device-code expiration."
    }
    $interval = 5
    $serverInterval = 0
    if ([int]::TryParse([string](Get-WscField $deviceCode 'interval'), [ref]$serverInterval) -and $serverInterval -gt 0) {
        $interval = [Math]::Max($serverInterval, 5)
    }
    $verificationUrl = [string](Get-WscField $deviceCode 'verification_uri_complete')
    if ([string]::IsNullOrWhiteSpace($verificationUrl)) { $verificationUrl = [string](Get-WscField $deviceCode 'verification_uri') }
    if ([string]::IsNullOrWhiteSpace($verificationUrl)) { throw "$server did not return a verification URL." }

    $message = [string](Get-WscField $deviceCode 'message')
    $userCode = [string](Get-WscField $deviceCode 'user_code')
    if ($message) { Write-WscItem Info $message -Icon Key }
    if ($userCode) { Write-WscItem Info ("Code: {0}  -  page: {1}" -f $userCode, [string](Get-WscField $deviceCode 'verification_uri')) -Icon Key }
    Write-WscItem Info 'Sign in with the account of the test: if the browser is already signed in with another account (work profile), open the page in a private window.' -Icon Key
    Write-WscItem Info ('Waiting for the sign-in in the browser (up to {0}).' -f (Format-WscDuration ([Math]::Min($expiresIn, [int]$Configuration.OAuthPollTimeoutSeconds)))) -Icon Clock
    # No browser in a service session, Server Core or SSH: the code is signed in from any other device.
    try { Start-Process -FilePath $verificationUrl -ErrorAction Stop | Out-Null }
    catch { Write-WscItem Info "No browser could be opened here ($($_.Exception.Message)): open $verificationUrl on any device and enter the code." -Icon Key }

    $deadline = [DateTimeOffset]::UtcNow.AddSeconds([Math]::Min($expiresIn, [int]$Configuration.OAuthPollTimeoutSeconds))
    while ([DateTimeOffset]::UtcNow -lt $deadline) {
        Wait-WscSeconds $interval
        $answer = Invoke-WscFormPost -HttpClient $HttpClient -Uri $Endpoints.TokenEndpoint -UserAgent $UserAgent -Collapse 'TokenPoll' -Fields ([ordered]@{
                grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id   = [string]$Configuration.ClientId
                device_code = $deviceCodeValue
            })
        if ($answer.StatusCode -eq 200) {
            $accessToken = [string](Get-WscField $answer.Json 'access_token')
            if ([string]::IsNullOrWhiteSpace($accessToken)) { throw "$server returned a token response without access_token." }
            return $accessToken
        }
        $code = [string](Get-WscField $answer.Json 'error')
        if (-not $code) { throw "$server token request failed: HTTP $($answer.StatusCode) without an OAuth error." }
        if ($code -eq 'authorization_pending') { continue }
        if ($code -eq 'slow_down') { $interval += 5; continue }
        if ($code -in 'access_denied', 'authorization_declined') { throw "$server sign-in was denied." }
        if ($code -eq 'expired_token') { throw "The $server device code expired before the sign-in was completed." }
        throw ('{0} token request failed: {1} - {2}' -f $server, $code, (& $explain ([string](Get-WscField $answer.Json 'error_description'))))
    }
    throw "$server did not return an access token before the configured timeout (Test.OAuthPollTimeoutSeconds)."
}

function Get-WscUrlFields {
    <# Parameters of the query (and fragment) of a URL, decoded. #>
    param([Parameter(Mandatory = $true)][string]$Url)

    $fields = @{}
    $start = $Url.IndexOfAny([char[]]'?#')
    if ($start -lt 0) { return $fields }
    foreach ($pair in $Url.Substring($start + 1).Split([char[]]'&#', [StringSplitOptions]::RemoveEmptyEntries)) {
        $parts = $pair.Split('=', 2)
        $fields[[Uri]::UnescapeDataString($parts[0])] = if ($parts.Count -gt 1) { [Uri]::UnescapeDataString($parts[1].Replace('+', ' ')) } else { '' }
    }
    return $fields
}

function Get-WscTokenClaims {
    <# Payload of a JWT access token as a hashtable, or $null if the token is not a readable JWT. The signature is not checked. #>
    param([Parameter(Mandatory = $true)][string]$AccessToken)

    $parts = $AccessToken.Split('.')
    if ($parts.Count -lt 2) { return $null }
    try {
        $json = [Text.Encoding]::UTF8.GetString((ConvertFrom-WscBase64Url $parts[1]))
        $claims = $json | ConvertFrom-Json -AsHashtable
        if ($claims -isnot [hashtable]) { return $null }
        return $claims
    }
    catch {
        return $null
    }
}

function Get-WscAuthorizeOutcome {
    <# What the AD FS authorization page answers to one request: sign-in page, unknown client, refused redirect URI... #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Response)

    $content = [string]$Response.Content
    $msis = [regex]::Match($content, 'MSIS\d{4}[^<]*')
    $text = if ($msis.Success) { [Net.WebUtility]::HtmlDecode($msis.Value).Trim() } else { $null }
    $outcome = { param([string]$Code, [string]$Text) [pscustomobject]@{ Code = $Code; Text = $Text } }
    if ($content -match 'MSIS9223') { return & $outcome 'UnknownClient' $text }
    if ($content -match 'MSIS9224') { return & $outcome 'RedirectRefused' $text }
    if ($Response.StatusCode -in 301, 302, 303 -and $Response.Location) {
        if ($Response.Location -match '[?&]code=') { return & $outcome 'SignedIn' 'AD FS signed in without a prompt (Windows integrated authentication) and returned an authorization code.' }
        if ($Response.Location -match '[?&]error=([^&]+)') {
            $description = if ($Response.Location -match '[?&]error_description=([^&]+)') { ': ' + [Uri]::UnescapeDataString($Matches[1].Replace('+', ' ')) } else { '' }
            return & $outcome 'Error' "AD FS returned the error $([Uri]::UnescapeDataString($Matches[1]))$description"
        }
        return & $outcome 'SignInPage' "AD FS continues the sign-in at $($Response.Location)."
    }
    if ($text) { return & $outcome 'Error' $text }
    if ($Response.StatusCode -eq 200 -and $content -match 'userNameInput|passwordInput|loginForm|idp_') { return & $outcome 'SignInPage' 'AD FS shows its sign-in page.' }
    return & $outcome 'Error' "HTTP $($Response.StatusCode) without a sign-in page."
}

function Resolve-WscEntraTenant {
    <#
        Tenant ID from the OpenID configuration Entra ID publishes for a tenant ID or a domain (traced,
        no sign-in). Returns TenantId (empty when the tenant is not found), Issuer, Url and Error.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Name)

    $url = "https://$($script:Entra.LoginHost)/$([Uri]::EscapeDataString($Name))/v2.0/.well-known/openid-configuration"
    $response = Invoke-WscWebRequest -HttpClient $Context.HttpClient -Uri $url -UserAgent $script:WscUserAgent -Headers @{ Accept = 'application/json' }
    $json = $null
    try { $json = $response.Content | ConvertFrom-Json -ErrorAction Stop } catch { $json = $null }
    $issuer = [string](Get-WscField $json 'issuer')
    $tenantId = [regex]::Match($issuer, '[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}').Value
    $problem = if ($response.StatusCode -eq 200 -and $tenantId) { $null }
    elseif ($json -and (Get-WscField $json 'error_description')) { ([string](Get-WscField $json 'error_description') -split "`r?`n")[0] }
    else { "HTTP $($response.StatusCode)" }
    [pscustomobject]@{ TenantId = if ($problem) { $null } else { $tenantId }; Issuer = $issuer; Url = $url; Error = $problem }
}

function Get-WscUserRealm {
    <#
        How Entra ID signs in a user, before any sign-in (traced): Managed (password hash
        synchronization, pass-through authentication), Federated (AD FS or another identity
        provider: AuthUrl) or Unknown (the domain is not a domain of a tenant).
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$User)

    $url = "https://$($script:Entra.LoginHost)/common/userrealm/$([Uri]::EscapeDataString($User))?api-version=2.0"
    $response = Invoke-WscWebRequest -HttpClient $Context.HttpClient -Uri $url -UserAgent $script:WscUserAgent -Headers @{ Accept = 'application/json' }
    $json = $null
    try { $json = $response.Content | ConvertFrom-Json -ErrorAction Stop } catch { $json = $null }
    [pscustomobject]@{
        HttpStatus          = $response.StatusCode
        NameSpaceType       = [string](Get-WscField $json 'NameSpaceType')
        DomainName          = [string](Get-WscField $json 'DomainName')
        FederationBrandName = [string](Get-WscField $json 'FederationBrandName')
        AuthUrl             = [string](Get-WscField $json 'AuthURL')
    }
}

function Get-WscEntraSignInOutcome {
    <#
        What the Entra ID authorization page shows: before the password Entra ID checks neither the
        client nor the redirect URI nor the resource (AADSTS50058: sign-in page), only the tenant.
    #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Response)

    $content = [string]$Response.Content
    $config = [regex]::Match($content, '\$Config=(\{.*?\});', 'Singleline')
    $json = $null
    if ($config.Success) { try { $json = $config.Groups[1].Value | ConvertFrom-Json -AsHashtable -ErrorAction Stop } catch { $json = $null } }
    $message = if ($json -and $json['strServiceExceptionMessage']) { [string]$json['strServiceExceptionMessage'] } else { ([regex]::Match($content, 'AADSTS\d+[^"\\<]{0,200}')).Value }
    $outcome = { param([string]$Code, [string]$Text) [pscustomobject]@{ Code = $Code; Text = $Text } }
    if ($Response.StatusCode -in 301, 302, 303 -and $Response.Location) { return & $outcome 'SignInPage' "Entra ID continues the sign-in at $($Response.Location)." }
    if ($message) { return & $outcome 'Error' $message }
    if ($Response.StatusCode -eq 200 -and ($content -match '"urlPost"' -or ($json -and $json['pgid'] -match 'SignIn'))) { return & $outcome 'SignInPage' 'Entra ID shows its sign-in page.' }
    return & $outcome 'Error' "HTTP $($Response.StatusCode) without a sign-in page."
}
