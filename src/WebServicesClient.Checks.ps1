<#
.SYNOPSIS
    Web Services Client for Exchange - Autodiscover, prerequisites, sign-in and orchestration (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    Invoke-WscMailboxTest runs the stages of a scenario in order with one shared context (token or
    credentials, HTTP client with its cookie container, EWS URL, access mode, folders, messages...).
    Every check adds one step with a status: Passed, Warning, Blocked, Failed or Skipped. A stage
    that fails or is blocked stops the scenario: the following stages are reported as Skipped.

    Sign-in, by authentication and context:
      OAuth  User         the user signs in with a Microsoft public client (Identity.ClientId)
             Delegated    the user signs in through your application (Identity.AppClientId)
             Application  your application signs in alone: client credentials, certificate or secret
      Basic               user name and password with every request
      Windows             Negotiate, NTLM or Kerberos handshake done and traced by the tool

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

#region Steps --------------------------------------------------------------------------------------

function Add-WscStep {
    <# Records one check: status, message, details and the time spent since the previous step. #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Context,
        [Parameter(Mandatory = $true)][string]$Stage,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][ValidateSet('Passed', 'Warning', 'Blocked', 'Failed', 'Skipped')][string]$Status,
        [Parameter(Mandatory = $true)][string]$Message,
        [System.Collections.IDictionary]$Details
    )

    $elapsed = $Context.Clock.Elapsed.TotalMilliseconds
    $Context.Clock.Restart()
    $copy = [ordered]@{}
    if ($Details) { foreach ($key in $Details.Keys) { if ($null -ne $Details[$key] -and [string]$Details[$key] -ne '') { $copy[$key] = $Details[$key] } } }
    $trace = Complete-WscTraceStep -Step ($Context.Steps.Count + 1) -Name $Name
    $Context.Steps.Add([pscustomobject]@{
            Stage        = $Stage
            Name         = $Name
            Status       = $Status
            Message      = $Message
            Details      = $copy
            Trace        = @($trace)
            DurationMs   = [int][Math]::Round($elapsed)
            TimestampUtc = [DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        })
    $item = @{ Passed = 'Ok'; Warning = 'Warn'; Blocked = 'Block'; Failed = 'Fail'; Skipped = 'Skip' }[$Status]
    Write-WscItem $item ('{0}: {1}' -f $Name, $Message)
}

#endregion

#region OAuth sign-in ------------------------------------------------------------------------------

function Get-WscRedirectUri {
    <# Redirect URI of the sign-in window: Identity.RedirectUri (Delegated), else the default of the server. #>
    param([Parameter(Mandatory = $true)][hashtable]$Configuration, [Parameter(Mandatory = $true)][pscustomobject]$Endpoints)
    if ([string]$Configuration.Context -eq 'Delegated' -and [string]$Configuration.RedirectUri) { return [string]$Configuration.RedirectUri }
    if ($Endpoints.Authority -eq 'EntraID') { return $script:SignInRedirect.EntraID }
    return $script:SignInRedirect.ADFS
}

function Get-WscSignInRequest {
    <# Authorization request of the sign-in window: authorization code with PKCE (S256), random state, login_hint, prompt=login. #>
    param([Parameter(Mandatory = $true)][hashtable]$Configuration, [Parameter(Mandatory = $true)][pscustomobject]$Endpoints)

    $verifier = ConvertTo-WscBase64Url ([Security.Cryptography.RandomNumberGenerator]::GetBytes(32))
    $challenge = ConvertTo-WscBase64Url ([Security.Cryptography.SHA256]::HashData([Text.Encoding]::ASCII.GetBytes($verifier)))
    $redirect = Get-WscRedirectUri -Configuration $Configuration -Endpoints $Endpoints
    $state = [guid]::NewGuid().ToString('N')
    $hint = if ([string]$Configuration.SignInUser -match '@') { [string]$Configuration.SignInUser } elseif (-not $Configuration.SignInUser) { [string]$Configuration.Mailbox } else { $null }
    $query = [ordered]@{
        response_type = 'code'; client_id = [string]$Configuration.ClientId; redirect_uri = $redirect; scope = $Endpoints.Scope; state = $state
        code_challenge = $challenge; code_challenge_method = 'S256'; prompt = 'login'
    }
    if ($hint) { $query.login_hint = $hint }
    [pscustomobject]@{
        Url         = $Endpoints.AuthorizeEndpoint + '?' + (@($query.Keys | ForEach-Object { '{0}={1}' -f $_, [Uri]::EscapeDataString([string]$query[$_]) }) -join '&')
        RedirectUri = $redirect
        State       = $state
        Verifier    = $verifier
        Hint        = $hint
    }
}

function Invoke-WscWindowAuthentication {
    <#
        Sign-in in the window (authorization code with PKCE), then the code exchanged for the token, traced.
        A confidential delegated application also sends its client secret with the code.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Context,
        [Parameter(Mandatory = $true)][pscustomobject]$Browser
    )

    $cfg = $Context.Config
    $ep = $Context.Endpoints
    $entra = $ep.Authority -eq 'EntraID'
    $server = if ($entra) { 'Entra ID' } else { 'AD FS' }
    $explain = {
        param([string]$Description)
        $first = ([string]$Description -split "`r?`n")[0]
        $hint = if ($entra) { Get-WscEntraErrorHint -Description $first -Resource $ep.Resource -ExchangeOnline:$ep.ExchangeOnline } else { $null }
        if ($hint) { "$first Cause: $hint" } else { $first }
    }
    $request = Get-WscSignInRequest -Configuration $cfg -Endpoints $ep
    $timeout = [int]$cfg.OAuthPollTimeoutSeconds
    $who = if ($request.Hint) { $request.Hint } else { 'the account of the test' }
    Write-WscItem Info ('Sign-in window ({0}, temporary profile): sign in as {1} on the {2} page - password, then MFA if asked. The window closes by itself.' -f $Browser.Name, $who, $server) -Icon Key
    Write-WscItem Info ('Waiting for the sign-in in the window (up to {0}).' -f (Format-WscDuration $timeout)) -Icon Clock
    $opened = "GET $($request.Url)`n`nOpened in the sign-in window: $($Browser.Name) in app mode, temporary profile deleted afterwards.`nThe sign-in pages (password, MFA) are exchanged by the browser and are not recorded."
    $clock = [Diagnostics.Stopwatch]::StartNew()
    try {
        $redirect = Invoke-WscBrowserAuthorization -Browser $Browser -Url $request.Url -RedirectUri $request.RedirectUri -TimeoutSeconds $timeout
    }
    catch {
        Add-WscTraceEntry -Method 'GET' -Url $request.Url -Request $opened -Response "No authorization code: $($_.Exception.Message)" -StatusCode $null -DurationMs $clock.ElapsedMilliseconds -Label 'sign-in window'
        if ($_.Exception -is [NotSupportedException]) { throw }
        throw "The $server sign-in in the window did not complete: $($_.Exception.Message)"
    }
    $masked = [regex]::Replace($redirect, '([?&#]code=)([^&#]+)', { param($m) $m.Groups[1].Value + "<$($m.Groups[2].Value.Length) characters, never written>" })
    Add-WscTraceEntry -Method 'GET' -Url $request.Url -Request $opened -Response "Redirect caught by the tool (not followed by the browser):`n$masked" `
        -StatusCode 302 -Reason 'Redirect caught' -DurationMs $clock.ElapsedMilliseconds -Label 'sign-in window'
    $fields = Get-WscUrlFields -Url $redirect
    if ($fields['error']) { throw ('{0} answered the sign-in with the error {1}: {2}' -f $server, $fields['error'], (& $explain ([string]$fields['error_description']))) }
    if ([string]$fields['state'] -ne $request.State) { throw "$server returned an authorization code with another state: it does not answer this sign-in." }
    $code = [string]$fields['code']
    if (-not $code) { throw "$server redirected to $($request.RedirectUri) without an authorization code." }

    $token = [ordered]@{ grant_type = 'authorization_code'; client_id = [string]$cfg.ClientId; code = $code; redirect_uri = $request.RedirectUri; code_verifier = $request.Verifier }
    if ($entra) { $token.scope = $ep.Scope }
    if ($Context.ClientSecret) { $token.client_secret = [Net.NetworkCredential]::new('', $Context.ClientSecret).Password }
    $answer = Invoke-WscFormPost -HttpClient $Context.HttpClient -Uri $ep.TokenEndpoint -Fields $token
    if ($answer.StatusCode -ne 200) {
        throw ('{0} refused the authorization code (HTTP {1}): {2} - {3}' -f $server, $answer.StatusCode, [string](Get-WscField $answer.Json 'error'), (& $explain ([string](Get-WscField $answer.Json 'error_description'))))
    }
    $accessToken = [string](Get-WscField $answer.Json 'access_token')
    if ([string]::IsNullOrWhiteSpace($accessToken)) { throw "$server returned a token response without access_token." }
    [pscustomobject]@{ AccessToken = $accessToken; RedirectUri = $request.RedirectUri }
}

function Test-WscAdfsWindowRedirect {
    <# AD FS: the authorization page must accept the redirect URI of the window for the client before the window opens. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $cfg = $Context.Config
    $redirect = Get-WscRedirectUri -Configuration $cfg -Endpoints $Context.Endpoints
    $query = [ordered]@{ response_type = 'code'; client_id = [string]$cfg.ClientId; redirect_uri = $redirect; scope = $Context.Endpoints.Scope; state = 'wsc-check' }
    $url = $Context.Endpoints.AuthorizeEndpoint + '?' + (@($query.Keys | ForEach-Object { '{0}={1}' -f $_, [Uri]::EscapeDataString([string]$query[$_]) }) -join '&')
    try { Get-WscAuthorizeOutcome -Response (Invoke-WscWebRequest -HttpClient $Context.HttpClient -Uri $url -UserAgent ([string]$cfg.UserAgent)) }
    catch { [pscustomobject]@{ Code = 'Error'; Text = "AD FS not reachable: $($_.Exception.Message)" } }
}

function Get-WscApplicationCertificate {
    <# Certificate of the application with its private key: Cert:\CurrentUser\My, then Cert:\LocalMachine\My. #>
    param([Parameter(Mandatory = $true)][string]$Thumbprint)

    foreach ($store in 'Cert:\CurrentUser\My', 'Cert:\LocalMachine\My') {
        $cert = Get-Item -LiteralPath (Join-Path $store $Thumbprint) -ErrorAction SilentlyContinue
        if ($cert) {
            if (-not $cert.HasPrivateKey) { throw "The certificate $Thumbprint ($store) has no private key: the application cannot sign its assertion." }
            return [pscustomobject]@{ Certificate = $cert; Store = $store }
        }
    }
    throw "Certificate $Thumbprint not found in Cert:\CurrentUser\My nor Cert:\LocalMachine\My (Identity.CertificateThumbprint)."
}

function New-WscClientAssertion {
    <# JWT signed with the certificate of the application (RS256, x5t), audience the token endpoint, valid 10 minutes. #>
    param([Parameter(Mandatory = $true)][Security.Cryptography.X509Certificates.X509Certificate2]$Certificate, [Parameter(Mandatory = $true)][string]$ClientId, [Parameter(Mandatory = $true)][string]$Audience)

    $now = [DateTimeOffset]::UtcNow
    $header = [ordered]@{ alg = 'RS256'; typ = 'JWT'; x5t = ConvertTo-WscBase64Url $Certificate.GetCertHash() }
    $payload = [ordered]@{ aud = $Audience; iss = $ClientId; sub = $ClientId; jti = [guid]::NewGuid().ToString(); nbf = $now.AddSeconds(-30).ToUnixTimeSeconds(); exp = $now.AddMinutes(10).ToUnixTimeSeconds() }
    $data = '{0}.{1}' -f (ConvertTo-WscBase64Url ([Text.Encoding]::UTF8.GetBytes(($header | ConvertTo-Json -Compress)))), (ConvertTo-WscBase64Url ([Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json -Compress))))
    $rsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if (-not $rsa) { throw "The certificate $($Certificate.Thumbprint) has no RSA private key usable here." }
    try {
        $signature = $rsa.SignData([Text.Encoding]::ASCII.GetBytes($data), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
    }
    finally { $rsa.Dispose() }
    return '{0}.{1}' -f $data, (ConvertTo-WscBase64Url $signature)
}

function Invoke-WscClientCredentials {
    <#
        Application sign-in (client credentials). Entra ID: scope https://<exchange>/.default with a
        certificate (client_assertion) or the client secret. AD FS: the resource of Exchange with the
        client secret (server application of an AD FS application group). Returns the token and how.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $cfg = $Context.Config
    $ep = $Context.Endpoints
    $entra = $ep.Authority -eq 'EntraID'
    $server = if ($entra) { 'Entra ID' } else { 'AD FS' }
    $fields = [ordered]@{ grant_type = 'client_credentials'; client_id = [string]$cfg.ClientId }
    $how = 'client secret'
    if ($entra) { $fields.scope = $ep.Scope } else { $fields.resource = $ep.Resource; $fields.scope = 'openid' }
    if ($entra -and [string]$cfg.CertificateThumbprint) {
        $found = Get-WscApplicationCertificate -Thumbprint ([string]$cfg.CertificateThumbprint)
        $fields.client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        $fields.client_assertion = New-WscClientAssertion -Certificate $found.Certificate -ClientId ([string]$cfg.ClientId) -Audience $ep.TokenEndpoint
        $how = "certificate $($found.Certificate.Thumbprint) ($($found.Store), $($found.Certificate.Subject), valid until $($found.Certificate.NotAfter.ToString('yyyy-MM-dd')))"
    }
    else {
        if (-not $Context.ClientSecret) { throw "The Application context needs the client secret of $($cfg.ClientId) (-ClientSecret, the prompt or the window)$(if ($entra) { ', or a certificate (Identity.CertificateThumbprint)' })." }
        $fields.client_secret = [Net.NetworkCredential]::new('', $Context.ClientSecret).Password
    }
    $answer = Invoke-WscFormPost -HttpClient $Context.HttpClient -Uri $ep.TokenEndpoint -Fields $fields
    if ($answer.StatusCode -ne 200) {
        $description = ([string](Get-WscField $answer.Json 'error_description') -split "`r?`n")[0]
        $hint = if ($entra) { Get-WscEntraErrorHint -Description $description -Resource $ep.Resource -ExchangeOnline:$ep.ExchangeOnline } else { $null }
        if (-not $entra -and $description -match 'MSIS9611|MSIS9612|MSIS9622') { $hint = 'the client is not a server application of AD FS with this secret, or it has no permission on the Exchange relying party (Grant-AdfsApplicationPermission).' }
        throw ('{0} refused the application sign-in (HTTP {1}): {2} - {3}{4}' -f $server, $answer.StatusCode, [string](Get-WscField $answer.Json 'error'), $description, $(if ($hint) { " Cause: $hint" }))
    }
    $token = [string](Get-WscField $answer.Json 'access_token')
    if (-not $token) { throw "$server returned a token response without access_token." }
    [pscustomobject]@{ AccessToken = $token; How = $how }
}

function Test-WscTokenClaims {
    <#
        Audience, permission and expiry of the token compared with the EWS resource. User and delegated
        tokens: scope EWS.AccessAsUser.All; application tokens: role full_access_as_app. Entra ID: tenant (tid).
    #>
    param(
        [AllowNull()][hashtable]$Claims,
        [Parameter(Mandatory = $true)][pscustomobject]$Endpoints,
        [Parameter(Mandatory = $true)][hashtable]$Configuration,
        [string]$TenantId,
        [DateTimeOffset]$Now = [DateTimeOffset]::UtcNow
    )

    if ($null -eq $Claims) {
        return [pscustomobject]@{ Status = 'Warning'; Message = 'The access token is not a readable JWT: its claims were not checked.'; Details = [ordered]@{} }
    }
    $app = [string]$Configuration.Context -eq 'Application'
    $first = { param([string[]]$Names) foreach ($n in $Names) { if ($Claims.ContainsKey($n) -and -not [string]::IsNullOrWhiteSpace([string]$Claims[$n])) { return [string]$Claims[$n] } }; return $null }
    $audiences = @($Claims['aud'] | Where-Object { $_ })
    $scope = & $first @('scp', 'scope')
    $roles = @($Claims['roles'] | Where-Object { $_ })
    $user = & $first @('upn', 'unique_name', 'email', 'preferred_username')
    $expires = $null
    $exp = 0L
    if ([long]::TryParse([string]$Claims['exp'], [ref]$exp)) { $expires = [DateTimeOffset]::FromUnixTimeSeconds($exp) }
    $details = [ordered]@{
        Issuer      = & $first @('iss')
        Audience    = $audiences -join ', '
        Scope       = $scope
        Roles       = $roles -join ', '
        User        = $user
        ClientId    = & $first @('appid', 'client_id', 'azp')
        AppName     = & $first @('app_displayname')
        TenantId    = & $first @('tid')
        ExpiresUtc  = if ($expires) { $expires.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ') } else { $null }
        MinutesLeft = if ($expires) { [int][Math]::Floor(($expires - $Now).TotalMinutes) } else { $null }
    }
    if ($expires -and $expires -le $Now) {
        return [pscustomobject]@{ Status = 'Failed'; Message = "The token expired at $($details.ExpiresUtc)."; Details = $details }
    }
    $issues = [Collections.Generic.List[string]]::new()
    if (-not $expires) { [void]$issues.Add('the token has no exp claim') }
    $expected = ([string]$Endpoints.Resource).TrimEnd('/')
    $cloud = @($audiences | Where-Object { $_ -ieq $script:Entra.ExchangeApp -or $_ -match 'outlook\.office(365)?\.com' }).Count
    $matched = @($audiences | Where-Object { ([string]$_).TrimEnd('/') -ieq $expected }).Count -or ($Endpoints.ExchangeOnline -and $cloud)
    if (-not $matched) {
        $why = if ($cloud) { ': the token is for Exchange Online, not for the on-premises URL' } else { '' }
        [void]$issues.Add("the audience '$($details.Audience)' is not the EWS resource '$($Endpoints.Resource)'$why. Exchange will reject the token (HTTP 401)")
    }
    if ($app) {
        if ($roles -notcontains 'full_access_as_app' -and $Endpoints.Authority -eq 'EntraID') {
            [void]$issues.Add("the token has no role full_access_as_app (roles: $(if ($roles) { $roles -join ', ' } else { 'none' })): grant the application permission full_access_as_app of Office 365 Exchange Online and the admin consent")
        }
        if ($scope) { [void]$issues.Add("the token has a scope ($scope): it is a delegated token, not an application token") }
    }
    elseif ($scope -notmatch '(^|\s)EWS\.AccessAsUser\.All(\s|$)') {
        [void]$issues.Add("the scope '$scope' does not contain EWS.AccessAsUser.All")
    }
    if ($details.ClientId -and $details.ClientId -ine [string]$Configuration.ClientId) {
        [void]$issues.Add("the token was issued to the client $($details.ClientId), not to $($Configuration.ClientId)")
    }
    if ($TenantId -and $details.TenantId -and $details.TenantId -ine $TenantId) {
        [void]$issues.Add("the token comes from the tenant $($details.TenantId), not from $TenantId")
    }
    $signIn = [string]$Configuration.SignInUser
    $expectedUser = if ($signIn) { $signIn } else { [string]$Configuration.Mailbox }
    $note = if (-not $app -and $user -and $expectedUser -match '@' -and $user -ine $expectedUser) { " Token user $user is not written like $expectedUser (UPN and SMTP address can differ)." } else { '' }
    if ($issues.Count) {
        return [pscustomobject]@{ Status = 'Warning'; Message = ('Token received, but ' + ($issues -join '; ') + '.' + $note); Details = $details }
    }
    $left = if ($null -ne $details.MinutesLeft) { " (valid for $($details.MinutesLeft) min)" } else { '' }
    $what = if ($app) { 'role full_access_as_app' } else { 'scope EWS.AccessAsUser.All' }
    return [pscustomobject]@{ Status = 'Passed'; Message = "Audience, $what$(if ($TenantId -and $details.TenantId) { ', tenant' }) and expiry match the EWS resource$left.$note"; Details = $details }
}

#endregion
