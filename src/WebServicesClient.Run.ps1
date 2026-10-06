<#
.SYNOPSIS
    Web Services Client for Exchange - the sign-in stage and the orchestration of a run (dot-sourced by WebServicesClient.psm1).

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

function Resolve-WscClientSettings {
    <#
        Settings of a run: defaults, then the configuration; the client ID used for OAuth (the Microsoft
        client for the User context, your application for Delegated and Application), and the access mode.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Configuration)

    $cfg = Get-WscDefaultConfiguration
    foreach ($key in $Configuration.Keys) { $cfg[$key] = $Configuration[$key] }
    $cfg.UserClientId = [string]$cfg.ClientId
    if ([string]$cfg.Context -in 'Delegated', 'Application') { $cfg.ClientId = [string]$cfg.AppClientId }
    $cfg.FreeBusyMailboxes = @($cfg.FreeBusyMailboxes | Where-Object { $_ })
    return $cfg
}

function Set-WscProtocol {
    <#
        The protocol of the run, once the EWS URL is known (Autodiscover or configuration). Auto: Microsoft Graph
        for a mailbox in Exchange Online (EWS is being retired there since October 2026), EWS for a mailbox
        on-premises (Graph does not reach it). Recomputes the endpoints: the token of Graph is for Graph.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage)

    $cfg = $Context.Config
    $wanted = [string]$cfg.Protocol
    $online = [bool]$Context.Endpoints.ExchangeOnline
    $use = if ($wanted -eq 'Graph' -or ($wanted -eq 'Auto' -and $online)) { 'Graph' } else { 'EWS' }
    $details = [ordered]@{ Protocol = $wanted; Used = $use; EwsUrl = $Context.Endpoints.EwsUrl; ExchangeOnline = $online }
    if ($use -eq 'Graph' -and [string]$cfg.Authentication -ne 'OAuth') {
        $use = 'EWS'
        $details.Used = 'EWS'
        $status = 'Warning'
        $message = "The mailbox is in Exchange Online, where Microsoft Graph replaces EWS, but Graph accepts only OAuth with Entra ID: the checks go on with EWS and $([string]$cfg.Authentication), which Exchange Online refuses. Run again with -Authentication OAuth -Authority EntraID."
    }
    elseif ($use -eq 'Graph') {
        $status = 'Passed'
        $message = if ($wanted -eq 'Graph') { 'Microsoft Graph, as asked (-Protocol Graph). It reaches only the mailboxes of Exchange Online.' }
        else { 'The mailbox is in Exchange Online: the operations go through Microsoft Graph. EWS is being retired in Exchange Online (disabled from October 2026, stopped in April 2027).' }
    }
    elseif ($online) {
        $status = 'Warning'
        $message = 'EWS forced (-Protocol EWS) for a mailbox in Exchange Online: EWS is being retired there; expect HTTP 403 (X-EWS-Policy-Reason) unless an administrator still allows it (EwsEnabled, EwsAllowedAppIDs).'
    }
    else {
        $status = 'Passed'
        $message = if ($wanted -eq 'EWS') { 'EWS, as asked (-Protocol EWS).' } else { 'The mailbox is on-premises: the operations go through EWS (Microsoft Graph does not reach on-premises mailboxes).' }
    }
    $Context.Protocol = $use
    $cfg.ProtocolUsed = $use
    if ($use -eq 'Graph') {
        if ([string]$cfg.Authority -eq 'Auto') { $cfg.Authority = 'EntraID' }
        $Context.Endpoints = Resolve-WscEndpoints -Configuration $cfg
        $details.GraphUrl = $Context.Endpoints.GraphUrl
        $details.Scope = $Context.Endpoints.Scope
    }
    Add-WscStep $Context $Stage 'Protocol' $status $message $details
}

function Invoke-WscStageSignIn {
    <# Sign-in of the run: OAuth (user, delegated application, application), Basic or Windows. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    switch ([string]$Context.Config.Authentication) {
        'Basic' { Invoke-WscSignInProbe -Context $Context -Name 'Basic sign-in' }
        'Windows' { Invoke-WscSignInProbe -Context $Context -Name 'Windows sign-in' }
        default { Invoke-WscSignInOAuth -Context $Context }
    }
}

function Invoke-WscSignInProbe {
    <#
        Basic or Windows: one GetServerTimeZones request (it needs no mailbox) with the credentials. The run stops at
        the first refusal, so that a wrong password is sent only once (account lockout).
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Name)

    $stage = 'SignIn'
    $cfg = $Context.Config
    $windows = [string]$cfg.Authentication -eq 'Windows'
    if ($Context.Endpoints.ExchangeOnline) {
        Add-WscStep $Context $stage $Name Failed "Exchange Online ($($Context.Endpoints.EwsHost)) accepts only OAuth for EWS: no password was sent. Test the mailbox with -Authentication OAuth -Authority EntraID." ([ordered]@{ EwsUrl = $Context.Endpoints.EwsUrl })
        $Context.Stop = $true
        return
    }
    if (-not $windows -and -not $Context.Credential) { throw 'Basic authentication needs a user name and a password (-Credential, the prompt or the window).' }
    $anchor = if ([string]$cfg.SignInUser -match '@') { [string]$cfg.SignInUser } elseif (-not $cfg.SignInUser) { [string]$cfg.Mailbox } else { $null }
    $answer = Invoke-WscEws -Context $Context -Operation 'GetServerTimeZones' -Body (New-WscGetServerTimeZonesBody) -Anchor $anchor -NoImpersonation
    $user = if ($Context.Credential) { $Context.Credential.UserName } else { "$([Environment]::UserDomainName)\$([Environment]::UserName) (current Windows account)" }
    $details = Get-WscEwsDetails -Answer $answer -More ([ordered]@{ User = $user })
    if ($windows -and $answer.Windows) {
        $w = $answer.Windows
        $details.Package = $w.Package
        $details.Legs = $w.Legs
        $details.Spn = $w.Spn
        $details.ChannelBinding = $w.ChannelBinding
        if ($w.ServerInfo) {
            $s = $w.ServerInfo
            $details.NtlmServer = $s.PSObject.Properties['DnsComputer'].Value
            $details.NtlmDomain = '{0} ({1})' -f $s.PSObject.Properties['NetBiosDomain'].Value, $s.PSObject.Properties['DnsDomain'].Value
        }
        $Context.Windows = [pscustomobject]@{ Package = $w.Package; Legs = $w.Legs; Spn = $w.Spn; Server = $details.NtlmServer; Domain = $details.NtlmDomain }
    }
    if ($answer.HttpStatus -eq 200 -and $answer.ResponseClass -eq 'Success') {
        if ($windows) {
            $pkg = [string]$details.Package
            $fallback = if ([string]$cfg.WindowsPackage -eq 'Negotiate' -and $pkg -eq 'NTLM') { ' Negotiate chose NTLM: no Kerberos ticket could be obtained for the SPN from here (KDC out of reach, SPN not registered, or a computer out of the domain).' } else { '' }
            $legs = if ($details.Legs -eq 1) { 'one request' } else { "$($details.Legs) requests" }
            Add-WscStep $Context $stage $Name Passed "Exchange accepted $user with $pkg ($legs, every leg in the trace).$fallback" $details
        }
        else { Add-WscStep $Context $stage $Name Passed "Exchange accepted the user name and password of $user (HTTP 200). They are sent with every request, protected only by TLS." $details }
        return
    }
    $Context.Stop = $true
    if ($answer.HttpStatus -eq 200) {
        Add-WscStep $Context $stage $Name Warning "Signed in, but GetServerTimeZones answered $($answer.ResponseCode): $(Get-WscEwsFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $false
        return
    }
    $text = Get-WscEwsFailureText -Context $Context -Answer $answer
    if ($answer.HttpStatus -eq 401 -and -not $windows) { $text += ' The run stops here: a wrong password is sent only once.' }
    if ($answer.HttpStatus -eq 401 -and $windows -and -not $Context.Credential) { $text += " The current Windows account ($user) is used: on a computer out of the domain give the account with -Credential." }
    Add-WscStep $Context $stage $Name Failed $text $details
}

function Invoke-WscSignInOAuth {
    <#
        OAuth sign-in. The authorization server is known before: AD FS (Target.AdfsUrl), Entra ID (tenant
        found by Discovery or here), or with Target.Authority Auto the one Exchange names for the mailbox.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'SignIn'
    $cfg = $Context.Config
    $app = [string]$cfg.Context -eq 'Application'
    if ([string]$cfg.Authority -eq 'Auto') {
        $challenge = Test-WscMailboxChallenge -Context $Context
        $server = Get-WscAuthorityInfo -Uri ([string]$challenge.AuthorizationUri)
        if ($server.Kind -eq 'ADFS') {
            Set-WscAuthority -Context $Context -Kind ADFS -AdfsRoot $server.AdfsRoot -Source 'Exchange challenge (authorization_uri)'
            Add-WscStep $Context $stage 'Authorization server' Passed "Exchange sends the clients of $($cfg.Mailbox) to AD FS $($server.AdfsRoot): the sign-in uses it." $challenge.Details
        }
        elseif ($server.Kind -eq 'EntraID') {
            Add-WscStep $Context $stage 'Authorization server' Passed "Exchange sends the clients of $($cfg.Mailbox) to Entra ID ($($challenge.AuthorizationUri)): the sign-in uses Entra ID." $challenge.Details
            if (-not (Add-WscEntraChecks -Context $Context -Stage $stage -Hint $server.Tenant -Source 'Exchange challenge (authorization_uri)')) { $Context.Stop = $true; return }
        }
        else {
            $why = if ($challenge.Status -eq 'Failed') { $challenge.Message } else { "Exchange names neither AD FS nor Entra ID ($($server.Name))." }
            Add-WscStep $Context $stage 'Authorization server' Failed "No authorization server to sign in with: $why" $challenge.Details
            $Context.Stop = $true
            return
        }
    }
    elseif ([string]$cfg.Authority -eq 'EntraID' -and -not $Context.TenantId) {
        if (-not (Add-WscEntraChecks -Context $Context -Stage $stage -Source 'Configuration')) { $Context.Stop = $true; return }
    }
    $ep = $Context.Endpoints
    $entra = $ep.Authority -eq 'EntraID'
    $server = if ($entra) { 'Entra ID' } else { 'AD FS' }
    if ($Context.AccessToken) {
        $Context.SignIn = 'Supplied'
        Add-WscStep $Context $stage 'Access token' Passed 'Access token supplied by the caller: sign-in skipped.' ([ordered]@{ Source = 'Caller' })
    }
    elseif ($app) {
        $result = Invoke-WscClientCredentials -Context $Context
        $Context.AccessToken = $result.AccessToken
        $Context.SignIn = 'ClientCredentials'
        Add-WscStep $Context $stage 'Application sign-in' Passed "Access token received from $server for the application $($cfg.ClientId) (client credentials, $($result.How)): no user, the application acts alone and impersonates $($cfg.Mailbox)." ([ordered]@{
                Source = "$server client credentials"; ClientId = $cfg.ClientId; Credential = $result.How; Scope = $(if ($entra) { $ep.Scope } else { "resource $($ep.Resource)" }); TokenEndpoint = $ep.TokenEndpoint })
    }
    else {
        $mode = Get-WscSignInMode -Configuration $cfg
        if ($mode.Mode -eq 'Window' -and -not $entra) {
            $check = Test-WscAdfsWindowRedirect -Context $Context
            if ($check.Code -notin 'SignInPage', 'SignedIn') {
                $redirect = Get-WscRedirectUri -Configuration $cfg -Endpoints $ep
                $why = "AD FS does not accept the redirect URI of the sign-in window ($redirect) for the client $($cfg.ClientId): $($check.Text)"
                if ([string]$cfg.SignIn -eq 'Window') { throw "$why Add it to the client (Set-AdfsNativeClientApplication -RedirectUri), or run with -SignIn DeviceCode." }
                Write-WscItem Info "$why Sign-in with a device code instead." -Icon Key
                $mode = [pscustomobject]@{ Mode = 'DeviceCode'; Browser = $null; Reason = 'redirect URI of the window not accepted by AD FS' }
            }
        }
        elseif ($mode.Mode -eq 'DeviceCode' -and $mode.Reason) {
            Write-WscItem Info "No sign-in window here ($($mode.Reason)): sign-in with a device code, on any device." -Icon Key
        }
        $window = $null
        if ($mode.Mode -eq 'Window') {
            try { $window = Invoke-WscWindowAuthentication -Context $Context -Browser $mode.Browser }
            catch [NotSupportedException] {
                if ([string]$cfg.SignIn -eq 'Window') { throw "$($_.Exception.Message) Run with -SignIn DeviceCode." }
                Write-WscItem Info "$($_.Exception.Message) Sign-in with a device code instead." -Icon Key
                $mode = [pscustomobject]@{ Mode = 'DeviceCode'; Browser = $null; Reason = $_.Exception.Message -replace '^The sign-in window could not start: ', 'could not start: ' }
            }
        }
        $Context.SignIn = $mode.Mode
        $who = if ([string]$cfg.Context -eq 'Delegated') { "through the application $($cfg.ClientId) (delegated permission, on behalf of the user)" } else { "with the client $($cfg.ClientId)" }
        if ($window) {
            $Context.AccessToken = $window.AccessToken
            Add-WscStep $Context $stage 'Sign-in window' Passed "Access token received from $server $who after the sign-in in the window (authorization code with PKCE)." ([ordered]@{
                    Source = "$server sign-in window"; Browser = $mode.Browser.Name; Flow = 'authorization code with PKCE'; RedirectUri = $window.RedirectUri; ClientId = $cfg.ClientId
                    Confidential = [bool]$Context.ClientSecret; Scope = $ep.Scope; AuthorizeEndpoint = $ep.AuthorizeEndpoint; TokenEndpoint = $ep.TokenEndpoint })
        }
        else {
            $Context.AccessToken = Invoke-WscDeviceCodeAuthentication -Configuration $cfg -Endpoints $ep -HttpClient $Context.HttpClient -UserAgent ([string]$cfg.UserAgent)
            $d = [ordered]@{ Source = "$server device code"; ClientId = $cfg.ClientId; Scope = $ep.Scope; TokenEndpoint = $ep.TokenEndpoint }
            if ($mode.Reason) { $d.SignInWindow = "not used: $($mode.Reason)" }
            Add-WscStep $Context $stage 'Device-code sign-in' Passed "Access token received from $server $who." $d
        }
    }
    $tenant = if ($entra) { [string]$Context.TenantId } else { $null }
    $check = if ($Context.Protocol -eq 'Graph') { Test-WscGraphTokenClaims -Claims (Get-WscTokenClaims -AccessToken $Context.AccessToken) -Configuration $cfg -TenantId $tenant -Access $Context.Access -Stages $Context.Stages }
    else { Test-WscTokenClaims -Claims (Get-WscTokenClaims -AccessToken $Context.AccessToken) -Endpoints $ep -Configuration $cfg -TenantId $tenant }
    $Context.Token = $check.Details
    Add-WscStep $Context $stage 'Token claims' $check.Status $check.Message $check.Details
    if ($check.Status -eq 'Failed') { $Context.Stop = $true }
}

function Invoke-WscMailboxTest {
    <#
    .SYNOPSIS
        Runs one scenario and returns the result (steps, folders, messages, free/busy, actions, trace).
    .PARAMETER Configuration
        Settings hashtable (Import-WscConfiguration). Missing keys take their default value.
    .PARAMETER TestType
        Scenario. Default: the TestType of the configuration (Test.DefaultType).
    .PARAMETER Credential
        Basic: user name and password. Windows: the account to use (default: the current Windows account).
    .PARAMETER ClientSecret
        Delegated (confidential application) or Application: the client secret. Never written anywhere.
    .PARAMETER AccessToken
        Token already obtained (integration only). It is never written anywhere.
    .PARAMETER ItemId
        ReplyMail, MoveMail, DeleteMail: the EWS ID of the message to work on.
    .PARAMETER ItemSubject
        ReplyMail, MoveMail, DeleteMail: the most recent Inbox message whose subject contains this text.
    .PARAMETER Quiet
        No console output (log and GUI still receive the lines).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Configuration,
        [string]$TestType,
        [pscredential]$Credential,
        [securestring]$ClientSecret,
        [string]$AccessToken,
        [string]$ItemId,
        [string]$ItemSubject,
        [switch]$Quiet
    )

    $cfg = Get-WscDefaultConfiguration
    foreach ($key in $Configuration.Keys) { $cfg[$key] = $Configuration[$key] }
    if ($TestType) { $cfg.TestType = $TestType }
    $validation = Test-WscConfiguration -Configuration $cfg
    if (-not $validation.IsValid) { throw ("Invalid configuration:`n - " + ($validation.Problems -join "`n - ")) }
    $cfg = Resolve-WscClientSettings -Configuration $cfg
    $auth = [string]$cfg.Authentication
    # Graph forced: no Autodiscover, the mailbox is reached through /users/{address}.
    $stages = @(Get-WscScenarioStages -TestType $cfg.TestType -Discovery $(if ($cfg.Protocol -eq 'Graph') { 'Manual' } else { $cfg.Discovery }))
    $signsIn = $stages -contains 'SignIn'
    if ($auth -eq 'Basic' -and $signsIn -and -not $Credential) { throw 'Basic authentication needs a user name and a password: -Credential (Get-Credential).' }
    if ($auth -eq 'OAuth' -and [string]$cfg.Context -eq 'Application' -and $signsIn -and -not $ClientSecret -and -not $AccessToken -and -not ([string]$cfg.CertificateThumbprint -and [string]$cfg.Authority -ne 'ADFS')) {
        throw 'The Application context needs the client secret (-ClientSecret) or, with Entra ID, a certificate (Identity.CertificateThumbprint).'
    }

    $previous = @{ Quiet = $script:Quiet; Trace = $script:WscTrace; Cookies = $script:WscCookies; Session = $script:WscWindowsSession }
    $script:Quiet = [bool]$Quiet
    $scenario = $script:Scenarios | Where-Object Name -eq $cfg.TestType
    $started = [DateTimeOffset]::UtcNow
    $context = @{
        Config          = $cfg
        Endpoints       = Resolve-WscEndpoints -Configuration $cfg
        Access          = Resolve-WscAccess -Configuration $cfg
        Protocol        = $null
        Stages          = $stages
        EwsUrlSource    = if ($cfg.Discovery -eq 'Manual') { 'Configuration (manual)' } else { $null }
        AuthoritySource = 'Configuration'
        TenantId        = $null
        AccessToken     = if ($auth -eq 'OAuth') { $AccessToken } else { $null }
        Credential      = if ($auth -in 'Basic', 'Windows') { $Credential } else { $null }
        ClientSecret    = $ClientSecret
        ItemId          = $ItemId
        ItemSubject     = $ItemSubject
        SignIn          = $null
        Token           = $null
        Windows         = $null
        WindowsSession  = $false
        WindowsPackageUsed = $null
        ChannelBindings = @{}
        NtlmServer      = $null
        OfferedSchemes  = @()
        HttpClient      = $null
        ServerVersion   = $null
        LastServer      = $null
        AffinityCookie  = $false
        Throttled       = 0
        Mailbox         = $null
        Folders         = @()
        Messages        = @()
        Message         = $null
        FreeBusy        = @()
        FreeBusyWindow  = $null
        TestFolder      = $null
        TestMessage     = $null
        Actions         = [Collections.Generic.List[object]]::new()
        Steps           = [Collections.Generic.List[object]]::new()
        Trace           = [Collections.Generic.List[object]]::new()
        Clock           = [Diagnostics.Stopwatch]::StartNew()
        Stop            = $false
    }
    $handler = [Net.Http.HttpClientHandler]::new()
    $handler.AllowAutoRedirect = $false
    $handler.UseDefaultCredentials = $false
    $handler.CookieContainer = [Net.CookieContainer]::new()
    $handler.AutomaticDecompression = [Net.DecompressionMethods]::GZip -bor [Net.DecompressionMethods]::Deflate
    # The legs of an NTLM handshake must share one TCP connection.
    if ($auth -eq 'Windows') { $handler.MaxConnectionsPerServer = 1 }
    $script:WscTrace = $context.Trace
    $script:WscCookies = $handler.CookieContainer
    $script:WscWindowsSession = $false
    $context.HttpClient = [Net.Http.HttpClient]::new($handler)
    $context.HttpClient.Timeout = [TimeSpan]::FromSeconds([int]$cfg.HttpTimeoutSeconds)
    $context.HttpClient.DefaultRequestHeaders.AcceptEncoding.ParseAdd('gzip, deflate')

    try {
        for ($i = 0; $i -lt $stages.Count; $i++) {
            $stage = $stages[$i]
            $title = Get-WscStageTitle -Stage $stage -Configuration $cfg
            $script:WscTraceStage = $stage
            Write-WscStep ($i + 1) $stages.Count $title -Icon $script:StageInfo[$stage].Icon
            if ($context.Stop) {
                Add-WscStep $context $stage $title Skipped 'Not run: an earlier step failed or was blocked.'
                continue
            }
            if (-not $context.Protocol -and $stage -ne 'Autodiscover') {
                try { Set-WscProtocol -Context $context -Stage $stage } catch { Add-WscStep $context $stage 'Protocol' Failed $_.Exception.Message; $context.Stop = $true }
                if ($context.Stop) { continue }
            }
            if (-not $context.Endpoints.EwsUrl -and $stage -ne 'Autodiscover' -and $context.Protocol -ne 'Graph') {
                Add-WscStep $context $stage $title Failed 'No EWS URL: give Target.EwsUrl (-EwsUrl) or use Autodiscover.'
                $context.Stop = $true
                continue
            }
            try {
                switch ($stage) {
                    'Autodiscover' { Invoke-WscStageAutodiscover -Context $context }
                    'Discovery' { Invoke-WscStageDiscovery -Context $context }
                    'SignIn' { Invoke-WscStageSignIn -Context $context }
                    default { Invoke-WscOperationStage -Context $context -Stage $stage -Title $title }
                }
            }
            catch {
                Add-WscStep $context $stage $title Failed $_.Exception.Message
                $context.Stop = $true
            }
        }
    }
    finally {
        $context.HttpClient.Dispose()
        $handler.Dispose()
        $script:Quiet = $previous.Quiet
        $script:WscTrace = $previous.Trace
        $script:WscCookies = $previous.Cookies
        $script:WscWindowsSession = $previous.Session
        $script:WscTraceStage = $null
    }

    $steps = @($context.Steps)
    $counts = [ordered]@{}
    foreach ($s in 'Passed', 'Warning', 'Blocked', 'Failed', 'Skipped') { $counts[$s] = @($steps | Where-Object Status -eq $s).Count }
    $firstProblem = $steps | Where-Object { $_.Status -in 'Failed', 'Blocked' } | Select-Object -First 1
    $completed = [DateTimeOffset]::UtcNow
    $ep = $context.Endpoints
    [pscustomobject]@{
        Tool            = 'Web Services Client for Exchange'
        Version         = $script:ToolVersion
        Status          = Get-WscOverallStatus -Steps $steps
        TestType        = $cfg.TestType
        Scenario        = $scenario.DisplayName
        StartedUtc      = $started.ToString('yyyy-MM-ddTHH:mm:ssZ')
        CompletedUtc    = $completed.ToString('yyyy-MM-ddTHH:mm:ssZ')
        DurationSeconds = [Math]::Round(($completed - $started).TotalSeconds, 1)
        Mailbox         = [string]$cfg.Mailbox
        SignInUser      = if ($context.Credential) { $context.Credential.UserName } elseif ($auth -eq 'Windows') { "$([Environment]::UserDomainName)\$([Environment]::UserName)" } elseif ($cfg.SignInUser) { [string]$cfg.SignInUser } else { $null }
        Authentication  = $auth
        Context         = if ($auth -eq 'OAuth') { [string]$cfg.Context } else { 'User' }
        Access          = $context.Access
        WindowsPackage  = if ($auth -eq 'Windows') { [string]$cfg.WindowsPackage } else { $null }
        Windows         = $context.Windows
        Protocol        = if ($context.Protocol) { $context.Protocol } else { [string]$cfg.Protocol }
        ProtocolSetting = [string]$cfg.Protocol
        GraphUrl        = $ep.GraphUrl
        StageTitles     = $($titles = [ordered]@{}; foreach ($s in $stages) { $titles[$s] = Get-WscStageTitle -Stage $s -Configuration $cfg }; $titles)
        Discovery       = [string]$cfg.Discovery
        EwsUrl          = $ep.EwsUrl
        EwsUrlSource    = $context.EwsUrlSource
        ExchangeOnline  = [bool]$ep.ExchangeOnline
        Authority       = if ($auth -ne 'OAuth' -or $cfg.Authority -eq 'Auto') { $null } else { [string]$cfg.Authority }
        AuthorityUrl    = if ($auth -ne 'OAuth') { $null } elseif ($ep.EntraRoot) { $ep.EntraRoot } else { $ep.AdfsRoot }
        AuthoritySource = if ($auth -eq 'OAuth' -and ($ep.EntraRoot -or $ep.AdfsRoot)) { $context.AuthoritySource } else { $null }
        TenantId        = $context.TenantId
        ClientId        = if ($auth -eq 'OAuth') { [string]$cfg.ClientId } else { $null }
        SignIn          = $context.SignIn
        Token           = $context.Token
        UserAgent       = [string]$cfg.UserAgent
        RequestServerVersion = [string]$cfg.RequestServerVersion
        Server          = [pscustomobject]@{
            Version = if ($context.ServerVersion) { $context.ServerVersion.Text } else { $null }
            Schema  = if ($context.ServerVersion) { $context.ServerVersion.Schema } else { $null }
            Route   = $context.LastServer
            Affinity = $context.AffinityCookie
            Throttled = $context.Throttled
        }
        AllowWrite      = [bool]$cfg.AllowWrite
        Folders         = @($context.Folders)
        Messages        = @($context.Messages)
        Message         = $context.Message
        FreeBusy        = @($context.FreeBusy | ForEach-Object { $_ | Select-Object Mailbox, ResponseClass, ResponseCode, MessageText, ViewType, WorkingHours, MergedFreeBusy, WorkingDays, WorkStart, WorkEnd, WorkOffsetMinutes, @{ n = 'Events'; e = { @($_.Events).Count } } })
        FreeBusyWindow  = $context.FreeBusyWindow
        FreeBusyEvents  = @($context.FreeBusy | ForEach-Object { @($_.Events) })
        Actions         = @($context.Actions)
        Counts          = $counts
        Steps           = $steps
        Trace           = @($context.Trace | Select-Object -Property * -ExcludeProperty Collapse, ResponseKey)
        Error           = if ($firstProblem) { "$($firstProblem.Name): $($firstProblem.Message)" } else { $null }
    }
}
