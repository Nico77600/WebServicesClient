<#
.SYNOPSIS
    Web Services Client for Exchange - HTTP layer and trace of the exchanges (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    Every HTTP request of the tool (EWS, Autodiscover, AD FS, Entra ID) goes through Invoke-WscHttp:
      - Send-WscHttpRequest sends it. It is the only function that touches the network; the tests
        and the documentation tool replace it with the simulator.
      - The exchange is added to the trace of the run: the request sent and the response received
        as text (start line, headers, body), SOAP and XML indented, JSON indented, HTML summarised,
        the Negotiate, NTLM and Kerberos tokens decoded (src\WebServicesClient.Windows.ps1).
    Add-WscStep attaches the exchanges recorded since the previous check to that check: the report
    shows, for every check, what the client sent and what it received.

    Secrets never reach the trace: access, refresh and ID tokens, client secrets and assertions,
    device and authorization codes are replaced by their length; for Basic only the user name is
    shown; for NTLM the user, domain and workstation, never the response to the challenge; a
    Kerberos ticket only by its length. Cookies: the affinity cookies of Exchange are shown (they
    name a back-end server, they authenticate nothing), every other cookie is masked.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

# Exchanges of the current run (List of objects), $null outside a run. Stage: current stage.
$script:WscTrace = $null
$script:WscTraceStage = $null
$script:WscTraceLimit = 30000
# Cookie container of the HTTP client of the run (affinity cookies sent back), and whether the connection is
# authenticated by a Windows handshake: both only to describe the requests in the trace.
$script:WscCookies = $null
$script:WscWindowsSession = $false
$script:WscTraceSecrets = @('access_token', 'refresh_token', 'id_token', 'device_code', 'code', 'client_secret', 'code_verifier', 'assertion', 'client_assertion', 'password')
# Cookies shown in the trace: server affinity of Exchange (on-premises and Exchange Online).
$script:WscAffinityCookies = @('X-BackEndOverrideCookie', 'X-BackEndCookie', 'X-BackEndCookie2', 'exchangecookie', 'X-OWA-CANARY-Affinity')
$script:WscReasons = @{
    200 = 'OK'; 301 = 'Moved Permanently'; 302 = 'Found'; 303 = 'See Other'; 307 = 'Temporary Redirect'; 308 = 'Permanent Redirect'
    400 = 'Bad Request'; 401 = 'Unauthorized'; 403 = 'Forbidden'; 404 = 'Not Found'; 429 = 'Too Many Requests'; 440 = 'Login Timeout'
    500 = 'Internal Server Error'; 502 = 'Bad Gateway'; 503 = 'Service Unavailable'; 504 = 'Gateway Timeout'
}
# Response headers that tell which Exchange servers answered (shown in the details of the checks).
$script:WscServerHeaders = @('X-FEServer', 'X-BEServer', 'X-CalculatedBETarget', 'X-DiagInfo', 'X-TargetBEServer', 'X-BackEndHttpStatus', 'request-id', 'client-request-id', 'X-MS-Diagnostics', 'X-ProxyBackendServerStatus', 'x-ms-ags-diagnostic')

function Send-WscHttpRequest {
    <# Sends one request, never follows a redirect. The only function of the tool that touches the network. #>
    param([Parameter(Mandatory = $true)][Net.Http.HttpClient]$HttpClient, [Parameter(Mandatory = $true)][Net.Http.HttpRequestMessage]$Request)

    $response = $HttpClient.SendAsync($Request).GetAwaiter().GetResult()
    try {
        $bytes = $response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
        $headers = @{}
        $lines = [Collections.Generic.List[string]]::new()
        foreach ($set in @($response.Headers.NonValidated, $response.Content.Headers.NonValidated)) {
            foreach ($header in $set) {
                $values = @($header.Value)
                $headers[$header.Key] = $values -join ','
                foreach ($value in $values) { [void]$lines.Add("$($header.Key): $value") }
            }
        }
        [pscustomobject]@{
            StatusCode  = [int]$response.StatusCode
            Reason      = $response.ReasonPhrase
            Headers     = $headers
            HeaderLines = @($lines)
            Challenges  = @(foreach ($challenge in $response.Headers.WwwAuthenticate) { $challenge.ToString() })
            Body        = [byte[]]$bytes
        }
    }
    finally {
        $response.Dispose()
    }
}

function Get-WscNetworkErrorText {
    <# Readable reason of a request that got no answer (network or TLS error), or $null. #>
    param([Parameter(Mandatory = $true)][Exception]$Exception)

    $chain = [Collections.Generic.List[Exception]]::new()
    for ($e = $Exception; $e; $e = $e.InnerException) { $chain.Add($e) }
    $outer = $chain | Where-Object {
        $_ -is [Net.Http.HttpRequestException] -or $_ -is [IO.IOException] -or $_ -is [Net.Sockets.SocketException] -or
        $_ -is [Security.Authentication.AuthenticationException] -or $_ -is [OperationCanceledException]
    } | Select-Object -First 1
    if (-not $outer) { return $null }
    $text = ($outer.Message -replace '[,.]?\s*see inner exception\.?\s*$', '').Trim()
    if ($outer -is [OperationCanceledException]) { $text = 'No answer within the timeout (Test.HttpTimeoutSeconds).' }
    $inner = $chain[$chain.Count - 1]
    if ($inner -ne $outer -and $inner -isnot [OperationCanceledException] -and $inner -isnot [TimeoutException] -and
        -not [string]::IsNullOrWhiteSpace($inner.Message) -and -not $text.Contains($inner.Message.Trim())) {
        $text = "$($text.TrimEnd('.')): $($inner.Message.Trim())"
    }
    return $text
}

function Invoke-WscHttp {
    <#
        Sends a request (Send-WscHttpRequest) and adds the exchange to the trace. Collapse: requests
        repeated with the same answer (token polling, waiting for a message) are counted on the first one.
    #>
    param(
        [Parameter(Mandatory = $true)][Net.Http.HttpClient]$HttpClient,
        [Parameter(Mandatory = $true)][Net.Http.HttpRequestMessage]$Request,
        [string]$Note,
        [string]$Collapse,
        [string]$Label,
        [string]$Operation
    )

    $body = if ($Request.Content) { [byte[]]$Request.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult() } else { [byte[]]@() }
    $clock = [Diagnostics.Stopwatch]::StartNew()
    try {
        $response = Send-WscHttpRequest -HttpClient $HttpClient -Request $Request
    }
    catch {
        $failure = $_.Exception
        while ($failure.InnerException) { $failure = $failure.InnerException }
        Add-WscTraceExchange -Request $Request -RequestBody $body -Failure $failure.Message -DurationMs $clock.ElapsedMilliseconds -Note $Note -Label $Label -Operation $Operation
        $text = Get-WscNetworkErrorText -Exception $_.Exception
        if ($text) { throw [Net.Http.HttpRequestException]::new($text, $_.Exception) }
        throw
    }
    Add-WscTraceExchange -Request $Request -RequestBody $body -Response $response -DurationMs $clock.ElapsedMilliseconds -Note $Note -Collapse $Collapse -Label $Label -Operation $Operation
    return $response
}

function Protect-WscCookieList {
    <# Cookie header value: the affinity cookies of Exchange shown, the others replaced by <hidden>. #>
    param([AllowEmptyString()][string]$Value, [switch]$SetCookie)

    $show = { param([string]$Name) $script:WscAffinityCookies -contains $Name.Trim() }
    if ($SetCookie) {
        $m = [regex]::Match($Value, '^\s*([^=;]+)=([^;]*)(.*)$')
        if (-not $m.Success) { return $Value }
        if (& $show $m.Groups[1].Value) { return $Value }
        return "$($m.Groups[1].Value)=<hidden, $($m.Groups[2].Value.Length) characters>$($m.Groups[3].Value)"
    }
    return [regex]::Replace($Value, '([^=;\s]+)=([^;]*)', {
            param($m)
            if (& $show $m.Groups[1].Value) { $m.Value } else { "$($m.Groups[1].Value)=<hidden>" }
        })
}

function Protect-WscHeaderValue {
    <# Header value as written to the trace: secrets replaced by their length, Windows tokens decoded, affinity cookies shown. #>
    param([Parameter(Mandatory = $true)][string]$Name, [AllowEmptyString()][string]$Value)

    if ($Name -ieq 'Authorization' -or $Name -ieq 'WWW-Authenticate') {
        $parts = $Value.Trim().Split(' ', 2)
        if ($parts.Count -lt 2 -or [string]::IsNullOrWhiteSpace($parts[1])) { return $Value }
        $scheme = $parts[0]
        if ($scheme -iin 'Negotiate', 'NTLM', 'Kerberos') {
            return "$scheme <$(Format-WscNegotiateSummary -Base64 $parts[1].Trim())>"
        }
        if ($Name -ieq 'WWW-Authenticate') { return $Value }
        if ($scheme -ieq 'Basic') {
            $decoded = $null
            try { $decoded = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($parts[1].Trim())) } catch { $decoded = $null }
            if ($decoded -and $decoded.Contains(':')) { return "Basic <user $($decoded.Split(':', 2)[0]), password never written>" }
            return "Basic <credentials: $($parts[1].Length) characters, never written>"
        }
        if ($parts[1] -eq $script:InvalidToken) { return $Value }
        return "$scheme <access token: $($parts[1].Length) characters, never written>"
    }
    if ($Name -ieq 'Set-Cookie') { return Protect-WscCookieList -Value $Value -SetCookie }
    if ($Name -ieq 'Cookie') { return Protect-WscCookieList -Value $Value }
    return $Value
}

function Limit-WscTraceText {
    param([AllowEmptyString()][string]$Text, [int]$Limit = $script:WscTraceLimit)
    if ($Text.Length -le $Limit) { return $Text }
    return $Text.Substring(0, $Limit) + "`n... (cut: $($Text.Length) characters in total)"
}

function Format-WscXml {
    <# XML (SOAP envelope, Autodiscover) indented; the text unchanged when it is not XML. #>
    param([Parameter(Mandatory = $true)][string]$Text)

    try {
        $doc = [Xml.XmlDocument]::new()
        $doc.PreserveWhitespace = $false
        $doc.LoadXml($Text)
        $builder = [Text.StringBuilder]::new()
        $settings = [Xml.XmlWriterSettings]::new()
        $settings.Indent = $true
        $settings.IndentChars = '  '
        $settings.OmitXmlDeclaration = $true
        $writer = [Xml.XmlWriter]::Create($builder, $settings)
        try { $doc.Save($writer) } finally { $writer.Dispose() }
        return $builder.ToString()
    }
    catch {
        return $Text
    }
}

function Format-WscFormBody {
    <# application/x-www-form-urlencoded body, one field per line, URL-decoded, secrets masked. #>
    param([AllowEmptyString()][string]$Text)

    $lines = foreach ($pair in $Text.Split('&')) {
        if (-not $pair) { continue }
        $name, $value = $pair.Split('=', 2)
        $name = [Uri]::UnescapeDataString($name.Replace('+', ' '))
        $value = if ($null -ne $value) { [Uri]::UnescapeDataString($value.Replace('+', ' ')) } else { '' }
        if ($name -in $script:WscTraceSecrets) { $value = "<$($value.Length) characters, never written>" }
        "$name=$value"
    }
    return "# form fields, one per line, URL-decoded`n" + (@($lines) -join "`n")
}

function Format-WscJsonBody {
    <# JSON body indented, secrets masked. Throws if the text is not JSON. #>
    param([Parameter(Mandatory = $true)][string]$Text)

    $value = $Text | ConvertFrom-Json -AsHashtable -ErrorAction Stop
    if ($value -is [System.Collections.IDictionary]) {
        foreach ($key in @($value.Keys)) {
            if ($key -in $script:WscTraceSecrets -and $null -ne $value[$key]) { $value[$key] = "<$($key): $(([string]$value[$key]).Length) characters, never written>" }
        }
    }
    return (ConvertTo-Json -InputObject $value -Depth 8)
}

function Format-WscHtmlSummary {
    <# An HTML page summarised: title, AD FS error (MSIS), Entra ID error (AADSTS), or the beginning of the visible text. #>
    param([AllowEmptyString()][string]$Text, [int]$Length)

    $lines = [Collections.Generic.List[string]]::new()
    [void]$lines.Add("# HTML page, $Length bytes: summary")
    $title = [regex]::Match($Text, '<title[^>]*>(.*?)</title>', 'IgnoreCase, Singleline')
    if ($title.Success) { [void]$lines.Add('Title: ' + [Net.WebUtility]::HtmlDecode($title.Groups[1].Value).Trim()) }
    $error = [regex]::Match($Text, '(MSIS\d{4}|AADSTS\d+)[^<"\\]{0,240}')
    if ($error.Success) { [void]$lines.Add('Error: ' + [Net.WebUtility]::HtmlDecode($error.Value).Trim()) }
    if ($Text -match 'passwordInput|userNameInput|loginForm') { [void]$lines.Add('Sign-in form: user name and password fields') }
    if (-not $error.Success -and $Text -notmatch 'passwordInput|userNameInput|loginForm') {
        $visible = [regex]::Replace($Text, '<(script|style)[^>]*>.*?</\1>', ' ', 'IgnoreCase, Singleline')
        $visible = [Net.WebUtility]::HtmlDecode([regex]::Replace($visible, '<[^>]+>', ' '))
        $visible = [regex]::Replace($visible, '\s+', ' ').Trim()
        if ($visible) { [void]$lines.Add('Text: ' + $(if ($visible.Length -gt 600) { $visible.Substring(0, 600) + '...' } else { $visible })) }
    }
    return $lines -join "`n"
}

function Format-WscTraceBody {
    <# Body of a request or a response as text: SOAP and XML indented, form and JSON readable, HTML summarised. #>
    param([AllowNull()][byte[]]$Body, [string]$ContentType)

    if ($null -eq $Body -or $Body.Length -eq 0) { return '' }
    $type = ([string]$ContentType).ToLowerInvariant()
    $text = [Text.Encoding]::UTF8.GetString($Body).TrimStart([char]0xFEFF)
    if ($type -like '*x-www-form-urlencoded*') { return Format-WscFormBody $text }
    if ($type -like '*json*' -or $text.TrimStart().StartsWith('{')) {
        try { return Format-WscJsonBody $text } catch { }
    }
    if ($type -like '*html*' -or $text -match '^\s*<(!doctype|html)') { return Format-WscHtmlSummary -Text $text -Length $Body.Length }
    if ($type -like '*xml*' -or $text.TrimStart().StartsWith('<')) {
        $kind = if ($text -match 'soap:Envelope|s:Envelope|Envelope xmlns') { 'SOAP' } else { 'XML' }
        return "# $kind, $($Body.Length) bytes, indented`n" + (Format-WscXml $text)
    }
    return Limit-WscTraceText $text 4000
}

function Get-WscTraceRequestText {
    <# The request as sent: start line, headers (secrets masked), blank line, body. #>
    param([Parameter(Mandatory = $true)][Net.Http.HttpRequestMessage]$Request, [byte[]]$Body, [string]$CookieHeader)

    $uri = $Request.RequestUri
    $lines = [Collections.Generic.List[string]]::new()
    [void]$lines.Add("$($Request.Method.Method) $($uri.PathAndQuery) HTTP/1.1")
    [void]$lines.Add("Host: $($uri.Authority)")
    $contentType = $null
    foreach ($set in @($Request.Headers.NonValidated, $(if ($Request.Content) { $Request.Content.Headers.NonValidated }))) {
        if ($null -eq $set) { continue }
        foreach ($header in $set) {
            if ($header.Key -ieq 'Content-Length') { continue }
            $value = $header.Value.ToString()
            if ($header.Key -ieq 'Content-Type') { $contentType = $value }
            [void]$lines.Add("$($header.Key): $(Protect-WscHeaderValue -Name $header.Key -Value $value)")
        }
    }
    # Cookies added by the cookie container of the client (affinity): not in the request object itself.
    if ($CookieHeader) { [void]$lines.Add("Cookie: $(Protect-WscHeaderValue -Name 'Cookie' -Value $CookieHeader)") }
    if ($Request.Content) { [void]$lines.Add("Content-Length: $($Body.Length)") }
    $text = $lines -join "`n"
    $bodyText = Format-WscTraceBody -Body $Body -ContentType $contentType
    if (-not $bodyText -and $uri.Query.Length -gt 1 -and ($uri.Query.Split('&').Count -ge 3)) {
        $bodyText = (Format-WscFormBody $uri.Query.TrimStart('?')).Replace('# form fields, one per line', '# query parameters, one per line')
    }
    if ($bodyText) { $text += "`n`n" + $bodyText }
    return Limit-WscTraceText $text
}

function Get-WscTraceResponseText {
    <# The response as received: status line, headers (secrets masked, Windows tokens decoded), blank line, body. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Response)

    $code = [int]$Response.StatusCode
    $reason = [string](Get-WscField $Response 'Reason')
    if (-not $reason) { $reason = if ($script:WscReasons.ContainsKey($code)) { $script:WscReasons[$code] } else { '' } }
    $lines = [Collections.Generic.List[string]]::new()
    [void]$lines.Add("HTTP/1.1 $code $reason".TrimEnd())
    $headerLines = @(Get-WscField $Response 'HeaderLines' | Where-Object { $_ })
    if (-not $headerLines.Count) {
        # Response without the raw header lines (simulator): its headers, one line per challenge.
        $headerLines = @(foreach ($c in @(Get-WscField $Response 'Challenges')) { "WWW-Authenticate: $c" })
        $headers = Get-WscField $Response 'Headers'
        if ($headers) { $headerLines += @(foreach ($key in ($headers.Keys | Sort-Object)) { if ($key -ine 'WWW-Authenticate') { "$($key): $($headers[$key])" } }) }
    }
    $contentType = $null
    foreach ($line in $headerLines) {
        $name, $value = ([string]$line).Split(':', 2)
        $value = ([string]$value).Trim()
        if ($name -ieq 'Content-Type') { $contentType = $value }
        [void]$lines.Add("$($name): $(Protect-WscHeaderValue -Name $name -Value $value)")
    }
    $text = $lines -join "`n"
    $bodyText = Format-WscTraceBody -Body ([byte[]](Get-WscField $Response 'Body')) -ContentType $contentType
    if ($bodyText) { $text += "`n`n" + $bodyText }
    return Limit-WscTraceText $text
}

function Add-WscTraceEntry {
    <# Adds one exchange to the trace of the run (nothing outside a run). #>
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$Request,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Response,
        [AllowNull()][object]$StatusCode,
        [string]$Reason,
        [long]$DurationMs,
        [string]$Note,
        [string]$Collapse,
        [string]$Label,
        [string]$Operation,
        [string]$Server
    )

    if ($null -eq $script:WscTrace) { return }
    if ($Collapse -and $script:WscTrace.Count) {
        $last = $script:WscTrace[$script:WscTrace.Count - 1]
        if ($last.Collapse -eq $Collapse -and $null -eq $last.Step -and $last.StatusCode -eq $StatusCode -and $last.ResponseKey -eq [string]$Response.GetHashCode()) {
            $last.Repeated++
            $last.Note = "Sent $($last.Repeated) times with the same answer (waiting): only the first one is shown."
            return
        }
    }
    $entry = [pscustomobject]@{
        Sequence     = $script:WscTrace.Count + 1
        TimestampUtc = [DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        Stage        = $script:WscTraceStage
        Step         = $null
        StepName     = $null
        Method       = $Method
        Url          = $Url
        Operation    = $Operation
        Label        = $Label
        StatusCode   = $StatusCode
        Reason       = $Reason
        Server       = $Server
        DurationMs   = [int]$DurationMs
        Repeated     = 1
        Note         = $Note
        Collapse     = $Collapse
        ResponseKey  = if ($Collapse) { [string]$Response.GetHashCode() } else { $null }
        Request      = $Request
        Response     = $Response
    }
    $script:WscTrace.Add($entry)
    $status = if ($null -ne $StatusCode) { "$StatusCode $Reason".TrimEnd() } else { 'no response' }
    $what = if ($Operation) { " [$Operation]" } else { '' }
    Write-WscLog 'INFO' ("HTTP #{0} {1} {2}{3} -> {4} ({5} ms)" -f $entry.Sequence, $Method, $Url, $what, $status, $entry.DurationMs)
}

function Get-WscResponseServer {
    <# The Exchange servers named by a response: front end, back end (X-FEServer, X-BEServer, X-CalculatedBETarget, X-DiagInfo). #>
    param([AllowNull()][pscustomobject]$Response)

    $headers = Get-WscField $Response 'Headers'
    if (-not $headers) { return $null }
    $get = { param([string]$Name) foreach ($k in $headers.Keys) { if ($k -ieq $Name) { return [string]$headers[$k] } }; $null }
    $fe = & $get 'X-FEServer'
    $be = & $get 'X-BEServer'
    if (-not $be) { $be = & $get 'X-CalculatedBETarget' }
    $diag = & $get 'X-DiagInfo'
    $ags = & $get 'x-ms-ags-diagnostic'
    if ($ags) {
        # Microsoft Graph: {"ServerInfo":{"DataCenter":"France Central","Slice":"E","Ring":"3","ScaleUnit":"001","RoleInstance":"PA1PEPF..."}}
        try {
            $info = ($ags | ConvertFrom-Json -ErrorAction Stop).ServerInfo
            return ("Graph {0} ({1}, ring {2})" -f $info.RoleInstance, $info.DataCenter, $info.Ring)
        }
        catch { return "Graph $ags" }
    }
    $parts = @()
    if ($fe) { $parts += "FE $fe" }
    if ($be) { $parts += "BE $be" }
    elseif ($diag) { $parts += "server $diag" }
    return ($parts -join ' > ')
}

function Add-WscTraceExchange {
    param(
        [Parameter(Mandatory = $true)][Net.Http.HttpRequestMessage]$Request,
        [byte[]]$RequestBody,
        [pscustomobject]$Response,
        [string]$Failure,
        [long]$DurationMs,
        [string]$Note,
        [string]$Collapse,
        [string]$Label,
        [string]$Operation
    )

    if ($null -eq $script:WscTrace) { return }
    $cookie = $null
    if ($script:WscCookies) {
        $cookie = $script:WscCookies.GetCookieHeader($Request.RequestUri)
        if ([string]::IsNullOrEmpty($cookie)) { $cookie = $null }
    }
    $requestText = Get-WscTraceRequestText -Request $Request -Body $RequestBody -CookieHeader $cookie
    # What makes this request different from its neighbours: the credentials sent, affinity, impersonation.
    $hints = [Collections.Generic.List[string]]::new()
    if ($Label) { $hints.Add($Label) }
    $operation = $Operation
    $isEws = $Request.RequestUri.AbsolutePath -match '/EWS/|/autodiscover/' -or $Request.RequestUri.Host -eq 'graph.microsoft.com'
    if ($isEws) {
        if (-not $operation -and $RequestBody -and $RequestBody.Length) {
            $m = [regex]::Match([Text.Encoding]::UTF8.GetString($RequestBody), '<soap:Body>\s*<(?:\w+:)?(\w+?)(?:Request)?[\s>/]')
            if ($m.Success) { $operation = $m.Groups[1].Value }
        }
        $auth = $Request.Headers.Authorization
        if (-not $auth) { $hints.Add($(if ($script:WscWindowsSession) { 'Windows session of the connection' } else { 'no credentials' })) }
        elseif (-not $auth.Parameter) { $hints.Add("empty $($auth.Scheme) header") }
        elseif ($auth.Scheme -ieq 'Basic') {
            $wrong = (Protect-WscHeaderValue -Name 'Authorization' -Value "Basic $($auth.Parameter)") -like "Basic <user $($script:InvalidBasicUserPrefix)*"
            $hints.Add($(if ($wrong) { 'wrong user name and password' } else { 'user name and password (Basic)' }))
        }
        elseif ($auth.Scheme -iin 'Negotiate', 'NTLM', 'Kerberos') { $hints.Add((Get-WscNegotiateKind -Base64 $auth.Parameter)) }
        elseif ($auth.Parameter -eq $script:InvalidToken) { $hints.Add('forged token') }
        else { $hints.Add('access token') }
        $values = $null
        if ($Request.Headers.TryGetValues('X-AnchorMailbox', [ref]$values)) { $hints.Add("anchor $(@($values) -join '')") }
        if ($RequestBody -and [Text.Encoding]::UTF8.GetString($RequestBody) -match '<t:ExchangeImpersonation>') { $hints.Add('impersonation') }
        if ($cookie -and $cookie -match 'X-BackEnd') { $hints.Add('affinity cookie') }
    }
    if ($RequestBody -and $RequestBody.Length -and $Request.Content -and "$($Request.Content.Headers.ContentType)" -like '*form-urlencoded*') {
        $grant = [regex]::Match([Text.Encoding]::UTF8.GetString($RequestBody), '(?:^|&)grant_type=([^&]*)')
        if ($grant.Success) { $hints.Add('grant ' + ([Uri]::UnescapeDataString($grant.Groups[1].Value) -replace '^urn:ietf:params:oauth:grant-type:', '')) }
    }
    $label = $hints -join ' + '
    if ($Response) {
        $code = [int]$Response.StatusCode
        $reason = [string](Get-WscField $Response 'Reason')
        if (-not $reason -and $script:WscReasons.ContainsKey($code)) { $reason = $script:WscReasons[$code] }
        Add-WscTraceEntry -Method $Request.Method.Method -Url $Request.RequestUri.AbsoluteUri -Request $requestText -Response (Get-WscTraceResponseText -Response $Response) `
            -StatusCode $code -Reason $reason -DurationMs $DurationMs -Note $Note -Collapse $Collapse -Label $label -Operation $operation -Server (Get-WscResponseServer -Response $Response)
    }
    else {
        Add-WscTraceEntry -Method $Request.Method.Method -Url $Request.RequestUri.AbsoluteUri -Request $requestText -Response "No response: $Failure" `
            -StatusCode $null -DurationMs $DurationMs -Note $Note -Label $label -Operation $operation
    }
}

function Complete-WscTraceStep {
    <# Attaches the exchanges recorded since the previous check to this check. Returns their numbers. #>
    param([Parameter(Mandatory = $true)][int]$Step, [Parameter(Mandatory = $true)][string]$Name)

    if ($null -eq $script:WscTrace) { return @() }
    $ids = foreach ($entry in $script:WscTrace) {
        if ($null -eq $entry.Step) { $entry.Step = $Step; $entry.StepName = $Name; $entry.Sequence }
    }
    return @($ids)
}
