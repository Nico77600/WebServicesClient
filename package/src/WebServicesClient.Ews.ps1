<#
.SYNOPSIS
    Web Services Client for Exchange - the EWS protocol: SOAP envelopes, headers, responses, throttling (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    Every EWS request is built and sent by Invoke-WscEws with the good practices of Microsoft:
      - RequestServerVersion in the SOAP header (Ews.RequestServerVersion), ServerVersionInfo read back.
      - X-AnchorMailbox: the mailbox the request works on, so that the front end routes it to the
        back end of that mailbox at once (on-premises 2013+ and Exchange Online).
      - X-PreferServerAffinity: true, and the affinity cookies returned (X-BackEndOverrideCookie,
        X-BackEndCookie) sent back by the cookie container of the run.
      - client-request-id (a new GUID per request) and return-client-request-id: true, so that the
        request is found in the logs of Exchange (request-id, client-request-id of the answer).
      - User-Agent of the tool (Ews.UserAgent), Accept-Encoding gzip and deflate.
      - ExchangeImpersonation header for impersonation, <t:Mailbox> in the folder IDs for delegate access.
      - Throttling: ErrorServerBusy (BackOffMilliseconds), HTTP 429 and 503 (Retry-After): the request
        is sent again after the time asked, Ews.MaxRetries times at most, and never after more than 60 s.
    Authentication: the bearer token (OAuth), the user name and password (Basic), or the Windows handshake
    (src\WebServicesClient.Windows.ps1).

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:EwsNs = @{
    soap = 'http://schemas.xmlsoap.org/soap/envelope/'
    m    = 'http://schemas.microsoft.com/exchange/services/2006/messages'
    t    = 'http://schemas.microsoft.com/exchange/services/2006/types'
    e    = 'http://schemas.microsoft.com/exchange/services/2006/errors'
}
$script:EwsMaxWaitSeconds = 60

function ConvertTo-WscXmlText {
    param([AllowNull()][AllowEmptyString()][string]$Text)
    if ($null -eq $Text) { return '' }
    return [Security.SecurityElement]::Escape($Text)
}

function Get-WscAnchorMailbox {
    <# X-AnchorMailbox of a request: the mailbox it works on. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    return [string]$Context.Config.Mailbox
}

function Get-WscFolderIdXml {
    <#
        Folder ID element: a distinguished folder (inbox, msgfolderroot, sentitems...) of the mailbox, with
        <t:Mailbox> for delegate access (impersonation and the own mailbox need none), or a folder ID.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [string]$Distinguished, [string]$FolderId)

    if ($FolderId) { return '<t:FolderId Id="{0}"/>' -f (ConvertTo-WscXmlText $FolderId) }
    if ($Context.Access -eq 'Delegate') {
        return '<t:DistinguishedFolderId Id="{0}"><t:Mailbox><t:EmailAddress>{1}</t:EmailAddress></t:Mailbox></t:DistinguishedFolderId>' -f $Distinguished, (ConvertTo-WscXmlText ([string]$Context.Config.Mailbox))
    }
    return '<t:DistinguishedFolderId Id="{0}"/>' -f $Distinguished
}

function New-WscSoapEnvelope {
    <# SOAP envelope: RequestServerVersion, ExchangeImpersonation when the mailbox is impersonated, and the body. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Body, [switch]$NoImpersonation)

    $headers = [Text.StringBuilder]::new()
    [void]$headers.Append(('<t:RequestServerVersion Version="{0}"/>' -f [string]$Context.Config.RequestServerVersion))
    if ($Context.Access -eq 'Impersonation' -and -not $NoImpersonation) {
        [void]$headers.Append(('<t:ExchangeImpersonation><t:ConnectingSID><t:SmtpAddress>{0}</t:SmtpAddress></t:ConnectingSID></t:ExchangeImpersonation>' -f (ConvertTo-WscXmlText ([string]$Context.Config.Mailbox))))
    }
    @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="$($script:EwsNs.soap)" xmlns:m="$($script:EwsNs.m)" xmlns:t="$($script:EwsNs.t)">
<soap:Header>$($headers.ToString())</soap:Header>
<soap:Body>$Body</soap:Body>
</soap:Envelope>
"@
}

function New-WscEwsRequestMessage {
    <# The HTTP request of one EWS operation, with the headers of the good practices. #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Context,
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$Soap,
        [Parameter(Mandatory = $true)][string]$Operation,
        [string]$Anchor,
        [ValidateSet('Context', 'None', 'EmptyBearer', 'Token', 'Basic')][string]$Credentials = 'Context',
        [string]$Token,
        [pscredential]$BasicCredential
    )

    $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Post, $Uri)
    $request.Content = [Net.Http.ByteArrayContent]::new([Text.Encoding]::UTF8.GetBytes($Soap))
    $request.Content.Headers.ContentType = [Net.Http.Headers.MediaTypeHeaderValue]::Parse('text/xml; charset=utf-8')
    $add = { param([string]$Name, [string]$Value) $null = $request.Headers.TryAddWithoutValidation($Name, $Value) }
    & $add 'Accept' 'text/xml'
    & $add 'User-Agent' ([string]$Context.Config.UserAgent)
    & $add 'SOAPAction' "`"http://schemas.microsoft.com/exchange/services/2006/messages/$Operation`""
    if ($Anchor) { & $add 'X-AnchorMailbox' $Anchor }
    if ($Context.Config.PreferServerAffinity) { & $add 'X-PreferServerAffinity' 'true' }
    & $add 'client-request-id' ([guid]::NewGuid().ToString())
    & $add 'return-client-request-id' 'true'
    $mode = if ($Credentials -eq 'Context') { [string]$Context.Config.Authentication } else { $Credentials }
    switch ($mode) {
        'OAuth' { if ($Context.AccessToken) { $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $Context.AccessToken) } }
        'Token' { $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $Token) }
        'EmptyBearer' { $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer') }
        'Basic' {
            $cred = if ($BasicCredential) { $BasicCredential } else { $Context.Credential }
            if ($cred) {
                # UTF-8, sent at once (no first anonymous request).
                $pair = '{0}:{1}' -f $cred.UserName, $cred.GetNetworkCredential().Password
                $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Basic', [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair)))
            }
        }
    }
    return $request
}

function Get-WscHeader {
    <# Value of a response header (case-insensitive), or $null. #>
    param([AllowNull()][pscustomobject]$Response, [Parameter(Mandatory = $true)][string]$Name)
    $headers = Get-WscField $Response 'Headers'
    if (-not $headers) { return $null }
    foreach ($k in $headers.Keys) { if ($k -ieq $Name) { return [string]$headers[$k] } }
    return $null
}

function ConvertFrom-WscSoapResponse {
    <#
        The answer of EWS read: XML document and namespaces, ServerVersionInfo, SOAP fault (code, text,
        EWS response code, BackOffMilliseconds), and the ResponseMessage elements (class, code, text).
    #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Response)

    $result = [ordered]@{
        HttpStatus = [int]$Response.StatusCode; Xml = $null; Ns = $null; ServerVersion = $null; Fault = $null
        Messages = @(); ResponseCode = $null; ResponseClass = $null; MessageText = $null
    }
    $body = [byte[]](Get-WscField $Response 'Body')
    if (-not $body -or -not $body.Length) { return [pscustomobject]$result }
    $text = [Text.Encoding]::UTF8.GetString($body).TrimStart([char]0xFEFF)
    if (-not $text.TrimStart().StartsWith('<')) { return [pscustomobject]$result }
    $xml = [Xml.XmlDocument]::new()
    try { $xml.LoadXml($text) } catch { return [pscustomobject]$result }
    $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
    foreach ($k in $script:EwsNs.Keys) { $ns.AddNamespace($k, $script:EwsNs[$k]) }
    $result.Xml = $xml
    $result.Ns = $ns
    $sv = $xml.SelectSingleNode('//t:ServerVersionInfo', $ns)
    if ($sv) {
        $get = { param([string]$n) $a = $sv.Attributes[$n]; if ($a) { $a.Value } else { $null } }
        $result.ServerVersion = [pscustomobject]@{
            Major = & $get 'MajorVersion'; Minor = & $get 'MinorVersion'; Build = & $get 'MajorBuildNumber'; Revision = & $get 'MinorBuildNumber'; Schema = & $get 'Version'
            Text  = '{0}.{1}.{2}.{3}' -f (& $get 'MajorVersion'), (& $get 'MinorVersion'), (& $get 'MajorBuildNumber'), (& $get 'MinorBuildNumber')
        }
    }
    $fault = $xml.SelectSingleNode('//soap:Fault', $ns)
    if ($fault) {
        $code = $fault.SelectSingleNode('detail/*[local-name()="ResponseCode"]')
        $msg = $fault.SelectSingleNode('detail/*[local-name()="Message"]')
        $backOff = $fault.SelectSingleNode('detail//*[local-name()="Value"][@Name="BackOffMilliseconds"]')
        $result.Fault = [pscustomobject]@{
            Code         = [string]$fault.SelectSingleNode('faultcode').InnerText
            Text         = [string]$fault.SelectSingleNode('faultstring').InnerText
            ResponseCode = if ($code) { $code.InnerText } else { $null }
            Message      = if ($msg) { $msg.InnerText } else { $null }
            BackOffMs    = if ($backOff) { [int]$backOff.InnerText } else { $null }
        }
        $result.ResponseCode = $result.Fault.ResponseCode
        $result.ResponseClass = 'Error'
        $result.MessageText = if ($result.Fault.Message) { $result.Fault.Message } else { $result.Fault.Text }
    }
    $messages = foreach ($node in $xml.SelectNodes('//m:ResponseMessages/*', $ns)) {
        [pscustomobject]@{
            Node          = $node
            ResponseClass = [string]$node.GetAttribute('ResponseClass')
            ResponseCode  = [string]$node.SelectSingleNode('m:ResponseCode', $ns).InnerText
            MessageText   = $(if ($node.SelectSingleNode('m:MessageText', $ns)) { [string]$node.SelectSingleNode('m:MessageText', $ns).InnerText } else { $null })
        }
    }
    $result.Messages = @($messages)
    if (-not $fault -and $result.Messages.Count) {
        $first = $result.Messages | Where-Object ResponseClass -ne 'Success' | Select-Object -First 1
        if (-not $first) { $first = $result.Messages[0] }
        $result.ResponseClass = $first.ResponseClass
        $result.ResponseCode = $first.ResponseCode
        $result.MessageText = $first.MessageText
    }
    return [pscustomobject]$result
}

function Invoke-WscEws {
    <#
        Sends one EWS operation (SOAP body without the envelope) with the authentication of the run and
        returns the HTTP response and the EWS answer read (ConvertFrom-WscSoapResponse). Throttled
        requests are sent again after the time Exchange asks. Never throws for an HTTP or EWS error:
        the caller reads HttpStatus, ResponseClass and ResponseCode.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Context,
        [Parameter(Mandatory = $true)][string]$Operation,
        [Parameter(Mandatory = $true)][string]$Body,
        [string]$Anchor,
        [string]$Uri,
        [ValidateSet('Context', 'None', 'EmptyBearer', 'Token', 'Basic')][string]$Credentials = 'Context',
        [string]$Token,
        [pscredential]$BasicCredential,
        [hashtable]$ExtraHeaders,
        [string]$Collapse,
        [switch]$NoImpersonation
    )

    Invoke-WscUiPump
    Assert-WscNotCancelled
    if (-not $Uri) { $Uri = $Context.Endpoints.EwsUrl }
    if (-not $PSBoundParameters.ContainsKey('Anchor')) { $Anchor = Get-WscAnchorMailbox -Context $Context }
    $soap = New-WscSoapEnvelope -Context $Context -Body $Body -NoImpersonation:$NoImpersonation
    $windows = $Credentials -eq 'Context' -and [string]$Context.Config.Authentication -eq 'Windows'
    # Read by the script block through the call stack (a closure would leave the module scope).
    $wscEwsBuild = @{ Context = $Context; Uri = $Uri; Soap = $soap; Operation = $Operation; Anchor = $Anchor; Credentials = $(if ($windows) { 'None' } else { $Credentials }); Token = $Token; Basic = $BasicCredential; Extra = $ExtraHeaders }
    $build = {
        $b = $wscEwsBuild
        $r = New-WscEwsRequestMessage -Context $b.Context -Uri $b.Uri -Soap $b.Soap -Operation $b.Operation -Anchor $b.Anchor -Credentials $b.Credentials -Token $b.Token -BasicCredential $b.Basic
        if ($b.Extra) { foreach ($k in $b.Extra.Keys) { $null = $r.Headers.TryAddWithoutValidation([string]$k, [string]$b.Extra[$k]) } }
        $r
    }
    $retries = [int]$Context.Config.MaxRetries
    $attempt = 0
    $windowsInfo = $null
    while ($true) {
        if ($windows) {
            $windowsInfo = Invoke-WscWindowsRequest -Context $Context -NewRequest $build -Label $Operation
            $response = $windowsInfo.Response
        }
        else {
            $request = & $build
            try { $response = Invoke-WscHttp -HttpClient $Context.HttpClient -Request $request -Label $Operation -Collapse $Collapse } finally { $request.Dispose() }
        }
        $answer = ConvertFrom-WscSoapResponse -Response $response
        $wait = $null
        if ($answer.Fault -and $answer.Fault.ResponseCode -eq 'ErrorServerBusy') { $wait = [Math]::Max(1, [int][Math]::Ceiling(([int]$answer.Fault.BackOffMs) / 1000.0)) }
        elseif ($response.StatusCode -in 429, 503) {
            $after = 0
            $wait = if ([int]::TryParse([string](Get-WscHeader $response 'Retry-After'), [ref]$after) -and $after -gt 0) { $after } else { 5 }
        }
        if ($null -eq $wait -or $attempt -ge $retries) { break }
        if ($wait -gt $script:EwsMaxWaitSeconds) {
            Write-WscItem Warn "Exchange throttles the requests and asks to wait $wait s (more than $($script:EwsMaxWaitSeconds) s): the request is not sent again."
            break
        }
        $attempt++
        $Context.Throttled++
        Write-WscItem Warn "Exchange throttles the requests ($(if ($answer.Fault) { 'ErrorServerBusy' } else { "HTTP $($response.StatusCode)" })): new attempt $attempt/$retries in $wait s, as asked."
        Wait-WscSeconds $wait
    }
    $server = Get-WscResponseServer -Response $response
    if ($server) { $Context.LastServer = $server }
    if ($answer.ServerVersion) { $Context.ServerVersion = $answer.ServerVersion }
    $setCookie = Get-WscHeader $response 'Set-Cookie'
    if ($setCookie -and $setCookie -match 'X-BackEnd(Override)?Cookie') { $Context.AffinityCookie = $true }
    [pscustomobject]@{
        Operation     = $Operation
        Response      = $response
        HttpStatus    = [int]$response.StatusCode
        Xml           = $answer.Xml
        Ns            = $answer.Ns
        Fault         = $answer.Fault
        Messages      = $answer.Messages
        ResponseClass = $answer.ResponseClass
        ResponseCode  = $answer.ResponseCode
        MessageText   = $answer.MessageText
        ServerVersion = $answer.ServerVersion
        Server        = $server
        RequestId     = Get-WscHeader $response 'request-id'
        Diagnostics   = Get-WscHeader $response 'x-ms-diagnostics'
        Windows       = $windowsInfo
        Retries       = $attempt
    }
}

function Get-WscEwsDetails {
    <# Details common to every EWS check: HTTP status, EWS code, servers, request ID, Exchange diagnostics. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Answer, [System.Collections.IDictionary]$More)

    $d = [ordered]@{ Operation = $Answer.Operation; HttpStatus = $Answer.HttpStatus; ResponseCode = $Answer.ResponseCode }
    if ($Answer.MessageText -and $Answer.ResponseClass -ne 'Success') { $d.MessageText = $Answer.MessageText }
    if ($Answer.Server) { $d.Servers = $Answer.Server }
    if ($Answer.RequestId) { $d.RequestId = $Answer.RequestId }
    if ($Answer.Diagnostics) { $d.Diagnostics = $Answer.Diagnostics }
    if ($Answer.Retries) { $d.Retries = $Answer.Retries }
    if ($More) { foreach ($k in $More.Keys) { $d[$k] = $More[$k] } }
    return $d
}

function Get-WscEwsFailureText {
    <#
        Why an EWS request failed, with what to check: HTTP status (401 with the Exchange diagnostics,
        403, 404, 500 without SOAP), SOAP fault or response code of the operation.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][pscustomobject]$Answer)

    $diag = if ($Answer.Diagnostics) { " Exchange diagnostics: $($Answer.Diagnostics)" } else { '' }
    $auth = [string]$Context.Config.Authentication
    $code = [string]$Answer.ResponseCode
    if ($Answer.HttpStatus -eq 401) {
        $why = switch ($auth) {
            'Basic' { 'Exchange refused the user name and password (wrong password, locked account, Basic disabled on the EWS virtual directory, or blocked by the authentication policy: BlockLegacyAuthWebServices).' }
            'Windows' { 'Exchange refused the Windows authentication. The Security log of the front end (event 4625) gives the reason: 0xC000006D with sub-status 0xC000006A wrong password or 0xC0000064 unknown user; 0xC000035B Extended Protection (channel binding refused: a reverse proxy or TLS inspection that presents another certificate, or tokenChecking Require); otherwise Windows authentication disabled on the EWS virtual directory, or a reverse proxy that does not let NTLM through.' }
            default {
                if ([string]$Context.Config.Context -eq 'Application') { 'Exchange refused the application token: application permission full_access_as_app with admin consent (Entra ID), audience of the token, application access policy.' }
                else { 'Exchange refused the token: audience, scope EWS.AccessAsUser.All, OAuth on the EWS virtual directory, authentication policy (BlockModernAuthWebServices), or EWS disabled for the user (EwsEnabled).' }
            }
        }
        return "HTTP 401: $why$diag"
    }
    if ($Answer.HttpStatus -eq 403) {
        $policy = Get-WscHeader $Answer.Response 'X-EWS-Policy-Reason'
        if ($policy -and $Context.Endpoints.ExchangeOnline) {
            return "HTTP 403, X-EWS-Policy-Reason '$policy': EWS is disabled in Exchange Online (retirement started October 2026, full stop April 2027). Use Microsoft Graph for this mailbox; until April 2027 an administrator can still allow it (Set-OrganizationConfig -EwsEnabled `$true and the client ID in EwsAllowedAppIDs).$diag"
        }
        $reason = if ($policy) { " X-EWS-Policy-Reason: $policy." } else { '' }
        return "HTTP 403: Exchange knows the account but refuses EWS (Get-CASMailbox | Format-List EwsEnabled, EwsAllowList, EwsApplicationAccessPolicy; Set-OrganizationConfig -EwsEnabled).$reason$diag"
    }
    if ($Answer.HttpStatus -eq 404) { return "HTTP 404: no EWS at $($Context.Endpoints.EwsUrl) (URL, publishing).$diag" }
    if ($Answer.HttpStatus -eq 302 -or $Answer.HttpStatus -eq 440) { return "HTTP $($Answer.HttpStatus): a reverse proxy or forms authentication answered instead of EWS (pre-authentication on /EWS).$diag" }
    if (-not $Answer.Xml) { return "HTTP $($Answer.HttpStatus) without an EWS answer (no SOAP).$diag" }
    $hint = switch -Regex ($code) {
        '^ErrorImpersonateUserDenied$' { 'the account may not impersonate this mailbox: ApplicationImpersonation role (on-premises), or full_access_as_app and an application access policy (Exchange Online).' }
        '^ErrorImpersonationDenied$' { 'impersonation refused: same causes as ErrorImpersonateUserDenied.' }
        '^ErrorAccessDenied$' { 'no permission on this mailbox or folder: Full Access or folder permissions for delegate access (Add-MailboxPermission, Add-MailboxFolderPermission).' }
        '^ErrorNonExistentMailbox$' { 'Exchange finds no mailbox with this address.' }
        '^ErrorMailboxMoveInProgress$|^ErrorMailboxStoreUnavailable$' { 'the mailbox is being moved or its database is not available: try again later.' }
        '^ErrorInvalidServerVersion$|^ErrorInvalidRequestServerVersion$' { "the server does not accept RequestServerVersion $($Context.Config.RequestServerVersion): set Ews.RequestServerVersion to Exchange2013_SP1." }
        '^ErrorSendAsDenied$' { 'the signed-in account has no Send As (or Send on Behalf) permission on this mailbox.' }
        '^ErrorQuotaExceeded$' { 'the mailbox is full.' }
        '^ErrorServerBusy$' { 'Exchange throttles the requests (EWS throttling policy): wait and try again.' }
        '^ErrorItemNotFound$' { 'the item is not (or no longer) in the mailbox: give another -ItemId.' }
        '^ErrorInvalidIdMalformed$|^ErrorInvalidIdNotAnItemAttachmentId$' { 'the item ID is not a valid EWS ID.' }
        '^ErrorMessageSizeExceeded$' { 'the message is bigger than the size limit.' }
        '^ErrorNoPublicFolderReplicaAvailable$|^ErrorProxyRequestNotAllowed$' { 'the request must go to another server: X-AnchorMailbox and the routing of the front end.' }
        default { $null }
    }
    $text = if ($Answer.MessageText) { $Answer.MessageText.Trim() } else { 'no message' }
    $where = if ($Answer.Fault) { "SOAP fault $code" } else { "$($Answer.Operation) answered $code" }
    return "$where ($text)$(if ($hint) { ": $hint" } else { '.' })$diag"
}

#region Operations: SOAP bodies ------------------------------------------------------------------

function New-WscGetServerTimeZonesBody {
    '<m:GetServerTimeZones ReturnFullTimeZoneData="false"><m:Ids><t:Id>UTC</t:Id></m:Ids></m:GetServerTimeZones>'
}

function New-WscGetFolderBody {
    param([Parameter(Mandatory = $true)][hashtable]$Context, [string[]]$Distinguished = @('inbox'))
    $ids = ($Distinguished | ForEach-Object { Get-WscFolderIdXml -Context $Context -Distinguished $_ }) -join ''
    '<m:GetFolder><m:FolderShape><t:BaseShape>Default</t:BaseShape><t:AdditionalProperties><t:FieldURI FieldURI="folder:FolderClass"/><t:FieldURI FieldURI="folder:ParentFolderId"/></t:AdditionalProperties></m:FolderShape><m:FolderIds>' + $ids + '</m:FolderIds></m:GetFolder>'
}

function New-WscFindFolderBody {
    <# Folders under a parent: Deep (the whole tree) or Shallow, optionally only the one with a display name. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [string]$Parent = 'msgfolderroot', [string]$ParentId, [ValidateSet('Deep', 'Shallow')][string]$Traversal = 'Deep', [string]$DisplayName, [int]$Offset = 0, [int]$Max = 500)

    $restriction = if ($DisplayName) { '<m:Restriction><t:IsEqualTo><t:FieldURI FieldURI="folder:DisplayName"/><t:FieldURIOrConstant><t:Constant Value="{0}"/></t:FieldURIOrConstant></t:IsEqualTo></m:Restriction>' -f (ConvertTo-WscXmlText $DisplayName) } else { '' }
    ('<m:FindFolder Traversal="{0}"><m:FolderShape><t:BaseShape>Default</t:BaseShape><t:AdditionalProperties><t:FieldURI FieldURI="folder:FolderClass"/><t:FieldURI FieldURI="folder:ParentFolderId"/></t:AdditionalProperties></m:FolderShape>' -f $Traversal) +
    ('<m:IndexedPageFolderView MaxEntriesReturned="{0}" Offset="{1}" BasePoint="Beginning"/>' -f $Max, $Offset) + $restriction +
    '<m:ParentFolderIds>' + (Get-WscFolderIdXml -Context $Context -Distinguished $Parent -FolderId $ParentId) + '</m:ParentFolderIds></m:FindFolder>'
}

function New-WscCreateFolderBody {
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Name, [string]$Parent = 'inbox')
    '<m:CreateFolder><m:ParentFolderId>' + (Get-WscFolderIdXml -Context $Context -Distinguished $Parent) + '</m:ParentFolderId><m:Folders><t:Folder><t:FolderClass>IPF.Note</t:FolderClass><t:DisplayName>' + (ConvertTo-WscXmlText $Name) + '</t:DisplayName></t:Folder></m:Folders></m:CreateFolder>'
}

function New-WscFindItemBody {
    <# The most recent items of a folder (headers only), optionally those whose subject contains a text. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [string]$Folder = 'inbox', [string]$FolderId, [int]$Max = 10, [string]$SubjectContains)

    $props = 'item:Subject', 'item:DateTimeReceived', 'message:From', 'message:IsRead', 'item:HasAttachments', 'item:Size', 'item:ItemClass', 'message:InternetMessageId'
    $restriction = if ($SubjectContains) { '<m:Restriction><t:Contains ContainmentMode="Substring" ContainmentComparison="IgnoreCase"><t:FieldURI FieldURI="item:Subject"/><t:Constant Value="{0}"/></t:Contains></m:Restriction>' -f (ConvertTo-WscXmlText $SubjectContains) } else { '' }
    '<m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>' + (($props | ForEach-Object { '<t:FieldURI FieldURI="{0}"/>' -f $_ }) -join '') + '</t:AdditionalProperties></m:ItemShape>' +
    ('<m:IndexedPageItemView MaxEntriesReturned="{0}" Offset="0" BasePoint="Beginning"/>' -f $Max) + $restriction +
    '<m:SortOrder><t:FieldOrder Order="Descending"><t:FieldURI FieldURI="item:DateTimeReceived"/></t:FieldOrder></m:SortOrder><m:ParentFolderIds>' + (Get-WscFolderIdXml -Context $Context -Distinguished $Folder -FolderId $FolderId) + '</m:ParentFolderIds></m:FindItem>'
}

function New-WscGetItemBody {
    param([Parameter(Mandatory = $true)][string]$ItemId)
    $props = 'item:Subject', 'item:DateTimeReceived', 'item:DateTimeSent', 'message:From', 'message:Sender', 'message:ToRecipients', 'message:CcRecipients', 'message:InternetMessageId', 'message:IsRead', 'item:HasAttachments', 'item:Size', 'item:Body', 'item:ParentFolderId'
    '<m:GetItem><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:BodyType>Text</t:BodyType><t:AdditionalProperties>' + (($props | ForEach-Object { '<t:FieldURI FieldURI="{0}"/>' -f $_ }) -join '') +
    '</t:AdditionalProperties></m:ItemShape><m:ItemIds><t:ItemId Id="' + (ConvertTo-WscXmlText $ItemId) + '"/></m:ItemIds></m:GetItem>'
}

function New-WscSendMessageBody {
    <# CreateItem SendAndSaveCopy. Delegate access: saved in the Sent Items of the mailbox and sent From it (Send As or Send on Behalf). #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Subject, [Parameter(Mandatory = $true)][string]$BodyText, [Parameter(Mandatory = $true)][string[]]$To)

    $recipients = ($To | ForEach-Object { '<t:Mailbox><t:EmailAddress>{0}</t:EmailAddress></t:Mailbox>' -f (ConvertTo-WscXmlText $_) }) -join ''
    $from = if ($Context.Access -eq 'Delegate') { '<t:From><t:Mailbox><t:EmailAddress>{0}</t:EmailAddress></t:Mailbox></t:From>' -f (ConvertTo-WscXmlText ([string]$Context.Config.Mailbox)) } else { '' }
    '<m:CreateItem MessageDisposition="SendAndSaveCopy"><m:SavedItemFolderId>' + (Get-WscFolderIdXml -Context $Context -Distinguished 'sentitems') + '</m:SavedItemFolderId><m:Items><t:Message>' +
    '<t:Subject>' + (ConvertTo-WscXmlText $Subject) + '</t:Subject><t:Body BodyType="Text">' + (ConvertTo-WscXmlText $BodyText) + '</t:Body><t:ToRecipients>' + $recipients + '</t:ToRecipients>' + $from +
    '</t:Message></m:Items></m:CreateItem>'
}

function New-WscReplyBody {
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$ItemId, [string]$ChangeKey, [Parameter(Mandatory = $true)][string]$BodyText)
    $ck = if ($ChangeKey) { ' ChangeKey="{0}"' -f (ConvertTo-WscXmlText $ChangeKey) } else { '' }
    '<m:CreateItem MessageDisposition="SendAndSaveCopy"><m:SavedItemFolderId>' + (Get-WscFolderIdXml -Context $Context -Distinguished 'sentitems') + '</m:SavedItemFolderId><m:Items><t:ReplyToItem>' +
    '<t:ReferenceItemId Id="' + (ConvertTo-WscXmlText $ItemId) + '"' + $ck + '/><t:NewBodyContent BodyType="Text">' + (ConvertTo-WscXmlText $BodyText) + '</t:NewBodyContent></t:ReplyToItem></m:Items></m:CreateItem>'
}

function New-WscMoveItemBody {
    param([Parameter(Mandatory = $true)][string]$ItemId, [Parameter(Mandatory = $true)][string]$FolderId)
    '<m:MoveItem><m:ToFolderId><t:FolderId Id="' + (ConvertTo-WscXmlText $FolderId) + '"/></m:ToFolderId><m:ItemIds><t:ItemId Id="' + (ConvertTo-WscXmlText $ItemId) + '"/></m:ItemIds></m:MoveItem>'
}

function New-WscDeleteItemBody {
    param([Parameter(Mandatory = $true)][string]$ItemId, [string]$DeleteType = 'MoveToDeletedItems')
    '<m:DeleteItem DeleteType="' + $DeleteType + '"><m:ItemIds><t:ItemId Id="' + (ConvertTo-WscXmlText $ItemId) + '"/></m:ItemIds></m:DeleteItem>'
}

function New-WscAvailabilityBody {
    <# GetUserAvailability for several mailboxes, times in UTC (time zone with a bias of 0), DetailedMerged view. #>
    param([Parameter(Mandatory = $true)][string[]]$Mailboxes, [Parameter(Mandatory = $true)][datetime]$StartUtc, [Parameter(Mandatory = $true)][datetime]$EndUtc, [int]$IntervalMinutes = 30)

    $zone = '<t:TimeZone><t:Bias>0</t:Bias><t:StandardTime><t:Bias>0</t:Bias><t:Time>00:00:00</t:Time><t:DayOrder>0</t:DayOrder><t:Month>0</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek></t:StandardTime>' +
        '<t:DaylightTime><t:Bias>0</t:Bias><t:Time>00:00:00</t:Time><t:DayOrder>0</t:DayOrder><t:Month>0</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek></t:DaylightTime></t:TimeZone>'
    $data = ($Mailboxes | ForEach-Object { '<t:MailboxData><t:Email><t:Address>{0}</t:Address><t:RoutingType>SMTP</t:RoutingType></t:Email><t:AttendeeType>Required</t:AttendeeType><t:ExcludeConflicts>false</t:ExcludeConflicts></t:MailboxData>' -f (ConvertTo-WscXmlText $_) }) -join ''
    $f = 'yyyy-MM-ddTHH:mm:ss'
    '<m:GetUserAvailabilityRequest>' + $zone + '<m:MailboxDataArray>' + $data + '</m:MailboxDataArray><t:FreeBusyViewOptions><t:TimeWindow>' +
    ('<t:StartTime>{0}</t:StartTime><t:EndTime>{1}</t:EndTime></t:TimeWindow><t:MergedFreeBusyIntervalInMinutes>{2}</t:MergedFreeBusyIntervalInMinutes><t:RequestedView>DetailedMerged</t:RequestedView>' -f $StartUtc.ToString($f), $EndUtc.ToString($f), $IntervalMinutes) +
    '</t:FreeBusyViewOptions></m:GetUserAvailabilityRequest>'
}

#endregion

#region Operations: answers read ------------------------------------------------------------------

function Get-WscNodeText {
    param([AllowNull()][Xml.XmlNode]$Node, [Parameter(Mandatory = $true)][string]$XPath, [Parameter(Mandatory = $true)][Xml.XmlNamespaceManager]$Ns)
    if (-not $Node) { return $null }
    $n = $Node.SelectSingleNode($XPath, $Ns)
    if ($n) { return $n.InnerText }
    return $null
}

function ConvertFrom-WscFolderNode {
    param([Parameter(Mandatory = $true)][Xml.XmlNode]$Node, [Parameter(Mandatory = $true)][Xml.XmlNamespaceManager]$Ns)
    $id = $Node.SelectSingleNode('t:FolderId', $Ns)
    $parent = $Node.SelectSingleNode('t:ParentFolderId', $Ns)
    [pscustomobject]@{
        DisplayName      = Get-WscNodeText $Node 't:DisplayName' $Ns
        Path             = $null
        FolderClass      = Get-WscNodeText $Node 't:FolderClass' $Ns
        Kind             = $Node.LocalName
        TotalCount       = Get-WscNodeText $Node 't:TotalCount' $Ns
        UnreadCount      = Get-WscNodeText $Node 't:UnreadCount' $Ns
        ChildFolderCount = Get-WscNodeText $Node 't:ChildFolderCount' $Ns
        FolderId         = if ($id) { $id.GetAttribute('Id') } else { $null }
        ChangeKey        = if ($id) { $id.GetAttribute('ChangeKey') } else { $null }
        ParentFolderId   = if ($parent) { $parent.GetAttribute('Id') } else { $null }
    }
}

function ConvertFrom-WscItemNode {
    param([Parameter(Mandatory = $true)][Xml.XmlNode]$Node, [Parameter(Mandatory = $true)][Xml.XmlNamespaceManager]$Ns)
    $id = $Node.SelectSingleNode('t:ItemId', $Ns)
    $mailbox = { param([string]$Path) $n = $Node.SelectSingleNode($Path, $Ns); if (-not $n) { return $null }; $name = Get-WscNodeText $n 't:Name' $Ns; $addr = Get-WscNodeText $n 't:EmailAddress' $Ns; if ($name -and $addr -and $name -ne $addr) { "$name <$addr>" } elseif ($addr) { $addr } else { $name } }
    $list = { param([string]$Path) @($Node.SelectNodes("$Path/t:Mailbox", $Ns) | ForEach-Object { Get-WscNodeText $_ 't:EmailAddress' $Ns }) -join ', ' }
    [pscustomobject]@{
        DateTimeReceived = Get-WscNodeText $Node 't:DateTimeReceived' $Ns
        From             = & $mailbox 't:From/t:Mailbox'
        Subject          = Get-WscNodeText $Node 't:Subject' $Ns
        IsRead           = Get-WscNodeText $Node 't:IsRead' $Ns
        HasAttachments   = Get-WscNodeText $Node 't:HasAttachments' $Ns
        Size             = Get-WscNodeText $Node 't:Size' $Ns
        ItemClass        = Get-WscNodeText $Node 't:ItemClass' $Ns
        To               = & $list 't:ToRecipients'
        Cc               = & $list 't:CcRecipients'
        Sender           = & $mailbox 't:Sender/t:Mailbox'
        InternetMessageId = Get-WscNodeText $Node 't:InternetMessageId' $Ns
        Body             = Get-WscNodeText $Node 't:Body' $Ns
        ItemId           = if ($id) { $id.GetAttribute('Id') } else { $null }
        ChangeKey        = if ($id) { $id.GetAttribute('ChangeKey') } else { $null }
    }
}

function ConvertFrom-WscAvailability {
    <# One object per mailbox of a GetUserAvailability answer: response code, working hours, busy periods, merged view. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Answer, [Parameter(Mandatory = $true)][string[]]$Mailboxes, [int]$IntervalMinutes = 30, [datetime]$StartUtc)

    $ns = $Answer.Ns
    $responses = @($Answer.Xml.SelectNodes('//m:FreeBusyResponseArray/m:FreeBusyResponse', $ns))
    for ($i = 0; $i -lt $Mailboxes.Count; $i++) {
        $r = if ($i -lt $responses.Count) { $responses[$i] } else { $null }
        $message = if ($r) { $r.SelectSingleNode('m:ResponseMessage', $ns) } else { $null }
        $view = if ($r) { $r.SelectSingleNode('m:FreeBusyView', $ns) } else { $null }
        $events = @(if ($view) {
                foreach ($e in $view.SelectNodes('t:CalendarEventArray/t:CalendarEvent', $ns)) {
                    $start = Get-WscNodeText $e 't:StartTime' $ns
                    $end = Get-WscNodeText $e 't:EndTime' $ns
                    [pscustomobject]@{
                        Mailbox  = $Mailboxes[$i]
                        StartUtc = $start
                        EndUtc   = $end
                        Start    = if ($start) { ([datetime]::SpecifyKind([datetime]$start, 'Utc')).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { $null }
                        End      = if ($end) { ([datetime]::SpecifyKind([datetime]$end, 'Utc')).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { $null }
                        BusyType = Get-WscNodeText $e 't:BusyType' $ns
                        Subject  = Get-WscNodeText $e 't:CalendarEventDetails/t:Subject' $ns
                        Location = Get-WscNodeText $e 't:CalendarEventDetails/t:Location' $ns
                        IsMeeting = Get-WscNodeText $e 't:CalendarEventDetails/t:IsMeeting' $ns
                        IsPrivate = Get-WscNodeText $e 't:CalendarEventDetails/t:IsPrivate' $ns
                    }
                }
            })
        $firstPeriod = if ($view) { $view.SelectSingleNode('t:WorkingHours/t:WorkingPeriodArray/t:WorkingPeriod', $ns) } else { $null }
        $periods = @(if ($view) {
                foreach ($p in $view.SelectNodes('t:WorkingHours/t:WorkingPeriodArray/t:WorkingPeriod', $ns)) {
                    $from = [int](Get-WscNodeText $p 't:StartTimeInMinutes' $ns); $to = [int](Get-WscNodeText $p 't:EndTimeInMinutes' $ns)
                    '{0} {1:00}:{2:00}-{3:00}:{4:00}' -f (Get-WscNodeText $p 't:DayOfWeek' $ns), [Math]::Floor($from / 60), ($from % 60), [Math]::Floor($to / 60), ($to % 60)
                }
            })
        [pscustomobject]@{
            Mailbox       = $Mailboxes[$i]
            ResponseClass = if ($message) { $message.GetAttribute('ResponseClass') } else { 'Error' }
            ResponseCode  = if ($message) { Get-WscNodeText $message 'm:ResponseCode' $ns } else { 'NoResponse' }
            MessageText   = if ($message) { Get-WscNodeText $message 'm:MessageText' $ns } else { 'No FreeBusyResponse for this mailbox.' }
            ViewType      = Get-WscNodeText $view 't:FreeBusyViewType' $ns
            MergedFreeBusy = Get-WscNodeText $view 't:MergedFreeBusy' $ns
            WorkingHours  = $periods -join '; '
            WorkingTimeZone = Get-WscNodeText $view 't:WorkingHours/t:TimeZone/t:Bias' $ns
            WorkingDays   = if ($firstPeriod) { @(([string](Get-WscNodeText $firstPeriod 't:DayOfWeek' $ns)).ToLowerInvariant().Split(' ', [StringSplitOptions]::RemoveEmptyEntries)) } else { @() }
            WorkStart     = if ($firstPeriod) { [int](Get-WscNodeText $firstPeriod 't:StartTimeInMinutes' $ns) } else { $null }
            WorkEnd       = if ($firstPeriod) { [int](Get-WscNodeText $firstPeriod 't:EndTimeInMinutes' $ns) } else { $null }
            WorkOffsetMinutes = if ($view) { Get-WscEwsZoneOffset -Zone $view.SelectSingleNode('t:WorkingHours/t:TimeZone', $ns) -Ns $ns -AtUtc $(if ($StartUtc) { $StartUtc } else { [datetime]::UtcNow }) } else { $null }
            Events        = $events
        }
    }
}

#endregion

function Get-WscEwsZoneOffset {
    <#
        UTC offset (minutes) of an EWS SerializableTimeZone (Bias, StandardTime, DaylightTime) at a UTC instant:
        local = UTC - (Bias + StandardTime/DaylightTime Bias). The transitions are relative rules (DayOrder 1-4 or
        5 = last DayOfWeek of Month, at Time). $null when the zone is missing.
    #>
    param([AllowNull()][Xml.XmlNode]$Zone, [Parameter(Mandatory = $true)][Xml.XmlNamespaceManager]$Ns, [Parameter(Mandatory = $true)][datetime]$AtUtc)

    if (-not $Zone) { return $null }
    $bias = [int](Get-WscNodeText $Zone 't:Bias' $Ns)
    $std = $Zone.SelectSingleNode('t:StandardTime', $Ns); $dst = $Zone.SelectSingleNode('t:DaylightTime', $Ns)
    $stdBias = if ($std) { [int](Get-WscNodeText $std 't:Bias' $Ns) } else { 0 }
    $dstBias = if ($dst) { [int](Get-WscNodeText $dst 't:Bias' $Ns) } else { 0 }
    $rule = {
        param([Xml.XmlNode]$N, [int]$Year)
        $month = [int](Get-WscNodeText $N 't:Month' $Ns); if ($month -lt 1) { return $null }
        $order = [int](Get-WscNodeText $N 't:DayOrder' $Ns); $dow = [DayOfWeek](Get-WscNodeText $N 't:DayOfWeek' $Ns)
        $time = [TimeSpan]::Parse((Get-WscNodeText $N 't:Time' $Ns))
        $d = [datetime]::new($Year, $month, 1); while ($d.DayOfWeek -ne $dow) { $d = $d.AddDays(1) }
        $d = $d.AddDays(7 * ([Math]::Min($order, 5) - 1)); while ($d.Month -ne $month) { $d = $d.AddDays(-7) }
        $d.Add($time)
    }
    $inDst = $false
    if ($std -and $dst -and $dstBias -ne $stdBias) {
        $y = $AtUtc.Year
        $toDst = & $rule $dst $y; $toStd = & $rule $std $y
        if ($toDst -and $toStd) {
            # Transition times are local wall-clock times of the zone before the change.
            $startUtc = $toDst.AddMinutes($bias + $stdBias); $endUtc = $toStd.AddMinutes($bias + $dstBias)
            $inDst = if ($startUtc -lt $endUtc) { $AtUtc -ge $startUtc -and $AtUtc -lt $endUtc } else { $AtUtc -ge $startUtc -or $AtUtc -lt $endUtc }
        }
    }
    return -($bias + $(if ($inDst) { $dstBias } else { $stdBias }))
}