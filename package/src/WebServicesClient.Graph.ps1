<#
.SYNOPSIS
    Web Services Client for Exchange - Microsoft Graph: requests and the operations of the toolbox for Exchange Online (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    EWS is being retired in Exchange Online (disabled from October 2026, stopped in April 2027): with
    Target.Protocol Auto, a mailbox in Exchange Online is tested through Microsoft Graph, an on-premises
    mailbox through EWS (Graph does not reach on-premises mailboxes). The scenarios and the report are
    the same; only the requests change:

      Endpoint      GET  /users/{mailbox}/mailFolders/inbox
      Folders       GET  /users/{mailbox}/mailFolders (+ childFolders, hidden folders, pages)
      ReadMail      GET  /users/{mailbox}/mailFolders/inbox/messages, then GET /messages/{id} (text body)
      FreeBusy      POST /users/{mailbox}/calendar/getSchedule
      CreateFolder  GET or POST /users/{mailbox}/mailFolders/inbox/childFolders
      SendMail      POST /users/{mailbox}/sendMail, then the message looked for in the Inbox
      ReplyMail     POST /users/{mailbox}/messages/{id}/reply
      MoveMail      POST /users/{mailbox}/messages/{id}/move
      DeleteMail    DELETE /messages/{id}, move to recoverableitemsdeletions, or permanentDelete

    Good practices of Graph: client-request-id and return-client-request-id, $select of the properties
    needed only, paging with @odata.nextLink, Prefer headers (text body, UTC), throttling (HTTP 429, 503,
    504 with Retry-After, Ews.MaxRetries attempts). The server that answered comes from x-ms-ags-diagnostic.
    Graph accepts only OAuth with Entra ID; the mailbox is always /users/{address} (no impersonation:
    an application token reaches the mailboxes its permissions and RBAC for Applications allow).

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:GraphRoot = 'https://graph.microsoft.com/v1.0'
$script:GraphAppId = '00000003-0000-0000-c000-000000000000'

function ConvertTo-WscODataText {
    <# A text inside an OData string literal ('' for '), URL-encoded. #>
    param([AllowEmptyString()][string]$Text)
    return [Uri]::EscapeDataString($Text.Replace("'", "''"))
}

function Get-WscGraphMailboxPath {
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    return '/users/' + [Uri]::EscapeDataString([string]$Context.Config.Mailbox)
}

function Invoke-WscGraph {
    <#
        One Microsoft Graph request with the token of the run (or a given token, or none), traced.
        Throttled requests are sent again after Retry-After. Never throws for an HTTP error: the caller
        reads HttpStatus, ErrorCode and ErrorMessage.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Context,
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Method = 'GET',
        [object]$Body,
        [hashtable]$Headers,
        [string]$Operation,
        [string]$Collapse,
        [ValidateSet('Context', 'None', 'Token')][string]$Credentials = 'Context',
        [string]$Token
    )

    Invoke-WscUiPump
    Assert-WscNotCancelled
    $uri = if ($Path -match '^https://') { $Path } else { "$($script:GraphRoot)$Path" }
    if (-not $Operation) { $Operation = "$Method " + (([Uri]$uri).AbsolutePath -replace '^/v1\.0/users/[^/]+/?', '' -replace '/[A-Za-z0-9_=-]{40,}', '/{id}') }
    $json = if ($null -ne $Body) { ConvertTo-Json -InputObject $Body -Depth 8 -Compress } else { $null }
    $retries = [int]$Context.Config.MaxRetries
    $attempt = 0
    while ($true) {
        $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::new($Method), $uri)
        try {
            $add = { param([string]$Name, [string]$Value) $null = $request.Headers.TryAddWithoutValidation($Name, $Value) }
            & $add 'Accept' 'application/json'
            & $add 'User-Agent' ([string]$Context.Config.UserAgent)
            & $add 'client-request-id' ([guid]::NewGuid().ToString())
            & $add 'return-client-request-id' 'true'
            if ($Headers) { foreach ($k in $Headers.Keys) { & $add ([string]$k) ([string]$Headers[$k]) } }
            switch ($Credentials) {
                'Context' { if ($Context.AccessToken) { $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $Context.AccessToken) } }
                'Token' { $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $Token) }
            }
            if ($null -ne $json) {
                $request.Content = [Net.Http.StringContent]::new($json, [Text.Encoding]::UTF8, 'application/json')
            }
            $response = Invoke-WscHttp -HttpClient $Context.HttpClient -Request $request -Operation $Operation -Collapse $Collapse
        }
        finally { $request.Dispose() }
        if ($response.StatusCode -notin 429, 503, 504 -or $attempt -ge $retries) { break }
        $after = 0
        $wait = if ([int]::TryParse([string](Get-WscHeader $response 'Retry-After'), [ref]$after) -and $after -gt 0) { $after } else { 5 }
        if ($wait -gt $script:EwsMaxWaitSeconds) { Write-WscItem Warn "Microsoft Graph throttles the requests and asks to wait $wait s (more than $($script:EwsMaxWaitSeconds) s): the request is not sent again."; break }
        $attempt++
        $Context.Throttled++
        Write-WscItem Warn "Microsoft Graph throttles the requests (HTTP $($response.StatusCode)): new attempt $attempt/$retries in $wait s, as asked."
        Wait-WscSeconds $wait
    }
    $data = $null
    $text = if ($response.Body -and $response.Body.Length) { [Text.Encoding]::UTF8.GetString([byte[]]$response.Body) } else { '' }
    if ($text.TrimStart().StartsWith('{')) { try { $data = $text | ConvertFrom-Json -AsHashtable -Depth 32 -ErrorAction Stop } catch { $data = $null } }
    $failure = if ($data -and $data['error']) { $data['error'] } else { $null }
    $server = Get-WscResponseServer -Response $response
    if ($server) { $Context.LastServer = $server }
    [pscustomobject]@{
        Operation    = $Operation
        Response     = $response
        HttpStatus   = [int]$response.StatusCode
        Json         = $data
        ErrorCode    = if ($failure) { [string]$failure['code'] } else { $null }
        ErrorMessage = if ($failure) { [string]$failure['message'] } else { $null }
        Server       = $server
        RequestId    = Get-WscHeader $response 'request-id'
        Retries      = $attempt
    }
}

function Get-WscGraphDetails {
    <# Details common to every Graph check: HTTP status, error, server, request ID. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Answer, [System.Collections.IDictionary]$More)
    $d = [ordered]@{ Operation = $Answer.Operation; HttpStatus = $Answer.HttpStatus }
    if ($Answer.ErrorCode) { $d.ErrorCode = $Answer.ErrorCode; $d.ErrorMessage = $Answer.ErrorMessage }
    if ($Answer.Server) { $d.Servers = $Answer.Server }
    if ($Answer.RequestId) { $d.RequestId = $Answer.RequestId }
    if ($Answer.Retries) { $d.Retries = $Answer.Retries }
    if ($More) { foreach ($k in $More.Keys) { $d[$k] = $More[$k] } }
    return $d
}

function Test-WscGraphOk {
    param([Parameter(Mandatory = $true)][pscustomobject]$Answer)
    return $Answer.HttpStatus -ge 200 -and $Answer.HttpStatus -lt 300
}

function Get-WscGraphFailureText {
    <# Why a Graph request failed, with what to check. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][pscustomobject]$Answer)

    $app = [string]$Context.Config.Context -eq 'Application'
    $what = if ($Answer.ErrorCode) { "$($Answer.ErrorCode): $($Answer.ErrorMessage)" } else { 'no error body' }
    $id = if ($Answer.RequestId) { " Request ID $($Answer.RequestId)." } else { '' }
    $hint = switch ($Answer.HttpStatus) {
        401 { 'Graph refused the token: audience https://graph.microsoft.com, expiry, tenant.' }
        403 {
            if ($app) { 'the application permission (Mail.Read / Mail.ReadWrite, Mail.Send, Calendars.Read) with admin consent, and the scope of RBAC for Applications or of the application access policy that limits the mailboxes.' }
            elseif ($Context.Access -eq 'Delegate') { 'the delegated permission of the token (Mail.Read.Shared, Mail.ReadWrite.Shared, Mail.Send.Shared) and the permission of the signed-in account on the mailbox (Full Access, folder permissions, Send As).' }
            else { 'the delegated permission of the token (scp: Mail.Read, Mail.ReadWrite, Mail.Send, Calendars.Read), consent, or a Conditional Access policy.' }
        }
        404 {
            switch -Regex ([string]$Answer.ErrorCode) {
                'MailboxNotEnabledForRESTAPI|MailboxNotSupportedForRESTAPI' { 'the mailbox is not in Exchange Online (on-premises, or not licensed): Graph cannot reach it. Use -Protocol EWS.' }
                'ErrorItemNotFound' { 'the item is not (or no longer) in the mailbox.' }
                'ErrorInvalidUser|ResourceNotFound|Request_ResourceNotFound' { 'no user or mailbox with this address in the tenant of the token.' }
                default { 'not found.' }
            }
        }
        429 { 'Microsoft Graph throttles the requests: wait and try again.' }
        default { $null }
    }
    return "HTTP $($Answer.HttpStatus) ($what)$(if ($hint) { ": $hint" } else { '.' })$id"
}

function ConvertFrom-WscGraphAddress {
    param([AllowNull()][object]$Recipient)
    if (-not $Recipient -or -not $Recipient['emailAddress']) { return $null }
    $name = [string]$Recipient['emailAddress']['name']; $address = [string]$Recipient['emailAddress']['address']
    if ($name -and $address -and $name -ne $address) { return "$name <$address>" }
    if ($address) { return $address }
    return $name
}

function ConvertFrom-WscGraphMessage {
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Message)
    [pscustomobject]@{
        DateTimeReceived  = [string]$Message['receivedDateTime']
        From              = ConvertFrom-WscGraphAddress $Message['from']
        Subject           = [string]$Message['subject']
        IsRead            = [string]$Message['isRead']
        HasAttachments    = [string]$Message['hasAttachments']
        Size              = $null
        ItemClass         = $null
        To                = (@($Message['toRecipients']) | Where-Object { $_ } | ForEach-Object { $_['emailAddress']['address'] }) -join ', '
        Cc                = (@($Message['ccRecipients']) | Where-Object { $_ } | ForEach-Object { $_['emailAddress']['address'] }) -join ', '
        Sender            = ConvertFrom-WscGraphAddress $Message['sender']
        InternetMessageId = [string]$Message['internetMessageId']
        Body              = if ($Message['body']) { [string]$Message['body']['content'] } else { $null }
        ItemId            = [string]$Message['id']
        ChangeKey         = [string]$Message['changeKey']
    }
}

function Get-WscGraphInboxMessages {
    <# Inbox messages: the most recent (Top), those with an exact subject (Filter), or those found by a search (Search, sorted here). #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [int]$Top = 10, [string]$Subject, [string]$Search, [string]$Collapse)

    $select = 'id,subject,from,receivedDateTime,isRead,hasAttachments,internetMessageId,changeKey'
    $base = "$(Get-WscGraphMailboxPath $Context)/mailFolders/inbox/messages?`$select=$select&`$top=$Top"
    $path = if ($Subject) { "$base&`$filter=subject eq '$(ConvertTo-WscODataText $Subject)'" }
    elseif ($Search) { "$base&`$search=" + [Uri]::EscapeDataString("`"$Search`"") }
    else { "$base&`$orderby=receivedDateTime%20desc" }
    $answer = Invoke-WscGraph -Context $Context -Path $path -Collapse $Collapse -Operation $(if ($Subject -or $Search) { 'GET inbox messages (search)' } else { 'GET inbox messages' })
    $items = @()
    if (Test-WscGraphOk $answer) {
        $items = @($answer.Json['value'] | ForEach-Object { ConvertFrom-WscGraphMessage $_ } | Sort-Object DateTimeReceived -Descending)
    }
    [pscustomobject]@{ Answer = $answer; Items = $items }
}

#region Stages --------------------------------------------------------------------------------------

function Invoke-WscGraphOperation {
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage)
    switch ($Stage) {
        'Endpoint' { Invoke-WscGraphEndpoint -Context $Context }
        'Folders' { Invoke-WscGraphFolders -Context $Context }
        'ReadMail' { Invoke-WscGraphReadMail -Context $Context }
        'FreeBusy' { Invoke-WscGraphFreeBusy -Context $Context }
        'CreateFolder' { [void](Get-WscGraphTestFolder -Context $Context -Stage 'CreateFolder' -Create) }
        'SendMail' { Invoke-WscGraphSendMail -Context $Context }
        'ReplyMail' { Invoke-WscGraphReplyMail -Context $Context }
        'MoveMail' { Invoke-WscGraphMoveMail -Context $Context }
        'DeleteMail' { Invoke-WscGraphDeleteMail -Context $Context }
    }
}

function Invoke-WscGraphEndpoint {
    <# First Graph request on the mailbox: its Inbox. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Endpoint'
    $answer = Invoke-WscGraph -Context $Context -Path "$(Get-WscGraphMailboxPath $Context)/mailFolders/inbox?`$select=id,displayName,totalItemCount,unreadItemCount,childFolderCount" -Operation 'GET mailFolders/inbox'
    $access = Get-WscAccessText -Context $Context
    $details = Get-WscGraphDetails -Answer $answer -More ([ordered]@{ Access = $Context.Access; Mailbox = "/users/$($Context.Config.Mailbox)" })
    if (-not (Test-WscGraphOk $answer)) {
        Add-WscStep $Context $stage 'GET mailFolders/inbox' Failed "Microsoft Graph did not open the Inbox ($access): $(Get-WscGraphFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $true
        return
    }
    $j = $answer.Json
    $inbox = [pscustomobject]@{ DisplayName = $j['displayName']; TotalCount = $j['totalItemCount']; UnreadCount = $j['unreadItemCount']; ChildFolderCount = $j['childFolderCount']; FolderId = $j['id'] }
    $Context.Inbox = $inbox
    $Context.Mailbox = [pscustomobject]@{ Address = [string]$Context.Config.Mailbox; Access = $Context.Access; Inbox = $inbox }
    $Context.ServerVersion = [pscustomobject]@{ Text = 'Microsoft Graph v1.0'; Schema = 'v1.0'; Major = $null }
    $details.Inbox = "$($inbox.DisplayName): $($inbox.TotalCount) item(s), $($inbox.UnreadCount) unread"
    Add-WscStep $Context $stage 'GET mailFolders/inbox' Passed "Microsoft Graph accepted the token and opened the Inbox through $($access): $($details.Inbox).$(if ($answer.Server) { " $($answer.Server)." })" $details
    $echo = Get-WscHeader $answer.Response 'client-request-id'
    $ids = [ordered]@{ RequestId = $answer.RequestId; ClientRequestId = $echo; Servers = $answer.Server }
    Add-WscStep $Context $stage 'Request IDs' Passed ("Graph routes the request to the mailbox itself (no X-AnchorMailbox, no affinity cookie). $(if ($echo) { 'It echoes client-request-id (return-client-request-id)' } else { 'It did not echo client-request-id' }); " +
        "request-id $($answer.RequestId) identifies the request for Microsoft support.") $ids
}

function Invoke-WscGraphFolders {
    <# The folder tree: mailFolders (hidden folders included), then the childFolders of every folder that has some, page by page. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Folders'
    $mb = Get-WscGraphMailboxPath $Context
    $select = 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
    $queue = [Collections.Generic.Queue[string]]::new()
    $queue.Enqueue("$mb/mailFolders?includeHiddenFolders=true&`$top=100&`$select=$select")
    $all = [Collections.Generic.List[object]]::new()
    $requests = 0
    $last = $null
    while ($queue.Count -and $requests -lt 200) {
        $requests++
        $last = Invoke-WscGraph -Context $Context -Path $queue.Dequeue() -Operation $(if ($requests -eq 1) { 'GET mailFolders' } else { 'GET childFolders' })
        if (-not (Test-WscGraphOk $last)) {
            Add-WscStep $Context $stage 'GET mailFolders' Failed "The folders were not listed: $(Get-WscGraphFailureText -Context $Context -Answer $last)" (Get-WscGraphDetails -Answer $last)
            $Context.Stop = $true
            return
        }
        foreach ($f in @($last.Json['value'])) {
            $all.Add([pscustomobject]@{
                    DisplayName = [string]$f['displayName']; Path = $null; FolderClass = $null; Kind = 'mailFolder'; TotalCount = $f['totalItemCount']; UnreadCount = $f['unreadItemCount']
                    ChildFolderCount = $f['childFolderCount']; FolderId = [string]$f['id']; ChangeKey = $null; ParentFolderId = [string]$f['parentFolderId']
                })
            if ([int]$f['childFolderCount'] -gt 0) { $queue.Enqueue("$mb/mailFolders/$([Uri]::EscapeDataString([string]$f['id']))/childFolders?includeHiddenFolders=true&`$top=100&`$select=$select") }
        }
        if ($last.Json['@odata.nextLink']) { $queue.Enqueue([string]$last.Json['@odata.nextLink']) }
    }
    $Context.Folders = Set-WscFolderPaths -Folders @($all)
    $items = ($all | ForEach-Object { [int]$_.TotalCount } | Measure-Object -Sum).Sum
    $details = Get-WscGraphDetails -Answer $last -More ([ordered]@{ Folders = $all.Count; Items = $items; Requests = $requests })
    $cut = if ($queue.Count) { " The listing stopped after $requests requests ($($queue.Count) left)." } else { '' }
    Add-WscStep $Context $stage 'GET mailFolders' $(if ($cut) { 'Warning' } else { 'Passed' }) "$($all.Count) mail folder(s) ($items item(s) in total) in $requests request(s): Folders tab of the report.$cut" $details
}

function Invoke-WscGraphReadMail {
    <# The most recent Inbox messages, then one message in full with a text body. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'ReadMail'
    $found = Get-WscGraphInboxMessages -Context $Context -Top ([int]$Context.Config.MessageCount)
    $answer = $found.Answer
    if (-not (Test-WscGraphOk $answer)) {
        Add-WscStep $Context $stage 'GET inbox messages' Failed "The Inbox messages were not listed: $(Get-WscGraphFailureText -Context $Context -Answer $answer)" (Get-WscGraphDetails -Answer $answer)
        $Context.Stop = $true
        return
    }
    $Context.Messages = @($found.Items | Select-Object DateTimeReceived, From, Subject, IsRead, HasAttachments, Size, ItemClass, InternetMessageId, ItemId)
    $total = if ($Context.Inbox) { $Context.Inbox.TotalCount } else { $null }
    $details = Get-WscGraphDetails -Answer $answer -More ([ordered]@{ Listed = $found.Items.Count; InInbox = $total })
    if (-not $found.Items.Count) { Add-WscStep $Context $stage 'GET inbox messages' Passed 'The Inbox is empty: nothing to read.' $details; return }
    Add-WscStep $Context $stage 'GET inbox messages' Passed "$($found.Items.Count) most recent message(s)$(if ($null -ne $total) { " of $total" }) in the Inbox: date, sender, subject (Messages tab)." $details

    $pick = $found.Items[0]
    $get = Invoke-WscGraph -Context $Context -Path "$(Get-WscGraphMailboxPath $Context)/messages/$([Uri]::EscapeDataString($pick.ItemId))?`$select=subject,from,sender,toRecipients,ccRecipients,receivedDateTime,internetMessageId,body" `
        -Headers @{ Prefer = 'outlook.body-content-type="text"' } -Operation 'GET message'
    $d = Get-WscGraphDetails -Answer $get
    if (-not (Test-WscGraphOk $get)) {
        Add-WscStep $Context $stage 'GET message' Failed "The message '$($pick.Subject)' was not read: $(Get-WscGraphFailureText -Context $Context -Answer $get)" $d
        return
    }
    $item = ConvertFrom-WscGraphMessage $get.Json
    $body = ([string]$item.Body -replace '\s+', ' ').Trim()
    $preview = if ($body.Length -gt 300) { $body.Substring(0, 300) + '...' } else { $body }
    $Context.Message = [pscustomobject]@{ Subject = $item.Subject; From = $item.From; Sender = $item.Sender; To = $item.To; Cc = $item.Cc; DateTimeReceived = $item.DateTimeReceived; InternetMessageId = $item.InternetMessageId; Size = $null; BodyLength = ([string]$item.Body).Length; BodyPreview = $preview; ItemId = $item.ItemId }
    foreach ($k in 'Subject', 'From', 'To', 'DateTimeReceived', 'InternetMessageId') { $d[$k] = $Context.Message.$k }
    $d.BodyLength = $Context.Message.BodyLength
    Add-WscStep $Context $stage 'GET message' Passed "Message read in full: '$($item.Subject)' from $($item.From), text body of $($Context.Message.BodyLength) characters (Prefer: outlook.body-content-type)." $d
}

function Invoke-WscGraphFreeBusy {
    <# getSchedule for the mailboxes of Test.FreeBusyMailboxes (else the mailbox), in UTC. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'FreeBusy'
    $cfg = $Context.Config
    $mailboxes = @($cfg.FreeBusyMailboxes | Where-Object { $_ })
    if (-not $mailboxes.Count) { $mailboxes = @([string]$cfg.Mailbox) }
    $start = (Get-Date).Date.ToUniversalTime()
    $end = $start.AddDays([int]$cfg.FreeBusyDays)
    $body = [ordered]@{
        schedules = $mailboxes; startTime = @{ dateTime = $start.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }
        endTime = @{ dateTime = $end.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }; availabilityViewInterval = [int]$cfg.FreeBusyIntervalMinutes
    }
    $answer = Invoke-WscGraph -Context $Context -Method POST -Path "$(Get-WscGraphMailboxPath $Context)/calendar/getSchedule" -Body $body -Headers @{ Prefer = 'outlook.timezone="UTC"' } -Operation 'POST calendar/getSchedule'
    $details = Get-WscGraphDetails -Answer $answer -More ([ordered]@{ Mailboxes = $mailboxes -join ', '; WindowUtc = '{0:yyyy-MM-dd HH:mm} to {1:yyyy-MM-dd HH:mm}' -f $start, $end; IntervalMinutes = $cfg.FreeBusyIntervalMinutes })
    if (-not (Test-WscGraphOk $answer)) {
        Add-WscStep $Context $stage 'POST getSchedule' Failed "Free/busy not read: $(Get-WscGraphFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $true
        return
    }
    $values = @($answer.Json['value'])
    $results = for ($i = 0; $i -lt $mailboxes.Count; $i++) {
        $v = $values | Where-Object { [string]$_['scheduleId'] -ieq $mailboxes[$i] } | Select-Object -First 1
        if (-not $v -and $i -lt $values.Count) { $v = $values[$i] }
        $err = if ($v) { $v['error'] } else { @{ responseCode = 'NoResponse'; message = 'No answer for this mailbox.' } }
        $wh = if ($v) { $v['workingHours'] } else { $null }
        $events = @(if ($v -and -not $err) {
                foreach ($e in @($v['scheduleItems'])) {
                    if (-not $e) { continue }
                    $s = [string]$e['start']['dateTime']; $f = [string]$e['end']['dateTime']
                    [pscustomobject]@{
                        Mailbox = $mailboxes[$i]; StartUtc = $s; EndUtc = $f
                        Start = if ($s) { ([datetime]::SpecifyKind([datetime]$s, 'Utc')).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { $null }
                        End = if ($f) { ([datetime]::SpecifyKind([datetime]$f, 'Utc')).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { $null }
                        BusyType = [string]$e['status']; Subject = [string]$e['subject']; Location = [string]$e['location']; IsMeeting = $null; IsPrivate = [string]$e['isPrivate']
                    }
                }
            })
        [pscustomobject]@{
            Mailbox        = $mailboxes[$i]
            ResponseClass  = if ($err) { 'Error' } else { 'Success' }
            ResponseCode   = if ($err) { [string]$err['responseCode'] } else { 'NoError' }
            MessageText    = if ($err) { [string]$err['message'] } else { $null }
            ViewType       = if ($err) { $null } else { 'availabilityView' }
            MergedFreeBusy = if ($v) { [string]$v['availabilityView'] } else { $null }
            WorkingDays    = if ($wh) { @($wh['daysOfWeek'] | ForEach-Object { ([string]$_).ToLowerInvariant() }) } else { @() }
            WorkStart      = if ($wh) { $t = [TimeSpan]::Parse(([string]$wh['startTime']).Substring(0, 8)); [int]$t.TotalMinutes } else { $null }
            WorkEnd        = if ($wh) { $t = [TimeSpan]::Parse(([string]$wh['endTime']).Substring(0, 8)); [int]$t.TotalMinutes } else { $null }
            WorkTimeZone   = if ($wh -and $wh['timeZone']) { [string]$wh['timeZone']['name'] } else { $null }
            WorkOffsetMinutes = if ($wh -and $wh['timeZone']) { try { [int][TimeZoneInfo]::FindSystemTimeZoneById([string]$wh['timeZone']['name']).GetUtcOffset($start).TotalMinutes } catch { $null } } else { $null }
            WorkingHours   = if ($wh) { '{0} {1}-{2} ({3})' -f ((@($wh['daysOfWeek']) -join ' ')), ([string]$wh['startTime']).Substring(0, 5), ([string]$wh['endTime']).Substring(0, 5), [string]$wh['timeZone']['name'] } else { $null }
            Events         = $events
        }
    }
    $Context.FreeBusy = @($results)
    $Context.FreeBusyWindow = [pscustomobject]@{ StartUtc = $start.ToString('yyyy-MM-ddTHH:mm:ssZ'); EndUtc = $end.ToString('yyyy-MM-ddTHH:mm:ssZ'); IntervalMinutes = [int]$cfg.FreeBusyIntervalMinutes; Days = [int]$cfg.FreeBusyDays }
    Write-WscFreeBusyGrid -Results $results -StartUtc $start -IntervalMinutes ([int]$cfg.FreeBusyIntervalMinutes) -Days ([int]$cfg.FreeBusyDays)
    foreach ($r in $results) {
        $d = [ordered]@{ Mailbox = $r.Mailbox; ResponseCode = $r.ResponseCode; Events = @($r.Events).Count; WorkingHours = $r.WorkingHours; AvailabilityView = $r.MergedFreeBusy }
        if ($r.ResponseClass -eq 'Success') {
            $kinds = @($r.Events | Group-Object { ([string]$_.BusyType).ToLowerInvariant() } | Sort-Object Name | ForEach-Object { '{0} {1}' -f $_.Count, $(switch ($_.Name) { 'oof' { 'away' } 'workingelsewhere' { 'elsewhere' } default { $_ } }) }) -join ', '
            $subjects = @($r.Events | Where-Object { $_.Subject }).Count
            $detail = if ($subjects) { ' with subjects (the caller may read the details)' } elseif (@($r.Events).Count) { ' without subjects (free/busy only for the caller)' } else { '' }
            Add-WscStep $Context $stage "Free/busy of $($r.Mailbox)" Passed "$(@($r.Events).Count) calendar item(s)$(if ($kinds) { " ($kinds)" })$detail. Working hours: $(if ($r.WorkingHours) { $r.WorkingHours } else { 'not returned' })." $d
        }
        else {
            $hint = if ("$($r.MessageText)" -match 'FederatedCrossForest|CrossForest|CrossSite|5027|5009') { ' - cross-premises free/busy: Exchange Online asks the on-premises Exchange through EWS (hybrid OAuth). Check that the on-premises EWS URL is published and reachable from Exchange Online, Get-IntraOrganizationConnector (TargetAddressDomains, DiscoveryEndpoint, Enabled), Test-OAuthConnectivity on-premises, and that the servers answer (a timeout right after a restart is normal).' } else { '' }
            Add-WscStep $Context $stage "Free/busy of $($r.Mailbox)" Warning "$($r.ResponseCode): $($r.MessageText)$hint" $d
        }
    }
}

function Get-WscGraphTestFolder {
    <# The test folder under the Inbox: found by name ($filter), or created when Create is set. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage, [switch]$Create)

    if ($Context.TestFolder) { return $Context.TestFolder }
    $name = [string]$Context.Config.FolderName
    $mb = Get-WscGraphMailboxPath $Context
    $find = Invoke-WscGraph -Context $Context -Path "$mb/mailFolders/inbox/childFolders?`$filter=displayName eq '$(ConvertTo-WscODataText $name)'&`$select=id,displayName" -Operation 'GET childFolders (filter)'
    if (-not (Test-WscGraphOk $find)) {
        Add-WscStep $Context $Stage 'Find the test folder' Failed "The Inbox subfolders were not searched: $(Get-WscGraphFailureText -Context $Context -Answer $find)" (Get-WscGraphDetails -Answer $find)
        $Context.Stop = $true
        return $null
    }
    $hit = @($find.Json['value']) | Where-Object { $_ } | Select-Object -First 1
    if ($hit) {
        $Context.TestFolder = [pscustomobject]@{ DisplayName = $name; FolderId = [string]$hit['id']; Path = "\Inbox\$name" }
        Add-WscStep $Context $Stage 'Find the test folder' Passed "The test folder '$name' exists under the Inbox: it is reused." (Get-WscGraphDetails -Answer $find -More ([ordered]@{ Folder = $name; FolderId = $Context.TestFolder.FolderId }))
        return $Context.TestFolder
    }
    if (-not $Create) {
        Add-WscStep $Context $Stage 'Find the test folder' Failed "No folder '$name' under the Inbox: run CreateFolder (or MailCycle) first." (Get-WscGraphDetails -Answer $find)
        $Context.Stop = $true
        return $null
    }
    $made = Invoke-WscGraph -Context $Context -Method POST -Path "$mb/mailFolders/inbox/childFolders" -Body @{ displayName = $name } -Operation 'POST childFolders'
    $d = Get-WscGraphDetails -Answer $made -More ([ordered]@{ Folder = $name; Parent = 'Inbox' })
    if (-not (Test-WscGraphOk $made)) {
        Add-WscStep $Context $Stage 'POST childFolders' Failed "The folder '$name' was not created: $(Get-WscGraphFailureText -Context $Context -Answer $made)" $d
        $Context.Stop = $true
        return $null
    }
    $Context.TestFolder = [pscustomobject]@{ DisplayName = $name; FolderId = [string]$made.Json['id']; Path = "\Inbox\$name" }
    $d.FolderId = $Context.TestFolder.FolderId
    Add-WscAction -Context $Context -Action 'CreateFolder' -Target "Inbox\$name" -Result 'Created' -ItemId $Context.TestFolder.FolderId
    Add-WscStep $Context $Stage 'POST childFolders' Passed "Folder '$name' created under the Inbox of $($Context.Config.Mailbox) (HTTP $($made.HttpStatus))." $d
    return $Context.TestFolder
}

function Invoke-WscGraphSendMail {
    <# sendMail from the mailbox, then the message looked for in the Inbox when the mailbox is a recipient. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'SendMail'
    $cfg = $Context.Config
    $to = if ([string]$cfg.Recipient) { [string]$cfg.Recipient } else { [string]$cfg.Mailbox }
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $subject = "$($script:TestSubjectPrefix) $stamp $([guid]::NewGuid().ToString('N').Substring(0, 6))"
    $how = switch ($Context.Access) { 'Delegate' { "sent as $($cfg.Mailbox) by the signed-in account (Send As, Mail.Send.Shared)" } 'Impersonation' { "sent as $($cfg.Mailbox) by the application (Mail.Send)" } default { "sent by $($cfg.Mailbox)" } }
    $message = [ordered]@{
        message = [ordered]@{ subject = $subject; body = @{ contentType = 'Text'; content = "Test message of Web Services Client for Exchange $($script:ToolVersion) through Microsoft Graph, $how, from $env:COMPUTERNAME at $stamp. It can be deleted." }; toRecipients = @(@{ emailAddress = @{ address = $to } }) }
        saveToSentItems = $true
    }
    $answer = Invoke-WscGraph -Context $Context -Method POST -Path "$(Get-WscGraphMailboxPath $Context)/sendMail" -Body $message -Operation 'POST sendMail'
    $details = Get-WscGraphDetails -Answer $answer -More ([ordered]@{ To = $to; Subject = $subject; SavedIn = 'Sent Items' })
    if (-not (Test-WscGraphOk $answer)) {
        Add-WscStep $Context $stage 'POST sendMail' Failed "The test message was not sent: $(Get-WscGraphFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $true
        return
    }
    Add-WscAction -Context $Context -Action 'SendMail' -Target $to -Result "Sent: $subject"
    $Context.TestMessage = [pscustomobject]@{ Subject = $subject; ItemId = $null; ChangeKey = $null }
    Add-WscStep $Context $stage 'POST sendMail' Passed "Test message accepted by Graph (HTTP $($answer.HttpStatus)) for $to ($how), a copy saved in Sent Items." $details
    if ($to -ine [string]$cfg.Mailbox) { return }

    $wait = [int]$cfg.DeliveryWaitSeconds
    $deadline = [DateTimeOffset]::UtcNow.AddSeconds($wait)
    $tries = 0
    $hit = $null
    do {
        $tries++
        $hit = (Get-WscGraphInboxMessages -Context $Context -Top 1 -Subject $subject -Collapse 'Delivery').Items | Select-Object -First 1
        if ($hit -or [DateTimeOffset]::UtcNow -ge $deadline) { break }
        Wait-WscSeconds 3
    } while ($true)
    $d = [ordered]@{ Subject = $subject; Searches = $tries }
    if ($hit) {
        $Context.TestMessage = $hit
        $d.ItemId = $hit.ItemId; $d.Received = $hit.DateTimeReceived
        Add-WscStep $Context $stage 'Delivery to the Inbox' Passed "The test message arrived in the Inbox of $($cfg.Mailbox) ($tries search(es))." $d
    }
    else {
        Add-WscStep $Context $stage 'Delivery to the Inbox' Warning "The test message is not in the Inbox after $wait s (Test.DeliveryWaitSeconds): transport, a transport rule or a slow delivery. The next steps look for it again." $d
    }
}

function Select-WscGraphTargetItem {
    <# The message to reply to, move or delete: -ItemId, the test message, -ItemSubject (search), or the last test message of the tool. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage)

    if ($Context.ItemId) { return [pscustomobject]@{ ItemId = $Context.ItemId; Subject = '(given by -ItemId)'; From = $null; Source = '-ItemId' } }
    if ($Context.TestMessage -and $Context.TestMessage.ItemId) { return $Context.TestMessage | Select-Object *, @{ n = 'Source'; e = { 'test message of this run' } } }
    if ($Context.ItemSubject) { $text = [string]$Context.ItemSubject; $source = "-ItemSubject '$text'"; $found = Get-WscGraphInboxMessages -Context $Context -Top 25 -Search "subject:$text" }
    elseif ($Context.TestMessage) { $text = $Context.TestMessage.Subject; $source = 'test message of this run'; $found = Get-WscGraphInboxMessages -Context $Context -Top 1 -Subject $text }
    else { $text = $script:TestSubjectPrefix; $source = 'last test message of the tool'; $found = Get-WscGraphInboxMessages -Context $Context -Top 25 -Search 'Web Services Client for Exchange' }
    if (-not (Test-WscGraphOk $found.Answer)) {
        Add-WscStep $Context $Stage 'Find the message' Failed "The Inbox was not searched: $(Get-WscGraphFailureText -Context $Context -Answer $found.Answer)" (Get-WscGraphDetails -Answer $found.Answer)
        $Context.Stop = $true
        return $null
    }
    $hit = $found.Items | Where-Object { $_.Subject -and $_.Subject.IndexOf($text, [StringComparison]::OrdinalIgnoreCase) -ge 0 } | Select-Object -First 1
    if (-not $hit) {
        Add-WscStep $Context $Stage 'Find the message' Failed "No Inbox message whose subject contains '$text' ($source): run SendMail first, or name the message with -ItemId or -ItemSubject." (Get-WscGraphDetails -Answer $found.Answer -More ([ordered]@{ Search = $text }))
        $Context.Stop = $true
        return $null
    }
    if (-not $Context.TestMessage -and -not $Context.ItemSubject) { $Context.TestMessage = $hit }
    return $hit | Select-Object *, @{ n = 'Source'; e = { $source } }
}

function Invoke-WscGraphReplyMail {
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    $stage = 'ReplyMail'
    $item = Select-WscGraphTargetItem -Context $Context -Stage $stage
    if (-not $item) { return }
    $comment = "Reply of Web Services Client for Exchange $($script:ToolVersion) through Microsoft Graph, from $env:COMPUTERNAME at $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))."
    $answer = Invoke-WscGraph -Context $Context -Method POST -Path "$(Get-WscGraphMailboxPath $Context)/messages/$([Uri]::EscapeDataString($item.ItemId))/reply" -Body @{ comment = $comment } -Operation 'POST message/reply'
    $details = Get-WscGraphDetails -Answer $answer -More ([ordered]@{ Message = $item.Subject; From = $item.From; Selected = $item.Source; ItemId = $item.ItemId })
    if (-not (Test-WscGraphOk $answer)) {
        Add-WscStep $Context $stage 'POST reply' Failed "No reply to '$($item.Subject)': $(Get-WscGraphFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $true
        return
    }
    Add-WscAction -Context $Context -Action 'ReplyMail' -Target $(if ($item.From) { $item.From } else { 'sender of the message' }) -Result "Replied to: $($item.Subject)" -ItemId $item.ItemId
    Add-WscStep $Context $stage 'POST reply' Passed "Reply sent to the sender of '$($item.Subject)' ($($item.Source)), saved in Sent Items." $details
}

function Invoke-WscGraphMoveMail {
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    $stage = 'MoveMail'
    $item = Select-WscGraphTargetItem -Context $Context -Stage $stage
    if (-not $item) { return }
    $folder = Get-WscGraphTestFolder -Context $Context -Stage $stage -Create
    if (-not $folder) { return }
    $answer = Invoke-WscGraph -Context $Context -Method POST -Path "$(Get-WscGraphMailboxPath $Context)/messages/$([Uri]::EscapeDataString($item.ItemId))/move" -Body @{ destinationId = $folder.FolderId } -Operation 'POST message/move'
    $details = Get-WscGraphDetails -Answer $answer -More ([ordered]@{ Message = $item.Subject; Selected = $item.Source; To = "Inbox\$($folder.DisplayName)"; ItemId = $item.ItemId })
    if (-not (Test-WscGraphOk $answer)) {
        Add-WscStep $Context $stage 'POST move' Failed "'$($item.Subject)' was not moved: $(Get-WscGraphFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $true
        return
    }
    $newId = if ($answer.Json) { [string]$answer.Json['id'] } else { $null }
    if ($newId) {
        $details.NewItemId = $newId
        $Context.TestMessage = [pscustomobject]@{ Subject = $item.Subject; ItemId = $newId; ChangeKey = $null; From = $item.From }
        if ($Context.ItemId -eq $item.ItemId) { $Context.ItemId = $newId }
    }
    Add-WscAction -Context $Context -Action 'MoveMail' -Target "Inbox\$($folder.DisplayName)" -Result "Moved: $($item.Subject)" -ItemId $newId
    Add-WscStep $Context $stage 'POST move' Passed "'$($item.Subject)' moved to Inbox\$($folder.DisplayName)$(if ($newId) { ': Graph returned the message with its new ID' })." $details
}

function Invoke-WscGraphDeleteMail {
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    $stage = 'DeleteMail'
    $item = Select-WscGraphTargetItem -Context $Context -Stage $stage
    if (-not $item) { return }
    $mode = [string]$Context.Config.DeleteMode
    $path = "$(Get-WscGraphMailboxPath $Context)/messages/$([Uri]::EscapeDataString($item.ItemId))"
    $answer = switch ($mode) {
        'SoftDelete' { Invoke-WscGraph -Context $Context -Method POST -Path "$path/move" -Body @{ destinationId = 'recoverableitemsdeletions' } -Operation 'POST message/move (recoverable items)' }
        'HardDelete' { Invoke-WscGraph -Context $Context -Method POST -Path "$path/permanentDelete" -Operation 'POST message/permanentDelete' }
        default { Invoke-WscGraph -Context $Context -Method DELETE -Path $path -Operation 'DELETE message' }
    }
    $details = Get-WscGraphDetails -Answer $answer -More ([ordered]@{ Message = $item.Subject; Selected = $item.Source; DeleteMode = $mode; ItemId = $item.ItemId })
    if (-not (Test-WscGraphOk $answer)) {
        Add-WscStep $Context $stage 'Delete' Failed "'$($item.Subject)' was not deleted: $(Get-WscGraphFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $true
        return
    }
    $where = switch ($mode) { 'MoveToDeletedItems' { 'moved to Deleted Items (DELETE)' } 'SoftDelete' { 'moved to Recoverable Items\Deletions' } default { 'permanently deleted (permanentDelete)' } }
    Add-WscAction -Context $Context -Action 'DeleteMail' -Target $mode -Result "Deleted: $($item.Subject)" -ItemId $item.ItemId
    Add-WscStep $Context $stage 'Delete' Passed "'$($item.Subject)' $where." $details
    $Context.TestMessage = $null
}

#endregion

#region Prerequisites and token ---------------------------------------------------------------------

function Invoke-WscStageDiscoveryGraph {
    <# Graph without sign-in: certificate of graph.microsoft.com, the challenge of Graph, a forged token refused. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Discovery'
    Add-WscTlsCheck -Context $Context -Stage $stage -HostName 'graph.microsoft.com' -Port 443
    $path = "$(Get-WscGraphMailboxPath $Context)/mailFolders/inbox?`$select=id"
    $anonymous = Invoke-WscGraph -Context $Context -Path $path -Credentials None -Operation 'GET mailFolders/inbox (no token)'
    $info = Get-WscChallengeInfo -Challenges @(Get-WscField $anonymous.Response 'Challenges')
    $d = Get-WscGraphDetails -Answer $anonymous -More ([ordered]@{ Schemes = $info.Schemes -join ', '; AuthorizationUri = $info.AuthorizationUri })
    if ($anonymous.HttpStatus -eq 401 -and $info.Bearer) { Add-WscStep $Context $stage 'Graph challenge' Passed "Microsoft Graph asks for an Entra ID token (Bearer$(if ($info.AuthorizationUri) { ", $($info.AuthorizationUri)" }))." $d }
    else { Add-WscStep $Context $stage 'Graph challenge' Warning "Microsoft Graph answered a request without a token with HTTP $($anonymous.HttpStatus) (401 expected): a proxy on the path?" $d }
    $forged = Invoke-WscGraph -Context $Context -Path $path -Credentials Token -Token $script:InvalidToken -Operation 'GET mailFolders/inbox (forged token)'
    $d = Get-WscGraphDetails -Answer $forged
    if ($forged.HttpStatus -eq 401) { Add-WscStep $Context $stage 'Forged token' Passed "A forged bearer token is refused (HTTP 401, $($forged.ErrorCode))." $d }
    else { Add-WscStep $Context $stage 'Forged token' Warning "A forged bearer token returned HTTP $($forged.HttpStatus) (401 expected)." $d }
}

function Test-WscGraphTokenClaims {
    <#
        A Graph token: audience Microsoft Graph, tenant, expiry, and the permissions the scenario needs -
        delegated (scp; the .Shared ones for delegate access) or application (roles).
    #>
    param([AllowNull()][hashtable]$Claims, [Parameter(Mandatory = $true)][hashtable]$Configuration, [string]$TenantId, [string]$Access, [string[]]$Stages)

    if ($null -eq $Claims) { return [pscustomobject]@{ Status = 'Warning'; Message = 'The access token is not a readable JWT: its claims were not checked.'; Details = [ordered]@{} } }
    $app = [string]$Configuration.Context -eq 'Application'
    $scopes = @(([string]$Claims['scp']).Split(' ', [StringSplitOptions]::RemoveEmptyEntries))
    $roles = @($Claims['roles'] | Where-Object { $_ })
    $granted = if ($app) { $roles } else { $scopes }
    $exp = 0L; $expires = $null
    if ([long]::TryParse([string]$Claims['exp'], [ref]$exp)) { $expires = [DateTimeOffset]::FromUnixTimeSeconds($exp) }
    $details = [ordered]@{
        Audience = @($Claims['aud']) -join ', '; Scope = $scopes -join ' '; Roles = $roles -join ', '; User = [string]$Claims['upn']; ClientId = [string]$Claims['appid']
        AppName = [string]$Claims['app_displayname']; TenantId = [string]$Claims['tid']; ExpiresUtc = if ($expires) { $expires.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ') } else { $null }
        MinutesLeft = if ($expires) { [int][Math]::Floor(($expires - [DateTimeOffset]::UtcNow).TotalMinutes) } else { $null }
    }
    if ($expires -and $expires -le [DateTimeOffset]::UtcNow) { return [pscustomobject]@{ Status = 'Failed'; Message = "The token expired at $($details.ExpiresUtc)."; Details = $details } }
    $issues = [Collections.Generic.List[string]]::new()
    if (@($Claims['aud']) -notcontains 'https://graph.microsoft.com' -and @($Claims['aud']) -notcontains $script:GraphAppId) { $issues.Add("the audience '$($details.Audience)' is not Microsoft Graph") }
    if ($TenantId -and $details.TenantId -and $details.TenantId -ine $TenantId) { $issues.Add("the token comes from the tenant $($details.TenantId), not from $TenantId") }
    $shared = -not $app -and $Access -eq 'Delegate'
    $needs = [ordered]@{}
    if (@($Stages | Where-Object { $_ -in 'Endpoint', 'Folders', 'ReadMail' }).Count) { $needs['read the mail'] = @('Mail.ReadBasic', 'Mail.Read', 'Mail.ReadWrite') }
    if (@($Stages | Where-Object { $_ -in 'CreateFolder', 'MoveMail', 'DeleteMail' }).Count) { $needs['change the mailbox'] = @('Mail.ReadWrite') }
    if (@($Stages | Where-Object { $_ -in 'SendMail', 'ReplyMail' }).Count) { $needs['send'] = @('Mail.Send') }
    if ($Stages -contains 'FreeBusy') { $needs['read free/busy'] = @('Calendars.ReadBasic', 'Calendars.Read', 'Calendars.ReadWrite') }
    if (@($Stages | Where-Object { $_ -in 'SeedData', 'CleanData' }).Count) { $needs['change the mailbox'] = @('Mail.ReadWrite'); $needs['write the calendar'] = @('Calendars.ReadWrite') }
    if ($Stages -contains 'SeedData') { $needs['send'] = @('Mail.Send') }
    $missing = foreach ($k in $needs.Keys) {
        $any = @($needs[$k] | ForEach-Object { if ($shared -and $_ -like 'Mail.*') { "$_.Shared" } else { $_ } })
        if (-not @($granted | Where-Object { $_ -in $any }).Count) { "$k ($($any -join ' or '))" }
    }
    if ($missing) {
        $kind = if ($app) { 'application permission (roles) with admin consent' } else { 'delegated permission (scp)' }
        $issues.Add("the token has no $kind to $($missing -join ', ')")
    }
    if ($issues.Count) { return [pscustomobject]@{ Status = 'Warning'; Message = 'Token received, but ' + ($issues -join '; ') + '.'; Details = $details } }
    $left = if ($null -ne $details.MinutesLeft) { " (valid for $($details.MinutesLeft) min)" } else { '' }
    return [pscustomobject]@{ Status = 'Passed'; Message = "Audience Microsoft Graph, tenant and the $(if ($app) { 'application' } else { 'delegated' }) permissions this scenario needs ($(@($granted | Where-Object { $_ -match '^(Mail|Calendars)\.' }) -join ', '))$left."; Details = $details }
}

#endregion
