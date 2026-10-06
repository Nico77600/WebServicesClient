<#
.SYNOPSIS
    Web Services Client for Exchange - simulated Exchange (EWS, Autodiscover), AD FS and Entra ID, for the tests and the documentation images.

.DESCRIPTION
    Dot-source this file. Get-SimHttpResponse answers every HTTP request of the tool like the real servers:
      Autodiscover v2 (JSON) and classic Autodiscover (POX, authenticated), AD FS and Entra ID (metadata,
      user realm, device code, token, client credentials), and EWS: 401 with the schemes of the state
      (Negotiate, NTLM, Basic, Bearer with authorization_uri), an NTLM type 2 challenge built byte by byte
      (server EXCH01, domain CONTOSO / contoso.test), Basic accounts, bearer tokens, SOAP answers of the
      operations (GetFolder, FindFolder, FindItem, GetItem, CreateFolder, CreateItem, MoveItem, DeleteItem,
      GetUserAvailability, GetServerTimeZones), ErrorServerBusy once, impersonation or delegate access refused.

      New-SimState          what the simulated organisation does
      New-SimToken          unsigned JWT with the claims of AD FS or Entra ID
      Get-SimHttpResponse   answer to one HTTP request of the tool
      Install-SimExchange   replaces the network inside the loaded module (Pester, documentation images)

    Nothing here is used by the tool itself, and it is not part of the package.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:SimTenantId = '7d4e2a91-3c5b-4f6e-8a1d-2b9c0e5f4a37'
$script:SimNs = 'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"'

function New-SimToken {
    param([string]$Audience = 'https://mail.contoso.test/', [string]$Scope = 'EWS.AccessAsUser.All', [string]$Upn = 'ews-test@contoso.test', [string]$AppId = 'd3590ed6-52b3-4102-aeff-aad2292ab01c', [string]$TenantId, [string[]]$Roles, [int]$ExpiresIn = 3600)
    $encode = { param($o) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($o | ConvertTo-Json -Compress))).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
    $payload = [ordered]@{ aud = $Audience; iss = 'http://adfs.contoso.test/adfs/services/trust'; appid = $AppId; exp = [DateTimeOffset]::UtcNow.AddSeconds($ExpiresIn).ToUnixTimeSeconds() }
    if ($Scope) { $payload.scp = $Scope; $payload.upn = $Upn }
    if ($Roles) { $payload.roles = $Roles; $payload.app_displayname = 'EWS test application' }
    if ($TenantId) { $payload.tid = $TenantId; $payload.iss = "https://sts.windows.net/$TenantId/" }
    '{0}.{1}.{2}' -f (& $encode @{ alg = 'RS256'; typ = 'JWT' }), (& $encode $payload), 'c2lnbmF0dXJl'
}

function New-SimState {
    <# Default: a healthy on-premises Exchange SE with AD FS, two Inbox messages (one subject is a spreadsheet formula). #>
    @{
        Calls             = [Collections.Generic.List[object]]::new()
        ValidTokens       = [Collections.Generic.List[string]]::new()
        TokenAudience     = 'https://mail.contoso.test/'
        TenantId          = $null
        AppRoles          = @('full_access_as_app')
        AutodiscoverV2    = 'https://mail.contoso.test/EWS/Exchange.asmx'
        AutodiscoverPox   = $false
        Schemes           = @('Negotiate', 'NTLM', 'Basic realm="mail.contoso.test"')
        AuthorizationUri  = 'https://adfs.contoso.test/adfs/oauth2/authorize'
        TrustedIssuers    = $null
        BasicUsers        = @{ 'ews-test@contoso.test' = 'Sim-Pa55word!'; 'CONTOSO\ews-test' = 'Sim-Pa55word!' }
        WindowsAccepted   = $true
        ServerBusyOnce    = $false
        DenyImpersonation = $false
        DenyDelegate      = $false
        ServerVersion     = @{ Major = 15; Minor = 2; Build = 2562; Revision = 17 }
        CertificateDays   = 365
        Mailbox           = 'ews-test@contoso.test'
        Folders           = [Collections.Generic.List[object]]@(
            @{ Id = 'F-INBOX'; Parent = 'F-ROOT'; Name = 'Inbox'; Class = 'IPF.Note'; Total = 2; Unread = 1; Children = 0 }
            @{ Id = 'F-SENT'; Parent = 'F-ROOT'; Name = 'Sent Items'; Class = 'IPF.Note'; Total = 0; Unread = 0; Children = 0 }
            @{ Id = 'F-CAL'; Parent = 'F-ROOT'; Name = 'Calendar'; Class = 'IPF.Appointment'; Total = 3; Unread = 0; Children = 0 }
        )
        Inbox             = [Collections.Generic.List[object]]@(
            @{ Id = 'M-1'; Subject = 'Budget 2027'; From = 'alice@contoso.test'; Received = '2026-10-05T08:00:00Z'; Read = 'false'; Body = 'Hello, here is the budget.' }
            @{ Id = 'M-2'; Subject = '=HYPERLINK("http://x")'; From = 'bob@contoso.test'; Received = '2026-10-04T08:00:00Z'; Read = 'true'; Body = 'Formula subject.' }
        )
        FreeBusy          = @{ 'ews-test@contoso.test' = 'Success'; 'room1@contoso.test' = 'Success'; 'nobody@contoso.test' = 'ErrorMailRecipientNotFound' }
        Next              = 10
        Calendar          = [Collections.Generic.List[object]]::new()
        Filed             = [Collections.Generic.List[object]]::new()
        EwsBlocked        = $false
        GraphForbidden    = $false
        GraphScopes       = 'Mail.ReadWrite Mail.Send Calendars.Read User.Read'
        GraphRoles        = @('Mail.ReadWrite', 'Mail.Send', 'Calendars.Read')
    }
}

function New-SimResponse([int]$Code, [string]$Body = '', [hashtable]$Headers = @{}, [string[]]$Challenges = @(), [string]$ContentType = 'text/xml; charset=utf-8') {
    if ($Body -and -not $Headers.ContainsKey('Content-Type')) { $Headers['Content-Type'] = $ContentType }
    [pscustomobject]@{ StatusCode = $Code; Headers = $Headers; Challenges = $Challenges; Body = [Text.Encoding]::UTF8.GetBytes($Body) }
}

function New-SimNtlmChallenge {
    <# NTLM type 2 message (MS-NLMP): target CONTOSO, AV pairs of EXCH01.contoso.test, version 10.0.26100. #>
    $u = { param([string]$s) [Text.Encoding]::Unicode.GetBytes($s) }
    $av = [Collections.Generic.List[byte]]::new()
    foreach ($pair in @(@(2, 'CONTOSO'), @(1, 'EXCH01'), @(4, 'contoso.test'), @(3, 'EXCH01.contoso.test'), @(5, 'contoso.test'))) {
        $v = & $u $pair[1]; $av.AddRange([BitConverter]::GetBytes([uint16]$pair[0])); $av.AddRange([BitConverter]::GetBytes([uint16]$v.Length)); $av.AddRange([byte[]]$v)
    }
    $av.AddRange([BitConverter]::GetBytes([uint16]7)); $av.AddRange([BitConverter]::GetBytes([uint16]8)); $av.AddRange([BitConverter]::GetBytes([DateTime]::UtcNow.ToFileTimeUtc()))
    $av.AddRange([byte[]](0, 0, 0, 0))
    $target = & $u 'CONTOSO'
    $m = [Collections.Generic.List[byte]]::new()
    $m.AddRange([byte[]](0x4E, 0x54, 0x4C, 0x4D, 0x53, 0x53, 0x50, 0x00)); $m.AddRange([BitConverter]::GetBytes([uint32]2))
    $m.AddRange([BitConverter]::GetBytes([uint16]$target.Length)); $m.AddRange([BitConverter]::GetBytes([uint16]$target.Length)); $m.AddRange([BitConverter]::GetBytes([uint32]56))
    $m.AddRange([BitConverter]::GetBytes([uint32]3800662549)); $m.AddRange([byte[]](1, 2, 3, 4, 5, 6, 7, 8)); $m.AddRange([byte[]]::new(8))
    $m.AddRange([BitConverter]::GetBytes([uint16]$av.Count)); $m.AddRange([BitConverter]::GetBytes([uint16]$av.Count)); $m.AddRange([BitConverter]::GetBytes([uint32](56 + $target.Length)))
    $m.AddRange([byte[]](10, 0)); $m.AddRange([BitConverter]::GetBytes([uint16]26100)); $m.AddRange([byte[]](0, 0, 0, 15))
    $m.AddRange([byte[]]$target); $m.AddRange($av)
    [Convert]::ToBase64String($m.ToArray())
}

function Get-SimSoap([hashtable]$State, [string]$Body) {
    $v = $State.ServerVersion
    "<?xml version=""1.0"" encoding=""utf-8""?><s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/""><s:Header><h:ServerVersionInfo MajorVersion=""$($v.Major)"" MinorVersion=""$($v.Minor)"" MajorBuildNumber=""$($v.Build)"" MinorBuildNumber=""$($v.Revision)"" Version=""V2018_01_08"" xmlns:h=""http://schemas.microsoft.com/exchange/services/2006/types"" xmlns=""http://schemas.microsoft.com/exchange/services/2006/types""/></s:Header><s:Body>$Body</s:Body></s:Envelope>"
}

function Get-SimFault([string]$Code, [string]$Message, [int]$BackOff = 0) {
    $extra = if ($BackOff) { "<t:MessageXml xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types""><t:Value Name=""BackOffMilliseconds"">$BackOff</t:Value></t:MessageXml>" } else { '' }
    "<?xml version=""1.0"" encoding=""utf-8""?><s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/""><s:Body><s:Fault><faultcode xmlns:a=""http://schemas.microsoft.com/exchange/services/2006/types"">a:$Code</faultcode><faultstring xml:lang=""en-US"">$Message</faultstring><detail><e:ResponseCode xmlns:e=""http://schemas.microsoft.com/exchange/services/2006/errors"">$Code</e:ResponseCode><e:Message xmlns:e=""http://schemas.microsoft.com/exchange/services/2006/errors"">$Message</e:Message>$extra</detail></s:Fault></s:Body></s:Envelope>"
}

function Get-SimItemXml([hashtable]$m, [switch]$Full) {
    $body = if ($Full) { "<t:Body BodyType=""Text"">$([Security.SecurityElement]::Escape($m.Body))</t:Body><t:ToRecipients><t:Mailbox><t:EmailAddress>ews-test@contoso.test</t:EmailAddress></t:Mailbox></t:ToRecipients>" } else { '' }
    "<t:Message><t:ItemId Id=""$($m.Id)"" ChangeKey=""CK-$($m.Id)""/><t:ItemClass>IPM.Note</t:ItemClass><t:Subject>$([Security.SecurityElement]::Escape($m.Subject))</t:Subject>$body<t:DateTimeReceived>$($m.Received)</t:DateTimeReceived><t:Size>2048</t:Size><t:HasAttachments>false</t:HasAttachments><t:From><t:Mailbox><t:Name>$($m.From)</t:Name><t:EmailAddress>$($m.From)</t:EmailAddress></t:Mailbox></t:From><t:IsRead>$($m.Read)</t:IsRead><t:InternetMessageId>&lt;$($m.Id)@contoso.test&gt;</t:InternetMessageId></t:Message>"
}

function Get-SimSchedule {
    <#
        A plausible week for a mailbox, from StartUtc (local midnight of today in UTC): meetings, a tentative one,
        a day out of office, an afternoon elsewhere - varied by mailbox. Events in UTC and the merged view
        (0 free, 1 tentative, 2 busy, 3 away, 4 elsewhere), one digit per interval, like Exchange returns it.
    #>
    param([string]$Mailbox, [datetime]$StartUtc, [int]$Days = 7, [int]$IntervalMinutes = 30)

    $seed = [int]$Mailbox.ToLowerInvariant()[0] % 3
    $plan = switch ($seed) {
        0 { @(@(0, 9, 0, 60, 'Busy', 'Weekly review', 'Room 1'), @(0, 14, 0, 90, 'Tentative', 'Design sync', 'Teams'), @(1, 8, 0, 600, 'OOF', 'Out of office', ''), @(2, 11, 0, 60, 'Busy', 'Customer call', 'Teams'), @(2, 16, 30, 30, 'Busy', '1:1', ''), @(3, 13, 0, 240, 'WorkingElsewhere', 'Working from home', ''), @(4, 10, 0, 120, 'Busy', 'Change board', 'Room 2')) }
        1 { @(@(0, 10, 0, 60, 'Busy', 'Project kick-off', 'Room 3'), @(0, 15, 0, 60, 'Busy', 'Exchange migration', 'Teams'), @(1, 9, 30, 90, 'Busy', 'Training', 'Room 1'), @(2, 14, 0, 120, 'Tentative', 'Workshop', 'Room 2'), @(3, 9, 0, 60, 'Busy', 'Weekly review', 'Room 1'), @(4, 13, 0, 300, 'OOF', 'Day off afternoon', '')) }
        default { @(@(0, 8, 30, 30, 'Busy', 'Stand-up', 'Teams'), @(0, 12, 0, 60, 'Busy', 'Lunch with team', ''), @(0, 16, 0, 60, 'Busy', 'Security review', 'Room 4'), @(1, 8, 30, 30, 'Busy', 'Stand-up', 'Teams'), @(2, 8, 30, 30, 'Busy', 'Stand-up', 'Teams'), @(2, 10, 0, 180, 'WorkingElsewhere', 'Customer site', 'Lyon'), @(3, 8, 30, 30, 'Busy', 'Stand-up', 'Teams'), @(3, 15, 0, 60, 'Tentative', 'Optional demo', 'Teams')) }
    }
    $localStart = $StartUtc.ToLocalTime().Date
    $events = foreach ($p in $plan) {
        if ($p[0] -ge $Days) { continue }
        $s = $localStart.AddDays($p[0]).AddHours($p[1]).AddMinutes($p[2]).ToUniversalTime()
        [pscustomobject]@{ Start = $s; End = $s.AddMinutes($p[3]); BusyType = $p[4]; Subject = $p[5]; Location = $p[6] }
    }
    $code = @{ Free = 0; Tentative = 1; Busy = 2; OOF = 3; WorkingElsewhere = 4 }
    $rank = @{ 3 = 5; 2 = 4; 1 = 3; 4 = 2; 0 = 1 }
    $slots = [int]($Days * 1440 / $IntervalMinutes)
    $view = [Text.StringBuilder]::new()
    for ($i = 0; $i -lt $slots; $i++) {
        $t0 = $StartUtc.AddMinutes($i * $IntervalMinutes); $t1 = $t0.AddMinutes($IntervalMinutes); $best = 0
        foreach ($e in $events) { if ($e.Start -lt $t1 -and $e.End -gt $t0) { $v = $code[$e.BusyType]; if ($rank[$v] -gt $rank[$best]) { $best = $v } } }
        [void]$view.Append($best)
    }
    [pscustomobject]@{ Events = @($events); View = $view.ToString() }
}

function Get-SimEwsAnswer([hashtable]$State, [string]$Operation, [string]$Soap) {
    $ok = { param([string]$Inner) Get-SimSoap $State $Inner }
    $success = 'ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode>'
    switch ($Operation) {
        'GetServerTimeZones' { return & $ok "<m:GetServerTimeZonesResponse $($script:SimNs)><m:ResponseMessages><m:GetServerTimeZonesResponseMessage $success<m:TimeZoneDefinitions><t:TimeZoneDefinition Id=""UTC"" Name=""(UTC) Coordinated Universal Time""/></m:TimeZoneDefinitions></m:GetServerTimeZonesResponseMessage></m:ResponseMessages></m:GetServerTimeZonesResponse>" }
        'GetFolder' {
            $f = $State.Folders[0]
            return & $ok "<m:GetFolderResponse $($script:SimNs)><m:ResponseMessages><m:GetFolderResponseMessage $success<m:Folders><t:Folder><t:FolderId Id=""$($f.Id)"" ChangeKey=""AQ""/><t:ParentFolderId Id=""F-ROOT""/><t:FolderClass>IPF.Note</t:FolderClass><t:DisplayName>Inbox</t:DisplayName><t:TotalCount>$($State.Inbox.Count)</t:TotalCount><t:ChildFolderCount>0</t:ChildFolderCount><t:UnreadCount>1</t:UnreadCount></t:Folder></m:Folders></m:GetFolderResponseMessage></m:ResponseMessages></m:GetFolderResponse>"
        }
        'FindFolder' {
            $parentId = [regex]::Match($Soap, '<m:ParentFolderIds><t:FolderId Id="([^"]*)"').Groups[1].Value
            $parentName = [regex]::Match($Soap, '<m:ParentFolderIds><t:DistinguishedFolderId Id="([^"]*)"').Groups[1].Value
            $list = @($State.Folders)
            if ($parentId) { $list = @($list | Where-Object { $_.Parent -eq $parentId }) }
            elseif ($parentName -eq 'inbox') { $list = @($list | Where-Object { $_.Parent -eq 'F-INBOX' }) }
            $name = [regex]::Match($Soap, 'FieldURI="folder:DisplayName"/><t:FieldURIOrConstant><t:Constant Value="([^"]*)"').Groups[1].Value
            if ($name) { $list = @($list | Where-Object { $_.Name -eq [Net.WebUtility]::HtmlDecode($name) }) }
            $xml = ($list | ForEach-Object { "<t:Folder><t:FolderId Id=""$($_.Id)"" ChangeKey=""AQ""/><t:ParentFolderId Id=""$($_.Parent)""/><t:FolderClass>$($_.Class)</t:FolderClass><t:DisplayName>$($_.Name)</t:DisplayName><t:TotalCount>$($_.Total)</t:TotalCount><t:ChildFolderCount>$($_.Children)</t:ChildFolderCount><t:UnreadCount>$($_.Unread)</t:UnreadCount></t:Folder>" }) -join ''
            return & $ok "<m:FindFolderResponse $($script:SimNs)><m:ResponseMessages><m:FindFolderResponseMessage $success<m:RootFolder IndexedPagingOffset=""$($list.Count)"" TotalItemsInView=""$($list.Count)"" IncludesLastItemInRange=""true""><t:Folders>$xml</t:Folders></m:RootFolder></m:FindFolderResponseMessage></m:ResponseMessages></m:FindFolderResponse>"
        }
        'CreateFolder' {
            $name = [Net.WebUtility]::HtmlDecode([regex]::Match($Soap, '<t:DisplayName>([^<]*)</t:DisplayName>').Groups[1].Value)
            $parent = [regex]::Match($Soap, '<m:ParentFolderId><t:FolderId Id="([^"]*)"').Groups[1].Value; if (-not $parent) { $parent = 'F-INBOX' }
            $id = "F-$($State.Next)"; $State.Next++
            $State.Folders.Add(@{ Id = $id; Parent = $parent; Name = $name; Class = 'IPF.Note'; Total = 0; Unread = 0; Children = 0 })
            return & $ok "<m:CreateFolderResponse $($script:SimNs)><m:ResponseMessages><m:CreateFolderResponseMessage $success<m:Folders><t:Folder><t:FolderId Id=""$id"" ChangeKey=""AQ""/></t:Folder></m:Folders></m:CreateFolderResponseMessage></m:ResponseMessages></m:CreateFolderResponse>"
        }
        'FindItem' {
            $inFolder = [regex]::Match($Soap, '<m:ParentFolderIds><t:FolderId Id="([^"]*)"').Groups[1].Value
            if ($inFolder) { return & $ok "<m:FindItemResponse $($script:SimNs)><m:ResponseMessages><m:FindItemResponseMessage $success<m:RootFolder TotalItemsInView=""0"" IncludesLastItemInRange=""true""><t:Items>$((@($State.Filed | Where-Object { $_.Folder -eq $inFolder }) | ForEach-Object { Get-SimItemXml $_ }) -join '')</t:Items></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>" }
            if ($Soap -match '<m:CalendarView') {
                $xml = ($State.Calendar | ForEach-Object { "<t:CalendarItem><t:ItemId Id=""$($_.Id)"" ChangeKey=""CK""/><t:Subject>$([Security.SecurityElement]::Escape($_.Subject))</t:Subject><t:Start>$($_.Start)</t:Start></t:CalendarItem>" }) -join ''
                return & $ok "<m:FindItemResponse $($script:SimNs)><m:ResponseMessages><m:FindItemResponseMessage $success<m:RootFolder TotalItemsInView=""$($State.Calendar.Count)"" IncludesLastItemInRange=""true""><t:Items>$xml</t:Items></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>"
            }
            $list = @($State.Inbox)
            $text = [regex]::Match($Soap, '<t:Contains[^>]*><t:FieldURI FieldURI="item:Subject"/><t:Constant Value="([^"]*)"').Groups[1].Value
            if ($text) { $text = [Net.WebUtility]::HtmlDecode($text); $list = @($list | Where-Object { $_.Subject.IndexOf($text, [StringComparison]::OrdinalIgnoreCase) -ge 0 }) }
            $max = [int][regex]::Match($Soap, 'MaxEntriesReturned="(\d+)"').Groups[1].Value
            $shown = @($list | Sort-Object { $_.Received } -Descending | Select-Object -First $max)
            $xml = ($shown | ForEach-Object { Get-SimItemXml $_ }) -join ''
            return & $ok "<m:FindItemResponse $($script:SimNs)><m:ResponseMessages><m:FindItemResponseMessage $success<m:RootFolder IndexedPagingOffset=""$($shown.Count)"" TotalItemsInView=""$($list.Count)"" IncludesLastItemInRange=""true""><t:Items>$xml</t:Items></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>"
        }
        'GetItem' {
            $id = [regex]::Match($Soap, '<t:ItemId Id="([^"]*)"').Groups[1].Value
            $m = $State.Inbox | Where-Object { $_.Id -eq $id } | Select-Object -First 1
            if (-not $m) { return & $ok "<m:GetItemResponse $($script:SimNs)><m:ResponseMessages><m:GetItemResponseMessage ResponseClass=""Error""><m:MessageText>The specified object was not found in the store.</m:MessageText><m:ResponseCode>ErrorItemNotFound</m:ResponseCode></m:GetItemResponseMessage></m:ResponseMessages></m:GetItemResponse>" }
            return & $ok "<m:GetItemResponse $($script:SimNs)><m:ResponseMessages><m:GetItemResponseMessage $success<m:Items>$(Get-SimItemXml $m -Full)</m:Items></m:GetItemResponseMessage></m:ResponseMessages></m:GetItemResponse>"
        }
        'CreateItem' {
            if ($Soap -match '<t:CalendarItem>') {
                $State.Calendar.Add(@{ Id = "C-$($State.Next)"; Subject = [Net.WebUtility]::HtmlDecode([regex]::Match($Soap, '<t:Subject>([^<]*)</t:Subject>').Groups[1].Value); Start = [regex]::Match($Soap, '<t:Start>([^<]*)<').Groups[1].Value; Status = [regex]::Match($Soap, '<t:LegacyFreeBusyStatus>([^<]*)<').Groups[1].Value }); $State.Next++
            }
            elseif ($Soap -match '<t:ReplyToItem>') {
                $id = [regex]::Match($Soap, '<t:ReferenceItemId Id="([^"]*)"').Groups[1].Value
                if (-not ($State.Inbox | Where-Object { $_.Id -eq $id })) { return & $ok "<m:CreateItemResponse $($script:SimNs)><m:ResponseMessages><m:CreateItemResponseMessage ResponseClass=""Error""><m:MessageText>The specified object was not found in the store.</m:MessageText><m:ResponseCode>ErrorItemNotFound</m:ResponseCode><m:Items/></m:CreateItemResponseMessage></m:ResponseMessages></m:CreateItemResponse>" }
            }
            else {
                $subject = [Net.WebUtility]::HtmlDecode([regex]::Match($Soap, '<t:Subject>([^<]*)</t:Subject>').Groups[1].Value)
                $to = [regex]::Match($Soap, '<t:ToRecipients><t:Mailbox><t:EmailAddress>([^<]*)<').Groups[1].Value
                if ($to -ieq $State.Mailbox) { $State.Inbox.Add(@{ Id = "M-$($State.Next)"; Subject = $subject; From = $State.Mailbox; Received = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'); Read = 'false'; Body = 'Test message.' }); $State.Next++ }
            }
            return & $ok "<m:CreateItemResponse $($script:SimNs)><m:ResponseMessages><m:CreateItemResponseMessage $success<m:Items/></m:CreateItemResponseMessage></m:ResponseMessages></m:CreateItemResponse>"
        }
        'MoveItem' {
            $id = [regex]::Match($Soap, '<m:ItemIds><t:ItemId Id="([^"]*)"').Groups[1].Value
            $m = $State.Inbox | Where-Object { $_.Id -eq $id } | Select-Object -First 1
            if (-not $m) { return & $ok "<m:MoveItemResponse $($script:SimNs)><m:ResponseMessages><m:MoveItemResponseMessage ResponseClass=""Error""><m:MessageText>Not found.</m:MessageText><m:ResponseCode>ErrorItemNotFound</m:ResponseCode><m:Items/></m:MoveItemResponseMessage></m:ResponseMessages></m:MoveItemResponse>" }
            $m.Id = "$($m.Id)-MOVED"
            return & $ok "<m:MoveItemResponse $($script:SimNs)><m:ResponseMessages><m:MoveItemResponseMessage $success<m:Items><t:Message><t:ItemId Id=""$($m.Id)"" ChangeKey=""CK2""/></t:Message></m:Items></m:MoveItemResponseMessage></m:ResponseMessages></m:MoveItemResponse>"
        }
        'DeleteItem' {
            $id = [regex]::Match($Soap, '<t:ItemId Id="([^"]*)"').Groups[1].Value
            $m = $State.Inbox | Where-Object { $_.Id -eq $id } | Select-Object -First 1
            if ($m) { [void]$State.Inbox.Remove($m) }
            $c = $State.Calendar | Where-Object { $_.Id -eq $id } | Select-Object -First 1
            if ($c) { [void]$State.Calendar.Remove($c) }
            return & $ok "<m:DeleteItemResponse $($script:SimNs)><m:ResponseMessages><m:DeleteItemResponseMessage $success</m:DeleteItemResponseMessage></m:ResponseMessages></m:DeleteItemResponse>"
        }
        'DeleteFolder' {
            $id = [regex]::Match($Soap, '<t:FolderId Id="([^"]*)"').Groups[1].Value
            foreach ($f in @($State.Folders | Where-Object { $_.Id -eq $id -or $_.Parent -eq $id })) { [void]$State.Folders.Remove($f) }
            return & $ok "<m:DeleteFolderResponse $($script:SimNs)><m:ResponseMessages><m:DeleteFolderResponseMessage $success</m:DeleteFolderResponseMessage></m:ResponseMessages></m:DeleteFolderResponse>"
        }
        'GetUserAvailability' {
            $addresses = @([regex]::Matches($Soap, '<t:Address>([^<]*)</t:Address>') | ForEach-Object { $_.Groups[1].Value })
            $start = [datetime]::SpecifyKind([datetime][regex]::Match($Soap, '<t:StartTime>([^<]*)<').Groups[1].Value, 'Utc')
            $end = [datetime]::SpecifyKind([datetime][regex]::Match($Soap, '<t:EndTime>([^<]*)<').Groups[1].Value, 'Utc')
            $interval = [int][regex]::Match($Soap, '<t:MergedFreeBusyIntervalInMinutes>(\d+)<').Groups[1].Value
            $responses = foreach ($a in $addresses) {
                $code = if ($State.FreeBusy.ContainsKey($a)) { $State.FreeBusy[$a] } else { 'Success' }
                if ($code -ne 'Success') { "<m:FreeBusyResponse><m:ResponseMessage ResponseClass=""Error""><m:MessageText>No mailbox found for $a.</m:MessageText><m:ResponseCode>$code</m:ResponseCode></m:ResponseMessage></m:FreeBusyResponse>"; continue }
                $f = 'yyyy-MM-ddTHH:mm:ss'
                $plan = Get-SimSchedule -Mailbox $a -StartUtc $start -Days ([int]($end - $start).TotalDays) -IntervalMinutes $interval
                $items = ($plan.Events | ForEach-Object { "<t:CalendarEvent><t:StartTime>$($_.Start.ToString($f))</t:StartTime><t:EndTime>$($_.End.ToString($f))</t:EndTime><t:BusyType>$($_.BusyType)</t:BusyType><t:CalendarEventDetails><t:Subject>$($_.Subject)</t:Subject><t:Location>$($_.Location)</t:Location><t:IsMeeting>true</t:IsMeeting><t:IsPrivate>false</t:IsPrivate></t:CalendarEventDetails></t:CalendarEvent>" }) -join ''
                "<m:FreeBusyResponse><m:ResponseMessage ResponseClass=""Success""><m:ResponseCode>NoError</m:ResponseCode></m:ResponseMessage><m:FreeBusyView><t:FreeBusyViewType>DetailedMerged</t:FreeBusyViewType><t:MergedFreeBusy>$($plan.View)</t:MergedFreeBusy><t:CalendarEventArray>$items</t:CalendarEventArray><t:WorkingHours><t:TimeZone><t:Bias>-60</t:Bias><t:StandardTime><t:Bias>0</t:Bias><t:Time>03:00:00</t:Time><t:DayOrder>5</t:DayOrder><t:Month>10</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek></t:StandardTime><t:DaylightTime><t:Bias>-60</t:Bias><t:Time>02:00:00</t:Time><t:DayOrder>5</t:DayOrder><t:Month>3</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek></t:DaylightTime></t:TimeZone><t:WorkingPeriodArray><t:WorkingPeriod><t:DayOfWeek>Monday Tuesday Wednesday Thursday Friday</t:DayOfWeek><t:StartTimeInMinutes>480</t:StartTimeInMinutes><t:EndTimeInMinutes>1020</t:EndTimeInMinutes></t:WorkingPeriod></t:WorkingPeriodArray></t:WorkingHours></m:FreeBusyView></m:FreeBusyResponse>"
            }
            return & $ok "<GetUserAvailabilityResponse xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages"" $($script:SimNs)><m:FreeBusyResponseArray>$($responses -join '')</m:FreeBusyResponseArray></GetUserAvailabilityResponse>"
        }
    }
    return Get-SimFault 'ErrorInvalidOperation' "Operation $Operation not simulated."
}

function Get-SimHttpResponse {
    <# Answer to one HTTP request of the tool. Every call is recorded in State.Calls. #>
    param([Parameter(Mandatory = $true)][Net.Http.HttpRequestMessage]$Request, [Parameter(Mandatory = $true)][hashtable]$State)

    $uri = $Request.RequestUri
    $body = if ($Request.Content) { [Text.Encoding]::UTF8.GetString($Request.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()) } else { '' }
    $auth = $Request.Headers.Authorization
    $operation = [regex]::Match($body, '<soap:Body>\s*<m:(\w+?)(?:Request)?[\s>]').Groups[1].Value
    $State.Calls.Add([pscustomobject]@{ Method = $Request.Method.Method; Url = $uri.AbsoluteUri; Operation = $operation; Scheme = if ($auth) { $auth.Scheme } else { $null }; Parameter = if ($auth) { $auth.Parameter } else { $null }; Body = $body })
    $json = { param([int]$Code, $Object) New-SimResponse $Code ($Object | ConvertTo-Json -Depth 5 -Compress) -ContentType 'application/json' }
    $path = $uri.AbsolutePath

    # ---- Autodiscover ----
    if ($path -match '/autodiscover/autodiscover\.json/v1\.0/') {
        if ($State.AutodiscoverV2 -and $uri.Host -like 'autodiscover.*') { return & $json 200 @{ Protocol = 'EWS'; Url = $State.AutodiscoverV2 } }
        return New-SimResponse 404 ''
    }
    if ($path -match '/autodiscover/autodiscover\.xml') {
        if ($Request.Method.Method -eq 'GET') { return New-SimResponse 404 '' }
        if (-not $State.AutodiscoverPox -or -not $auth) { return New-SimResponse 401 '' -Challenges @('Negotiate', 'NTLM', 'Basic realm="autodiscover"') }
        if ($auth.Scheme -in 'Negotiate', 'NTLM' -and ([Convert]::FromBase64String($auth.Parameter))[8] -eq 1) { return New-SimResponse 401 '' -Challenges @("$($auth.Scheme) $(New-SimNtlmChallenge)") }
        return New-SimResponse 200 "<?xml version=""1.0"" encoding=""utf-8""?><Autodiscover xmlns=""http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006""><Response xmlns=""http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a""><Account><Protocol><Type>EXCH</Type><EwsUrl>https://internal.contoso.test/EWS/Exchange.asmx</EwsUrl></Protocol><Protocol><Type>EXPR</Type><EwsUrl>https://mail.contoso.test/EWS/Exchange.asmx</EwsUrl></Protocol></Account></Response></Autodiscover>"
    }

    # ---- AD FS and Entra ID ----
    if ($path -match '\.well-known/openid-configuration$') {
        if ($uri.Host -eq 'login.microsoftonline.com') { return & $json 200 @{ issuer = "https://login.microsoftonline.com/$($script:SimTenantId)/v2.0"; token_endpoint = 'x' } }
        return & $json 200 @{ issuer = 'https://adfs.contoso.test/adfs'; token_endpoint = 'https://adfs.contoso.test/adfs/oauth2/token'; device_authorization_endpoint = 'https://adfs.contoso.test/adfs/oauth2/devicecode' }
    }
    if ($path -match '/userrealm/') { return & $json 200 @{ NameSpaceType = 'Managed'; DomainName = 'contoso.test' } }
    if ($path -match '/devicecode$') {
        $State.DeviceScope = [Uri]::UnescapeDataString(([regex]::Match($body, '(?:^|&)scope=([^&]*)').Groups[1].Value).Replace('+', ' '))
        return & $json 200 @{ device_code = 'DC-123'; user_code = 'ABCD-EFGH'; verification_uri = 'https://microsoft.com/devicelogin'; expires_in = 900; interval = 5; message = 'Open the page and enter the code ABCD-EFGH.' } }
    if ($path -match '/oauth2/(v2\.0/)?authorize$') { return New-SimResponse 200 '<html><title>Sign In</title><form id="loginForm"><input id="userNameInput"/></form></html>' -ContentType 'text/html' }
    if ($path -match '/oauth2/(v2\.0/)?token$') {
        $fields = @{}; foreach ($p in $body.Split('&')) { $k, $v = $p.Split('=', 2); $fields[[Uri]::UnescapeDataString($k)] = [Uri]::UnescapeDataString(([string]$v).Replace('+', ' ')) }
        $tenant = if ($uri.Host -eq 'login.microsoftonline.com') { $script:SimTenantId } else { $null }
        $scope = if ($fields.scope) { [string]$fields.scope } elseif ($State.ContainsKey('DeviceScope')) { [string]$State.DeviceScope } else { '' }
        $graph = $scope -like 'https://graph.microsoft.com/*'
        $audience = if ($graph) { 'https://graph.microsoft.com' } else { $State.TokenAudience }
        if ($fields.grant_type -eq 'client_credentials') {
            if ($fields.client_secret -and $fields.client_secret -ne 'Sim-Secret!') { return & $json 401 @{ error = 'invalid_client'; error_description = 'AADSTS7000215: Invalid client secret provided.' } }
            $token = New-SimToken -Audience $audience -Scope '' -Roles $(if ($graph) { $State.GraphRoles } else { $State.AppRoles }) -AppId $fields.client_id -TenantId $tenant
        }
        else { $token = New-SimToken -Audience $audience -Scope $(if ($graph) { $State.GraphScopes } else { 'EWS.AccessAsUser.All' }) -AppId $fields.client_id -TenantId $tenant }
        $State.ValidTokens.Add($token)
        return & $json 200 @{ access_token = $token; token_type = 'Bearer'; expires_in = 3600 }
    }

    # ---- Microsoft Graph ----
    if ($uri.Host -eq 'graph.microsoft.com') { return Get-SimGraphResponse -State $State -Request $Request -Body $body }

    # ---- EWS ----
    if ($path -match '/EWS/Exchange\.asmx$') {
        $headers = @{ 'X-FEServer' = 'EXCH01'; 'X-BEServer' = 'EXCH02'; 'request-id' = [guid]::NewGuid().ToString() }
        $cid = $null; if ($Request.Headers.TryGetValues('client-request-id', [ref]$cid)) { $headers['client-request-id'] = @($cid)[0] }
        $refuse = { param([string[]]$Extra) New-SimResponse 401 '' $headers (@($State.Schemes) + @($Extra)) }
        $bearer = "Bearer client_id=""00000002-0000-0ff1-ce00-000000000000"", trusted_issuers=""$(if ($State.TrustedIssuers) { $State.TrustedIssuers } else { '00000001-0000-0000-c000-000000000000@*' })"", authorization_uri=""$($State.AuthorizationUri)"""
        # A connection authenticated by a Windows handshake stays authenticated (IIS, connection-based NTLM).
        if (-not $auth -and -not $State.WindowsConnection) { return & $refuse @() }
        if ($auth.Scheme -eq 'Bearer' -and -not $auth.Parameter) { return New-SimResponse 401 '' $headers @(@($State.Schemes) + $bearer) }
        if ($auth.Scheme -eq 'Bearer' -and -not $State.ValidTokens.Contains($auth.Parameter)) { $headers['x-ms-diagnostics'] = '2000001;reason="The token is invalid.";error_category="invalid_token"'; return & $refuse @($bearer) }
        if ($auth.Scheme -eq 'Basic') {
            $pair = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($auth.Parameter)).Split(':', 2)
            if (-not $State.BasicUsers.ContainsKey($pair[0]) -or $State.BasicUsers[$pair[0]] -ne $pair[1]) { return & $refuse @() }
        }
        if ($auth.Scheme -in 'Negotiate', 'NTLM') {
            $raw = [Convert]::FromBase64String($auth.Parameter)
            if ($raw.Length -gt 8 -and $raw[8] -eq 1) { return New-SimResponse 401 '' $headers @("$($auth.Scheme) $(New-SimNtlmChallenge)") }
            if (-not $State.WindowsAccepted) { return & $refuse @() }
            $State.WindowsConnection = $true
        }
        if ($State.EwsBlocked) { $headers['X-EWS-Policy-Reason'] = 'EWS is blocked by policy for this user or tenant'; return New-SimResponse 403 '' $headers }
        if (-not $State.ContainsKey('AffinitySet')) { $State.AffinitySet = $true; $headers['Set-Cookie'] = 'X-BackEndOverrideCookie=EXCH02.contoso.test~1942062522; path=/EWS; secure' }
        if ($State.ServerBusyOnce) { $State.ServerBusyOnce = $false; return New-SimResponse 500 (Get-SimFault 'ErrorServerBusy' 'The server cannot service this request right now.' 1000) $headers }
        if ($State.DenyImpersonation -and $body -match '<t:ExchangeImpersonation>') { return New-SimResponse 500 (Get-SimFault 'ErrorImpersonateUserDenied' 'The account does not have permission to impersonate the requested user.') $headers }
        if ($State.DenyDelegate -and $body -match '<t:DistinguishedFolderId Id="\w+"><t:Mailbox>') { return New-SimResponse 200 (Get-SimSoap $State "<m:GetFolderResponse $($script:SimNs)><m:ResponseMessages><m:GetFolderResponseMessage ResponseClass=""Error""><m:MessageText>Access is denied.</m:MessageText><m:ResponseCode>ErrorAccessDenied</m:ResponseCode></m:GetFolderResponseMessage></m:ResponseMessages></m:GetFolderResponse>") $headers }
        return New-SimResponse 200 (Get-SimEwsAnswer $State $operation $body) $headers
    }
    if ($uri.Scheme -eq 'http') { return New-SimResponse 404 '' }
    return New-SimResponse 404 '' -ContentType 'text/html'
}

function Get-SimGraphResponse {
    <# Microsoft Graph v1.0 for the mailbox of the state: mailFolders, messages, sendMail, reply, move, delete, getSchedule. #>
    param([hashtable]$State, [Net.Http.HttpRequestMessage]$Request, [string]$Body)

    $uri = $Request.RequestUri
    $path = [Uri]::UnescapeDataString($uri.AbsolutePath)
    $query = [Uri]::UnescapeDataString($uri.Query)
    $auth = $Request.Headers.Authorization
    $method = $Request.Method.Method
    $headers = @{ 'request-id' = [guid]::NewGuid().ToString(); 'x-ms-ags-diagnostic' = '{"ServerInfo":{"DataCenter":"France Central","Slice":"E","Ring":"3","ScaleUnit":"001","RoleInstance":"PA1PEPF0000SIM"}}' }
    $cid = $null; if ($Request.Headers.TryGetValues('client-request-id', [ref]$cid)) { $headers['client-request-id'] = @($cid)[0] }
    $send = { param([int]$Code, $Object) $h = $headers.Clone(); if ($null -eq $Object) { New-SimResponse $Code '' $h } else { New-SimResponse $Code ($Object | ConvertTo-Json -Depth 8 -Compress) $h -ContentType 'application/json' } }
    $fail = { param([int]$Code, [string]$Err, [string]$Message) & $send $Code @{ error = @{ code = $Err; message = $Message } } }
    if (-not $auth) { $h = $headers.Clone(); return New-SimResponse 401 '{"error":{"code":"InvalidAuthenticationToken","message":"Access token is empty."}}' $h @('Bearer realm="", authorization_uri="https://login.microsoftonline.com/common/oauth2/authorize", client_id="00000003-0000-0000-c000-000000000000"') -ContentType 'application/json' }
    if (-not $State.ValidTokens.Contains($auth.Parameter)) { return & $fail 401 'InvalidAuthenticationToken' 'Access token validation failure. Invalid audience.' }
    if ($State.GraphForbidden) { return & $fail 403 'ErrorAccessDenied' 'Access is denied. Check credentials and try again.' }
    if ($path -notmatch '^/v1\.0/users/([^/]+)(/.*)?$') { return & $fail 400 'BadRequest' 'Unsupported.' }
    $rest = [string]$Matches[2]
    $json = if ($Body) { $Body | ConvertFrom-Json -AsHashtable } else { @{} }
    $folderJson = { param($f) @{ id = $f.Id; displayName = $f.Name; parentFolderId = $f.Parent; childFolderCount = @($State.Folders | Where-Object { $_.Parent -eq $f.Id }).Count; totalItemCount = $f.Total; unreadItemCount = $f.Unread } }
    $messageJson = { param($m, [switch]$Full) $o = @{ id = $m.Id; subject = $m.Subject; receivedDateTime = $m.Received; isRead = ($m.Read -eq 'true'); hasAttachments = $false; internetMessageId = "<$($m.Id)@contoso.test>"; changeKey = "CK-$($m.Id)"; from = @{ emailAddress = @{ name = $m.From; address = $m.From } } }; if ($Full) { $o.body = @{ contentType = 'text'; content = $m.Body }; $o.toRecipients = @(@{ emailAddress = @{ address = $State.Mailbox } }) }; $o }
    switch -Regex ("$method $rest") {
        '^GET /mailFolders/inbox$' { $f = $State.Folders[0]; return & $send 200 @{ id = $f.Id; displayName = 'Inbox'; totalItemCount = $State.Inbox.Count; unreadItemCount = 1; childFolderCount = @($State.Folders | Where-Object { $_.Parent -eq 'F-INBOX' }).Count } }
        '^GET /mailFolders$' { return & $send 200 @{ value = @($State.Folders | Where-Object { $_.Parent -eq 'F-ROOT' } | ForEach-Object { & $folderJson $_ }) } }
        '^GET /mailFolders/(inbox|[^/]+)/childFolders$' {
            $parent = if ($Matches[1] -eq 'inbox') { 'F-INBOX' } else { $Matches[1] }
            $list = @($State.Folders | Where-Object { $_.Parent -eq $parent })
            $name = [regex]::Match($query, "displayName eq '((?:[^']|'')*)'").Groups[1].Value
            if ($name) { $list = @($list | Where-Object { $_.Name -eq $name.Replace("''", "'") }) }
            return & $send 200 @{ value = @($list | ForEach-Object { & $folderJson $_ }) }
        }
        '^POST /mailFolders/(inbox|[^/]+)/childFolders$' { $parent = if ($Matches[1] -eq 'inbox') { 'F-INBOX' } else { $Matches[1] }; $id = "F-$($State.Next)"; $State.Next++; $f = @{ Id = $id; Parent = $parent; Name = $json.displayName; Class = 'IPF.Note'; Total = 0; Unread = 0; Children = 0 }; $State.Folders.Add($f); return & $send 201 (& $folderJson $f) }
        '^GET /mailFolders/([^/]+)/messages$' {
            $list = @($State.Inbox)
            $exact = [regex]::Match($query, "subject eq '((?:[^']|'')*)'").Groups[1].Value
            $search = [regex]::Match($query, '\$search="(?:subject:)?([^"]*)"').Groups[1].Value
            if ($exact) { $exact = $exact.Replace("''", "'"); $list = @($list | Where-Object { $_.Subject -eq $exact }) }
            elseif ($search) { $list = @($list | Where-Object { $_.Subject.IndexOf($search, [StringComparison]::OrdinalIgnoreCase) -ge 0 }) }
            $top = [int][regex]::Match($query, '\$top=(\d+)').Groups[1].Value
            return & $send 200 @{ value = @($list | Sort-Object { $_.Received } -Descending | Select-Object -First $top | ForEach-Object { & $messageJson $_ }) }
        }
        '^GET /messages/([^/]+)$' { $m = $State.Inbox | Where-Object { $_.Id -eq $Matches[1] } | Select-Object -First 1; if (-not $m) { return & $fail 404 'ErrorItemNotFound' 'The specified object was not found in the store.' }; return & $send 200 (& $messageJson $m -Full) }
        '^POST /sendMail$' {
            $to = $json.message.toRecipients[0].emailAddress.address
            if ($to -ieq $State.Mailbox) { $State.Inbox.Add(@{ Id = "M-$($State.Next)"; Subject = $json.message.subject; From = $State.Mailbox; Received = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'); Read = 'false'; Body = $json.message.body.content }); $State.Next++ }
            return & $send 202 $null
        }
        '^POST /messages/([^/]+)/reply$' { if (-not ($State.Inbox | Where-Object { $_.Id -eq $Matches[1] })) { return & $fail 404 'ErrorItemNotFound' 'Not found.' }; return & $send 202 $null }
        '^POST /messages/([^/]+)/move$' {
            $m = $State.Inbox | Where-Object { $_.Id -eq $Matches[1] } | Select-Object -First 1
            if (-not $m) { return & $fail 404 'ErrorItemNotFound' 'Not found.' }
            $m.Id = "$($m.Id)-MOVED"; return & $send 201 (& $messageJson $m)
        }
        '^(DELETE /messages/([^/]+)|POST /messages/([^/]+)/permanentDelete)$' {
            $id = if ($Matches[2]) { $Matches[2] } else { $Matches[3] }
            $m = $State.Inbox | Where-Object { $_.Id -eq $id } | Select-Object -First 1
            if ($m) { [void]$State.Inbox.Remove($m) }
            return & $send 204 $null
        }
        '^GET /calendarView$' { return & $send 200 @{ value = @($State.Calendar | ForEach-Object { @{ id = $_.Id; subject = $_.Subject } }) } }
        '^POST /events$' { $State.Calendar.Add(@{ Id = "C-$($State.Next)"; Subject = $json.subject; Start = $json.start.dateTime; Status = $json.showAs }); $State.Next++; return & $send 201 @{ id = "C-$($State.Next - 1)" } }
        '^DELETE /events/([^/]+)$' { $c = $State.Calendar | Where-Object { $_.Id -eq $Matches[1] } | Select-Object -First 1; if ($c) { [void]$State.Calendar.Remove($c) }; return & $send 204 $null }
        '^DELETE /mailFolders/([^/]+)$' { $id = $Matches[1]; foreach ($f in @($State.Folders | Where-Object { $_.Id -eq $id -or $_.Parent -eq $id })) { [void]$State.Folders.Remove($f) }; return & $send 204 $null }
        '^POST /calendar/getSchedule$' {
            $fbStart = [datetime]::SpecifyKind([datetime]$json.startTime.dateTime, 'Utc'); $fbEnd = [datetime]::SpecifyKind([datetime]$json.endTime.dateTime, 'Utc')
            $values = foreach ($a in @($json.schedules)) {
                $code = if ($State.FreeBusy.ContainsKey($a)) { $State.FreeBusy[$a] } else { 'Success' }
                if ($code -ne 'Success') { @{ scheduleId = $a; availabilityView = ''; error = @{ message = "No mailbox for $a."; responseCode = $code } }; continue }
                $plan = Get-SimSchedule -Mailbox $a -StartUtc $fbStart -Days ([int]($fbEnd - $fbStart).TotalDays) -IntervalMinutes ([int]$json.availabilityViewInterval)
                $status = @{ Busy = 'busy'; Tentative = 'tentative'; OOF = 'oof'; WorkingElsewhere = 'workingElsewhere'; Free = 'free' }
                @{ scheduleId = $a; availabilityView = $plan.View; scheduleItems = @($plan.Events | ForEach-Object { @{ isPrivate = $false; status = $status[$_.BusyType]; subject = $_.Subject; location = $_.Location; start = @{ dateTime = $_.Start.ToString('yyyy-MM-ddTHH:mm:ss.0000000'); timeZone = 'UTC' }; end = @{ dateTime = $_.End.ToString('yyyy-MM-ddTHH:mm:ss.0000000'); timeZone = 'UTC' } } })
                   workingHours = @{ daysOfWeek = @('monday', 'tuesday', 'wednesday', 'thursday', 'friday'); startTime = '08:00:00.0000000'; endTime = '17:00:00.0000000'; timeZone = @{ name = 'Romance Standard Time' } } }
            }
            return & $send 200 @{ value = @($values) }
        }
    }
    return & $fail 400 'BadRequest' "Not simulated: $method $rest"
}

function Get-SimCertificate([string]$HostName, [int]$Port, [int]$Days) {
    [pscustomobject]@{ HostName = $HostName; Port = $Port; Reachable = $true; Valid = $true; Interrupted = $false; Subject = "CN=$HostName"; Issuer = 'CN=Contoso Test CA'; NotAfterUtc = [DateTime]::UtcNow.AddDays($Days).ToString('yyyy-MM-ddTHH:mm:ssZ'); DaysLeft = $Days; Protocol = 'Tls13'; Error = $null }
}

function Get-SimWindowRedirect([string]$Url, [string]$RedirectUri, [hashtable]$State) {
    $returned = [regex]::Match($Url, '[?&]state=([^&]+)').Groups[1].Value
    "$RedirectUri`?code=SIM-CODE-123&state=$returned"
}

function Install-SimExchange {
    <#
        Replaces the network inside the loaded module: Send-WscHttpRequest answers with Get-SimHttpResponse,
        the TLS check with Get-SimCertificate, the sign-in window with Get-SimWindowRedirect, the Kerberos
        check finds no KDC, the waits are short. Remove-Module restores them.
    #>
    param([Parameter(Mandatory = $true)][psmoduleinfo]$Module, [Parameter(Mandatory = $true)][hashtable]$State, [switch]$Slow)

    $responder = ${function:Get-SimHttpResponse}
    $certificate = ${function:Get-SimCertificate}
    $window = ${function:Get-SimWindowRedirect}
    . $Module {
        param($State, $Responder, $Certificate, $Window, $Slow)
        $rsa = [Security.Cryptography.RSA]::Create(2048)
        $script:SimServerCertificate = [Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=mail.contoso.test', $rsa, 'SHA256', [Security.Cryptography.RSASignaturePadding]::Pkcs1).CreateSelfSigned([DateTimeOffset]::UtcNow.AddDays(-1), [DateTimeOffset]::UtcNow.AddDays(365))
        $script:SimState = $State; $script:SimResponder = $Responder; $script:SimCertificate = $Certificate; $script:SimWindow = $Window; $script:SimSlow = [bool]$Slow
        function script:Send-WscHttpRequest { param($HttpClient, $Request) if ($script:SimSlow) { Start-Sleep -Milliseconds (Get-Random -Minimum 40 -Maximum 160) }; & $script:SimResponder -Request $Request -State $script:SimState }
        function script:Get-WscTlsCertificate { param([string]$HostName, [int]$Port = 443, [int]$TimeoutSeconds) & $script:SimCertificate $HostName $Port $script:SimState.CertificateDays }
        function script:Start-Process { }
        function script:Wait-WscSeconds { param([double]$Seconds) }
        function script:Test-WscKerberosRealm { param([string]$Domain, [int]$TimeoutSeconds) [pscustomobject]@{ Domain = $Domain; Kdc = $null; Port88 = $false; Error = "no SRV record _kerberos._tcp.$Domain"; DomainJoined = $false; ComputerDomain = 'WORKGROUP' } }
        function script:Get-WscAutodiscoverSrvHost { param([string]$Domain) $null }
        function script:Get-WscTlsServerCertificate { param([string]$HostName, [int]$Port, [int]$TimeoutSeconds) $script:SimServerCertificate }
        function script:Test-WscDesktopSession { $true }
        function script:Find-WscBrowser { param([switch]$All) [pscustomobject]@{ Name = 'Microsoft Edge'; Path = 'msedge.exe' } }
        function script:Test-WscBrowserPolicyBlock { $false }
        function script:Invoke-WscBrowserAuthorization { param($Browser, $Url, $RedirectUri, $TimeoutSeconds) & $script:SimWindow -Url $Url -RedirectUri $RedirectUri -State $script:SimState }
    } $State $responder $certificate $window $Slow
}
