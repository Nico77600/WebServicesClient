<#
.SYNOPSIS
    Web Services Client for Exchange - the EWS operations of the toolbox (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    Endpoint      GetFolder on the Inbox: Exchange version, servers, routing and affinity, access to the mailbox
    Folders       FindFolder (Deep) from the root of the mailbox: the whole folder tree with paths
    ReadMail      FindItem (most recent Inbox messages), then GetItem of one message (body preview)
    FreeBusy      GetUserAvailability for one or several mailboxes (DetailedMerged)
    CreateFolder  the test folder under the Inbox (found again when it exists)
    SendMail      CreateItem SendAndSaveCopy, then waits for the message in the Inbox
    ReplyMail     ReplyToItem on the chosen message
    MoveMail      MoveItem of the chosen message to the test folder
    DeleteMail    DeleteItem of the chosen message (Test.DeleteMode)
    The writing operations need AllowWrite. Reply, move and delete work on the message named with
    -ItemId or -ItemSubject, else on the test message of this run or the last one of the tool
    (subject starting with [Web Services Client for Exchange]): never on an arbitrary message.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:WriteStages = @('CreateFolder', 'SendMail', 'ReplyMail', 'MoveMail', 'DeleteMail', 'SeedData', 'CleanData')

function Get-WscExchangeProduct {
    <# Product name of a ServerVersionInfo: Exchange Online, Exchange Server SE, 2019, 2016, 2013. #>
    param([AllowNull()][pscustomobject]$Version)
    if (-not $Version -or -not $Version.Major) { return $null }
    $major = [int]$Version.Major; $minor = [int]$Version.Minor; $build = [int]$Version.Build
    if ($major -eq 15 -and $minor -ge 20) { return 'Exchange Online' }
    if ($major -eq 15 -and $minor -eq 2) { if ($build -ge 2562) { return 'Exchange Server SE' } else { return 'Exchange Server 2019' } }
    if ($major -eq 15 -and $minor -eq 1) { return 'Exchange Server 2016' }
    if ($major -eq 15 -and $minor -eq 0) { return 'Exchange Server 2013' }
    return "Exchange $major.$minor"
}

function Get-WscAccessText {
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    $mailbox = [string]$Context.Config.Mailbox
    switch ($Context.Access) {
        'Delegate' { "delegate access to $mailbox" }
        'Impersonation' { "impersonation of $mailbox" }
        default { "the own mailbox $mailbox" }
    }
}

function Add-WscAction {
    <# A change made to the mailbox, listed in the report (Actions). #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Action, [string]$Target, [string]$Result, [string]$ItemId)
    $Context.Actions.Add([pscustomobject]@{ TimeUtc = [DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'); Action = $Action; Target = $Target; Result = $Result; ItemId = $ItemId })
}

function Invoke-WscOperationStage {
    <# Runs one operation stage; the writing ones are blocked without AllowWrite. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage, [Parameter(Mandatory = $true)][string]$Title)

    if ($Stage -in $script:WriteStages -and -not $Context.Config.AllowWrite) {
        Add-WscStep $Context $Stage $Title Blocked "This step changes the mailbox $($Context.Config.Mailbox): it runs only with -AllowWrite (Test.AllowWrite), on a test mailbox." ([ordered]@{ AllowWrite = $false })
        $Context.Stop = $true
        return
    }
    if ($Stage -eq 'SeedData') { if ($Context.Protocol -eq 'Graph') { Invoke-WscSeedGraph -Context $Context } else { Invoke-WscSeedEws -Context $Context }; return }
    if ($Stage -eq 'CleanData') { if ($Context.Protocol -eq 'Graph') { Invoke-WscCleanGraph -Context $Context } else { Invoke-WscCleanEws -Context $Context }; return }
    if ($Context.Protocol -eq 'Graph') { Invoke-WscGraphOperation -Context $Context -Stage $Stage; return }
    switch ($Stage) {
        'Endpoint' { Invoke-WscStageEndpoint -Context $Context }
        'Folders' { Invoke-WscStageFolders -Context $Context }
        'ReadMail' { Invoke-WscStageReadMail -Context $Context }
        'FreeBusy' { Invoke-WscStageFreeBusy -Context $Context }
        'CreateFolder' { Invoke-WscStageCreateFolder -Context $Context }
        'SendMail' { Invoke-WscStageSendMail -Context $Context }
        'ReplyMail' { Invoke-WscStageReplyMail -Context $Context }
        'MoveMail' { Invoke-WscStageMoveMail -Context $Context }
        'DeleteMail' { Invoke-WscStageDeleteMail -Context $Context }
    }
}

function Invoke-WscStageEndpoint {
    <# First EWS request on the mailbox: GetFolder of the Inbox, with the access mode of the run. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Endpoint'
    $answer = Invoke-WscEws -Context $Context -Operation 'GetFolder' -Body (New-WscGetFolderBody -Context $Context -Distinguished 'inbox')
    $access = Get-WscAccessText -Context $Context
    $details = Get-WscEwsDetails -Answer $answer -More ([ordered]@{ Access = $Context.Access; AnchorMailbox = Get-WscAnchorMailbox -Context $Context; RequestServerVersion = $Context.Config.RequestServerVersion })
    if ($answer.HttpStatus -ne 200 -or $answer.ResponseClass -ne 'Success') {
        Add-WscStep $Context $stage 'GetFolder (Inbox)' Failed "EWS did not open the Inbox ($access): $(Get-WscEwsFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $true
        return
    }
    $node = $answer.Xml.SelectSingleNode('//m:Folders/*', $answer.Ns)
    $inbox = if ($node) { ConvertFrom-WscFolderNode -Node $node -Ns $answer.Ns } else { $null }
    $product = Get-WscExchangeProduct -Version $answer.ServerVersion
    if ($answer.ServerVersion) { $details.ExchangeVersion = "$product $($answer.ServerVersion.Text)"; $details.Schema = $answer.ServerVersion.Schema }
    if ($inbox) { $details.Inbox = "$($inbox.DisplayName): $($inbox.TotalCount) item(s), $($inbox.UnreadCount) unread"; $Context.Inbox = $inbox }
    $Context.Mailbox = [pscustomobject]@{ Address = [string]$Context.Config.Mailbox; Access = $Context.Access; Inbox = $inbox }
    $accepted = switch ([string]$Context.Config.Authentication) { 'Basic' { 'the user name and password' } 'Windows' { "the Windows authentication ($($Context.WindowsPackageUsed))" } default { 'the token' } }
    $version = if ($product) { " $product $($answer.ServerVersion.Text)," } else { '' }
    $route = if ($answer.Server) { " $($answer.Server)." } else { '' }
    Add-WscStep $Context $stage 'GetFolder (Inbox)' Passed "EWS accepted $accepted and opened the Inbox through $($access):$version $($details.Inbox).$route" $details

    $cookies = @()
    if ($script:WscCookies) { $cookies = @($script:WscCookies.GetCookies([Uri]$Context.Endpoints.EwsUrl) | Where-Object { $_.Name -in $script:WscAffinityCookies } | ForEach-Object { $_.Name }) }
    $routing = [ordered]@{ AnchorMailbox = $details.AnchorMailbox; PreferServerAffinity = [bool]$Context.Config.PreferServerAffinity; AffinityCookies = $cookies -join ', '; Servers = $answer.Server; RequestId = $answer.RequestId; ClientRequestId = Get-WscHeader $answer.Response 'client-request-id' }
    $echo = if ($routing.ClientRequestId) { 'Exchange echoes client-request-id (return-client-request-id): the request can be found in its logs.' } else { 'Exchange did not echo client-request-id.' }
    if ($cookies.Count) {
        Add-WscStep $Context $stage 'Routing and affinity' Passed "X-AnchorMailbox routes to the mailbox; Exchange returned the affinity cookie $($cookies -join ', '), sent back with every next request. $echo" $routing
    }
    else {
        Add-WscStep $Context $stage 'Routing and affinity' Passed "X-AnchorMailbox routes every request to the mailbox$(if ($answer.Server) { " ($($answer.Server))" }); no affinity cookie was needed. $echo" $routing
    }
}

function Set-WscFolderPaths {
    <#
        Path and depth of every folder from the parent IDs. Folder IDs are case-sensitive (base64): an
        ordinal dictionary, never a PowerShell hashtable, which ignores the case and mixes two folders.
    #>
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Folders)

    $byId = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
    foreach ($f in $Folders) { if ($f.FolderId -and -not $byId.ContainsKey($f.FolderId)) { $byId[$f.FolderId] = $f } }
    foreach ($f in $Folders) {
        $names = [Collections.Generic.List[string]]::new(); $names.Add([string]$f.DisplayName)
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal); [void]$seen.Add([string]$f.FolderId)
        $p = [string]$f.ParentFolderId
        while ($p -and $byId.ContainsKey($p) -and $seen.Add($p)) { $names.Insert(0, [string]$byId[$p].DisplayName); $p = [string]$byId[$p].ParentFolderId }
        $f.Path = '\' + ($names -join '\')
        $f | Add-Member -NotePropertyName Depth -NotePropertyValue ($names.Count - 1) -Force
    }
    # Tree order: a folder right after its parent, siblings by name.
    return @($Folders | Sort-Object { ($_.Path.Split('\') | ForEach-Object { $_.ToLowerInvariant() }) -join [char]1 })
}
function Invoke-WscStageFolders {
    <# The folder tree of the mailbox (FindFolder Deep from msgfolderroot, pages of 500), with paths. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Folders'
    $all = [Collections.Generic.List[object]]::new()
    $offset = 0
    $answer = $null
    for ($page = 0; $page -lt 10; $page++) {
        $answer = Invoke-WscEws -Context $Context -Operation 'FindFolder' -Body (New-WscFindFolderBody -Context $Context -Traversal Deep -Offset $offset -Max 500)
        if ($answer.HttpStatus -ne 200 -or $answer.ResponseClass -ne 'Success') {
            Add-WscStep $Context $stage 'FindFolder' Failed "The folders were not listed: $(Get-WscEwsFailureText -Context $Context -Answer $answer)" (Get-WscEwsDetails -Answer $answer)
            $Context.Stop = $true
            return
        }
        foreach ($n in $answer.Xml.SelectNodes('//t:Folders/*', $answer.Ns)) { $all.Add((ConvertFrom-WscFolderNode -Node $n -Ns $answer.Ns)) }
        $root = $answer.Xml.SelectSingleNode('//m:RootFolder', $answer.Ns)
        if (-not $root -or $root.GetAttribute('IncludesLastItemInRange') -ne 'false') { break }
        $offset = [int]$root.GetAttribute('IndexedPagingOffset')
    }
    $Context.Folders = Set-WscFolderPaths -Folders @($all)
    $mail = @($all | Where-Object { [string]$_.FolderClass -eq 'IPF.Note' -or -not $_.FolderClass }).Count
    $items = ($all | ForEach-Object { [int]$_.TotalCount } | Measure-Object -Sum).Sum
    $details = Get-WscEwsDetails -Answer $answer -More ([ordered]@{ Folders = $all.Count; MailFolders = $mail; Items = $items; Pages = $page + 1 })
    Add-WscStep $Context $stage 'FindFolder' Passed "$($all.Count) folder(s) in the mailbox ($mail mail folder(s), $items item(s) in total): Folders tab of the report." $details
}

function Find-WscMessages {
    <# FindItem in a folder of the mailbox; returns the answer and the messages read. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [string]$Folder = 'inbox', [string]$FolderId, [int]$Max = 10, [string]$SubjectContains, [string]$Collapse)
    $answer = Invoke-WscEws -Context $Context -Operation 'FindItem' -Body (New-WscFindItemBody -Context $Context -Folder $Folder -FolderId $FolderId -Max $Max -SubjectContains $SubjectContains) -Collapse $Collapse
    $items = @()
    if ($answer.HttpStatus -eq 200 -and $answer.ResponseClass -eq 'Success') {
        $items = @(foreach ($n in $answer.Xml.SelectNodes('//t:Items/*', $answer.Ns)) { ConvertFrom-WscItemNode -Node $n -Ns $answer.Ns })
    }
    [pscustomobject]@{ Answer = $answer; Items = $items }
}

function Invoke-WscStageReadMail {
    <# The most recent Inbox messages (headers), then one message in full: recipients and a preview of the body. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'ReadMail'
    $count = [int]$Context.Config.MessageCount
    $found = Find-WscMessages -Context $Context -Max $count
    $answer = $found.Answer
    if ($answer.HttpStatus -ne 200 -or $answer.ResponseClass -ne 'Success') {
        Add-WscStep $Context $stage 'FindItem (Inbox)' Failed "The Inbox messages were not listed: $(Get-WscEwsFailureText -Context $Context -Answer $answer)" (Get-WscEwsDetails -Answer $answer)
        $Context.Stop = $true
        return
    }
    $Context.Messages = @($found.Items | Select-Object DateTimeReceived, From, Subject, IsRead, HasAttachments, Size, ItemClass, InternetMessageId, ItemId)
    $root = $answer.Xml.SelectSingleNode('//m:RootFolder', $answer.Ns)
    $total = if ($root) { $root.GetAttribute('TotalItemsInView') } else { $null }
    $details = Get-WscEwsDetails -Answer $answer -More ([ordered]@{ Listed = $found.Items.Count; InInbox = $total })
    if (-not $found.Items.Count) {
        Add-WscStep $Context $stage 'FindItem (Inbox)' Passed 'The Inbox is empty: nothing to read.' $details
        return
    }
    Add-WscStep $Context $stage 'FindItem (Inbox)' Passed "$($found.Items.Count) most recent message(s) of $total in the Inbox: date, sender, subject (Messages tab)." $details

    $pick = $found.Items | Where-Object { [string]$_.ItemClass -like 'IPM.Note*' } | Select-Object -First 1
    if (-not $pick) { $pick = $found.Items[0] }
    $get = Invoke-WscEws -Context $Context -Operation 'GetItem' -Body (New-WscGetItemBody -ItemId $pick.ItemId)
    $d = Get-WscEwsDetails -Answer $get
    if ($get.HttpStatus -ne 200 -or $get.ResponseClass -ne 'Success') {
        Add-WscStep $Context $stage 'GetItem' Failed "The message '$($pick.Subject)' was not read: $(Get-WscEwsFailureText -Context $Context -Answer $get)" $d
        return
    }
    $node = $get.Xml.SelectSingleNode('//m:Items/*', $get.Ns)
    $item = ConvertFrom-WscItemNode -Node $node -Ns $get.Ns
    $body = ([string]$item.Body -replace '\s+', ' ').Trim()
    $preview = if ($body.Length -gt 300) { $body.Substring(0, 300) + '...' } else { $body }
    $Context.Message = [pscustomobject]@{ Subject = $item.Subject; From = $item.From; Sender = $item.Sender; To = $item.To; Cc = $item.Cc; DateTimeReceived = $item.DateTimeReceived; InternetMessageId = $item.InternetMessageId; Size = $item.Size; BodyLength = ([string]$item.Body).Length; BodyPreview = $preview; ItemId = $item.ItemId }
    foreach ($k in 'Subject', 'From', 'To', 'DateTimeReceived', 'InternetMessageId') { $d[$k] = $Context.Message.$k }
    $d.BodyLength = $Context.Message.BodyLength
    Add-WscStep $Context $stage 'GetItem' Passed "Message read in full: '$($item.Subject)' from $($item.From), body of $($Context.Message.BodyLength) characters (preview in the report)." $d
}

function Invoke-WscStageFreeBusy {
    <# Free/busy of the mailboxes of Test.FreeBusyMailboxes (else the mailbox), from today for Test.FreeBusyDays days. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'FreeBusy'
    $cfg = $Context.Config
    $mailboxes = @($cfg.FreeBusyMailboxes | Where-Object { $_ })
    if (-not $mailboxes.Count) { $mailboxes = @([string]$cfg.Mailbox) }
    $start = (Get-Date).Date.ToUniversalTime()
    $end = $start.AddDays([int]$cfg.FreeBusyDays)
    $answer = Invoke-WscEws -Context $Context -Operation 'GetUserAvailability' -Body (New-WscAvailabilityBody -Mailboxes $mailboxes -StartUtc $start -EndUtc $end -IntervalMinutes ([int]$cfg.FreeBusyIntervalMinutes))
    $details = Get-WscEwsDetails -Answer $answer -More ([ordered]@{ Mailboxes = $mailboxes -join ', '; WindowUtc = '{0:yyyy-MM-dd HH:mm} to {1:yyyy-MM-dd HH:mm}' -f $start, $end; IntervalMinutes = $cfg.FreeBusyIntervalMinutes })
    if ($answer.HttpStatus -ne 200 -or -not $answer.Xml -or $answer.Fault) {
        Add-WscStep $Context $stage 'GetUserAvailability' Failed "Free/busy not read: $(Get-WscEwsFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $true
        return
    }
    $results = @(ConvertFrom-WscAvailability -Answer $answer -Mailboxes $mailboxes -IntervalMinutes ([int]$cfg.FreeBusyIntervalMinutes) -StartUtc $start)
    $Context.FreeBusy = $results
    $Context.FreeBusyWindow = [pscustomobject]@{ StartUtc = $start.ToString('yyyy-MM-ddTHH:mm:ssZ'); EndUtc = $end.ToString('yyyy-MM-ddTHH:mm:ssZ'); IntervalMinutes = [int]$cfg.FreeBusyIntervalMinutes; Days = [int]$cfg.FreeBusyDays }
    Write-WscFreeBusyGrid -Results $results -StartUtc $start -IntervalMinutes ([int]$cfg.FreeBusyIntervalMinutes) -Days ([int]$cfg.FreeBusyDays)
    foreach ($r in $results) {
        $d = [ordered]@{ Mailbox = $r.Mailbox; ResponseCode = $r.ResponseCode; ViewType = $r.ViewType; Events = @($r.Events).Count; WorkingHours = $r.WorkingHours; MergedFreeBusy = $r.MergedFreeBusy }
        if ($r.ResponseClass -eq 'Success') {
            $kinds = @($r.Events | Group-Object { ([string]$_.BusyType).ToLowerInvariant() } | Sort-Object Name | ForEach-Object { '{0} {1}' -f $_.Count, $(switch ($_.Name) { 'oof' { 'away' } 'workingelsewhere' { 'elsewhere' } default { $_ } }) }) -join ', '
            $detail = if ($r.ViewType -match 'Detailed') { 'with subjects and locations (detailed view allowed)' } elseif ($r.ViewType) { "view $($r.ViewType) (the calendar permission of the caller limits the details)" } else { '' }
            Add-WscStep $Context $stage "Free/busy of $($r.Mailbox)" Passed "$(@($r.Events).Count) calendar event(s)$(if ($kinds) { " ($kinds)" }), $detail. Working hours: $(if ($r.WorkingHours) { $r.WorkingHours } else { 'not returned' })." $d
        }
        else {
            $hint = switch -Regex ([string]$r.ResponseCode) {
                'ErrorMailRecipientNotFound|ErrorNonExistentMailbox' { 'Exchange finds no recipient with this address in its directory. For a mailbox of Exchange Online seen from on-premises (hybrid), the on-premises directory needs a remote mailbox (Enable-RemoteMailbox, or created by the migration) with this address; a cloud-only account has none.' }
                'ErrorFreeBusyGenerationFailed|ErrorProxyRequestProcessingFailed|ErrorAvailabilityConfigNotFound|ErrorNoFreeBusyAccess' { 'the free/busy of this mailbox is not available here: cross-premises or cross-organisation free/busy (organisation relationship, OAuth between on-premises and Exchange Online, Get-AvailabilityAddressSpace) or a calendar permission.' }
                'ErrorFreeBusyDLLimitReached' { 'a distribution group with too many members.' }
                default { $null }
            }
            Add-WscStep $Context $stage "Free/busy of $($r.Mailbox)" Warning "$($r.ResponseCode): $($r.MessageText)$(if ($hint) { " - $hint" })" $d
        }
    }
}

function Get-WscTestFolder {
    <# The test folder under the Inbox: found by name, or created when Create is set. Returns the folder or $null. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage, [switch]$Create)

    if ($Context.TestFolder) { return $Context.TestFolder }
    $name = [string]$Context.Config.FolderName
    $find = Invoke-WscEws -Context $Context -Operation 'FindFolder' -Body (New-WscFindFolderBody -Context $Context -Parent 'inbox' -Traversal Shallow -DisplayName $name -Max 10)
    if ($find.HttpStatus -ne 200 -or $find.ResponseClass -ne 'Success') {
        Add-WscStep $Context $Stage 'Find the test folder' Failed "The Inbox subfolders were not searched: $(Get-WscEwsFailureText -Context $Context -Answer $find)" (Get-WscEwsDetails -Answer $find)
        $Context.Stop = $true
        return $null
    }
    $node = $find.Xml.SelectSingleNode('//t:Folders/*', $find.Ns)
    if ($node) {
        $Context.TestFolder = ConvertFrom-WscFolderNode -Node $node -Ns $find.Ns
        Add-WscStep $Context $Stage 'Find the test folder' Passed "The test folder '$name' exists under the Inbox: it is reused." (Get-WscEwsDetails -Answer $find -More ([ordered]@{ Folder = $name; FolderId = $Context.TestFolder.FolderId }))
        return $Context.TestFolder
    }
    if (-not $Create) {
        Add-WscStep $Context $Stage 'Find the test folder' Failed "No folder '$name' under the Inbox: run CreateFolder (or MailCycle) first." (Get-WscEwsDetails -Answer $find -More ([ordered]@{ Folder = $name }))
        $Context.Stop = $true
        return $null
    }
    $made = Invoke-WscEws -Context $Context -Operation 'CreateFolder' -Body (New-WscCreateFolderBody -Context $Context -Name $name)
    $d = Get-WscEwsDetails -Answer $made -More ([ordered]@{ Folder = $name; Parent = 'Inbox' })
    if ($made.HttpStatus -ne 200 -or $made.ResponseClass -ne 'Success') {
        Add-WscStep $Context $Stage 'CreateFolder' Failed "The folder '$name' was not created: $(Get-WscEwsFailureText -Context $Context -Answer $made)" $d
        $Context.Stop = $true
        return $null
    }
    $id = $made.Xml.SelectSingleNode('//t:FolderId', $made.Ns)
    $Context.TestFolder = [pscustomobject]@{ DisplayName = $name; FolderId = $id.GetAttribute('Id'); Path = "\Inbox\$name" }
    $d.FolderId = $Context.TestFolder.FolderId
    Add-WscAction -Context $Context -Action 'CreateFolder' -Target "Inbox\$name" -Result 'Created' -ItemId $Context.TestFolder.FolderId
    Add-WscStep $Context $Stage 'CreateFolder' Passed "Folder '$name' created under the Inbox of $($Context.Config.Mailbox)." $d
    return $Context.TestFolder
}

function Invoke-WscStageCreateFolder {
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    [void](Get-WscTestFolder -Context $Context -Stage 'CreateFolder' -Create)
}

function Invoke-WscStageSendMail {
    <# Sends the test message, then waits for it in the Inbox when the mailbox is a recipient. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'SendMail'
    $cfg = $Context.Config
    $to = if ([string]$cfg.Recipient) { [string]$cfg.Recipient } else { [string]$cfg.Mailbox }
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $subject = "$($script:TestSubjectPrefix) $stamp $([guid]::NewGuid().ToString('N').Substring(0, 6))"
    $how = switch ($Context.Access) { 'Delegate' { "sent From $($cfg.Mailbox) by the signed-in account (Send As or Send on Behalf)" } 'Impersonation' { "sent as $($cfg.Mailbox) (impersonation)" } default { "sent by $($cfg.Mailbox)" } }
    $text = "Test message of Web Services Client for Exchange $($script:ToolVersion), $how, from $env:COMPUTERNAME at $stamp. It can be deleted."
    $answer = Invoke-WscEws -Context $Context -Operation 'CreateItem' -Body (New-WscSendMessageBody -Context $Context -Subject $subject -BodyText $text -To @($to))
    $details = Get-WscEwsDetails -Answer $answer -More ([ordered]@{ To = $to; Subject = $subject; SavedIn = 'Sent Items'; From = $(if ($Context.Access -eq 'Delegate') { $cfg.Mailbox } else { $null }) })
    if ($answer.HttpStatus -ne 200 -or $answer.ResponseClass -ne 'Success') {
        Add-WscStep $Context $stage 'CreateItem (send)' Failed "The test message was not sent: $(Get-WscEwsFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $true
        return
    }
    Add-WscAction -Context $Context -Action 'SendMail' -Target $to -Result "Sent: $subject"
    $Context.TestMessage = [pscustomobject]@{ Subject = $subject; ItemId = $null; ChangeKey = $null }
    Add-WscStep $Context $stage 'CreateItem (send)' Passed "Test message sent to $to ($how), a copy saved in Sent Items." $details
    if ($to -ine [string]$cfg.Mailbox) { return }

    # Delivery: FindItem on the subject until the message is in the Inbox (requests repeated with the same answer are collapsed in the trace).
    $wait = [int]$cfg.DeliveryWaitSeconds
    $deadline = [DateTimeOffset]::UtcNow.AddSeconds($wait)
    $hit = $null
    $tries = 0
    do {
        $tries++
        $found = Find-WscMessages -Context $Context -Max 1 -SubjectContains $subject -Collapse 'Delivery'
        $hit = $found.Items | Select-Object -First 1
        if ($hit -or [DateTimeOffset]::UtcNow -ge $deadline) { break }
        Wait-WscSeconds 3
    } while ($true)
    $d = [ordered]@{ Subject = $subject; Searches = $tries; WaitedSeconds = [int]($wait - [Math]::Max(0, ($deadline - [DateTimeOffset]::UtcNow).TotalSeconds)) }
    if ($hit) {
        $Context.TestMessage = $hit
        $d.ItemId = $hit.ItemId; $d.Received = $hit.DateTimeReceived
        Add-WscStep $Context $stage 'Delivery to the Inbox' Passed "The test message arrived in the Inbox of $($cfg.Mailbox) ($tries search(es))." $d
    }
    else {
        Add-WscStep $Context $stage 'Delivery to the Inbox' Warning "The test message is not in the Inbox after $wait s (Test.DeliveryWaitSeconds): transport queue, transport rule, or a slow delivery. The next steps look for it again." $d
    }
}

function Select-WscTargetItem {
    <#
        The message a reply, a move or a delete works on: -ItemId, else the most recent Inbox message whose
        subject contains -ItemSubject, else the test message of this run, else the last test message of the tool.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Stage)

    if ($Context.ItemId) { return [pscustomobject]@{ ItemId = $Context.ItemId; ChangeKey = $null; Subject = '(given by -ItemId)'; Source = '-ItemId' } }
    if ($Context.TestMessage -and $Context.TestMessage.ItemId) { return $Context.TestMessage | Select-Object *, @{ n = 'Source'; e = { 'test message of this run' } } }
    $text = if ($Context.ItemSubject) { $Context.ItemSubject } elseif ($Context.TestMessage) { $Context.TestMessage.Subject } else { $script:TestSubjectPrefix }
    $source = if ($Context.ItemSubject) { "-ItemSubject '$text'" } elseif ($Context.TestMessage) { 'test message of this run' } else { 'last test message of the tool' }
    $found = Find-WscMessages -Context $Context -Max 1 -SubjectContains $text
    if ($found.Answer.HttpStatus -ne 200 -or $found.Answer.ResponseClass -ne 'Success') {
        Add-WscStep $Context $Stage 'Find the message' Failed "The Inbox was not searched: $(Get-WscEwsFailureText -Context $Context -Answer $found.Answer)" (Get-WscEwsDetails -Answer $found.Answer)
        $Context.Stop = $true
        return $null
    }
    $hit = $found.Items | Select-Object -First 1
    if (-not $hit) {
        Add-WscStep $Context $Stage 'Find the message' Failed "No Inbox message whose subject contains '$text' ($source): run SendMail first, or name the message with -ItemId or -ItemSubject." (Get-WscEwsDetails -Answer $found.Answer -More ([ordered]@{ Search = $text }))
        $Context.Stop = $true
        return $null
    }
    if (-not $Context.TestMessage -and -not $Context.ItemSubject) { $Context.TestMessage = $hit }
    return $hit | Select-Object *, @{ n = 'Source'; e = { $source } }
}

function Invoke-WscStageReplyMail {
    <# ReplyToItem on the chosen message: the reply goes to its sender and is saved in Sent Items. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'ReplyMail'
    $item = Select-WscTargetItem -Context $Context -Stage $stage
    if (-not $item) { return }
    $text = "Reply of Web Services Client for Exchange $($script:ToolVersion) from $env:COMPUTERNAME at $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) (ReplyToItem)."
    $answer = Invoke-WscEws -Context $Context -Operation 'CreateItem' -Body (New-WscReplyBody -Context $Context -ItemId $item.ItemId -ChangeKey $item.ChangeKey -BodyText $text)
    $details = Get-WscEwsDetails -Answer $answer -More ([ordered]@{ Message = $item.Subject; From = $item.From; Selected = $item.Source; ItemId = $item.ItemId })
    if ($answer.HttpStatus -ne 200 -or $answer.ResponseClass -ne 'Success') {
        Add-WscStep $Context $stage 'ReplyToItem' Failed "No reply to '$($item.Subject)': $(Get-WscEwsFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $true
        return
    }
    Add-WscAction -Context $Context -Action 'ReplyMail' -Target $(if ($item.From) { $item.From } else { 'sender of the message' }) -Result "Replied to: $($item.Subject)" -ItemId $item.ItemId
    Add-WscStep $Context $stage 'ReplyToItem' Passed "Reply sent to the sender of '$($item.Subject)' ($($item.Source)), saved in Sent Items." $details
}

function Invoke-WscStageMoveMail {
    <# MoveItem of the chosen message to the test folder (found, or created with AllowWrite). #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'MoveMail'
    $item = Select-WscTargetItem -Context $Context -Stage $stage
    if (-not $item) { return }
    $folder = Get-WscTestFolder -Context $Context -Stage $stage -Create
    if (-not $folder) { return }
    $answer = Invoke-WscEws -Context $Context -Operation 'MoveItem' -Body (New-WscMoveItemBody -ItemId $item.ItemId -FolderId $folder.FolderId)
    $details = Get-WscEwsDetails -Answer $answer -More ([ordered]@{ Message = $item.Subject; Selected = $item.Source; To = "Inbox\$($folder.DisplayName)"; ItemId = $item.ItemId })
    if ($answer.HttpStatus -ne 200 -or $answer.ResponseClass -ne 'Success') {
        Add-WscStep $Context $stage 'MoveItem' Failed "'$($item.Subject)' was not moved: $(Get-WscEwsFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $true
        return
    }
    $moved = $answer.Xml.SelectSingleNode('//m:Items/*/t:ItemId', $answer.Ns)
    $newId = if ($moved) { $moved.GetAttribute('Id') } else { $null }
    if ($newId) {
        $details.NewItemId = $newId
        if ($Context.TestMessage -and ($Context.TestMessage.ItemId -eq $item.ItemId -or -not $Context.TestMessage.ItemId)) { $Context.TestMessage = [pscustomobject]@{ Subject = $item.Subject; ItemId = $newId; ChangeKey = $moved.GetAttribute('ChangeKey'); From = $item.From } }
        if ($Context.ItemId -eq $item.ItemId) { $Context.ItemId = $newId }
    }
    Add-WscAction -Context $Context -Action 'MoveMail' -Target "Inbox\$($folder.DisplayName)" -Result "Moved: $($item.Subject)" -ItemId $newId
    Add-WscStep $Context $stage 'MoveItem' Passed "'$($item.Subject)' moved to Inbox\$($folder.DisplayName)$(if ($newId) { ': EWS returned its new ID (an item gets a new ID in another folder)' })." $details
}

function Invoke-WscStageDeleteMail {
    <# DeleteItem of the chosen message, with Test.DeleteMode. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'DeleteMail'
    $item = Select-WscTargetItem -Context $Context -Stage $stage
    if (-not $item) { return }
    $mode = [string]$Context.Config.DeleteMode
    $answer = Invoke-WscEws -Context $Context -Operation 'DeleteItem' -Body (New-WscDeleteItemBody -ItemId $item.ItemId -DeleteType $mode)
    $details = Get-WscEwsDetails -Answer $answer -More ([ordered]@{ Message = $item.Subject; Selected = $item.Source; DeleteType = $mode; ItemId = $item.ItemId })
    if ($answer.HttpStatus -ne 200 -or $answer.ResponseClass -ne 'Success') {
        Add-WscStep $Context $stage 'DeleteItem' Failed "'$($item.Subject)' was not deleted: $(Get-WscEwsFailureText -Context $Context -Answer $answer)" $details
        $Context.Stop = $true
        return
    }
    $where = switch ($mode) { 'MoveToDeletedItems' { 'moved to Deleted Items' } 'SoftDelete' { 'soft-deleted (Recoverable Items)' } default { 'hard-deleted (purged)' } }
    Add-WscAction -Context $Context -Action 'DeleteMail' -Target $mode -Result "Deleted: $($item.Subject)" -ItemId $item.ItemId
    Add-WscStep $Context $stage 'DeleteItem' Passed "'$($item.Subject)' $where." $details
    $Context.TestMessage = $null
}
