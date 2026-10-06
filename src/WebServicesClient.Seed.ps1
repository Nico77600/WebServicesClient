<#
.SYNOPSIS
    Web Services Client for Exchange - test data: fill a test mailbox so that the read-only scenarios show something, and clean it (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    SeedData (needs AllowWrite), through EWS or Microsoft Graph like every operation:
      - folders Inbox\Projects, Inbox\Projects\Migration and Inbox\Archive 2026 (found again when they exist);
      - six messages sent to the mailbox itself (real delivered messages, subject starting with [Test data]),
        one of them moved to Projects\Migration once delivered;
      - seven calendar items over the next five working days, without attendees: busy, tentative, out of
        office, working elsewhere - what the free/busy view of the report shows.
    Messages and calendar items already there (subject [Test data]) are not created again.
    CleanData deletes the [Test data] messages of the Inbox, the calendar items of the window and the folders
    (moved to Deleted Items). Nothing else is touched.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:SeedPrefix = '[Test data]'

function Get-WscSeedPlan {
    <# What SeedData creates. Calendar items on the next five working days from today, local time. #>
    $messages = @(
        @{ Subject = 'Budget 2027 - first draft'; Body = "Hello,`n`nHere is the first draft of the 2027 budget: licences, Azure, support. Comments before Friday please.`n`nRegards" }
        @{ Subject = 'Exchange SE migration - weekly status'; Body = "Status of the week:`n- EXCH03 and EXCH04 patched`n- 120 mailboxes moved`n- Next: decommission of the 2016 servers"; MoveTo = 'Projects\Migration' }
        @{ Subject = 'Room booking for the design workshop'; Body = 'Room 1 is booked Wednesday 14:00-16:00. The projector is in the cupboard.' }
        @{ Subject = 'Security review - action items'; Body = "Actions:`n1. Enable Extended Protection everywhere`n2. Remove Basic authentication for EWS`n3. Review the application access policies" }
        @{ Subject = 'Team lunch on Friday'; Body = 'Lunch at 12:30 at the usual place. Tell me if you come.' }
        @{ Subject = 'Mailbox audit - October report'; Body = 'The October audit report is ready: 3 mailboxes with Full Access to review.' }
    )
    $events = @(
        @{ Day = 0; Start = '09:00'; Minutes = 60; Status = 'Busy'; Subject = 'Weekly review'; Location = 'Room 1' }
        @{ Day = 0; Start = '14:00'; Minutes = 90; Status = 'Tentative'; Subject = 'Design sync'; Location = 'Teams' }
        @{ Day = 1; Start = '08:00'; Minutes = 600; Status = 'OOF'; Subject = 'Out of office'; Location = '' }
        @{ Day = 2; Start = '11:00'; Minutes = 60; Status = 'Busy'; Subject = 'Customer call'; Location = 'Teams' }
        @{ Day = 2; Start = '16:30'; Minutes = 30; Status = 'Busy'; Subject = 'One to one'; Location = '' }
        @{ Day = 3; Start = '13:00'; Minutes = 240; Status = 'WorkingElsewhere'; Subject = 'Working from home'; Location = 'Home' }
        @{ Day = 4; Start = '10:00'; Minutes = 120; Status = 'Busy'; Subject = 'Change board'; Location = 'Room 2' }
    )
    # The next five working days, today included.
    $days = [Collections.Generic.List[datetime]]::new()
    for ($d = (Get-Date).Date; $days.Count -lt 5; $d = $d.AddDays(1)) { if ($d.DayOfWeek -notin 'Saturday', 'Sunday') { $days.Add($d) } }
    $items = foreach ($e in $events) {
        $start = $days[$e.Day].Add([TimeSpan]::Parse($e.Start))
        [pscustomobject]@{ Subject = "$($script:SeedPrefix) $($e.Subject)"; StartUtc = $start.ToUniversalTime(); EndUtc = $start.AddMinutes($e.Minutes).ToUniversalTime(); Status = $e.Status; Location = $e.Location }
    }
    [pscustomobject]@{
        Folders  = @('Projects', 'Projects\Migration', 'Archive 2026')
        Messages = @($messages | ForEach-Object { [pscustomobject]@{ Subject = "$($script:SeedPrefix) $($_.Subject)"; Body = $_.Body; MoveTo = $_['MoveTo'] } })
        Events   = @($items)
        FromUtc  = $days[0].ToUniversalTime()
        ToUtc    = $days[4].AddDays(1).ToUniversalTime()
    }
}

#region EWS -----------------------------------------------------------------------------------------

function Find-WscChildFolder {
    <# A folder by name under a parent (distinguished name or folder ID), through EWS; created when Create is set. Returns ID, Created, Answer. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Name, [string]$Parent = 'inbox', [string]$ParentId, [switch]$Create)

    $find = Invoke-WscEws -Context $Context -Operation 'FindFolder' -Body (New-WscFindFolderBody -Context $Context -Parent $Parent -ParentId $ParentId -Traversal Shallow -DisplayName $Name -Max 5)
    if ($find.HttpStatus -ne 200 -or $find.ResponseClass -ne 'Success') { return [pscustomobject]@{ Id = $null; Created = $false; Answer = $find } }
    $node = $find.Xml.SelectSingleNode('//t:Folders/*/t:FolderId', $find.Ns)
    if ($node) { return [pscustomobject]@{ Id = $node.GetAttribute('Id'); Created = $false; Answer = $find } }
    if (-not $Create) { return [pscustomobject]@{ Id = $null; Created = $false; Answer = $find } }
    $parentXml = Get-WscFolderIdXml -Context $Context -Distinguished $Parent -FolderId $ParentId
    $body = '<m:CreateFolder><m:ParentFolderId>' + $parentXml + '</m:ParentFolderId><m:Folders><t:Folder><t:FolderClass>IPF.Note</t:FolderClass><t:DisplayName>' + (ConvertTo-WscXmlText $Name) + '</t:DisplayName></t:Folder></m:Folders></m:CreateFolder>'
    $made = Invoke-WscEws -Context $Context -Operation 'CreateFolder' -Body $body
    $id = if ($made.HttpStatus -eq 200 -and $made.ResponseClass -eq 'Success') { $made.Xml.SelectSingleNode('//t:FolderId', $made.Ns).GetAttribute('Id') } else { $null }
    [pscustomobject]@{ Id = $id; Created = [bool]$id; Answer = $made }
}

function Find-WscSeedCalendarItems {
    <# [Test data] calendar items between two UTC dates (CalendarView), through EWS. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [datetime]$FromUtc, [datetime]$ToUtc)
    $f = 'yyyy-MM-ddTHH:mm:ssZ'
    $body = '<m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties><t:FieldURI FieldURI="item:Subject"/><t:FieldURI FieldURI="calendar:Start"/></t:AdditionalProperties></m:ItemShape>' +
        ('<m:CalendarView MaxEntriesReturned="500" StartDate="{0}" EndDate="{1}"/>' -f $FromUtc.ToString($f), $ToUtc.ToString($f)) + '<m:ParentFolderIds>' + (Get-WscFolderIdXml -Context $Context -Distinguished 'calendar') + '</m:ParentFolderIds></m:FindItem>'
    $answer = Invoke-WscEws -Context $Context -Operation 'FindItem' -Body $body
    $items = @()
    if ($answer.HttpStatus -eq 200 -and $answer.ResponseClass -eq 'Success') {
        $items = @(foreach ($n in $answer.Xml.SelectNodes('//t:Items/t:CalendarItem', $answer.Ns)) {
                $subject = Get-WscNodeText $n 't:Subject' $answer.Ns
                if ([string]$subject -like "$([WildcardPattern]::Escape($script:SeedPrefix))*") { [pscustomobject]@{ Id = $n.SelectSingleNode('t:ItemId', $answer.Ns).GetAttribute('Id'); Subject = $subject } }
            })
    }
    [pscustomobject]@{ Answer = $answer; Items = $items }
}

function Invoke-WscSeedEws {
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'SeedData'
    $plan = Get-WscSeedPlan
    $cfg = $Context.Config
    # Folders
    $ids = @{}
    $made = [Collections.Generic.List[string]]::new()
    foreach ($path in $plan.Folders) {
        $parts = $path.Split('\'); $name = $parts[-1]; $parentPath = ($parts[0..($parts.Count - 2)] -join '\')
        $r = if ($parts.Count -eq 1) { Find-WscChildFolder -Context $Context -Name $name -Parent 'inbox' -Create } else { Find-WscChildFolder -Context $Context -Name $name -ParentId $ids[$parentPath] -Create }
        if (-not $r.Id) {
            Add-WscStep $Context $stage 'Folders' Failed "The folder Inbox\$path was not created: $(Get-WscEwsFailureText -Context $Context -Answer $r.Answer)" (Get-WscEwsDetails -Answer $r.Answer)
            $Context.Stop = $true; return
        }
        $ids[$path] = $r.Id
        if ($r.Created) { $made.Add("Inbox\$path"); Add-WscAction -Context $Context -Action 'CreateFolder' -Target "Inbox\$path" -Result 'Created (test data)' -ItemId $r.Id }
    }
    Add-WscStep $Context $stage 'Folders' Passed "$(if ($made.Count) { "Created: $($made -join ', ')" } else { 'Already there' }) - Inbox\$($plan.Folders -join ', Inbox\')." ([ordered]@{ Folders = ($plan.Folders | ForEach-Object { "Inbox\$_" }) -join ', '; Created = $made.Count })

    # Messages
    $existing = Find-WscMessages -Context $Context -Max 50 -SubjectContains $script:SeedPrefix
    $have = @($existing.Items | ForEach-Object { $_.Subject })
    foreach ($folder in @($plan.Messages | Where-Object { $_.MoveTo } | ForEach-Object { $_.MoveTo } | Select-Object -Unique)) {
        $have += @((Find-WscMessages -Context $Context -FolderId $ids[$folder] -Max 50 -SubjectContains $script:SeedPrefix).Items | ForEach-Object { $_.Subject })
    }
    $toSend = @($plan.Messages | Where-Object { $_.Subject -notin $have })
    foreach ($m in $toSend) {
        $answer = Invoke-WscEws -Context $Context -Operation 'CreateItem' -Body (New-WscSendMessageBody -Context $Context -Subject $m.Subject -BodyText $m.Body -To @([string]$cfg.Mailbox))
        if ($answer.HttpStatus -ne 200 -or $answer.ResponseClass -ne 'Success') {
            Add-WscStep $Context $stage 'Messages' Failed "'$($m.Subject)' was not sent: $(Get-WscEwsFailureText -Context $Context -Answer $answer)" (Get-WscEwsDetails -Answer $answer)
            $Context.Stop = $true; return
        }
        Add-WscAction -Context $Context -Action 'SendMail' -Target ([string]$cfg.Mailbox) -Result "Sent: $($m.Subject)"
    }
    Add-WscStep $Context $stage 'Messages' Passed "$($toSend.Count) message(s) sent to $($cfg.Mailbox)$(if ($have.Count) { ", $($have.Count) already in the Inbox" })." ([ordered]@{ Sent = $toSend.Count; AlreadyThere = $have.Count })

    # One message filed in Projects\Migration once delivered
    # Filed only when sent by this run: a message sent earlier is already where it belongs.
    $move = $plan.Messages | Where-Object { $_.MoveTo -and $_.Subject -in @($toSend.Subject) } | Select-Object -First 1
    if ($move) {
        $deadline = [DateTimeOffset]::UtcNow.AddSeconds([int]$cfg.DeliveryWaitSeconds)
        do { $hit = (Find-WscMessages -Context $Context -Max 1 -SubjectContains $move.Subject -Collapse 'SeedDelivery').Items | Select-Object -First 1; if ($hit -or [DateTimeOffset]::UtcNow -ge $deadline) { break }; Wait-WscSeconds 3 } while ($true)
        if ($hit) {
            $answer = Invoke-WscEws -Context $Context -Operation 'MoveItem' -Body (New-WscMoveItemBody -ItemId $hit.ItemId -FolderId $ids[$move.MoveTo])
            $status = if ($answer.HttpStatus -eq 200 -and $answer.ResponseClass -eq 'Success') { 'Passed' } else { 'Warning' }
            if ($status -eq 'Passed') { Add-WscAction -Context $Context -Action 'MoveMail' -Target "Inbox\$($move.MoveTo)" -Result "Moved: $($move.Subject)" }
            Add-WscStep $Context $stage 'File a message' $status $(if ($status -eq 'Passed') { "'$($move.Subject)' filed in Inbox\$($move.MoveTo)." } else { "Not filed: $(Get-WscEwsFailureText -Context $Context -Answer $answer)" }) (Get-WscEwsDetails -Answer $answer)
        }
        else { Add-WscStep $Context $stage 'File a message' Warning "'$($move.Subject)' is not in the Inbox yet (already filed, or delivery slower than $($cfg.DeliveryWaitSeconds) s)." ([ordered]@{ Subject = $move.Subject }) }
    }

    # Calendar
    $found = Find-WscSeedCalendarItems -Context $Context -FromUtc $plan.FromUtc -ToUtc $plan.ToUtc
    $already = @($found.Items | ForEach-Object { $_.Subject })
    $created = 0
    foreach ($e in $plan.Events) {
        if ($e.Subject -in $already) { continue }
        $f = 'yyyy-MM-ddTHH:mm:ssZ'
        $body = '<m:CreateItem SendMeetingInvitations="SendToNone"><m:SavedItemFolderId>' + (Get-WscFolderIdXml -Context $Context -Distinguished 'calendar') + '</m:SavedItemFolderId><m:Items><t:CalendarItem>' +
            '<t:Subject>' + (ConvertTo-WscXmlText $e.Subject) + '</t:Subject><t:Body BodyType="Text">Test data of Web Services Client for Exchange (free/busy view). It can be deleted.</t:Body>' +
            ('<t:Start>{0}</t:Start><t:End>{1}</t:End><t:LegacyFreeBusyStatus>{2}</t:LegacyFreeBusyStatus>' -f $e.StartUtc.ToString($f), $e.EndUtc.ToString($f), $e.Status) +
            $(if ($e.Location) { '<t:Location>' + (ConvertTo-WscXmlText $e.Location) + '</t:Location>' }) + '</t:CalendarItem></m:Items></m:CreateItem>'
        $answer = Invoke-WscEws -Context $Context -Operation 'CreateItem' -Body $body
        if ($answer.HttpStatus -ne 200 -or $answer.ResponseClass -ne 'Success') {
            Add-WscStep $Context $stage 'Calendar' Failed "'$($e.Subject)' was not created: $(Get-WscEwsFailureText -Context $Context -Answer $answer)" (Get-WscEwsDetails -Answer $answer)
            $Context.Stop = $true; return
        }
        $created++
        Add-WscAction -Context $Context -Action 'CreateEvent' -Target "$($e.StartUtc.ToLocalTime().ToString('ddd dd HH:mm')) $($e.Status)" -Result "Created: $($e.Subject)"
    }
    Add-WscStep $Context $stage 'Calendar' Passed "$created calendar item(s) created over the next five working days (busy, tentative, out of office, working elsewhere)$(if ($already.Count) { ", $($already.Count) already there" }): run FreeBusy or ReadOnly to see them." ([ordered]@{ Created = $created; AlreadyThere = $already.Count; From = $plan.FromUtc.ToLocalTime().ToString('yyyy-MM-dd'); To = $plan.ToUtc.ToLocalTime().AddDays(-1).ToString('yyyy-MM-dd') })
}

function Invoke-WscCleanEws {
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'CleanData'
    $plan = Get-WscSeedPlan
    $messages = Find-WscMessages -Context $Context -Max 100 -SubjectContains $script:SeedPrefix
    $n = 0
    foreach ($m in @($messages.Items | Where-Object { $_.Subject -like "$([WildcardPattern]::Escape($script:SeedPrefix))*" })) {
        $a = Invoke-WscEws -Context $Context -Operation 'DeleteItem' -Body (New-WscDeleteItemBody -ItemId $m.ItemId -DeleteType 'MoveToDeletedItems')
        if ($a.HttpStatus -eq 200 -and $a.ResponseClass -eq 'Success') { $n++; Add-WscAction -Context $Context -Action 'DeleteMail' -Target 'MoveToDeletedItems' -Result "Deleted: $($m.Subject)" }
    }
    Add-WscStep $Context $stage 'Messages' Passed "$n [Test data] message(s) of the Inbox moved to Deleted Items." ([ordered]@{ Deleted = $n })
    $found = Find-WscSeedCalendarItems -Context $Context -FromUtc $plan.FromUtc.AddDays(-14) -ToUtc $plan.ToUtc.AddDays(14)
    $c = 0
    foreach ($e in $found.Items) {
        $a = Invoke-WscEws -Context $Context -Operation 'DeleteItem' -Body ('<m:DeleteItem DeleteType="MoveToDeletedItems" SendMeetingCancellations="SendToNone"><m:ItemIds><t:ItemId Id="' + (ConvertTo-WscXmlText $e.Id) + '"/></m:ItemIds></m:DeleteItem>')
        if ($a.HttpStatus -eq 200 -and $a.ResponseClass -eq 'Success') { $c++; Add-WscAction -Context $Context -Action 'DeleteEvent' -Target 'MoveToDeletedItems' -Result "Deleted: $($e.Subject)" }
    }
    Add-WscStep $Context $stage 'Calendar' Passed "$c [Test data] calendar item(s) deleted (four weeks around today)." ([ordered]@{ Deleted = $c })
    $f = 0
    foreach ($top in 'Projects', 'Archive 2026') {
        $r = Find-WscChildFolder -Context $Context -Name $top -Parent 'inbox'
        if (-not $r.Id) { continue }
        $a = Invoke-WscEws -Context $Context -Operation 'DeleteFolder' -Body ('<m:DeleteFolder DeleteType="MoveToDeletedItems"><m:FolderIds><t:FolderId Id="' + (ConvertTo-WscXmlText $r.Id) + '"/></m:FolderIds></m:DeleteFolder>')
        if ($a.HttpStatus -eq 200 -and $a.ResponseClass -eq 'Success') { $f++; Add-WscAction -Context $Context -Action 'DeleteFolder' -Target "Inbox\$top" -Result 'Moved to Deleted Items' }
    }
    Add-WscStep $Context $stage 'Folders' Passed "$f test folder(s) (Inbox\Projects, Inbox\Archive 2026) moved to Deleted Items with their content." ([ordered]@{ Deleted = $f })
}

#endregion

#region Microsoft Graph -----------------------------------------------------------------------------

function Find-WscGraphChildFolder {
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Name, [string]$ParentPath, [switch]$Create)
    if (-not $ParentPath) { $ParentPath = "$(Get-WscGraphMailboxPath $Context)/mailFolders/inbox" }
    $find = Invoke-WscGraph -Context $Context -Path "$ParentPath/childFolders?`$filter=displayName eq '$(ConvertTo-WscODataText $Name)'&`$select=id,displayName" -Operation 'GET childFolders (filter)'
    if (-not (Test-WscGraphOk $find)) { return [pscustomobject]@{ Id = $null; Created = $false; Answer = $find } }
    $hit = @($find.Json['value']) | Where-Object { $_ } | Select-Object -First 1
    if ($hit) { return [pscustomobject]@{ Id = [string]$hit['id']; Created = $false; Answer = $find } }
    if (-not $Create) { return [pscustomobject]@{ Id = $null; Created = $false; Answer = $find } }
    $made = Invoke-WscGraph -Context $Context -Method POST -Path "$ParentPath/childFolders" -Body @{ displayName = $Name } -Operation 'POST childFolders'
    [pscustomobject]@{ Id = $(if (Test-WscGraphOk $made) { [string]$made.Json['id'] } else { $null }); Created = (Test-WscGraphOk $made); Answer = $made }
}

function Find-WscGraphSeedEvents {
    param([Parameter(Mandatory = $true)][hashtable]$Context, [datetime]$FromUtc, [datetime]$ToUtc)
    $f = 'yyyy-MM-ddTHH:mm:ssZ'
    $answer = Invoke-WscGraph -Context $Context -Path "$(Get-WscGraphMailboxPath $Context)/calendarView?startDateTime=$($FromUtc.ToString($f))&endDateTime=$($ToUtc.ToString($f))&`$select=id,subject&`$top=500" -Operation 'GET calendarView'
    $items = @()
    if (Test-WscGraphOk $answer) { $items = @($answer.Json['value'] | Where-Object { [string]$_['subject'] -like "$([WildcardPattern]::Escape($script:SeedPrefix))*" } | ForEach-Object { [pscustomobject]@{ Id = [string]$_['id']; Subject = [string]$_['subject'] } }) }
    [pscustomobject]@{ Answer = $answer; Items = $items }
}

function Invoke-WscSeedGraph {
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'SeedData'
    $plan = Get-WscSeedPlan
    $cfg = $Context.Config
    $mb = Get-WscGraphMailboxPath $Context
    $ids = @{}
    $made = [Collections.Generic.List[string]]::new()
    foreach ($path in $plan.Folders) {
        $parts = $path.Split('\'); $name = $parts[-1]; $parentPath = ($parts[0..($parts.Count - 2)] -join '\')
        $parent = if ($parts.Count -eq 1) { $null } else { "$mb/mailFolders/$([Uri]::EscapeDataString($ids[$parentPath]))" }
        $r = Find-WscGraphChildFolder -Context $Context -Name $name -ParentPath $parent -Create
        if (-not $r.Id) {
            Add-WscStep $Context $stage 'Folders' Failed "The folder Inbox\$path was not created: $(Get-WscGraphFailureText -Context $Context -Answer $r.Answer)" (Get-WscGraphDetails -Answer $r.Answer)
            $Context.Stop = $true; return
        }
        $ids[$path] = $r.Id
        if ($r.Created) { $made.Add("Inbox\$path"); Add-WscAction -Context $Context -Action 'CreateFolder' -Target "Inbox\$path" -Result 'Created (test data)' -ItemId $r.Id }
    }
    Add-WscStep $Context $stage 'Folders' Passed "$(if ($made.Count) { "Created: $($made -join ', ')" } else { 'Already there' }) - Inbox\$($plan.Folders -join ', Inbox\')." ([ordered]@{ Folders = ($plan.Folders | ForEach-Object { "Inbox\$_" }) -join ', '; Created = $made.Count })

    $existing = Get-WscGraphInboxMessages -Context $Context -Top 50 -Search 'Test data'
    $have = @($existing.Items | Where-Object { $_.Subject -like "$([WildcardPattern]::Escape($script:SeedPrefix))*" } | ForEach-Object { $_.Subject })
    foreach ($folder in @($plan.Messages | Where-Object { $_.MoveTo } | ForEach-Object { $_.MoveTo } | Select-Object -Unique)) {
        $filed = Invoke-WscGraph -Context $Context -Path "$mb/mailFolders/$([Uri]::EscapeDataString($ids[$folder]))/messages?`$select=subject&`$top=50" -Operation 'GET folder messages'
        if (Test-WscGraphOk $filed) { $have += @($filed.Json['value'] | ForEach-Object { [string]$_['subject'] }) }
    }
    $toSend = @($plan.Messages | Where-Object { $_.Subject -notin $have })
    foreach ($m in $toSend) {
        $answer = Invoke-WscGraph -Context $Context -Method POST -Path "$mb/sendMail" -Body ([ordered]@{ message = [ordered]@{ subject = $m.Subject; body = @{ contentType = 'Text'; content = $m.Body }; toRecipients = @(@{ emailAddress = @{ address = [string]$cfg.Mailbox } }) }; saveToSentItems = $true }) -Operation 'POST sendMail'
        if (-not (Test-WscGraphOk $answer)) {
            Add-WscStep $Context $stage 'Messages' Failed "'$($m.Subject)' was not sent: $(Get-WscGraphFailureText -Context $Context -Answer $answer)" (Get-WscGraphDetails -Answer $answer)
            $Context.Stop = $true; return
        }
        Add-WscAction -Context $Context -Action 'SendMail' -Target ([string]$cfg.Mailbox) -Result "Sent: $($m.Subject)"
    }
    Add-WscStep $Context $stage 'Messages' Passed "$($toSend.Count) message(s) sent to $($cfg.Mailbox)$(if ($have.Count) { ", $($have.Count) already in the Inbox" })." ([ordered]@{ Sent = $toSend.Count; AlreadyThere = $have.Count })

    # Filed only when sent by this run: a message sent earlier is already where it belongs.
    $move = $plan.Messages | Where-Object { $_.MoveTo -and $_.Subject -in @($toSend.Subject) } | Select-Object -First 1
    if ($move) {
        $deadline = [DateTimeOffset]::UtcNow.AddSeconds([int]$cfg.DeliveryWaitSeconds)
        do { $hit = (Get-WscGraphInboxMessages -Context $Context -Top 1 -Subject $move.Subject -Collapse 'SeedDelivery').Items | Select-Object -First 1; if ($hit -or [DateTimeOffset]::UtcNow -ge $deadline) { break }; Wait-WscSeconds 3 } while ($true)
        if ($hit) {
            $answer = Invoke-WscGraph -Context $Context -Method POST -Path "$mb/messages/$([Uri]::EscapeDataString($hit.ItemId))/move" -Body @{ destinationId = $ids[$move.MoveTo] } -Operation 'POST message/move'
            $ok = Test-WscGraphOk $answer
            if ($ok) { Add-WscAction -Context $Context -Action 'MoveMail' -Target "Inbox\$($move.MoveTo)" -Result "Moved: $($move.Subject)" }
            Add-WscStep $Context $stage 'File a message' $(if ($ok) { 'Passed' } else { 'Warning' }) $(if ($ok) { "'$($move.Subject)' filed in Inbox\$($move.MoveTo)." } else { "Not filed: $(Get-WscGraphFailureText -Context $Context -Answer $answer)" }) (Get-WscGraphDetails -Answer $answer)
        }
        else { Add-WscStep $Context $stage 'File a message' Warning "'$($move.Subject)' is not in the Inbox yet (already filed, or delivery slower than $($cfg.DeliveryWaitSeconds) s)." ([ordered]@{ Subject = $move.Subject }) }
    }

    $found = Find-WscGraphSeedEvents -Context $Context -FromUtc $plan.FromUtc -ToUtc $plan.ToUtc
    $already = @($found.Items | ForEach-Object { $_.Subject })
    $created = 0
    $show = @{ Busy = 'busy'; Tentative = 'tentative'; OOF = 'oof'; WorkingElsewhere = 'workingElsewhere' }
    foreach ($e in $plan.Events) {
        if ($e.Subject -in $already) { continue }
        $event = [ordered]@{
            subject = $e.Subject; body = @{ contentType = 'text'; content = 'Test data of Web Services Client for Exchange (free/busy view). It can be deleted.' }
            start = @{ dateTime = $e.StartUtc.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }; end = @{ dateTime = $e.EndUtc.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }
            showAs = $show[$e.Status]; isReminderOn = $false
        }
        if ($e.Location) { $event.location = @{ displayName = $e.Location } }
        $answer = Invoke-WscGraph -Context $Context -Method POST -Path "$mb/events" -Body $event -Operation 'POST events'
        if (-not (Test-WscGraphOk $answer)) {
            Add-WscStep $Context $stage 'Calendar' Failed "'$($e.Subject)' was not created: $(Get-WscGraphFailureText -Context $Context -Answer $answer)" (Get-WscGraphDetails -Answer $answer)
            $Context.Stop = $true; return
        }
        $created++
        Add-WscAction -Context $Context -Action 'CreateEvent' -Target "$($e.StartUtc.ToLocalTime().ToString('ddd dd HH:mm')) $($e.Status)" -Result "Created: $($e.Subject)"
    }
    Add-WscStep $Context $stage 'Calendar' Passed "$created calendar item(s) created over the next five working days (busy, tentative, out of office, working elsewhere)$(if ($already.Count) { ", $($already.Count) already there" }): run FreeBusy or ReadOnly to see them." ([ordered]@{ Created = $created; AlreadyThere = $already.Count; From = $plan.FromUtc.ToLocalTime().ToString('yyyy-MM-dd'); To = $plan.ToUtc.ToLocalTime().AddDays(-1).ToString('yyyy-MM-dd') })
}

function Invoke-WscCleanGraph {
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'CleanData'
    $plan = Get-WscSeedPlan
    $mb = Get-WscGraphMailboxPath $Context
    $messages = Get-WscGraphInboxMessages -Context $Context -Top 100 -Search 'Test data'
    $n = 0
    foreach ($m in @($messages.Items | Where-Object { $_.Subject -like "$([WildcardPattern]::Escape($script:SeedPrefix))*" })) {
        $a = Invoke-WscGraph -Context $Context -Method DELETE -Path "$mb/messages/$([Uri]::EscapeDataString($m.ItemId))" -Operation 'DELETE message'
        if (Test-WscGraphOk $a) { $n++; Add-WscAction -Context $Context -Action 'DeleteMail' -Target 'Deleted Items' -Result "Deleted: $($m.Subject)" }
    }
    Add-WscStep $Context $stage 'Messages' Passed "$n [Test data] message(s) of the Inbox moved to Deleted Items." ([ordered]@{ Deleted = $n })
    $found = Find-WscGraphSeedEvents -Context $Context -FromUtc $plan.FromUtc.AddDays(-14) -ToUtc $plan.ToUtc.AddDays(14)
    $c = 0
    foreach ($e in $found.Items) {
        $a = Invoke-WscGraph -Context $Context -Method DELETE -Path "$mb/events/$([Uri]::EscapeDataString($e.Id))" -Operation 'DELETE event'
        if (Test-WscGraphOk $a) { $c++; Add-WscAction -Context $Context -Action 'DeleteEvent' -Target 'Deleted Items' -Result "Deleted: $($e.Subject)" }
    }
    Add-WscStep $Context $stage 'Calendar' Passed "$c [Test data] calendar item(s) deleted (four weeks around today)." ([ordered]@{ Deleted = $c })
    $f = 0
    foreach ($top in 'Projects', 'Archive 2026') {
        $r = Find-WscGraphChildFolder -Context $Context -Name $top
        if (-not $r.Id) { continue }
        $a = Invoke-WscGraph -Context $Context -Method DELETE -Path "$mb/mailFolders/$([Uri]::EscapeDataString($r.Id))" -Operation 'DELETE mailFolder'
        if (Test-WscGraphOk $a) { $f++; Add-WscAction -Context $Context -Action 'DeleteFolder' -Target "Inbox\$top" -Result 'Moved to Deleted Items' }
    }
    Add-WscStep $Context $stage 'Folders' Passed "$f test folder(s) (Inbox\Projects, Inbox\Archive 2026) moved to Deleted Items with their content." ([ordered]@{ Deleted = $f })
}

#endregion
