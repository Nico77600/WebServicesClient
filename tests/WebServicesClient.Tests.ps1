#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.1.0' }
<#
    Web Services Client for Exchange - automated tests (Pester 6.1 or later).
    Author  : Nicolas Fabert
    Version : 1.0.0

    Run:  .\Run-Tests.ps1      (or Invoke-Pester -Path .\tests -Output Detailed)

    No Exchange, AD FS or Entra ID is needed: tests\WebServicesClient.Simulator.ps1 replaces the network
    inside the module (Send-WscHttpRequest) and answers like the real servers, NTLM challenge included.
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'WebServicesClient.Simulator.ps1')
    $script:Base = @{
        Mailbox = 'ews-test@contoso.test'; Discovery = 'Manual'; EwsUrl = 'https://mail.contoso.test/EWS/Exchange.asmx'; AdfsUrl = 'https://adfs.contoso.test/adfs'
        SignIn = 'DeviceCode'; OutputPath = (Join-Path $script:RepoRoot 'artifacts\test-reports'); LogPath = (Join-Path $script:RepoRoot 'artifacts\test-logs')
    }
    function Start-Sim([hashtable]$Changes = @{}) {
        Import-Module (Join-Path $script:Root 'WebServicesClient.psd1') -Force
        $script:Module = Get-Module WebServicesClient
        $script:State = New-SimState
        foreach ($k in $Changes.Keys) { $script:State[$k] = $Changes[$k] }
        Install-SimExchange -Module $script:Module -State $script:State
    }
    function Invoke-Sim([string]$Type, [hashtable]$Overrides = @{}, [hashtable]$State = @{}, [pscredential]$Credential, [securestring]$ClientSecret, [string]$ItemSubject, [string]$ItemId) {
        Start-Sim $State
        $cfg = $script:Base.Clone(); foreach ($k in $Overrides.Keys) { $cfg[$k] = $Overrides[$k] }
        Invoke-WscMailboxTest -Configuration $cfg -TestType $Type -Credential $Credential -ClientSecret $ClientSecret -ItemSubject $ItemSubject -ItemId $ItemId -Quiet
    }
    function Get-Step($Result, [string]$Name) { $Result.Steps | Where-Object Name -eq $Name | Select-Object -First 1 }
    $script:Basic = [pscredential]::new('ews-test@contoso.test', (ConvertTo-SecureString 'Sim-Pa55word!' -AsPlainText -Force))
    $script:Secret = ConvertTo-SecureString 'Sim-Secret!' -AsPlainText -Force
    $script:App = '11111111-2222-3333-4444-555555555555'
}

Describe 'Configuration' {
    BeforeAll { Import-Module (Join-Path $script:Root 'WebServicesClient.psd1') -Force }

    It 'reads the sections and resolves relative paths from the tool folder' {
        $c = Import-WscConfiguration
        $c.OutputPath | Should -Be (Join-Path $script:Root 'reports')
        $c.TestType | Should -Be 'ReadOnly'
        $c.AllowWrite | Should -BeFalse
        $c.Discovery | Should -Be 'Autodiscover'
    }

    It 'lists unknown sections, unknown keys and invalid values together' {
        $path = Join-Path $script:RepoRoot 'artifacts\bad.config.psd1'
        [void][IO.Directory]::CreateDirectory((Split-Path $path))
        "@{ Target = @{ EwsUrl = 'http://mail/owa'; Mailbx = 'a' }; Extra = @{}; Test = @{ MessageCount = 500; DefaultType = 'Nope' }; Identity = @{ Authentication = 'Digest' } }" | Set-Content $path
        $text = ({ Import-WscConfiguration -Path $path } | Should -Throw -PassThru).Exception.Message
        $text | Should -Match "Unknown key 'Target.Mailbx'"
        $text | Should -Match "Unknown section 'Extra'"
        $text | Should -Match 'Target.EwsUrl must be an https'
        $text | Should -Match 'Test.MessageCount must be'
        $text | Should -Match 'Test.DefaultType must be one of'
        $text | Should -Match "Identity.Authentication must be 'OAuth', 'Basic' or 'Windows'"
        Remove-Item $path
    }

    It 'exposes the scenarios with their stages and writing flags' {
        $catalog = Get-WscTestCatalog
        $catalog.Name | Should -Be @('Discovery', 'SignIn', 'Endpoint', 'Folders', 'ReadMail', 'FreeBusy', 'CreateFolder', 'SendMail', 'ReplyMail', 'MoveMail', 'DeleteMail', 'SeedData', 'CleanData', 'ReadOnly', 'MailCycle', 'Full')
        ($catalog | Where-Object Name -eq 'ReadOnly').Writes | Should -BeFalse
        ($catalog | Where-Object Name -eq 'MailCycle').Writes | Should -BeTrue
        ($catalog | Where-Object Name -eq 'Discovery').SignIn | Should -BeFalse
        Get-WscScenarioStages -TestType 'ReadMail' -Discovery Autodiscover | Should -Be @('Autodiscover', 'SignIn', 'Endpoint', 'ReadMail')
        Get-WscScenarioStages -TestType 'ReadMail' -Discovery Manual | Should -Be @('SignIn', 'Endpoint', 'ReadMail')
    }

    It 'refuses the combinations that cannot work' {
        $defaults = & (Get-Module WebServicesClient) { Get-WscDefaultConfiguration }
        foreach ($k in $script:Base.Keys) { $defaults[$k] = $script:Base[$k] }
        $c = $defaults.Clone(); $c.Context = 'Delegated'
        (Test-WscConfiguration -Configuration $c).Problems -join ' ' | Should -Match 'AppClientId'
        $c = $defaults.Clone(); $c.Authentication = 'Basic'; $c.Context = 'Application'
        (Test-WscConfiguration -Configuration $c).Problems -join ' ' | Should -Match 'Basic and Windows sign in as a user'
        $c = $defaults.Clone(); $c.EwsUrl = 'https://outlook.office365.com/EWS/Exchange.asmx'; $c.Authentication = 'Windows'
        (Test-WscConfiguration -Configuration $c).Problems -join ' ' | Should -Match 'does not offer Windows authentication'
        $c = $defaults.Clone(); $c.Context = 'Application'; $c.AppClientId = $script:App; $c.Access = 'Delegate'
        (Test-WscConfiguration -Configuration $c).Problems -join ' ' | Should -Match 'no mailbox of its own'
    }

    It 'derives the access mode from the context and the signing-in account' {
        $m = Get-Module WebServicesClient
        & $m { Resolve-WscAccess -Configuration @{ Access = 'Auto'; Context = 'User'; SignInUser = ''; Mailbox = 'a@b.c' } } | Should -Be 'Self'
        & $m { Resolve-WscAccess -Configuration @{ Access = 'Auto'; Context = 'User'; SignInUser = 'x@b.c'; Mailbox = 'a@b.c' } } | Should -Be 'Delegate'
        & $m { Resolve-WscAccess -Configuration @{ Access = 'Auto'; Context = 'Application'; SignInUser = ''; Mailbox = 'a@b.c' } } | Should -Be 'Impersonation'
    }

    It 'builds the scope of the token for each context and server' {
        $m = Get-Module WebServicesClient
        (& $m { Resolve-WscEndpoints -Configuration @{ Authority = 'ADFS'; AdfsUrl = 'https://adfs.contoso.test/adfs'; EwsUrl = 'https://mail.contoso.test/EWS/Exchange.asmx'; Context = 'User' } }).Scope | Should -Be 'openid https://mail.contoso.test//EWS.AccessAsUser.All'
        (& $m { Resolve-WscEndpoints -Configuration @{ Authority = 'EntraID'; TenantId = 't'; EwsUrl = 'https://outlook.office365.com/EWS/Exchange.asmx'; Context = 'User' } }).Scope | Should -Be 'https://outlook.office365.com/EWS.AccessAsUser.All'
        (& $m { Resolve-WscEndpoints -Configuration @{ Authority = 'EntraID'; TenantId = 't'; EwsUrl = 'https://outlook.office365.com/EWS/Exchange.asmx'; Context = 'Application' } }).Scope | Should -Be 'https://outlook.office365.com/.default'
    }
}

Describe 'Windows tokens' {
    BeforeAll { Import-Module (Join-Path $script:Root 'WebServicesClient.psd1') -Force; . (Join-Path $PSScriptRoot 'WebServicesClient.Simulator.ps1') }

    It 'decodes an NTLM challenge: server, domains and version' {
        $info = & (Get-Module WebServicesClient) { param($t) Get-WscNegotiateInfo -Base64 $t } (New-SimNtlmChallenge)
        $info.Ntlm.Type | Should -Be 2
        $info.Ntlm.DnsComputer | Should -Be 'EXCH01.contoso.test'
        $info.Ntlm.NetBiosDomain | Should -Be 'CONTOSO'
        $info.Ntlm.ServerVersion | Should -Be 'Windows 10.0 build 26100'
    }

    It 'never writes the NTLM response, only the user, the domain and the workstation' {
        $o = [Net.Security.NegotiateAuthenticationClientOptions]::new(); $o.Package = 'NTLM'; $o.Credential = [Net.NetworkCredential]::new('ews-test', 'Secret-123', 'CONTOSO'); $o.TargetName = 'HTTP/mail.contoso.test'
        $a = [Net.Security.NegotiateAuthentication]::new($o); $s = [Net.Security.NegotiateAuthenticationStatusCode]::GenericFailure
        $null = $a.GetOutgoingBlob([NullString]::Value, [ref]$s)
        $type3 = $a.GetOutgoingBlob((New-SimNtlmChallenge), [ref]$s)
        $a.Dispose()
        $text = & (Get-Module WebServicesClient) { param($t) Format-WscNegotiateSummary -Base64 $t } $type3
        $text | Should -Match 'NTLM type 3'
        $text | Should -Match 'User ews-test'
        $text | Should -Match 'Domain CONTOSO'
        $text | Should -Match 'never written'
        $text | Should -Not -Match 'Secret-123'
    }
}

Describe 'Scenarios against the simulated Exchange' {
    AfterEach { Remove-Module WebServicesClient -ErrorAction SilentlyContinue }

    It 'ReadOnly with AD FS (device code): folders, messages, free/busy and the server version' {
        $r = Invoke-Sim 'ReadOnly'
        $r.Status | Should -Be 'Passed'
        (Get-Step $r 'Device-code sign-in').Status | Should -Be 'Passed'
        (Get-Step $r 'Token claims').Status | Should -Be 'Passed'
        (Get-Step $r 'GetFolder (Inbox)').Message | Should -Match 'Exchange Server SE 15\.2\.2562\.17'
        $r.Folders.Count | Should -Be 3
        $r.Messages.Count | Should -Be 2
        $r.Message.BodyPreview | Should -Match 'budget'
        @($r.FreeBusyEvents).Count | Should -BeGreaterThan 3
        (Get-Step $r 'Routing and affinity').Message | Should -Match 'X-AnchorMailbox routes'
    } -Skip:$false

    It 'sends X-AnchorMailbox, client-request-id, RequestServerVersion and the affinity cookie on the EWS requests' {
        $r = Invoke-Sim 'Folders'
        $ews = @($r.Trace | Where-Object Operation)
        $ews.Count | Should -BeGreaterThan 1
        $ews[0].Request | Should -Match 'X-AnchorMailbox: ews-test@contoso.test'
        $ews[0].Request | Should -Match 'client-request-id: [0-9a-f-]{36}'
        $ews[0].Request | Should -Match 'return-client-request-id: true'
        $ews[0].Request | Should -Match 'RequestServerVersion Version="Exchange2016"'
        ($ews.Response -join '') | Should -Match 'Set-Cookie: X-BackEndOverrideCookie=EXCH02\.contoso\.test~1942062522'
        $ews[0].Server | Should -Be 'FE EXCH01 > BE EXCH02'
    }

    It 'Entra ID application (client secret): role full_access_as_app and the impersonation header' {
        $r = Invoke-Sim 'ReadMail' @{ Authority = 'EntraID'; Context = 'Application'; AppClientId = $script:App } -ClientSecret $script:Secret
        $r.Status | Should -Be 'Passed'
        $r.Access | Should -Be 'Impersonation'
        (Get-Step $r 'Token claims').Details.Roles | Should -Be 'full_access_as_app'
        ($r.Trace | Where-Object Operation -eq 'FindItem').Request | Should -Match '<t:ExchangeImpersonation>\s*<t:ConnectingSID>\s*<t:SmtpAddress>ews-test@contoso.test'
        ($r.Trace | Where-Object { $_.Url -match '/token$' }).Request | Should -Match 'client_secret=<11 characters, never written>'
    }

    It 'a wrong client secret stops at the sign-in with the cause' {
        $r = Invoke-Sim 'ReadMail' @{ Authority = 'EntraID'; Context = 'Application'; AppClientId = $script:App } -ClientSecret (ConvertTo-SecureString 'wrong' -AsPlainText -Force)
        $r.Status | Should -Be 'Failed'
        $r.Error | Should -Match 'AADSTS7000215.*client secret is wrong or expired'
    }

    It 'impersonation refused gives the role to check' {
        $r = Invoke-Sim 'Endpoint' @{ Authority = 'EntraID'; Context = 'Application'; AppClientId = $script:App } -State @{ DenyImpersonation = $true } -ClientSecret $script:Secret
        $r.Error | Should -Match 'ErrorImpersonateUserDenied.*ApplicationImpersonation'
    }

    It 'delegated application with another signing-in account opens the mailbox as a delegate' {
        $r = Invoke-Sim 'Folders' @{ Authority = 'EntraID'; Context = 'Delegated'; AppClientId = $script:App; SignInUser = 'alice@contoso.test' }
        $r.Access | Should -Be 'Delegate'
        ($r.Trace | Where-Object Operation -eq 'FindFolder').Request | Should -Match '<t:DistinguishedFolderId Id="msgfolderroot">\s*<t:Mailbox>\s*<t:EmailAddress>ews-test@contoso.test'
        $r = Invoke-Sim 'Folders' @{ SignInUser = 'alice@contoso.test' } -State @{ DenyDelegate = $true }
        $r.Error | Should -Match 'ErrorAccessDenied.*Full Access'
    }

    It 'Basic: password accepted, wrong password refused with a user that does not exist' {
        $r = Invoke-Sim 'ReadOnly' @{ Authentication = 'Basic' } -Credential $script:Basic
        (Get-Step $r 'Basic sign-in').Status | Should -Be 'Passed'
        (Get-Step $r 'Wrong password').Status | Should -Be 'Passed'
        $all = ($r.Trace | ForEach-Object { $_.Request }) -join "`n"
        $all | Should -Not -Match 'Sim-Pa55word'
        $all | Should -Match 'Basic <user ews-test@contoso.test, password never written>'
    }

    It 'Windows NTLM: three legs traced and decoded, server named, nothing secret written' {
        $cred = [pscredential]::new('CONTOSO\ews-test', (ConvertTo-SecureString 'Sim-Pa55word!' -AsPlainText -Force))
        $r = Invoke-Sim 'Endpoint' @{ Authentication = 'Windows'; WindowsPackage = 'NTLM' } -Credential $cred
        $r.Status | Should -Be 'Passed'
        $r.Windows.Package | Should -Be 'NTLM'
        $r.Windows.Legs | Should -Be 2
        $r.Windows.Server | Should -Be 'EXCH01.contoso.test'
        $legs = @($r.Trace | Where-Object { $_.Label -match 'leg' })
        $legs[0].Request | Should -Match 'NTLM type 1 \(negotiate\)'
        $legs[0].Response | Should -Match 'NTLM type 2 \(challenge\).*DnsComputer EXCH01\.contoso\.test'
        $legs[1].Request | Should -Match 'NTLM type 3 \(authenticate\).*User ews-test'
        $legs[1].Request | Should -Match 'Spn HTTP/mail\.contoso\.test; ChannelBinding present'
        (Get-Step $r 'Windows sign-in').Details.ChannelBinding | Should -Match 'tls-server-end-point, SHA-256 of CN=mail\.contoso\.test'
        (($r.Trace | ForEach-Object { $_.Request }) -join '') | Should -Not -Match 'Sim-Pa55word'
        # The next request reuses the authenticated connection.
        ($r.Trace | Where-Object Operation -eq 'GetFolder' | Select-Object -Last 1).Label | Should -Match 'Windows session of the connection'
    }

    It 'Windows Discovery: NTLM challenge without credentials and the Kerberos realm' {
        $r = Invoke-Sim 'Discovery' @{ Authentication = 'Windows' }
        (Get-Step $r 'NTLM challenge').Message | Should -Match 'server EXCH01\.contoso\.test, domain CONTOSO \(contoso\.test\)'
        (Get-Step $r 'Kerberos realm').Status | Should -Be 'Warning'
        (Get-Step $r 'Kerberos realm').Message | Should -Match 'fall back to NTLM'
    }

    It 'writing scenarios are blocked without AllowWrite' {
        $r = Invoke-Sim 'MailCycle'
        $r.Status | Should -Be 'Blocked'
        (Get-Step $r 'Create the test folder (CreateFolder)').Status | Should -Be 'Blocked'
        $script:State.Calls | Where-Object Operation -eq 'CreateFolder' | Should -BeNullOrEmpty
    }

    It 'MailCycle: folder created, test message sent and received, replied to, moved, deleted' {
        $r = Invoke-Sim 'MailCycle' @{ AllowWrite = $true }
        $r.Status | Should -Be 'Passed'
        $r.Actions.Action | Should -Be @('CreateFolder', 'SendMail', 'ReplyMail', 'MoveMail', 'DeleteMail')
        (Get-Step $r 'Delivery to the Inbox').Status | Should -Be 'Passed'
        (Get-Step $r 'MoveItem').Details.NewItemId | Should -Match '-MOVED$'
        ($script:State.Inbox | Where-Object { $_.Subject -like '`[Web Services Client for Exchange`]*' }) | Should -BeNullOrEmpty
    }

    It 'ReplyMail on a message chosen by its subject' {
        $r = Invoke-Sim 'ReplyMail' @{ AllowWrite = $true } -ItemSubject 'Budget'
        $r.Status | Should -Be 'Passed'
        (Get-Step $r 'ReplyToItem').Details.Selected | Should -Match "-ItemSubject 'Budget'"
        ($r.Trace | Where-Object Operation -eq 'CreateItem').Request | Should -Match '<t:ReferenceItemId Id="M-1" ChangeKey="CK-M-1"\s*/>'
    }

    It 'never replies to an arbitrary message: without a test message the step fails' {
        $r = Invoke-Sim 'DeleteMail' @{ AllowWrite = $true }
        $r.Status | Should -Be 'Failed'
        $r.Error | Should -Match 'run SendMail first'
    }

    It 'folder paths use case-sensitive IDs, and folders come in tree order with their depth' {
        Import-Module (Join-Path $script:Root 'WebServicesClient.psd1') -Force; $m = Get-Module WebServicesClient
        $f = { param($n, $id, $p) [pscustomobject]@{ DisplayName = $n; FolderId = $id; ParentFolderId = $p; Path = $null } }
        $list = @((& $f 'Inbox' 'AbC' 'ROOT'), (& $f 'Contacts' 'aBc' 'ROOT'), (& $f 'Projects' 'x1' 'AbC'), (& $f 'GAL Contacts' 'x2' 'aBc'), (& $f 'Migration' 'x3' 'x1'))
        $out = & $m { param($l) Set-WscFolderPaths -Folders $l } $list
        ($out | Where-Object DisplayName -eq 'Migration').Path | Should -Be '\Inbox\Projects\Migration'
        ($out | Where-Object DisplayName -eq 'GAL Contacts').Path | Should -Be '\Contacts\GAL Contacts'
        ($out | Where-Object DisplayName -eq 'Migration').Depth | Should -Be 2
        $out.DisplayName | Should -Be @('Contacts', 'GAL Contacts', 'Inbox', 'Projects', 'Migration')
    }

    It 'working hours: the UTC offset of the zone of the mailbox, summer and winter' {
        Import-Module (Join-Path $script:Root 'WebServicesClient.psd1') -Force; $m = Get-Module WebServicesClient
        $xml = [xml]'<t:TimeZone xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"><t:Bias>-60</t:Bias><t:StandardTime><t:Bias>0</t:Bias><t:Time>03:00:00</t:Time><t:DayOrder>5</t:DayOrder><t:Month>10</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek></t:StandardTime><t:DaylightTime><t:Bias>-60</t:Bias><t:Time>02:00:00</t:Time><t:DayOrder>5</t:DayOrder><t:Month>3</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek></t:DaylightTime></t:TimeZone>'
        $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable); $ns.AddNamespace('t', 'http://schemas.microsoft.com/exchange/services/2006/types')
        & $m { param($z, $n) Get-WscEwsZoneOffset -Zone $z -Ns $n -AtUtc ([datetime]'2026-10-06T08:00:00') } $xml.DocumentElement $ns | Should -Be 120
        & $m { param($z, $n) Get-WscEwsZoneOffset -Zone $z -Ns $n -AtUtc ([datetime]'2026-12-15T08:00:00') } $xml.DocumentElement $ns | Should -Be 60
        & $m { param($z, $n) Get-WscEwsZoneOffset -Zone $z -Ns $n -AtUtc ([datetime]'2026-10-25T00:30:00') } $xml.DocumentElement $ns | Should -Be 120
        & $m { param($z, $n) Get-WscEwsZoneOffset -Zone $z -Ns $n -AtUtc ([datetime]'2026-10-25T01:30:00') } $xml.DocumentElement $ns | Should -Be 60
    }
    It 'free/busy: merged view, working hours and the window for the calendar view' {
        $r = Invoke-Sim 'FreeBusy'
        $fb = $r.FreeBusy[0]
        $fb.MergedFreeBusy.Length | Should -Be (7 * 48)
        $fb.WorkingDays | Should -Contain 'monday'
        $fb.WorkStart | Should -Be 480
        $r.FreeBusyWindow.IntervalMinutes | Should -Be 30
        @($r.FreeBusyEvents.BusyType | Select-Object -Unique).Count | Should -BeGreaterThan 2
    }

    It 'free/busy of several mailboxes, one unknown' {
        $r = Invoke-Sim 'FreeBusy' @{ FreeBusyMailboxes = @('ews-test@contoso.test', 'room1@contoso.test', 'nobody@contoso.test') }
        (Get-Step $r 'Free/busy of room1@contoso.test').Status | Should -Be 'Passed'
        (Get-Step $r 'Free/busy of nobody@contoso.test').Message | Should -Match 'ErrorMailRecipientNotFound'
        @($r.FreeBusyEvents).Count | Should -BeGreaterThan 8
    }

    It 'waits and sends again after ErrorServerBusy' {
        $r = Invoke-Sim 'Endpoint' -State @{ ServerBusyOnce = $true }
        $r.Status | Should -Be 'Passed'
        $r.Server.Throttled | Should -Be 1
    }

    It 'Autodiscover v2 finds the URL; POX with Basic when v2 does not answer; the typed URL as a last resort' {
        $r = Invoke-Sim 'Endpoint' @{ Discovery = 'Autodiscover'; EwsUrl = '' }
        (Get-Step $r 'EWS URL').Message | Should -Match 'Autodiscover v2'
        $r = Invoke-Sim 'Endpoint' @{ Discovery = 'Autodiscover'; EwsUrl = ''; Authentication = 'Basic' } -State @{ AutodiscoverV2 = $null; AutodiscoverPox = $true } -Credential $script:Basic
        (Get-Step $r 'EWS URL').Message | Should -Match 'classic Autodiscover \(POX\).*mail\.contoso\.test'
        $r = Invoke-Sim 'Endpoint' @{ Discovery = 'Autodiscover' } -State @{ AutodiscoverV2 = $null }
        (Get-Step $r 'EWS URL').Status | Should -Be 'Warning'
        (Get-Step $r 'GetFolder (Inbox)').Status | Should -Be 'Passed'
    }

    It 'Discovery with Entra ID: tenant, OAuth for the mailbox, forged token refused' {
        $r = Invoke-Sim 'Discovery' @{ Authority = 'EntraID' } -State @{ AuthorizationUri = 'https://login.windows.net/common/oauth2/authorize'; TrustedIssuers = "00000001-0000-0000-c000-000000000000@$($script:SimTenantId)" }
        (Get-Step $r 'OAuth for the mailbox').Status | Should -Be 'Passed'
        (Get-Step $r 'Tenant trusted by Exchange').Status | Should -Be 'Passed'
        (Get-Step $r 'Forged token').Status | Should -Be 'Passed'
    }
}

Describe 'Microsoft Graph for Exchange Online' {
    AfterEach { Remove-Module WebServicesClient -ErrorAction SilentlyContinue }
    BeforeAll { $script:Online = @{ AutodiscoverV2 = 'https://outlook.office365.com/EWS/Exchange.asmx'; AuthorizationUri = 'https://login.microsoftonline.com/common/oauth2/authorize' } }

    It 'Auto chooses Graph for a mailbox in Exchange Online and explains why' {
        $r = Invoke-Sim 'ReadOnly' @{ Discovery = 'Autodiscover'; EwsUrl = ''; Authority = 'EntraID' } -State $script:Online
        $r.Status | Should -Be 'Passed'
        $r.Protocol | Should -Be 'Graph'
        (Get-Step $r 'Protocol').Message | Should -Match 'EWS is being retired'
        (Get-Step $r 'Graph challenge').Status | Should -Be 'Passed'
        (Get-Step $r 'Token claims').Message | Should -Match 'Audience Microsoft Graph'
        $r.Folders.Count | Should -Be 3
        $r.Messages.Count | Should -Be 2
        @($r.FreeBusyEvents).Count | Should -BeGreaterThan 3
        $r.StageTitles['Folders'] | Should -Be 'Folders (mailFolders)'
        ($r.Trace | Where-Object { $_.Url -match 'graph.microsoft.com' -and $_.Operation -eq 'GET inbox messages' }).Request | Should -Match 'client-request-id: [0-9a-f-]{36}'
        ($r.Trace | Where-Object { $_.Url -match '/devicecode$' }).Request | Should -Match 'scope=https://graph.microsoft.com/.default'
    }

    It 'Auto keeps EWS for a mailbox on-premises' {
        $r = Invoke-Sim 'Endpoint'
        $r.Protocol | Should -Be 'EWS'
        (Get-Step $r 'Protocol').Message | Should -Match 'Microsoft Graph does not reach on-premises'
    }

    It 'MailCycle through Graph: folder, sendMail, delivery, reply, move, delete' {
        $r = Invoke-Sim 'MailCycle' @{ Protocol = 'Graph'; Authority = 'EntraID'; AllowWrite = $true }
        $r.Status | Should -Be 'Passed'
        $r.Actions.Action | Should -Be @('CreateFolder', 'SendMail', 'ReplyMail', 'MoveMail', 'DeleteMail')
        ($r.Trace | Where-Object Operation -eq 'POST sendMail').Request | Should -Match '"saveToSentItems":\s*true'
        ($script:State.Inbox | Where-Object { $_.Subject -like '`[Web Services Client for Exchange`]*' }) | Should -BeNullOrEmpty
    }

    It 'application through Graph: roles, /users/{mailbox}, no impersonation header' {
        $r = Invoke-Sim 'ReadMail' @{ Protocol = 'Graph'; Authority = 'EntraID'; Context = 'Application'; AppClientId = $script:App } -ClientSecret $script:Secret
        $r.Status | Should -Be 'Passed'
        (Get-Step $r 'Token claims').Details.Roles | Should -Match 'Mail.ReadWrite'
        ($r.Trace | Where-Object Operation -eq 'GET inbox messages').Url | Should -Match '/v1.0/users/ews-test%40contoso.test/mailFolders/inbox/messages'
    }

    It 'a missing permission is named before the request fails' {
        $r = Invoke-Sim 'SendMail' @{ Protocol = 'Graph'; Authority = 'EntraID'; AllowWrite = $true } -State @{ GraphScopes = 'Mail.Read User.Read'; GraphForbidden = $true }
        (Get-Step $r 'Token claims').Message | Should -Match 'no delegated permission \(scp\) to send \(Mail.Send\)'
        $r.Error | Should -Match 'HTTP 403 \(ErrorAccessDenied'
    }

    It 'EWS forced on Exchange Online: warned, and the 403 names the retirement' {
        $r = Invoke-Sim 'Endpoint' @{ Protocol = 'EWS'; Discovery = 'Manual'; EwsUrl = 'https://outlook.office365.com/EWS/Exchange.asmx'; Authority = 'EntraID' } -State @{ EwsBlocked = $true; TokenAudience = 'https://outlook.office365.com/' }
        (Get-Step $r 'Protocol').Status | Should -Be 'Warning'
        $r.Error | Should -Match 'X-EWS-Policy-Reason.*retirement'
    }

    It 'Graph refuses Basic, Windows and AD FS in the configuration' {
        Import-Module (Join-Path $script:Root 'WebServicesClient.psd1') -Force
        $c = & (Get-Module WebServicesClient) { Get-WscDefaultConfiguration }
        $c.Protocol = 'Graph'; $c.Authentication = 'Basic'
        (Test-WscConfiguration -Configuration $c).Problems -join ' ' | Should -Match 'Graph accepts only OAuth'
        $c.Authentication = 'OAuth'; $c.Authority = 'ADFS'
        (Test-WscConfiguration -Configuration $c).Problems -join ' ' | Should -Match 'Graph accepts only Entra ID'
    }
}

Describe 'Test data' {
    AfterEach { Remove-Module WebServicesClient -ErrorAction SilentlyContinue }

    It 'SeedData through EWS: folders, six messages (one filed), seven calendar items with four statuses; nothing twice' {
        Start-Sim
        $cfg = $script:Base.Clone(); $cfg.AllowWrite = $true
        $r = Invoke-WscMailboxTest -Configuration $cfg -TestType SeedData -Quiet
        $r.Status | Should -Be 'Passed'
        ($script:State.Folders | Where-Object { $_.Name -in 'Projects', 'Migration', 'Archive 2026' }).Count | Should -Be 3
        ($script:State.Folders | Where-Object Name -eq 'Migration').Parent | Should -Be ($script:State.Folders | Where-Object Name -eq 'Projects').Id
        @($r.Actions | Where-Object Action -eq 'SendMail').Count | Should -Be 6
        (Get-Step $r 'File a message').Status | Should -Be 'Passed'
        $script:State.Calendar.Count | Should -Be 7
        @($script:State.Calendar.Status | Select-Object -Unique).Count | Should -Be 4
        ($r.Trace | Where-Object { $_.Request -match '<t:CalendarItem>' } | Select-Object -First 1).Request | Should -Match 'SendMeetingInvitations="SendToNone"'
        $again = Invoke-WscMailboxTest -Configuration $cfg -TestType SeedData -Quiet
        @($again.Actions).Count | Should -Be 0
        $script:State.Calendar.Count | Should -Be 7
    }

    It 'CleanData removes only the test data' {
        Start-Sim
        $cfg = $script:Base.Clone(); $cfg.AllowWrite = $true
        $null = Invoke-WscMailboxTest -Configuration $cfg -TestType SeedData -Quiet
        $r = Invoke-WscMailboxTest -Configuration $cfg -TestType CleanData -Quiet
        $r.Status | Should -Be 'Passed'
        $script:State.Calendar.Count | Should -Be 0
        ($script:State.Inbox | Where-Object { $_.Subject -like '`[Test data`]*' }) | Should -BeNullOrEmpty
        ($script:State.Inbox | Where-Object Subject -eq 'Budget 2027') | Should -Not -BeNullOrEmpty
        ($script:State.Folders | Where-Object { $_.Name -in 'Projects', 'Migration', 'Archive 2026' }) | Should -BeNullOrEmpty
    }

    It 'SeedData through Graph, and blocked without AllowWrite' {
        $r = Invoke-Sim 'SeedData' @{ Protocol = 'Graph'; Authority = 'EntraID'; AllowWrite = $true } -State @{ GraphScopes = 'Mail.ReadWrite Mail.Send Calendars.ReadWrite' }
        $r.Status | Should -Be 'Passed'
        $script:State.Calendar.Count | Should -Be 7
        ($r.Trace | Where-Object Operation -eq 'POST events' | Select-Object -First 1).Request | Should -Match '"showAs":\s*"busy"'
        $r = Invoke-Sim 'SeedData'
        $r.Status | Should -Be 'Blocked'
    }
}

Describe 'Report' {
    AfterAll { Remove-Module WebServicesClient -ErrorAction SilentlyContinue }

    It 'writes the CSV, JSON and HTML files without a token, a secret or a formula' {
        $r = Invoke-Sim 'ReadMail' @{ Authority = 'EntraID'; Context = 'Application'; AppClientId = $script:App } -ClientSecret $script:Secret
        $report = Export-WscReport -Result $r -OutputPath (Join-Path $script:RepoRoot 'artifacts\test-reports')
        foreach ($k in 'Steps', 'Folders', 'Messages', 'FreeBusy', 'Actions', 'Trace', 'Summary', 'Html') { Test-Path $report.Files[$k] | Should -BeTrue }
        $all = (Get-ChildItem $report.Directory -File | ForEach-Object { Get-Content $_.FullName -Raw }) -join "`n"
        foreach ($t in $script:State.ValidTokens) { $all.Contains($t) | Should -BeFalse }
        $all | Should -Not -Match 'Sim-Secret!'
        (Get-Content $report.Files.Messages -Raw) | Should -Match """'=HYPERLINK"
        (Get-Content $report.Files.Html -Raw) | Should -Not -Match '\{\{[A-Z_]+\}\}'
    }
}

Describe 'Window' {
    BeforeAll { Import-Module (Join-Path $script:Root 'WebServicesClient.psd1') -Force }

    It 'shows the fields of the chosen sign-in and blocks a run without the password' {
        $f = New-WscTestForm -Configuration (Import-WscConfiguration) -Theme Light
        $c = $f.Controls
        $c.Authentication.SelectedIndex = 3
        $c.WindowsPanel.Visibility | Should -Be 'Visible'
        $c.PasswordPanel.Visibility | Should -Be 'Collapsed'
        $c.CurrentAccount.IsChecked = $false
        & (Get-Module WebServicesClient) { Update-WscGuiScenario }
        $c.PasswordPanel.Visibility | Should -Be 'Visible'
        $c.Authentication.SelectedIndex = 1; $c.Context.SelectedIndex = 2
        $c.AppPanel.Visibility | Should -Be 'Visible'
        $c.SignInUser.IsEnabled | Should -BeFalse
        $c.TestType.SelectedItem = 'MailCycle'
        $c.Warning.Text | Should -Match 'blocked until'
        $c.Authentication.SelectedIndex = 2; $c.TestType.SelectedItem = 'ReadMail'
        $c.Run.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
        ($f.Lines -join "`n") | Should -Match 'Enter the password'
        $f.Form.Close()
    }
}
