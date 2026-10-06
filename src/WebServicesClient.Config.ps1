<#
.SYNOPSIS
    Web Services Client for Exchange - configuration and scenarios (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    The configuration file has sections (Target, Identity, Ews, Test, Report, Logging), like the other
    tools. It is flattened into one settings hashtable used by the CLI, the GUI and the tests;
    unknown sections or keys and invalid values are all reported at once. Secrets are never part of
    it: passwords and client secrets are asked at run time and stay in memory.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

# Section.Key of the configuration file -> key of the settings hashtable.
$script:ConfigSchema = [ordered]@{
    Target   = [ordered]@{ Mailbox = 'Mailbox'; SignInUser = 'SignInUser'; Protocol = 'Protocol'; Discovery = 'Discovery'; EwsUrl = 'EwsUrl'; AdfsUrl = 'AdfsUrl'; Authority = 'Authority'; TenantId = 'TenantId' }
    Identity = [ordered]@{
        Authentication = 'Authentication'; Context = 'Context'; ClientId = 'ClientId'; AppClientId = 'AppClientId'; RedirectUri = 'RedirectUri'
        CertificateThumbprint = 'CertificateThumbprint'; WindowsPackage = 'WindowsPackage'; Access = 'Access'
    }
    Ews      = [ordered]@{ RequestServerVersion = 'RequestServerVersion'; UserAgent = 'UserAgent'; PreferServerAffinity = 'PreferServerAffinity'; MaxRetries = 'MaxRetries' }
    Test     = [ordered]@{
        DefaultType = 'TestType'; SignIn = 'SignIn'; AllowWrite = 'AllowWrite'; MessageCount = 'MessageCount'; FolderName = 'FolderName'; Recipient = 'Recipient'
        DeleteMode = 'DeleteMode'; DeliveryWaitSeconds = 'DeliveryWaitSeconds'; FreeBusyMailboxes = 'FreeBusyMailboxes'; FreeBusyDays = 'FreeBusyDays'
        FreeBusyIntervalMinutes = 'FreeBusyIntervalMinutes'; OAuthPollTimeoutSeconds = 'OAuthPollTimeoutSeconds'; HttpTimeoutSeconds = 'HttpTimeoutSeconds'
        CertificateWarningDays = 'CertificateWarningDays'
    }
    Report   = [ordered]@{ OutputPath = 'OutputPath'; FilePrefix = 'ReportPrefix'; Formats = 'ReportFormats'; CsvDelimiter = 'CsvDelimiter' }
    Logging  = [ordered]@{ Path = 'LogPath'; RetentionDays = 'LogRetentionDays' }
}

# Entra ID: the same flows for Exchange on-premises with hybrid modern authentication (HMA) and Exchange Online.
$script:Entra = @{
    LoginHost   = 'login.microsoftonline.com'
    Hosts       = @('login.microsoftonline.com', 'login.windows.net', 'login.microsoft.com', 'sts.windows.net', 'login.microsoftonline.us', 'login.partner.microsoftonline.cn', 'login.chinacloudapi.cn')
    EvoStsId    = '00000001-0000-0000-c000-000000000000'
    # Office 365 Exchange Online: the on-premises URLs are its service principal names (HMA).
    ExchangeApp = '00000002-0000-0ff1-ce00-000000000000'
    OnlineHosts = @('outlook.office365.com', 'outlook.office.com', 'outlook.office365.us', 'outlook-dod.office365.us', 'partner.outlook.cn')
}

# Redirect URIs of the sign-in window when Identity.RedirectUri is empty: the native-client page of Entra ID
# (Microsoft Office d3590ed6...), urn:ietf:wg:oauth:2.0:oob for AD FS (application group of the Exchange documentation).
$script:SignInRedirect = @{
    EntraID = 'https://login.microsoftonline.com/common/oauth2/nativeclient'
    ADFS    = 'urn:ietf:wg:oauth:2.0:oob'
}

# Prefix of the subject of the test messages: the tool replies to, moves and deletes only these, unless an item is named.
$script:TestSubjectPrefix = '[Web Services Client for Exchange]'

# Scenarios and the stages they run, in order. A stage runs only if no earlier stage failed or was blocked.
# Writes: the scenario changes the mailbox and needs AllowWrite.
$script:Scenarios = @(
    @{ Name = 'Discovery'; DisplayName = 'Prerequisites without sign-in'; Stages = @('Discovery')
       Description = 'EWS URL (Autodiscover or manual), TLS certificate, authentication offered by EWS, authorization server (AD FS or Entra ID), OAuth offered to the mailbox, forged token or wrong password refused. No sign-in, nothing changed.' }
    @{ Name = 'SignIn'; DisplayName = 'Sign-in'; Stages = @('SignIn')
       Description = 'OAuth sign-in (user, delegated application or application) and the claims of the token, or the Basic or Windows sign-in (NTLM, Kerberos) with every leg of the handshake.' }
    @{ Name = 'Endpoint'; DisplayName = 'EWS endpoint'; Stages = @('SignIn', 'Endpoint')
       Description = 'Sign-in, then a first EWS request on the mailbox: Exchange version, front-end and back-end servers, affinity, access to the mailbox.' }
    @{ Name = 'Folders'; DisplayName = 'List the folders'; Stages = @('SignIn', 'Endpoint', 'Folders')
       Description = 'Folder tree of the mailbox (FindFolder): name, path, class, item and unread counts.' }
    @{ Name = 'ReadMail'; DisplayName = 'Read messages'; Stages = @('SignIn', 'Endpoint', 'ReadMail')
       Description = 'The most recent Inbox messages (FindItem), then one message in full (GetItem): sender, recipients, body preview.' }
    @{ Name = 'FreeBusy'; DisplayName = 'Free/busy'; Stages = @('SignIn', 'Endpoint', 'FreeBusy')
       Description = 'Free/busy of one or several mailboxes (GetUserAvailability): working hours, busy periods, merged view.' }
    @{ Name = 'CreateFolder'; DisplayName = 'Create a folder'; Stages = @('SignIn', 'Endpoint', 'CreateFolder'); Writes = $true
       Description = 'Creates the test folder under the Inbox (CreateFolder), or finds it if it exists.' }
    @{ Name = 'SendMail'; DisplayName = 'Send a message'; Stages = @('SignIn', 'Endpoint', 'SendMail'); Writes = $true
       Description = 'Sends a test message (CreateItem, SendAndSaveCopy) to the recipient (the mailbox itself by default) and waits for it in the Inbox.' }
    @{ Name = 'ReplyMail'; DisplayName = 'Reply to a message'; Stages = @('SignIn', 'Endpoint', 'ReplyMail'); Writes = $true
       Description = 'Replies to a given message (-ItemId, -ItemSubject) or to the last test message of the tool (ReplyToItem).' }
    @{ Name = 'MoveMail'; DisplayName = 'Move a message'; Stages = @('SignIn', 'Endpoint', 'MoveMail'); Writes = $true
       Description = 'Moves a given message or the last test message of the tool to the test folder (MoveItem).' }
    @{ Name = 'DeleteMail'; DisplayName = 'Delete a message'; Stages = @('SignIn', 'Endpoint', 'DeleteMail'); Writes = $true
       Description = 'Deletes a given message or the last test message of the tool (DeleteItem, Test.DeleteMode).' }
    @{ Name = 'SeedData'; DisplayName = 'Fill the test mailbox'; Stages = @('SignIn', 'Endpoint', 'SeedData'); Writes = $true
       Description = 'Creates test data so that the read-only scenarios show something: folders Projects, Projects\Migration and Archive 2026, six messages sent to the mailbox (one filed), seven calendar items over the next five working days (busy, tentative, out of office, working elsewhere). Subjects start with [Test data]; nothing is created twice.' }
    @{ Name = 'CleanData'; DisplayName = 'Remove the test data'; Stages = @('SignIn', 'Endpoint', 'CleanData'); Writes = $true
       Description = 'Deletes what SeedData created: [Test data] messages of the Inbox, [Test data] calendar items, the test folders (moved to Deleted Items).' }
    @{ Name = 'ReadOnly'; DisplayName = 'Read-only diagnostic'; Stages = @('Discovery', 'SignIn', 'Endpoint', 'Folders', 'ReadMail', 'FreeBusy')
       Description = 'Every check that changes nothing: prerequisites, sign-in, endpoint, folders, messages and free/busy.' }
    @{ Name = 'MailCycle'; DisplayName = 'Mail cycle'; Stages = @('SignIn', 'Endpoint', 'CreateFolder', 'SendMail', 'ReplyMail', 'MoveMail', 'DeleteMail'); Writes = $true
       Description = 'Creates the test folder, sends a test message, replies to it, moves it to the test folder, then deletes it: only items the tool created.' }
    @{ Name = 'Full'; DisplayName = 'Complete diagnostic'; Stages = @('Discovery', 'SignIn', 'Endpoint', 'Folders', 'ReadMail', 'FreeBusy', 'CreateFolder', 'SendMail', 'ReplyMail', 'MoveMail', 'DeleteMail'); Writes = $true
       Description = 'Everything in order: the read-only diagnostic, then the mail cycle.' }
)

$script:StageInfo = [ordered]@{
    Autodiscover = @{ Title = 'Autodiscover: where is EWS'; Icon = 'Search' }
    Discovery    = @{ Title = 'Prerequisites without sign-in'; Icon = 'Search' }
    SignIn       = @{ Title = 'Sign-in'; Icon = 'Key' }
    Endpoint     = @{ Title = 'EWS endpoint and mailbox'; GraphTitle = 'Microsoft Graph and mailbox'; Icon = 'Server' }
    Folders      = @{ Title = 'Folders (FindFolder)'; GraphTitle = 'Folders (mailFolders)'; Icon = 'Folder' }
    ReadMail     = @{ Title = 'Read messages (FindItem, GetItem)'; GraphTitle = 'Read messages (messages)'; Icon = 'Mail' }
    FreeBusy     = @{ Title = 'Free/busy (GetUserAvailability)'; GraphTitle = 'Free/busy (getSchedule)'; Icon = 'Clock' }
    CreateFolder = @{ Title = 'Create the test folder (CreateFolder)'; GraphTitle = 'Create the test folder (childFolders)'; Icon = 'Folder' }
    SendMail     = @{ Title = 'Send a test message (CreateItem)'; GraphTitle = 'Send a test message (sendMail)'; Icon = 'Mail' }
    ReplyMail    = @{ Title = 'Reply to a message (ReplyToItem)'; GraphTitle = 'Reply to a message (reply)'; Icon = 'Mail' }
    MoveMail     = @{ Title = 'Move a message (MoveItem)'; GraphTitle = 'Move a message (move)'; Icon = 'Folder' }
    SeedData     = @{ Title = 'Test data (folders, messages, calendar)'; GraphTitle = 'Test data (folders, messages, calendar)'; Icon = 'Mail' }
    CleanData    = @{ Title = 'Remove the test data'; GraphTitle = 'Remove the test data'; Icon = 'Block' }
    DeleteMail   = @{ Title = 'Delete a message (DeleteItem)'; GraphTitle = 'Delete a message (delete)'; Icon = 'Block' }
}

function Get-WscTestCatalog {
    <# The scenarios: name, description, stages, whether they sign in and whether they change the mailbox. #>
    [CmdletBinding()]
    param()

    foreach ($s in $script:Scenarios) {
        [pscustomobject]@{
            Name        = $s.Name
            DisplayName = $s.DisplayName
            Description = $s.Description
            Stages      = @($s.Stages)
            SignIn      = $s.Stages -contains 'SignIn'
            Writes      = [bool]($s.ContainsKey('Writes') -and $s.Writes)
        }
    }
}

function Get-WscScenarioStages {
    <# Stages a scenario runs: Autodiscover first when the EWS URL is found with Autodiscover. #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$TestType, [string]$Discovery = 'Autodiscover')

    $scenario = $script:Scenarios | Where-Object { $_.Name -eq $TestType } | Select-Object -First 1
    if (-not $scenario) { return }
    if ($Discovery -eq 'Autodiscover') { 'Autodiscover' }
    foreach ($stage in $scenario.Stages) { $stage }
}

function Get-WscStageTitle {
    <# Title of a stage in the console, the window and the report; the sign-in depends on the authentication. #>
    param([Parameter(Mandatory = $true)][string]$Stage, [hashtable]$Configuration = @{})

    if ($Stage -ne 'SignIn') {
        $info = $script:StageInfo[$Stage]
        if ([string]$Configuration['ProtocolUsed'] -eq 'Graph' -and $info.ContainsKey('GraphTitle')) { return $info.GraphTitle }
        return $info.Title
    }
    switch ([string]$Configuration.Authentication) {
        'Basic' { return 'Basic sign-in (user name and password)' }
        'Windows' { return "Windows sign-in ($([string]$Configuration.WindowsPackage))" }
    }
    $server = switch ([string]$Configuration.Authority) { 'EntraID' { 'Entra ID' } 'Auto' { 'OAuth' } default { 'AD FS' } }
    switch ([string]$Configuration.Context) {
        'Application' { return "$server sign-in of the application (client credentials)" }
        'Delegated' { return "$server sign-in through the application (delegated)" }
        default { return "$server sign-in of the user" }
    }
}

function Get-WscDefaultConfiguration {
    @{
        Mailbox                 = 'ews-test@contoso.test'
        SignInUser              = ''
        Protocol                = 'Auto'
        ProtocolUsed            = ''
        Discovery               = 'Autodiscover'
        EwsUrl                  = 'https://mail.contoso.test/EWS/Exchange.asmx'
        AdfsUrl                 = 'https://adfs.contoso.test/adfs'
        Authority               = 'ADFS'
        TenantId                = ''
        Authentication          = 'OAuth'
        Context                 = 'User'
        ClientId                = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
        AppClientId             = ''
        RedirectUri             = ''
        CertificateThumbprint   = ''
        WindowsPackage          = 'Negotiate'
        Access                  = 'Auto'
        RequestServerVersion    = 'Exchange2016'
        UserAgent               = 'WebServicesClient/1.0'
        PreferServerAffinity    = $true
        MaxRetries              = 2
        TestType                = 'ReadOnly'
        SignIn                  = 'Auto'
        AllowWrite              = $false
        MessageCount            = 10
        FolderName              = 'Web Services Client for Exchange'
        Recipient               = ''
        DeleteMode              = 'MoveToDeletedItems'
        DeliveryWaitSeconds     = 60
        FreeBusyMailboxes       = @()
        FreeBusyDays            = 7
        FreeBusyIntervalMinutes = 30
        OAuthPollTimeoutSeconds = 600
        HttpTimeoutSeconds      = 60
        CertificateWarningDays  = 30
        OutputPath              = '.\reports'
        ReportPrefix            = 'WebServicesClient'
        ReportFormats           = @('Csv', 'Html')
        CsvDelimiter            = ';'
        LogPath                 = '.\logs'
        LogRetentionDays        = 14
    }
}

function Resolve-WscPath {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Root)
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath($Path, $Root)
}

function Test-WscExchangeOnlineUrl {
    <# Whether a URL is one of Exchange Online (outlook.office365.com and the other clouds). #>
    param([AllowEmptyString()][AllowNull()][string]$Url)

    $uri = $null
    if (-not $Url -or -not [Uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$uri)) { return $false }
    return $script:Entra.OnlineHosts -contains $uri.Host.ToLowerInvariant()
}

function Resolve-WscAccess {
    <#
        How the mailbox is opened: Self, Delegate or Impersonation. Auto: Impersonation for an application,
        Delegate when another account signs in, Self otherwise.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Configuration)

    $access = [string]$Configuration.Access
    if ($access -and $access -ne 'Auto') { return $access }
    if ([string]$Configuration.Context -eq 'Application') { return 'Impersonation' }
    $user = [string]$Configuration.SignInUser
    if ($user -and $user -ine [string]$Configuration.Mailbox) { return 'Delegate' }
    return 'Self'
}

function Test-WscConfiguration {
    <# Checks a settings hashtable (flattened configuration) and lists every problem. #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Configuration)

    $problems = [Collections.Generic.List[string]]::new()
    $c = $Configuration
    $auth = [string]$c.Authentication
    $oauth = $auth -eq 'OAuth'
    $context = [string]$c.Context
    $manual = [string]$c.Discovery -eq 'Manual'
    $guid = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
    $number = {
        param([string]$Key, [string]$Label, [int]$Min, [int]$Max)
        $n = 0
        if (-not [int]::TryParse([string]$c[$Key], [ref]$n) -or $n -lt $Min -or $n -gt $Max) { [void]$problems.Add("$Label must be a whole number between $Min and $Max.") }
    }
    $address = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

    if ([string]::IsNullOrWhiteSpace([string]$c.Mailbox)) { [void]$problems.Add('Target.Mailbox is required.') }
    elseif ([string]$c.Mailbox -notmatch $address) { [void]$problems.Add('Target.Mailbox must be an SMTP address (user@domain).') }
    if ([string]$c.SignInUser -and [string]$c.SignInUser -notmatch '^([^@\s:\\]+@[^@\s:\\]+|[^@\s:\\]+\\[^@\s:\\]+)$') {
        [void]$problems.Add('Target.SignInUser must be empty (= Target.Mailbox), a UPN (user@domain) or DOMAIN\user.')
    }
    if ([string]$c.Discovery -notin 'Autodiscover', 'Manual') { [void]$problems.Add("Target.Discovery must be 'Autodiscover' or 'Manual'.") }
    $ews = [string]$c.EwsUrl
    $protocol = [string]$c.Protocol
    if ($protocol -notin 'Auto', 'EWS', 'Graph') { [void]$problems.Add("Target.Protocol must be 'Auto', 'EWS' or 'Graph'.") }
    if ($manual -and -not $ews -and $protocol -ne 'Graph') { [void]$problems.Add('Target.EwsUrl is required with Target.Discovery Manual.') }
    if ($protocol -eq 'Graph' -and -not $oauth) { [void]$problems.Add('Microsoft Graph accepts only OAuth with Entra ID: set Identity.Authentication to OAuth, or Target.Protocol to EWS or Auto.') }
    if ($protocol -eq 'Graph' -and $oauth -and [string]$c.Authority -eq 'ADFS') { [void]$problems.Add('Microsoft Graph accepts only Entra ID tokens: set Target.Authority to EntraID (or Auto).') }
    if ($ews) {
        $uri = $null
        if (-not [Uri]::TryCreate($ews, [UriKind]::Absolute, [ref]$uri)) { [void]$problems.Add('Target.EwsUrl is not an absolute URL.') }
        else {
            if ($uri.Scheme -ne 'https') { [void]$problems.Add('Target.EwsUrl must be an https:// URL.') }
            if ($uri.AbsolutePath -notmatch '^/EWS/Exchange\.asmx$') { [void]$problems.Add('Target.EwsUrl must target /EWS/Exchange.asmx.') }
        }
    }
    if ($auth -notin 'OAuth', 'Basic', 'Windows') { [void]$problems.Add("Identity.Authentication must be 'OAuth', 'Basic' or 'Windows'.") }
    if ([string]$c.Authority -notin 'ADFS', 'EntraID', 'Auto') { [void]$problems.Add("Target.Authority must be 'ADFS', 'EntraID' or 'Auto'.") }
    if ($context -notin 'User', 'Delegated', 'Application') { [void]$problems.Add("Identity.Context must be 'User', 'Delegated' or 'Application'.") }
    if ([string]$c.Access -notin 'Auto', 'Self', 'Delegate', 'Impersonation') { [void]$problems.Add("Identity.Access must be 'Auto', 'Self', 'Delegate' or 'Impersonation'.") }
    if ([string]$c.WindowsPackage -notin 'Negotiate', 'NTLM', 'Kerberos') { [void]$problems.Add("Identity.WindowsPackage must be 'Negotiate', 'NTLM' or 'Kerberos'.") }
    if ([string]$c.SignIn -notin 'Auto', 'Window', 'DeviceCode') { [void]$problems.Add("Test.SignIn must be 'Auto', 'Window' or 'DeviceCode'.") }
    if ($oauth -and [string]$c.Authority -eq 'ADFS') {
        $adfs = ([string]$c.AdfsUrl).TrimEnd('/')
        if (-not $adfs) { [void]$problems.Add('Target.AdfsUrl is required with Target.Authority ADFS.') }
        elseif ($adfs -notmatch '^https://[^/\s]+/' -or -not $adfs.EndsWith('/adfs', [StringComparison]::OrdinalIgnoreCase)) {
            [void]$problems.Add('Target.AdfsUrl must be an https:// URL ending with /adfs (for example https://adfs.contoso.com/adfs).')
        }
    }
    if ([string]$c.TenantId -and [string]$c.TenantId -notmatch '^([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}|[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+)$') {
        [void]$problems.Add('Target.TenantId must be empty (= the domain of the mailbox), a tenant ID (GUID) or a domain of the tenant.')
    }
    if ($oauth -and $context -eq 'User' -and [string]$c.ClientId -notmatch $guid) { [void]$problems.Add('Identity.ClientId must be a client ID (GUID).') }
    if ($oauth -and $context -in 'Delegated', 'Application' -and [string]$c.AppClientId -notmatch $guid) {
        [void]$problems.Add("Identity.AppClientId (the client ID of your application) is required by the $context context.")
    }
    if (-not $oauth -and $context -ne 'User') { [void]$problems.Add("The $context context uses OAuth: Basic and Windows sign in as a user (Identity.Context User).") }
    if ($oauth -and $context -eq 'Application' -and [string]$c.Access -in 'Self', 'Delegate') {
        [void]$problems.Add('An application has no mailbox of its own: with the Application context, Identity.Access must be Auto or Impersonation.')
    }
    if ([string]$c.CertificateThumbprint -and [string]$c.CertificateThumbprint -notmatch '^[0-9A-Fa-f]{40}$') { [void]$problems.Add('Identity.CertificateThumbprint must be empty or a SHA-1 thumbprint (40 hexadecimal characters).') }
    if ([string]$c.RedirectUri -and [string]$c.RedirectUri -notmatch '^[A-Za-z][A-Za-z0-9+.-]*:\S+$') { [void]$problems.Add('Identity.RedirectUri must be empty or an absolute URI.') }
    $online = Test-WscExchangeOnlineUrl -Url $ews
    if ($online -and $manual -and $auth -eq 'Basic' -and @($script:Scenarios | Where-Object { $_.Name -eq [string]$c.TestType -and $_.Stages -contains 'SignIn' }).Count) {
        [void]$problems.Add('Exchange Online no longer accepts Basic authentication for EWS: use OAuth with Entra ID (Target.Authority EntraID). Discovery with Basic shows what Exchange Online offers, without a password.')
    }
    if ($online -and $manual -and $oauth -and [string]$c.Authority -eq 'ADFS') {
        [void]$problems.Add('Exchange Online accepts only Entra ID tokens: with this Target.EwsUrl set Target.Authority to EntraID (or Auto).')
    }
    if ($online -and $manual -and $auth -eq 'Windows') { [void]$problems.Add('Exchange Online does not offer Windows authentication (NTLM, Kerberos): use OAuth with Entra ID.') }
    if ([string]$c.RequestServerVersion -notmatch '^(Exchange2007_SP1|Exchange2010(_SP[12])?|Exchange2013(_SP1)?|Exchange2016|V2015_10_05|V2016_01_06|V2017_04_14|V2018_01_08)$') {
        [void]$problems.Add('Ews.RequestServerVersion must be an EWS schema version, for example Exchange2013_SP1 or Exchange2016.')
    }
    if ([string]::IsNullOrWhiteSpace([string]$c.UserAgent) -or [string]$c.UserAgent -match '[\r\n]') { [void]$problems.Add('Ews.UserAgent is required, on one line.') }
    if ($c.PreferServerAffinity -isnot [bool]) { [void]$problems.Add('Ews.PreferServerAffinity must be $true or $false.') }
    if ($c.AllowWrite -isnot [bool]) { [void]$problems.Add('Test.AllowWrite must be $true or $false.') }
    & $number 'MaxRetries' 'Ews.MaxRetries' 0 5
    if ([string]$c.TestType -notin @($script:Scenarios.Name)) { [void]$problems.Add("Test.DefaultType must be one of: $($script:Scenarios.Name -join ', ').") }
    & $number 'MessageCount' 'Test.MessageCount' 1 100
    & $number 'DeliveryWaitSeconds' 'Test.DeliveryWaitSeconds' 0 600
    & $number 'FreeBusyDays' 'Test.FreeBusyDays' 1 42
    & $number 'FreeBusyIntervalMinutes' 'Test.FreeBusyIntervalMinutes' 5 1440
    & $number 'OAuthPollTimeoutSeconds' 'Test.OAuthPollTimeoutSeconds' 30 3600
    & $number 'HttpTimeoutSeconds' 'Test.HttpTimeoutSeconds' 5 300
    & $number 'CertificateWarningDays' 'Test.CertificateWarningDays' 0 365
    & $number 'LogRetentionDays' 'Logging.RetentionDays' 1 365
    if ([string]::IsNullOrWhiteSpace([string]$c.FolderName) -or [string]$c.FolderName -match '[\\/\r\n]' -or ([string]$c.FolderName).Length -gt 100) { [void]$problems.Add('Test.FolderName is required: one line, no \ or /, 100 characters at most.') }
    if ([string]$c.Recipient -and [string]$c.Recipient -notmatch $address) { [void]$problems.Add('Test.Recipient must be empty (= the mailbox) or an SMTP address.') }
    if ([string]$c.DeleteMode -notin 'MoveToDeletedItems', 'SoftDelete', 'HardDelete') { [void]$problems.Add("Test.DeleteMode must be 'MoveToDeletedItems', 'SoftDelete' or 'HardDelete'.") }
    $freeBusy = @($c.FreeBusyMailboxes | Where-Object { $_ })
    if ($freeBusy.Count -gt 100) { [void]$problems.Add('Test.FreeBusyMailboxes: 100 mailboxes at most (limit of GetUserAvailability).') }
    foreach ($m in $freeBusy) { if ([string]$m -notmatch $address) { [void]$problems.Add("Test.FreeBusyMailboxes: '$m' is not an SMTP address.") } }
    $formats = @($c.ReportFormats)
    if ($formats.Count -eq 0 -or @($formats | Where-Object { $_ -notin 'Csv', 'Html' }).Count) { [void]$problems.Add("Report.Formats must contain 'Csv', 'Html' or both.") }
    if ([string]$c.CsvDelimiter -notin ';', ',', "`t") { [void]$problems.Add("Report.CsvDelimiter must be ';', ',' or a tab.") }
    if ([string]::IsNullOrWhiteSpace([string]$c.OutputPath)) { [void]$problems.Add('Report.OutputPath is required.') }
    if ([string]::IsNullOrWhiteSpace([string]$c.ReportPrefix)) { [void]$problems.Add('Report.FilePrefix is required.') }
    if ([string]::IsNullOrWhiteSpace([string]$c.LogPath)) { [void]$problems.Add('Logging.Path is required.') }

    [pscustomobject]@{ IsValid = $problems.Count -eq 0; Problems = @($problems) }
}

function Import-WscConfiguration {
    <#
        Reads config\WebServicesClient.config.psd1 (sections), applies the defaults, resolves the relative
        paths from the tool folder, checks everything and returns the settings hashtable.
    #>
    [CmdletBinding()]
    param(
        [string]$Path = (Join-Path $script:ToolRoot 'config\WebServicesClient.config.psd1'),
        [string]$Root = $script:ToolRoot
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Configuration file not found: $Path" }
    $settings = Get-WscDefaultConfiguration
    $problems = [Collections.Generic.List[string]]::new()
    $data = Import-PowerShellDataFile -LiteralPath $Path
    foreach ($section in $data.Keys) {
        if (-not $script:ConfigSchema.Contains($section)) {
            [void]$problems.Add("Unknown section '$section'. Sections: $($script:ConfigSchema.Keys -join ', ').")
            continue
        }
        if ($data[$section] -isnot [hashtable]) { [void]$problems.Add("Section '$section' must be a @{ } block."); continue }
        foreach ($key in $data[$section].Keys) {
            if (-not $script:ConfigSchema[$section].Contains($key)) {
                [void]$problems.Add("Unknown key '$section.$key'. Keys of $($section): $($script:ConfigSchema[$section].Keys -join ', ').")
                continue
            }
            $settings[$script:ConfigSchema[$section][$key]] = $data[$section][$key]
        }
    }
    $settings.FreeBusyMailboxes = @($settings.FreeBusyMailboxes | Where-Object { $_ })
    foreach ($key in 'OutputPath', 'LogPath') {
        if (-not [string]::IsNullOrWhiteSpace([string]$settings[$key])) { $settings[$key] = Resolve-WscPath -Path ([string]$settings[$key]) -Root $Root }
    }
    $settings.ConfigPath = [IO.Path]::GetFullPath($Path)
    foreach ($p in (Test-WscConfiguration -Configuration $settings).Problems) {
        # Checked again once the context is known (command line, window): the client ID of an application is given there.
        if ($p -like 'Identity.AppClientId*') { continue }
        [void]$problems.Add($p)
    }
    if ($problems.Count) { throw ("Invalid configuration ($Path):`n - " + ($problems -join "`n - ")) }
    return $settings
}

function Resolve-WscEndpoints {
    <#
        URLs derived from the configuration and the EWS URL found. Authority EntraID: the endpoints of the
        tenant (Microsoft identity platform v2.0). Scope: the Exchange resource and the permission of the
        context - EWS.AccessAsUser.All for a user or a delegated application, .default for an application.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Configuration)

    $authority = if ([string]$Configuration['Authority']) { [string]$Configuration['Authority'] } else { 'ADFS' }
    $tenant = [string]$Configuration['TenantId']
    $adfsRoot = ([string]$Configuration['AdfsUrl']).TrimEnd('/')
    $adfsUri = if ($adfsRoot) { [Uri]$adfsRoot } else { $null }
    $ewsUrl = [string]$Configuration['EwsUrl']
    $ewsUri = $null
    if ($ewsUrl -and -not [Uri]::TryCreate($ewsUrl, [UriKind]::Absolute, [ref]$ewsUri)) { $ewsUri = $null }
    $resource = if ($ewsUri) { $ewsUri.GetLeftPart([UriPartial]::Authority) + '/' } else { $null }
    $entraRoot = if ($authority -eq 'EntraID' -and $tenant) { "https://$($script:Entra.LoginHost)/$tenant" } else { $null }
    $app = [string]$Configuration['Context'] -eq 'Application'
    $graph = [string]$Configuration['ProtocolUsed'] -eq 'Graph'
    if ($graph) {
        # Microsoft Graph: the resource of the token is Graph, the permissions are those granted to the client (.default).
        $authority = 'EntraID'
        $resource = 'https://graph.microsoft.com/'
        $entraRoot = if ($tenant) { "https://$($script:Entra.LoginHost)/$tenant" } else { $null }
    }
    $scope = if (-not $resource) { $null }
    elseif ($graph) { 'https://graph.microsoft.com/.default' }
    elseif ($authority -eq 'EntraID') { if ($app) { "$($resource).default" } else { "$($resource)EWS.AccessAsUser.All" } }
    else {
        # AD FS: the Web API identifier ends with '/', and the resource-qualified scope adds another '/'.
        if ($app) { 'openid' } else { "openid $($resource)/EWS.AccessAsUser.All" }
    }
    [pscustomobject]@{
        Authority          = $authority
        TenantId           = $tenant
        AdfsRoot           = $adfsRoot
        EntraRoot          = $entraRoot
        Protocol           = if ($graph) { 'Graph' } else { 'EWS' }
        GraphUrl           = if ($graph) { 'https://graph.microsoft.com/v1.0' } else { $null }
        EwsUrl             = if ($ewsUri) { $ewsUri.AbsoluteUri } else { $null }
        ExchangeOnline     = Test-WscExchangeOnlineUrl -Url $ewsUrl
        Resource           = $resource
        Scope              = $scope
        AuthorizeEndpoint  = if ($entraRoot) { "$entraRoot/oauth2/v2.0/authorize" } elseif ($adfsRoot) { "$adfsRoot/oauth2/authorize" } else { $null }
        DeviceCodeEndpoint = if ($entraRoot) { "$entraRoot/oauth2/v2.0/devicecode" } elseif ($adfsRoot) { "$adfsRoot/oauth2/devicecode" } else { $null }
        TokenEndpoint      = if ($entraRoot) { "$entraRoot/oauth2/v2.0/token" } elseif ($adfsRoot) { "$adfsRoot/oauth2/token" } else { $null }
        MetadataEndpoint   = if ($entraRoot) { "$entraRoot/v2.0/.well-known/openid-configuration" } elseif ($adfsRoot) { "$adfsRoot/.well-known/openid-configuration" } else { $null }
        AdfsHost           = if ($adfsUri) { $adfsUri.Host } else { $null }
        AdfsPort           = if ($adfsUri) { $adfsUri.Port } else { $null }
        EwsHost            = if ($ewsUri) { $ewsUri.Host } else { $null }
        EwsPort            = if ($ewsUri) { $ewsUri.Port } else { $null }
    }
}
