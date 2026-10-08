<#
.SYNOPSIS
    Web Services Client for Exchange - test toolbox for Exchange mailboxes through EWS (on-premises) and Microsoft Graph (Exchange Online), step by step, with the HTTP trace of every request.

.DESCRIPTION
    Runs one scenario and tells exactly which step works and which one does not:

      Discovery     no sign-in: EWS URL, TLS certificate, authentication offered, authorization server,
                    OAuth offered to the mailbox, NTLM challenge and Kerberos realm, forged token refused
      SignIn        OAuth (user, delegated application, application), Basic or Windows (NTLM, Kerberos)
      Endpoint      + first EWS request on the mailbox: version, servers, routing and affinity
      Folders       + the folder tree (FindFolder)
      ReadMail      + the most recent Inbox messages (FindItem) and one message in full (GetItem)
      FreeBusy      + free/busy of one or several mailboxes (GetUserAvailability)
      CreateFolder  + the test folder under the Inbox                                   (writes)
      SendMail      + a test message, then its delivery to the Inbox                       (writes)
      ReplyMail     + a reply to a given message or to the test message                    (writes)
      MoveMail      + a given message or the test message moved to the test folder         (writes)
      DeleteMail    + a given message or the test message deleted                          (writes)
      ReadOnly      Discovery, sign-in, endpoint, folders, messages, free/busy
      MailCycle     create the folder, send, reply, move, delete: only what the tool created  (writes)
      Full          ReadOnly, then MailCycle                                                (writes)
      SeedData      fills a test mailbox: folders, 6 messages, 7 calendar items             (writes)
      CleanData     removes what SeedData created                                           (writes)

    Protocol: EWS for a mailbox on-premises, Microsoft Graph for a mailbox in Exchange Online (Auto), where EWS
    is being retired. The scenarios, the checks and the report are the same.

    Authentication: OAuth with AD FS (Exchange 2019 CU13+ / SE), with Entra ID (hybrid modern authentication
    or Exchange Online), Basic, or Windows (Negotiate, NTLM, Kerberos: the handshake is traced leg by leg).
    Context (OAuth): User, Delegated (your application on behalf of the user) or Application (client
    credentials, impersonation). The EWS URL comes from Autodiscover or is given (Discovery Manual).

    Writes CSV, JSON and HTML report files in a new folder, and a daily log file. Everything is set in
    config\WebServicesClient.config.psd1; the parameters below override it. Tokens, client secrets and
    passwords stay in memory and are never written anywhere.

.PARAMETER TestType
    Scenario. Default: Test.DefaultType.
.PARAMETER Mailbox
    The mailbox the tests work on (Target.Mailbox).
.PARAMETER SignInUser
    The account that signs in, when it is not the mailbox: delegate access or impersonation (Target.SignInUser).
.PARAMETER Protocol
    Auto (Microsoft Graph for a mailbox in Exchange Online, EWS on-premises), EWS or Graph (Target.Protocol).
    EWS is being retired in Exchange Online (disabled from October 2026, stopped in April 2027); Graph does not
    reach on-premises mailboxes and accepts only OAuth with Entra ID.
.PARAMETER Discovery
    Autodiscover (like Outlook) or Manual (the EWS URL given). Default: Target.Discovery.
.PARAMETER EwsUrl
    EWS URL, https://<server>/EWS/Exchange.asmx (Target.EwsUrl). With Autodiscover, used only when Autodiscover gives no answer.
.PARAMETER Authentication
    OAuth, Basic or Windows (Identity.Authentication).
.PARAMETER Authority
    OAuth: ADFS, EntraID (HMA or Exchange Online) or Auto (the server Exchange names) (Target.Authority).
.PARAMETER AdfsUrl
    AD FS root, ends with /adfs (Target.AdfsUrl).
.PARAMETER TenantId
    Entra ID: tenant ID or a domain of the tenant (Target.TenantId).
.PARAMETER Context
    OAuth: User, Delegated or Application (Identity.Context).
.PARAMETER AppClientId
    Delegated or Application: the client ID of your application (Identity.AppClientId).
.PARAMETER ClientId
    User context: the public client used (Identity.ClientId, Microsoft Office by default).
.PARAMETER RedirectUri
    Delegated: a redirect URI of your application (Identity.RedirectUri).
.PARAMETER CertificateThumbprint
    Application with Entra ID: certificate of the application (Identity.CertificateThumbprint).
.PARAMETER ClientSecret
    Application, or a confidential delegated application: the client secret. Without it (and without a
    certificate) the Application context asks for it. Never written anywhere.
.PARAMETER Access
    Auto, Self, Delegate or Impersonation (Identity.Access).
.PARAMETER WindowsPackage
    Windows: Negotiate, NTLM or Kerberos (Identity.WindowsPackage).
.PARAMETER Credential
    Basic: user name and password (asked when missing). Windows: the account to use (default: the
    current Windows account; give it on a computer out of the domain).
.PARAMETER SignIn
    OAuth user sign-in: Auto, Window or DeviceCode (Test.SignIn).
.PARAMETER AllowWrite
    Lets the writing steps run (create a folder, send, reply, move, delete). Test mailbox only.
.PARAMETER ItemId
    ReplyMail, MoveMail, DeleteMail: the EWS ID of the message (Messages tab of a ReadMail report).
.PARAMETER ItemSubject
    ReplyMail, MoveMail, DeleteMail: the most recent Inbox message whose subject contains this text.
.PARAMETER FreeBusyMailboxes
    FreeBusy: the mailboxes whose free/busy is read (Test.FreeBusyMailboxes; default: the mailbox).
.PARAMETER FreeBusyDays
    FreeBusy: number of days from today (Test.FreeBusyDays).
.PARAMETER Recipient
    SendMail: recipient of the test message (Test.Recipient; default: the mailbox).
.PARAMETER MessageCount
    ReadMail: messages listed (Test.MessageCount).
.PARAMETER Gui
    Opens the window: same scenarios, same report, progress shown live.
.PARAMETER OutputPath
    Overrides Report.OutputPath.
.PARAMETER NoReport
    No report file (console and log only).
.PARAMETER AccessToken
    Token already obtained (integration only): the sign-in is skipped. It can stay in the PowerShell history.
.PARAMETER ConfigPath
    Configuration file. Default: config\WebServicesClient.config.psd1.

.EXAMPLE
    .\Invoke-WebServicesClient.ps1 -TestType Discovery
    Prerequisites without any sign-in.

.EXAMPLE
    .\Invoke-WebServicesClient.ps1 -TestType ReadOnly -Mailbox ews-test@contoso.com -Authority ADFS
    Exchange on-premises with AD FS: Autodiscover, sign-in in a window, folders, messages, free/busy.

.EXAMPLE
    .\Invoke-WebServicesClient.ps1 -TestType MailCycle -Authority EntraID -Mailbox ews-test@contoso.com -AllowWrite
    Hybrid modern authentication (or Exchange Online, found by Autodiscover): create the test folder, send a
    message to the mailbox, reply to it, move it, delete it.

.EXAMPLE
    .\Invoke-WebServicesClient.ps1 -TestType FreeBusy -Discovery Manual -EwsUrl https://outlook.office365.com/EWS/Exchange.asmx -Authority EntraID -FreeBusyMailboxes room1@contoso.com,alice@contoso.com
    Free/busy of two mailboxes from Exchange Online.

.EXAMPLE
    .\Invoke-WebServicesClient.ps1 -TestType ReadMail -Context Application -Authority EntraID -AppClientId 11111111-2222-3333-4444-555555555555 -CertificateThumbprint 0123456789ABCDEF0123456789ABCDEF01234567 -Mailbox ews-test@contoso.com
    An application in its own context (client credentials with a certificate) that impersonates the mailbox.

.EXAMPLE
    .\Invoke-WebServicesClient.ps1 -TestType ReadMail -Context Delegated -Authority EntraID -AppClientId 11111111-2222-3333-4444-555555555555 -Mailbox shared@contoso.com -SignInUser alice@contoso.com
    Your application on behalf of Alice, who opens a shared mailbox with her permissions (delegate access).

.EXAMPLE
    .\Invoke-WebServicesClient.ps1 -TestType ReadOnly -Authentication Windows -WindowsPackage NTLM -Credential CONTOSO\ews-test -Discovery Manual -EwsUrl https://mail.contoso.com/EWS/Exchange.asmx
    NTLM from a computer out of the domain: the three legs of the handshake are in the report.

.EXAMPLE
    .\Invoke-WebServicesClient.ps1 -TestType ReplyMail -ItemSubject 'Budget 2027' -AllowWrite
    Replies to the most recent Inbox message whose subject contains "Budget 2027".

.EXAMPLE
    .\Invoke-WebServicesClient.ps1 -Gui

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
    Exit codes : 0 = Passed, 1 = Failed, 2 = finished with warnings or blocked (writes not authorised).
    Documentation : docs\WebServicesClient-Guide.html (source: docs\WebServicesClient-Guide.md)
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [ValidateSet('Discovery', 'SignIn', 'Endpoint', 'Folders', 'ReadMail', 'FreeBusy', 'CreateFolder', 'SendMail', 'ReplyMail', 'MoveMail', 'DeleteMail', 'SeedData', 'CleanData', 'ReadOnly', 'MailCycle', 'Full')]
    [string]$TestType,
    [string]$Mailbox,
    [string]$SignInUser,
    [ValidateSet('Auto', 'EWS', 'Graph')][string]$Protocol,
    [ValidateSet('Autodiscover', 'Manual')][string]$Discovery,
    [string]$EwsUrl,
    [ValidateSet('OAuth', 'Basic', 'Windows')][string]$Authentication,
    [ValidateSet('ADFS', 'EntraID', 'Auto')][string]$Authority,
    [string]$AdfsUrl,
    [string]$TenantId,
    [ValidateSet('User', 'Delegated', 'Application')][string]$Context,
    [string]$AppClientId,
    [string]$ClientId,
    [string]$RedirectUri,
    [string]$CertificateThumbprint,
    [securestring]$ClientSecret,
    [ValidateSet('Auto', 'Self', 'Delegate', 'Impersonation')][string]$Access,
    [ValidateSet('Negotiate', 'NTLM', 'Kerberos')][string]$WindowsPackage,
    [pscredential]$Credential,
    [ValidateSet('Auto', 'Window', 'DeviceCode')][string]$SignIn,
    [switch]$AllowWrite,
    [string]$ItemId,
    [string]$ItemSubject,
    [string[]]$FreeBusyMailboxes,
    [int]$FreeBusyDays,
    [string]$Recipient,
    [int]$MessageCount,
    [switch]$Gui,
    [string]$OutputPath,
    [switch]$NoReport,
    [string]$AccessToken,
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config\WebServicesClient.config.psd1')
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$exitCode = 1
$moduleLoaded = $false

try {
    Import-Module (Join-Path $PSScriptRoot 'WebServicesClient.psd1') -Force
    $moduleLoaded = $true

    # ---- configuration, then command-line overrides ---------------------------------------------
    $settings = Import-WscConfiguration -Path $ConfigPath -Root $PSScriptRoot
    foreach ($name in 'TestType', 'Mailbox', 'SignInUser', 'Protocol', 'Discovery', 'EwsUrl', 'Authentication', 'Authority', 'AdfsUrl', 'TenantId', 'Context', 'AppClientId', 'ClientId',
        'RedirectUri', 'CertificateThumbprint', 'Access', 'WindowsPackage', 'SignIn', 'Recipient') {
        if ($PSBoundParameters.ContainsKey($name)) { $settings[$name] = [string]$PSBoundParameters[$name] }
    }
    if ($PSBoundParameters.ContainsKey('FreeBusyMailboxes')) { $settings.FreeBusyMailboxes = @($FreeBusyMailboxes | ForEach-Object { $_ -split '[,;\s]+' } | Where-Object { $_ }) }
    if ($PSBoundParameters.ContainsKey('FreeBusyDays')) { $settings.FreeBusyDays = $FreeBusyDays }
    if ($PSBoundParameters.ContainsKey('MessageCount')) { $settings.MessageCount = $MessageCount }
    if ($PSBoundParameters.ContainsKey('AllowWrite')) { $settings.AllowWrite = [bool]$AllowWrite }
    if ($EwsUrl -and -not $PSBoundParameters.ContainsKey('Discovery') -and -not $PSBoundParameters.ContainsKey('Mailbox')) { $settings.Discovery = 'Manual' }
    # Another mailbox than the one of the configuration: its EWS URL is not the one of the configuration either.
    if ($Mailbox -and -not $EwsUrl -and $settings.Discovery -eq 'Autodiscover') { $settings.EwsUrl = '' }
    if ($OutputPath) { $settings.OutputPath = [IO.Path]::GetFullPath($OutputPath, (Get-Location).Path) }
    $check = Test-WscConfiguration -Configuration $settings
    if (-not $check.IsValid) { throw ("Invalid value:`n - " + ($check.Problems -join "`n - ")) }

    $logPath = Start-WscLog -Directory $settings.LogPath -RetentionDays $settings.LogRetentionDays

    if ($Gui) {
        Write-WscLog 'STEP' '=== Web Services Client for Exchange - window opened ==='
        Show-WscTestGui -Configuration $settings
        $exitCode = 0
    }
    else {
        Write-WscRunBanner -Settings $settings -Credential $Credential -LogPath $logPath -NoReport:$NoReport
        if ($AccessToken) { Write-WscItem Warn 'An access token was passed on the command line: it can stay in the PowerShell history.' }
        $scenario = Get-WscTestCatalog | Where-Object Name -eq $settings.TestType
        # Secrets are asked here, kept in memory for this run only.
        if ($settings.Authentication -eq 'Basic' -and -not $Credential -and $scenario.SignIn) {
            $user = if ($settings.SignInUser) { [string]$settings.SignInUser } else { [string]$settings.Mailbox }
            Write-WscItem Info "Basic authentication: password of $user (UPN or DOMAIN\user), sent with every request and never written." -Icon Key
            $Credential = Get-Credential -UserName $user -Message "Web Services Client for Exchange - Basic authentication for $($settings.Mailbox)"
            if (-not $Credential) { throw 'No user name and password entered: the Basic test cannot run.' }
        }
        if ($settings.Authentication -eq 'OAuth' -and $settings.Context -eq 'Application' -and $scenario.SignIn -and -not $ClientSecret -and -not $AccessToken -and
            -not ($settings.CertificateThumbprint -and $settings.Authority -ne 'ADFS')) {
            Write-WscItem Info "Application context: client secret of $($settings.AppClientId), kept in memory and never written." -Icon Key
            $ClientSecret = Read-Host -AsSecureString -Prompt 'Client secret'
            if (-not $ClientSecret -or $ClientSecret.Length -eq 0) { throw 'No client secret entered: the application cannot sign in.' }
        }

        $result = Invoke-WscMailboxTest -Configuration $settings -TestType $settings.TestType -Credential $Credential -ClientSecret $ClientSecret -AccessToken $AccessToken -ItemId $ItemId -ItemSubject $ItemSubject

        $reportText = 'none (-NoReport)'
        if (-not $NoReport) {
            $report = Export-WscReport -Result $result -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Formats $settings.ReportFormats -Delimiter $settings.CsvDelimiter
            $reportText = if ($report.Files.Contains('Html')) { $report.Files.Html } else { $report.Directory }
        }
        Write-WscRunSummary -Result $result -ReportText $reportText -LogPath $logPath
        $exitCode = switch ($result.Status) { 'Passed' { 0 } 'Failed' { 1 } default { 2 } }
    }
}
catch {
    if ($moduleLoaded) { Write-WscItem Fail $_.Exception.Message; Write-Host '' }
    else { Write-Host "Web Services Client for Exchange: $($_.Exception.Message)" -ForegroundColor Red }
    $exitCode = 1
}
finally {
    if ($moduleLoaded) { Stop-WscLog }
    [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
}
exit $exitCode
