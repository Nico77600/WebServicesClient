#
#  Web Services Client for Exchange - configuration file
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : 1.0.0
#
#  This file is read by Invoke-WebServicesClient.ps1 and by the window (-Gui). It is a PowerShell
#  data file: text between quotes, $true / $false, numbers, and @( ) for lists.
#  Lines starting with # are comments. Relative paths are relative to the tool folder.
#  Every value is checked at start; all the problems are listed at once.
#
#  No secret here: the OAuth sign-in is interactive (sign-in window or device code), the client
#  secret of an application and the passwords (Basic, Windows) are asked at each run and never
#  written. Replace every contoso.test value before the first run.
#
@{
    # ---------------------------------------------------------------------
    # Which mailbox, and how the EWS URL is found.
    #   Mailbox:    the mailbox the tests work on (SMTP address).
    #   SignInUser: the account that signs in. Empty = Mailbox. Another account = delegate access
    #               (Full Access or folder permissions) or impersonation (Identity.Access).
    #   Protocol:   'Auto' (Microsoft Graph for a mailbox in Exchange Online, where EWS is being retired
    #               since October 2026; EWS for a mailbox on-premises, which Graph does not reach),
    #               'EWS' or 'Graph' (Graph needs OAuth with Entra ID).
    #   Discovery:  'Autodiscover' (like Outlook: Autodiscover v2, then the classic Autodiscover)
    #               or 'Manual' (EwsUrl below). With Autodiscover, EwsUrl is used only when
    #               Autodiscover gives no answer.
    #   Authority:  where the OAuth sign-in happens.
    #     'ADFS'    AD FS at AdfsUrl (Exchange Server 2019 CU13+ / SE with AD FS)
    #     'EntraID' Entra ID: Exchange on-premises with hybrid modern authentication (HMA), or
    #               Exchange Online (EwsUrl = https://outlook.office365.com/EWS/Exchange.asmx)
    #     'Auto'    the server Exchange names in its challenge, like a client
    # ---------------------------------------------------------------------
    Target = @{
        Mailbox    = 'ews-test@contoso.test'
        SignInUser = ''
        Protocol   = 'Auto'                                         # Auto | EWS | Graph
        Discovery  = 'Autodiscover'                                 # Autodiscover | Manual
        EwsUrl     = 'https://mail.contoso.test/EWS/Exchange.asmx'
        AdfsUrl    = 'https://adfs.contoso.test/adfs'               # AD FS root, ends with /adfs
        Authority  = 'ADFS'                                         # ADFS | EntraID | Auto
        TenantId   = ''                                             # EntraID: tenant ID or domain (empty = domain of the mailbox)
    }

    # ---------------------------------------------------------------------
    # Who signs in, and how.
    #   Authentication: 'OAuth' (token of AD FS or Entra ID), 'Basic' (user name and password with
    #     every request) or 'Windows' (Negotiate, NTLM or Kerberos: the handshake is done and traced
    #     by the tool; current Windows account, or the account given with -Credential).
    #   Context (OAuth):
    #     'User'        the user signs in with a Microsoft client (ClientId, Microsoft Office)
    #     'Delegated'   the user signs in through YOUR application (AppClientId, delegated permission
    #                   EWS.AccessAsUser.All): the application acts on behalf of the user
    #     'Application' the application signs in alone (client credentials: certificate or client
    #                   secret, permission full_access_as_app) and impersonates the mailbox
    #   Access: how the mailbox is opened.
    #     'Auto'          Self when SignInUser is empty or is the mailbox, Delegate otherwise,
    #                     Impersonation for the Application context
    #     'Self'          the own mailbox of the signed-in user
    #     'Delegate'      another mailbox through permissions (<t:Mailbox> in the folder IDs)
    #     'Impersonation' ExchangeImpersonation header (ApplicationImpersonation role on-premises,
    #                     Application context in Exchange Online)
    #   WindowsPackage: 'Negotiate' (Kerberos, else NTLM, like Windows), 'NTLM' or 'Kerberos'.
    # ---------------------------------------------------------------------
    Identity = @{
        Authentication        = 'OAuth'                                  # OAuth | Basic | Windows
        Context               = 'User'                                   # User | Delegated | Application
        ClientId              = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'   # User: public client allowed for EWS (Microsoft Office)
        AppClientId           = ''                                       # Delegated / Application: client ID of your application
        RedirectUri           = ''                                       # Delegated: a redirect URI of your application (empty = default of the server)
        CertificateThumbprint = ''                                       # Application with Entra ID: certificate in Cert:\CurrentUser\My or LocalMachine\My (empty = client secret asked)
        WindowsPackage        = 'Negotiate'                              # Negotiate | NTLM | Kerberos
        Access                = 'Auto'                                   # Auto | Self | Delegate | Impersonation
    }

    # ---------------------------------------------------------------------
    # EWS requests (good practices of Microsoft, see the guide).
    #   RequestServerVersion: schema asked in every request (Exchange2013_SP1, Exchange2016...).
    #   UserAgent: identifies the tool in the IIS, HttpProxy and Exchange Online logs.
    #   PreferServerAffinity: X-PreferServerAffinity on the first request, then the affinity cookie
    #     (X-BackEndOverrideCookie) Exchange returns is sent back.
    #   MaxRetries: new attempts after ErrorServerBusy, HTTP 429 or 503 (waiting the time Exchange asks).
    # ---------------------------------------------------------------------
    Ews = @{
        RequestServerVersion = 'Exchange2016'
        UserAgent            = 'WebServicesClient/1.0'
        PreferServerAffinity = $true
        MaxRetries           = 2
    }

    # ---------------------------------------------------------------------
    # How it is tested.
    #   DefaultType: Discovery | SignIn | Endpoint | Folders | ReadMail | FreeBusy | CreateFolder |
    #                SendMail | ReplyMail | MoveMail | DeleteMail | SeedData | CleanData |
    #                ReadOnly | MailCycle | Full
    #   SignIn: OAuth sign-in. 'Window' (Edge or Chrome, temporary profile), 'DeviceCode' (a code
    #     typed on any device) or 'Auto' (the window when this session can show one).
    #   AllowWrite: $true lets the scenarios that change the mailbox run (create a folder, send,
    #     reply, move, delete). Keep $false outside a test mailbox. -AllowWrite on the command line.
    #   FolderName: test folder, created under the Inbox (CreateFolder) and target of MoveMail.
    #   Recipient: recipient of the test message (empty = the mailbox itself).
    #   DeleteMode: MoveToDeletedItems | SoftDelete | HardDelete.
    #   DeliveryWaitSeconds: longest wait for the test message to arrive in the Inbox.
    #   FreeBusyMailboxes: mailboxes whose free/busy is read (empty = the mailbox).
    # ---------------------------------------------------------------------
    Test = @{
        DefaultType             = 'ReadOnly'
        SignIn                  = 'Auto'          # Auto | Window | DeviceCode
        AllowWrite              = $false
        MessageCount            = 10              # messages listed by ReadMail (1-100)
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
    }

    # ---------------------------------------------------------------------
    # Report files (one sub-folder per execution), written locally only.
    # ---------------------------------------------------------------------
    Report = @{
        OutputPath   = '.\reports'
        FilePrefix   = 'WebServicesClient'
        Formats      = @('Csv', 'Html')   # a Summary.json file is always written as well
        CsvDelimiter = ';'                # ';' opens directly in Excel with French regional settings
    }

    # ---------------------------------------------------------------------
    # Execution log files (one per day, no token, no colour).
    # ---------------------------------------------------------------------
    Logging = @{
        Path          = '.\logs'
        RetentionDays = 14
    }
}
