# Web Services Client for Exchange

A test toolbox for Exchange mailboxes, on-premises through EWS and in Exchange Online through Microsoft Graph.

This folder contains everything needed to run the tool: `Invoke-WebServicesClient.ps1`, the module, the configuration, the report template and the guides. Tests and build tools stay outside it, in the repository.

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows. Unblock them once, from this folder:
>
> ```powershell
> Get-ChildItem . -Recurse -File | Unblock-File
> ```

## Requirements
- Windows 10 / 11 or Windows Server 2016 to 2025.
- PowerShell 7.4 or later.
- Microsoft Edge or Google Chrome for the OAuth sign-in window, or device code sign-in.
- HTTPS to EWS or `graph.microsoft.com`, Autodiscover, and AD FS or `login.microsoftonline.com`.
- A test mailbox and, for applications, an app registration with admin consent.
- Exchange Server 2019 CU13 or later, Exchange Server SE, Exchange Online, or Basic / Windows authentication on EWS.

## Quick start
```powershell
notepad .\config\WebServicesClient.config.psd1     # the test mailbox, the EWS URL, the AD FS URL

.\Invoke-WebServicesClient.ps1 -TestType Discovery                                   # no sign-in: Autodiscover, certificates, which sign-in Exchange offers
.\Invoke-WebServicesClient.ps1 -TestType ReadOnly                                    # sign-in, folders, messages, free/busy - changes nothing
.\Invoke-WebServicesClient.ps1 -TestType ReadOnly -Authority EntraID -Mailbox alice@contoso.com     # HMA, or Exchange Online through Graph
.\Invoke-WebServicesClient.ps1 -TestType ReadOnly -Authentication Windows -WindowsPackage NTLM -Credential CONTOSO\alice
.\Invoke-WebServicesClient.ps1 -TestType FreeBusy -FreeBusyMailboxes alice@contoso.com,room1@contoso.com
.\Invoke-WebServicesClient.ps1 -TestType MailCycle -AllowWrite                       # test mailbox: send, reply, move, delete
.\Invoke-WebServicesClient.ps1 -Gui                                                  # the same in a window
```

## Content
| Item | Role |
|---|---|
| `Invoke-WebServicesClient.ps1` | Entry script. |
| `WebServicesClient.psd1` | Module manifest. |
| `WebServicesClient.psm1` | Module loader. |
| `config\` | Example configuration. |
| `docs\` | User and developer guide files. |
| `src\` | Module implementation. |
| `templates\` | HTML report template. |
| `LICENSE` | MIT license. |
| `THIRD-PARTY-NOTICES.md` | Third-party notices. |

## Documentation
- [User guide](docs/WebServicesClient-UserGuide.md) - also `docs/WebServicesClient-UserGuide.html`, a single file to open locally
- [Developer guide](docs/WebServicesClient-Guide.md) - also `docs/WebServicesClient-Guide.html`, a single file to open locally

Project page, releases and change log: https://github.com/Nico77600/WebServicesClient

License: [MIT](LICENSE).
