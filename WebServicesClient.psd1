#
#  Web Services Client for Exchange - module manifest
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : see ModuleVersion
#
#  Loaded by Invoke-WebServicesClient.ps1 (Import-Module by path).
#
@{
    RootModule        = 'WebServicesClient.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'bcdf1a8c-5c1b-45cb-a542-3cfe56b3062c'
    Author            = 'Nicolas Fabert'
    Copyright         = '(c) 2026 Nicolas Fabert. MIT License.'
    Description       = 'Web Services Client for Exchange: step-by-step test toolbox for Exchange mailboxes through EWS (on-premises) and Microsoft Graph (Exchange Online), with OAuth (AD FS, Entra ID for HMA or Exchange Online), Basic or Windows authentication (NTLM, Kerberos), as a user, a delegated application or an application - Autodiscover, folders, read, send, reply, move, delete, free/busy, test data - with the HTTP trace of every request and CSV, JSON and HTML reports.'
    PowerShellVersion = '7.4'

    # Functions called by Invoke-WebServicesClient.ps1, the tests and the documentation tool. The other functions stay internal.
    FunctionsToExport = @(
        'Import-WscConfiguration', 'Test-WscConfiguration', 'Get-WscTestCatalog', 'Get-WscScenarioStages'
        'Invoke-WscMailboxTest', 'Export-WscReport', 'Show-WscTestGui', 'New-WscTestForm'
        'Start-WscLog', 'Stop-WscLog', 'Write-WscLog', 'Write-WscBanner', 'Write-WscStep', 'Write-WscItem', 'Write-WscSummary', 'Write-WscRunBanner', 'Write-WscRunSummary', 'Format-WscDuration'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags       = @('Exchange', 'EWS', 'MicrosoftGraph', 'OAuth', 'ADFS', 'EntraID', 'NTLM', 'Kerberos', 'FreeBusy', 'Diagnostic')
            LicenseUri = 'https://opensource.org/licenses/MIT'
        }
    }
}
