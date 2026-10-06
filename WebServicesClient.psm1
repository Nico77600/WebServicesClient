<#
.SYNOPSIS
    Web Services Client for Exchange - PowerShell module.

.DESCRIPTION
    Loads the parts of the tool, in the order of an execution:

        src\WebServicesClient.Console.ps1   console output and log file (same rules as EAS OAuth Mailbox)
        src\WebServicesClient.Config.ps1    configuration file and scenarios
        src\WebServicesClient.Http.ps1      HTTP requests and their trace (request sent, response received)
        src\WebServicesClient.Windows.ps1   Windows authentication (Negotiate, NTLM, Kerberos), tokens decoded
        src\WebServicesClient.Ews.ps1       EWS protocol: SOAP envelopes, headers, responses, throttling
        src\WebServicesClient.OAuth.ps1     helpers and OAuth building blocks shared with EAS OAuth Mailbox
        src\WebServicesClient.Checks.ps1    steps, OAuth sign-in (window, application), token claims
        src\WebServicesClient.Discovery.ps1 Autodiscover and the prerequisites without sign-in
        src\WebServicesClient.Graph.ps1     Microsoft Graph: the same operations for Exchange Online
        src\WebServicesClient.Seed.ps1      test data: fill a test mailbox (messages, folders, calendar) and clean it
        src\WebServicesClient.Run.ps1       protocol, sign-in stage and Invoke-WscMailboxTest
        src\WebServicesClient.Operations.ps1  the EWS operations: endpoint, folders, read, free/busy, create, send, reply, move, delete
        src\WebServicesClient.Browser.ps1   sign-in window (Edge or Chrome, DevTools protocol)
        src\WebServicesClient.Report.ps1    CSV, JSON and HTML report
        src\WebServicesClient.Gui.ps1       WPF window (Fluent theme of Windows 11)

    The access token, the client secret and the passwords stay in memory: they are never written to
    the console, the log or the report.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
    History : see CHANGELOG.md
#>
#Requires -Version 7.4
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Net.Http

$script:ToolVersion = '1.0.0'
$script:ToolRoot = $PSScriptRoot
$script:WscUserAgent = 'WebServicesClient/1.0'
# User name of the deliberately wrong Basic credentials (Discovery): a user that does not exist,
# so that no real account is ever locked by the test.
$script:InvalidBasicUserPrefix = 'wsc-invalid-'
$script:LogWriter = $null
$script:LogPath = $null
$script:Quiet = $false
# GUI hooks, set only while the window runs a test: Sink (progress lines), Pump (keeps the window responsive), Cancel.
$script:Ui = $null

foreach ($part in 'Console', 'Config', 'Http', 'Windows', 'Ews', 'OAuth', 'Checks', 'Discovery', 'Operations', 'Graph', 'Seed', 'Run', 'Browser', 'Report', 'Gui') {
    . (Join-Path $PSScriptRoot "src\WebServicesClient.$part.ps1")
}

# Unsigned token with a wrong audience: Exchange must answer 401 (Discovery scenario).
$script:InvalidToken = '{0}.{1}.{2}' -f (ConvertTo-WscBase64Url ([Text.Encoding]::UTF8.GetBytes('{"alg":"none","typ":"JWT"}'))),
    (ConvertTo-WscBase64Url ([Text.Encoding]::UTF8.GetBytes('{"aud":"https://invalid.example/","iss":"web-services-client-for-exchange","exp":1}'))),
    (ConvertTo-WscBase64Url ([Text.Encoding]::UTF8.GetBytes('invalid')))

Export-ModuleMember -Function @(
    'Import-WscConfiguration', 'Test-WscConfiguration', 'Get-WscTestCatalog', 'Get-WscScenarioStages'
    'Invoke-WscMailboxTest', 'Export-WscReport', 'Show-WscTestGui', 'New-WscTestForm'
    'Start-WscLog', 'Stop-WscLog', 'Write-WscLog', 'Write-WscBanner', 'Write-WscStep', 'Write-WscItem', 'Write-WscSummary', 'Write-WscRunBanner', 'Write-WscRunSummary', 'Format-WscDuration'
)
