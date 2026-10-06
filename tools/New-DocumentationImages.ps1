<#
.SYNOPSIS
    Renders the screenshots of the guides (docs\images\wsc-*.png) from the real tool, against the simulated Exchange.

.DESCRIPTION
    No lab is needed and the images always match the current code: the console, the window and the report
    are produced by the module itself, with tests\WebServicesClient.Simulator.ps1 standing in for AD FS,
    Entra ID, Exchange and Microsoft Graph (Install-SimExchange replaces only the network calls). Paths and
    names are anonymised (C:\Tools\WebServicesClient, contoso.test).

        wsc-console.png          console of a ReadOnly run with AD FS (stages, free/busy map, summary card)
        wsc-console-ntlm.png     console of a ReadOnly run with Windows authentication (NTLM legs)
        wsc-console-online.png   console of a ReadOnly run against Exchange Online (Microsoft Graph)
        wsc-gui.png              the window after a ReadOnly run, with its progress box
        wsc-gui-dark.png         the window after a ReadOnly run with Windows authentication, dark theme
        wsc-report-overview.png  header, tiles and scope of the HTML report
        wsc-report-checks.png    the checks, grouped by stage, with the requests under each one
        wsc-report-freebusy.png  the free/busy view (scheduling assistant)
        wsc-report-folders.png   the Folders tab (folder tree)
        wsc-report-exchange.png  a request sent and the response received (HTTP trace)
        wsc-report-ntlm.png      the NTLM challenge of the Windows run, decoded

    The HTML pages are captured with Microsoft Edge (headless), the window with RenderTargetBitmap (WPF).
    Run tools\Build-Documentation.ps1 afterwards: the guides embed the images.

.PARAMETER OutputFolder
    Default: docs\images next to the tools folder.

.PARAMETER KeepWork
    Keeps the work folder (reports, HTML pages) and shows its path.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0  (from EAS OAuth Mailbox 1.2.1)
    Part of : Web Services Client for Exchange (repository tool, not in the package)
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [string]$OutputFolder,
    [switch]$KeepWork
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'docs\images' }
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge not found: it takes the screenshots (headless mode).' }
$work = Join-Path ([IO.Path]::GetTempPath()) ('wsc-doc-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work, $OutputFolder -Force | Out-Null
$shown = 'C:\Tools\WebServicesClient'

#region Edge helpers (same as EAS OAuth Mailbox) --------------------------------------------------
function Save-Screenshot([string]$Html, [string]$Png, [int]$Width, [int]$Height) {
    $url = 'file:///' + ($Html -replace '\\', '/')
    if (Test-Path $Png) { Remove-Item $Png -Force }
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$(Join-Path $work 'edge-profile')`"", "--window-size=$Width,$Height", '--force-device-scale-factor=1', '--virtual-time-budget=3000', "--screenshot=`"$Png`"", "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden
    $deadline = (Get-Date).AddSeconds(45)
    while (-not (Test-Path $Png) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
    if (-not $proc.WaitForExit(10000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    if (-not (Test-Path $Png)) { throw "Screenshot not written: $Png" }
}

function Get-PageHeight([string]$Html, [int]$Width) {
    # The page writes its height in body[data-h]; Edge returns the DOM with --dump-dom.
    $url = 'file:///' + ($Html -replace '\\', '/')
    $dom = Join-Path $work ('dom-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.html')
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$(Join-Path $work 'edge-profile')`"", "--window-size=$Width,2000", '--virtual-time-budget=3000', '--dump-dom', "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden -RedirectStandardOutput $dom
    if (-not $proc.WaitForExit(45000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    $m = $null
    for ($i = 0; $i -lt 20 -and -not ($m -and $m.Success); $i++) {
        $stream = [IO.File]::Open($dom, 'Open', 'Read', 'ReadWrite')
        try { $text = [IO.StreamReader]::new($stream).ReadToEnd() } finally { $stream.Dispose() }
        $m = [regex]::Match($text, 'data-h="(\d+)"')
        if (-not $m.Success) { Start-Sleep -Milliseconds 250 }
    }
    if (-not $m.Success) { throw "Height not measured: $Html" }
    return [int]$m.Groups[1].Value
}

function Save-Page([string]$Html, [string]$Name, [int]$Width, [int]$Height = 0) {
    if ($Height -le 0) { $Height = Get-PageHeight $Html $Width }
    $png = Join-Path $OutputFolder "$Name.png"
    Save-Screenshot $Html $png $Width $Height
    Write-Host ("  {0,-26} {1} x {2}" -f "$Name.png", $Width, $Height)
}
#endregion

#region Simulated organisation --------------------------------------------------------------------
# Console theme is chosen when the module loads: colours and emoji, as in Windows Terminal.
$env:WSC_FORCE_COLOR = '1'
$env:WSC_ICONS = 'Emoji'
Remove-Module WebServicesClient -ErrorAction SilentlyContinue
Import-Module (Join-Path $root 'WebServicesClient.psd1') -Force
$module = Get-Module WebServicesClient
. (Join-Path $root 'tests\WebServicesClient.Simulator.ps1')

$folders = @(
    @{ Id = 'F-INBOX'; Parent = 'F-ROOT'; Name = 'Inbox'; Class = 'IPF.Note'; Total = 6; Unread = 2; Children = 2 }
    @{ Id = 'F-PROJ'; Parent = 'F-INBOX'; Name = 'Projects'; Class = 'IPF.Note'; Total = 2; Unread = 0; Children = 1 }
    @{ Id = 'F-MIG'; Parent = 'F-PROJ'; Name = 'Migration'; Class = 'IPF.Note'; Total = 1; Unread = 0; Children = 0 }
    @{ Id = 'F-ARCH'; Parent = 'F-INBOX'; Name = 'Archive 2026'; Class = 'IPF.Note'; Total = 1; Unread = 0; Children = 0 }
    @{ Id = 'F-DRAFTS'; Parent = 'F-ROOT'; Name = 'Drafts'; Class = 'IPF.Note'; Total = 0; Unread = 0; Children = 0 }
    @{ Id = 'F-SENT'; Parent = 'F-ROOT'; Name = 'Sent Items'; Class = 'IPF.Note'; Total = 4; Unread = 0; Children = 0 }
    @{ Id = 'F-DEL'; Parent = 'F-ROOT'; Name = 'Deleted Items'; Class = 'IPF.Note'; Total = 1; Unread = 0; Children = 0 }
    @{ Id = 'F-CAL'; Parent = 'F-ROOT'; Name = 'Calendar'; Class = 'IPF.Appointment'; Total = 7; Unread = 0; Children = 0 }
    @{ Id = 'F-CON'; Parent = 'F-ROOT'; Name = 'Contacts'; Class = 'IPF.Contact'; Total = 3; Unread = 0; Children = 0 }
    @{ Id = 'F-TASK'; Parent = 'F-ROOT'; Name = 'Tasks'; Class = 'IPF.Task'; Total = 0; Unread = 0; Children = 0 }
    @{ Id = 'F-JUNK'; Parent = 'F-ROOT'; Name = 'Junk Email'; Class = 'IPF.Note'; Total = 0; Unread = 0; Children = 0 }
)
$day = [datetime]::UtcNow.Date
$inbox = @(
    @{ Id = 'M-1'; Subject = 'Agenda - migration workshop'; From = 'alice.martin@contoso.test'; Received = $day.AddHours(7).AddMinutes(58).ToString('s') + 'Z'; Read = 'false'; Body = 'Hello, here is the agenda of Thursday.' }
    @{ Id = 'M-2'; Subject = 'RE: EWS application access policy'; From = 'bob.durand@contoso.test'; Received = $day.AddHours(-7).ToString('s') + 'Z'; Read = 'false'; Body = 'The policy is in place for the archiving tool.' }
    @{ Id = 'M-3'; Subject = 'Maintenance window this weekend'; From = 'exchange@contoso.test'; Received = $day.AddDays(-1).AddHours(9).ToString('s') + 'Z'; Read = 'true'; Body = 'The DAG will be patched on Saturday.' }
    @{ Id = 'M-4'; Subject = 'Free/busy with Exchange Online'; From = 'carla.petit@contoso.test'; Received = $day.AddDays(-1).AddHours(6).ToString('s') + 'Z'; Read = 'true'; Body = 'Can you check the organisation relationship?' }
    @{ Id = 'M-5'; Subject = 'Room 1 booking policy'; From = 'facilities@contoso.test'; Received = $day.AddDays(-2).AddHours(14).ToString('s') + 'Z'; Read = 'true'; Body = 'Room 1 now accepts meetings up to 4 hours.' }
    @{ Id = 'M-6'; Subject = 'Welcome to the pilot'; From = 'servicedesk@contoso.test'; Received = $day.AddDays(-3).AddHours(8).ToString('s') + 'Z'; Read = 'true'; Body = 'Your mailbox is in the pilot group.' }
)
function New-DocState([hashtable]$Changes = @{}) {
    $s = New-SimState
    $s.Folders = [Collections.Generic.List[object]]@($folders | ForEach-Object { $_.Clone() })
    $s.Inbox = [Collections.Generic.List[object]]@($inbox | ForEach-Object { $_.Clone() })
    $s.FreeBusy = @{ 'ews-test@contoso.test' = 'Success'; 'room1@contoso.test' = 'Success'; 'alice.martin@contoso.test' = 'Success' }
    $s.CertificateDays = 240
    foreach ($k in $Changes.Keys) { $s[$k] = $Changes[$k] }
    return $s
}

$settings = Import-WscConfiguration
$settings.TestType = 'ReadOnly'
$settings.FreeBusyMailboxes = @('ews-test@contoso.test', 'room1@contoso.test', 'alice.martin@contoso.test')
$settings.MessageCount = 6
$settings.OutputPath = Join-Path $work 'reports'
$settings.LogPath = Join-Path $work 'logs'
$stamp = Get-Date -Format 'yyyyMMdd'

function Invoke-DocRun {
    # One run as the command line shows it: banner, stages, summary card; the console records are returned.
    param([hashtable]$State, [hashtable]$Settings, [pscredential]$Credential, [string]$Name)
    Install-SimExchange -Module $module -State $State
    $displayed = $Settings.Clone(); $displayed.OutputPath = "$shown\reports"
    $log = "$shown\logs\WebServicesClient_$stamp.log"
    $records = & {
        Write-WscRunBanner -Settings $displayed -Credential $Credential -LogPath $log
        $script:DocResult = Invoke-WscMailboxTest -Configuration $Settings -TestType $Settings.TestType -Credential $Credential
        Write-WscRunSummary -Result $script:DocResult -ReportText "$shown\reports\WebServicesClient_$($Settings.TestType)_$stamp-0930\WebServicesClient.html" -LogPath $log
    } 6>&1
    $report = Export-WscReport -Result $script:DocResult -OutputPath (Join-Path $work "reports-$Name")
    Write-Host ("  {0,-8} {1}, {2} checks, {3}" -f $Name, $script:DocResult.Status, $script:DocResult.Steps.Count, $script:DocResult.Protocol)
    [pscustomobject]@{ Records = $records; Result = $script:DocResult; Html = $report.Files.Html }
}

Write-Host 'Runs (simulated Exchange)...'
$adfs = Invoke-DocRun -State (New-DocState) -Settings $settings -Name 'adfs'

$ntlmSettings = $settings.Clone()
$ntlmSettings.Authentication = 'Windows'; $ntlmSettings.WindowsPackage = 'NTLM'
$ntlmCredential = [pscredential]::new('CONTOSO\ews-test', (ConvertTo-SecureString 'Sim-Pa55word!' -AsPlainText -Force))
$ntlm = Invoke-DocRun -State (New-DocState) -Settings $ntlmSettings -Credential $ntlmCredential -Name 'ntlm'

$onlineSettings = $settings.Clone()
$onlineSettings.Authority = 'EntraID'; $onlineSettings.AdfsUrl = ''; $onlineSettings.EwsUrl = ''
$online = Invoke-DocRun -State (New-DocState @{ AutodiscoverV2 = 'https://outlook.office365.com/EWS/Exchange.asmx'; AuthorizationUri = 'https://login.microsoftonline.com/common/oauth2/authorize' }) -Settings $onlineSettings -Name 'online'
#endregion

#region Console images ----------------------------------------------------------------------------
function ConvertFrom-Ansi([string]$Line) {
    # SGR codes used by the console theme: 0 reset, 1 bold, 90 dim, 97 white, 38;2;r;g;b and 48;2;r;g;b.
    $out = [Text.StringBuilder]::new()
    $fg = $null; $bg = $null; $bold = $false
    foreach ($part in [regex]::Split($Line, '(\x1b\[[0-9;]*m)')) {
        if ($part -match '^\x1b\[([0-9;]*)m$') {
            $codes = @($Matches[1].Split(';') | ForEach-Object { if ($_ -eq '') { 0 } else { [int]$_ } })
            for ($i = 0; $i -lt $codes.Count; $i++) {
                switch ($codes[$i]) {
                    0 { $fg = $null; $bg = $null; $bold = $false }
                    1 { $bold = $true }
                    22 { $bold = $false }
                    39 { $fg = $null }
                    49 { $bg = $null }
                    90 { $fg = '#8a8a8a' }
                    97 { $fg = '#ffffff' }
                    38 { $fg = 'rgb({0},{1},{2})' -f $codes[$i + 2], $codes[$i + 3], $codes[$i + 4]; $i += 4 }
                    48 { $bg = 'rgb({0},{1},{2})' -f $codes[$i + 2], $codes[$i + 3], $codes[$i + 4]; $i += 4 }
                }
            }
            continue
        }
        if ($part -eq '') { continue }
        $style = @()
        if ($fg) { $style += "color:$fg" }
        if ($bg) { $style += "background:$bg" }
        if ($bold) { $style += 'font-weight:700' }
        $text = [Net.WebUtility]::HtmlEncode($part)
        [void]$out.Append($(if ($style) { "<span style=""$($style -join ';')"">$text</span>" } else { $text }))
    }
    return $out.ToString()
}

function Save-Console([object[]]$Records, [string]$Command, [string]$Name) {
    $lines = foreach ($r in $Records) {
        $data = if ($r -is [Management.Automation.InformationRecord]) { $r.MessageData } else { $r }
        $text = if ($data -is [Management.Automation.HostInformationMessage]) { [string]$data.Message } else { [string]$data }
        $text.Replace($settings.OutputPath, "$shown\reports").Replace($settings.LogPath, "$shown\logs")
    }
    $body = ($lines | ForEach-Object { ConvertFrom-Ansi $_ }) -join "`n"
    $console = @"
<!doctype html><html><head><meta charset="utf-8"><style>
body { margin:0; background:#ffffff; font-family:"Segoe UI", sans-serif; }
.win { width:1180px; margin:0; border-radius:10px; overflow:hidden; background:#0c0c0c; border:1px solid #2b2b2b; }
.bar { display:flex; align-items:center; gap:10px; height:38px; padding:0 14px; background:#202020; color:#d0d0d0; font-size:12.5px; }
.tab { padding:6px 14px; background:#0c0c0c; border-radius:8px 8px 0 0; margin-top:8px; }
pre { margin:0; padding:14px 18px 18px; color:#cccccc; font:13.5px/1.42 "Cascadia Mono", Consolas, monospace; white-space:pre-wrap; word-break:break-all; }
.prompt { color:#cccccc; }
</style></head><body><div class="win"><div class="bar"><span class="tab">PowerShell 7.5</span></div>
<pre><span class="prompt">PS $shown&gt; $([Net.WebUtility]::HtmlEncode($Command))</span>
$body</pre></div>
<script>document.body.setAttribute('data-h', Math.ceil(document.querySelector('.win').getBoundingClientRect().height) + 2);</script></body></html>
"@
    $page = Join-Path $work "$Name.html"
    [IO.File]::WriteAllText($page, $console, [Text.UTF8Encoding]::new($false))
    Save-Page $page $Name 1182
}

Write-Host 'Images:'
Save-Console $adfs.Records '.\Invoke-WebServicesClient.ps1 -TestType ReadOnly -FreeBusyMailboxes ews-test@contoso.test,room1@contoso.test,alice.martin@contoso.test' 'wsc-console'
Save-Console $ntlm.Records '.\Invoke-WebServicesClient.ps1 -TestType ReadOnly -Authentication Windows -WindowsPackage NTLM -Credential CONTOSO\ews-test' 'wsc-console-ntlm'
Save-Console $online.Records '.\Invoke-WebServicesClient.ps1 -TestType ReadOnly -Authority EntraID' 'wsc-console-online'
#endregion

#region Report images -----------------------------------------------------------------------------
function Get-ReportHtml([string]$Path) {
    [IO.File]::ReadAllText($Path).Replace([Net.WebUtility]::HtmlEncode($settings.OutputPath), "$shown\reports").Replace($settings.OutputPath.Replace('\', '\\'), "$shown\reports".Replace('\', '\\'))
}
$adfsHtml = Get-ReportHtml $adfs.Html
$ntlmHtml = Get-ReportHtml $ntlm.Html
function Save-ReportView([string]$Name, [string]$Css, [string]$Script, [int]$Height = 0, [string]$Html = $adfsHtml) {
    # Height of the body itself: documentElement.scrollHeight is never smaller than the window.
    $measure = "document.body.setAttribute('data-h', Math.ceil(document.body.getBoundingClientRect().height));"
    $inject = "<style>$Css</style><script>window.addEventListener('load', () => { $Script; setTimeout(() => { $measure }, 50); });</script></body>"
    $page = Join-Path $work "$Name.html"
    [IO.File]::WriteAllText($page, $Html.Replace('</body>', $inject), [Text.UTF8Encoding]::new($false))
    Save-Page $page $Name 1280 $Height
}
$only = { param([string]$Keep) "header, section.block, footer { display:none !important; } $Keep { display:block !important; } body { padding-top:20px; padding-bottom:8px; }" }
Save-ReportView 'wsc-report-overview' 'section.block:nth-of-type(n+2), footer { display:none !important; } body { padding-bottom:8px; }' ''
Save-ReportView 'wsc-report-checks' 'header, section.block:nth-of-type(1), section.block:nth-of-type(n+3), footer { display:none !important; } body { padding-top:20px; padding-bottom:8px; }' ''
Save-ReportView 'wsc-report-freebusy' (& $only '#fb-block') ''
Save-ReportView 'wsc-report-folders' ((& $only '#data') + ' .scroll { max-height:none !important; }') "Array.from(document.querySelectorAll('#tabs button')).find(b => b.textContent.startsWith('Folders')).click()"
Save-ReportView 'wsc-report-exchange' 'body { min-height:980px; } dialog { max-height:none; } .dialog-body { max-height:none; }' "Array.from(document.querySelectorAll('.timeline .xchg li')).find(li => li.textContent.includes('FindItem')).click()" 980
Save-ReportView 'wsc-report-ntlm' 'body { min-height:980px; } dialog { max-height:none; } .dialog-body { max-height:none; }' "Array.from(document.querySelectorAll('.timeline .xchg li')).find(li => li.textContent.includes('NTLM type 1')).click()" 980 -Html $ntlmHtml
#endregion

#region Window images -----------------------------------------------------------------------------
Write-Host 'Window runs (simulated Exchange)...'
function Save-Window([hashtable]$Configuration, [hashtable]$State, [string]$Name, [string]$Theme = 'Light', [scriptblock]$Prepare) {
    Install-SimExchange -Module $module -State $State
    $window = New-WscTestForm -Configuration $Configuration -Theme $Theme
    $form = $window.Form
    # Off screen, at its design size (the work area of this computer does not matter for the image).
    $form.WindowStartupLocation = 'Manual'
    $form.Left = -6000; $form.Top = -6000; $form.Width = 1180; $form.Height = 860
    $form.ShowActivated = $false; $form.ShowInTaskbar = $false
    $form.Show()
    if ($Prepare) { & $Prepare $window }
    & $module { Invoke-WscGuiRun } 6>$null
    # Anonymised paths in the progress and the footer.
    for ($i = 0; $i -lt $window.Items.Count; $i++) {
        $item = $window.Items[$i]
        if ($item.Text -and $item.Text.Contains($Configuration.OutputPath)) { $copy = $item.PSObject.Copy(); $copy.Text = $item.Text.Replace($Configuration.OutputPath, "$shown\reports"); $window.Items[$i] = $copy }
    }
    $window.Controls.Footer.Text = $window.Controls.Footer.Text.Replace($Configuration.OutputPath, "$shown\reports")
    $window.Controls.LogScroll.ScrollToHome()
    & $module { Invoke-WscGuiPump }
    $form.UpdateLayout()
    $content = $form.Content
    $bitmap = [Windows.Media.Imaging.RenderTargetBitmap]::new([int][Math]::Ceiling($content.ActualWidth), [int][Math]::Ceiling($content.ActualHeight), 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($content)
    $encoder = [Windows.Media.Imaging.PngBitmapEncoder]::new()
    $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream = [IO.File]::Create((Join-Path $OutputFolder "$Name.png"))
    try { $encoder.Save($stream) } finally { $stream.Dispose() }
    Write-Host ("  {0,-26} {1} x {2} ({3})" -f "$Name.png", $bitmap.PixelWidth, $bitmap.PixelHeight, $Theme)
    $form.Close()
}
Save-Window $settings.Clone() (New-DocState) 'wsc-gui'
Save-Window $ntlmSettings.Clone() (New-DocState) 'wsc-gui-dark' -Theme Dark -Prepare {
    param($w)
    $w.Controls.CurrentAccount.IsChecked = $false
    $w.Controls.UserName.Text = 'CONTOSO\ews-test'
    $w.Controls.Password.Password = 'Sim-Pa55word!'
    & $module { Update-WscGuiScenario }
}
#endregion

Remove-Module WebServicesClient -Force
Remove-Item Env:\WSC_FORCE_COLOR, Env:\WSC_ICONS -ErrorAction SilentlyContinue
if ($KeepWork) { Write-Host "Work folder: $work" } else { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
