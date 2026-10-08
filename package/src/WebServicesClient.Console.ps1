<#
.SYNOPSIS
    Web Services Client for Exchange - console output and log file (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    Same rules as EAS OAuth Mailbox and Exchange Log Report:
      - ANSI colours are disabled when the output is redirected or NO_COLOR is set;
        WSC_FORCE_COLOR=1 forces them.
      - Icons: emoji in Windows Terminal / VS Code, symbols of the classic console fonts elsewhere.
        WSC_ICONS = Emoji | Symbols | Ascii forces a style.
      - Every line shown is also written to the daily log file, without colours or icons.
      - During a GUI run, the same lines are sent to the progress box of the window.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:C = @{ Reset = ''; Bold = ''; Dim = ''; Accent = ''; AccentBg = ''; Green = ''; Yellow = ''; Red = ''; White = '' }
if ($env:WSC_FORCE_COLOR -eq '1' -or (-not [Console]::IsOutputRedirected -and -not $env:NO_COLOR)) {
    $e = [char]27
    $script:C = @{
        Reset = "$e[0m"; Bold = "$e[1m"; Dim = "$e[90m"; White = "$e[97m"
        Accent = "$e[38;2;214;62;115m"; AccentBg = "$e[48;2;177;31;75m$e[97m"
        Green = "$e[38;2;80;200;120m"; Yellow = "$e[38;2;240;200;90m"; Red = "$e[38;2;240;90;90m"
    }
}
$script:IconStyle = if ($env:WSC_ICONS -in 'Emoji', 'Symbols', 'Ascii') { $env:WSC_ICONS }
    elseif ([Console]::IsOutputRedirected) { 'Symbols' }
    elseif ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') { 'Emoji' }
    else { 'Symbols' }

function Get-WscIconSet {
    <# Icons of one console style. Symbols: only characters of the classic console fonts. #>
    param([Parameter(Mandatory = $true)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)

    $u = { param([int]$Code) [char]::ConvertFromUtf32($Code) }
    switch ($Style) {
        'Emoji' {
            return @{
                Logo = & $u 0x1F4EC; Ok = & $u 0x2705; Warn = (& $u 0x26A0) + [char]0xFE0F; Fail = & $u 0x274C; Info = & $u 0x1F539
                Skip = & $u 0x23E9; Block = & $u 0x26D4; Key = & $u 0x1F511; Server = & $u 0x1F5A5; Shield = & $u 0x1F512
                Folder = & $u 0x1F4C1; Mail = & $u 0x1F4E8; People = & $u 0x1F465; File = & $u 0x1F4C4; Log = & $u 0x1F4DD
                Report = & $u 0x1F4CA; Done = & $u 0x1F389; Target = & $u 0x1F3AF; Search = & $u 0x1F50E; Clock = & $u 0x23F3
                Settings = & $u 0x1F527
            }
        }
        'Symbols' {
            return @{
                Logo = & $u 0x2666; Ok = & $u 0x221A; Warn = & $u 0x25B2; Fail = & $u 0x00D7; Info = & $u 0x2022
                Skip = & $u 0x00BB; Block = & $u 0x25A0; Key = & $u 0x00A7; Server = & $u 0x2261; Shield = & $u 0x25CA
                Folder = & $u 0x2302; Mail = '@'; People = & $u 0x2192; File = & $u 0x25AC; Log = & $u 0x00B6
                Report = & $u 0x2261; Done = & $u 0x221A; Target = & $u 0x25D9; Search = & $u 0x25BA; Clock = & $u 0x25CB
                Settings = & $u 0x263C
            }
        }
        default {
            return @{
                Logo = '*'; Ok = '+'; Warn = '!'; Fail = 'x'; Info = '-'; Skip = '>'; Block = '#'; Key = 'k'; Server = '='
                Shield = 'o'; Folder = '>'; Mail = '@'; People = '&'; File = '-'; Log = '='; Report = '='; Done = '*'
                Target = 'o'; Search = '?'; Clock = '~'; Settings = '%'
            }
        }
    }
}

function Get-WscFrameSet {
    <# Rounded corners in modern terminals (emoji style), square corners elsewhere (present in every console font). #>
    param([Parameter(Mandatory = $true)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)

    if ($Style -eq 'Ascii') {
        return @{ TopLeft = [char]'+'; TopRight = [char]'+'; BottomLeft = [char]'+'; BottomRight = [char]'+'; Horizontal = [char]'-'; Vertical = [char]'|' }
    }
    if ($Style -eq 'Symbols') {
        return @{ TopLeft = [char]0x250C; TopRight = [char]0x2510; BottomLeft = [char]0x2514; BottomRight = [char]0x2518; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
    }
    return @{ TopLeft = [char]0x256D; TopRight = [char]0x256E; BottomLeft = [char]0x2570; BottomRight = [char]0x256F; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
}

$script:Icons = Get-WscIconSet $script:IconStyle
$script:Frame = Get-WscFrameSet $script:IconStyle
$script:IconPad = if ($script:IconStyle -eq 'Emoji') { ' ' } else { '  ' }

function Get-WscIcon { param([Parameter(Mandatory = $true)][string]$Name) return $script:Icons[$Name] + $script:IconPad }

function Format-WscDuration {
    param([Parameter(Mandatory = $true)][double]$Seconds)

    $inv = [Globalization.CultureInfo]::InvariantCulture
    $t = [TimeSpan]::FromTicks([long]([Math]::Max(0.0, $Seconds) * 10000000))
    if ($t.TotalHours -ge 1) { return [string]::Format($inv, '{0} h {1:00} min', [int][Math]::Floor($t.TotalHours), $t.Minutes) }
    if ($t.TotalMinutes -ge 1) { return [string]::Format($inv, '{0} min {1:00} s', $t.Minutes, $t.Seconds) }
    return [string]::Format($inv, '{0:0.0} s', $t.TotalSeconds)
}

function Send-WscUi {
    <# Forwards a console line to the GUI progress box while a GUI run is in progress. #>
    param([string]$Status, [string]$Text)
    if ($script:Ui -and $script:Ui.Sink) { & $script:Ui.Sink $Status $Text }
}

function Start-WscLog {
    <# Opens (or continues) today's log file and deletes the log files older than the retention. #>
    param([Parameter(Mandatory = $true)][string]$Directory, [int]$RetentionDays = 14)

    Stop-WscLog
    [void][IO.Directory]::CreateDirectory($Directory)
    $script:LogPath = Join-Path $Directory ('WebServicesClient_{0:yyyyMMdd}.log' -f (Get-Date))
    $stream = [IO.FileStream]::new($script:LogPath, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
    $script:LogWriter = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
    $script:LogWriter.AutoFlush = $true
    $limit = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -LiteralPath $Directory -Filter 'WebServicesClient_*.log' -File -ErrorAction SilentlyContinue |
        Where-Object LastWriteTime -lt $limit | Remove-Item -Force -ErrorAction SilentlyContinue
    return $script:LogPath
}

function Stop-WscLog {
    if ($script:LogWriter) { $script:LogWriter.Dispose(); $script:LogWriter = $null }
}

function Write-WscLog {
    <# One line in the log file only. The log never contains colours, icons or tokens. #>
    param(
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP')][string]$Level = 'INFO',
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Message
    )
    if ($script:LogWriter) { $script:LogWriter.WriteLine(('{0:yyyy-MM-ddTHH:mm:ss.fffzzz} [{1,-5}] {2}' -f (Get-Date), $Level, $Message)) }
}

function Write-WscBanner {
    <# Title card at the start of an execution, followed by the context rows (label -> @(Icon, Text)). #>
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Subtitle,
        [System.Collections.Specialized.OrderedDictionary]$Details
    )

    Write-WscLog 'STEP' "=== $Title v$($script:ToolVersion) ==="
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $v = $Details[$key]
            Write-WscLog 'INFO' ('{0}: {1}' -f $key, $(if ($v -is [array]) { $v[1] } else { $v }))
        }
    }
    if ($script:Quiet) { return }
    $C = $script:C; $F = $script:Frame; $width = 74
    $right = "v$($script:ToolVersion) $([char]0x00B7) Nicolas Fabert"
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }
    $left = "  $($script:Icons.Logo)  $Title"
    $gap = [Math]::Max(1, $width - ($left.Length - $script:Icons.Logo.Length + $iconWidth) - $right.Length - 2)
    Write-Host ''
    Write-Host ('  {0}{1}{2}{3}{4}' -f $C.Accent, $F.TopLeft, [string]::new($F.Horizontal, $width), $F.TopRight, $C.Reset)
    Write-Host ('  {0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}' -f $C.Accent, $F.Vertical, $C.Reset, $C.Bold, $left, $C.Reset, [string]::new(' ', $gap), $C.Dim, $right, '  ', ($C.Accent + $F.Vertical), $C.Reset)
    if ($Subtitle) {
        $sub = "     $Subtitle"
        if ($sub.Length -gt $width - 2) { $sub = $sub.Substring(0, $width - 5) + '...' }
        Write-Host ('  {0}{1}{2}{3}{4}{5}{0}{6}{2}' -f $C.Accent, $F.Vertical, $C.Reset, $C.Dim, $sub.PadRight($width), $C.Reset, $F.Vertical)
    }
    Write-Host ('  {0}{1}{2}{3}{4}' -f $C.Accent, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $value = $Details[$key]
            $icon, $text = if ($value -is [array]) { (Get-WscIcon $value[0]), $value[1] } else { '   ', $value }
            Write-Host ('     {0}{1}{2,-11}{3} {4}' -f $icon, $C.Dim, $key, $C.Reset, $text)
        }
    }
}

function Write-WscStep {
    <# Step header with a coloured number pill and an icon:  ─ 3/5 ─ 🔑  Sign-in #>
    param(
        [Parameter(Mandatory = $true)][int]$Number,
        [Parameter(Mandatory = $true)][int]$Total,
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Icon = 'Info'
    )

    Write-WscLog 'STEP' "[$Number/$Total] $Title"
    Send-WscUi 'Step' "[$Number/$Total] $Title"
    if ($script:Quiet) { return }
    $C = $script:C
    Write-Host ''
    Write-Host ('  {0} {1}/{2} {3} {4}{5}{6}{3}' -f $C.AccentBg, $Number, $Total, $C.Reset, (Get-WscIcon $Icon), $C.Bold, $Title)
}

function Write-WscItem {
    <# One indented result line with a status icon, also written to the log and to the GUI. #>
    param(
        [ValidateSet('Ok', 'Warn', 'Fail', 'Info', 'Skip', 'Block')][string]$Status = 'Info',
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
        [string]$Icon
    )

    $level = @{ Ok = 'OK'; Warn = 'WARN'; Fail = 'ERROR'; Info = 'INFO'; Skip = 'INFO'; Block = 'WARN' }[$Status]
    Write-WscLog $level $Text
    Send-WscUi $Status $Text
    if ($script:Quiet) { return }
    $color = @{ Ok = $script:C.Green; Warn = $script:C.Yellow; Fail = $script:C.Red; Info = ''; Skip = $script:C.Dim; Block = $script:C.Yellow }[$Status]
    $symbol = Get-WscIcon $(if ($Icon) { $Icon } else { $Status })
    $textColor = if ($Status -in 'Warn', 'Fail', 'Skip', 'Block') { $color } else { '' }
    Write-Host ('      {0}{1}{2}{3}{4}{2}' -f $color, $symbol, $script:C.Reset, $textColor, $Text)
}

function Write-WscSummary {
    <# Final summary card (label -> @(Icon, Text)). #>
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][System.Collections.Specialized.OrderedDictionary]$Values,
        [ValidateSet('Ok', 'Warn', 'Fail')][string]$Status = 'Ok'
    )

    foreach ($key in $Values.Keys) {
        $v = $Values[$key]
        Write-WscLog 'INFO' ('Summary - {0}: {1}' -f $key, $(if ($v -is [array]) { $v[1] } else { $v }))
    }
    if ($script:Quiet) { return }
    $C = $script:C; $F = $script:Frame; $width = 74
    $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red }[$Status]
    $icon = $script:Icons[@{ Ok = 'Done'; Warn = 'Warn'; Fail = 'Fail' }[$Status]]
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }
    $head = " $icon  $Title "
    $rest = [Math]::Max(2, $width - 1 - ($head.Length - $icon.Length + $iconWidth))
    Write-Host ''
    Write-Host ('  {0}{1}{2}{3}{4}{0}{5}{6}{7}' -f $color, $F.TopLeft, $F.Horizontal, $C.Bold, $head, ($C.Reset + $color), ([string]::new($F.Horizontal, $rest) + $F.TopRight), $C.Reset)
    foreach ($key in $Values.Keys) {
        $value = $Values[$key]
        $rowIcon, $text = if ($value -is [array]) { (Get-WscIcon $value[0]), $value[1] } else { '   ', $value }
        Write-Host ('    {0}{1}{2,-10}{3} {4}' -f $rowIcon, $C.Dim, $key, $C.Reset, $text)
    }
    Write-Host ('  {0}{1}{2}{3}{4}' -f $color, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    Write-Host ''
}

function Get-WscSignInText {
    <# One line that says who signs in and how (banner, window footer). #>
    param([Parameter(Mandatory = $true)][hashtable]$Configuration, [pscredential]$Credential)

    $c = $Configuration
    $dot = [char]0x00B7
    switch ([string]$c.Authentication) {
        'Basic' {
            $user = if ($Credential) { $Credential.UserName } elseif ($c.SignInUser) { $c.SignInUser } else { $c.Mailbox }
            return "Basic $dot user $user $dot password never written"
        }
        'Windows' {
            $user = if ($Credential) { $Credential.UserName } else { "current Windows account $([Environment]::UserDomainName)\$([Environment]::UserName)" }
            return "Windows $([string]$c.WindowsPackage) $dot $user $dot handshake traced"
        }
    }
    $server = switch ([string]$c.Authority) {
        'EntraID' { "Entra ID ($(if (Test-WscExchangeOnlineUrl -Url ([string]$c.EwsUrl)) { 'Exchange Online' } else { 'HMA or Exchange Online' }))" }
        'Auto' { 'the server Exchange names (AD FS or Entra ID)' }
        default { "AD FS $([string]$c.AdfsUrl)" }
    }
    switch ([string]$c.Context) {
        'Application' {
            $cred = if ([string]$c.CertificateThumbprint) { "certificate $($c.CertificateThumbprint)" } else { 'client secret' }
            return "Application $($c.AppClientId) $dot $server $dot $cred"
        }
        'Delegated' { return "User through the application $($c.AppClientId) $dot $server" }
        default { return "User $dot $server $dot client $($c.ClientId)" }
    }
}

function Write-WscRunBanner {
    <# Title card of a command-line run: scenario, mailbox, sign-in, access, EWS URL, writes, report and log. #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Settings,
        [pscredential]$Credential,
        [string]$LogPath,
        [switch]$NoReport
    )

    $scenario = Get-WscTestCatalog | Where-Object Name -eq $Settings.TestType
    $dot = [char]0x00B7
    $banner = [ordered]@{}
    $banner['Scenario'] = @('Target', "$($scenario.Name) $dot $($scenario.DisplayName)")
    $banner['Mailbox'] = @('People', $Settings.Mailbox)
    $access = Resolve-WscAccess -Configuration $Settings
    $accessText = switch ($access) {
        'Delegate' { "Delegate access as $(if ($Settings.SignInUser) { $Settings.SignInUser } else { 'the signed-in user' })" }
        'Impersonation' { 'Impersonation (ExchangeImpersonation header)' }
        default { 'Own mailbox of the signed-in user' }
    }
    $banner['Access'] = @('Shield', $accessText)
    $signIn = Get-WscSignInText -Configuration $Settings -Credential $Credential
    if ($scenario.SignIn -and [string]$Settings.Authentication -eq 'OAuth' -and [string]$Settings.Context -ne 'Application') {
        $how = try { if ((Get-WscSignInMode -Configuration $Settings).Mode -eq 'Window') { 'sign-in window' } else { 'device code' } } catch { 'sign-in window not available here' }
        $signIn = "$signIn $dot $how"
    }
    $banner['Sign-in'] = @('Key', $(if ($scenario.SignIn) { $signIn } else { "$signIn $dot no sign-in in this scenario" }))
    $banner['Protocol'] = @('Server', $(switch ([string]$Settings.Protocol) {
                'Graph' { 'Microsoft Graph (Exchange Online), forced' }
                'EWS' { 'EWS, forced' }
                default { 'Auto: Microsoft Graph for Exchange Online, EWS on-premises' }
            }))
    if ([string]$Settings.Protocol -ne 'Graph') {
        $banner['EWS'] = @('Server', $(if ([string]$Settings.Discovery -eq 'Autodiscover') { "Autodiscover$(if ($Settings.EwsUrl) { " (else $($Settings.EwsUrl))" })" } else { "$($Settings.EwsUrl) (manual)" }))
    }
    if ($scenario.Writes) {
        $banner['Writes'] = @('Shield', $(if ($Settings.AllowWrite) { "AUTHORISED $dot folder '$($Settings.FolderName)' $dot delete: $($Settings.DeleteMode)" } else { 'not authorised: the writing steps are blocked (-AllowWrite)' }))
    }
    $banner['Report'] = @('Report', $(if ($NoReport) { 'none (-NoReport)' } else { $Settings.OutputPath }))
    if ($LogPath) { $banner['Log'] = @('Log', $LogPath) }
    $subtitle = switch ([string]$Settings.Authentication) { 'Basic' { 'Basic' } 'Windows' { 'Windows (NTLM, Kerberos)' } default { "OAuth $dot $([string]$Settings.Context)" } }
    Write-WscBanner -Title 'Web Services Client for Exchange' -Subtitle "EWS and Microsoft Graph $dot $subtitle $dot test toolbox" -Details $banner
}

function Write-WscRunSummary {
    <# Final card of a command-line run: status, counts, first issue, report, log and what to do next. #>
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Result,
        [string]$ReportText = 'none (-NoReport)',
        [string]$LogPath
    )

    $dot = [char]0x00B7
    $n = $Result.Counts
    $values = [ordered]@{}
    $values['Status'] = @($(switch ($Result.Status) { 'Passed' { 'Ok' } 'Failed' { 'Fail' } 'Blocked' { 'Block' } default { 'Warn' } }), "$($Result.Status) $dot $($Result.TestType)")
    $values['Checks'] = @('Target', ('{0} passed {5} {1} warning(s) {5} {2} blocked {5} {3} failed {5} {4} skipped' -f $n.Passed, $n.Warning, $n.Blocked, $n.Failed, $n.Skipped, $dot))
    if ($Result.Protocol -eq 'Graph') { $values['Protocol'] = @('Server', "Microsoft Graph $dot $($Result.GraphUrl)") }
    elseif ($Result.Server -and $Result.Server.Version) { $values['Exchange'] = @('Server', "$($Result.Server.Version) $dot $($Result.EwsUrl)") }
    if ($Result.Error) { $values['First issue'] = @('Fail', $Result.Error) }
    $values['Duration'] = @('Clock', (Format-WscDuration $Result.DurationSeconds))
    $values['Report'] = @('Report', $ReportText)
    if ($LogPath) { $values['Log'] = @('Log', $LogPath) }
    $values['Next'] = @('Info', $(switch ($Result.Status) {
                'Passed' { 'Nothing to do. The HTTP trace of the report shows every request and response.' }
                'Blocked' { 'A writing step needs -AllowWrite (test mailbox only), or the item to work on: -ItemId, -ItemSubject.' }
                'Warning' { 'Read the warnings in the report: each one gives the cause and what to check.' }
                default { 'Open the report: the first Failed step gives the HTTP status, the EWS response code and the Exchange diagnostics.' }
            }))
    $title = switch ($Result.Status) { 'Passed' { 'Test passed' } 'Warning' { 'Test finished with warnings' } 'Blocked' { 'Test blocked' } default { 'Test failed' } }
    $card = switch ($Result.Status) { 'Passed' { 'Ok' } 'Failed' { 'Fail' } default { 'Warn' } }
    Write-WscSummary -Title $title -Values $values -Status $card
}
function Write-WscFreeBusyGrid {
    <#
        Free/busy as a small map in the console: one line per day and mailbox, one cell per hour from 07:00
        to 20:00 (local time of this computer), the strongest status of the hour. Also in the log, as text.
        Free  .   Tentative ~   Busy #   Out of office O   Working elsewhere w   No data ?
    #>
    param([Parameter(Mandatory = $true)][object[]]$Results, [Parameter(Mandatory = $true)][datetime]$StartUtc, [int]$IntervalMinutes = 30, [int]$Days = 7, [int]$FirstHour = 7, [int]$LastHour = 20)

    $C = $script:C
    $ascii = $script:IconStyle -eq 'Ascii'
    $glyph = if ($ascii) { @{ 0 = ' . '; 1 = ' ~ '; 2 = '###'; 3 = 'OOO'; 4 = ' w '; 5 = ' ? ' } }
             else { @{ 0 = [string]::new([char]0x00B7, 1).PadLeft(2).PadRight(3); 1 = [string]::new([char]0x2592, 3); 2 = [string]::new([char]0x2588, 3); 3 = [string]::new([char]0x2593, 3); 4 = [string]::new([char]0x2591, 3); 5 = ' ? ' } }
    $color = @{ 0 = $C.Dim; 1 = $C.Yellow; 2 = "$([char]27)[38;2;71;158;245m"; 3 = "$([char]27)[38;2;180;140;230m"; 4 = "$([char]27)[38;2;90;200;200m"; 5 = $C.Dim }
    if (-not $C.Reset) { $color = @{ 0 = ''; 1 = ''; 2 = ''; 3 = ''; 4 = ''; 5 = '' } }
    # Strongest first: out of office, busy, tentative, working elsewhere, free, no data.
    $rank = @{ 3 = 6; 2 = 5; 1 = 4; 4 = 3; 0 = 2; 5 = 1 }
    $startLocal = $StartUtc.ToLocalTime()
    $ruler = (($FirstHour..($LastHour - 1)) | ForEach-Object { '{0:00} ' -f $_ }) -join ''
    $shown = $false
    foreach ($r in $Results) {
        $view = [string]$r.MergedFreeBusy
        if (-not $view) { continue }
        # Exchange counts 'working elsewhere' as free in the merged view: the calendar items complete it.
        $codes = @{ free = 0; tentative = 1; busy = 2; oof = 3; workingelsewhere = 4 }
        $items = @($r.Events | Where-Object { $_ -and $_.StartUtc } | ForEach-Object { [pscustomobject]@{ S = [datetime]::SpecifyKind([datetime]$_.StartUtc, 'Utc'); E = [datetime]::SpecifyKind([datetime]$_.EndUtc, 'Utc'); C = $codes[([string]$_.BusyType).ToLowerInvariant()] } } | Where-Object { $null -ne $_.C })
        if (-not $shown -and -not $script:Quiet) { Write-Host ('      {0}{1,-12}{2}{3}' -f $C.Dim, 'Free/busy', $ruler, $C.Reset) }
        $shown = $true
        if (-not $script:Quiet) { Write-Host ('      {0}{1}{2}' -f $C.Bold, $r.Mailbox, $C.Reset) }
        Write-WscLog 'INFO' "Free/busy map of $($r.Mailbox) (hours $FirstHour-$LastHour, . free ~ tentative # busy O away w elsewhere)"
        for ($d = 0; $d -lt $Days; $d++) {
            $day = $startLocal.AddDays($d)
            $cells = [Text.StringBuilder]::new(); $plain = [Text.StringBuilder]::new()
            for ($h = $FirstHour; $h -lt $LastHour; $h++) {
                $from = $day.Date.AddHours($h).ToUniversalTime(); $to = $from.AddHours(1)
                $best = 5
                for ($t = $from; $t -lt $to; $t = $t.AddMinutes($IntervalMinutes)) {
                    $i = [int][Math]::Floor(($t - $StartUtc).TotalMinutes / $IntervalMinutes)
                    if ($i -ge 0 -and $i -lt $view.Length) { $v = [int][string]$view[$i]; if ($rank[$v] -gt $rank[$best] -or $best -eq 5) { $best = $v } }
                }
                foreach ($e in $items) { if ($e.S -lt $to -and $e.E -gt $from -and ($rank[$e.C] -gt $rank[$best] -or $best -eq 5)) { $best = $e.C } }
                [void]$cells.Append($color[$best] + $glyph[$best] + $C.Reset)
                [void]$plain.Append(@{ 0 = '.'; 1 = '~'; 2 = '#'; 3 = 'O'; 4 = 'w'; 5 = '?' }[$best])
            }
            $label = $day.ToString('ddd dd MMM', [Globalization.CultureInfo]::GetCultureInfo('en-GB'))
            if (-not $script:Quiet) { Write-Host ('      {0}{1,-12}{2}{3}' -f $C.Dim, $label, $C.Reset, $cells.ToString()) }
            Write-WscLog 'INFO' ('  {0}  {1}' -f $label, $plain.ToString())
        }
    }
    if ($shown -and -not $script:Quiet) {
        $legend = if ($ascii) { '. free  ~ tentative  # busy  O away  w elsewhere' } else { "$($glyph[0].Trim()) free  $([char]0x2592) tentative  $([char]0x2588) busy  $([char]0x2593) away  $([char]0x2591) elsewhere" }
        Write-Host ('      {0}{1,-12}{2}{3}' -f $C.Dim, '', $legend, $C.Reset)
    }
}