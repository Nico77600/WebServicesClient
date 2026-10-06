#Requires -Version 7.4

<#
.SYNOPSIS
    Renders the graphics of the GitHub README from the Web Services Client for Exchange guide, in a light
    and a dark version: banner, principles, how it works, EWS or Microsoft Graph.

.DESCRIPTION
    GitHub renders Markdown only: the custom blocks of the guide (cards, flow) and its theme are lost. This
    tool renders them as images with the CSS of the built HTML guide and the icons of
    tools\Build-Documentation.ps1, so that the README and the guide always look the same. The README shows
    them with <picture>, which picks the light or dark image from the theme of the reader.

    Sources:
      docs\WebServicesClient-Guide.md        the cards block of the introduction, the version
      docs\WebServicesClient-Guide.html      the CSS (run tools\Build-Documentation.ps1 first)
      tools\Build-Documentation.ps1          the icons

    Screenshots: Microsoft Edge in headless mode, with a temporary profile, 2x resolution. Only local files
    are opened. Output: docs\images\readme-<name>-light.png and readme-<name>-dark.png.

.PARAMETER OutputFolder
    Default: docs\images next to the tools folder.

.PARAMETER KeepWork
    Keeps the work folder (the HTML pages of the graphics) and shows its path.

.EXAMPLE
    .\tools\Build-Documentation.ps1; .\tools\New-ReadmeImages.ps1

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0  (from EAS OAuth Mailbox 1.2.1)
    Part of : Web Services Client for Exchange (repository tool, not in the package)
#>
[CmdletBinding()]
param(
    [string]$OutputFolder,
    [switch]$KeepWork
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'docs\images' }

#region Assets of the guide ------------------------------------------------------------------------
function ConvertTo-ReadmeInline([string]$Text) {
    # Inline Markdown of a guide block (code, bold, italic) -> HTML.
    $h = [System.Net.WebUtility]::HtmlEncode($Text.Trim())
    $h = [regex]::Replace($h, '`([^`]+)`', '<code>$1</code>')
    $h = [regex]::Replace($h, '\*\*([^*]+)\*\*', '<strong>$1</strong>')
    return [regex]::Replace($h, '(?<![\w*])\*([^*\s][^*]*)\*(?![\w*])', '<em>$1</em>')
}

function Get-ReadmeAssets {
    param([string]$Root)
    $builder = Join-Path $Root 'tools\Build-Documentation.ps1'
    $guideHtml = Join-Path $Root 'docs\WebServicesClient-Guide.html'
    $guideMd = Join-Path $Root 'docs\WebServicesClient-Guide.md'
    if (-not (Test-Path $guideHtml)) { throw 'docs\WebServicesClient-Guide.html not found: run tools\Build-Documentation.ps1 first (it holds the CSS of the graphics).' }
    # Icons: the $Icons table of the documentation builder, read without running the builder.
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($builder, [ref]$null, [ref]$null)
    $assign = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$Icons' }, $true)
    if (-not $assign) { throw "Icon table not found in $builder." }
    $md = [IO.File]::ReadAllText($guideMd) -replace "`r`n", "`n"
    $blocks = foreach ($m in [regex]::Matches($md, '(?s)```(flow|cards)\n(.*?)\n```')) {
        [pscustomobject]@{ Kind = $m.Groups[1].Value; Lines = @($m.Groups[2].Value -split "`n" | Where-Object { $_.Trim() }) }
    }
    [pscustomobject]@{
        Icons   = & ([scriptblock]::Create($assign.Right.Extent.Text))
        Css     = [regex]::Match([IO.File]::ReadAllText($guideHtml), '(?s)<style>(.*?)</style>').Groups[1].Value
        Version = [regex]::Match($md, '(?m)^version:\s*(\S+)').Groups[1].Value
        Flows   = @($blocks | Where-Object Kind -eq 'flow')
        Cards   = @($blocks | Where-Object Kind -eq 'cards')
    }
}

function Get-ReadmeIcon([string]$Name, [string]$Class = 'icon') {
    $path = $assets.Icons[$Name]; if (-not $path) { $path = $assets.Icons['info'] }
    "<svg class=""$Class"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round"" stroke-linejoin=""round"">$path</svg>"
}

function ConvertTo-ReadmeFlow([string[]]$Lines, [switch]$Vertical) {
    # Vertical: the nodes are stacked, icon on the left, with a downward arrow and its label.
    $items = foreach ($l in $Lines) {
        $icon, $title, $sub = $l.Split('|', 3).ForEach({ $_.Trim() })
        $title = [System.Net.WebUtility]::HtmlEncode($title); $sub = [System.Net.WebUtility]::HtmlEncode($sub)
        if ($Vertical) {
            if ($icon -eq 'arrow') {
                $note = if ($sub) { "<span class=""flow-sub"">$sub</span>" } else { '' }
                "<div class=""rb-varrow""><svg viewBox=""0 0 12 30""><path d=""M6 1v26M1 21l5 6 5-6"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""/></svg><span class=""flow-label"">$title</span>$note</div>"
            } else {
                "<div class=""rb-vnode""><div class=""flow-icon"">$(Get-ReadmeIcon $icon)</div><div><div class=""flow-title"">$title</div><div class=""flow-text"">$sub</div></div></div>"
            }
        } elseif ($icon -eq 'arrow') {
            $class = if ($title -or $sub) { 'flow-arrow' } else { 'flow-arrow rb-bare' }
            "<div class=""$class""><span class=""flow-label"">$title</span><svg viewBox=""0 0 40 12""><path d=""M0 6h36M31 1l6 5-6 5"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""/></svg><span class=""flow-sub"">$sub</span></div>"
        } else {
            "<div class=""flow-node""><div class=""flow-icon"">$(Get-ReadmeIcon $icon)</div><div class=""flow-title"">$title</div><div class=""flow-text"">$sub</div></div>"
        }
    }
    $class = if ($Vertical) { 'flow rb-vflow' } else { 'flow rb-flow' }
    "<div class=""$class"">$($items -join '')</div>"
}

function ConvertTo-ReadmeCards([string[]]$Lines, [string]$Class = '') {
    $items = foreach ($l in $Lines) {
        $icon, $title, $text = $l.Split('|', 3).ForEach({ $_.Trim() })
        "<div class=""card-item""><div class=""card-icon"">$(Get-ReadmeIcon $icon)</div><div><div class=""card-title"">$(ConvertTo-ReadmeInline $title)</div><div class=""card-text"">$(ConvertTo-ReadmeInline $text)</div></div></div>"
    }
    "<div class=""cards $Class"">$($items -join '')</div>"
}

function Get-ReadmePill([string]$Text, [string]$Tone) { "<span class=""rb-pill"" style=""--tone: var(--cp-$Tone)"">$Text</span>" }
#endregion

#region Styles of the graphics, on top of the CSS of the guide -------------------------------------
$Script:ReadmeCss = @'
html, body { background: #ffffff; }
html[data-theme="dark"], html[data-theme="dark"] body { background: #0d1117; }
:root { --cp-info: #0078d4; --cp-violet: #7c3aed; --cp-teal: #0d9488; }
html[data-theme="dark"] { --cp-info: #4da6ff; --cp-violet: #a78bfa; --cp-teal: #2dd4bf; }
body { display: block; margin: 0; padding: 0; }
.canvas { padding: 6px; }
.rb-pill { display: inline-block; padding: 1px 10px; margin: 8px 6px 0 0; border-radius: 999px; font-size: 11.5px; font-weight: 600; line-height: 1.6;
  color: var(--tone); background: color-mix(in srgb, var(--tone) 11%, transparent); border: 1px solid color-mix(in srgb, var(--tone) 38%, transparent); }
.rb-caption { font-size: 11.5px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: var(--cp-accent); margin: 0 0 8px 4px; }
.rb-caption span { color: var(--cp-text-muted); font-weight: 600; letter-spacing: 0.04em; text-transform: none; font-size: 12.5px; }
/* Before / after */
.rb-bench { display: grid; gap: 14px; padding: 20px 22px; border-radius: 16px; background: var(--cp-surface); border: 1px solid var(--cp-border); }
.rb-row { display: grid; grid-template-columns: 250px minmax(0, 1fr); gap: 18px; align-items: center; }
.rb-row .label { font-size: 13.5px; font-weight: 650; color: var(--cp-text); } .rb-row .label span { display: block; font-weight: 500; font-size: 12px; color: var(--cp-text-muted); margin-top: 2px; }
.rb-bars { display: grid; gap: 6px; }
.rb-bar { display: flex; align-items: center; gap: 10px; font-size: 12.5px; color: var(--cp-text-muted); }
.rb-bar i { display: block; height: 18px; border-radius: 6px; min-width: 6px; }
.rb-bar.old i { background: color-mix(in srgb, var(--cp-text-muted) 45%, transparent); } .rb-bar.new i { background: var(--cp-accent); }
.rb-bar b { color: var(--cp-text); font-weight: 650; } .rb-bar .bad { color: var(--cp-danger); font-weight: 600; }
.rb-legend { display: flex; gap: 18px; font-size: 12px; color: var(--cp-text-muted); margin: 0 0 4px 4px; } .rb-legend i { display: inline-block; width: 12px; height: 12px; border-radius: 4px; margin-right: 6px; vertical-align: -1px; }
/* Banner */
.rb-hero { margin: 0; padding: 32px 36px 30px; }
.rb-hero-grid { position: relative; display: grid; grid-template-columns: minmax(0, 1fr) 240px; gap: 34px; align-items: center; }
.rb-hero h1 { font-size: 35px; }
.rb-hero .lead { margin: 18px 0 0; font-size: 17px; max-width: none; }
.rb-hero .badges { margin: 20px 0 0; }
.rb-stats { position: relative; display: grid; gap: 10px; }
.rb-stat { display: flex; align-items: center; gap: 14px; padding: 12px 16px; border-radius: 14px; background: var(--cp-panel-strong); border: 1px solid var(--cp-border); box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08); }
.rb-stat b { font-size: 30px; line-height: 1; color: var(--cp-accent); font-weight: 750; min-width: 40px; text-align: center; }
.rb-stat span { font-size: 13px; color: var(--cp-text-muted); line-height: 1.35; }
.rb-stat strong { display: block; color: var(--cp-text); font-size: 14px; }
/* Cards and flows */
.cards { margin: 0; }
.rb-cards2 { grid-template-columns: 1fr 1fr; }
.rb-flow { margin: 0; flex-wrap: nowrap; padding: 18px; gap: 4px; }
.rb-flow .flow-node { flex: 1 1 0; min-width: 0; padding: 14px 10px; }
.rb-flow .flow-title { font-size: 13.5px; overflow-wrap: anywhere; }
.rb-flow .flow-arrow { min-width: 0; width: 84px; flex: 0 0 84px; }
.rb-flow .flow-arrow.rb-bare { width: 46px; flex-basis: 46px; }
.rb-flow .flow-sub { max-width: 84px; }
.rb-space { height: 18px; }
/* How it works: vertical pipeline and the three modes */
.rb-hiw { display: grid; grid-template-columns: minmax(0, 1.08fr) minmax(0, 1fr); gap: 16px; align-items: stretch; }
.rb-col { display: flex; flex-direction: column; }
.rb-vflow { flex: 1; flex-direction: column; flex-wrap: nowrap; align-items: stretch; justify-content: center; gap: 0; margin: 0; padding: 16px 18px; }
.rb-vnode { display: flex; align-items: center; gap: 14px; padding: 11px 16px; border-radius: 12px; background: var(--cp-surface); border: 1px solid var(--cp-border); }
.rb-vnode .flow-icon { margin: 0; flex-shrink: 0; }
.rb-vnode .flow-text { margin-top: 1px; }
.rb-varrow { display: flex; align-items: center; gap: 10px; min-height: 36px; padding-left: 31px; }
.rb-varrow svg { width: 12px; height: 28px; color: var(--cp-accent); flex-shrink: 0; }
.rb-varrow .flow-sub { max-width: none; font-size: 12px; }
.rb-modes { flex: 1; display: flex; flex-direction: column; gap: 10px; }
.rb-modes .card-item { flex: 1; align-items: center; }
.rb-modes .card-title { display: flex; align-items: center; gap: 8px; }
.rb-chip { font-size: 11px; font-weight: 600; padding: 0 8px; border-radius: 999px; border: 1px solid var(--cp-border); color: var(--cp-text-muted); }
.rb-chip.hot { color: var(--cp-accent-fg); background: var(--cp-accent); border-color: var(--cp-accent); }
'@
#endregion

#region Rendering (Microsoft Edge, headless) -------------------------------------------------------
function Save-Screenshot([string]$Html, [string]$Png, [int]$Width, [int]$Height, [int]$Scale = 1) {
    $url = 'file:///' + ($Html -replace '\\', '/')
    $profilePath = Join-Path $work 'edge-profile'
    if (Test-Path $Png) { Remove-Item $Png -Force }
    # Start-Process, not &: an Edge helper process can keep the output pipe open after the capture.
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$profilePath`"", "--window-size=$Width,$Height", "--force-device-scale-factor=$Scale", "--screenshot=`"$Png`"", "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden
    $deadline = (Get-Date).AddSeconds(45)
    while (-not (Test-Path $Png) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
    if (-not $proc.WaitForExit(10000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    if (-not (Test-Path $Png)) { throw "Screenshot not written: $Png" }
}

function Get-PageHeight([string]$Html, [int]$Width) {
    # Height of the .canvas element: the page writes it in body[data-h], read with --dump-dom.
    $url = 'file:///' + ($Html -replace '\\', '/')
    $dom = Join-Path $work ('dom-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.html')
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$(Join-Path $work 'edge-profile')`"", "--window-size=$Width,2000", '--dump-dom', "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden -RedirectStandardOutput $dom
    if (-not $proc.WaitForExit(45000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    # Edge helper processes inherit the output handle: read in shared mode, retry until written.
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

function New-ReadmeGraphic {
    # One graphic, light and dark: HTML page -> height measured by Edge -> 2x screenshot.
    param([string]$Name, [string]$Body, [int]$Width)
    $pages = @{}
    foreach ($theme in 'light', 'dark') {
        $html = "<!doctype html><html lang=""en"" data-theme=""$theme""><head><meta charset=""utf-8""><style>$($assets.Css)`n$($Script:ReadmeCss)</style></head>" +
            "<body><div class=""canvas"" style=""width:$($Width)px"">$Body</div><script>document.body.setAttribute('data-h', Math.ceil(document.querySelector('.canvas').getBoundingClientRect().height));</script></body></html>"
        $pages[$theme] = Join-Path $work "readme-$Name-$theme.html"
        [IO.File]::WriteAllText($pages[$theme], $html, [Text.UTF8Encoding]::new($false))
    }
    $height = Get-PageHeight $pages['light'] $Width
    foreach ($theme in 'light', 'dark') { Save-Screenshot $pages[$theme] (Join-Path $OutputFolder "readme-$Name-$theme.png") $Width $height 2 }
}
#endregion

#region Main ---------------------------------------------------------------------------------------
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge not found: it takes the screenshots (headless mode).' }
$work = Join-Path ([IO.Path]::GetTempPath()) ('wsc-readme-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work, $OutputFolder -Force | Out-Null
$Script:assets = Get-ReadmeAssets -Root $root
if ($assets.Cards.Count -lt 1) { throw 'The guide must hold the cards block of the introduction.' }
$mid = '&middot;'

try {
    Write-Host 'Rendering the README graphics (light and dark, 2x)...'

    # Banner: the hero of the guide, with the key figures.
    $badges = @(
        "<span class=""badge badge-accent"">Version $($assets.Version)</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'terminal' 'icon-sm')PowerShell 7.4+</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'server' 'icon-sm')Exchange 2019 &middot; SE &middot; Online</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'key' 'icon-sm')AD FS &middot; Entra ID &middot; Basic &middot; NTLM &middot; Kerberos</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'tag' 'icon-sm')MIT license</span>"
    ) -join ''
    $banner = "<header class=""hero rb-hero""><div class=""rb-hero-grid""><div>" +
        "<div class=""hero-top""><div class=""hero-logo"">$(Get-ReadmeIcon 'mail')</div><div><div class=""eyebrow"">EWS $mid Microsoft Graph $mid OAuth $mid Basic $mid Windows</div><h1>Web Services Client for Exchange</h1></div></div>" +
        "<p class=""lead"">A <strong>test toolbox for Exchange mailboxes</strong>, on-premises through EWS and in Exchange Online through Microsoft Graph: it signs in like a real client, as a user or an application, does what a client does &mdash; folders, read, send, reply, move, delete, <strong>free/busy</strong> &mdash; and shows <strong>every request sent and every response received</strong>.</p>" +
        "<div class=""badges"">$badges</div></div>" +
        "<div class=""rb-stats"">" +
        "<div class=""rb-stat""><b>4</b><span><strong>ways to sign in</strong>AD FS, Entra ID, Basic, Windows</span></div>" +
        "<div class=""rb-stat""><b>3</b><span><strong>contexts</strong>user, delegated app, application</span></div>" +
        "<div class=""rb-stat""><b>0</b><span><strong>token written</strong>secrets masked in the trace</span></div>" +
        "</div></div></header>"
    New-ReadmeGraphic -Name 'banner' -Body $banner -Width 1080

    # Principles: the cards block of the introduction of the guide.
    New-ReadmeGraphic -Name 'principles' -Body (ConvertTo-ReadmeCards $assets.Cards[0].Lines 'rb-cards2') -Width 1080

    # How it works: the stages as a vertical pipeline, and four ways to use them.
    $stages = @(
        'search | Autodiscover | where is EWS, like Outlook — or the URL given'
        'arrow | no sign-in |'
        'shield | Prerequisites | certificate, authentication offered, sign-in server, forged token'
        'arrow | sign in | window, device code, certificate, password, NTLM'
        'key | Sign-in | OAuth token and its claims, Basic, or the Windows handshake leg by leg'
        'arrow | X-AnchorMailbox |'
        'server | Endpoint | the Inbox opened: version, front end and back end, affinity'
        'arrow | same connection |'
        'mail | Operations | folders, messages, free/busy, then the writes with -AllowWrite'
    )
    $uses = @(
        [pscustomobject]@{ Icon = 'search'; Name = 'Discovery'; Chip = '<span class="rb-chip">no sign-in</span>'; Text = 'Before opening EWS to users or to an application: publishing, certificates, <strong>which sign-in does Exchange offer this mailbox?</strong>'; Pills = (Get-ReadmePill 'Autodiscover' 'teal') + (Get-ReadmePill 'NTLM challenge' 'info') }
        [pscustomobject]@{ Icon = 'check'; Name = 'ReadOnly'; Chip = '<span class="rb-chip hot">changes nothing</span>'; Text = 'What a client sees: <strong>folders, messages and free/busy</strong> of one or several mailboxes, in a scheduling-assistant view.'; Pills = (Get-ReadmePill 'Folder tree' 'success') + (Get-ReadmePill 'Everyone free' 'violet') }
        [pscustomobject]@{ Icon = 'refresh'; Name = 'MailCycle &middot; Full'; Chip = '<span class="rb-chip">-AllowWrite</span>'; Text = 'On a test mailbox: test folder, <strong>send, reply, move, delete</strong>; <code>SeedData</code> fills an empty mailbox.'; Pills = (Get-ReadmePill 'Delivery checked' 'success') + (Get-ReadmePill 'Changes tab' 'warning') }
        [pscustomobject]@{ Icon = 'file'; Name = 'HTTP trace'; Chip = '<span class="rb-chip">every scenario</span>'; Text = 'Under each check, <strong>the request sent and the response received</strong>: SOAP and JSON indented, NTLM and Kerberos decoded.'; Pills = (Get-ReadmePill 'Tokens masked' 'success') + (Get-ReadmePill 'Trace.csv' 'teal') }
    )
    $useHtml = ($uses | ForEach-Object { "<div class=""card-item""><div class=""card-icon"">$(Get-ReadmeIcon $_.Icon)</div><div><div class=""card-title"">$($_.Name) $($_.Chip)</div><div class=""card-text"">$($_.Text)</div><div>$($_.Pills)</div></div></div>" }) -join ''
    $howItWorks = "<div class=""rb-hiw""><div class=""rb-col""><div class=""rb-caption"">The stages <span>$mid always in this order</span></div>$(ConvertTo-ReadmeFlow $stages -Vertical)</div>" +
        "<div class=""rb-col""><div class=""rb-caption"">Four ways to use them <span>$mid read-only for AD FS, Entra ID and Exchange</span></div><div class=""rb-modes"">$useHtml</div></div></div>"
    New-ReadmeGraphic -Name 'how-it-works' -Body $howItWorks -Width 1080

    # EWS or Microsoft Graph: the choice of -Protocol Auto.
    $protocols = @(
        'server | On-premises → EWS | Exchange 2019 / SE with AD FS, Entra ID (HMA), Basic or Windows: SOAP, `X-AnchorMailbox`, affinity cookies, impersonation or delegate access.'
        'cloud | Exchange Online → Microsoft Graph | EWS is retired there (HTTP 403 `X-EWS-Policy-Reason`): `mailFolders`, `messages`, `sendMail`, `reply`, `move`, `getSchedule`, with an Entra ID token for Graph.'
    )
    $protocolHtml = "<div class=""rb-caption"">EWS or Microsoft Graph <span>$mid -Protocol Auto chooses once Autodiscover has found the mailbox: the same scenarios, the same report</span></div>$(ConvertTo-ReadmeCards $protocols 'rb-cards2')"
    New-ReadmeGraphic -Name 'protocols' -Body $protocolHtml -Width 1080

    Get-ChildItem $OutputFolder -Filter 'readme-*.png' | Select-Object Name, @{ n = 'KB'; e = { [math]::Round($_.Length / 1KB) } } | Format-Table -AutoSize | Out-String | Write-Host
} finally {
    # Edge helper processes of the temporary profile, if any are left.
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like "*$work*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    if ($KeepWork) { Write-Host "Work folder: $work" } else { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
}
#endregion