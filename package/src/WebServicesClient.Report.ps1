<#
.SYNOPSIS
    Web Services Client for Exchange - report files (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    One folder per execution, <FilePrefix>_<Scenario>_<yyyyMMdd-HHmmss>, with:
      <prefix>-Steps.csv      every check: stage, status, message, details, duration
      <prefix>-Folders.csv    folders returned by FindFolder (path, class, counts)
      <prefix>-Messages.csv   Inbox messages returned by FindItem (date, sender, subject)
      <prefix>-FreeBusy.csv   busy periods returned by GetUserAvailability, one line per event
      <prefix>-Actions.csv    what the run changed in the mailbox (create, send, reply, move, delete)
      <prefix>-Trace.csv      every HTTP request sent and the response received (secrets masked)
      <prefix>-Summary.json   the whole result, for scripts
      <prefix>.html           self-contained dashboard (templates\Report.template.html)
    CSV files: UTF-8 with BOM, configurable delimiter, cells starting with = + - @ prefixed with an
    apostrophe (no formula injection in Excel). Tokens and passwords are never part of the result.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:ReportColumns = [ordered]@{
    Steps    = @('Stage', 'Name', 'Status', 'Message', 'Details', 'DurationMs', 'TimestampUtc', 'Trace')
    Folders  = @('DisplayName', 'Depth', 'FolderClass', 'TotalCount', 'UnreadCount', 'ChildFolderCount', 'Path', 'FolderId', 'ParentFolderId')
    Messages = @('DateTimeReceived', 'From', 'Subject', 'IsRead', 'HasAttachments', 'Size', 'ItemClass', 'InternetMessageId', 'ItemId')
    FreeBusy = @('Mailbox', 'Start', 'End', 'BusyType', 'Subject', 'Location', 'IsMeeting', 'IsPrivate', 'StartUtc', 'EndUtc')
    Actions  = @('TimeUtc', 'Action', 'Target', 'Result', 'ItemId')
    Trace    = @('Sequence', 'TimestampUtc', 'Stage', 'Step', 'StepName', 'Method', 'Url', 'Operation', 'Label', 'StatusCode', 'Reason', 'Server', 'DurationMs', 'Repeated', 'Note', 'Request', 'Response')
}

function Format-WscDetails {
    <# Details of a step on one line: Key=Value; Key=Value (empty values left out). #>
    param([AllowNull()][System.Collections.IDictionary]$Details)

    if ($null -eq $Details) { return '' }
    $parts = foreach ($key in $Details.Keys) {
        $value = $Details[$key]
        if ($null -eq $value -or ([string]$value) -eq '') { continue }
        if ($value -is [array]) { $value = $value -join ', ' }
        '{0}={1}' -f $key, $value
    }
    return (@($parts) -join '; ')
}

function Format-WscCsvCell {
    param([AllowNull()][object]$Value, [Parameter(Mandatory = $true)][string]$Delimiter)

    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Collections.IDictionary]) { $text = Format-WscDetails $Value }
    elseif ($Value -is [array]) { $text = $Value -join ', ' }
    elseif ($Value -is [bool]) { $text = if ($Value) { 'True' } else { 'False' } }
    elseif ($Value -is [string]) {
        $text = $Value
        if ($text -match '^[=+\-@\t\r]') { $text = "'" + $text }
    }
    else { $text = [string]$Value }
    if ($text.Contains($Delimiter) -or $text.Contains('"') -or $text -match '[\r\n]') { $text = '"' + $text.Replace('"', '""') + '"' }
    return $text
}

function Write-WscCsv {
    param([AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory = $true)][string[]]$Columns, [Parameter(Mandatory = $true)][string]$Path, [string]$Delimiter = ';')

    $builder = [Text.StringBuilder]::new()
    [void]$builder.AppendLine((@($Columns | ForEach-Object { Format-WscCsvCell $_ $Delimiter }) -join $Delimiter))
    foreach ($row in @($Rows)) {
        [void]$builder.AppendLine((@($Columns | ForEach-Object { Format-WscCsvCell (Get-WscField $row $_) $Delimiter }) -join $Delimiter))
    }
    [IO.File]::WriteAllText($Path, $builder.ToString(), [Text.UTF8Encoding]::new($true))
}

function ConvertTo-WscEmbeddedJson {
    <# JSON safe inside a <script type="application/json"> block. #>
    param([AllowNull()][object]$Value)
    $json = ConvertTo-Json -InputObject $Value -Depth 8 -Compress
    if ([string]::IsNullOrEmpty($json)) { $json = 'null' }
    return $json.Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
}

function New-WscRunFolder {
    param([Parameter(Mandatory = $true)][string]$OutputPath, [Parameter(Mandatory = $true)][string]$Prefix, [string]$TestType = 'Run')
    $base = Join-Path $OutputPath ('{0}_{1}_{2}' -f $Prefix, $TestType, (Get-Date).ToString('yyyyMMdd-HHmmss'))
    $path = $base
    $n = 2
    while (Test-Path -LiteralPath $path) { $path = "$base-$n"; $n++ }
    [void][IO.Directory]::CreateDirectory($path)
    return $path
}

function Export-WscReport {
    <#
    .SYNOPSIS
        Writes the CSV, JSON and HTML files of one result in a new folder under OutputPath.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Result,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [string]$Prefix = 'WebServicesClient',
        [ValidateSet('Csv', 'Html')][string[]]$Formats = @('Csv', 'Html'),
        [ValidateSet(';', ',', "`t")][string]$Delimiter = ';'
    )

    $testType = [string](Get-WscField $Result 'TestType')
    $runPath = New-WscRunFolder -OutputPath $OutputPath -Prefix $Prefix -TestType $(if ($testType) { $testType } else { 'Run' })
    $data = [ordered]@{
        Steps    = @(Get-WscField $Result 'Steps')
        Folders  = @(Get-WscField $Result 'Folders')
        Messages = @(Get-WscField $Result 'Messages')
        FreeBusy = @(Get-WscField $Result 'FreeBusyEvents')
        Actions  = @(Get-WscField $Result 'Actions')
        Trace    = @(Get-WscField $Result 'Trace')
    }
    $files = [ordered]@{}
    if ($Formats -contains 'Csv') {
        foreach ($name in $data.Keys) {
            $files[$name] = Join-Path $runPath "$Prefix-$name.csv"
            Write-WscCsv -Rows $data[$name] -Columns $script:ReportColumns[$name] -Path $files[$name] -Delimiter $Delimiter
        }
    }
    $files.Summary = Join-Path $runPath "$Prefix-Summary.json"
    [IO.File]::WriteAllText($files.Summary, (ConvertTo-Json -InputObject $Result -Depth 8), [Text.UTF8Encoding]::new($false))

    if ($Formats -contains 'Html') {
        $summary = [ordered]@{}
        foreach ($key in 'Tool', 'Version', 'Status', 'TestType', 'Scenario', 'StartedUtc', 'CompletedUtc', 'DurationSeconds', 'Mailbox', 'SignInUser', 'Authentication', 'Context', 'Access',
            'WindowsPackage', 'Windows', 'Protocol', 'ProtocolSetting', 'GraphUrl', 'StageTitles', 'Discovery', 'EwsUrl', 'EwsUrlSource', 'ExchangeOnline', 'Authority', 'AuthorityUrl', 'AuthoritySource', 'TenantId', 'ClientId', 'SignIn', 'Token',
            'UserAgent', 'RequestServerVersion', 'Server', 'AllowWrite', 'Message', 'FreeBusy', 'FreeBusyWindow', 'Counts', 'Error') {
            $summary[$key] = Get-WscField $Result $key
        }
        $template = [IO.File]::ReadAllText((Join-Path $script:ToolRoot 'templates\Report.template.html'))
        $status = if ($summary.Status) { $summary.Status } else { 'Unknown' }
        $html = $template.Replace('{{TITLE}}', [Net.WebUtility]::HtmlEncode("Web Services Client for Exchange | $status | $testType"))
        $html = $html.Replace('{{SUMMARY_JSON}}', (ConvertTo-WscEmbeddedJson $summary))
        foreach ($name in $data.Keys) { $html = $html.Replace("{{$($name.ToUpperInvariant())_JSON}}", (ConvertTo-WscEmbeddedJson @($data[$name]))) }
        if ($html -match '\{\{[A-Z_]+\}\}') { throw "Report template marker not replaced: $($Matches[0])" }
        $files.Html = Join-Path $runPath "$Prefix.html"
        [IO.File]::WriteAllText($files.Html, $html, [Text.UTF8Encoding]::new($true))
    }
    [pscustomobject]@{ Directory = $runPath; Files = $files }
}
