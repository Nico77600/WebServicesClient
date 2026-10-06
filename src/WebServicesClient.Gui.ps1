<#
.SYNOPSIS
    Web Services Client for Exchange - graphical interface (dot-sourced by WebServicesClient.psm1).

.DESCRIPTION
    WPF window, same design as EAS OAuth Mailbox: Fluent theme of Windows 11 with .NET 9 or later
    (PowerShell 7.5+), light or dark like Windows, the accent colour of the report; classic WPF controls
    with the same colours on PowerShell 7.4.

    Layout: a header; on the left the sign-in (four cards: OAuth AD FS, OAuth Entra ID, Basic, Windows),
    the context (user, delegated application, application) and the target (mailbox, signing-in account,
    Autodiscover or manual URL); on the right the scenario with its options (writes, item, free/busy
    mailboxes) and the progress, one line per check; at the bottom the actions.

    It runs exactly the same engine as the command line (Invoke-WscMailboxTest, Export-WscReport). The
    run happens on the window thread, kept responsive while waiting (Invoke-WscUiPump), and can be
    cancelled. Close is the cancel button of the window (Esc too); a Closing handler is attached only
    while a test runs. Ctrl+C is ignored in the console while the window is open.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:Gui = $null
$script:GuiAuthentication = @(
    [pscustomobject]@{ Text = 'OAuth - On-prem AD FS'; Kind = 'OAuth'; Title = 'On-prem AD FS'; Detail = ''; Glyph = 0xE8D7; Authentication = 'OAuth'; Authority = 'ADFS'
        Hint = 'Exchange 2019 CU13+ or SE with AD FS: the token comes from AD FS.' }
    [pscustomobject]@{ Text = 'OAuth - Entra ID'; Kind = 'OAuth'; Title = 'Entra ID'; Detail = 'HMA or Exchange Online'; Glyph = 0xE753; Authentication = 'OAuth'; Authority = 'EntraID'
        Hint = 'Entra ID signs in (password, MFA, Conditional Access): Exchange on-prem with hybrid modern authentication, or Exchange Online.' }
    [pscustomobject]@{ Text = 'Basic - On-prem'; Kind = 'Basic'; Title = 'On-prem'; Detail = ''; Glyph = 0xE77B; Authentication = 'Basic'; Authority = $null
        Hint = 'User name and password with every request, protected only by TLS. Exchange Online no longer accepts it.' }
    [pscustomobject]@{ Text = 'Windows - NTLM / Kerberos'; Kind = 'Windows'; Title = 'NTLM / Kerberos'; Detail = 'On-prem'; Glyph = 0xE7EF; Authentication = 'Windows'; Authority = $null
        Hint = 'Negotiate, NTLM or Kerberos: the handshake is done by the tool and every leg is in the trace. Kerberos needs a domain controller in reach.' }
)
$script:GuiContexts = @(
    [pscustomobject]@{ Name = 'User'; Text = 'User - signs in with a Microsoft client' }
    [pscustomobject]@{ Name = 'Delegated'; Text = 'Delegated - your application on behalf of the user' }
    [pscustomobject]@{ Name = 'Application'; Text = 'Application - client credentials, impersonation' }
)
$script:GuiProtocols = @(
    [pscustomobject]@{ Name = 'Auto'; Text = 'Auto - Graph for Exchange Online, EWS on-prem' }
    [pscustomobject]@{ Name = 'EWS'; Text = 'EWS' }
    [pscustomobject]@{ Name = 'Graph'; Text = 'Microsoft Graph (Exchange Online)' }
)
$script:GuiIconFont = 'Segoe Fluent Icons, Segoe MDL2 Assets'

function Get-WscGuiXaml {
    <# The window. Colours come from the Fluent theme resources (or the classic fallback of Set-WscGuiTheme). #>
    @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="1240" Height="900" MinWidth="1000" MinHeight="640" WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="True">
  <Window.Resources>
    <Style x:Key="WscCard" TargetType="Border">
      <Setter Property="Background" Value="{DynamicResource CardBackgroundFillColorDefaultBrush}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource CardStrokeColorDefaultBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="18,14,18,16"/>
      <Setter Property="Margin" Value="0,0,0,12"/>
    </Style>
    <Style x:Key="WscCardTitle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,10"/>
      <Setter Property="Foreground" Value="{DynamicResource TextFillColorPrimaryBrush}"/>
    </Style>
    <Style x:Key="WscLabel" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="{DynamicResource TextFillColorSecondaryBrush}"/>
      <Setter Property="Margin" Value="0,8,0,4"/>
    </Style>
    <Style x:Key="WscIcon" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Segoe Fluent Icons, Segoe MDL2 Assets"/>
    </Style>
    <Style x:Key="WscMethod" TargetType="ListBoxItem">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListBoxItem">
            <Border x:Name="Card" Margin="0,0,8,8" Padding="12,10,10,10" CornerRadius="6" BorderThickness="2"
                    Background="{DynamicResource ControlFillColorDefaultBrush}" BorderBrush="{DynamicResource ControlStrokeColorDefaultBrush}">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Card" Property="Background" Value="{DynamicResource ControlFillColorSecondaryBrush}"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="Card" Property="BorderBrush" Value="{DynamicResource AccentFillColorDefaultBrush}"/>
                <Setter TargetName="Card" Property="Background" Value="{DynamicResource WscAccentSoft}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid x:Name="Root" Background="{DynamicResource ApplicationBackgroundBrush}">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Grid x:Name="Header" Margin="24,18,24,14">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <Border Width="46" Height="46" CornerRadius="10" Background="{DynamicResource WscBrand}" VerticalAlignment="Center">
        <TextBlock Style="{StaticResource WscIcon}" Text="&#xE715;" FontSize="22" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
      </Border>
      <StackPanel Grid.Column="1" Margin="14,0,0,0" VerticalAlignment="Center">
        <TextBlock Text="EXCHANGE MAILBOX TOOLBOX - EWS AND MICROSOFT GRAPH" FontSize="11" FontWeight="SemiBold" Foreground="{DynamicResource WscBrandText}"/>
        <TextBlock Text="Web Services Client for Exchange" FontSize="24" FontWeight="SemiBold" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
        <TextBlock FontSize="13" Foreground="{DynamicResource TextFillColorSecondaryBrush}" TextTrimming="CharacterEllipsis"
                   Text="Choose how to sign in, the mailbox and a scenario, then run it: same checks, same HTTP trace and same report as the command line."/>
      </StackPanel>
      <TextBlock x:Name="Version" Grid.Column="2" FontSize="12" Foreground="{DynamicResource TextFillColorSecondaryBrush}" VerticalAlignment="Top"/>
    </Grid>

    <Grid Grid.Row="1" Margin="24,0,24,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="480"/>
        <ColumnDefinition Width="16"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <ScrollViewer x:Name="SettingsScroll" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="0,0,4,0">
        <StackPanel>
          <Border x:Name="Method" Style="{StaticResource WscCard}">
            <StackPanel>
              <TextBlock Text="Sign-in" Style="{StaticResource WscCardTitle}"/>
              <ListBox x:Name="Authentication" ItemContainerStyle="{StaticResource WscMethod}" BorderThickness="0" Background="Transparent" Padding="0"
                       ScrollViewer.HorizontalScrollBarVisibility="Disabled" ScrollViewer.VerticalScrollBarVisibility="Disabled">
                <ListBox.ItemsPanel>
                  <ItemsPanelTemplate><UniformGrid Columns="2"/></ItemsPanelTemplate>
                </ListBox.ItemsPanel>
              </ListBox>
              <TextBlock x:Name="MethodHint" TextWrapping="Wrap" Margin="0,2,0,0" FontSize="12" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
              <StackPanel x:Name="OAuthPanel">
                <TextBlock Text="Context" Style="{StaticResource WscLabel}"/>
                <ComboBox x:Name="Context"/>
                <StackPanel x:Name="AppPanel">
                  <TextBlock Text="Client ID of your application" Style="{StaticResource WscLabel}"/>
                  <TextBox x:Name="AppClientId"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="12"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel>
                      <TextBlock x:Name="SecretLabel" Text="Client secret (never written)" Style="{StaticResource WscLabel}"/>
                      <PasswordBox x:Name="ClientSecret"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2">
                      <TextBlock x:Name="CertificateLabel" Text="or certificate thumbprint" Style="{StaticResource WscLabel}"/>
                      <TextBox x:Name="CertificateThumbprint"/>
                    </StackPanel>
                  </Grid>
                </StackPanel>
                <StackPanel x:Name="AdfsPanel">
                  <TextBlock Text="AD FS URL (ends with /adfs)" Style="{StaticResource WscLabel}"/>
                  <TextBox x:Name="AdfsUrl"/>
                </StackPanel>
              </StackPanel>
              <StackPanel x:Name="WindowsPanel">
                <TextBlock Text="Windows package" Style="{StaticResource WscLabel}"/>
                <ComboBox x:Name="WindowsPackage"/>
                <CheckBox x:Name="CurrentAccount" Margin="0,10,0,0" Content="Use the current Windows account"/>
              </StackPanel>
              <Grid x:Name="PasswordPanel">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="12"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Text="User (UPN or DOMAIN\user)" Style="{StaticResource WscLabel}"/>
                  <TextBox x:Name="UserName"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Text="Password (never written)" Style="{StaticResource WscLabel}"/>
                  <PasswordBox x:Name="Password"/>
                </StackPanel>
              </Grid>
            </StackPanel>
          </Border>

          <Border x:Name="Target" Style="{StaticResource WscCard}">
            <StackPanel>
              <TextBlock Text="Mailbox" Style="{StaticResource WscCardTitle}" Margin="0,0,0,2"/>
              <TextBlock Text="Mailbox tested (SMTP)" Style="{StaticResource WscLabel}"/>
              <TextBox x:Name="Mailbox"/>
              <TextBlock x:Name="SignInUserLabel" Text="Account that signs in (empty = the mailbox; another = delegate access)" Style="{StaticResource WscLabel}"/>
              <TextBox x:Name="SignInUser"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="12"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Text="Find EWS with" Style="{StaticResource WscLabel}"/>
                  <ComboBox x:Name="Discovery"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Text="Open the mailbox as" Style="{StaticResource WscLabel}"/>
                  <ComboBox x:Name="Access"/>
                </StackPanel>
              </Grid>
              <TextBlock Text="Protocol" Style="{StaticResource WscLabel}"/>
              <ComboBox x:Name="Protocol"/>
              <TextBlock x:Name="EwsUrlLabel" Text="EWS URL" Style="{StaticResource WscLabel}"/>
              <TextBox x:Name="EwsUrl"/>
            </StackPanel>
          </Border>
        </StackPanel>
      </ScrollViewer>

      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Border x:Name="Scenario" Style="{StaticResource WscCard}">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="230"/>
              <ColumnDefinition Width="16"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Text="Scenario" Style="{StaticResource WscCardTitle}" Grid.ColumnSpan="3"/>
            <ComboBox x:Name="TestType" Grid.Row="1" VerticalAlignment="Top"/>
            <TextBlock x:Name="Description" Grid.Row="1" Grid.Column="2" TextWrapping="Wrap" VerticalAlignment="Center" FontSize="12" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
            <Grid Grid.Row="2" Grid.ColumnSpan="3" Margin="0,4,0,0">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <StackPanel>
                <TextBlock x:Name="ItemLabel" Text="Message to reply to, move or delete: subject contains (empty = test message of the tool)" TextWrapping="Wrap" Style="{StaticResource WscLabel}"/>
                <TextBox x:Name="ItemSubject"/>
              </StackPanel>
              <StackPanel Grid.Column="2">
                <TextBlock x:Name="FreeBusyLabel" Text="Free/busy of (addresses separated by ; empty = the mailbox)" TextWrapping="Wrap" Style="{StaticResource WscLabel}"/>
                <TextBox x:Name="FreeBusyMailboxes"/>
              </StackPanel>
            </Grid>
            <WrapPanel Grid.Row="3" Grid.ColumnSpan="3" Margin="0,10,0,0">
              <CheckBox x:Name="AllowWrite" Margin="0,0,24,0" Content="Allow changes to the mailbox (test mailbox only)"/>
              <CheckBox x:Name="DeviceCode" Content="Sign in with a device code (no sign-in window)"/>
            </WrapPanel>
            <Border x:Name="WarningBox" Grid.Row="4" Grid.ColumnSpan="3" Margin="0,10,0,0" Padding="12,8" CornerRadius="6" BorderThickness="1">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="WarningIcon" Style="{StaticResource WscIcon}" FontSize="16" Margin="0,1,10,0" VerticalAlignment="Top"/>
                <TextBlock x:Name="Warning" Grid.Column="1" TextWrapping="Wrap" FontSize="12" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
              </Grid>
            </Border>
          </Grid>
        </Border>
        <Border x:Name="Progress" Grid.Row="1" Style="{StaticResource WscCard}">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Progress" Style="{StaticResource WscCardTitle}"/>
              <Border x:Name="StatusPill" Grid.Column="1" CornerRadius="10" Padding="10,2" VerticalAlignment="Top">
                <TextBlock x:Name="Status" FontSize="12" FontWeight="SemiBold"/>
              </Border>
            </Grid>
            <Border x:Name="CodeBanner" Grid.Row="1" Visibility="Collapsed" Margin="0,0,0,12" Padding="14,10" CornerRadius="6" BorderThickness="1"
                    Background="{DynamicResource WscAccentSoft}" BorderBrush="{DynamicResource AccentFillColorDefaultBrush}">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Text="Device code: type it on the sign-in page, from any device" FontSize="12" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
                  <TextBlock x:Name="CodeText" FontSize="26" FontWeight="SemiBold" FontFamily="Cascadia Mono, Consolas" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
                  <TextBlock x:Name="CodePage" FontSize="12" Foreground="{DynamicResource AccentTextFillColorPrimaryBrush}"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                  <Button x:Name="CopyCode" Content="Copy the code" Margin="0,0,8,0"/>
                  <Button x:Name="OpenPage" Content="Open the page"/>
                </StackPanel>
              </Grid>
            </Border>
            <TextBlock x:Name="LogEmpty" Grid.Row="2" FontSize="13" Foreground="{DynamicResource TextFillColorTertiaryBrush}" TextWrapping="Wrap"
                       Text="The checks appear here as they run. The report shows, for every check, the HTTP request sent and the response received."/>
            <ScrollViewer x:Name="LogScroll" Grid.Row="2" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
              <ItemsControl x:Name="Log" Margin="0,0,12,0">
                <ItemsControl.ItemTemplate>
                  <DataTemplate>
                    <Grid Margin="{Binding Margin}">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="22"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock Text="{Binding Glyph}" Foreground="{Binding Brush}" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="13" Margin="0,3,0,0" VerticalAlignment="Top"/>
                      <TextBlock Grid.Column="1" Text="{Binding Text}" TextWrapping="Wrap" FontSize="{Binding Size}" FontWeight="{Binding Weight}" Foreground="{Binding TextBrush}"/>
                      <TextBlock Grid.Column="2" Text="{Binding Time}" FontSize="11" Margin="10,2,0,0" Foreground="{DynamicResource TextFillColorTertiaryBrush}"/>
                    </Grid>
                  </DataTemplate>
                </ItemsControl.ItemTemplate>
              </ItemsControl>
            </ScrollViewer>
          </Grid>
        </Border>
      </Grid>
    </Grid>

    <Border x:Name="Actions" Grid.Row="2" Padding="24,12" BorderThickness="0,1,0,0"
            BorderBrush="{DynamicResource DividerStrokeColorDefaultBrush}" Background="{DynamicResource LayerFillColorDefaultBrush}">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal">
          <Button x:Name="Run" MinWidth="150" Padding="16,6" Margin="0,0,8,0">
            <StackPanel Orientation="Horizontal">
              <TextBlock Style="{StaticResource WscIcon}" Text="&#xE768;" Margin="0,2,8,0"/>
              <TextBlock Text="Run the test"/>
            </StackPanel>
          </Button>
          <Button x:Name="Cancel" Content="Cancel" MinWidth="96" IsEnabled="False"/>
        </StackPanel>
        <TextBlock x:Name="Footer" Grid.Column="1" Margin="16,0" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" FontSize="12"
                   Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
        <StackPanel Grid.Column="2" Orientation="Horizontal">
          <Button x:Name="OpenReport" Content="Open the report" Margin="0,0,8,0" IsEnabled="False"/>
          <Button x:Name="OpenFolder" Content="Open the folder" Margin="0,0,8,0" IsEnabled="False"/>
          <Button x:Name="Close" Content="Close" MinWidth="96" IsCancel="True"/>
        </StackPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
'@
}

function New-WscTestForm {
    <#
    .SYNOPSIS
        Builds the window (without showing it). Used by Show-WscTestGui, the tests and the documentation tool.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Configuration, [ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System')

    $look = Initialize-WscGuiTheme -Theme $Theme
    $catalog = @(Get-WscTestCatalog)
    $window = [Windows.Markup.XamlReader]::Parse((Get-WscGuiXaml))
    $window.Title = "Web Services Client for Exchange $($script:ToolVersion)"
    Set-WscGuiTheme -Window $window -Theme $look
    $controls = @{}
    foreach ($name in 'Root', 'Header', 'Version', 'Method', 'Authentication', 'MethodHint', 'OAuthPanel', 'Context', 'AppPanel', 'AppClientId', 'SecretLabel', 'ClientSecret', 'CertificateLabel',
        'CertificateThumbprint', 'AdfsPanel', 'AdfsUrl', 'WindowsPanel', 'WindowsPackage', 'CurrentAccount', 'PasswordPanel', 'UserName', 'Password', 'Target', 'Mailbox', 'SignInUserLabel',
        'SignInUser', 'Discovery', 'Access', 'Protocol', 'EwsUrlLabel', 'EwsUrl', 'Scenario', 'TestType', 'Description', 'ItemLabel', 'ItemSubject', 'FreeBusyLabel', 'FreeBusyMailboxes', 'AllowWrite',
        'DeviceCode', 'WarningBox', 'WarningIcon', 'Warning', 'Progress', 'StatusPill', 'Status', 'CodeBanner', 'CodeText', 'CodePage', 'CopyCode', 'OpenPage', 'LogEmpty', 'LogScroll',
        'Log', 'Actions', 'Run', 'Cancel', 'Footer', 'OpenReport', 'OpenFolder', 'Close', 'SettingsScroll') {
        $controls[$name] = $window.FindName($name)
    }
    if ($look.Fluent) { $controls.Run.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'AccentButtonStyle') }
    else {
        $controls.Run.SetResourceReference([Windows.Controls.Control]::BackgroundProperty, 'AccentFillColorDefaultBrush')
        $controls.Run.Foreground = [Windows.Media.Brushes]::White
    }
    $controls.Version.Text = "v$($script:ToolVersion)  " + [char]0x00B7 + '  Nicolas Fabert'
    foreach ($method in $script:GuiAuthentication) { [void]$controls.Authentication.Items.Add((New-WscGuiMethodCard -Method $method)) }
    $current = switch ([string]$Configuration.Authentication) { 'Basic' { 'Basic' } 'Windows' { 'Windows' } default { if ([string]$Configuration.Authority -eq 'EntraID') { 'EntraID' } else { 'ADFS' } } }
    $index = [Array]::FindIndex([object[]]$script:GuiAuthentication, [Predicate[object]] { param($x) ($x.Authentication -eq $current) -or ($x.Authority -eq $current) })
    $controls.Authentication.SelectedIndex = [Math]::Max(0, $index)
    foreach ($c in $script:GuiContexts) { [void]$controls.Context.Items.Add($c.Text) }
    $controls.Context.SelectedIndex = [Math]::Max(0, [Array]::FindIndex([object[]]$script:GuiContexts, [Predicate[object]] { param($x) $x.Name -eq [string]$Configuration.Context }))
    foreach ($p in 'Negotiate', 'NTLM', 'Kerberos') { [void]$controls.WindowsPackage.Items.Add($p) }
    $controls.WindowsPackage.SelectedItem = [string]$Configuration.WindowsPackage
    foreach ($d in 'Autodiscover', 'Manual') { [void]$controls.Discovery.Items.Add($d) }
    $controls.Discovery.SelectedItem = [string]$Configuration.Discovery
    foreach ($a in 'Auto', 'Self', 'Delegate', 'Impersonation') { [void]$controls.Access.Items.Add($a) }
    foreach ($p in $script:GuiProtocols) { [void]$controls.Protocol.Items.Add($p.Text) }
    $controls.Protocol.SelectedIndex = [Math]::Max(0, [Array]::FindIndex([object[]]$script:GuiProtocols, [Predicate[object]] { param($x) $x.Name -eq [string]$Configuration.Protocol }))
    $controls.Access.SelectedItem = [string]$Configuration.Access
    foreach ($key in 'AdfsUrl', 'EwsUrl', 'Mailbox', 'SignInUser', 'AppClientId', 'CertificateThumbprint') { $controls[$key].Text = [string]$Configuration[$key] }
    $controls.FreeBusyMailboxes.Text = (@($Configuration.FreeBusyMailboxes) -join '; ')
    $controls.CurrentAccount.IsChecked = $true
    foreach ($s in $catalog) { [void]$controls.TestType.Items.Add($s.Name) }
    $controls.AllowWrite.IsChecked = [bool]$Configuration.AllowWrite
    $controls.DeviceCode.IsChecked = [string]$Configuration.SignIn -eq 'DeviceCode'

    $items = [Collections.ObjectModel.ObservableCollection[object]]::new()
    $controls.Log.ItemsSource = $items
    $script:Gui = @{
        Form = $window; Controls = $controls; Configuration = $Configuration.Clone(); Catalog = $catalog; Theme = $look
        Running = $false; LastReport = $null; LastFolder = $null
        Items = $items; Lines = [Collections.Generic.List[string]]::new(); CodePage = $null
        ClosingGuard = [ComponentModel.CancelEventHandler] {
            param($sender, $e)
            $e.Cancel = $true
            if ($script:Ui) { $script:Ui.Cancel = $true }
            Add-WscGuiLine 'Warn' 'A test is running: it stops at the next check, then the window can be closed.'
        }
    }
    Set-WscGuiStatus 'Ready' 'Ready'
    $controls.Footer.Text = "Reports: $($Configuration.OutputPath)"

    foreach ($c in 'TestType', 'Authentication', 'Context', 'Discovery', 'Access', 'WindowsPackage', 'Protocol') { $controls[$c].Add_SelectionChanged({ Update-WscGuiScenario }) }
    foreach ($c in 'DeviceCode', 'AllowWrite', 'CurrentAccount') { $controls[$c].Add_Click({ Update-WscGuiScenario }) }
    foreach ($c in 'EwsUrl', 'SignInUser', 'Mailbox') { $controls[$c].Add_LostFocus({ Update-WscGuiScenario }) }
    $controls.Run.Add_Click({ Invoke-WscGuiRun })
    $controls.Cancel.Add_Click({
            if ($script:Ui) { $script:Ui.Cancel = $true; Add-WscGuiLine 'Warn' 'Cancellation requested: the run stops at the next check.' }
        })
    $controls.OpenReport.Add_Click({ if ($script:Gui.LastReport) { Start-Process -FilePath $script:Gui.LastReport } })
    $controls.OpenFolder.Add_Click({ if ($script:Gui.LastFolder) { Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$($script:Gui.LastFolder)`"" } })
    $controls.CopyCode.Add_Click({ [Windows.Clipboard]::SetText([string]$script:Gui.Controls.CodeText.Text) })
    $controls.OpenPage.Add_Click({ if ($script:Gui.CodePage) { Start-Process -FilePath $script:Gui.CodePage } })

    $area = [Windows.SystemParameters]::WorkArea
    $window.Width = [Math]::Min($window.Width, $area.Width)
    $window.Height = [Math]::Min($window.Height, $area.Height)
    $window.MinWidth = [Math]::Min($window.MinWidth, $area.Width)
    $window.MinHeight = [Math]::Min($window.MinHeight, $area.Height)

    $controls.TestType.SelectedItem = if ($catalog.Name -contains [string]$Configuration.TestType) { [string]$Configuration.TestType } else { 'ReadOnly' }
    Update-WscGuiScenario
    [pscustomobject]@{ Form = $window; Controls = $controls; Lines = $script:Gui.Lines; Items = $items }
}

function Get-WscGuiChoice {
    <# The sign-in card, the context and the scenario chosen in the window. #>
    $g = $script:Gui
    $c = $g.Controls
    [pscustomobject]@{
        Method   = $script:GuiAuthentication[[Math]::Max(0, $c.Authentication.SelectedIndex)]
        Context  = $script:GuiContexts[[Math]::Max(0, $c.Context.SelectedIndex)].Name
        Scenario = $g.Catalog | Where-Object Name -eq ([string]$c.TestType.SelectedItem) | Select-Object -First 1
    }
}

function Update-WscGuiScenario {
    $g = $script:Gui
    if (-not $g) { return }
    $c = $g.Controls
    $choice = Get-WscGuiChoice
    $s = $choice.Scenario
    if (-not $s) { return }
    $m = $choice.Method
    $oauth = $m.Authentication -eq 'OAuth'
    $app = $oauth -and $choice.Context -eq 'Application'
    $delegated = $oauth -and $choice.Context -eq 'Delegated'
    $windows = $m.Authentication -eq 'Windows'
    $basic = $m.Authentication -eq 'Basic'
    $c.Description.Text = "$($s.DisplayName). $($s.Description)"
    $c.MethodHint.Text = $m.Hint
    $show = { param($Control, [bool]$Visible) $Control.Visibility = if ($Visible) { 'Visible' } else { 'Collapsed' } }
    & $show $c.OAuthPanel $oauth
    & $show $c.AppPanel ($app -or $delegated)
    & $show $c.AdfsPanel ($oauth -and $m.Authority -eq 'ADFS')
    & $show $c.WindowsPanel $windows
    $current = $windows -and [bool]$c.CurrentAccount.IsChecked
    & $show $c.PasswordPanel ($basic -or ($windows -and -not $current))
    $c.SecretLabel.Text = if ($app) { 'Client secret (never written)' } else { 'Client secret if confidential (never written)' }
    $c.CertificateThumbprint.IsEnabled = $app -and $m.Authority -eq 'EntraID'
    $c.SignInUser.IsEnabled = -not $app
    $c.SignInUserLabel.Text = if ($app) { 'Account that signs in: none, the application signs in alone' } else { 'Account that signs in (empty = the mailbox; another = delegate access)' }
    $protocol = $script:GuiProtocols[[Math]::Max(0, $c.Protocol.SelectedIndex)].Name
    $manual = [string]$c.Discovery.SelectedItem -eq 'Manual'
    $c.EwsUrlLabel.Text = if ($manual) { 'EWS URL (https://<server>/EWS/Exchange.asmx)' } else { 'EWS URL used only if Autodiscover gives no answer (optional)' }
    $stages = @($s.Stages)
    $c.ItemSubject.IsEnabled = @($stages | Where-Object { $_ -in 'ReplyMail', 'MoveMail', 'DeleteMail' }).Count -gt 0
    $c.FreeBusyMailboxes.IsEnabled = $stages -contains 'FreeBusy'
    $c.AllowWrite.IsEnabled = $s.Writes
    $c.DeviceCode.IsEnabled = $oauth -and -not $app -and $s.SignIn
    $how = if ($c.DeviceCode.IsChecked) { 'sign-in with a code typed on any device' } else { 'sign-in in a window (password, MFA)' }
    $online = Test-WscExchangeOnlineUrl -Url ([string]$c.EwsUrl.Text)
    if ($protocol -eq 'Graph' -and -not ($oauth -and $m.Authority -eq 'EntraID')) {
        Set-WscGuiNotice 'Microsoft Graph accepts only OAuth with Entra ID: choose the Entra ID card, or the protocol Auto or EWS.' Critical
    }
    elseif ($online -and ($basic -or $windows) -and $s.SignIn) {
        Set-WscGuiNotice 'Exchange Online accepts only OAuth with Entra ID for EWS: choose the Entra ID card.' Critical
    }
    elseif ($online -and $oauth -and $m.Authority -eq 'ADFS') {
        Set-WscGuiNotice 'Exchange Online accepts only Entra ID tokens: choose the Entra ID card.' Critical
    }
    elseif ($s.Writes -and -not $c.AllowWrite.IsChecked) {
        Set-WscGuiNotice 'This scenario changes the mailbox (folder, messages): the writing steps are blocked until "Allow changes to the mailbox" is checked. Use a test mailbox.' Caution
    }
    elseif ($s.Writes) {
        Set-WscGuiNotice "Changes allowed: test folder '$($g.Configuration.FolderName)', a test message sent to $(if ($g.Configuration.Recipient) { $g.Configuration.Recipient } else { 'the mailbox' }), reply, move and delete ($($g.Configuration.DeleteMode)). Only the message named above, or the test message of the tool." Caution
    }
    elseif ($app) {
        Set-WscGuiNotice 'The application signs in alone (client credentials) and impersonates the mailbox: permission full_access_as_app (Entra ID), certificate or client secret.' Info
    }
    elseif ($windows) {
        $who = if ($current) { "the current account $([Environment]::UserDomainName)\$([Environment]::UserName)" } else { 'the account above' }
        Set-WscGuiNotice "Windows $([string]$c.WindowsPackage.SelectedItem) with $($who): every leg of the handshake is in the report. Out of the domain, give the account and use NTLM." Info
    }
    elseif (-not $s.SignIn) {
        Set-WscGuiNotice 'No sign-in: only requests without credentials, a forged token or a user that does not exist are sent.' Info
    }
    else {
        $ctx = if ($delegated) { 'through your application, on behalf of the user' } else { 'with a Microsoft client' }
        Set-WscGuiNotice "$(if ($basic) { 'Basic: user name and password with every request (TLS only)' } else { "$($m.Title) $ctx, $how" }). Read-only: nothing is changed in the mailbox." Info
    }
}

function Invoke-WscGuiRun {
    $g = $script:Gui
    $c = $g.Controls
    $choice = Get-WscGuiChoice
    $cfg = $g.Configuration.Clone()
    foreach ($key in 'AdfsUrl', 'EwsUrl', 'Mailbox', 'SignInUser', 'AppClientId', 'CertificateThumbprint') { $cfg[$key] = $c[$key].Text.Trim() }
    $cfg.TestType = [string]$choice.Scenario.Name
    $cfg.Authentication = $choice.Method.Authentication
    if ($choice.Method.Authority) { $cfg.Authority = $choice.Method.Authority }
    $cfg.Context = if ($cfg.Authentication -eq 'OAuth') { $choice.Context } else { 'User' }
    $cfg.Discovery = [string]$c.Discovery.SelectedItem
    $cfg.Protocol = $script:GuiProtocols[[Math]::Max(0, $c.Protocol.SelectedIndex)].Name
    $cfg.Access = [string]$c.Access.SelectedItem
    $cfg.WindowsPackage = [string]$c.WindowsPackage.SelectedItem
    $cfg.AllowWrite = $c.AllowWrite.IsEnabled -and [bool]$c.AllowWrite.IsChecked
    $cfg.FreeBusyMailboxes = @(([string]$c.FreeBusyMailboxes.Text) -split '[,;\s]+' | Where-Object { $_ })
    if ($c.DeviceCode.IsChecked) { $cfg.SignIn = 'DeviceCode' } elseif ([string]$cfg.SignIn -eq 'DeviceCode') { $cfg.SignIn = 'Auto' }
    if ($cfg.Context -eq 'Application') { $cfg.SignInUser = '' }

    Clear-WscGuiProgress
    $problems = @((Test-WscConfiguration -Configuration $cfg).Problems)
    $credential = $null
    $secret = $null
    if ($c.PasswordPanel.Visibility -eq 'Visible' -and $choice.Scenario.SignIn) {
        $user = if ($c.UserName.Text.Trim()) { $c.UserName.Text.Trim() } elseif ($cfg.SignInUser) { $cfg.SignInUser } else { [string]$cfg.Mailbox }
        if ($c.Password.SecurePassword.Length -eq 0) { $problems += 'Enter the password.' }
        else { $credential = [pscredential]::new($user, $c.Password.SecurePassword.Copy()) }
    }
    if ($c.AppPanel.Visibility -eq 'Visible' -and $c.ClientSecret.SecurePassword.Length) { $secret = $c.ClientSecret.SecurePassword.Copy() }
    if ($cfg.Context -eq 'Application' -and $choice.Scenario.SignIn -and -not $secret -and -not ($cfg.CertificateThumbprint -and $cfg.Authority -eq 'EntraID')) { $problems += 'Enter the client secret of the application, or a certificate thumbprint (Entra ID).' }
    if ($problems.Count) {
        foreach ($problem in $problems) { Add-WscGuiLine 'Fail' $problem }
        Set-WscGuiStatus 'Fix the values on the left.' 'Failed'
        return
    }
    $item = $c.ItemSubject.Text.Trim()

    $script:Ui = @{ Sink = { param($Status, $Text) Add-WscGuiLine $Status $Text }; Pump = { Invoke-WscGuiPump }; Cancel = $false }
    try {
        $g.Running = $true
        $g.Form.add_Closing($g.ClosingGuard)
        foreach ($b in 'Run', 'OpenReport', 'OpenFolder', 'Close') { $c[$b].IsEnabled = $false }
        $c.Cancel.IsEnabled = $true
        Set-WscGuiStatus "Running $($cfg.TestType)..." 'Running'
        $c.Footer.Text = "$($choice.Method.Text)  " + [char]0x00B7 + "  $($cfg.Context)  " + [char]0x00B7 + "  $($cfg.Mailbox)"
        Write-WscLog 'STEP' "GUI run: $($cfg.TestType) ($($cfg.Authentication), $($cfg.Context)) for $($cfg.Mailbox)"
        $result = Invoke-WscMailboxTest -Configuration $cfg -TestType $cfg.TestType -Credential $credential -ClientSecret $secret -ItemSubject $item
        $report = Export-WscReport -Result $result -OutputPath $cfg.OutputPath -Prefix $cfg.ReportPrefix -Formats $cfg.ReportFormats -Delimiter $cfg.CsvDelimiter
        $g.LastFolder = $report.Directory
        $g.LastReport = Get-WscField $report.Files 'Html'
        $c.Footer.Text = "Report: $($report.Directory)"
        $n = $result.Counts
        Set-WscGuiStatus ("{0}  {1}  {2}/{3} checks passed" -f $result.Status, [char]0x00B7, $n.Passed, @($result.Steps).Count) $result.Status
    }
    catch {
        Add-WscGuiLine 'Fail' $_.Exception.Message
        Set-WscGuiStatus 'Failed - see the progress' 'Failed'
    }
    finally {
        $g.Form.remove_Closing($g.ClosingGuard)
        $script:Ui = $null
        $g.Running = $false
        foreach ($b in 'Run', 'Close') { $c[$b].IsEnabled = $true }
        $c.Cancel.IsEnabled = $false
        $c.OpenReport.IsEnabled = [bool]$g.LastReport
        $c.OpenFolder.IsEnabled = [bool]$g.LastFolder
        if ($credential) { $credential.Password.Dispose() }
        if ($secret) { $secret.Dispose() }
    }
}

#region Theme, progress and window helpers (same as EAS OAuth Mailbox) -------------------------------


function New-WscGuiBrush {
    param([Parameter(Mandatory = $true)][string]$Color)
    $brush = [Windows.Media.SolidColorBrush]::new([Windows.Media.ColorConverter]::ConvertFromString($Color))
    $brush.Freeze()
    return $brush
}

function Initialize-WscGuiTheme {
    <#
        Loads WPF and applies the theme to the application: Fluent (.NET 9+), light or dark as Windows
        (System), or Light / Dark for the documentation images. Returns Fluent and Dark.
    #>
    param([ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System')

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
    $dark = $Theme -eq 'Dark'
    if ($Theme -eq 'System') { $dark = Test-WscGuiDarkMode }
    # One application per process: created once, never shut down by a closed window.
    $app = [Windows.Application]::Current
    if (-not $app) {
        $app = [Windows.Application]::new()
        $app.ShutdownMode = [Windows.ShutdownMode]::OnExplicitShutdown
    }
    $fluent = $null -ne [Windows.Application].GetProperty('ThemeMode')
    if ($fluent) {
        # ThemeMode is the Fluent theme of WPF (.NET 9 and later); its name is experimental in .NET 9.
        # Light or Dark, never System: without the setting (Windows Server 2016) WPF would pick dark and the
        # colours of the window (light) would not match.
        $app.ThemeMode = [Windows.ThemeMode]::new($(if ($dark) { 'Dark' } else { 'Light' }))
    }
    [pscustomobject]@{ Fluent = $fluent; Dark = $dark; Application = $app }
}

function Test-WscGuiDarkMode {
    <# Windows shows the applications in dark mode (AppsUseLightTheme = 0); light when the setting is missing. #>
    $personalize = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction SilentlyContinue
    return [string](Get-WscField $personalize 'AppsUseLightTheme') -eq '0'
}

function Set-WscGuiTheme {
    <#
        Colours of the window: the accent of the report on the Fluent accent resources, the status colours,
        and with the classic theme (.NET 8) the Fluent resources the window uses, with the report palette.
    #>
    param([Parameter(Mandatory = $true)][Windows.Window]$Window, [Parameter(Mandatory = $true)][pscustomobject]$Theme)

    $dark = $Theme.Dark
    $r = $Window.Resources
    $set = { param([string[]]$Keys, [string]$LightColor, [string]$DarkColor) $b = New-WscGuiBrush $(if ($dark) { $DarkColor } else { $LightColor }); foreach ($k in $Keys) { $r[$k] = $b } }
    if (-not $Theme.Fluent) {
        & $set 'ApplicationBackgroundBrush' '#F7F4EF' '#202020'
        & $set 'CardBackgroundFillColorDefaultBrush' '#FFFFFF' '#2B2B2B'
        & $set 'CardStrokeColorDefaultBrush', 'ControlStrokeColorDefaultBrush' '#DEDEDE' '#3D3D3D'
        & $set 'ControlFillColorDefaultBrush' '#FFFFFF' '#2D2D2D'
        & $set 'ControlFillColorSecondaryBrush' '#F5F5F5' '#323232'
        & $set 'DividerStrokeColorDefaultBrush' '#DEDEDE' '#3D3D3D'
        & $set 'LayerFillColorDefaultBrush' '#FCFBF8' '#262626'
        & $set 'TextFillColorPrimaryBrush' '#242424' '#FFFFFF'
        & $set 'TextFillColorSecondaryBrush' '#5C5C5C' '#C5C5C5'
        & $set 'TextFillColorTertiaryBrush' '#8A8A8A' '#9A9A9A'
    }
    # The accent of the report instead of the accent colour of Windows.
    & $set 'AccentFillColorDefaultBrush', 'AccentButtonBackground', 'AccentButtonBorderBrush' '#B11F4B' '#FD8EA1'
    & $set 'AccentFillColorSecondaryBrush', 'AccentButtonBackgroundPointerOver' '#E6B11F4B' '#E6FD8EA1'
    & $set 'AccentFillColorTertiaryBrush', 'AccentButtonBackgroundPressed' '#CCB11F4B' '#CCFD8EA1'
    & $set 'AccentTextFillColorPrimaryBrush' '#9A1A41' '#FD8EA1'
    & $set 'WscBrand' '#B11F4B' '#B11F4B'
    & $set 'WscBrandText' '#B11F4B' '#FD8EA1'
    & $set 'WscAccentSoft' '#14B11F4B' '#33FD8EA1'
    & $set 'WscSuccess' '#16A34A' '#4ADE80'
    & $set 'WscCaution' '#D97706' '#FBBF24'
    & $set 'WscCritical' '#DC2626' '#F87171'
    & $set 'WscInfoBackground' '#F3F3F3' '#2E2E2E'
    & $set 'WscInfoBorder' '#E0E0E0' '#3D3D3D'
    & $set 'WscCautionBackground' '#FFF7E8' '#33FBBF24'
    & $set 'WscCautionBorder' '#F5D7A1' '#66FBBF24'
    & $set 'WscCriticalBackground' '#FDECEC' '#33F87171'
    & $set 'WscCriticalBorder' '#F4B4B4' '#66F87171'
    & $set 'WscSuccessBackground' '#EAF7EE' '#334ADE80'
}

function Invoke-WscGuiPump {
    <# Lets the window repaint and handle clicks during a run (the WPF equivalent of DoEvents). #>
    $frame = [Windows.Threading.DispatcherFrame]::new()
    [void][Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Background,
        [Windows.Threading.DispatcherOperationCallback] { param($f) $f.Continue = $false; $null }, $frame)
    [Windows.Threading.Dispatcher]::PushFrame($frame)
}

function New-WscGuiMethodCard {
    <# Content of one sign-in method card: icon, kind (OAuth, Basic), name. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Method)

    $panel = [Windows.Controls.StackPanel]::new()
    $icon = [Windows.Controls.TextBlock]::new()
    $icon.Text = [string][char]$Method.Glyph
    $icon.FontFamily = [Windows.Media.FontFamily]::new($script:GuiIconFont)
    $icon.FontSize = 18
    $icon.Margin = [Windows.Thickness]::new(0, 0, 0, 6)
    $icon.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'WscBrandText')
    $kind = [Windows.Controls.TextBlock]::new()
    $kind.Text = $Method.Kind
    $kind.FontSize = 11
    $kind.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'TextFillColorSecondaryBrush')
    $name = [Windows.Controls.TextBlock]::new()
    $name.Text = $Method.Title
    $name.FontSize = 13
    $name.FontWeight = [Windows.FontWeights]::SemiBold
    $name.TextWrapping = [Windows.TextWrapping]::Wrap
    $name.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'TextFillColorPrimaryBrush')
    foreach ($part in $icon, $kind, $name) { [void]$panel.Children.Add($part) }
    if ($Method.Detail) {
        $detail = [Windows.Controls.TextBlock]::new()
        $detail.Text = $Method.Detail
        $detail.FontSize = 11
        $detail.Margin = [Windows.Thickness]::new(0, 2, 0, 0)
        $detail.TextWrapping = [Windows.TextWrapping]::Wrap
        $detail.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'TextFillColorSecondaryBrush')
        [void]$panel.Children.Add($detail)
    }
    $item = [Windows.Controls.ListBoxItem]::new()
    $item.Content = $panel
    $item.ToolTip = $Method.Text
    return $item
}

function Set-WscGuiNotice {
    <# The coloured box under the scenario: Info, Caution or Critical. #>
    param([Parameter(Mandatory = $true)][string]$Text, [ValidateSet('Info', 'Caution', 'Critical')][string]$Level = 'Info')

    $c = $script:Gui.Controls
    $c.Warning.Text = $Text
    $c.WarningIcon.Text = [string][char]$(switch ($Level) { 'Caution' { 0xE7BA } 'Critical' { 0xEA39 } default { 0xE946 } })
    $c.WarningIcon.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, $(switch ($Level) { 'Caution' { 'WscCaution' } 'Critical' { 'WscCritical' } default { 'TextFillColorSecondaryBrush' } }))
    $c.WarningBox.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, $(switch ($Level) { 'Caution' { 'WscCautionBackground' } 'Critical' { 'WscCriticalBackground' } default { 'WscInfoBackground' } }))
    $c.WarningBox.SetResourceReference([Windows.Controls.Border]::BorderBrushProperty, $(switch ($Level) { 'Caution' { 'WscCautionBorder' } 'Critical' { 'WscCriticalBorder' } default { 'WscInfoBorder' } }))
}

function Add-WscGuiLine {
    <# One line of the progress: icon and colour of its status; a device code also shows in its own box. #>
    param([string]$Status, [string]$Text)

    $g = $script:Gui
    if (-not $g) { return }
    $glyphs = @{ Step = 0xE76C; Ok = 0xE73E; Warn = 0xE7BA; Fail = 0xEA39; Info = 0xE946; Skip = 0xE72A; Block = 0xE733 }
    $colours = @{ Step = 'WscBrandText'; Ok = 'WscSuccess'; Warn = 'WscCaution'; Fail = 'WscCritical'; Info = 'TextFillColorSecondaryBrush'; Skip = 'TextFillColorTertiaryBrush'; Block = 'WscCaution' }
    $key = if ($glyphs.ContainsKey($Status)) { $Status } else { 'Info' }
    $step = $Status -eq 'Step'
    $shown = if ($step) { $Text -replace '^\[(\d+/\d+)\]\s*', '$1   ' } else { $Text }
    $window = $g.Form
    $g.Items.Add([pscustomobject]@{
            Glyph     = [string][char]$glyphs[$key]
            Brush     = $window.TryFindResource($colours[$key])
            Text      = $shown
            TextBrush = $window.TryFindResource($(if ($key -in 'Info', 'Skip') { 'TextFillColorSecondaryBrush' } else { 'TextFillColorPrimaryBrush' }))
            Weight    = if ($step) { [Windows.FontWeights]::SemiBold } else { [Windows.FontWeights]::Normal }
            Size      = if ($step) { 14 } else { 13 }
            Margin    = if ($step) { [Windows.Thickness]::new(0, $(if ($g.Items.Count) { 12 } else { 0 }), 0, 4) } else { [Windows.Thickness]::new(0, 2, 0, 2) }
            Time      = (Get-Date).ToString('HH:mm:ss')
        })
    $g.Lines.Add("[$Status] $Text")
    # The device code: big, with Copy and Open buttons (the console line is "Code: XXXX  -  page: URL").
    $code = [regex]::Match($Text, '^Code:\s*(\S+)\s+-\s+page:\s*(\S+)')
    if ($code.Success) {
        $g.Controls.CodeText.Text = $code.Groups[1].Value
        $g.Controls.CodePage.Text = $code.Groups[2].Value
        $g.CodePage = $code.Groups[2].Value
        $g.Controls.CodeBanner.Visibility = 'Visible'
    }
    $g.Controls.LogEmpty.Visibility = 'Collapsed'
    $g.Controls.LogScroll.ScrollToEnd()
    Invoke-WscGuiPump
}

function Clear-WscGuiProgress {
    $g = $script:Gui
    $g.Items.Clear()
    $g.Lines.Clear()
    $g.CodePage = $null
    $g.Controls.CodeBanner.Visibility = 'Collapsed'
    $g.Controls.LogEmpty.Visibility = 'Visible'
}

function Set-WscGuiStatus {
    <# The status pill of the progress: Ready, Running, Passed, Warning, Blocked or Failed. #>
    param([string]$Text, [string]$Status)

    $c = $script:Gui.Controls
    $c.Status.Text = $Text
    $pair = switch ($Status) {
        'Passed' { 'WscSuccess', 'WscSuccessBackground' }
        'Failed' { 'WscCritical', 'WscCriticalBackground' }
        { $_ -in 'Warning', 'Blocked' } { 'WscCaution', 'WscCautionBackground' }
        'Running' { 'WscBrandText', 'WscAccentSoft' }
        default { 'TextFillColorSecondaryBrush', 'WscInfoBackground' }
    }
    $c.Status.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, $pair[0])
    $c.StatusPill.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, $pair[1])
}

function Show-WscTestGui {
    <#
    .SYNOPSIS
        Opens the window. Default configuration: config\WebServicesClient.config.psd1 of the tool folder.
    #>
    [CmdletBinding()]
    param([hashtable]$Configuration)

    if (-not $Configuration) { $Configuration = Import-WscConfiguration }
    $window = New-WscTestForm -Configuration $Configuration
    # Ctrl+C in the console would stop the command that owns the window: the window then could not
    # run any of its PowerShell handlers. Ctrl+C is ignored while the window is open.
    $previousCtrlC = $null
    try { if (-not [Console]::IsInputRedirected) { $previousCtrlC = [Console]::TreatControlCAsInput; [Console]::TreatControlCAsInput = $true } } catch { $previousCtrlC = $null }
    try {
        [void]$window.Form.ShowDialog()
    }
    finally {
        if ($null -ne $previousCtrlC) { try { [Console]::TreatControlCAsInput = $previousCtrlC } catch { } }
        $script:Gui = $null
    }
}

#endregion
