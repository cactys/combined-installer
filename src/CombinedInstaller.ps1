# ============================================================
#  Combined Installer — WPF GUI
#  Requires: PowerShell 5.1+, Administrator rights
#
#  .env          — secrets and settings (not in Git)
#  src/apps.json — application list    (can be in Git)
# ============================================================

# --- Auto-elevate to Administrator ---
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Script root
$ScriptRoot = if ($PSScriptRoot -and $PSScriptRoot -ne '') {
    $PSScriptRoot
} else {
    Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}

# Load .env
function Import-EnvFile {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        [System.Windows.MessageBox]::Show("Config file not found:`n$Path`n`nCopy .env.example to .env and fill in the variables.", "Configuration Error", "OK", "Error")
        exit 1
    }
    $cfg = @{}
    foreach ($line in (Get-Content $Path -Encoding UTF8)) {
        $line = $line.Trim()
        if ($line -eq '' -or $line.StartsWith('#')) { continue }
        if ($line -match '^([^=]+)=(.*)$') {
            $cfg[$Matches[1].Trim()] = [System.Environment]::ExpandEnvironmentVariables($Matches[2].Trim())
        }
    }
    return $cfg
}

# Load apps.json
function Import-AppsJson {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        [System.Windows.MessageBox]::Show("apps.json not found:`n$Path", "Configuration Error", "OK", "Error")
        exit 1
    }
    try {
        $apps = (Get-Content $Path -Raw -Encoding UTF8) | ConvertFrom-Json
        foreach ($app in $apps) {
            if (-not $app.name -or -not $app.path) { throw "Each entry must have 'name' and 'path'." }
        }
        return $apps
    } catch {
        [System.Windows.MessageBox]::Show("Error reading apps.json:`n$_", "Configuration Error", "OK", "Error")
        exit 1
    }
}

$cfg  = Import-EnvFile  -Path (Join-Path $ScriptRoot ".env")
$Apps = Import-AppsJson -Path (Join-Path $ScriptRoot "src\apps.json")

$NAS_BASE = $cfg["NAS_BASE"]
$NAS_USER = $cfg["NAS_USER"]
$NAS_PASS = $cfg["NAS_PASS"]
$TEMP_DIR = $cfg["TEMP_DIR"]
$LOG_FILE = $cfg["LOG_FILE"]

$APP_TITLE    = if ($cfg["APP_TITLE"])    { $cfg["APP_TITLE"] }    else { "Combined Installer" }
$APP_SUBTITLE = if ($cfg["APP_SUBTITLE"]) { $cfg["APP_SUBTITLE"] } else { "Automatic software installation" }

$DefenderPaths     = if ($cfg["DEFENDER_EXCLUDE_PATHS"])     { $cfg["DEFENDER_EXCLUDE_PATHS"]     -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @() }
$DefenderProcesses = if ($cfg["DEFENDER_EXCLUDE_PROCESSES"]) { $cfg["DEFENDER_EXCLUDE_PROCESSES"] -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @() }

# XAML
[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Combined Installer"
    Width="900" Height="620"
    MinWidth="800" MinHeight="550"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize"
    Background="#0D1210"
    FontFamily="Segoe UI">
    <Window.Resources>
        <Style x:Key="BtnInstall" TargetType="Button">
            <Setter Property="Background" Value="#22C55E"/>
            <Setter Property="Foreground" Value="#052e16"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Height" Value="44"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="10" Padding="20,0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#16A34A"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Background" Value="#1A2920"/><Setter Property="Foreground" Value="#3D5245"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="BtnClose" TargetType="Button">
            <Setter Property="Background" Value="#141F1A"/>
            <Setter Property="Foreground" Value="#4B6B58"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Height" Value="44"/>
            <Setter Property="Width" Value="100"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="10">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#1A2920"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="AppCard" TargetType="Border">
            <Setter Property="Background" Value="#141F1A"/>
            <Setter Property="CornerRadius" Value="10"/>
            <Setter Property="Margin" Value="0,0,0,6"/>
            <Setter Property="Padding" Value="12,10"/>
        </Style>
        <Style x:Key="SlimScrollBar" TargetType="ScrollBar">
            <Setter Property="Width" Value="4"/>
            <Setter Property="MinWidth" Value="4"/>
            <Setter Property="Background" Value="#0D1210"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid Background="{TemplateBinding Background}" Width="4">
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.DecreaseRepeatButton><RepeatButton Opacity="0" Height="0"/></Track.DecreaseRepeatButton>
                                <Track.IncreaseRepeatButton><RepeatButton Opacity="0" Height="0"/></Track.IncreaseRepeatButton>
                                <Track.Thumb>
                                    <Thumb>
                                        <Thumb.Template>
                                            <ControlTemplate TargetType="Thumb">
                                                <Border CornerRadius="2" Background="#22C55E"/>
                                            </ControlTemplate>
                                        </Thumb.Template>
                                    </Thumb>
                                </Track.Thumb>
                            </Track>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SlimScrollViewer" TargetType="ScrollViewer">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollViewer">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <ScrollContentPresenter Grid.Column="0"/>
                            <ScrollBar x:Name="PART_VerticalScrollBar" Style="{StaticResource SlimScrollBar}"
                                       Grid.Column="1"
                                       Value="{TemplateBinding VerticalOffset}"
                                       Maximum="{TemplateBinding ScrollableHeight}"
                                       ViewportSize="{TemplateBinding ViewportHeight}"
                                       Visibility="{TemplateBinding ComputedVerticalScrollBarVisibility}"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="MainProgress" TargetType="ProgressBar">
            <Setter Property="Height" Value="6"/>
            <Setter Property="Background" Value="#1A2920"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ProgressBar">
                        <Grid>
                            <Border Background="#1A2920" CornerRadius="3"/>
                            <Border x:Name="PART_Track" CornerRadius="3"/>
                            <Border x:Name="PART_Indicator" HorizontalAlignment="Left" CornerRadius="3">
                                <Border.Background>
                                    <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                                        <GradientStop Color="#22C55E" Offset="0"/>
                                        <GradientStop Color="#86EFAC" Offset="1"/>
                                    </LinearGradientBrush>
                                </Border.Background>
                            </Border>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="AppCheckBox" TargetType="CheckBox">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <Border x:Name="Box" Width="22" Height="22" CornerRadius="6"
                                BorderThickness="2" BorderBrush="#2D5240" Background="Transparent">
                            <TextBlock x:Name="Check" Text="&#10003;" FontSize="13" FontWeight="Bold"
                                       Foreground="#22C55E" HorizontalAlignment="Center"
                                       VerticalAlignment="Center" Visibility="Collapsed"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Box"   Property="Background"  Value="#15803D"/>
                                <Setter TargetName="Box"   Property="BorderBrush" Value="#22C55E"/>
                                <Setter TargetName="Check" Property="Visibility"  Value="Visible"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Box" Property="BorderBrush" Value="#22C55E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" MinWidth="320"/>
            <ColumnDefinition Width="6"/>
            <ColumnDefinition Width="*" MinWidth="220"/>
        </Grid.ColumnDefinitions>

        <!-- LEFT PANEL -->
        <Grid Grid.Column="0" Margin="24,24,16,24">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <StackPanel Grid.Row="0" Margin="0,0,0,6">
                <TextBlock x:Name="TitleBlock"    Foreground="White"   FontSize="22" FontWeight="Bold"/>
                <TextBlock x:Name="SubtitleBlock" Foreground="#4B6B58" FontSize="12" Margin="0,3,0,0"/>
            </StackPanel>
            <Border Grid.Row="1" Height="1" Background="#1F2E27" Margin="0,0,0,12"/>
            <Grid Grid.Row="2" Margin="12,0,8,6">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="110"/>
                    <ColumnDefinition Width="30"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="Application" Foreground="#3D5245" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center" Grid.Column="0"/>
                <TextBlock Text="Mode"        Foreground="#3D5245" FontSize="11" FontWeight="SemiBold" HorizontalAlignment="Center" VerticalAlignment="Center" Grid.Column="1"/>
                <TextBlock Text="Install"     Foreground="#3D5245" FontSize="11" FontWeight="SemiBold" HorizontalAlignment="Center" VerticalAlignment="Center" Grid.Column="2"/>
            </Grid>
            <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto" Style="{StaticResource SlimScrollViewer}">
                <StackPanel x:Name="AppList"/>
            </ScrollViewer>
            <Grid Grid.Row="4" Margin="0,14,0,6">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <ProgressBar x:Name="ProgressMain" Style="{StaticResource MainProgress}" Minimum="0" Maximum="100" Value="0"/>
                <TextBlock x:Name="ProgressPercent" Text="0%" Foreground="#22C55E" FontSize="11"
                           FontWeight="SemiBold" VerticalAlignment="Center" Margin="8,0,0,0" Grid.Column="1"/>
            </Grid>
            <Grid Grid.Row="5">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock x:Name="StatusText" Grid.Row="0" Text="Ready" Foreground="#4B6B58"
                           FontSize="11" Margin="0,0,0,8" TextTrimming="CharacterEllipsis"/>
                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="BtnInstall" Content="Install selected" Style="{StaticResource BtnInstall}" Grid.Column="0"/>
                    <Button x:Name="BtnClose"   Content="Close"           Style="{StaticResource BtnClose}"   Grid.Column="2"/>
                </Grid>
            </Grid>
        </Grid>

        <!-- DRAGGABLE SPLITTER -->
        <GridSplitter Grid.Column="1"
                      Width="6"
                      HorizontalAlignment="Stretch"
                      VerticalAlignment="Stretch"
                      Background="#1F2E27"
                      Cursor="SizeWE"
                      ResizeBehavior="PreviousAndNext"
                      ResizeDirection="Columns"/>

        <!-- RIGHT PANEL - LOG -->
        <Grid Grid.Column="2" Background="#0A100D">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid Grid.Row="0" Background="#0D1210" Margin="0,0,0,1">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="Installation Log" Foreground="#4B6B58" FontSize="11" FontWeight="SemiBold"
                           VerticalAlignment="Center" Margin="14,10,0,10"/>
                <Button x:Name="BtnClearLog" Content="Clear" Grid.Column="1"
                        Background="Transparent" Foreground="#3D5245" BorderThickness="0"
                        FontSize="10" Cursor="Hand" Margin="0,0,10,0" Padding="6,2"/>
            </Grid>
            <Border Grid.Row="0" Height="1" VerticalAlignment="Bottom" Background="#1F2E27"/>
            <ScrollViewer Grid.Row="1" x:Name="LogScroll"
                          VerticalScrollBarVisibility="Auto"
                          Style="{StaticResource SlimScrollViewer}"
                          Padding="14,10,6,10">
                <TextBlock x:Name="LogText" Foreground="#4B6B58"
                           FontSize="11" FontFamily="Consolas"
                           TextWrapping="Wrap" LineHeight="18"/>
            </ScrollViewer>
            <Border Grid.Row="2" Background="#0D1210" Padding="14,6" BorderThickness="0,1,0,0" BorderBrush="#1F2E27">
                <TextBlock x:Name="LogPathText" Foreground="#2D4035" FontSize="9"
                           FontFamily="Consolas" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/>
            </Border>
        </Grid>
    </Grid>
</Window>
"@

# Create window
$Reader       = [System.Xml.XmlNodeReader]::new($XAML)
$Window       = [Windows.Markup.XamlReader]::Load($Reader)
$AppListCtrl  = $Window.FindName("AppList")
$ProgressMain = $Window.FindName("ProgressMain")
$ProgressPct  = $Window.FindName("ProgressPercent")
$StatusText   = $Window.FindName("StatusText")
$LogText      = $Window.FindName("LogText")
$LogScroll    = $Window.FindName("LogScroll")
$LogPathText  = $Window.FindName("LogPathText")
$BtnInstall   = $Window.FindName("BtnInstall")
$BtnClose     = $Window.FindName("BtnClose")
$BtnClearLog  = $Window.FindName("BtnClearLog")

$Window.FindName("TitleBlock").Text    = $APP_TITLE
$Window.FindName("SubtitleBlock").Text = $APP_SUBTITLE
$LogPathText.Text                      = "Log: $LOG_FILE"

$BtnClearLog.Add_Click({ $LogText.Text = "" })

# App states
$AppStates = @{}

# Toggle builder
function New-Toggle {
    param([bool]$IsSilent)
    $c = [System.Windows.Media.BrushConverter]::new()
    $container = [System.Windows.Controls.StackPanel]::new()
    $container.Orientation = "Horizontal"; $container.VerticalAlignment = "Center"
    $container.Cursor = [System.Windows.Input.Cursors]::Hand
    $label = [System.Windows.Controls.TextBlock]::new()
    $label.FontSize = 10; $label.FontWeight = "SemiBold"; $label.Width = 50
    $label.TextAlignment = "Right"; $label.VerticalAlignment = "Center"
    $label.Margin = [System.Windows.Thickness]::new(0,0,7,0)
    $track = [System.Windows.Controls.Border]::new()
    $track.Width = 34; $track.Height = 18
    $track.CornerRadius = [System.Windows.CornerRadius]::new(9)
    $thumb = [System.Windows.Controls.Border]::new()
    $thumb.Width = 12; $thumb.Height = 12
    $thumb.CornerRadius = [System.Windows.CornerRadius]::new(6)
    $thumb.Background = $c.ConvertFrom("#FFFFFF")
    $track.Child = $thumb
    $container.Children.Add($label)
    $container.Children.Add($track)
    if ($IsSilent) {
        $label.Text = "Silent"; $label.Foreground = $c.ConvertFrom("#22C55E")
        $track.Background = $c.ConvertFrom("#15803D")
        $thumb.Margin = [System.Windows.Thickness]::new(18,3,3,3)
    } else {
        $label.Text = "Manual"; $label.Foreground = $c.ConvertFrom("#4B6B58")
        $track.Background = $c.ConvertFrom("#1F3329")
        $thumb.Margin = [System.Windows.Thickness]::new(3,3,18,3)
    }
    return @{ Container = $container; Track = $track; Thumb = $thumb; Label = $label }
}

# Build app cards
foreach ($app in $Apps) {
    $icon          = if ($app.icon) { $app.icon } else { "package" }
    $desc          = if ($app.desc) { $app.desc } else { "" }
    $defaultSilent = if ($app.PSObject.Properties["silent"] -and $app.silent -eq $false) { $false } else { $true }
    $AppStates[$app.name] = @{ Checked = $true; Silent = $defaultSilent }

    $card = [System.Windows.Controls.Border]::new(); $card.Style = $Window.Resources["AppCard"]
    $grid = [System.Windows.Controls.Grid]::new()
    $cL = [System.Windows.Controls.ColumnDefinition]::new(); $cL.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $cT = [System.Windows.Controls.ColumnDefinition]::new(); $cT.Width = [System.Windows.GridLength]::Auto
    $cC = [System.Windows.Controls.ColumnDefinition]::new(); $cC.Width = [System.Windows.GridLength]::Auto
    $grid.ColumnDefinitions.Add($cL); $grid.ColumnDefinitions.Add($cT); $grid.ColumnDefinitions.Add($cC)

    $left = [System.Windows.Controls.StackPanel]::new()
    $left.Orientation = "Horizontal"; $left.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($left, 0)

    $iconBadge = [System.Windows.Controls.Border]::new()
    $iconBadge.Width = 36; $iconBadge.Height = 36
    $iconBadge.CornerRadius = [System.Windows.CornerRadius]::new(9)
    $iconBadge.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#1E3D2F")
    $iconBadge.VerticalAlignment = "Center"
    $iconBadge.Margin = [System.Windows.Thickness]::new(0,0,10,0)
    $iconBlock = [System.Windows.Controls.TextBlock]::new()
    $iconBlock.Text = $icon; $iconBlock.FontSize = 17
    $iconBlock.HorizontalAlignment = "Center"; $iconBlock.VerticalAlignment = "Center"
    $iconBadge.Child = $iconBlock

    $textStack = [System.Windows.Controls.StackPanel]::new(); $textStack.VerticalAlignment = "Center"
    $nameBlock = [System.Windows.Controls.TextBlock]::new()
    $nameBlock.Text = $app.name
    $nameBlock.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#D1FAE5")
    $nameBlock.FontSize = 13; $nameBlock.FontWeight = "SemiBold"
    $descBlock = [System.Windows.Controls.TextBlock]::new()
    $descBlock.Text = $desc
    $descBlock.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#4B6B58")
    $descBlock.FontSize = 11
    $textStack.Children.Add($nameBlock); $textStack.Children.Add($descBlock)
    $left.Children.Add($iconBadge); $left.Children.Add($textStack)

    $toggle = New-Toggle -IsSilent $defaultSilent
    $tc = $toggle.Container; $tc.Margin = [System.Windows.Thickness]::new(10,0,12,0)
    [System.Windows.Controls.Grid]::SetColumn($tc, 1)
    $tc.Tag = @{ AppName = $app.name; Track = $toggle.Track; Thumb = $toggle.Thumb; Label = $toggle.Label }
    $tc.Add_MouseLeftButtonUp({
        param($s,$e); $t = $s.Tag
        $ns = -not $AppStates[$t.AppName].Silent; $AppStates[$t.AppName].Silent = $ns
        $conv = [System.Windows.Media.BrushConverter]::new()
        if ($ns) {
            $t.Label.Text = "Silent"; $t.Label.Foreground = $conv.ConvertFrom("#22C55E")
            $t.Track.Background = $conv.ConvertFrom("#15803D"); $t.Thumb.Margin = [System.Windows.Thickness]::new(18,3,3,3)
        } else {
            $t.Label.Text = "Manual"; $t.Label.Foreground = $conv.ConvertFrom("#4B6B58")
            $t.Track.Background = $conv.ConvertFrom("#1F3329"); $t.Thumb.Margin = [System.Windows.Thickness]::new(3,3,18,3)
        }
    })

    $chk = [System.Windows.Controls.CheckBox]::new()
    $chk.IsChecked = $true; $chk.Style = $Window.Resources["AppCheckBox"]
    $chk.Margin = [System.Windows.Thickness]::new(6,0,0,0)
    [System.Windows.Controls.Grid]::SetColumn($chk, 2)
    $chkTip = [System.Windows.Controls.ToolTip]::new()
    $chkTip.Content = "Include in installation"
    $chkTip.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#141F1A")
    $chkTip.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#86EFAC")
    $chkTip.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#1F2E27")
    $chk.ToolTip = $chkTip
    $appNameRef = $app.name
    $chk.Add_Checked({   $AppStates[$appNameRef].Checked = $true  })
    $chk.Add_Unchecked({ $AppStates[$appNameRef].Checked = $false })
    $AppStates[$app.name].CheckBox = $chk

    $grid.Children.Add($left); $grid.Children.Add($tc); $grid.Children.Add($chk)
    $card.Child = $grid; $AppListCtrl.Children.Add($card)
}

# Install button
$BtnInstall.Add_Click({
    $BtnInstall.IsEnabled = $false
    $BtnClose.IsEnabled   = $false

    # Collect only CHECKED apps with current toggle state
    $selected = @($Apps | Where-Object { $AppStates[$_.name].Checked -eq $true } | ForEach-Object {
        [PSCustomObject]@{
            name   = $_.name
            path   = $_.path
            args   = if ($_.args) { [string]$_.args } else { "" }
            silent = [bool]$AppStates[$_.name].Silent
        }
    })

    if ($selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please select at least one application.", "Nothing selected", "OK", "Warning")
        $BtnInstall.IsEnabled = $true; $BtnClose.IsEnabled = $true
        return
    }

    $total = $selected.Count
    $step  = [math]::Floor(100 / ($total * 2))

    # Ensure log directory exists before runspace starts
    $logDir = Split-Path $LOG_FILE -Parent
    if ($logDir -and -not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir  -Force | Out-Null }
    if (-not (Test-Path $TEMP_DIR))            { New-Item -ItemType Directory -Path $TEMP_DIR -Force | Out-Null }
    Add-Content -Path $LOG_FILE -Value "=== Combined Installer $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" -Encoding UTF8

    # Snapshot all needed values into local vars before passing to runspace
    $rsWindow      = $Window;       $rsLogText    = $LogText;     $rsLogScroll  = $LogScroll
    $rsStatus      = $StatusText;   $rsProgress   = $ProgressMain; $rsProgPct   = $ProgressPct
    $rsBtnI        = $BtnInstall;   $rsBtnC       = $BtnClose
    $rsLogFile     = $LOG_FILE;     $rsTempDir    = $TEMP_DIR
    $rsNasBase     = $NAS_BASE;     $rsNasUser    = $NAS_USER;    $rsNasPass    = $NAS_PASS
    $rsDefPaths    = $DefenderPaths; $rsDefProcs  = $DefenderProcesses
    $rsSelected    = $selected;     $rsTotal      = $total;       $rsStep       = $step

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = "STA"; $rs.ThreadOptions = "ReuseThread"; $rs.Open()

    "rsWindow","rsLogText","rsLogScroll","rsStatus","rsProgress","rsProgPct",
    "rsBtnI","rsBtnC","rsLogFile","rsTempDir","rsNasBase","rsNasUser","rsNasPass",
    "rsDefPaths","rsDefProcs","rsSelected","rsTotal","rsStep" | ForEach-Object {
        $rs.SessionStateProxy.SetVariable($_, (Get-Variable $_ -ValueOnly))
    }

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs

    $ps.AddScript({

        function WriteLog([string]$msg) {
            $ts = (Get-Date).ToString("HH:mm:ss")
            $line = "[$ts] $msg"
            # File write — no UI, always works
            try { Add-Content -Path $rsLogFile -Value $line -Encoding UTF8 } catch {}
            # UI update
            $rsWindow.Dispatcher.Invoke([Action]{
                $rsLogText.Text += "$line`n"
                $rsLogScroll.ScrollToEnd()
            })
        }

        function SetStatus([string]$msg) {
            $rsWindow.Dispatcher.Invoke([Action]{ $rsStatus.Text = $msg })
        }

        function SetProgress([int]$pct) {
            $rsWindow.Dispatcher.Invoke([Action]{
                $rsProgress.Value = $pct
                $rsProgPct.Text   = "$pct%"
            })
        }

        # ── SSL: ignore certificate errors (self-signed Synology cert) ──
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        [System.Net.ServicePointManager]::SecurityProtocol = (
            [System.Net.SecurityProtocolType]::Tls12 -bor
            [System.Net.SecurityProtocolType]::Tls11 -bor
            [System.Net.SecurityProtocolType]::Tls
        )

        # ── Helper: create WebClient with SSL bypass + auth ───────────────
        function New-NasClient {
            $wc = New-Object System.Net.WebClient
            $wc.Headers.Add("User-Agent", "CombinedInstaller/1.0")
            if ($rsNasUser -and $rsNasPass) {
                $wc.Credentials = New-Object System.Net.NetworkCredential($rsNasUser, $rsNasPass)
            }
            return $wc
        }

        # ── Synology File Station API: get session SID ────────────────────
        function Get-SynologySid {
            # Auth endpoint is always at NAS root, not inside the shared folder
            $nasRoot  = $rsNasBase -replace '/webdav.*','' -replace '(/[^/]+){0}$',''
            # Extract just host:port — strip any path after port
            if ($rsNasBase -match '^(https?://[^/]+)') { $nasRoot = $Matches[1] }
            $loginUrl = $nasRoot.TrimEnd('/') + '/webapi/auth.cgi' +
                        '?api=SYNO.API.Auth&version=3&method=login' +
                        '&account=' + [Uri]::EscapeDataString($rsNasUser) +
                        '&passwd='  + [Uri]::EscapeDataString($rsNasPass) +
                        '&session=FileStation&format=sid'
            WriteLog "   Auth URL: $nasRoot/webapi/auth.cgi"
            $resp = Invoke-RestMethod -Uri $loginUrl -Method Get -ErrorAction Stop
            if (-not $resp.success) { throw "Synology auth failed: code $($resp.error.code)" }
            return @{ Sid = $resp.data.sid; Root = $nasRoot }
        }

        # ── Main download dispatcher ──────────────────────────────────────
        function DownloadFile([string]$nasPath, [string]$destPath) {

            # ── 1. SMB / UNC ──────────────────────────────────────────────
            if ($rsNasBase -match '^\\\\') {
                $smb = Join-Path $rsNasBase $nasPath
                WriteLog "   Method: SMB  →  $smb"
                Copy-Item -Path $smb -Destination $destPath -Force
                return
            }

            $base = $rsNasBase.TrimEnd('/')

            # ── 2. WebDAV ─────────────────────────────────────────────────
            # NAS_BASE already contains the webdav root path (e.g. https://host:5006/webdav)
            # nasPath is relative inside that share (e.g. /installers/7zip/file.exe)
            $webdavUrl = $base + '/' + $nasPath.TrimStart('/')
            WriteLog "   Method: WebDAV"
            WriteLog "   URL:    $webdavUrl"
            try {
                $wc = New-NasClient
                $wc.DownloadFile($webdavUrl, $destPath)
                return
            } catch {
                WriteLog "   WebDAV failed: $($_.Exception.Message) — trying File Station API..."
            }

            # ── 3. File Station API fallback ──────────────────────────────
            WriteLog "   Method: File Station API"
            try {
                $auth        = Get-SynologySid
                $encodedPath = [Uri]::EscapeDataString($nasPath)
                $apiUrl      = $auth.Root.TrimEnd('/') + '/webapi/entry.cgi' +
                               '?api=SYNO.FileStation.Download&version=2&method=download' +
                               '&path=' + $encodedPath + '&mode=download&_sid=' + $auth.Sid
                WriteLog "   API:    $($auth.Root)/webapi/entry.cgi"
                $wc = New-NasClient
                $wc.DownloadFile($apiUrl, $destPath)
            } catch {
                throw "File Station API failed: $($_.Exception.Message)"
            }
        }

        WriteLog "Selected: $rsTotal app(s)"

        # Defender
        SetStatus "Configuring Windows Defender..."
        WriteLog "Adding Defender exclusions..."
        foreach ($p in (@($rsTempDir) + $rsDefPaths)) {
            if ($p) { Add-MpPreference -ExclusionPath $p -ErrorAction SilentlyContinue }
        }
        foreach ($proc in $rsDefProcs) {
            if ($proc) { Add-MpPreference -ExclusionProcess $proc -ErrorAction SilentlyContinue }
        }
        WriteLog "OK Defender exclusions set"

        $current = 0
        foreach ($app in $rsSelected) {
            $current++
            # app.path = NAS path e.g. /software/7zip.exe
            $nasPath   = $app.path
            $localName = Split-Path $nasPath -Leaf
            $destFile  = Join-Path $rsTempDir $localName
            $modeLabel = if ($app.silent) { "silent" } else { "manual" }

            WriteLog ""
            WriteLog "[$current/$rsTotal] >>> $($app.name)"

            # Download
            SetStatus "[$current/$rsTotal] Downloading: $($app.name)..."
            WriteLog "Downloading [$(Split-Path $nasPath -Leaf)]..."
            try {
                DownloadFile -nasPath $nasPath -destPath $destFile
                $sizeMB = [math]::Round((Get-Item $destFile).Length / 1MB, 2)
                WriteLog "OK Downloaded — ${sizeMB} MB"
            } catch {
                WriteLog "ERR Download failed: $($_.Exception.Message)"
                SetProgress ([int](($current * 2 - 1) * $rsStep))
                continue
            }

            SetProgress ([int](($current * 2 - 1) * $rsStep))

            # Install
            SetStatus "[$current/$rsTotal] Installing ($modeLabel): $($app.name)..."
            WriteLog "Installing [$modeLabel]..."
            try {
                $procArgs = @{ FilePath = $destFile; Wait = $true; PassThru = $true; ErrorAction = "Stop" }
                if ($app.args -and $app.args.Trim() -ne "") { $procArgs.ArgumentList = $app.args }
                if ($app.silent) { $procArgs.WindowStyle = "Hidden" }

                $proc = Start-Process @procArgs

                if ($proc.ExitCode -eq 0) {
                    WriteLog "OK Installed successfully"
                } elseif ($proc.ExitCode -eq 3010) {
                    WriteLog "OK Installed — reboot required"
                } else {
                    WriteLog "!! Finished with exit code $($proc.ExitCode)"
                }
            } catch {
                WriteLog "ERR Install error: $($_.Exception.Message)"
            }

            # Remove installer file, keep log
            try { Remove-Item $destFile -Force -ErrorAction SilentlyContinue } catch {}

            SetProgress ([int]($current * 2 * $rsStep))
        }

        SetProgress 100
        SetStatus "Installation complete"
        WriteLog ""
        WriteLog "=== Done. Log: $rsLogFile ==="

        $rsWindow.Dispatcher.Invoke([Action]{
            $rsBtnI.IsEnabled = $false
            $rsBtnI.Content   = "Installation complete"
            $rsBtnC.IsEnabled = $true
        })

    }) | Out-Null

    $ps.BeginInvoke() | Out-Null
})

$BtnClose.Add_Click({ $Window.Close() })

$Window.ShowDialog() | Out-Null