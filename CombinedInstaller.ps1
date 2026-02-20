# ============================================================
#  Combined Installer — WPF GUI + Progress Bar
#  Requires: PowerShell 5.1+, Administrator rights
#
#  .env      — secrets and settings (not in Git)
#  apps.json — application list    (can be in Git)
# ============================================================

# --- Auto-elevate to Administrator ---
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ============================================================
#  Script root (works both as .ps1 and compiled .exe)
# ============================================================
$ScriptRoot = if ($PSScriptRoot -and $PSScriptRoot -ne '') {
    $PSScriptRoot
} else {
    Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}

# ============================================================
#  Load .env
# ============================================================
function Import-EnvFile {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        [System.Windows.MessageBox]::Show(
            "Config file not found:`n$Path`n`nCopy .env.example to .env and fill in the variables.",
            "Configuration Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
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

# ============================================================
#  Load apps.json
# ============================================================
function Import-AppsJson {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        [System.Windows.MessageBox]::Show(
            "apps.json not found:`n$Path`n`nCreate it following the README example.",
            "Configuration Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        exit 1
    }

    try {
        $json = Get-Content $Path -Raw -Encoding UTF8
        $apps = $json | ConvertFrom-Json
        foreach ($app in $apps) {
            if (-not $app.name -or -not $app.file) {
                throw "Each entry must have 'name' and 'file' fields."
            }
        }
        return $apps
    } catch {
        [System.Windows.MessageBox]::Show(
            "Error reading apps.json:`n$_`n`nCheck JSON syntax.",
            "Configuration Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        exit 1
    }
}

# --- Read configs ---
$cfg  = Import-EnvFile  -Path (Join-Path $ScriptRoot ".env")
$Apps = Import-AppsJson -Path (Join-Path $ScriptRoot "apps.json")

# ============================================================
#  Variables from .env
# ============================================================
$NAS_BASE = $cfg["NAS_BASE"]
$NAS_USER = $cfg["NAS_USER"]
$NAS_PASS = $cfg["NAS_PASS"]
$TEMP_DIR = $cfg["TEMP_DIR"]
$LOG_FILE = $cfg["LOG_FILE"]

$APP_TITLE    = if ($cfg["APP_TITLE"])    { $cfg["APP_TITLE"] }    else { "Combined Installer" }
$APP_SUBTITLE = if ($cfg["APP_SUBTITLE"]) { $cfg["APP_SUBTITLE"] } else { "Automatic software installation" }

$DefenderPaths     = if ($cfg["DEFENDER_EXCLUDE_PATHS"])     { $cfg["DEFENDER_EXCLUDE_PATHS"]     -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @() }
$DefenderProcesses = if ($cfg["DEFENDER_EXCLUDE_PROCESSES"]) { $cfg["DEFENDER_EXCLUDE_PROCESSES"] -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @() }

# ============================================================
#  XAML  — green theme
#  Background:  #0D1210  (very dark green-black)
#  Card:        #141F1A  (dark green)
#  Accent:      #22C55E  (green-500)
#  Accent dark: #16A34A  (green-600, hover)
#  Toggle ON:   #15803D  (green-700)
#  Separator:   #1F2E27
#  Text dim:    #4B6B58
# ============================================================
[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Combined Installer"
    Width="700" Height="630"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanMinimize"
    Background="#0D1210"
    FontFamily="Segoe UI">

    <Window.Resources>

        <!-- Install button -->
        <Style x:Key="BtnInstall" TargetType="Button">
            <Setter Property="Background" Value="#22C55E"/>
            <Setter Property="Foreground" Value="#052e16"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Height" Value="48"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="10" Padding="24,0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#16A34A"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#1A2920"/>
                                <Setter Property="Foreground" Value="#3D5245"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Slim scrollbar: 3px track, green thumb -->
        <Style x:Key="SlimScrollBar" TargetType="ScrollBar">
            <Setter Property="Width" Value="3"/>
            <Setter Property="MinWidth" Value="3"/>
            <Setter Property="Background" Value="#141F1A"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid x:Name="Bg" Background="{TemplateBinding Background}" Width="3">
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.DecreaseRepeatButton>
                                    <RepeatButton Opacity="0" Height="0"/>
                                </Track.DecreaseRepeatButton>
                                <Track.IncreaseRepeatButton>
                                    <RepeatButton Opacity="0" Height="0"/>
                                </Track.IncreaseRepeatButton>
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
                            <ScrollContentPresenter Grid.Column="0" Margin="0,0,4,0"/>
                            <ScrollBar x:Name="PART_VerticalScrollBar"
                                       Style="{StaticResource SlimScrollBar}"
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

        <!-- App card -->
        <Style x:Key="AppCard" TargetType="Border">
            <Setter Property="Background" Value="#141F1A"/>
            <Setter Property="CornerRadius" Value="10"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
            <Setter Property="Padding" Value="16,12"/>
        </Style>

        <!-- Progress bar -->
        <Style x:Key="MainProgress" TargetType="ProgressBar">
            <Setter Property="Height" Value="8"/>
            <Setter Property="Background" Value="#1A2920"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ProgressBar">
                        <Grid>
                            <Border Background="#1A2920" CornerRadius="4"/>
                            <Border x:Name="PART_Track" CornerRadius="4"/>
                            <Border x:Name="PART_Indicator" HorizontalAlignment="Left" CornerRadius="4">
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


        <!-- Custom checkbox: green rounded square with checkmark -->
        <Style x:Key="AppCheckBox" TargetType="CheckBox">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <Border x:Name="Box"
                                Width="22" Height="22"
                                CornerRadius="6"
                                BorderThickness="2"
                                BorderBrush="#2D5240"
                                Background="Transparent">
                            <TextBlock x:Name="Check"
                                       Text="&#10003;"
                                       FontSize="13"
                                       FontWeight="Bold"
                                       Foreground="#22C55E"
                                       HorizontalAlignment="Center"
                                       VerticalAlignment="Center"
                                       Visibility="Collapsed"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Box"   Property="Background"    Value="#15803D"/>
                                <Setter TargetName="Box"   Property="BorderBrush"   Value="#22C55E"/>
                                <Setter TargetName="Check" Property="Visibility"    Value="Visible"/>
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

    <Grid Margin="32,28,32,28">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,6">
            <TextBlock x:Name="TitleBlock"    Foreground="White"   FontSize="26" FontWeight="Bold"/>
            <TextBlock x:Name="SubtitleBlock" Foreground="#4B6B58" FontSize="13" Margin="0,4,0,0"/>
        </StackPanel>

        <Border Grid.Row="1" Height="1" Background="#1F2E27" Margin="0,0,0,14"/>

        <!-- Column headers -->
        <Grid Grid.Row="2" Margin="16,0,16,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="130"/>
                <ColumnDefinition Width="28"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Application"  Foreground="#3D5245" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center" Grid.Column="0"/>
            <TextBlock Text="Mode"         Foreground="#3D5245" FontSize="11" FontWeight="SemiBold" HorizontalAlignment="Center" VerticalAlignment="Center" Grid.Column="1"/>
            <TextBlock Text="Install" Foreground="#3D5245" FontSize="11" FontWeight="SemiBold" HorizontalAlignment="Center" VerticalAlignment="Center" Grid.Column="2"/>
        </Grid>

        <!-- App list -->
        <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto" Style="{StaticResource SlimScrollViewer}">
            <StackPanel x:Name="AppList"/>
        </ScrollViewer>

        <!-- Progress bar -->
        <Grid Grid.Row="4" Margin="0,18,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <ProgressBar x:Name="ProgressMain" Style="{StaticResource MainProgress}"
                         Minimum="0" Maximum="100" Value="0" Grid.Column="0"/>
            <TextBlock x:Name="ProgressPercent" Text="0%"
                       Foreground="#22C55E" FontSize="12" FontWeight="SemiBold"
                       VerticalAlignment="Center" Margin="12,0,0,0" Grid.Column="1"/>
        </Grid>

        <!-- Status -->
        <TextBlock x:Name="StatusText" Grid.Row="5"
                   Foreground="#4B6B58" FontSize="12"
                   Margin="0,0,0,4" TextTrimming="CharacterEllipsis"/>

        <!-- Log -->
        <Border Grid.Row="6" Background="#0A100D" CornerRadius="8"
                Height="80" Margin="0,4,0,16" Padding="12,8">
            <ScrollViewer x:Name="LogScroll" VerticalScrollBarVisibility="Auto">
                <TextBlock x:Name="LogText" Foreground="#3D5245"
                           FontSize="11" FontFamily="Consolas" TextWrapping="Wrap"/>
            </ScrollViewer>
        </Border>

        <!-- Buttons -->
        <Grid Grid.Row="7">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="BtnInstall" Content="Install selected"
                    Style="{StaticResource BtnInstall}" Grid.Column="0"/>
            <Button x:Name="BtnClose" Content="Close" Grid.Column="2"
                    Width="110" Height="48" Background="#141F1A" Foreground="#4B6B58"
                    BorderThickness="0" FontSize="14" Cursor="Hand">
                <Button.Template>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="10">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#1A2920"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Button.Template>
            </Button>
        </Grid>
    </Grid>
</Window>
"@

# ============================================================
#  Create window
# ============================================================
$Reader      = [System.Xml.XmlNodeReader]::new($XAML)
$Window      = [Windows.Markup.XamlReader]::Load($Reader)
$AppListCtrl = $Window.FindName("AppList")
$ProgressMain= $Window.FindName("ProgressMain")
$ProgressPct = $Window.FindName("ProgressPercent")
$StatusText  = $Window.FindName("StatusText")
$LogText     = $Window.FindName("LogText")
$LogScroll   = $Window.FindName("LogScroll")
$BtnInstall  = $Window.FindName("BtnInstall")
$BtnClose    = $Window.FindName("BtnClose")

# --- Set title/subtitle programmatically to avoid encoding issues ---
$Window.FindName("TitleBlock").Text    = $APP_TITLE
$Window.FindName("SubtitleBlock").Text = $APP_SUBTITLE

# ============================================================
#  App states
# ============================================================
$AppStates = @{}

# ============================================================
#  Toggle switch builder
#  Silent ON  → green track, thumb right, label "Silent"
#  Silent OFF → grey track,  thumb left,  label "Manual"
# ============================================================
function New-Toggle {
    param([bool]$IsSilent)

    $conv = [System.Windows.Media.BrushConverter]::new()

    $container = [System.Windows.Controls.StackPanel]::new()
    $container.Orientation       = "Horizontal"
    $container.VerticalAlignment = "Center"
    $container.Cursor            = [System.Windows.Input.Cursors]::Hand

    # Label
    $label = [System.Windows.Controls.TextBlock]::new()
    $label.FontSize      = 11
    $label.FontWeight    = "SemiBold"
    $label.Width         = 56
    $label.TextAlignment = "Right"
    $label.VerticalAlignment = "Center"
    $label.Margin        = [System.Windows.Thickness]::new(0,0,8,0)

    # Track
    $track = [System.Windows.Controls.Border]::new()
    $track.Width        = 38
    $track.Height       = 20
    $track.CornerRadius = [System.Windows.CornerRadius]::new(10)

    # Thumb
    $thumb = [System.Windows.Controls.Border]::new()
    $thumb.Width        = 14
    $thumb.Height       = 14
    $thumb.CornerRadius = [System.Windows.CornerRadius]::new(7)
    $thumb.Background   = $conv.ConvertFrom("#FFFFFF")

    $track.Child = $thumb

    $container.Children.Add($label)
    $container.Children.Add($track)

    # Apply initial state
    $c = [System.Windows.Media.BrushConverter]::new()
    if ($IsSilent) {
        $label.Text       = "Silent"
        $label.Foreground = $c.ConvertFrom("#22C55E")
        $track.Background = $c.ConvertFrom("#15803D")
        $thumb.Margin     = [System.Windows.Thickness]::new(20,3,3,3)
    } else {
        $label.Text       = "Manual"
        $label.Foreground = $c.ConvertFrom("#4B6B58")
        $track.Background = $c.ConvertFrom("#1F3329")
        $thumb.Margin     = [System.Windows.Thickness]::new(3,3,20,3)
    }

    return @{ Container = $container; Track = $track; Thumb = $thumb; Label = $label }
}

# ============================================================
#  Build app cards from apps.json
# ============================================================
foreach ($app in $Apps) {
    $icon          = if ($app.icon) { $app.icon } else { "📦" }
    $desc          = if ($app.desc) { $app.desc } else { "" }
    $defaultSilent = if ($app.PSObject.Properties["silent"] -and $app.silent -eq $false) { $false } else { $true }

    $AppStates[$app.name] = @{ Checked = $true; Silent = $defaultSilent }

    # Card
    $card       = [System.Windows.Controls.Border]::new()
    $card.Style = $Window.Resources["AppCard"]

    # Grid: [icon+text] | [toggle] | [checkbox]
    $grid = [System.Windows.Controls.Grid]::new()
    $cL = [System.Windows.Controls.ColumnDefinition]::new(); $cL.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $cT = [System.Windows.Controls.ColumnDefinition]::new(); $cT.Width = [System.Windows.GridLength]::Auto
    $cC = [System.Windows.Controls.ColumnDefinition]::new(); $cC.Width = [System.Windows.GridLength]::Auto
    $grid.ColumnDefinitions.Add($cL)
    $grid.ColumnDefinitions.Add($cT)
    $grid.ColumnDefinitions.Add($cC)

    # Icon + text
    $left = [System.Windows.Controls.StackPanel]::new()
    $left.Orientation       = "Horizontal"
    $left.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($left, 0)

    # Icon inside rounded badge
    $iconBadge                   = [System.Windows.Controls.Border]::new()
    $iconBadge.Width             = 40
    $iconBadge.Height            = 40
    $iconBadge.CornerRadius      = [System.Windows.CornerRadius]::new(10)
    $iconBadge.Background        = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#1E3D2F")
    $iconBadge.VerticalAlignment = "Center"
    $iconBadge.Margin            = [System.Windows.Thickness]::new(0,0,12,0)

    $iconBlock                      = [System.Windows.Controls.TextBlock]::new()
    $iconBlock.Text                 = $icon
    $iconBlock.FontSize             = 20
    $iconBlock.HorizontalAlignment  = "Center"
    $iconBlock.VerticalAlignment    = "Center"
    $iconBadge.Child               = $iconBlock

    $textStack                   = [System.Windows.Controls.StackPanel]::new()
    $textStack.VerticalAlignment = "Center"

    $nameBlock            = [System.Windows.Controls.TextBlock]::new()
    $nameBlock.Text       = $app.name
    $nameBlock.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#D1FAE5")
    $nameBlock.FontSize   = 14
    $nameBlock.FontWeight = "SemiBold"

    $descBlock            = [System.Windows.Controls.TextBlock]::new()
    $descBlock.Text       = $desc
    $descBlock.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#4B6B58")
    $descBlock.FontSize   = 12

    $textStack.Children.Add($nameBlock)
    $textStack.Children.Add($descBlock)
    $left.Children.Add($iconBadge)
    $left.Children.Add($textStack)

    # Toggle
    $toggle          = New-Toggle -IsSilent $defaultSilent
    $toggleContainer = $toggle.Container
    $toggleContainer.Margin = [System.Windows.Thickness]::new(12,0,16,0)
    [System.Windows.Controls.Grid]::SetColumn($toggleContainer, 1)

    $toggleContainer.Tag = @{
        AppName = $app.name
        Track   = $toggle.Track
        Thumb   = $toggle.Thumb
        Label   = $toggle.Label
    }

    $toggleContainer.Add_MouseLeftButtonUp({
        param($s, $e)
        $t        = $s.Tag
        $newState = -not $AppStates[$t.AppName].Silent
        $AppStates[$t.AppName].Silent = $newState

        $conv = [System.Windows.Media.BrushConverter]::new()
        if ($newState) {
            $t.Label.Text       = "Silent"
            $t.Label.Foreground = $conv.ConvertFrom("#22C55E")
            $t.Track.Background = $conv.ConvertFrom("#15803D")
            $t.Thumb.Margin     = [System.Windows.Thickness]::new(20,3,3,3)
        } else {
            $t.Label.Text       = "Manual"
            $t.Label.Foreground = $conv.ConvertFrom("#4B6B58")
            $t.Track.Background = $conv.ConvertFrom("#1F3329")
            $t.Thumb.Margin     = [System.Windows.Thickness]::new(3,3,20,3)
        }
    })

    # Checkbox
    $chk                   = [System.Windows.Controls.CheckBox]::new()
    $chk.IsChecked         = $true
    $chk.Style             = $Window.Resources["AppCheckBox"]
    $chk.Margin            = [System.Windows.Thickness]::new(8,0,0,0)
    $chkTip                = [System.Windows.Controls.ToolTip]::new()
    $chkTip.Content        = "Include in installation"
    $chkTip.Background     = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#141F1A")
    $chkTip.Foreground     = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#86EFAC")
    $chkTip.BorderBrush    = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#1F2E27")
    $chk.ToolTip           = $chkTip
    [System.Windows.Controls.Grid]::SetColumn($chk, 2)

    $appNameRef = $app.name
    $chk.Add_Checked({   $AppStates[$appNameRef].Checked = $true  })
    $chk.Add_Unchecked({ $AppStates[$appNameRef].Checked = $false })
    $AppStates[$app.name].CheckBox = $chk

    $grid.Children.Add($left)
    $grid.Children.Add($toggleContainer)
    $grid.Children.Add($chk)
    $card.Child = $grid
    $AppListCtrl.Children.Add($card)
}

# ============================================================
#  Helpers
# ============================================================
function AppendLog($msg) {
    $ts = (Get-Date).ToString("HH:mm:ss")
    $Window.Dispatcher.Invoke({
        $LogText.Text += "[$ts] $msg`n"
        $LogScroll.ScrollToEnd()
        Add-Content -Path $LOG_FILE -Value "[$ts] $msg" -ErrorAction SilentlyContinue
    })
}
function SetStatus($msg)   { $Window.Dispatcher.Invoke({ $StatusText.Text = $msg }) }
function SetProgress($pct) {
    $Window.Dispatcher.Invoke({
        $ProgressMain.Value = $pct
        $ProgressPct.Text   = "$pct%"
    })
}

# ============================================================
#  Download from NAS
# ============================================================
function Get-RemoteFile {
    param([string]$FileName, [string]$DestPath)

    $src = "$NAS_BASE/$FileName"
    if ($src -match '^\\\\') {
        Copy-Item -Path $src -Destination $DestPath -Force
        return
    }

    $wc = [System.Net.WebClient]::new()
    if ($NAS_USER -and $NAS_PASS) {
        $wc.Credentials = [System.Net.NetworkCredential]::new($NAS_USER, $NAS_PASS)
    }
    $wc.DownloadFile($src, $DestPath)
}

# ============================================================
#  Install button
# ============================================================
$BtnInstall.Add_Click({
    $BtnInstall.IsEnabled = $false
    $BtnClose.IsEnabled   = $false

    $selected = $Apps | Where-Object { $AppStates[$_.name].Checked -eq $true } | ForEach-Object {
        [PSCustomObject]@{
            name   = $_.name
            file   = $_.file
            args   = $_.args
            silent = $AppStates[$_.name].Silent
        }
    }

    if ($selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please select at least one application.", "Nothing selected", "OK", "Warning")
        $BtnInstall.IsEnabled = $true
        $BtnClose.IsEnabled   = $true
        return
    }

    $total = $selected.Count
    $step  = [math]::Floor(100 / ($total * 2))

    # --- Передаём все переменные явно в Runspace ---
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = "STA"
    $rs.ThreadOptions  = "ReuseThread"
    $rs.Open()

    # Передаём переменные окружения в runspace
    $rs.SessionStateProxy.SetVariable("Window",            $Window)
    $rs.SessionStateProxy.SetVariable("LogText",           $LogText)
    $rs.SessionStateProxy.SetVariable("LogScroll",         $LogScroll)
    $rs.SessionStateProxy.SetVariable("StatusText",        $StatusText)
    $rs.SessionStateProxy.SetVariable("ProgressMain",      $ProgressMain)
    $rs.SessionStateProxy.SetVariable("ProgressPct",       $ProgressPct)
    $rs.SessionStateProxy.SetVariable("BtnInstall",        $BtnInstall)
    $rs.SessionStateProxy.SetVariable("BtnClose",          $BtnClose)
    $rs.SessionStateProxy.SetVariable("selected",          $selected)
    $rs.SessionStateProxy.SetVariable("total",             $total)
    $rs.SessionStateProxy.SetVariable("step",              $step)
    $rs.SessionStateProxy.SetVariable("NAS_BASE",          $NAS_BASE)
    $rs.SessionStateProxy.SetVariable("NAS_USER",          $NAS_USER)
    $rs.SessionStateProxy.SetVariable("NAS_PASS",          $NAS_PASS)
    $rs.SessionStateProxy.SetVariable("TEMP_DIR",          $TEMP_DIR)
    $rs.SessionStateProxy.SetVariable("LOG_FILE",          $LOG_FILE)
    $rs.SessionStateProxy.SetVariable("DefenderPaths",     $DefenderPaths)
    $rs.SessionStateProxy.SetVariable("DefenderProcesses", $DefenderProcesses)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs

    $ps.AddScript({

        # --- Helpers внутри runspace ---
        function AppendLog($msg) {
            $ts = (Get-Date).ToString("HH:mm:ss")
            $Window.Dispatcher.Invoke([Action]{
                $LogText.Text += "[$ts] $msg`n"
                $LogScroll.ScrollToEnd()
                Add-Content -Path $LOG_FILE -Value "[$ts] $msg" -ErrorAction SilentlyContinue
            })
        }
        function SetStatus($msg) {
            $Window.Dispatcher.Invoke([Action]{ $StatusText.Text = $msg })
        }
        function SetProgress($pct) {
            $Window.Dispatcher.Invoke([Action]{
                $ProgressMain.Value = $pct
                $ProgressPct.Text   = "$pct%"
            })
        }
        function Get-RemoteFile {
            param([string]$FileName, [string]$DestPath)
            $src = "$NAS_BASE/$FileName"
            if ($src -match '^\\\\') {
                Copy-Item -Path $src -Destination $DestPath -Force
                return
            }
            $wc = [System.Net.WebClient]::new()
            if ($NAS_USER -and $NAS_PASS) {
                $wc.Credentials = [System.Net.NetworkCredential]::new($NAS_USER, $NAS_PASS)
            }
            $wc.DownloadFile($src, $DestPath)
        }

        # --- Создаём папки ---
        New-Item -ItemType Directory -Path $TEMP_DIR -Force | Out-Null
        New-Item -ItemType Directory -Path (Split-Path $LOG_FILE) -Force | Out-Null
        "" | Set-Content -Path $LOG_FILE -Encoding UTF8

        AppendLog "=== Installation started ==="
        AppendLog "NAS: $NAS_BASE"
        AppendLog "Selected: $total apps"

        # Defender exclusions
        SetStatus "Configuring Windows Defender..."
        AppendLog "Adding Defender exclusions..."
        foreach ($p in (@($TEMP_DIR) + $DefenderPaths)) {
            if ($p) { Add-MpPreference -ExclusionPath $p -ErrorAction SilentlyContinue }
        }
        foreach ($proc in $DefenderProcesses) {
            if ($proc) { Add-MpPreference -ExclusionProcess $proc -ErrorAction SilentlyContinue }
        }
        AppendLog "OK Exclusions added"

        $current = 0
        foreach ($app in $selected) {
            $current++
            $destFile  = "$TEMP_DIR\$($app.file)"
            $modeLabel = if ($app.silent) { "silent" } else { "manual" }

            # Download
            SetStatus "[$current/$total] Downloading: $($app.name)..."
            AppendLog ">> Downloading $($app.name)..."
            try {
                Get-RemoteFile -FileName $app.file -DestPath $destFile
                AppendLog "OK Downloaded: $($app.file)"
            } catch {
                AppendLog "ERR Download failed $($app.name): $_"
                SetProgress ([int](($current * 2 - 1) * $step))
                continue
            }

            SetProgress ([int](($current * 2 - 1) * $step))

            # Install
            SetStatus "[$current/$total] Installing ($modeLabel): $($app.name)..."
            AppendLog ".. Installing $($app.name) [$modeLabel]..."
            try {
                if ($app.silent) {
                    $proc = Start-Process -FilePath $destFile -ArgumentList $app.args -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop
                } else {
                    $proc = Start-Process -FilePath $destFile -ArgumentList $app.args -Wait -PassThru -ErrorAction Stop
                }
                if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
                    AppendLog "OK Installed: $($app.name) (ExitCode: $($proc.ExitCode))"
                } else {
                    AppendLog "!! $($app.name) exited with code $($proc.ExitCode)"
                }
            } catch {
                AppendLog "ERR Install failed $($app.name): $_"
            }

            SetProgress ([int]($current * 2 * $step))
        }

        # Done
        SetProgress 100
        SetStatus "Installation complete"
        AppendLog "=== Done. Log: $LOG_FILE ==="
        AppendLog "Cleaning up temp files (keeping *_log.txt)..."
        Get-ChildItem -Path $TEMP_DIR -Recurse -Force -File |
            Where-Object { $_.Extension -ne "_log.txt" } |
            Remove-Item -Force -ErrorAction SilentlyContinue
        AppendLog "OK All done"

        $Window.Dispatcher.Invoke([Action]{
            $BtnInstall.IsEnabled = $false
            $BtnInstall.Content   = "Installation complete"
            $BtnClose.IsEnabled   = $true
        })

    }) | Out-Null

    # Запускаем асинхронно
    $ps.BeginInvoke() | Out-Null
})

$BtnClose.Add_Click({ $Window.Close() })

$Window.ShowDialog() | Out-Null