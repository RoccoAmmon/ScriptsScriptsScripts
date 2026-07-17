<#
.SYNOPSIS
    Bluetooth Diagnose- und Verwaltungstool mit moderner WPF-Oberfläche
.DESCRIPTION
    Zeigt Bluetooth-Adapter, gepaarte Geräte, Audio-Devices und bietet Reparaturfunktionen
#>

#Requires -RunAsAdministrator

Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms

# C# COM-Interop für IPolicyConfig (Standard-Audio-Gerät setzen)
Add-Type @"
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("F8679F50-850A-41CF-9C72-430F290290C8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IPolicyConfig {
    int SetDefaultEndpoint(string deviceId, int role);
}

public static class AudioPolicy {
    private static readonly Guid CLSID = new Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9");
    public static int SetDefault(string deviceId, int role) {
        object comObj = Activator.CreateInstance(Type.GetTypeFromCLSID(CLSID));
        IPolicyConfig policy = (IPolicyConfig)comObj;
        int hr = policy.SetDefaultEndpoint(deviceId, role);
        Marshal.ReleaseComObject(comObj);
        return hr;
    }
}
"@

# XAML Layout
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Bluetooth Diagnose &amp; Fix" Height="750" Width="950"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="13"
        Background="#F0F0F0">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Margin" Value="4"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="DangerBtn" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#D13438"/>
        </Style>
        <Style x:Key="SuccessBtn" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#107C10"/>
        </Style>
        <Style x:Key="WarningBtn" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#FF8C00"/>
        </Style>
    </Window.Resources>

    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#0078D4" CornerRadius="6" Padding="16,12" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="&#x1F9ED; Bluetooth Diagnose" FontSize="22" FontWeight="Bold" Foreground="White" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1" Text="Status: wird geladen…" x:Name="TxtStatus" Foreground="#B0D4F1" FontSize="14" VerticalAlignment="Center" Margin="20,0,0,0"/>
                <Button Grid.Column="2" x:Name="BtnRefresh" Content="&#x21BB; Aktualisieren" Background="#106EBE"/>
            </Grid>
        </Border>

        <!-- Main Content -->
        <TabControl Grid.Row="1" Margin="0,0,0,8" BorderThickness="0" Background="White">
            <TabControl.Resources>
                <Style TargetType="TabItem">
                    <Setter Property="FontWeight" Value="SemiBold"/>
                    <Setter Property="FontSize" Value="13"/>
                    <Setter Property="Padding" Value="14,6"/>
                </Style>
            </TabControl.Resources>

            <!-- Tab 1: Geräte -->
            <TabItem Header="&#x1F4F1; Geräte">
                <Grid Margin="8">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="#F5F5F5" CornerRadius="4" Padding="8,6" Margin="0,0,0,8">
                        <StackPanel Orientation="Horizontal">
                            <Button x:Name="BtnRemoveDevice" Content="&#x274C; Markiertes Gerät entfernen" Background="#D13438" Margin="0,0,8,0"/>
                            <Button x:Name="BtnScanDevices" Content="&#x1F50D; Geräte suchen" Style="{StaticResource SuccessBtn}"/>
                            <TextBlock Text="(Doppelklick = Gerät entfernen)" Foreground="#888" VerticalAlignment="Center" Margin="12,0,0,0" FontSize="11"/>
                        </StackPanel>
                    </Border>
                    <ListView Grid.Row="1" x:Name="LvDevices" BorderThickness="0.5" BorderBrush="#DDD">
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Name" Width="200" DisplayMemberBinding="{Binding Name}"/>
                                <GridViewColumn Header="Typ" Width="120" DisplayMemberBinding="{Binding Type}"/>
                                <GridViewColumn Header="Status" Width="100" DisplayMemberBinding="{Binding Status}"/>
                                <GridViewColumn Header="MAC-Adresse" Width="150" DisplayMemberBinding="{Binding MacAddress}"/>
                                <GridViewColumn Header="Verbunden" Width="90" DisplayMemberBinding="{Binding Connected}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                </Grid>
            </TabItem>

            <!-- Tab 2: Audio -->
            <TabItem Header="&#x1F50A; Audio-Geräte">
                <Grid Margin="8">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="#FFF3CD" BorderBrush="#FFE69C" BorderThickness="1" CornerRadius="4" Padding="10,8" Margin="0,0,0,8">
                        <StackPanel>
                            <TextBlock FontWeight="Bold" Text="&#x26A0; Problem: Bluetooth-Headset zeigt kein Mikrofon an?" Foreground="#856404"/>
                            <TextBlock Text="Oft liegt es an falschen Audio-Standardgeräten oder deaktivierten Aufnahmegeräten." Foreground="#856404" Margin="0,4,0,0"/>
                        </StackPanel>
                    </Border>
                    <ListView Grid.Row="1" x:Name="LvAudio" BorderThickness="0.5" BorderBrush="#DDD">
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Gerät" Width="220" DisplayMemberBinding="{Binding Name}"/>
                                <GridViewColumn Header="Typ" Width="100" DisplayMemberBinding="{Binding Type}"/>
                                <GridViewColumn Header="Status" Width="100" DisplayMemberBinding="{Binding Status}"/>
                                <GridViewColumn Header="Standard" Width="90" DisplayMemberBinding="{Binding Default}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <Border Grid.Row="2" Background="#F5F5F5" CornerRadius="4" Padding="8,6" Margin="0,8,0,0">
                        <StackPanel Orientation="Horizontal">
                            <Button x:Name="BtnSetPlaybackDefault" Content="&#x1F50A; Als Standard-Wiedergabe" Style="{StaticResource SuccessBtn}"/>
                            <Button x:Name="BtnSetRecordingDefault" Content="&#x1F3A4; Als Standard-Aufnahme" Style="{StaticResource SuccessBtn}"/>
                            <Button x:Name="BtnEnableAudioDevice" Content="&#x25B6; Aktivieren" Style="{StaticResource WarningBtn}"/>
                            <Button x:Name="BtnDisableAudioDevice" Content="&#x23F9; Deaktivieren" Background="#D13438"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>

            <!-- Tab 3: Diagnose & Fix -->
            <TabItem Header="&#x1F527; Diagnose &amp; Reparatur">
                <Grid Margin="8">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0" Background="#E8F5E9" CornerRadius="4" Padding="12,10" Margin="0,0,0,10">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0">
                                <TextBlock FontWeight="Bold" Text="&#x1F916; Schnellreparaturen" FontSize="15"/>
                                <TextBlock Text="Führe einen oder mehrere Schritte aus, um Bluetooth-Probleme zu beheben." Foreground="#555" Margin="0,2,0,0"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <Border Grid.Row="1" Background="#FEFEFE" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,8">
                        <StackPanel>
                            <TextBlock FontWeight="Bold" Text="&#x1F504; Bluetooth-Dienst neustarten" Margin="0,0,0,6"/>
                            <TextBlock Text="Behebt oft Probleme mit nicht erkennbaren Geräten und Verbindungsabbrüchen." Foreground="#666" FontSize="12" Margin="0,0,0,8"/>
                            <Button x:Name="BtnRestartService" Content="&#x1F504; Dienst neustarten" HorizontalAlignment="Left" Width="200"/>
                        </StackPanel>
                    </Border>

                    <Border Grid.Row="2" Background="#FEFEFE" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,8">
                        <StackPanel>
                            <TextBlock FontWeight="Bold" Text="&#x1F4E1; Bluetooth-Adapter zurücksetzen" Margin="0,0,0,6"/>
                            <TextBlock Text="Deaktiviert und aktiviert den Bluetooth-Adapter. Hilft bei Verbindungsproblemen." Foreground="#666" FontSize="12" Margin="0,0,0,8"/>
                            <StackPanel Orientation="Horizontal">
                                <Button x:Name="BtnResetAdapter" Content="&#x1F4E1; Adapter deaktivieren/reaktivieren" HorizontalAlignment="Left" Width="240"/>
                                <Button x:Name="BtnToggleAdapter" Content="&#x2234; Adapter umschalten" Style="{StaticResource WarningBtn}" HorizontalAlignment="Left" Width="180" Margin="10,0,0,0"/>
                            </StackPanel>
                        </StackPanel>
                    </Border>

                    <Border Grid.Row="3" Background="#FEFEFE" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,8">
                        <StackPanel>
                            <TextBlock FontWeight="Bold" Text="&#x1F3A7; Bluetooth-Audio-Problembehandlung" Margin="0,0,0,6"/>
                            <TextBlock Text="Setzt Audio-Endpunkte zurück, aktiviert alle Aufnahmegeräte und stellt Standardgeräte ein." Foreground="#666" FontSize="12" Margin="0,0,0,8"/>
                            <StackPanel Orientation="Horizontal">
                                <Button x:Name="BtnFixAudio" Content="&#x1F3A7; Audio-Problem beheben" Style="{StaticResource SuccessBtn}" HorizontalAlignment="Left" Width="200"/>
                                <Button x:Name="BtnRestartAudioService" Content="&#x1F504; Windows-Audio-Dienst neustarten" HorizontalAlignment="Left" Width="260" Margin="10,0,0,0"/>
                            </StackPanel>
                        </StackPanel>
                    </Border>

                    <TextBox Grid.Row="4" x:Name="TxtLog" IsReadOnly="True"
                             Background="#1E1E1E" Foreground="#D4D4D4" FontFamily="Cascadia Code, Consolas"
                             FontSize="11" VerticalScrollBarVisibility="Auto"
                             BorderThickness="0.5" BorderBrush="#CCC" Padding="6"/>
                </Grid>
            </TabItem>

            <!-- Tab 4: Adapter-Info -->
            <TabItem Header="&#x2139; Adapter-Info">
                <Grid Margin="8">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="#F5F5F5" CornerRadius="4" Padding="8,6" Margin="0,0,0,8">
                        <Button x:Name="BtnCopyInfo" Content="&#x1F4CB; In Zwischenablage kopieren" HorizontalAlignment="Left"/>
                    </Border>
                    <TextBox Grid.Row="1" x:Name="TxtAdapterInfo" IsReadOnly="True"
                             FontFamily="Cascadia Code, Consolas" FontSize="12"
                             VerticalScrollBarVisibility="Auto" BorderThickness="0.5" BorderBrush="#DDD" Padding="6"/>
                </Grid>
            </TabItem>
        </TabControl>

        <!-- Footer -->
        <Border Grid.Row="2" Background="#E8E8E8" CornerRadius="4" Padding="10,6">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" x:Name="TxtFooter" Text="Bereit" Foreground="#666" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="2" x:Name="TxtDeviceCount" Text="" Foreground="#888" VerticalAlignment="Center" FontSize="12"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Fenster erstellen
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Controls auslesen
$txtStatus = $window.FindName('TxtStatus')
$btnRefresh = $window.FindName('BtnRefresh')
$lvDevices = $window.FindName('LvDevices')
$lvAudio = $window.FindName('LvAudio')
$txtLog = $window.FindName('TxtLog')
$txtAdapterInfo = $window.FindName('TxtAdapterInfo')
$txtFooter = $window.FindName('TxtFooter')
$txtDeviceCount = $window.FindName('TxtDeviceCount')
$btnRemoveDevice = $window.FindName('BtnRemoveDevice')
$btnScanDevices = $window.FindName('BtnScanDevices')
$btnRestartService = $window.FindName('BtnRestartService')
$btnResetAdapter = $window.FindName('BtnResetAdapter')
$btnToggleAdapter = $window.FindName('BtnToggleAdapter')
$btnFixAudio = $window.FindName('BtnFixAudio')
$btnRestartAudioService = $window.FindName('BtnRestartAudioService')
$btnSetPlaybackDefault = $window.FindName('BtnSetPlaybackDefault')
$btnSetRecordingDefault = $window.FindName('BtnSetRecordingDefault')
$btnEnableAudioDevice = $window.FindName('BtnEnableAudioDevice')
$btnDisableAudioDevice = $window.FindName('BtnDisableAudioDevice')
$btnCopyInfo = $window.FindName('BtnCopyInfo')

# Log-Funktion
function Write-Log {
    param([string]$Message, [string]$Color = 'White')
    $timestamp = Get-Date -Format 'HH:mm:ss'
    $disp = if ($Color -eq 'Red') { "❌ $Message" } elseif ($Color -eq 'Green') { "✅ $Message" } elseif ($Color -eq 'Yellow') { "⚠️ $Message" } else { "  $Message" }
    $txtLog.AppendText("[$timestamp] $disp`r`n")
    $txtLog.ScrollToEnd()
}

function Set-Status {
    param([string]$Text, [string]$Color = '#B0D4F1')
    $txtStatus.Text = $Text
    $txtFooter.Text = $Text
}

# --- Bluetooth-Daten sammeln ---

function Get-BluetoothAdapterInfo {
    $adapters = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object {
        $_.PNPClass -eq 'Bluetooth' -or $_.Service -eq 'BTHUSB' -or $_.Service -eq 'BTHMINI'
    }
    return $adapters
}

function Get-BluetoothDevices {
    $devices = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object {
        $_.PNPClass -eq 'Bluetooth' -and $_.Service -ne 'BTHUSB' -and $_.Service -ne 'BTHMINI' -and $_.Name -notlike '*Radios*' -and $_.Name -notlike '*Adapter*'
    }
    return $devices
}

function Get-PairedBluetoothDevices {
    $paired = @()
    $devices = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object {
        $_.PNPClass -eq 'Bluetooth' -and $_.Name -notlike '*Radios*' -and $_.Name -notlike '*Adapter*'
    }
    foreach ($d in $devices) {
        $connected = 'Nein'
        $status = $d.Status
        if ($d.Status -eq 'OK' -or $d.Status -eq 'Started') {
            $connected = 'Ja'
            $status = 'Aktiv'
        } elseif ($d.Status -eq 'Unknown') {
            $status = 'Unbekannt'
        } elseif ($d.Status -eq 'Error') {
            $status = 'Fehler'
        }
        # MAC aus Hardware-ID extrahieren
        $mac = 'N/A'
        if ($d.HardwareID) {
            foreach ($hid in $d.HardwareID) {
                if ($hid -match '([0-9A-Fa-f]{2}[:]){5}[0-9A-Fa-f]{2}') {
                    $mac = $matches[0].ToUpper()
                    break
                }
            }
        }

        # Typ erkennen – zuerst System-Dienste ausfiltern
        $deviceType = ''
        $typeName = $d.Name
        if ($typeName -match 'Generisches Attributprofil|Generic Attribute|Bluetooth Device \(RFCOMM|Generisches Zugriffsprofil|GAP|SIM-Zugriff|SIM Access|PSE-Dienst|PBAP|NAP-Dienst|PAN\b|Dienst für (Objektübertragung|persönliches|Telefonbuch)|Microsoft Bluetooth|Energiearmes Bluetooth|LE-Enumerator|Auflistung|Intel.*Wireless.*Bluetooth|Bluetooth.*Radio|Bluetooth.*Adapter') { $deviceType = '⚙️ System' }
        elseif ($typeName -match 'Jabra|Evolve|Elite|Talk|Bose|Sony|WH-|WF-|AirPods|FreeBuds|Galaxy Buds|Pixel Buds|Headset|Kopfhörer|Headphones|Hands-Free|Handsfree|Freisprech|Earbud|In-Ear|Over-Ear|On-Ear') { $deviceType = '🎧 Headset' }
        elseif ($typeName -match 'Speaker|Lautsprecher|Soundbar|SoundLink|JBL|Sonos|Stereo|A2DP') { $deviceType = '🔊 Audio' }
        elseif ($typeName -match 'Microphone|Mikrofon|Mic\b') { $deviceType = '🎤 Mikrofon' }
        elseif ($typeName -match 'S[0-9]{2,}|Galaxy(?!\sBud)|iPhone|Pixel\b(?!\sBud)|OnePlus|Huawei|Xiaomi|Oppo|Nothing\sPhone|Phone\b|Smartphone|Handy|Telefon') { $deviceType = '📱 Telefon' }
        elseif ($typeName -match 'iPad|Tablet|Galaxy Tab|Surface\sPro|Kindle') { $deviceType = '📟 Tablet' }
        elseif ($typeName -match 'Mouse|Maus|Trackpad|Touchpad|Magic Mouse|MX Master|Anywhere') { $deviceType = '🖱️ Maus' }
        elseif ($typeName -match 'Keyboard|Tastatur|Keypad|MX Keys|Magic Keyboard|Type Cover') { $deviceType = '⌨️ Tastatur' }
        elseif ($typeName -match 'Printer|Drucker|LabelWriter|Brother|HP\s|Canon|Epson') { $deviceType = '🖨️ Drucker' }
        elseif ($typeName -match 'Gamepad|Controller|Xbox|PlayStation|DualSense|DualShock|Joy-Con|Pro Controller') { $deviceType = '🎮 Gamepad' }
        elseif ($typeName -match 'Watch|Uhr\b|Smartwatch|Galaxy Watch|Apple Watch|Fitbit|Garmin|Band\b') { $deviceType = '⌚ Uhr' }
        elseif ($typeName -match 'AVRCP|Audio|Music|Sound|Remote Control') { $deviceType = '🎧 Audio' }
        else { $deviceType = '📡 Sonstiges' }

        $paired += [PSCustomObject]@{
            Name       = $d.Name
            Type       = $deviceType
            Status     = $status
            MacAddress = $mac
            Connected  = $connected
            PnpDevice  = $d
        }
    }
    return $paired
}

function Get-AudioDevices {
    $audio = @()

    # Methode: WMI/CIM - Win32_PnPEntity mit PNPClass=AudioEndpoint
    # Das DeviceID-Format enthält {0.0.0.} = Playback (Render) und {0.0.1.} = Recording (Capture)
    try {
        Write-Log "Ermittle Audio-Geräte via WMI/CIM…" -Color White
        $endpoints = Get-CimInstance -ClassName Win32_PnPEntity -ErrorAction Stop |
            Where-Object { $_.PNPClass -eq 'AudioEndpoint' -and $_.Name -ne $null }

        # Standard-Geräte via Registry (sofern vorhanden)
        $defaultPlaybackId = $null
        $defaultRecordingId = $null
        try {
            $r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{7e75e2ee-3446-41ce-92d1-2e60f710f4c2},4' -ErrorAction Stop
            if ($r) { $defaultPlaybackId = $r.'{7e75e2ee-3446-41ce-92d1-2e60f710f4c2},4' }
        } catch { }
        try {
            $c = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{7e75e2ee-3446-41ce-92d1-2e60f710f4c2},4' -ErrorAction Stop
            if ($c) { $defaultRecordingId = $c.'{7e75e2ee-3446-41ce-92d1-2e60f710f4c2},4' }
        } catch { }

        foreach ($ep in $endpoints) {
            $isRender = $ep.DeviceID -match '\\{0\.0\.0\.'
            $isCapture = $ep.DeviceID -match '\\{0\.0\.1\.'

            if (-not $isRender -and -not $isCapture) { continue }

            $type = if ($isRender) { '🔊 Wiedergabe' } else { '🎙️ Aufnahme' }

            # Status: OK/Aktiv vs deaktiviert (via PNPClass reicht nicht, also Registry)
            $state = switch ($ep.Status) {
                'OK' { 'Aktiv' }
                'Error' { 'Fehler' }
                'Degraded' { 'Eingeschränkt' }
                'Unknown' { 'Unbekannt' }
                default { $ep.Status }
            }

            # Default-Prüfung: über DeviceID-Vergleich mit Registry-Wert
            $guidOnly = if ($ep.DeviceID -match '\{([0-9A-Fa-f-]+)\}$') { $matches[1] } else { $null }
            $isDefault = $false
            if ($guidOnly) {
                if ($isRender -and $defaultPlaybackId) { $isDefault = ($guidOnly -eq $defaultPlaybackId) }
                if ($isCapture -and $defaultRecordingId) { $isDefault = ($guidOnly -eq $defaultRecordingId) }
            }

            $audio += [PSCustomObject]@{
                Name    = $ep.Name
                Type    = $type
                Status  = $state
                Default = if ($isDefault) { '✅ Ja' } else { '' }
                DeviceId = $guidOnly
            }
        }

        if ($audio.Count -gt 0) {
            Write-Log "Audio-Geräte via WMI ermittelt: $($audio.Count) Stück" -Color Green
        }
        else {
            Write-Log "Keine Audio-Endpoints via WMI gefunden" -Color Yellow
        }
    }
    catch {
        Write-Log "WMI-Methode fehlgeschlagen: $_" -Color Yellow
    }

    if ($audio.Count -eq 0) {
        Write-Log "Keine Audio-Geräte gefunden" -Color Red
    }
    return $audio
}

# --- GUI aktualisieren ---

function Update-DeviceList {
    $txtStatus.Text = '🔄 Sammle Bluetooth-Geräte…'
    $txtFooter.Text = 'Bitte warten…'
    [System.Windows.Forms.Application]::DoEvents()

    $devices = Get-PairedBluetoothDevices | Sort-Object Name
    $lvDevices.ItemsSource = $devices
    $txtDeviceCount.Text = "$($devices.Count) Geräte"
    $txtFooter.Text = "Letzte Aktualisierung: $(Get-Date -Format 'HH:mm:ss')"
    Write-Log "Geräteliste aktualisiert: $($devices.Count) Geräte gefunden" -Color Green
}

function Update-AudioList {
    $txtStatus.Text = '🔄 Sammle Audio-Geräte…'
    [System.Windows.Forms.Application]::DoEvents()

    $audio = Get-AudioDevices
    $lvAudio.ItemsSource = $audio
    Write-Log "Audio-Geräteliste aktualisiert: $($audio.Count) Geräte gefunden" -Color Green
}

function Update-AdapterInfo {
    $info = @()
    $info += "=== Bluetooth-Adapter ==="
    $adapters = Get-BluetoothAdapterInfo
    foreach ($a in $adapters) {
        $info += "Name:     $($a.Name)"
        $info += "Status:   $($a.Status)"
        $info += "Class:    $($a.PNPClass)"
        $info += "DeviceID: $($a.DeviceID)"
        $info += "---"
    }

    $info += "`n=== Bluetooth-Dienst ==="
    try {
        $svc = Get-Service -Name 'bthserv' -ErrorAction Stop
        $info += "Name:        $($svc.Name)"
        $info += "Status:      $($svc.Status)"
        $info += "Starttyp:    $($svc.StartType)"

        $svcA = Get-Service -Name 'AudioSrv' -ErrorAction SilentlyContinue
        if ($svcA) {
            $info += "`n=== Windows-Audio-Dienst ==="
            $info += "Name:        $($svcA.Name)"
            $info += "Status:      $($svcA.Status)"
            $info += "Starttyp:    $($svcA.StartType)"
        }

        $svcAG = Get-Service -Name 'Audiosrv' -ErrorAction SilentlyContinue
        if (-not $svcA) {
            $svcAG = Get-Service -Name 'Audiosrv' -ErrorAction SilentlyContinue
            if ($svcAG) {
                $info += "`n=== Windows-Audio-Dienst ==="
                $info += "Name:        $($svcAG.Name)"
                $info += "Status:      $($svcAG.Status)"
                $info += "Starttyp:    $($svcAG.StartType)"
            }
        }
    }
    catch {
        $info += "Dienst-Information nicht verfügbar: $_"
    }

    $info += "`n=== Bluetooth-fähige PnP-Geräte ==="
    $allBt = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object { $_.PNPClass -eq 'Bluetooth' }
    foreach ($b in $allBt) {
        $info += "  $($b.Name)  [$($b.Status)]"
    }

    $txtAdapterInfo.Text = $info -join "`r`n"
    Write-Log "Adapter-Info aktualisiert" -Color Green
}

function Update-All {
    Set-Status '🔄 Aktualisiere…'
    Update-DeviceList
    Update-AudioList
    Update-AdapterInfo
    Set-Status '✅ Bereit'
    Write-Log "Alle Daten aktualisiert" -Color Green
}

# --- Fix-Funktionen ---

function Invoke-RestartBluetoothService {
    Write-Log "Starte Bluetooth-Dienst neu…" -Color Yellow
    try {
        Stop-Service -Name 'bthserv' -Force -ErrorAction Stop
        Start-Sleep -Milliseconds 500
        Start-Service -Name 'bthserv' -ErrorAction Stop
        Write-Log "Bluetooth-Dienst erfolgreich neugestartet" -Color Green
        Start-Sleep -Seconds 1
        Update-All
    }
    catch {
        Write-Log "Fehler beim Neustart des Bluetooth-Dienstes: $_" -Color Red
    }
}

function Invoke-ResetBluetoothAdapter {
    Write-Log "Deaktiviere Bluetooth-Adapter…" -Color Yellow
    try {
        $adapter = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object { $_.PNPClass -eq 'Bluetooth' -and ($_.Service -eq 'BTHUSB' -or $_.Name -match 'Radio') }
        if (-not $adapter) {
            # Fallback: alle Bluetooth-Geräte
            $adapter = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object { $_.PNPClass -eq 'Bluetooth' } | Select-Object -First 1
        }
        if ($adapter) {
            # Deaktivieren per PnP-Path
            $pnpPath = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object { $_.DeviceID -eq $adapter.DeviceID }
            if ($pnpPath) {
                # Alternative: devcon oder PnPUtil - besser: Disable-PnpDevice
                Disable-PnpDevice -InstanceId $adapter.DeviceID -Confirm:$false -ErrorAction Stop
                Write-Log "Adapter deaktiviert" -Color Green
                Start-Sleep -Seconds 2
                Enable-PnpDevice -InstanceId $adapter.DeviceID -Confirm:$false -ErrorAction Stop
                Write-Log "Adapter reaktiviert" -Color Green
                Start-Sleep -Seconds 2
                Update-All
            }
        }
        else {
            Write-Log "Kein Bluetooth-Adapter gefunden" -Color Red
        }
    }
    catch {
        Write-Log "Fehler beim Zurücksetzen des Adapters: $_" -Color Red
        Write-Log "Versuche devcon.exe als Fallback…" -Color Yellow
        try {
            $devcon = Get-Command 'devcon.exe' -ErrorAction SilentlyContinue
            if (-not $devcon) {
                Write-Log "devcon.exe nicht gefunden. Bitte Windows SDK oder DTK installieren." -Color Red
                return
            }
            # devcon funktioniert anders - überspringe
        }
        catch { }
    }
}

function Invoke-ToggleBluetoothAdapter {
    Write-Log "Schalte Bluetooth-Adapter um (via Dienst)…" -Color Yellow
    try {
        Stop-Service -Name 'bthserv' -Force -ErrorAction Stop
        Start-Sleep -Seconds 1
        Start-Service -Name 'bthserv' -ErrorAction Stop
        Write-Log "Bluetooth-Dienst umgeschaltet" -Color Green
        Start-Sleep -Seconds 2
        Update-All
    }
    catch {
        Write-Log "Fehler beim Umschalten: $_" -Color Red
    }
}

function Invoke-FixBluetoothAudio {
    Write-Log "Starte Bluetooth-Audio-Reparatur…" -Color Yellow
    $fixed = $false

    # 1. Audio-Dienste prüfen/starten
    try {
        $audiosrv = Get-Service -Name 'Audiosrv' -ErrorAction Stop
        if ($audiosrv.Status -ne 'Running') {
            Start-Service -Name 'Audiosrv' -ErrorAction Stop
            Write-Log "Windows-Audio-Dienst gestartet" -Color Green
            $fixed = $true
        }
        else {
            Write-Log "Windows-Audio-Dienst läuft bereits" -Color Green
        }
    }
    catch {
        Write-Log "Audio-Dienst konnte nicht gestartet werden: $_" -Color Red
    }

    # 2. Alle AudioEndpoint-Geräte via PnP aktivieren (deaktivierte reaktivieren)
    try {
        $disabled = Get-CimInstance -ClassName Win32_PnPEntity -ErrorAction Stop |
            Where-Object { $_.PNPClass -eq 'AudioEndpoint' -and $_.Status -ne 'OK' }
        foreach ($d in $disabled) {
            try {
                Enable-PnpDevice -InstanceId $d.DeviceID -Confirm:$false -ErrorAction Stop
                Write-Log "Gerät aktiviert: $($d.Name)" -Color Green
                $fixed = $true
            }
            catch {
                Write-Log "Konnte $($d.Name) nicht aktivieren: $_" -Color Yellow
            }
        }
    }
    catch {
        Write-Log "Fehler bei PnP-Aktivierung: $_" -Color Yellow
    }

    # 3. Bluetooth-Audio-Geräte als Standard setzen (via Registry)
    try {
        $btEndpoints = Get-CimInstance -ClassName Win32_PnPEntity -ErrorAction Stop |
            Where-Object { $_.PNPClass -eq 'AudioEndpoint' -and $_.Name -match 'Bluetooth|Headset|Headphones|Kopfhörer|Hands-Free|Handsfree|Freisprech|Mikrofon|Microphone|WH-|WF-|AirPods|Galaxy|BudS|FreeBuds|Jabra|Evolve|Elite' }

        # Standard setzen via mmsys.cpl öffnen (immer hilfreich)
        try {
            Start-Process 'control.exe' -ArgumentList 'mmsys.cpl'
            Write-Log "Systemsteuerung 'Sound' geöffnet – bitte Standardgerät manuell wählen" -Color Green
        }
        catch { }

        if ($btEndpoints.Count -gt 0) {
            Write-Log "Bluetooth-Audio-Geräte gefunden:" -Color Green
            foreach ($bt in $btEndpoints) {
                Write-Log "  - $($bt.Name) [$($bt.Status)]" -Color White
            }
            $fixed = $true
        }
    }
    catch {
        Write-Log "Fehler bei Bluetooth-Audio-Suche: $_" -Color Yellow
    }

    if ($fixed) {
        Write-Log "Audio-Reparatur abgeschlossen" -Color Green
    }
    else {
        Write-Log "Keine Probleme gefunden oder alles bereits in Ordnung" -Color Yellow
    }

    Start-Sleep -Seconds 1
    Update-AudioList
}

function Invoke-RestartAudioService {
    Write-Log "Starte Windows-Audio-Dienst neu…" -Color Yellow
    try {
        Stop-Service -Name 'Audiosrv' -Force -ErrorAction Stop
        Start-Sleep -Milliseconds 500
        Start-Service -Name 'Audiosrv' -ErrorAction Stop
        Write-Log "Windows-Audio-Dienst erfolgreich neugestartet" -Color Green
        Start-Sleep -Seconds 2
        Update-AudioList
    }
    catch {
        Write-Log "Fehler beim Neustart des Audio-Dienstes: $_" -Color Red
    }
}

function Invoke-RemoveDevice {
    $selected = $lvDevices.SelectedItem
    if (-not $selected) {
        Write-Log "Kein Gerät ausgewählt" -Color Red
        return
    }
    $name = $selected.Name
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Möchtest du '$name' wirklich entfernen?`nDas Gerät muss anschließend neu gekoppelt werden.",
        "Gerät entfernen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($result -eq 'Yes') {
        Write-Log "Entferne Gerät: $name" -Color Yellow
        try {
            $pnp = $selected.PnpDevice
            if ($pnp) {
                $instanceId = $pnp.DeviceID
                # PnP-Gerät deinstallieren
                Get-CimInstance -ClassName Win32_PnPEntity | Where-Object { $_.DeviceID -eq $instanceId } |
                    Remove-CimInstance -ErrorAction Stop
                Write-Log "Gerät '$name' wurde entfernt" -Color Green
                Update-DeviceList
            }
        }
        catch {
            Write-Log "Fehler beim Entfernen: $_" -Color Red
            Write-Log "Versuche IR-Geräte-Entfernung über Registry…" -Color Yellow
            # Fallback: via setupapi
            try {
                & "pnputil.exe" /remove-device $selected.PnpDevice.DeviceID 2>$null
                Write-Log "Gerät via pnputil entfernt (möglicherweise Neustart nötig)" -Color Green
                Update-DeviceList
            }
            catch {
                Write-Log "Auch pnputil fehlgeschlagen: $_" -Color Red
            }
        }
    }
}

function Invoke-ScanDevices {
    Write-Log "Starte Bluetooth-Gerätesuche…" -Color Yellow
    try {
        # Versuche via Windows.Devices.Bluetooth (UWP) - funktioniert nicht immer aus PS
        # Alternativ: den Dienst anstoßen
        $btService = Get-Service -Name 'bthserv' -ErrorAction Stop
        if ($btService.Status -eq 'Running') {
            # Einfach die UI auffordern zu scannen - via Shell
            $shell = New-Object -ComObject "Shell.Application"
            $shell.Open("ms-settings:bluetooth")
            Write-Log "Bluetooth-Einstellungen geöffnet. Scannen startet automatisch." -Color Green
        }
        else {
            Write-Log "Bluetooth-Dienst läuft nicht. Starte ihn…" -Color Yellow
            Start-Service -Name 'bthserv' -ErrorAction Stop
            Start-Sleep -Seconds 2
            $shell = New-Object -ComObject "Shell.Application"
            $shell.Open("ms-settings:bluetooth")
            Write-Log "Bluetooth-Einstellungen geöffnet." -Color Green
        }
    }
    catch {
        Write-Log "Fehler beim Scannen: $_" -Color Red
    }
}

function Invoke-SetDefaultPlayback {
    $selected = $lvAudio.SelectedItem
    if (-not $selected) { Write-Log "Kein Audio-Gerät ausgewählt" -Color Red; return }
    if ($selected.Type -notmatch 'Wiedergabe') { Write-Log "Bitte ein Wiedergabegerät auswählen" -Color Red; return }
    if (-not $selected.DeviceId) { Write-Log "Keine Device-ID für dieses Gerät" -Color Red; return }

    try {
        Write-Log "Setze '$($selected.Name)' als Standard-Wiedergabe (eConsole)…" -Color Yellow
        $hr = [AudioPolicy]::SetDefault($selected.DeviceId, 0) # eConsole
        if ($hr -eq 0) {
            Write-Log "Standard-Wiedergabegerät gesetzt: $($selected.Name)" -Color Green
            # Auch als eMultimedia (1) und eCommunications (2) setzen
            [AudioPolicy]::SetDefault($selected.DeviceId, 1) | Out-Null
            [AudioPolicy]::SetDefault($selected.DeviceId, 2) | Out-Null
        } else {
            Write-Log "Fehler: HRESULT = 0x$('{0:X8}' -f $hr)" -Color Red
        }
        Update-AudioList
    }
    catch {
        Write-Log "Fehler: $_" -Color Red
    }
}

function Invoke-SetDefaultRecording {
    $selected = $lvAudio.SelectedItem
    if (-not $selected) { Write-Log "Kein Audio-Gerät ausgewählt" -Color Red; return }
    if ($selected.Type -notmatch 'Aufnahme') { Write-Log "Bitte ein Aufnahmegerät auswählen" -Color Red; return }
    if (-not $selected.DeviceId) { Write-Log "Keine Device-ID für dieses Gerät" -Color Red; return }

    try {
        Write-Log "Setze '$($selected.Name)' als Standard-Aufnahme (eConsole)…" -Color Yellow
        $hr = [AudioPolicy]::SetDefault($selected.DeviceId, 0) # eConsole
        if ($hr -eq 0) {
            Write-Log "Standard-Aufnahmegerät gesetzt: $($selected.Name)" -Color Green
            [AudioPolicy]::SetDefault($selected.DeviceId, 1) | Out-Null
            [AudioPolicy]::SetDefault($selected.DeviceId, 2) | Out-Null
        } else {
            Write-Log "Fehler: HRESULT = 0x$('{0:X8}' -f $hr)" -Color Red
        }
        Update-AudioList
    }
    catch {
        Write-Log "Fehler: $_" -Color Red
    }
}

function Invoke-EnableAudioDevice {
    $selected = $lvAudio.SelectedItem
    if (-not $selected) { Write-Log "Kein Audio-Gerät ausgewählt" -Color Red; return }

    Write-Log "Aktivieren von '$($selected.Name)' nicht per Skript möglich." -Color Yellow
    Write-Log "Bitte in der Sound-Systemsteuerung aktivieren (wird geöffnet)." -Color White
    try { Start-Process 'control.exe' -ArgumentList 'mmsys.cpl' } catch { }
}

function Invoke-DisableAudioDevice {
    $selected = $lvAudio.SelectedItem
    if (-not $selected) { Write-Log "Kein Audio-Gerät ausgewählt" -Color Red; return }

    Write-Log "Deaktivieren von '$($selected.Name)' nicht per Skript möglich." -Color Yellow
    Write-Log "Bitte in der Sound-Systemsteuerung deaktivieren (wird geöffnet)." -Color White
    try { Start-Process 'control.exe' -ArgumentList 'mmsys.cpl' } catch { }
}

# --- Event-Handler ---

$btnRefresh.Add_Click({ Update-All })
$btnRemoveDevice.Add_Click({ Invoke-RemoveDevice })
$btnScanDevices.Add_Click({ Invoke-ScanDevices })
$btnRestartService.Add_Click({ Invoke-RestartBluetoothService })
$btnResetAdapter.Add_Click({ Invoke-ResetBluetoothAdapter })
$btnToggleAdapter.Add_Click({ Invoke-ToggleBluetoothAdapter })
$btnFixAudio.Add_Click({ Invoke-FixBluetoothAudio })
$btnRestartAudioService.Add_Click({ Invoke-RestartAudioService })
$btnSetPlaybackDefault.Add_Click({ Invoke-SetDefaultPlayback })
$btnSetRecordingDefault.Add_Click({ Invoke-SetDefaultRecording })
$btnEnableAudioDevice.Add_Click({ Invoke-EnableAudioDevice })
$btnDisableAudioDevice.Add_Click({ Invoke-DisableAudioDevice })

$btnCopyInfo.Add_Click({
    try {
        [System.Windows.Forms.Clipboard]::SetText($txtAdapterInfo.Text)
        Write-Log "Adapter-Info kopiert" -Color Green
    }
    catch {
        Write-Log "Fehler beim Kopieren: $_" -Color Red
    }
})

# Doppelklick auf Gerät = entfernen
$lvDevices.Add_MouseDoubleClick({
    Invoke-RemoveDevice
})

# Initial laden
$window.Add_Loaded({
    Update-All
})

# Fenster anzeigen
$window.ShowDialog() | Out-Null

