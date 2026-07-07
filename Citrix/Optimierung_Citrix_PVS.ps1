#================================================================================
# Skript   : PVS2507-Optimierung.ps1
# Zweck    : Netzwerk- und Windows-Optimierung für Citrix PVS Target Devices
#            inkl. Sicherung, GUI mit Live-Log und vollständigem Rollback.
# Autor    : Rocco Ammon
# Version  : 4.8
# Hinweis  : IP, Domäne und Servername sind bereits vorkonfiguriert.
#            Muss als Administrator ausgeführt werden!
#================================================================================

#--------------------------------------------------------------------------------
# 1. VARIABLEN-DEFINITION (am Anfang, gemäß Best Practice)
#--------------------------------------------------------------------------------
$Global:LogVerzeichnis  = "C:\ScriptLog"
$Global:LogDatei        = Join-Path $LogVerzeichnis "PVS2507_Optimierung.log"
$Global:BackupDatei     = Join-Path $LogVerzeichnis "PVS2507_Backup.json"
$Global:ClassGuid       = '{4D36E972-E325-11CE-BFC1-08002bE10318}'  # Netzwerkadapter-Klasse
$Global:GuiLogBox       = $null   # Referenz auf die RichTextBox (wird später gesetzt)
$Global:LogForm         = $null   # Referenz auf das Formular (für Refresh)

# Zu deaktivierende Offload-Eigenschaften (Anzeigenamen)
$Global:OffloadFeatures = @(
    'IPv4 Checksum Offload',
    'TCP Checksum Offload (IPv4)',
    'TCP Checksum Offload (IPv6)',
    'UDP Checksum Offload (IPv4)',
    'UDP Checksum Offload (IPv6)',
    'Jumbo Packet'
)

# Sicherstellen, dass das Log-Verzeichnis existiert
if (-not (Test-Path $Global:LogVerzeichnis)) {
    New-Item -Path $Global:LogVerzeichnis -ItemType Directory -Force | Out-Null
}

#--------------------------------------------------------------------------------
# 2. HILFSFUNKTION: Logging (Datei + Konsole + Live-GUI)
#--------------------------------------------------------------------------------
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Nachricht,
        [ValidateSet('INFO','SUCCESS','WARN','ERROR')][string]$Level = 'INFO'
    )

    $zeit    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $zeile   = "[$zeit] [$Level] $Nachricht"

    # 2a) In Log-Datei schreiben
    try { Add-Content -Path $Global:LogDatei -Value $zeile -Encoding UTF8 } catch { }

    # 2b) In Konsole schreiben (mit Farbe)
    switch ($Level) {
        'SUCCESS' { Write-Host $zeile -ForegroundColor DarkGreen }
        'WARN'    { Write-Host $zeile -ForegroundColor DarkOrange }
        'ERROR'   { Write-Host $zeile -ForegroundColor Red }
        default   { Write-Host $zeile -ForegroundColor Black }
    }

    # 2c) LIVE in die GUI-Box schreiben (farblich passend)
    if ($Global:GuiLogBox -ne $null) {
        try {
            $farbe = switch ($Level) {
                'SUCCESS' { [System.Drawing.Color]::FromArgb(46,139,87) }   # SeaGreen
                'WARN'    { [System.Drawing.Color]::FromArgb(230,126,34) }  # freundliches Orange
                'ERROR'   { [System.Drawing.Color]::FromArgb(192,57,43) }   # gedämpftes Rot
                default   { [System.Drawing.Color]::FromArgb(52,73,94) }    # dunkles Blaugrau
            }
            $Global:GuiLogBox.SelectionStart  = $Global:GuiLogBox.TextLength
            $Global:GuiLogBox.SelectionLength = 0
            $Global:GuiLogBox.SelectionColor  = $farbe
            $Global:GuiLogBox.AppendText("$zeile`r`n")
            $Global:GuiLogBox.ScrollToCaret()
            # UI sofort aktualisieren, damit der Log LIVE erscheint
            [System.Windows.Forms.Application]::DoEvents()
        } catch { }
    }
}

#--------------------------------------------------------------------------------
# 3. FUNKTION: NIC-Energieverwaltung deaktivieren (robust, 3 Methoden)
#--------------------------------------------------------------------------------
function Disable-PVSNicPowerManagement {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$AdapterName,
        [Parameter(Mandatory = $false)][hashtable]$BackupStore
    )

    $erfolg       = $false
    $rebootNoetig = $false

    try {
        $netAdapter = Get-NetAdapter -Name $AdapterName -ErrorAction Stop
        $pnpId      = $netAdapter.PnPDeviceID
        Write-Log "  Bearbeite NIC-Energieverwaltung: $AdapterName" "INFO"

        #=== METHODE 1: WMI MSPower_DeviceEnable (kein Reboot nötig) ===========
        try {
            $powerMgmt = Get-CimInstance -Namespace 'root\wmi' -ClassName 'MSPower_DeviceEnable' -ErrorAction Stop |
                         Where-Object { $_.InstanceName -match [regex]::Escape($pnpId) }

            if ($powerMgmt) {
                if ($BackupStore) { $BackupStore["NIC_PowerMgmt_$AdapterName"] = $powerMgmt.Enable }

                if ($powerMgmt.Enable -eq $true) {
                    Set-CimInstance -InputObject $powerMgmt -Property @{ Enable = $false } -ErrorAction Stop | Out-Null
                    Write-Log "  Energieverwaltung via WMI deaktiviert (kein Neustart nötig)." "SUCCESS"
                } else {
                    Write-Log "  Energieverwaltung war via WMI bereits deaktiviert." "INFO"
                }
                $erfolg = $true
            }
        }
        catch { Write-Log "  WMI-Methode nicht verfügbar: $($_.Exception.Message)" "WARN" }

        #=== METHODE 2: Registry IdleInWorkingState (Reboot nötig) =============
        if (-not $erfolg) {
            try {
                $basisPfad = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\$Global:ClassGuid"
                $subKeys   = Get-ChildItem -Path $basisPfad -ErrorAction Stop |
                             Where-Object { $_.PSChildName -match '^\d{4}$' }

                foreach ($key in $subKeys) {
                    $props     = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                    $devInstId = $props.DeviceInstanceID
                    if ($devInstId) {
                        $wdfPfad = "HKLM:\SYSTEM\CurrentControlSet\Enum\$devInstId\Device Parameters\WDF"
                        if (Test-Path -Path $wdfPfad) {
                            $idleWert = (Get-ItemProperty -Path $wdfPfad -ErrorAction SilentlyContinue).IdleInWorkingState
                            if ($BackupStore) {
                                $BackupStore["NIC_IdleInWorkingState_$($key.PSChildName)"] = @{
                                    Pfad = $wdfPfad; Wert = $idleWert
                                }
                            }
                            if ($null -eq $idleWert) {
                                New-ItemProperty -Path $wdfPfad -Name 'IdleInWorkingState' -Value 0 -PropertyType DWord -Force -ErrorAction Stop | Out-Null
                            } else {
                                Set-ItemProperty -Path $wdfPfad -Name 'IdleInWorkingState' -Value 0 -ErrorAction Stop | Out-Null
                            }
                            Write-Log "  Registry-Wert 'IdleInWorkingState' gesetzt (=0). NEUSTART erforderlich." "WARN"
                            $erfolg = $true; $rebootNoetig = $true
                        }
                    }
                }
            }
            catch { Write-Log "  Registry-Methode fehlgeschlagen: $($_.Exception.Message)" "WARN" }
        }

        #=== METHODE 3: Cmdlet-Fallback ========================================
        if (-not $erfolg) {
            try {
                Disable-NetAdapterPowerManagement -Name $AdapterName -ErrorAction Stop
                Write-Log "  Energieverwaltung via Disable-NetAdapterPowerManagement deaktiviert." "SUCCESS"
                $erfolg = $true
            }
            catch { Write-Log "  Disable-NetAdapterPowerManagement nicht möglich: $($_.Exception.Message)" "WARN" }
        }

        if (-not $erfolg) {
            Write-Log "  NIC-Energieverwaltung auf $AdapterName konnte mit keiner Methode geändert werden." "WARN"
        }
    }
    catch { Write-Log "  Kritischer Fehler bei $AdapterName : $($_.Exception.Message)" "ERROR" }

    return [PSCustomObject]@{ Erfolg = $erfolg; RebootNoetig = $rebootNoetig }
}

#--------------------------------------------------------------------------------
# 4. FUNKTION: Optimierung durchführen (inkl. Backup)
#--------------------------------------------------------------------------------
function Start-Optimierung {
    Write-Log "=== Optimierung gestartet ===" "INFO"
    $rebootGesamt = $false

    try {
        #-- 4a) Backup-Struktur anlegen ----------------------------------------
        Write-Log "Beginne Sicherung des Ausgangszustands..." "INFO"
        $backup = @{}
        $backup["Zeitstempel"]   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        $backup["Computername"]  = $env:COMPUTERNAME
        $backup["Adapter"]       = @{}

        $adapters = Get-NetAdapter -Physical -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' }

        foreach ($adapter in $adapters) {
            $name = $adapter.Name
            $backup["Adapter"][$name] = @{ Offloads = @{}; IPv6 = $null; NetBIOS = $null }

            Write-Log "Optimiere Adapter: $name" "INFO"

            #-- Offload-Eigenschaften deaktivieren -----------------------------
            foreach ($feature in $Global:OffloadFeatures) {
                try {
                    $adv = Get-NetAdapterAdvancedProperty -Name $name -DisplayName $feature -ErrorAction Stop
                    $backup["Adapter"][$name]["Offloads"][$feature] = $adv.DisplayValue
                    Set-NetAdapterAdvancedProperty -Name $name -DisplayName $feature -DisplayValue 'Disabled' -ErrorAction Stop
                    Write-Log "  Deaktiviert: $feature" "SUCCESS"
                }
                catch { Write-Log "  Konnte '$feature' auf '$name' nicht ändern: $($_.Exception.Message)" "WARN" }
            }

            #-- IPv6 deaktivieren ----------------------------------------------
            try {
                $ipv6Binding = Get-NetAdapterBinding -Name $name -ComponentID 'ms_tcpip6' -ErrorAction Stop
                $backup["Adapter"][$name]["IPv6"] = $ipv6Binding.Enabled
                Disable-NetAdapterBinding -Name $name -ComponentID 'ms_tcpip6' -ErrorAction Stop
                Write-Log "  IPv6 deaktiviert auf $name" "SUCCESS"
            }
            catch { Write-Log "  IPv6 konnte nicht deaktiviert werden: $($_.Exception.Message)" "WARN" }

            #-- NetBIOS über TCP/IP deaktivieren -------------------------------
            try {
                $nic = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ErrorAction Stop |
                       Where-Object { $_.Description -eq $adapter.InterfaceDescription -and $_.IPEnabled }
                if ($nic) {
                    $backup["Adapter"][$name]["NetBIOS"] = $nic.TcpipNetbiosOptions
                    Invoke-CimMethod -InputObject $nic -MethodName SetTcpipNetbios -Arguments @{ TcpipNetbiosOptions = [uint32]2 } -ErrorAction Stop | Out-Null
                    Write-Log "  NetBIOS über TCP/IP deaktiviert" "SUCCESS"
                }
            }
            catch { Write-Log "  NetBIOS konnte nicht deaktiviert werden: $($_.Exception.Message)" "WARN" }

            #-- NIC-Energieverwaltung (robust) ---------------------------------
            $nicResult = Disable-PVSNicPowerManagement -AdapterName $name -BackupStore $backup
            if ($nicResult.RebootNoetig) { $rebootGesamt = $true }
        }

        #-- 4b) Globales TCP Task Offload --------------------------------------
        try {
            $tcpGlobal = Get-NetOffloadGlobalSetting -ErrorAction Stop
            $backup["TCP_TaskOffload"] = $tcpGlobal.TaskOffload
            Set-NetOffloadGlobalSetting -TaskOffload Disabled -ErrorAction Stop
            Write-Log "Globales TCP Task Offload deaktiviert" "SUCCESS"
        }
        catch { Write-Log "TCP Task Offload konnte nicht geändert werden: $($_.Exception.Message)" "WARN" }

        #-- 4c) Energieplan 'Höchstleistung' -----------------------------------
        try {
            $aktiverPlan = (powercfg /getactivescheme) -replace '.*GUID:\s*([a-f0-9\-]+).*','$1'
            $backup["Energieplan"] = $aktiverPlan.Trim()
            powercfg /setactive SCHEME_MIN | Out-Null   # SCHEME_MIN = Höchstleistung
            Write-Log "Energieplan 'Höchstleistung' aktiviert" "SUCCESS"
        }
        catch { Write-Log "Energieplan konnte nicht geändert werden: $($_.Exception.Message)" "WARN" }

        #-- 4d) Ruhezustand deaktivieren ---------------------------------------
        try {
            $backup["Ruhezustand"] = $true   # war zuvor aktiv (Standard)
            powercfg /hibernate off | Out-Null
            Write-Log "Ruhezustand deaktiviert" "SUCCESS"
        }
        catch { Write-Log "Ruhezustand konnte nicht deaktiviert werden: $($_.Exception.Message)" "WARN" }

        #-- 4e) Backup speichern -----------------------------------------------
        $backup | ConvertTo-Json -Depth 8 | Set-Content -Path $Global:BackupDatei -Encoding UTF8
        Write-Log "Sicherung erfolgreich gespeichert unter: $Global:BackupDatei" "SUCCESS"

        if ($rebootGesamt) {
            Write-Log "HINWEIS: Ein NEUSTART ist erforderlich, damit alle NIC-Einstellungen wirksam werden." "WARN"
        }
        Write-Log "=== Optimierung abgeschlossen ===" "SUCCESS"
    }
    catch { Write-Log "Kritischer Fehler während der Optimierung: $($_.Exception.Message)" "ERROR" }
}

#--------------------------------------------------------------------------------
# 5. FUNKTION: Rollback (stellt Ausgangszustand aus Backup wieder her)
#--------------------------------------------------------------------------------
function Start-Rollback {
    Write-Log "=== Rollback gestartet ===" "INFO"

    try {
        if (-not (Test-Path $Global:BackupDatei)) {
            Write-Log "Keine Sicherungsdatei gefunden unter: $Global:BackupDatei" "ERROR"
            return
        }

        $backup = Get-Content -Path $Global:BackupDatei -Raw | ConvertFrom-Json
        Write-Log "Sicherung vom $($backup.Zeitstempel) wird wiederhergestellt..." "INFO"

        #-- 5a) Adaptereinstellungen zurücksetzen ------------------------------
        foreach ($name in $backup.Adapter.PSObject.Properties.Name) {
            $aData = $backup.Adapter.$name
            Write-Log "Stelle Adapter wieder her: $name" "INFO"

            # Offloads
            foreach ($feature in $aData.Offloads.PSObject.Properties.Name) {
                try {
                    $wert = $aData.Offloads.$feature
                    Set-NetAdapterAdvancedProperty -Name $name -DisplayName $feature -DisplayValue $wert -ErrorAction Stop
                    Write-Log "  Wiederhergestellt: $feature = $wert" "SUCCESS"
                }
                catch { Write-Log "  Konnte '$feature' nicht wiederherstellen: $($_.Exception.Message)" "WARN" }
            }

            # IPv6
            if ($aData.IPv6 -eq $true) {
                try {
                    Enable-NetAdapterBinding -Name $name -ComponentID 'ms_tcpip6' -ErrorAction Stop
                    Write-Log "  IPv6 wieder aktiviert" "SUCCESS"
                }
                catch { Write-Log "  IPv6 konnte nicht aktiviert werden: $($_.Exception.Message)" "WARN" }
            }

            # NetBIOS
            if ($null -ne $aData.NetBIOS) {
                try {
                    $adapter = Get-NetAdapter -Name $name -ErrorAction Stop
                    $nic = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration |
                           Where-Object { $_.Description -eq $adapter.InterfaceDescription -and $_.IPEnabled }
                    if ($nic) {
                        Invoke-CimMethod -InputObject $nic -MethodName SetTcpipNetbios -Arguments @{ TcpipNetbiosOptions = [uint32]$aData.NetBIOS } -ErrorAction Stop | Out-Null
                        Write-Log "  NetBIOS-Einstellung wiederhergestellt ($($aData.NetBIOS))" "SUCCESS"
                    }
                }
                catch { Write-Log "  NetBIOS konnte nicht wiederhergestellt werden: $($_.Exception.Message)" "WARN" }
            }

            # NIC-Energieverwaltung via WMI wiederherstellen
            $wmiKey = "NIC_PowerMgmt_$name"
            if ($backup.PSObject.Properties.Name -contains $wmiKey) {
                try {
                    $adapter = Get-NetAdapter -Name $name -ErrorAction Stop
                    $pm = Get-CimInstance -Namespace 'root\wmi' -ClassName 'MSPower_DeviceEnable' -ErrorAction Stop |
                          Where-Object { $_.InstanceName -match [regex]::Escape($adapter.PnPDeviceID) }
                    if ($pm) {
                        Set-CimInstance -InputObject $pm -Property @{ Enable = [bool]$backup.$wmiKey } -ErrorAction Stop | Out-Null
                        Write-Log "  NIC-Energieverwaltung (WMI) wiederhergestellt: $($backup.$wmiKey)" "SUCCESS"
                    }
                }
                catch { Write-Log "  WMI-Energieverwaltung Rollback fehlgeschlagen: $($_.Exception.Message)" "WARN" }
            }
        }

        #-- 5b) IdleInWorkingState (Registry) zurücksetzen ---------------------
        foreach ($prop in $backup.PSObject.Properties) {
            if ($prop.Name -like 'NIC_IdleInWorkingState_*') {
                try {
                    $pfad = $prop.Value.Pfad
                    $wert = $prop.Value.Wert
                    if ($null -eq $wert) {
                        # Wert existierte vorher nicht -> entfernen
                        Remove-ItemProperty -Path $pfad -Name 'IdleInWorkingState' -ErrorAction SilentlyContinue
                        Write-Log "  Registry 'IdleInWorkingState' entfernt (war ursprünglich nicht vorhanden). NEUSTART nötig." "WARN"
                    } else {
                        Set-ItemProperty -Path $pfad -Name 'IdleInWorkingState' -Value $wert -ErrorAction Stop
                        Write-Log "  Registry 'IdleInWorkingState' auf Ausgangswert ($wert) zurückgesetzt. NEUSTART nötig." "WARN"
                    }
                }
                catch { Write-Log "  IdleInWorkingState Rollback fehlgeschlagen: $($_.Exception.Message)" "WARN" }
            }
        }

        #-- 5c) TCP Task Offload -----------------------------------------------
        if ($backup.PSObject.Properties.Name -contains 'TCP_TaskOffload') {
            try {
                Set-NetOffloadGlobalSetting -TaskOffload $backup.TCP_TaskOffload -ErrorAction Stop
                Write-Log "TCP Task Offload wiederhergestellt: $($backup.TCP_TaskOffload)" "SUCCESS"
            }
            catch { Write-Log "TCP Task Offload Rollback fehlgeschlagen: $($_.Exception.Message)" "WARN" }
        }

        #-- 5d) Energieplan ----------------------------------------------------
        if ($backup.PSObject.Properties.Name -contains 'Energieplan') {
            try {
                powercfg /setactive $backup.Energieplan | Out-Null
                Write-Log "Energieplan wiederhergestellt: $($backup.Energieplan)" "SUCCESS"
            }
            catch { Write-Log "Energieplan Rollback fehlgeschlagen: $($_.Exception.Message)" "WARN" }
        }

        #-- 5e) Ruhezustand ----------------------------------------------------
        if ($backup.Ruhezustand -eq $true) {
            try {
                powercfg /hibernate on | Out-Null
                Write-Log "Ruhezustand wieder aktiviert" "SUCCESS"
            }
            catch { Write-Log "Ruhezustand Rollback fehlgeschlagen: $($_.Exception.Message)" "WARN" }
        }

        Write-Log "=== Rollback abgeschlossen ===" "SUCCESS"
    }
    catch { Write-Log "Kritischer Fehler während des Rollbacks: $($_.Exception.Message)" "ERROR" }
}

#================================================================================
# 6. GRAFISCHE OBERFLÄCHE (helle, freundliche GUI mit Live-Changelog)
#================================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Farbschema: hell & freundlich ---
$farbeHintergrund = [System.Drawing.Color]::FromArgb(245,248,250)  # sehr helles Blaugrau
$farbeAkzent      = [System.Drawing.Color]::FromArgb(52,152,219)   # freundliches Blau
$farbeAkzentGruen = [System.Drawing.Color]::FromArgb(46,204,113)   # frisches Grün
$farbeAkzentOrange= [System.Drawing.Color]::FromArgb(230,126,34)   # warmes Orange
$farbeText        = [System.Drawing.Color]::FromArgb(44,62,80)     # dunkles Blaugrau

# --- Hauptfenster ---
$form = New-Object System.Windows.Forms.Form
$form.Text            = "PVS Target Device Optimierung – v4.8"
$form.Size            = New-Object System.Drawing.Size(720, 560)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $farbeHintergrund
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
$Global:LogForm       = $form

# --- Kopfbereich / Titel ---
$lblTitel = New-Object System.Windows.Forms.Label
$lblTitel.Text      = "Netzwerk- & Windows-Optimierung"
$lblTitel.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 15, [System.Drawing.FontStyle]::Bold)
$lblTitel.ForeColor = $farbeText
$lblTitel.Location  = New-Object System.Drawing.Point(20, 15)
$lblTitel.AutoSize  = $true
$form.Controls.Add($lblTitel)

$lblUntertitel = New-Object System.Windows.Forms.Label
$lblUntertitel.Text      = "Optimierung durchführen oder auf den gesicherten Zustand zurücksetzen."
$lblUntertitel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$lblUntertitel.ForeColor = [System.Drawing.Color]::FromArgb(127,140,141)
$lblUntertitel.Location  = New-Object System.Drawing.Point(22, 48)
$lblUntertitel.AutoSize  = $true
$form.Controls.Add($lblUntertitel)

# --- Button: Optimierung starten ---
$btnOptimieren = New-Object System.Windows.Forms.Button
$btnOptimieren.Text      = "▶  Optimierung starten"
$btnOptimieren.Size      = New-Object System.Drawing.Size(200, 42)
$btnOptimieren.Location  = New-Object System.Drawing.Point(20, 80)
$btnOptimieren.BackColor = $farbeAkzentGruen
$btnOptimieren.ForeColor = [System.Drawing.Color]::White
$btnOptimieren.FlatStyle = "Flat"
$btnOptimieren.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Bold)
$btnOptimieren.FlatAppearance.BorderSize = 0
$btnOptimieren.Cursor    = "Hand"
$form.Controls.Add($btnOptimieren)

# --- Button: Rollback ---
$btnRollback = New-Object System.Windows.Forms.Button
$btnRollback.Text      = "↺  Rollback (Wiederherstellen)"
$btnRollback.Size      = New-Object System.Drawing.Size(220, 42)
$btnRollback.Location  = New-Object System.Drawing.Point(235, 80)
$btnRollback.BackColor = $farbeAkzentOrange
$btnRollback.ForeColor = [System.Drawing.Color]::White
$btnRollback.FlatStyle = "Flat"
$btnRollback.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Bold)
$btnRollback.FlatAppearance.BorderSize = 0
$btnRollback.Cursor    = "Hand"
$form.Controls.Add($btnRollback)

# --- Button: Log leeren ---
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text      = "Log leeren"
$btnClear.Size      = New-Object System.Drawing.Size(110, 42)
$btnClear.Location  = New-Object System.Drawing.Point(470, 80)
$btnClear.BackColor = [System.Drawing.Color]::White
$btnClear.ForeColor = $farbeText
$btnClear.FlatStyle = "Flat"
$btnClear.Cursor    = "Hand"
$btnClear.FlatAppearance.BorderColor = $farbeAkzent
$form.Controls.Add($btnClear)

# --- Label: Live-Changelog ---
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text      = "Live-Changelog:"
$lblLog.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Bold)
$lblLog.ForeColor = $farbeText
$lblLog.Location  = New-Object System.Drawing.Point(20, 140)
$lblLog.AutoSize  = $true
$form.Controls.Add($lblLog)

# --- RichTextBox: Live-Log-Ausgabe (heller Hintergrund) ---
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location   = New-Object System.Drawing.Point(20, 165)
$logBox.Size       = New-Object System.Drawing.Size(660, 300)
$logBox.BackColor  = [System.Drawing.Color]::White
$logBox.Font       = New-Object System.Drawing.Font("Consolas", 9)
$logBox.ReadOnly   = $true
$logBox.BorderStyle= "FixedSingle"
$logBox.ScrollBars = "Vertical"
$form.Controls.Add($logBox)
$Global:GuiLogBox = $logBox    # global verfügbar machen für Write-Log

# --- Statusleiste unten ---
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text      = "Bereit."
$lblStatus.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(127,140,141)
$lblStatus.Location  = New-Object System.Drawing.Point(20, 475)
$lblStatus.AutoSize  = $true
$form.Controls.Add($lblStatus)

#--------------------------------------------------------------------------------
# 7. EVENT-HANDLER
#--------------------------------------------------------------------------------
$btnOptimieren.Add_Click({
    $btnOptimieren.Enabled = $false; $btnRollback.Enabled = $false
    $lblStatus.Text = "Optimierung läuft..."
    Start-Optimierung
    $lblStatus.Text = "Optimierung abgeschlossen."
    $btnOptimieren.Enabled = $true; $btnRollback.Enabled = $true
})

$btnRollback.Add_Click({
    $antwort = [System.Windows.Forms.MessageBox]::Show(
        "Möchten Sie wirklich den gesicherten Ausgangszustand wiederherstellen?",
        "Rollback bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)

    if ($antwort -eq [System.Windows.Forms.DialogResult]::Yes) {
        $btnOptimieren.Enabled = $false; $btnRollback.Enabled = $false
        $lblStatus.Text = "Rollback läuft..."
        Start-Rollback
        $lblStatus.Text = "Rollback abgeschlossen."
        $btnOptimieren.Enabled = $true; $btnRollback.Enabled = $true
    }
})

$btnClear.Add_Click({
    $logBox.Clear()
    $lblStatus.Text = "Log-Anzeige geleert."
})

# --- Startmeldung LIVE anzeigen ---
$form.Add_Shown({
    Write-Log "GUI gestartet. Bereit für Optimierung oder Rollback." "INFO"
})

# --- Fenster anzeigen ---
[void]$form.ShowDialog()
