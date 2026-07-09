<#
.SYNOPSIS
    Exportiert Application Error/Hang/Popup, Service Control Manager, Windows Resource Exhaustion
    Eventlog-Einträge aller Server einer OU als HTML.

.DESCRIPTION
    Sammelt von allen Servern in einer angegebenen OU die Eventlog-Einträge mit den Quellen
    "Application Error", "Application Hang", "Application Popup", "Service Control Manager" und
    "Windows Resource Exhaustion" und exportiert sie
    übersichtlich als HTML-Datei (Uhrzeit, Server, Fehler). Bei Angabe von -Interval läuft
    das Skript durchgehend und aktualisiert die HTML im angegebenen Minuten-Takt; die HTML-Seite
    aktualisiert sich dann automatisch im Browser. Zusätzlich werden freier Speicherplatz D:,
    Größe der mcsdif.vhdx und freier Arbeitsspeicher als Ampelanzeige (rot/gelb/grün) dargestellt.
    Fehlermeldungen enthalten
    farbige Hervorhebungen: EXE-Pfade (grün), DLL-Pfade (rot), Ausnahmecodes (gelb) und
    "Nicht genügend virtueller Speicher" (rot). Zeigt aktive Citrix Sessions pro Server,
    CPU-Auslastung, Auslagerungsdatei-Belegung und FSLogix-Dienststatus (farbig).
    TOP 10 der größten Speicherfresser (WorkingSet) sowie eine TOP 10
    des Session-RAM (nach Benutzer gruppiert) aller Server an.
    Klick auf die Prozess-Anzahl öffnet die Prozessliste in einem In-Page-Modal (kein Popup!).

.PARAMETER SearchBase
    DistinguishedName der OU, in der nach Servern gesucht wird.

.PARAMETER OutputPath
    Pfad zur Ausgabe-HTML-Datei. Standard: Skriptverzeichnis\SystemStatusReport.html

.PARAMETER DaysBack
    Nur Einträge der letzten X Tage berücksichtigen. Standard: 7

.PARAMETER Interval
    Intervall in Minuten für automatische Wiederholung (0 = einmalig).

.EXAMPLE
    .\System-Status-Application-Fehler-Report.ps1 -SearchBase "OU=Servers,DC=domain,DC=local" -DaysBack 14

.NOTES
    Version  : 1.10
    Autor    : Rocco Ammon
    Änderung : CPU-Auslastung, Auslagerungsdatei (Pagefile), FSLogix-Dienststatus
               in die Ampel-Tabelle aufgenommen.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$SearchBase,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [int]$DaysBack = 7,

    [Parameter(Mandatory = $false)]
    [int]$Interval = 0
)

# =========================================================================
# Region: Konfiguration / Variablen (immer am Anfang)
# =========================================================================
$ErrorActionPreference = 'Stop'
$ScriptDir             = if ($PSScriptRoot) { $PSScriptRoot } else { '.' }
if (-not $OutputPath) { $OutputPath = Join-Path -Path $ScriptDir -ChildPath "SystemStatusReport.html" }
$firstRun              = $true

# --- Logging-Konfiguration (Standard: C:\ScriptLog) ---
$LogDir  = 'C:\ScriptLog'
$LogFile = Join-Path -Path $LogDir -ChildPath 'System-Status-Application-Fehler-Report.log'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')] [string]$Level = 'INFO'
    )
    try {
        if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
        $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        Add-Content -Path $LogFile -Value $line -Encoding UTF8
    }
    catch {
        Write-Warning "Logging fehlgeschlagen: $($_.Exception.Message)"
    }
}

Write-Log "Skriptstart. SearchBase='$SearchBase', DaysBack=$DaysBack, Interval=$Interval"

do {
try {
    $StartTime = (Get-Date).AddDays(-$DaysBack)

    # Pro LogName die zu suchenden Provider (Quellen)
    $LogQueries = @(
        @{ LogName = 'Application'; Providers = @('Application Error', 'Application Hang') }
        @{ LogName = 'System';      Providers = @('Application Popup') }
        @{ LogName = 'System';      Providers = @('Service Control Manager'); EventIDs = @(7031, 7032, 7034, 7000, 7009) }
        @{ LogName = 'System';      Providers = @('Windows Resource Exhaustion'); EventIDs = @(2001, 2004, 2005, 2006, 2020) }
    )

    # =====================================================================
    # Region: AD abfragen
    # =====================================================================
    Write-Host "Suche Computer in OU: $SearchBase ..." -ForegroundColor Cyan
    Write-Log  "Suche Computer in OU: $SearchBase"
    try {
        $Computers = Get-ADComputer -Filter 'enabled -eq "true"' -SearchBase $SearchBase `
            -Properties Name -ErrorAction Stop | Select-Object -ExpandProperty Name
    }
    catch {
        Write-Error "Fehler bei AD-Abfrage: $_"
        Write-Log   "Fehler bei AD-Abfrage: $($_.Exception.Message)" 'ERROR'
        break
    }

    if (-not $Computers) {
        Write-Warning "Keine Computer in der angegebenen OU gefunden."
        Write-Log     "Keine Computer in der angegebenen OU gefunden." 'WARN'
        break
    }

    Write-Host "$($Computers.Count) Server gefunden, starte parallele Abfragen (ThrottleLimit=15) ..." -ForegroundColor Cyan
    Write-Log  "$($Computers.Count) Server gefunden."

    # =====================================================================
    # Region: Parallele Server-Abfragen via Invoke-Command
    # =====================================================================
    $LogQueriesJson = $LogQueries | ConvertTo-Json -Compress

    $ServerScriptBlock = {
        param($StartTime, $LogQueriesJson)

        $LogQueries = $LogQueriesJson | ConvertFrom-Json
        $events = @()

        $reachable = Test-Connection -ComputerName $env:COMPUTERNAME -Count 1 -Quiet -ErrorAction SilentlyContinue
        if (-not $reachable) {
            return [PSCustomObject]@{
                Computer         = $env:COMPUTERNAME
                Reachable        = $false
                Events           = $events
                FreeGB           = $null; TotalGB = $null; FreePct = $null
                VhdxExists       = $false; VhdxSizeGB = $null
                MemFreeMB        = $null; MemTotalMB = $null; MemFreePct = $null
                Sessions         = 0
                TopProcs         = @()
                SessionRAM       = @()
                CpuPct           = $null
                PageFileMB       = $null; PageFileTotalMB = $null; PageFileUsagePct = $null
                FslogixServices  = @()
            }
        }

        foreach ($Query in $LogQueries) {
            $LogName  = $Query.LogName
            $Provider = $Query.Providers

            $XPathFilter = "*[System[{0}]]" -f (
                ($Provider | ForEach-Object { "Provider[@Name='$_']" }) -join ' or '
            )

            try {
                $rawEvents = Get-WinEvent -LogName $LogName -FilterXPath $XPathFilter `
                    -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

                if (-not $rawEvents) { continue }

                $filtered = $rawEvents | Where-Object {
                    $_.TimeCreated -ge $StartTime -and
                    (-not $Query.EventIDs -or ($_.Id -in $Query.EventIDs))
                }

                foreach ($Event in $filtered) {
                    $msg = ($Event.Message -replace '\r?\n', ' ')
                    if ($msg -match 'Windows Update|wuauserv') { continue }
                    $events += [PSCustomObject]@{
                        Time     = $Event.TimeCreated
                        Server   = $env:COMPUTERNAME
                        Level    = switch ($Event.Level) {
                            1 { 'Critical' }
                            2 { 'Error' }
                            3 { 'Warning' }
                            4 { 'Information' }
                            default { "Level $($Event.Level)" }
                        }
                        Provider = $Event.ProviderName
                        Message  = $msg
                    }
                }
            }
            catch {}
        }

        $drive = $null; $mem = $null; $sessionCount = 0
        try { $drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='D:'" -ErrorAction Stop } catch {}
        try { $mem   = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop } catch {}

        try {
            $explorer = Get-CimInstance Win32_Process -Filter "Name = 'explorer.exe'" -ErrorAction SilentlyContinue
            if ($explorer) { $sessionCount = @($explorer).Count }
        } catch {}
        if ($sessionCount -eq 0) {
            $lines = query user 2>$null
            if ($lines -and $lines.Count -gt 1) { $sessionCount = $lines.Count - 1 }
        }

        $f = Get-Item 'D:\mcsdif.vhdx' -ErrorAction SilentlyContinue
        $vhdxExists = if ($f) { $true } else { $false }
        $vhdxSizeGB = if ($f) { [math]::Round($f.Length / 1GB, 2) } else { $null }

        $top = @()
        try {
            $top = Get-Process | Sort-Object WorkingSet64 -Descending |
                Select-Object -First 5 Name, Id, @{N='WS_MB';E={[math]::Round($_.WorkingSet64 / 1MB)}}
        } catch {}

        $sessionRAM = @()
        try {
            $sessionRAM = Get-Process -IncludeUserName -ErrorAction SilentlyContinue |
                Where-Object { $_.WorkingSet64 -gt 0 } |
                Group-Object UserName |
                Where-Object { $_.Name -and $_.Name -notlike '*SYSTEM' -and $_.Name -notlike '*LOCAL SERVICE' -and $_.Name -notlike '*NETWORK SERVICE' } |
                ForEach-Object {
                    $userName = $_.Name
                    $procList = $_.Group | Sort-Object WorkingSet64 -Descending |
                        Select-Object Name, @{N='WS_MB';E={[math]::Round($_.WorkingSet64 / 1MB)}}
                    [PSCustomObject]@{
                        User          = $userName
                        SessionRAM_MB = [math]::Round(($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)
                        ProcessCount  = $_.Count
                        ProcessList   = $procList | ConvertTo-Json -Compress
                    }
                } | Sort-Object SessionRAM_MB -Descending | Select-Object -First 10
        } catch {}

        $cpuPct = $null
        try {
            $cpu = Get-CimInstance Win32_Processor -ErrorAction Stop
            if ($cpu) { $cpuPct = [math]::Round(($cpu | Measure-Object -Property LoadPercentage -Average).Average, 1) }
        } catch {}

        $pfAllocatedMB = $null; $pfCurrentMB = $null; $pfUsagePct = $null
        try {
            $pageFile = Get-CimInstance Win32_PageFileUsage -ErrorAction Stop
            if ($pageFile) {
                $pfAllocatedMB = ($pageFile | Measure-Object -Property AllocatedBaseSize -Sum).Sum
                $pfCurrentMB   = ($pageFile | Measure-Object -Property CurrentUsage -Sum).Sum
                $pfUsagePct    = if ($pfAllocatedMB -gt 0) { [math]::Round(($pfCurrentMB / $pfAllocatedMB) * 100, 1) } else { $null }
            }
        } catch {}

        $fslogixServices = @()
        try {
            $fslogixServices = Get-Service |
                Where-Object { $_.DisplayName -like '*FSLogix*' -or $_.Name -like '*fsl*' -or $_.Name -like '*frx*' } |
                Select-Object Name, DisplayName, Status, StartType -ErrorAction Stop
        } catch {}

        if ($drive) {
            $freeGB  = [math]::Round($drive.FreeSpace / 1GB, 2)
            $totalGB = [math]::Round($drive.Size / 1GB, 2)
            $freePct = if ($drive.Size -gt 0) { [math]::Round(($drive.FreeSpace / $drive.Size) * 100, 1) } else { 0 }
        } else { $freeGB = $null; $totalGB = $null; $freePct = $null }

        [PSCustomObject]@{
            Computer         = $env:COMPUTERNAME
            Reachable        = $true
            Events           = $events
            FreeGB           = $freeGB
            TotalGB          = $totalGB
            FreePct          = $freePct
            VhdxExists       = $vhdxExists
            VhdxSizeGB       = $vhdxSizeGB
            MemFreeMB        = if ($mem) { [math]::Round($mem.FreePhysicalMemory / 1024, 1) } else { $null }
            MemTotalMB       = if ($mem) { [math]::Round($mem.TotalVisibleMemorySize / 1024, 1) } else { $null }
            MemFreePct       = if ($mem -and $mem.TotalVisibleMemorySize -gt 0) { [math]::Round(($mem.FreePhysicalMemory / $mem.TotalVisibleMemorySize) * 100, 1) } else { $null }
            Sessions         = $sessionCount
            TopProcs         = $top
            SessionRAM       = $sessionRAM
            CpuPct           = $cpuPct
            PageFileMB       = $pfCurrentMB
            PageFileTotalMB  = $pfAllocatedMB
            PageFileUsagePct = $pfUsagePct
            FslogixServices  = $fslogixServices
        }
    }

    $allResults = Invoke-Command -ComputerName $Computers -ScriptBlock $ServerScriptBlock `
        -ArgumentList $StartTime, $LogQueriesJson -ThrottleLimit 15 -ErrorAction SilentlyContinue

    # =====================================================================
    # Region: Ergebnisse einsammeln
    # =====================================================================
    $Results           = [System.Collections.ArrayList]::new()
    $DriveResults      = [System.Collections.ArrayList]::new()
    $TopProcesses      = [System.Collections.ArrayList]::new()
    $SessionRAMResults = [System.Collections.ArrayList]::new()
    $totalEvents       = 0

    foreach ($result in $allResults) {
        if (-not $result) { continue }
        Write-Host "  $($result.Computer) fertig" -ForegroundColor Green

        if (-not $result.Reachable) {
            Write-Warning "$($result.Computer) | nicht erreichbar"
            Write-Log     "$($result.Computer) nicht erreichbar" 'WARN'
            continue
        }

        if ($result.Events -and $result.Events.Count -gt 0) {
            foreach ($evt in $result.Events) {
                [void]$Results.Add($evt)
                $totalEvents++
            }
        }

        [void]$DriveResults.Add([PSCustomObject]@{
            Server           = $result.Computer
            FreeGB           = $result.FreeGB
            TotalGB          = $result.TotalGB
            FreePct          = $result.FreePct
            VhdxExists       = $result.VhdxExists
            VhdxSizeGB       = $result.VhdxSizeGB
            MemFreeMB        = $result.MemFreeMB
            MemTotalMB       = $result.MemTotalMB
            MemFreePct       = $result.MemFreePct
            Sessions         = $result.Sessions
            CpuPct           = $result.CpuPct
            PageFileMB       = $result.PageFileMB
            PageFileTotalMB  = $result.PageFileTotalMB
            PageFileUsagePct = $result.PageFileUsagePct
            FslogixServices  = $result.FslogixServices
        })

        if ($result.TopProcs -and $result.TopProcs.Count -gt 0) {
            foreach ($p in $result.TopProcs) {
                [void]$TopProcesses.Add([PSCustomObject]@{ Server = $result.Computer; Name = $p.Name; PID = $p.Id; WS_MB = $p.WS_MB })
            }
        }

        if ($result.SessionRAM -and $result.SessionRAM.Count -gt 0) {
            foreach ($s in $result.SessionRAM) {
                [void]$SessionRAMResults.Add([PSCustomObject]@{
                    Server        = $result.Computer
                    User          = $s.User
                    SessionRAM_MB = $s.SessionRAM_MB
                    ProcessCount  = $s.ProcessCount
                    ProcessList   = $s.ProcessList
                })
            }
        }
    }

    $processedServers = @($allResults | Where-Object { $_ -and $_.Computer } | ForEach-Object { $_.Computer })
    foreach ($comp in $Computers) {
        if ($comp -notin $processedServers) {
            Write-Warning "$comp | Keine Verbindung (WinRM nicht erreichbar)"
            Write-Log     "$comp keine WinRM-Verbindung" 'WARN'
        }
    }

    $Top10         = $TopProcesses | Sort-Object WS_MB -Descending | Select-Object -First 10
    $TopSessionRAM = $SessionRAMResults | Sort-Object SessionRAM_MB -Descending | Select-Object -First 10

    # =====================================================================
    # Region: Neue-Einträge-Erkennung (Piepton)
    # =====================================================================
    $CacheFile = Join-Path -Path $ScriptDir -ChildPath "SystemStatusReport_Cache.json"
    $lastCount = 0
    if (Test-Path $CacheFile) {
        try {
            $cache = Get-Content $CacheFile -Raw | ConvertFrom-Json
            $lastCount = $cache.LastEventCount
        } catch {}
    }
    if ($Results.Count -gt $lastCount) {
        [System.Console]::Beep(800, 400)
        Start-Sleep -Milliseconds 200
        [System.Console]::Beep(1000, 400)
    }
    @{ LastEventCount = $Results.Count; Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') } |
        ConvertTo-Json | Set-Content $CacheFile -Encoding UTF8

    Write-Host "Erstelle HTML-Report ..." -ForegroundColor Cyan
    Write-Log  "Erstelle HTML-Report. Event-Einträge: $($Results.Count)"

    # =====================================================================
    # Region: HTML erstellen
    # =====================================================================
    $HasErrors = $Results.Count -gt 0
    $HasDrives = $DriveResults.Count -gt 0

    if (-not $HasErrors -and -not $HasDrives) {
        Write-Warning "Keine Daten gefunden – Report wird nicht erstellt."
        Write-Log     "Keine Daten gefunden." 'WARN'
        break
    }

    # ------------------- Style-Block -------------------
    $StyleBlock = [System.Text.StringBuilder]::new()
    [void]$StyleBlock.AppendLine('<style>')
    [void]$StyleBlock.AppendLine('    body { font-family: "Segoe UI", Arial, sans-serif; margin: 20px; }')
    [void]$StyleBlock.AppendLine('    h1 { color: #333; }')
    [void]$StyleBlock.AppendLine('    .summary { margin-bottom: 15px; color: #666; }')
    [void]$StyleBlock.AppendLine('    .filterBar { display: flex; gap: 10px; margin-bottom: 10px; flex-wrap: wrap; }')
    [void]$StyleBlock.AppendLine('    .filterBar label { font-weight: 600; margin-right: 4px; }')
    [void]$StyleBlock.AppendLine('    .filterBar select { padding: 6px; border: 1px solid #ccc; border-radius: 4px; min-width: 180px; }')
    [void]$StyleBlock.AppendLine('    #searchBox { width: 300px; padding: 6px; border: 1px solid #ccc; border-radius: 4px; }')
    [void]$StyleBlock.AppendLine('    table { border-collapse: collapse; width: 100%; margin-bottom: 30px; }')
    [void]$StyleBlock.AppendLine('    th { background-color: #007acc; color: white; padding: 8px; text-align: left; cursor: pointer; user-select: none; }')
    [void]$StyleBlock.AppendLine('    th::after { content: " \25B4\25BE"; font-size: 16px; opacity: 0.3; }')
    [void]$StyleBlock.AppendLine('    th.asc::after { content: " \25B4"; opacity: 1; }')
    [void]$StyleBlock.AppendLine('    th.desc::after { content: " \25BE"; opacity: 1; }')
    [void]$StyleBlock.AppendLine('    td { padding: 6px 8px; border-bottom: 1px solid #ddd; vertical-align: top; }')
    [void]$StyleBlock.AppendLine('    tr:nth-child(even) { background-color: #f5f5f5; }')
    [void]$StyleBlock.AppendLine('    tr:hover { background-color: #e0f0ff; }')
    [void]$StyleBlock.AppendLine('    .level-Error { border-left: 4px solid #e74c3c; }')
    [void]$StyleBlock.AppendLine('    .level-Warning { border-left: 4px solid #f39c12; }')
    [void]$StyleBlock.AppendLine('    .level-Information { border-left: 4px solid #27ae60; }')
    [void]$StyleBlock.AppendLine('    .level-Offline { border-left: 4px solid #95a5a6; color: #999; }')
    [void]$StyleBlock.AppendLine('    td.msg { max-width: 600px; word-wrap: break-word; }')
    [void]$StyleBlock.AppendLine('    .hidden { display: none; }')
    [void]$StyleBlock.AppendLine('    .ampel { display: inline-block; padding: 3px 12px; border-radius: 4px; font-weight: 600; font-size: 13px; min-width: 110px; text-align: center; }')
    [void]$StyleBlock.AppendLine('    .ampel-green { background-color: #27ae60; color: white; }')
    [void]$StyleBlock.AppendLine('    .ampel-yellow { background-color: #f39c12; color: white; }')
    [void]$StyleBlock.AppendLine('    .ampel-red { background-color: #e74c3c; color: white; }')
    [void]$StyleBlock.AppendLine('    .ampel-gray { background-color: #95a5a6; color: white; }')
    [void]$StyleBlock.AppendLine('    td.lv-Error { background-color: #fce4e4; }')
    [void]$StyleBlock.AppendLine('    td.lv-Warning { background-color: #fef9e7; }')
    [void]$StyleBlock.AppendLine('    td.lv-Information { background-color: #e8f4fd; }')
    [void]$StyleBlock.AppendLine('    td.msg-oom { background-color: #f8d7da; font-weight: 600; }')
    [void]$StyleBlock.AppendLine('    span.exe { background-color: #d4edda; color: #000; font-weight: 700; padding: 1px 4px; border-radius: 3px; }')
    [void]$StyleBlock.AppendLine('    span.dll { background-color: #f8d7da; color: #000; font-weight: 700; padding: 1px 4px; border-radius: 3px; }')
    [void]$StyleBlock.AppendLine('    span.excode { background-color: #fff3cd; color: #000; font-weight: 700; padding: 1px 4px; border-radius: 3px; }')
    [void]$StyleBlock.AppendLine('    #topTable td:first-child, #sessionRamTable td:first-child { font-weight: 700; color: #007acc; }')
    [void]$StyleBlock.AppendLine('    #topTable td:nth-child(5), #sessionRamTable td:nth-child(5) { font-weight: 700; text-align: right; }')
    [void]$StyleBlock.AppendLine('    .sessions { text-align: center; font-weight: 600; }')
    [void]$StyleBlock.AppendLine('    td.sessions-highlight { background-color: #e8f4fd; font-weight: 600; }')
    # --- NEU: Modal-CSS (fehlte bisher komplett!) ---
    [void]$StyleBlock.AppendLine('    .modal-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000; display: flex; align-items: center; justify-content: center; }')
    [void]$StyleBlock.AppendLine('    .modal-box { background: #fff; border-radius: 8px; padding: 20px 24px; max-width: 700px; width: 90%; max-height: 80vh; overflow-y: auto; box-shadow: 0 10px 40px rgba(0,0,0,0.3); position: relative; }')
    [void]$StyleBlock.AppendLine('    .modal-box h2 { margin-top: 0; color: #333; }')
    [void]$StyleBlock.AppendLine('    .modal-box table { margin-bottom: 0; }')
    [void]$StyleBlock.AppendLine('    .modal-box td:last-child { font-weight: 700; text-align: right; }')
    [void]$StyleBlock.AppendLine('    .modal-box tr.sum td { font-weight: 700; background-color: #e8f4fd; }')
    [void]$StyleBlock.AppendLine('    .modal-close { position: absolute; top: 10px; right: 16px; font-size: 28px; font-weight: 700; color: #888; cursor: pointer; line-height: 1; }')
    [void]$StyleBlock.AppendLine('    .modal-close:hover { color: #e74c3c; }')
    [void]$StyleBlock.AppendLine('</style>')

    # ------------------- Script-Block -------------------
    $ScriptBlock = [System.Text.StringBuilder]::new()
    [void]$ScriptBlock.AppendLine('<script>')
    [void]$ScriptBlock.AppendLine('function makeSortable(table) {')
    [void]$ScriptBlock.AppendLine('  if (!table) return;')
    [void]$ScriptBlock.AppendLine('  var rows = Array.prototype.filter.call(table.querySelectorAll("tr"), function (r) { return r.querySelector("td"); });')
    [void]$ScriptBlock.AppendLine('  table.querySelectorAll("th").forEach(function (th, col) {')
    [void]$ScriptBlock.AppendLine('    th.addEventListener("click", function () {')
    [void]$ScriptBlock.AppendLine('      var isAsc = th.classList.contains("asc");')
    [void]$ScriptBlock.AppendLine('      table.querySelectorAll("th").forEach(function (h) { h.classList.remove("asc", "desc"); });')
    [void]$ScriptBlock.AppendLine('      th.classList.add(isAsc ? "desc" : "asc");')
    [void]$ScriptBlock.AppendLine('      rows.sort(function (a, b) {')
    [void]$ScriptBlock.AppendLine('        var aVal = (a.cells[col] ? a.cells[col].textContent : "").trim();')
    [void]$ScriptBlock.AppendLine('        var bVal = (b.cells[col] ? b.cells[col].textContent : "").trim();')
    [void]$ScriptBlock.AppendLine('        return isAsc ? bVal.localeCompare(aVal) : aVal.localeCompare(bVal);')
    [void]$ScriptBlock.AppendLine('      });')
    [void]$ScriptBlock.AppendLine('      rows.forEach(function (r) { table.appendChild(r); });')
    [void]$ScriptBlock.AppendLine('    });')
    [void]$ScriptBlock.AppendLine('  });')
    [void]$ScriptBlock.AppendLine('}')
    [void]$ScriptBlock.AppendLine('document.addEventListener("DOMContentLoaded", function () {')
    [void]$ScriptBlock.AppendLine('  makeSortable(document.getElementById("ampelTable"));')
[void]$ScriptBlock.AppendLine('  // showSessionProcs + Modal-Handler (VOR early return, immer definiert)')
[void]$ScriptBlock.AppendLine('  window.showSessionProcs = function (idx) {')
[void]$ScriptBlock.AppendLine('    try {')
[void]$ScriptBlock.AppendLine('      var jsonEl = document.getElementById("sessionProcessData");')
[void]$ScriptBlock.AppendLine('      var modal  = document.getElementById("sessionModal");')
[void]$ScriptBlock.AppendLine('      if (!jsonEl || !modal) return;')
[void]$ScriptBlock.AppendLine('      var allData = JSON.parse(jsonEl.textContent);')
[void]$ScriptBlock.AppendLine('      var data = allData[idx];')
[void]$ScriptBlock.AppendLine('      if (!data) return;')
[void]$ScriptBlock.AppendLine('      var procs = JSON.parse(data.ProcessList);')
[void]$ScriptBlock.AppendLine('      if (!procs) return;')
[void]$ScriptBlock.AppendLine('      if (!Array.isArray(procs)) { procs = [procs]; }')
[void]$ScriptBlock.AppendLine('      document.getElementById("modalUser").textContent   = data.User;')
[void]$ScriptBlock.AppendLine('      document.getElementById("modalServer").textContent = data.Server;')
[void]$ScriptBlock.AppendLine('      var body = document.getElementById("modalTableBody");')
[void]$ScriptBlock.AppendLine('      body.innerHTML = "";')
[void]$ScriptBlock.AppendLine('      var total = 0;')
[void]$ScriptBlock.AppendLine('      for (var j = 0; j < procs.length; j++) {')
[void]$ScriptBlock.AppendLine('        var tr = document.createElement("tr");')
[void]$ScriptBlock.AppendLine('        tr.innerHTML = "<td>" + (j + 1) + "</td><td>" + procs[j].Name + "</td><td>" + procs[j].WS_MB + " MB</td>";')
[void]$ScriptBlock.AppendLine('        body.appendChild(tr);')
[void]$ScriptBlock.AppendLine('        total += procs[j].WS_MB;')
[void]$ScriptBlock.AppendLine('      }')
[void]$ScriptBlock.AppendLine('      var sum = document.createElement("tr");')
[void]$ScriptBlock.AppendLine('      sum.className = "sum";')
[void]$ScriptBlock.AppendLine('      sum.innerHTML = "<td></td><td>Gesamt</td><td>" + total + " MB</td>";')
[void]$ScriptBlock.AppendLine('      body.appendChild(sum);')
[void]$ScriptBlock.AppendLine('      modal.style.display = "flex";')
[void]$ScriptBlock.AppendLine('    } catch(e) {}')
[void]$ScriptBlock.AppendLine('  };')
[void]$ScriptBlock.AppendLine('  var modal = document.getElementById("sessionModal");')
[void]$ScriptBlock.AppendLine('  if (modal) {')
[void]$ScriptBlock.AppendLine('    var mc = document.getElementById("modalClose");')
[void]$ScriptBlock.AppendLine('    if (mc) mc.addEventListener("click", function () { modal.style.display = "none"; });')
[void]$ScriptBlock.AppendLine('    modal.addEventListener("click", function (e) { if (e.target === modal) modal.style.display = "none"; });')
[void]$ScriptBlock.AppendLine('    document.addEventListener("keydown", function (e) { if (e.key === "Escape") modal.style.display = "none"; });')
[void]$ScriptBlock.AppendLine('  }')
[void]$ScriptBlock.AppendLine('  var errTable = document.getElementById("errorTable");')
    [void]$ScriptBlock.AppendLine('  if (!errTable) return;')
    [void]$ScriptBlock.AppendLine('  var rows = Array.prototype.filter.call(errTable.querySelectorAll("tr"), function (r) { return r.querySelector("td"); });')
    [void]$ScriptBlock.AppendLine('  var searchBox = document.getElementById("searchBox");')
    [void]$ScriptBlock.AppendLine('  function getUnique(col) { var set = {}; rows.forEach(function (r) { var val = (r.cells[col] ? r.cells[col].textContent : "").trim(); if (val) set[val] = true; }); return Object.keys(set).sort(); }')
    [void]$ScriptBlock.AppendLine('  function populateFilter(id, col) { var sel = document.getElementById(id); if (!sel) return; getUnique(col).forEach(function (v) { var opt = document.createElement("option"); opt.value = v; opt.textContent = v; sel.appendChild(opt); }); sel.addEventListener("change", applyFilters); }')
    [void]$ScriptBlock.AppendLine('  populateFilter("filterServer", 1);')
    [void]$ScriptBlock.AppendLine('  populateFilter("filterQuelle", 3);')
    [void]$ScriptBlock.AppendLine('  errTable.querySelectorAll("th").forEach(function (th, col) {')
    [void]$ScriptBlock.AppendLine('    th.addEventListener("click", function () {')
    [void]$ScriptBlock.AppendLine('      var isAsc = th.classList.contains("asc");')
    [void]$ScriptBlock.AppendLine('      errTable.querySelectorAll("th").forEach(function (h) { h.classList.remove("asc", "desc"); });')
    [void]$ScriptBlock.AppendLine('      th.classList.add(isAsc ? "desc" : "asc");')
    [void]$ScriptBlock.AppendLine('      rows.sort(function (a, b) {')
    [void]$ScriptBlock.AppendLine('        var aVal = (a.cells[col] ? a.cells[col].textContent : "").trim();')
    [void]$ScriptBlock.AppendLine('        var bVal = (b.cells[col] ? b.cells[col].textContent : "").trim();')
    [void]$ScriptBlock.AppendLine('        return isAsc ? bVal.localeCompare(aVal) : aVal.localeCompare(bVal);')
    [void]$ScriptBlock.AppendLine('      });')
    [void]$ScriptBlock.AppendLine('      rows.forEach(function (r) { errTable.appendChild(r); });')
    [void]$ScriptBlock.AppendLine('    });')
    [void]$ScriptBlock.AppendLine('  });')
    [void]$ScriptBlock.AppendLine('  function applyFilters() {')
    [void]$ScriptBlock.AppendLine('    var q = searchBox ? searchBox.value.toLowerCase() : "";')
    [void]$ScriptBlock.AppendLine('    var sv = document.getElementById("filterServer").value;')
    [void]$ScriptBlock.AppendLine('    var qv = document.getElementById("filterQuelle").value;')
    [void]$ScriptBlock.AppendLine('    rows.forEach(function (r) {')
    [void]$ScriptBlock.AppendLine('      var cells = Array.prototype.slice.call(r.cells);')
    [void]$ScriptBlock.AppendLine('      var textMatch = q === "" || cells.some(function (c) { return c.textContent.toLowerCase().indexOf(q) !== -1; });')
    [void]$ScriptBlock.AppendLine('      var serverMatch = sv === "" || (cells[1] ? cells[1].textContent.trim() : "") === sv;')
    [void]$ScriptBlock.AppendLine('      var quelleMatch = qv === "" || (cells[3] ? cells[3].textContent.trim() : "") === qv;')
    [void]$ScriptBlock.AppendLine('      r.classList.toggle("hidden", !(textMatch && serverMatch && quelleMatch));')
    [void]$ScriptBlock.AppendLine('    });')
    [void]$ScriptBlock.AppendLine('  }')
    [void]$ScriptBlock.AppendLine('  if (searchBox) searchBox.addEventListener("input", applyFilters);')
    [void]$ScriptBlock.AppendLine('  var mr = document.querySelector("meta[http-equiv=refresh]");')
    [void]$ScriptBlock.AppendLine('  if (mr) {')
    [void]$ScriptBlock.AppendLine('    var totalSec = parseInt(mr.getAttribute("content"), 10);')
    [void]$ScriptBlock.AppendLine('    var el = document.getElementById("countdown");')
    [void]$ScriptBlock.AppendLine('    if (el) {')
    [void]$ScriptBlock.AppendLine('      var startTime = new Date().getTime();')
    [void]$ScriptBlock.AppendLine('      (function tick() {')
    [void]$ScriptBlock.AppendLine('        var elapsed = Math.floor((new Date().getTime() - startTime) / 1000);')
    [void]$ScriptBlock.AppendLine('        var rem = Math.max(0, totalSec - elapsed);')
    [void]$ScriptBlock.AppendLine('        var m = Math.floor(rem / 60);')
    [void]$ScriptBlock.AppendLine('        var s = rem % 60;')
    [void]$ScriptBlock.AppendLine('        el.textContent = m + " Min " + (s < 10 ? "0" : "") + s + " Sek";')
    [void]$ScriptBlock.AppendLine('        setTimeout(tick, 1000);')
    [void]$ScriptBlock.AppendLine('      })();')
    [void]$ScriptBlock.AppendLine('    }')
    [void]$ScriptBlock.AppendLine('  }')
    [void]$ScriptBlock.AppendLine('  (function() {')
    [void]$ScriptBlock.AppendLine('    var tbl = document.getElementById("errorTable");')
    [void]$ScriptBlock.AppendLine('    if (!tbl) return;')
    [void]$ScriptBlock.AppendLine('    var rowCount = tbl.querySelectorAll("tr").length;')
    [void]$ScriptBlock.AppendLine('    var prevCount = sessionStorage.getItem("eventRowCount");')
    [void]$ScriptBlock.AppendLine('    if (prevCount !== null && parseInt(prevCount, 10) < rowCount) {')
    [void]$ScriptBlock.AppendLine('      try {')
    [void]$ScriptBlock.AppendLine('        var actx = new (window.AudioContext || window.webkitAudioContext)();')
    [void]$ScriptBlock.AppendLine('        var osc = actx.createOscillator();')
    [void]$ScriptBlock.AppendLine('        var gain = actx.createGain();')
    [void]$ScriptBlock.AppendLine('        osc.connect(gain);')
    [void]$ScriptBlock.AppendLine('        gain.connect(actx.destination);')
    [void]$ScriptBlock.AppendLine('        osc.frequency.value = 800;')
    [void]$ScriptBlock.AppendLine('        gain.gain.value = 0.3;')
    [void]$ScriptBlock.AppendLine('        osc.start(0);')
    [void]$ScriptBlock.AppendLine('        osc.stop(actx.currentTime + 0.3);')
    [void]$ScriptBlock.AppendLine('      } catch(e) {}')
    [void]$ScriptBlock.AppendLine('    }')
    [void]$ScriptBlock.AppendLine('    sessionStorage.setItem("eventRowCount", rowCount);')
    [void]$ScriptBlock.AppendLine('  })();')
    [void]$ScriptBlock.AppendLine('});')
    [void]$ScriptBlock.AppendLine('</script>')

    $HtmlHead = $StyleBlock.ToString() + "`r`n" + $ScriptBlock.ToString()

    # ------------------- Event-Tabelle -------------------
    $HtmlBodyErrors = $Results | Sort-Object Time -Descending | Select-Object @{N='Zeit';E={
        if ($_.Time) { $_.Time.ToString('dd.MM.yyyy HH:mm:ss') } else { '-' }
    }}, @{N='Server';E={ $_.Server }}, @{N='Level';E={ $_.Level }}, @{N='Quelle';E={ $_.Provider }}, @{N='Fehler';E={
        $msg = $_.Message
        if ($msg -match '(Nicht genügend|Zu wenig|Out of memory|Virtual memory)') { $msg = '__OOM__' + $msg }
        $msg = $msg -replace '(?:[A-Za-z]:\\[^<>:"|?*\n]+\.exe\b|\b\w+\.exe\b)', '__EXE__$&__/EXE__'
        $msg = $msg -replace '(?:[A-Za-z]:\\[^<>:"|?*\n]+\.dll\b|\b\w+\.dll\b)', '__DLL__$&__/DLL__'
        $msg = $msg -replace '\b0x[a-fA-F0-9]+\b', '__EXCODE__$&__/EXCODE__'
        $msg
    }} | ConvertTo-Html -Fragment

    $HtmlBodyErrors = $HtmlBodyErrors -replace '<table>', '<table id="errorTable">'
    $HtmlBodyErrors = $HtmlBodyErrors -replace '<td>Error</td>',       '<td class="lv-Error">Error</td>'
    $HtmlBodyErrors = $HtmlBodyErrors -replace '<td>Warning</td>',     '<td class="lv-Warning">Warning</td>'
    $HtmlBodyErrors = $HtmlBodyErrors -replace '<td>Information</td>', '<td class="lv-Information">Information</td>'
    $HtmlBodyErrors = $HtmlBodyErrors -replace '<td>__OOM__', '<td class="msg-oom">'
    $HtmlBodyErrors = $HtmlBodyErrors -replace '__EXE__', '<span class="exe">'
    $HtmlBodyErrors = $HtmlBodyErrors -replace '__/EXE__', '</span>'
    $HtmlBodyErrors = $HtmlBodyErrors -replace '__DLL__', '<span class="dll">'
    $HtmlBodyErrors = $HtmlBodyErrors -replace '__/DLL__', '</span>'
    $HtmlBodyErrors = $HtmlBodyErrors -replace '__EXCODE__', '<span class="excode">'
    $HtmlBodyErrors = $HtmlBodyErrors -replace '__/EXCODE__', '</span>'

    # ------------------- HTML zusammenbauen -------------------
    $Html = New-Object System.Text.StringBuilder
    [void]$Html.AppendLine('<!DOCTYPE html>')
    [void]$Html.AppendLine('<html>')
    [void]$Html.AppendLine('<head>')
    [void]$Html.AppendLine('    <meta charset="UTF-8">')
    if ($Interval -gt 0) {
        $refreshSeconds = $Interval * 60 + 30
        [void]$Html.AppendLine("    <meta http-equiv=""refresh"" content=""$refreshSeconds"">")
    }
    [void]$Html.AppendLine('    <title>System-Status &amp; Application-Fehler Report</title>')
    [void]$Html.AppendLine($HtmlHead)
    [void]$Html.AppendLine('</head>')
    [void]$Html.AppendLine('<body>')
    [void]$Html.AppendLine('    <h1>System-Status &amp; Application-Fehler Report</h1>')

    $SummaryLine  = '    <div class="summary">'
    $SummaryLine += 'Erstellt: ' + (Get-Date -Format 'dd.MM.yyyy HH:mm') + ' | '
    if ($Interval -gt 0) {
        $SummaryLine += 'Letzte Aktualisierung: <span id="lastUpdate">' + (Get-Date -Format 'HH:mm:ss') + '</span> | '
        $SummaryLine += 'Nächste Aktualisierung in: <span id="countdown">--</span> | '
    }
    $SummaryLine += 'Zeitraum: ' + $StartTime.ToString('dd.MM.yyyy') + ' - ' + (Get-Date -Format 'dd.MM.yyyy') + ' | '
    $SummaryLine += 'Server: ' + $Computers.Count + ' | Event-Einträge: ' + $Results.Count
    $SummaryLine += '    </div>'
    [void]$Html.AppendLine($SummaryLine)

    # ---- Ampel-Tabelle ----
    if ($HasDrives) {
        [void]$Html.AppendLine('    <h2>Statusübersicht</h2>')
        [void]$Html.AppendLine('    <table id="ampelTable">')
        [void]$Html.AppendLine('        <tr><th>Server</th><th>Freier Speicher D:</th><th>mcsdif.vhdx</th><th>Freier Arbeitsspeicher</th><th>CPU %</th><th>Auslagerungsdatei</th><th>Sessions</th><th>FSLogix-Dienste</th></tr>')

        foreach ($drv in ($DriveResults | Sort-Object Server)) {
            $serverName = $drv.Server

            if ($drv.FreePct -ne $null) {
                if ($drv.FreePct -lt 10)      { $ampelSpeicher = 'ampel-red';    $statusSpeicher = "$($drv.FreePct)% frei" }
                elseif ($drv.FreePct -lt 20)  { $ampelSpeicher = 'ampel-yellow'; $statusSpeicher = "$($drv.FreePct)% frei" }
                else                          { $ampelSpeicher = 'ampel-green';  $statusSpeicher = "$($drv.FreePct)% frei" }
            } else { $ampelSpeicher = 'ampel-gray'; $statusSpeicher = 'n/a' }

            if ($drv.VhdxExists -and $drv.VhdxSizeGB -ne $null) {
                if ($drv.VhdxSizeGB -gt 30)      { $ampelDatei = 'ampel-red';    $statusDatei = "$($drv.VhdxSizeGB) GB" }
                elseif ($drv.VhdxSizeGB -gt 20)  { $ampelDatei = 'ampel-yellow'; $statusDatei = "$($drv.VhdxSizeGB) GB" }
                else                             { $ampelDatei = 'ampel-green';  $statusDatei = "$($drv.VhdxSizeGB) GB" }
            } elseif ($drv.VhdxExists) { $ampelDatei = 'ampel-gray'; $statusDatei = 'Größe unbekannt' }
            else { $ampelDatei = 'ampel-gray'; $statusDatei = 'Datei fehlt' }

            if ($drv.MemFreePct -ne $null) {
                if ($drv.MemFreePct -lt 10)     { $ampelMem = 'ampel-red';    $statusMem = "$($drv.MemFreePct)% frei" }
                elseif ($drv.MemFreePct -lt 20) { $ampelMem = 'ampel-yellow'; $statusMem = "$($drv.MemFreePct)% frei" }
                else                            { $ampelMem = 'ampel-green';  $statusMem = "$($drv.MemFreePct)% frei" }
            } else { $ampelMem = 'ampel-gray'; $statusMem = 'n/a' }

            # CPU-Ampel
            if ($drv.CpuPct -ne $null) {
                if ($drv.CpuPct -ge 90)      { $ampelCpu = 'ampel-red';    $statusCpu = "$($drv.CpuPct)%" }
                elseif ($drv.CpuPct -ge 75)  { $ampelCpu = 'ampel-yellow'; $statusCpu = "$($drv.CpuPct)%" }
                else                         { $ampelCpu = 'ampel-green';  $statusCpu = "$($drv.CpuPct)%" }
            } else { $ampelCpu = 'ampel-gray'; $statusCpu = 'n/a' }

            # Pagefile-Ampel
            if ($drv.PageFileUsagePct -ne $null) {
                if ($drv.PageFileUsagePct -ge 90)      { $ampelPf = 'ampel-red';    $statusPf = "$($drv.PageFileMB) MB ($($drv.PageFileUsagePct)%)" }
                elseif ($drv.PageFileUsagePct -ge 75)  { $ampelPf = 'ampel-yellow'; $statusPf = "$($drv.PageFileMB) MB ($($drv.PageFileUsagePct)%)" }
                else                                   { $ampelPf = 'ampel-green';  $statusPf = "$($drv.PageFileMB) MB ($($drv.PageFileUsagePct)%)" }
            } else { $ampelPf = 'ampel-gray'; $statusPf = 'n/a' }

            $sessionDisplay = if ($drv.Sessions -and $drv.Sessions -gt 0) { $drv.Sessions } else { '-' }

            # FSLogix-Status farbig darstellen
            $fslDisplay = @()
            if ($drv.FslogixServices -and $drv.FslogixServices.Count -gt 0) {
                foreach ($svc in $drv.FslogixServices) {
                    $color = if ($svc.Status -eq 'Running') { '#27ae60' } else { '#e74c3c' }
                    $fslDisplay += "<span style=""color:$color;font-weight:600"">$($svc.Name)=$($svc.Status)</span>"
                }
            }
            $fslHtml = if ($fslDisplay) { $fslDisplay -join '<br>' } else { '<span class="ampel ampel-gray" style="font-size:11px">keine</span>' }

            [void]$Html.AppendLine("        <tr><td>$serverName</td><td><span class=""ampel $ampelSpeicher"">$statusSpeicher</span></td><td><span class=""ampel $ampelDatei"">$statusDatei</span></td><td><span class=""ampel $ampelMem"">$statusMem</span></td><td><span class=""ampel $ampelCpu"">$statusCpu</span></td><td><span class=""ampel $ampelPf"">$statusPf</span></td><td>$sessionDisplay</td><td style=""font-size:12px;white-space:nowrap"">$fslHtml</td></tr>")
        }
        [void]$Html.AppendLine('    </table>')
    }

    # ---- TOP 10 Speicherfresser ----
    if ($Top10 -and $Top10.Count -gt 0) {
        [void]$Html.AppendLine('    <h2>TOP 10 Speicherfresser (WorkingSet)</h2>')
        [void]$Html.AppendLine('    <table id="topTable">')
        [void]$Html.AppendLine('        <tr><th>#</th><th>Server</th><th>Prozess</th><th>PID</th><th>Arbeitsspeicher (MB)</th></tr>')
        $rank = 0
        foreach ($p in $Top10) {
            $rank++
            [void]$Html.AppendLine("        <tr><td>$rank</td><td>$($p.Server)</td><td>$($p.Name)</td><td>$($p.PID)</td><td>$($p.WS_MB) MB</td></tr>")
        }
        [void]$Html.AppendLine('    </table>')
    }

    # ---- TOP 10 Session-RAM (mit Klick auf Prozess-Anzahl -> Modal) ----
    if ($TopSessionRAM -and $TopSessionRAM.Count -gt 0) {
        $sessionDataForJs = $TopSessionRAM | ForEach-Object {
            @{ User = $_.User; Server = $_.Server; ProcessList = $_.ProcessList }
        }
        $sessionDataJson = $sessionDataForJs | ConvertTo-Json -Compress
        [void]$Html.AppendLine('    <div id="sessionProcessData" style="display:none">' + $sessionDataJson + '</div>')

        [void]$Html.AppendLine('    <!-- Modal für Session-Detail -->')
        [void]$Html.AppendLine('    <div id="sessionModal" class="modal-overlay" style="display:none">')
        [void]$Html.AppendLine('      <div class="modal-box">')
        [void]$Html.AppendLine('        <span id="modalClose" class="modal-close">&times;</span>')
        [void]$Html.AppendLine('        <h2>Session: <span id="modalUser"></span> (<span id="modalServer"></span>)</h2>')
        [void]$Html.AppendLine('        <table><thead><tr><th>#</th><th>Prozess</th><th>Arbeitsspeicher (MB)</th></tr></thead><tbody id="modalTableBody"></tbody></table>')
        [void]$Html.AppendLine('      </div>')
        [void]$Html.AppendLine('    </div>')
        [void]$Html.AppendLine('    <h2>TOP 10 Session-RAM (WorkingSet)</h2>')
        [void]$Html.AppendLine('    <table id="sessionRamTable">')
        [void]$Html.AppendLine('        <tr><th>#</th><th>Server</th><th>Benutzer</th><th>Prozesse</th><th>Arbeitsspeicher (MB)</th></tr>')
        $rank = 0
        foreach ($s in $TopSessionRAM) {
            [void]$Html.AppendLine("        <tr><td>$($rank + 1)</td><td>$($s.Server)</td><td>$($s.User)</td><td><a href=""javascript:void(0)"" onclick=""showSessionProcs($rank)"" style=""font-weight:700;color:#007acc;cursor:pointer"">$($s.ProcessCount)</a></td><td>$($s.SessionRAM_MB) MB</td></tr>")
            $rank++
        }
        [void]$Html.AppendLine('    </table>')
    }

    # ---- Eventlog-Tabelle ----
    if ($HasErrors) {
        [void]$Html.AppendLine('    <h2>Eventlog-Einträge</h2>')
        [void]$Html.AppendLine('    <div class="filterBar" id="filterBar">')
        [void]$Html.AppendLine('        <div><label for="filterServer">Server:</label><select id="filterServer"><option value="">Alle</option></select></div>')
        [void]$Html.AppendLine('        <div><label for="filterQuelle">Quelle:</label><select id="filterQuelle"><option value="">Alle</option></select></div>')
        [void]$Html.AppendLine('        <div><label for="searchBox">Suche:</label><input type="text" id="searchBox" placeholder="Beliebiger Text ..." /></div>')
        [void]$Html.AppendLine('    </div>')
        [void]$Html.AppendLine($HtmlBodyErrors)
    }

    [void]$Html.AppendLine('</body>')
    [void]$Html.AppendLine('</html>')

    # =====================================================================
    # Region: Ausgabe
    # =====================================================================
    try {
        $OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
        $Html.ToString() | Out-File -FilePath $OutputPath -Encoding utf8 -Force
        Write-Host "Report erstellt: $OutputPath" -ForegroundColor Green
        Write-Log  "Report erstellt: $OutputPath"

        if ((Get-Item $OutputPath).Length -gt 0 -and $firstRun) {
            Start-Process $OutputPath
            $firstRun = $false
        }
    }
    catch {
        Write-Error "Fehler beim Schreiben der Datei: $_"
        Write-Log   "Fehler beim Schreiben der Datei: $($_.Exception.Message)" 'ERROR'
        break
    }

    # =====================================================================
    # Region: Loop-Ende (bei -Interval > 0)
    # =====================================================================
    if ($Interval -gt 0) {
        $nextRun = (Get-Date).AddMinutes($Interval)
        Write-Host "Nächste Aktualisierung: $($nextRun.ToString('dd.MM.yyyy HH:mm'))" -ForegroundColor Cyan
        Write-Log  "Warte bis nächste Aktualisierung: $($nextRun.ToString('dd.MM.yyyy HH:mm'))"
        Start-Sleep -Seconds ($Interval * 60)
    }
}
catch {
    # Übergeordnete Fehlerbehandlung für den gesamten Durchlauf
    Write-Error "Unerwarteter Fehler im Hauptdurchlauf: $_"
    Write-Log   "Unerwarteter Fehler im Hauptdurchlauf: $($_.Exception.Message)" 'ERROR'
    if ($Interval -le 0) { break }
    Start-Sleep -Seconds 30
}
} while ($Interval -gt 0)

Write-Log "Skriptende."
