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
    "Nicht genügend virtueller Speicher" (rot). Zeigt aktive Citrix Sessions pro Server
    und eine TOP 10 der größten Speicherfresser (WorkingSet) aller Server an.

.PARAMETER SearchBase
    DistinguishedName der OU, in der nach Servern gesucht wird.

.PARAMETER OutputPath
    Pfad zur Ausgabe-HTML-Datei. Standard: Skriptverzeichnis\SystemStatusReport.html

.PARAMETER DaysBack
    Nur Einträge der letzten X Tage berücksichtigen. Standard: 7

.PARAMETER Interval
    Intervall in Minuten für automatische Wiederholung (0 = einmalig). Bei &gt; 0 läuft das
    Skript durchgehend und aktualisiert die HTML-Datei im angegebenen Intervall.

.EXAMPLE
    .\System-Status-Application-Fehler-Report.ps1 -SearchBase "OU=Servers,DC=domain,DC=local"
    -OutputPath "C:\Reports\SystemStatus.html" -DaysBack 14

.EXAMPLE
    .\System-Status-Application-Fehler-Report.ps1 -SearchBase "OU=Servers,DC=domain,DC=local" -Interval 30

.NOTES
    Version  : 1.5
    Autor    : Rocco Ammon
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

# Region: Konfiguration
$ErrorActionPreference = 'Stop'
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { '.' }
if (-not $OutputPath) { $OutputPath = Join-Path -Path $ScriptDir -ChildPath "SystemStatusReport.html" }
$firstRun = $true

do {
    $StartTime = (Get-Date).AddDays(-$DaysBack)

# Pro LogName die zu suchenden Provider (Quellen)
$LogQueries = @(
    @{ LogName = 'Application'; Providers = @('Application Error', 'Application Hang') }
    @{ LogName = 'System';      Providers = @('Application Popup') }
    @{ LogName = 'System';      Providers = @('Service Control Manager'); EventIDs = @(7031, 7032, 7034, 7000, 7009) }
    @{ LogName = 'System';      Providers = @('Windows Resource Exhaustion'); EventIDs = @(2001, 2004, 2005, 2006, 2020) }
)

# Region: AD abfragen
Write-Host "Suche Computer in OU: $SearchBase ..." -ForegroundColor Cyan
try {
    $Computers = Get-ADComputer -Filter 'enabled -eq "true"' -SearchBase $SearchBase `
        -Properties Name -ErrorAction Stop | Select-Object -ExpandProperty Name
}
catch {
    Write-Error "Fehler bei AD-Abfrage: $_"
    break
}

if (-not $Computers) {
    Write-Warning "Keine Computer in der angegebenen OU gefunden."
    break
}

Write-Host "$($Computers.Count) Server gefunden, sammle Eventlogs ..." -ForegroundColor Cyan

# Region: Eventlogs sammeln
$Results = [System.Collections.ArrayList]::new()
$serverCount = $Computers.Count
$currentServer = 0

foreach ($Computer in $Computers) {
    $currentServer++
    Write-Host "[$currentServer/$serverCount] $Computer ..." -ForegroundColor Yellow

    # Prüfen ob Server erreichbar ist
    $reachable = Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $reachable) {
        Write-Warning "$Computer | nicht erreichbar"
        continue
    }

    foreach ($Query in $LogQueries) {
        $LogName  = $Query.LogName
        $Provider = $Query.Providers

        $XPathFilter = "*[System[{0}]]" -f (
            ($Provider | ForEach-Object { "Provider[@Name='$_']" }) -join ' or '
        )

        try {
            # -ErrorAction Stop NICHT verwenden, da sonst Warnungen wie
            # "Beschreibungszeichenfolge" (unvollständige Message-Resolution)
            # die Ereignisse unterdrücken würden
            $rawEvents = Get-WinEvent -ComputerName $Computer -FilterXPath $XPathFilter `
                -LogName $LogName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue `
                -ErrorVariable getErr

            if (-not $rawEvents) {
                if ($getErr -and $getErr[0].Exception.Message -match 'Keine Ereignisse gefunden|No events were found') {
                    Write-Verbose "$Computer | $LogName | keine Fehler im Eventlog"
                } elseif ($getErr) {
                    Write-Warning "$Computer | $LogName | Fehler: $($getErr[0].Exception.Message)"
                }
                continue
            }

            $Events = $rawEvents | Where-Object {
                    $_.TimeCreated -ge $StartTime -and
                    (-not $Query.EventIDs -or ($_.Id -in $Query.EventIDs))
                }

            foreach ($Event in $Events) {
                $msg = ($Event.Message -replace '\r?\n', ' ')
                # Dienstfehler mit "Windows Update" oder "wuauserv" ignorieren
                if ($msg -match 'Windows Update|wuauserv') { continue }
                [void]$Results.Add([PSCustomObject]@{
                    Time     = $Event.TimeCreated
                    Server   = $Computer
                    Level    = switch ($Event.Level) {
                        1 { 'Critical' }
                        2 { 'Error' }
                        3 { 'Warning' }
                        4 { 'Information' }
                        default { "Level $($Event.Level)" }
                    }
                    Provider = $Event.ProviderName
                    Message  = $msg
                })
            }
        }
        catch [Exception] {
            Write-Warning "$Computer | $LogName | Fehler: $_"
        }
    }
}

# Region: Neuen-Einträge-Erkennung (Piepton)
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
@{ LastEventCount = $Results.Count; Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') } | ConvertTo-Json | Set-Content $CacheFile -Encoding UTF8

Write-Host "Eventlogs ausgewertet, sammle Systeminfos (Speicher, Dateien, RAM) ..." -ForegroundColor Cyan

# Region: Festplatten-Status & Arbeitsspeicher sammeln (D:, mcsdif.vhdx, RAM)
$DriveResults = [System.Collections.ArrayList]::new()

$currentServer = 0
foreach ($Computer in $Computers) {
    $currentServer++
    Write-Host "[$currentServer/$serverCount] $Computer ..." -ForegroundColor Yellow
    Write-Verbose "Sammle Systeminfos von $Computer ..."
    $drive = $null
    $fileInfo = $null

    try {
        $drive = Get-CimInstance -ComputerName $Computer -ClassName Win32_LogicalDisk `
            -Filter "DeviceID='D:'" -ErrorAction Stop
    }
    catch {
        Write-Warning "$Computer | Festplatte D: nicht ermittelbar: $_"
    }

    try {
        $fileInfo = Invoke-Command -ComputerName $Computer -ScriptBlock {
            $f = Get-Item 'D:\mcsdif.vhdx' -ErrorAction SilentlyContinue
            if ($f) { [PSCustomObject]@{ Exists = $true; Size = $f.Length } }
            else { [PSCustomObject]@{ Exists = $false; Size = $null } }
        } -ErrorAction Stop
    }
    catch {
        Write-Warning "$Computer | Datei mcsdif.vhdx nicht ermittelbar: $_"
        $fileInfo = [PSCustomObject]@{ Exists = $false; Size = $null }
    }

    # Arbeitsspeicher abfragen
    $mem = $null
    try {
        $mem = Get-CimInstance -ComputerName $Computer -ClassName Win32_OperatingSystem -ErrorAction Stop
    }
    catch {
        Write-Warning "$Computer | Arbeitsspeicher nicht ermittelbar: $_"
    }

    # Sessions zählen (mehrere Methoden)
    $sessionCount = 0
    # 1. explorer.exe pro Benutzersitzung (zuverlässig auf Terminalservern)
    try {
        $procs = Get-CimInstance -ComputerName $Computer -ClassName Win32_Process `
            -Filter "Name = 'explorer.exe'" -ErrorAction SilentlyContinue
        if ($procs) { $sessionCount = @($procs).Count }
    } catch {}
    # 2. Fallback: query session zählt nur aktive Sessions
    if ($sessionCount -eq 0) {
        try {
            $qsRaw = Invoke-Command -ComputerName $Computer -ScriptBlock {
                @(query session 2>$null | Select-String 'Active').Count
            } -ErrorAction SilentlyContinue
            if ($qsRaw) { $sessionCount = $qsRaw }
        } catch {}
    }

    if ($drive) {
        $freeGB  = [math]::Round($drive.FreeSpace / 1GB, 2)
        $totalGB = [math]::Round($drive.Size / 1GB, 2)
        $freePct = if ($drive.Size -gt 0) { [math]::Round(($drive.FreeSpace / $drive.Size) * 100, 1) } else { 0 }
    } else {
        $freeGB  = $null
        $totalGB = $null
        $freePct = $null
    }

    [void]$DriveResults.Add([PSCustomObject]@{
        Server         = $Computer
        FreeGB         = $freeGB
        TotalGB        = $totalGB
        FreePct        = $freePct
        VhdxExists     = $fileInfo.Exists
        VhdxSizeGB     = if ($fileInfo.Size) { [math]::Round($fileInfo.Size / 1GB, 2) } else { $null }
        MemFreeMB      = if ($mem) { [math]::Round($mem.FreePhysicalMemory / 1024, 1) } else { $null }
        MemTotalMB     = if ($mem) { [math]::Round($mem.TotalVisibleMemorySize / 1024, 1) } else { $null }
        MemFreePct     = if ($mem -and $mem.TotalVisibleMemorySize -gt 0) { [math]::Round(($mem.FreePhysicalMemory / $mem.TotalVisibleMemorySize) * 100, 1) } else { $null }
        Sessions       = $sessionCount
    })
}

Write-Host "Sammle Top-Prozesse (Speicherfresser) ..." -ForegroundColor Cyan

# Region: TOP 10 Speicherfresser sammeln
$TopProcesses = [System.Collections.ArrayList]::new()
$currentServer = 0
foreach ($Computer in $Computers) {
    $currentServer++
    Write-Host "[$currentServer/$serverCount] $Computer - Top-Prozesse ..." -ForegroundColor Yellow
    try {
        $procs = Invoke-Command -ComputerName $Computer -ScriptBlock {
            Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 5 Name, Id, @{N='WS_MB';E={[math]::Round($_.WorkingSet64 / 1MB, 1)}}
        } -ErrorAction SilentlyContinue
        if ($procs) {
            foreach ($p in $procs) {
                [void]$TopProcesses.Add([PSCustomObject]@{
                    Server = $Computer
                    Name   = $p.Name
                    PID    = $p.Id
                    WS_MB  = $p.WS_MB
                })
            }
        }
    }
    catch {
        Write-Warning "$Computer | Top-Prozesse nicht ermittelbar: $_"
    }
}
$Top10 = $TopProcesses | Sort-Object WS_MB -Descending | Select-Object -First 10

Write-Host "Erstelle HTML-Report ..." -ForegroundColor Cyan

# Region: HTML erstellen
$HasErrors = $Results.Count -gt 0
$HasDrives  = $DriveResults.Count -gt 0

if (-not $HasErrors -and -not $HasDrives) {
    Write-Warning "Keine Daten gefunden – Report wird nicht erstellt."
    break
}

# Region: HTML-Bausteine aufbauen (keine Here-Strings)
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
[void]$StyleBlock.AppendLine('    #topTable td:first-child { font-weight: 700; color: #007acc; }')
[void]$StyleBlock.AppendLine('    #topTable td:nth-child(5) { font-weight: 700; text-align: right; }')
[void]$StyleBlock.AppendLine('    .sessions { text-align: center; font-weight: 600; }')
[void]$StyleBlock.AppendLine('    td.sessions-highlight { background-color: #e8f4fd; font-weight: 600; }')
[void]$StyleBlock.AppendLine('</style>')

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
[void]$ScriptBlock.AppendLine('  var errTable = document.querySelector("table:not(#ampelTable)");')
[void]$ScriptBlock.AppendLine('  if (!errTable) return;')
[void]$ScriptBlock.AppendLine('  var rows = Array.prototype.filter.call(errTable.querySelectorAll("tr"), function (r) { return r.querySelector("td"); });')
[void]$ScriptBlock.AppendLine('  var searchBox = document.getElementById("searchBox");')
[void]$ScriptBlock.AppendLine('  function getUnique(col) { var set = {}; rows.forEach(function (r) { var val = (r.cells[col] ? r.cells[col].textContent : "").trim(); if (val) set[val] = true; }); return Object.keys(set).sort(); }')
[void]$ScriptBlock.AppendLine('  function populateFilter(id, col) { var sel = document.getElementById(id); if (!sel) return; getUnique(col).forEach(function (v) { var opt = document.createElement("option"); opt.value = v; opt.textContent = v; sel.appendChild(opt); }); sel.addEventListener("change", applyFilters); }')
[void]$ScriptBlock.AppendLine('  populateFilter("filterServer", 1);')
[void]$ScriptBlock.AppendLine('  populateFilter("filterQuelle", 3);')
[void]$ScriptBlock.AppendLine('  // Sortierung für Error-Tabelle (nutzt gleiches rows-Array wie Filter)')
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
[void]$ScriptBlock.AppendLine('  // Countdown für nächste Aktualisierung')
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
[void]$ScriptBlock.AppendLine('  // Piepton bei neuen Event-Einträgen (Web Audio API)')
[void]$ScriptBlock.AppendLine('  (function() {')
[void]$ScriptBlock.AppendLine('    var tbl = document.querySelector("table:not(#ampelTable)");')
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

$HtmlBodyErrors = $Results | Sort-Object Time -Descending | Select-Object @{N='Zeit';E={
    if ($_.Time) { $_.Time.ToString('dd.MM.yyyy HH:mm:ss') } else { '-' }
}}, @{N='Server';E={ $_.Server }}, @{N='Level';E={ $_.Level }}, @{N='Quelle';E={ $_.Provider }}, @{N='Fehler';E={
    $msg = $_.Message
    # OOM-Marker setzen (für späteres CSS)
    if ($msg -match '(Nicht genügend|Zu wenig|Out of memory|Virtual memory)') {
        $msg = '__OOM__' + $msg
    }
    # EXE-Namen/Pfade (z.B. Winword.exe oder C:\...\Winword.exe) mit Marker umschließen
    $msg = $msg -replace '(?:[A-Za-z]:\\[^<>:"|?*\n]+\.exe\b|\b\w+\.exe\b)', '__EXE__$&__/EXE__'
    # DLL-Namen/Pfade (z.B. KERNELBASE.dll oder C:\...\PBVM.dll) mit Marker umschließen
    $msg = $msg -replace '(?:[A-Za-z]:\\[^<>:"|?*\n]+\.dll\b|\b\w+\.dll\b)', '__DLL__$&__/DLL__'
    # Ausnahmecodes (z.B. 0xe0000008) mit Marker umschließen
    $msg = $msg -replace '\b0x[a-fA-F0-9]+\b', '__EXCODE__$&__/EXCODE__'
    $msg
}} | ConvertTo-Html -Fragment

# Level-Zellen farblich markieren
$HtmlBodyErrors = $HtmlBodyErrors -replace '<td>Error</td>',       '<td class="lv-Error">Error</td>'
$HtmlBodyErrors = $HtmlBodyErrors -replace '<td>Warning</td>',     '<td class="lv-Warning">Warning</td>'
$HtmlBodyErrors = $HtmlBodyErrors -replace '<td>Information</td>', '<td class="lv-Information">Information</td>'

# Fehlerzelle bei virtueller Speicher hellrot (Marker durch Klasse ersetzen)
$HtmlBodyErrors = $HtmlBodyErrors -replace '<td>__OOM__', '<td class="msg-oom">'

# EXE-Namen/Pfade hervorheben (Marker durch span ersetzen)
$HtmlBodyErrors = $HtmlBodyErrors -replace '__EXE__', '<span class="exe">'
$HtmlBodyErrors = $HtmlBodyErrors -replace '__/EXE__', '</span>'

# DLL-Namen hervorheben
$HtmlBodyErrors = $HtmlBodyErrors -replace '__DLL__', '<span class="dll">'
$HtmlBodyErrors = $HtmlBodyErrors -replace '__/DLL__', '</span>'

# Ausnahmecodes hervorheben
$HtmlBodyErrors = $HtmlBodyErrors -replace '__EXCODE__', '<span class="excode">'
$HtmlBodyErrors = $HtmlBodyErrors -replace '__/EXCODE__', '</span>'

$HtmlBodyDrives = $DriveResults | Sort-Object Server | Select-Object @{N='Server';E={ $_.Server }},
    @{N='D: Frei (GB)';E={ if ($_.FreeGB -ne $null) { $_.FreeGB } else { 'n/a' } }},
    @{N='D: Gesamt (GB)';E={ if ($_.TotalGB -ne $null) { $_.TotalGB } else { 'n/a' } }},
    @{N='D: Frei %';E={ if ($_.FreePct -ne $null) { $_.FreePct } else { 'n/a' } }},
    @{N='mcsdif.vhdx';E={ if ($_.VhdxExists) { 'vorhanden' } else { 'fehlt' } }},
    @{N='mcsdif.vhdx Größe (GB)';E={ if ($_.VhdxSizeGB -ne $null) { $_.VhdxSizeGB } else { '-' } }} |
    ConvertTo-Html -Fragment

# HTML mit StringBuilder aufbauen (vermeidet Here-String-Probleme)
$Html = New-Object System.Text.StringBuilder

[void]$Html.AppendLine('<!DOCTYPE html>')
[void]$Html.AppendLine('<html>')
[void]$Html.AppendLine('<head>')
[void]$Html.AppendLine('    <meta charset="UTF-8">')
if ($Interval -gt 0) {
    $refreshSeconds = $Interval * 60 + 30  # 30 Sek. Puffer nach Skript-Neulauf
    [void]$Html.AppendLine("    <meta http-equiv=""refresh"" content=""$refreshSeconds"">")
}
    [void]$Html.AppendLine('    <title>System-Status &amp; Application-Fehler Report</title>')
[void]$Html.AppendLine($HtmlHead)
[void]$Html.AppendLine('</head>')
[void]$Html.AppendLine('<body>')
[void]$Html.AppendLine('    <h1>System-Status &amp; Application-Fehler Report</h1>')

$SummaryLine = '    <div class="summary">'
$SummaryLine += 'Erstellt: ' + (Get-Date -Format 'dd.MM.yyyy HH:mm') + ' | '
if ($Interval -gt 0) {
    $SummaryLine += 'Letzte Aktualisierung: <span id="lastUpdate">' + (Get-Date -Format 'HH:mm:ss') + '</span> | '
    $SummaryLine += 'Nächste Aktualisierung in: <span id="countdown">--</span> | '
}
$SummaryLine += 'Zeitraum: ' + $StartTime.ToString('dd.MM.yyyy') + ' - ' + (Get-Date -Format 'dd.MM.yyyy') + ' | '
$SummaryLine += 'Server: ' + $Computers.Count + ' | Event-Einträge: ' + $Results.Count
$SummaryLine += '    </div>'
[void]$Html.AppendLine($SummaryLine)

# Ampel-Tabelle (Speicher & Datei)
if ($HasDrives) {
    [void]$Html.AppendLine('    <h2>Statusübersicht</h2>')
    [void]$Html.AppendLine('    <table id="ampelTable">')
    [void]$Html.AppendLine('        <tr><th>Server</th><th>Freier Speicher D:</th><th>mcsdif.vhdx</th><th>Freier Arbeitsspeicher</th><th>Sessions</th></tr>')

    foreach ($drv in ($DriveResults | Sort-Object Server)) {
        $serverName = $drv.Server

        if ($drv.FreePct -ne $null) {
            if ($drv.FreePct -lt 10)  { $ampelSpeicher = 'ampel-red';    $statusSpeicher = "$($drv.FreePct)% frei" }
            elseif ($drv.FreePct -lt 20) { $ampelSpeicher = 'ampel-yellow'; $statusSpeicher = "$($drv.FreePct)% frei" }
            else                          { $ampelSpeicher = 'ampel-green';  $statusSpeicher = "$($drv.FreePct)% frei" }
        } else {
            $ampelSpeicher = 'ampel-gray'
            $statusSpeicher = 'n/a'
        }

        if ($drv.VhdxExists -and $drv.VhdxSizeGB -ne $null) {
            if ($drv.VhdxSizeGB -gt 30)      { $ampelDatei = 'ampel-red';    $statusDatei = "$($drv.VhdxSizeGB) GB" }
            elseif ($drv.VhdxSizeGB -gt 20)  { $ampelDatei = 'ampel-yellow'; $statusDatei = "$($drv.VhdxSizeGB) GB" }
            else                             { $ampelDatei = 'ampel-green';  $statusDatei = "$($drv.VhdxSizeGB) GB" }
        } elseif ($drv.VhdxExists) {
            $ampelDatei = 'ampel-gray'
            $statusDatei = 'Größe unbekannt'
        } else {
            $ampelDatei = 'ampel-gray'
            $statusDatei = 'Datei fehlt'
        }

        # Ampel für Arbeitsspeicher
        if ($drv.MemFreePct -ne $null) {
            if ($drv.MemFreePct -lt 10)       { $ampelMem = 'ampel-red';    $statusMem = "$($drv.MemFreePct)% frei" }
            elseif ($drv.MemFreePct -lt 20)   { $ampelMem = 'ampel-yellow'; $statusMem = "$($drv.MemFreePct)% frei" }
            else                              { $ampelMem = 'ampel-green';  $statusMem = "$($drv.MemFreePct)% frei" }
        } else {
            $ampelMem = 'ampel-gray'
            $statusMem = 'n/a'
        }

        # Sessions als Zahl (grau bei 0)
        $sessionDisplay = if ($drv.Sessions -and $drv.Sessions -gt 0) { $drv.Sessions } else { '-' }

        [void]$Html.AppendLine("        <tr><td>$serverName</td><td><span class=""ampel $ampelSpeicher"">$statusSpeicher</span></td><td><span class=""ampel $ampelDatei"">$statusDatei</span></td><td><span class=""ampel $ampelMem"">$statusMem</span></td><td>$sessionDisplay</td></tr>")
    }

    [void]$Html.AppendLine('    </table>')
}

# TOP 10 Speicherfresser
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

# Region: Ausgabe
try {
    $OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    $Html.ToString() | Out-File -FilePath $OutputPath -Encoding utf8 -Force
    Write-Host "Report erstellt: $OutputPath" -ForegroundColor Green

    if ((Get-Item $OutputPath).Length -gt 0 -and $firstRun) {
        Start-Process $OutputPath
        $firstRun = $false
    }
}
catch {
    Write-Error "Fehler beim Schreiben der Datei: $_"
    break
}

# Region: Loop-Ende (bei -Interval > 0)
if ($Interval -gt 0) {
    $nextRun = (Get-Date).AddMinutes($Interval)
    Write-Host "Nächste Aktualisierung: $($nextRun.ToString('dd.MM.yyyy HH:mm'))" -ForegroundColor Cyan
    Start-Sleep -Seconds ($Interval * 60)
}
} while ($Interval -gt 0)
