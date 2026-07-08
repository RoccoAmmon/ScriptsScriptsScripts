<#
.SYNOPSIS
    Exportiert Application Error/Hang/Popup Eventlog-Einträge aller Server einer OU als HTML.

.DESCRIPTION
    Sammelt von allen Servern in einer angegebenen OU die Eventlog-Einträge mit den Quellen
    "Application Error", "Application Hang" und "Application Popup" und exportiert sie
    übersichtlich als HTML-Datei (Uhrzeit, Server, Fehler). Zusätzlich werden freier
    Speicherplatz D: und Größe der mcsdif.vhdx als Ampelanzeige (rot/gelb/grün) dargestellt.

.PARAMETER SearchBase
    DistinguishedName der OU, in der nach Servern gesucht wird.

.PARAMETER OutputPath
    Pfad zur Ausgabe-HTML-Datei. Standard: Skriptverzeichnis\ApplicationErrors.html

.PARAMETER DaysBack
    Nur Einträge der letzten X Tage berücksichtigen. Standard: 7

.EXAMPLE
    .\Application_Error_Eventlog.ps1 -SearchBase "OU=Servers,DC=domain,DC=local"
    -OutputPath "C:\Reports\errors.html" -DaysBack 14

.NOTES
    Version  : 1.0
    Autor    : Rocco Ammon
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$SearchBase,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath "ApplicationErrors.html"),

    [Parameter(Mandatory = $false)]
    [int]$DaysBack = 7
)

# Region: Konfiguration
$ErrorActionPreference = 'Stop'
$StartTime = (Get-Date).AddDays(-$DaysBack)

# Pro LogName die zu suchenden Provider (Quellen)
$LogQueries = @(
    @{ LogName = 'Application'; Providers = @('Application Error', 'Application Hang') }
    @{ LogName = 'System';      Providers = @('Application Popup') }
)

# Region: AD abfragen
try {
    Write-Verbose "Suche Computer in OU: $SearchBase"
    $Computers = Get-ADComputer -Filter 'enabled -eq "true"' -SearchBase $SearchBase `
        -Properties Name -ErrorAction Stop | Select-Object -ExpandProperty Name
}
catch {
    Write-Error "Fehler bei AD-Abfrage: $_"
    exit 1
}

if (-not $Computers) {
    Write-Warning "Keine Computer in der angegebenen OU gefunden."
    exit 0
}

Write-Verbose "Gefundene Server: $($Computers.Count)"

# Region: Eventlogs sammeln
$Results = [System.Collections.ArrayList]::new()

foreach ($Computer in $Computers) {
    Write-Verbose "Verarbeite $Computer ..."

    foreach ($Query in $LogQueries) {
        $LogName  = $Query.LogName
        $Provider = $Query.Providers

        $XPathFilter = "*[System[{0}]]" -f (
            ($Provider | ForEach-Object { "Provider[@Name='$_']" }) -join ' or '
        )

        try {
            $Events = Get-WinEvent -ComputerName $Computer -FilterXPath $XPathFilter `
                -LogName $LogName -ErrorAction Stop | Where-Object {
                    $_.TimeCreated -ge $StartTime
                }

            foreach ($Event in $Events) {
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
                    Message  = ($Event.Message -replace '\r?\n', ' ')
                })
            }
        }
        catch [Exception] {
            Write-Warning "$Computer | $LogName nicht erreichbar oder keine Berechtigung: $_"
        }
    }
}

# Region: Festplatten-Status sammeln (freier Speicher D:, mcsdif.vhdx)
$DriveResults = [System.Collections.ArrayList]::new()

foreach ($Computer in $Computers) {
    Write-Verbose "Sammle Festplatteninfos von $Computer ..."
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
    })
}

# Region: HTML erstellen
$HasErrors = $Results.Count -gt 0
$HasDrives  = $DriveResults.Count -gt 0

if (-not $HasErrors -and -not $HasDrives) {
    Write-Warning "Keine Daten gefunden – Report wird nicht erstellt."
    exit 0
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
[void]$StyleBlock.AppendLine('</style>')

$ScriptBlock = [System.Text.StringBuilder]::new()
[void]$ScriptBlock.AppendLine('<script>')
[void]$ScriptBlock.AppendLine('document.addEventListener("DOMContentLoaded", function () {')
[void]$ScriptBlock.AppendLine('  var table = document.querySelector("table:not(#ampelTable)");')
[void]$ScriptBlock.AppendLine('  if (!table) return;')
[void]$ScriptBlock.AppendLine('  var rows = Array.prototype.filter.call(table.querySelectorAll("tr"), function (r) { return r.querySelector("td"); });')
[void]$ScriptBlock.AppendLine('  var searchBox = document.getElementById("searchBox");')
[void]$ScriptBlock.AppendLine('  function getUnique(col) { var set = {}; rows.forEach(function (r) { var val = (r.cells[col] ? r.cells[col].textContent : "").trim(); if (val) set[val] = true; }); return Object.keys(set).sort(); }')
[void]$ScriptBlock.AppendLine('  function populateFilter(id, col) { var sel = document.getElementById(id); if (!sel) return; getUnique(col).forEach(function (v) { var opt = document.createElement("option"); opt.value = v; opt.textContent = v; sel.appendChild(opt); }); sel.addEventListener("change", applyFilters); }')
[void]$ScriptBlock.AppendLine('  populateFilter("filterServer", 1);')
[void]$ScriptBlock.AppendLine('  populateFilter("filterQuelle", 3);')
[void]$ScriptBlock.AppendLine('  table.querySelectorAll("th").forEach(function (th, col) {')
[void]$ScriptBlock.AppendLine('    th.addEventListener("click", function () {')
[void]$ScriptBlock.AppendLine('      var isAsc = th.classList.contains("asc");')
[void]$ScriptBlock.AppendLine('      table.querySelectorAll("th").forEach(function (h) { h.classList.remove("asc", "desc"); });')
[void]$ScriptBlock.AppendLine('      th.classList.add(isAsc ? "desc" : "asc");')
[void]$ScriptBlock.AppendLine('      rows.sort(function (a, b) { var aVal = (a.cells[col] ? a.cells[col].textContent : "").trim(); var bVal = (b.cells[col] ? b.cells[col].textContent : "").trim(); return isAsc ? bVal.localeCompare(aVal) : aVal.localeCompare(bVal); });')
[void]$ScriptBlock.AppendLine('      rows.forEach(function (r) { table.appendChild(r); });')
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
[void]$ScriptBlock.AppendLine('});')
[void]$ScriptBlock.AppendLine('</script>')

$HtmlHead = $StyleBlock.ToString() + "`r`n" + $ScriptBlock.ToString()

$HtmlBodyErrors = $Results | Sort-Object Time -Descending | Select-Object @{N='Zeit';E={
    if ($_.Time) { $_.Time.ToString('dd.MM.yyyy HH:mm:ss') } else { '-' }
}}, @{N='Server';E={ $_.Server }}, @{N='Level';E={ $_.Level }}, @{N='Quelle';E={ $_.Provider }}, @{N='Fehler';E={
    $_.Message
}} | ConvertTo-Html -Fragment

# Level-Zellen farblich markieren
$HtmlBodyErrors = $HtmlBodyErrors -replace '<td>Error</td>',       '<td class="lv-Error">Error</td>'
$HtmlBodyErrors = $HtmlBodyErrors -replace '<td>Warning</td>',     '<td class="lv-Warning">Warning</td>'
$HtmlBodyErrors = $HtmlBodyErrors -replace '<td>Information</td>', '<td class="lv-Information">Information</td>'

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
[void]$Html.AppendLine('    <title>Application Error / Hang / Popup - Report</title>')
[void]$Html.AppendLine($HtmlHead)
[void]$Html.AppendLine('</head>')
[void]$Html.AppendLine('<body>')
[void]$Html.AppendLine('    <h1>Application-Fehler Report</h1>')

$SummaryLine = '    <div class="summary">'
$SummaryLine += 'Erstellt: ' + (Get-Date -Format 'dd.MM.yyyy HH:mm') + ' | '
$SummaryLine += 'Zeitraum: ' + $StartTime.ToString('dd.MM.yyyy') + ' - ' + (Get-Date -Format 'dd.MM.yyyy') + ' | '
$SummaryLine += 'Server: ' + $Computers.Count + ' | Event-Einträge: ' + $Results.Count
$SummaryLine += '    </div>'
[void]$Html.AppendLine($SummaryLine)

# Ampel-Tabelle (Speicher & Datei)
if ($HasDrives) {
    [void]$Html.AppendLine('    <h2>Statusübersicht</h2>')
    [void]$Html.AppendLine('    <table id="ampelTable">')
    [void]$Html.AppendLine('        <tr><th>Server</th><th>Freier Speicher D:</th><th>mcsdif.vhdx</th></tr>')

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

        [void]$Html.AppendLine("        <tr><td>$serverName</td><td><span class=""ampel $ampelSpeicher"">$statusSpeicher</span></td><td><span class=""ampel $ampelDatei"">$statusDatei</span></td></tr>")
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

    if ((Get-Item $OutputPath).Length -gt 0) {
        Start-Process $OutputPath
    }
}
catch {
    Write-Error "Fehler beim Schreiben der Datei: $_"
    exit 1
}
