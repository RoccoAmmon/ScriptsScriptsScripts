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

# Region: HTML erstellen
if ($Results.Count -eq 0) {
    Write-Warning "Keine passenden Eventlog-Einträge gefunden."
    exit 0
}

$HtmlHead = @"
<style>
    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; }
    h1 { color: #333; }
    .summary { margin-bottom: 15px; color: #666; }
    .filterBar { display: flex; gap: 10px; margin-bottom: 10px; flex-wrap: wrap; }
    .filterBar label { font-weight: 600; margin-right: 4px; }
    .filterBar select { padding: 6px; border: 1px solid #ccc; border-radius: 4px; min-width: 180px; }
    #searchBox { width: 300px; padding: 6px; border: 1px solid #ccc; border-radius: 4px; }
    table { border-collapse: collapse; width: 100%; }
    th { background-color: #007acc; color: white; padding: 8px; text-align: left; cursor: pointer; user-select: none; }
    th::after { content: ' \25B4\25BE'; font-size: 16px; opacity: 0.3; }
    th.asc::after { content: ' \25B4'; opacity: 1; }
    th.desc::after { content: ' \25BE'; opacity: 1; }
    td { padding: 6px 8px; border-bottom: 1px solid #ddd; vertical-align: top; }
    tr:nth-child(even) { background-color: #f5f5f5; }
    tr:hover { background-color: #e0f0ff; }
    .level-Error { border-left: 4px solid #e74c3c; }
    .level-Warning { border-left: 4px solid #f39c12; }
    .level-Information { border-left: 4px solid #27ae60; }
    .level-Offline { border-left: 4px solid #95a5a6; color: #999; }
    td.Message { max-width: 600px; word-wrap: break-word; }
    .hidden { display: none; }
</style>
<script>
    document.addEventListener('DOMContentLoaded', function () {
        const table = document.querySelector('table');
        // Nur Zeilen mit Daten (<td>), Header-Zeile (<th>) ausnehmen
        const rows = Array.from(table.querySelectorAll('tr')).filter(function (r) {
            return r.querySelector('td');
        });
        const filterBar = document.getElementById('filterBar');
        const searchBox = document.getElementById('searchBox');

        // Dropdowns aus Tabellendaten befüllen
        function getUnique(col) {
            const set = new Set();
            rows.forEach(function (r) {
                const val = (r.cells[col]?.textContent || '').trim();
                if (val) set.add(val);
            });
            return Array.from(set).sort();
        }

        function populateFilter(id, col) {
            const sel = document.getElementById(id);
            getUnique(col).forEach(function (v) {
                const opt = document.createElement('option');
                opt.value = v;
                opt.textContent = v;
                sel.appendChild(opt);
            });
            sel.addEventListener('change', applyFilters);
        }

        populateFilter('filterServer', 1);  // Server ist Spalte 1 (0-basiert: Zeit=0, Server=1, Level=2, Quelle=3, Fehler=4)
        populateFilter('filterQuelle', 3);  // Quelle ist Spalte 3

        // Sortieren per Klick auf th
        table.querySelectorAll('th').forEach(function (th, col) {
            th.addEventListener('click', function () {
                const isAsc = th.classList.contains('asc');
                table.querySelectorAll('th').forEach(function (h) { h.classList.remove('asc', 'desc'); });
                th.classList.add(isAsc ? 'desc' : 'asc');

                rows.sort(function (a, b) {
                    const aVal = (a.cells[col]?.textContent || '').trim();
                    const bVal = (b.cells[col]?.textContent || '').trim();
                    const aNum = parseFloat(aVal.replace(',', '.'));
                    const bNum = parseFloat(bVal.replace(',', '.'));
                    if (!isNaN(aNum) && !isNaN(bNum)) return isAsc ? bNum - aNum : aNum - bNum;
                    return isAsc ? bVal.localeCompare(aVal) : aVal.localeCompare(bVal);
                });
                rows.forEach(function (r) { table.appendChild(r); });
            });
        });

        // Kombinierte Filterfunktion
        function applyFilters() {
            const q = searchBox.value.toLowerCase();
            const sv = document.getElementById('filterServer').value;
            const qv = document.getElementById('filterQuelle').value;

            rows.forEach(function (r) {
                const cells = Array.from(r.cells);
                const textMatch = q === '' || cells.some(function (c) { return c.textContent.toLowerCase().includes(q); });
                const serverMatch = sv === '' || (cells[1]?.textContent || '').trim() === sv;
                const quelleMatch = qv === '' || (cells[3]?.textContent || '').trim() === qv;
                r.classList.toggle('hidden', !(textMatch && serverMatch && quelleMatch));
            });
        }

        searchBox.addEventListener('input', applyFilters);
    });
</script>
"@

$HtmlBody = $Results | Sort-Object Time -Descending | Select-Object @{N='Zeit';E={
    if ($_.Time) { $_.Time.ToString('dd.MM.yyyy HH:mm:ss') } else { '-' }
}}, @{N='Server';E={ $_.Server }}, @{N='Level';E={ $_.Level }}, @{N='Quelle';E={ $_.Provider }}, @{N='Fehler';E={
    $_.Message
}} | ConvertTo-Html -Fragment

$Html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Application Error / Hang / Popup - Report</title>
    $HtmlHead
</head>
<body>
    <h1>Application-Fehler Report</h1>
    <div class="summary">
        Erstellt: $(Get-Date -Format 'dd.MM.yyyy HH:mm') |
        Zeitraum: $($StartTime.ToString('dd.MM.yyyy')) - $(Get-Date -Format 'dd.MM.yyyy') |
        Server: $($Computers.Count) | Einträge: $($Results.Count)
    </div>
    <div class="filterBar" id="filterBar">
        <div><label for="filterServer">Server:</label><select id="filterServer"><option value="">Alle</option></select></div>
        <div><label for="filterQuelle">Quelle:</label><select id="filterQuelle"><option value="">Alle</option></select></div>
        <div><label for="searchBox">Suche:</label><input type="text" id="searchBox" placeholder="Beliebiger Text ..." /></div>
    </div>
    $HtmlBody
</body>
</html>
"@

# Region: Ausgabe
try {
    $OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    $Html | Out-File -FilePath $OutputPath -Encoding utf8 -Force
    Write-Host "Report erstellt: $OutputPath" -ForegroundColor Green

    # Öffnen im Browser
    if ((Get-Item $OutputPath).Length -gt 0) {
        Start-Process $OutputPath
    }
}
catch {
    Write-Error "Fehler beim Schreiben der Datei: $_"
    exit 1
}
