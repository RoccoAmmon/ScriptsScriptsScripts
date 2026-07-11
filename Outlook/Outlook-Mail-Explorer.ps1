<#
==================================================================================
 SCRIPT   : Outlook-Mail-Explorer.ps1
 VERSION  : 1.2
 AUTOR    : Rocco Ammon
 DATUM    : 2026-07-12
==================================================================================
 BESCHREIBUNG:
   Grafischer Outlook-Mail-Explorer (WinForms) mit Filter nach Suchwort, Datumsbereich
   und Anhang. Durchsucht ein oder mehrere Postfächer (auch alle gleichzeitig)
   und listet Treffer in einer sortierbaren Liste. Pro Mail: Vorschau des
   Inhalts, Anhänge öffnen/speichern, Weiterleitung (ganze Mail oder nur
   Anhänge). Prüft, ob eine Mail bereits an die eingetragene Adresse
   weitergeleitet wurde und zeigt dies in der Trefferliste an.

 FEATURES:
   - Suche über alle oder ausgewählte Postfächer
   - Filter: Suchwort (Betreff + Text), Datumsbereich, nur mit Anhang
   - Ergebnisliste mit Datum, Absender, Betreff, Anhang, Weiterleitungsstatus, Postfach
   - Sortierung per Klick auf Spaltenüberschrift
   - Mehrfachauswahl (Strg+Klick) für Sammel-Weiterleitung
   - Vorschau des Mail-Inhalts bei Klick
   - Anhänge öffnen (Doppelklick oder Button)
   - Weiterleitung: ganze Mail oder nur Anhänge (umschaltbar)
   - PDF-Filter: nur PDF-Anhänge bei Weiterleitung (optional)
   - Navigations-Buttons (▲/▼) zum Durchblättern der Treffer
   - Weiterleitungs-Historie (erkennt doppelte Sendungen)
   - Live-Update der Weiterleitungs-Spalte bei Adressänderung
   - Auto-Update der Weiterl.-Spalte nach dem Weiterleiten
   - Suchabbruch jederzeit möglich
   - Fortschrittsanzeige mit ProgressBar
   - Logging nach C:\ScriptLog\

 HINWEIS:
   Benötigt eine lokal installierte und konfigurierte Outlook-Instanz
   (COM-Automatisierung). Logging erfolgt unter C:\ScriptLog.
==================================================================================
#>

#Region ================== VARIABLEN / KONFIGURATION ==================
# Zentrale Konfiguration - hier bei Bedarf anpassen
$LogVerzeichnis        = "C:\ScriptLog"                                   # Log-Ablage (Standard)
$LogDatei              = Join-Path $LogVerzeichnis "Outlook_Mailsuche.log"
$StandardSuchwort      = "Rechnung"                                       # Vorbelegung Suchfeld
$StandardWeiterleitung = "empfaenger@firma.de"                            # Vorbelegung Zieladresse
$UpdateIntervall       = 25                                               # GUI-Update alle X Mails
$TempAnhangPfad        = Join-Path $env:TEMP "OutlookMailsuche_Anhaenge"  # Temp-Ordner für Anhänge
$FensterSkalierung     = 0.85                                             # 85 % der Bildschirmgröße

# Globale Sammel-Variable für Treffer (Mail-Objekte referenzieren)
$Global:GefundeneMails = @{}

# Steuerflag: wird auf $true gesetzt, um eine laufende Suche abzubrechen
$Global:SucheAbbrechen = $false
#EndRegion


#Region ================== VORBEREITUNG / ASSEMBLIES ==================
try {
    # Log-Verzeichnis anlegen, falls nicht vorhanden
    if (-not (Test-Path -Path $LogVerzeichnis)) {
        New-Item -Path $LogVerzeichnis -ItemType Directory -Force | Out-Null
    }
    # Temp-Ordner für Anhänge anlegen
    if (-not (Test-Path -Path $TempAnhangPfad)) {
        New-Item -Path $TempAnhangPfad -ItemType Directory -Force | Out-Null
    }

    # Windows-Forms-Bibliotheken laden (für GUI)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # IComparer für ListView-Sortierung per Spaltenklick
    Add-Type @"
    using System;
    using System.Collections;
    using System.Windows.Forms;

    public class ListViewColumnSorter : IComparer {
        public int Column { get; set; }
        public SortOrder SortDirection { get; set; }

        public ListViewColumnSorter() {
            Column = 0;
            SortDirection = SortOrder.Ascending;
        }

        public ListViewColumnSorter(int column, SortOrder direction) {
            Column = column;
            SortDirection = direction;
        }

        public int Compare(object x, object y) {
            ListViewItem itemX = (ListViewItem)x;
            ListViewItem itemY = (ListViewItem)y;
            string textX = itemX.SubItems[Math.Min(Column, itemX.SubItems.Count - 1)].Text;
            string textY = itemY.SubItems[Math.Min(Column, itemY.SubItems.Count - 1)].Text;

            // Datum-Sortierung für Spalte 0
            if (Column == 0) {
                DateTime dateX, dateY;
                if (DateTime.TryParse(textX, out dateX) && DateTime.TryParse(textY, out dateY)) {
                    return SortDirection == SortOrder.Ascending
                        ? DateTime.Compare(dateX, dateY) : DateTime.Compare(dateY, dateX);
                }
            }

            // Numerische Sortierung
            double dblX, dblY;
            if (double.TryParse(textX, out dblX) && double.TryParse(textY, out dblY)) {
                return SortDirection == SortOrder.Ascending
                    ? dblX.CompareTo(dblY) : dblY.CompareTo(dblX);
            }

            // Standard: String-Vergleich
            int strResult = string.Compare(textX, textY, StringComparison.CurrentCultureIgnoreCase);
            return SortDirection == SortOrder.Ascending ? strResult : -strResult;
        }
    }
"@ -ReferencedAssemblies System.Windows.Forms
}
catch {
    Write-Host "Kritischer Fehler bei der Vorbereitung: $($_.Exception.Message)" -ForegroundColor Red
    return
}
#EndRegion


#Region ================== HILFSFUNKTIONEN ==================

<#
.SYNOPSIS
    Schreibt einen Eintrag in die Log-Datei unter C:\ScriptLog.
.PARAMETER Text
    Der zu protokollierende Text.
.PARAMETER Level
    Log-Level: INFO, WARN, ERROR, STATUS.
#>
function Write-Log {
    param (
        [Parameter(Mandatory = $true)][string] $Text,
        [ValidateSet("INFO","WARN","ERROR","STATUS")][string] $Level = "INFO"
    )
    try {
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Text" |
            Out-File -FilePath $LogDatei -Append -Encoding UTF8
    }
    catch {
        # Logging darf niemals das Script abbrechen
        Write-Host "Logging fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

<#
.SYNOPSIS
    Aktualisiert die Fortschrittsanzeige in der GUI und schreibt einen Log-Eintrag.
.PARAMETER Statustext
    Der anzuzeigende Statustext.
.PARAMETER Prozent
    Fortschritt in Prozent (0-100).
#>
function Update-Status {
    param (
        [Parameter(Mandatory = $true)][string] $Statustext,
        [Parameter(Mandatory = $true)][ValidateRange(0,100)][int] $Prozent
    )
    try {
        $lblStatus.Text    = $Statustext
        $progressBar.Value = $Prozent
        [System.Windows.Forms.Application]::DoEvents()   # GUI reaktionsfähig halten
        Write-Log -Text "$Statustext ($Prozent%)" -Level STATUS
    }
    catch {
        Write-Log -Text "Status-Update fehlgeschlagen: $($_.Exception.Message)" -Level WARN
    }
}

<#
.SYNOPSIS
    Liefert das aktuell in der Ergebnisliste ausgewählte Original-MailItem
    (aufgelöst über den im .Tag gespeicherten GUID-Schlüssel).
#>
function Get-AusgewaehlteMail {
    if ($ergebnisListe.SelectedItems.Count -eq 0) { return $null }
    $key = $ergebnisListe.SelectedItems[0].Tag
    return $Global:GefundeneMails[$key]
}

#EndRegion


#Region ================== OUTLOOK-VERBINDUNG ==================
try {
    Write-Log -Text "Script gestartet. Verbinde mit Outlook..." -Level INFO

    # Outlook-COM-Objekt erzeugen und MAPI-Namespace holen
    $outlook   = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")

    # Alle verfügbaren Postfächer (Stores) ermitteln
    $postfaecher = @()
    foreach ($store in $namespace.Stores) {
        $postfaecher += $store
    }

    if ($postfaecher.Count -eq 0) {
        throw "Es wurden keine Outlook-Postfächer gefunden."
    }
    Write-Log -Text "$($postfaecher.Count) Postfach/Postfächer gefunden." -Level INFO
}
catch {
    $msg = "Verbindung zu Outlook fehlgeschlagen: $($_.Exception.Message)"
    Write-Log -Text $msg -Level ERROR
    [System.Windows.Forms.MessageBox]::Show($msg,"Fehler",'OK','Error')
    return
}
#EndRegion


#Region ================== GUI-AUFBAU (DYNAMISCH, 85% BILDSCHIRM) ==================

# --- Bildschirmgröße ermitteln und Fenster auf 85 % skalieren ---
$screen      = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$fensterBrt  = [int]($screen.Width  * $FensterSkalierung)
$fensterHoe  = [int]($screen.Height * $FensterSkalierung)

# --- Hauptfenster ---
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Outlook-Mailsuche"
$form.Size            = New-Object System.Drawing.Size($fensterBrt, $fensterHoe)
$form.MinimumSize     = New-Object System.Drawing.Size(760, 640)   # sinnvolle Untergrenze
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "Sizable"     # Fenster darf frei skaliert werden
$form.MaximizeBox     = $true
$form.AutoScaleMode   = 'Font'
$form.Font            = New-Object System.Drawing.Font("Tahoma", 10, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)

# --- Label + Textbox: Suchwort (oben, feste Position) ---
$lblSuchwort = New-Object System.Windows.Forms.Label
$lblSuchwort.Location = New-Object System.Drawing.Point(15, 15)
$lblSuchwort.Size     = New-Object System.Drawing.Size(120, 20)
$lblSuchwort.Text     = "Suchwort:"
$form.Controls.Add($lblSuchwort)

$txtSuchwort = New-Object System.Windows.Forms.TextBox
$txtSuchwort.Location = New-Object System.Drawing.Point(140, 13)
$txtSuchwort.Size     = New-Object System.Drawing.Size(200, 20)
$txtSuchwort.Text     = $StandardSuchwort
$form.Controls.Add($txtSuchwort)

# --- Checkbox: nur Mails mit Anhang ---
$chkNurAnhang = New-Object System.Windows.Forms.CheckBox
$chkNurAnhang.Location = New-Object System.Drawing.Point(360, 13)
$chkNurAnhang.Size     = New-Object System.Drawing.Size(200, 20)
$chkNurAnhang.Text     = "Nur Mails mit Anhang"
$chkNurAnhang.Checked  = $true
$form.Controls.Add($chkNurAnhang)

# --- Datumsbereich: von / bis ---
$lblVon = New-Object System.Windows.Forms.Label
$lblVon.Location = New-Object System.Drawing.Point(15, 48)
$lblVon.Size     = New-Object System.Drawing.Size(120, 20)
$lblVon.Text     = "Datum von:"
$form.Controls.Add($lblVon)

$dtpVon = New-Object System.Windows.Forms.DateTimePicker
$dtpVon.Location = New-Object System.Drawing.Point(140, 46)
$dtpVon.Size     = New-Object System.Drawing.Size(200, 20)
$dtpVon.Format   = "Short"
$dtpVon.Value    = (Get-Date).AddMonths(-1)          # Vorbelegung: letzter Monat
$form.Controls.Add($dtpVon)

$lblBis = New-Object System.Windows.Forms.Label
$lblBis.Location = New-Object System.Drawing.Point(360, 48)
$lblBis.Size     = New-Object System.Drawing.Size(70, 20)
$lblBis.Text     = "Datum bis:"
$form.Controls.Add($lblBis)

$dtpBis = New-Object System.Windows.Forms.DateTimePicker
$dtpBis.Location = New-Object System.Drawing.Point(435, 46)
$dtpBis.Size     = New-Object System.Drawing.Size(200, 20)
$dtpBis.Format   = "Short"
$dtpBis.Value    = (Get-Date)                        # Vorbelegung: heute
$form.Controls.Add($dtpBis)

# --- Label + Liste: Postfach-Auswahl ---
$lblPostfach = New-Object System.Windows.Forms.Label
$lblPostfach.Location = New-Object System.Drawing.Point(15, 82)
$lblPostfach.Size     = New-Object System.Drawing.Size(300, 20)
$lblPostfach.Text     = "Postfächer (Mehrfachauswahl möglich):"
$form.Controls.Add($lblPostfach)

$lstPostfaecher = New-Object System.Windows.Forms.CheckedListBox
$lstPostfaecher.Location    = New-Object System.Drawing.Point(15, 105)
$lstPostfaecher.Size        = New-Object System.Drawing.Size(320, 90)
$lstPostfaecher.CheckOnClick = $true
foreach ($pf in $postfaecher) {
    [void]$lstPostfaecher.Items.Add($pf.DisplayName)
}
$form.Controls.Add($lstPostfaecher)

# --- Checkbox: alle Postfächer durchsuchen ---
$chkAllePostfaecher = New-Object System.Windows.Forms.CheckBox
$chkAllePostfaecher.Location = New-Object System.Drawing.Point(345, 105)
$chkAllePostfaecher.Size     = New-Object System.Drawing.Size(250, 20)
$chkAllePostfaecher.Text     = "Alle Postfächer durchsuchen"
$form.Controls.Add($chkAllePostfaecher)

# Alle Postfächer-Checkbox aktiviert/deaktiviert die Einzelliste
$chkAllePostfaecher.Add_CheckedChanged({
    $lstPostfaecher.Enabled = -not $chkAllePostfaecher.Checked
})

# --- Weiterleitungs-Adresse ---
$lblZiel = New-Object System.Windows.Forms.Label
$lblZiel.Location = New-Object System.Drawing.Point(345, 135)
$lblZiel.Size     = New-Object System.Drawing.Size(250, 20)
$lblZiel.Text     = "Weiterleiten an (E-Mail):"
$form.Controls.Add($lblZiel)

$txtZiel = New-Object System.Windows.Forms.TextBox
$txtZiel.Location = New-Object System.Drawing.Point(345, 158)
$txtZiel.Size     = New-Object System.Drawing.Size(320, 20)
$txtZiel.Text     = $StandardWeiterleitung
$form.Controls.Add($txtZiel)

# --- Bei Änderung der Ziel-Adresse: Weiterleitungs-Spalte aktualisieren ---
$txtZiel.Add_TextChanged({
    $ziel = $txtZiel.Text.Trim()
    $histDatei = Join-Path $LogVerzeichnis "Outlook_Weiterleitungen.json"
    $hist = @()
    if (Test-Path $histDatei) {
        try {
            $hist = @(Get-Content -Path $histDatei -Raw -Encoding UTF8 | ConvertFrom-Json)
        } catch { $hist = @() }
    }
    foreach ($item in $ergebnisListe.Items) {
        $key = $item.Tag
        $mail = $Global:GefundeneMails[$key]
        if (-not $mail) { continue }
        $weg = ""
        if ($mail.EntryID) {
            foreach ($h in $hist) {
                if ($h.MailId -eq $mail.EntryID -and $h.Empfaenger -eq $ziel) {
                    $weg = "Ja"
                    break
                }
            }
        }
        if ($item.SubItems.Count -ge 5) { $item.SubItems[4].Text = $weg }
    }
})

# --- Such-Button ---
$btnSuchen = New-Object System.Windows.Forms.Button
$btnSuchen.Location  = New-Object System.Drawing.Point(345, 185)
$btnSuchen.Size      = New-Object System.Drawing.Size(150, 30)
$btnSuchen.Text      = "Suche starten"
$btnSuchen.BackColor = [System.Drawing.Color]::LightSteelBlue
$form.Controls.Add($btnSuchen)

# --- Abbrechen-Button (nur während laufender Suche aktiv) ---
$btnAbbrechen = New-Object System.Windows.Forms.Button
$btnAbbrechen.Location  = New-Object System.Drawing.Point(505, 185)
$btnAbbrechen.Size      = New-Object System.Drawing.Size(160, 30)
$btnAbbrechen.Text      = "Suche abbrechen"
$btnAbbrechen.BackColor = [System.Drawing.Color]::LightSalmon
$btnAbbrechen.Enabled   = $false
$form.Controls.Add($btnAbbrechen)

# --- Ergebnisliste (ListView) - wächst mit dem Fenster ---
$lblErgebnis = New-Object System.Windows.Forms.Label
$lblErgebnis.Location = New-Object System.Drawing.Point(15, 225)
$lblErgebnis.Size     = New-Object System.Drawing.Size(300, 20)
$lblErgebnis.Text     = "Ergebnisse:"
$form.Controls.Add($lblErgebnis)

$ergebnisListe = New-Object System.Windows.Forms.ListView
$ergebnisListe.Location      = New-Object System.Drawing.Point(15, 248)
$ergebnisListe.Size          = New-Object System.Drawing.Size(($fensterBrt - 55), ($fensterHoe - 565))
$ergebnisListe.View          = 'Details'
$ergebnisListe.FullRowSelect = $true
$ergebnisListe.GridLines     = $true
$ergebnisListe.MultiSelect   = $true
# Anker: oben+links+rechts fixiert, Höhe wächst mit -> Top,Bottom,Left,Right
$ergebnisListe.Anchor        = 'Top','Bottom','Left','Right'
[void]$ergebnisListe.Columns.Add("Datum", 120)
[void]$ergebnisListe.Columns.Add("Absender", 200)
[void]$ergebnisListe.Columns.Add("Betreff", 320)
[void]$ergebnisListe.Columns.Add("Anhang", 70)
[void]$ergebnisListe.Columns.Add("Weiterl.", 65)
[void]$ergebnisListe.Columns.Add("Postfach", 150)
$form.Controls.Add($ergebnisListe)

# --- Sortierung per Spaltenklick ---
$ergebnisListe.Add_ColumnClick({
    param($sender, $e)
    $sorter = $sender.ListViewItemSorter
    if ($sorter -and $sorter.Column -eq $e.Column) {
        # Gleiche Spalte: Richtung umkehren
        $sorter.SortDirection = if ($sorter.SortDirection -eq 'Ascending') { 'Descending' } else { 'Ascending' }
    } else {
        # Neue Spalte: aufsteigend sortieren
        $sender.ListViewItemSorter = New-Object ListViewColumnSorter($e.Column, 'Ascending')
    }
    $sender.Sort()
})

# --- Navigations-Buttons (vor/zurück in der Trefferliste) ---
$btnVorherige = New-Object System.Windows.Forms.Button
$btnVorherige.Location = New-Object System.Drawing.Point(15, ($fensterHoe - 300))
$btnVorherige.Size     = New-Object System.Drawing.Size(30, 28)
$btnVorherige.Text     = "▲"
$btnVorherige.Anchor   = 'Bottom','Left'
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($btnVorherige, "Vorherige Mail")
$form.Controls.Add($btnVorherige)

$btnNaechste = New-Object System.Windows.Forms.Button
$btnNaechste.Location = New-Object System.Drawing.Point(50, ($fensterHoe - 300))
$btnNaechste.Size     = New-Object System.Drawing.Size(30, 28)
$btnNaechste.Text     = "▼"
$btnNaechste.Anchor   = 'Bottom','Left'
$toolTip.SetToolTip($btnNaechste, "Nächste Mail")
$form.Controls.Add($btnNaechste)

# --- Aktions-Buttons unter der Liste (am unteren Rand verankert) ---
$btnInhalt = New-Object System.Windows.Forms.Button
$btnInhalt.Location = New-Object System.Drawing.Point(95, ($fensterHoe - 300))
$btnInhalt.Size     = New-Object System.Drawing.Size(150, 28)
$btnInhalt.Text     = "Inhalt anzeigen"
$btnInhalt.Anchor   = 'Bottom','Left'
$form.Controls.Add($btnInhalt)

$btnAnhang = New-Object System.Windows.Forms.Button
$btnAnhang.Location = New-Object System.Drawing.Point(255, ($fensterHoe - 300))
$btnAnhang.Size     = New-Object System.Drawing.Size(150, 28)
$btnAnhang.Text     = "Anhang öffnen"
$btnAnhang.Anchor   = 'Bottom','Left'
$form.Controls.Add($btnAnhang)

$btnWeiterleiten = New-Object System.Windows.Forms.Button
$btnWeiterleiten.Location  = New-Object System.Drawing.Point(415, ($fensterHoe - 300))
$btnWeiterleiten.Size      = New-Object System.Drawing.Size(180, 28)
$btnWeiterleiten.Text      = "Ausgewählte weiterleiten"
$btnWeiterleiten.BackColor = [System.Drawing.Color]::PaleGreen
$btnWeiterleiten.Anchor    = 'Bottom','Left'
$form.Controls.Add($btnWeiterleiten)

# --- Checkbox: Nur Anhang weiterleiten (neben dem Weiterleiten-Button) ---
$chkNurAnhangWeiterleiten = New-Object System.Windows.Forms.CheckBox
$chkNurAnhangWeiterleiten.Location = New-Object System.Drawing.Point(620, ($fensterHoe - 298))
$chkNurAnhangWeiterleiten.Size     = New-Object System.Drawing.Size(170, 22)
$chkNurAnhangWeiterleiten.Text     = "Nur Anhang weiterleiten"
$chkNurAnhangWeiterleiten.Checked  = $false
$chkNurAnhangWeiterleiten.Anchor   = 'Bottom','Left'
$form.Controls.Add($chkNurAnhangWeiterleiten)

# --- Checkbox: Nur PDFs weiterleiten (nur aktiv wenn "Nur Anhang" aktiv) ---
$chkNurPDF = New-Object System.Windows.Forms.CheckBox
$chkNurPDF.Location = New-Object System.Drawing.Point(620, ($fensterHoe - 274))
$chkNurPDF.Size     = New-Object System.Drawing.Size(170, 22)
$chkNurPDF.Text     = "Nur PDF-Anhänge"
$chkNurPDF.Checked  = $true
$chkNurPDF.Anchor   = 'Bottom','Left'
$chkNurPDF.Enabled  = $chkNurAnhangWeiterleiten.Checked
$form.Controls.Add($chkNurPDF)

# Nur-PDF-Checkbox deaktivieren wenn "Nur Anhang" ausgehakt ist
$chkNurAnhangWeiterleiten.Add_CheckedChanged({
    $chkNurPDF.Enabled = $chkNurAnhangWeiterleiten.Checked
})

# --- Fortschrittsanzeige (unten verankert) ---
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(15, ($fensterHoe - 265))
$lblStatus.Size      = New-Object System.Drawing.Size(($fensterBrt - 55), 20)
$lblStatus.Text      = "Bereit."
$lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
$lblStatus.Anchor    = 'Bottom','Left','Right'
$form.Controls.Add($lblStatus)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(15, ($fensterHoe - 243))
$progressBar.Size     = New-Object System.Drawing.Size(($fensterBrt - 55), 22)
$progressBar.Minimum  = 0
$progressBar.Maximum  = 100
$progressBar.Style    = 'Continuous'
$progressBar.Anchor   = 'Bottom','Left','Right'
$form.Controls.Add($progressBar)

# --- Vorschau-Textbox (Mailinhalt) - links, unten verankert ---
$txtVorschau = New-Object System.Windows.Forms.TextBox
$txtVorschau.Location   = New-Object System.Drawing.Point(15, ($fensterHoe - 190))
$txtVorschau.Size       = New-Object System.Drawing.Size(($fensterBrt - 245), 150)
$txtVorschau.Multiline  = $true
$txtVorschau.ScrollBars = 'Vertical'
$txtVorschau.ReadOnly   = $true
$txtVorschau.Anchor     = 'Bottom','Left','Right'
$form.Controls.Add($txtVorschau)

# --- Label + Liste: Anhänge (rechts neben der Vorschau) ---
$lblAnhaenge = New-Object System.Windows.Forms.Label
$lblAnhaenge.Location = New-Object System.Drawing.Point(($fensterBrt - 220), ($fensterHoe - 215))
$lblAnhaenge.Size     = New-Object System.Drawing.Size(180, 25)
$lblAnhaenge.Text     = "Anhänge:"
$lblAnhaenge.Font     = New-Object System.Drawing.Font($lblAnhaenge.Font, 'Bold')
$lblAnhaenge.Anchor   = 'Bottom','Right'
$form.Controls.Add($lblAnhaenge)

$lstAnhaenge = New-Object System.Windows.Forms.ListBox
$lstAnhaenge.Location = New-Object System.Drawing.Point(($fensterBrt - 220), ($fensterHoe - 190))
$lstAnhaenge.Size     = New-Object System.Drawing.Size(180, 150)
$lstAnhaenge.Anchor   = 'Bottom','Right'
$form.Controls.Add($lstAnhaenge)

#EndRegion


#Region ================== SUCHLOGIK ==================
$btnSuchen.Add_Click({
    try {
        # --- Eingaben auslesen ---
        $suchwort  = $txtSuchwort.Text.Trim()
        $datumVon  = $dtpVon.Value.Date
        $datumBis  = $dtpBis.Value.Date.AddDays(1).AddSeconds(-1)   # bis Tagesende
        $nurAnhang = $chkNurAnhang.Checked

        if ([string]::IsNullOrWhiteSpace($suchwort)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte ein Suchwort eingeben.","Hinweis",'OK','Warning')
            return
        }
        if ($datumVon -gt $datumBis) {
            [System.Windows.Forms.MessageBox]::Show("Das Start-Datum liegt nach dem End-Datum.","Hinweis",'OK','Warning')
            return
        }

        # --- Zu durchsuchende Postfächer bestimmen ---
        if ($chkAllePostfaecher.Checked) {
            $zuDurchsuchen = $postfaecher
        }
        else {
            $zuDurchsuchen = @()
            foreach ($idx in $lstPostfaecher.CheckedIndices) {
                $zuDurchsuchen += $postfaecher[$idx]
            }
        }

        if ($zuDurchsuchen.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte mindestens ein Postfach auswählen (oder 'Alle').","Hinweis",'OK','Warning')
            return
        }

        # --- Vorbereitung ---
        $ergebnisListe.Items.Clear()
        $Global:GefundeneMails.Clear()
        $txtVorschau.Clear()
        $lstAnhaenge.Items.Clear()

        # Abbruch-Flag zuruecksetzen und Buttons umschalten
        $Global:SucheAbbrechen = $false
        $btnSuchen.Enabled     = $false     # Doppelklick verhindern
        $btnAbbrechen.Enabled  = $true       # Abbrechen freigeben

        # Weiterleitungs-Historie für Prüfung laden (einmalig)
        $weiterleitungsHistorie = @()
        $weiterleitungsDatei = Join-Path $LogVerzeichnis "Outlook_Weiterleitungen.json"
        try {
            if (Test-Path $weiterleitungsDatei) {
                $weiterleitungsHistorie = Get-Content -Path $weiterleitungsDatei -Raw -Encoding UTF8 | ConvertFrom-Json
            }
        } catch { Write-Log -Text "Fehler beim Laden der Weiterleitungs-Historie: $($_.Exception.Message)" -Level WARN }

        Update-Status -Statustext "Suche wird gestartet..." -Prozent 0
        Write-Log -Text "Suche gestartet. Wort='$suchwort', Von=$datumVon, Bis=$datumBis, NurAnhang=$nurAnhang" -Level INFO

        $anzahlPostfaecher = $zuDurchsuchen.Count
        $postfachIndex     = 0
        $trefferGesamt     = 0

        foreach ($postfach in $zuDurchsuchen) {

            # --- Abbruch auf Postfach-Ebene pruefen ---
            if ($Global:SucheAbbrechen) { break }
            $postfachIndex++

            try {
                # Posteingang des jeweiligen Postfachs holen
                $rootFolder = $postfach.GetRootFolder()
                $mails      = $rootFolder.Folders | Where-Object { $_.Name -match 'Posteingang|Inbox' } | Select-Object -First 1

                if (-not $mails) {
                    # Fallback: Root-Ordner selbst verwenden
                    $items = $rootFolder.Items
                }
                else {
                    $items = $mails.Items
                }

                # Nach Empfangszeit sortieren (Performance beim Durchlauf)
                $items.Sort("[ReceivedTime]", $true)

                $gesamtMails = $items.Count
                $mailIndex   = 0

                Update-Status -Statustext "Postfach $postfachIndex von $anzahlPostfaecher : '$($postfach.DisplayName)' wird durchsucht..." `
                              -Prozent ([int](($postfachIndex - 1) / $anzahlPostfaecher * 100))

                foreach ($mail in $items) {

                    # --- Abbruch auf Mail-Ebene pruefen ---
                    if ($Global:SucheAbbrechen) { break }
                    [System.Windows.Forms.Application]::DoEvents()   # Abbrechen-Klick verarbeiten

                    $mailIndex++

                    # --- Fortschritt regelmäßig aktualisieren ---
                    if (($mailIndex % $UpdateIntervall) -eq 0 -or $mailIndex -eq $gesamtMails) {
                        $prozentGesamt = [int]( ( ($postfachIndex - 1) + ($mailIndex / [math]::Max($gesamtMails,1)) ) / $anzahlPostfaecher * 100 )
                        if ($prozentGesamt -gt 100) { $prozentGesamt = 100 }
                        Update-Status -Statustext ("Postfach '{0}' ({1}/{2}) - Mail {3} von {4} - Treffer: {5}" -f `
                                      $postfach.DisplayName, $postfachIndex, $anzahlPostfaecher, $mailIndex, $gesamtMails, $trefferGesamt) `
                                      -Prozent $prozentGesamt
                    }

                    try {
                        # Nur echte E-Mail-Objekte prüfen (MailItem = Class 43)
                        if ($mail.Class -ne 43) { continue }

                        # --- Datumsfilter ---
                        $empfangen = $mail.ReceivedTime
                        if ($empfangen -lt $datumVon -or $empfangen -gt $datumBis) { continue }

                        # --- Anhang-Filter ---
                        $hatAnhang = ($mail.Attachments.Count -gt 0)
                        if ($nurAnhang -and -not $hatAnhang) { continue }

                        # --- Suchwort-Filter (Betreff ODER Text) ---
                        $betreff = if ($mail.Subject) { $mail.Subject } else { "" }
                        $body    = if ($mail.Body)    { $mail.Body }    else { "" }
                        if ($betreff -notmatch [regex]::Escape($suchwort) -and
                            $body    -notmatch [regex]::Escape($suchwort)) { continue }

                        # --- Prüfen ob bereits weitergeleitet ---
                        $bereitsWeg = ""
                        $zielAdr = $txtZiel.Text.Trim()
                        if ($zielAdr -match '^[^@\s]+@[^@\s]+\.[^@\s]+$' -and $mail.EntryID) {
                            foreach ($eh in $weiterleitungsHistorie) {
                                if ($eh.MailId -eq $mail.EntryID -and $eh.Empfaenger -eq $zielAdr) {
                                    $bereitsWeg = "Ja"
                                    break
                                }
                            }
                        }

                        # --- Treffer in Liste eintragen ---
                        $item = New-Object System.Windows.Forms.ListViewItem($empfangen.ToString("dd.MM.yyyy HH:mm"))
                        [void]$item.SubItems.Add([string]$mail.SenderName)
                        [void]$item.SubItems.Add([string]$betreff)
                        [void]$item.SubItems.Add($(if ($hatAnhang) { "Ja" } else { "Nein" }))
                        [void]$item.SubItems.Add($bereitsWeg)
                        [void]$item.SubItems.Add([string]$postfach.DisplayName)

                        # Eindeutigen Schlüssel vergeben und Mail-Objekt merken
                        $key = [Guid]::NewGuid().ToString()
                        $item.Tag = $key
                        $Global:GefundeneMails[$key] = $mail

                        [void]$ergebnisListe.Items.Add($item)
                        $trefferGesamt++
                    }
                    catch {
                        Write-Log -Text "Fehler beim Prüfen einer Mail: $($_.Exception.Message)" -Level WARN
                        continue
                    }
                }
            }
            catch {
                Write-Log -Text "Fehler beim Zugriff auf Postfach '$($postfach.DisplayName)': $($_.Exception.Message)" -Level ERROR
                continue
            }
        }

        # --- Abschlussmeldung je nach Abbruch-Status ---
        if ($Global:SucheAbbrechen) {
            Update-Status -Statustext "Suche abgebrochen. $trefferGesamt Treffer bis zum Abbruch." -Prozent 100
            Write-Log -Text "Suche abgebrochen nach $trefferGesamt Treffern." -Level WARN
        }
        else {
            Update-Status -Statustext "Suche abgeschlossen. $trefferGesamt Treffer gefunden." -Prozent 100
            Write-Log -Text "Suche abgeschlossen. $trefferGesamt Treffer." -Level INFO
        }
    }
    catch {
        $fehlerText = "Fehler während der Suche: $($_.Exception.Message)"
        Update-Status -Statustext $fehlerText -Prozent 0
        Write-Log -Text $fehlerText -Level ERROR
        [System.Windows.Forms.MessageBox]::Show($fehlerText,"Fehler",'OK','Error')
    }
    finally {
        # Buttons IMMER zuruecksetzen - auch bei Fehler oder Abbruch
        $btnSuchen.Enabled     = $true
        $btnAbbrechen.Enabled  = $false
        $Global:SucheAbbrechen = $false
    }
})
#EndRegion


#Region ================== ABBRECHEN ==================
$btnAbbrechen.Add_Click({
    $Global:SucheAbbrechen = $true
    Update-Status -Statustext "Abbruch angefordert - bitte warten..." -Prozent $progressBar.Value
    Write-Log -Text "Suche wurde vom Benutzer abgebrochen." -Level WARN
})
#EndRegion


#Region ================== VORSCHAU / ANHANG / WEITERLEITEN ==================

# --- Vorschau + Anhangsliste automatisch bei Auswahl anzeigen ---
$aktionVorschau = {
    try {
        $mail = Get-AusgewaehlteMail
        # Felder immer erst leeren, damit keine alten Daten stehen bleiben
        $txtVorschau.Clear()
        $lstAnhaenge.Items.Clear()

        if (-not $mail) { return }

        # --- Mailinhalt in die Vorschau schreiben ---
        $txtVorschau.Text = "Von: $($mail.SenderName)`r`n" +
                            "Betreff: $($mail.Subject)`r`n" +
                            "Empfangen: $($mail.ReceivedTime)`r`n" +
                            "--------------------------------------------------`r`n" +
                            "$($mail.Body)"

        # --- Anhangsnamen rechts auflisten ---
        if ($mail.Attachments.Count -gt 0) {
            foreach ($att in $mail.Attachments) {
                [void]$lstAnhaenge.Items.Add($att.FileName)
            }
        }
        else {
            [void]$lstAnhaenge.Items.Add("(keine Anhänge)")
        }

        Write-Log -Text "Vorschau geladen: '$($mail.Subject)'" -Level INFO
    }
    catch {
        Write-Log -Text "Fehler bei der Vorschau: $($_.Exception.Message)" -Level ERROR
    }
}

# Vorschau sofort bei Auswahl (Klick / Pfeiltasten) - und per Button verfügbar
$ergebnisListe.Add_SelectedIndexChanged($aktionVorschau)
$btnInhalt.Add_Click($aktionVorschau)

# --- Navigation: Vorherige / Nächste Mail ---
$btnVorherige.Add_Click({
    if ($ergebnisListe.SelectedItems.Count -gt 0 -and $ergebnisListe.Items.Count -gt 0) {
        $idx = $ergebnisListe.Items.IndexOf($ergebnisListe.SelectedItems[0])
        if ($idx -gt 0) {
            $ergebnisListe.Items[$idx].Selected = $false
            $ergebnisListe.Items[$idx - 1].Selected = $true
            $ergebnisListe.Items[$idx - 1].Focused = $true
            $ergebnisListe.EnsureVisible($idx - 1)
        }
    }
})

$btnNaechste.Add_Click({
    if ($ergebnisListe.SelectedItems.Count -gt 0 -and $ergebnisListe.Items.Count -gt 0) {
        $idx = $ergebnisListe.Items.IndexOf($ergebnisListe.SelectedItems[0])
        if ($idx -lt $ergebnisListe.Items.Count - 1) {
            $ergebnisListe.Items[$idx].Selected = $false
            $ergebnisListe.Items[$idx + 1].Selected = $true
            $ergebnisListe.Items[$idx + 1].Focused = $true
            $ergebnisListe.EnsureVisible($idx + 1)
        }
    }
})

# --- Anhang öffnen ---
# Öffnet den in der Anhangsliste markierten Anhang; ist keiner markiert, werden alle geöffnet.
$btnAnhang.Add_Click({
    try {
        $mail = Get-AusgewaehlteMail
        if (-not $mail) {
            [System.Windows.Forms.MessageBox]::Show("Bitte zuerst eine Mail auswählen.","Hinweis",'OK','Information')
            return
        }
        if ($mail.Attachments.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Diese Mail hat keinen Anhang.","Hinweis",'OK','Information')
            return
        }

        # Falls in der Anhangsliste ein konkreter Anhang markiert ist, nur diesen öffnen
        $ausgewaehlterName = $null
        if ($lstAnhaenge.SelectedItem -and $lstAnhaenge.SelectedItem -ne "(keine Anhänge)") {
            $ausgewaehlterName = [string]$lstAnhaenge.SelectedItem
        }

        foreach ($att in $mail.Attachments) {
            if ($ausgewaehlterName -and $att.FileName -ne $ausgewaehlterName) { continue }

            $zielPfad = Join-Path $TempAnhangPfad $att.FileName
            $att.SaveAsFile($zielPfad)
            Start-Process -FilePath $zielPfad
            Write-Log -Text "Anhang geöffnet: $($att.FileName)" -Level INFO
        }
    }
    catch {
        Write-Log -Text "Fehler beim Öffnen des Anhangs: $($_.Exception.Message)" -Level ERROR
        [System.Windows.Forms.MessageBox]::Show("Anhang konnte nicht geöffnet werden: $($_.Exception.Message)","Fehler",'OK','Error')
    }
})

# --- Doppelklick in der Anhangsliste öffnet direkt den markierten Anhang ---
$lstAnhaenge.Add_DoubleClick({ $btnAnhang.PerformClick() })

# --- Weiterleiten ---
$btnWeiterleiten.Add_Click({
    try {
        $zielAdresse = $txtZiel.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($zielAdresse)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Bitte zuerst eine Ziel-E-Mail-Adresse eingeben.",
                "Hinweis", 'OK', 'Warning')
            return
        }

        # --- E-Mail-Format grob validieren ---
        if ($zielAdresse -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            [System.Windows.Forms.MessageBox]::Show(
                "Die Adresse '$zielAdresse' scheint ungueltig zu sein.",
                "Ungueltige Adresse", 'OK', 'Warning')
            return
        }

        # --- Pruefen, ob Mails in der Liste ausgewaehlt sind ---
        if ($null -eq $ergebnisListe.SelectedItems -or $ergebnisListe.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Bitte zuerst eine oder mehrere Mails auswaehlen (Strg+Klick).",
                "Hinweis", 'OK', 'Information')
            return
        }

        $weiterleitungsDatei = Join-Path $LogVerzeichnis "Outlook_Weiterleitungen.json"

        # --- Alle ausgewaehlten Mails ermitteln und auf bereits weitergeleitet pruefen ---
        $zuSenden = @()       # Mails, die noch nicht weitergeleitet wurden
        $bereitsWeg = @()     # Mails, die schon weitergeleitet wurden
        # Historie einmal laden (fuer Pruefung UND spaeteres Speichern)
        $historieGesamt = @()
        if (Test-Path $weiterleitungsDatei) {
            try {
                $historieGesamt = @(Get-Content -Path $weiterleitungsDatei -Raw -Encoding UTF8 | ConvertFrom-Json)
            } catch {
                $historieGesamt = @()
            }
        }
        foreach ($lvItem in $ergebnisListe.SelectedItems) {
            $key   = $lvItem.Tag
            $mail  = $Global:GefundeneMails[$key]
            if ($null -eq $mail) { continue }

            $schonWeg = $false
            if (-not [string]::IsNullOrWhiteSpace($mail.EntryID)) {
                foreach ($h in $historieGesamt) {
                    if ($h.MailId -eq $mail.EntryID -and $h.Empfaenger -eq $zielAdresse) {
                        $schonWeg = $true
                        break
                    }
                }
            }

            if ($schonWeg) { $bereitsWeg += $mail } else { $zuSenden += $mail }
        }

        # --- Nachfragen wenn einige bereits weitergeleitet wurden ---
        if ($bereitsWeg.Count -gt 0 -and $zuSenden.Count -gt 0) {
            $meldung = "$($bereitsWeg.Count) von $($ergebnisListe.SelectedItems.Count) Mails wurden bereits an '$zielAdresse' weitergeleitet.`n`nDie restlichen $($zuSenden.Count) trotzdem senden?"
            $antwort = [System.Windows.Forms.MessageBox]::Show($meldung, "Bereits weitergeleitet", 'YesNo', 'Warning')
            if ($antwort -eq 'No') { return }
        }
        elseif ($bereitsWeg.Count -gt 0 -and $zuSenden.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Alle ausgewaehlten Mails wurden bereits an '$zielAdresse' weitergeleitet.",
                "Hinweis", 'OK', 'Information')
            return
        }

        # --- Temporäres Verzeichnis für Anhänge (einmalig anlegen) ---
        $tmpDir = Join-Path $env:TEMP "OutlookMailsuche_Weiterleitung"
        if (-not (Test-Path $tmpDir)) {
            New-Item -Path $tmpDir -ItemType Directory -Force | Out-Null
        }

        $erfolgreich = 0

        :weiterleitung foreach ($originalMail in $zuSenden) {
            try {
                if ($chkNurAnhangWeiterleiten.Checked) {
                    # --- Nur Anhänge weiterleiten ---
                    if ($originalMail.Attachments.Count -eq 0) {
                        Write-Log -Text "Mail '$($originalMail.Subject)' hat keine Anhaenge - uebersprungen." -Level WARN
                        continue weiterleitung
                    }

                    # Anhänge speichern (optional nur PDFs)
                    $savedFiles = @()
                    foreach ($att in $originalMail.Attachments) {
                        if ($chkNurPDF.Checked) {
                            $ext = [System.IO.Path]::GetExtension($att.FileName)
                            if ($ext -notmatch '\.pdf') {
                                Write-Log -Text "Anhang '$($att.FileName)' ist keine PDF - uebersprungen." -Level INFO
                                continue
                            }
                        }
                        $fileName = $att.FileName -replace '[<>:"/\\|?*]', '_'
                        $filePath = Join-Path $tmpDir $fileName
                        $counter = 1
                        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
                        $extension = [System.IO.Path]::GetExtension($fileName)
                        while (Test-Path $filePath) {
                            $fileName = "$baseName($counter)$extension"
                            $filePath = Join-Path $tmpDir $fileName
                            $counter++
                        }
                        $att.SaveAsFile($filePath)
                        $savedFiles += $filePath
                    }

                    # Keine PDFs (bei aktiviertem Filter) → Mail überspringen
                    if ($chkNurPDF.Checked -and $savedFiles.Count -eq 0) {
                        Write-Log -Text "Mail '$($originalMail.Subject)' enthaelt keine PDF-Anhaenge - uebersprungen." -Level WARN
                        continue weiterleitung
                    }

                    # Neue Mail mit Anhängen erstellen und senden
                    $newMail = $outlook.CreateItem(0)
                    $newMail.Subject = "Weitergeleitete Anhaenge: $($originalMail.Subject)"
                    $newMail.Body = "Anhaenge aus der Mail '$($originalMail.Subject)' vom $($originalMail.ReceivedTime.ToString('dd.MM.yyyy')) werden hiermit weitergeleitet."
                    foreach ($file in $savedFiles) {
                        $newMail.Attachments.Add($file) | Out-Null
                    }
                    $newMail.Recipients.Add($zielAdresse) | Out-Null
                    $newMail.Recipients.ResolveAll() | Out-Null
                    $newMail.Send()

                    # Temp-Dateien aufräumen
                    foreach ($file in $savedFiles) {
                        Remove-Item $file -Force -ErrorAction SilentlyContinue
                    }

                    Write-Log -Text "Anhaenge aus '$($originalMail.Subject)' ($($savedFiles.Count) Datei(en)) an '$zielAdresse' weitergeleitet." -Level INFO
                }
                else {
                    # --- Ganze Mail weiterleiten ---
                    $weiterleitung = $originalMail.Forward()
                    $weiterleitung.Recipients.Add($zielAdresse) | Out-Null
                    if (-not $weiterleitung.Recipients.ResolveAll()) {
                        Write-Log -Text "Empfaenger '$zielAdresse' nicht vollstaendig aufloesbar - wird trotzdem versendet." -Level WARN
                    }
                    $weiterleitung.Send()
                    Write-Log -Text "Mail '$($originalMail.Subject)' weitergeleitet an '$zielAdresse'." -Level INFO
                }

                # --- Historie-Eintrag merken ---
                $historieGesamt += @{
                    MailId     = $originalMail.EntryID
                    Empfaenger = $zielAdresse
                    Datum      = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
                $erfolgreich++
            }
            catch {
                Write-Log -Text "Fehler beim Weiterleiten von '$($originalMail.Subject)': $($_.Exception.Message)" -Level ERROR
            }
        }

        # --- Historie speichern (alle Eintraege auf einmal) ---
        if ($historieGesamt.Count -gt 0) {
            try {
                $historieGesamt | ConvertTo-Json -Depth 1 | Out-File -FilePath $weiterleitungsDatei -Encoding UTF8
            }
            catch {
                $fehlerMsg = "Fehler beim Speichern der Weiterleitungs-Historie: $($_.Exception.Message)"
                Write-Log -Text $fehlerMsg -Level WARN
                [System.Windows.Forms.MessageBox]::Show($fehlerMsg, "Warnung", 'OK', 'Warning')
            }
        }

        # --- Weiterleitungs-Spalte aktualisieren ---
        if ($erfolgreich -gt 0) {
            foreach ($item in $ergebnisListe.Items) {
                $key  = $item.Tag
                $mail = $Global:GefundeneMails[$key]
                if (-not $mail) { continue }
                $weg = ""
                if ($mail.EntryID) {
                    foreach ($h in $historieGesamt) {
                        if ($h.MailId -eq $mail.EntryID -and $h.Empfaenger -eq $zielAdresse) {
                            $weg = "Ja"
                            break
                        }
                    }
                }
                if ($item.SubItems.Count -ge 5) { $item.SubItems[4].Text = $weg }
            }
        }

        # --- Zusammenfassung ---
        $lblStatus.Text = "$erfolgreich von $($zuSenden.Count) Mail(s) an '$zielAdresse' weitergeleitet."
        [System.Windows.Forms.MessageBox]::Show(
            "$erfolgreich von $($zuSenden.Count) Mail(s) wurden erfolgreich an '$zielAdresse' weitergeleitet.",
            "Erfolg",'OK','Information')
    }
    catch {
        # --- Zentrale Fehlerbehandlung mit Logging ---
        $fehlerText = "Fehler beim Weiterleiten: $($_.Exception.Message)"
        Write-Log -Text $fehlerText -Level ERROR
        [System.Windows.Forms.MessageBox]::Show($fehlerText, "Fehler", 'OK', 'Error')
    }
})

#EndRegion


#Region ================== START ==================
try {
    [void]$form.ShowDialog()
}
finally {
    # COM-Objekte sauber freigeben
    if ($outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
    Write-Log -Text "Script beendet." -Level INFO
}
#EndRegion
