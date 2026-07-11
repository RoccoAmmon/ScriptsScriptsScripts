# ScriptsScriptsScripts

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![License](https://img.shields.io/badge/License-MIT-green)
![Maintenance](https://img.shields.io/badge/Maintained-yes-brightgreen)
![GitHub last commit](https://img.shields.io/github/last-commit/RoccoAmmon/ScriptsScriptsScripts)
![GitHub repo size](https://img.shields.io/github/repo-size/RoccoAmmon/ScriptsScriptsScripts)

> Sammlung von PowerShell-Skripten für die Citrix-/Windows-Systemadministration und Outlook-Automation – mit Fokus auf **Monitoring**, **Optimierung** und **Automatisierung**.

---

## 🚀 Quick Start

```powershell
# System-Status-Report über alle Server einer OU (einmalig)
.\Citrix\System-Status-Application-Fehler-Report.ps1 -SearchBase "OU=Servers,DC=domain,DC=local"

# Citrix PVS Target Device optimieren (als Administrator)
.\Citrix\Optimierung_Citrix_PVS.ps1
```

---

## 📦 Enthaltene Skripte

### 🖥️ Citrix

| Skript | Beschreibung |
|--------|-------------|
| [`Outlook-Mail-Explorer.ps1`](./Outlook/Outlook-Mail-Explorer.ps1) | 📧 **Outlook-Mail-Explorer mit WinForms-GUI v1.2** – Durchsucht ein oder mehrere Postfächer nach Mails mit Suchwort, Datumsbereich und Anhang-Filter. **Sortierbare Ergebnisspalten** (Klick auf Kopf), **Mehrfachauswahl** (Strg+Klick) für Sammel-Weiterleitung, **Postfach-Spalte**. Ergebnisliste mit **Weiterleitungs-Status** (erkennt doppelte Sendungen). **Vorschau** des Mail-Inhalts per Klick, **Anhänge öffnen/speichern**, **Weiterleitung** ganzer Mails oder nur der Anhänge (umschaltbar). **PDF-Filter** für reine PDF-Weiterleitung. **Navigations-Buttons** (▲/▼) zum Durchblättern. Live-Update der Weiterleitungs-Spalte bei Adressänderung. Suchabbruch, Fortschrittsanzeige, Logging. 📖 [Doku →](https://github.com/RoccoAmmon/ScriptsScriptsScripts/wiki/Outlook-Suche) |
| [`System-Status-Application-Fehler-Report.ps1`](./Citrix/System-Status-Application-Fehler-Report.ps1) | Erweiterter System-Status-Report – sammelt **Application Error/Hang/Popup**, **Service Control Manager** (Dienstabstürze) und **Windows Resource Exhaustion** (Speichermangel) von allen Servern einer OU. 🟢🟡🔴 Ampelanzeige für Speicherplatz D:, mcsdif.vhdx-Größe, freien Arbeitsspeicher, **CPU-Auslastung**, **Auslagerungsdatei** und **FSLogix-Dienststatus**. Zeigt aktive **Citrix Sessions**, **TOP 10 Speicherfresser** (WorkingSet) und **TOP 10 Session-RAM** (nach Benutzer gruppiert). **Medico Update-Version** live abgerufen (mit Rot-Markierung bei veralteten Versionen). Farbige EXE/DLL/Ausnahmecode/OOM-Hervorhebungen. 🔄 Auto-Refresh per `-Interval`, 🔔 Piepton bei neuen Events, 🔍 interaktive Filter im Browser. 🎛️ Alle Schwellwerte als Variablen anpassbar. 📖 [Doku →](https://github.com/RoccoAmmon/ScriptsScriptsScripts/wiki/System-Status-Report) |
| [`Optimierung_Citrix_PVS.ps1`](./Citrix/Optimierung_Citrix_PVS.ps1) | Netzwerk- und Windows-Optimierung für Citrix PVS Target Devices. Deaktiviert Offload-Features (Checksum, Jumbo Packet), NIC-Energieverwaltung (3-Stufen-Fallback), IPv6, NetBIOS, Task Offload. Setzt Energieplan auf Höchstleistung, deaktiviert Ruhezustand. 🗃️ Vollständiges JSON-Backup + 🔄 **Rollback** aller Änderungen. 📖 [Doku →](https://github.com/RoccoAmmon/ScriptsScriptsScripts/wiki/Citrix-PVS-Optimierung) |

---

## ✨ Features auf einen Blick

| Feature | System-Status-Report | PVS-Optimierung | Outlook-Suche |
|---------|:---:|:---:|:---:|
| Eventlog-Sammlung (remote) | ✅ | – | – |
| Live-Ampel-Anzeige (HDD, RAM, VHDX, CPU, Pagefile) | ✅ | – | – |
| FSLogix & weitere Dienste (Cortex, WEM, Broker) | ✅ | – | – |
| Medico Update-Version (live abgerufen) | ✅ | – | – |
| Auto-Refresh (Intervall-Modus) | ✅ | – | – |
| TOP 10 Auswertungen (Prozesse, Session-RAM) | ✅ | – | – |
| In-Page-Modal für Prozessdetails | ✅ | – | – |
| Interaktive Filter + Sortierung im Browser | ✅ | – | – |
| Farbige Fehler-Hervorhebungen (EXE/DLL/Excode) | ✅ | – | – |
| Anpassbare Schwellwerte (Variablen am Anfang) | ✅ | – | – |
| WinForms-GUI mit Live-Log | – | ✅ | ✅ |
| 3-stufige NIC-Energieverwaltung | – | ✅ | – |
| Rollback aller Änderungen | – | ✅ | – |
| JSON-Backup | – | ✅ | – |
| Outlook-Postfach-Suche (alle/mehrere) | – | – | ✅ |
| Anhang-Filter & Vorschau | – | – | ✅ |
| PDF-Filter (Nur PDF-Anhänge) | – | – | ✅ |
| Navigations-Buttons (▲/▼) | – | – | ✅ |
| Weiterleitungs-Historie (duplex-Erkennung) | – | – | ✅ |
| Anhänge einzeln öffnen/speichern | – | – | ✅ |
| Suchabbruch + Fortschritt | – | – | ✅ |

---

## 📖 Wiki

Ausführliche Dokumentation zu allen Skripten findest du im **[📚 Wiki](https://github.com/RoccoAmmon/ScriptsScriptsScripts/wiki)**:

| Seite | Inhalt |
|-------|--------|
| [🏠 Home](https://github.com/RoccoAmmon/ScriptsScriptsScripts/wiki/Home) | Übersicht und Navigation |
| [📊 System-Status Report](https://github.com/RoccoAmmon/ScriptsScriptsScripts/wiki/System-Status-Report) | Vollständige Dokumentation des Monitoring-Reports |
| [🖥️ Citrix PVS Optimierung](https://github.com/RoccoAmmon/ScriptsScriptsScripts/wiki/Citrix-PVS-Optimierung) | Details zur PVS Target Device Optimierung |
| [⚙️ Voraussetzungen](https://github.com/RoccoAmmon/ScriptsScriptsScripts/wiki/Voraussetzungen) | Systemanforderungen, Module, Encoding |
| [📧 Outlook-Suche](https://github.com/RoccoAmmon/ScriptsScriptsScripts/wiki/Outlook-Suche) | Dokumentation zur Outlook-Mailsuche |
| [🐛 Fehlerbehebung](https://github.com/RoccoAmmon/ScriptsScriptsScripts/wiki/Fehlerbehebung) | Bekannte Probleme und Lösungen |
| [📜 Changelog](https://github.com/RoccoAmmon/ScriptsScriptsScripts/wiki/Changelog) | Versionshistorie |

---

## ⚙️ Voraussetzungen

| Anforderung | Details |
|------------|---------|
| **PowerShell** | Windows PowerShell 5.1 oder höher |
| **Rechte** | Administratorrechte für ausführende Skripte (Citrix); lokale Benutzerrechte für Outlook-Suche |
| **Module** | ActiveDirectory-Modul (für System-Status-Report) |
| **Outlook** | Lokal installierte und konfigurierte Outlook-Instanz (für Outlook-Suche) |
| **WinRM** | Muss auf Zielservern aktiviert sein (für Remote-Abfragen) |
| **Encoding** | Skripte mit **UTF-8-BOM** speichern (PS 5.1-Kompatibilität) |

---

## 🏗️ Repository-Struktur

```
ScriptsScriptsScripts/
├── Citrix/
│   ├── Optimierung_Citrix_PVS.ps1
│   └── System-Status-Application-Fehler-Report.ps1
├── Outlook/
│   └── Outlook-Mail-Explorer.ps1
├── README.md
├── LICENSE
└── .gitignore
```

---

## 📄 Lizenz

MIT License © 2026 Rocco Ammon – siehe [`LICENSE`](LICENSE).
