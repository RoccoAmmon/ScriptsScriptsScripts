# ScriptsScriptsScripts

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![License](https://img.shields.io/badge/License-MIT-green)
![Maintenance](https://img.shields.io/badge/Maintained-yes-brightgreen)
![GitHub last commit](https://img.shields.io/github/last-commit/RoccoAmmon/ScriptsScriptsScripts)
![GitHub repo size](https://img.shields.io/github/repo-size/RoccoAmmon/ScriptsScriptsScripts)

> Sammlung von PowerShell-Skripten für die Citrix-/Windows-Systemadministration – mit Fokus auf **Monitoring**, **Optimierung** und **Automatisierung**.

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
| [`System-Status-Application-Fehler-Report.ps1`](.wiki/System-Status-Report.md) | Erweiterter System-Status-Report – sammelt **Application Error/Hang/Popup**, **Service Control Manager** (Dienstabstürze) und **Windows Resource Exhaustion** (Speichermangel) von allen Servern einer OU. 🟢🟡🔴 Ampelanzeige für Speicherplatz D:, mcsdif.vhdx-Größe **und** freien Arbeitsspeicher. Zeigt aktive **Citrix Sessions**, **TOP 10 Speicherfresser** (WorkingSet) und **TOP 10 Session-RAM** (nach Benutzer gruppiert). Farbige EXE/DLL/Ausnahmecode/OOM-Hervorhebungen. 🔄 Auto-Refresh per `-Interval`, 🔔 Piepton bei neuen Events, 🔍 interaktive Filter im Browser. |
| [`Optimierung_Citrix_PVS.ps1`](.wiki/Citrix-PVS-Optimierung.md) | Netzwerk- und Windows-Optimierung für Citrix PVS Target Devices. Deaktiviert Offload-Features (Checksum, Jumbo Packet), NIC-Energieverwaltung (3-Stufen-Fallback), IPv6, NetBIOS, Task Offload. Setzt Energieplan auf Höchstleistung, deaktiviert Ruhezustand. 🗃️ Vollständiges JSON-Backup + 🔄 **Rollback** aller Änderungen. |

---

## ✨ Features auf einen Blick

| Feature | System-Status-Report | PVS-Optimierung |
|---------|:---:|:---:|
| Eventlog-Sammlung (remote) | ✅ | – |
| Live-Ampel-Anzeige (HDD, RAM, VHDX) | ✅ | – |
| Auto-Refresh (Intervall-Modus) | ✅ | – |
| TOP 10 Auswertungen (Prozesse, Session-RAM) | ✅ | – |
| In-Page-Modal für Prozessdetails | ✅ | – |
| Interaktive Filter + Sortierung im Browser | ✅ | – |
| Farbige Fehler-Hervorhebungen (EXE/DLL/Excode) | ✅ | – |
| WinForms-GUI mit Live-Log | – | ✅ |
| 3-stufige NIC-Energieverwaltung | – | ✅ |
| Rollback aller Änderungen | – | ✅ |
| JSON-Backup | – | ✅ |

---

## 📖 Wiki

Ausführliche Dokumentation zu allen Skripten findest du im **[📚 Wiki](.wiki/Home.md)**:

| Seite | Inhalt |
|-------|--------|
| [🏠 Home](.wiki/Home.md) | Übersicht und Navigation |
| [📊 System-Status Report](.wiki/System-Status-Report.md) | Vollständige Dokumentation des Monitoring-Reports |
| [🖥️ Citrix PVS Optimierung](.wiki/Citrix-PVS-Optimierung.md) | Details zur PVS Target Device Optimierung |
| [⚙️ Voraussetzungen](.wiki/Voraussetzungen.md) | Systemanforderungen, Module, Encoding |
| [🐛 Fehlerbehebung](.wiki/Fehlerbehebung.md) | Bekannte Probleme und Lösungen |
| [📜 Changelog](.wiki/Changelog.md) | Versionshistorie |

---

## ⚙️ Voraussetzungen

| Anforderung | Details |
|------------|---------|
| **PowerShell** | Windows PowerShell 5.1 oder höher |
| **Rechte** | Administratorrechte für ausführende Skripte |
| **Module** | ActiveDirectory-Modul (für System-Status-Report) |
| **WinRM** | Muss auf Zielservern aktiviert sein (für Remote-Abfragen) |
| **Encoding** | Skripte mit **UTF-8-BOM** speichern (PS 5.1-Kompatibilität) |

---

## 🏗️ Repository-Struktur

```
ScriptsScriptsScripts/
├── .wiki/                          # 📚 Wiki-Dokumentation (Markdown)
│   ├── Home.md
│   ├── System-Status-Report.md
│   ├── Citrix-PVS-Optimierung.md
│   ├── Voraussetzungen.md
│   ├── Fehlerbehebung.md
│   └── Changelog.md
├── Citrix/
│   ├── Optimierung_Citrix_PVS.ps1
│   └── System-Status-Application-Fehler-Report.ps1
├── README.md
├── LICENSE
└── .gitignore
```

---

## 📄 Lizenz

MIT License © 2026 Rocco Ammon – siehe [`LICENSE`](LICENSE).
