# ScriptsScriptsScripts

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![License](https://img.shields.io/badge/License-MIT-green)
![Maintenance](https://img.shields.io/badge/Maintained-yes-brightgreen)

Sammlung verschiedener PowerShell-Skripte für die Systemadministration.

## Enthaltene Skripte

### Citrix

| Skript | Beschreibung |
|--------|-------------|
| `System-Status-Application-Fehler-Report.ps1` | Erweiterter System-Status-Report: Überwacht Application Error/Hang/Popup, **Service Control Manager** (Dienstabstürze) und **Windows Resource Exhaustion** (Speichermangel). Ampelanzeige für Speicherplatz D:, mcsdif.vhdx-Größe **und freien Arbeitsspeicher**. Farbige Hervorhebungen (EXE/DLL/Ausnahmecodes/OOM). Unterstützt automatische Aktualisierung per -Interval. |
| `Optimierung_Citrix_PVS.ps1` | Netzwerk- und Windows-Optimierung für Citrix PVS Target Devices. Deaktiviert Offload-Features, erstellt Sicherungen und unterstützt Rollback. |

## Voraussetzungen

- Windows PowerShell 5.1 oder höher
- Administratorrechte für ausführende Skripte
- Je nach Skript: entsprechende Module (ActiveDirectory, Citrix PVS etc.)
- Skripte mit UTF-8-BOM speichern (Windows PowerShell 5.1-Kompatibilität)

## Verwendung

```powershell
# System-Status-Report einmalig erstellen
.\Citrix\System-Status-Application-Fehler-Report.ps1 -SearchBase "OU=Servers,DC=domain,DC=local"

# System-Status-Report alle 30 Min. automatisch aktualisieren
.\Citrix\System-Status-Application-Fehler-Report.ps1 -SearchBase "OU=Servers,DC=domain,DC=local" -Interval 30

# Citrix PVS Optimierung
.\Citrix\Optimierung_Citrix_PVS.ps1
```

## Lizenz

Siehe `LICENSE`.
