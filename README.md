# ScriptsScriptsScripts

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![License](https://img.shields.io/badge/License-MIT-green)
![Maintenance](https://img.shields.io/badge/Maintained-yes-brightgreen)

Sammlung verschiedener PowerShell-Skripte für die Systemadministration.

## Enthaltene Skripte

### Root

| Skript | Beschreibung |
|--------|-------------|
| `Application_Error_Eventlog.ps1` | Exportiert Application Error/Hang/Popup Eventlog-Einträge aller Server einer OU als HTML inkl. Ampelanzeige für Speicherplatz D: und mcsdif.vhdx-Größe. |

### Citrix

| Skript | Beschreibung |
|--------|-------------|
| `Application_Error_Eventlog.ps1` | Exportiert Application Error/Hang/Popup Eventlog-Einträge aller Server einer OU als HTML inkl. Ampelanzeige für Speicherplatz D: und mcsdif.vhdx-Größe. |
| `Optimierung_Citrix_PVS.ps1` | Netzwerk- und Windows-Optimierung für Citrix PVS Target Devices. Deaktiviert Offload-Features, erstellt Sicherungen und unterstützt Rollback. |

## Voraussetzungen

- Windows PowerShell 5.1 oder höher
- Administratorrechte für ausführende Skripte
- Je nach Skript: entsprechende Module (ActiveDirectory, Citrix PVS etc.)
- Skripte mit UTF-8-BOM speichern (Windows PowerShell 5.1-Kompatibilität)

## Verwendung

```powershell
# Eventlog-Report erstellen
.\Application_Error_Eventlog.ps1 -SearchBase "OU=Servers,DC=domain,DC=local"

# Citrix PVS Optimierung
.\Citrix\Optimierung_Citrix_PVS.ps1
```

## Lizenz

Siehe `LICENSE`.
