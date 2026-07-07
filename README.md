# ScriptsScriptsScripts

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![License](https://img.shields.io/badge/License-MIT-green)
![Maintenance](https://img.shields.io/badge/Maintained-yes-brightgreen)

Sammlung verschiedener PowerShell-Skripte für die Systemadministration.

## Enthaltene Skripte

### Citrix

| Skript | Beschreibung |
|--------|-------------|
| `Optimierung_Citrix_PVS.ps1` | Netzwerk- und Windows-Optimierung für Citrix PVS Target Devices. Deaktiviert Offload-Features, erstellt Sicherungen und unterstützt Rollback. |

## Voraussetzungen

- Windows PowerShell 5.1 oder höher
- Administratorrechte für ausführende Skripte
- Je nach Skript: entsprechende Module (ActiveDirectory, Exchange, Citrix PVS etc.)

## Verwendung

```powershell
# Beispiel: Citrix PVS Optimierung
.\Citrix\Optimierung_Citrix_PVS.ps1
```

## Lizenz

Siehe `LICENSE`.
