# 🖥️ Citrix PVS Optimierung

**Skript:** `Citrix/Optimierung_Citrix_PVS.ps1`  
**Version:** 4.8 · **Autor:** Rocco Ammon

---

## 📋 Übersicht

Führt eine Netzwerk- und Windows-Optimierung für **Citrix PVS Target Devices** durch.
Das Skript ist mit einer GUI ausgestattet, die ein Live-Log anzeigt, erstellt vorab eine
Sicherung und bietet einen **vollständigen Rollback**.

---

## ⚙️ Was wird gemacht?

- 🔌 Deaktivierung von Offload-Features am Netzwerkadapter:
  - IPv4 / TCP / UDP Checksum Offload
  - Jumbo Packet
- 💾 Automatische Sicherung der Original-Konfiguration (`PVS2507_Backup.json`)
- 🖼️ WinForms-GUI mit Live-Log (`RichTextBox`)
- ↩️ Rollback über gesicherte Werte

---

## 🚀 Verwendung

```powershell
# Als Administrator ausführen!
.\Citrix\Optimierung_Citrix_PVS.ps1
```

> ⚠️ **Wichtig:** Das Skript muss zwingend als **Administrator** gestartet werden.
> IP, Domäne und Servername sind bereits vorkonfiguriert.

---

## 📁 Log & Backup

| Datei | Pfad |
|-------|------|
| Log | `C:\ScriptLog\PVS2507_Optimierung.log` |
| Backup | `C:\ScriptLog\PVS2507_Backup.json` |

---

<sub>© Rocco Ammon · MIT-Lizenz</sub>
