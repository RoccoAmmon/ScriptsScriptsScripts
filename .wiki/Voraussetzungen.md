# ⚙️ Voraussetzungen

Damit alle Skripte reibungslos laufen, solltest du folgendes beachten:

---

## 🖥️ System

- **Windows PowerShell 5.1** oder höher
- **Administratorrechte** (für die meisten Skripte zwingend)
- Je nach Skript: entsprechende Module
  - `ActiveDirectory`
  - Citrix PVS-Komponenten

---

## 📝 Encoding

> ⚠️ Skripte **immer als UTF-8-BOM** speichern!
> Das ist nötig für die Windows PowerShell 5.1-Kompatibilität (Umlaute, Sonderzeichen).

---

## 📦 Module installieren (Beispiel)

```powershell
# Remote Server Administration Tools (RSAT) für AD-Modul
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

---

<sub>© Rocco Ammon · MIT-Lizenz</sub>
