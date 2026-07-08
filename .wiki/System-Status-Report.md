# 📊 System-Status Application-Fehler-Report

**Skript:** `Citrix/System-Status-Application-Fehler-Report.ps1`  
**Version:** 1.3 · **Autor:** Rocco Ammon

---

## 📋 Übersicht

Sammelt von allen Servern einer OU die relevanten Eventlog-Einträge und exportiert sie
als übersichtliche **HTML-Datei**. Optional läuft das Skript im Intervall-Modus und
aktualisiert die Seite automatisch im Browser.

---

## 🔍 Erfasste Quellen

- Application Error
- Application Hang
- Application Popup
- Service Control Manager (Dienstabstürze)
- Windows Resource Exhaustion (Speichermangel)

---

## 🟢🟡🔴 Ampel-Anzeigen

| Metrik | Quelle |
|--------|--------|
| Freier Speicherplatz `D:` | WMI |
| Größe `mcsdif.vhdx` | Dateisystem |
| Freier Arbeitsspeicher | WMI |

---

## 🎨 Farbige Hervorhebung

- 🟢 EXE-Pfade
- 🔴 DLL-Pfade
- 🟡 Ausnahmecodes
- 🔴 „Nicht genügend virtueller Speicher"

---

## 🚀 Verwendung

```powershell
# Einmalig, letzte 7 Tage
.\Citrix\System-Status-Application-Fehler-Report.ps1 `
    -SearchBase "OU=Servers,DC=domain,DC=local"

# Alle 30 Min. automatisch aktualisieren
.\Citrix\System-Status-Application-Fehler-Report.ps1 `
    -SearchBase "OU=Servers,DC=domain,DC=local" -Interval 30
```

### Parameter

| Parameter | Pflicht | Standard | Beschreibung |
|-----------|---------|----------|--------------|
| `-SearchBase` | ✅ | – | DN der OU |
| `-OutputPath` | ❌ | `SystemStatusReport.html` | Ausgabe-Pfad |
| `-DaysBack` | ❌ | `7` | Tage zurück |
| `-Interval` | ❌ | `0` | Minuten (0 = einmalig) |

---

<sub>© Rocco Ammon · MIT-Lizenz</sub>
