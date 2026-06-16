# KopfnotenTool – Übersicht und Funktionen

## Beschreibung

**KopfnotenTool** ist eine Desktop-Anwendung zum Import, zur Verwaltung, Auswertung und Weiterverarbeitung von Kopfnoten aus dem Schulportal Hessen (SPH). Es unterstützt den automatisierten SPH-Download, den manuellen Excel-Import, die Bearbeitung von Noten, statistische Auswertungen und den Word-/Excel-Export.

Das Tool richtet sich vor allem an Integrierte Gesamtschulen (IGS) in Hessen, die ein Beiblatt zum Zeugnis mit Kopfnoten aller Fächer benötigen.

**Aktuelle Version:** [v1.1.0](https://github.com/jpospi/kopfnotentool/releases/tag/v1.1.0)

## Hauptfunktionen

### Import (Tab „Import“)

- **SPH Download & Import** – direkter Login im Schulportal Hessen, automatische Klassenerkennung und Import in die lokale Datenbank
- **Autoimport bis 9 Züge pro Jahrgang** (Klassen `05a` … `05i`, analog für J6–J10)
- **Manueller Excel-Import** – lokale SPH-Dateien auswählen und einlesen
- **Live Import-Log** während SPH-Download und Verarbeitung
- **Backup-Klassenangaben** (Fallback bei fehlgeschlagener Autoerkennung) unter `Datei → Backup-Klassenangaben…`

### Datenbank (Tab „Datenbank“)

- Übersicht aller Lernenden mit Filter (Klasse, Name, Lehrkraft, Status)
- Status **„Vollständig“** / **„Unvollständig“** je Lernendem
- **SPH-Abgleich** je Lernendem (Ampelfarben und Spalte „SPH-Abgleich“)
- Noten bearbeiten (Doppelklick), inkl. Sondernoten **GB** und **NF**
- Lernende **deaktivieren** bzw. deaktivierte Lernende verwalten (bleiben über Re-Import erhalten)
- **Fehlliste exportieren** (fehlende Fächer inkl. Lehrkräfte)

### Analyse (Tab „Analyse“)

- KPIs: Lernende, Gesamtschnitt, Vollständigkeit, Klassenanzahl
- Tabellen: Klassen, Jahrgänge, Fächer-Ranking
- **Top Lernende:** Top 10 Schule, Top 3 je Jahrgang, Klassenbeste je Klasse
- **Vergleichsperioden** (Mehrfachauswahl) für Entwicklung über Halbjahre hinweg
- Deaktivierte Lernende sind in allen Kennzahlen ausgeschlossen

### Export (Tab „Export“)

- Serienbrief-Export als Word (`.docx`) mit vorhandener Template-Datei
- Excel-Export fehlender Noten
- Template-Auswahl über Dateidialog im Export-Tab (kein separater Template-Manager in der UI)

### Datei-Menü

- Datenbank öffnen, **importieren**, **exportieren**, sichern
- Backup-Klassenangaben für SPH-Fallback

## Installation

### Windows (empfohlen)

Fertige Builds stehen als GitHub-Release bereit:

- **Installer:** [KopfnotenTool-Setup.exe](https://github.com/jpospi/kopfnotentool/releases/latest) (empfohlen)
- **Portable:** `KopfnotenTool.exe` (ohne Installation)

### Aus dem Quellcode

```bash
git clone https://github.com/jpospi/kopfnotentool.git
cd kopfnotentool

python3 -m venv .venv
# Windows: .venv\Scripts\activate
# Linux/macOS: source .venv/bin/activate

pip install -r requirements.txt
python app.py
```

## Windows: EXE und Installer selbst bauen

```powershell
.\build_windows.ps1
```

Ergebnis: `dist\KopfnotenTool.exe`

Installer (Inno Setup 6):

```powershell
iscc "installer\KopfnotenTool.iss"
```

Ergebnis: `dist\KopfnotenTool-Setup.exe`

Beim Setup können u. a. konfiguriert werden:

- Datenpfad (Basisordner)
- Import-, Word- und Excel-Exportpfade
- Datenbank- und Backup-Ordner

Die Werte werden in `kopfnotentool.paths.json` neben der EXE gespeichert.

## Nutzung

### SPH-Login und Import

1. Tab **Import** öffnen
2. Bei Bedarf anmelden (`Anmeldung ändern`)
3. **SPH Download & Import** starten – Klassen werden automatisch erkannt (bis `…i`)
4. Alternativ: Excel-Dateien manuell auswählen und importieren

Zugangsdaten werden lokal verschlüsselt gespeichert (nicht im Klartext). Für den SPH-Download sind Tooladmin-Rechte im Kopfnotenmodul erforderlich.

### Datenbank und Bearbeitung

- Periode (Schuljahr/Halbjahr) im Tab **Datenbank** wählen
- Filter nutzen oder Doppelklick auf **Fächer** zum Noten-Editor
- Tab **Analyse** für Statistiken und Rankings

### Word-Export

Im Tab **Export** eine bestehende `.docx`-Template-Datei auswählen. Ein visueller Template-Designer ist in v1.1.0 **deaktiviert** (Code bleibt im Projekt für spätere Wiederaufnahme).

## Konfiguration

- Laufzeitpfade: `kopfnotentool.paths.json`
- Beispiel: `kopfnotentool.paths.example.json`
- SPH-Konfiguration und Backup-Klassen: `sph_config.json` (Datenordner)

## Hinweise zur Version 1.1.0

- Neuer **Analyse-Tab** mit KPIs, Rankings und Periodenvergleich
- Import-UI: SPH oben, manueller Import kompakt unten
- SPH-Abgleich pro Lernendem (nicht klassenweise)
- Datenbank-Import/-Export über das Datei-Menü
- Autoimport für **neunzügige Jahrgänge** (a–i)
- Template-Manager/Designer in der Oberfläche **vorübergehend deaktiviert**

## Lizenz

Dieses Projekt ist unter der **MIT-Lizenz** veröffentlicht (siehe `LICENSE`).
