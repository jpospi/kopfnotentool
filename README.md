# KopfnotenTool – Übersicht und Funktionen

## Beschreibung

**KopfnotenTool** ist eine Desktop‑Anwendung zum Import, Verwaltung, Auswertung und Weiterverarbeitung von Kopfnoten aus dem Schulportal Hessen. Es ermöglicht einen manuellen und direkten Import von Excel‑Dateien aus dem SPH, das Analysieren und Bearbeiten von Kopfnoten sowie das Erstellen von Export‑Templates und die Weiterverarbeitung der Noten in Serienbriefen. 
Das Tool richtet sich vor allem an Integrierte Gesamtschulen (IGS) in Hessen, die ein Beiblatt zum Zeugnis mit Kopfnoten aller Fächer benötigen.

## Hauptfunktionen

1. **SPH‑Login & Excel‑Import**
   - Direkter Login im Schulportal Hessen (SPH) zum automatisierten Abruf von Kopfnotendaten.
   - Einlesen von Klassen‑Excel‑Dateien aus dem SPH.
   - Automatisches Erkennen von Fach‑ und Lehrer‑Informationen.
   - Speicherung der Daten in einer gesicherten, lokalen SQLite‑Datenbank.

2. **Analyse‑Tab**
   - Anzeige von Lernenden, Klassen und Noten‑Statistiken.
   - Berechnung des Status **„Vollständig“** oder **„Unvollständig“** für alle Lernenden (basierend auf Fach‑ und Notenzahl).
   - Identifikation fehlender Noten und Erstellung einer Kontroll‑Lücken‑Liste.
   - Bearbeiten und Nachtragen von Kopfnoten.

3. **Export‑Tab**
   - Export in Serienbrief/en mit allen Kopfnoten als Klassenssatz.
   - Export von detaillierten Listen mit fehlenden Noten nach Excel.
   - Automatisches Anlegen von Klassen‑Sheets und einer Gesamt‑Übersicht.

4. **Template‑Designer**
   - Visueller Designer für Word‑Templates (horizontales 3‑Zeilen‑Layout).
   - Dynamische Erstellung von Tabellen mit variabler Fach‑Anzahl pro Schüler.
   - Speicherung als `.docx`‑Datei.


## Installation

```bash
# Repository klonen
git clone https://github.com/jpospi/kopfnotentool.git
cd kopfnotentool

# Virtuelle Umgebung erstellen und aktivieren
python3 -m venv .venv
source .venv/bin/activate

# Abhängigkeiten installieren
pip install -r requirements.txt
```

## Windows: EXE Build und Installer

```powershell
# im Projektordner
.\build_windows.ps1
```

Danach liegt die EXE unter `dist/KopfnotenTool.exe`.

Für einen Installer:

1. [Inno Setup](https://jrsoftware.org/isdl.php) installieren
2. `installer/KopfnotenTool.iss` öffnen
3. Build starten (`Compile`)

Der Installer fragt beim Setup gezielt nach:

- Datenpfad (Basisordner)
- Standard-Importpfad
- Word-Exportpfad
- Excel-Exportpfad
- Datenbank-Ordner
- Backup-Ordner für Datenbanksicherungen

Diese Werte werden in `kopfnotentool.paths.json` neben der EXE gespeichert.

## Nutzung

```bash
# Anwendung starten
python app.py
```

### Schulportal-Login (SPH) und Import
Das Tool bietet eine integrierte Anbindung an das Schulportal Hessen.
Sie können sich direkt mit Ihren SPH-Zugangsdaten anmelden.
Das Passwort wird **nicht** im Klartext gespeichert. Es wird lokal verschlüsselt (via `keyring` und `cryptography`) abgelegt, sodass Sie sich bei jedem Neustart auch offline anmelden können.
Zum direkt Download der Excel-Listen müssen Sie Tooladmin für das Kopfnotenmodul sein und angeben, wie viele Klassen es pro Jahrgang gibt. Es besteht aber auch die Möglichkeit, die Listen aus dem SPH herunterzuladen und manuell zu importieren.


### Analyse und Bearbeitung
Der Analyse-Tab bietet eine zentrale Übersicht über den Status aller eingelesenen Schüler.
Das System berechnet automatisch, ob ein Schüler "Vollständig" (alle erforderlichen Noten vorhanden) oder "Unvollständig" ist.
Suchen und Filterfunktionen ermöglichen es Ihnen, gezielt nach unvollständigen Schülern zu suchen und zu filtern.
Durch Doppelklick auf einen Schüler (Spalte "Fächer") öffnet sich ein Editor, in dem Sie fehlende Noten direkt manuell nachtragen können.


## Konfiguration

- Laufzeitpfade werden über `kopfnotentool.paths.json` gesteuert.
- Eine Beispielkonfiguration liegt in `kopfnotentool.paths.example.json`.
- Wichtige Pfade: Import, Word-/Excel-Export, Datenbankpfad, Backup-Ordner, Template-/Log-/Temp-Verzeichnisse.
- Im Menü `Datei` gibt es den Punkt **Datenbank sichern** (Backup in den konfigurierten Backup-Ordner).

## Lizenz

Dieses Projekt ist unter der **MIT‑Lizenz** veröffentlicht (siehe `LICENSE`).
