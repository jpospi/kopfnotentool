# KopfnotenTool – Übersicht und Funktionen

## Beschreibung

**KopfnotenTool** ist eine Desktop‑Anwendung zum Import, Verwaltung, Auswertung und Weiterverarbeitung von Kopfnoten aus dem Schulportal Hessen. Es ermöglicht einen manuellen und direkten Import von Excel‑Dateien aus dem SPH, das Analysieren und Bearbeiten von Kopfnoten sowie das Erstellen von Export‑Templates und die Weiterverarbeitung der Noten in Serienbriefen. 
Das Tool richtet sich vor allem an Integrierte Gesamtschulen (IGS) in Hessen, die ein Beiblatt zum Zeugnis mit Kopfnoten aller Fächer benötigen.

## Hauptfunktionen

1. **Excel‑Import**
   - Einlesen von Klassen‑Excel‑Dateien aus dem SPH.
   - Automatisches Erkennen von Fach‑ und Lehrer‑Informationen.
   - Speicherung der Daten in einer gesicherten, lokalen SQLite‑Datenbank (`kopfnoten_secure.db`).

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

5. **SPH‑Downloader**
   - Authentifizierung über das Schulportal Hessen (SPH) mittels `lanisapi`.
   - Herunterladen von Klassen‑Excel‑Dateien für ausgewählte Schulen.
   - Caching der Schulliste für schnellere Wiederholungen.


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

## Nutzung

```bash
# Anwendung starten
python app.py
```

## Konfiguration

- Die Datenbank wird im Verzeichnis `output_database/` verschlüsselt angelegt.
- Export‑Dateien landen im Verzeichnis `output_excel/` (Listen) und  `output_word/` (Serienbriefe).
- Templates werden im Ordner `templates/` gespeichert.
- SPH‑Konfiguration (`sph_config.json`) enthält zuletzt genutzte Schule und Anmeldedaten (verschlüsselt).

## Lizenz

Dieses Projekt ist unter der **MIT‑Lizenz** veröffentlicht (siehe `LICENSE`).
