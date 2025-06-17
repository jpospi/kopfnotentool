import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import sqlite3
import threading
import logging
import sys
import os
import json
import shutil
import tempfile
import zipfile
import re
import hashlib
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT

# Logging-Konfiguration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("logs/kopfnoten_gui.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)

# Globale Konstanten
FAECHER_MAPPING = {
    "Deu": "Deutsch",
    "Mat": "Mathematik",
    "Eng": "Englisch",
    "Eth": "Ethik",
    "Ges": "Geschichte",
    "Kun": "Kunst",
    "Mus": "Musik",
    "Nawi": "Naturwissenschaften",
    "Spo": "Sport",
    "Inf": "Informatik",
    "Phy": "Physik",
    "Che": "Chemie",
    "Bio": "Biologie",
    "Gl": "Gesellschaftskunde",
    "Fra": "Französisch",
    "Spa": "Spanisch",
    "Al": "Arbeitslehre",
    "Rel": "Religion",
}


class LinuxPathManager:
    """Linux-spezifische Pfad-Verwaltung"""

    @staticmethod
    def ensure_directory(path: Path) -> Path:
        """Erstellt Verzeichnis mit korrekten Linux-Berechtigungen"""
        path = Path(path)
        try:
            path.mkdir(parents=True, exist_ok=True)
            os.chmod(path, 0o755)
            return path
        except PermissionError:
            logging.error(f"Berechtigung verweigert für: {path}")
            raise
        except Exception as e:
            logging.error(f"Fehler beim Erstellen von {path}: {e}")
            raise

    @staticmethod
    def check_file_permissions(file_path: Path) -> bool:
        """Prüft Dateiberechtigungen unter Linux"""
        file_path = Path(file_path)
        if not file_path.exists():
            return False

        try:
            with open(file_path, "rb") as f:
                f.read(1)
            return True
        except PermissionError:
            logging.warning(f"Keine Berechtigung für: {file_path}")
            return False
        except Exception:
            return False


class SimpleTemplateDesigner:
    """Vereinfachter Template-Designer ohne komplexe Validierung"""

    def __init__(self, parent):
        self.parent = parent
        self.logger = logging.getLogger("template_designer")

    def create_template_designer_window(self):
        """Öffnet Template-Designer-Fenster"""
        designer_window = tk.Toplevel(self.parent)
        designer_window.title("Template-Designer")
        designer_window.geometry("800x600")

        # Header
        header_frame = ttk.Frame(designer_window)
        header_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(
            header_frame, text="Template-Designer", font=("Arial", 16, "bold")
        ).pack(anchor=tk.W)
        ttk.Label(
            header_frame,
            text="Erstellen Sie einfache Templates für den horizontalen Export",
            font=("Arial", 10),
        ).pack(anchor=tk.W)

        # Template-Optionen
        options_frame = ttk.LabelFrame(designer_window, text="Template-Einstellungen")
        options_frame.pack(fill=tk.X, padx=10, pady=5)

        # Template-Typ Auswahl
        type_frame = ttk.Frame(options_frame)
        type_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(type_frame, text="Template-Typ:").pack(side=tk.LEFT)

        template_type = tk.StringVar(value="horizontal")
        ttk.Radiobutton(
            type_frame,
            text="Horizontal (3 Zeilen)",
            variable=template_type,
            value="horizontal",
        ).pack(side=tk.LEFT, padx=10)

        # Spaltenanzahl
        cols_frame = ttk.Frame(options_frame)
        cols_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(cols_frame, text="Max. Spalten:").pack(side=tk.LEFT)
        max_cols = tk.StringVar(value="15")
        ttk.Spinbox(cols_frame, from_=10, to=20, textvariable=max_cols, width=5).pack(
            side=tk.LEFT, padx=5
        )

        # Vorschau
        preview_frame = ttk.LabelFrame(designer_window, text="Vorschau")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        preview_text = scrolledtext.ScrolledText(
            preview_frame, height=20, font=("Courier", 10)
        )
        preview_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        def update_preview():
            """Aktualisiert die Template-Vorschau"""
            template_content = self.generate_template_content(
                template_type.get(), int(max_cols.get())
            )
            preview_text.delete(1.0, tk.END)
            preview_text.insert(1.0, template_content)

        # Buttons
        button_frame = ttk.Frame(designer_window)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(
            button_frame, text="Vorschau aktualisieren", command=update_preview
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            button_frame,
            text="Template erstellen",
            command=lambda: self.create_template_file(
                template_type.get(), int(max_cols.get()), designer_window
            ),
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            button_frame, text="Schließen", command=designer_window.destroy
        ).pack(side=tk.RIGHT, padx=5)

        # Initiale Vorschau
        update_preview()

    def generate_template_content(self, template_type: str, max_cols: int) -> str:
        """Generiert Template-Inhalt als Text-Vorschau"""
        if template_type == "horizontal":
            content = (
                """KOPFNOTEN - KLASSE {{klasse}}
Export-Datum: {{export_datum}}

Horizontale Tabelle (3 Zeilen × dynamische Spaltenanzahl):

{% for schueler in schueler_liste %}
Schüler: {{schueler.name}}

Fach:     {% for fach in schueler.faecher_spalten %}{{fach}}{% if not loop.last %} | {% endif %}{% endfor %}
AV-Note:  {% for note in schueler.av_noten %}{{note}}{% if not loop.last %} | {% endif %}{% endfor %}
SV-Note:  {% for note in schueler.sv_noten %}{{note}}{% if not loop.last %} | {% endif %}{% endfor %}

{% if not schueler.ist_letzter %}---{% endif %}
{% endfor %}
"""
            )
        else:
            content = """KOPFNOTEN - KLASSE {{klasse}}
Export-Datum: {{export_datum}}

Horizontale Tabelle (3 Zeilen × dynamische Spaltenanzahl):

{% for schueler in schueler_liste %}
Schüler: {{schueler.name}}

Fach:     {% for fach in schueler.faecher_spalten %}{{fach}}{% if not loop.last %} | {% endif %}{% endfor %}
AV-Note:  {% for note in schueler.av_noten %}{{note}}{% if not loop.last %} | {% endif %}{% endfor %}
SV-Note:  {% for note in schueler.sv_noten %}{{note}}{% if not loop.last %} | {% endif %}{% endfor %}

{% if not schueler.ist_letzter %}---{% endif %}
{% endfor %}
"""
        return content

    @staticmethod
    def create_working_horizontal_template(filename: str, max_cols: int = 15):
        """Erstellt funktionierendes horizontales Template für DocxTemplate mit optimaler Spaltenbreite
        und dynamischer Anpassung an die tatsächliche Anzahl der Fächer"""
        try:
            from docxtpl import DocxTemplate
            from docx import Document
            from docx.shared import Pt, Inches
            from docx.enum.table import WD_TABLE_ALIGNMENT

            # Neue DOCX-Datei erstellen
            doc = Document()

            # Titel
            title = doc.add_heading()
            title_run = title.add_run("KOPFNOTEN - KLASSE ")
            title_run.bold = True
            # Template-Variable für Klasse
            title.add_run("{{ klasse }}")

            # Export-Datum
            date_para = doc.add_paragraph()
            date_para.add_run("Export-Datum: ")
            date_para.add_run("{{ export_datum }}")

            # Leerzeile
            doc.add_paragraph()

            # Schüler-Schleife
            doc.add_paragraph("{% for schueler in schueler_liste %}")

            # Schüler-Name als Überschrift
            doc.add_heading("{{ schueler.name }}", level=2)

            # Dynamische Tabelle mit 3 Zeilen erstellen
            # Die Spaltenanzahl wird auf max_cols begrenzt, aber die Breite wird optimal verteilt
            table = doc.add_table(rows=3, cols=max_cols + 1)
            table.style = "Table Grid"

            # WICHTIG: Deaktiviere automatische Anpassung für volle Kontrolle über Spaltenbreite
            table.autofit = False
            table.allow_autofit = False

            # Erste Spalte für Bezeichnungen
            bezeichnungen = ["Fach", "AV", "SV"]
            for i, bezeichnung in enumerate(bezeichnungen):
                cell = table.cell(i, 0)
                cell.text = bezeichnung
                # Formatierung fett
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True

            # Zeile 1: Fach-Namen (mit Fallback für leere Spalten)
            for i in range(max_cols):
                cell = table.cell(0, i + 1)  # +1 wegen Beschriftungsspalte
                # Index-basierte Zugriffe mit sicherem Fallback
                cell.text = f'{{{{ schueler.faecher_spalten[{i}] if schueler.faecher_spalten|length > {i} else "" }}}}'

            # Zeile 2: AV-Noten
            for i in range(max_cols):
                cell = table.cell(1, i + 1)  # +1 wegen Beschriftungsspalte
                cell.text = f'{{{{ schueler.av_noten[{i}] if schueler.av_noten|length > {i} else "" }}}}'

            # Zeile 3: SV-Noten
            for i in range(max_cols):
                cell = table.cell(2, i + 1)  # +1 wegen Beschriftungsspalte
                cell.text = f'{{{{ schueler.sv_noten[{i}] if schueler.sv_noten|length > {i} else "" }}}}'

            # OPTIMIERTE SPALTENBREITENVERTEILUNG
            # Berechnung der optimalen Spaltenbreiten basierend auf verfügbarem Platz

            # Standard-Seitengeometrie (Letter-Format)
            seitenbreite_total = Inches(8.5)  # Standard Letter-Breite
            linker_rand = Inches(1.0)  # Linker Rand
            rechter_rand = Inches(1.0)  # Rechter Rand
            verfuegbare_breite = seitenbreite_total - linker_rand - rechter_rand  # = 6.5 Inches

            # Beschriftungsspalte bekommt feste, optimale Breite
            beschriftung_breite = Inches(1.0)

            # Verbleibende Breite für Fächerspalten
            faecher_breite_total = verfuegbare_breite - beschriftung_breite  # = 5.5 Inches

            # Breite pro Fächerspalte (gleichmäßig verteilt)
            fach_spalten_breite = faecher_breite_total / max_cols

            # Setze Spaltenbreiten
            for i in range(table.columns.__len__()):
                if i == 0:
                    # Beschriftungsspalte: feste optimale Breite
                    table.columns[i].width = beschriftung_breite
                else:
                    # Fächerspalten: gleichmäßig verteilte Breite
                    table.columns[i].width = fach_spalten_breite

            # Seitenumbruch zwischen Schülern (außer beim letzten)
            doc.add_paragraph("{% if not loop.last %}")
            doc.add_page_break()
            doc.add_paragraph("{% endif %}")

            # Ende der Schüler-Schleife
            doc.add_paragraph("{% endfor %}")

            # Datei speichern
            doc.save(filename)

            # Note: This function is called statically, so messagebox calls are standalone
            messagebox.showinfo(
                "Template erstellt",
                f"Template erfolgreich erstellt:\n"
                f"{Path(filename).name}\n\n"
                f"Typ: horizontal mit optimaler Spaltenbreite\n"
                f"Max. Spalten: {max_cols}\n"
                f"Beschriftungsbreite: 1.0''\n"
                f"Fächerspaltenbreite: {round(fach_spalten_breite.inches, 2)}'' pro Spalte"
            )

        except Exception as e:
            logging.error(f"Fehler beim Erstellen des Templates: {e}")
            messagebox.showerror("Template-Fehler", f"Fehler beim Erstellen: {e}")

    def create_template_file(self, template_type: str, max_cols: int, parent_window):
        """Erstellt Template-Datei über Dialog"""
        try:
            # Dateiname abfragen
            filename = filedialog.asksaveasfilename(
                title="Template speichern",
                defaultextension=".docx",
                filetypes=[("Word-Dokument", "*.docx")],
                initialdir="templates",
            )

            if not filename:
                return

            # Template erstellen
            if template_type == "horizontal":
                self.create_working_horizontal_template(filename, max_cols)
                parent_window.destroy()
            else:
                messagebox.showinfo(
                    "Info", "Vertikale Templates werden nicht unterstützt."
                )

        except Exception as e:
            self.logger.error(f"Fehler beim Erstellen der Template-Datei: {e}")
            messagebox.showerror("Template-Fehler", f"Fehler: {e}")


class KopfnotenImporter:
    """Import-Klasse für Excel-Dateien"""

    def __init__(self, db_path: str):
        self.db_path = Path(db_path)
        self.conn = None
        self.faecher_cache = {}
        self.logger = logging.getLogger("importer")

    def __enter__(self):
        self.conn = self._create_database()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.conn:
            self.conn.close()

    def _create_database(self) -> sqlite3.Connection:
        """Erstellt die Datenbank mit normalisierten Tabellen"""
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA foreign_keys = ON")

        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS schueler (
                schueler_id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                klasse TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(name, klasse)
            );
            
            CREATE TABLE IF NOT EXISTS faecher (
                fach_id INTEGER PRIMARY KEY AUTOINCREMENT,
                fach_kurz TEXT NOT NULL,
                fach_lang TEXT NOT NULL,
                fach_typ TEXT,
                ist_wahlpflicht BOOLEAN DEFAULT 0,
                wahlpflicht_gruppe TEXT,
                UNIQUE(fach_kurz, fach_typ, wahlpflicht_gruppe)
            );
            
            CREATE TABLE IF NOT EXISTS noten (
                noten_id INTEGER PRIMARY KEY AUTOINCREMENT,
                schueler_id INTEGER NOT NULL,
                fach_id INTEGER NOT NULL,
                note_av INTEGER CHECK(note_av BETWEEN 1 AND 6),
                note_sv INTEGER CHECK(note_sv BETWEEN 1 AND 6),
                ist_wahlpflicht_belegung BOOLEAN DEFAULT 0,
                schuljahr TEXT DEFAULT '2024/2025',
                halbjahr INTEGER DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (schueler_id) REFERENCES schueler(schueler_id),
                FOREIGN KEY (fach_id) REFERENCES faecher(fach_id),
                UNIQUE(schueler_id, fach_id, schuljahr, halbjahr)
            );
            
            CREATE INDEX IF NOT EXISTS idx_schueler_klasse ON schueler(klasse);
            CREATE INDEX IF NOT EXISTS idx_noten_schueler ON noten(schueler_id);
        """
        )

        conn.commit()
        return conn

    def _parse_note_mit_wahlpflicht(self, note_str: str) -> Tuple[Optional[int], bool]:
        """Extrahiert Note und Wahlpflicht-Flag aus einem Notenwert"""
        if pd.isna(note_str) or note_str == "":
            return None, False

        note_str = str(note_str).strip()
        ist_wahlpflicht = "(W)" in note_str

        if note_str.startswith("-"):
            return None, ist_wahlpflicht
        if "GB" in note_str.upper():
            return None, ist_wahlpflicht

        decimal_match = re.search(r"(\d+\.?\d*)", note_str)
        if decimal_match:
            try:
                note_float = float(decimal_match.group(1))
                note = int(round(note_float))
                if 1 <= note <= 6:
                    return note, ist_wahlpflicht
            except ValueError:
                pass

        digit_match = re.search(r"(\d)", note_str)
        if digit_match:
            try:
                note = int(digit_match.group(1))
                if 1 <= note <= 6:
                    return note, ist_wahlpflicht
            except ValueError:
                pass

        return None, ist_wahlpflicht

    def _extract_wahlpflicht_gruppe(self, fach_name: str) -> Tuple[str, Optional[str]]:
        """Extrahiert Wahlpflichtgruppe aus Fachnamen"""
        patterns = [
            (r"\(WPU1\)", "WPU1"),
            (r"\(WPU2\)", "WPU2"),
            (r"\(WPU 1\)", "WPU1"),
            (r"\(WPU 2\)", "WPU2"),
            (r"^WP1 ", "WP1"),
            (r"^WP2 ", "WP2"),
            (r"\(WPU\s*1\)", "WPU1"),
            (r"\(WPU\s*2\)", "WPU2"),
        ]

        for pattern, gruppe in patterns:
            if re.search(pattern, fach_name):
                fach_clean = re.sub(pattern, "", fach_name).strip()
                return fach_clean, gruppe

        return fach_name, None

    def _get_or_create_fach(
        self,
        fach_kurz: str,
        fach_typ: str = None,
        ist_wahlpflicht: bool = False,
        wahlpflicht_gruppe: str = None,
    ) -> int:
        """Holt oder erstellt ein Fach und gibt die ID zurück"""
        cache_key = f"{fach_kurz}_{fach_typ}_{wahlpflicht_gruppe}"
        if cache_key in self.faecher_cache:
            return self.faecher_cache[cache_key]

        fach_clean, wp_gruppe = self._extract_wahlpflicht_gruppe(fach_kurz)
        if wp_gruppe:
            wahlpflicht_gruppe = wp_gruppe
            ist_wahlpflicht = True
            fach_kurz = fach_clean

        cursor = self.conn.execute(
            """SELECT fach_id FROM faecher 
               WHERE fach_kurz = ? 
               AND (fach_typ = ? OR (fach_typ IS NULL AND ? IS NULL))
               AND (wahlpflicht_gruppe = ? OR (wahlpflicht_gruppe IS NULL AND ? IS NULL))""",
            (fach_kurz, fach_typ, fach_typ, wahlpflicht_gruppe, wahlpflicht_gruppe),
        )
        result = cursor.fetchone()

        if result:
            fach_id = result[0]
        else:
            fach_lang = FAECHER_MAPPING.get(fach_kurz, fach_kurz)
            cursor = self.conn.execute(
                """INSERT INTO faecher (fach_kurz, fach_lang, fach_typ, ist_wahlpflicht, wahlpflicht_gruppe) 
                   VALUES (?, ?, ?, ?, ?)""",
                (fach_kurz, fach_lang, fach_typ, ist_wahlpflicht, wahlpflicht_gruppe),
            )
            fach_id = cursor.lastrowid

        self.faecher_cache[cache_key] = fach_id
        return fach_id

    def _get_or_create_schueler(self, name: str, klasse: str) -> int:
        """Holt oder erstellt einen Schüler und gibt die ID zurück"""
        cursor = self.conn.execute(
            "SELECT schueler_id FROM schueler WHERE name = ? AND klasse = ?",
            (name, klasse),
        )
        result = cursor.fetchone()

        if result:
            return result[0]
        else:
            cursor = self.conn.execute(
                "INSERT INTO schueler (name, klasse) VALUES (?, ?)", (name, klasse)
            )
            return cursor.lastrowid

    def import_excel_file(self, file_path: str):
        """Importiert eine Excel-Datei"""
        file_path = Path(file_path)
        klasse = (
            file_path.stem.split("_")[-1] if "_" in file_path.stem else file_path.stem
        )

        self.logger.info(f"Importiere Datei: {file_path.name} (Klasse: {klasse})")

        try:
            df = pd.read_excel(file_path, engine="openpyxl")

            if "Name" not in df.columns or "Art" not in df.columns:
                raise ValueError("Spalten 'Name' oder 'Art' nicht gefunden")

            meta_columns = ["Name", "Art", "KN", "Abstg."]
            fach_info = []

            for idx, col in enumerate(df.columns):
                if col not in meta_columns and pd.notna(col):
                    fach_info.append((idx, col))

            fach_columns_clean = []
            rel_count = 0

            for idx, col_name in fach_info:
                if col_name == "Rel":
                    rel_count += 1
                    if rel_count == 1:
                        fach_columns_clean.append((idx, col_name, "evangelisch"))
                    else:
                        fach_columns_clean.append((idx, col_name, "katholisch"))
                else:
                    fach_columns_clean.append((idx, col_name, None))

            schueler_noten = {}

            for row_idx, row in df.iterrows():
                name = row.iloc[df.columns.get_loc("Name")]
                art = row.iloc[df.columns.get_loc("Art")]

                if pd.isna(name) or pd.isna(art):
                    continue

                if name not in schueler_noten:
                    schueler_noten[name] = {"AV": {}, "SV": {}}

                for col_idx, fach_kurz, fach_typ in fach_columns_clean:
                    note_raw = row.iloc[col_idx]
                    note, ist_wahlpflicht = self._parse_note_mit_wahlpflicht(note_raw)

                    if note is not None or ist_wahlpflicht:
                        schueler_noten[name][art][(fach_kurz, fach_typ)] = {
                            "note": note,
                            "ist_wahlpflicht": ist_wahlpflicht,
                        }

            schueler_count = 0
            noten_count = 0

            for name, noten_data in schueler_noten.items():
                schueler_id = self._get_or_create_schueler(name, klasse)
                schueler_count += 1

                faecher_gesamt = set()
                faecher_gesamt.update(noten_data.get("AV", {}).keys())
                faecher_gesamt.update(noten_data.get("SV", {}).keys())

                for fach_kurz, fach_typ in faecher_gesamt:
                    fach_id = self._get_or_create_fach(fach_kurz, fach_typ)

                    av_data = noten_data.get("AV", {}).get((fach_kurz, fach_typ), {})
                    sv_data = noten_data.get("SV", {}).get((fach_kurz, fach_typ), {})

                    note_av = av_data.get("note")
                    note_sv = sv_data.get("note")
                    ist_wahlpflicht_belegung = av_data.get(
                        "ist_wahlpflicht", False
                    ) or sv_data.get("ist_wahlpflicht", False)

                    if (
                        note_av is not None
                        or note_sv is not None
                        or ist_wahlpflicht_belegung
                    ):
                        cursor = self.conn.execute(
                            """SELECT noten_id FROM noten 
                               WHERE schueler_id = ? AND fach_id = ? 
                               AND schuljahr = '2024/2025' AND halbjahr = 1""",
                            (schueler_id, fach_id),
                        )
                        existing = cursor.fetchone()

                        if existing:
                            self.conn.execute(
                                """UPDATE noten 
                                   SET note_av = ?, note_sv = ?, ist_wahlpflicht_belegung = ?
                                   WHERE noten_id = ?""",
                                (
                                    note_av,
                                    note_sv,
                                    ist_wahlpflicht_belegung,
                                    existing[0],
                                ),
                            )
                        else:
                            self.conn.execute(
                                """INSERT INTO noten 
                                   (schueler_id, fach_id, note_av, note_sv, ist_wahlpflicht_belegung) 
                                   VALUES (?, ?, ?, ?, ?)""",
                                (
                                    schueler_id,
                                    fach_id,
                                    note_av,
                                    note_sv,
                                    ist_wahlpflicht_belegung,
                                ),
                            )
                        noten_count += 1

            self.conn.commit()
            self.logger.info(
                f"Verarbeitet: {schueler_count} Schüler mit {noten_count} Noteneinträgen"
            )

        except Exception as e:
            self.logger.error(f"Fehler beim Import von {file_path.name}: {str(e)}")
            self.conn.rollback()
            raise


class OptimizedKopfnotenExporter:
    """Optimierter Exporter für horizontale 3-Zeilen-Tabellen mit korrekter erster Spalte"""

    def __init__(self, db_path: str):
        self.db_path = Path(db_path)
        self.conn = None
        self.logger = logging.getLogger("exporter")

        if not self.db_path.exists():
            raise FileNotFoundError(f"Datenbank nicht gefunden: {self.db_path}")

    def __enter__(self):
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.conn:
            self.conn.close()

    def export_horizontal_tables(
        self, output_dir: Path, template_path: Path, klassen_liste: List[str], schueler_id: Optional[int] = None
    ) -> Dict[str, Any]:
        """Exportiert horizontale 3-Zeilen-Tabellen für ausgewählte Klassen oder einen einzelnen Schüler

        Args:
            output_dir (Path): Ausgabeverzeichnis
            template_path (Path): Template-Pfad
            klassen_liste (List[str]): Liste der Klassen zum Exportieren
            schueler_id (Optional[int]): Optionale ID eines einzelnen Schülers

        Returns:
            Dict[str, Any]: Export-Ergebnis
        """
        output_dir = Path(output_dir)
        template_path = Path(template_path)

        output_dir.mkdir(parents=True, exist_ok=True)

        summary = {
            "gesamt_dateien": 0,
            "gesamt_fehler": 0,
            "klassen_details": {},
            "schueler_details": {},
            "start_time": datetime.now(),
            "export_mode": "horizontal_optimized",
        }

        try:
            # Wenn eine Schüler-ID angegeben ist, exportiere nur diesen Schüler
            if schueler_id is not None:
                self.logger.info(f"Exportiere einzelnen Schüler: ID {schueler_id}")
                
                # Hole Schülerdaten
                cursor = self.conn.execute(
                    "SELECT name, klasse FROM schueler WHERE schueler_id = ?", 
                    (schueler_id,)
                )
                schueler_data = cursor.fetchone()
                
                if not schueler_data:
                    raise ValueError(f"Schüler mit ID {schueler_id} nicht gefunden")
                
                schueler_name = schueler_data["name"]
                klasse = schueler_data["klasse"]
                
                # Exportiere den einzelnen Schüler
                schueler_result = self._export_einzelschueler_horizontal(
                    schueler_id, schueler_name, klasse, output_dir, template_path
                )
                
                summary["schueler_details"][schueler_id] = schueler_result
                if schueler_result["datei_erstellt"]:
                    summary["gesamt_dateien"] += 1
                else:
                    summary["gesamt_fehler"] += 1
            
            # Sonst exportiere alle ausgewählten Klassen
            else:
                for klasse in klassen_liste:
                    self.logger.info(f"Exportiere Klasse horizontal: {klasse}")

                    klassen_result = self._export_klasse_horizontal_optimized(
                        klasse, output_dir, template_path
                    )

                    summary["klassen_details"][klasse] = klassen_result
                    if klassen_result["datei_erstellt"]:
                        summary["gesamt_dateien"] += 1
                    else:
                        summary["gesamt_fehler"] += 1

        except Exception as e:
            self.logger.error(f"Fehler beim Export: {e}")
            summary["gesamt_fehler"] += 1

        summary["end_time"] = datetime.now()
        summary["duration"] = summary["end_time"] - summary["start_time"]

        return summary

    def _export_klasse_horizontal_optimized(
        self, klasse: str, output_dir: Path, template_path: Path
    ) -> Dict[str, Any]:
        """Exportiert eine Klasse als horizontale Tabelle (optimiert)

        Args:
            klasse (str): Zu exportierende Klasse
            output_dir (Path): Ausgabeverzeichnis
            template_path (Path): Template-Pfad

        Returns:
            Dict[str, Any]: Export-Ergebnis
        """
        result = {
            "datei_erstellt": False,
            "output_file": None,
            "schueler_count": 0,
            "faecher_count": 0,
            "fehler": None,
        }

        try:
            # Template laden
            doc = DocxTemplate(str(template_path))

            # Schüler-Daten sammeln
            schueler_liste = self._get_schueler_horizontal_optimized(klasse)

            if not schueler_liste:
                raise ValueError(f"Keine Schüler in Klasse {klasse} gefunden")

            # Maximalwert berechnen
            max_faecher = max(s["faecher_anzahl"] for s in schueler_liste)

            # Context erstellen
            context = {
                "klasse": klasse,
                "export_datum": datetime.now().strftime("%d.%m.%Y"),
                "schueler_liste": schueler_liste,
                "max_faecher": max_faecher,
                "schueler": schueler_liste[0] if schueler_liste else None,
            }

            # Template verarbeiten
            doc.render(context)

            # Datei speichern
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = output_dir / f"Kopfnoten_{klasse}_horizontal_{timestamp}.docx"
            doc.save(output_file)

            # Ergebnis aktualisieren
            result.update(
                {
                    "datei_erstellt": True,
                    "output_file": str(output_file),
                    "schueler_count": len(schueler_liste),
                    "faecher_count": max_faecher,
                }
            )

            self.logger.info(f"Erfolgreich exportiert: {output_file.name}")

        except Exception as e:
            error_msg = f"Exportfehler {klasse}: {str(e)}"
            self.logger.error(error_msg)
            result["fehler"] = error_msg

        return result

    def _export_einzelschueler_horizontal(
        self, schueler_id: int, schueler_name: str, klasse: str, output_dir: Path, template_path: Path
    ) -> Dict[str, Any]:
        """Exportiert einen einzelnen Schüler als horizontale Tabelle

        Args:
            schueler_id (int): ID des zu exportierenden Schülers
            schueler_name (str): Name des Schülers
            klasse (str): Klasse des Schülers
            output_dir (Path): Ausgabeverzeichnis
            template_path (Path): Template-Pfad

        Returns:
            Dict[str, Any]: Export-Ergebnis
        """
        result = {
            "datei_erstellt": False,
            "output_file": None,
            "faecher_count": 0,
            "fehler": None,
        }

        try:
            # Template laden
            doc = DocxTemplate(str(template_path))

            # Schüler-Daten sammeln
            schueler_data = self._get_einzelschueler_horizontal(schueler_id, schueler_name, klasse)

            if not schueler_data:
                raise ValueError(f"Keine Daten für Schüler {schueler_name} (ID: {schueler_id}) gefunden")

            # Liste mit nur diesem Schüler erstellen
            schueler_liste = [schueler_data]
            
            # Context erstellen
            context = {
                "klasse": klasse,
                "export_datum": datetime.now().strftime("%d.%m.%Y"),
                "schueler_liste": schueler_liste,
                "max_faecher": schueler_data["faecher_anzahl"],
                "schueler": schueler_data,
            }

            # Template verarbeiten
            doc.render(context)

            # Datei speichern
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = output_dir / f"Kopfnoten_{schueler_name.replace(' ', '_')}_{timestamp}.docx"
            doc.save(output_file)

            # Ergebnis aktualisieren
            result.update(
                {
                    "datei_erstellt": True,
                    "output_file": str(output_file),
                    "faecher_count": schueler_data["faecher_anzahl"],
                }
            )

            self.logger.info(f"Erfolgreich exportiert: {output_file.name}")

        except Exception as e:
            error_msg = f"Exportfehler für Schüler {schueler_name}: {str(e)}"
            self.logger.error(error_msg)
            result["fehler"] = error_msg

        return result

    def _get_einzelschueler_horizontal(
        self, schueler_id: int, schueler_name: str, klasse: str
    ) -> Dict[str, Any]:
        """Sammelt Daten für einen einzelnen Schüler für horizontale Darstellung"""
        # Fächer für diesen Schüler laden
        cursor = self.conn.execute(
            """
            SELECT 
                f.fach_kurz,
                n.note_av,
                n.note_sv,
                n.ist_wahlpflicht_belegung,
                f.wahlpflicht_gruppe
            FROM noten n
            JOIN faecher f ON n.fach_id = f.fach_id
            WHERE n.schueler_id = ?
            ORDER BY f.fach_lang
            """,
            (schueler_id,),
        )

        # Arrays für horizontale Darstellung
        faecher_spalten = []
        av_noten = []
        sv_noten = []

        for row in cursor.fetchall():
            fach_kurz = row["fach_kurz"]

            # Wahlpflicht-Kennzeichnung
            if row["ist_wahlpflicht_belegung"]:
                if row["wahlpflicht_gruppe"]:
                    fach_kurz += f"({row['wahlpflicht_gruppe']})"
                else:
                    fach_kurz += "(W)"

            faecher_spalten.append(fach_kurz)

            # Noten formatieren
            av_note = str(row["note_av"]) if row["note_av"] else "-"
            sv_note = str(row["note_sv"]) if row["note_sv"] else "-"

            av_noten.append(av_note)
            sv_noten.append(sv_note)

        schueler_data = {
            "name": schueler_name,
            "klasse": klasse,
            "faecher_spalten": faecher_spalten,
            "av_noten": av_noten,
            "sv_noten": sv_noten,
            "faecher_anzahl": len(faecher_spalten),
            "ist_letzter": True,  # Beim Einzelschüler gibt es nur einen Schüler
        }

        return schueler_data

    def _get_schueler_horizontal_optimized(self, klasse: str) -> List[Dict[str, Any]]:
        """Sammelt Schülerdaten für optimierte horizontale Darstellung"""
        cursor = self.conn.execute(
            """
            SELECT schueler_id, name 
            FROM schueler 
            WHERE klasse = ? 
            ORDER BY name
        """,
            (klasse,),
        )

        schueler_liste = []
        schueler_rows = cursor.fetchall()

        for i, schueler in enumerate(schueler_rows):
            schueler_id = schueler["schueler_id"]
            schueler_name = schueler["name"]

            # Fächer für diesen Schüler laden
            cursor = self.conn.execute(
                """
                SELECT 
                    f.fach_kurz,
                    n.note_av,
                    n.note_sv,
                    n.ist_wahlpflicht_belegung,
                    f.wahlpflicht_gruppe
                FROM noten n
                JOIN faecher f ON n.fach_id = f.fach_id
                WHERE n.schueler_id = ?
                ORDER BY f.fach_lang
            """,
                (schueler_id,),
            )

            # Arrays für horizontale Darstellung
            faecher_spalten = []
            av_noten = []
            sv_noten = []

            for row in cursor.fetchall():
                fach_kurz = row["fach_kurz"]

                # Wahlpflicht-Kennzeichnung
                if row["ist_wahlpflicht_belegung"]:
                    if row["wahlpflicht_gruppe"]:
                        fach_kurz += f"({row['wahlpflicht_gruppe']})"
                    else:
                        fach_kurz += "(W)"

                faecher_spalten.append(fach_kurz)

                # Noten formatieren
                av_note = str(row["note_av"]) if row["note_av"] else "-"
                sv_note = str(row["note_sv"]) if row["note_sv"] else "-"

                av_noten.append(av_note)
                sv_noten.append(sv_note)

            schueler_data = {
                "name": schueler_name,
                "klasse": klasse,
                "faecher_spalten": faecher_spalten,
                "av_noten": av_noten,
                "sv_noten": sv_noten,
                "faecher_anzahl": len(faecher_spalten),
                "ist_letzter": (i == len(schueler_rows) - 1),
            }

            schueler_liste.append(schueler_data)

        return schueler_liste


class SimplifiedGradeEditor:
    """Vereinfachter Noten-Editor"""

    def __init__(self, parent, db_path: str):
        self.parent = parent
        self.db_path = db_path
        self.logger = logging.getLogger("grade_editor")

    def open_grade_editor(self, student_id: int, student_name: str, student_class: str):
        """Öffnet vereinfachten Noten-Editor"""
        try:
            # Editor-Fenster erstellen
            editor_window = tk.Toplevel(self.parent)
            editor_window.title(f"Noten bearbeiten - {student_name} ({student_class})")
            editor_window.geometry("700x500")
            editor_window.resizable(True, True)

            # Header
            header_frame = ttk.Frame(editor_window)
            header_frame.pack(fill=tk.X, padx=10, pady=5)

            ttk.Label(
                header_frame,
                text=f"Schüler: {student_name}",
                font=("Arial", 14, "bold"),
            ).pack(side=tk.LEFT)
            ttk.Label(
                header_frame, text=f"Klasse: {student_class}", font=("Arial", 12)
            ).pack(side=tk.RIGHT)

            # Noten-Frame
            notes_frame = ttk.LabelFrame(editor_window, text="Noten bearbeiten")
            notes_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

            # Lade aktuelle Noten
            grades_data = self._load_student_grades(student_id)

            # Erstelle Eingabefelder
            self._create_grade_inputs(
                notes_frame, grades_data, student_id, editor_window
            )

        except Exception as e:
            self.logger.error(f"Fehler beim Öffnen des Noten-Editors: {e}")
            messagebox.showerror("Editor-Fehler", f"Fehler: {e}")

    def _load_student_grades(self, student_id: int) -> List[Dict]:
        """Lädt Noten eines Schülers (korrigierte Version)"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row  # Für Spaltenzugriff per Name
                cursor = conn.execute(
                    """
                    SELECT 
                        n.noten_id,
                        f.fach_lang,
                        f.fach_kurz,
                        n.note_av,
                        n.note_sv,
                        n.ist_wahlpflicht_belegung,
                        f.wahlpflicht_gruppe
                    FROM noten n
                    JOIN faecher f ON n.fach_id = f.fach_id
                    WHERE n.schueler_id = ?
                    ORDER BY f.fach_lang
                """,
                    (student_id,),
                )

                return [dict(row) for row in cursor.fetchall()]
        except Exception as e:
            self.logger.error(f"Fehler beim Laden der Noten: {e}")
            return []

    def _create_grade_inputs(
        self, parent_frame, grades_data: List[Dict], student_id: int, editor_window
    ):
        """Erstellt Eingabefelder für Noten"""
        # Scrollbarer Frame
        canvas = tk.Canvas(parent_frame)
        scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Header für Eingabefelder
        header_frame = ttk.Frame(scrollable_frame)
        header_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(header_frame, text="Fach", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, padx=5
        )
        ttk.Label(header_frame, text="AV-Note", font=("Arial", 10, "bold")).grid(
            row=0, column=1, padx=5
        )
        ttk.Label(header_frame, text="SV-Note", font=("Arial", 10, "bold")).grid(
            row=0, column=2, padx=5
        )
        ttk.Label(header_frame, text="Wahlpflicht", font=("Arial", 10, "bold")).grid(
            row=0, column=3, padx=5
        )

        # Eingabefelder für jedes Fach
        grade_vars = {}

        for i, grade in enumerate(grades_data, 1):
            row_frame = ttk.Frame(scrollable_frame)
            row_frame.pack(fill=tk.X, padx=5, pady=2)

            # Fachname
            fach_name = grade["fach_lang"]
            if grade["wahlpflicht_gruppe"]:
                fach_name += f" ({grade['wahlpflicht_gruppe']})"

            ttk.Label(row_frame, text=fach_name, width=20).grid(
                row=0, column=0, sticky=tk.W, padx=5
            )

            # AV-Note
            av_var = tk.StringVar(
                value=str(grade["note_av"]) if grade["note_av"] else ""
            )
            av_spinbox = ttk.Spinbox(
                row_frame, from_=1, to=6, textvariable=av_var, width=5
            )
            av_spinbox.grid(row=0, column=1, padx=5)

            # SV-Note
            sv_var = tk.StringVar(
                value=str(grade["note_sv"]) if grade["note_sv"] else ""
            )
            sv_spinbox = ttk.Spinbox(
                row_frame, from_=1, to=6, textvariable=sv_var, width=5
            )
            sv_spinbox.grid(row=0, column=2, padx=5)

            # Wahlpflicht-Checkbox
            wp_var = tk.BooleanVar(value=grade["ist_wahlpflicht_belegung"])
            wp_check = ttk.Checkbutton(row_frame, variable=wp_var)
            wp_check.grid(row=0, column=3, padx=5)

            # Speichere Variablen für späteren Zugriff
            grade_vars[grade["noten_id"]] = {
                "av_var": av_var,
                "sv_var": sv_var,
                "wp_var": wp_var,
                "fach_name": fach_name,
            }

        # Scrollbar und Canvas positionieren
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Buttons
        button_frame = ttk.Frame(editor_window)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        def save_all_grades():
            """Speichert alle Noten"""
            try:
                with sqlite3.connect(self.db_path) as conn:
                    saved_count = 0
                    for noten_id, vars_dict in grade_vars.items():
                        av_text = vars_dict["av_var"].get().strip()
                        sv_text = vars_dict["sv_var"].get().strip()
                        wp_value = vars_dict["wp_var"].get()

                        # Validiere und konvertiere Noten
                        av_value = None
                        sv_value = None

                        if av_text:
                            try:
                                av_value = int(av_text)
                                if not (1 <= av_value <= 6):
                                    raise ValueError(
                                        f"AV-Note für {vars_dict['fach_name']} muss zwischen 1 und 6 liegen"
                                    )
                            except ValueError as e:
                                messagebox.showerror("Ungültige Eingabe", str(e))
                                return

                        if sv_text:
                            try:
                                sv_value = int(sv_text)
                                if not (1 <= sv_value <= 6):
                                    raise ValueError(
                                        f"SV-Note für {vars_dict['fach_name']} muss zwischen 1 und 6 liegen"
                                    )
                            except ValueError as e:
                                messagebox.showerror("Ungültige Eingabe", str(e))
                                return

                        # Speichere in Datenbank
                        conn.execute(
                            """
                            UPDATE noten 
                            SET note_av = ?, note_sv = ?, ist_wahlpflicht_belegung = ?
                            WHERE noten_id = ?
                        """,
                            (av_value, sv_value, wp_value, noten_id),
                        )

                        saved_count += 1

                    conn.commit()

                    messagebox.showinfo(
                        "Gespeichert", f"{saved_count} Noten erfolgreich gespeichert!"
                    )
                    editor_window.destroy()

                    # Refresh parent window if possible
                    if hasattr(self.parent, "refresh_analysis_data"):
                        self.parent.refresh_analysis_data()

            except Exception as e:
                self.logger.error(f"Fehler beim Speichern der Noten: {e}")
                messagebox.showerror("Speicher-Fehler", f"Fehler: {e}")

        ttk.Button(
            button_frame, text="Alle Noten speichern", command=save_all_grades
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Abbrechen", command=editor_window.destroy).pack(
            side=tk.RIGHT, padx=5
        )


class StatusManager:
    """Einfache Status-Verwaltung"""

    def __init__(self, parent):
        self.parent = parent
        self.status_label = None
        self.progress_bar = None

    def setup_status_bar(self, parent_frame):
        """Erstellt Statusleiste"""
        status_frame = ttk.Frame(parent_frame)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=2)

        self.status_label = ttk.Label(status_frame, text="Bereit", relief=tk.SUNKEN)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.progress_bar = ttk.Progressbar(status_frame, mode="indeterminate")
        self.progress_bar.pack(side=tk.RIGHT, padx=(5, 0))

        return status_frame

    def set_status(self, message: str, progress: bool = False):
        """Setzt Status"""
        if self.status_label:
            self.status_label.config(
                text=f"{datetime.now().strftime('%H:%M:%S')} - {message}"
            )

        if progress and self.progress_bar:
            self.progress_bar.start()
        elif not progress and self.progress_bar:
            self.progress_bar.stop()

    def clear_status(self):
        """Setzt Status zurück"""
        self.set_status("Bereit", False)


class KopfnotenGUI:
    """Hauptklasse der optimierten GUI-Anwendung"""

    def __init__(self):
        self.root = tk.Tk()
        self.setup_application()

        # Manager
        self.path_manager = LinuxPathManager()
        self.status_manager = StatusManager(self)
        self.template_designer = SimpleTemplateDesigner(self.root)

        # Pfade
        self.db_path = Path("output_database/kopfnoten_secure.db")

        # GUI-Variablen
        self.template_var = tk.StringVar()
        self.output_var = tk.StringVar(value="output_word")
        self.export_running = False

        # GUI-Komponenten
        self.notebook = None
        self.import_listbox = None
        self.export_listbox = None
        self.export_log = None
        self.analysis_tree = None
        self.stats_text = None
        self.selected_schueler_var = tk.StringVar(value="")

        # Setup
        self.create_gui()
        self.load_initial_data()
        self.setup_linux_environment()

    def setup_application(self):
        """Grundlegende Anwendungseinrichtung"""
        self.root.title("AES-Kopfnoten-Manager v2.0 - Optimiert")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Style für bessere Optik
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")

    def setup_linux_environment(self):
        """Linux-spezifische Umgebungseinrichtung"""
        directories = [
            Path("logs"),
            Path("templates"),
            Path("output_word"),
            Path("output_database"),
            Path("temp"),
        ]

        for directory in directories:
            try:
                self.path_manager.ensure_directory(directory)
                logging.info(f"Verzeichnis bereit: {directory}")
            except Exception as e:
                logging.error(f"Fehler beim Erstellen von {directory}: {e}")

    def create_gui(self):
        """Erstellt die GUI"""
        # Menü
        self.create_menu()

        # Hauptcontainer
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Status-Bar
        self.status_manager.setup_status_bar(self.root)

        # Notebook für Tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Tabs erstellen
        self.create_import_tab()
        self.create_analysis_tab()
        self.create_export_tab()
        self.create_template_tab()

        # Initial Status
        self.status_manager.set_status("Anwendung gestartet")

    def create_menu(self):
        """Erstellt vereinfachtes Menü"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Datei-Menü
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Datei", menu=file_menu)
        file_menu.add_command(label="Datenbank öffnen", command=self.open_database)
        file_menu.add_command(label="Datenbank-Info", command=self.show_database_info)
        file_menu.add_separator()
        file_menu.add_command(label="Beenden", command=self.root.quit)

        # Werkzeuge-Menü
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Werkzeuge", menu=tools_menu)
        tools_menu.add_command(
            label="Template-Designer",
            command=self.template_designer.create_template_designer_window,
        )
        tools_menu.add_command(label="Logs anzeigen", command=self.show_logs)
        tools_menu.add_separator()
        tools_menu.add_command(
            label="Berechtigungen prüfen", command=self.check_permissions
        )

        # Hilfe-Menü
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Hilfe", menu=help_menu)
        help_menu.add_command(label="Über", command=self.show_about)
        help_menu.add_command(label="Linux-Hilfe", command=self.show_linux_help)

    def create_import_tab(self):
        """Erstellt Import-Tab"""
        import_frame = ttk.Frame(self.notebook)
        self.notebook.add(import_frame, text="📥 Import")

        # Header
        header_frame = ttk.LabelFrame(import_frame, text="Excel-Dateien importieren")
        header_frame.pack(fill=tk.X, padx=10, pady=5)

        # Buttons
        button_frame = ttk.Frame(header_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(
            button_frame,
            text="Excel-Dateien auswählen",
            command=self.select_excel_files,
        ).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(
            button_frame, text="Alle importieren", command=self.import_all_files
        ).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(
            button_frame, text="Auswahl löschen", command=self.clear_import_selection
        ).pack(side=tk.LEFT)

        # Liste
        list_frame = ttk.LabelFrame(import_frame, text="Ausgewählte Dateien")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        listbox_frame = ttk.Frame(list_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.import_listbox = tk.Listbox(listbox_frame, selectmode=tk.EXTENDED)
        scrollbar_import = ttk.Scrollbar(
            listbox_frame, orient=tk.VERTICAL, command=self.import_listbox.yview
        )
        self.import_listbox.config(yscrollcommand=scrollbar_import.set)

        self.import_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_import.pack(side=tk.RIGHT, fill=tk.Y)

        # Import-Log
        log_frame = ttk.LabelFrame(import_frame, text="Import-Log")
        log_frame.pack(fill=tk.X, padx=10, pady=5)

        self.import_log = scrolledtext.ScrolledText(
            log_frame, height=8, state=tk.DISABLED
        )
        self.import_log.pack(fill=tk.X, padx=5, pady=5)

    def create_analysis_tab(self):
        """Erstellt Analyse-Tab"""
        analysis_frame = ttk.Frame(self.notebook)
        self.notebook.add(analysis_frame, text="🔍 Analyse")

        # Filter
        filter_frame = ttk.LabelFrame(analysis_frame, text="Filter und Suche")
        filter_frame.pack(fill=tk.X, padx=10, pady=5)

        filter_controls = ttk.Frame(filter_frame)
        filter_controls.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(filter_controls, text="Klasse:").pack(side=tk.LEFT)
        self.class_filter = ttk.Combobox(filter_controls, width=10, state="readonly")
        self.class_filter.pack(side=tk.LEFT, padx=(5, 15))

        ttk.Label(filter_controls, text="Schüler:").pack(side=tk.LEFT)
        self.student_search = ttk.Entry(filter_controls, width=20)
        self.student_search.pack(side=tk.LEFT, padx=(5, 15))

        ttk.Button(filter_controls, text="Suchen", command=self.search_students).pack(
            side=tk.LEFT, padx=(0, 5)
        )
        ttk.Button(
            filter_controls, text="Filter zurücksetzen", command=self.reset_filters
        ).pack(side=tk.LEFT)

        # Daten
        data_frame = ttk.LabelFrame(analysis_frame, text="Schüler und Noten")
        data_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        tree_frame = ttk.Frame(data_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.analysis_tree = ttk.Treeview(tree_frame, show="tree headings")
        v_scrollbar = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.analysis_tree.yview
        )
        h_scrollbar = ttk.Scrollbar(
            tree_frame, orient=tk.HORIZONTAL, command=self.analysis_tree.xview
        )

        self.analysis_tree.config(
            yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set
        )

        self.analysis_tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Bearbeitung
        edit_frame = ttk.Frame(data_frame)
        edit_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(
            edit_frame, text="Noten bearbeiten", command=self.edit_selected_grade
        ).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(
            edit_frame, text="Schüler löschen", command=self.delete_selected_student
        ).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(
            edit_frame, text="Daten aktualisieren", command=self.refresh_analysis_data
        ).pack(side=tk.LEFT)
        
        # Neue Schaltfläche für Einzelschüler-Export
        ttk.Button(
            edit_frame, text="Schüler exportieren", command=self.export_selected_student
        ).pack(side=tk.RIGHT, padx=5)

    def create_export_tab(self):
        """Erstellt vereinfachten Export-Tab"""
        export_frame = ttk.Frame(self.notebook)
        self.notebook.add(export_frame, text="📤 Export")

        # Export-Optionen
        options_frame = ttk.LabelFrame(export_frame, text="Export-Optionen")
        options_frame.pack(fill=tk.X, padx=10, pady=5)

        # Template-Auswahl
        template_frame = ttk.Frame(options_frame)
        template_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(template_frame, text="Template:").pack(side=tk.LEFT)
        template_entry = ttk.Entry(
            template_frame, textvariable=self.template_var, width=50
        )
        template_entry.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)

        ttk.Button(
            template_frame, text="Durchsuchen", command=self.select_template
        ).pack(side=tk.RIGHT, padx=(5, 0))

        # Ausgabe-Verzeichnis
        output_frame = ttk.Frame(options_frame)
        output_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(output_frame, text="Ausgabe:").pack(side=tk.LEFT)
        output_entry = ttk.Entry(output_frame, textvariable=self.output_var, width=50)
        output_entry.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)

        ttk.Button(
            output_frame, text="Durchsuchen", command=self.select_output_dir
        ).pack(side=tk.RIGHT)

        # Klassenauswahl
        class_frame = ttk.LabelFrame(export_frame, text="Klassenauswahl")
        class_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        class_controls = ttk.Frame(class_frame)
        class_controls.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(
            class_controls, text="Alle auswählen", command=self.select_all_classes
        ).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(
            class_controls, text="Auswahl umkehren", command=self.invert_class_selection
        ).pack(side=tk.LEFT, padx=(0, 5))

        # Export-Button
        self.export_btn = ttk.Button(
            class_controls,
            text="🚀 Horizontalen Export starten",
            command=self.start_optimized_export,
        )
        self.export_btn.pack(side=tk.RIGHT, padx=(5, 0))

        # Content
        content_frame = ttk.Frame(class_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Klassen-Liste
        list_frame = ttk.Frame(content_frame)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        ttk.Label(list_frame, text="Verfügbare Klassen:").pack(anchor=tk.W)

        self.export_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED)
        list_scrollbar = ttk.Scrollbar(
            list_frame, orient=tk.VERTICAL, command=self.export_listbox.yview
        )
        self.export_listbox.config(yscrollcommand=list_scrollbar.set)

        self.export_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Export-Log
        log_frame = ttk.Frame(content_frame)
        log_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))

        ttk.Label(log_frame, text="Export-Log:").pack(anchor=tk.W)

        self.export_log = scrolledtext.ScrolledText(
            log_frame, width=50, state=tk.DISABLED
        )
        self.export_log.pack(fill=tk.BOTH, expand=True)

        # Progress-Bar
        self.export_progress = ttk.Progressbar(export_frame, mode="indeterminate")
        self.export_progress.pack(fill=tk.X, padx=10, pady=5)

    def create_template_tab(self):
        """Erstellt Template-Tab"""
        template_frame = ttk.Frame(self.notebook)
        self.notebook.add(template_frame, text="📝 Templates")

        # Template-Designer
        designer_frame = ttk.LabelFrame(template_frame, text="Template-Designer")
        designer_frame.pack(fill=tk.X, padx=10, pady=5)

        designer_text = ttk.Label(
            designer_frame,
            text="Erstellen Sie einfache Templates für den horizontalen Export.\n"
            "Der Designer hilft Ihnen bei der Erstellung von Templates ohne komplexe Validierung.",
        )
        designer_text.pack(padx=10, pady=10)

        ttk.Button(
            designer_frame,
            text="Template-Designer öffnen",
            command=self.template_designer.create_template_designer_window,
        ).pack(pady=5)

        # Template-Übersicht
        overview_frame = ttk.LabelFrame(template_frame, text="Verfügbare Templates")
        overview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.template_list = tk.Listbox(overview_frame)
        template_scroll = ttk.Scrollbar(
            overview_frame, orient=tk.VERTICAL, command=self.template_list.yview
        )
        self.template_list.config(yscrollcommand=template_scroll.set)

        self.template_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        template_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=5)

        # Template-Buttons
        template_buttons = ttk.Frame(overview_frame)
        template_buttons.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(
            template_buttons, text="Aktualisieren", command=self.refresh_template_list
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            template_buttons,
            text="Template verwenden",
            command=self.use_selected_template,
        ).pack(side=tk.LEFT, padx=5)

    # ===================== NEUE FUNKTIONEN =====================
    
    def export_selected_student(self):
        """Exportiert ausgewählten Schüler"""
        selection = self.analysis_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Keine Auswahl", "Bitte wählen Sie einen Schüler aus."
            )
            return

        item = self.analysis_tree.item(selection[0])
        values = item["values"]
        student_id = values[0]
        student_name = values[1]
        student_class = values[2]

        # Template-Prüfung
        template_path = Path(self.template_var.get().strip())
        if not template_path or not template_path.exists():
            messagebox.showerror(
                "Template fehlt", "Bitte wählen Sie eine gültige Template-Datei aus."
            )
            # Zur Export-Registerkarte wechseln für Template-Auswahl
            self.notebook.select(2)
            return

        # Ausgabeverzeichnis prüfen
        output_dir = Path(self.output_var.get().strip())
        if not output_dir:
            output_dir = Path("output_word")
            self.output_var.set(str(output_dir))

        try:
            self.path_manager.ensure_directory(output_dir)
        except Exception as e:
            messagebox.showerror(
                "Ausgabe-Fehler",
                f"Ausgabeverzeichnis konnte nicht erstellt werden:\n{e}",
            )
            return

        # Export-UI vorbereiten
        self.export_running = True
        self.status_manager.set_status(f"Exportiere Schüler: {student_name}...", True)
        
        # Export in separatem Thread
        export_thread = threading.Thread(
            target=self.run_student_export,
            args=(student_id, student_name, student_class, template_path, output_dir),
            daemon=True,
        )
        export_thread.start()

    def run_student_export(
        self, 
        student_id: int, 
        student_name: str, 
        student_class: str, 
        template_path: Path, 
        output_dir: Path
    ):
        """Führt Schüler-Export in separatem Thread aus"""
        try:
            with OptimizedKopfnotenExporter(self.db_path) as exporter:
                start_time = datetime.now()
                
                # Export durchführen für einzelnen Schüler
                summary = exporter.export_horizontal_tables(
                    output_dir, template_path, [student_class], student_id
                )
                
                end_time = datetime.now()
                duration = end_time - start_time
                
                # Erfolgs-Meldung
                if summary["gesamt_dateien"] > 0:
                    schueler_result = list(summary["schueler_details"].values())[0]
                    success_msg = (
                        f"✅ Export für {student_name} erfolgreich!\n\n"
                        f"Dauer: {duration.total_seconds():.1f} Sekunden\n"
                        f"Datei: {Path(schueler_result['output_file']).name}\n"
                        f"Ausgabe: {output_dir}\n\n"
                        f"Format: Horizontale Tabelle"
                    )
                    
                    # GUI-Thread für MessageBox
                    self.root.after(
                        100, lambda: messagebox.showinfo(f"Export für {student_name}", success_msg)
                    )
                else:
                    error_msg = f"❌ Export für {student_name} fehlgeschlagen"
                    self.root.after(
                        100, lambda: messagebox.showerror("Export-Fehler", error_msg)
                    )
                    
        except Exception as e:
            error_msg = f"❌ Export für {student_name} fehlgeschlagen: {str(e)}"
            logging.error(f"Student-Export-Fehler: {e}")
            
            self.root.after(
                100, lambda: messagebox.showerror("Export-Fehler", error_msg)
            )
            
        finally:
            self.root.after(100, lambda: self.status_manager.clear_status())
            self.root.after(100, lambda: setattr(self, "export_running", False))

    # ===================== OPTIMIERTE FUNKTIONEN =====================

    def start_optimized_export(self):
        """Startet optimierten horizontalen Export"""
        if self.export_running:
            messagebox.showwarning("Export läuft", "Es läuft bereits ein Export!")
            return

        # Validierungen
        selected_indices = self.export_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning(
                "Keine Auswahl", "Bitte wählen Sie mindestens eine Klasse aus."
            )
            return

        template_path = Path(self.template_var.get().strip())
        if not template_path or not template_path.exists():
            messagebox.showerror(
                "Template fehlt", "Bitte wählen Sie eine gültige Template-Datei aus."
            )
            return

        output_dir = Path(self.output_var.get().strip())
        if not output_dir:
            output_dir = Path("output_word")
            self.output_var.set(str(output_dir))

        try:
            self.path_manager.ensure_directory(output_dir)
        except Exception as e:
            messagebox.showerror(
                "Ausgabe-Fehler",
                f"Ausgabeverzeichnis konnte nicht erstellt werden:\n{e}",
            )
            return

        selected_classes = [self.export_listbox.get(i) for i in selected_indices]

        # Export-UI vorbereiten
        self.export_running = True
        self.export_btn.config(state=tk.DISABLED, text="Export läuft...")
        self.export_progress.start()
        self.clear_export_log()

        self.log_to_export(
            f"Starte optimierten horizontalen Export für {len(selected_classes)} Klassen"
        )
        self.status_manager.set_status("Optimierter Export läuft...", True)

        # Export in separatem Thread
        export_thread = threading.Thread(
            target=self.run_optimized_export,
            args=(selected_classes, template_path, output_dir),
            daemon=True,
        )
        export_thread.start()

    def run_optimized_export(
        self, klassen: List[str], template_path: Path, output_dir: Path
    ):
        """Führt optimierten Export aus"""
        try:
            with OptimizedKopfnotenExporter(self.db_path) as exporter:
                start_time = datetime.now()
                self.log_to_export(f"Start: {start_time.strftime('%H:%M:%S')}")

                # Export durchführen
                summary = exporter.export_horizontal_tables(
                    output_dir, template_path, klassen
                )

                end_time = datetime.now()
                duration = end_time - start_time

                # Detaillierte Log-Ausgabe
                self.log_to_export(f"\nOptimierter Export Details:")
                for klasse, details in summary["klassen_details"].items():
                    if details["datei_erstellt"]:
                        self.log_to_export(
                            f"✅ {klasse}: {details['schueler_count']} Schüler, "
                            f"max. {details['faecher_count']} Fächer"
                        )
                        self.log_to_export(
                            f"   Datei: {Path(details['output_file']).name}"
                        )
                    else:
                        self.log_to_export(
                            f"❌ {klasse}: {details.get('fehler', 'Unbekannter Fehler')}"
                        )

                # Erfolg-Meldung
                success_msg = (
                    f"✅ Optimierter horizontaler Export erfolgreich!\n\n"
                    f"Dateien erstellt: {summary['gesamt_dateien']}\n"
                    f"Fehler: {summary['gesamt_fehler']}\n"
                    f"Dauer: {duration.total_seconds():.1f} Sekunden\n"
                    f"Ausgabe: {output_dir}\n\n"
                    f"Format: Horizontale Tabellen mit Beschriftungsspalte"
                )

                self.log_to_export(success_msg)

                # GUI-Thread für MessageBox
                self.root.after(
                    100, lambda: messagebox.showinfo("Export erfolgreich", success_msg)
                )

        except Exception as e:
            error_msg = f"❌ Optimierter Export fehlgeschlagen: {str(e)}"
            self.log_to_export(error_msg)
            logging.error(f"Export-Fehler: {e}")

            self.root.after(
                100, lambda: messagebox.showerror("Export-Fehler", error_msg)
            )

        finally:
            self.root.after(100, self.reset_export_ui)

    def edit_selected_grade(self):
        """Öffnet vereinfachten Noten-Editor"""
        selection = self.analysis_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Keine Auswahl", "Bitte wählen Sie einen Schüler aus."
            )
            return

        item = self.analysis_tree.item(selection[0])
        values = item["values"]
        student_id = values[0]
        student_name = values[1]
        student_class = values[2]

        # Verwende vereinfachten Editor
        grade_editor = SimplifiedGradeEditor(self.root, str(self.db_path))
        grade_editor.open_grade_editor(student_id, student_name, student_class)

    def refresh_template_list(self):
        """Aktualisiert Template-Liste"""
        self.template_list.delete(0, tk.END)

        template_dir = Path("templates")
        if template_dir.exists():
            for template_file in template_dir.glob("*.docx"):
                self.template_list.insert(tk.END, template_file.name)

    def use_selected_template(self):
        """Verwendet ausgewähltes Template"""
        selection = self.template_list.curselection()
        if not selection:
            messagebox.showwarning(
                "Keine Auswahl", "Bitte wählen Sie ein Template aus."
            )
            return

        template_name = self.template_list.get(selection[0])
        template_path = Path("templates") / template_name

        if template_path.exists():
            self.template_var.set(str(template_path))
            messagebox.showinfo(
                "Template gewählt", f"Template ausgewählt: {template_name}"
            )
            # Wechsle zum Export-Tab
            self.notebook.select(2)

    # ===================== UTILITY-FUNKTIONEN =====================

    def reset_export_ui(self):
        """Setzt Export-UI zurück"""
        self.export_running = False
        self.export_btn.config(state=tk.NORMAL, text="🚀 Horizontalen Export starten")
        self.export_progress.stop()
        self.status_manager.clear_status()

    def select_template(self):
        """Template-Datei auswählen"""
        filename = filedialog.askopenfilename(
            title="Template-Datei auswählen",
            filetypes=[("Word-Dokumente", "*.docx"), ("Alle Dateien", "*.*")],
            initialdir=str(Path("templates").resolve()),
        )

        if filename:
            self.template_var.set(filename)
            self.log_to_export(f"Template ausgewählt: {Path(filename).name}")

    def select_output_dir(self):
        """Ausgabeverzeichnis auswählen"""
        dirname = filedialog.askdirectory(
            title="Ausgabeverzeichnis auswählen",
            initialdir=str(Path("output_word").resolve()),
        )

        if dirname:
            self.output_var.set(dirname)
            self.log_to_export(f"Ausgabeverzeichnis: {dirname}")

    def select_excel_files(self):
        """Excel-Dateien für Import auswählen"""
        filenames = filedialog.askopenfilenames(
            title="Excel-Dateien auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xls"), ("Alle Dateien", "*.*")],
        )

        if filenames:
            for filename in filenames:
                if filename not in self.import_listbox.get(0, tk.END):
                    self.import_listbox.insert(tk.END, filename)

            self.log_to_import(f"{len(filenames)} Datei(en) hinzugefügt")

    def import_all_files(self):
        """Importiert alle ausgewählten Excel-Dateien"""
        if self.import_listbox.size() == 0:
            messagebox.showwarning(
                "Keine Dateien", "Bitte wählen Sie zuerst Excel-Dateien aus."
            )
            return

        files = list(self.import_listbox.get(0, tk.END))
        self.status_manager.set_status(f"Importiere {len(files)} Dateien...", True)

        import_thread = threading.Thread(
            target=self.run_import, args=(files,), daemon=True
        )
        import_thread.start()

    def run_import(self, files: List[str]):
        """Führt Import in separatem Thread aus"""
        try:
            self.path_manager.ensure_directory(self.db_path.parent)

            with KopfnotenImporter(str(self.db_path)) as importer:
                successful = 0
                for file_path in files:
                    try:
                        self.log_to_import(f"Importiere: {Path(file_path).name}")
                        importer.import_excel_file(file_path)
                        successful += 1
                        self.log_to_import(f"✅ Erfolgreich: {Path(file_path).name}")
                    except Exception as e:
                        self.log_to_import(f"❌ Fehler bei {Path(file_path).name}: {e}")
                        logging.error(f"Import-Fehler für {file_path}: {e}")

                self.log_to_import(
                    f"\nImport abgeschlossen: {successful}/{len(files)} erfolgreich"
                )
                self.root.after(100, self.refresh_all_data)

        except Exception as e:
            self.log_to_import(f"❌ Import-Fehler: {e}")
            logging.error(f"Allgemeiner Import-Fehler: {e}")
        finally:
            self.root.after(100, lambda: self.status_manager.clear_status())

    def clear_import_selection(self):
        """Löscht Auswahl der Import-Dateien"""
        self.import_listbox.delete(0, tk.END)
        self.log_to_import("Auswahl gelöscht")

    def select_all_classes(self):
        """Wählt alle Klassen aus"""
        self.export_listbox.selection_set(0, tk.END)

    def invert_class_selection(self):
        """Kehrt Klassenauswahl um"""
        for i in range(self.export_listbox.size()):
            if i in self.export_listbox.curselection():
                self.export_listbox.selection_clear(i)
            else:
                self.export_listbox.selection_set(i)

    def log_to_import(self, message: str):
        """Fügt Nachricht zum Import-Log hinzu"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.update_log_widget(self.import_log, f"[{timestamp}] {message}")

    def log_to_export(self, message: str):
        """Fügt Nachricht zum Export-Log hinzu"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.update_log_widget(self.export_log, f"[{timestamp}] {message}")

    def clear_export_log(self):
        """Leert das Export-Log"""
        self.export_log.config(state=tk.NORMAL)
        self.export_log.delete(1.0, tk.END)
        self.export_log.config(state=tk.DISABLED)

    def update_log_widget(self, widget, message: str):
        """Thread-sichere Log-Widget-Aktualisierung"""

        def update():
            widget.config(state=tk.NORMAL)
            widget.insert(tk.END, message + "\n")
            widget.see(tk.END)
            widget.config(state=tk.DISABLED)

        if threading.current_thread() == threading.main_thread():
            update()
        else:
            self.root.after(10, update)

    def load_initial_data(self):
        """Lädt initiale Daten"""
        if self.db_path.exists():
            self.refresh_all_data()
        else:
            self.log_to_import(
                "Keine Datenbank gefunden. Bitte importieren Sie zuerst Excel-Dateien."
            )

    def refresh_all_data(self):
        """Aktualisiert alle Daten"""
        try:
            self.load_classes_for_export()
            self.load_classes_for_analysis()
            self.refresh_analysis_data()
            self.refresh_template_list()
            self.status_manager.set_status("Daten aktualisiert")
        except Exception as e:
            logging.error(f"Fehler beim Aktualisieren: {e}")
            self.status_manager.set_status(f"Fehler: {e}")

    def load_classes_for_export(self):
        """Lädt Klassen für Export"""
        try:
            if not self.db_path.exists():
                return

            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.execute(
                    "SELECT DISTINCT klasse FROM schueler ORDER BY klasse"
                )
                classes = [row[0] for row in cursor.fetchall()]

            self.export_listbox.delete(0, tk.END)
            for class_name in classes:
                self.export_listbox.insert(tk.END, class_name)

            self.log_to_export(f"{len(classes)} Klassen gefunden")

        except Exception as e:
            logging.error(f"Fehler beim Laden der Klassen: {e}")

    def load_classes_for_analysis(self):
        """Lädt Klassen für Analyse"""
        try:
            if not self.db_path.exists():
                return

            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.execute(
                    "SELECT DISTINCT klasse FROM schueler ORDER BY klasse"
                )
                classes = ["Alle"] + [row[0] for row in cursor.fetchall()]

            self.class_filter["values"] = classes
            if classes:
                self.class_filter.set(classes[0])

        except Exception as e:
            logging.error(f"Fehler beim Laden der Klassen für Analyse: {e}")

    def refresh_analysis_data(self):
        """Aktualisiert Analyse-Daten"""
        try:
            if not self.db_path.exists():
                return

            self.analysis_tree.delete(*self.analysis_tree.get_children())

            columns = [
                "ID",
                "Name",
                "Klasse",
                "Fächer",
                "AV-Noten",
                "SV-Noten",
                "Vollständig",
            ]
            self.analysis_tree["columns"] = columns
            self.analysis_tree.column("#0", width=0, stretch=False)

            for col in columns:
                self.analysis_tree.heading(col, text=col)
                self.analysis_tree.column(col, width=100)

            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.execute(
                    """
                    SELECT 
                        s.schueler_id,
                        s.name,
                        s.klasse,
                        COUNT(n.noten_id) as faecher_count,
                        SUM(CASE WHEN n.note_av IS NOT NULL THEN 1 ELSE 0 END) as av_count,
                        SUM(CASE WHEN n.note_sv IS NOT NULL THEN 1 ELSE 0 END) as sv_count,
                        CASE WHEN COUNT(n.noten_id) = SUM(CASE WHEN n.note_av IS NOT NULL AND n.note_sv IS NOT NULL THEN 1 ELSE 0 END) 
                             THEN 'Ja' ELSE 'Nein' END as vollstaendig
                    FROM schueler s
                    LEFT JOIN noten n ON s.schueler_id = n.schueler_id
                    GROUP BY s.schueler_id, s.name, s.klasse
                    ORDER BY s.klasse, s.name
                """
                )

                for row in cursor.fetchall():
                    self.analysis_tree.insert("", tk.END, values=row)

        except Exception as e:
            logging.error(f"Fehler beim Aktualisieren der Analyse-Daten: {e}")

    def search_students(self):
        """Sucht Schüler"""
        try:
            class_filter = self.class_filter.get()
            student_filter = self.student_search.get().strip()

            self.analysis_tree.delete(*self.analysis_tree.get_children())

            query = """
                SELECT 
                    s.schueler_id,
                    s.name,
                    s.klasse,
                    COUNT(n.noten_id) as faecher_count,
                    SUM(CASE WHEN n.note_av IS NOT NULL THEN 1 ELSE 0 END) as av_count,
                    SUM(CASE WHEN n.note_sv IS NOT NULL THEN 1 ELSE 0 END) as sv_count,
                    CASE WHEN COUNT(n.noten_id) = SUM(CASE WHEN n.note_av IS NOT NULL AND n.note_sv IS NOT NULL THEN 1 ELSE 0 END) 
                         THEN 'Ja' ELSE 'Nein' END as vollstaendig
                FROM schueler s
                LEFT JOIN noten n ON s.schueler_id = n.schueler_id
                WHERE 1=1
            """
            params = []

            if class_filter and class_filter != "Alle":
                query += " AND s.klasse = ?"
                params.append(class_filter)

            if student_filter:
                query += " AND s.name LIKE ?"
                params.append(f"%{student_filter}%")

            query += (
                " GROUP BY s.schueler_id, s.name, s.klasse ORDER BY s.klasse, s.name"
            )

            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.execute(query, params)
                results = cursor.fetchall()

                for row in results:
                    self.analysis_tree.insert("", tk.END, values=row)

                self.status_manager.set_status(f"{len(results)} Schüler gefunden")

        except Exception as e:
            logging.error(f"Fehler bei der Suche: {e}")
            messagebox.showerror("Suchfehler", f"Fehler: {e}")

    def reset_filters(self):
        """Setzt Filter zurück"""
        self.class_filter.set("Alle")
        self.student_search.delete(0, tk.END)
        self.refresh_analysis_data()

    def delete_selected_student(self):
        """Löscht ausgewählten Schüler"""
        selection = self.analysis_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Keine Auswahl", "Bitte wählen Sie einen Schüler aus."
            )
            return

        item = self.analysis_tree.item(selection[0])
        values = item["values"]
        student_id = values[0]
        student_name = values[1]

        if messagebox.askyesno(
            "Löschen bestätigen", f"Schüler '{student_name}' wirklich löschen?"
        ):
            try:
                with sqlite3.connect(self.db_path) as conn:
                    conn.execute(
                        "DELETE FROM noten WHERE schueler_id = ?", (student_id,)
                    )
                    conn.execute(
                        "DELETE FROM schueler WHERE schueler_id = ?", (student_id,)
                    )
                    conn.commit()

                messagebox.showinfo(
                    "Gelöscht", f"Schüler '{student_name}' wurde gelöscht."
                )
                self.refresh_analysis_data()
                self.load_classes_for_export()

            except Exception as e:
                logging.error(f"Fehler beim Löschen: {e}")
                messagebox.showerror("Lösch-Fehler", f"Fehler: {e}")

    # ===================== HILFS-FUNKTIONEN =====================

    def open_database(self):
        """Öffnet Datenbank"""
        filename = filedialog.askopenfilename(
            title="Datenbank öffnen",
            filetypes=[("SQLite-Datenbank", "*.db"), ("Alle Dateien", "*.*")],
            initialdir=str(Path("output_database").resolve()),
        )

        if filename:
            self.db_path = Path(filename)
            self.refresh_all_data()
            messagebox.showinfo(
                "Datenbank geöffnet", f"Datenbank geöffnet: {self.db_path.name}"
            )

    def show_database_info(self):
        """Zeigt Datenbank-Informationen"""
        if not self.db_path.exists():
            messagebox.showwarning("Keine Datenbank", "Keine Datenbank gefunden.")
            return

        try:
            with sqlite3.connect(self.db_path) as conn:
                schueler_count = conn.execute(
                    "SELECT COUNT(*) FROM schueler"
                ).fetchone()[0]
                klassen_count = conn.execute(
                    "SELECT COUNT(DISTINCT klasse) FROM schueler"
                ).fetchone()[0]
                faecher_count = conn.execute("SELECT COUNT(*) FROM faecher").fetchone()[
                    0
                ]
                noten_count = conn.execute("SELECT COUNT(*) FROM noten").fetchone()[0]

                db_size = self.db_path.stat().st_size / (1024 * 1024)

                info_text = f"""Datenbank-Informationen:

Datei: {self.db_path.name}
Größe: {db_size:.2f} MB

Inhalt:
• Schüler: {schueler_count}
• Klassen: {klassen_count}
• Fächer: {faecher_count}
• Noten: {noten_count}"""

                messagebox.showinfo("Datenbank-Info", info_text)

        except Exception as e:
            messagebox.showerror("Datenbankfehler", f"Fehler: {e}")

    def show_logs(self):
        """Zeigt Logs"""
        log_dir = Path("logs")
        if not log_dir.exists():
            messagebox.showinfo("Keine Logs", "Keine Log-Dateien gefunden.")
            return

        log_files = list(log_dir.glob("*.log"))
        if not log_files:
            messagebox.showinfo("Keine Logs", "Keine Log-Dateien gefunden.")
            return

        latest_log = max(log_files, key=lambda f: f.stat().st_mtime)

        try:
            with open(latest_log, "r", encoding="utf-8") as f:
                content = f.read()

            log_window = tk.Toplevel(self.root)
            log_window.title(f"Logs - {latest_log.name}")
            log_window.geometry("800x600")

            log_text = scrolledtext.ScrolledText(
                log_window, wrap=tk.WORD, font=("Courier", 9)
            )
            log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            log_text.insert(tk.END, content)
            log_text.config(state=tk.DISABLED)
            log_text.see(tk.END)

            ttk.Button(log_window, text="Schließen", command=log_window.destroy).pack(
                pady=5
            )

        except Exception as e:
            messagebox.showerror("Log-Fehler", f"Fehler: {e}")

    def check_permissions(self):
        """Prüft Berechtigungen"""
        try:
            dirs = [
                Path("templates"),
                Path("output_word"),
                Path("output_database"),
                Path("logs"),
            ]
            report = []

            for directory in dirs:
                if directory.exists():
                    readable = os.access(directory, os.R_OK)
                    writable = os.access(directory, os.W_OK)
                    executable = os.access(directory, os.X_OK)
                    status = "✅" if (readable and writable and executable) else "❌"
                    report.append(
                        f"{status} {directory.name}: R:{readable} W:{writable} X:{executable}"
                    )
                else:
                    report.append(f"❓ {directory.name}: Existiert nicht")

            messagebox.showinfo(
                "Berechtigungen", "Berechtigungsprüfung:\n\n" + "\n".join(report)
            )

        except Exception as e:
            messagebox.showerror("Prüfungsfehler", f"Fehler: {e}")

    def show_about(self):
        """Zeigt Über-Dialog"""
        about_text = """AES-Kopfnoten-Manager v2.1 - Optimiert
Entwickelt für IGS in Hessen
© 2025"""

        messagebox.showinfo("Über", about_text)

    def show_linux_help(self):
        """Zeigt Linux-Hilfe"""
        help_text = """Linux-Hilfe - Optimierte Version

SCHNELLSTART:
1. chmod 755 templates/ output_word/ output_database/ logs/
2. python kopfnotenapp.py

FEATURES:
• Template-Designer: Einfache Template-Erstellung
• Optimierter Export: Horizontale 3-Zeilen-Tabellen mit Beschriftungsspalte
• Einzelschüler-Export: Export für einzelne ausgewählte Schüler
• Vereinfachte Noten-Bearbeitung

PROBLEMLÖSUNG:
• Bei Berechtigungsfehlern: sudo chown $USER:$USER -R ./
• Templates in LibreOffice Writer bearbeiten
• Logs prüfen: tail -f logs/kopfnoten_gui.log"""

        messagebox.showinfo("Linux-Hilfe", help_text)

    def on_closing(self):
        """Behandelt Schließen"""
        if self.export_running:
            messagebox.showwarning(
                "Export läuft", "Bitte warten Sie bis der Export abgeschlossen ist."
            )
            return

        if messagebox.askokcancel("Beenden", "Anwendung beenden?"):
            self.root.destroy()


def main():
    """Hauptfunktion"""
    try:
        # Verzeichnisse erstellen
        Path("logs").mkdir(exist_ok=True)
        Path("templates").mkdir(exist_ok=True)
        Path("output_word").mkdir(exist_ok=True)
        Path("output_database").mkdir(exist_ok=True)

        # Anwendung starten
        app = KopfnotenGUI()
        app.root.mainloop()

    except Exception as e:
        logging.error(f"Kritischer Anwendungsfehler: {e}")
        print(f"Fehler beim Starten: {e}")


if __name__ == "__main__":
    main()