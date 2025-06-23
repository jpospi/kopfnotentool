1. Ergänze die GUI um ein Eingabefeld für das Export-Datum
Füge im Export-Tab (in create_export_tab) ein neues Eingabefeld für das Datum hinzu, z. B.:

Python
# in create_export_tab (nach self.output_var etc.)
self.export_date_var = tk.StringVar(value=datetime.now().strftime("%d.%m.%Y"))

# Im GUI-Code für Export-Optionen:
date_frame = ttk.Frame(options_frame)
date_frame.pack(fill=tk.X, padx=5, pady=5)
ttk.Label(date_frame, text="Export-Datum:").pack(side=tk.LEFT)
date_entry = ttk.Entry(date_frame, textvariable=self.export_date_var, width=15)
date_entry.pack(side=tk.LEFT, padx=(5, 5))
ttk.Button(date_frame, text="Heute", command=lambda: self.export_date_var.set(datetime.now().strftime("%d.%m.%Y"))).pack(side=tk.LEFT)
2. Übergebe das Datum an die Exportfunktion
Passe die Methoden start_optimized_export, run_optimized_export, export_selected_student und run_student_export so an, dass sie das Datum übergeben:

Beispiel für start_optimized_export:
Python
def start_optimized_export(self):
    ...
    export_date = self.export_date_var.get().strip()
    ...
    export_thread = threading.Thread(
        target=self.run_optimized_export,
        args=(selected_classes, template_path, output_dir, export_date),
        daemon=True,
    )
    export_thread.start()
und entsprechend:

Python
def run_optimized_export(self, klassen, template_path, output_dir, export_date):
    ...
    summary = exporter.export_horizontal_tables(
        output_dir, template_path, klassen, export_date=export_date
    )
Dasselbe für Einzelschüler-Export:
Python
def export_selected_student(self):
    ...
    export_date = self.export_date_var.get().strip()
    ...
    export_thread = threading.Thread(
        target=self.run_student_export,
        args=(student_id, student_name, student_class, template_path, output_dir, export_date),
        daemon=True,
    )
    export_thread.start()
und:

Python
def run_student_export(self, student_id, student_name, student_class, template_path, output_dir, export_date):
    ...
    summary = exporter.export_horizontal_tables(
        output_dir, template_path, [student_class], student_id, export_date=export_date
    )
3. Passe die Export-Logik an (OptimizedKopfnotenExporter):
Füge in der Methode export_horizontal_tables das neue Argument hinzu und übergib es an die Kontextdaten:

Python
def export_horizontal_tables(
    self, output_dir: Path, template_path: Path, klassen_liste: List[str], schueler_id: Optional[int] = None, export_date: Optional[str] = None
) -> Dict[str, Any]:
    ...
    if not export_date:
        export_date = datetime.now().strftime("%d.%m.%Y")
    ...
    context = {
        "klasse": klasse,
        "export_datum": export_date,
        ...
    }
    ...