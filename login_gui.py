import tkinter as tk
from tkinter import ttk, messagebox
import logging
import queue
from typing import Optional, Dict, Any, Callable

class LoginWindow:
    """
    Modal Login Window.
    Blocks the main application startup until successful login.
    """
    def __init__(self, root, credential_manager):
        self.root = root
        self.cm = credential_manager
        self.logger = logging.getLogger("login_gui")
        self.all_schools = None
        self.search_after = None # For debouncing
        self.queue = queue.Queue()
        
        self.result = False # Login Success Status
        
        self.window = tk.Toplevel(root)
        self.window.title("Anmeldung - Schulportal Hessen")
        self.window.geometry("600x600")
        self.window.resizable(True, True)
        
        # Modal interactions
        # self.window.transient(root) # Disabled: Causes issues on Linux if root is withdrawn
        self.window.deiconify()
        self.window.attributes("-topmost", True)
        self.window.lift()
        self.window.focus_force()
        # self.window.attributes("-topmost", False) # Disable topmost after showing? 
        # Better leave it for a moment or remove if annoying.
        
        self.window.grab_set()
        
        self.center_window()
        self.create_widgets()
        
        # Pre-fill
        self.load_previous_school()
        
        # Handle X button
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Start queue processing
        self.process_queue()

    def center_window(self):
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'+{x}+{y}')

    def create_widgets(self):
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        ttk.Label(main_frame, text="Anmeldung", font=("Segoe UI", 16, "bold")).pack(pady=(0, 20))
        
        # Info Text
        info_text = (
            "Bitte melden Sie sich mit Ihren Lanis-Zugangsdaten an.\n"
            "Diese werden lokal verschlüsselt gespeichert.\n"
            "Für den ersten Login ist eine Internetverbindung erforderlich."
        )
        ttk.Label(main_frame, text=info_text, justify=tk.CENTER, foreground="#555").pack(pady=(0, 20))
        
        # Form Container
        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=10) # Expand for listbox
        
        # 1. School Search
        ttk.Label(form_frame, text="Schule (Name/Ort):", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(5, 0))
        
        search_frame = ttk.Frame(form_frame)
        search_frame.pack(fill=tk.X, pady=(2, 10))
        
        self.school_search_var = tk.StringVar()
        self.school_search_entry = ttk.Entry(search_frame, textvariable=self.school_search_var)
        self.school_search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.school_search_entry.bind("<Return>", lambda e: self.start_search_thread())
        self.school_search_entry.bind("<KeyRelease>", lambda e: self.start_search_thread())
        
        ttk.Button(search_frame, text="🔍", width=3, command=self.start_search_thread).pack(side=tk.RIGHT)
        
        # School List
        list_frame = ttk.Frame(form_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.school_listbox = tk.Listbox(list_frame, height=5)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.school_listbox.yview)
        self.school_listbox.config(yscrollcommand=scrollbar.set)
        
        self.school_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.school_listbox.bind("<<ListboxSelect>>", self.on_school_select)
        
        # Hidden School ID
        self.selected_school_id = None
        self.selected_school_label = ttk.Label(form_frame, text="Keine Schule ausgewählt", foreground="red")
        self.selected_school_label.pack(anchor=tk.W, pady=(0, 10))
        
        # 2. User
        ttk.Label(form_frame, text="Benutzername (Kürzel):", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        self.user_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.user_var).pack(fill=tk.X, pady=(2, 10))
        
        # 3. Password
        ttk.Label(form_frame, text="Passwort:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        self.pw_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.pw_var, show="*").pack(fill=tk.X, pady=(2, 20))
        
        # Login Button
        self.login_btn = ttk.Button(main_frame, text="Anmelden", command=self.perform_login_thread)
        self.login_btn.pack(fill=tk.X, ipady=5)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="", foreground="blue")
        self.status_label.pack(pady=10)

    def process_queue(self):
        """Processes UI updates from background threads"""
        try:
            while True:
                task = self.queue.get_nowait()
                if callable(task):
                    task()
                self.queue.task_done()
        except queue.Empty:
            pass
        self.window.after(100, self.process_queue)

    def queue_ui(self, func: Callable, *args, **kwargs):
        """Queues a UI update function to be run on the main thread"""
        self.queue.put(lambda: func(*args, **kwargs))

    def start_search_thread(self):
        """Starts search with debounce"""
        import threading
        if self.search_after:
            self.window.after_cancel(self.search_after)
            
        term = self.school_search_var.get()
        if len(term) < 2:
            self.school_listbox.delete(0, tk.END)
            self.status_label.config(text="")
            return

        def run():
            self.queue_ui(self.status_label.config, text="Suche läuft...", foreground="blue")
            self.search_schools(term)

        self.search_after = self.window.after(300, lambda: threading.Thread(target=run, daemon=True).start())

    def search_schools(self, term):
        if len(term) < 2:
            return
            
        self.queue_ui(self.school_listbox.delete, 0, tk.END)
        self.queue_ui(self.school_listbox.insert, tk.END, "Suche läuft...")
            
        try:
            if not self.all_schools:
                from sph_downloader import SPHDownloader
                dl = SPHDownloader()
                # Sort schools by name for better UX
                self.all_schools = sorted(dl.get_schools(), key=lambda x: x["name"])
            
            schools = self.all_schools
            
            # Post-process in thread, then update UI
            found = []
            term_lower = term.lower()
            
            # 1. Check for exact ID match first (Schulnummer)
            for s in schools:
                if term_lower == str(s["id"]):
                     label = f"{s['name']} - {s['city']} [{s['id']}]"
                     found.append((label, s))
            
            # 2. Then check name/city/contains ID
            for s in schools:
                # Avoid duplicates from exact ID match
                if any(x[1]["id"] == s["id"] for x in found):
                    continue
                    
                if term_lower in s["name"].lower() or term_lower in s["city"].lower() or term_lower in str(s["id"]):
                     label = f"{s['name']} - {s['city']} [{s['id']}]"
                     found.append((label, s))
            
            # Update UI via queue
            self.queue_ui(self.update_school_list, found)
                
        except Exception as e:
            self.queue_ui(self.status_label.config, text=f"Suchfehler: {e}", foreground="red")

    def update_school_list(self, found_schools_tuples):
        self.school_listbox.delete(0, tk.END)
        self.found_schools = []
        
        if not found_schools_tuples:
            self.school_listbox.insert(tk.END, "Keine Treffer.")
        else:
            for label, s in found_schools_tuples:
                self.school_listbox.insert(tk.END, label)
                self.found_schools.append(s)
        
        self.status_label.config(text=f"{len(found_schools_tuples)} Schulen gefunden.")

    def on_school_select(self, event):
        selection = self.school_listbox.curselection()
        if selection and hasattr(self, "found_schools") and selection[0] < len(self.found_schools):
            s = self.found_schools[selection[0]]
            self.selected_school_id = s["id"]
            self.selected_school_label.config(text=f"Ausgewählt: {s['name']} ({s['id']})", foreground="green")

    def load_previous_school(self):
        """Tries to pre-fill from recent config or secrets"""
        # We can check sph_config.json for convenience
        try:
            import json
            from pathlib import Path
            p = Path("sph_config.json")
            if p.exists():
                with open(p, "r") as f:
                    data = json.load(f)
                    if "school" in data:
                        self.selected_school_id = data["school"]
                        self.selected_school_label.config(text=f"Gespeichert: ID {self.selected_school_id}", foreground="blue")
                    if "user" in data:
                        self.user_var.set(data["user"])
        except:
            pass

    def perform_login_thread(self):
        """Starts login in thread"""
        import threading
        
        school = self.selected_school_id
        user = self.user_var.get().strip()
        pw = self.pw_var.get().strip() # Still on main thread
        
        self.status_label.config(text="Anmeldung läuft...", foreground="blue")
        self.login_btn.config(state=tk.DISABLED)
        
        t = threading.Thread(target=self.perform_login_logic, args=(school, user, pw), daemon=True)
        t.start()

    def perform_login_logic(self, school, user, pw):
        if not school:
            def warn():
                messagebox.showwarning("Fehler", "Bitte eine Schule auswählen.")
                self.login_btn.config(state=tk.NORMAL)
                self.status_label.config(text="Bitte Schule wählen.")
            self.queue_ui(warn)
            return
            
        success, msg = self.cm.login(school, user, pw)
        
        self.queue_ui(self._handle_login_result, success, msg)

    def _handle_login_result(self, success, msg):
        self.login_btn.config(state=tk.NORMAL)
        if success:
            self.status_label.config(text=msg, foreground="green")
            self.result = True
            # Short delay to show success message
            self.window.after(500, self.window.destroy)
        else:
            self.status_label.config(text=msg, foreground="red")
            messagebox.showerror("Fehler", msg)

    def on_close(self):
        if not self.result:
            self.window.destroy()

