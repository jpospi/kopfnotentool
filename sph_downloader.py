import requests
import logging
import os
import json
import time
from pathlib import Path
from lanisapi import LanisClient, LanisAccount, LanisCookie

class SPHDownloader:
    BASE_URL = "https://start.schulportal.hessen.de"
    
    def __init__(self, output_dir="temp", logger=None):
        self.logger = logger or logging.getLogger("sph_downloader")
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True, parents=True)
        self.client = None
        
    def get_schools(self):
        """Lädt die Schulliste via LanisAPI (oder Cache)"""
        cache_file = self.output_dir / "schools.json"
        
        # Cache prüfen (1 Tag gültig)
        if cache_file.exists():
            if time.time() - cache_file.stat().st_mtime < 86400:
                try:
                    with open(cache_file, "r", encoding="utf-8") as f:
                        return json.load(f)
                except:
                    pass

        try:
            self.logger.info("Lade Schulliste via LanisAPI...")
            # LanisAPI get_schools is static-like or works with dummy client
            # We use a dummy account to init client if needed, or just standard init
            dummy_client = LanisClient(LanisAccount("0", "", ""))
            lanis_schools = dummy_client.get_schools()
            
            schools = []
            for s in lanis_schools:
                # LanisAPI returns dicts with Id, Name, Ort
                sid = str(s.get("Id", ""))
                name = s.get("Name", "")
                city = s.get("Ort", "")
                
                schools.append({
                    "id": sid,
                    "name": name,
                    "city": city,
                    "label": f"{name} ({city}) [{sid}]"
                })
            
            # Cache speichern
            with open(cache_file, "w", encoding="utf-8") as f:
                json.dump(schools, f, ensure_ascii=False)
                
            return schools
            
        except Exception as e:
            self.logger.error(f"Fehler beim Laden der Schulliste: {e}")
            return []

    def login(self, school_id, username, password):
        """Führt Login via LanisAPI durch"""
        self.logger.info(f"Versuche Login bei Schule {school_id} als {username} via LanisAPI")
        
        try:
            account = LanisAccount(school_id, username, password)
            self.client = LanisClient(account)
            self.client.authenticate()
            
            self.logger.info("Login erfolgreich.")
            return True
        except Exception as e:
            import traceback
            self.logger.error(f"Login fehlgeschlagen: {e}\n{traceback.format_exc()}")
            raise ConnectionError(f"Login fehlgeschlagen: {e}")

    def download_class_list(self, class_name, year_level, output_dir):
        """Lädt Liste für eine Klasse herunter"""
        if not self.client:
             raise ConnectionError("Nicht eingeloggt.")

        # URL Format:
        url = f"{self.BASE_URL}/meinunterricht.php"
        params = {
            "a": "klassenlehrerAVSV",
            "k": class_name,
            "b": "xlsx"
        }
        
        try:
            # Transfer Cookies from LanisClient (httpx) to requests
            import requests
            session = requests.Session()
            
            # Use the official property to get cookies (school_id and session_id)
            if self.client:
                cookies = self.client.authentication_cookies
                session.cookies.set("i", cookies.school_id)
                session.cookies.set("sid", cookies.session_id)
                self.logger.info(f"Cookies gesetzt: i={cookies.school_id}, sid={cookies.session_id}")
            
            # Also set User-Agent to match
            session.headers.update({
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
            })
            
            r = session.get(url, params=params)
            r.raise_for_status()
            
            # Content-Check
            if "application/vnd.openxmlformats" not in r.headers.get("Content-Type", "") and len(r.content) < 1000:
                self.logger.warning(f"Warnung: Download für {class_name} scheint kein gültiges Excel zu sein.")
            
            file_path = Path(output_dir) / f"Klasse_{class_name}.xlsx"
            with open(file_path, "wb") as f:
                f.write(r.content)
            
            return file_path
        except Exception as e:
            self.logger.error(f"Fehler beim Download {class_name}: {e}")
            return None
