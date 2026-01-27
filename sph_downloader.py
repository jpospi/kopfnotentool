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
        self.school_id = None
        
    def get_schools(self):
        """Lädt die Schulliste via SPH Endpoint (oder Cache)"""
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
            self.logger.info("Lade Schulliste vom SPH Exporteur...")
            # Direct fetch to avoid library overhead/bugs in get_schools
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
            }
            url = "https://startcache.schulportal.hessen.de/exporteur.php?a=schoollist"
            r = requests.get(url, headers=headers, timeout=10)
            r.raise_for_status()
            
            data = r.json()
            schools = []
            
            # Data is a list of categories: [{"Kategorie": "...", "Schulen": [...]}, ...]
            for group in data:
                if "Schulen" in group:
                    for s in group["Schulen"]:
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
            # WORKAROUND for lanisapi GC bug:
            # The library has a global httpx client that can be closed by __del__ of previous clients.
            from lanisapi.helpers.request import Request as LanisRequest
            import httpx
            try:
                # Test if client is closed
                LanisRequest.client.post("https://example.com", timeout=0.001)
            except RuntimeError as re:
                if "closed" in str(re).lower():
                    self.logger.info("LanisAPI internal client was closed. Re-opening...")
                    LanisRequest.client = httpx.Client(timeout=httpx.Timeout(30.0, connect=60.0))
            except Exception:
                pass # Other errors are fine, we just want to catch "client closed"

            account = LanisAccount(school_id, username, password)
            self.client = LanisClient(account)
            self.client.authenticate()
            
            if not self.client.authenticated:
                 raise ConnectionError("Authentifizierung abgeschlossen, aber Client ist nicht als 'authenticated' markiert (Handshake Fehler?).")

            self.school_id = school_id
            self.logger.info("Login erfolgreich.")
            return True
        except Exception as e:
            import traceback
            err_msg = str(e)
            self.logger.error(f"Login fehlgeschlagen: {err_msg}\n{traceback.format_exc()}")
            
            # Simple heuristic for auth/connection
            if any(x in err_msg.lower() for x in ["auth", "login", "credentials", "zugangsdaten", "passwort"]):
                 raise ValueError(f"Login fehlgeschlagen: {err_msg}")
            else:
                 raise ConnectionError(f"Verbindungsfehler: {err_msg}")

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
                # We extract cookies manually from the client's internal Request helper
                # to be sure we get the current state.
                from lanisapi.helpers.request import Request as LanisRequest
                cookies = LanisRequest.get_cookies()
                
                # SPH expects 'i' and 'sid'
                # httpx cookies can be accessed via .get(name)
                # We try to get them without being too picky about domains
                sid = cookies.get("sid", domain="")
                i = cookies.get("i", domain="") or self.school_id
                
                if not sid:
                     self.logger.error("Download fehlgeschlagen: Keine Session-ID (sid) gefunden.")
                     raise ConnectionError("Keine aktive Session (sid fehlt).")
                
                session.cookies.set("i", i)
                session.cookies.set("sid", sid)
                self.logger.info(f"Cookies gesetzt: i={i}, sid={sid[:8]}...")
            
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
