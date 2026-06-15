import requests
import logging
import os
import json
import time
import re
import zipfile
from pathlib import Path
from lanisapi import LanisClient, LanisAccount, LanisCookie
from app_paths import load_app_paths
from lxml import html

class SPHDownloader:
    BASE_URL = "https://start.schulportal.hessen.de"
    _lanis_sid_patch_applied = False
    _lanis_cryptor_patch_applied = False

    @staticmethod
    def _looks_like_xlsx_bytes(content: bytes) -> bool:
        """Checks if bytes look like a valid XLSX (ZIP-based OOXML) file."""
        if not content or len(content) < 4:
            return False
        # XLSX is a ZIP container and starts with PK\x03\x04.
        if not content.startswith(b"PK"):
            return False
        # Extra safety: verify ZIP structure.
        try:
            import io
            with zipfile.ZipFile(io.BytesIO(content), "r") as zf:
                return "[Content_Types].xml" in zf.namelist()
        except Exception:
            return False
    
    def __init__(self, output_dir=None, logger=None):
        self.logger = logger or logging.getLogger("sph_downloader")
        if output_dir is None:
            output_dir = str(load_app_paths().temp_dir)
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
            self._apply_lanisapi_sid_patch()

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

    def _apply_lanisapi_sid_patch(self):
        """
        Patcht einen bekannten lanisapi-Bug (v0.4.1) beim Parsen von Set-Cookie.
        Der Upstream-Code zerlegt den Header mit festen Indizes und crasht mit
        'IndexError: list index out of range', obwohl die Verbindung funktioniert.
        """
        if SPHDownloader._lanis_sid_patch_applied:
            return

        try:
            import httpx
            from http.cookies import SimpleCookie
            from lanisapi.helpers.request import Request as LanisRequest
            import lanisapi.helpers.authentication as lanis_auth
            import lanisapi.client as lanis_client

            def _robust_get_authentication_sid(url: str, cookies: httpx.Cookies, schoolid: str) -> httpx.Cookies:
                response = LanisRequest.head(url, cookies=cookies)

                final_cookies = httpx.Cookies()
                final_cookies.set("i", schoolid)

                sid = None

                # 1) Preferred: parse cookie jar directly
                try:
                    sid = response.cookies.get("sid")
                except Exception:
                    sid = None

                # 2) Fallback: parse all Set-Cookie headers robustly
                if not sid:
                    try:
                        set_cookie_headers = response.headers.get_list("set-cookie")
                    except Exception:
                        raw_header = response.headers.get("set-cookie")
                        set_cookie_headers = [raw_header] if raw_header else []

                    for header in set_cookie_headers:
                        if not header:
                            continue
                        cookie = SimpleCookie()
                        try:
                            cookie.load(header)
                        except Exception:
                            pass
                        if "sid" in cookie:
                            sid = cookie["sid"].value
                            break

                        # Last-resort regex for unusual combined header formats
                        match = re.search(r"(?:^|[;,]\s*)sid=([^;,\s]+)", header)
                        if match:
                            sid = match.group(1)
                            break

                if not sid:
                    raise ConnectionError(
                        "SPH-Session-ID konnte nicht aus der Serverantwort gelesen werden."
                    )

                final_cookies.set("sid", sid)
                return final_cookies

            # Patch both helper module and already imported symbol in client module.
            lanis_auth.get_authentication_sid = _robust_get_authentication_sid
            lanis_client.get_authentication_sid = _robust_get_authentication_sid

            SPHDownloader._lanis_sid_patch_applied = True
            self.logger.info("LanisAPI SID-Workaround aktiviert.")
        except Exception as e:
            # Non-fatal: keep old behavior if patching fails.
            self.logger.warning(f"LanisAPI SID-Workaround konnte nicht aktiviert werden: {e}")

        # Also patch sporadic RSA plaintext length issue in lanisapi handshake.
        self._apply_lanisapi_cryptor_patch()

    def _apply_lanisapi_cryptor_patch(self):
        """
        Patcht einen sporadischen lanisapi-Fehler:
        ValueError("Plaintext is too long.") beim RSA-Handshake.
        Dann wird ein kompakteres Secret erzeugt und erneut versucht.
        """
        if SPHDownloader._lanis_cryptor_patch_applied:
            return

        try:
            import base64
            import secrets
            from Cryptodome.Cipher import PKCS1_v1_5
            from Cryptodome.PublicKey import RSA
            import lanisapi.helpers.cryptor as lanis_cryptor

            original_encrypt_key = lanis_cryptor.Cryptor._encrypt_key

            def _patched_encrypt_key(self, public_key: str) -> str:
                try:
                    return original_encrypt_key(self, public_key)
                except ValueError as e:
                    if "Plaintext is too long" not in str(e):
                        raise

                    rsa = PKCS1_v1_5.new(RSA.import_key(public_key))

                    # Try progressively smaller secrets until encryption fits.
                    for plain_len in (16, 12, 8, 6, 4):
                        compact_plain = secrets.token_hex((plain_len + 1) // 2)[:plain_len]
                        compact_secret = self.encrypt(compact_plain, compact_plain)
                        try:
                            encrypted = base64.b64encode(rsa.encrypt(compact_secret.encode())).decode()
                            self.secret = compact_secret
                            return encrypted
                        except ValueError:
                            continue

                    raise ValueError(
                        "Plaintext is too long (auch mit kompaktem Secret)."
                    )

            lanis_cryptor.Cryptor._encrypt_key = _patched_encrypt_key
            SPHDownloader._lanis_cryptor_patch_applied = True
            self.logger.info("LanisAPI Cryptor-Workaround aktiviert.")
        except Exception as e:
            self.logger.warning(f"LanisAPI Cryptor-Workaround konnte nicht aktiviert werden: {e}")

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

            content_type = (r.headers.get("Content-Type", "") or "").lower()
            looks_like_xlsx = self._looks_like_xlsx_bytes(r.content)
            if not looks_like_xlsx:
                preview = (r.text[:300] if r.text else "").replace("\n", " ").replace("\r", " ")
                self.logger.warning(
                    f"Download für {class_name} ist kein gültiges XLSX "
                    f"(content-type='{content_type}', bytes={len(r.content)}). Vorschau: {preview}"
                )
                return None

            file_path = Path(output_dir) / f"Klasse_{class_name}.xlsx"
            with open(file_path, "wb") as f:
                f.write(r.content)

            # Final on-disk validation before returning.
            if not zipfile.is_zipfile(file_path):
                self.logger.warning(f"Download für {class_name} wurde verworfen: Datei ist kein ZIP/XLSX.")
                try:
                    file_path.unlink(missing_ok=True)
                except Exception:
                    pass
                return None
            
            return file_path
        except Exception as e:
            self.logger.error(f"Fehler beim Download {class_name}: {e}")
            return None

    def fetch_missing_submissions_overview(self):
        """
        Lädt die SPH-Seite 'Fehlende Abgaben' für alle Zweigstufen und liefert
        einen Klassen-Überblick im Ampelsystem:
          - gruen  => alles erfolgt
          - gelb   => tlw. fehlend
          - rot    => fehlend
        Es werden KEINE Fachnamen/WPU-Bezeichnungen verändert.
        """
        if not self.client:
            raise ConnectionError("Nicht eingeloggt.")

        from lanisapi.helpers.request import Request as LanisRequest

        cookies = LanisRequest.get_cookies()
        sid = cookies.get("sid", domain="")
        i = cookies.get("i", domain="") or self.school_id
        if not sid:
            raise ConnectionError("Keine aktive Session (sid fehlt).")

        session = requests.Session()
        session.cookies.set("i", i)
        session.cookies.set("sid", sid)
        session.headers.update(
            {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
            }
        )

        url = f"{self.BASE_URL}/kopfnoten.php?a=fehlende"
        branches = ["IGS~5", "IGS~6", "IGS~7", "IGS~8", "IGS~9", "NDHS/S1~30"]
        class_overview = {}

        for branch in branches:
            response = session.post(
                url, data={"a": "fehlende", "zweigstufe": branch}, timeout=30
            )
            response.raise_for_status()

            doc = html.fromstring(response.text)
            rows = doc.xpath("//table[@id='kopfnotenTable']/tbody/tr")

            for row in rows:
                lerngruppe = " ".join(row.xpath("string(td[1])").split())
                status_text = " ".join(row.xpath("string(td[3])").split()).lower()

                class_matches = list(re.finditer(r"\b(\d{2}[a-z]|daz\d+)\b", lerngruppe, re.IGNORECASE))
                if not class_matches:
                    continue

                first_match = class_matches[0]
                subject_text = lerngruppe[: first_match.start()].strip()
                if status_text == "erfolgt":
                    row_color = "gruen"
                elif "tlw" in status_text:
                    row_color = "gelb"
                elif "fehlend" in status_text:
                    row_color = "rot"
                else:
                    row_color = "unbekannt"

                # Eine Lerngruppe kann mehrere Klassen enthalten (z. B. 07c/07d).
                for match in class_matches:
                    klasse = match.group(1).upper()
                    class_token = match.group(1)
                    current = class_overview.get(
                        klasse,
                        {"rows": [], "has_red": False, "has_yellow": False, "all_green": True},
                    )
                    current["rows"].append(
                        {
                            "lerngruppe": lerngruppe,
                            "klasse": klasse,
                            "klasse_token": class_token,
                            "fach_raw": subject_text,
                            "status": status_text,
                            "farbe": row_color,
                        }
                    )

                    if row_color == "rot":
                        current["has_red"] = True
                    if row_color == "gelb":
                        current["has_yellow"] = True
                    if row_color != "gruen":
                        current["all_green"] = False

                    class_overview[klasse] = current

        return class_overview
