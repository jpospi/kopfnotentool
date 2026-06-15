import os
import json
import base64
import logging
from pathlib import Path
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from typing import Optional, Tuple, Dict

from sph_downloader import SPHDownloader
from app_paths import load_app_paths

logger = logging.getLogger("credentials")

class CredentialManager:
    """
    Manages secure storage and retrieval of SPH credentials.
    Uses AES-GCM encryption where the key is derived from the user's password.
    This allows offline verification: if the password can decrypt the file, it's correct.
    """
    
    SECRET_FILE = "secret.dat"
    ITERATIONS = 480000 # High iteration count for security
    
    def __init__(self, data_dir: Optional[str] = None):
        if data_dir is None:
            data_dir = str(load_app_paths().temp_dir)
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(exist_ok=True, parents=True)
        self.secret_path = self.data_dir / self.SECRET_FILE
        self.credentials = None # Cached (school_id, username, password)
        self.session_cookies = None
        
    def _derive_key(self, password: str, salt: bytes) -> bytes:
        """Derive a 256-bit key from the password using PBKDF2."""
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=self.ITERATIONS,
        )
        return kdf.derive(password.encode("utf-8"))

    def login(self, school_id: str, username: str, password: str) -> Tuple[bool, str]:
        """
        Attempts to login. 
        If online: authenticates against SPH, and if successful, saves encrypted credentials.
        If offline (connection error): attempts to decrypt local storage with provided password.
        
        Returns: (success, message)
        """
        if not school_id or not username or not password:
             return False, "Bitte alle Felder ausfüllen."

        # 1. Try Online Login
        try:
            downloader = SPHDownloader(logger=logger)
            # SPHDownloader.login returns True or raises Exception
            downloader.login(school_id, username, password)
            
            # If successful: Save credentials encrypted
            self._save_credentials(school_id, username, password)
            self.credentials = (school_id, username, password)
            # We could also pull the cookies from the downloader if needed
            if downloader.client:
                self.session_cookies = downloader.client.authentication_cookies
                
            return True, "Login erfolgreich (Online)."
            
        except ConnectionError as e:
            # Always try online first; offline fallback only if local credentials exist.
            logger.info(f"Online login failed: {e}")
            if self.secret_path.exists():
                logger.info("Offline data found, trying offline verification.")
                return self._verify_offline(school_id, username, password)
            return False, (
                "Online-Anmeldung fehlgeschlagen und keine Offline-Daten vorhanden. "
                "Bitte Internetverbindung prüfen und erneut anmelden."
            )
            
        except ValueError as e:
             # Specifically for authentication failures
             logger.error(f"Auth error: {e}")
             return False, str(e)
             
        except Exception as e:
            # Other errors
            logger.error(f"Login error: {e}")
            if self.secret_path.exists():
                logger.info("Trying offline verification after unexpected online error.")
                return self._verify_offline(school_id, username, password)
            return False, f"Login fehlgeschlagen: {e}"

    def _save_credentials(self, school_id: str, username: str, password: str):
        """Encrypts and saves credentials."""
        try:
            salt = os.urandom(16)
            key = self._derive_key(password, salt)
            aesgcm = AESGCM(key)
            nonce = os.urandom(12)
            
            data = json.dumps({
                "school": school_id,
                "user": username,
                "password": password # Storing pw to allow re-login flow if needed
            }).encode("utf-8")
            
            ciphertext = aesgcm.encrypt(nonce, data, None)
            
            # File format: salt (16) + nonce (12) + ciphertext
            with open(self.secret_path, "wb") as f:
                f.write(salt + nonce + ciphertext)
                
            logger.info("Credentials secured locally.")
        except Exception as e:
            logger.error(f"Failed to save credentials: {e}")

    def _verify_offline(self, school_id: str, username: str, password: str) -> Tuple[bool, str]:
        """Verifies credentials by attempting to decrypt the local file."""
        if not self.secret_path.exists():
            return False, "Keine Offline-Daten vorhanden. Bitte einmalig online anmelden."
            
        try:
            with open(self.secret_path, "rb") as f:
                content = f.read()
                
            if len(content) < 28: # 16 salt + 12 nonce
                return False, "Beschädigte Offline-Daten."
                
            salt = content[:16]
            nonce = content[16:28]
            ciphertext = content[28:]
            
            key = self._derive_key(password, salt)
            aesgcm = AESGCM(key)
            
            plaintext = aesgcm.decrypt(nonce, ciphertext, None)
            data = json.loads(plaintext.decode("utf-8"))
            
            # Verify data matches (though decryption success implies password is correct)
            # We also check username/school to ensure it's the right account
            if data["school"] != school_id or data["user"] != username:
                return False, "Gespeicherte Daten gehören zu einem anderen Nutzer."
                
            self.credentials = (data["school"], data["user"], data["password"])
            return True, "Login erfolgreich (Offline)."
            
        except Exception as e:
            logger.warning(f"Offline decryption failed: {e}")
            return False, "Falsches Passwort oder beschädigte Daten."

    def get_saved_info(self) -> Dict[str, str]:
        """
        Attempts to recover School/User info WITHOUT password if we stored it separately?
        Wait, we decided to store everything encrypted.
        So we can't get it unless we ask.
        BUT, the `sph_config.json` ALREADY stores recent school and username in plaintext (from previous implementation).
        We can continue to use that for pre-filling the UI.
        
        This method is strictly for when we have already unlocked the detailed secure data.
        """
        if self.credentials:
            return {"school": self.credentials[0], "user": self.credentials[1]}
        return {}
