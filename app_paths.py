import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Any


APP_NAME = "KopfnotenTool"
CONFIG_FILENAME = "kopfnotentool.paths.json"


@dataclass(frozen=True)
class AppPaths:
    data_root: Path
    import_dir: Path
    output_word_dir: Path
    output_excel_dir: Path
    templates_dir: Path
    logs_dir: Path
    temp_dir: Path
    database_path: Path
    backup_dir: Path
    sph_config_path: Path
    config_file: Path

    def ensure_runtime_dirs(self) -> None:
        dirs = [
            self.data_root,
            self.import_dir,
            self.output_word_dir,
            self.output_excel_dir,
            self.templates_dir,
            self.logs_dir,
            self.temp_dir,
            self.database_path.parent,
            self.backup_dir,
        ]
        for directory in dirs:
            directory.mkdir(parents=True, exist_ok=True)


def _default_data_root() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
        return base / APP_NAME
    return Path.home() / ".kopfnotentool"


def _default_config_file() -> Path:
    # If frozen, prefer config beside the executable (installer writes it there).
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / CONFIG_FILENAME
    # Development: keep config in project root unless overridden.
    return Path.cwd() / CONFIG_FILENAME


def _resolve_path(value: str, fallback: Path) -> Path:
    if not value:
        return fallback
    p = Path(value).expanduser()
    if p.is_absolute():
        return p
    return (fallback.parent / p).resolve()


def load_app_paths() -> AppPaths:
    env_config = os.environ.get("KOPFNOTEN_CONFIG_FILE", "").strip()
    config_file = Path(env_config).expanduser() if env_config else _default_config_file()

    raw: Dict[str, Any] = {}
    if config_file.exists():
        try:
            raw = json.loads(config_file.read_text(encoding="utf-8"))
        except Exception:
            raw = {}

    env_data_root = os.environ.get("KOPFNOTEN_DATA_ROOT", "").strip()
    data_root_default = Path(env_data_root).expanduser() if env_data_root else _default_data_root()
    data_root = _resolve_path(str(raw.get("data_root", "")), data_root_default)

    import_dir = _resolve_path(str(raw.get("import_dir", "")), data_root / "input_excel")
    output_word_dir = _resolve_path(str(raw.get("output_word_dir", "")), data_root / "output_word")
    output_excel_dir = _resolve_path(str(raw.get("output_excel_dir", "")), data_root / "output_excel")
    templates_dir = _resolve_path(str(raw.get("templates_dir", "")), data_root / "templates")
    logs_dir = _resolve_path(str(raw.get("logs_dir", "")), data_root / "logs")
    temp_dir = _resolve_path(str(raw.get("temp_dir", "")), data_root / "temp")
    database_path = _resolve_path(str(raw.get("database_path", "")), data_root / "output_database" / "kopfnoten_secure.db")
    backup_dir = _resolve_path(str(raw.get("backup_dir", "")), data_root / "db_backup")
    sph_config_path = _resolve_path(str(raw.get("sph_config_path", "")), data_root / "sph_config.json")

    paths = AppPaths(
        data_root=data_root,
        import_dir=import_dir,
        output_word_dir=output_word_dir,
        output_excel_dir=output_excel_dir,
        templates_dir=templates_dir,
        logs_dir=logs_dir,
        temp_dir=temp_dir,
        database_path=database_path,
        backup_dir=backup_dir,
        sph_config_path=sph_config_path,
        config_file=config_file,
    )
    paths.ensure_runtime_dirs()
    return paths

