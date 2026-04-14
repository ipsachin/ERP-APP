from __future__ import annotations

import os
import sys
from importlib import import_module
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:
    def load_dotenv(env_path: Path) -> bool:
        if not env_path.exists():
            return False
        for raw_line in env_path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key:
                os.environ.setdefault(key, value)
        return True


BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
DEFAULT_ENV_FILE = APP_DIR / ".env"
DEFAULT_CA_CERT = BUNDLE_DIR / "certs" / "ca-certificate.crt"
_last_loaded_env_path: Path | None = None


def find_env_file() -> Path:
    candidates = [
        APP_DIR / ".env",
        BUNDLE_DIR / ".env",
        Path.cwd() / ".env",
        APP_DIR / ".env.example",
        BUNDLE_DIR / ".env.example",
        Path.cwd() / ".env.example",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return DEFAULT_ENV_FILE


def load_project_env(env_file: Path | None = None) -> Path:
    global _last_loaded_env_path
    env_path = env_file or find_env_file()
    load_dotenv(env_path)
    _last_loaded_env_path = env_path
    return env_path


def resolve_config_path(value: str, default_path: Path) -> str:
    raw = (value or "").strip()
    if not raw:
        return str(default_path)
    path = Path(raw).expanduser()
    if path.is_absolute():
        return str(path)
    relative_bases = []
    if _last_loaded_env_path is not None:
        relative_bases.append(_last_loaded_env_path.parent)
    relative_bases.extend([APP_DIR, BUNDLE_DIR, Path.cwd()])
    for base in relative_bases:
        candidate = (base / path).resolve()
        if candidate.exists():
            return str(candidate)
    return str((APP_DIR / path).resolve())


def require_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def get_database_settings() -> dict[str, str]:
    sslrootcert = resolve_config_path(os.getenv("PGSSLROOTCERT", ""), DEFAULT_CA_CERT)
    return {
        "host": require_env("PGHOST"),
        "port": require_env("PGPORT"),
        "dbname": require_env("PGDATABASE"),
        "user": require_env("PGUSER"),
        "password": require_env("PGPASSWORD"),
        "sslmode": os.getenv("PGSSLMODE", "verify-full").strip() or "verify-full",
        "sslrootcert": sslrootcert,
        "connect_timeout": os.getenv("PGCONNECT_TIMEOUT", "10").strip() or "10",
    }


def connect(**overrides):
    settings = get_database_settings()
    settings.update({k: v for k, v in overrides.items() if v is not None})
    psycopg = import_module("psycopg")
    return psycopg.connect(**settings)
