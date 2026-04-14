from __future__ import annotations

import os
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


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_ENV_FILE = ROOT_DIR / ".env"
DEFAULT_CA_CERT = ROOT_DIR / "certs" / "ca-certificate.crt"


def load_project_env(env_file: Path | None = None) -> Path:
    env_path = env_file or DEFAULT_ENV_FILE
    load_dotenv(env_path)
    return env_path


def require_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def get_database_settings() -> dict[str, str]:
    sslrootcert = os.getenv("PGSSLROOTCERT", "").strip() or str(DEFAULT_CA_CERT)
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
