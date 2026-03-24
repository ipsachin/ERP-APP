from __future__ import annotations

import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from db_config import connect, get_database_settings, load_project_env


def main() -> int:
    try:
        env_path = load_project_env()
        kwargs = get_database_settings()
        cert_path = Path(kwargs["sslrootcert"])
        if not cert_path.exists():
            raise FileNotFoundError(f"CA certificate not found: {cert_path}")

        print(f"Loaded environment from: {env_path}")
        print("Connecting to PostgreSQL with TLS verification...")
        print(f"Host: {kwargs['host']}")
        print(f"Port: {kwargs['port']}")
        print(f"Database: {kwargs['dbname']}")
        print(f"User: {kwargs['user']}")
        print(f"SSL mode: {kwargs['sslmode']}")
        print(f"CA cert: {cert_path}")

        with connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    select
                        current_database(),
                        current_user,
                        version(),
                        now()
                    """
                )
                database_name, current_user, version_text, server_time = cur.fetchone()

        print("\nConnection successful.")
        print(f"Connected database: {database_name}")
        print(f"Authenticated user: {current_user}")
        print(f"Server time: {server_time}")
        print(f"Server version: {version_text.split(',')[0]}")
        return 0
    except Exception as exc:
        print("\nConnection failed.")
        print(f"{type(exc).__name__}: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
