from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from urllib.error import URLError
from urllib.request import urlopen


ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from db_config import load_project_env


def main() -> int:
    env_path = load_project_env()
    base_url = os.getenv("ERP_API_BASE_URL", "http://127.0.0.1:3000").rstrip("/")
    api_prefix = os.getenv("ERP_API_PREFIX", "").strip()
    health_url = f"{base_url}{api_prefix}/"

    print(f"Loaded environment from: {env_path}")
    print(f"Checking PostgREST root: {health_url}")

    try:
        with urlopen(health_url, timeout=10) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except URLError as exc:
        print(f"PostgREST health check failed: {exc}")
        return 1

    print("PostgREST health check succeeded.")
    print(json.dumps(payload, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
