from __future__ import annotations

import os

from db_config import load_project_env


load_project_env()


class ApiConfig:
    TITLE = os.getenv("ERP_API_TITLE", "Liquimech ERP API")
    VERSION = os.getenv("ERP_API_VERSION", "0.1.0")
    HOST = os.getenv("ERP_API_HOST", "127.0.0.1")
    PORT = int(os.getenv("ERP_API_PORT", "8000"))
    PREFIX = os.getenv("ERP_API_PREFIX", "/api/v1")
    CORS_ORIGINS = [origin.strip() for origin in os.getenv("ERP_API_CORS_ORIGINS", "").split(",") if origin.strip()]
