from __future__ import annotations

from fastapi import APIRouter

from backend.db import ping_database


router = APIRouter(tags=["health"])


@router.get("/health")
def health() -> dict[str, object]:
    return {
        "status": "ok",
        "database": ping_database(),
    }
