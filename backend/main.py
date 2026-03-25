from __future__ import annotations

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from backend.config import ApiConfig
from backend.routers.attachments import router as attachments_router
from backend.routers.health import router as health_router
from backend.routers.modules import router as modules_router
from backend.routers.products import router as products_router
from backend.routers.projects import router as projects_router
from backend.routers.repository import router as repository_router


app = FastAPI(
    title=ApiConfig.TITLE,
    version=ApiConfig.VERSION,
)

if ApiConfig.CORS_ORIGINS:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=ApiConfig.CORS_ORIGINS,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )


@app.get("/")
def root() -> dict[str, str]:
    return {
        "service": ApiConfig.TITLE,
        "version": ApiConfig.VERSION,
        "docs": "/docs",
    }


app.include_router(health_router, prefix=ApiConfig.PREFIX)
app.include_router(modules_router, prefix=ApiConfig.PREFIX)
app.include_router(products_router, prefix=ApiConfig.PREFIX)
app.include_router(projects_router, prefix=ApiConfig.PREFIX)
app.include_router(attachments_router, prefix=ApiConfig.PREFIX)
app.include_router(repository_router, prefix=ApiConfig.PREFIX)
