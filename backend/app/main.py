from contextlib import asynccontextmanager
import os
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.config import get_settings
from app.database import init_db
from app.routers import auth, batches, catalog, engines, health, orders, project1, reviews
from app.routers import settings as settings_router

_DEFAULT_DIST = Path(__file__).resolve().parents[2] / "frontend" / "dist"
DIST_DIR = Path(os.environ["ERP_DIST_DIR"]) if os.environ.get("ERP_DIST_DIR") else _DEFAULT_DIST


@asynccontextmanager
async def lifespan(_app: FastAPI):
    init_db()
    yield


def create_app() -> FastAPI:
    settings = get_settings()
    app = FastAPI(title="订单智转", version="1.0.0", lifespan=lifespan)
    web_mode = os.environ.get("ERP_UI_MODE", "").strip().lower() == "web"
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"] if web_mode else settings.cors_origin_list,
        allow_credentials=not web_mode,
        allow_methods=["*"],
        allow_headers=["*"],
    )
    app.include_router(health.router)
    app.include_router(auth.router)
    app.include_router(orders.router)
    app.include_router(batches.router)
    app.include_router(engines.router)
    app.include_router(project1.router)
    app.include_router(reviews.router)
    app.include_router(catalog.router)
    app.include_router(settings_router.router)

    if DIST_DIR.is_dir() and (DIST_DIR / "index.html").is_file():
        assets = DIST_DIR / "assets"
        if assets.is_dir():
            app.mount("/assets", StaticFiles(directory=str(assets)), name="assets")

        @app.get("/")
        def spa_root():
            return FileResponse(DIST_DIR / "index.html")

        @app.get("/{full_path:path}")
        def spa_fallback(full_path: str):
            # API routes already matched above; serve static or SPA shell.
            candidate = DIST_DIR / full_path
            if full_path and candidate.is_file():
                return FileResponse(candidate)
            return FileResponse(DIST_DIR / "index.html")
    else:

        @app.get("/")
        def root() -> dict:
            return {"service": "订单智转", "health": "/health", "docs": "/docs"}

    return app


app = create_app()
