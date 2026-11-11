from fastapi import APIRouter, Depends
from sqlalchemy import text
from sqlalchemy.orm import Session

from app.config import get_settings
from app.database import get_db

router = APIRouter(tags=["health"])


@router.get("/health")
def health(db: Session = Depends(get_db)) -> dict:
    db.execute(text("SELECT 1"))
    url = get_settings().resolved_database_url
    if url.startswith("postgresql"):
        database = "postgresql"
    elif url.startswith("sqlite"):
        database = "sqlite"
    else:
        database = "other"
    return {"status": "ok", "service": "订单智转", "database": database, "phase": "0"}
