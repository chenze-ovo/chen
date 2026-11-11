from fastapi import APIRouter, Depends

from app.deps import get_current_user
from app.models import User
from app.services.project1_runner import project1_status
from app.services.project2_stub import project2_status

router = APIRouter(prefix="/api/engines", tags=["engines"])


@router.get("/project1/status")
def project1(_user: User = Depends(get_current_user)) -> dict:
    return project1_status()


@router.get("/project2/status")
def project2(_user: User = Depends(get_current_user)) -> dict:
    return project2_status()
