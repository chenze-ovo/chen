from __future__ import annotations

import jwt
from fastapi import Depends, HTTPException
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from sqlalchemy import select
from sqlalchemy.orm import Session

from app.database import get_db
from app.models import User
from app.security import decode_token, hash_password
from app.models import utcnow

bearer = HTTPBearer(auto_error=False)


def _ensure_admin(db: Session) -> User:
    """无 token 时落到库里的 admin；没有则现场建一个。"""
    user = db.scalar(select(User).where(User.username == "admin"))
    if user is None:
        user = User(
            username="admin",
            password_hash=hash_password("admin123"),
            role="admin",
            display_name="系统管理员",
            is_active=True,
            created_at=utcnow(),
        )
        db.add(user)
        db.commit()
        db.refresh(user)
    elif not user.is_active:
        user.is_active = True
        db.commit()
        db.refresh(user)
    return user


def get_current_user(
    creds: HTTPAuthorizationCredentials | None = Depends(bearer),
    db: Session = Depends(get_db),
) -> User:
    # 无登录模式：缺 token / 无效 token 一律当 admin，避免 401 踢人
    if creds is None or not creds.credentials:
        return _ensure_admin(db)
    try:
        payload = decode_token(creds.credentials)
        user_id = int(payload["uid"])
    except (jwt.PyJWTError, KeyError, TypeError, ValueError):
        return _ensure_admin(db)
    user = db.get(User, user_id)
    if user is None or not user.is_active:
        return _ensure_admin(db)
    return user


def require_admin(user: User = Depends(get_current_user)) -> User:
    if user.role != "admin":
        raise HTTPException(status_code=403, detail="需要管理员权限")
    return user


def require_roles(*roles: str):
    def checker(user: User = Depends(get_current_user)) -> User:
        if user.role == "admin" or user.role in roles:
            return user
        raise HTTPException(status_code=403, detail="没有权限执行此操作")

    return checker
