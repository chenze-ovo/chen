"""运行配置。数据库默认 SQLite，正式库用 PostgreSQL。"""
from __future__ import annotations

import os
from functools import lru_cache
from pathlib import Path

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict

_DEFAULT_ROOT = Path(__file__).resolve().parents[2]


def _default_root() -> Path:
    raw = os.environ.get("ERP_REPO_ROOT")
    return Path(raw) if raw else _DEFAULT_ROOT


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=str(_DEFAULT_ROOT / ".env"),
        env_file_encoding="utf-8",
        extra="ignore",
    )

    database_url: str | None = None
    jwt_secret: str = "dev-only-change-in-production"
    jwt_expire_minutes: int = 60 * 12
    cors_origins: str = "http://127.0.0.1:8742,http://localhost:8742"
    erp_repo_root: Path = Field(default_factory=_default_root)
    erp_data_dir: Path | None = None
    order_convert_src: Path | None = None
    project2_routing: Path | None = None

    @property
    def data_dir(self) -> Path:
        if self.erp_data_dir:
            return Path(self.erp_data_dir)
        return self.erp_repo_root / "data"

    @property
    def resolved_database_url(self) -> str:
        if self.database_url:
            return self.database_url
        return "sqlite:///" + str((self.data_dir / "erp.db").resolve())

    @property
    def order_convert_src_path(self) -> Path:
        if self.order_convert_src:
            return Path(self.order_convert_src)
        return (
            self.erp_repo_root
            / "engines"
            / "project1"
            / "vendor"
            / "order-convert"
            / "src"
        )

    @property
    def project2_routing_path(self) -> Path:
        if self.project2_routing:
            return Path(self.project2_routing)
        return self.erp_repo_root / "engines" / "project2" / "routing.yaml"

    @property
    def cors_origin_list(self) -> list[str]:
        return [item.strip() for item in self.cors_origins.split(",") if item.strip()]


@lru_cache
def get_settings() -> Settings:
    return Settings()
