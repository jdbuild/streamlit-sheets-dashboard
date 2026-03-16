from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import os

from dotenv import load_dotenv


load_dotenv()


def _default_app_url() -> str:
    return os.getenv("APP_URL", "http://localhost:8501")


def _looks_unset(value: str | None) -> bool:
    if value is None:
        return True
    normalized = value.strip()
    if not normalized:
        return True
    return normalized.startswith("test-")


@dataclass(slots=True)
class AppConfig:
    app_url: str = _default_app_url()
    app_state_dir: Path = Path(os.getenv("APP_STATE_DIR", "app_state"))
    sqlite_path: Path = Path(os.getenv("APP_STATE_DIR", "app_state")) / "metadata.db"
    canonical_template_sheet_id: str | None = os.getenv("CANONICAL_TEMPLATE_SHEET_ID")
    google_client_id: str | None = os.getenv("GOOGLE_CLIENT_ID")
    google_client_secret: str | None = os.getenv("GOOGLE_CLIENT_SECRET")
    google_project_id: str | None = os.getenv("GOOGLE_PROJECT_ID")
    google_oauth_scopes: tuple[str, ...] = (
        "openid",
        "https://www.googleapis.com/auth/userinfo.email",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    )

    @property
    def sqlite_url(self) -> str:
        return f"sqlite:///{self.sqlite_path}"

    @property
    def google_oauth_redirect_uri(self) -> str:
        return f"{self.app_url.rstrip('/')}/"

    def ensure_state_dir(self) -> None:
        self.app_state_dir.mkdir(parents=True, exist_ok=True)

    def missing_env_items(self) -> list[str]:
        missing: list[str] = []
        if _looks_unset(self.google_client_id):
            missing.append("GOOGLE_CLIENT_ID")
        if _looks_unset(self.google_client_secret):
            missing.append("GOOGLE_CLIENT_SECRET")
        if _looks_unset(self.google_project_id):
            missing.append("GOOGLE_PROJECT_ID")
        if _looks_unset(self.canonical_template_sheet_id):
            missing.append("CANONICAL_TEMPLATE_SHEET_ID")
        return missing


def get_config() -> AppConfig:
    config = AppConfig()
    config.ensure_state_dir()
    return config
