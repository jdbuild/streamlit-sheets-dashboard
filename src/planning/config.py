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
    google_authorized_user_path: Path = Path(os.getenv("GOOGLE_AUTHORIZED_USER_PATH", "google-authorized-user.json"))

    @property
    def sqlite_url(self) -> str:
        return f"sqlite:///{self.sqlite_path}"

    def ensure_state_dir(self) -> None:
        self.app_state_dir.mkdir(parents=True, exist_ok=True)

    def missing_env_items(self) -> list[str]:
        missing: list[str] = []
        if _looks_unset(self.canonical_template_sheet_id):
            missing.append("CANONICAL_TEMPLATE_SHEET_ID")
        if not self.google_authorized_user_path.is_file():
            missing.append(f"GOOGLE_AUTHORIZED_USER_PATH ({self.google_authorized_user_path})")
        return missing


def get_config() -> AppConfig:
    config = AppConfig()
    config.ensure_state_dir()
    return config
