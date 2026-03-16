from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import json
import sqlite3
from pathlib import Path


@dataclass(slots=True)
class WorkspaceRecord:
    user_email: str
    google_sheet_id: str
    workspace_state: str
    created_at: str
    last_synced_at: str | None


class MetadataStore:
    def __init__(self, db_path: str | Path):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        return connection

    def _initialize(self) -> None:
        with self._connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS workspace_registry (
                    user_email TEXT PRIMARY KEY,
                    google_sheet_id TEXT NOT NULL,
                    workspace_state TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    last_synced_at TEXT
                );
                CREATE TABLE IF NOT EXISTS oauth_tokens (
                    user_email TEXT PRIMARY KEY,
                    credentials_json TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                """
            )

    def get_workspace(self, user_email: str) -> WorkspaceRecord | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT user_email, google_sheet_id, workspace_state, created_at, last_synced_at "
                "FROM workspace_registry WHERE user_email = ?",
                (user_email,),
            ).fetchone()
        if row is None:
            return None
        return WorkspaceRecord(**dict(row))

    def upsert_workspace(self, user_email: str, google_sheet_id: str, workspace_state: str = "ready") -> WorkspaceRecord:
        now = datetime.now(timezone.utc).isoformat()
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO workspace_registry (user_email, google_sheet_id, workspace_state, created_at, last_synced_at)
                VALUES (?, ?, ?, ?, NULL)
                ON CONFLICT(user_email) DO UPDATE SET
                    google_sheet_id = excluded.google_sheet_id,
                    workspace_state = excluded.workspace_state
                """,
                (user_email, google_sheet_id, workspace_state, now),
            )
        return self.get_workspace(user_email)

    def set_last_synced(self, user_email: str) -> None:
        with self._connect() as conn:
            conn.execute(
                "UPDATE workspace_registry SET last_synced_at = ? WHERE user_email = ?",
                (datetime.now(timezone.utc).isoformat(), user_email),
            )

    def store_google_credentials(self, user_email: str, credentials_payload: dict) -> None:
        now = datetime.now(timezone.utc).isoformat()
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO oauth_tokens (user_email, credentials_json, updated_at)
                VALUES (?, ?, ?)
                ON CONFLICT(user_email) DO UPDATE SET
                    credentials_json = excluded.credentials_json,
                    updated_at = excluded.updated_at
                """,
                (user_email, json.dumps(credentials_payload), now),
            )

    def get_google_credentials(self, user_email: str) -> dict | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT credentials_json FROM oauth_tokens WHERE user_email = ?",
                (user_email,),
            ).fetchone()
        if row is None:
            return None
        return json.loads(row["credentials_json"])
