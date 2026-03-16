from __future__ import annotations

from dataclasses import dataclass
import json

from google_auth_oauthlib.flow import Flow

from planning.config import AppConfig


GOOGLE_SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


@dataclass(slots=True)
class OAuthStart:
    authorization_url: str
    state: str


def _client_config(config: AppConfig) -> dict:
    if not config.google_client_id or not config.google_client_secret:
        raise RuntimeError("Google OAuth client configuration is incomplete.")
    return {
        "web": {
            "client_id": config.google_client_id,
            "project_id": config.google_project_id or "streamlit-sheets-dashboard",
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_secret": config.google_client_secret,
            "redirect_uris": [config.google_oauth_redirect_uri],
        }
    }


def build_google_oauth_start(config: AppConfig) -> OAuthStart:
    flow = Flow.from_client_config(
        _client_config(config),
        scopes=GOOGLE_SCOPES,
        redirect_uri=config.google_oauth_redirect_uri,
    )
    authorization_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return OAuthStart(authorization_url=authorization_url, state=state)


def exchange_google_oauth_code(config: AppConfig, state: str, code: str) -> dict:
    flow = Flow.from_client_config(
        _client_config(config),
        scopes=GOOGLE_SCOPES,
        state=state,
        redirect_uri=config.google_oauth_redirect_uri,
    )
    flow.fetch_token(code=code)
    credentials = flow.credentials
    return json.loads(credentials.to_json())
