from __future__ import annotations

import pandas as pd
import streamlit as st

from planning.config import get_config
from planning.google_workspace import GoogleWorkspaceClient
from planning.metadata_store import MetadataStore
from planning.sync_service import PlanningSyncService
from planning.ui import (
    render_project_creation,
    render_workspace_setup_needed,
    workspace_title_for_user,
)


st.set_page_config(page_title="Resource Planning", layout="wide")

config = get_config()
metadata_store = MetadataStore(config.sqlite_path)
LOCAL_USER_EMAIL = "local-user@localhost"


@st.cache_resource
def metadata_connection(sqlite_url: str):
    return st.connection("metadata_db", type="sql", url=sqlite_url)


def _session_email() -> str:
    return LOCAL_USER_EMAIL


def _get_google_client() -> GoogleWorkspaceClient | None:
    credentials_path = config.google_authorized_user_path
    if not credentials_path.is_file():
        return None
    return GoogleWorkspaceClient.from_authorized_user_file(credentials_path)


def _startup_checks(user_email: str) -> list[dict[str, str]]:
    checks: list[dict[str, str]] = []
    missing_env = config.missing_env_items()
    checks.append(
        {
            "Check": "Environment",
            "Status": "ok" if not missing_env else "action_required",
            "Details": "All required .env values are configured."
            if not missing_env
            else f"Missing or placeholder values: {', '.join(missing_env)}",
        }
    )
    checks.append(
        {
            "Check": "App mode",
            "Status": "ok",
            "Details": f"Single-user local mode is active as `{user_email}`.",
        }
    )
    has_google_credentials = config.google_authorized_user_path.is_file()
    checks.append(
        {
            "Check": "Google Workspace access",
            "Status": "ok" if has_google_credentials else "action_required",
            "Details": f"Authorized-user credentials found at `{config.google_authorized_user_path}`."
            if has_google_credentials
            else f"Add a Google authorized-user JSON file at `{config.google_authorized_user_path}`.",
        }
    )
    workspace = metadata_store.get_workspace(user_email)
    checks.append(
        {
            "Check": "Workspace assignment",
            "Status": "ok" if workspace else "action_required",
            "Details": f"Workspace configured: {workspace.google_sheet_id}"
            if workspace
            else "Use 'Workspace einrichten' to create the local workspace copy.",
        }
    )
    return checks


def _render_startup_diagnostics(user_email: str) -> None:
    checks = _startup_checks(user_email)
    blocked = [item for item in checks if item["Status"] != "ok"]
    with st.expander("Startup Diagnostics", expanded=bool(blocked)):
        if blocked:
            st.warning("Configuration is incomplete. Finish the action-required items below before the full app flow will work.")
        else:
            st.success("Startup checks are complete.")
        st.dataframe(pd.DataFrame(checks), hide_index=True, width="stretch")
        if blocked:
            st.markdown(
                "\n".join(
                    [
                        "- Set required values in `.env` and restart the app.",
                        "- Keep `google-authorized-user.json` in the repo root or set `GOOGLE_AUTHORIZED_USER_PATH`.",
                        "- Run the migration script once and set `CANONICAL_TEMPLATE_SHEET_ID`.",
                    ]
                )
            )


def _ensure_workspace(user_email: str, client: GoogleWorkspaceClient) -> None:
    if not config.canonical_template_sheet_id:
        st.error("CANONICAL_TEMPLATE_SHEET_ID is not configured.")
        return
    workspace_id = client.copy_spreadsheet(
        config.canonical_template_sheet_id,
        workspace_title_for_user(user_email),
    )
    metadata_store.upsert_workspace(user_email, workspace_id)
    st.success("Workspace created.")


def _sync_workspace(user_email: str, client: GoogleWorkspaceClient, google_sheet_id: str) -> None:
    workbook = client.spreadsheet_to_workbook(google_sheet_id)
    sync_service = PlanningSyncService()
    sync_result = sync_service.sync_workbook(workbook)
    metadata_store.set_last_synced(user_email)
    st.session_state["sync_result"] = sync_result
    st.session_state["sync_service"] = sync_service


def _render_dashboard(sync_service: PlanningSyncService) -> None:
    sync_result = st.session_state.get("sync_result")
    if sync_result is None:
        return
    st.subheader("Sync Summary")
    metrics = pd.DataFrame(
        [
            {"Metric": "Run ID", "Value": sync_result.run_id},
            {"Metric": "Status", "Value": sync_result.status},
            {"Metric": "Projects", "Value": sync_result.project_count},
            {"Metric": "Fact Rows", "Value": sync_result.fact_row_count},
            {"Metric": "Issues", "Value": sync_result.issue_count},
        ]
    )
    st.dataframe(metrics, hide_index=True, width="stretch")
    issues = sync_service.issues()
    if not issues.empty:
        st.error("Analytics are blocked until the inconsistency log is clean.")
        st.subheader("To-Do List")
        st.data_editor(issues, disabled=True, hide_index=True, width="stretch")
        return
    st.subheader("Monthly Capacity")
    st.data_editor(sync_service.analytics_monthly_capacity(), disabled=True, hide_index=True, width="stretch")
    st.line_chart(sync_service.analytics_monthly_capacity(), x="month_date", y="total_hours", color="userid")
    st.subheader("Budget")
    st.data_editor(sync_service.analytics_budget(), disabled=True, hide_index=True, width="stretch")
    st.bar_chart(sync_service.analytics_budget(), x="month_date", y="budget", color="sheet_title")
    st.subheader("FTE Load")
    st.data_editor(sync_service.analytics_fte(), disabled=True, hide_index=True, width="stretch")
    st.line_chart(sync_service.analytics_fte(), x="month_date", y="fte_load", color="userid")


def main() -> None:
    metadata_connection(config.sqlite_url)
    user_email = _session_email()
    _render_startup_diagnostics(user_email)
    google_client = _get_google_client()
    if google_client is None:
        st.error(
            "Google access is not ready because the authorized-user JSON file is missing. "
            "Place it at `google-authorized-user.json` or set `GOOGLE_AUTHORIZED_USER_PATH`."
        )
        return
    workspace = metadata_store.get_workspace(user_email)
    if workspace is None:
        if render_workspace_setup_needed(user_email):
            _ensure_workspace(user_email, google_client)
            st.rerun()
        return
    st.title("Resource Planning")
    st.caption(f"Workspace: `{workspace.google_sheet_id}`")
    action_column, project_column = st.columns([1, 1])
    with action_column:
        if st.button("Sync from Google Sheets", type="primary"):
            _sync_workspace(user_email, google_client, workspace.google_sheet_id)
    with project_column:
        new_sheet_name = render_project_creation()
        if new_sheet_name:
            google_client.duplicate_project_template(workspace.google_sheet_id, new_sheet_name)
            st.success(f"Created project sheet `{new_sheet_name}`.")
    sync_service = st.session_state.get("sync_service")
    if sync_service is not None:
        _render_dashboard(sync_service)
    else:
        st.info("Run a sync to load analytics.")


if __name__ == "__main__":
    main()
