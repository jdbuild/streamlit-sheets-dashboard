from __future__ import annotations

from datetime import date

import pandas as pd
import streamlit as st

from planning.config import get_config
from planning.google_workspace import GoogleWorkspaceClient
from planning.sync_service import PlanningSyncService
from planning.ui import (
    render_project_creation,
)


st.set_page_config(page_title="Resource Planning", layout="wide")

config = get_config()


def _get_google_client() -> GoogleWorkspaceClient | None:
    credentials_path = config.google_authorized_user_path
    if not credentials_path.is_file():
        return None
    return GoogleWorkspaceClient.from_authorized_user_file(credentials_path)


def _startup_checks() -> list[dict[str, str]]:
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
            "Details": "Single-user local mode is active.",
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
    checks.append(
        {
            "Check": "Target spreadsheet",
            "Status": "ok" if config.canonical_template_sheet_id else "action_required",
            "Details": f"Using canonical spreadsheet `{config.canonical_template_sheet_id}`."
            if config.canonical_template_sheet_id
            else "Set `CANONICAL_TEMPLATE_SHEET_ID` to the Google Sheet the app should read and modify.",
        }
    )
    return checks


def _render_startup_diagnostics() -> None:
    checks = _startup_checks()
    blocked = [item for item in checks if item["Status"] != "ok"]
    with st.expander("1. Startup Diagnostics", expanded=bool(blocked)):
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
                        "- Set `CANONICAL_TEMPLATE_SHEET_ID` to the spreadsheet you want the app to use directly.",
                    ]
                )
            )


def _sync_workspace(client: GoogleWorkspaceClient, google_sheet_id: str) -> None:
    workbook = client.spreadsheet_to_workbook(google_sheet_id)
    sync_service = PlanningSyncService()
    sync_result = sync_service.sync_workbook(workbook)
    st.session_state["sync_result"] = sync_result
    st.session_state["sync_service"] = sync_service


def _render_sync_results(sync_service: PlanningSyncService) -> None:
    sync_result = st.session_state.get("sync_result")
    if sync_result is None:
        return
    issues = sync_service.issues()
    errors = issues[issues["severity"] == "error"] if not issues.empty else issues
    warnings = issues[issues["severity"] != "error"] if not issues.empty else issues
    metrics = pd.DataFrame(
        [
            {"Metric": "Run ID", "Value": sync_result.run_id},
            {"Metric": "Status", "Value": sync_result.status},
            {"Metric": "Projects", "Value": sync_result.project_count},
            {"Metric": "Fact Rows", "Value": sync_result.fact_row_count},
            {"Metric": "Issues", "Value": sync_result.issue_count},
            {"Metric": "Errors", "Value": 0 if errors.empty else len(errors.index)},
            {"Metric": "Warnings", "Value": 0 if warnings.empty else len(warnings.index)},
        ]
    )
    metrics["Value"] = metrics["Value"].astype(str)
    st.dataframe(metrics, hide_index=True, width="stretch")
    if not errors.empty:
        st.error("Sync completed with blocking errors. Step 3 stays blocked until the error list is clean.")
        st.subheader("Errors")
        st.data_editor(errors, disabled=True, hide_index=True, width="stretch")
    elif not warnings.empty:
        st.warning("Sync completed with warnings. Step 3 analytics remains available.")
        st.subheader("Warnings")
        st.data_editor(warnings, disabled=True, hide_index=True, width="stretch")
    else:
        st.success("Sync completed without issues.")


def _apply_analytics_filters(
    dataframe: pd.DataFrame,
    selected_userids: list[str],
    date_range: tuple,
) -> pd.DataFrame:
    filtered = dataframe.copy()
    if "userid" in filtered.columns:
        filtered = filtered[filtered["userid"].isin(selected_userids)]
    if "month_date" in filtered.columns and not filtered.empty:
        month_dates = pd.to_datetime(filtered["month_date"]).dt.date
        filtered = filtered[(month_dates >= date_range[0]) & (month_dates <= date_range[1])]
    return filtered


def _round_numeric_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    rounded = dataframe.copy()
    numeric_columns = rounded.select_dtypes(include="number").columns
    if len(numeric_columns) > 0:
        rounded[numeric_columns] = rounded[numeric_columns].round(0)
    return rounded


def _monthly_capacity_pivot(dataframe: pd.DataFrame) -> pd.DataFrame:
    if dataframe.empty:
        return dataframe
    normalized = dataframe.copy()
    normalized["month_date"] = pd.to_datetime(normalized["month_date"])
    normalized["project"] = normalized["sheet_title"].str.removeprefix("P-")
    month_columns = sorted(normalized["month_date"].drop_duplicates())
    detail = normalized.pivot_table(
        index=["userid", "project"],
        columns="month_date",
        values="total_hours",
        aggfunc="sum",
        fill_value=0,
    )
    detail = detail.reindex(columns=month_columns, fill_value=0).reset_index()
    detail.columns = ["userid", "project", *[month.strftime("%m/%y") for month in month_columns]]
    month_labels = [month.strftime("%m/%y") for month in month_columns]
    rows: list[dict] = []
    for userid, group in detail.groupby("userid", sort=True):
        for _, row in group.sort_values("project").iterrows():
            rows.append(row.to_dict())
        total_row = {"userid": userid, "project": "Total"}
        for label in month_labels:
            total_row[label] = float(group[label].sum())
        rows.append(total_row)
    return _round_numeric_columns(pd.DataFrame(rows, columns=["userid", "project", *month_labels]))


def _render_analytics_step(sync_service: PlanningSyncService) -> None:
    sync_result = st.session_state.get("sync_result")
    if sync_result is None:
        return
    issues = sync_service.issues()
    errors = issues[issues["severity"] == "error"] if not issues.empty else issues
    st.subheader("3. Visual Analytics")
    if not errors.empty:
        st.info("Run step 2 sync without blocking errors to unlock analytics.")
        return
    people_df = sync_service.people()
    if people_df.empty:
        st.info("No people were found in the synced planning rows.")
        return
    capacity_df = sync_service.analytics_monthly_capacity_detail()
    fte_df = sync_service.analytics_fte()
    budget_detail_df = sync_service.analytics_budget_detail()
    all_userids = people_df["userid"].tolist()
    default_userids = ["vleung"] if "vleung" in all_userids else all_userids
    label_by_userid = {
        row.userid: f"{row.userid} - {row.person_name}" if row.person_name not in (None, "", row.userid) else row.userid
        for row in people_df.itertuples(index=False)
    }
    all_dates = pd.concat(
        [
            pd.to_datetime(capacity_df["month_date"]),
            pd.to_datetime(fte_df["month_date"]),
            pd.to_datetime(budget_detail_df["month_date"]),
        ],
        ignore_index=True,
    )
    min_date = all_dates.min().date()
    max_date = all_dates.max().date()
    preferred_start = max(min_date, date(2026, 1, 1))
    preferred_end = min(max_date, date(2027, 12, 1))
    default_date_range = (
        (preferred_start, preferred_end) if preferred_start <= preferred_end else (min_date, max_date)
    )
    filter_column, info_column = st.columns([2, 1])
    with filter_column:
        select_column, clear_column = st.columns([1, 1])
        with select_column:
            if st.button("Select all people", use_container_width=True):
                st.session_state["analytics_people"] = all_userids
        with clear_column:
            if st.button("Clear people", use_container_width=True):
                st.session_state["analytics_people"] = []
        selected_userids = st.multiselect(
            "People",
            options=all_userids,
            default=st.session_state.get("analytics_people", default_userids),
            format_func=lambda userid: label_by_userid.get(userid, userid),
            key="analytics_people",
            placeholder="Select one or more people",
        )
        selected_dates = st.date_input(
            "Time interval",
            value=default_date_range,
            min_value=min_date,
            max_value=max_date,
        )
    with info_column:
        st.metric("People found", len(all_userids))
        st.metric("Selected", len(selected_userids))
    if isinstance(selected_dates, tuple) and len(selected_dates) == 2:
        date_range = selected_dates
    else:
        date_range = (min_date, max_date)
    if not selected_userids:
        st.info("Select at least one person to view analytics.")
        return
    capacity_filtered = _apply_analytics_filters(capacity_df, selected_userids, date_range)
    fte_filtered = _apply_analytics_filters(fte_df, selected_userids, date_range)
    budget_filtered = _apply_analytics_filters(budget_detail_df, selected_userids, date_range)
    capacity_pivot_df = _monthly_capacity_pivot(capacity_filtered)
    budget_chart_df = (
        budget_filtered.groupby(["sheet_title", "month_date"], as_index=False)["budget"].sum()
        if not budget_filtered.empty
        else budget_filtered
    )
    st.subheader("Monthly Capacity")
    st.data_editor(capacity_pivot_df, disabled=True, hide_index=True, width="stretch")
    st.subheader("Budget")
    budget_chart_df = _round_numeric_columns(budget_chart_df)
    st.data_editor(budget_chart_df, disabled=True, hide_index=True, width="stretch")
    st.bar_chart(budget_chart_df, x="month_date", y="budget", color="sheet_title")
    st.subheader("FTE Load")
    fte_filtered = _round_numeric_columns(fte_filtered)
    st.data_editor(fte_filtered, disabled=True, hide_index=True, width="stretch")
    st.line_chart(fte_filtered, x="month_date", y="fte_load", color="userid")


def main() -> None:
    st.title("Resource Planning")
    _render_startup_diagnostics()
    google_client = _get_google_client()
    if google_client is None:
        st.error(
            "Google access is not ready because the authorized-user JSON file is missing. "
            "Place it at `google-authorized-user.json` or set `GOOGLE_AUTHORIZED_USER_PATH`."
        )
        return
    if not config.canonical_template_sheet_id:
        st.error("`CANONICAL_TEMPLATE_SHEET_ID` is not configured.")
        return
    target_sheet_id = config.canonical_template_sheet_id
    st.caption(f"Google Sheet: `{target_sheet_id}`")
    sync_service = st.session_state.get("sync_service")
    sync_result = st.session_state.get("sync_result")
    sync_expanded = sync_service is None or (sync_result is not None and sync_result.status != "clean")
    with st.expander("2. Sync", expanded=sync_expanded):
        action_column, project_column = st.columns([1, 1])
        with action_column:
            if st.button("Sync from Google Sheets", type="primary"):
                _sync_workspace(google_client, target_sheet_id)
        with project_column:
            new_sheet_name = render_project_creation()
            if new_sheet_name:
                google_client.duplicate_project_template(target_sheet_id, new_sheet_name)
                st.success(f"Created project sheet `{new_sheet_name}`.")
        sync_service = st.session_state.get("sync_service")
        if sync_service is not None:
            _render_sync_results(sync_service)
        else:
            st.info("Run sync to load results and analytics.")
    sync_service = st.session_state.get("sync_service")
    if sync_service is not None:
        _render_analytics_step(sync_service)
    else:
        st.subheader("3. Visual Analytics")
        st.info("Run step 2 sync to load analytics.")


if __name__ == "__main__":
    main()
