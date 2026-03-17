from __future__ import annotations

import streamlit as st

from planning.google_workspace import build_workspace_title, slugify_project_name


def render_workspace_setup_needed(user_email: str) -> bool:
    st.title("Workspace Setup")
    st.write(f"No workspace found for `{user_email}`.")
    return st.button("Workspace einrichten", type="primary")


def render_project_creation() -> str | None:
    with st.form("new-project-form", clear_on_submit=True):
        project_name = st.text_input("New project name", placeholder="Example Project")
        submitted = st.form_submit_button("+ Create project")
    if submitted and project_name.strip():
        return f"P-{slugify_project_name(project_name)}"
    return None


def workspace_title_for_user(user_email: str) -> str:
    return build_workspace_title(user_email)
