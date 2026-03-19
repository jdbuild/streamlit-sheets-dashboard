from __future__ import annotations

import streamlit as st

from planning.google_workspace import slugify_project_name


def render_project_creation() -> str | None:
    with st.form("new-project-form", clear_on_submit=True):
        project_name = st.text_input("New project name", placeholder="Example Project")
        submitted = st.form_submit_button("+ Create project")
    if submitted and project_name.strip():
        return f"P-{slugify_project_name(project_name)}"
    return None
