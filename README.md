# Streamlit Sheets Dashboard

Resource planning web app built with Streamlit, DuckDB, and Google Sheets.

## What is implemented

- Streamlit app shell with login, Google OAuth handoff, workspace bootstrap, sync trigger, and blocked analytics state
- Startup diagnostics panel that shows missing configuration and the next required steps
- XLSM parser for `data/sandbox.xlsm`
- In-memory DuckDB sync pipeline with dimensions, fact table, and `inconsistency_log`
- Google Sheets / Drive integration scaffolding
- Migration script to create a canonical Google Sheets template from the local XLSM file
- Dockerfile and local helper scripts

## Local commands

Run tests:

```bash
./test.sh
```

Run the app:

```bash
./run.sh
```

Install or refresh dependencies:

```bash
./install.sh
```

## Docker run

Build and start the container:

```bash
docker compose up --build
```

Run it in the background:

```bash
docker compose up --build -d
```

Stop it:

```bash
docker compose down
```

The container exposes:

- `http://localhost:8501`

Persistent app metadata and stored Google OAuth tokens are written to the named Docker volume:

- `app_state`

## Environment setup

Copy the example env file:

```bash
cp .env.example .env
```

Set these values in `.env`:

```env
APP_URL=http://localhost:8501
APP_STATE_DIR=app_state
CANONICAL_TEMPLATE_SHEET_ID=
GOOGLE_CLIENT_ID=
GOOGLE_CLIENT_SECRET=
GOOGLE_PROJECT_ID=
```

When running in Docker, keep:

```env
APP_URL=http://localhost:8501
```

The app ignores:

- `.env`
- `.venv/`
- `.streamlit/secrets.toml`

## Google setup

Two separate auth layers are assumed:

1. Streamlit OIDC for app login via `st.login`
2. Google OAuth for Sheets and Drive access

### 1. Streamlit login

Configure your Streamlit auth settings according to the current `st.login` documentation:

- https://docs.streamlit.io/develop/api-reference/user/st.login

For local development, the redirect base should match:

- `APP_URL=http://localhost:8501`

### 2. Google OAuth client

Create a Google OAuth client in Google Cloud Console and allow at least these scopes:

- `openid`
- `https://www.googleapis.com/auth/userinfo.email`
- `https://www.googleapis.com/auth/spreadsheets`
- `https://www.googleapis.com/auth/drive`

Use the redirect URI:

- `http://localhost:8501/`

Then place the client ID and secret in `.env`.

## Canonical template migration

Before workspace onboarding can work, create the canonical Google Spreadsheet template once from the local XLSM file.

The migration script is:

- [scripts/migrate_to_gsheets.py](/home/jakob/dev/streamlit-sheets-dashboard/scripts/migrate_to_gsheets.py)

It expects a Google authorized-user credentials JSON file. Example:

```bash
python3 scripts/migrate_to_gsheets.py \
  --source data/sandbox.xlsm \
  --credentials /path/to/google-authorized-user.json \
  --title "Planning Template"
```

The script prints the new spreadsheet ID. Put that value into:

- `CANONICAL_TEMPLATE_SHEET_ID`

If you run the migration from inside Docker later, mount the credentials file into the container first. For now, the simplest path is to run the migration on the host.

## First run flow

1. Start the app with `./run.sh`
2. Log into the app through Streamlit OIDC
3. Authorize Google Sheets / Drive access
4. Click `Workspace einrichten`
5. The app copies the canonical spreadsheet into the user account
6. Click `Sync from Google Sheets`

For Docker, replace step 1 with:

1. Start the app with `docker compose up --build`

## Startup diagnostics

On every app start, open the `Startup Diagnostics` panel near the top of the page.

It checks:

- whether required `.env` values are missing or still placeholders
- whether Streamlit login is active
- whether Google Workspace access has been connected
- whether a workspace has already been assigned to the signed-in user

If something is missing, the panel tells you what to do next before the full app flow will work.

## Current MVP constraints

- In-app editing is not enabled yet
- `st.data_editor` is used read-only for review and analytics tables
- Analytics stay blocked while `inconsistency_log` contains errors
- Some real project sheets use month headers in row `19`, others in row `4`; the parser handles both
- Docker runtime is prepared for the app itself, but Google OAuth / Streamlit OIDC provider configuration must still be valid for `http://localhost:8501`

## Data model notes

- Source workbook field `fhId` is normalized to `userid`
- `P-XYTemplate` is used for duplication and workspace bootstrap only
- Only the block start rows `20, 35, ..., 170` are transformed into planning facts
- Monthly planning cells are read from `T:CM`

## References

- Streamlit `st.login`: https://docs.streamlit.io/develop/api-reference/user/st.login
- Streamlit `st.connection`: https://docs.streamlit.io/develop/api-reference/connections/st.connection
- Streamlit `st.data_editor`: https://docs.streamlit.io/develop/api-reference/data/st.data_editor
- DuckDB `httpfs`: https://duckdb.org/docs/stable/core_extensions/httpfs/overview.html
- DuckDB Python DB-API: https://duckdb.org/docs/stable/clients/python/dbapi.html
- gspread docs: https://docs.gspread.org/en/latest/
- python-dotenv: https://pypi.org/project/python-dotenv/
