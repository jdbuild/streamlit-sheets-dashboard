# Streamlit Sheets Dashboard

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Python 3.12+](https://img.shields.io/badge/Python-3.12%2B-3776AB.svg?logo=python&logoColor=white)](./pyproject.toml)
[![Streamlit](https://img.shields.io/badge/Streamlit-App-FF4B4B.svg?logo=streamlit&logoColor=white)](https://streamlit.io/)
[![Google Sheets](https://img.shields.io/badge/Google%20Sheets-API-34A853.svg?logo=googlesheets&logoColor=white)](https://developers.google.com/sheets/api)
[![DuckDB](https://img.shields.io/badge/DuckDB-Analytics-FECD45.svg)](https://duckdb.org/)

Resource planning web app built with Streamlit, DuckDB, and Google Sheets for a single local user.

## What is implemented

- Streamlit app shell with local single-user Google credentials, direct canonical-sheet sync, and blocked analytics state
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

Persistent app metadata is written to the named Docker volume:

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
GOOGLE_AUTHORIZED_USER_PATH=google-authorized-user.json
```

When running in Docker, keep:

```env
APP_URL=http://localhost:8501
```

Place the authorized-user JSON used by the migration script at the repo root as:

- `google-authorized-user.json`

Or point to it explicitly with:

- `GOOGLE_AUTHORIZED_USER_PATH=/absolute/path/to/google-authorized-user.json`

The app ignores:

- `.env`
- `.venv/`
- `.streamlit/secrets.toml`

## Google setup

This project now runs in single-user local mode. The app and the migration script both use the same Google authorized-user credentials JSON.

Required:

- a Google authorized-user JSON file with Sheets and Drive access
- the file available at `google-authorized-user.json` or via `GOOGLE_AUTHORIZED_USER_PATH`

## Canonical template migration

Before the app can run, create the canonical Google Spreadsheet once from the local XLSM file.

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

The app uses that spreadsheet directly for sync and project creation. It does not create a separate workspace copy.

If you run the migration from inside Docker later, mount the credentials file into the container first. For now, the simplest path is to run the migration on the host.

## First run flow

1. Start the app with `./run.sh`
2. Open the app at `http://localhost:8501`
3. Click `Sync from Google Sheets`
4. Use `+ Create project` to duplicate `P-XYTemplate` inside that spreadsheet

For Docker, replace step 1 with:

1. Start the app with `docker compose up --build`

## Startup diagnostics

On every app start, open the `Startup Diagnostics` panel near the top of the page.

It checks:

- whether required `.env` values are missing or still placeholders
- whether the authorized-user JSON is available
- whether the canonical spreadsheet ID is configured

If something is missing, the panel tells you what to do next before the full app flow will work.

## Current MVP constraints

- In-app editing is not enabled yet
- `st.data_editor` is used read-only for review and analytics tables
- Analytics stay blocked while `inconsistency_log` contains errors
- Some real project sheets use month headers in row `19`, others in row `4`; the parser handles both
- Docker runtime is prepared for the app itself, but the authorized-user JSON must still be mounted or made available inside the container

## Data model notes

- Source workbook field `fhId` is normalized to `userid`
- `P-XYTemplate` is used for duplication inside the canonical spreadsheet
- Person blocks start at rows `20, 31, 42, ...` with block height `11`
- The block userid is read from `B{block_start + 1}`
- The 8 WP rows in each person block are the fact source for `project + person + wp + month`
- The summary row at `block_start + 8` is used to validate the summed WP rows for that person block
- Monthly planning cells are read from `T:CM`; empty cells stay empty

## ERD

The intended analytical grain is:

- one fact row per `project + person + wp + month`

Logical ERD:

```text
Project (1) ----< PlanningFact >---- (1) Person
                     |
                     v
                    WP
                     |
                     v
                   Month
```

Entity summary:

- `Project`
  - source: one Google Sheet tab per project such as `P-CT-Extention`
- `Person`
  - source: person block metadata including `person_slot`, display name, and `userid`
- `WP`
  - source: WP labels in the project block and the WP dimension table in `A210:C218`
- `Month`
  - source: month headers in `T:CM`
- `PlanningFact`
  - source: the 8 WP rows inside each person block
  - grain: monthly hours for one person on one WP in one project

Validation rule:

- the person summary row at `block_start + 8` is not the fact source
- it is used to validate that the sum of the 8 WP rows matches the expected per-person monthly total

## References

- Streamlit `st.connection`: https://docs.streamlit.io/develop/api-reference/connections/st.connection
- Streamlit `st.data_editor`: https://docs.streamlit.io/develop/api-reference/data/st.data_editor
- DuckDB `httpfs`: https://duckdb.org/docs/stable/core_extensions/httpfs/overview.html
- DuckDB Python DB-API: https://duckdb.org/docs/stable/clients/python/dbapi.html
- gspread docs: https://docs.gspread.org/en/latest/
- python-dotenv: https://pypi.org/project/python-dotenv/

## License

MIT. See [LICENSE](/home/jakob/dev/streamlit-sheets-dashboard/LICENSE).
