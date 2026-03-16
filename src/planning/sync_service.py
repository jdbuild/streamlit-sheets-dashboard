from __future__ import annotations

from dataclasses import asdict
from uuid import uuid4
import re

import duckdb
import pandas as pd

from planning.models import GLOBAL_MONTHLY_HOURS_CAP, SyncResult, TEMPLATE_SHEET
from planning.xlsm_parser import ParsedWorkbook, ParsedProjectBlock, parse_openpyxl_workbook


def _project_id(sheet_title: str) -> str:
    normalized = re.sub(r"[^a-z0-9]+", "-", sheet_title.lower())
    return normalized.strip("-")


def _parse_wp_label(value: str | None) -> tuple[str | None, str | None]:
    if not value:
        return None, None
    parts = value.split(" ", 1)
    if len(parts) == 1:
        return parts[0].strip(), None
    return parts[0].strip(), parts[1].strip()


class PlanningSyncService:
    def __init__(self) -> None:
        self.connection = duckdb.connect(":memory:")
        try:
            self.connection.execute("INSTALL httpfs")
            self.connection.execute("LOAD httpfs")
        except duckdb.Error:
            # The sync pipeline does not require httpfs for the MVP path.
            pass

    def sync_workbook(self, workbook) -> SyncResult:
        parsed = parse_openpyxl_workbook(workbook)
        return self.sync_parsed_workbook(parsed)

    def sync_parsed_workbook(self, parsed: ParsedWorkbook) -> SyncResult:
        run_id = str(uuid4())
        self._create_schema()
        dim_users_df = pd.DataFrame([asdict(row) for row in parsed.user_dimensions])
        dim_wps_df = pd.DataFrame(
            [{**asdict(row), "wp_label_normalized": row.wp_label_normalized} for row in parsed.wp_dimensions]
        )
        dim_projects_df = pd.DataFrame(
            [
                {
                    "project_id": _project_id(sheet_title),
                    "sheet_title": sheet_title,
                    "project_name": sheet_title.removeprefix("P-"),
                    "is_template": sheet_title == TEMPLATE_SHEET,
                    "project_order": index,
                }
                for index, sheet_title in enumerate(parsed.project_sheet_names, start=1)
                if sheet_title != TEMPLATE_SHEET
            ]
        )
        facts_df, issues_df = self._build_fact_rows(parsed.project_blocks, dim_users_df, dim_wps_df, run_id)
        self._load_table("dim_users", dim_users_df)
        self._load_table("dim_wps", dim_wps_df)
        self._load_table("dim_projects", dim_projects_df)
        self._load_table("fact_planning", facts_df)
        self._load_table("inconsistency_log", issues_df)
        self._run_post_load_checks(run_id)
        issue_count = int(self.connection.execute("SELECT COUNT(*) FROM inconsistency_log").fetchone()[0])
        fact_count = int(self.connection.execute("SELECT COUNT(*) FROM fact_planning").fetchone()[0])
        project_count = len(dim_projects_df.index)
        return SyncResult(
            run_id=run_id,
            status="clean" if issue_count == 0 else "blocked",
            issue_count=issue_count,
            project_count=project_count,
            fact_row_count=fact_count,
        )

    def analytics_monthly_capacity(self) -> pd.DataFrame:
        return self.connection.execute(
            """
            SELECT month_date, userid, SUM(COALESCE(hours, 0)) AS total_hours
            FROM fact_planning
            GROUP BY month_date, userid
            ORDER BY month_date, userid
            """
        ).fetchdf()

    def analytics_budget(self) -> pd.DataFrame:
        return self.connection.execute(
            """
            SELECT
                fp.sheet_title,
                fp.month_date,
                SUM(COALESCE(fp.hours, 0) * COALESCE(du.hourly_rate_overhead, du.hourly_rate, 0)) AS budget
            FROM fact_planning fp
            LEFT JOIN dim_users du ON fp.userid = du.userid
            GROUP BY fp.sheet_title, fp.month_date
            ORDER BY fp.sheet_title, fp.month_date
            """
        ).fetchdf()

    def analytics_fte(self) -> pd.DataFrame:
        return self.connection.execute(
            f"""
            SELECT
                month_date,
                userid,
                SUM(COALESCE(hours, 0)) / {GLOBAL_MONTHLY_HOURS_CAP} AS fte_load
            FROM fact_planning
            GROUP BY month_date, userid
            ORDER BY month_date, userid
            """
        ).fetchdf()

    def issues(self) -> pd.DataFrame:
        return self.connection.execute("SELECT * FROM inconsistency_log ORDER BY severity, sheet_title, cell_ref").fetchdf()

    def _build_fact_rows(
        self,
        blocks: list[ParsedProjectBlock],
        dim_users_df: pd.DataFrame,
        dim_wps_df: pd.DataFrame,
        run_id: str,
    ) -> tuple[pd.DataFrame, pd.DataFrame]:
        users_by_slot = {}
        if not dim_users_df.empty:
            users_by_slot = dim_users_df.set_index("person_slot").to_dict("index")
        wp_codes = set(dim_wps_df["wp_code"].tolist()) if not dim_wps_df.empty else set()
        fact_rows: list[dict] = []
        issue_rows: list[dict] = []
        for block in blocks:
            userid = None
            person_slot = block.person_slot
            if person_slot and person_slot in users_by_slot:
                userid = users_by_slot[person_slot]["userid"]
            else:
                issue_rows.append(
                    self._issue(
                        run_id,
                        "UNKNOWN_USERID",
                        "error",
                        block.sheet_title,
                        f"A{block.source_row}",
                        None,
                        person_slot,
                        None,
                        None,
                        "Unable to resolve userid from person_slot.",
                    )
                )
            wp_code, wp_shortname = _parse_wp_label(block.wp_label_raw)
            if wp_code is None:
                issue_rows.append(
                    self._issue(
                        run_id,
                        "WP_LABEL_PARSE_ERROR",
                        "error",
                        block.sheet_title,
                        f"E{block.source_row}",
                        userid,
                        person_slot,
                        None,
                        None,
                        "Missing WP label.",
                    )
                )
            elif wp_code not in wp_codes:
                issue_rows.append(
                    self._issue(
                        run_id,
                        "UNKNOWN_WP",
                        "error",
                        block.sheet_title,
                        f"E{block.source_row}",
                        userid,
                        person_slot,
                        wp_code,
                        None,
                        f"Unknown WP code '{wp_code}'.",
                    )
                )
            for month_index, (month_date, source_cell, hours) in enumerate(block.monthly_values, start=1):
                fact_rows.append(
                    {
                        "project_id": _project_id(block.sheet_title),
                        "sheet_title": block.sheet_title,
                        "userid": userid,
                        "person_slot": person_slot,
                        "person_name": block.person_name,
                        "role": block.role,
                        "wp_code": wp_code,
                        "wp_shortname": wp_shortname,
                        "month_date": month_date,
                        "month_index": month_index,
                        "hours": hours,
                        "source_cell": source_cell,
                        "source_row": block.source_row,
                        "load_run_id": run_id,
                    }
                )
        return pd.DataFrame(fact_rows), pd.DataFrame(issue_rows)

    def _issue(
        self,
        run_id: str,
        error_code: str,
        severity: str,
        sheet_title: str,
        cell_ref: str | None,
        userid: str | None,
        person_slot: str | None,
        wp_code: str | None,
        month_date,
        message: str,
    ) -> dict:
        return {
            "load_run_id": run_id,
            "error_code": error_code,
            "severity": severity,
            "sheet_title": sheet_title,
            "cell_ref": cell_ref,
            "userid": userid,
            "person_slot": person_slot,
            "wp_code": wp_code,
            "month_date": month_date,
            "message": message,
        }

    def _create_schema(self) -> None:
        self.connection.execute("DROP TABLE IF EXISTS dim_users")
        self.connection.execute("DROP TABLE IF EXISTS dim_wps")
        self.connection.execute("DROP TABLE IF EXISTS dim_projects")
        self.connection.execute("DROP TABLE IF EXISTS fact_planning")
        self.connection.execute("DROP TABLE IF EXISTS inconsistency_log")
        self.connection.execute(
            """
            CREATE TABLE dim_users (
                userid TEXT,
                person_slot TEXT,
                person_name TEXT,
                gender TEXT,
                role TEXT,
                hourly_rate DOUBLE,
                hourly_rate_overhead DOUBLE,
                hour_divider DOUBLE,
                contract_hours DOUBLE,
                sandbox_pct DOUBLE,
                source_sheet_title TEXT,
                source_row INTEGER
            )
            """
        )
        self.connection.execute(
            """
            CREATE TABLE dim_wps (
                wp_code TEXT,
                wp_shortname TEXT,
                wp_long_name TEXT,
                wp_label_normalized TEXT
            )
            """
        )
        self.connection.execute(
            """
            CREATE TABLE dim_projects (
                project_id TEXT,
                sheet_title TEXT,
                project_name TEXT,
                is_template BOOLEAN,
                project_order INTEGER
            )
            """
        )
        self.connection.execute(
            """
            CREATE TABLE fact_planning (
                project_id TEXT,
                sheet_title TEXT,
                userid TEXT,
                person_slot TEXT,
                person_name TEXT,
                role TEXT,
                wp_code TEXT,
                wp_shortname TEXT,
                month_date DATE,
                month_index INTEGER,
                hours DOUBLE,
                source_cell TEXT,
                source_row INTEGER,
                load_run_id TEXT
            )
            """
        )
        self.connection.execute(
            """
            CREATE TABLE inconsistency_log (
                load_run_id TEXT,
                error_code TEXT,
                severity TEXT,
                sheet_title TEXT,
                cell_ref TEXT,
                userid TEXT,
                person_slot TEXT,
                wp_code TEXT,
                month_date DATE,
                message TEXT
            )
            """
        )

    def _load_table(self, table_name: str, dataframe: pd.DataFrame) -> None:
        if dataframe.empty:
            return
        self.connection.register("frame_tmp", dataframe)
        self.connection.execute(f"INSERT INTO {table_name} SELECT * FROM frame_tmp")
        self.connection.unregister("frame_tmp")

    def _run_post_load_checks(self, run_id: str) -> None:
        self.connection.execute(
            f"""
            INSERT INTO inconsistency_log
            SELECT
                '{run_id}' AS load_run_id,
                'DUPLICATE_USERID' AS error_code,
                'error' AS severity,
                source_sheet_title AS sheet_title,
                NULL AS cell_ref,
                userid,
                NULL AS person_slot,
                NULL AS wp_code,
                NULL AS month_date,
                'Duplicate userid in dimension table.' AS message
            FROM dim_users
            GROUP BY userid, source_sheet_title
            HAVING userid IS NOT NULL AND COUNT(*) > 1
            """
        )
        self.connection.execute(
            f"""
            INSERT INTO inconsistency_log
            SELECT
                '{run_id}' AS load_run_id,
                'GLOBAL_MONTH_CAP_EXCEEDED' AS error_code,
                'error' AS severity,
                NULL AS sheet_title,
                NULL AS cell_ref,
                userid,
                NULL AS person_slot,
                NULL AS wp_code,
                month_date,
                'Monthly hour cap exceeded.' AS message
            FROM fact_planning
            WHERE userid IS NOT NULL
            GROUP BY userid, month_date
            HAVING SUM(COALESCE(hours, 0)) > {GLOBAL_MONTHLY_HOURS_CAP}
            """
        )
