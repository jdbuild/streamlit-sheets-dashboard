from __future__ import annotations

from dataclasses import dataclass
from datetime import date


BLOCK_START_ROWS = (20, 35, 50, 65, 80, 95, 110, 125, 140, 155, 170)
MONTH_START_COL = "T"
MONTH_END_COL = "CM"
USER_RANGE = "A195:J207"
WP_RANGE = "A210:C218"
MONTH_HEADER_ROW = 19
PERSON_BLOCK_HEIGHT = 15
GLOBAL_MONTHLY_HOURS_CAP = 143.33
TEMPLATE_SHEET = "P-XYTemplate"


@dataclass(slots=True)
class UserDimensionRow:
    person_slot: str
    person_name: str
    userid: str
    gender: str | None
    role: str | None
    hourly_rate: float | None
    hourly_rate_overhead: float | None
    hour_divider: float | None
    contract_hours: float | None
    sandbox_pct: float | None
    source_sheet_title: str
    source_row: int


@dataclass(slots=True)
class WPDimensionRow:
    wp_code: str
    wp_shortname: str | None
    wp_long_name: str | None

    @property
    def wp_label_normalized(self) -> str:
        if self.wp_shortname:
            return f"{self.wp_code} {self.wp_shortname}"
        return self.wp_code


@dataclass(slots=True)
class PlanningFactRow:
    project_id: str
    sheet_title: str
    userid: str | None
    person_slot: str | None
    person_name: str | None
    role: str | None
    wp_code: str | None
    wp_shortname: str | None
    month_date: date
    month_index: int
    hours: float | None
    source_cell: str
    source_row: int


@dataclass(slots=True)
class SyncResult:
    run_id: str
    status: str
    issue_count: int
    project_count: int
    fact_row_count: int
